# ============================================================
#  assign_names.py – экспорт блоков, распределение портов,
#  формирование NAME. Версия 4.1
#
#  Что делает файл:
#    1. Проверяет настройки (конфликты портов, дубликаты).
#    2. Подключается к AutoCAD и собирает данные по всем
#       включённым типам блоков.
#    3. Проверяет, хватает ли портов в каждом шкафу.
#    4. Распределяет порты по патч-панелям:
#         - для основных блоков (камеры, AP и т.д.),
#         - для блоков особого режима (розетки, RJ-45),
#           которые идут ПОСЛЕ основных.
#    5. Формирует два CSV-файла:
#         - export_final.csv  – полный отчёт по всем блокам,
#         - import.csv        – Handle + новое имя.
# ============================================================

import sys
import os
import importlib.util
import pandas as pd
import pythoncom
import time
import gc
import win32com.client
from win32com.client import VARIANT


# ------------------------------------------------------------
# 1. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ПАРСИНГА НАСТРОЕК
# ------------------------------------------------------------

def parse_skip_ports(skip_list):
    """
    Преобразует список строк SKIP_PORTS в словарь.
    Формат строки: "шкаф,номер_панели,список_портов"
    Пример: "3.1C6,1,25-29" -> {("3.1C6", 1): [(25, 29)]}
    """
    result = {}
    for line in skip_list:
        if not line.strip() or line.startswith("#"):
            continue
        parts = line.split(",")
        if len(parts) < 3:
            continue
        cabinet = parts[0].strip()
        try:
            panel = int(parts[1].strip())
        except ValueError:
            continue
        ports_str = ",".join(parts[2:]).strip()
        items = []
        for token in ports_str.split(","):
            token = token.strip()
            if "-" in token:
                start, end = map(int, token.split("-"))
                items.append((start, end))
            else:
                items.append(int(token))
        result[(cabinet, panel)] = items
    return result


def parse_port_ranges(port_ranges):
    """
    Преобразует список диапазонов [[1, 18], [25, 42]]
    в плоский список портов [1, 2, ..., 18, 25, ..., 42].
    """
    result = []
    for r in port_ranges:
        if isinstance(r, (list, tuple)) and len(r) == 2:
            if r[0] == r[1]:
                result.append(r[0])
            else:
                result.extend(range(r[0], r[1] + 1))
        elif isinstance(r, int):
            result.append(r)
    return result


def get_available_ports(cabinet, panel_number, allowed_ports, skip_ports_dict):
    """
    Возвращает список доступных портов для указанного шкафа и панели,
    исключая порты из SKIP_PORTS.
    """
    skip_list_raw = skip_ports_dict.get((cabinet, panel_number), [])
    skip_ports = set()
    for item in skip_list_raw:
        if isinstance(item, (tuple, list)) and len(item) == 2:
            start, end = item
            skip_ports.update(range(start, end + 1))
        else:
            skip_ports.add(item)
    return [p for p in allowed_ports if p not in skip_ports]


# ------------------------------------------------------------
# 2. ВАЛИДАЦИЯ НАСТРОЕК
# ------------------------------------------------------------

def validate_port_conflicts(block_configs):
    """
    Проверяет, что диапазоны портов у включённых блоков
    не пересекаются между собой.
    """
    errors = []
    enabled = [b for b in block_configs if b.get("enabled") and b.get("block_name")]
    for i in range(len(enabled)):
        for j in range(i + 1, len(enabled)):
            ports_i = set(parse_port_ranges(enabled[i]["port_ranges"]))
            ports_j = set(parse_port_ranges(enabled[j]["port_ranges"]))
            overlap = ports_i & ports_j
            if overlap:
                errors.append(
                    f"Блоки '{enabled[i]['block_name']}' и '{enabled[j]['block_name']}' "
                    f"имеют пересекающиеся порты: {sorted(overlap)}"
                )
    return errors


def validate_special_mode(block_configs, special_enabled, special_names):
    """
    Проверяет, что блоки из особого режима не дублируют
    блоки, включённые в основном режиме.
    """
    errors = []
    if not special_enabled:
        return errors
    enabled_names = {b["block_name"] for b in block_configs
                     if b.get("enabled") and b.get("block_name")}
    for name in special_names:
        if name in enabled_names:
            errors.append(
                f"Блок '{name}' включён ОДНОВРЕМЕННО в основном режиме и в особом. "
                f"Отключите его в таблице основных блоков или уберите из особого режима."
            )
    return errors


# ------------------------------------------------------------
# 3. РАСЧЁТ СТАРТОВОЙ ПОЗИЦИИ ДЛЯ ОСОБОГО РЕЖИМА
# ------------------------------------------------------------

def get_used_ports_on_panel(df_all, cabinet, panel, tag_cabinet):
    """
    Возвращает множество портов, занятых основными блоками
    на указанной панели указанного шкафа.

    Используется для определения последнего занятого порта
    при «продолжении с последней патч-панели».
    """
    mask = (df_all["_cabinet"] == cabinet) & (df_all["_panel"] == panel)
    df_panel = df_all[mask]
    return set(df_panel["_port"].astype(int).tolist())


def get_special_start_point(df_all, cabinet, enabled_blocks, skip_ports_dict,
                             max_panels, continue_last_panel, skip_count,
                             special_port_ranges):
    """
    Определяет стартовую точку для особого режима в конкретном шкафу.

    Возвращает:
        (start_panel, min_port_on_start_panel):
            start_panel             — номер панели, с которой начинать;
            min_port_on_start_panel — минимальный порт на этой панели,
                                      с которого разрешено начинать.
                                      -1 означает «без ограничения» —
                                      брать первый доступный порт.

    Логика:
        ЕСЛИ continue_last_panel = False:
            start_panel = max_used_panel + 1
            min_port = -1 (без ограничений, берём первый доступный порт)
        ИНАЧЕ:
            start_panel = max_used_panel (последняя, где есть основные блоки)
            last_port = максимальный занятый порт на этой панели
            min_port = last_port + 1 + skip_count
            Если min_port выходит за пределы панели — переходим на
            следующую панель с min_port = -1.
    """
    # --- 1. Находим максимальную панель, занятую основными блоками ---
    max_used_panel = 0
    for block_cfg in enabled_blocks:
        bname = block_cfg["block_name"]
        df_b = df_all[(df_all["_device_type"] == bname) &
                      (df_all["_cabinet"] == cabinet)]
        if df_b.empty:
            continue
        # Считаем, сколько панелей занял этот блок в этом шкафу
        allowed = parse_port_ranges(block_cfg["port_ranges"])
        total = 0
        for panel in range(1, max_panels + 1):
            available = get_available_ports(cabinet, panel, allowed, skip_ports_dict)
            total += len(available)
            if total >= len(df_b):
                if panel > max_used_panel:
                    max_used_panel = panel
                break

    # --- 2. Если особый режим не продолжает последнюю панель ---
    if not continue_last_panel:
        return max_used_panel + 1, -1

    # --- 3. Особый режим продолжает последнюю панель ---
    # Если вообще нет основных блоков — стартуем с 1-й панели
    if max_used_panel == 0:
        return 1, -1

    # Находим последний занятый порт на последней панели
    used_ports = get_used_ports_on_panel(df_all, cabinet, max_used_panel,
                                         "_cabinet")
    if not used_ports:
        # Странная ситуация: на панели нет занятых портов.
        # Считаем, что стартуем с 1-й панели без ограничений.
        return max_used_panel, -1

    last_port = max(used_ports)
    min_port = last_port + 1 + skip_count

    # Проверяем: есть ли на этой панели доступные порты >= min_port
    allowed_special = parse_port_ranges(special_port_ranges)
    available_on_panel = get_available_ports(cabinet, max_used_panel,
                                             allowed_special, skip_ports_dict)
    candidates = [p for p in available_on_panel if p >= min_port]

    if candidates:
        # Есть подходящие порты — продолжаем с этой панели
        return max_used_panel, min_port
    else:
        # На этой панели места нет — переходим на следующую
        return max_used_panel + 1, -1


# ------------------------------------------------------------
# 4. ПРОВЕРКА ЁМКОСТИ ШКАФОВ
# ------------------------------------------------------------

def check_capacity(cabinet, num_devices, port_ranges, skip_ports_dict,
                   max_panels, start_panel=1, min_port=-1):
    """
    Проверяет, хватает ли портов в шкафу для всех устройств.

    Параметры:
        cabinet        — имя шкафа
        num_devices    — количество устройств
        port_ranges    — диапазоны портов для этого типа блока
        skip_ports_dict — словарь пропусков
        max_panels     — максимум панелей
        start_panel    — с какой панели начинать
        min_port       — минимальный порт на стартовой панели
                         (-1 = без ограничений)

    Возвращает:
        (enough, total_ports, details)
    """
    allowed = parse_port_ranges(port_ranges)
    total = 0
    details = []
    for panel in range(start_panel, max_panels + 1):
        available = get_available_ports(cabinet, panel, allowed, skip_ports_dict)
        # На стартовой панели применяем фильтр по min_port
        if panel == start_panel and min_port > 0:
            available = [p for p in available if p >= min_port]
        total += len(available)
        details.append(f"    Панель {panel}: {len(available)} портов")
        if total >= num_devices:
            break
    return total >= num_devices, total, details


def get_max_used_panel(cabinet, num_devices, port_ranges, skip_ports_dict, max_panels):
    """Определяет, на какой панели закончатся устройства."""
    allowed = parse_port_ranges(port_ranges)
    total = 0
    for panel in range(1, max_panels + 1):
        available = get_available_ports(cabinet, panel, allowed, skip_ports_dict)
        total += len(available)
        if total >= num_devices:
            return True, panel
    return False, max_panels


# ------------------------------------------------------------
# 5. ПОЛУЧЕНИЕ БЛОКОВ ЧЕРЕЗ SELECTIONSET
# ------------------------------------------------------------

def get_blocks_by_name(doc, block_name, max_retries=3):
    """
    Возвращает список вхождений блоков с заданным именем.
    Использует SelectionSet с фильтром и повторные попытки
    при COM-ошибках «Вызов был отклонен».
    """
    for attempt in range(max_retries):
        try:
            acSelectionSetAll = 5
            ss_name = f"TempSel_{block_name}"
            try:
                ss = doc.SelectionSets.Add(ss_name)
            except:
                ss = doc.SelectionSets.Item(ss_name)
                ss.Clear()

            filter_type = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 2])
            filter_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT", block_name])

            ss.Select(acSelectionSetAll, None, None, filter_type, filter_data)
            blocks = [blk for blk in ss]
            ss.Clear()
            return blocks
        except Exception as e:
            if attempt < max_retries - 1:
                wait = 1.0 + attempt * 1.0
                print(f"    COM-ошибка при чтении '{block_name}', "
                      f"повтор через {wait:.0f} сек (попытка {attempt+2}/{max_retries})...")
                time.sleep(wait)
                gc.collect()
            else:
                raise e


# ------------------------------------------------------------
# 6. СБОР ДАННЫХ ПО ОДНОМУ ТИПУ БЛОКА
# ------------------------------------------------------------

def collect_blocks_data(doc, block_config, tag_cabinet, tag_floor, tag_name):
    """
    Собирает данные по одному типу блока:
    Handle, X, Y, Z, значения атрибутов (шкаф, этаж, имя),
    имя типа блока (_device_type) и отображаемое имя (_display_name).
    """
    block_name = block_config["block_name"]
    display_name = block_config.get("display_name", block_name)
    data = []
    try:
        target_blocks = get_blocks_by_name(doc, block_name)
    except Exception as e:
        print(f"Ошибка при поиске блоков '{block_name}': {e}")
        return data

    for obj in target_blocks:
        try:
            handle = obj.Handle
            if not obj.HasAttributes:
                continue
            ins = obj.InsertionPoint
            x, y, z = ins[0], ins[1], ins[2]
            atts = obj.GetAttributes()
            if isinstance(atts, VARIANT):
                att_list = list(atts.value)
            else:
                att_list = list(atts)
            row = {
                "Handle": handle,
                "X": x,
                "Y": y,
                "Z": z,
                tag_cabinet: "",
                tag_floor: "",
                tag_name: "",
                "_device_type": block_name,
                "_display_name": display_name,
            }
            for attr in att_list:
                tag = attr.TagString
                if tag in row:
                    row[tag] = attr.TextString
            data.append(row)
        except Exception as e:
            print(f"Ошибка при обработке блока {block_name}: {e}")
            continue
    return data


# ------------------------------------------------------------
# 7. ОБРАБОТКА ОДНОГО ТИПА БЛОКА
# ------------------------------------------------------------

def process_device_type(df, device_type, port_ranges, skip_ports_dict,
                        max_panels, name_format, tag_cabinet, tag_floor,
                        start_panel=1, min_port=-1):
    """
    Обрабатывает один тип блока: сортирует блоки внутри шкафа,
    распределяет порты, формирует новое имя.

    Параметры:
        start_panel — с какой панели начинать
        min_port    — минимальный порт на стартовой панели
                      (-1 = без ограничений, используется в особом режиме
                      при «продолжении с последней панели»)
    """
    allowed_ports = parse_port_ranges(port_ranges)
    result_rows = []
    max_used_panel = start_panel - 1

    if df.empty:
        return pd.DataFrame(), max_used_panel

    for cabinet, group in df.groupby(tag_cabinet):
        # Сортировка внутри шкафа: этаж (по возрастанию),
        # Y (сверху вниз), X (слева направо)
        group_sorted = group.sort_values(
            by=[tag_floor, "Y", "X"],
            ascending=[True, False, True]
        )

        panel = start_panel
        available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
        # Если мы на стартовой панели и есть ограничение по min_port —
        # отбрасываем порты, меньшие min_port
        if panel == start_panel and min_port > 0:
            available = [p for p in available if p >= min_port]
        # Ищем первую панель, где есть доступные порты
        while not available and panel <= max_panels:
            panel += 1
            available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
        if not available:
            raise Exception(
                f"Шкаф {cabinet}, тип '{device_type}': "
                f"нет доступных портов в диапазоне {start_panel}..{max_panels}."
            )

        port_index = 0
        for idx, row in group_sorted.iterrows():
            if port_index >= len(available):
                panel += 1
                port_index = 0
                available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                while not available and panel <= max_panels:
                    panel += 1
                    available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                if not available:
                    raise Exception(
                        f"Шкаф {cabinet}, тип '{device_type}': "
                        f"не хватило панелей (макс. {max_panels})."
                    )

            port = available[port_index]
            new_row = row.copy()
            new_row["_panel"] = panel
            new_row["_port"] = port
            new_row["_name"] = name_format.format(
                cabinet=cabinet, panel=panel, port=port
            )
            result_rows.append(new_row)
            port_index += 1

        if panel > max_used_panel:
            max_used_panel = panel

    return pd.DataFrame(result_rows), max_used_panel


# ------------------------------------------------------------
# 8. ОСНОВНАЯ ФУНКЦИЯ
# ------------------------------------------------------------

def main(wait_for_exit=True):
    """
    Главная функция: загружает config, подключается к AutoCAD,
    валидирует настройки, собирает данные, проверяет ёмкость,
    распределяет порты, формирует NAME, сохраняет CSV.
    """
    # --- Динамическая загрузка config ---
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(application_path, 'config.py')
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    pythoncom.CoInitialize()
    start_total = time.time()

    try:
        # --- 1. Читаем настройки ---
        TAG_CABINET = config.TAG_CABINET
        TAG_FLOOR = config.TAG_FLOOR
        TAG_NAME = config.TAG_NAME
        NAME_FORMAT = config.NAME_FORMAT
        MAX_PANELS = config.MAX_PANELS_PER_CABINET
        SKIP_PORTS_LIST = config.SKIP_PORTS
        BLOCK_CONFIGS = config.BLOCK_CONFIGS
        SPECIAL_MODE_ENABLED = getattr(config, "SPECIAL_MODE_ENABLED", False)
        SPECIAL_BLOCK_NAMES = getattr(config, "SPECIAL_BLOCK_NAMES", [])
        SPECIAL_PORT_RANGES = getattr(config, "SPECIAL_PORT_RANGES", [])
        # Новые параметры особого режима
        SPECIAL_CONTINUE_LAST_PANEL = getattr(config, "SPECIAL_CONTINUE_LAST_PANEL", False)
        SPECIAL_SKIP_PORTS = getattr(config, "SPECIAL_SKIP_PORTS", 0)

        # --- 2. Валидация настроек ---
        print("Проверка настроек...")
        errors = validate_port_conflicts(BLOCK_CONFIGS)
        special_errors = validate_special_mode(
            BLOCK_CONFIGS, SPECIAL_MODE_ENABLED, SPECIAL_BLOCK_NAMES
        )
        errors.extend(special_errors)

        if errors:
            print("\n❌ Обнаружены ошибки в настройках:")
            for err in errors:
                print(f"   • {err}")
            print("\nИсправьте настройки и запустите снова.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit(1)
        print("Ошибок в настройках не обнаружено.\n")

        # --- 3. Подключение к AutoCAD ---
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # --- 4. Парсим пропуски портов ---
        skip_ports_dict = parse_skip_ports(SKIP_PORTS_LIST)
        print("Загруженные пропуски портов:")
        if skip_ports_dict:
            for key, value in skip_ports_dict.items():
                print(f"  {key} -> {value}")
        else:
            print("  Пропуски портов отсутствуют.")

        # --- 5. Собираем данные по всем включённым блокам ---
        all_data = []
        enabled_blocks = [b for b in BLOCK_CONFIGS if b.get("enabled") and b.get("block_name")]
        print(f"\nВключено блоков для обработки: {len(enabled_blocks)}")
        for block_cfg in enabled_blocks:
            bname = block_cfg["block_name"]
            print(f"  Чтение блоков '{bname}'...")
            data = collect_blocks_data(doc, block_cfg, TAG_CABINET, TAG_FLOOR, TAG_NAME)
            print(f"    Найдено: {len(data)}")
            all_data.extend(data)
            # Пауза между типами блоков — лечит COM-ошибку «Вызов был отклонен»
            time.sleep(0.5)
            gc.collect()

        if not all_data:
            print("Не найдено ни одного блока. Завершение.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit()

        # --- 6. Общий DataFrame ---
        df_all = pd.DataFrame(all_data)
        print(f"\nВсего собрано блоков: {len(df_all)}")

        # --- 7. Фильтрация ---
        df_all = df_all.dropna(subset=[TAG_CABINET, TAG_FLOOR])
        df_all = df_all[df_all[TAG_CABINET].astype(str).str.strip() != ""]
        df_all = df_all[df_all[TAG_FLOOR].astype(str).str.strip() != ""]
        df_all[TAG_FLOOR] = pd.to_numeric(df_all[TAG_FLOOR], errors='coerce')
        df_all = df_all.dropna(subset=[TAG_FLOOR])
        print(f"После фильтрации (есть шкаф и этаж): {len(df_all)}")

        df_all["_cabinet"] = df_all[TAG_CABINET]
        df_all["_floor"] = df_all[TAG_FLOOR]

        # --- 8. ОБРАБОТКА ОСНОВНЫХ БЛОКОВ ---
        # (сначала обрабатываем их, потому что для особого режима
        # нужно знать, какие панели/порты уже заняты)
        result_dfs = []

        for block_cfg in enabled_blocks:
            bname = block_cfg["block_name"]
            print(f"\n--- Обработка '{bname}' ---")
            df_type = df_all[df_all["_device_type"] == bname].copy()
            if df_type.empty:
                print(f"  Нет данных для '{bname}'.")
                continue

            processed_df, max_panel = process_device_type(
                df_type, bname, block_cfg["port_ranges"],
                skip_ports_dict, MAX_PANELS, NAME_FORMAT,
                TAG_CABINET, TAG_FLOOR,
                start_panel=1
            )
            result_dfs.append(processed_df)
            print(f"  Обработано: {len(processed_df)}, последняя панель: {max_panel}")

        # Собираем промежуточный DataFrame с уже обработанными
        # основными блоками — нужен для определения последних
        # занятых портов.
        df_main_processed = pd.concat(result_dfs, ignore_index=True) if result_dfs else pd.DataFrame()

        # --- 9. ПРОВЕРКА ЁМКОСТИ ШКАФОВ ---
        print("\nПроверка ёмкости шкафов...")
        capacity_errors = []

        # 9.1. Ёмкость для основных блоков
        for block_cfg in enabled_blocks:
            bname = block_cfg["block_name"]
            df_type = df_all[df_all["_device_type"] == bname]
            if df_type.empty:
                continue
            for cabinet, group in df_type.groupby(TAG_CABINET):
                num_devices = len(group)
                enough, total, details = check_capacity(
                    cabinet, num_devices, block_cfg["port_ranges"],
                    skip_ports_dict, MAX_PANELS, start_panel=1
                )
                if not enough:
                    err = (
                        f"Шкаф '{cabinet}', блок '{bname}': "
                        f"устройств {num_devices}, доступно портов {total} "
                        f"(при макс. {MAX_PANELS} панелях)."
                    )
                    capacity_errors.append((err, details))

        # 9.2. Ёмкость для особого режима
        if SPECIAL_MODE_ENABLED and SPECIAL_BLOCK_NAMES and SPECIAL_PORT_RANGES:
            special_names = set(SPECIAL_BLOCK_NAMES)
            df_special_all = df_all[df_all["_device_type"].isin(special_names)]
            if not df_special_all.empty:
                for cabinet, group in df_special_all.groupby(TAG_CABINET):
                    num_devices = len(group)
                    # Определяем стартовую точку с учётом новых параметров
                    start_panel, min_port = get_special_start_point(
                        df_main_processed, cabinet, enabled_blocks,
                        skip_ports_dict, MAX_PANELS,
                        SPECIAL_CONTINUE_LAST_PANEL, SPECIAL_SKIP_PORTS,
                        SPECIAL_PORT_RANGES
                    )
                    enough, total, details = check_capacity(
                        cabinet, num_devices, SPECIAL_PORT_RANGES,
                        skip_ports_dict, MAX_PANELS,
                        start_panel=start_panel, min_port=min_port
                    )
                    if not enough:
                        err = (
                            f"Шкаф '{cabinet}', особый режим: "
                            f"устройств {num_devices}, доступно портов {total} "
                            f"начиная с панели {start_panel}"
                            + (f", порт >= {min_port}" if min_port > 0 else "")
                            + "."
                        )
                        capacity_errors.append((err, details))

        if capacity_errors:
            print("\n❌ Не хватает портов в шкафах:")
            for err, details in capacity_errors:
                print(f"   • {err}")
                for d in details:
                    print(d)
            print("\nУвеличьте MAX_PANELS_PER_CABINET или проверьте пропуски/диапазоны.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit(1)
        print("Ёмкости шкафов достаточно.")

        # --- 10. ОСОБЫЙ РЕЖИМ ---
        if SPECIAL_MODE_ENABLED and SPECIAL_BLOCK_NAMES and SPECIAL_PORT_RANGES:
            print("\n=== ОСОБЫЙ РЕЖИМ ===")

            # Собираем данные по блокам особого режима
            special_all = []
            for sp_name in SPECIAL_BLOCK_NAMES:
                print(f"  Чтение блоков '{sp_name}'...")
                cfg = next(
                    (b for b in BLOCK_CONFIGS if b["block_name"] == sp_name),
                    {"block_name": sp_name, "port_ranges": SPECIAL_PORT_RANGES}
                )
                data = collect_blocks_data(doc, cfg, TAG_CABINET, TAG_FLOOR, TAG_NAME)
                print(f"    Найдено: {len(data)}")
                special_all.extend(data)
                time.sleep(0.5)
                gc.collect()

            if special_all:
                df_special = pd.DataFrame(special_all)
                df_special = df_special.dropna(subset=[TAG_CABINET, TAG_FLOOR])
                df_special = df_special[df_special[TAG_CABINET].astype(str).str.strip() != ""]
                df_special = df_special[df_special[TAG_FLOOR].astype(str).str.strip() != ""]
                df_special[TAG_FLOOR] = pd.to_numeric(df_special[TAG_FLOOR], errors='coerce')
                df_special = df_special.dropna(subset=[TAG_FLOOR])
                df_special["_cabinet"] = df_special[TAG_CABINET]
                df_special["_floor"] = df_special[TAG_FLOOR]
                df_special["_special_order"] = df_special["_device_type"].apply(
                    lambda x: SPECIAL_BLOCK_NAMES.index(x) if x in SPECIAL_BLOCK_NAMES else 999
                )

                for cabinet, group in df_special.groupby("_cabinet"):
                    # Определяем стартовую точку с учётом новых параметров
                    start_panel, min_port = get_special_start_point(
                        df_main_processed, cabinet, enabled_blocks,
                        skip_ports_dict, MAX_PANELS,
                        SPECIAL_CONTINUE_LAST_PANEL, SPECIAL_SKIP_PORTS,
                        SPECIAL_PORT_RANGES
                    )

                    # Сортируем блоки особого режима
                    group_sorted = group.sort_values(
                        by=["_special_order", TAG_FLOOR, "Y", "X"],
                        ascending=[True, True, False, True]
                    )

                    allowed_ports = parse_port_ranges(SPECIAL_PORT_RANGES)
                    panel = start_panel
                    available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                    # Фильтр по min_port на стартовой панели
                    if panel == start_panel and min_port > 0:
                        available = [p for p in available if p >= min_port]
                    while not available and panel <= MAX_PANELS:
                        panel += 1
                        available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                    if not available:
                        print(f"  ⚠ Шкаф {cabinet}: нет доступных портов для особого режима.")
                        continue

                    port_index = 0
                    special_rows = []
                    for idx, row in group_sorted.iterrows():
                        if port_index >= len(available):
                            panel += 1
                            port_index = 0
                            available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                            while not available and panel <= MAX_PANELS:
                                panel += 1
                                available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                            if not available:
                                raise Exception(
                                    f"Шкаф {cabinet}, особый режим: "
                                    f"не хватило панелей (макс. {MAX_PANELS})."
                                )
                        port = available[port_index]
                        new_row = row.copy()
                        new_row["_panel"] = panel
                        new_row["_port"] = port
                        new_row["_name"] = NAME_FORMAT.format(
                            cabinet=cabinet, panel=panel, port=port
                        )
                        special_rows.append(new_row)
                        port_index += 1

                    if special_rows:
                        result_dfs.append(pd.DataFrame(special_rows))
                        print(f"  Шкаф {cabinet}: обработано {len(special_rows)} блоков, "
                              f"старт с панели {start_panel}, последняя панель: {panel}")

        # --- 11. Финальный DataFrame ---
        df_final = pd.concat(result_dfs, ignore_index=True)
        print(f"\nВсего обработано блоков: {len(df_final)}")

        # --- 12. export_final.csv ---
        export_cols = [
            "Handle", TAG_CABINET, TAG_FLOOR,
            "_panel", "_port", "_name", "_device_type", "_display_name"
        ]
        export_cols = [c for c in export_cols if c in df_final.columns]
        df_export = df_final[export_cols].copy()
        df_export = df_export.rename(columns={
            TAG_CABINET: "Шкаф",
            TAG_FLOOR: "Этаж",
            "_panel": "Панель",
            "_port": "Порт",
            "_name": "Имя",
            "_device_type": "Тип",
        })
        df_export.to_csv("export_final.csv", index=False, encoding="utf-8-sig")
        print("Сохранён export_final.csv")

        # --- 13. import.csv ---
        df_import = df_final[["Handle", "_name"]].rename(columns={"_name": TAG_NAME})
        df_import.to_csv("import.csv", index=False, encoding="utf-8-sig")
        print("Сохранён import.csv")

        elapsed = time.time() - start_total
        print(f"\nОбщее время выполнения: {elapsed:.2f} сек")

        gc.collect()

    except Exception as e:
        print(f"\n❌ Ошибка в основном коде: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")


if __name__ == "__main__":
    main(wait_for_exit=True)
