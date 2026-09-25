# ============================================================
#  mark_new_equipment.py – доназначение имён для НОВОГО оборудования.
#  Версия 4.2
#
#  Что делает:
#    1. Подключается к AutoCAD.
#    2. Читает блоки из BLOCK_CONFIGS.
#    3. Отделяет НОВЫЕ блоки (NAME == PLACEHOLDER_MARKER) от старых.
#    4. Строит карту занятых портов по старым блокам.
#    5. Назначает свободные порты новым блокам:
#         • FILL_GAPS=False — строго после максимума занятых портов.
#         • FILL_GAPS=True  — заполнять свободные «дыры».
#    6. Формирует import.csv только для НОВЫХ блоков.
#    7. Обновляет export_final.csv (старые + новые).
#
#  Важно:
#    • Старые блоки НЕ перемаркировываются, NAME у них не меняется.
#    • Если портов не хватает — программа завершается с ошибкой.
# ============================================================

import sys
import os
import importlib.util
import time
import gc
import pandas as pd
import pythoncom
import win32com.client
from win32com.client import VARIANT


# ------------------------------------------------------------
# 1. Вспомогательные функции (те же, что в read_current_data)
# ------------------------------------------------------------

def parse_name_to_panel_port(name):
    """
    Парсит имя вида "8.1A1/01.4" → (panel_int, port_int).
    Возвращает (None, None), если не получилось.
    """
    if not name or "/" not in name or "." not in name:
        return None, None
    try:
        after_slash = name.split("/", 1)[1]
        panel_str, port_str = after_slash.split(".", 1)
        return int(panel_str), int(port_str)
    except:
        return None, None


def parse_skip_ports(skip_list):
    """
    Преобразует список строк SKIP_PORTS в словарь.
    Формат строки: "шкаф,номер_панели,список_портов".
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
    Преобразует [[1, 18], [25, 42]] в плоский список [1..18, 25..42].
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


def get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict):
    """
    Возвращает список разрешённых портов на панели,
    исключая пропущенные в SKIP_PORTS.
    """
    skip_raw = skip_ports_dict.get((cabinet, panel), [])
    skip = set()
    for item in skip_raw:
        if isinstance(item, (tuple, list)) and len(item) == 2:
            skip.update(range(item[0], item[1] + 1))
        else:
            skip.add(item)
    return [p for p in allowed_ports if p not in skip]


# ------------------------------------------------------------
# 2. Чтение блоков из AutoCAD
# ------------------------------------------------------------

def get_blocks_by_name(doc, block_name, max_retries=3):
    """Возвращает вхождения блоков с заданным именем."""
    for attempt in range(max_retries):
        try:
            ss_name = f"MarkNew_{block_name}"
            try:
                ss = doc.SelectionSets.Add(ss_name)
            except:
                ss = doc.SelectionSets.Item(ss_name)
                ss.Clear()
            filter_type = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 2])
            filter_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                                  ["INSERT", block_name])
            ss.Select(5, None, None, filter_type, filter_data)
            blocks = [blk for blk in ss]
            ss.Clear()
            return blocks
        except Exception as e:
            if attempt < max_retries - 1:
                wait = 1.0 + attempt * 1.0
                print(f"    COM-ошибка чтения '{block_name}', пауза {wait:.0f} сек...")
                time.sleep(wait)
                gc.collect()
            else:
                raise e


def collect_blocks_data(doc, block_config, tag_cabinet, tag_floor, tag_name):
    """
    Собирает данные по блоку. Возвращает список словарей.
    Дополнительно парсит панель и порт из NAME.
    """
    block_name = block_config["block_name"]
    display_name = block_config.get("display_name", block_name)
    data = []

    try:
        target_blocks = get_blocks_by_name(doc, block_name)
    except Exception as e:
        print(f"  Ошибка чтения '{block_name}': {e}")
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
                "X": x, "Y": y,
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
            # Парсим NAME
            panel, port = parse_name_to_panel_port(row.get(tag_name, ""))
            row["_panel"] = panel if panel is not None else 0
            row["_port"] = port if port is not None else 0
            data.append(row)
        except Exception as e:
            print(f"  Ошибка обработки блока {block_name}: {e}")
            continue
    return data


# ------------------------------------------------------------
# 3. Поиск следующего свободного порта
# ------------------------------------------------------------

def find_next_free_port(cabinet, allowed_ports, skip_ports_dict,
                        max_panels, occupied_set, start_after=None):
    """
    Ищет следующий свободный порт для нового оборудования.

    Параметры:
        cabinet         — имя шкафа
        allowed_ports   — список разрешённых портов для типа блока
        skip_ports_dict — пропуски портов
        max_panels      — максимум панелей в шкафу
        occupied_set    — множество занятых портов: {(panel, port), ...}
        start_after     — (panel, port), с которого начинать поиск
                          (для строгого режима — максимум занятых).
                          None — искать с самого начала.

    Возвращает (panel, port) или (None, None), если не нашлось.
    """
    for panel in range(1, max_panels + 1):
        available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
        for port in available:
            # В строгом режиме — пропускаем всё, что <= start_after
            if start_after is not None and (panel, port) <= start_after:
                continue
            # Порт уже занят?
            if (panel, port) in occupied_set:
                continue
            return panel, port
    return None, None


# ------------------------------------------------------------
# 4. Основная функция
# ------------------------------------------------------------

def main(wait_for_exit=True):
    """
    Доназначение имён для нового оборудования.
    """
    start_total = time.time()

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
    try:
        # --- 1. Настройки ---
        TAG_CABINET = config.TAG_CABINET
        TAG_FLOOR = config.TAG_FLOOR
        TAG_NAME = config.TAG_NAME
        NAME_FORMAT = config.NAME_FORMAT
        MAX_PANELS = config.MAX_PANELS_PER_CABINET
        SKIP_PORTS_LIST = config.SKIP_PORTS
        BLOCK_CONFIGS = config.BLOCK_CONFIGS

        # Новые параметры доназначения
        PLACEHOLDER = getattr(config, "PLACEHOLDER_MARKER", "Пустой")
        FILL_GAPS = getattr(config, "FILL_GAPS", False)

        # --- 2. Подключение к AutoCAD ---
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")
        mode_str = "заполнять свободные порты" if FILL_GAPS else "строго после максимума"
        print(f"РЕЖИМ ДОНАЗНАЧЕНИЯ: {mode_str}")
        print(f"Маркер нового блока: '{PLACEHOLDER}'\n")

        # --- 3. Сбор данных ---
        all_data = []
        enabled_blocks = [b for b in BLOCK_CONFIGS
                          if b.get("enabled") and b.get("block_name")]
        print(f"Включено блоков для чтения: {len(enabled_blocks)}")

        for block_cfg in enabled_blocks:
            bname = block_cfg["block_name"]
            print(f"  Чтение блоков '{bname}'...")
            data = collect_blocks_data(doc, block_cfg, TAG_CABINET, TAG_FLOOR, TAG_NAME)
            print(f"    Найдено: {len(data)}")
            all_data.extend(data)
            time.sleep(0.3)
            gc.collect()

        if not all_data:
            print("Не найдено ни одного блока. Завершение.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit()

        df = pd.DataFrame(all_data)
        print(f"\nВсего собрано блоков: {len(df)}")

        # --- 4. Фильтрация ---
        df = df.dropna(subset=[TAG_CABINET, TAG_FLOOR])
        df = df[df[TAG_CABINET].astype(str).str.strip() != ""]
        df = df[df[TAG_FLOOR].astype(str).str.strip() != ""]
        df[TAG_FLOOR] = pd.to_numeric(df[TAG_FLOOR], errors='coerce')
        df = df.dropna(subset=[TAG_FLOOR])
        print(f"После фильтрации: {len(df)}")

        # --- 5. Разделяем на старые и новые ---
        # Новый = NAME равен маркеру-пустышке.
        # Пустая строка тоже считается новой.
        new_mask = df[TAG_NAME].astype(str).str.strip().isin(
            ["", PLACEHOLDER]
        )
        df_new = df[new_mask].copy()
        df_old = df[~new_mask].copy()

        print(f"\nПромаркированных (старых) блоков: {len(df_old)}")
        print(f"Новых (с маркером '{PLACEHOLDER}'): {len(df_new)}")

        if df_new.empty:
            print("\nНовых блоков нет — нечего доназначать.")
            print("Формирую export_final.csv только из старых блоков.")
            # Просто сохраним export_final.csv
            export_cols = ["Handle", "X", "Y", TAG_CABINET, TAG_FLOOR,
                           "_panel", "_port", TAG_NAME, "_device_type", "_display_name"]
            export_cols = [c for c in export_cols if c in df_old.columns]
            df_export = df_old[export_cols].rename(columns={
                TAG_CABINET: "Шкаф", TAG_FLOOR: "Этаж",
                "_panel": "Панель", "_port": "Порт",
                TAG_NAME: "Имя", "_device_type": "Тип",
            })
            df_export.to_csv("export_final.csv", index=False, encoding="utf-8-sig")
            print("Сохранён export_final.csv")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            return

        # --- 6. Парсим пропуски ---
        skip_ports_dict = parse_skip_ports(SKIP_PORTS_LIST)

        # --- 7. Строим карту занятых портов по старым блокам ---
        # occupied_all: {(cabinet, panel, port), ...} — все занятые порты
        # occupied_by_type: {(cabinet, block_type): set((panel, port))}
        occupied_all = set()
        occupied_by_type = {}

        for _, row in df_old.iterrows():
            if row["_panel"] > 0 and row["_port"] > 0:
                cab = str(row[TAG_CABINET])
                key = (cab, row["_panel"], row["_port"])
                occupied_all.add(key)
                type_key = (cab, row["_device_type"])
                occupied_by_type.setdefault(type_key, set()).add(
                    (row["_panel"], row["_port"])
                )

        print(f"Занятых портов (по старым блокам): {len(occupied_all)}")

        # --- 8. Распределение новых блоков ---
        # Готовим справочник конфигураций по имени блока
        cfg_by_name = {b["block_name"]: b for b in enabled_blocks}

        # Сортируем новые блоки: по типу, шкафу, этажу, Y, X
        df_new = df_new.sort_values(
            by=["_device_type", TAG_CABINET, TAG_FLOOR, "Y", "X"],
            ascending=[True, True, True, False, True]
        ).reset_index(drop=True)

        new_rows = []     # сюда будем складывать назначенные (panel, port, name)
        errors = []       # ошибки (не хватило портов)

        # Группируем по (тип блока, шкаф)
        grouped = df_new.groupby(["_device_type", TAG_CABINET])

        for (block_type, cabinet), group in grouped:
            print(f"\n  Обработка: тип '{block_type}', шкаф '{cabinet}' "
                  f"({len(group)} новых блоков)")
            cfg = cfg_by_name.get(block_type)
            if not cfg:
                print(f"    Пропуск: нет конфигурации для '{block_type}'")
                continue

            allowed_ports = parse_port_ranges(cfg["port_ranges"])
            if not allowed_ports:
                print(f"    Пропуск: пустой список портов")
                continue

            # Занятые порты этого типа в этом шкафу
            type_key = (cabinet, block_type)
            type_occupied = occupied_by_type.get(type_key, set())

            # Определяем начальную точку (для строгого режима)
            start_after = None
            if not FILL_GAPS and type_occupied:
                # Максимум (panel, port) среди старых блоков этого типа
                max_panel = max(p for p, _ in type_occupied)
                max_port = max(po for pa, po in type_occupied if pa == max_panel)
                start_after = (max_panel, max_port)
                print(f"    Старт строго после порта: панель {max_panel}, порт {max_port}")

            # Занятые порты, которые надо исключить при поиске:
            # — все старые блоки (не только этого типа), чтобы не наложиться
            #   на чужой диапазон
            occupied_for_search = set()
            for (c, p, po) in occupied_all:
                if c == cabinet:
                    occupied_for_search.add((p, po))

            # Назначаем каждому новому блоку порт
            assigned = 0
            for idx, row in group.iterrows():
                panel, port = find_next_free_port(
                    cabinet, allowed_ports, skip_ports_dict,
                    MAX_PANELS, occupied_for_search,
                    start_after=start_after
                )
                if panel is None:
                    errors.append(
                        f"Шкаф '{cabinet}', тип '{block_type}': "
                        f"не хватило портов для {len(group) - assigned} блоков."
                    )
                    break

                # Фиксируем назначенный порт
                occupied_for_search.add((panel, port))
                type_occupied.add((panel, port))
                occupied_all.add((cabinet, panel, port))

                new_row = row.copy()
                new_row["_panel"] = panel
                new_row["_port"] = port
                new_row[TAG_NAME] = NAME_FORMAT.format(
                    cabinet=cabinet, panel=panel, port=port
                )
                new_row["_assigned"] = True
                new_rows.append(new_row)

                # Для строгого режима обновляем start_after, чтобы
                # следующий блок искал уже после назначенного
                if not FILL_GAPS:
                    start_after = (panel, port)

                assigned += 1

            print(f"    Назначено портов: {assigned}")

        # --- 9. Если были ошибки — выходим ---
        if errors:
            print("\n❌ Не хватило портов:")
            for err in errors:
                print(f"   • {err}")
            print("\nУвеличьте MAX_PANELS_PER_CABINET или проверьте диапазоны.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit(1)

        # --- 10. Формируем итоговый DataFrame (старые + новые) ---
        df_new_assigned = pd.DataFrame(new_rows) if new_rows else pd.DataFrame()
        # Объединяем: старые блоки — как есть, новые — с назначенными
        # panel/port/name.
        df_final = pd.concat([df_old, df_new_assigned], ignore_index=True)
        print(f"\nВсего блоков в финале: {len(df_final)} "
              f"(старых {len(df_old)}, новых {len(df_new_assigned)})")

        # --- 11. Сохраняем export_final.csv ---
        export_cols = ["Handle", "X", "Y", TAG_CABINET, TAG_FLOOR,
                       "_panel", "_port", TAG_NAME, "_device_type", "_display_name"]
        export_cols = [c for c in export_cols if c in df_final.columns]
        df_export = df_final[export_cols].rename(columns={
            TAG_CABINET: "Шкаф", TAG_FLOOR: "Этаж",
            "_panel": "Панель", "_port": "Порт",
            TAG_NAME: "Имя", "_device_type": "Тип",
        })
        df_export.to_csv("export_final.csv", index=False, encoding="utf-8-sig")
        print("Сохранён export_final.csv (старые + новые)")

        # --- 12. Формируем import.csv только для НОВЫХ блоков ---
        if new_rows:
            df_import = df_new_assigned[["Handle", TAG_NAME]].copy()
            df_import.to_csv("import.csv", index=False, encoding="utf-8-sig")
            print(f"Сохранён import.csv (только новые — {len(df_import)} блоков)")
        else:
            print("Новых блоков нет — import.csv не создан.")

        elapsed = time.time() - start_total
        print(f"\nОбщее время выполнения: {elapsed:.2f} сек")

    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")


if __name__ == "__main__":
    main(wait_for_exit=True)
