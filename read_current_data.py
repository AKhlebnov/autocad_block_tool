# ============================================================
#  read_current_data.py – чтение данных из AutoCAD БЕЗ изменений.
#  Версия 4.2
#
#  Что делает:
#    1. Подключается к AutoCAD.
#    2. Читает блоки из BLOCK_CONFIGS (только включённые).
#    3. НЕ распределяет порты, НЕ перемаркировывает.
#    4. Пытается распарсить текущее NAME (например, 8.1A1/01.4)
#       и восстановить из него Панель и Порт.
#    5. Сохраняет всё в export_final.csv — тот же файл, что
#       используется программой в обычном режиме.
#
#  Когда использовать:
#    • Проект уже готов и промаркирован, но нет export_final.csv.
#    • Нельзя перемаркировывать (например, временная схема).
#    • Нужны данные для заполнения выносок, отчётов и т.д.
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
# 1. Парсинг NAME → (panel, port)
# ------------------------------------------------------------

def parse_name_to_panel_port(name):
    """
    Пытается извлечь из имени вида "8.1A1/01.4" номер панели и порта.
    Возвращает (panel_int, port_int) или (None, None), если не получилось.

    Ожидаемый формат: <шкаф>/<панель>.<порт>
    Панель может быть 1-2 цифры, порт — 1-2 цифры.
    """
    if not name or "/" not in name or "." not in name:
        return None, None
    try:
        after_slash = name.split("/", 1)[1]   # "01.4"
        panel_str, port_str = after_slash.split(".", 1)
        return int(panel_str), int(port_str)
    except:
        return None, None


# ------------------------------------------------------------
# 2. Получение блоков через SelectionSet
# ------------------------------------------------------------

def get_blocks_by_name(doc, block_name, max_retries=3):
    """
    Возвращает список вхождений блоков с заданным именем.
    С повторами при COM-ошибке «Вызов был отклонен».
    """
    for attempt in range(max_retries):
        try:
            ss_name = f"ReadData_{block_name}"
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
                print(f"    COM-ошибка при чтении '{block_name}', "
                      f"повтор через {wait:.0f} сек...")
                time.sleep(wait)
                gc.collect()
            else:
                raise e


# ------------------------------------------------------------
# 3. Сбор данных по одному типу блока (только чтение)
# ------------------------------------------------------------

def collect_blocks_data(doc, block_config, tag_cabinet, tag_floor, tag_name):
    """
    Собирает данные по одному типу блока БЕЗ изменений:
    Handle, X, Y, шкаф, этаж, текущее NAME, тип блока, display_name.
    Также пытается распарсить панель/порт из NAME.
    """
    block_name = block_config["block_name"]
    display_name = block_config.get("display_name", block_name)
    data = []

    try:
        target_blocks = get_blocks_by_name(doc, block_name)
    except Exception as e:
        print(f"  Ошибка при чтении блоков '{block_name}': {e}")
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

            # Пытаемся распарсить NAME → панель и порт
            current_name = row.get(tag_name, "")
            panel, port = parse_name_to_panel_port(current_name)
            row["_panel"] = panel if panel is not None else 0
            row["_port"] = port if port is not None else 0

            data.append(row)
        except Exception as e:
            print(f"  Ошибка при обработке блока {block_name}: {e}")
            continue
    return data


# ------------------------------------------------------------
# 4. Основная функция
# ------------------------------------------------------------

def main(wait_for_exit=True):
    """
    Чтение данных из AutoCAD без изменений.
    Результат сохраняется в export_final.csv — тот же файл,
    что используется в обычном режиме.
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
        BLOCK_CONFIGS = config.BLOCK_CONFIGS

        # --- 2. Подключение к AutoCAD ---
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")
        print("РЕЖИМ: только чтение (без перемаркировки)\n")

        # --- 3. Сбор данных по всем включённым блокам ---
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

        # --- 4. DataFrame ---
        df = pd.DataFrame(all_data)
        print(f"\nВсего собрано блоков: {len(df)}")

        # --- 5. Фильтрация (как в основном скрипте) ---
        df = df.dropna(subset=[TAG_CABINET, TAG_FLOOR])
        df = df[df[TAG_CABINET].astype(str).str.strip() != ""]
        df = df[df[TAG_FLOOR].astype(str).str.strip() != ""]
        df[TAG_FLOOR] = pd.to_numeric(df[TAG_FLOOR], errors='coerce')
        df = df.dropna(subset=[TAG_FLOOR])
        print(f"После фильтрации: {len(df)}")

        # --- 6. Статистика по парсингу NAME ---
        parsed_ok = ((df["_panel"] > 0) & (df["_port"] > 0)).sum()
        print(f"Успешно распарсено имён (панель/порт): {parsed_ok} из {len(df)}")

        # --- 7. Сохранение в export_final.csv ---
        export_cols = [
            "Handle", "X", "Y", TAG_CABINET, TAG_FLOOR,
            "_panel", "_port", TAG_NAME, "_device_type", "_display_name"
        ]
        export_cols = [c for c in export_cols if c in df.columns]
        df_export = df[export_cols].copy()
        df_export = df_export.rename(columns={
            TAG_CABINET: "Шкаф",
            TAG_FLOOR: "Этаж",
            "_panel": "Панель",
            "_port": "Порт",
            TAG_NAME: "Имя",
            "_device_type": "Тип",
        })
        df_export.to_csv("export_final.csv", index=False, encoding="utf-8-sig")
        print("Сохранён export_final.csv (только чтение, без изменений в чертеже)")

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
