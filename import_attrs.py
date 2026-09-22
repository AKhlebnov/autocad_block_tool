# ============================================================
#  import_attrs.py – импорт имён (NAME) в атрибуты блоков AutoCAD.
#  Версия 3.0
# ============================================================

import sys
import os
import pandas as pd
import win32com.client
import pythoncom
import time


def main(wait_for_exit=True):
    """Основная функция импорта имён в AutoCAD."""

    # --- 1. Динамическая загрузка config ---
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(application_path, 'config.py')
    import importlib.util
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    TAG_NAME = config.TAG_NAME

    # --- 2. Проверка файла import.csv ---
    import_file = "import.csv"
    if not os.path.exists(import_file):
        print(f"Ошибка: файл {import_file} не найден.")
        print("Сначала запустите экспорт и расчёт имён (assign_names.py).")
        if wait_for_exit:
            input("\nНажмите Enter для выхода...")
        sys.exit(1)

    print(f"Читаем файл: {import_file}")

    # --- 3. Чтение данных ---
    df = pd.read_csv(import_file, encoding="utf-8-sig")
    df = df.fillna("")

    print(f"Записей для импорта: {len(df)}")

    # --- 4. Подключение к AutoCAD ---
    pythoncom.CoInitialize()
    start_total = time.time()
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # --- 5. Основной цикл импорта ---
        processed_count = 0    # успешно обработано
        updated_count = 0      # изменили NAME
        not_found_count = 0    # Handle не найден
        error_count = 0        # ошибки

        for _, row in df.iterrows():
            handle = row["Handle"]
            new_value = str(row[TAG_NAME])
            try:
                obj = doc.HandleToObject(handle)

                if obj.ObjectName != "AcDbBlockReference" or not obj.HasAttributes:
                    continue

                atts = obj.GetAttributes()
                if hasattr(atts, "value"):
                    att_list = list(atts.value)
                else:
                    att_list = list(atts)

                changed = False
                for attr in att_list:
                    if attr.TagString == TAG_NAME:
                        if str(attr.TextString) != new_value:
                            attr.TextString = new_value
                            changed = True
                        break

                processed_count += 1
                if changed:
                    updated_count += 1

            except Exception as e:
                if "не найден" in str(e).lower() or "not found" in str(e).lower():
                    not_found_count += 1
                else:
                    print(f"  Ошибка при обновлении блока {handle}: {e}")
                    error_count += 1

        # --- 6. Итоги ---
        print(f"\nОбработано блоков: {processed_count}")
        print(f"Из них изменили NAME: {updated_count}")
        if not_found_count:
            print(f"Не найдено в чертеже: {not_found_count}")
        if error_count:
            print(f"Ошибок при обновлении: {error_count}")

        # --- 7. Регенерация ---
        doc.Regen(0)
        print("Чертёж обновлён (Regen).")

        elapsed = time.time() - start_total
        print(f"\nОбщее время выполнения: {elapsed:.2f} сек")

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
