import sys
import os
import importlib.util
import pandas as pd
import time
import win32com.client
import pythoncom   # для работы с COM в потоках

def main(wait_for_exit=True):
    """
    Импорт ранее рассчитанных имён (NAME) в атрибуты блоков AutoCAD.
    wait_for_exit – ожидание нажатия Enter в консольном режиме.
    """
    # Динамическая загрузка config (свежая из файла)
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(application_path, 'config.py')
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    # Инициализируем COM для этого потока (необходимо при вызове из GUI)
    pythoncom.CoInitialize()
    start_time = time.time()
    try:
        # Подключаемся к AutoCAD
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # Имя файла для импорта формируем на основе BLOCK_NAME
        BLOCK_NAME = config.BLOCK_NAME
        import_file = f'{BLOCK_NAME}_import.csv'

        # Проверка существования файла
        if not os.path.exists(import_file):
            print(f"Ошибка: файл {import_file} не найден. Сначала запустите экспорт.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit(1)

        print(f"Читаем файл: {import_file}")
        # Читаем файл импорта (должен быть создан assign_names.py)
        df = pd.read_csv(import_file, encoding='utf-8-sig')
        df = df.fillna('')

        updated_count = 0
        for _, row in df.iterrows():
            handle = row['Handle']
            try:
                obj = doc.HandleToObject(handle)
                if obj.ObjectName == "AcDbBlockReference" and obj.HasAttributes:
                    atts = obj.GetAttributes()
                    if hasattr(atts, 'value'):
                        att_list = list(atts.value)
                    else:
                        att_list = list(atts)
                    for attr in att_list:
                        tag = attr.TagString
                        if tag in row.index and str(attr.TextString) != str(row[tag]):
                            attr.TextString = str(row[tag])
                    updated_count += 1
            except Exception as e:
                print(f"Не удалось обновить блок {handle}: {e}")

        print(f"Обновлено блоков: {updated_count}")
        doc.Regen(0)
        elapsed = time.time() - start_time
        print(f"Общее время выполнения: {elapsed:.2f} сек")

    except Exception as e:
        print(f"\n❌ Ошибка в процессе импорта: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Освобождаем ресурсы COM
        pythoncom.CoUninitialize()
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")

if __name__ == "__main__":
    main(wait_for_exit=True)
