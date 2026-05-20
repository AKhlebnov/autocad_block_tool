import sys               # работа с системой: аргументы командной строки, флаг frozen (собран ли в exe)
import os                # пути к файлам и папкам
import importlib.util    # динамическая загрузка модуля по произвольному пути (нужен для config.py рядом с exe)
import pandas as pd      # таблицы (DataFrame) — сортировка, фильтрация, группировка
import pythoncom
import win32com.client   # доступ к COM-объектам Windows, в частности к AutoCAD
from win32com.client import VARIANT  # особый тип данных COM для массивов атрибутов

# ------------------------------------------------------------
# 1. Динамическая загрузка config (чтобы работало в EXE)
# ------------------------------------------------------------
if getattr(sys, 'frozen', False):
    # Программа собрана в exe — ищем config.py рядом с exe-файлом
    application_path = os.path.dirname(sys.executable)
else:
    # Запуск из исходного кода — берём папку текущего скрипта
    application_path = os.path.dirname(os.path.abspath(__file__))

config_path = os.path.join(application_path, 'config.py')
spec = importlib.util.spec_from_file_location("config", config_path)
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)   # теперь config доступен как модуль

# ------------------------------------------------------------
# 2. Вспомогательная функция для получения доступных портов
# ------------------------------------------------------------
def get_available_ports(cabinet, panel_number, allowed_ports):
    """
    Возвращает список портов для указанного шкафа и панели,
    исключая порты, заданные в SKIP_PORTS.
    Поддерживает одиночные порты и диапазоны (кортежи/списки из двух чисел).
    """
    skip_list_raw = config.SKIP_PORTS.get((cabinet, panel_number), [])
    skip_ports = set()
    for item in skip_list_raw:
        if isinstance(item, (tuple, list)) and len(item) == 2:
            # Диапазон (начало, конец) включительно
            start, end = item
            skip_ports.update(range(start, end + 1))
        else:
            # Одиночный порт
            skip_ports.add(item)
    # Возвращаем порты из allowed_ports, которых нет в skip_ports
    return [p for p in allowed_ports if p not in skip_ports]

def main(wait_for_exit=True):
    # Инициализируем COM для этого потока
    pythoncom.CoInitialize()
    try:
        # ------------------------------------------------------------
        # 3. Подключение к AutoCAD
        # ------------------------------------------------------------
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # ------------------------------------------------------------
        # 4. Настройки из config
        # ------------------------------------------------------------
        BLOCK_NAME = config.BLOCK_NAME
        TAG_CABINET = config.TAG_CABINET   # атрибут, где временно хранится имя шкафа
        TAG_FLOOR = config.TAG_FLOOR       # атрибут, где временно хранится номер этажа
        TAG_NAME = config.TAG_NAME         # итоговый атрибут с именем камеры
        PORT_RANGES = config.PORT_RANGES   # допустимые диапазоны портов на одной панели
        NAME_FORMAT = config.NAME_FORMAT   # шаблон имени, например "{cabinet}/{panel:02d}.{port}"
        SORT_BY = config.SORT_BY           # колонки для сортировки
        SORT_ASCENDING = config.SORT_ASCENDING   # порядок сортировки (True/False)

        # Формируем плоский список всех допустимых портов для одной панели
        ALLOWED_PORTS = []
        for r in PORT_RANGES:
            ALLOWED_PORTS.extend(r)

        # ------------------------------------------------------------
        # 5. Сбор данных из чертежа
        # ------------------------------------------------------------
        data = []
        blocks_found = 0

        for obj in doc.ModelSpace:
            if obj.ObjectName == "AcDbBlockReference" and obj.EffectiveName == BLOCK_NAME:
                blocks_found += 1
                handle = obj.Handle

                if not obj.HasAttributes:
                    print(f"Блок {handle} не имеет атрибутов, пропускаем")
                    continue

                # Получаем координаты вставки (нужны для сортировки)
                ins = obj.InsertionPoint
                x, y, z = ins[0], ins[1], ins[2]

                # Получаем атрибуты (преобразуем VARIANT в обычный список, если нужно)
                try:
                    atts = obj.GetAttributes()
                    if isinstance(atts, VARIANT):
                        att_list = list(atts.value)
                    else:
                        att_list = list(atts)
                except Exception as e:
                    print(f"Ошибка получения атрибутов блока {handle}: {e}")
                    continue

                # Словарь для одной строки таблицы (пока с пустыми значениями)
                row = {
                    'Handle': handle,
                    'X': x,
                    'Y': y,
                    'Z': z,
                    TAG_CABINET: "",
                    TAG_FLOOR: "",
                    TAG_NAME: "",
                }
                # Заполняем значения атрибутов, если тег совпадает с ключами выше
                for attr in att_list:
                    tag = attr.TagString
                    if tag in row:
                        row[tag] = attr.TextString

                data.append(row)

        print(f'Обработано блоков - {blocks_found} шт.')

        if not data:
            print("Не найдено ни одного блока с атрибутами. Завершение.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit()

        df = pd.DataFrame(data)
        print(f"Экспортировано блоков: {len(df)}")
        df.to_csv('camera_export_raw.csv', index=False, encoding='utf-8-sig')

        # ------------------------------------------------------------
        # 6. Очистка и фильтрация
        # ------------------------------------------------------------
        # Удаляем строки, где не указан шкаф или этаж
        df = df.dropna(subset=[TAG_CABINET, TAG_FLOOR])
        # Преобразуем этаж в число (если не число -> NaN)
        df[TAG_FLOOR] = pd.to_numeric(df[TAG_FLOOR], errors='coerce')
        df = df.dropna(subset=[TAG_FLOOR])
        df[TAG_FLOOR] = df[TAG_FLOOR].astype(int)
        print(f"После фильтрации (есть шкаф и этаж) осталось блоков: {len(df)}")

        # ------------------------------------------------------------
        # 7. Сортировка
        # ------------------------------------------------------------
        df_sorted = df.sort_values(by=SORT_BY, ascending=SORT_ASCENDING).reset_index(drop=True)

        # ------------------------------------------------------------
        # 8. Назначение панелей, портов и формирование NAME
        # ------------------------------------------------------------
        df_sorted['_PANEL'] = 0   # временная колонка для номера панели
        df_sorted['_PORT'] = 0    # временная колонка для номера порта

        MAX_PANELS = config.MAX_PANELS_PER_CABINET   # максимальное число панелей в шкафу (из конфига)

        def get_skipped_ports_string(allowed_ports, available_ports):
            """
            Возвращает строку с перечислением пропущенных портов (диапазонами).
            Например: " (пропущены 25-28,30)"
            """
            skipped = sorted(set(allowed_ports) - set(available_ports))
            if not skipped:
                return ""
            # Преобразуем список в строку с диапазонами
            ranges = []
            start = skipped[0]
            end = skipped[0]
            for p in skipped[1:]:
                if p == end + 1:
                    end = p
                else:
                    ranges.append(f"{start}-{end}" if start != end else str(start))
                    start = end = p
            ranges.append(f"{start}-{end}" if start != end else str(start))
            return f" (пропущены {','.join(ranges)})"

        def check_cabinet_capacity(cabinet, num_cameras, max_panels, allowed_ports):
            """
            Проверяет, хватит ли портов в шкафу с учётом пропусков и максимального числа панелей.
            Возвращает (enough, total_ports, details)
                enough - bool (хватает ли портов)
                total_ports - общее количество доступных портов
                details - список строк с информацией о каждой панели
            """
            total_ports = 0
            details = []
            for panel in range(1, max_panels + 1):
                available = get_available_ports(cabinet, panel, allowed_ports)
                num_ports = len(available)
                total_ports += num_ports
                skip_str = get_skipped_ports_string(allowed_ports, available)
                details.append(f"  Панель {panel}: {num_ports} портов{skip_str}")
                # Если уже набрали достаточно портов, дальше можно не проверять (экономия времени)
                if total_ports >= num_cameras:
                    break
            enough = total_ports >= num_cameras
            return enough, total_ports, details

        # Предварительная проверка для каждого шкафа (до основного цикла)
        for cabinet, group in df_sorted.groupby(TAG_CABINET):
            num_cameras = len(group)
            enough, total_ports, details = check_cabinet_capacity(cabinet, num_cameras, MAX_PANELS, ALLOWED_PORTS)
            if not enough:
                print(f"\n❌ Шкаф {cabinet}:")
                print(f"   Камер: {num_cameras}, доступно портов: {total_ports} (при макс. {MAX_PANELS} панелях).")
                print("   Детали по панелям:")
                for d in details:
                    print(d)
                print("\n   Решения:")
                print("   - Увеличьте MAX_PANELS_PER_CABINET в config.py, если в шкафу реально больше панелей.")
                print(f"   - Перенесите {num_cameras - total_ports} камер в другой шкаф.")
                print("   - Проверьте пропуски портов (SKIP_PORTS).\n")
                if wait_for_exit:
                    input("\nНажмите Enter для выхода...")
                sys.exit(1)

        # Основной цикл назначения панелей и портов (теперь мы уверены, что портов хватит)
        for cabinet, group in df_sorted.groupby(TAG_CABINET):
            print(f"Обрабатывается шкаф: {cabinet}, камер: {len(group)}")
            panel_number = 1
            # Поиск первой панели с портами
            while True:
                available_ports = get_available_ports(cabinet, panel_number, ALLOWED_PORTS)
                if available_ports:
                    break
                panel_number += 1
            port_index = 0
            idx_list = group.index.tolist()
            total_cameras = len(idx_list)
            cameras_done = 0

            for idx in idx_list:
                if port_index >= len(available_ports):
                    panel_number += 1
                    port_index = 0
                    # Ищем следующую панель с портами
                    while True:
                        available_ports = get_available_ports(cabinet, panel_number, ALLOWED_PORTS)
                        if available_ports:
                            break
                        panel_number += 1

                port = available_ports[port_index]
                df_sorted.at[idx, '_PANEL'] = panel_number
                df_sorted.at[idx, '_PORT'] = port
                df_sorted.at[idx, TAG_NAME] = NAME_FORMAT.format(
                    cabinet=cabinet,
                    panel=panel_number,
                    port=port
                )
                port_index += 1
                cameras_done += 1

            print(f"  Последняя использованная панель: {panel_number}, всего камер: {total_cameras}")

        # ------------------------------------------------------------
        # 9. Сохранение файла для импорта (только Handle и NAME)
        # ------------------------------------------------------------
        df_import = df_sorted[['Handle', TAG_NAME]].copy()
        df_import.to_csv('camera_import.csv', index=False, encoding='utf-8-sig')
        print("Файл для импорта сохранён как camera_import.csv")

        # Для отладки – полный результат (со всеми временными колонками и координатами)
        df_sorted.to_csv('camera_export_final.csv', index=False, encoding='utf-8-sig')

    except Exception as e:
        # --- Обработка ошибок ---
        print(f"\n❌ Ошибка в основном коде: {e}")
        # Можно добавить traceback для отладки
        import traceback
        traceback.print_exc()
    finally:
        # Освобождаем ресурсы COM
        pythoncom.CoUninitialize()
        # Ожидание нажатия Enter для закрытия окна, если wait_for_exit=True
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")

if __name__ == "__main__":
    main(wait_for_exit=True)
