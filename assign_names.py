# Работа с системой: аргументы командной строки, флаг frozen (собран ли в exe)
import sys
# Пути к файлам и папкам
import os
# Динамическая загрузка модуля по произвольному пути (нужен для config.py рядом с exe)
import importlib.util
# Библиотека для работы с таблицами (DataFrame) — сортировка, фильтрация, группировка
import pandas as pd
# Для инициализации COM в потоках (нужно при вызове из GUI)
import pythoncom
# Для замеров времени выполнения
import time
# Доступ к COM-объектам Windows, в частности к AutoCAD
import win32com.client
# Особый тип данных COM для массивов атрибутов
from win32com.client import VARIANT


# ------------------------------------------------------------
# Вспомогательная функция для получения доступных портов
# ------------------------------------------------------------
def get_available_ports(cabinet, panel_number, allowed_ports, skip_ports_dict):
    """
    Возвращает список портов для указанного шкафа и панели,
    исключая порты, заданные в SKIP_PORTS.
    Поддерживает одиночные порты и диапазоны (кортежи/списки из двух чисел).

    Параметры:
        cabinet (str) – имя шкафа
        panel_number (int) – номер патч-панели
        allowed_ports (list[int]) – список всех разрешённых портов для одной панели
        skip_ports_dict (dict) – словарь пропусков (config.SKIP_PORTS)
    """
    # Получаем список пропусков для данной пары (шкаф, панель) или пустой список
    skip_list_raw = skip_ports_dict.get((cabinet, panel_number), [])
    # Множество для хранения номеров портов, которые нужно пропустить
    skip_ports = set()
    # Обрабатываем каждый элемент из списка пропусков
    for item in skip_list_raw:
        # Если элемент — кортеж или список из двух чисел, значит это диапазон (начало, конец)
        if isinstance(item, (tuple, list)) and len(item) == 2:
            start, end = item
            # Добавляем в множество все числа от start до end включительно
            skip_ports.update(range(start, end + 1))
        else:
            # Иначе это одиночный порт – просто добавляем его
            skip_ports.add(item)
    # Возвращаем список портов из allowed_ports, которых нет в skip_ports
    return [p for p in allowed_ports if p not in skip_ports]


# ------------------------------------------------------------
# Функция быстрого получения блоков через SelectionSet
# ------------------------------------------------------------
def get_blocks_by_name(doc, block_name):
    """
    Использует SelectionSet с фильтром, чтобы получить все вхождения блоков
    с заданным именем. Возвращает обычный список Python.
    """
    # Константа: acSelectionSetAll = 5 (выбрать все объекты в чертеже)
    acSelectionSetAll = 5
    # Уникальное имя временной выборки
    ss_name = "TempBlockSelect"
    try:
        # Пытаемся создать новую выборку
        ss = doc.SelectionSets.Add(ss_name)
    except:
        # Если выборка уже существует, используем её и очищаем
        ss = doc.SelectionSets.Item(ss_name)
        ss.Clear()
    # Создаём фильтр:
    # filter_type: массив целых чисел (VT_ARRAY | VT_I2) – коды DXF (0 и 2)
    # filter_data: массив строк (VT_ARRAY | VT_VARIANT) – "INSERT" (тип блок) и имя блока
    filter_type = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 2])
    filter_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT", block_name])
    # Выполняем выборку по всему чертежу
    ss.Select(acSelectionSetAll, None, None, filter_type, filter_data)
    # Преобразуем COM-коллекцию в обычный список Python
    blocks = [blk for blk in ss]
    # Удаляем временную выборку, чтобы не засорять чертёж
    ss.Delete()
    return blocks


# ------------------------------------------------------------
# Основная функция
# ------------------------------------------------------------
def main(wait_for_exit=True):
    """
    Главная функция: загружает config, подключается к AutoCAD,
    собирает блоки, сортирует, назначает номера панелей и портов,
    формирует NAME и сохраняет файлы.
    wait_for_exit – если True, в конце будет ожидание нажатия Enter (для консольного режима).
    """
    # ----- 1. Динамическая загрузка config (свежая из файла) -----
    # Это гарантирует, что даже при работе из exe настройки будут актуальны.
    if getattr(sys, 'frozen', False):
        # Если программа собрана в exe, ищем config.py рядом с exe-файлом
        application_path = os.path.dirname(sys.executable)
    else:
        # Если запущена как скрипт, берём папку текущего скрипта
        application_path = os.path.dirname(os.path.abspath(__file__))
    # Полный путь к файлу config.py
    config_path = os.path.join(application_path, 'config.py')
    # Загружаем config как модуль Python
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    # Инициализируем COM для этого потока (необходимо при вызове из GUI)
    pythoncom.CoInitialize()
    # Засекаем время начала работы
    start_total = time.time()

    try:
        # ----- 2. Подключение к AutoCAD -----
        # Получаем активный экземпляр AutoCAD
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        # Получаем активный документ (чертёж)
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # ----- 3. Получение настроек из config -----
        # Имя блока (например, "camera", "WiFi", "SOCKET_1P")
        BLOCK_NAME = config.BLOCK_NAME
        # Префикс для имён выходных файлов (совпадает с именем блока)
        FILE_PREFIX = BLOCK_NAME
        # Атрибут, где хранится имя шкафа (обычно MAC)
        TAG_CABINET = config.TAG_CABINET
        # Атрибут, где хранится номер этажа (обычно IP)
        TAG_FLOOR = config.TAG_FLOOR
        # Итоговый атрибут имени устройства
        TAG_NAME = config.TAG_NAME
        # Разрешённые диапазоны портов на одной панели (например, [[1,19],[25,42]])
        PORT_RANGES = config.PORT_RANGES
        # Шаблон имени, например "{cabinet}/{panel:02d}.{port}"
        NAME_FORMAT = config.NAME_FORMAT
        # Колонки для сортировки (например, [TAG_CABINET, TAG_FLOOR, 'Y', 'X'])
        SORT_BY = config.SORT_BY
        # Порядок сортировки (True – по возрастанию, False – по убыванию)
        SORT_ASCENDING = config.SORT_ASCENDING
        # Словарь пропусков портов (например, {('3.1C6',1):[(25,29)]})
        SKIP_PORTS_DICT = config.SKIP_PORTS

        # Выводим информацию о загруженных пропусках (для контроля)
        print("Загруженные пропуски портов (SKIP_PORTS_DICT):")
        if SKIP_PORTS_DICT:
            for key, value in SKIP_PORTS_DICT.items():
                print(f"  {key} -> {value}")
        else:
            print("Пропуски портов отсутствуют.")

        # ----- 4. Преобразование PORT_RANGES в плоский список ALLOWED_PORTS -----
        # ALLOWED_PORTS – все числа портов, которые могут быть использованы на одной панели.
        # Например, [[1,19],[25,42]] превратится в [1,2,...,19,25,...,42].
        ALLOWED_PORTS = []
        for r in PORT_RANGES:
            # Если элемент — список или кортеж из двух чисел
            if isinstance(r, (list, tuple)) and len(r) == 2:
                if r[0] == r[1]:
                    # Одиночный порт, представленный как [19,19]
                    ALLOWED_PORTS.append(r[0])
                else:
                    # Диапазон – добавляем все числа от r[0] до r[1]
                    ALLOWED_PORTS.extend(range(r[0], r[1] + 1))
            elif isinstance(r, int):
                # Уже число – просто добавляем
                ALLOWED_PORTS.append(r)
            else:
                # На случай, если r — что-то другое (попробуем преобразовать)
                try:
                    ALLOWED_PORTS.append(int(r))
                except:
                    pass

        # ------------------------------------------------------------
        # 5. Сбор данных из чертежа
        # ------------------------------------------------------------
        # Получаем все блоки нужного имени одним быстрым запросом
        target_blocks = get_blocks_by_name(doc, BLOCK_NAME)
        blocks_found = len(target_blocks)
        print(f"Найдено целевых блоков '{BLOCK_NAME}': {blocks_found}")

        # Список для хранения данных о каждом успешно обработанном блоке
        data = []

        # Единый блок try – любая COM-ошибка (включая переключение чертежа) прерывает выполнение.
        # Это самый надёжный способ: если что-то пошло не так, мы просто завершаем программу
        # с понятным сообщением, без сложных повторных попыток.
        try:
            # Перебираем каждый блок из выборки
            for obj in target_blocks:
                # Получаем уникальный Handle блока
                handle = obj.Handle
                # Проверяем, есть ли у блока атрибуты (если нет – пропускаем)
                if not obj.HasAttributes:
                    continue
                # Координаты вставки (X, Y, Z) – нужны для сортировки
                ins = obj.InsertionPoint
                x, y, z = ins[0], ins[1], ins[2]
                # Получаем коллекцию атрибутов блока
                atts = obj.GetAttributes()
                # Преобразуем VARIANT в обычный список Python
                if isinstance(atts, VARIANT):
                    att_list = list(atts.value)
                else:
                    att_list = list(atts)
                # Формируем словарь для строки таблицы (пока пустые значения)
                row = {
                    'Handle': handle,
                    'X': x,
                    'Y': y,
                    'Z': z,
                    TAG_CABINET: "",
                    TAG_FLOOR: "",
                    TAG_NAME: "",
                }
                # Заполняем значения атрибутов, если тег совпадает с ключами row
                for attr in att_list:
                    tag = attr.TagString
                    if tag in row:
                        row[tag] = attr.TextString
                # Добавляем готовую строку в список
                data.append(row)
        except Exception as e:
            # Если произошла любая ошибка (например, переключение чертежа) – выводим сообщение и выходим
            print(f"\n❌ Ошибка при обработке блоков: {e}")
            print("Убедитесь, что во время работы скрипта вы не переключаете чертежи.")
            print("Запустите скрипт заново, оставаясь в исходном чертеже.")
            sys.exit(1)

        # Выводим статистику
        print(f"Успешно обработано блоков с атрибутами: {len(data)}")
        if not data:
            print("Не найдено ни одного блока с атрибутами. Завершение.")
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            sys.exit()

        # Преобразуем список словарей в pandas DataFrame (таблицу)
        df = pd.DataFrame(data)
        print(f"Экспортировано блоков в DataFrame: {len(df)}")

        # ------------------------------------------------------------
        # 6. Очистка и фильтрация
        # ------------------------------------------------------------
        # Удаляем строки, где не указан шкаф или этаж (такие блоки нельзя распределить)
        df = df.dropna(subset=[TAG_CABINET, TAG_FLOOR])
        # Преобразуем этаж в число (если не число – становится NaN)
        df[TAG_FLOOR] = pd.to_numeric(df[TAG_FLOOR], errors='coerce')
        # Удаляем строки, где этаж стал NaN (не число)
        df = df.dropna(subset=[TAG_FLOOR])
        # НЕ приводим к int, чтобы можно было использовать дробные этажи (например, 1.5)
        print(f"После фильтрации (есть шкаф и этаж) осталось блоков: {len(df)}")

        # ------------------------------------------------------------
        # 7. Сортировка
        # ------------------------------------------------------------
        # Сортируем по заданным колонкам и порядку, сбрасываем индексы
        df_sorted = df.sort_values(by=SORT_BY, ascending=SORT_ASCENDING).reset_index(drop=True)

        # ------------------------------------------------------------
        # 8. Назначение панелей, портов и формирование NAME
        # ------------------------------------------------------------
        # Временные колонки для номера панели и номера порта
        df_sorted['_PANEL'] = 0
        df_sorted['_PORT'] = 0
        # Максимальное количество патч-панелей в одном шкафу (из config)
        MAX_PANELS = config.MAX_PANELS_PER_CABINET

        # Вспомогательная функция для красивого вывода пропущенных портов (диапазонами)
        def get_skipped_ports_string(allowed_ports, available_ports):
            """
            Возвращает строку с перечислением пропущенных портов (диапазонами).
            Например: " (пропущены 25-28,30)"
            """
            # Находим порты, которые есть в allowed_ports, но отсутствуют в available_ports
            skipped = sorted(set(allowed_ports) - set(available_ports))
            if not skipped:
                return ""
            # Преобразуем список в строку с диапазонами (например, [1,2,3,5] -> "1-3,5")
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

        # Функция проверки ёмкости шкафа (достаточно ли портов для всех камер)
        def check_cabinet_capacity(cabinet, num_cameras, max_panels, allowed_ports, skip_ports_dict):
            """
            Проверяет, хватит ли портов в шкафу с учётом пропусков и максимального числа панелей.
            Возвращает:
                enough (bool) – хватает ли портов
                total_ports (int) – общее количество доступных портов
                details (list) – список строк с информацией о каждой панели
            """
            total_ports = 0
            details = []
            # Перебираем все панели от 1 до max_panels
            for panel in range(1, max_panels + 1):
                available = get_available_ports(cabinet, panel, allowed_ports, skip_ports_dict)
                num_ports = len(available)
                total_ports += num_ports
                skip_str = get_skipped_ports_string(allowed_ports, available)
                details.append(f"  Панель {panel}: {num_ports} портов{skip_str}")
                # Как только набрали достаточно портов для всех камер – дальше можно не проверять
                if total_ports >= num_cameras:
                    break
            enough = total_ports >= num_cameras
            return enough, total_ports, details

        # Предварительная проверка для каждого шкафа (до основного цикла распределения)
        for cabinet, group in df_sorted.groupby(TAG_CABINET):
            num_cameras = len(group)
            enough, total_ports, details = check_cabinet_capacity(
                cabinet, num_cameras, MAX_PANELS, ALLOWED_PORTS, SKIP_PORTS_DICT
            )
            if not enough:
                # Если портов не хватает – выводим подробную информацию и завершаем программу
                print(f"\n❌ Шкаф {cabinet}:")
                print(f"   Устройств: {num_cameras}, доступно портов: {total_ports} (при макс. {MAX_PANELS} панелях).")
                print("   Детали по панелям:")
                for d in details:
                    print(d)
                print("\n   Решения:")
                print("   - Увеличьте MAX_PANELS_PER_CABINET в config.py, если в шкафу реально больше панелей.")
                print(f"   - Перенесите {num_cameras - total_ports} устройств в другой шкаф.")
                print("   - Проверьте пропуски портов (SKIP_PORTS).\n")
                if wait_for_exit:
                    input("\nНажмите Enter для выхода...")
                sys.exit(1)

        # Основной цикл назначения панелей и портов (теперь мы уверены, что портов хватит)
        for cabinet, group in df_sorted.groupby(TAG_CABINET):
            print(f"Обрабатывается шкаф: {cabinet}, устройств: {len(group)}")
            panel_number = 1  # начинаем с первой патч-панели
            # Ищем первую панель, у которой есть хотя бы один доступный порт
            while True:
                available_ports = get_available_ports(cabinet, panel_number, ALLOWED_PORTS, SKIP_PORTS_DICT)
                if available_ports:
                    break
                panel_number += 1
            port_index = 0
            # Индексы строк этой группы в общем DataFrame
            idx_list = group.index.tolist()
            total_cameras = len(idx_list)
            cameras_done = 0

            for idx in idx_list:
                # Если в текущей панели закончились порты – переходим на следующую
                if port_index >= len(available_ports):
                    panel_number += 1
                    port_index = 0
                    # Ищем следующую панель с доступными портами
                    while True:
                        available_ports = get_available_ports(cabinet, panel_number, ALLOWED_PORTS, SKIP_PORTS_DICT)
                        if available_ports:
                            break
                        panel_number += 1

                # Берём порт из текущего списка доступных портов
                port = available_ports[port_index]
                # Сохраняем номер панели и порт во временные колонки
                df_sorted.at[idx, '_PANEL'] = panel_number
                df_sorted.at[idx, '_PORT'] = port
                # Формируем итоговое имя по шаблону (например, "3.1C6/01.25")
                df_sorted.at[idx, TAG_NAME] = NAME_FORMAT.format(
                    cabinet=cabinet,
                    panel=panel_number,
                    port=port
                )
                port_index += 1
                cameras_done += 1

            print(f"  Последняя использованная панель: {panel_number}, всего устройств: {total_cameras}")

        # ------------------------------------------------------------
        # 9. Сохранение файлов
        # ------------------------------------------------------------
        # Файл для импорта в AutoCAD (только Handle и NAME)
        df_import = df_sorted[['Handle', TAG_NAME]].copy()
        df_import.to_csv(f'{FILE_PREFIX}_import.csv', index=False, encoding='utf-8-sig')
        print(f"Файл для импорта сохранён как {FILE_PREFIX}_import.csv")

        # Полный файл со всеми данными (для отладки и для отчёта)
        df_sorted.to_csv(f'{FILE_PREFIX}_export_final.csv', index=False, encoding='utf-8-sig')

        # Выводим общее время выполнения
        elapsed = time.time() - start_total
        print(f"Общее время выполнения: {elapsed:.2f} сек")

    except Exception as e:
        # Обработка любых ошибок, возникших на уровне выше (например, при подключении к AutoCAD)
        print(f"\n❌ Ошибка в основном коде: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Освобождаем ресурсы COM (обязательно, даже при ошибке)
        pythoncom.CoUninitialize()
        # Если скрипт запущен в консольном режиме – ждём нажатия Enter
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")


if __name__ == "__main__":
    main(wait_for_exit=True)
