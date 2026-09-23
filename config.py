# ============================================================
#  config.py – настройки программы нумерации блоков (v4.1)
#  Файл автоматически сохраняется при работе из GUI.
#  Можно редактировать вручную, но при первом же нажатии
#  кнопки в GUI файл перезапишется значениями из интерфейса.
# ============================================================

# --- Общие теги атрибутов (одинаковые для всех блоков) ---
TAG_CABINET = 'MAC'    # атрибут с именем шкафа
TAG_FLOOR   = 'IP'    # атрибут с номером этажа
TAG_NAME    = 'NAME'    # итоговое имя блока
TAG_ICON    = 'ICON'    # на будущее (тип камеры)
TAG_SERIAL  = 'SERIAL_NUMBER'    # на будущее (серийный номер)

# --- Шаблон итогового имени ---
NAME_FORMAT = '{cabinet}/{panel:02d}.{port}'

# --- Максимальное количество патч-панелей в одном шкафу ---
MAX_PANELS_PER_CABINET = 20

# --- Общие пропуски портов ---
# Формат каждой строки: "шкаф,номер_панели,список_портов"
SKIP_PORTS = [
]

# --- Список основных блоков ---
# enabled:      True/False — участвует ли блок в обработке
# block_name:   имя блока в AutoCAD
# display_name: отображаемое имя для кабельного журнала
# port_ranges:  список диапазонов портов
# poe:          True/False — питание PoE (для кроссировочной таблицы)
BLOCK_CONFIGS = [
    {'enabled': True, 'block_name': 'camera', 'display_name': 'Видеокамера', 'port_ranges': [[1, 18], [25, 42]], 'poe': True},
    {'enabled': True, 'block_name': 'AP', 'display_name': 'Точка доступа Wi-Fi', 'port_ranges': [[19, 23], [43, 47]], 'poe': True},
    {'enabled': False, 'block_name': 'Socket_1p', 'display_name': 'Розетка', 'port_ranges': [[1, 18], [25, 42]], 'poe': False},
    {'enabled': False, 'block_name': 'Socket_RJ-45', 'display_name': 'Разъём RJ-45', 'port_ranges': [[1, 18], [25, 42]], 'poe': False},
    {'enabled': False, 'block_name': '', 'display_name': '', 'port_ranges': [], 'poe': False},
]

# ============================================================
#  ОСОБЫЙ РЕЖИМ
# ============================================================
SPECIAL_MODE_ENABLED = True
SPECIAL_BLOCK_NAMES = ['Socket_1p', 'Socket_RJ-45']
SPECIAL_PORT_RANGES = [[1, 23], [25, 47]]
SPECIAL_CONTINUE_LAST_PANEL = True
SPECIAL_SKIP_PORTS = 5
