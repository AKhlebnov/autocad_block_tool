# ============================================================
#  config.py – настройки программы нумерации блоков (v4.2)
#  Файл автоматически сохраняется при работе из GUI.
# ============================================================

# --- Общие теги атрибутов ---
TAG_CABINET = 'MAC'
TAG_FLOOR   = 'IP'
TAG_NAME    = 'NAME'
TAG_ICON    = 'ICON'
TAG_SERIAL  = 'SERIAL_NUMBER'

# --- Шаблон итогового имени ---
NAME_FORMAT = '{cabinet}/{panel:02d}.{port}'

# --- Максимум патч-панелей ---
MAX_PANELS_PER_CABINET = 20

# --- Пропуски портов ---
SKIP_PORTS = [
    '8.1B5,1,25-30',
    '8.1A1,1,25-30',
    '8.1A2,1,25-30',
    '8.1A4,1,25-30',
    '8.1A5,1,25-30',
    '8.1A6,1,25-30',
]

# --- Список блоков ---
BLOCK_CONFIGS = [
    {'enabled': True, 'block_name': 'camera', 'display_name': 'Видеокамера', 'port_ranges': [[1, 18], [25, 42]], 'poe': True},
    {'enabled': True, 'block_name': 'AP', 'display_name': 'Точка доступа Wi-Fi', 'port_ranges': [[19, 23], [43, 47]], 'poe': True},
    {'enabled': False, 'block_name': 'Socket_1p', 'display_name': 'Розетка', 'port_ranges': [[1, 18], [25, 42]], 'poe': False},
    {'enabled': False, 'block_name': 'Socket_RJ-45', 'display_name': 'Разъём RJ-45', 'port_ranges': [[1, 18], [25, 42]], 'poe': False},
    {'enabled': False, 'block_name': '', 'display_name': '', 'port_ranges': [], 'poe': False},
]

# --- Особый режим ---
SPECIAL_MODE_ENABLED = False
SPECIAL_BLOCK_NAMES = ['Socket_1p', 'Socket_RJ-45']
SPECIAL_PORT_RANGES = [[1, 23], [25, 47]]
SPECIAL_CONTINUE_LAST_PANEL = False
SPECIAL_SKIP_PORTS = 0

# --- Доназначение нового оборудования ---
PLACEHOLDER_MARKER = 'Пустой'
FILL_GAPS = False

# --- Выноски стояков ---
STOYAK_FRAMES_LAYER = '_WB_CAB_AREAS'
STOYAK_LINE_FORMAT = "{count} UTP 4x2x0.5 с отм. {floor} этажа"
STOYAK_FIRST_FLOOR_LINE = "{count} UTP 4x2x0.5 с отм. +2.500"
STOYAK_FINAL_LINE = "на отм. 0.000"
STOYAK_REGULAR_LINE = "{count} UTP 4x2x0.5\nна отм. 0.000"
