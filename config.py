# ============================================================
#  config.py – настройки программы нумерации блоков (v4.0)
#  Файл автоматически сохраняется при работе из GUI.
# ============================================================

TAG_CABINET = 'MAC'
TAG_FLOOR   = 'IP'
TAG_NAME    = 'NAME'
TAG_ICON    = 'ICON'
TAG_SERIAL  = 'SERIAL_NUMBER'

NAME_FORMAT = '{cabinet}/{panel:02d}.{port}'

MAX_PANELS_PER_CABINET = 20

SKIP_PORTS = [
]

BLOCK_CONFIGS = [
    {'enabled': True, 'block_name': 'camera', 'display_name': 'Видеокамера', 'port_ranges': [[1, 18], [25, 42]]},
    {'enabled': True, 'block_name': 'AP', 'display_name': 'Точка доступа Wi-Fi', 'port_ranges': [[19, 23], [43, 47]]},
    {'enabled': False, 'block_name': 'Socket_1p', 'display_name': 'Розетка', 'port_ranges': [[1, 18], [25, 42]]},
    {'enabled': False, 'block_name': 'Socket_RJ-45', 'display_name': 'Разъём RJ-45', 'port_ranges': [[1, 18], [25, 42]]},
    {'enabled': False, 'block_name': '', 'display_name': '', 'port_ranges': []},
]

SPECIAL_MODE_ENABLED = True
SPECIAL_BLOCK_NAMES = ['Socket_1p', 'Socket_RJ-45']
SPECIAL_PORT_RANGES = [[1, 23], [25, 47]]
SPECIAL_CONTINUE_LAST_PANEL = True
SPECIAL_SKIP_PORTS = 3
