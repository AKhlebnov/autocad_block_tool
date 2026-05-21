# config.py – автоматически сохранён из GUI

BLOCK_NAME = 'SOCKET_1P'

TAG_CABINET = 'MAC'
TAG_FLOOR = 'IP'
TAG_NAME = 'NAME'
TAG_ICON = 'ICON'
TAG_SERIAL = 'SERIAL_NUMBER'

PORT_RANGES = [[1, 23], [25, 47]]

NAME_FORMAT = '{cabinet}/{panel:02d}.{port}'

TEMP_ATTRIBUTES = ['MAC', 'IP']

SORT_BY = ['MAC', 'IP', 'Y', 'X']
SORT_ASCENDING = [True, True, False, True]

MAX_PANELS_PER_CABINET = 20

SKIP_PORTS = {('АБЧ2.1', 1): [(1, 47)], ('АБЧ2.1', 2): [(1, 47)], ('АБЧ2.1', 3): [(1, 47)], ('АБЧ2.1', 4): [(1, 47)], ('АБЧ2.1', 5): [(1, 47)], ('АБЧ2.1', 6): [(1, 22)]}
