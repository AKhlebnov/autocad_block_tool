# config.py – автоматически сохранён из GUI

BLOCK_NAME = 'ТД'

TAG_CABINET = 'MAC'
TAG_FLOOR = 'IP'
TAG_NAME = 'NAME'
TAG_ICON = 'ICON'
TAG_SERIAL = 'SERIAL_NUMBER'

PORT_RANGES = [[19, 23], [43, 47]]

NAME_FORMAT = '{cabinet}/{panel:02d}.{port}'

TEMP_ATTRIBUTES = ['MAC', 'IP']

SORT_BY = ['MAC', 'IP', 'Y', 'X']
SORT_ASCENDING = [True, True, False, True]

MAX_PANELS_PER_CABINET = 20

SKIP_PORTS = {}
