# ============================================================
#  cable_journal.py – генерация листа "Кабельный журнал"
#  для отчёта по проекту. Версия 4.2
# ============================================================

from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# Светло-голубой фон для шапки — тот же, что и на остальных листах
HEADER_FILL = PatternFill(
    start_color="D9E1F2",
    end_color="D9E1F2",
    fill_type="solid"
)


def add_cable_journal_sheet(writer, df):
    """
    Добавляет в книгу Excel лист "Кабельный журнал".

    Структура шапки (4 строки):
        Строка 1: | Позиция | Обозначение кабеля | Трасса (C:D) | | Тип кабеля | Длина, м | Примечание |
        Строка 2: |         |                    | Обозначение оборудования (C:D) | | | | |
        Строка 3: |         |                    | Начало, Шкаф_Патч-панель_Порт | Конец, Тип оборудования, номер | | | |
        Строка 4: |    1    |         2          |              3               |              4                |    5     |    6    |    7    |

    Данные (с 5-й строки):
        - Позиция:              1, 2, 3, ...
        - Обозначение кабеля:   = имя оборудования
        - Начало:               "8.1A1_1_4"
        - Конец:                "Видеокамера\\n8.1A1/01.4"
        - Тип кабеля:           "UTP 4x2x0.5 cat.5e"
        - Длина, м:             "-"
        - Примечание:           пусто
    """
    wb = writer.book
    ws = wb.create_sheet("Кабельный журнал")

    # --- Стили ---
    header_font = Font(bold=True, size=11)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ============================================================
    # ШАПКА (4 строки)
    # ============================================================
    ws.cell(row=1, column=1, value="Позиция")
    ws.cell(row=1, column=2, value="Обозначение кабеля")
    ws.cell(row=1, column=3, value="Трасса")
    ws.cell(row=1, column=5, value="Тип кабеля")
    ws.cell(row=1, column=6, value="Длина, м")
    ws.cell(row=1, column=7, value="Примечание")

    ws.cell(row=2, column=3, value="Обозначение оборудования")

    ws.cell(row=3, column=3, value="Начало,\nШкаф_Патч-панель_Порт")
    ws.cell(row=3, column=4, value="Конец,\nТип оборудования, номер")

    # Строка 4: нумерация столбцов
    for col in range(1, 8):
        ws.cell(row=4, column=col, value=col)

    # Объединение ячеек шапки
    ws.merge_cells('A1:A4')
    ws.merge_cells('B1:B4')
    ws.merge_cells('C1:D1')
    ws.merge_cells('C2:D2')
    ws.merge_cells('E1:E4')
    ws.merge_cells('F1:F4')
    ws.merge_cells('G1:G4')

    # Стилизация шапки (строки 1-4) — жирный + цвет фона + границы
    for row in range(1, 5):
        for col in range(1, 8):
            cell = ws.cell(row=row, column=col)
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            cell.fill = HEADER_FILL

    # ============================================================
    # ДАННЫЕ (с 5-й строки)
    # ============================================================
    CABLE_TYPE = "UTP 4x2x0.5 cat.5e"

    position = 1
    current_row = 5

    for _, row in df.iterrows():
        cabinet = row["Шкаф"]
        panel = int(row["Панель"])
        port = int(row["Порт"])
        name = row["Имя"]
        display_name = row.get("_display_name", "") or row.get("Тип", "")

        ws.cell(row=current_row, column=1, value=position)
        ws.cell(row=current_row, column=2, value=name)
        ws.cell(row=current_row, column=3, value=f"{cabinet}_{panel}_{port}")
        ws.cell(row=current_row, column=4, value=f"{display_name}\n{name}")
        ws.cell(row=current_row, column=5, value=CABLE_TYPE)
        ws.cell(row=current_row, column=6, value="-")

        for col in range(1, 8):
            cell = ws.cell(row=current_row, column=col)
            cell.border = thin_border
            if col in (1, 5, 6):
                cell.alignment = center_align
            else:
                cell.alignment = left_align

        position += 1
        current_row += 1

    # ============================================================
    # ШИРИНА КОЛОНОК
    # ============================================================
    widths = {
        'A': 10,   # Позиция
        'B': 22,   # Обозначение кабеля
        'C': 26,   # Начало
        'D': 26,   # Конец
        'E': 22,   # Тип кабеля
        'F': 10,   # Длина
        'G': 20,   # Примечание
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    # Высота строк с данными (для двухстрочного «Конца»)
    for row in range(5, current_row):
        ws.row_dimensions[row].height = 30

    return ws
