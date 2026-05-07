import sys
import os
import importlib.util
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------------------------------------------------
# 1. Загрузка config (как в assign_names.py)
# ------------------------------------------------------------
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

config_path = os.path.join(application_path, 'config.py')
spec = importlib.util.spec_from_file_location("config", config_path)
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)

# ------------------------------------------------------------
# 2. Настройки из config
# ------------------------------------------------------------
TAG_CABINET = config.TAG_CABINET   # например "MAC"
TAG_FLOOR = config.TAG_FLOOR       # например "IP"
TAG_NAME = config.TAG_NAME         # "NAME"
BLOCK_NAME = config.BLOCK_NAME
# Имена промежуточных колонок, которые создаёт assign_names.py
PANEL_COL = '_PANEL'
PORT_COL = '_PORT'

INPUT_CSV = 'camera_export_final.csv'
OUTPUT_EXCEL = f'{BLOCK_NAME}_cable_report.xlsx'

# ------------------------------------------------------------
# 3. Чтение данных
# ------------------------------------------------------------
df = pd.read_csv(INPUT_CSV, encoding='utf-8-sig')

# Проверяем наличие необходимых колонок
required_cols = [TAG_CABINET, TAG_FLOOR, TAG_NAME, PANEL_COL, PORT_COL]
for col in required_cols:
    if col not in df.columns:
        print(f"Ошибка: в файле {INPUT_CSV} нет колонки {col}")
        sys.exit(1)

# Преобразуем типы
df[PANEL_COL] = df[PANEL_COL].astype(int)
df[PORT_COL] = df[PORT_COL].astype(int)

# ------------------------------------------------------------
# 4. Формирование листа "По шкафам" (сводка)
# ------------------------------------------------------------
summary_data = []

for cabinet, group in df.groupby(TAG_CABINET):
    total_devices = len(group)
    panels = sorted(group[PANEL_COL].unique())
    num_panels = len(panels)
    
    panel_details = []
    for panel in panels:
        ports = sorted(group[group[PANEL_COL] == panel][PORT_COL].tolist())
        # Преобразуем список портов в строку с диапазонами
        ranges = []
        start = ports[0]
        end = ports[0]
        for p in ports[1:]:
            if p == end + 1:
                end = p
            else:
                ranges.append(f"{start}-{end}" if start != end else str(start))
                start = end = p
        ranges.append(f"{start}-{end}" if start != end else str(start))
        panel_details.append(f"Панель {panel}: {','.join(ranges)}")
    
    summary_data.append({
        'Шкаф': cabinet,
        'Всего устройств': total_devices,
        'Количество патч-панелей': num_panels,
        'Детали по панелям': '; '.join(panel_details)
    })

df_summary = pd.DataFrame(summary_data)

# ------------------------------------------------------------
# 5. Формирование листа "Детали по панелям"
# ------------------------------------------------------------
detail_rows = []
for cabinet, group in df.groupby(TAG_CABINET):
    for panel, panel_group in group.groupby(PANEL_COL):
        ports = sorted(panel_group[PORT_COL].tolist())
        ranges = []
        start = ports[0]
        end = ports[0]
        for p in ports[1:]:
            if p == end + 1:
                end = p
            else:
                ranges.append(f"{start}-{end}" if start != end else str(start))
                start = end = p
        ranges.append(f"{start}-{end}" if start != end else str(start))
        port_str = ','.join(ranges)
        detail_rows.append({
            'Шкаф': cabinet,
            'Номер панели': panel,
            'Занятые порты': port_str,
            'Количество портов': len(ports)
        })

df_detail = pd.DataFrame(detail_rows)

# ------------------------------------------------------------
# 6. Запись в Excel с форматированием
# ------------------------------------------------------------
with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    df_summary.to_excel(writer, sheet_name='По шкафам', index=False)
    df_detail.to_excel(writer, sheet_name='Детали по панелям', index=False)
    
    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_len + 2, 50)
            ws.column_dimensions[col_letter].width = adjusted_width

print(f"Отчёт сохранён в {OUTPUT_EXCEL}")

input("\n✅ Готово! Нажмите Enter для выхода...")
