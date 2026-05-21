import sys
import os
import glob
import time
import importlib.util
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

def main(wait_for_exit=True):
    start_time = time.time()
    # ------------------------------------------------------------
    # 0. Загружаем config, чтобы узнать имена колонок для шкафа, этажа, имени
    # ------------------------------------------------------------
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(application_path, 'config.py')
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    TAG_CABINET = config.TAG_CABINET   # например "MAC"
    TAG_FLOOR = config.TAG_FLOOR       # например "IP"
    TAG_NAME = config.TAG_NAME         # "NAME"

    # ------------------------------------------------------------
    # 1. Находим все файлы *_export_final.csv в текущей папке
    # ------------------------------------------------------------
    csv_files = glob.glob("*_export_final.csv")
    if not csv_files:
        print("Ошибка: не найдено ни одного файла *_export_final.csv. Сначала выполните экспорт.")
        if wait_for_exit:
            input("\nНажмите Enter для выхода...")
        sys.exit(1)

    # ------------------------------------------------------------
    # 2. Объединяем данные из всех файлов, добавляя колонку "Тип оборудования"
    # ------------------------------------------------------------
    all_data = []
    for file in csv_files:
        # Имя типа берём как часть до "_export_final.csv"
        device_type = file.replace("_export_final.csv", "")
        df = pd.read_csv(file, encoding='utf-8-sig')
        # Переименовываем колонки TAG_CABINET, TAG_FLOOR, TAG_NAME в стандартные
        df.rename(columns={TAG_CABINET: 'CABINET', TAG_FLOOR: 'FLOOR', TAG_NAME: 'NAME'}, inplace=True)
        # Добавляем колонку с типом оборудования
        df['Тип оборудования'] = device_type
        all_data.append(df)

    df_combined = pd.concat(all_data, ignore_index=True)

    # ------------------------------------------------------------
    # 3. Проверяем наличие необходимых колонок
    # ------------------------------------------------------------
    required_cols = ['CABINET', 'FLOOR', 'NAME', '_PANEL', '_PORT', 'Тип оборудования']
    for col in required_cols:
        if col not in df_combined.columns:
            print(f"Ошибка: в объединённых данных нет колонки {col}")
            print("Доступные колонки:", list(df_combined.columns))
            sys.exit(1)

    # ------------------------------------------------------------
    # 4. Сводка по шкафам (объединённая)
    # ------------------------------------------------------------
    summary_data = []
    for cabinet, group in df_combined.groupby('CABINET'):
        total_devices = len(group)
        panels = sorted(group['_PANEL'].unique())
        num_panels = len(panels)
        panel_details = []
        for panel in panels:
            panel_group = group[group['_PANEL'] == panel]
            ports = sorted(panel_group['_PORT'].tolist())
            # Преобразуем в диапазоны
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
            types = sorted(panel_group['Тип оборудования'].unique())
            panel_details.append(f"Панель {panel}: {','.join(ranges)} (типы: {','.join(types)})")
        summary_data.append({
            'Шкаф': cabinet,
            'Всего устройств': total_devices,
            'Количество патч-панелей': num_panels,
            'Детали по панелям': '; '.join(panel_details)
        })

    df_summary = pd.DataFrame(summary_data)

    # ------------------------------------------------------------
    # 5. Детали по панелям (с типом оборудования)
    # ------------------------------------------------------------
    detail_rows = []
    for cabinet, group in df_combined.groupby('CABINET'):
        for panel, panel_group in group.groupby('_PANEL'):
            ports = sorted(panel_group['_PORT'].tolist())
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
            types = ', '.join(panel_group['Тип оборудования'].unique())
            detail_rows.append({
                'Шкаф': cabinet,
                'Номер панели': panel,
                'Занятые порты': port_str,
                'Количество портов': len(ports),
                'Тип оборудования': types
            })
    df_detail = pd.DataFrame(detail_rows)

    # ------------------------------------------------------------
    # 6. Попортовый список для всех панелей всех шкафов
    # ------------------------------------------------------------
    port_rows = []
    max_port = 48  # стандартная патч-панель 48 портов
    for (cabinet, panel), panel_group in df_combined.groupby(['CABINET', '_PANEL']):
        taken = dict(zip(panel_group['_PORT'], panel_group['NAME']))
        # Для каждого порта от 1 до max_port
        for port in range(1, max_port+1):
            if port in taken:
                status = taken[port]
                dev_type = panel_group[panel_group['_PORT'] == port]['Тип оборудования'].values[0]
            else:
                status = "Свободный порт"
                dev_type = ""
            port_rows.append({
                'Шкаф': cabinet,
                'Панель': panel,
                'Порт': port,
                'Оборудование': status,
                'Тип оборудования': dev_type
            })
    df_ports = pd.DataFrame(port_rows)

    # ------------------------------------------------------------
    # 7. Запись в Excel
    # ------------------------------------------------------------
    output_excel = "Комбинированный_отчёт.xlsx"
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Сводка по шкафам', index=False)
            df_detail.to_excel(writer, sheet_name='Детали по панелям', index=False)
            df_ports.to_excel(writer, sheet_name='Порты', index=False)

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

        print(f"Отчёт сохранён в {output_excel}")
    except PermissionError:
        print(f"\n❌ Ошибка: не удалось сохранить файл {output_excel}. Возможно, он открыт в другой программе. Закройте файл и повторите.")
    except Exception as e:
        print(f"\n❌ Ошибка при создании отчёта: {e}")
        import traceback
        traceback.print_exc()

    elapsed = time.time() - start_time
    print(f"Общее время выполнения: {elapsed:.2f} сек")

    if wait_for_exit:
        input("\n✅ Готово! Нажмите Enter для выхода...")

if __name__ == "__main__":
    main(wait_for_exit=True)
