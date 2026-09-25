# ============================================================
#  fill_stoyak_notes.py – заполнение выносок стояков мезонина.
#  Версия 4.2
#
#  Логика:
#    1. Читает export_final.csv (данные по всем блокам).
#    2. Собирает рамки (LWPOLYLINE) на слое STOYAK_FRAMES_LAYER.
#    3. Для каждой мультивыноски берёт точку стрелки, находит
#       рамку, внутри которой она лежит.
#    4. Внутри рамки определяет (шкаф, этаж) по большинству
#       блоков.
#    5. Обновляет текст выносок:
#         • этажи 2+ → «N UTP 4x2x0.5 / на отм. 0.000»
#         • этаж 1  → многострочный сводный текст по всем этажам.
# ============================================================

import sys
import os
import importlib.util
import time
import gc
import pandas as pd
import pythoncom
import win32com.client
from win32com.client import VARIANT


# ------------------------------------------------------------
# 1. Вспомогательные функции
# ------------------------------------------------------------

def var_to_list(v):
    """Распаковка VARIANT / SAFEARRAY в обычный список Python."""
    if hasattr(v, 'value'):
        return list(v.value)
    return list(v)


def get_mleader_arrow_points(obj):
    """
    Возвращает список 2D-точек стрелок MLeader.
    Берём ТОЛЬКО ПЕРВУЮ вершину каждой линии — это стрелка
    (проверено тестом TESTMLP).
    """
    points = []
    i = 0
    while i < 20:
        try:
            v = obj.GetLeaderLineVertices(i)
        except:
            break
        vlist = var_to_list(v)
        if vlist and len(vlist) >= 3:
            points.append((vlist[0], vlist[1]))
        i += 1
    return points


def point_inside_polygon(pt, poly):
    """Проверка: точка внутри полигона (ray casting)."""
    x, y = pt
    n = len(poly)
    if n < 3:
        return False
    inside = False
    p1x, p1y = poly[0]
    for i in range(n + 1):
        p2x, p2y = poly[i % n]
        if y > min(p1y, p2y):
            if y <= max(p1y, p2y):
                if x <= max(p1x, p2x):
                    if p1y != p2y:
                        xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                    if p1x == p2x or x <= xinters:
                        inside = not inside
        p1x, p1y = p2x, p2y
    return inside


def get_frame_vertices(frame_ent):
    """Возвращает список 2D-вершин рамки (LWPOLYLINE)."""
    coords = frame_ent.Coordinates
    pts = []
    i = 0
    while i < len(coords):
        pts.append((coords[i], coords[i+1]))
        i += 2
    return pts


# ------------------------------------------------------------
# 2. Основная функция
# ------------------------------------------------------------

def main(wait_for_exit=True):
    """Заполнение выносок стояков мезонина через рамки."""
    start_total = time.time()

    # --- 1. Загрузка config ---
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(application_path, 'config.py')
    spec = importlib.util.spec_from_file_location("config", config_path)
    config = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(config)

    frames_layer = getattr(config, "STOYAK_FRAMES_LAYER", "_WB_CAB_AREAS")
    line_format = getattr(config, "STOYAK_LINE_FORMAT",
                          "{count} UTP 4x2x0.5 с отм. {floor} этажа")
    first_floor_line = getattr(config, "STOYAK_FIRST_FLOOR_LINE",
                               "{count} UTP 4x2x0.5 с отм. +2.500")
    final_line = getattr(config, "STOYAK_FINAL_LINE", "на отм. 0.000")
    regular_line = getattr(config, "STOYAK_REGULAR_LINE",
                           "{count} UTP 4x2x0.5\nна отм. 0.000")

    # --- 2. Проверяем export_final.csv ---
    if not os.path.exists("export_final.csv"):
        print("Ошибка: файл export_final.csv не найден.")
        print("Сначала выполните экспорт/чтение данных.")
        if wait_for_exit:
            input("\nНажмите Enter для выхода...")
        sys.exit(1)

    df = pd.read_csv("export_final.csv", encoding="utf-8-sig")

    if "X" not in df.columns or "Y" not in df.columns:
        print("Ошибка: в export_final.csv нет колонок X, Y.")
        if wait_for_exit:
            input("\nНажмите Enter для выхода...")
        sys.exit(1)

    df["X"] = pd.to_numeric(df["X"], errors="coerce")
    df["Y"] = pd.to_numeric(df["Y"], errors="coerce")
    df = df.dropna(subset=["X", "Y"])
    df["Этаж"] = pd.to_numeric(df["Этаж"], errors="coerce").fillna(0).astype(int)

    # --- Краткая статистика по CSV ---
    print(f"Загружено блоков из export_final.csv: {len(df)}")
    print(f"Уникальных пар (Шкаф, Этаж): "
          f"{df.groupby(['Шкаф', 'Этаж']).ngroups}")
    print(f"Уникальных шкафов: {df['Шкаф'].nunique()}")
    print(f"Уникальных этажей: {sorted(df['Этаж'].unique())}")

    # --- 3. Подключение к AutoCAD ---
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"Подключен к чертежу: {doc.Name}")

        # --- 4. Собираем рамки ---
        print(f"\nЧтение рамок на слое '{frames_layer}'...")
        ss_frames_name = "TempFramesSel"
        try:
            ss_frames = doc.SelectionSets.Add(ss_frames_name)
        except:
            ss_frames = doc.SelectionSets.Item(ss_frames_name)
            ss_frames.Clear()

        ft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 8])
        fd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                     ["LWPOLYLINE", frames_layer])
        ss_frames.Select(5, None, None, ft, fd)

        frames_count = ss_frames.Count
        print(f"Найдено рамок: {frames_count}")

        if frames_count == 0:
            print(f"\n❌ На слое '{frames_layer}' нет рамок (LWPOLYLINE).")
            ss_frames.Delete()
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            return

        frames = []
        for i in range(frames_count):
            frame_obj = ss_frames.Item(i)
            verts = get_frame_vertices(frame_obj)
            if len(verts) >= 3:
                frames.append((verts, frame_obj.Handle))
        ss_frames.Delete()
        print(f"Успешно прочитано рамок: {len(frames)}")

        # --- 5. Все MLeader ---
        ss_name = "TempMLeaderSel"
        try:
            ss = doc.SelectionSets.Add(ss_name)
        except:
            ss = doc.SelectionSets.Item(ss_name)
            ss.Clear()

        filter_type = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
        filter_data = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["MULTILEADER"])
        ss.Select(5, None, None, filter_type, filter_data)

        ml_count = ss.Count
        print(f"\nНайдено мультивыносок в чертеже: {ml_count}")

        if ml_count == 0:
            print("Нет выносок для обработки.")
            ss.Delete()
            if wait_for_exit:
                input("\nНажмите Enter для выхода...")
            return

        # --- 6. Привязка выносок через рамки ---
        print("\nПривязка выносок через рамки...")
        ml_info = {}          # {handle_ml: (cabinet, floor)}
        skipped_ml = []       # [(handle, причина), ...]

        for i in range(ml_count):
            obj = ss.Item(i)
            handle_ml = obj.Handle

            arrow_points = get_mleader_arrow_points(obj)
            if not arrow_points:
                skipped_ml.append((handle_ml, "нет точек стрелки"))
                continue

            # Ищем рамку, в которой лежит хотя бы одна стрелка
            matched_frame = None
            for verts, frame_handle in frames:
                for ap in arrow_points:
                    if point_inside_polygon(ap, verts):
                        matched_frame = (verts, frame_handle)
                        break
                if matched_frame:
                    break

            if not matched_frame:
                skipped_ml.append((handle_ml, "стрелка не в рамке"))
                continue

            # Внутри рамки — собираем блоки
            verts, _ = matched_frame
            blocks_in_frame = df[
                df.apply(lambda r: point_inside_polygon((r["X"], r["Y"]), verts), axis=1)
            ]
            if blocks_in_frame.empty:
                skipped_ml.append((handle_ml, "в рамке нет блоков"))
                continue

            # Определяем (шкаф, этаж) по большинству
            counts = blocks_in_frame.groupby(["Шкаф", "Этаж"]).size()
            cab, flr = counts.idxmax()
            ml_info[handle_ml] = (str(cab), int(flr))

        ss.Delete()

        # Итог по привязке
        print(f"\nСвязано выносок: {len(ml_info)} из {ml_count}")
        if skipped_ml:
            print(f"Пропущено выносок: {len(skipped_ml)}")
            for handle_ml, reason in skipped_ml:
                print(f"  MLeader {handle_ml}: {reason}")

        # --- 7. Группируем выноски по (шкаф, этаж) ---
        groups = {}
        for handle_ml, (cab, flr) in ml_info.items():
            groups.setdefault((cab, flr), []).append(handle_ml)

        # --- 8. Считаем блоки по (шкаф, этаж) ---
        counts_all = {}
        for _, row in df.iterrows():
            key = (str(row["Шкаф"]), int(row["Этаж"]))
            counts_all[key] = counts_all.get(key, 0) + 1

        # --- 9. Обновляем выноски ---
        print("\nОбновление выносок...")
        updated_count = 0

        for (cab, flr), handles in groups.items():
            count = counts_all.get((cab, flr), 0)
            if count == 0:
                continue

            if flr == 1:
                # Сводная выноска первого этажа
                lines = []
                cab_floors = sorted(
                    [f for (c, f) in counts_all.keys() if c == cab and f > 1],
                    reverse=True
                )
                for f in cab_floors:
                    c = counts_all[(cab, f)]
                    lines.append(line_format.format(count=c, floor=f))
                lines.append(first_floor_line.format(count=count))
                lines.append(final_line)
                new_text = "\n".join(lines)
            else:
                new_text = regular_line.format(count=count)

            for handle_ml in handles:
                try:
                    ml_obj = doc.HandleToObject(handle_ml)
                    ml_obj.TextString = new_text
                    updated_count += 1
                except Exception as e:
                    print(f"  Ошибка обновления {handle_ml}: {e}")

            print(f"  {cab}, этаж {flr}: {count} UTP (выносок: {len(handles)})")

        doc.Regen(0)
        elapsed = time.time() - start_total
        print(f"\nОбновлено выносок: {updated_count}")
        print(f"Общее время выполнения: {elapsed:.2f} сек")

    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()
        if wait_for_exit:
            input("\n✅ Готово! Нажмите Enter для выхода...")


if __name__ == "__main__":
    main(wait_for_exit=True)