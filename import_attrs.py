import sys
import os
import pandas as pd
import win32com.client

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
print(f"Подключен к чертежу: {doc.Name}")

df = pd.read_csv('camera_import.csv', encoding='utf-8-sig')
df = df.fillna('')

updated_count = 0
for _, row in df.iterrows():
    handle = row['Handle']
    try:
        obj = doc.HandleToObject(handle)
        if obj.ObjectName == "AcDbBlockReference" and obj.HasAttributes:
            atts = obj.GetAttributes()
            if hasattr(atts, 'value'):
                att_list = list(atts.value)
            else:
                att_list = list(atts)
            for attr in att_list:
                tag = attr.TagString
                if tag in row.index and str(attr.TextString) != str(row[tag]):
                    attr.TextString = str(row[tag])
            updated_count += 1
    except Exception as e:
        print(f"Не удалось обновить блок {handle}: {e}")

print(f"Обновлено блоков: {updated_count}")
doc.Regen(0)

input("\n✅ Готово! Нажмите Enter для выхода...")
