import sys
import os
import importlib.util
import win32com.client

# Загрузка config
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

config_path = os.path.join(application_path, 'config.py')
spec = importlib.util.spec_from_file_location("config", config_path)
config = importlib.util.module_from_spec(spec)
spec.loader.exec_module(config)

acad = win32com.client.GetActiveObject("AutoCAD.Application")
doc = acad.ActiveDocument
print(f"Подключен к чертежу: {doc.Name}")

BLOCK_NAME = config.BLOCK_NAME
TEMP_ATTRS = config.TEMP_ATTRIBUTES

cleared = 0
for obj in doc.ModelSpace:
    if obj.ObjectName == "AcDbBlockReference" and obj.EffectiveName == BLOCK_NAME:
        if not obj.HasAttributes:
            continue
        try:
            atts = obj.GetAttributes()
            if hasattr(atts, 'value'):
                att_list = list(atts.value)
            else:
                att_list = list(atts)
            for attr in att_list:
                if attr.TagString in TEMP_ATTRS and attr.TextString != "":
                    attr.TextString = ""
                    cleared += 1
        except Exception as e:
            print(f"Ошибка при очистке блока {obj.Handle}: {e}")

print(f"Очищено временных атрибутов: {cleared}")
doc.Regen(0)
