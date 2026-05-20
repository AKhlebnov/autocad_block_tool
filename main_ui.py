import sys
import os
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
import io

# Добавляем текущую папку в путь, чтобы импортировать наши модули
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Импортируем наши скрипты как модули
import assign_names
import import_attrs
import generate_report

class RedirectText(io.StringIO):
    """Класс для перенаправления stdout и stderr в виджет Text"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()

def run_threaded(func, *args, **kwargs):
    """Запускает функцию в отдельном потоке, чтобы не блокировать GUI"""
    def wrapper():
        try:
            func(*args, **kwargs)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
        finally:
            # Восстанавливаем стандартный вывод (опционально)
            pass
    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()

class AutoCADApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoCAD Camera Tool")
        self.root.geometry("800x600")

        # Фрейм с кнопками
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        btn_assign = tk.Button(btn_frame, text="1. Экспорт и расчёт имён", command=self.assign_names, width=25, bg="lightblue")
        btn_assign.pack(side=tk.LEFT, padx=5)

        btn_import = tk.Button(btn_frame, text="2. Импорт в AutoCAD", command=self.import_attrs, width=25, bg="lightgreen")
        btn_import.pack(side=tk.LEFT, padx=5)

        btn_report = tk.Button(btn_frame, text="3. Сгенерировать отчёт", command=self.generate_report, width=25, bg="lightyellow")
        btn_report.pack(side=tk.LEFT, padx=5)

        # Текстовое поле для вывода
        self.log = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=80, height=30)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Перенаправляем вывод
        sys.stdout = RedirectText(self.log)
        sys.stderr = RedirectText(self.log)

    def assign_names(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск экспорта и расчёта имён ===\n\n")
        run_threaded(assign_names.main, wait_for_exit=False)

    def import_attrs(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск импорта в AutoCAD ===\n\n")
        run_threaded(import_attrs.main, wait_for_exit=False)

    def generate_report(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск генерации отчёта ===\n\n")
        run_threaded(generate_report.main, wait_for_exit=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoCADApp(root)
    root.mainloop()
