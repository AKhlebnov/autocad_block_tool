import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import io

# Добавляем текущую папку в путь, чтобы импортировать наши модули
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Импортируем наши скрипты как модули
import assign_names
import import_attrs
import generate_report

# Глобальная ссылка на конфиг (будет загружена при старте)
config = None

def parse_port_ranges(port_str):
    """
    Преобразует строку вида '1-19,25-42' в список,
    где каждый элемент либо число (одиночный порт), либо список [start, end].
    Пример: "19,21-23,43-47" -> [19, [21,23], [43,47]]
    """
    ranges = []
    parts = port_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            ranges.append([start, end])
        else:
            num = int(part)
            ranges.append(num)
    return ranges

def parse_skip_ports(skip_text):
    """
    Преобразует текст из поля SKIP_PORTS в словарь для config.SKIP_PORTS.
    Формат строки: cabinet,panel,ports
    Пример: 1-1.A09,1,25-34
           1-1.B13,2,25-30,40
    Возвращает {(cabinet, panel): list_of_skip_items}
    """
    skip_dict = {}
    for line in skip_text.strip().splitlines():
        if not line.strip() or line.startswith('#'):
            continue
        parts = line.split(',')
        if len(parts) < 3:
            continue
        cabinet = parts[0].strip()
        try:
            panel = int(parts[1].strip())
        except ValueError:
            continue
        ports_str = ','.join(parts[2:]).strip()
        # Разбираем порты: отдельные числа и диапазоны
        skip_items = []
        for token in ports_str.split(','):
            token = token.strip()
            if '-' in token:
                start, end = map(int, token.split('-'))
                skip_items.append((start, end))
            else:
                skip_items.append(int(token))
        skip_dict[(cabinet, panel)] = skip_items
    return skip_dict

def format_skip_ports(skip_dict):
    """Обратное преобразование словаря SKIP_PORTS в текст для отображения"""
    lines = []
    for (cabinet, panel), items in skip_dict.items():
        port_strs = []
        for item in items:
            if isinstance(item, tuple):
                port_strs.append(f"{item[0]}-{item[1]}")
            else:
                port_strs.append(str(item))
        lines.append(f"{cabinet},{panel},{','.join(port_strs)}")
    return '\n'.join(lines)

class RedirectText(io.StringIO):
    """Класс для перенаправления stdout и stderr в виджет Text"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()

def run_threaded(func, on_done=None, *args, **kwargs):
    """Запускает функцию в отдельном потоке, после завершения вызывает on_done (если передан)"""
    def wrapper():
        try:
            func(*args, **kwargs)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
        finally:
            if on_done:
                on_done()
    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()

class AutoCADApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoCAD Block Numbering Tool RWB")
        self.root.geometry("900x750")

        # Загружаем оригинальный config
        global config
        import config as cfg
        config = cfg
        # Сохраняем исходные значения (на случай, если понадобятся, но теперь не используем)
        self.original_config = {
            'BLOCK_NAME': config.BLOCK_NAME,
            'PORT_RANGES': config.PORT_RANGES,
            'SKIP_PORTS': config.SKIP_PORTS,
        }

        # Создаём вкладки
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Вкладка настроек
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Настройки")

        # Вкладка выполнения
        self.run_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.run_frame, text="Выполнение")

        # --- Настройки ---
        # BLOCK_NAME
        ttk.Label(self.settings_frame, text="Имя блока (BLOCK_NAME):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.block_name_var = tk.StringVar(value=config.BLOCK_NAME)
        self.block_name_combo = ttk.Combobox(self.settings_frame, textvariable=self.block_name_var,
                                             values=["camera", "WiFi", "SOCKET_1P"], state="normal")
        self.block_name_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

        # PORT_RANGES
        ttk.Label(self.settings_frame, text="Диапазоны портов (PORT_RANGES):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.port_ranges_var = tk.StringVar()
        # Преобразуем текущие PORT_RANGES в строку для отображения
        current_ranges = []
        for r in config.PORT_RANGES:
            if isinstance(r, (list, tuple)) and len(r) == 2:
                if r[0] == r[1]:
                    current_ranges.append(str(r[0]))
                else:
                    current_ranges.append(f"{r[0]}-{r[1]}")
            else:
                # r — это число (одиночный порт)
                current_ranges.append(str(r))
        self.port_ranges_var.set(','.join(current_ranges))
        self.port_ranges_combo = ttk.Combobox(self.settings_frame, textvariable=self.port_ranges_var,
                                              values=["1-18,25-42", "19-23,43-47"], state="normal")
        self.port_ranges_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(self.settings_frame, text="Формат: 1-18,25-42 (диапазоны через запятую)").grid(row=2, column=1, sticky=tk.W, padx=5, pady=0)

        # SKIP_PORTS
        ttk.Label(self.settings_frame, text="Пропуски портов (SKIP_PORTS):").grid(row=3, column=0, sticky=tk.NW, padx=5, pady=5)
        self.skip_ports_text = tk.Text(self.settings_frame, height=8, width=50)
        self.skip_ports_text.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        # Создаём контекстное меню для поля ввода пропусков
        self.skip_ports_menu = tk.Menu(self.skip_ports_text, tearoff=0)
        self.skip_ports_menu.add_command(label="Копировать", command=self._copy_from_skip_ports)
        self.skip_ports_menu.add_command(label="Вставить", command=self._paste_to_skip_ports)

        def show_skip_menu(event):
            self.skip_ports_menu.post(event.x_root, event.y_root)

        self.skip_ports_text.bind("<Button-3>", show_skip_menu)
        self.skip_ports_text.bind("<Control-v>", lambda e: self._paste_to_skip_ports())
        self.skip_ports_text.bind("<Control-c>", lambda e: self._copy_from_skip_ports())

        # Заполняем текущими значениями
        self.skip_ports_text.insert(tk.END, format_skip_ports(config.SKIP_PORTS))

        # Пример формата SKIP_PORTS (выделяемый и копируемый)
        example_frame = ttk.Frame(self.settings_frame)
        example_frame.grid(row=4, column=1, sticky=tk.W, padx=5, pady=0)
        ttk.Label(example_frame, text="Формат: шкаф,номер_панели,порты\nПример (можно скопировать):").pack(anchor=tk.W)
        self.example_text = tk.Text(example_frame, height=3, width=50, wrap=tk.WORD, bg='lightyellow')
        self.example_text.insert(tk.END, "1-1.A09,1,25-34\n1-1.B13,2,25-30,40")
        self.example_text.config(state=tk.DISABLED)  # запрещаем редактирование, но выделять можно
        self.example_text.pack(anchor=tk.W, pady=2)

        # Контекстное меню для копирования текста примера
        example_menu = tk.Menu(self.example_text, tearoff=0)
        example_menu.add_command(label="Копировать", command=self._copy_example_text)
        self.example_text.bind("<Button-3>", lambda event: example_menu.post(event.x_root, event.y_root))

        # Кнопка сохранения настроек в config.py (ручное сохранение)
        self.save_btn = ttk.Button(self.settings_frame, text="Сохранить настройки в config.py", command=self.save_settings)
        self.save_btn.grid(row=5, column=1, sticky=tk.W, padx=5, pady=10)

        # --- Вкладка выполнения ---
        btn_frame = ttk.Frame(self.run_frame)
        btn_frame.pack(pady=10)

        self.btn_assign = tk.Button(btn_frame, text="1. Экспорт и расчёт имён", command=self.assign_names,
                                    bg="lightblue", width=25)
        self.btn_assign.pack(side=tk.LEFT, padx=5)

        self.btn_import = tk.Button(btn_frame, text="2. Импорт в AutoCAD", command=self.import_attrs,
                                    bg="lightgreen", width=25)
        self.btn_import.pack(side=tk.LEFT, padx=5)

        self.btn_report = tk.Button(btn_frame, text="3. Сгенерировать отчёт", command=self.generate_report,
                                    bg="lightyellow", width=25)
        self.btn_report.pack(side=tk.LEFT, padx=5)

        # Область вывода логов
        self.log = scrolledtext.ScrolledText(self.run_frame, wrap=tk.WORD, width=80, height=30)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.log.bind("<Button-3>", self._show_context_menu)   # правая кнопка мыши

        # Перенаправляем stdout/stderr
        sys.stdout = RedirectText(self.log)
        sys.stderr = RedirectText(self.log)

        # Прогресс-бар
        self.progress = ttk.Progressbar(self.run_frame, mode='indeterminate', length=400)
        self.progress.pack(pady=5)

    # ------ Методы для работы с контекстным меню и буфером обмена ------
    def _copy_example_text(self):
        """Копирует весь текст из виджета примера (readonly)"""
        try:
            text = self.example_text.get(1.0, tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except:
            pass

    def _copy_from_skip_ports(self):
        """Копирует выделенный текст из поля SKIP_PORTS в буфер обмена"""
        try:
            selected = self.skip_ports_text.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
        except tk.TclError:
            pass

    def _paste_to_skip_ports(self):
        """Вставляет текст из буфера обмена в поле SKIP_PORTS в позицию курсора"""
        try:
            clipboard_text = self.root.clipboard_get()
            self.skip_ports_text.insert(tk.INSERT, clipboard_text)
        except tk.TclError:
            pass

    def _show_context_menu(self, event):
        """Показывает контекстное меню при правом клике на логе"""
        menu = tk.Menu(self.log, tearoff=0)
        menu.add_command(label="Копировать", command=self._copy_from_log)
        menu.post(event.x_root, event.y_root)

    def _copy_from_log(self):
        """Копирует выделенный текст из лога в буфер обмена"""
        try:
            selected = self.log.selection_get()
            self.log.clipboard_clear()
            self.log.clipboard_append(selected)
        except tk.TclError:
            pass

    # ------ Метод сохранения настроек (раньше отсутствовал) ------
    def save_settings(self, silent=False):
        """
        Сохраняет текущие настройки из GUI в файл config.py.
        silent=True — не показывать всплывающее окно (используется при автоматическом сохранении).
        """
        try:
            new_block_name = self.block_name_var.get().strip()
            port_ranges_str = self.port_ranges_var.get().strip()
            new_port_ranges = parse_port_ranges(port_ranges_str)
            skip_text = self.skip_ports_text.get(1.0, tk.END)
            new_skip_ports = parse_skip_ports(skip_text)

            config_path = os.path.join(os.path.dirname(__file__), 'config.py')
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write("# config.py – автоматически сохранён из GUI\n\n")
                f.write(f"BLOCK_NAME = {repr(new_block_name)}\n\n")
                f.write(f"TAG_CABINET = {repr(config.TAG_CABINET)}\n")
                f.write(f"TAG_FLOOR = {repr(config.TAG_FLOOR)}\n")
                f.write(f"TAG_NAME = {repr(config.TAG_NAME)}\n")
                f.write(f"TAG_ICON = {repr(config.TAG_ICON)}\n")
                f.write(f"TAG_SERIAL = {repr(config.TAG_SERIAL)}\n\n")
                f.write(f"PORT_RANGES = {repr(new_port_ranges)}\n\n")
                f.write(f"NAME_FORMAT = {repr(config.NAME_FORMAT)}\n\n")
                f.write(f"TEMP_ATTRIBUTES = {repr(config.TEMP_ATTRIBUTES)}\n\n")
                f.write(f"SORT_BY = {repr(config.SORT_BY)}\n")
                f.write(f"SORT_ASCENDING = {repr(config.SORT_ASCENDING)}\n\n")
                f.write(f"MAX_PANELS_PER_CABINET = {config.MAX_PANELS_PER_CABINET}\n\n")
                f.write(f"SKIP_PORTS = {repr(new_skip_ports)}\n")

            # Обновляем текущий модуль config (чтобы в памяти были новые значения)
            import importlib
            importlib.reload(config)

            if not silent:
                messagebox.showinfo("Сохранено", "Настройки сохранены в config.py")
        except Exception as e:
            if not silent:
                messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}")
            else:
                # При тихом сохранении выводим ошибку в лог
                self.log.insert(tk.END, f"\n❌ Ошибка сохранения настроек: {e}\n")

    # ------ Методы для запуска задач (с прогресс-баром) ------
    def assign_names(self):
        """Запуск экспорта и расчёта имён"""
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск экспорта и расчёта имён ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)
        def on_done():
            self.progress.stop()
        run_threaded(assign_names.main, wait_for_exit=False, on_done=on_done)

    def import_attrs(self):
        """Запуск импорта в AutoCAD"""
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск импорта в AutoCAD ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)
        def on_done():
            self.progress.stop()
        run_threaded(import_attrs.main, wait_for_exit=False, on_done=on_done)

    def generate_report(self):
        """Запуск генерации отчёта"""
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск генерации отчёта ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)
        def on_done():
            self.progress.stop()
        run_threaded(generate_report.main, wait_for_exit=False, on_done=on_done)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoCADApp(root)
    root.mainloop()
