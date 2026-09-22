# ============================================================
#  main_ui.py – графический интерфейс программы нумерации.
#  Версия 4.0
#  Работает с таблицей блоков, особым режимом, единым отчётом.
#  Вкладка «Настройки» прокручивается.
# ============================================================

import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import io

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import assign_names
import import_attrs
import generate_report

config = None


# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------

def parse_port_ranges(port_str):
    """'1-18,25-42' -> [[1,18],[25,42]]"""
    ranges = []
    for part in port_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            start, end = map(int, part.split('-'))
            ranges.append([start, end])
        else:
            ranges.append([int(part), int(part)])
    return ranges


def format_port_ranges(ranges):
    """[[1,18],[25,42]] -> '1-18,25-42'"""
    parts = []
    for r in ranges:
        if isinstance(r, (list, tuple)) and len(r) == 2:
            if r[0] == r[1]:
                parts.append(str(r[0]))
            else:
                parts.append(f"{r[0]}-{r[1]}")
        else:
            parts.append(str(r))
    return ','.join(parts)


def parse_skip_ports_text(text):
    """Многострочный текст -> список строк (без пустых и комментариев)."""
    result = []
    for line in text.splitlines():
        line = line.strip()
        if line and not line.startswith('#'):
            result.append(line)
    return result


class RedirectText(io.StringIO):
    """Перенаправление stdout/stderr в виджет Text."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()


def run_threaded(func, on_done=None, *args, **kwargs):
    """Запуск функции в отдельном потоке."""
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


# ============================================================
# Основной класс приложения
# ============================================================

class AutoCADApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoCAD Block Numbering Tool RWB v4.0")
        self.root.geometry("1150x900")

        if getattr(sys, 'frozen', False):
            self.config_dir = os.path.dirname(sys.executable)
        else:
            self.config_dir = os.path.dirname(os.path.abspath(__file__))

        # Иконка окна
        try:
            icon_path = os.path.join(self.config_dir, 'rwb.png')
            self.icon_img = tk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, self.icon_img)
        except:
            try:
                icon_path = os.path.join(self.config_dir, 'rwb.ico')
                self.root.iconbitmap(icon_path)
            except:
                pass

        global config
        import config as cfg
        config = cfg

        # Меню «Справка»
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        about_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=about_menu)
        about_menu.add_command(label="О программе", command=self.show_about)

        # Вкладки
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ============================================================
        #  Вкладка «Настройки» — прокручиваемая
        # ============================================================
        self.settings_container = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_container, text="Настройки")

        self.settings_canvas = tk.Canvas(self.settings_container, highlightthickness=0)
        self.settings_scrollbar = ttk.Scrollbar(self.settings_container,
                                                orient="vertical",
                                                command=self.settings_canvas.yview)
        self.settings_canvas.configure(yscrollcommand=self.settings_scrollbar.set)

        self.settings_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.settings_frame = ttk.Frame(self.settings_canvas)
        self.settings_canvas_window = self.settings_canvas.create_window(
            (0, 0), window=self.settings_frame, anchor="nw"
        )

        self.settings_frame.bind(
            "<Configure>",
            lambda e: self.settings_canvas.configure(
                scrollregion=self.settings_canvas.bbox("all")
            )
        )
        self.settings_canvas.bind(
            "<Configure>",
            lambda e: self.settings_canvas.itemconfig(
                self.settings_canvas_window, width=e.width
            )
        )
        self.settings_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # ============================================================
        #  Вкладка «Выполнение»
        # ============================================================
        self.run_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.run_frame, text="Выполнение")

        self._build_settings_tab()
        self._build_run_tab()

        sys.stdout = RedirectText(self.log)
        sys.stderr = RedirectText(self.log)

    def _on_mousewheel(self, event):
        """Прокрутка колёсиком мыши работает, только если мышь над вкладкой настроек."""
        try:
            x, y = self.root.winfo_pointerxy()
            widget = self.root.winfo_containing(x, y)
            parent = widget
            while parent is not None:
                if parent == self.settings_container:
                    self.settings_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    return
                parent = parent.master if hasattr(parent, 'master') else None
        except:
            pass

    # ------------------------------------------------------------
    # Вкладка «Настройки»
    # ------------------------------------------------------------
    def _build_settings_tab(self):
        # --- Заголовок ---
        ttk.Label(self.settings_frame, text="Список блоков для обработки:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(10, 5))

        # --- Таблица блоков ---
        self.table_frame = ttk.Frame(self.settings_frame)
        self.table_frame.pack(anchor=tk.W, padx=5, pady=5)

        ttk.Label(self.table_frame, text="Вкл", width=5).grid(row=0, column=0, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Имя блока", width=18).grid(row=0, column=1, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Отображаемое имя", width=22).grid(row=0, column=2, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Диапазоны портов", width=28).grid(row=0, column=3, padx=5, pady=4)

        self.block_rows = []
        block_presets = ["camera", "AP", "Socket_1p", "Socket_RJ-45", ""]
        display_presets = ["Видеокамера", "Точка доступа Wi-Fi", "Розетка", "Разъём RJ-45", ""]
        port_presets = ["1-18,25-42", "19-23,43-47", "1-23,25-47"]

        block_configs = getattr(config, "BLOCK_CONFIGS", [])
        while len(block_configs) < 5:
            block_configs.append({"enabled": False, "block_name": "",
                                  "display_name": "", "port_ranges": []})

        for i in range(5):
            cfg = block_configs[i]
            enabled_var = tk.BooleanVar(value=cfg.get("enabled", False))
            name_var = tk.StringVar(value=cfg.get("block_name", ""))
            display_var = tk.StringVar(value=cfg.get("display_name", ""))
            port_var = tk.StringVar(value=format_port_ranges(cfg.get("port_ranges", [])))

            tk.Checkbutton(self.table_frame, variable=enabled_var).grid(row=1+i, column=0, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=name_var, values=block_presets,
                         width=16, state="normal").grid(row=1+i, column=1, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=display_var, values=display_presets,
                         width=20, state="normal").grid(row=1+i, column=2, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=port_var, values=port_presets,
                         width=26, state="normal").grid(row=1+i, column=3, padx=5, pady=4)

            self.block_rows.append((enabled_var, name_var, display_var, port_var))

        self.add_row_button = ttk.Button(self.table_frame, text="+ Добавить строку",
                                         command=self._add_block_row)
        self.add_row_button.grid(row=1+len(self.block_rows), column=0, columnspan=4,
                                 sticky=tk.W, padx=5, pady=5)

        # --- Пропуски портов ---
        ttk.Label(self.settings_frame, text="Пропуски портов:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(10, 5))

        self.skip_ports_text = tk.Text(self.settings_frame, height=6, width=90, wrap=tk.WORD)
        self.skip_ports_text.pack(anchor=tk.W, padx=5, pady=2)

        skip_list = getattr(config, "SKIP_PORTS", [])
        self.skip_ports_text.insert(tk.END, '\n'.join(skip_list))

        # Жёлтый пример
        example_frame = ttk.Frame(self.settings_frame)
        example_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(example_frame,
                  text="Формат: шкаф,номер_панели,порты\nПример (можно скопировать):").pack(anchor=tk.W)
        self.example_text = tk.Text(example_frame, height=3, width=90, wrap=tk.WORD, bg='lightyellow')
        self.example_text.insert(tk.END, "3.1C6,1,25-29\n3.1C6,2,25-30,40")
        self.example_text.config(state=tk.DISABLED)
        self.example_text.pack(anchor=tk.W, pady=2)

        example_menu = tk.Menu(self.example_text, tearoff=0)
        example_menu.add_command(label="Копировать", command=self._copy_example_text)
        self.example_text.bind("<Button-3>", lambda e: example_menu.post(e.x_root, e.y_root))

        # Контекстное меню для пропусков
        self.skip_ports_menu = tk.Menu(self.skip_ports_text, tearoff=0)
        self.skip_ports_menu.add_command(label="Копировать", command=self._copy_from_skip_ports)
        self.skip_ports_menu.add_command(label="Вставить", command=self._paste_to_skip_ports)

        self.skip_ports_text.bind("<Button-3>",
                                  lambda e: self.skip_ports_menu.post(e.x_root, e.y_root))
        self.skip_ports_text.bind("<Control-v>", lambda e: self._paste_to_skip_ports())
        self.skip_ports_text.bind("<Control-c>", lambda e: self._copy_from_skip_ports())

        # --- Максимум патч-панелей ---
        max_frame = ttk.Frame(self.settings_frame)
        max_frame.pack(anchor=tk.W, padx=5, pady=(10, 5))
        ttk.Label(max_frame, text="Максимум патч-панелей в шкафу:",
                  font=('TkDefaultFont', 10, 'bold')).pack(side=tk.LEFT)
        self.max_panels_var = tk.StringVar(value=str(getattr(config, "MAX_PANELS_PER_CABINET", 20)))
        ttk.Entry(max_frame, textvariable=self.max_panels_var, width=10).pack(side=tk.LEFT, padx=5)

        # --- ОСОБЫЙ РЕЖИМ ---
        ttk.Label(self.settings_frame, text="Особый режим (обработка после основных блоков):",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(10, 5))

        # Галочка включения особого режима
        self.special_enabled_var = tk.BooleanVar(
            value=getattr(config, "SPECIAL_MODE_ENABLED", False))
        tk.Checkbutton(self.settings_frame,
                       text="Обрабатывать блоки особого режима",
                       variable=self.special_enabled_var).pack(anchor=tk.W, padx=5)

        # Имена блоков особого режима
        names_frame = ttk.Frame(self.settings_frame)
        names_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(names_frame, text="Имена блоков (через запятую):").pack(side=tk.LEFT)
        special_names = getattr(config, "SPECIAL_BLOCK_NAMES", [])
        self.special_names_var = tk.StringVar(value=", ".join(special_names))
        ttk.Entry(names_frame, textvariable=self.special_names_var,
                  width=55).pack(side=tk.LEFT, padx=5)

        # Диапазоны портов для особого режима
        ranges_frame = ttk.Frame(self.settings_frame)
        ranges_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(ranges_frame, text="Диапазоны портов:").pack(side=tk.LEFT)
        special_ranges = getattr(config, "SPECIAL_PORT_RANGES", [])
        self.special_ranges_var = tk.StringVar(value=format_port_ranges(special_ranges))
        ttk.Entry(ranges_frame, textvariable=self.special_ranges_var,
                  width=55).pack(side=tk.LEFT, padx=5)

        # === НОВОЕ: галочка «Продолжить с последней патч-панели» ===
        self.special_continue_var = tk.BooleanVar(
            value=getattr(config, "SPECIAL_CONTINUE_LAST_PANEL", False))
        tk.Checkbutton(self.settings_frame,
                       text="Продолжить с последней патч-панели (не начинать новую)",
                       variable=self.special_continue_var).pack(anchor=tk.W, padx=5, pady=(5, 2))

        # === НОВОЕ: поле «Отступить от последнего порта» ===
        skip_after_frame = ttk.Frame(self.settings_frame)
        skip_after_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(skip_after_frame,
                  text="Отступить от последнего порта (сколько портов пропустить):").pack(side=tk.LEFT)
        self.special_skip_ports_var = tk.StringVar(
            value=str(getattr(config, "SPECIAL_SKIP_PORTS", 0)))
        ttk.Entry(skip_after_frame, textvariable=self.special_skip_ports_var,
                  width=6).pack(side=tk.LEFT, padx=5)

        # Кнопка ручного сохранения
        ttk.Button(self.settings_frame, text="Сохранить настройки в config.py",
                   command=self.save_settings).pack(anchor=tk.W, padx=5, pady=15)

    # ------------------------------------------------------------
    # Вкладка «Выполнение»
    # ------------------------------------------------------------
    def _build_run_tab(self):
        self.btn_full = tk.Button(self.run_frame,
                                  text="⚡ Полный цикл: экспорт → импорт → отчёт",
                                  command=self.run_full_cycle,
                                  bg="lightcoral", width=45, height=2)
        self.btn_full.pack(pady=(10, 5))

        sub_frame = ttk.Frame(self.run_frame)
        sub_frame.pack(pady=5)

        self.btn_assign = tk.Button(sub_frame, text="1. Экспорт и расчёт имён",
                                    command=self.assign_names, bg="lightblue", width=25)
        self.btn_assign.pack(side=tk.LEFT, padx=5)

        self.btn_import = tk.Button(sub_frame, text="2. Импорт в AutoCAD",
                                    command=self.import_attrs, bg="lightgreen", width=25)
        self.btn_import.pack(side=tk.LEFT, padx=5)

        self.btn_report = tk.Button(sub_frame, text="3. Сгенерировать отчёт",
                                    command=self.generate_report, bg="lightyellow", width=25)
        self.btn_report.pack(side=tk.LEFT, padx=5)

        self.log = scrolledtext.ScrolledText(self.run_frame, wrap=tk.WORD, width=120, height=35)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.log.bind("<Button-3>", self._show_context_menu)

        self.progress = ttk.Progressbar(self.run_frame, mode='indeterminate', length=600)
        self.progress.pack(pady=5)

    # ------------------------------------------------------------
    # Добавление строки в таблицу блоков
    # ------------------------------------------------------------
    def _add_block_row(self):
        i = len(self.block_rows)
        block_presets = ["camera", "AP", "Socket_1p", "Socket_RJ-45", ""]
        display_presets = ["Видеокамера", "Точка доступа Wi-Fi", "Розетка", "Разъём RJ-45", ""]
        port_presets = ["1-18,25-42", "19-23,43-47", "1-23,25-47"]

        enabled_var = tk.BooleanVar(value=False)
        name_var = tk.StringVar(value="")
        display_var = tk.StringVar(value="")
        port_var = tk.StringVar(value="")

        tk.Checkbutton(self.table_frame, variable=enabled_var).grid(row=1+i, column=0, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=name_var, values=block_presets,
                     width=16, state="normal").grid(row=1+i, column=1, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=display_var, values=display_presets,
                     width=20, state="normal").grid(row=1+i, column=2, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=port_var, values=port_presets,
                     width=26, state="normal").grid(row=1+i, column=3, padx=5, pady=4)

        self.block_rows.append((enabled_var, name_var, display_var, port_var))
        self.add_row_button.grid_configure(row=1 + len(self.block_rows))

        self.settings_frame.update_idletasks()
        self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all"))

    # ------------------------------------------------------------
    # Копирование / вставка
    # ------------------------------------------------------------
    def _copy_example_text(self):
        try:
            text = self.example_text.get(1.0, tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except:
            pass

    def _copy_from_skip_ports(self):
        try:
            selected = self.skip_ports_text.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
        except tk.TclError:
            pass

    def _paste_to_skip_ports(self):
        try:
            clipboard_text = self.root.clipboard_get()
            self.skip_ports_text.insert(tk.INSERT, clipboard_text)
        except tk.TclError:
            pass

    def _show_context_menu(self, event):
        menu = tk.Menu(self.log, tearoff=0)
        menu.add_command(label="Копировать", command=self._copy_from_log)
        menu.post(event.x_root, event.y_root)

    def _copy_from_log(self):
        try:
            selected = self.log.selection_get()
            self.log.clipboard_clear()
            self.log.clipboard_append(selected)
        except tk.TclError:
            pass

    # ------------------------------------------------------------
    # О программе
    # ------------------------------------------------------------
    def show_about(self):
        messagebox.showinfo(
            "О программе",
            "AutoCAD Block Numbering Tool\n"
            "Версия 4.0\n"
            "Разработано Александром Хлебновым\n"
            "2026 г.\n\n"
            "Инструмент для массовой маркировки блоков AutoCAD\n"
            "с распределением портов на патч-панелях.\n"
            "Поддерживает несколько типов блоков за один запуск."
        )

    # ------------------------------------------------------------
    # Сохранение настроек в config.py
    # ------------------------------------------------------------
    def save_settings(self, silent=False):
        try:
            block_configs = []
            for enabled_var, name_var, display_var, port_var in self.block_rows:
                enabled = enabled_var.get()
                name = name_var.get().strip()
                display = display_var.get().strip()
                port_str = port_var.get().strip()

                if not name and not port_str:
                    block_configs.append({
                        "enabled": False, "block_name": "",
                        "display_name": "", "port_ranges": []
                    })
                    continue

                try:
                    port_ranges = parse_port_ranges(port_str) if port_str else []
                except Exception as e:
                    raise Exception(f"Ошибка в диапазонах портов для блока '{name}': {e}")

                block_configs.append({
                    "enabled": bool(enabled),
                    "block_name": name,
                    "display_name": display,
                    "port_ranges": port_ranges,
                })

            skip_text = self.skip_ports_text.get(1.0, tk.END)
            skip_list = parse_skip_ports_text(skip_text)

            try:
                max_panels = int(self.max_panels_var.get().strip())
            except:
                max_panels = 20

            special_enabled = self.special_enabled_var.get()
            special_names = [s.strip() for s in self.special_names_var.get().split(",") if s.strip()]
            try:
                special_ranges = (parse_port_ranges(self.special_ranges_var.get().strip())
                                  if self.special_ranges_var.get().strip() else [])
            except:
                special_ranges = []

            # Новые параметры особого режима
            special_continue = self.special_continue_var.get()
            try:
                special_skip = int(self.special_skip_ports_var.get().strip())
            except:
                special_skip = 0

            config_path = os.path.join(self.config_dir, 'config.py')
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write("# ============================================================\n")
                f.write("#  config.py – настройки программы нумерации блоков (v4.0)\n")
                f.write("#  Файл автоматически сохраняется при работе из GUI.\n")
                f.write("# ============================================================\n\n")
                f.write(f"TAG_CABINET = {repr(config.TAG_CABINET)}\n")
                f.write(f"TAG_FLOOR   = {repr(config.TAG_FLOOR)}\n")
                f.write(f"TAG_NAME    = {repr(config.TAG_NAME)}\n")
                f.write(f"TAG_ICON    = {repr(config.TAG_ICON)}\n")
                f.write(f"TAG_SERIAL  = {repr(config.TAG_SERIAL)}\n\n")
                f.write(f"NAME_FORMAT = {repr(config.NAME_FORMAT)}\n\n")
                f.write(f"MAX_PANELS_PER_CABINET = {max_panels}\n\n")
                f.write("SKIP_PORTS = [\n")
                for line in skip_list:
                    f.write(f"    {repr(line)},\n")
                f.write("]\n\n")
                f.write("BLOCK_CONFIGS = [\n")
                for cfg in block_configs:
                    f.write(f"    {repr(cfg)},\n")
                f.write("]\n\n")
                f.write(f"SPECIAL_MODE_ENABLED = {repr(special_enabled)}\n")
                f.write(f"SPECIAL_BLOCK_NAMES = {repr(special_names)}\n")
                f.write(f"SPECIAL_PORT_RANGES = {repr(special_ranges)}\n")
                f.write(f"SPECIAL_CONTINUE_LAST_PANEL = {repr(special_continue)}\n")
                f.write(f"SPECIAL_SKIP_PORTS = {repr(special_skip)}\n")

            if not silent:
                messagebox.showinfo("Сохранено", "Настройки сохранены в config.py")
        except Exception as e:
            if not silent:
                messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}")
            else:
                self.log.insert(tk.END, f"\n❌ Ошибка сохранения настроек: {e}\n")

    # ------------------------------------------------------------
    # Запуск задач
    # ------------------------------------------------------------
    def assign_names(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск экспорта и расчёта имён ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def on_done():
            self.progress.stop()

        run_threaded(assign_names.main, wait_for_exit=False, on_done=on_done)

    def import_attrs(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск импорта в AutoCAD ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def on_done():
            self.progress.stop()

        run_threaded(import_attrs.main, wait_for_exit=False, on_done=on_done)

    def generate_report(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== Запуск генерации отчёта ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def on_done():
            self.progress.stop()

        run_threaded(generate_report.main, wait_for_exit=False, on_done=on_done)

    def run_full_cycle(self):
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== ПОЛНЫЙ ЦИКЛ: экспорт → импорт → отчёт ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def task():
            print("--- Этап 1: Экспорт и расчёт имён ---")
            assign_names.main(wait_for_exit=False)
            print("\n--- Этап 2: Импорт в AutoCAD ---")
            import_attrs.main(wait_for_exit=False)
            print("\n--- Этап 3: Генерация отчёта ---")
            generate_report.main(wait_for_exit=False)
            print("\n=== Полный цикл завершён ===")

        def on_done():
            self.progress.stop()

        run_threaded(task, on_done=on_done)


# ============================================================
# Точка входа
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoCADApp(root)
    root.mainloop()
