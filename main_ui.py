# ============================================================
#  main_ui.py – графический интерфейс программы нумерации.
#  Версия 4.2
#
#  Что нового в v4.2:
#   • На вкладке «Дополнительно» появилось поле для имени слоя
#     с рамками (STOYAK_FRAMES_LAYER).
#   • Заполнение выносок стояков теперь работает через рамки,
#     радиус поиска больше не используется.
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
import read_current_data
import mark_new_equipment
import fill_stoyak_notes

config = None


# ------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------

def parse_port_ranges(port_str):
    """'1-18,25-42' → [[1,18],[25,42]]."""
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
    """[[1,18],[25,42]] → '1-18,25-42'."""
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
    """Многострочный текст → список строк."""
    result = []
    for line in text.splitlines():
        line = line.strip()
        if line and not line.startswith('#'):
            result.append(line)
    return result


class RedirectText(io.StringIO):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        try:
            self.text_widget.insert(tk.END, string)
            self.text_widget.see(tk.END)
            self.text_widget.update_idletasks()
        except:
            pass


def run_threaded(func, on_done=None, *args, **kwargs):
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
        self.root.title("AutoCAD Block Numbering Tool RWB v4.2")
        self.root.geometry("1250x920")

        if getattr(sys, 'frozen', False):
            self.config_dir = os.path.dirname(sys.executable)
        else:
            self.config_dir = os.path.dirname(os.path.abspath(__file__))

        # Иконка
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

        # Меню
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        about_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=about_menu)
        about_menu.add_command(label="О программе", command=self.show_about)

        # Вкладки
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Вкладка «Настройки»
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

        # Вкладка «Выполнение»
        self.run_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.run_frame, text="Выполнение")

        # Вкладка «Дополнительно»
        self.extra_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.extra_frame, text="Дополнительно")

        self._build_settings_tab()
        self._build_run_tab()
        self._build_extra_tab()

        sys.stdout = RedirectText(self.log)
        sys.stderr = RedirectText(self.log)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        try:
            self.save_settings(silent=True)
        except:
            pass
        self.root.destroy()

    def _on_mousewheel(self, event):
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
        # --- Список блоков ---
        ttk.Label(self.settings_frame, text="Список блоков для обработки:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(10, 5))

        self.table_frame = ttk.Frame(self.settings_frame)
        self.table_frame.pack(anchor=tk.W, padx=5, pady=5)

        ttk.Label(self.table_frame, text="Вкл", width=5).grid(row=0, column=0, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Имя блока", width=18).grid(row=0, column=1, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Отображаемое имя", width=22).grid(row=0, column=2, padx=5, pady=4)
        ttk.Label(self.table_frame, text="Диапазоны портов", width=28).grid(row=0, column=3, padx=5, pady=4)
        ttk.Label(self.table_frame, text="PoE", width=5).grid(row=0, column=4, padx=5, pady=4)

        self.block_rows = []
        block_presets = ["camera", "AP", "Socket_1p", "Socket_RJ-45", ""]
        display_presets = ["Видеокамера", "Точка доступа Wi-Fi", "Розетка", "Разъём RJ-45", ""]
        port_presets = ["1-18,25-42", "19-23,43-47", "1-23,25-47"]

        block_configs = getattr(config, "BLOCK_CONFIGS", [])
        while len(block_configs) < 5:
            block_configs.append({"enabled": False, "block_name": "",
                                  "display_name": "", "port_ranges": [], "poe": False})

        for i in range(5):
            cfg = block_configs[i]
            enabled_var = tk.BooleanVar(value=cfg.get("enabled", False))
            name_var = tk.StringVar(value=cfg.get("block_name", ""))
            display_var = tk.StringVar(value=cfg.get("display_name", ""))
            port_var = tk.StringVar(value=format_port_ranges(cfg.get("port_ranges", [])))
            poe_var = tk.BooleanVar(value=cfg.get("poe", False))

            tk.Checkbutton(self.table_frame, variable=enabled_var).grid(row=1+i, column=0, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=name_var, values=block_presets,
                         width=16, state="normal").grid(row=1+i, column=1, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=display_var, values=display_presets,
                         width=20, state="normal").grid(row=1+i, column=2, padx=5, pady=4)
            ttk.Combobox(self.table_frame, textvariable=port_var, values=port_presets,
                         width=26, state="normal").grid(row=1+i, column=3, padx=5, pady=4)
            tk.Checkbutton(self.table_frame, variable=poe_var).grid(row=1+i, column=4, padx=5, pady=4)

            self.block_rows.append((enabled_var, name_var, display_var, port_var, poe_var))

        self.add_row_button = ttk.Button(self.table_frame, text="+ Добавить строку",
                                         command=self._add_block_row)
        self.add_row_button.grid(row=1+len(self.block_rows), column=0, columnspan=5,
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
        self.example_text.insert(tk.END, "8.1A1,1,25-29\n8.1A1,2,25-30,40")
        self.example_text.config(state=tk.DISABLED)
        self.example_text.pack(anchor=tk.W, pady=2)

        example_menu = tk.Menu(self.example_text, tearoff=0)
        example_menu.add_command(label="Копировать", command=self._copy_example_text)
        self.example_text.bind("<Button-3>", lambda e: example_menu.post(e.x_root, e.y_root))

        # Контекстное меню пропусков
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

        self.special_enabled_var = tk.BooleanVar(
            value=getattr(config, "SPECIAL_MODE_ENABLED", False))
        tk.Checkbutton(self.settings_frame,
                       text="Обрабатывать блоки особого режима",
                       variable=self.special_enabled_var).pack(anchor=tk.W, padx=5)

        names_frame = ttk.Frame(self.settings_frame)
        names_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(names_frame, text="Имена блоков (через запятую):").pack(side=tk.LEFT)
        special_names = getattr(config, "SPECIAL_BLOCK_NAMES", [])
        self.special_names_var = tk.StringVar(value=", ".join(special_names))
        ttk.Entry(names_frame, textvariable=self.special_names_var,
                  width=55).pack(side=tk.LEFT, padx=5)

        ranges_frame = ttk.Frame(self.settings_frame)
        ranges_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(ranges_frame, text="Диапазоны портов:").pack(side=tk.LEFT)
        special_ranges = getattr(config, "SPECIAL_PORT_RANGES", [])
        self.special_ranges_var = tk.StringVar(value=format_port_ranges(special_ranges))
        ttk.Entry(ranges_frame, textvariable=self.special_ranges_var,
                  width=55).pack(side=tk.LEFT, padx=5)

        self.special_continue_var = tk.BooleanVar(
            value=getattr(config, "SPECIAL_CONTINUE_LAST_PANEL", False))
        tk.Checkbutton(self.settings_frame,
                       text="Продолжить с последней патч-панели (не начинать новую)",
                       variable=self.special_continue_var).pack(anchor=tk.W, padx=5, pady=(5, 2))

        skip_after_frame = ttk.Frame(self.settings_frame)
        skip_after_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(skip_after_frame,
                  text="Отступить от последнего порта (сколько портов пропустить):").pack(side=tk.LEFT)
        self.special_skip_ports_var = tk.StringVar(
            value=str(getattr(config, "SPECIAL_SKIP_PORTS", 0)))
        ttk.Entry(skip_after_frame, textvariable=self.special_skip_ports_var,
                  width=6).pack(side=tk.LEFT, padx=5)

        # ============================================================
        #  ДОНАЗНАЧЕНИЕ НОВОГО ОБОРУДОВАНИЯ
        # ============================================================
        ttk.Label(self.settings_frame, text="Доназначение нового оборудования:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(15, 5))

        ttk.Label(self.settings_frame,
                  text="Блоки с указанным ниже NAME считаются «новыми» и участвуют в доназначении.",
                  foreground="gray").pack(anchor=tk.W, padx=5, pady=(0, 5))

        placeholder_frame = ttk.Frame(self.settings_frame)
        placeholder_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(placeholder_frame, text="Маркер нового блока (PLACEHOLDER_MARKER):",
                  width=42).pack(side=tk.LEFT)
        self.placeholder_var = tk.StringVar(
            value=getattr(config, "PLACEHOLDER_MARKER", "Пустой"))
        ttk.Entry(placeholder_frame, textvariable=self.placeholder_var,
                  width=20).pack(side=tk.LEFT, padx=5)

        self.fill_gaps_var = tk.BooleanVar(value=getattr(config, "FILL_GAPS", False))
        tk.Checkbutton(self.settings_frame,
                       text="Заполнять свободные порты (дыры), а не только после максимума",
                       variable=self.fill_gaps_var).pack(anchor=tk.W, padx=5, pady=(5, 2))

        # ============================================================
        #  АТРИБУТЫ БЛОКОВ И ШАБЛОН ИМЕНИ
        # ============================================================
        ttk.Label(self.settings_frame, text="Атрибуты блоков и шаблон имени:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=5, pady=(15, 5))

        ttk.Label(self.settings_frame,
                  text="Имена атрибутов в блоках AutoCAD. По умолчанию: MAC, IP, NAME.",
                  foreground="gray").pack(anchor=tk.W, padx=5, pady=(0, 5))

        attr_presets = ["MAC", "IP", "NAME", "ICON", "SERIAL_NUMBER"]

        tag_cabinet_frame = ttk.Frame(self.settings_frame)
        tag_cabinet_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(tag_cabinet_frame, text="Атрибут с именем шкафа (TAG_CABINET):", width=38).pack(side=tk.LEFT)
        self.tag_cabinet_var = tk.StringVar(value=getattr(config, "TAG_CABINET", "MAC"))
        ttk.Combobox(tag_cabinet_frame, textvariable=self.tag_cabinet_var,
                     values=attr_presets, width=18, state="normal").pack(side=tk.LEFT, padx=5)

        tag_floor_frame = ttk.Frame(self.settings_frame)
        tag_floor_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(tag_floor_frame, text="Атрибут с номером этажа (TAG_FLOOR):", width=38).pack(side=tk.LEFT)
        self.tag_floor_var = tk.StringVar(value=getattr(config, "TAG_FLOOR", "IP"))
        ttk.Combobox(tag_floor_frame, textvariable=self.tag_floor_var,
                     values=attr_presets, width=18, state="normal").pack(side=tk.LEFT, padx=5)

        tag_name_frame = ttk.Frame(self.settings_frame)
        tag_name_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(tag_name_frame, text="Атрибут для итогового имени (TAG_NAME):", width=38).pack(side=tk.LEFT)
        self.tag_name_var = tk.StringVar(value=getattr(config, "TAG_NAME", "NAME"))
        ttk.Combobox(tag_name_frame, textvariable=self.tag_name_var,
                     values=attr_presets, width=18, state="normal").pack(side=tk.LEFT, padx=5)

        tag_icon_frame = ttk.Frame(self.settings_frame)
        tag_icon_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(tag_icon_frame, text="Атрибут типа/иконки (TAG_ICON):", width=38).pack(side=tk.LEFT)
        self.tag_icon_var = tk.StringVar(value=getattr(config, "TAG_ICON", "ICON"))
        ttk.Combobox(tag_icon_frame, textvariable=self.tag_icon_var,
                     values=attr_presets, width=18, state="normal").pack(side=tk.LEFT, padx=5)

        tag_serial_frame = ttk.Frame(self.settings_frame)
        tag_serial_frame.pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(tag_serial_frame, text="Атрибут серийного номера (TAG_SERIAL):", width=38).pack(side=tk.LEFT)
        self.tag_serial_var = tk.StringVar(value=getattr(config, "TAG_SERIAL", "SERIAL_NUMBER"))
        ttk.Combobox(tag_serial_frame, textvariable=self.tag_serial_var,
                     values=attr_presets, width=18, state="normal").pack(side=tk.LEFT, padx=5)

        name_format_frame = ttk.Frame(self.settings_frame)
        name_format_frame.pack(anchor=tk.W, padx=5, pady=(10, 2))
        ttk.Label(name_format_frame, text="Шаблон имени (NAME_FORMAT):", width=38).pack(side=tk.LEFT)
        self.name_format_var = tk.StringVar(
            value=getattr(config, "NAME_FORMAT", "{cabinet}/{panel:02d}.{port}"))
        ttk.Entry(name_format_frame, textvariable=self.name_format_var,
                  width=30).pack(side=tk.LEFT, padx=5)

        ttk.Label(self.settings_frame,
                  text="Доступные подстановки: {cabinet} — шкаф, {panel} — номер панели, "
                       "{port} — номер порта.\n"
                       "Пример: {cabinet}/{panel:02d}.{port} даёт имя 8.1A1/01.4",
                  foreground="gray", justify=tk.LEFT).pack(anchor=tk.W, padx=5, pady=(0, 10))

    # ------------------------------------------------------------
    # Вкладка «Выполнение»
    # ------------------------------------------------------------
    def _build_run_tab(self):
        # Полный цикл
        self.btn_full = tk.Button(self.run_frame,
                                  text="⚡ Полный цикл: экспорт → импорт → отчёт",
                                  command=self.run_full_cycle,
                                  bg="lightcoral", width=45, height=2)
        self.btn_full.pack(pady=(10, 5))

        # Три основных кнопки
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

        # --- Новые кнопки v4.2 ---
        extra_frame = ttk.Frame(self.run_frame)
        extra_frame.pack(pady=(10, 5))

        self.btn_read = tk.Button(extra_frame,
                                  text="📖 Считать данные (без изменений)",
                                  command=self.read_current_data,
                                  bg="lightsteelblue", width=35, height=2)
        self.btn_read.pack(side=tk.LEFT, padx=5)

        self.btn_mark_new = tk.Button(extra_frame,
                                      text="🆕 Промаркировать новое оборудование",
                                      command=self.mark_new_equipment,
                                      bg="khaki", width=35, height=2)
        self.btn_mark_new.pack(side=tk.LEFT, padx=5)

        # Лог
        self.log = scrolledtext.ScrolledText(self.run_frame, wrap=tk.WORD, width=120, height=32)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.log.bind("<Button-3>", self._show_context_menu)

        # Прогресс-бар
        self.progress = ttk.Progressbar(self.run_frame, mode='indeterminate', length=600)
        self.progress.pack(pady=5)

    # ------------------------------------------------------------
    # Вкладка «Дополнительно»
    # ------------------------------------------------------------
    def _build_extra_tab(self):
        ttk.Label(self.extra_frame, text="Дополнительные функции:",
                  font=('TkDefaultFont', 10, 'bold')).pack(anchor=tk.W, padx=10, pady=(15, 10))

        # Кнопка «Заполнить выноски стояков»
        self.btn_stoyak = tk.Button(
            self.extra_frame,
            text="Заполнить выноски стояков мезонина",
            command=self.fill_stoyak_notes,
            bg="lightsteelblue",
            width=45, height=2
        )
        self.btn_stoyak.pack(pady=10)

        # Пояснение
        ttk.Label(self.extra_frame,
                  text="Автоматически заполняет мультивыноски стояков:\n"
                       "• на верхних этажах — количеством кабелей этого этажа,\n"
                       "• на первом этаже — сводкой по всем этажам.\n\n"
                       "Привязка выносок к (шкаф, этаж) идёт через рамки\n"
                       "на указанном ниже слое.\n\n"
                       "Требуется предварительно выполнить «Экспорт и расчёт имён»\n"
                       "или «Считать данные (без изменений)».",
                  foreground="gray", justify=tk.LEFT).pack(anchor=tk.W, padx=10, pady=5)

        # --- Имя слоя с рамками (новое в v4.2) ---
        layer_frame = ttk.Frame(self.extra_frame)
        layer_frame.pack(anchor=tk.W, padx=10, pady=(15, 5))
        ttk.Label(layer_frame, text="Слой с рамками (STOYAK_FRAMES_LAYER):",
                  font=('TkDefaultFont', 10, 'bold')).pack(side=tk.LEFT)
        self.stoyak_frames_layer_var = tk.StringVar(
            value=getattr(config, "STOYAK_FRAMES_LAYER", "_WB_CAB_AREAS"))
        ttk.Entry(layer_frame, textvariable=self.stoyak_frames_layer_var,
                  width=30).pack(side=tk.LEFT, padx=5)

        # Лог дополнительной вкладки
        self.extra_log = scrolledtext.ScrolledText(self.extra_frame, wrap=tk.WORD,
                                                    width=120, height=25)
        self.extra_log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.extra_log.bind("<Button-3>", self._show_context_menu_extra)

        # Прогресс-бар
        self.extra_progress = ttk.Progressbar(self.extra_frame, mode='indeterminate', length=600)
        self.extra_progress.pack(pady=5)

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
        poe_var = tk.BooleanVar(value=False)

        tk.Checkbutton(self.table_frame, variable=enabled_var).grid(row=1+i, column=0, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=name_var, values=block_presets,
                     width=16, state="normal").grid(row=1+i, column=1, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=display_var, values=display_presets,
                     width=20, state="normal").grid(row=1+i, column=2, padx=5, pady=4)
        ttk.Combobox(self.table_frame, textvariable=port_var, values=port_presets,
                     width=26, state="normal").grid(row=1+i, column=3, padx=5, pady=4)
        tk.Checkbutton(self.table_frame, variable=poe_var).grid(row=1+i, column=4, padx=5, pady=4)

        self.block_rows.append((enabled_var, name_var, display_var, port_var, poe_var))
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

    def _show_context_menu_extra(self, event):
        menu = tk.Menu(self.extra_log, tearoff=0)
        menu.add_command(label="Копировать", command=self._copy_from_extra_log)
        menu.post(event.x_root, event.y_root)

    def _copy_from_extra_log(self):
        try:
            selected = self.extra_log.selection_get()
            self.extra_log.clipboard_clear()
            self.extra_log.clipboard_append(selected)
        except tk.TclError:
            pass

    # ------------------------------------------------------------
    # О программе
    # ------------------------------------------------------------
    def show_about(self):
        messagebox.showinfo(
            "О программе",
            "AutoCAD Block Numbering Tool\n"
            "Версия 4.2\n"
            "Разработано Александром Хлебновым\n"
            "2026 г.\n\n"
            "Инструмент для массовой маркировки блоков AutoCAD\n"
            "с распределением портов на патч-панелях.\n"
            "Поддерживает несколько типов блоков, особый режим,\n"
            "доназначение нового оборудования и заполнение\n"
            "выносок стояков мезонина через рамки."
        )

    # ------------------------------------------------------------
    # Сохранение настроек
    # ------------------------------------------------------------
    def save_settings(self, silent=False):
        try:
            # --- BLOCK_CONFIGS ---
            block_configs = []
            for enabled_var, name_var, display_var, port_var, poe_var in self.block_rows:
                enabled = enabled_var.get()
                name = name_var.get().strip()
                display = display_var.get().strip()
                port_str = port_var.get().strip()
                poe = poe_var.get()

                if not name and not port_str:
                    block_configs.append({
                        "enabled": False, "block_name": "",
                        "display_name": "", "port_ranges": [], "poe": False
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
                    "poe": bool(poe),
                })

            # --- SKIP_PORTS ---
            skip_text = self.skip_ports_text.get(1.0, tk.END)
            skip_list = parse_skip_ports_text(skip_text)

            # --- MAX_PANELS ---
            try:
                max_panels = int(self.max_panels_var.get().strip())
            except:
                max_panels = 20

            # --- Особый режим ---
            special_enabled = self.special_enabled_var.get()
            special_names = [s.strip() for s in self.special_names_var.get().split(",") if s.strip()]
            try:
                special_ranges = (parse_port_ranges(self.special_ranges_var.get().strip())
                                  if self.special_ranges_var.get().strip() else [])
            except:
                special_ranges = []
            special_continue = self.special_continue_var.get()
            try:
                special_skip = int(self.special_skip_ports_var.get().strip())
            except:
                special_skip = 0

            # --- Доназначение ---
            placeholder = self.placeholder_var.get().strip() or "Пустой"
            fill_gaps = self.fill_gaps_var.get()

            # --- Имя слоя с рамками ---
            frames_layer = self.stoyak_frames_layer_var.get().strip() or "_WB_CAB_AREAS"

            # --- Теги и шаблон ---
            tag_cabinet = self.tag_cabinet_var.get().strip() or "MAC"
            tag_floor = self.tag_floor_var.get().strip() or "IP"
            tag_name = self.tag_name_var.get().strip() or "NAME"
            tag_icon = self.tag_icon_var.get().strip() or "ICON"
            tag_serial = self.tag_serial_var.get().strip() or "SERIAL_NUMBER"
            name_format = self.name_format_var.get().strip() or "{cabinet}/{panel:02d}.{port}"

            # --- Запись config.py ---
            config_path = os.path.join(self.config_dir, 'config.py')
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write("# ============================================================\n")
                f.write("#  config.py – настройки программы нумерации блоков (v4.2)\n")
                f.write("#  Файл автоматически сохраняется при работе из GUI.\n")
                f.write("# ============================================================\n\n")

                f.write("# --- Общие теги атрибутов ---\n")
                f.write(f"TAG_CABINET = {repr(tag_cabinet)}\n")
                f.write(f"TAG_FLOOR   = {repr(tag_floor)}\n")
                f.write(f"TAG_NAME    = {repr(tag_name)}\n")
                f.write(f"TAG_ICON    = {repr(tag_icon)}\n")
                f.write(f"TAG_SERIAL  = {repr(tag_serial)}\n\n")

                f.write("# --- Шаблон итогового имени ---\n")
                f.write(f"NAME_FORMAT = {repr(name_format)}\n\n")

                f.write("# --- Максимум патч-панелей ---\n")
                f.write(f"MAX_PANELS_PER_CABINET = {max_panels}\n\n")

                f.write("# --- Пропуски портов ---\n")
                f.write("SKIP_PORTS = [\n")
                for line in skip_list:
                    f.write(f"    {repr(line)},\n")
                f.write("]\n\n")

                f.write("# --- Список блоков ---\n")
                f.write("BLOCK_CONFIGS = [\n")
                for cfg in block_configs:
                    f.write(f"    {repr(cfg)},\n")
                f.write("]\n\n")

                f.write("# --- Особый режим ---\n")
                f.write(f"SPECIAL_MODE_ENABLED = {repr(special_enabled)}\n")
                f.write(f"SPECIAL_BLOCK_NAMES = {repr(special_names)}\n")
                f.write(f"SPECIAL_PORT_RANGES = {repr(special_ranges)}\n")
                f.write(f"SPECIAL_CONTINUE_LAST_PANEL = {repr(special_continue)}\n")
                f.write(f"SPECIAL_SKIP_PORTS = {repr(special_skip)}\n\n")

                f.write("# --- Доназначение нового оборудования ---\n")
                f.write(f"PLACEHOLDER_MARKER = {repr(placeholder)}\n")
                f.write(f"FILL_GAPS = {repr(fill_gaps)}\n\n")

                f.write("# --- Выноски стояков ---\n")
                f.write(f"STOYAK_FRAMES_LAYER = {repr(frames_layer)}\n")
                f.write('STOYAK_LINE_FORMAT = "{count} UTP 4x2x0.5 с отм. {floor} этажа"\n')
                f.write('STOYAK_FIRST_FLOOR_LINE = "{count} UTP 4x2x0.5 с отм. +2.500"\n')
                f.write('STOYAK_FINAL_LINE = "на отм. 0.000"\n')
                f.write('STOYAK_REGULAR_LINE = "{count} UTP 4x2x0.5\\nна отм. 0.000"\n')

            if not silent:
                messagebox.showinfo("Сохранено", "Настройки сохранены в config.py")
        except Exception as e:
            if not silent:
                messagebox.showerror("Ошибка", f"Не удалось сохранить настройки:\n{e}")
            else:
                try:
                    self.log.insert(tk.END, f"\n❌ Ошибка сохранения настроек: {e}\n")
                except:
                    pass

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

    def read_current_data(self):
        """Чтение данных из AutoCAD без изменений."""
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== ЧТЕНИЕ ДАННЫХ ИЗ ЧЕРТЕЖА (без изменений) ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def on_done():
            self.progress.stop()

        run_threaded(read_current_data.main, wait_for_exit=False, on_done=on_done)

    def mark_new_equipment(self):
        """Промаркировать новое оборудование (с NAME = маркер)."""
        self.log.delete(1.0, tk.END)
        self.log.insert(tk.END, "=== ДОНАЗНАЧЕНИЕ НОВОГО ОБОРУДОВАНИЯ ===\n\n")
        self.progress.start()
        self.save_settings(silent=True)

        def on_done():
            self.progress.stop()

        run_threaded(mark_new_equipment.main, wait_for_exit=False, on_done=on_done)

    def fill_stoyak_notes(self):
        """Заполнение выносок стояков мезонина."""
        self.extra_log.delete(1.0, tk.END)
        self.extra_log.insert(tk.END, "=== Заполнение выносок стояков мезонина ===\n\n")
        self.extra_progress.start()
        self.save_settings(silent=True)

        old_stdout = sys.stdout
        sys.stdout = RedirectText(self.extra_log)

        def on_done():
            sys.stdout = old_stdout
            self.extra_progress.stop()

        run_threaded(fill_stoyak_notes.main, wait_for_exit=False, on_done=on_done)


# ============================================================
# Точка входа
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoCADApp(root)
    root.mainloop()
