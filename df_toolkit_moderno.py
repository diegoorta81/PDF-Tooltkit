# df_toolkit_moderno.py
"""
PDF Toolkit Moderno - Versión final con zona de logs compartida
- Título: "PDF Toolkit Moderno"
- Logs diarios en carpeta configurada en config.json (df_toolkit_YYYY-MM-DD.log) (modo append)
- Zona de logs compartida en la parte inferior (altura inicial ≈10 líneas, expandible)
- Logs más recientes arriba, cada línea con hora [HH:MM:SS]
- Botones: "Copiar log" y "Limpiar log" (limpia la vista, no los archivos)
- Mantiene menú superior, menú lateral y funcionalidades de las pestañas
"""

import os
import sys
import json
import threading
import queue
import logging
from datetime import datetime, date
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar, BooleanVar, Toplevel, LEFT

import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledText as TBScrolledText

# pdfminer + odfpy
from pdfminer.high_level import extract_text
from odf.opendocument import OpenDocumentText
from odf.text import P, Span

# Intentar importar Toast y Icon de forma flexible
try:
    from ttkbootstrap.toast import ToastNotification
except Exception:
    ToastNotification = None

# icon handling: try multiple imports/apis
try:
    from ttkbootstrap.icons import get_icon  # type: ignore
except Exception:
    get_icon = None
    try:
        from ttkbootstrap.icons import Icon  # type: ignore
    except Exception:
        Icon = None

# ---------------------------
# Base dir: compatible con .exe empaquetado
# ---------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

CONFIG_FILE = BASE_DIR / "config.json"
DEFAULT_CONFIG = {
    "log_folder": str(BASE_DIR / "logs"),
    "result_folder": str(BASE_DIR / "resultados"),
    "theme": "flatly"
}

# ---------------------------
# Config load/save
# ---------------------------
def load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                merged = {**DEFAULT_CONFIG, **cfg}
                return merged
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

config = load_config()
os.makedirs(config["log_folder"], exist_ok=True)
os.makedirs(config["result_folder"], exist_ok=True)

# ---------------------------
# Logging (daily file handler)
# ---------------------------
def init_logging_daily(log_folder: str):
    log_folder = Path(log_folder)
    log_folder.mkdir(parents=True, exist_ok=True)
    today = date.today().strftime("%Y-%m-%d")
    log_filename = log_folder / f"df_toolkit_{today}.log"

    # remove existing handlers to avoid duplicate logs
    root_logger = logging.getLogger()
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)

    # basic file handler (append mode)
    fh = logging.FileHandler(str(log_filename), mode="a", encoding="utf-8")
    fh.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    fh.setFormatter(formatter)
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(fh)

    return str(log_filename)

# initialize logging file for today (append mode)
CURRENT_LOG_PATH = init_logging_daily(config.get("log_folder", DEFAULT_CONFIG["log_folder"]))

def write_log_file(msg: str, level: str = "info"):
    if level == "info":
        logging.info(msg)
    elif level == "error":
        logging.error(msg)
    else:
        logging.debug(msg)

# ---------------------------
# Utilidades
# ---------------------------
def unique_filename(path: str) -> str:
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base}_{counter}{ext}"
        counter += 1
    return new_path

def open_file_with_default_app(path: str):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass

def get_icon_compat(name: str, size: int = 16):
    try:
        if get_icon is not None:
            try:
                return get_icon(name, size=size)
            except Exception:
                pass
        if 'Icon' in globals() and Icon is not None:
            try:
                return Icon(name, size=(size, size))
            except Exception:
                try:
                    return Icon().get(name, size=size)
                except Exception:
                    return None
        return None
    except Exception:
        return None

# ---------------------------
# Worker controlable
# ---------------------------
class Worker(threading.Thread):
    def __init__(self, target, args=()):
        super().__init__(daemon=True)
        self._target = target
        self._args = args
        self.stop_event = threading.Event()

    def stop(self):
        self.stop_event.set()

    def run(self):
        try:
            self._target(self.stop_event, *self._args)
        except Exception as e:
            write_log_file(f"Error en hilo: {e}", level="error")

# ---------------------------
# Ventana de configuración
# ---------------------------
class ConfigWindow(Toplevel):
    def __init__(self, parent, current_config: dict, apply_callback):
        super().__init__(parent)
        self.title("Configuración de carpetas")
        self.geometry("640x260")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.apply_callback = apply_callback

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Carpeta de logs:").grid(row=0, column=0, sticky="w", padx=6, pady=(6,2))
        self.logs_var = StringVar(value=current_config.get("log_folder", DEFAULT_CONFIG["log_folder"]))
        ttk.Entry(frm, textvariable=self.logs_var, width=60).grid(row=1, column=0, columnspan=2, padx=6, pady=2, sticky="we")
        ttk.Button(frm, text="Seleccionar...", bootstyle="secondary-outline", command=lambda: self._select_folder(self.logs_var)).grid(row=1, column=2, padx=6, pady=2)

        ttk.Label(frm, text="Carpeta de resultados:").grid(row=2, column=0, sticky="w", padx=6, pady=(10,2))
        self.result_var = StringVar(value=current_config.get("result_folder", DEFAULT_CONFIG["result_folder"]))
        ttk.Entry(frm, textvariable=self.result_var, width=60).grid(row=3, column=0, columnspan=2, padx=6, pady=2, sticky="we")
        ttk.Button(frm, text="Seleccionar...", bootstyle="secondary-outline", command=lambda: self._select_folder(self.result_var)).grid(row=3, column=2, padx=6, pady=2)

        ttk.Label(frm, text="Tema:").grid(row=4, column=0, sticky="w", padx=6, pady=(10,2))
        self.theme_var = StringVar(value=current_config.get("theme", DEFAULT_CONFIG["theme"]))
        themes = ttk.Style().theme_names()
        theme_combo = ttk.Combobox(frm, values=themes, textvariable=self.theme_var, width=24, state="readonly")
        theme_combo.grid(row=4, column=1, padx=6, sticky="w")
        theme_combo.bind("<<ComboboxSelected>>", lambda e: ttk.Style().theme_use(self.theme_var.get()))

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, pady=12)
        ttk.Button(btns, text="Guardar", bootstyle="success", command=self._save).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Cancelar", bootstyle="secondary", command=self.destroy).pack(side=LEFT, padx=6)

        frm.columnconfigure(0, weight=0)
        frm.columnconfigure(1, weight=1)

    def _select_folder(self, var: StringVar):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _save(self):
        cfg_log = self.logs_var.get().strip() or DEFAULT_CONFIG["log_folder"]
        cfg_res = self.result_var.get().strip() or DEFAULT_CONFIG["result_folder"]
        cfg_theme = self.theme_var.get().strip() or DEFAULT_CONFIG["theme"]
        cfg = {
            "log_folder": cfg_log,
            "result_folder": cfg_res,
            "theme": cfg_theme
        }
        try:
            os.makedirs(cfg_log, exist_ok=True)
            os.makedirs(cfg_res, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pueden crear las carpetas seleccionadas: {e}")
            return
        save_config(cfg)
        try:
            self.apply_callback(cfg)
        except Exception:
            pass
        try:
            if ToastNotification:
                ToastNotification(title="Configuración", message="Guardada correctamente", duration=1800, alert=True).show_toast()
        except Exception:
            pass
        self.destroy()

# ---------------------------
# Aplicación principal
# ---------------------------
class DFToolkitModern:
    def __init__(self, root: ttk.Window):
        self.root = root
        self.root.title("PDF Toolkit Moderno")
        self.root.geometry("1100x720")
        self.root.minsize(980, 600)

        # style
        theme = config.get("theme", DEFAULT_CONFIG["theme"])
        try:
            self.style = ttk.Style(theme=theme)
        except Exception:
            self.style = ttk.Style()

        # comms
        self.queue = queue.Queue()
        self.active_worker = None

        # state vars
        self.status_var = StringVar(value="Listo ✅")
        self.progress_var = IntVar(value=0)

        # paths controlled by config window
        self.logs_path = config.get("log_folder", DEFAULT_CONFIG["log_folder"])
        self.result_folder = config.get("result_folder", DEFAULT_CONFIG["result_folder"])
        os.makedirs(self.logs_path, exist_ok=True)
        os.makedirs(self.result_folder, exist_ok=True)

        # re-init logging to match config folder (append mode)
        global CURRENT_LOG_PATH
        CURRENT_LOG_PATH = init_logging_daily(self.logs_path)

        # selected pdf (search)
        self.selected_pdf = StringVar(value="")

        # progress variables (per feature)
        self.progress_search_var = IntVar(value=0)
        self.label_search_status = StringVar(value="0 / 0 páginas")

        self.progress_number_var = IntVar(value=0)
        self.label_number_status = StringVar(value="0 / 0 páginas")

        self.progress_merge_var = IntVar(value=0)
        self.label_merge_status = StringVar(value="0 / 0 archivos")

        self.progress_extract_var = IntVar(value=0)
        self.label_extract_status = StringVar(value="0 / 0 páginas")

        self.progress_pdf2odt_var = IntVar(value=0)
        self.label_pdf2odt_status = StringVar(value="0 / 0")

        # build UI
        self._build_menu()
        self._build_ui()

        # load today's log file into view
        self._load_today_log_into_view()

        # process queue periodically
        self.root.after(150, self._process_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------
    # Menú (configuración)
    # ---------------------------
    def _build_menu(self):
        menubar = ttk.Menu(self.root)
        file_menu = ttk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Configuración", command=self._open_config_window)
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self.root.quit)
        menubar.add_cascade(label="Archivo", menu=file_menu)

        help_menu = ttk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Acerca de", command=lambda: messagebox.showinfo("Acerca de", "PDF Toolkit Moderno\nby Diego Orta"))
        menubar.add_cascade(label="Ayuda", menu=help_menu)

        self.root.config(menu=menubar)

    def _open_config_window(self):
        def apply_cfg(cfg):
            self.logs_path = cfg.get("log_folder", self.logs_path)
            self.result_folder = cfg.get("result_folder", self.result_folder)
            try:
                os.makedirs(self.logs_path, exist_ok=True)
                os.makedirs(self.result_folder, exist_ok=True)
            except Exception:
                pass
            try:
                global CURRENT_LOG_PATH
                CURRENT_LOG_PATH = init_logging_daily(self.logs_path)
                write_log_file("Configuración actualizada")
                # reload today's log view from new folder
                self._load_today_log_into_view()
            except Exception:
                pass

        ConfigWindow(self.root, config, apply_callback=apply_cfg)

    # ---------------------------
    # UI principal (sidebar + content + shared log)
    # ---------------------------
    def _build_ui(self):
        container = ttk.Frame(self.root, padding=8)
        container.pack(fill="both", expand=True)

        # Sidebar
        sidebar = ttk.Frame(container, width=240)
        sidebar.pack(side="left", fill="y", padx=(0,8))
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="PDF Toolkit", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(8,6), padx=8)
        ttk.Label(sidebar, text="Acciones principales", bootstyle="secondary").pack(anchor="w", padx=8, pady=(0,6))

        ico_search = get_icon_compat("search", size=16)
        ico_number = get_icon_compat("hash", size=16)
        ico_merge = get_icon_compat("files", size=16)
        ico_extract = get_icon_compat("file-export", size=16)
        ico_pdf2odt = get_icon_compat("file-text", size=16)
        ico_settings = get_icon_compat("gear", size=14)
        ico_results = get_icon_compat("folder", size=14)

        ttk.Button(sidebar, text=" Buscar texto", image=ico_search, compound=LEFT, bootstyle="info-outline", command=lambda: self._show_frame("search")).pack(fill="x", padx=8, pady=6)
        ttk.Button(sidebar, text=" Numerar páginas", image=ico_number, compound=LEFT, bootstyle="info-outline", command=lambda: self._show_frame("number")).pack(fill="x", padx=8, pady=6)
        ttk.Button(sidebar, text=" Unir PDFs", image=ico_merge, compound=LEFT, bootstyle="info-outline", command=lambda: self._show_frame("merge")).pack(fill="x", padx=8, pady=6)
        ttk.Button(sidebar, text=" Extraer páginas", image=ico_extract, compound=LEFT, bootstyle="info-outline", command=lambda: self._show_frame("extract")).pack(fill="x", padx=8, pady=6)
        ttk.Button(sidebar, text=" PDF → ODT", image=ico_pdf2odt, compound=LEFT, bootstyle="info-outline", command=lambda: self._show_frame("pdf2odt")).pack(fill="x", padx=8, pady=6)

        ttk.Separator(sidebar).pack(fill="x", pady=8, padx=8)
        ttk.Button(sidebar, text=" Configuración", image=ico_settings, compound=LEFT, bootstyle="secondary-outline", command=self._open_config_window).pack(fill="x", padx=8, pady=6)
        ttk.Button(sidebar, text=" Resultados", image=ico_results, compound=LEFT, bootstyle="secondary-outline", command=self.abrir_carpeta_resultados).pack(fill="x", padx=8, pady=6)

        # Right side: content + shared log (use grid to allow proportional resizing)
        right_frame = ttk.Frame(container)
        right_frame.pack(side="left", fill="both", expand=True)

        right_frame.grid_rowconfigure(0, weight=3)  # main content area
        right_frame.grid_rowconfigure(1, weight=1)  # log area (expandible)
        right_frame.grid_columnconfigure(0, weight=1)

        # Main content frames (stacked, only one visible)
        main_content = ttk.Frame(right_frame)
        main_content.grid(row=0, column=0, sticky="nsew")

        self.frames = {}
        for name in ("search", "number", "merge", "extract", "pdf2odt"):
            f = ttk.Frame(main_content)
            f.pack(fill="both", expand=True)
            f.pack_forget()
            self.frames[name] = f

        self._build_search_frame(self.frames["search"])
        self._build_number_frame(self.frames["number"])
        self._build_merge_frame(self.frames["merge"])
        self._build_extract_frame(self.frames["extract"])
        self._build_pdf2odt_frame(self.frames["pdf2odt"])

        self._show_frame("search")

        # Shared log area (bottom)
        log_frame = ttk.Labelframe(right_frame, text="Logs (día actual)", padding=6)
        log_frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(6,0))

        # Buttons for log actions
        btns_frame = ttk.Frame(log_frame)
        btns_frame.pack(fill="x", pady=(0,6))

        ttk.Button(btns_frame, text="Copiar log", bootstyle="outline-secondary", command=self._copy_log).pack(side=LEFT, padx=4)
        ttk.Button(btns_frame, text="Limpiar log", bootstyle="outline-danger", command=self._clear_log_view).pack(side=LEFT, padx=4)
        ttk.Button(btns_frame, text="Abrir carpeta logs", bootstyle="outline-secondary", command=lambda: open_file_with_default_app(self.logs_path)).pack(side=LEFT, padx=4)

        # Shared scrolled text: height initial ~10 lines, expandible
        self.shared_log = TBScrolledText(log_frame, height=10)
        self.shared_log.pack(fill="both", expand=True)
        # make non-editable by default
        try:
            self.shared_log.configure(state="disabled")
        except Exception:
            # fallback to text widget config method
            try:
                self.shared_log.config(state="disabled")
            except Exception:
                pass

        # Statusbar
        statusbar = ttk.Frame(self.root)
        statusbar.pack(fill="x", side="bottom", pady=4)
        ttk.Label(statusbar, textvariable=self.status_var).pack(side="left", padx=8)
        self.global_progress = ttk.Progressbar(statusbar, variable=self.progress_var, maximum=100, bootstyle="info-striped")
        self.global_progress.pack(side="left", fill="x", expand=True, padx=8)

    def _show_frame(self, name: str):
        for n, f in self.frames.items():
            if n == name:
                f.pack(fill="both", expand=True)
            else:
                f.pack_forget()

    # ---------------------------
    # Frame builders (no per-tab log text areas)
    # ---------------------------
    def _build_search_frame(self, frame):
        frm_top = ttk.LabelFrame(frame, text="Buscar texto en PDF", padding=10)
        frm_top.pack(fill="x", padx=8, pady=8)

        ttk.Label(frm_top, text="Archivo PDF:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(frm_top, textvariable=self.selected_pdf, width=70).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(frm_top, text="Seleccionar...", bootstyle="secondary", command=self._select_search_pdf).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(frm_top, text="Texto 1:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.search_text1 = StringVar()
        ttk.Entry(frm_top, textvariable=self.search_text1, width=50).grid(row=1, column=1, padx=4, pady=2)
        ttk.Label(frm_top, text="Texto 2:").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        self.search_text2 = StringVar()
        ttk.Entry(frm_top, textvariable=self.search_text2, width=50).grid(row=2, column=1, padx=4, pady=2)
        ttk.Label(frm_top, text="Texto 3:").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        self.search_text3 = StringVar()
        ttk.Entry(frm_top, textvariable=self.search_text3, width=50).grid(row=3, column=1, padx=4, pady=2)

        self.case_sensitive = BooleanVar(value=False)
        ttk.Checkbutton(frm_top, text="Respetar mayúsculas/minúsculas", variable=self.case_sensitive).grid(row=4, column=1, sticky="w", padx=4, pady=2)
        self.require_all = BooleanVar(value=False)
        ttk.Checkbutton(frm_top, text="Buscar todos los textos (AND) / cualquiera (OR)", variable=self.require_all).grid(row=5, column=1, sticky="w", padx=4, pady=2)
        self.search_preview = BooleanVar(value=True)
        ttk.Checkbutton(frm_top, text="Abrir PDF resultante al terminar", variable=self.search_preview).grid(row=6, column=1, sticky="w", padx=4, pady=2)

        frm_actions = ttk.Frame(frame)
        frm_actions.pack(fill="x", padx=8, pady=4)
        self.btn_search_start = ttk.Button(frm_actions, text="Iniciar búsqueda", bootstyle="primary", command=self._start_search)
        self.btn_search_start.pack(side=LEFT, padx=6)
        self.btn_search_stop = ttk.Button(frm_actions, text="Detener", bootstyle="danger-outline", command=self._stop_worker, state="disabled")
        self.btn_search_stop.pack(side=LEFT, padx=6)

        # progress area (kept)
        frm_progress = ttk.Frame(frame)
        frm_progress.pack(fill="x", padx=8, pady=6)
        ttk.Progressbar(frm_progress, variable=self.progress_search_var, maximum=100).pack(fill="x", padx=4)
        ttk.Label(frm_progress, textvariable=self.label_search_status).pack(anchor="w", padx=4)

    def _select_search_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.selected_pdf.set(path)

    def _build_number_frame(self, frame):
        frm = ttk.LabelFrame(frame, text="Numeración de páginas", padding=10)
        frm.pack(fill="x", padx=8, pady=8)

        ttk.Label(frm, text="Archivo PDF:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.number_pdf_path = StringVar()
        ttk.Entry(frm, textvariable=self.number_pdf_path, width=60).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(frm, text="Seleccionar...", bootstyle="secondary", command=self._select_number_pdf).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(frm, text="Número inicial:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.number_start = IntVar(value=1)
        ttk.Entry(frm, textvariable=self.number_start, width=10).grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Página de comienzo:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.page_start = IntVar(value=1)
        ttk.Entry(frm, textvariable=self.page_start, width=10).grid(row=2, column=1, sticky="w", padx=4, pady=4)

        self.use_number_initial = BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Iniciar numeración en la página de comienzo", variable=self.use_number_initial).grid(row=3, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Coordenada X (pts desde izquierda):").grid(row=4, column=0, sticky="w", padx=4, pady=4)
        self.coord_x = IntVar(value=50)
        ttk.Entry(frm, textvariable=self.coord_x, width=10).grid(row=4, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Coordenada Y (pts desde fondo):").grid(row=5, column=0, sticky="w", padx=4, pady=4)
        self.coord_y = IntVar(value=50)
        ttk.Entry(frm, textvariable=self.coord_y, width=10).grid(row=5, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Prefijo texto:").grid(row=6, column=0, sticky="w", padx=4, pady=4)
        self.prefix_text = StringVar(value="Página ")
        ttk.Entry(frm, textvariable=self.prefix_text, width=30).grid(row=6, column=1, sticky="w", padx=4, pady=4)

        self.number_preview = BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Abrir PDF resultante al terminar", variable=self.number_preview).grid(row=7, column=1, sticky="w", padx=4, pady=4)

        frm_btns = ttk.Frame(frame)
        frm_btns.pack(fill="x", padx=8, pady=4)
        self.btn_number_start = ttk.Button(frm_btns, text="Numerar", bootstyle="primary", command=self._start_numbering)
        self.btn_number_start.pack(side=LEFT, padx=6)
        self.btn_number_stop = ttk.Button(frm_btns, text="Detener", bootstyle="danger-outline", command=self._stop_worker, state="disabled")
        self.btn_number_stop.pack(side=LEFT, padx=6)

        frm_progress = ttk.Frame(frame)
        frm_progress.pack(fill="x", padx=8, pady=6)
        ttk.Progressbar(frm_progress, variable=self.progress_number_var, maximum=100).pack(fill="x", padx=4)
        ttk.Label(frm_progress, textvariable=self.label_number_status).pack(anchor="w", padx=4)

    def _select_number_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.number_pdf_path.set(path)

    def _build_merge_frame(self, frame):
        frm = ttk.LabelFrame(frame, text="Unir PDFs", padding=10)
        frm.pack(fill="both", expand=True, padx=8, pady=8)

        left = ttk.Frame(frm)
        left.pack(side="left", fill="both", expand=True, padx=6, pady=6)

        frame_list = ttk.Frame(left)
        frame_list.pack(fill="both", expand=True)
        scroll_y = ttk.Scrollbar(frame_list, orient="vertical")
        self.listbox_files = tk.Listbox(frame_list, height=14, yscrollcommand=scroll_y.set, selectmode=tk.SINGLE)
        scroll_y.config(command=self.listbox_files.yview)
        self.listbox_files.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        frm_list_btns = ttk.Frame(left)
        frm_list_btns.pack(pady=4)
        ttk.Button(frm_list_btns, text="Añadir archivos...", command=self._add_merge_files).grid(row=0, column=0, padx=2)
        ttk.Button(frm_list_btns, text="Quitar seleccionado", command=self._remove_merge_selected).grid(row=0, column=1, padx=2)
        ttk.Button(frm_list_btns, text="Subir", command=self._move_merge_up).grid(row=1, column=0, padx=2, pady=2)
        ttk.Button(frm_list_btns, text="Bajar", command=self._move_merge_down).grid(row=1, column=1, padx=2, pady=2)

        right = ttk.Frame(frm)
        right.pack(side="left", fill="both", expand=True, padx=6)
        ttk.Label(right, text="Archivo de salida:").pack(anchor="w")
        self.merge_output = StringVar(value="unido_resultado.pdf")
        ttk.Entry(right, textvariable=self.merge_output, width=60).pack(fill="x", padx=4, pady=4)
        self.merge_preview = BooleanVar(value=True)
        ttk.Checkbutton(right, text="Abrir PDF resultante al terminar", variable=self.merge_preview).pack(anchor="w", padx=4, pady=4)
        frm_btns = ttk.Frame(right)
        frm_btns.pack(pady=4)
        self.btn_merge_start = ttk.Button(frm_btns, text="Unir", bootstyle="primary", command=self._start_merge)
        self.btn_merge_start.pack(side=LEFT, padx=6)
        self.btn_merge_stop = ttk.Button(frm_btns, text="Detener", bootstyle="danger-outline", command=self._stop_worker, state="disabled")
        self.btn_merge_stop.pack(side=LEFT, padx=6)

        frm_progress = ttk.Frame(frm)
        frm_progress.pack(fill="x", padx=8, pady=6)
        ttk.Progressbar(frm_progress, variable=self.progress_merge_var, maximum=100).pack(fill="x", padx=4)
        ttk.Label(frm_progress, textvariable=self.label_merge_status).pack(anchor="w", padx=4)

    def _add_merge_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for p in paths:
            self.listbox_files.insert("end", p)

    def _remove_merge_selected(self):
        sel = self.listbox_files.curselection()
        if sel:
            self.listbox_files.delete(sel[0])

    def _move_merge_up(self):
        sel = self.listbox_files.curselection()
        if sel and sel[0] > 0:
            i = sel[0]
            item = self.listbox_files.get(i)
            self.listbox_files.delete(i)
            self.listbox_files.insert(i-1, item)
            self.listbox_files.selection_set(i-1)

    def _move_merge_down(self):
        sel = self.listbox_files.curselection()
        if sel and sel[0] < self.listbox_files.size() - 1:
            i = sel[0]
            item = self.listbox_files.get(i)
            self.listbox_files.delete(i)
            self.listbox_files.insert(i+1, item)
            self.listbox_files.selection_set(i+1)

    def _build_extract_frame(self, frame):
        frm_top = ttk.LabelFrame(frame, text="Extraer páginas", padding=10)
        frm_top.pack(fill="x", padx=8, pady=8)

        ttk.Label(frm_top, text="Archivo PDF:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.extract_pdf_path = StringVar()
        ttk.Entry(frm_top, textvariable=self.extract_pdf_path, width=60).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(frm_top, text="Seleccionar...", bootstyle="secondary", command=self._select_extract_pdf).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(frm_top, text="Rangos de páginas:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.page_ranges = StringVar()
        ttk.Entry(frm_top, textvariable=self.page_ranges, width=60).grid(row=1, column=1, padx=4, pady=4)
        ttk.Label(frm_top, text="Ej: 1-3,6-7").grid(row=1, column=2, sticky="w")

        self.extract_preview = BooleanVar(value=True)
        ttk.Checkbutton(frm_top, text="Abrir PDF resultante al terminar", variable=self.extract_preview).grid(row=2, column=1, sticky="w", padx=4, pady=4)

        frm_actions = ttk.Frame(frame)
        frm_actions.pack(fill="x", padx=8, pady=4)
        self.btn_extract_start = ttk.Button(frm_actions, text="Extraer", bootstyle="primary", command=self._start_extract)
        self.btn_extract_start.pack(side=LEFT, padx=6)
        self.btn_extract_stop = ttk.Button(frm_actions, text="Detener", bootstyle="danger-outline", command=self._stop_worker, state="disabled")
        self.btn_extract_stop.pack(side=LEFT, padx=6)

        frm_progress = ttk.Frame(frame)
        frm_progress.pack(fill="x", padx=8, pady=6)
        ttk.Progressbar(frm_progress, variable=self.progress_extract_var, maximum=100).pack(fill="x", padx=4)
        ttk.Label(frm_progress, textvariable=self.label_extract_status).pack(anchor="w", padx=4)

    def _select_extract_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.extract_pdf_path.set(path)

    def _build_pdf2odt_frame(self, frame):
        frm_top = ttk.LabelFrame(frame, text="Convertir PDF a ODT (pdfminer.six + odfpy)", padding=10)
        frm_top.pack(fill="x", padx=8, pady=8)

        ttk.Label(frm_top, text="Archivo PDF:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.pdf2odt_path = StringVar()
        ttk.Entry(frm_top, textvariable=self.pdf2odt_path, width=60).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(frm_top, text="Seleccionar...", bootstyle="secondary", command=self._select_pdf2odt).grid(row=0, column=2, padx=4, pady=4)

        self.pdf2odt_preview = BooleanVar(value=True)
        ttk.Checkbutton(frm_top, text="Abrir ODT al terminar", variable=self.pdf2odt_preview).grid(row=1, column=1, sticky="w", padx=4, pady=2)

        frm_actions = ttk.Frame(frame)
        frm_actions.pack(fill="x", padx=8, pady=4)
        self.btn_pdf2odt_start = ttk.Button(frm_actions, text="Convertir a ODT", bootstyle="primary", command=self._start_pdf2odt)
        self.btn_pdf2odt_start.pack(side=LEFT, padx=6)
        self.btn_pdf2odt_stop = ttk.Button(frm_actions, text="Detener", bootstyle="danger-outline", command=self._stop_worker, state="disabled")
        self.btn_pdf2odt_stop.pack(side=LEFT, padx=6)

        frm_progress = ttk.Frame(frame)
        frm_progress.pack(fill="x", padx=8, pady=6)
        ttk.Progressbar(frm_progress, variable=self.progress_pdf2odt_var, maximum=100).pack(fill="x", padx=4)
        ttk.Label(frm_progress, textvariable=self.label_pdf2odt_status).pack(anchor="w", padx=4)

    def _select_pdf2odt(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf2odt_path.set(path)

    # ---------------------------
    # Buttons state & worker control
    # ---------------------------
    def _set_buttons_state(self, started: bool):
        for b in [getattr(self, 'btn_search_start', None), getattr(self, 'btn_number_start', None), getattr(self, 'btn_merge_start', None), getattr(self, 'btn_extract_start', None), getattr(self, 'btn_pdf2odt_start', None)]:
            if b:
                try:
                    b.config(state="disabled" if started else "normal")
                except Exception:
                    pass
        for b in [getattr(self, 'btn_search_stop', None), getattr(self, 'btn_number_stop', None), getattr(self, 'btn_merge_stop', None), getattr(self, 'btn_extract_stop', None), getattr(self, 'btn_pdf2odt_stop', None)]:
            if b:
                try:
                    b.config(state="normal" if started else "disabled")
                except Exception:
                    pass
        if started:
            self.progress_var.set(0)
            self.status_var.set("Procesando...")
        else:
            self.progress_var.set(0)
            self.status_var.set("Listo ✅")

    def _stop_worker(self):
        if self.active_worker and self.active_worker.is_alive():
            self.active_worker.stop()
            self.status_var.set("🟥 Deteniendo proceso...")
            # posten logs for user
            self.queue.put(("log", "system", "🟥 Proceso detenido por el usuario."))
            # worker will post ("done", None) when finishes

    # ---------------------------
    # Queue processor (centralizado)
    # ---------------------------
    def _process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    tab, value, current, total = msg[1], msg[2], msg[3], msg[4]
                    if tab == "search":
                        self.progress_search_var.set(value)
                        self.label_search_status.set(f"{current} / {total} páginas")
                    elif tab == "number":
                        self.progress_number_var.set(value)
                        self.label_number_status.set(f"{current} / {total} páginas")
                    elif tab == "merge":
                        self.progress_merge_var.set(value)
                        self.label_merge_status.set(f"{current} / {total} archivos")
                    elif tab == "extract":
                        self.progress_extract_var.set(value)
                        self.label_extract_status.set(f"{current} / {total} páginas")
                    elif tab == "pdf2odt":
                        self.progress_pdf2odt_var.set(value)
                        self.label_pdf2odt_status.set(f"{current} / {total}")
                    # global progress average
                    try:
                        vals = []
                        for v in (self.progress_search_var.get(), self.progress_number_var.get(), self.progress_merge_var.get(), self.progress_extract_var.get(), self.progress_pdf2odt_var.get()):
                            if v > 0:
                                vals.append(v)
                        if vals:
                            self.progress_var.set(sum(vals)//len(vals))
                    except Exception:
                        pass
                elif kind == "log":
                    tipo, texto = msg[1], msg[2]
                    # write to file via logging
                    write_log_file(f"[{tipo}] {texto}")
                    # insert to shared log view with timestamp, newest on top
                    now = datetime.now().strftime("%H:%M:%S")
                    line = f"[{now}] [{tipo}] {texto}"
                    try:
                        self.shared_log.configure(state="normal")
                    except Exception:
                        try:
                            self.shared_log.config(state="normal")
                        except Exception:
                            pass
                    # insert at beginning so newest appear first
                    try:
                        self.shared_log.insert("1.0", line + "\n")
                        self.shared_log.see("1.0")
                        try:
                            self.shared_log.configure(state="disabled")
                        except Exception:
                            try:
                                self.shared_log.config(state="disabled")
                            except Exception:
                                pass
                    except Exception:
                        # fallback: append at end
                        try:
                            self.shared_log.insert("end", line + "\n")
                            self.shared_log.see("end")
                        except Exception:
                            pass
                elif kind == "done":
                    self._set_buttons_state(False)
                    self.status_var.set("Listo ✅")
                    try:
                        if ToastNotification:
                            ToastNotification(title="Operación", message="Finalizado", duration=1500, alert=True).show_toast()
                    except Exception:
                        pass
                    self.active_worker = None
                elif kind == "error":
                    messagebox.showerror("Error", msg[1])
                    write_log_file("Error: " + msg[1], level="error")
                    self._set_buttons_state(False)
                    self.active_worker = None
        except queue.Empty:
            pass
        self.root.after(150, self._process_queue)

    # ---------------------------
    # Launchers
    # ---------------------------
    def _start_search(self):
        path = self.selected_pdf.get().strip()
        texts = [self.search_text1.get().strip(), self.search_text2.get().strip(), self.search_text3.get().strip()]
        texts = [t for t in texts if t]
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Selecciona un PDF válido.")
            return
        if not texts:
            messagebox.showerror("Error", "Introduce al menos un texto a buscar.")
            return
        # clear only the shared view (not the file)
        self._clear_log_view()
        self._set_buttons_state(True)
        args = (path, texts, self.case_sensitive.get(), self.require_all.get(), self.search_preview.get(), self.queue)
        self.active_worker = Worker(target=self._task_search_and_extract, args=args)
        self.active_worker.start()
        try:
            self.btn_search_stop.config(state="normal")
        except Exception:
            pass

    def _start_numbering(self):
        path = self.number_pdf_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Selecciona un PDF válido.")
            return
        self._clear_log_view()
        self._set_buttons_state(True)
        args = (
            path,
            self.number_start.get(),
            self.page_start.get(),
            self.use_number_initial.get(),
            self.coord_x.get(),
            self.coord_y.get(),
            self.prefix_text.get(),
            self.number_preview.get(),
            self.queue
        )
        self.active_worker = Worker(target=self._task_add_page_numbers, args=args)
        self.active_worker.start()
        try:
            self.btn_number_stop.config(state="normal")
        except Exception:
            pass

    def _start_merge(self):
        files = [self.listbox_files.get(i) for i in range(self.listbox_files.size())]
        if not files:
            messagebox.showerror("Error", "Añade archivos a la lista.")
            return
        output_name = self.merge_output.get().strip()
        if not output_name:
            messagebox.showerror("Error", "Introduce nombre de salida.")
            return
        if not os.path.isabs(output_name):
            output_name = os.path.join(self.result_folder, output_name)
        output_name = unique_filename(output_name)
        self._clear_log_view()
        self._set_buttons_state(True)
        self.active_worker = Worker(target=self._task_merge_pdfs, args=(files, output_name, self.merge_preview.get(), self.queue))
        self.active_worker.start()
        try:
            self.btn_merge_stop.config(state="normal")
        except Exception:
            pass

    def _start_extract(self):
        path = self.extract_pdf_path.get().strip()
        ranges = self.page_ranges.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Selecciona un PDF válido.")
            return
        if not ranges:
            messagebox.showerror("Error", "Introduce rangos de páginas.")
            return
        self._clear_log_view()
        self._set_buttons_state(True)
        args = (path, ranges, self.extract_preview.get(), self.queue)
        self.active_worker = Worker(target=self._task_extract_pages, args=args)
        self.active_worker.start()
        try:
            self.btn_extract_stop.config(state="normal")
        except Exception:
            pass

    def _start_pdf2odt(self):
        path = self.pdf2odt_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "Selecciona un PDF válido.")
            return
        self._clear_log_view()
        self._set_buttons_state(True)
        args = (path, self.pdf2odt_preview.get(), self.queue)
        self.active_worker = Worker(target=self._task_pdf_to_odt, args=args)
        self.active_worker.start()
        try:
            self.btn_pdf2odt_stop.config(state="normal")
        except Exception:
            pass

    # ---------------------------
    # Tasks (fitz + pdfminer + odfpy)
    # ---------------------------
    def _task_search_and_extract(self, stop_event: threading.Event, pdf_path: str, texts: list, case_sensitive: bool, require_all: bool, preview: bool, q: queue.Queue):
        try:
            q.put(("log", "search", f"Iniciando búsqueda en: {os.path.basename(pdf_path)}"))
            doc = fitz.open(pdf_path)
            result_doc = fitz.open()
            total = doc.page_count
            texts_comp = texts if case_sensitive else [t.lower() for t in texts]
            matched_count = 0

            for i in range(total):
                if stop_event.is_set():
                    q.put(("log", "search", "🟥 Proceso detenido por el usuario."))
                    q.put(("done", None))
                    doc.close(); result_doc.close()
                    return
                page = doc.load_page(i)
                page_text = page.get_text("text") or ""
                search_text = page_text if case_sensitive else page_text.lower()
                found = [t in search_text for t in texts_comp]
                if (require_all and all(found)) or (not require_all and any(found)):
                    result_doc.insert_pdf(doc, from_page=i, to_page=i)
                    matched_count += 1
                    q.put(("log", "search", f"Coincidencia en página {i+1}"))
                q.put(("progress", "search", int((i+1)/total*100), i+1, total))

            if matched_count > 0:
                out_name = os.path.splitext(os.path.basename(pdf_path))[0] + "_resultado.pdf"
                out_path = unique_filename(os.path.join(self.result_folder, out_name))
                result_doc.save(out_path)
                q.put(("log", "search", f"✅ Archivo generado: {out_path}"))
                if preview:
                    open_file_with_default_app(out_path)
            else:
                q.put(("log", "search", "⚠️ No se encontraron coincidencias."))
            doc.close(); result_doc.close()
            q.put(("done", None))
        except Exception as e:
            q.put(("error", str(e)))
            write_log_file(f"Error search task: {e}", level="error")

    def _task_add_page_numbers(self, stop_event: threading.Event, pdf_path: str, start_num: int, page_start: int, use_number_initial: bool, x: int, y_from_bottom: int, prefix: str, preview: bool, q: queue.Queue):
        try:
            q.put(("log", "number", f"Numerando archivo: {os.path.basename(pdf_path)}"))
            doc = fitz.open(pdf_path)
            total = doc.page_count
            for i in range(total):
                if stop_event.is_set():
                    q.put(("log", "number", "🟥 Proceso detenido por el usuario."))
                    q.put(("done", None))
                    doc.close()
                    return
                page = doc.load_page(i)
                if (i + 1) >= page_start:
                    if use_number_initial:
                        page_number = start_num + (i - page_start + 1)
                    else:
                        page_number = i + 1
                else:
                    page_number = None
                if page_number is not None:
                    rect = page.rect
                    y_pos = rect.height - y_from_bottom
                    text = f"{prefix}{page_number}"
                    font_size = 12
                    page.insert_text((x, y_pos), text, fontsize=font_size, fontname="helv", fill=(0, 0, 0))
                q.put(("progress", "number", int((i+1)/total*100), i+1, total))
            out_name = os.path.splitext(os.path.basename(pdf_path))[0] + "_numerado.pdf"
            out_path = unique_filename(os.path.join(self.result_folder, out_name))
            doc.save(out_path)
            q.put(("log", "number", f"✅ Archivo generado: {out_path}"))
            if preview:
                open_file_with_default_app(out_path)
            doc.close()
            q.put(("done", None))
        except Exception as e:
            q.put(("error", str(e)))
            write_log_file(f"Error numbering task: {e}", level="error")

    def _task_merge_pdfs(self, stop_event: threading.Event, files: list, output_name: str, preview: bool, q: queue.Queue):
        try:
            q.put(("log", "merge", f"Unión de {len(files)} archivos en: {output_name}"))
            if not output_name.lower().endswith(".pdf"):
                output_name = output_name + "_unido.pdf"
            else:
                base = os.path.splitext(output_name)[0]
                output_name = base + "_unido.pdf"
            out_path = unique_filename(output_name)
            newdoc = fitz.open()
            total = len(files)
            for i, fpath in enumerate(files):
                if stop_event.is_set():
                    q.put(("log", "merge", "🟥 Proceso detenido por el usuario."))
                    q.put(("done", None))
                    newdoc.close()
                    return
                try:
                    src = fitz.open(fpath)
                    newdoc.insert_pdf(src)
                    src.close()
                    q.put(("log", "merge", f"Añadido: {os.path.basename(fpath)}"))
                except Exception as e:
                    q.put(("log", "merge", f"⚠️ Error añadiendo {os.path.basename(fpath)}: {e}"))
                    write_log_file(f"Error merging {fpath}: {e}", level="error")
                q.put(("progress", "merge", int((i+1)/total*100), i+1, total))
            newdoc.save(out_path)
            newdoc.close()
            q.put(("log", "merge", f"✅ Archivo generado: {out_path}"))
            if preview:
                open_file_with_default_app(out_path)
            q.put(("done", None))
        except Exception as e:
            q.put(("error", str(e)))
            write_log_file(f"Error merge task: {e}", level="error")

    def _task_extract_pages(self, stop_event: threading.Event, pdf_path: str, ranges_str: str, preview: bool, q: queue.Queue):
        try:
            q.put(("log", "extract", f"Extrayendo páginas de: {os.path.basename(pdf_path)}"))
            doc = fitz.open(pdf_path)
            total_pages = doc.page_count

            pages_to_extract = []
            for part in ranges_str.split(','):
                part = part.strip()
                if not part:
                    continue
                if '-' in part:
                    try:
                        start, end = part.split('-', 1)
                        start_i = int(start) - 1
                        end_i = int(end) - 1
                        if start_i <= end_i:
                            pages_to_extract.extend(range(start_i, end_i+1))
                    except Exception:
                        continue
                else:
                    try:
                        p = int(part) - 1
                        pages_to_extract.append(p)
                    except Exception:
                        continue

            pages_to_extract = sorted(set([p for p in pages_to_extract if 0 <= p < total_pages]))
            if not pages_to_extract:
                q.put(("log", "extract", "⚠️ No hay páginas válidas en la selección."))
                q.put(("done", None))
                doc.close()
                return

            newdoc = fitz.open()
            for idx, p in enumerate(pages_to_extract):
                if stop_event.is_set():
                    q.put(("log", "extract", "🟥 Proceso detenido por el usuario."))
                    q.put(("done", None))
                    doc.close(); newdoc.close()
                    return
                newdoc.insert_pdf(doc, from_page=p, to_page=p)
                q.put(("progress", "extract", int((idx+1)/len(pages_to_extract)*100), idx+1, len(pages_to_extract)))
                q.put(("log", "extract", f"Página {p+1} añadida"))

            out_name = os.path.splitext(os.path.basename(pdf_path))[0] + "_extraido.pdf"
            out_path = unique_filename(os.path.join(self.result_folder, out_name))
            newdoc.save(out_path)
            newdoc.close(); doc.close()
            q.put(("log", "extract", f"✅ Archivo generado: {out_path}"))
            if preview:
                open_file_with_default_app(out_path)
            q.put(("done", None))
        except Exception as e:
            q.put(("error", str(e)))
            write_log_file(f"Error extract task: {e}", level="error")

    def _task_pdf_to_odt(self, stop_event: threading.Event, pdf_path: str, preview: bool, q: queue.Queue):
        try:
            q.put(("log", "pdf2odt", f"Iniciando conversión: {os.path.basename(pdf_path)}"))
            total_pages_hint = 1
            try:
                doc = fitz.open(pdf_path)
                total_pages_hint = max(1, doc.page_count)
                doc.close()
            except Exception:
                total_pages_hint = 1

            texto = extract_text(pdf_path) or ""
            lines = texto.splitlines()
            q.put(("progress", "pdf2odt", 10, 1, total_pages_hint))

            odt = OpenDocumentText()
            total = max(1, len(lines))
            for idx, linea in enumerate(lines):
                if stop_event.is_set():
                    q.put(("log", "pdf2odt", "🟥 Proceso detenido por el usuario."))
                    q.put(("done", None))
                    return
                p = P()
                p.addElement(Span(text=linea))
                odt.text.addElement(p)
                progress_val = 10 + int((idx+1)/total*80)
                q.put(("progress", "pdf2odt", progress_val, idx+1, total))

            base = os.path.splitext(os.path.basename(pdf_path))[0]
            out_name = f"{base}.odt"
            out_path = unique_filename(os.path.join(self.result_folder, out_name))
            odt.save(out_path)
            q.put(("log", "pdf2odt", f"✅ Archivo generado: {out_path}"))
            q.put(("progress", "pdf2odt", 100, total, total))
            if preview:
                open_file_with_default_app(out_path)
            q.put(("done", None))
        except Exception as e:
            q.put(("error", str(e)))
            write_log_file(f"Error pdf2odt task: {e}", level="error")

    # ---------------------------
    # Misc UI helpers (logs)
    # ---------------------------
    def _clear_log_view(self):
        try:
            self.shared_log.configure(state="normal")
        except Exception:
            try:
                self.shared_log.config(state="normal")
            except Exception:
                pass
        try:
            self.shared_log.delete("1.0", "end")
        except Exception:
            pass
        try:
            self.shared_log.configure(state="disabled")
        except Exception:
            try:
                self.shared_log.config(state="disabled")
            except Exception:
                pass

    def _copy_log(self):
        try:
            self.shared_log.configure(state="normal")
        except Exception:
            try:
                self.shared_log.config(state="normal")
            except Exception:
                pass
        try:
            text = self.shared_log.get("1.0", "end").strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception:
            pass
        try:
            self.shared_log.configure(state="disabled")
        except Exception:
            try:
                self.shared_log.config(state="disabled")
            except Exception:
                pass

    def _load_today_log_into_view(self):
        # read CURRENT_LOG_PATH and load into shared_log (most recent on top)
        try:
            if not CURRENT_LOG_PATH:
                return
            p = Path(CURRENT_LOG_PATH)
            if not p.exists():
                return
            with p.open("r", encoding="utf-8") as f:
                lines = f.read().splitlines()
            # show newest first
            try:
                self.shared_log.configure(state="normal")
            except Exception:
                try:
                    self.shared_log.config(state="normal")
                except Exception:
                    pass
            for line in reversed(lines):
                try:
                    self.shared_log.insert("1.0", line + "\n")
                except Exception:
                    try:
                        self.shared_log.insert("end", line + "\n")
                    except Exception:
                        pass
            try:
                self.shared_log.configure(state="disabled")
            except Exception:
                try:
                    self.shared_log.config(state="disabled")
                except Exception:
                    pass
        except Exception:
            pass

    def abrir_carpeta_resultados(self):
        carpeta_resultados = self.result_folder or config.get("result_folder", DEFAULT_CONFIG["result_folder"])
        if os.path.exists(carpeta_resultados):
            open_file_with_default_app(carpeta_resultados)
        else:
            messagebox.showwarning("Resultados", f"La carpeta de resultados no existe:\n{carpeta_resultados}")

    # ---------------------------
    # Close handler
    # ---------------------------
    def _on_close(self):
        if self.active_worker and self.active_worker.is_alive():
            if messagebox.askyesno("Salir", "Hay una operación en curso. ¿Deseas detenerla y salir?"):
                self.active_worker.stop()
                self.root.after(500, self.root.destroy)
            else:
                return
        else:
            self.root.destroy()

# ---------------------------
# Run
# ---------------------------
if __name__ == "__main__":
    app = ttk.Window(themename=config.get("theme", DEFAULT_CONFIG["theme"]))
    DFToolkitModern(app)
    app.mainloop()
