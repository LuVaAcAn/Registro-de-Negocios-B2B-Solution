#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Registro de Negocios — Aplicación de escritorio (v2, diseño moderno)
---------------------------------------------------------------------
Guarda toda la base de datos en un archivo local (negocios.db, SQLite)
ubicado en la misma carpeta que este programa. Los datos NO se pierden
al cerrar el programa: quedan escritos en ese archivo y se vuelven a
cargar automáticamente la próxima vez que lo abras.

Novedades de esta versión:
    - Diseño blanco / moderno (paleta clara, acentos en azul).
    - Panel redimensionable (arrastra el divisor central) y ventana
      con soporte completo de resize.
    - Formulario con scroll (rueda del mouse) para que nada quede
      oculto sin importar el tamaño de la ventana.
    - Atajos de teclado (ver barra inferior de la app).
    - Optimizaciones: SQLite en modo WAL, índices, refresco parcial
      de la interfaz (ya no se reconstruye todo en cada cambio).

Requisitos:
    - Python 3.8 o superior (tkinter y sqlite3 vienen incluidos)
    - openpyxl  ->  pip install openpyxl   (solo para exportar a Excel)

Ejecutar:
    python app.py   (o "python3 app.py" según tu sistema)
"""

import os
import sys
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ---------------------------------------------------------------------------
# Rutas y base de datos
# ---------------------------------------------------------------------------

APP_DIR = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "negocios.db")

ESTADOS = ["Activo", "Potencial", "Inactivo"]

COLUMNS = [
    ("code", "Código", 90),
    ("negocio", "Negocio", 190),
    ("contacto", "Contacto", 130),
    ("telefono", "Teléfono", 120),
    ("correo", "Correo", 190),
    ("direccion", "Dirección", 190),
    ("ciudad", "Ciudad", 100),
    ("estado", "Estado", 90),
    ("fecha", "Fecha", 90),
]
GENERAL_COLUMNS = [("category", "Rubro", 130)] + COLUMNS


def today_str():
    return datetime.date.today().strftime("%d/%m/%Y")


def sanitize_prefix(text):
    return "".join(ch for ch in text.upper() if ch.isalnum())[:8]


def pad_number(n):
    return f"{n:02d}" if n < 100 else str(n)


# ---------------------------------------------------------------------------
# Capa de datos (SQLite) — optimizada
# ---------------------------------------------------------------------------

class Database:
    def __init__(self, path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        # ---- Optimización de velocidad ----
        # WAL permite lecturas/escrituras concurrentes más rápidas y evita
        # bloqueos largos; synchronous=NORMAL reduce fsyncs innecesarios
        # sin arriesgar la integridad de los datos en uso normal de escritorio.
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.conn.execute("PRAGMA synchronous = NORMAL")
        self.conn.execute("PRAGMA temp_store = MEMORY")
        self.conn.execute("PRAGMA cache_size = -8000")  # ~8MB de caché
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._create_tables()
        self._seed_if_empty()

    def _create_tables(self):
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                prefix TEXT NOT NULL,
                counter INTEGER NOT NULL DEFAULT 1
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS businesses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT UNIQUE NOT NULL,
                category_id INTEGER NOT NULL,
                negocio TEXT NOT NULL,
                contacto TEXT,
                telefono TEXT NOT NULL,
                correo TEXT NOT NULL,
                direccion TEXT NOT NULL,
                ciudad TEXT,
                estado TEXT NOT NULL DEFAULT 'Activo',
                notas TEXT,
                fecha TEXT,
                FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
            )
        """)
        # Índices para acelerar los JOIN y filtros por categoría/código.
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_businesses_category ON businesses(category_id)")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_businesses_code ON businesses(code)")
        self.conn.commit()

    def _seed_if_empty(self):
        cur = self.conn.execute("SELECT COUNT(*) FROM categories")
        if cur.fetchone()[0] == 0:
            self.conn.execute(
                "INSERT INTO categories (name, prefix, counter) VALUES (?,?,?)",
                ("Cafetería", "TSCA", 2),
            )
            cat_id = self.conn.execute(
                "SELECT id FROM categories WHERE name = ?", ("Cafetería",)
            ).fetchone()[0]
            self.conn.execute("""
                INSERT INTO businesses
                (code, category_id, negocio, contacto, telefono, correo, direccion, ciudad, estado, notas, fecha)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, ("TSCA01", cat_id, "The Standard Cafetería", "María Gómez", "+51 987 654 321",
                  "maria.gomez@thestandard.com", "Av. Los Álamos 123, San Isidro", "Lima",
                  "Activo", "Registro de ejemplo — puedes eliminarlo.", today_str()))
            self.conn.commit()

    # ---- categorías ----
    def get_categories(self):
        rows = self.conn.execute("SELECT id, name, prefix, counter FROM categories ORDER BY name").fetchall()
        return {name: {"id": cid, "prefix": prefix, "counter": counter} for cid, name, prefix, counter in rows}

    def create_category(self, name, prefix):
        cur = self.conn.execute(
            "INSERT INTO categories (name, prefix, counter) VALUES (?,?,1)", (name, prefix)
        )
        self.conn.commit()
        return cur.lastrowid

    def delete_category(self, cat_id):
        self.conn.execute("DELETE FROM businesses WHERE category_id = ?", (cat_id,))
        self.conn.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
        self.conn.commit()

    def bump_counter(self, cat_id, new_counter):
        self.conn.execute("UPDATE categories SET counter = ? WHERE id = ?", (new_counter, cat_id))
        self.conn.commit()

    # ---- negocios ----
    def get_businesses(self):
        return self.conn.execute("""
            SELECT b.code, c.name as category, b.negocio, b.contacto, b.telefono, b.correo,
                   b.direccion, b.ciudad, b.estado, b.notas, b.fecha
            FROM businesses b JOIN categories c ON c.id = b.category_id
            ORDER BY c.name, b.code
        """).fetchall()

    def get_businesses_grouped(self):
        """Devuelve (lista_completa, dict_por_rubro) en una sola pasada,
        en vez de filtrar la lista completa una vez por cada pestaña."""
        rows = self.get_businesses()
        grouped = {}
        for row in rows:
            grouped.setdefault(row["category"], []).append(row)
        return rows, grouped

    def insert_business(self, code, category_id, data):
        self.conn.execute("""
            INSERT INTO businesses
            (code, category_id, negocio, contacto, telefono, correo, direccion, ciudad, estado, notas, fecha)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, (code, category_id, data["negocio"], data["contacto"], data["telefono"], data["correo"],
              data["direccion"], data["ciudad"], data["estado"], data["notas"], today_str()))
        self.conn.commit()

    def update_business(self, code, data):
        self.conn.execute("""
            UPDATE businesses SET negocio=?, contacto=?, telefono=?, correo=?, direccion=?, ciudad=?, estado=?, notas=?
            WHERE code = ?
        """, (data["negocio"], data["contacto"], data["telefono"], data["correo"],
              data["direccion"], data["ciudad"], data["estado"], data["notas"], code))
        self.conn.commit()

    def delete_business(self, code):
        self.conn.execute("DELETE FROM businesses WHERE code = ?", (code,))
        self.conn.commit()

    def get_business(self, code):
        row = self.conn.execute("""
            SELECT b.code, c.name as category, b.negocio, b.contacto, b.telefono, b.correo,
                   b.direccion, b.ciudad, b.estado, b.notas, b.fecha
            FROM businesses b JOIN categories c ON c.id = b.category_id
            WHERE b.code = ?
        """, (code,)).fetchone()
        return row


# ---------------------------------------------------------------------------
# Paleta — diseño blanco / moderno
# ---------------------------------------------------------------------------

BG          = "#FFFFFF"   # fondo general
SURFACE     = "#FFFFFF"   # tarjetas
SURFACE_ALT = "#F7F8FA"   # paneles secundarios / fondo de ventana
BORDER      = "#E4E7EC"
TEXT        = "#101828"
TEXT_SOFT   = "#667085"
TEXT_FAINT  = "#98A2B3"

ACCENT       = "#2563EB"  # azul principal (misma familia que SentiMYPE)
ACCENT_HOVER = "#1D4ED8"
ACCENT_SOFT  = "#EFF4FF"

SUCCESS      = "#12805C"
SUCCESS_SOFT = "#E7F6EF"
DANGER       = "#B42318"
DANGER_SOFT  = "#FEF3F2"
WARN_SOFT    = "#FFFAEB"
WARN_TEXT    = "#B54708"

FONT_BASE   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_LABEL  = ("Segoe UI", 9, "bold")
FONT_TITLE  = ("Segoe UI", 18, "bold")
FONT_SUB    = ("Segoe UI", 12, "bold")
FONT_MONO   = ("Consolas", 13, "bold")


# ---------------------------------------------------------------------------
# Widgets reutilizables
# ---------------------------------------------------------------------------

class ModernButton(tk.Button):
    """Botón plano con estado hover, pensado para el look moderno blanco."""

    def __init__(self, master, bg=ACCENT, fg="white", hover=None, **kwargs):
        self._bg = bg
        self._hover = hover or self._darken(bg)
        self._fg = fg
        kwargs.setdefault("highlightthickness", 0)
        super().__init__(
            master, bg=bg, fg=fg, activebackground=self._hover, activeforeground=fg,
            relief="flat", bd=0, cursor="hand2", font=kwargs.pop("font", FONT_LABEL),
            padx=kwargs.pop("padx", 14), pady=kwargs.pop("pady", 8),
            **kwargs,
        )
        self.bind("<Enter>", lambda e: self.configure(bg=self._hover))
        self.bind("<Leave>", lambda e: self.configure(bg=self._bg))

    @staticmethod
    def _darken(hex_color, factor=0.88):
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return hex_color
        r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
        r, g, b = (max(0, int(c * factor)) for c in (r, g, b))
        return f"#{r:02x}{g:02x}{b:02x}"


class ScrollableFrame(tk.Frame):
    """Frame con scroll vertical (rueda del mouse incluida) — usado para que
    el formulario nunca oculte campos sin importar el tamaño de la ventana."""

    def __init__(self, master, bg=SURFACE, **kwargs):
        super().__init__(master, bg=bg, **kwargs)

        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        for widget in (self.canvas, self.inner):
            widget.bind("<Enter>", lambda e: self._bind_mousewheel())
            widget.bind("<Leave>", lambda e: self._unbind_mousewheel())

    def _on_inner_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # El contenido interno siempre ocupa el ancho disponible -> se adapta
        # correctamente cuando el usuario redimensiona la ventana o el panel.
        self.canvas.itemconfigure(self._window, width=event.width)

    def _bind_mousewheel(self):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)      # Windows / macOS
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)  # Linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _unbind_mousewheel(self):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        self.canvas.yview_scroll(-1 if event.num == 4 else 1, "units")


class Card(tk.Frame):
    """Contenedor con look de 'tarjeta' blanca moderna (borde suave, sin sombra dura)."""

    def __init__(self, master, **kwargs):
        super().__init__(master, bg=SURFACE, highlightbackground=BORDER,
                          highlightthickness=1, bd=0, **kwargs)


def styled_entry(parent, **kwargs):
    e = tk.Entry(parent, relief="flat", bd=0, font=FONT_BASE, bg=SURFACE_ALT, fg=TEXT,
                 insertbackground=TEXT, highlightthickness=1,
                 highlightbackground=BORDER, highlightcolor=ACCENT, **kwargs)

    def on_focus_in(_e):
        e.configure(highlightbackground=ACCENT, highlightcolor=ACCENT, bg=SURFACE)

    def on_focus_out(_e):
        e.configure(highlightbackground=BORDER, bg=SURFACE_ALT)

    e.bind("<FocusIn>", on_focus_in)
    e.bind("<FocusOut>", on_focus_out)
    return e


# ---------------------------------------------------------------------------
# Interfaz gráfica
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Registro de Negocios — Base de datos local")
        self.geometry("1220x760")
        self.configure(bg=SURFACE_ALT)
        self.minsize(880, 560)

        self.db = Database(DB_PATH)
        self.categories = {}
        self.tab_widgets = {}       # nombre_rubro -> {"frame", "tree"}
        self.field_order = []       # orden de tabulación entre campos del form

        self._build_style()
        self._build_layout()
        self._build_shortcuts()
        self.refresh_all(rebuild_tabs=True)

        # Enfocar el primer campo al abrir, listo para escribir de inmediato.
        self.after(80, lambda: self.rubro_combo.focus_set())

    # ================================================================
    # Estilo ttk
    # ================================================================
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TPanedwindow", background=SURFACE_ALT)
        style.configure("TNotebook", background=SURFACE_ALT, borderwidth=0, tabmargins=(0, 6, 0, 0))
        style.configure("TNotebook.Tab", padding=(14, 9), font=FONT_SMALL,
                         background=SURFACE_ALT, foreground=TEXT_SOFT, borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", SURFACE)],
                  foreground=[("selected", ACCENT)])

        style.configure("Treeview", rowheight=28, font=FONT_SMALL, background=SURFACE,
                         fieldbackground=SURFACE, foreground=TEXT, borderwidth=0)
        style.configure("Treeview.Heading", font=FONT_LABEL, background=SURFACE_ALT,
                         foreground=TEXT_SOFT, relief="flat", borderwidth=0)
        style.map("Treeview",
                  background=[("selected", ACCENT_SOFT)],
                  foreground=[("selected", ACCENT)])
        style.map("Treeview.Heading", background=[("active", SURFACE_ALT)])

        style.configure("TCombobox", fieldbackground=SURFACE_ALT, background=SURFACE,
                         foreground=TEXT, arrowcolor=TEXT_SOFT, borderwidth=0)
        style.configure("Vertical.TScrollbar", background=SURFACE_ALT, troughcolor=SURFACE,
                         bordercolor=SURFACE, arrowcolor=TEXT_FAINT)
        style.configure("Horizontal.TScrollbar", background=SURFACE_ALT, troughcolor=SURFACE,
                         bordercolor=SURFACE, arrowcolor=TEXT_FAINT)
        style.configure("TPanedwindow", background=SURFACE_ALT)

    # ================================================================
    # Layout general — con panel redimensionable
    # ================================================================
    def _build_layout(self):
        header = tk.Frame(self, bg=SURFACE_ALT)
        header.pack(fill="x", padx=20, pady=(16, 4))

        title_box = tk.Frame(header, bg=SURFACE_ALT)
        title_box.pack(side="left", anchor="w")
        tk.Label(title_box, text="Registro de Negocios", font=FONT_TITLE, bg=SURFACE_ALT, fg=TEXT)\
            .pack(anchor="w")
        tk.Label(title_box, text=f"Base de datos local: {DB_PATH}", font=("Segoe UI", 8),
                 bg=SURFACE_ALT, fg=TEXT_FAINT).pack(anchor="w")

        self.stats_var = tk.StringVar()
        tk.Label(header, textvariable=self.stats_var, font=FONT_SMALL, bg=SURFACE_ALT,
                 fg=TEXT_SOFT).pack(side="right", anchor="e")

        # ---- Panel central redimensionable (arrastra el divisor) ----
        self.paned = ttk.PanedWindow(self, orient="horizontal")
        self.paned.pack(fill="both", expand=True, padx=20, pady=(10, 8))

        form_holder = tk.Frame(self.paned, bg=SURFACE_ALT)
        dir_holder = tk.Frame(self.paned, bg=SURFACE_ALT)
        self.paned.add(form_holder, weight=0)
        self.paned.add(dir_holder, weight=1)

        self._build_form(form_holder)
        self._build_directory(dir_holder)

        # ---- Barra inferior de atajos de teclado ----
        footer = tk.Frame(self, bg=SURFACE_ALT)
        footer.pack(fill="x", padx=20, pady=(0, 12))
        shortcuts_text = (
            "⌨  Ctrl+Enter registrar   ·   Ctrl+N nuevo   ·   Ctrl+E editar   ·   "
            "Supr eliminar   ·   Ctrl+S exportar   ·   F5 actualizar   ·   Esc limpiar mensaje"
        )
        tk.Label(footer, text=shortcuts_text, font=("Segoe UI", 8), bg=SURFACE_ALT, fg=TEXT_FAINT)\
            .pack(anchor="w")

    # ================================================================
    # Atajos de teclado
    # ================================================================
    def _build_shortcuts(self):
        self.bind_all("<Control-Return>", lambda e: self.submit_business())
        self.bind_all("<Control-n>", lambda e: self.focus_new_entry())
        self.bind_all("<Control-N>", lambda e: self.focus_new_entry())
        self.bind_all("<Control-e>", lambda e: self.edit_selected())
        self.bind_all("<Control-E>", lambda e: self.edit_selected())
        self.bind_all("<Control-s>", lambda e: self.export_excel())
        self.bind_all("<Control-S>", lambda e: self.export_excel())
        self.bind_all("<F5>", lambda e: self.refresh_all(rebuild_tabs=True))
        self.bind_all("<Escape>", self._on_escape)
        self.bind_all("<Delete>", self._on_delete_key)

    def _on_escape(self, event):
        widget = self.focus_get()
        # Si el foco está en un campo de texto, Escape solo limpia el mensaje.
        self.msg_var.set("")
        if isinstance(widget, (tk.Entry, tk.Text)):
            return
        self.focus_set()

    def _on_delete_key(self, event):
        widget = self.focus_get()
        # Solo eliminar si el foco está sobre la tabla, para no chocar con
        # el borrado de texto dentro de los campos del formulario.
        if isinstance(widget, ttk.Treeview):
            self.delete_selected()

    def focus_new_entry(self):
        self.rubro_combo.focus_set()

    # ================================================================
    # Formulario (izquierda) — con scroll
    # ================================================================
    def _build_form(self, parent):
        card = Card(parent)
        card.pack(fill="both", expand=True)

        header = tk.Frame(card, bg=SURFACE)
        header.pack(fill="x", padx=20, pady=(18, 6))
        tk.Label(header, text="Ficha de registro", font=FONT_SUB, bg=SURFACE, fg=TEXT).pack(anchor="w")
        tk.Label(header, text="Los campos con * son obligatorios.", font=("Segoe UI", 8),
                 bg=SURFACE, fg=TEXT_FAINT).pack(anchor="w", pady=(2, 0))

        scroll = ScrollableFrame(card, bg=SURFACE)
        scroll.pack(fill="both", expand=True)
        inner = scroll.inner
        inner.configure(padx=20, pady=4)

        # Rubro
        tk.Label(inner, text="Rubro *", font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT).pack(anchor="w", pady=(8, 2))
        self.rubro_var = tk.StringVar()
        self.rubro_combo = ttk.Combobox(inner, textvariable=self.rubro_var, state="readonly", font=FONT_BASE)
        self.rubro_combo.pack(fill="x", pady=(0, 4), ipady=3)
        self.rubro_combo.bind("<<ComboboxSelected>>", self.on_rubro_change)

        self.new_cat_frame = tk.Frame(inner, bg=WARN_SOFT, highlightbackground="#F2C94C",
                                       highlightthickness=1, bd=0)
        tk.Label(self.new_cat_frame, text="Rubro nuevo: indica el nombre y las siglas para su código\n"
                                          "(ej. Cafetería → TSCA: TS de The Standard + CA del rubro).",
                 bg=WARN_SOFT, fg=WARN_TEXT, font=("Segoe UI", 8), justify="left")\
            .pack(anchor="w", padx=10, pady=(8, 6))
        tk.Label(self.new_cat_frame, text="Nombre del rubro", bg=WARN_SOFT, fg=WARN_TEXT,
                 font=FONT_LABEL).pack(anchor="w", padx=10)
        self.new_cat_name = styled_entry(self.new_cat_frame)
        self.new_cat_name.pack(fill="x", padx=10, pady=(2, 8), ipady=3)
        tk.Label(self.new_cat_frame, text="Siglas para el código", bg=WARN_SOFT, fg=WARN_TEXT,
                 font=FONT_LABEL).pack(anchor="w", padx=10)
        self.new_cat_prefix = styled_entry(self.new_cat_frame)
        self.new_cat_prefix.pack(fill="x", padx=10, pady=(2, 10), ipady=3)
        self.new_cat_prefix.bind("<KeyRelease>", lambda e: self.update_code_preview())
        self.new_cat_name.bind("<KeyRelease>", lambda e: self.update_code_preview())

        self.fields = {}
        field_defs = [
            ("negocio", "Nombre del negocio *", True),
            ("contacto", "Nombre de contacto", False),
            ("telefono", "Número de contacto *", True),
            ("correo", "Correo electrónico *", True),
            ("direccion", "Dirección *", True),
            ("ciudad", "Ciudad", False),
        ]
        for key, label, _required in field_defs:
            tk.Label(inner, text=label, font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT).pack(anchor="w", pady=(8, 2))
            entry = styled_entry(inner)
            entry.pack(fill="x", pady=(0, 2), ipady=4)
            self.fields[key] = entry
            self.field_order.append(entry)

        tk.Label(inner, text="Estado del cliente", font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT)\
            .pack(anchor="w", pady=(8, 2))
        self.estado_var = tk.StringVar(value="Activo")
        ttk.Combobox(inner, textvariable=self.estado_var, values=ESTADOS, state="readonly", font=FONT_BASE)\
            .pack(fill="x", pady=(0, 2), ipady=3)

        tk.Label(inner, text="Notas", font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT).pack(anchor="w", pady=(8, 2))
        self.notas_text = tk.Text(inner, height=3, relief="flat", bd=0, font=FONT_SMALL,
                                   bg=SURFACE_ALT, fg=TEXT, highlightthickness=1,
                                   highlightbackground=BORDER, highlightcolor=ACCENT, wrap="word")
        self.notas_text.pack(fill="x", pady=(0, 4))

        # Enter avanza de un campo a otro; el último envía el formulario.
        self._wire_enter_navigation()

        preview_frame = tk.Frame(inner, bg=SURFACE_ALT, highlightbackground=BORDER, highlightthickness=1)
        preview_frame.pack(fill="x", pady=(14, 10))
        tk.Label(preview_frame, text="PRÓXIMO CÓDIGO", font=("Segoe UI", 7, "bold"), bg=SURFACE_ALT,
                 fg=TEXT_FAINT).pack(anchor="w", padx=12, pady=(8, 0))
        self.code_preview_var = tk.StringVar(value="—")
        tk.Label(preview_frame, textvariable=self.code_preview_var, font=FONT_MONO,
                 bg=ACCENT_SOFT, fg=ACCENT, padx=12, pady=5).pack(anchor="w", padx=12, pady=(2, 10))

        self.submit_btn = ModernButton(inner, text="Registrar negocio  (Ctrl+Enter)",
                                        bg=ACCENT, fg="white", command=self.submit_business)
        self.submit_btn.pack(fill="x", pady=(2, 8))

        self.msg_var = tk.StringVar()
        self.msg_label = tk.Label(inner, textvariable=self.msg_var, font=FONT_SMALL,
                                   bg=SURFACE, fg=SUCCESS, wraplength=300, justify="left")
        self.msg_label.pack(anchor="w", pady=(0, 16))

    def _wire_enter_navigation(self):
        ordered = self.field_order
        for i, entry in enumerate(ordered):
            nxt = ordered[i + 1] if i + 1 < len(ordered) else None
            if nxt is not None:
                entry.bind("<Return>", lambda e, n=nxt: (n.focus_set(), "break"))
            else:
                entry.bind("<Return>", lambda e: (self.submit_business(), "break"))

    # ================================================================
    # Directorio (derecha)
    # ================================================================
    def _build_directory(self, parent):
        right = tk.Frame(parent, bg=SURFACE_ALT)
        right.pack(fill="both", expand=True)

        top = tk.Frame(right, bg=SURFACE_ALT)
        top.pack(fill="x", pady=(0, 8))
        tk.Label(top, text="Directorio por rubro", font=FONT_SUB, bg=SURFACE_ALT, fg=TEXT).pack(side="left")

        ModernButton(top, text="⬇  Exportar a Excel  (Ctrl+S)", bg=SUCCESS, hover="#0E6B4C",
                     command=self.export_excel).pack(side="right")

        notebook_card = Card(right)
        notebook_card.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(notebook_card)
        self.notebook.pack(fill="both", expand=True, padx=4, pady=4)

        actions = tk.Frame(right, bg=SURFACE_ALT)
        actions.pack(fill="x", pady=(10, 0))
        ModernButton(actions, text="✏  Editar  (Ctrl+E)", bg=SURFACE, fg=TEXT, hover=SURFACE_ALT,
                     highlightbackground=BORDER, highlightthickness=1,
                     command=self.edit_selected).pack(side="left", padx=(0, 10))
        ModernButton(actions, text="🗑  Eliminar  (Supr)", bg=SURFACE, fg=DANGER, hover=DANGER_SOFT,
                     highlightbackground=BORDER, highlightthickness=1,
                     command=self.delete_selected).pack(side="left")

    # -----------------------------------------------------------------
    # Refresco de datos — optimizado (evita reconstruir todo cada vez)
    # -----------------------------------------------------------------
    def refresh_all(self, rebuild_tabs=False):
        self.categories = self.db.get_categories()
        self.refresh_rubro_combo()
        self.refresh_notebook(rebuild_tabs=rebuild_tabs or self._categories_changed())
        self.update_stats()

    def _categories_changed(self):
        return set(self.categories.keys()) != set(self.tab_widgets.keys())

    def update_stats(self):
        businesses, _grouped = self.db.get_businesses_grouped()
        self.stats_var.set(f"{len(self.categories)} rubros   ·   {len(businesses)} negocios")

    def refresh_rubro_combo(self):
        names = list(self.categories.keys())
        values = [f"{n} — {self.categories[n]['prefix']}" for n in names] + ["+ Nueva categoría (rubro)"]
        current = self.rubro_var.get()
        self.rubro_combo["values"] = values
        if not current and values:
            self.rubro_combo.current(0)
        self.on_rubro_change()

    def selected_rubro_name(self):
        val = self.rubro_var.get()
        if val == "+ Nueva categoría (rubro)":
            return None
        return val.split(" — ")[0] if val else None

    def on_rubro_change(self, event=None):
        is_new = self.rubro_var.get() == "+ Nueva categoría (rubro)"
        if is_new:
            self.new_cat_frame.pack(fill="x", pady=(0, 10), after=self.rubro_combo)
        else:
            self.new_cat_frame.forget()
        self.update_code_preview()

    def update_code_preview(self):
        if self.rubro_var.get() == "+ Nueva categoría (rubro)":
            prefix = sanitize_prefix(self.new_cat_prefix.get())
            self.code_preview_var.set(f"{prefix}01" if prefix else "— · —")
        else:
            name = self.selected_rubro_name()
            cat = self.categories.get(name)
            if cat:
                self.code_preview_var.set(cat["prefix"] + pad_number(cat["counter"]))
            else:
                self.code_preview_var.set("—")

    # -----------------------------------------------------------------
    # Notebook / pestañas por rubro + vista general
    # -----------------------------------------------------------------
    def refresh_notebook(self, rebuild_tabs=False):
        businesses, grouped = self.db.get_businesses_grouped()

        if rebuild_tabs:
            self._rebuild_tabs(grouped)
        else:
            # Optimización clave: en vez de destruir y recrear cada pestaña
            # y cada Treeview (costoso), solo se limpia y reinserta el
            # contenido de las tablas ya existentes.
            self._refresh_tab_contents(businesses, grouped)

    def _rebuild_tabs(self, grouped):
        current_tab_text = None
        try:
            current_tab_text = self.notebook.tab(self.notebook.select(), "text")
        except tk.TclError:
            pass

        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        self.tab_widgets = {}

        all_rows = [row for rows in grouped.values() for row in rows]

        general_frame = tk.Frame(self.notebook, bg=SURFACE)
        self.notebook.add(general_frame, text="📋 Vista general")
        general_tree = self._build_tree(general_frame, GENERAL_COLUMNS, all_rows)
        self.tab_widgets["__general__"] = {"frame": general_frame, "tree": general_tree}

        for name, meta in self.categories.items():
            frame = tk.Frame(self.notebook, bg=SURFACE)
            self.notebook.add(frame, text=f"{name} ({meta['prefix']})")

            toolbar = tk.Frame(frame, bg=SURFACE)
            toolbar.pack(fill="x", padx=10, pady=(10, 0))
            tk.Label(toolbar, text=f"{name} · siglas {meta['prefix']}", font=FONT_LABEL,
                     bg=SURFACE, fg=TEXT).pack(side="left")
            ModernButton(toolbar, text="🗑 Eliminar rubro", bg=DANGER_SOFT, fg=DANGER, hover="#FEE4E2",
                         font=("Segoe UI", 8, "bold"), padx=8, pady=3,
                         command=lambda n=name: self.delete_category(n)).pack(side="right")

            rows = grouped.get(name, [])
            tree = self._build_tree(frame, COLUMNS, rows)
            self.tab_widgets[name] = {"frame": frame, "tree": tree}

        if current_tab_text:
            for tab_id in self.notebook.tabs():
                if self.notebook.tab(tab_id, "text") == current_tab_text:
                    self.notebook.select(tab_id)
                    break

    def _refresh_tab_contents(self, all_rows, grouped):
        general = self.tab_widgets.get("__general__")
        if general:
            self._fill_tree(general["tree"], GENERAL_COLUMNS, all_rows)

        for name, widgets in self.tab_widgets.items():
            if name == "__general__":
                continue
            rows = grouped.get(name, [])
            self._fill_tree(widgets["tree"], COLUMNS, rows)

    def _build_tree(self, parent, columns, rows):
        wrap = tk.Frame(parent, bg=SURFACE)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        col_ids = [c[0] for c in columns]
        tree = ttk.Treeview(wrap, columns=col_ids, show="headings", selectmode="browse")
        for cid, heading, width in columns:
            tree.heading(cid, text=heading)
            tree.column(cid, width=width, anchor="w")

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        # Scroll con rueda del mouse también dentro de la tabla.
        tree.bind("<MouseWheel>", lambda e: tree.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        tree.bind("<Button-4>", lambda e: tree.yview_scroll(-1, "units"))
        tree.bind("<Button-5>", lambda e: tree.yview_scroll(1, "units"))

        self._fill_tree(tree, columns, rows, wrap=wrap)
        tree.bind("<Double-1>", lambda e: self.edit_selected())
        parent.tree = tree
        return tree

    def _fill_tree(self, tree, columns, rows, wrap=None):
        col_ids = [c[0] for c in columns]
        selected = tree.selection()
        selected_code = selected[0] if selected else None

        tree.delete(*tree.get_children())
        for row in rows:
            values = [row[cid] if row[cid] is not None else "" for cid in col_ids]
            tree.insert("", "end", iid=row["code"], values=values)

        if selected_code and tree.exists(selected_code):
            tree.selection_set(selected_code)

        # Mensaje de "sin datos" (se agrega/quita sin reconstruir la tabla).
        wrap = wrap or tree.master
        empty_label = getattr(wrap, "_empty_label", None)
        if not rows:
            if empty_label is None:
                empty_label = tk.Label(wrap, text="Todavía no hay negocios aquí. Regístralos desde la ficha de la izquierda.",
                                        bg=SURFACE, fg=TEXT_FAINT, font=FONT_SMALL)
                empty_label.grid(row=2, column=0, columnspan=2, pady=20)
                wrap._empty_label = empty_label
        elif empty_label is not None:
            empty_label.destroy()
            wrap._empty_label = None

    def get_active_tree(self):
        current = self.notebook.select()
        if not current:
            return None
        frame = self.nametowidget(current)
        return getattr(frame, "tree", None)

    def get_selected_code(self):
        tree = self.get_active_tree()
        if not tree:
            return None
        sel = tree.selection()
        return sel[0] if sel else None

    # -----------------------------------------------------------------
    # Registrar negocio
    # -----------------------------------------------------------------
    def submit_business(self):
        self.msg_var.set("")
        rubro_val = self.rubro_var.get()

        if rubro_val == "+ Nueva categoría (rubro)":
            name = self.new_cat_name.get().strip()
            prefix = sanitize_prefix(self.new_cat_prefix.get())
            if not name or not prefix:
                return self.show_msg("Indica el nombre del rubro y sus siglas antes de continuar.", error=True)
            existing = next((k for k in self.categories if k.lower() == name.lower()), None)
            if existing:
                cat_name = existing
            else:
                self.db.create_category(name, prefix)
                self.categories = self.db.get_categories()
                cat_name = name
        else:
            cat_name = self.selected_rubro_name()
            if not cat_name:
                return self.show_msg("Selecciona un rubro.", error=True)

        data = {k: e.get().strip() for k, e in self.fields.items()}
        data["notas"] = self.notas_text.get("1.0", "end").strip()

        if not data["negocio"] or not data["telefono"] or not data["correo"] or not data["direccion"]:
            return self.show_msg("Nombre del negocio, número de contacto, correo y dirección son obligatorios.", error=True)

        data["estado"] = self.estado_var.get()

        cat = self.categories[cat_name]
        code = cat["prefix"] + pad_number(cat["counter"])
        self.db.insert_business(code, cat["id"], data)
        self.db.bump_counter(cat["id"], cat["counter"] + 1)

        for e in self.fields.values():
            e.delete(0, "end")
        self.notas_text.delete("1.0", "end")
        self.estado_var.set("Activo")

        new_category = cat_name not in self.tab_widgets
        self.refresh_all(rebuild_tabs=new_category)
        self.show_msg(f"Negocio registrado con el código {code}.")
        self.fields["negocio"].focus_set()

    def show_msg(self, text, error=False):
        self.msg_var.set(text)
        self.msg_label.configure(fg=DANGER if error else SUCCESS)

    # -----------------------------------------------------------------
    # Editar / eliminar negocio
    # -----------------------------------------------------------------
    def edit_selected(self):
        code = self.get_selected_code()
        if not code:
            return messagebox.showinfo("Editar", "Selecciona primero un negocio en la tabla.")
        row = self.db.get_business(code)
        if not row:
            return
        self.open_edit_dialog(row)

    def open_edit_dialog(self, row):
        dialog = tk.Toplevel(self)
        dialog.title(f"Editar negocio — {row['code']}")
        dialog.configure(bg=SURFACE)
        dialog.geometry("420x600")
        dialog.minsize(360, 420)
        dialog.transient(self)
        dialog.grab_set()
        dialog.bind("<Escape>", lambda e: dialog.destroy())

        tk.Label(dialog, text=f"{row['code']} · {row['category']}", font=("Consolas", 10, "bold"),
                 bg=ACCENT_SOFT, fg=ACCENT).pack(fill="x", padx=16, pady=(16, 10), ipady=6)

        scroll = ScrollableFrame(dialog, bg=SURFACE)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)
        inner = scroll.inner
        inner.configure(padx=16)

        entries = {}
        ordered_entries = []
        field_defs = [
            ("negocio", "Nombre del negocio *"),
            ("contacto", "Nombre de contacto"),
            ("telefono", "Número de contacto *"),
            ("correo", "Correo electrónico *"),
            ("direccion", "Dirección *"),
            ("ciudad", "Ciudad"),
        ]
        for key, label in field_defs:
            tk.Label(inner, text=label, font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT)\
                .pack(anchor="w", pady=(10, 2))
            entry = styled_entry(inner)
            entry.insert(0, row[key] or "")
            entry.pack(fill="x", pady=2, ipady=4)
            entries[key] = entry
            ordered_entries.append(entry)

        tk.Label(inner, text="Estado del cliente", font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT)\
            .pack(anchor="w", pady=(10, 2))
        estado_var = tk.StringVar(value=row["estado"])
        ttk.Combobox(inner, textvariable=estado_var, values=ESTADOS, state="readonly", font=FONT_BASE)\
            .pack(fill="x", pady=2, ipady=3)

        tk.Label(inner, text="Notas", font=FONT_LABEL, bg=SURFACE, fg=TEXT_SOFT).pack(anchor="w", pady=(10, 2))
        notas_text = tk.Text(inner, height=4, relief="flat", bd=0, font=FONT_SMALL, bg=SURFACE_ALT,
                              fg=TEXT, highlightthickness=1, highlightbackground=BORDER,
                              highlightcolor=ACCENT, wrap="word")
        notas_text.insert("1.0", row["notas"] or "")
        notas_text.pack(fill="x", pady=(2, 16))

        def save(event=None):
            data = {k: e.get().strip() for k, e in entries.items()}
            data["estado"] = estado_var.get()
            data["notas"] = notas_text.get("1.0", "end").strip()
            if not data["negocio"] or not data["telefono"] or not data["correo"] or not data["direccion"]:
                messagebox.showerror("Datos incompletos",
                                      "Nombre del negocio, número de contacto, correo y dirección son obligatorios.",
                                      parent=dialog)
                return
            self.db.update_business(row["code"], data)
            dialog.destroy()
            self.refresh_all()
            self.show_msg(f"Datos de {data['negocio']} ({row['code']}) actualizados.")

        for i, entry in enumerate(ordered_entries):
            if i + 1 < len(ordered_entries):
                nxt = ordered_entries[i + 1]
                entry.bind("<Return>", lambda e, n=nxt: (n.focus_set(), "break"))
            else:
                entry.bind("<Return>", lambda e: (save(), "break"))
        dialog.bind("<Control-Return>", save)

        btns = tk.Frame(dialog, bg=SURFACE)
        btns.pack(fill="x", padx=16, pady=16)
        ModernButton(btns, text="Cancelar", bg=SURFACE, fg=TEXT_SOFT, hover=SURFACE_ALT,
                     highlightbackground=BORDER, highlightthickness=1,
                     command=dialog.destroy).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ModernButton(btns, text="Guardar cambios  (Ctrl+Enter)", bg=ACCENT, fg="white",
                     command=save).pack(side="left", expand=True, fill="x")

        entries["negocio"].focus_set()

    def delete_selected(self):
        code = self.get_selected_code()
        if not code:
            return messagebox.showinfo("Eliminar", "Selecciona primero un negocio en la tabla.")
        row = self.db.get_business(code)
        if not messagebox.askyesno("Eliminar negocio", f"¿Eliminar el registro de \"{row['negocio']}\" ({code})?"):
            return
        self.db.delete_business(code)
        self.refresh_all()

    # -----------------------------------------------------------------
    # Eliminar rubro
    # -----------------------------------------------------------------
    def delete_category(self, name):
        meta = self.categories.get(name)
        if not meta:
            return
        _all_rows, grouped = self.db.get_businesses_grouped()
        count = len(grouped.get(name, []))
        if count > 0:
            question = (f"El rubro \"{name}\" tiene {count} negocio(s) registrado(s). "
                        f"Si lo eliminas, también se eliminarán esos registros.\n\n¿Deseas continuar?")
        else:
            question = f"¿Eliminar el rubro \"{name}\"?"
        if not messagebox.askyesno("Eliminar rubro", question):
            return
        self.db.delete_category(meta["id"])
        self.rubro_var.set("")
        self.refresh_all(rebuild_tabs=True)

    # -----------------------------------------------------------------
    # Exportar a Excel
    # -----------------------------------------------------------------
    def export_excel(self):
        if not OPENPYXL_OK:
            messagebox.showerror(
                "Falta un paquete",
                "Para exportar a Excel necesitas instalar openpyxl.\n\n"
                "Abre una terminal y ejecuta:\n\npip install openpyxl"
            )
            return

        path = filedialog.asksaveasfilename(
            title="Guardar directorio de negocios",
            defaultextension=".xlsx",
            initialfile="Directorio_Negocios.xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
        )
        if not path:
            return

        businesses, grouped = self.db.get_businesses_grouped()
        wb = openpyxl.Workbook()
        ws_resumen = wb.active
        ws_resumen.title = "Resumen"
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")

        ws_resumen.append(["Rubro", "Siglas", "Negocios registrados", "Próximo código"])
        for c in ws_resumen[1]:
            c.font = header_font
            c.fill = header_fill
        for name, meta in self.categories.items():
            count = len(grouped.get(name, []))
            ws_resumen.append([name, meta["prefix"], count, meta["prefix"] + pad_number(meta["counter"])])
        for i, w in enumerate([22, 12, 20, 16], start=1):
            ws_resumen.column_dimensions[get_column_letter(i)].width = w

        headers = ["Código", "Nombre del Negocio", "Nombre de Contacto", "Número de Contacto",
                   "Correo Electrónico", "Dirección", "Ciudad", "Estado", "Fecha de Registro", "Notas"]
        used_names = {"Resumen"}
        for name in self.categories:
            safe = "".join(ch for ch in name if ch not in '[]*?/\\:')[:31] or "Rubro"
            final = safe
            i = 2
            while final in used_names:
                final = f"{safe[:28]}_{i}"
                i += 1
            used_names.add(final)

            ws = wb.create_sheet(final)
            ws.append(headers)
            for c in ws[1]:
                c.font = header_font
                c.fill = header_fill
                c.alignment = Alignment(horizontal="center", wrap_text=True)
            rows = grouped.get(name, [])
            for b in rows:
                ws.append([b["code"], b["negocio"], b["contacto"], b["telefono"], b["correo"],
                           b["direccion"], b["ciudad"], b["estado"], b["fecha"], b["notas"]])
            widths = [10, 26, 20, 18, 28, 30, 14, 12, 16, 30]
            for i, w in enumerate(widths, start=1):
                ws.column_dimensions[get_column_letter(i)].width = w
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:J{max(1, len(rows) + 1)}"

        wb.save(path)
        messagebox.showinfo("Exportado", f"Archivo guardado en:\n{path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
