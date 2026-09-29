import json
import math
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd

CONFIG_FILE = "config.json"
GROUPS_FILE = "groups.json"
ADDRESSES_FILE = "adresses.json"

KOROB_COLORS = ["белый", "коричневый", "серебро", "бежевый", "махагон", "золотой дуб"]

# ---------- Палитра ----------
BG_DARK       = "#0d1b2a"
BG_PANEL      = "#1b263b"
BG_INPUT      = "#22304a"
BG_CARD_ISO   = "#1f3a5f"
BG_CARD_VIT   = "#164e4a"
FG_TEXT       = "#e0e1dd"
FG_MUTED      = "#9bb1c9"
ACCENT        = "#4ea8de"
ACCENT_HOVER  = "#63b6e6"
ACCENT_VIT    = "#3ddc97"
DANGER        = "#e05661"
BORDER        = "#415a77"

FONT_MAIN     = ("Segoe UI", 10)
FONT_BOLD     = ("Segoe UI", 10, "bold")
FONT_TITLE    = ("Segoe UI", 11, "bold")
FONT_BIG      = ("Segoe UI", 22, "bold")


# ---------- Файлы ----------
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"price_path": ""}


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------- A1-нотация ----------
def col_to_index(col_letters):
    col = 0
    for ch in col_letters.upper():
        if not ch.isalpha():
            raise ValueError(f"Некорректная колонка: {col_letters}")
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col


def parse_cell(ref):
    ref = ref.strip().upper()
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    col_part, row_part = ref[:i], ref[i:]
    if not col_part or not row_part.isdigit():
        raise ValueError(f"Некорректная ячейка: {ref}")
    return int(row_part), col_to_index(col_part)


def parse_range(rng):
    if (isinstance(rng, (list, tuple)) and len(rng) == 4
            and all(isinstance(v, int) for v in rng)):
        return tuple(rng)
    if isinstance(rng, str):
        parts = rng.replace(" ", "").split(":")
        if len(parts) != 2:
            raise ValueError(f"Некорректный диапазон: {rng}")
        start, end = parts
    else:
        start, end = rng
    r1, c1 = parse_cell(start)
    r2, c2 = parse_cell(end)
    if r1 > r2:
        r1, r2 = r2, r1
    if c1 > c2:
        c1, c2 = c2, c1
    return r1, c1, r2, c2


# ---------- Группы ----------
def load_groups():
    raw = load_json(GROUPS_FILE)
    reverse = {}
    applies_shift = {}
    for tipo, info in raw["types"].items():
        applies_shift[tipo] = info.get("applies_shift", False)
        reverse[tipo] = {}
        for base_group, colors in info["base_group_colors"].items():
            for c in colors:
                reverse[tipo][c] = base_group
    return {
        "base_groups": raw["base_groups"],
        "shift_map": raw["shift_map"],
        "reverse": reverse,
        "applies_shift": applies_shift,
    }


def resolve_group(groups_data, tipo, color, korob):
    base = groups_data["reverse"].get(tipo, {}).get(color)
    if base is None:
        return None
    if not groups_data["applies_shift"].get(tipo, False):
        return base
    base_groups = groups_data["base_groups"]
    if base not in base_groups:
        return None
    idx = base_groups.index(base)
    row = groups_data["shift_map"].get(korob)
    if row is None or idx >= len(row):
        return None
    return row[idx]


def get_colors(groups_data, tipo):
    return list(groups_data["reverse"].get(tipo, {}).keys())


# ---------- Прайс ----------
def read_sheet(path, sheet_name):
    engine = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
    return pd.read_excel(path, sheet_name=sheet_name, header=None, engine=engine)


def _frange(start, end, step):
    n = round((end - start) / step) + 1
    return [round(start + i * step, 1) for i in range(n)]


def build_prices_for_type(df, tipo_addr):
    grid = tipo_addr["grid"]
    widths = _frange(grid["width_start"], grid["width_end"], grid["step"])
    heights = _frange(grid["height_start"], grid["height_end"], grid["step"])

    result = {}
    for group, rng in tipo_addr["groups"].items():
        r1, c1, r2, c2 = parse_range(rng)
        sub = df.iloc[r1 - 1: r2, c1 - 1: c2]
        table = {}
        for i, h in enumerate(heights):
            if i >= len(sub.index):
                break
            row = sub.iloc[i]
            for j, w in enumerate(widths):
                if j >= len(row):
                    break
                val = row.iloc[j]
                try:
                    table[(h, w)] = float(val)
                except (TypeError, ValueError):
                    continue
        result[group] = table
    return result


def build_all_prices(price_path):
    if not price_path or not os.path.exists(price_path):
        return {}
    addresses = load_json(ADDRESSES_FILE)
    cache = {}
    for tipo, tipo_addr in addresses.items():
        sheet = tipo_addr["sheet"]
        try:
            df = read_sheet(price_path, sheet)
        except Exception as e:
            raise RuntimeError(f"Лист '{sheet}' ({tipo}): {e}")
        cache[tipo] = build_prices_for_type(df, tipo_addr)
    return cache


def mm_to_m(value_mm):
    return math.ceil(value_mm / 100) / 10


def find_price(price_data, groups_data, tipo, color, korob, w_mm, h_mm):
    group = resolve_group(groups_data, tipo, color, korob)
    if group is None:
        return None
    table = price_data.get(tipo, {}).get(group, {})
    return table.get((mm_to_m(h_mm), mm_to_m(w_mm)))


# ---------- Тест ----------
def test_price(price_path):
    if not price_path:
        return False, "Путь к прайсу не указан."
    if not os.path.exists(price_path):
        return False, f"Файл не найден: {price_path}"

    try:
        addresses = load_json(ADDRESSES_FILE)
    except Exception as e:
        return False, f"Ошибка чтения adresses.json: {e}"

    engine = "xlrd" if price_path.lower().endswith(".xls") else "openpyxl"
    try:
        xls = pd.ExcelFile(price_path, engine=engine)
    except Exception as e:
        return False, f"Не удалось открыть прайс: {e}"

    available = set(xls.sheet_names)
    missing = {t: a["sheet"] for t, a in addresses.items() if a["sheet"] not in available}
    if missing:
        lines = [f"  • {t}: ожидается лист '{s}'" for t, s in missing.items()]
        return False, ("Не найдены листы в прайсе:\n" + "\n".join(lines) +
                       "\n\nДоступные листы: " + ", ".join(xls.sheet_names))

    problems = []
    for tipo, tipo_addr in addresses.items():
        sheet = tipo_addr["sheet"]
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception as e:
            problems.append(f"{tipo} / '{sheet}': не читается ({e})")
            continue

        grid = tipo_addr["grid"]
        n_w = len(_frange(grid["width_start"], grid["width_end"], grid["step"]))
        n_h = len(_frange(grid["height_start"], grid["height_end"], grid["step"]))

        for group, rng in tipo_addr["groups"].items():
            r1, c1, r2, c2 = parse_range(rng)
            sub = df.iloc[r1 - 1: r2, c1 - 1: c2]
            if sub.shape[0] < n_h or sub.shape[1] < n_w:
                problems.append(f"{tipo}/{group}: диапазон {sub.shape[0]}×{sub.shape[1]} "
                                f"меньше сетки {n_h}×{n_w}")
                continue
            block = sub.iloc[:n_h, :n_w]
            n_empty = int(block.isna().sum().sum())
            non_numeric = []
            for i in range(n_h):
                for j in range(n_w):
                    v = block.iat[i, j]
                    if pd.isna(v):
                        continue
                    if not isinstance(v, (int, float)):
                        try:
                            float(str(v).replace(",", "."))
                        except ValueError:
                            non_numeric.append((i, j, v))
            if n_empty or non_numeric:
                parts = [f"{tipo}/{group}:"]
                if n_empty:
                    parts.append(f"пустых ячеек — {n_empty}")
                if non_numeric:
                    sample = non_numeric[:3]
                    shown = ", ".join(f"[{i+1},{j+1}]='{v}'" for i, j, v in sample)
                    more = "" if len(non_numeric) <= 3 else f" и ещё {len(non_numeric)-3}"
                    parts.append(f"не числа: {shown}{more}")
                problems.append(" ".join(parts))

    if problems:
        return False, "Найдены проблемы в диапазонах:\n" + "\n".join(problems)
    return True, "Прайс корректен: все листы найдены, все ячейки — числа."


# ---------- Модель ----------
class Card:
    def __init__(self, tipo, color, w_mm, h_mm, korob, price, group=None):
        self.tipo = tipo
        self.color = color
        self.w_mm = int(round(w_mm))
        self.h_mm = int(round(h_mm))
        self.korob = korob
        self.price = price
        self.group = group


# ---------- Стили ----------
def setup_styles(root):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(bg=BG_DARK)

    style.configure(".", background=BG_DARK, foreground=FG_TEXT,
                    fieldbackground=BG_INPUT, font=FONT_MAIN)
    style.configure("TFrame", background=BG_DARK)
    style.configure("Panel.TFrame", background=BG_PANEL)
    style.configure("TLabel", background=BG_DARK, foreground=FG_TEXT, font=FONT_MAIN)
    style.configure("Panel.TLabel", background=BG_PANEL, foreground=FG_TEXT)
    style.configure("Muted.TLabel", background=BG_PANEL, foreground=FG_MUTED,
                    font=("Segoe UI", 9))
    style.configure("Header.TLabel", background=BG_PANEL, foreground=FG_TEXT,
                    font=FONT_TITLE)

    style.configure("TLabelframe", background=BG_PANEL, foreground=FG_TEXT,
                    bordercolor=BORDER, relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=BG_PANEL, foreground=ACCENT,
                    font=FONT_BOLD)
    style.configure("FieldLabel.TLabel",
                    background=BG_PANEL,
                    foreground=FG_MUTED,
                    font=("Segoe UI", 9))
    style.configure("Summary.TLabel",
                    background=BG_PANEL, foreground=FG_MUTED,
                    font=("Segoe UI", 11))
    style.configure("SummaryValue.TLabel",
                    background=BG_PANEL, foreground=ACCENT,
                    font=("Segoe UI", 16, "bold"))
    
    
    # Высокие поля ввода
    style.configure("Tall.TEntry",
                    fieldbackground=BG_INPUT, foreground=FG_TEXT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    insertcolor=FG_TEXT, padding=(8, 10), justify='center')
    style.map("Tall.TEntry", fieldbackground=[("focus", "#28395a")])

    style.configure("TCombobox",
                    fieldbackground=BG_INPUT, background=BG_INPUT,
                    foreground=FG_TEXT, arrowcolor=ACCENT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    padding=(8, 8))
    style.map("TCombobox",
              fieldbackground=[("readonly", BG_INPUT), ("focus", "#28395a")],
              foreground=[("readonly", FG_TEXT)],
              background=[("readonly", BG_INPUT)])
    
    
    # Combobox — высокий (как кнопки)
    style.configure("Tall.TCombobox",
                    fieldbackground=BG_INPUT, background=BG_INPUT,
                    foreground=FG_TEXT, arrowcolor=ACCENT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    padding=(8, 10))
    style.map("Tall.TCombobox",
              fieldbackground=[("readonly", BG_INPUT), ("focus", "#28395a")],
              foreground=[("readonly", FG_TEXT)],
              background=[("readonly", BG_INPUT)])
    
    
    root.option_add("*TCombobox*Listbox.background", BG_INPUT)
    root.option_add("*TCombobox*Listbox.foreground", FG_TEXT)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", "#0d1b2a")



    style.configure("Accent.TButton",
                    background=ACCENT, foreground="#0d1b2a",
                    font=FONT_BOLD, borderwidth=0, focuscolor=ACCENT,
                    padding=(12, 10))
    style.map("Accent.TButton",
              background=[("active", ACCENT_HOVER), ("pressed", ACCENT)])

    style.configure("Ghost.TButton",
                    background=BG_PANEL, foreground=FG_TEXT,
                    borderwidth=0, focuscolor=BG_PANEL, padding=(8, 6))
    style.map("Ghost.TButton",
              background=[("active", "#2a3b5a"), ("pressed", BG_PANEL)])

    style.configure("Toggle.TButton",
                    background=BG_INPUT, foreground=ACCENT,
                    font=FONT_BOLD, borderwidth=0, focuscolor=BG_INPUT,
                    padding=(10, 10))
    style.map("Toggle.TButton",
              background=[("active", "#2a3b5a"), ("pressed", BG_INPUT)])


# ---------- Tooltip ----------
class ToolTip:
    def __init__(self, widget, text, bg=BG_PANEL, fg=FG_TEXT):
        self.widget = widget
        self.text = text
        self.bg = bg
        self.fg = fg
        self.tip = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<Button-1>", self._hide, add="+")

    def _schedule(self, _):
        self._cancel()
        self._after_id = self.widget.after(400, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, bg=self.bg, fg=self.fg,
                 font=("Segoe UI", 9), padx=8, pady=4,
                 highlightthickness=1, highlightbackground=BORDER, bd=0).pack()

    def _hide(self, _=None):
        self._cancel()
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ---------- Иконки ----------
def make_badge(parent, text, bg_color, fg_color="#e0e1dd", size=34):
    cv = tk.Canvas(parent, width=size, height=size, bg=parent["bg"],
                   highlightthickness=0, bd=0)
    cv.create_oval(1, 1, size - 1, size - 1, fill=bg_color, outline="")
    cv.create_text(size / 2, size / 2 + 1, text=text, fill=fg_color,
                   font=("Segoe UI", 10, "bold"))
    return cv


def make_icon_button(parent, symbol, bg, hover_bg, fg, command,
                     size=30, tooltip=None):
    cv = tk.Canvas(parent, width=size, height=size, bg=bg,
                   highlightthickness=0, bd=0, cursor="hand2")
    rect = cv.create_rectangle(1, 1, size - 1, size - 1,
                               fill=bg, outline=BORDER)
    txt = cv.create_text(size / 2, size / 2, text=symbol, fill=fg,
                         font=("Segoe UI", 13, "bold"))

    def on_enter(_):
        cv.itemconfig(rect, fill=hover_bg)
        cv.itemconfig(txt, fill="#ffffff")

    def on_leave(_):
        cv.itemconfig(rect, fill=bg)
        cv.itemconfig(txt, fill=fg)

    cv.bind("<Enter>", on_enter, add="+")
    cv.bind("<Leave>", on_leave, add="+")
    cv.bind("<Button-1>", lambda e: command())

    if tooltip:
        ToolTip(cv, tooltip, bg=BG_PANEL, fg=FG_TEXT)
    return cv


# ---------- Карточка ----------
class CardWidget(tk.Frame):
    def __init__(self, parent, card, groups_data, on_edit, on_delete):
        super().__init__(parent, highlightthickness=1,
                         highlightbackground=BORDER, bd=0)
        self.card = card
        self.groups_data = groups_data
        self.on_edit = on_edit
        self.on_delete = on_delete
        self.bg = BG_CARD_ISO
        self._render()

    def _render(self):
        # фон зависит от текущего типа
        self.bg = BG_CARD_ISO if self.card.tipo == "isolite" else BG_CARD_VIT
        self.configure(bg=self.bg)

        for w in self.winfo_children():
            w.destroy()
        c = self.card

        # ---------- Справа (в порядке вызова справа → налево) ----------
        # 1) Кнопки — самые правые
        btns = tk.Frame(self, bg=self.bg)
        btns.pack(side="right", padx=(4, 12), pady=8)

        edit_btn = make_icon_button(
            btns, "⚙", self.bg, ACCENT, FG_TEXT,
            command=lambda: self.on_edit(self), size=32,
            tooltip="Редактировать",
        )
        edit_btn.pack(side="top", pady=(0, 4))

        del_btn = make_icon_button(
            btns, "✕", self.bg, DANGER, FG_TEXT,
            command=lambda: self.on_delete(self), size=32,
            tooltip="Удалить",
        )
        del_btn.pack(side="top")

        # 2) Цена
        price_text = f"{int(round(c.price))} ₽" if c.price is not None else "— ₽"
        tk.Label(
            self, text=price_text, bg=self.bg, fg=ACCENT,
            font=FONT_BIG, anchor="e", justify="right",
        ).pack(side="right", padx=(4, 16), pady=8)

        # 3) Размеры
        tk.Label(
            self, text=f"{c.w_mm} × {c.h_mm} мм",
            bg=self.bg, fg=FG_TEXT,
            font=("Segoe UI", 14, "bold"), anchor="e",
        ).pack(side="right", padx=(4, 20), pady=8)

        # ---------- Слева ----------
        left = tk.Frame(self, bg=self.bg)
        left.pack(side="left", fill="both", expand=True, padx=(12, 4), pady=8)

        # бейдж
        badge_text = "ISO" if c.tipo == "isolite" else "VIT"
        badge_color = ACCENT if c.tipo == "isolite" else ACCENT_VIT
        badge = make_badge(left, badge_text, badge_color, size=42)
        badge.pack(side="left", padx=(0, 12))

        # столбец "группа / короб"
        group_col = tk.Frame(left, bg=self.bg)
        group_col.pack(side="left", anchor="w")

        group_text = f"группа: {c.group}" if getattr(c, "group", None) else "группа: —"
        tk.Label(group_col, text=group_text, bg=self.bg, fg=FG_TEXT,
                 font=("Segoe UI", 10, "bold"), anchor="w"
                 ).pack(anchor="w")
        tk.Label(group_col, text=f"короб: {c.korob}", bg=self.bg, fg=FG_MUTED,
                 font=("Segoe UI", 10), anchor="w"
                 ).pack(anchor="w")

        # столбец "цвет" — вертикально выровнен по центру двух строк
        color_col = tk.Frame(left, bg=self.bg)
        color_col.pack(side="left", anchor="center", padx=(48, 0))

        tk.Label(color_col, text=c.color, bg=self.bg, fg=FG_TEXT,
                 font=("Segoe UI", 13, "bold"), anchor="center"
                 ).pack(anchor="center")
        
        group = resolve_group(self.groups_data, c.tipo, c.color, c.korob) \
            if self.groups_data else None
        group_text = f"группа: {group}" if group else "группа: —"
        
        # tk.Label(labels, text=group_text, bg=self.bg, fg=FG_MUTED,
        #          font=("Segoe UI", 10), anchor="w"
        #          ).pack(anchor="w")

        # tk.Label(labels, text=f"короб: {c.korob}", bg=self.bg, fg=FG_MUTED,
        #          font=("Segoe UI", 10), anchor="w"
        #          ).pack(anchor="w")

    def refresh(self):
        self._render()


# ---------- Диалог ----------
class EditDialog(tk.Toplevel):
    def __init__(self, parent, card, groups_data, on_save):
        super().__init__(parent)
        self.title("Редактирование")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.card = card
        self.groups_data = groups_data
        self.on_save = on_save

        self.tipo_var = tk.StringVar(value=card.tipo)
        self.color_var = tk.StringVar(value=card.color)
        self.korob_var = tk.StringVar(value=card.korob)
        self.w_var = tk.StringVar(value=str(card.w_mm))
        self.h_var = tk.StringVar(value=str(card.h_mm))

        frm = ttk.Frame(self, style="Panel.TFrame", padding=16)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Редактирование", style="Header.TLabel"
                  ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(frm, text="Тип:", style="Panel.TLabel"
                  ).grid(row=1, column=0, sticky="w", pady=4)
        tipo_cb = ttk.Combobox(frm, textvariable=self.tipo_var,
                               values=list(groups_data["reverse"].keys()),
                               state="readonly", width=24)
        tipo_cb.grid(row=1, column=1, sticky="ew", pady=4)
        tipo_cb.bind("<<ComboboxSelected>>", self._update_colors)

        ttk.Label(frm, text="Цвет:", style="Panel.TLabel"
                  ).grid(row=2, column=0, sticky="w", pady=4)
        self.color_cb = ttk.Combobox(frm, textvariable=self.color_var,
                                     state="readonly", width=24)
        self.color_cb.grid(row=2, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Короб:", style="Panel.TLabel"
                  ).grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(frm, textvariable=self.korob_var,
                     values=KOROB_COLORS, state="readonly", width=24
                     ).grid(row=3, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Ширина (мм):", style="Panel.TLabel"
                  ).grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.w_var, width=26).grid(row=4, column=1, sticky="ew", pady=4)

        ttk.Label(frm, text="Высота (мм):", style="Panel.TLabel"
                  ).grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.h_var, width=26).grid(row=5, column=1, sticky="ew", pady=4)

        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm, style="Panel.TFrame")
        btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Отмена", style="Ghost.TButton",
                   command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="Сохранить", style="Accent.TButton",
                   command=self._save).pack(side="right")

        self._update_colors()
        self.update_idletasks()
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(px, 0)}+{max(py, 0)}")

    def _update_colors(self, *_):
        colors = get_colors(self.groups_data, self.tipo_var.get())
        self.color_cb["values"] = colors
        if self.color_var.get() not in colors and colors:
            self.color_var.set(colors[0])

    def _save(self):
        try:
            w = int(round(float(self.w_var.get().replace(",", "."))))
            h = int(round(float(self.h_var.get().replace(",", "."))))
        except ValueError:
            messagebox.showerror("Ошибка", "Ширина и высота должны быть числами (мм)")
            return
        self.card.tipo = self.tipo_var.get()
        self.card.color = self.color_var.get()
        self.card.korob = self.korob_var.get()
        self.card.w_mm = w
        self.card.h_mm = h
        self.on_save(self.card)
        self.destroy()


# ---------- Приложение ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Расчёт цен")
        self.geometry("1020x660")
        self.minsize(920, 580)

        setup_styles(self)

        self.config_data = load_config()
        self.groups_data = load_groups()
        self.price_cache = {}
        self.cards = []

        self._build_ui()
        self._update_summary()
        
        if self.config_data.get("price_path"):
            self._rebuild_price_cache(silent=True)

    def _update_summary(self):
            count = len(self.cards)
            total = sum(
                c.price for c in self.cards
                if c.price is not None
            )
            # print("DBG cards:", len(self.cards), [c.price for c in self.cards])
            self.summary_count_var.set(f"Позиций: {count}")
            self.summary_total_var.set(f"Итого: {int(round(total))} ₽")

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=14, pady=(14, 6))

        inputs = ttk.LabelFrame(top, text="Добавить карточку", padding=12)
        inputs.pack(side="left", fill="both", expand=True)

                # ---------- Подблок 1: Параметры ----------
        params = ttk.Frame(inputs, style="Panel.TFrame")
        params.pack(fill="x", pady=(0, 10))

        ttk.Label(params, text="Параметры", style="Muted.TLabel"
                  ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # подписи над элементами
        ttk.Label(params, text="Тип", style="FieldLabel.TLabel"
                  ).grid(row=1, column=0, sticky="w", padx=(0, 6))
        ttk.Label(params, text="Цвет", style="FieldLabel.TLabel"
                  ).grid(row=1, column=1, sticky="w", padx=(0, 6))
        ttk.Label(params, text="Короб", style="FieldLabel.TLabel"
                  ).grid(row=1, column=2, sticky="w")

        self.tipo_var = tk.StringVar(value="isolite")
        self.color_var = tk.StringVar()
        self.korob_var = tk.StringVar(value=KOROB_COLORS[0])
        self.w_var = tk.StringVar()
        self.h_var = tk.StringVar()

        ttk.Button(params, textvariable=self.tipo_var, style="Toggle.TButton",
                   width=10, command=self._toggle_tipo
                   ).grid(row=2, column=0, sticky="ew", padx=(0, 6))

        self.color_cb = ttk.Combobox(params, textvariable=self.color_var,
                                     state="readonly", width=16,
                                     style="Tall.TCombobox")
        self.color_cb.grid(row=2, column=1, sticky="ew", padx=(0, 6))

        ttk.Combobox(params, textvariable=self.korob_var,
                     values=KOROB_COLORS, state="readonly", width=14,
                     style="Tall.TCombobox"
                     ).grid(row=2, column=2, sticky="ew")

        for i in range(3):
            params.columnconfigure(i, weight=1)

                       # ---------- Подблок 2: Размеры ----------
        dims = ttk.Frame(inputs, style="Panel.TFrame")
        dims.pack(fill="x")

        ttk.Label(dims, text="Размеры и добавление", style="Muted.TLabel"
                  ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        # подписи
        ttk.Label(dims, text="Ширина, мм", style="FieldLabel.TLabel"
                  ).grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Label(dims, text="Высота, мм", style="FieldLabel.TLabel"
                  ).grid(row=1, column=1, sticky="w", padx=(0, 8))
        ttk.Label(dims, text="", style="FieldLabel.TLabel"
                  ).grid(row=1, column=2, sticky="w")
        ttk.Label(dims, text="", style="FieldLabel.TLabel"
                  ).grid(row=1, column=3, sticky="w")

        # поля ввода
        self.w_entry = ttk.Entry(dims, textvariable=self.w_var, style="Tall.TEntry")
        self.w_entry.grid(row=2, column=0, sticky="ew", padx=(0, 8))
        self._placeholder(self.w_entry, self.w_var, "Например, 500")

        self.h_entry = ttk.Entry(dims, textvariable=self.h_var, style="Tall.TEntry")
        self.h_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8))
        self._placeholder(self.h_entry, self.h_var, "Например, 1200")

        # кнопка очистки (иконка)
        clear_holder = tk.Frame(dims, bg=BG_PANEL, height=42)
        clear_holder.grid(row=2, column=2, sticky="ns", padx=(0, 8))
        clear_holder.grid_propagate(False)

        clear_btn = make_icon_button(
            clear_holder, "⟲", BG_PANEL, ACCENT, FG_TEXT,
            command=self._clear_size_fields, size=42,
            tooltip="Очистить поля",
        )
        clear_btn.pack()

        # кнопка добавления — растягиваем на всё оставшееся место
        ttk.Button(dims, text="+ Добавить", style="Accent.TButton",
                   command=self._add_card).grid(row=2, column=3, sticky="ew")

        # ширина инпутов — примерно четверть от доступной ширины
        dims.columnconfigure(0, weight=1, uniform="size", minsize=100)
        dims.columnconfigure(1, weight=1, uniform="size", minsize=100)
        dims.columnconfigure(2, weight=0, minsize=52)   # кнопка очистки
        dims.columnconfigure(3, weight=3)               # растягивающаяся кнопка

        self._update_colors()

        # Настройки
        settings = ttk.LabelFrame(top, text="Настройки", padding=12)
        settings.pack(side="right", fill="y", padx=(10, 0))

        self.path_var = tk.StringVar(value=self.config_data.get("price_path", ""))
        ttk.Entry(settings, textvariable=self.path_var, width=30
                  ).pack(fill="x", pady=(0, 6))
        ttk.Button(settings, text="Выбрать файл", style="Ghost.TButton",
                   command=self._choose_file).pack(fill="x", pady=2)
        ttk.Button(settings, text="Сохранить путь", style="Ghost.TButton",
                   command=self._save_path).pack(fill="x", pady=2)
        ttk.Button(settings, text="Тест", style="Ghost.TButton",
                   command=self._run_test).pack(fill="x", pady=2)
        ttk.Button(settings, text="Обновить цены", style="Ghost.TButton",
                   command=self._rebuild_price_cache).pack(fill="x", pady=2)

        # Список карточек
        list_frame = ttk.LabelFrame(self, text="Карточки", padding=8)
        list_frame.pack(fill="both", expand=True, padx=14, pady=(6, 14))

        # --- Итого: упаковываем ДО канваса с side="bottom" ---
        summary = tk.Frame(list_frame, bg=BG_PANEL)
        summary.pack(side="bottom", fill="x", pady=(8, 0))

        # тонкая линия-разделитель сверху
        tk.Frame(summary, bg=BORDER, height=1).pack(fill="x", pady=(0, 8))

        row = tk.Frame(summary, bg=BG_PANEL)
        row.pack(fill="x")

        self.summary_count_var = tk.StringVar(value="Позиций: 0")
        self.summary_total_var = tk.StringVar(value="Итого: 0 ₽")

        tk.Label(row, textvariable=self.summary_count_var,
                 bg=BG_PANEL, fg=FG_MUTED,
                 font=("Segoe UI", 11), anchor="w"
                 ).pack(side="left")

        tk.Label(row, textvariable=self.summary_total_var,
                 bg=BG_PANEL, fg=ACCENT,
                 font=("Segoe UI", 16, "bold"), anchor="e"
                 ).pack(side="right")

        # --- Прокручиваемая область ---
        self.canvas = tk.Canvas(list_frame, highlightthickness=0, bg=BG_PANEL)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.cards_frame = tk.Frame(self.canvas, bg=BG_PANEL)
        self.cards_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self._win = self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self._bind_mousewheel()

        

    # ---------- Плейсхолдеры ----------
    def _placeholder(self, entry, var, text):
        entry.insert(0, text)
        entry.configure(foreground=FG_MUTED)

        def on_focus_in(_):
            if entry.get() == text:
                entry.delete(0, "end")
                entry.configure(foreground=FG_TEXT)

        def on_focus_out(_):
            if not entry.get():
                entry.insert(0, text)
                entry.configure(foreground=FG_MUTED)

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _read_entry(self, var, placeholder):
        v = var.get().strip()
        if v == placeholder:
            return None
        return v

    # ---------- Скролл ----------
    def _bind_mousewheel(self):
        def on_wheel(e):
            delta = -1 if e.delta > 0 else 1
            self.canvas.yview_scroll(delta * 2, "units")
        self.canvas.bind_all("<MouseWheel>", on_wheel)

    # ---------- UI callbacks ----------
    def _toggle_tipo(self):
        self.tipo_var.set("viteo" if self.tipo_var.get() == "isolite" else "isolite")
        self._update_colors()

    def _update_colors(self):
        colors = get_colors(self.groups_data, self.tipo_var.get())
        self.color_cb["values"] = colors
        if colors and self.color_var.get() not in colors:
            self.color_var.set(colors[0])

    def _choose_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if path:
            self.path_var.set(path)

    def _save_path(self):
        self.config_data["price_path"] = self.path_var.get()
        save_config(self.config_data)
        self._rebuild_price_cache(silent=True)

    def _rebuild_price_cache(self, silent=False):
        path = self.path_var.get()
        self.config_data["price_path"] = path
        save_config(self.config_data)
        try:
            self.price_cache = build_all_prices(path)
            if not silent:
                messagebox.showinfo("Готово", "Цены загружены.")
        except Exception as e:
            if not silent:
                messagebox.showerror("Ошибка", str(e))

    def _run_test(self):
        ok, msg = test_price(self.path_var.get())
        (messagebox.showinfo if ok else messagebox.showwarning)("Тест", msg)
    def _clear_size_fields(self):
            for entry, text in (
                (self.w_entry, "Например, 500"),
                (self.h_entry, "Например, 1200"),
            ):
                entry.delete(0, "end")
                entry.insert(0, text)
                entry.configure(foreground=FG_MUTED)
            if not self.color_var.get():
                messagebox.showerror("Ошибка", "Выберите цвет")
            return
    def _add_card(self):
        

        w_raw = self._read_entry(self.w_var, "Ширина, мм")
        h_raw = self._read_entry(self.h_var, "Высота, мм")
        if w_raw is None or h_raw is None:
            messagebox.showerror("Ошибка", "Введите ширину и высоту (мм)")
            return

        try:
            w = int(round(float(w_raw.replace(",", "."))))
            h = int(round(float(h_raw.replace(",", "."))))
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные ширину и высоту (мм)")
            return

        tipo = self.tipo_var.get()
        color = self.color_var.get()
        korob = self.korob_var.get()

        price = None
        if self.price_cache:
            price = find_price(self.price_cache, self.groups_data,
                               tipo, color, korob, w, h)

        group = resolve_group(self.groups_data, tipo, color, korob)
        card = Card(tipo, color, w, h, korob, price, group=group)
        self.cards.append(card)
        widget = CardWidget(self.cards_frame, card, self.groups_data,
                            self._edit_card, self._delete_card)
        widget.pack(fill="x", pady=4, padx=4)
        
        # print("ADD:", len(self.cards), card.price)
        # self._update_summary()
        # print("AFTER:", self.summary_count_var.get(), self.summary_total_var.get())
        self._update_summary()
        
        
    def _edit_card(self, widget):
        card = widget.card

        def on_save(c):
            c.group = resolve_group(self.groups_data, c.tipo, c.color, c.korob)
            if self.price_cache:
                c.price = find_price(self.price_cache, self.groups_data,
                                     c.tipo, c.color, c.korob, c.w_mm, c.h_mm)
            widget.refresh()
            self._update_summary()
            
        EditDialog(self, card, self.groups_data, on_save)

    def _delete_card(self, widget):
        self.cards.remove(widget.card)
        widget.destroy()
        self._update_summary()


if __name__ == "__main__":
    app = App()
    app.mainloop()