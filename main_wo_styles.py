import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import pandas as pd

CONFIG_FILE = "config.json"
GROUPS_FILE = "groups.json"
ADDRESSES_FILE = "adresses.json"

KOROB_COLORS = ["белый", "коричневый", "серебро", "бежевый", "махагон", "золотой дуб"]


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
    """'A' → 1, 'B' → 2, 'AA' → 27."""
    col = 0
    for ch in col_letters.upper():
        if not ch.isalpha():
            raise ValueError(f"Некорректная колонка: {col_letters}")
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col


def parse_cell(ref):
    """'B3' → (row=3, col=2). Возвращает (row, col) в 1-based."""
    ref = ref.strip().upper()
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    col_part = ref[:i]
    row_part = ref[i:]
    if not col_part or not row_part.isdigit():
        raise ValueError(f"Некорректная ячейка: {ref}")
    return int(row_part), col_to_index(col_part)


def parse_range(rng):
    """
    Принимает:
      - ['B3','Q24']  — пара углов
      - 'B3:Q24'      — строка Excel
      - [3,2,24,17]   — старый числовой формат
    Возвращает (row_start, col_start, row_end, col_end) в 1-based.
    """
    # старый формат
    if (
        isinstance(rng, (list, tuple))
        and len(rng) == 4
        and all(isinstance(v, int) for v in rng)
    ):
        return tuple(rng)
    # строка 'B3:Q24'
    if isinstance(rng, str):
        parts = rng.replace(" ", "").split(":")
        if len(parts) != 2:
            raise ValueError(f"Некорректный диапазон: {rng}")
        start, end = parts
        r1, c1 = parse_cell(start)
        r2, c2 = parse_cell(end)
    else:
        # пара углов ['B3','Q24']
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
    base_groups = raw["base_groups"]
    shift_map = raw["shift_map"]

    reverse = {}  # {тип: {цвет: базовая_группа}}
    applies_shift = {}  # {тип: bool}
    for tipo, info in raw["types"].items():
        applies_shift[tipo] = info.get("applies_shift", False)
        reverse[tipo] = {}
        for base_group, colors in info["base_group_colors"].items():
            for c in colors:
                reverse[tipo][c] = base_group

    return {
        "base_groups": base_groups,
        "shift_map": shift_map,
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
    """Читает лист без заголовков. Требует openpyxl (xlsx) или xlrd (xls)."""
    engine = "xlrd" if path.lower().endswith(".xls") else "openpyxl"
    return pd.read_excel(path, sheet_name=sheet_name, header=None, engine=engine)


def _frange(start, end, step):
    """Аккуратный список чисел с плавающей точкой (без 0.30000000004)."""
    n = round((end - start) / step) + 1
    return [round(start + i * step, 1) for i in range(n)]


def build_prices_for_type(df, tipo_addr):
    grid = tipo_addr["grid"]
    widths = _frange(grid["width_start"], grid["width_end"], grid["step"])
    heights = _frange(grid["height_start"], grid["height_end"], grid["step"])

    result = {}
    for group, rng in tipo_addr["groups"].items():
        r1, c1, r2, c2 = parse_range(rng)
        sub = df.iloc[r1 - 1 : r2, c1 - 1 : c2]
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
                    price = round(float(val))
                except (TypeError, ValueError):
                    continue
                table[(h, w)] = price
        result[group] = table
    return result


def build_all_prices(price_path):
    """{тип: {группа: {(h, w): price}}}"""
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
    """
    Миллиметры → метры с округлением ВВЕРХ до 1 знака после запятой.
    1200 → 1.2, 1201 → 1.3, 1150 → 1.2, 1140 → 1.2, 1100 → 1.1.
    """
    # ceil(1150 / 100) = ceil(11.5) = 12 → 1.2
    return math.ceil(value_mm / 100) / 10


def find_price(price_data, groups_data, tipo, color, korob, w_mm, h_mm):
    group = resolve_group(groups_data, tipo, color, korob)
    if group is None:
        return None
    table = price_data.get(tipo, {}).get(group, {})
    key = (mm_to_m(h_mm), mm_to_m(w_mm))
    return table.get(key)


# ---------- Тест ----------
def test_price(price_path):
    """
    Двухэтапный тест:
      1) Проверка наличия файла и листов в прайсе.
      2) Проверка, что во всех диапазонах сетки лежат числа (не пусто, не текст).
    Возвращает (ok: bool, message: str).
    """
    # --- Шаг 0: файл ---
    if not price_path:
        return False, "Путь к прайсу не указан."
    if not os.path.exists(price_path):
        return False, f"Файл не найден: {price_path}"

    # --- Шаг 1: наличие листов ---
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
    required = {tipo: addr["sheet"] for tipo, addr in addresses.items()}
    missing = {
        tipo: sheet for tipo, sheet in required.items() if sheet not in available
    }

    if missing:
        lines = [
            f"  • {tipo}: ожидается лист '{sheet}'" for tipo, sheet in missing.items()
        ]
        return False, (
            "Не найдены листы в прайсе:\n"
            + "\n".join(lines)
            + "\n\nДоступные листы: "
            + ", ".join(xls.sheet_names)
        )

    # --- Шаг 2: проверка данных в диапазонах ---
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
            sub = df.iloc[r1 - 1 : r2, c1 - 1 : c2]
            # размер блока
            if sub.shape[0] < n_h or sub.shape[1] < n_w:
                problems.append(
                    f"{tipo}/{group}: диапазон {sub.shape[0]}×{sub.shape[1]} "
                    f"меньше сетки {n_h}×{n_w}"
                )
                continue

            block = sub.iloc[:n_h, :n_w]

            # пустые ячейки
            empty_mask = block.isna()
            n_empty = int(empty_mask.sum().sum())

            # нечисловые ячейки
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
                    more = (
                        "" if len(non_numeric) <= 3 else f" и ещё {len(non_numeric)-3}"
                    )
                    parts.append(f"не числа: {shown}{more}")
                problems.append(" ".join(parts))

    if problems:
        return False, "Найдены проблемы в диапазонах:\n" + "\n".join(problems)
    return True, "Прайс корректен: все листы найдены, все ячейки — числа."


# ---------- Модель ----------
class Card:
    def __init__(self, tipo, color, w_mm, h_mm, korob, price):
        self.tipo = tipo
        self.color = color
        self.w_mm = int(round(w_mm))  # храним как ввёл пользователь (мм)
        self.h_mm = int(round(h_mm))
        self.korob = korob
        self.price = price


class CardWidget(ttk.Frame):
    def __init__(self, parent, card, on_edit, on_delete):
        super().__init__(parent, borderwidth=1, relief="solid", padding=6)
        self.card = card
        self.on_edit = on_edit
        self.on_delete = on_delete
        self._render()

    def _render(self):
        for w in self.winfo_children():
            w.destroy()
        c = self.card
        price = f"{int(round(c.price))} ₽" if c.price is not None else "—"
        text = (
            f"[{c.tipo}] {c.color} [короб: {c.korob}] - "
            f"{c.w_mm} x {c.h_mm} мм  —  {price}"
        )
        ttk.Label(self, text=text, anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Button(self, text="Удалить", command=lambda: self.on_delete(self)).pack(
            side="right", padx=2
        )
        ttk.Button(self, text="Редактировать", command=lambda: self.on_edit(self)).pack(
            side="right", padx=2
        )

    def refresh(self):
        self._render()


class EditDialog(tk.Toplevel):
    def __init__(self, parent, card, groups_data, on_save):
        super().__init__(parent)
        self.title("Редактирование")
        self.card = card
        self.groups_data = groups_data
        self.on_save = on_save

        self.tipo_var = tk.StringVar(value=card.tipo)
        self.color_var = tk.StringVar(value=card.color)
        self.korob_var = tk.StringVar(value=card.korob)
        self.w_var = tk.StringVar(value=str(card.w_mm))
        self.h_var = tk.StringVar(value=str(card.h_mm))

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Тип:").grid(row=0, column=0, sticky="w", pady=2)
        tipo_cb = ttk.Combobox(
            frm,
            textvariable=self.tipo_var,
            values=list(groups_data["reverse"].keys()),
            state="readonly",
        )
        tipo_cb.grid(row=0, column=1, sticky="ew", pady=2)
        tipo_cb.bind("<<ComboboxSelected>>", self._update_colors)

        ttk.Label(frm, text="Цвет:").grid(row=1, column=0, sticky="w", pady=2)
        self.color_cb = ttk.Combobox(frm, textvariable=self.color_var, state="readonly")
        self.color_cb.grid(row=1, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Короб:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Combobox(
            frm, textvariable=self.korob_var, values=KOROB_COLORS, state="readonly"
        ).grid(row=2, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Ширина (мм):").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(frm, textvariable=self.w_var).grid(row=3, column=1, sticky="ew", pady=2)

        ttk.Label(frm, text="Высота (мм):").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(frm, textvariable=self.h_var).grid(row=4, column=1, sticky="ew", pady=2)

        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(self, padding=(10, 0, 10, 10))
        btns.pack(fill="x")
        ttk.Button(btns, text="Сохранить", command=self._save).pack(
            side="right", padx=2
        )
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right", padx=2)

        self._update_colors()

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
        self.geometry("950x600")

        self.config_data = load_config()
        self.groups_data = load_groups()
        self.price_cache = {}
        self.cards = []

        self._build_ui()

        if self.config_data.get("price_path"):
            self._rebuild_price_cache(silent=True)

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        inputs = ttk.LabelFrame(top, text="Добавить карточку", padding=10)
        inputs.pack(side="left", fill="both", expand=True)

        self.tipo_var = tk.StringVar(value="isolite")
        self.color_var = tk.StringVar()
        self.korob_var = tk.StringVar(value=KOROB_COLORS[0])
        self.w_var = tk.StringVar()
        self.h_var = tk.StringVar()

        ttk.Button(inputs, textvariable=self.tipo_var, command=self._toggle_tipo).grid(
            row=0, column=0, sticky="ew", padx=4, pady=4
        )

        self.color_cb = ttk.Combobox(
            inputs, textvariable=self.color_var, state="readonly"
        )
        self.color_cb.grid(row=0, column=1, sticky="ew", padx=4, pady=4)

        ttk.Combobox(
            inputs, textvariable=self.korob_var, values=KOROB_COLORS, state="readonly"
        ).grid(row=0, column=2, sticky="ew", padx=4, pady=4)

        ttk.Entry(inputs, textvariable=self.w_var).grid(
            row=1, column=0, sticky="ew", padx=4, pady=4
        )
        ttk.Entry(inputs, textvariable=self.h_var).grid(
            row=1, column=1, sticky="ew", padx=4, pady=4
        )
        ttk.Button(inputs, text="Добавить", command=self._add_card).grid(
            row=1, column=2, sticky="ew", padx=4, pady=4
        )

        for i in range(3):
            inputs.columnconfigure(i, weight=1)

        self._update_colors()

        settings = ttk.LabelFrame(top, text="Настройки", padding=10)
        settings.pack(side="right", fill="y", padx=(10, 0))

        self.path_var = tk.StringVar(value=self.config_data.get("price_path", ""))
        ttk.Entry(settings, textvariable=self.path_var, width=28).pack(fill="x", pady=2)
        ttk.Button(settings, text="Выбрать файл", command=self._choose_file).pack(
            fill="x", pady=2
        )
        ttk.Button(settings, text="Сохранить путь", command=self._save_path).pack(
            fill="x", pady=2
        )
        ttk.Button(settings, text="Тест", command=self._run_test).pack(fill="x", pady=2)
        ttk.Button(
            settings, text="Обновить цены", command=self._rebuild_price_cache
        ).pack(fill="x", pady=2)

        list_frame = ttk.LabelFrame(self, text="Карточки", padding=5)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(list_frame, highlightthickness=0)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.cards_frame = ttk.Frame(self.canvas)
        self.cards_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    # ---- UI callbacks ----
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

    def _add_card(self):
        if not self.color_var.get():
            messagebox.showerror("Ошибка", "Выберите цвет")
            return
        try:
            w = int(round(float(self.w_var.get().replace(",", "."))))
            h = int(round(float(self.h_var.get().replace(",", "."))))
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

        card = Card(tipo, color, w, h, korob, price)
        self.cards.append(card)
        widget = CardWidget(self.cards_frame, card, self._edit_card, self._delete_card)
        widget.pack(fill="x", pady=2, padx=2)

    def _edit_card(self, widget):
        card = widget.card

        def on_save(c):
            if self.price_cache:
                c.price = find_price(self.price_cache, self.groups_data,
                                    c.tipo, c.color, c.korob, c.w_mm, c.h_mm)
            widget.refresh()

        EditDialog(self, card, self.groups_data, on_save)

        

    def _delete_card(self, widget):
        self.cards.remove(widget.card)
        widget.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
