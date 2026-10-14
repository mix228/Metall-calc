# -*- coding: utf-8 -*-
"""
Хранилище истории расчётов в SQLite.
Файл базы: data/calc.db
"""

import os
import sys
import sqlite3
from datetime import datetime


def is_termux():
    """True, если приложение запущено в Termux на Android."""
    prefix = os.environ.get("PREFIX", "")
    return "com.termux" in prefix or "com.termux" in sys.prefix.lower()


def get_data_dir():
    """Каталог данных (совпадает с логикой app.py). В Termux $HOME может быть
    недоступен на запись без termux-setup-storage — берём private dir."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if is_termux():
        d = os.environ.get("TERMUX_APP_PRIVATE_DATA_DIR")
        if not d:
            prefix = os.environ.get("PREFIX", "")
            pkg = os.path.basename(os.path.dirname(prefix)) or "com.termux"
            d = f"/data/data/{pkg}/app_private_files"
        if os.path.isdir(d) and os.access(d, os.W_OK):
            return os.path.join(d, "calc_data")
    return os.path.join(base_dir, "data")


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = get_data_dir()
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, "calc.db")


SCHEMA = """
CREATE TABLE IF NOT EXISTS postings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    drawing TEXT NOT NULL,
    shop TEXT DEFAULT '',
    material_name TEXT DEFAULT '',
    material_source TEXT DEFAULT '',
    diameter_mm REAL,
    length_mm REAL,
    quantity INTEGER,
    mass_one REAL,
    mass_total REAL,
    price_per_kg REAL,
    material_cost REAL,
    time_min REAL,
    price_per_min REAL,
    labor_cost_one REAL,
    labor_cost_total REAL,
    unit_cost REAL,
    total_cost REAL,
    raw_text TEXT DEFAULT '',
    unit TEXT DEFAULT '',
    required_units REAL,
    kg_per_unit REAL,
    stock_units_before REAL
);
"""

# Журнал складских движений (приход / списание / корректировка).
STOCK_SCHEMA = """
CREATE TABLE IF NOT EXISTS stock_moves (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    move_type TEXT NOT NULL,
    row_key TEXT DEFAULT '',
    material_label TEXT DEFAULT '',
    unit TEXT DEFAULT '',
    qty_units REAL,
    qty_kg REAL,
    stock_after REAL,
    basis TEXT DEFAULT '',
    posting_id INTEGER,
    shortage INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_stock_moves_posting ON stock_moves(posting_id);
"""


def _ensure_columns(conn, table, cols):
    """Мягкая миграция: добавляет отсутствующие колонки в существующую таблицу."""
    existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
    for name, ddl in cols:
        if name not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def init_db():
    """Создаёт папку data, таблицы postings и stock_moves, если их ещё нет."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(SCHEMA)
        conn.executescript(STOCK_SCHEMA)
        _ensure_columns(conn, "postings", [
            ("unit", "TEXT DEFAULT ''"),
            ("required_units", "REAL"),
            ("kg_per_unit", "REAL"),
            ("stock_units_before", "REAL"),
        ])


def _fmt_dim(v):
    """60.0 -> 60, 60.5 -> 60.5 (для красивого вывода в таблице)."""
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return v
    return int(f) if f == int(f) else f


def add_posting(data: dict) -> int:
    """
    Сохраняет подтверждённую запись в базу.
    Возвращает id новой строки.
    """
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            """
            INSERT INTO postings (
                created_at, drawing, shop, material_name, material_source,
                diameter_mm, length_mm, quantity, mass_one, mass_total,
                price_per_kg, material_cost, time_min, price_per_min,
                labor_cost_one, labor_cost_total, unit_cost, total_cost, raw_text,
                unit, required_units, kg_per_unit, stock_units_before
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                str(data.get("drawing") or ""),
                str(data.get("shop") or ""),
                str(data.get("material_name") or ""),
                str(data.get("material_source") or ""),
                data.get("diameter_mm"),
                data.get("length_mm"),
                data.get("quantity"),
                data.get("mass_one"),
                data.get("mass_total"),
                data.get("price_per_kg"),
                data.get("material_cost"),
                data.get("time_min"),
                data.get("price_per_min"),
                data.get("labor_cost_one"),
                data.get("labor_cost_total"),
                data.get("unit_cost"),
                data.get("total_cost"),
                str(data.get("raw_text") or ""),
                str(data.get("unit") or ""),
                data.get("required_units"),
                data.get("kg_per_unit"),
                data.get("stock_units_before"),
            ),
        )
        return cur.lastrowid


def fetch_history(limit=500):
    """
    Возвращает историю (свежие записи первыми)
    в виде списка словарей с ключами, которые ожидает шаблон index.html.
    """
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM postings ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()

    result = []
    for r in rows:
        result.append({
            "id": r["id"],
            "timestamp": r["created_at"],
            "drawing": r["drawing"],
            "shop": r["shop"],
            "material_name": r["material_name"],
            "material_source": r["material_source"],
            "diameter_mm": _fmt_dim(r["diameter_mm"]),
            "length_mm": _fmt_dim(r["length_mm"]),
            "quantity": r["quantity"],
            "mass_one": r["mass_one"],
            "mass_total": r["mass_total"],
            "price_per_kg": r["price_per_kg"],
            "material_cost": r["material_cost"],
            "time_min": r["time_min"],
            "price_per_min": r["price_per_min"],
            "labor_cost_one": r["labor_cost_one"],
            "labor_cost_total": r["labor_cost_total"],
            "unit_cost": r["unit_cost"],
            "total_cost": r["total_cost"],
            "raw_text": r["raw_text"],
            "unit": r["unit"] or "",
            "required_units": r["required_units"],
            "kg_per_unit": r["kg_per_unit"],
            "stock_units_before": r["stock_units_before"],
        })
    return result


def count_history():
    """Общее количество записей в истории."""
    with sqlite3.connect(DB_PATH) as conn:
        return conn.execute("SELECT COUNT(*) FROM postings").fetchone()[0]
    
def delete_posting(record_id: int) -> bool:
    """
    Удаляет запись из базы по id.
    Возвращает True, если строка была удалена.
    """
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("DELETE FROM postings WHERE id = ?", (record_id,))
        return cur.rowcount > 0

# ============================================================
# ЖУРНАЛ СКЛАДСКИХ ДВИЖЕНИЙ (stock_moves)
# ============================================================

def add_stock_move(move: dict) -> int:
    """Записывает движение в журнал. Возвращает id строки."""
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            """
            INSERT INTO stock_moves (
                created_at, move_type, row_key, material_label, unit,
                qty_units, qty_kg, stock_after, basis, posting_id, shortage
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                str(move.get("move_type") or ""),
                str(move.get("row_key") or ""),
                str(move.get("material_label") or ""),
                str(move.get("unit") or ""),
                move.get("qty_units"),
                move.get("qty_kg"),
                move.get("stock_after"),
                str(move.get("basis") or ""),
                move.get("posting_id"),
                1 if move.get("shortage") else 0,
            ),
        )
        return cur.lastrowid


def fetch_stock_moves(limit=200):
    """Последние движения (свежие первыми)."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM stock_moves ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def get_posting(record_id: int):
    """Возвращает запись истории по id или None."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        r = conn.execute(
            "SELECT * FROM postings WHERE id = ?", (record_id,)
        ).fetchone()
    return dict(r) if r else None


def moves_for_posting(record_id: int):
    """Движения журнала, привязанные к записи истории."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM stock_moves WHERE posting_id = ? ORDER BY id",
            (record_id,),
        ).fetchall()
    return [dict(r) for r in rows]
