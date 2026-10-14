# -*- coding: utf-8 -*-
"""
Хранилище истории расчётов в SQLite.
Файл базы: data/calc.db
"""

import os
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
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
    raw_text TEXT DEFAULT ''
);
"""


def init_db():
    """Создаёт папку data и таблицу postings, если их ещё нет."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(SCHEMA)


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
                labor_cost_one, labor_cost_total, unit_cost, total_cost, raw_text
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
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