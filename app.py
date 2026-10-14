# -*- coding: utf-8 -*-

import os
import re
import math
import json
import csv
import logging
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, url_for

import db

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
MATERIALS_XLSX = os.path.join(DATA_DIR, "materials.xlsx")
MATERIALS_CSV = os.path.join(DATA_DIR, "materials.csv")

logging.basicConfig(level=logging.INFO)


# ============================================================
# НАСТРОЙКИ (config.json)
# ============================================================

DEFAULT_CONFIG = {
    "price_per_min": 10.0,
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                file_cfg = json.load(f)
            if isinstance(file_cfg, dict):
                cfg.update(file_cfg)
        except Exception:
            logging.exception("Не удалось прочитать config.json, использую значения по умолчанию")

    # Защита от битых значений в config.json
    if not isinstance(cfg.get("price_per_min"), (int, float)):
        cfg["price_per_min"] = DEFAULT_CONFIG["price_per_min"]

    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ============================================================
# ДИАПАЗОНЫ ДОПУСТИМЫХ ЗНАЧЕНИЙ
# ============================================================

FIELD_LIMITS = {
    "time_min": (0.0, 10000.0),          # минут на деталь
    "price_per_min": (0.0, 100000.0),    # руб за минуту
    "quantity": (1, 1000000),            # штук
    "dimension_mm": (0.1, 1000000.0),    # мм
}


def check_range(label, value, lo, hi):
    if value is None:
        return None
    if value < lo or value > hi:
        return f"{label}: значение вне допустимого диапазона. Разрешено от {lo:g} до {hi:g}."
    return None


# ============================================================
# ВСТРОЕННЫЙ ТЕСТОВЫЙ СПРАВОЧНИК (fallback)
# ============================================================

MATERIALS = {
    "сталь 45": {"density": 7850, "price_per_kg": 85.0},
    "сталь 3": {"density": 7850, "price_per_kg": 70.0},
    "сталь 20": {"density": 7850, "price_per_kg": 75.0},
    "сталь 09г2с": {"density": 7850, "price_per_kg": 80.0},
    "нержавейка": {"density": 7900, "price_per_kg": 250.0},
    "алюминий": {"density": 2700, "price_per_kg": 320.0},
    "латунь": {"density": 8500, "price_per_kg": 600.0},
    "медь": {"density": 8900, "price_per_kg": 900.0},
    "капролон": {"density": 1140, "price_per_kg": 400.0},
    "текстолит": {"density": 1350, "price_per_kg": 600.0},
}

MATERIALS_DB = []
MATERIALS_INFO = {"source": "встроенный (файл не найден)", "count": 0}
MATERIALS_SOURCE_KIND = "none"
MATERIALS_COLMAP = {}
MATERIALS_HEADER_FIELDS = []
MATERIALS_CSV_FIELDS = []
MATERIALS_CSV_RAWNAMES = []

DEFAULT_HEADERS = [
    "Тип заготовки", "Материал", "Марка", "Размер",
    "Плотность, кг/м3", "Цена за кг, руб", "Вес 1 м, кг", "Примечание",
]
DEFAULT_HEADER_FIELDS = [
    "profile", "material", "grade", "size",
    "density", "price_per_kg", "weight_per_m", "note",
]
NUM_FIELDS = {"density", "price_per_kg", "weight_per_m"}


# ============================================================
# ЗАГРУЗКА СПРАВОЧНИКА
# ============================================================

PROFILE_SYNONYMS = {
    "круг": "round", "круглый": "round", "пруток": "round",
    "лист": "sheet", "пластина": "sheet", "плита": "sheet",
    "труба": "pipe", "трубка": "pipe",
    "квадрат": "square",
    "уголок": "angle",
    "швеллер": "channel",
    "шестигранник": "hex",
}

HEADER_MAP = {
    "тип заготовки": "profile", "тип": "profile", "профиль": "profile",
    "материал": "material", "наименование": "material",
    "марка": "grade", "марка материала": "grade",
    "размер": "size", "размеры": "size", "габарит": "size", "габариты": "size",
    "плотность, кг/м3": "density", "плотность": "density",
    "цена за кг, руб": "price_per_kg", "цена за кг": "price_per_kg", "цена": "price_per_kg",
    "вес 1 м, кг": "weight_per_m", "вес 1 м": "weight_per_m",
    "примечание": "note",
}


def normalize_header(h):
    return re.sub(r"\s+", " ", str(h or "").lower().replace("ё", "е").strip())


def to_float(v):
    if v is None:
        return None
    s = str(v).replace(",", ".").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def build_row(raw: dict):
    profile_raw = str(raw.get("profile") or "").strip().lower()
    profile = PROFILE_SYNONYMS.get(profile_raw, profile_raw or None)

    return {
        "profile": profile,
        "material": str(raw.get("material") or "").strip(),
        "grade": str(raw.get("grade") or "").strip(),
        "size": str(raw.get("size") or "").strip(),
        "density": to_float(raw.get("density")),
        "price_per_kg": to_float(raw.get("price_per_kg")),
        "weight_per_m": to_float(raw.get("weight_per_m")),
        "note": str(raw.get("note") or "").strip(),
        "source_row": None,
    }


def load_materials_db():
    global MATERIALS_DB, MATERIALS_INFO, MATERIALS_SOURCE_KIND
    global MATERIALS_COLMAP, MATERIALS_HEADER_FIELDS
    global MATERIALS_CSV_FIELDS, MATERIALS_CSV_RAWNAMES

    rows = []
    source = "встроенный (файл не найден)"
    kind = "none"
    MATERIALS_COLMAP = {}
    MATERIALS_HEADER_FIELDS = []
    MATERIALS_CSV_FIELDS = []
    MATERIALS_CSV_RAWNAMES = []

    if os.path.exists(MATERIALS_XLSX):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(MATERIALS_XLSX, data_only=True)
            ws = wb["Материалы"] if "Материалы" in wb.sheetnames else wb.active

            header_cells = [normalize_header(c.value) for c in ws[1]]
            col_map = {}
            for idx, h in enumerate(header_cells):
                if h in HEADER_MAP:
                    col_map[HEADER_MAP[h]] = idx

            MATERIALS_COLMAP = {f: i + 1 for f, i in col_map.items()}
            fields_by_col = [None] * len(header_cells)
            for f, i in col_map.items():
                fields_by_col[i] = f
            MATERIALS_HEADER_FIELDS = fields_by_col

            for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if row is None or all(v is None for v in row):
                    continue
                raw = {field: row[idx] if idx < len(row) else None
                       for field, idx in col_map.items()}
                r = build_row(raw)
                if r["material"] or r["grade"]:
                    r["source_row"] = row_num
                    rows.append(r)

            source = "data/materials.xlsx"
            kind = "xlsx"
        except Exception as e:
            logging.exception("Ошибка чтения materials.xlsx")
            source = f"ошибка чтения data/materials.xlsx: {e}"

    elif os.path.exists(MATERIALS_CSV):
        try:
            with open(MATERIALS_CSV, encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                MATERIALS_CSV_RAWNAMES = list(reader.fieldnames or [])
                MATERIALS_CSV_FIELDS = [
                    HEADER_MAP.get(normalize_header(fn), normalize_header(fn))
                    for fn in MATERIALS_CSV_RAWNAMES
                ]
                for row_num, row in enumerate(reader, start=2):
                    raw = {HEADER_MAP.get(normalize_header(k), normalize_header(k)): v
                           for k, v in row.items()}
                    r = build_row(raw)
                    if r["material"] or r["grade"]:
                        r["source_row"] = row_num
                        rows.append(r)
            source = "data/materials.csv"
            kind = "csv"
        except Exception as e:
            logging.exception("Ошибка чтения materials.csv")
            source = f"ошибка чтения data/materials.csv: {e}"

    MATERIALS_DB = rows
    MATERIALS_INFO = {"source": source, "count": len(rows)}
    MATERIALS_SOURCE_KIND = kind
    return MATERIALS_INFO


def find_material(profile_type, material_name, dims):
    mat_low = re.sub(r"\s+", " ", (material_name or "").lower().strip())

    if MATERIALS_DB:
        candidates = []
        for r in MATERIALS_DB:
            if r["profile"] and profile_type and r["profile"] != profile_type:
                continue

            key = re.sub(r"\s+", " ", f"{r['material']} {r['grade']}".lower().strip())

            if key == mat_low or mat_low in key or key in mat_low:
                candidates.append(r)
            elif r["grade"] and r["grade"].lower() == mat_low:
                candidates.append(r)

        if candidates:
            def size_score(r):
                s = (r["size"] or "").replace(",", ".").lower()
                if not s or not dims:
                    return 0
                score = 0
                for v in dims.values():
                    if v is None:
                        continue
                    if f"{v:g}" in s:
                        score += 1
                return score

            candidates.sort(key=size_score, reverse=True)
            r = candidates[0]

            return {
                "name": re.sub(r"\s+", " ", f"{r['material']} {r['grade']}".strip()),
                "density": r["density"],
                "price_per_kg": r["price_per_kg"],
                "weight_per_m": r["weight_per_m"],
                "size": r["size"],
                "source": MATERIALS_INFO["source"],
            }, None

        return None, f"Материал '{material_name}' не найден в справочнике ({MATERIALS_INFO['source']})."

    m = MATERIALS.get(mat_low)
    if m:
        return {
            "name": mat_low,
            "density": m["density"],
            "price_per_kg": m["price_per_kg"],
            "weight_per_m": None,
            "size": None,
            "source": "встроенный",
        }, None

    return None, f"Для материала '{material_name}' нет цены/плотности во встроенном справочнике."


# ============================================================
# РЕДАКТИРОВАНИЕ СПРАВОЧНИКА
# ============================================================

def _row_values(values, fields, empty):
    out = []
    for f in fields:
        v = values.get(f)
        if f in NUM_FIELDS:
            v = to_float(v)
            out.append(empty if v is None else v)
        else:
            s = (v or "").strip() if isinstance(v, str) else ("" if v is None else str(v))
            out.append(s if s != "" else empty)
    return out


def _open_xlsx_ws():
    from openpyxl import load_workbook
    wb = load_workbook(MATERIALS_XLSX)
    ws = wb["Материалы"] if "Материалы" in wb.sheetnames else wb.active
    return wb, ws


def add_material(values):
    if MATERIALS_SOURCE_KIND == "csv":
        os.makedirs(DATA_DIR, exist_ok=True)
        if not os.path.exists(MATERIALS_CSV):
            with open(MATERIALS_CSV, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).append(DEFAULT_HEADERS)
            load_materials_db()
        fields = MATERIALS_CSV_FIELDS or DEFAULT_HEADER_FIELDS
        with open(MATERIALS_CSV, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).append(_row_values(values, fields, empty=""))
        load_materials_db()
        return "Строка добавлена"

    from openpyxl import load_workbook, Workbook

    if MATERIALS_SOURCE_KIND == "none":
        os.makedirs(DATA_DIR, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "Материалы"
        ws.append(DEFAULT_HEADERS)
        fields = DEFAULT_HEADER_FIELDS
    else:
        wb, ws = _open_xlsx_ws()
        fields = MATERIALS_HEADER_FIELDS or DEFAULT_HEADER_FIELDS

    ws.append(_row_values(values, fields, empty=None))
    wb.save(MATERIALS_XLSX)
    load_materials_db()
    return "Строка добавлена"


def update_material(row_idx, values):
    if not row_idx:
        return "Неизвестная строка"

    if MATERIALS_SOURCE_KIND == "xlsx":
        wb, ws = _open_xlsx_ws()
        for field, col in MATERIALS_COLMAP.items():
            v = values.get(field)
            if field in NUM_FIELDS:
                v = to_float(v)
            else:
                v = (v or "").strip() if isinstance(v, str) else v
                v = v if v not in ("", None) else None
            ws.cell(row=row_idx, column=col).value = v
        wb.save(MATERIALS_XLSX)
        load_materials_db()
        return "Строка сохранена"

    if MATERIALS_SOURCE_KIND == "csv":
        with open(MATERIALS_CSV, encoding="utf-8-sig") as f:
            all_rows = list(csv.reader(f))
        idx = row_idx - 1
        if 0 < idx < len(all_rows):
            all_rows[idx] = [
                str(x) if x is not None else ""
                for x in _row_values(values, MATERIALS_CSV_FIELDS, empty="")
            ]
        with open(MATERIALS_CSV, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(all_rows)
        load_materials_db()
        return "Строка сохранена"

    return "Справочник не из файла — редактирование недоступно"


def delete_material(row_idx):
    if not row_idx:
        return "Неизвестная строка"

    if MATERIALS_SOURCE_KIND == "xlsx":
        wb, ws = _open_xlsx_ws()
        ws.delete_rows(row_idx)
        wb.save(MATERIALS_XLSX)
        load_materials_db()
        return "Строка удалена"

    if MATERIALS_SOURCE_KIND == "csv":
        with open(MATERIALS_CSV, encoding="utf-8-sig") as f:
            all_rows = list(csv.reader(f))
        idx = row_idx - 1
        if 0 < idx < len(all_rows):
            all_rows.pop(idx)
        with open(MATERIALS_CSV, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(all_rows)
        load_materials_db()
        return "Строка удалена"

    return "Справочник не из файла — удаление недоступно"


# ============================================================
# ПАРСЕРЫ
# ============================================================

def normalize_text(text: str) -> str:
    text = text.lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


SHOP_RE = re.compile(
    r"\b(цех|участок|уч\.?|отделение|производство)\s*(?:№)?\s*:?\s*"
    r"([0-9]{1,4}|[A-Za-zА-Яа-я][A-Za-zА-Яа-я\-]{0,24})",
    re.IGNORECASE,
)


def parse_shop(text: str):
    m = SHOP_RE.search(text)
    if not m:
        return None

    kind_raw = m.group(1).lower().rstrip(".")
    if kind_raw == "цех":
        kind = "цех"
    elif kind_raw.startswith("уч"):
        kind = "участок"
    else:
        kind = kind_raw

    return {"kind": kind, "value": m.group(2).strip()}


DRAWING_CODE = r"[0-9A-Za-zА-Яа-я]+(?:[.\-/][0-9A-Za-zА-Яа-я]+)*"
DRAWING_SUFFIX = r"(?:\s*(?:СБ|ДСП|СП|ВО|И))?"

DRAWING_MARKER_RE = re.compile(
    r"(?:чертеж|черт\.?|кд|№|деталь|позиция|поз\.?|обозначение)\s*:?\s*("
    + DRAWING_CODE + DRAWING_SUFFIX + r")",
    re.IGNORECASE,
)

DRAWING_FULL_RE = re.compile(
    r"^(" + DRAWING_CODE + DRAWING_SUFFIX + r")$",
    re.IGNORECASE,
)

DIMENSION_IN_SEGMENT_RE = re.compile(r"\d\s*(?:х|x|×|\*|на)\s*\d", re.IGNORECASE)

SKIP_WORDS = [
    "круг", "пруток", "круглый", "лист", "труба", "квадрат",
    "уголок", "швеллер", "сталь", "ст.", "шт",
]


def split_segments(text: str):
    return [s.strip() for s in re.split(r"[,;]+", text) if s.strip()]


def extract_drawing(text: str):
    clean = SHOP_RE.sub(" ", text)

    m = DRAWING_MARKER_RE.search(clean)
    if m and re.search(r"\d", m.group(1)):
        code = m.group(1).strip()
        rest = clean[:m.start(1)] + " " + clean[m.end(1):]
        return code, rest

    for seg in split_segments(clean):
        low = seg.lower()

        if any(w in low for w in SKIP_WORDS):
            continue

        if DIMENSION_IN_SEGMENT_RE.search(seg):
            continue

        code = seg.strip()
        if not DRAWING_FULL_RE.match(code):
            continue

        has_sep = bool(re.search(r"[.\-/]", code))
        has_letter = bool(re.search(r"[A-Za-zА-Яа-я]", code))
        has_digit = bool(re.search(r"\d", code))

        ok = (
            has_sep
            or (has_letter and has_digit)
            or (has_digit and not has_letter and len(code) <= 6)
        )

        if ok:
            rest = clean.replace(seg, " ", 1)
            return code, rest

    return None, clean


def parse_profile_type(text: str):
    if re.search(r"\b(?:круг|пруток|круглый)\b", text):
        return "round"
    return None


def parse_material(text: str):
    m = re.search(r"\b(капролон|текстолит|алюминий|латунь|медь|нержавейка)\b", text)
    if m:
        return m.group(1)

    m = re.search(r"\b(?:сталь|ст\.?)\s*([0-9]{1,3}|[0-9а-з]+)", text)
    if m:
        return f"сталь {m.group(1)}"

    m = re.search(r"\bкруг\s+([0-9]{1,3})(?!\s*(?:х|x|×|\*|на|–|-))\b", text)
    if m:
        return f"сталь {m.group(1)}"

    return None


def parse_quantity(text: str):
    m = re.search(r"(\d+)\s*(?:шт\.?|штук|штуки|шт)", text)
    if m:
        return int(m.group(1))

    m = re.search(r"(?:количество|кол-во|кол\.?)\s*:?\s*(\d+)", text)
    if m:
        return int(m.group(1))

    words = {
        "один": 1, "одна": 1, "одно": 1,
        "два": 2, "две": 2,
        "три": 3, "четыре": 4, "пять": 5,
        "шесть": 6, "семь": 7, "восемь": 8,
        "девять": 9, "десять": 10,
    }

    m = re.search(
        r"\b(один|одна|одно|два|две|три|четыре|пять|шесть|семь|восемь|девять|десять)"
        r"\s*(?:шт\.?|штук|штуки|шт)",
        text,
    )
    if m:
        return words.get(m.group(1))

    return None


def parse_time_minutes(text: str):
    m = re.search(
        r"(?:штучное\s*время|шт\s*время|время|нормо-?время)\s*:?\s*"
        r"(\d+(?:[.,]\d+)?)\s*(?:мин|мин\.|минут|минуты)?",
        text,
    )
    if m:
        return float(m.group(1).replace(",", "."))

    m = re.search(r"\b(\d+(?:[.,]\d+)?)\s*(?:мин|мин\.|минут|минуты)\b", text)
    if m:
        return float(m.group(1).replace(",", "."))

    return None


def to_mm(value, unit=None):
    val = float(str(value).replace(",", "."))
    if unit:
        unit = unit.lower()
        if unit in {"м", "m", "метр", "метра", "метров"}:
            return val * 1000
        if unit in {"см", "cm"}:
            return val * 10
    return val


def parse_round_dimensions(text: str):
    dims = {}

    pair = re.search(
        r"(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?\s*(?:х|x|×|\*|на)\s*"
        r"(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text,
        flags=re.IGNORECASE,
    )

    if pair:
        dims["diameter_mm"] = to_mm(pair.group(1), pair.group(2))
        dims["length_mm"] = to_mm(pair.group(3), pair.group(4))

    m = re.search(
        r"(?:диаметр|диам\.?|d|∅|ø)\s*:?\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text, flags=re.IGNORECASE,
    )
    if m:
        dims["diameter_mm"] = to_mm(m.group(1), m.group(2))

    m = re.search(
        r"(?:длина|дл\.?|l)\s*:?\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text, flags=re.IGNORECASE,
    )
    if m:
        dims["length_mm"] = to_mm(m.group(1), m.group(2))

    if "length_mm" not in dims:
        m = re.search(r"\b(\d+(?:[.,]\d+)?)\s*(?:м|метр|метра|метров)\b", text, flags=re.IGNORECASE)
        if m:
            dims["length_mm"] = to_mm(m.group(1), "м")

    return dims


def calc_round_mass_kg(diameter_mm: float, length_mm: float, density: float) -> float:
    volume_mm3 = math.pi / 4 * diameter_mm ** 2 * length_mm
    volume_m3 = volume_mm3 / 1_000_000_000
    return volume_m3 * density


# ============================================================
# МАРШРУТЫ FLASK
# ============================================================

API_ERROR_TEXT = (
    "При обработке запроса произошла внутренняя ошибка сервера. "
    "Попробуйте ещё раз. Если ошибка повторяется — обратитесь к администратору."
)


@app.route("/")
def index():
    return render_template(
        "index.html",
        history=db.fetch_history(),
        settings=load_config(),
        materials_info=MATERIALS_INFO,
    )


def _calculate_impl():
    data = request.form.get("input_text", "")
    shop_from_form = request.form.get("shop", "").strip()
    time_from_form = request.form.get("time_min", "").strip()
    price_min_from_form = request.form.get("price_per_min", "").strip()

    if not data.strip():
        return jsonify({"success": False, "errors": ["Введите текст"], "warnings": []})

    text = normalize_text(data)

    drawing, text_core = extract_drawing(text)

    profile_type = parse_profile_type(text_core)
    material_name = parse_material(text_core)
    quantity = parse_quantity(text_core)
    shop_parsed = parse_shop(text)
    time_parsed = parse_time_minutes(text_core)

    dims = {}
    if profile_type == "round":
        dims = parse_round_dimensions(text_core)

    errors = []
    warnings = []

    if not drawing:
        errors.append("Не найден номер чертежа.")

    if profile_type != "round":
        errors.append("В тестовой версии пока поддерживается только круг.")

    if not material_name:
        errors.append("Не найден материал.")

    if not quantity:
        quantity = 1
        warnings.append("Количество не указано. Использую 1 шт.")

    # --- проверка количества по диапазону ---
    err = check_range("Количество", quantity, *FIELD_LIMITS["quantity"])
    if err:
        errors.append(err)

    if "diameter_mm" not in dims or "length_mm" not in dims:
        errors.append("Не найдены размеры. Нужны диаметр и длина.")

    # --- проверка размеров по диапазону ---
    for key, label in (("diameter_mm", "Диаметр (мм)"), ("length_mm", "Длина (мм)")):
        v = dims.get(key)
        if v is not None:
            err = check_range(label, v, *FIELD_LIMITS["dimension_mm"])
            if err:
                errors.append(err)

    config = load_config()

    # --- штучное время: поле формы важнее текста ---
    if time_from_form:
        time_min = to_float(time_from_form)
        if time_min is None:
            errors.append("Штучное время: не удалось распознать число. Введите например 25 или 25,5.")
        else:
            err = check_range("Штучное время (мин)", time_min, *FIELD_LIMITS["time_min"])
            if err:
                errors.append(err)
    else:
        time_min = time_parsed
        if time_min is not None:
            err = check_range("Штучное время (мин)", time_min, *FIELD_LIMITS["time_min"])
            if err:
                errors.append(err)

    # --- стоимость минуты: поле формы важнее config ---
    if price_min_from_form:
        price_per_min = to_float(price_min_from_form)
        if price_per_min is None:
            errors.append("Стоимость минуты: не удалось распознать число. Введите например 12 или 12,5.")
        else:
            err = check_range("Стоимость минуты (руб)", price_per_min, *FIELD_LIMITS["price_per_min"])
            if err:
                errors.append(err)
    else:
        price_per_min = config["price_per_min"]

    material = None
    if material_name:
        material, mat_err = find_material(profile_type, material_name, dims)
        if material is None:
            errors.append(mat_err)
        elif material["density"] is None or material["price_per_kg"] is None:
            errors.append(
                f"Для материала '{material['name']}' в справочнике не заполнены плотность или цена."
            )

    if errors:
        return jsonify({"success": False, "errors": errors, "warnings": warnings})

    # Цех / участок: поле формы важнее распознанного из текста
    if shop_from_form:
        shop_display = shop_from_form
    elif shop_parsed:
        shop_display = f"{shop_parsed['kind']} {shop_parsed['value']}"
    else:
        shop_display = ""

    mass_one = calc_round_mass_kg(
        diameter_mm=dims["diameter_mm"],
        length_mm=dims["length_mm"],
        density=material["density"],
    )

    mass_total = mass_one * quantity
    material_cost = mass_total * material["price_per_kg"]

    if time_min:
        labor_cost_one = time_min * price_per_min
        labor_cost_total = labor_cost_one * quantity
    else:
        labor_cost_one = 0.0
        labor_cost_total = 0.0
        warnings.append("Штучное время не указано. Стоимость работы не учтена.")

    total_cost = material_cost + labor_cost_total
    unit_cost = (material_cost / quantity) + labor_cost_one if quantity else total_cost

    result = {
        "drawing": drawing,
        "shop": shop_display,
        "material_name": material["name"],
        "material_source": material["source"],
        "diameter_mm": dims["diameter_mm"],
        "length_mm": dims["length_mm"],
        "quantity": quantity,
        "mass_one": round(mass_one, 3),
        "mass_total": round(mass_total, 3),
        "price_per_kg": material["price_per_kg"],
        "material_cost": round(material_cost, 2),
        "time_min": time_min,
        "price_per_min": price_per_min,
        "labor_cost_one": round(labor_cost_one, 2),
        "labor_cost_total": round(labor_cost_total, 2),
        "unit_cost": round(unit_cost, 2),
        "total_cost": round(total_cost, 2),
        "raw_text": data,
        "warnings": warnings,
    }

    return jsonify({"success": True, "result": result})


@app.route("/calculate", methods=["POST"])
def calculate():
    try:
        return _calculate_impl()
    except Exception:
        logging.exception("Ошибка в /calculate")
        return jsonify({
            "success": False,
            "errors": [API_ERROR_TEXT],
            "warnings": [],
        }), 500


def _confirm_impl():
    data = request.get_json(silent=True)

    if not isinstance(data, dict):
        return jsonify({
            "success": False,
            "errors": ["Запись получена в неверном формате. Обновите страницу и попробуйте ещё раз."],
        }), 400

    drawing = str(data.get("drawing") or "").strip()
    if not drawing:
        return jsonify({"success": False, "errors": ["В записи нет номера чертежа."]}), 400

    def num(field, limits=None, label=None):
        v = to_float(data.get(field))
        if v is None:
            return None, None
        if limits:
            err = check_range(label or field, v, *limits)
            if err:
                return v, err
        return v, None

    field_errors = []

    quantity, err = num("quantity", FIELD_LIMITS["quantity"], "Количество")
    if err:
        field_errors.append(err)
    if quantity is None:
        quantity = 1

    mass_one, err = num("mass_one")
    if err:
        field_errors.append(err)
    mass_total, err = num("mass_total")
    if err:
        field_errors.append(err)
    material_cost, err = num("material_cost")
    if err:
        field_errors.append(err)
    time_min, err = num("time_min", FIELD_LIMITS["time_min"], "Штучное время (мин)")
    if err:
        field_errors.append(err)
    price_per_min, err = num("price_per_min", FIELD_LIMITS["price_per_min"], "Стоимость минуты (руб)")
    if err:
        field_errors.append(err)
    labor_cost_total, err = num("labor_cost_total")
    if err:
        field_errors.append(err)
    total_cost, err = num("total_cost")
    if err:
        field_errors.append(err)

    if field_errors:
        return jsonify({"success": False, "errors": field_errors}), 400

    record = {
        "drawing": drawing,
        "shop": str(data.get("shop") or ""),
        "material_name": str(data.get("material_name") or ""),
        "material_source": str(data.get("material_source") or ""),
        "diameter_mm": data.get("diameter_mm", ""),
        "length_mm": data.get("length_mm", ""),
        "quantity": int(quantity),
        "mass_one": mass_one,
        "mass_total": mass_total,
        "price_per_kg": to_float(data.get("price_per_kg")),
        "material_cost": material_cost,
        "time_min": time_min,
        "price_per_min": price_per_min,
        "labor_cost_one": to_float(data.get("labor_cost_one")),
        "labor_cost_total": labor_cost_total,
        "unit_cost": to_float(data.get("unit_cost")),
        "total_cost": total_cost,
        "raw_text": str(data.get("raw_text") or ""),
    }

    try:
        record_id = db.add_posting(record)
    except Exception:
        logging.exception("Ошибка сохранения истории в SQLite")
        return jsonify({
            "success": False,
            "errors": ["Не удалось сохранить запись в базу истории. Попробуйте ещё раз."],
        }), 500

    record["id"] = record_id

    return jsonify({"success": True, "message": "Запись сохранена", "record": record})


@app.route("/confirm", methods=["POST"])
def confirm():
    try:
        return _confirm_impl()
    except Exception:
        logging.exception("Ошибка в /confirm")
        return jsonify({"success": False, "errors": [API_ERROR_TEXT]}), 500


@app.route("/reload_materials", methods=["POST"])
def reload_materials():
    try:
        info = load_materials_db()
        return jsonify(info)
    except Exception:
        logging.exception("Ошибка в /reload_materials")
        return jsonify({
            "source": "ошибка перезагрузки справочника",
            "count": 0,
        }), 500


@app.route("/delete", methods=["POST"])
def delete_record():
    data = request.get_json(silent=True) or {}
    raw_id = data.get("id")
    try:
        record_id = int(raw_id)
    except (TypeError, ValueError):
        return jsonify({"success": False, "errors": ["Некорректный идентификатор записи."]}), 400

    try:
        deleted = db.delete_posting(record_id)
    except Exception:
        logging.exception("Ошибка удаления записи из SQLite")
        return jsonify({"success": False, "errors": ["Не удалось удалить запись. Попробуйте ещё раз."]}), 500

    if not deleted:
        return jsonify({"success": False, "errors": ["Запись с таким идентификатором не найдена."]}), 404

    return jsonify({"success": True, "message": "Запись удалена."})


@app.route("/materials")
def materials_json():
    return jsonify({"info": MATERIALS_INFO, "rows": MATERIALS_DB[:300]})


@app.route("/materials_page")
def materials_page():
    return render_template(
        "materials.html",
        rows=MATERIALS_DB,
        info=MATERIALS_INFO,
        source_kind=MATERIALS_SOURCE_KIND,
        msg=request.args.get("msg", ""),
    )


@app.route("/materials/add", methods=["POST"])
def materials_add():
    try:
        values = {k: request.form.get(k, "") for k in DEFAULT_HEADER_FIELDS}
        msg = add_material(values)
    except Exception:
        logging.exception("Ошибка добавления строки справочника")
        msg = "Не удалось добавить строку: внутренняя ошибка. Попробуйте ещё раз."
    return redirect(url_for("materials_page", msg=msg))


@app.route("/materials/update", methods=["POST"])
def materials_update():
    try:
        row = to_float(request.form.get("row"))
        values = {k: request.form.get(k, "") for k in DEFAULT_HEADER_FIELDS}
        msg = update_material(int(row) if row else None, values)
    except Exception:
        logging.exception("Ошибка сохранения строки справочника")
        msg = "Не удалось сохранить строку: внутренняя ошибка. Попробуйте ещё раз."
    return redirect(url_for("materials_page", msg=msg))


@app.route("/materials/delete", methods=["POST"])
def materials_delete():
    try:
        row = to_float(request.form.get("row"))
        msg = delete_material(int(row) if row else None)
    except Exception:
        logging.exception("Ошибка удаления строки справочника")
        msg = "Не удалось удалить строку: внутренняя ошибка. Попробуйте ещё раз."
    return redirect(url_for("materials_page", msg=msg))


# ============================================================
# СТАРТ
# ============================================================

db.init_db()
load_materials_db()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)