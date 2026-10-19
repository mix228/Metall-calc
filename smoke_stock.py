# -*- coding: utf-8 -*-
"""Smoke-тест складского контура через Flask test client."""
import os, shutil, tempfile

TMP = tempfile.mkdtemp(prefix="stock_smoke_")
os.environ["CALC_DATA_DIR"] = TMP  # если поддерживается app.py; иначе подменим ниже

import app as A
import db

# Принудительно перенаправляем данные во временный каталог
A.DATA_DIR = TMP
A.PROJECT_DATA_DIR = TMP
A.MATERIALS_XLSX = os.path.join(TMP, "materials.xlsx")
A.MATERIALS_CSV = os.path.join(TMP, "materials.csv")
A.UNITS_JSON = os.path.join(TMP, "units.json")
db.DB_PATH = os.path.join(TMP, "calc.db")

from openpyxl import Workbook
wb = Workbook(); ws = wb.active; ws.title = "Материалы"
ws.append(["Тип заготовки","Материал","Марка","Размер","Плотность, кг/м3",
           "Цена за кг, руб","Вес 1 м, кг","Примечание"])   # СТАРЫЙ файл без складских колонок
ws.append(["круг","Сталь","45","60",7850,85,22.2,""])
ws.append([None,"Труба","12х1",None,None,90,0.27,""])
wb.save(A.MATERIALS_XLSX)
if os.path.exists(A.UNITS_JSON): os.remove(A.UNITS_JSON)

db.init_db()
A.load_materials_db()

def get(url): return A.app.test_client().get(url)
cl = A.app.test_client()

print("== Миграция идемпотентна ==")
hdrs = [c.value for c in __import__("openpyxl").load_workbook(A.MATERIALS_XLSX).active[1]]
assert "Единица учёта" in hdrs and "Остаток на складе" in hdrs and \
       "Кг на 1 единицу учёта" in hdrs and "Накладная" in hdrs, hdrs
print("колонки добавлены:", hdrs[-4:])
A.load_materials_db(); A.load_materials_db()  # повторный вызов ничего не ломает
hdrs2 = [c.value for c in __import__("openpyxl").load_workbook(A.MATERIALS_XLSX).active[1]]
assert hdrs2 == hdrs, (hdrs, hdrs2)
print("повторная миграция — без изменений OK")
units = A.load_units()
assert units == A.DEFAULT_UNITS, units
print("units.json создан со стартовым набором:", units)

print("\n== GET /stock -> 200 (не заглушка) ==")
r = cl.get("/stock")
body = r.get_data(as_text=True)
assert r.status_code == 200, r.status_code
assert "Ошибка склада" not in body
assert "Учёт остатков" in body or "учёт остатков" in body
for marker in ("Приход по накладной", "журнал", "Единицы учёта", "Недостач"):
    assert marker.lower() in body.lower(), marker
print("status:", r.status_code, ", блоки найдены OK")
r2 = cl.get("/materials_page"); assert r2.status_code == 200
print("GET /materials_page ->", r2.status_code)

print("\n== Синхронизация: задаём единицу/кг-на-ед на /materials_page, видим на /stock ==")
row = A.find_row_by_key("Сталь|45|60")
r = cl.post("/materials/update", data={
    "row": str(row["source_row"]), "profile": "круг", "material": "Сталь",
    "grade": "45", "size": "60", "density": "7850", "price_per_kg": "85",
    "weight_per_m": "22.2", "note": "", "unit": "м", "stock_units": "10",
    "kg_per_unit": "22.2", "invoice": ""})
assert r.status_code == 302
b = cl.get("/stock").get_data(as_text=True)
assert "10" in b and "222" in b, "остаток в кг (10*22.2=222) не виден на /stock"
print("значение с /materials_page видно на /stock OK (10 ед. = 222 кг)")

print("\n== Приход по накладной меняет остаток и пишет журнал ==")
before = A.find_row_by_key("Сталь|45|60")["stock_units"]
r = cl.post("/stock/receive", data={"key": "Сталь|45|60", "qty": "5", "invoice": "НК-1 от 20.10.2026"})
assert r.status_code == 302
after = A.find_row_by_key("Сталь|45|60")["stock_units"]
print(f"остаток {before} -> {after}")
assert after == before + 5
moves = db.fetch_stock_moves(10)
assert moves and moves[0]["move_type"] == "приход" and moves[0]["basis"] == "НК-1 от 20.10.2026"
print("журнал: последнее движение =", moves[0]["move_type"], moves[0]["qty_units"], moves[0]["unit"])

print("\n== /calculate: поля склада + предупреждение при нехватке, расчёт не блокируется ==")
payload = {"data": "ЧР-100\nСталь 45 круг Ø60\nДлина 1000 мм\nКол-во 1 шт\nШв 10 мин\nСтавка 50 руб"}
import json
r = cl.post("/calculate", data=json.dumps(payload), content_type="application/json")
res = r.get_json()["result"]
mass_total = res["mass_total"]
req = round(mass_total / 22.2, 6)
print("mass_total =", mass_total, "kg; stock_units =", res["stock_units"],
      "required_units =", res["required_units"], "stock_ok =", res["stock_ok"])
assert res["stock_tracked"] and abs(res["required_units"] - req) < 1e-6
assert res["stock_ok"] is True  # 15 м в остатке, нужно меньше

# большой расход -> жёлтое предупреждение, но success=True
payload2 = dict(payload); payload2["data"] = payload["data"].replace("Кол-во 1 шт", "Кол-во 50 шт")
r = cl.post("/calculate", data=json.dumps(payload2), content_type="application/json")
res2 = r.get_json()["result"]
assert r.status_code == 200 and r.get_json()["success"] and res2["stock_ok"] is False
print("предупреждение:", res2["stock_warning"])

print("\n== /confirm: списание <= остатка (зелёное) ==")
r = cl.post("/confirm", data=json.dumps(res), content_type="application/json")
j = r.get_json(); assert j["success"], j
rid = j["record"]["id"]
st1 = A.find_row_by_key("Сталь|45|60")["stock_units"]
print(f"запись №{rid}: остаток 15 -> {st1}, списано {res['required_units']} м")
mv = [m for m in db.moves_for_posting(rid)]
assert mv and mv[0]["move_type"] == "списание" and not mv[0]["shortage"]
assert st1 < 15

print("\n== /confirm: расход > остатка -> минус, флаг недостачи, блок «Недостачи» ==")
r = cl.post("/confirm", data=json.dumps(res2), content_type="application/json")
j2 = r.get_json(); assert j2["success"], j2
rid2 = j2["record"]["id"]
st2 = A.find_row_by_key("Сталь|45|60")["stock_units"]
print(f"запись №{rid2}: остаток стал {st2} (минус), msg: {j2['stock_messages']}")
assert st2 < 0
mv2 = db.moves_for_posting(rid2)
assert mv2 and mv2[0]["move_type"] == "списание" and mv2[0]["shortage"]
sb = cl.get("/stock").get_data(as_text=True)
assert "недостача" in sb.lower()
print("флаг недостачи в журнале и блок «Недостачи» на странице OK")

print("\n== /delete записи с недостачей восстанавливает остаток ==")
r = cl.post("/delete", data=json.dumps({"id": rid2}), content_type="application/json")
j3 = r.get_json(); assert j3["success"], j3
st3 = A.find_row_by_key("Сталь|45|60")["stock_units"]
print(f"остаток {st2} -> {st3}; msgs: {j3['stock_messages']}")
assert abs(st3 - st1) < 1e-6
cancel = [m for m in db.fetch_stock_moves(5) if m["move_type"] == "корректировка" and "отмена" in (m["basis"] or "")]
assert cancel, "нет отменяющего движения"
print("отменяющее движение записано:", cancel[0]["basis"], cancel[0]["qty_units"])

print("\n== Единицы: добавить «упак.», удалить; неиспользуемую «компл» можно удалить ==")
r = cl.post("/stock/unit/add", data={"name": "упак."}); assert r.status_code == 302
assert "упак." in A.load_units()
r = cl.post("/stock/unit/delete", data={"name": "м"})  # используется Сталь 45
b = cl.get("/stock").get_data(as_text=True)
assert "используется" in b, "удаление используемой единицы не отклонено"
print("удаление используемой единицы отклонено OK")
r = cl.post("/stock/unit/delete", data={"name": "упак."}); assert "упак." not in A.load_units()
print("добавление/удаление неиспользуемой единицы OK")

print("\n== Корректировка кнопкой из строки ==")
r = cl.post("/stock/adjust", data={"key": "Сталь|45|60", "new_value": "100", "reason": "инвентаризация"})
assert A.find_row_by_key("Сталь|45|60")["stock_units"] == 100
print("корректировка до 100 OK")

print("\n== Перезапуск: остаток (xlsx) и журнал (SQLite) на месте ==")
A.load_materials_db()
assert A.find_row_by_key("Сталь|45|60")["stock_units"] == 100
assert db.fetch_stock_moves(500)
print("после перечитывания всё на месте OK")

print("\nSMOKE TEST: ВСЁ ЗЕЛЁНОЕ ✅")
