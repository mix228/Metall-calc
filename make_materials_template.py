# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook

os.makedirs("data", exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "Материалы"

headers = [
    "Тип заготовки", "Материал", "Марка", "Размер",
    "Плотность, кг/м3", "Цена за кг, руб", "Вес 1 м, кг", "Примечание",
]
ws.append(headers)

rows = [
    ["круг", "Сталь", "45", "60", 7850, 85, None, ""],
    ["круг", "Сталь", "45", "80", 7850, 85, None, ""],
    ["круг", "Сталь", "3", "40", 7850, 70, None, ""],
    ["круг", "Сталь", "20", "50", 7850, 75, None, ""],
    ["лист", "Сталь", "3", "3х1000х2000", 7850, 70, None, ""],
    ["труба", "Сталь", "20", "57х3,5", 7850, 90, 4.62, ""],
    ["круг", "Капролон", "", "50", 1140, 400, None, ""],
    ["лист", "Текстолит", "", "10х500х500", 1350, 600, None, ""],
]

for r in rows:
    ws.append(r)

wb.save(os.path.join("data", "materials.xlsx"))
print("Создан файл data/materials.xlsx")