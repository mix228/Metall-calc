# -*- coding: utf-8 -*-
from doc_builder import build_kalkulyatsiya

groups = [
    {
        "no": 1,
        "product": "Поддон под 630 тару",
        "shop": "3",
        "drawing": "",
        "sub_part": "",
        "time_min": 960,
        "qty": 1,
        "materials": [
            {"unit": "м", "name": "Уголок металл. 75х75х6мм", "consumption": 2.0, "price": 458.2},
            {"unit": "м", "name": "Уголок металл. 50х50х5мм", "consumption": 16.8, "price": 220.49},
            {"unit": "м", "name": "Труба проф. 40х20х2мм", "consumption": 0.6, "price": 132.15},
        ],
    },
    {
        "no": 2,
        "product": "Пиноль",
        "shop": "5",
        "drawing": "P-Drum_Twister-1",
        "sub_part": "",
        "time_min": 720,
        "qty": 10,
        "materials": [
            {"unit": "кг", "name": "Круг Ст.45 130мм", "consumption": 7.4, "price": 59.02},
        ],
    },
    {
        "no": 3,
        "product": "План установки крутильной машины MKD 6-12-18/500",
        "shop": "3",
        "drawing": "ПТ 12835.00.00 СБ",
        "sub_part": "Болты",
        "time_min": 400,
        "qty": 36,
        "materials": [
            {"unit": "кг", "name": "Круг Ст.3 24мм", "consumption": 170.9, "price": 91.8},
            {"unit": "кг", "name": "Круг Ст.3 20мм", "consumption": 11.8, "price": 81.54},
            {"unit": "кг", "name": "Круг Ст.3 16мм", "consumption": 5.2, "price": 91.8},
        ],
    },
]

sign = {
    "approver_post": "Главный механик",
    "approver": "Бурыгин Е.Е.",
    "prepared": "Дроздов И.А.",
    "checked": "Сомкин М.Е.",
}

build_kalkulyatsiya(
    path="Калькуляция_тест.xlsx",
    year=2026,
    month=2,
    groups=groups,
    sign=sign,
    rate_per_min=9.18,
)
print("Готово: Калькуляция_тест.xlsx")
