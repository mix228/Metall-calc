# -*- coding: utf-8 -*-
"""
Генератор документа «Калькуляция» в формате предприятия.
Версия 2: оформление приближено к оригиналу Калькуляция №1.xlsx:
  - блок «УТВЕРЖДАЮ» — справа вверху;
  - заголовок — по центру;
  - «Итого:» — у предпоследнего столбца, формат 0.00;
  - подписи «Подготовил/Проверил» — справа внизу;
  - формулы: материал=расход*цена; работы=время*ставка;
    деталь=СУММ(материалы группы)+работа; общая=деталь*кол-во;
    Итого=СУММЕСЛИ(служебная "+"; общая).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MONTHS_NOM = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}
MONTHS_GEN = {
    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
    5: "май", 6: "июнь", 7: "июль", 8: "август",
    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь",
}

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BASE_HEADERS = [
    "№ п/п",
    "Наименование изделия",
    "№ Цеха",
    "Наименование чертежа",
    "Ед. изм.",
    "Наименование материала",
    "Расход материала 1шт. (кг./м.)",
    "Цена за 1 кг. (руб.)",
    "Стоимость материала 1 шт. (руб.) (Цзаг.=Кзаг*Цкг)",
    "Время изготовления (мин. за 1шт.) (t)",
    "Стоимость работ (Цшт=t*НДС)",
    "Стоимость детали      (руб за шт.)",
    "Кол-во шт. (К)",
    "Общая стоимость Цобщ=(Цшт*К+Цзаг.)",
]


def build_kalkulyatsiya(path, year, month, groups, sign, rate_per_min,
                        with_subpart=None, reserve_rows=10):
    """
    groups: список групп (изделий). Группа:
      {
        "no": 1, "product": "...", "shop": "3", "drawing": "...",
        "sub_part": "Болты",        # может быть ""
        "time_min": 960, "qty": 1,
        "materials": [
            {"unit": "кг", "name": "...", "consumption": 7.4, "price": 59.02},
        ],
      }
    sign: {"approver_post", "approver", "prepared", "checked"}
    """
    if with_subpart is None:
        with_subpart = any((g.get("sub_part") or "") for g in groups)

    headers = list(BASE_HEADERS)
    if with_subpart:
        headers.insert(4, "Наименование чертежа ")

    off = 1 if with_subpart else 0
    col = {
        "no": 1, "product": 2, "shop": 3, "drawing": 4,
        "sub": 5 if with_subpart else None,
        "unit": 5 + off, "material": 6 + off, "consum": 7 + off,
        "price": 8 + off, "mat_cost": 9 + off, "time": 10 + off,
        "work": 11 + off, "detail": 12 + off, "qty": 13 + off,
        "total": 14 + off, "helper": 15 + off,
    }
    CL = lambda k: get_column_letter(col[k])
    ncols = col["total"]

    wb = Workbook()
    ws = wb.active
    ws.title = f"{MONTHS_NOM[month]} {year}г"

    def merge(row, text, bold=False, align="left", c1=None, c2=None):
        a = c1 if c1 else 1
        b = c2 if c2 else ncols
        if b > a:
            ws.merge_cells(start_row=row, start_column=a, end_row=row, end_column=b)
        c = ws.cell(row=row, column=a, value=text)
        c.font = Font(bold=bold)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)

    money_fmt = "0.00"
    rate = round(rate_per_min, 2)

    # ---------- блок утверждения: справа вверху ----------
    r = 1
    merge(r, "УТВЕРЖДАЮ", c1=ncols - 2, c2=ncols, align="right"); r += 1
    merge(r, sign.get("approver_post", "Главный механик"), c1=ncols - 2, c2=ncols, align="right"); r += 1
    merge(r, f"______________{sign.get('approver', '')}", c1=ncols - 2, c2=ncols, align="right"); r += 1
    merge(r, f'"____"_____{year}г.', c1=ncols - 2, c2=ncols, align="right"); r += 2

    # ---------- заголовок: по центру ----------
    merge(r, f"Калькуляция  за {MONTHS_GEN[month]} {year}г.", bold=True, align="center"); r += 1

    # ---------- шапка таблицы ----------
    header_row = r
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[header_row].height = 45
    r += 1

    # ---------- данные ----------
    data_first = r

    for g in groups:
        mats = g["materials"]
        g_first = r
        g_last = r + len(mats) - 1
        for mi, m in enumerate(mats):
            ws[f"{CL('no')}{r}"] = g["no"]
            ws[f"{CL('product')}{r}"] = g.get("product", "")
            ws[f"{CL('shop')}{r}"] = g.get("shop", "")
            ws[f"{CL('drawing')}{r}"] = g.get("drawing", "")
            if with_subpart:
                ws[f"{CL('sub')}{r}"] = g.get("sub_part", "")
            ws[f"{CL('unit')}{r}"] = m.get("unit", "кг")
            ws[f"{CL('material')}{r}"] = m.get("name", "")
            ws[f"{CL('consum')}{r}"] = m.get("consumption")
            ws[f"{CL('price')}{r}"] = m.get("price")
            ws[f"{CL('mat_cost')}{r}"] = f"={CL('consum')}{r}*{CL('price')}{r}"
            ws[f"{CL('time')}{r}"] = g.get("time_min")
            ws[f"{CL('work')}{r}"] = f"={CL('time')}{r}*{rate}"
            ws[f"{CL('detail')}{r}"] = (
                f"=SUM({CL('mat_cost')}{g_first}:{CL('mat_cost')}{g_last})"
                f"+{CL('work')}{g_first}"
            )
            ws[f"{CL('qty')}{r}"] = g.get("qty", 1)
            ws[f"{CL('total')}{r}"] = f"={CL('detail')}{r}*{CL('qty')}{r}"
            if mi == 0:
                ws[f"{CL('helper')}{r}"] = "+"
            r += 1

    # резервные пустые строки с формулами (0.00, как в оригинале)
    for _ in range(reserve_rows):
        ws[f"{CL('mat_cost')}{r}"] = f"={CL('consum')}{r}*{CL('price')}{r}"
        ws[f"{CL('work')}{r}"] = f"={CL('time')}{r}*{rate}"
        ws[f"{CL('detail')}{r}"] = f"={CL('mat_cost')}{r}+{CL('work')}{r}"
        ws[f"{CL('total')}{r}"] = f"={CL('detail')}{r}*{CL('qty')}{r}"
        r += 1
    data_last = r - 1

    # ---------- Итого: метка у предпоследнего столбца ----------
    total_row = r
    lc = ws.cell(row=total_row, column=ncols - 1, value="Итого:")
    lc.font = Font(bold=True)
    lc.alignment = Alignment(horizontal="left", vertical="center")
    tc = ws.cell(
        row=total_row,
        column=ncols,
        value=(
            f'=SUMIF({CL("helper")}{data_first}:{CL("helper")}{data_last},"+",'
            f'{CL("total")}{data_first}:{CL("total")}{data_last})'
        ),
    )
    tc.font = Font(bold=True)
    tc.number_format = money_fmt
    r += 2

    # ---------- ставка минуты: слева ----------
    rate_text = (
        f"Стоимость работы за минуту без НДС "
        f"{round(rate * 60):.0f}/60={str(rate).replace('.', ',')}р/мин"
    )
    merge(r, rate_text); r += 2

    # ---------- подписи: справа внизу ----------
    merge(r, f"Подготовил___________________________{sign.get('prepared', '')}",
          c1=max(1, ncols - 3), c2=ncols, align="left"); r += 1
    merge(r, f"Проверил_____________________________{sign.get('checked', '')}",
          c1=max(1, ncols - 3), c2=ncols, align="left")

    # ---------- оформление таблицы ----------
    for row in range(header_row, data_last + 1):
        for i in range(1, ncols + 2):
            ws.cell(row=row, column=i).border = BORDER
        for k in ("mat_cost", "work", "detail", "total"):
            ws[f"{CL(k)}{row}"].number_format = money_fmt
        for k in ("no", "shop", "unit", "consum", "price", "time", "qty", "helper"):
            ws[f"{CL(k)}{row}"].alignment = Alignment(horizontal="center")
        for k in ("product", "drawing", "sub", "material"):
            if col[k]:
                ws[f"{CL(k)}{row}"].alignment = Alignment(horizontal="left", wrap_text=True)

    widths = {
        "no": 6, "product": 28, "shop": 8, "drawing": 22, "sub": 14,
        "unit": 8, "material": 30, "consum": 10, "price": 10,
        "mat_cost": 12, "time": 10, "work": 12, "detail": 12,
        "qty": 8, "total": 14, "helper": 4,
    }
    for k, w in widths.items():
        if col[k]:
            ws.column_dimensions[CL(k)].width = w

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(path)
    return path