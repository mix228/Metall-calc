# -*- coding: utf-8 -*-

import os
import re
import math
import logging
import asyncio

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart, Command
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup


# ============================================================
# НАСТРОЙКИ
# ============================================================

# Сюда вставьте токен от BotFather
BOT_TOKEN = os.getenv("BOT_TOKEN", "")

# Временный тестовый справочник материалов.
# Позже заменим его на импорт из Материалы.xls / базу данных.
MATERIALS = {
    "сталь 45": {
        "density": 7850,
        "price_per_kg": 85.0,
    },
    "сталь 3": {
        "density": 7850,
        "price_per_kg": 70.0,
    },
    "сталь 20": {
        "density": 7850,
        "price_per_kg": 75.0,
    },
    "сталь 09г2с": {
        "density": 7850,
        "price_per_kg": 80.0,
    },
    "нержавейка": {
        "density": 7900,
        "price_per_kg": 250.0,
    },
    "алюминий": {
        "density": 2700,
        "price_per_kg": 320.0,
    },
    "латунь": {
        "density": 8500,
        "price_per_kg": 600.0,
    },
    "медь": {
        "density": 8900,
        "price_per_kg": 900.0,
    },
    "капролон": {
        "density": 1140,
        "price_per_kg": 400.0,
    },
    "текстолит": {
        "density": 1350,
        "price_per_kg": 600.0,
    },
}

# Временное хранилище неподтверждённых записей.
# Для теста достаточно. Позже заменим на нормальную базу/состояние.
pending_records = {}


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def normalize_text(text: str) -> str:
    """
    Приводим текст к нижнему регистру, меняем ё на е,
    убираем лишние пробелы.
    """
    text = text.lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ============================================================
# ПАРСЕРЫ
# ============================================================

def parse_drawing(text: str):
    """
    Ищем номер чертежа.

    Поддерживаем варианты:
    - чертёж 125
    - чертеж 125
    - черт. 125
    - черт 125
    - №125
    - кд 125
    """
    m = re.search(
        r"(?:чертеж|черт\.?|кд|№)\s*:?\s*(\d+)",
        text,
        flags=re.IGNORECASE,
    )
    if m:
        return m.group(1)
    return None


def parse_profile_type(text: str):
    """
    Определяем тип заготовки.

    Пока поддерживаем только круг.
    """
    if re.search(r"\b(?:круг|пруток|круглый)\b", text):
        return "round"

    return None


def parse_material(text: str):
    """
    Ищем материал.

    Поддерживаем варианты:
    - сталь 45
    - ст45
    - ст.45
    - круг 45
    - капролон
    - текстолит
    - алюминий
    """

    # Сначала ищем неметаллические материалы / цветмет
    m = re.search(
        r"\b(капролон|текстолит|алюминий|латунь|медь|нержавейка)\b",
        text,
    )
    if m:
        return m.group(1)

    # Варианты: сталь 45, ст45, ст.45, сталь3
    m = re.search(
        r"\b(?:сталь|ст\.?)\s*([0-9]{1,3}|[0-9а-з]+)",
        text,
    )
    if m:
        return f"сталь {m.group(1)}"

    # Вариант: круг 45, если после числа нет явного размера через х/на
    m = re.search(
        r"\bкруг\s+([0-9]{1,3})(?!\s*(?:х|x|×|\*|на|–|-))\b",
        text,
    )
    if m:
        return f"сталь {m.group(1)}"

    return None


def parse_quantity(text: str):
    """
    Ищем количество.

    Поддерживаем:
    - 3 шт
    - 3 шт.
    - 3 штуки
    - количество 3
    - кол-во 3
    - три штуки
    """

    # Числовой вариант
    m = re.search(
        r"(\d+)\s*(?:шт\.?|штук|штуки|шт)",
        text,
    )
    if m:
        return int(m.group(1))

    # Количество: 3
    m = re.search(
        r"(?:количество|кол-во|кол\.?)\s*:?\s*(\d+)",
        text,
    )
    if m:
        return int(m.group(1))

    # Словами: три штуки
    words = {
        "один": 1,
        "одна": 1,
        "одно": 1,
        "два": 2,
        "две": 2,
        "три": 3,
        "четыре": 4,
        "пять": 5,
        "шесть": 6,
        "семь": 7,
        "восемь": 8,
        "девять": 9,
        "десять": 10,
    }

    m = re.search(
        r"\b(один|одна|одно|два|две|три|четыре|пять|шесть|семь|восемь|девять|десять)\s*(?:шт\.?|штук|штуки|шт)",
        text,
    )
    if m:
        return words.get(m.group(1))

    return None


def to_mm(value, unit=None):
    """
    Приводим размер к миллиметрам.
    """
    val = float(str(value).replace(",", "."))

    if unit:
        unit = unit.lower()

        if unit in {"м", "m", "метр", "метра", "метров"}:
            return val * 1000

        if unit in {"см", "cm"}:
            return val * 10

    return val


def parse_round_dimensions(text: str):
    """
    Ищем диаметр и длину для круга.

    Поддерживаем:
    - 60х500
    - 60x500
    - 60 на 500
    - диаметр 60, длина 500
    - Ø60 × 500
    - 60 мм х 500 мм
    - длина 0.5 м
    """

    dims = {}

    # Пара размеров: 60х500, 60 на 500, 60*500
    pair = re.search(
        r"(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?\s*(?:х|x|×|\*|на|–|-)\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text,
        flags=re.IGNORECASE,
    )

    if pair:
        dims["diameter_mm"] = to_mm(pair.group(1), pair.group(2))
        dims["length_mm"] = to_mm(pair.group(3), pair.group(4))

    # Явный диаметр
    m = re.search(
        r"(?:диаметр|диам\.?|d|∅|ø)\s*:?\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text,
        flags=re.IGNORECASE,
    )
    if m:
        dims["diameter_mm"] = to_mm(m.group(1), m.group(2))

    # Явная длина
    m = re.search(
        r"(?:длина|дл\.?|l)\s*:?\s*(\d+(?:[.,]\d+)?)\s*(мм|см|м|mm|m)?",
        text,
        flags=re.IGNORECASE,
    )
    if m:
        dims["length_mm"] = to_mm(m.group(1), m.group(2))

    # Если длина была указана отдельно в метрах: 6 м, 0.5 м
    if "length_mm" not in dims:
        m = re.search(
            r"\b(\d+(?:[.,]\d+)?)\s*(?:м|метр|метра|метров)\b",
            text,
            flags=re.IGNORECASE,
        )
        if m:
            dims["length_mm"] = to_mm(m.group(1), "м")

    return dims


# ============================================================
# РАСЧЁТ
# ============================================================

def calc_round_mass_kg(diameter_mm: float, length_mm: float, density: float) -> float:
    """
    Масса круга.

    диаметр в мм,
    длина в мм,
    плотность в кг/м³.
    """
    volume_mm3 = math.pi / 4 * diameter_mm ** 2 * length_mm
    volume_m3 = volume_mm3 / 1_000_000_000

    return volume_m3 * density


# ============================================================
# ФОРМАТИРОВАНИЕ СООБЩЕНИЙ
# ============================================================

def format_record(data: dict) -> str:
    return "\n".join([
        f"Чертёж: {data['drawing']}",
        f"Материал: {data['material_name']}",
        f"Заготовка: круг Ø{data['diameter_mm']:g} × {data['length_mm']:g} мм",
        f"Количество: {data['quantity']} шт.",
        f"Масса заготовки: {data['mass_one']:.3f} кг",
        f"Общая масса: {data['mass_total']:.3f} кг",
        f"Цена материала: {data['price_per_kg']:.2f} руб/кг",
        f"Стоимость материала: {data['material_cost']:.2f} руб",
    ])


def format_preview(data: dict) -> str:
    return format_record(data) + "\n\nЗаписать?"


def make_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="Да", callback_data="confirm"),
                InlineKeyboardButton(text="Нет", callback_data="cancel"),
            ]
        ]
    )


# ============================================================
# ЛОГИКА ОБРАБОТКИ ТЕКСТА
# ============================================================

async def process_text(message: types.Message, raw_text: str):
    text = normalize_text(raw_text)

    drawing = parse_drawing(text)
    profile_type = parse_profile_type(text)
    material_name = parse_material(text)
    quantity = parse_quantity(text)

    warnings = []
    errors = []

    if not drawing:
        errors.append("Не найден номер чертежа.")

    if profile_type != "round":
        errors.append("В тестовой версии пока поддерживается только круг.")

    if not material_name:
        errors.append("Не найден материал.")

    if not quantity:
        quantity = 1
        warnings.append("Количество не указано. Использую 1 шт.")

    dims = {}
    if profile_type == "round":
        dims = parse_round_dimensions(text)

    if "diameter_mm" not in dims or "length_mm" not in dims:
        errors.append("Не найдены размеры. Нужны диаметр и длина.")

    material = MATERIALS.get(material_name)

    if material_name and not material:
        errors.append(
            f"Материал '{material_name}' распознан, но в тестовом справочнике нет цены/плотности."
        )

    if errors:
        answer = []
        answer.append("Не получилось полностью распознать запрос.")
        answer.append("")
        answer.append("Ошибки:")
        for err in errors:
            answer.append(f"- {err}")

        if warnings:
            answer.append("")
            answer.append("Предупреждения:")
            for warn in warnings:
                answer.append(f"- {warn}")

        answer.append("")
        answer.append("Пример правильного ввода:")
        answer.append("Чертёж 125, круг, сталь 45, 60 на 500, 3 шт")

        await message.answer("\n".join(answer))
        return

    mass_one = calc_round_mass_kg(
        diameter_mm=dims["diameter_mm"],
        length_mm=dims["length_mm"],
        density=material["density"],
    )

    mass_total = mass_one * quantity
    material_cost = mass_total * material["price_per_kg"]

    record = {
        "drawing": drawing,
        "material_name": material_name,
        "diameter_mm": dims["diameter_mm"],
        "length_mm": dims["length_mm"],
        "quantity": quantity,
        "mass_one": mass_one,
        "mass_total": mass_total,
        "price_per_kg": material["price_per_kg"],
        "material_cost": material_cost,
    }

    pending_records[message.chat.id] = record

    preview = format_preview(record)

    if warnings:
        preview = "Предупреждения:\n" + "\n".join(f"- {w}" for w in warnings) + "\n\n" + preview

    await message.answer(
        preview,
        reply_markup=make_keyboard(),
    )


# ============================================================
# TELEGRAM BOT
# ============================================================

dp = Dispatcher()


@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    await message.answer(
        "Привет! Это тестовый бот расчёта материалов.\n\n"
        "Отправьте мне фразу, например:\n"
        "Чертёж 125, круг, сталь 45, 60 на 500, 3 шт\n\n"
        "Или:\n"
        "Чертёж 125. Заготовка — круг, сталь 45, диаметр 60, длина 500, три штуки"
    )


@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    await message.answer(
        "Формат ввода:\n"
        "Чертёж 125, круг, сталь 45, 60 на 500, 3 шт\n\n"
        "Пока тестовая версия понимает только круг.\n"
        "После расчёта нужно подтвердить запись кнопками Да/Нет."
    )


@dp.message(Command("add"))
async def cmd_add(message: types.Message):
    """
    Команда вида:
    /add Чертёж 125, круг, сталь 45, 60х500, 3 шт
    """
    parts = message.text.split(maxsplit=1)

    if len(parts) < 2:
        await message.answer(
            "Используйте так:\n"
            "/add Чертёж 125, круг, сталь 45, 60х500, 3 шт"
        )
        return

    await process_text(message, parts[1])


@dp.message(F.text)
async def any_text(message: types.Message):
    """
    Обычный текстовый ввод.
    """
    await process_text(message, message.text)


@dp.callback_query(F.data == "confirm")
async def confirm_record(callback: types.CallbackQuery):
    if not callback.message:
        await callback.answer()
        return

    record = pending_records.pop(callback.message.chat.id, None)

    if not record:
        await callback.answer("Нет активного подтверждения.", show_alert=True)
        return

    text = (
        "✅ Тестовая запись подтверждена.\n"
        "На следующем этапе здесь будет реальная запись в Excel.\n\n"
        + format_record(record)
    )

    await callback.message.edit_text(text)
    await callback.answer("Подтверждено")


@dp.callback_query(F.data == "cancel")
async def cancel_record(callback: types.CallbackQuery):
    if not callback.message:
        await callback.answer()
        return

    pending_records.pop(callback.message.chat.id, None)

    await callback.message.edit_text("❌ Запись отменена.")
    await callback.answer("Отменено")


# ============================================================
# ЗАПУСК
# ============================================================

async def main():
    logging.basicConfig(level=logging.INFO)

    bot = Bot(token=BOT_TOKEN)

    await dp.start_polling(bot)


if __name__ == "__main__":
    if BOT_TOKEN == "ВСТАВЬТЕ_ТОКЕН":
        raise SystemExit(
            "Ошибка: не указан токен бота.\n"
            "Вставьте токен от BotFather в переменную BOT_TOKEN."
        )

    asyncio.run(main())
