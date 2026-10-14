<div align="center">

# 🔧 Калькулятор материалов

**Веб-приложение для расчёта стоимости материалов и подготовки калькуляций**

_Работает на ПК и прямо на Android-телефоне через Termux · Мобильная версия интерфейса · Без облака — все данные у вас_

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-2.x-000?style=flat-square&logo=flask&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-3-003b57?style=flat-square&logo=sqlite&logoColor=white)
![Excel](https://img.shields.io/badge/XLSX-openpyxl-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Android](https://img.shields.io/badge/Termux-ready-333?style=flat-square&logo=android&logoColor=white)
![Mobile](https://img.shields.io/badge/mobile-responsive-6C5CE7?style=flat-square)

</div>

---

## ✨ Возможности

| | |
|---|---|
| 🧮 **Быстрый расчёт** | Ввод размеров и материала — стоимость считается мгновенно |
| 📚 **Справочник материалов** | Добавление, редактирование и удаление материалов (плотность, цена за кг) |
| 🗂 **История калькуляций** | Все расчёты сохраняются в SQLite, можно подтвердить или удалить запись |
| 📄 **Экспорт в Excel** | Формирование файла «Калькуляция» в формате, близком к оригинальному документу |
| 📱 **Мобильный интерфейс** | Адаптивная вёрстка: крупные touch-кнопки, карточки вместо таблиц, без автозума на iOS |
| 🤖 **Telegram-бот** | Отдельный режим работы через бота (`bot.py`, aiogram) |
| 🔒 **Локальные данные** | Никакого облака — база и справочник хранятся в папке проекта |

---

## 📁 Структура проекта

```
.
├── app.py                        # Flask-сервер: расчёты, история, справочник
├── db.py                         # Работа с SQLite (история калькуляций)
├── doc_builder.py                # Генератор Excel-файла «Калькуляция»
├── bot.py                        # Telegram-бот (aiogram) — отдельный режим
├── make_sample_doc.py            # Утилита: тестовый Калькуляция_тест.xlsx
├── make_materials_template.py    # Утилита: шаблон data/materials.xlsx
│
├── templates/
│   ├── index.html                # Страница калькулятора
│   └── materials.html            # Страница справочника материалов
├── static/
│   ├── style.css / theme.css     # Стили + светлая/тёмная тема
│   └── theme.js                  # Переключение темы
│
├── data/
│   ├── calc.db                   # База истории калькуляций (SQLite)
│   └── materials.xlsx            # Справочник материалов
│
├── requirements.txt              # Зависимости (Flask, openpyxl)
├── setup_termux.sh               # Установка на Android (Termux)
├── run_termux.sh                 # Запуск на Android (+ --daemon)
└── README_TERMUX.md              # Подробное руководство по Termux
```

**Как связаны модули:** `app.py` — центральная точка входа; он использует `db.py`
(`data/calc.db`) и справочник `data/materials.xlsx`, рендерит `templates/*` со
стилями из `static/`. `make_sample_doc.py → doc_builder.py` генерирует Excel.
`bot.py` работает автономно и серверу не нужен.

---

## 🚀 Запуск на компьютере (ПК / Mac / Linux)

### 1. Установите Python

Нужен **Python 3.8+**. Проверьте:

```bash
python3 --version
```

### 2. Скачайте проект

```bash
git clone <адрес вашего репозитория> && cd <папка проекта>
```

(или просто распакуйте архив с проектом и перейдите в его папку)

### 3. Установите зависимости

<details open>
<summary><b>Windows</b></summary>

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```
</details>

<details>
<summary><b>macOS / Linux</b></summary>

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```
</details>

### 4. Запустите сервер

```bash
python app.py
```

Откройте в браузере: **http://localhost:5000**

Чтобы пользоваться с других устройств в вашей Wi-Fi сети, сервер при старте
печатает адрес вида `http://192.168.x.x:5000` — откройте его на телефоне.

<details>
<summary><b>Необязательные настройки запуска</b></summary>

```bash
PORT=8080 python app.py       # другой порт
HOST=0.0.0.0 python app.py    # доступ из локальной сети (по умолчанию уже открыт)
DEBUG=1 python app.py         # режим отладки (только для разработки!)
```
</details>

---

## 📱 Запуск на телефоне (Android + Termux)

Приложение специально оптимизировано для работы прямо на Android:
данные автоматически размещаются в доступном для записи каталоге,
экран не «засыпает» во время работы сервера, а интерфейс адаптирован под палец.

### Шаг 1 — Установите Termux

⚠️ Берите Termux из **F-Droid** или **GitHub**, а **не из Google Play**
(там устаревшая нерабочая версия).

👉 [F-Droid: Termux](https://f-droid.org/packages/com.termux/)

### Шаг 2 — Установите и запустите проект

```bash
pkg update -y
pkg install git python -y

git clone <адрес вашего репозитория>
cd <папка проекта>

bash setup_termux.sh     # создаёт .venv и ставит зависимости
bash run_termux.sh       # запускает сервер
```

### Шаг 3 — Откройте приложение

В браузере телефона: **http://localhost:5000**

### Фоновый режим (чтобы сервер жил после закрытия терминала)

```bash
bash run_termux.sh --daemon      # запуск в tmux/nohup
tmux attach -t calc              # посмотреть логи
tmux kill-session -t calc        # остановить
```

### Автозапуск при подключении зарядки (Termux:Boot)

```bash
mkdir -p ~/.termux/boot
cat > ~/.termux/boot/start-calc <<'EOF'
#!/data/data/com.termux/files/usr/bin/bash
termux-wake-lock
cd ~/<папка проекта> && bash run_termux.sh --daemon
EOF
chmod +x ~/.termux/boot/start-calc
```

> 📖 Полное описание всех тонкостей Termux-версии (wakelock, private-каталог
> данных, порты, безопасность) — в **[README_TERMUX.md](README_TERMUX.md)**.

---

## ⚙️ Как это работает под капотом

```
Браузер (ПК / телефон)
        │  HTTP
        ▼
   app.py (Flask) ──────► templates/*.html + static/*  (интерфейс)
        │
        ├───────────────► data/calc.db        (история — SQLite, db.py)
        ├───────────────► data/materials.xlsx (справочник — openpyxl)
        └───────────────► doc_builder.py      (экспорт калькуляции в Excel)
```

* На старте приложение определяет окружение: в **Termux** каталог данных
  автоматически переносится в private-папку приложения (в `$HOME` запись может
  быть запрещена), на **ПК** используется обычная папка `data/`.
* Справочник материалов сохраняется **атомарно** (временный файл + `os.replace`)
  — обрыв процесса не может повредить `materials.xlsx`.
* `debug`-режим (перезагрузчик и консоль с `exec()`) по умолчанию **выключен** —
  безопасно для запуска в локальной сети.

---

## 🤖 Telegram-бот (опционально)

`bot.py` — самостоятельный режим расчёта в Telegram (aiogram).

```bash
pip install aiogram
BOT_TOKEN="ваш_токен_от_BotFather" python bot.py
```

> Токен задаётся через переменную окружения `BOT_TOKEN` — не храните его в коде.

---

## ❓ Частые проблемы

| Проблема | Решение |
|---|---|
| Порт 5000 занят | `PORT=8080 python app.py` (на iPhone Safari порт 5000 иногда блокируется) |
| Телефон не видит сервер ПК | Убедитесь, что устройства в одной Wi-Fi сети, и что файрвол разрешает выбранный порт |
| Termux: `attempt to write a readonly database` | Обновите код — приложение само выбирает private-каталог; либо выполните `termux-setup-storage` |
| Сервер умер при закрытии терминала | Используйте `bash run_termux.sh --daemon` и пакет `termux-wake-lock`/`termux-api` |
| `pip install` идёт долго на телефоне | Это нормально: используются только лёгкие чистые Python-пакеты (Flask, openpyxl) — нативные сборки не нужны |

---

## 📄 Лицензия

Используйте свободно в личных и рабочих целях.

<div align="center">

Сделано для тех, кто считает материалы — за столом 💻 и на ходу 📱

</div>
