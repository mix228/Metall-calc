#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
# Установка проекта "Калькуляция" в Termux (Android)
# Использование:  bash setup_termux.sh
# ============================================================
set -e
cd "$(dirname "$0")"

echo "==> Обновляем пакеты Termux..."
pkg update -y

echo "==> Устанавливаем Python, git и зависимости сборки..."
pkg install -y python git build-essential

# termux-api опционален: даёт termux-wake-lock (сервер не убьёт Android в спящем режиме)
if pkg list-installed 2>/dev/null | grep -q '^termux-api '; then
    echo "==> termux-api установлен — wakelock будет включаться автоматически."
else
    echo "==> (необязательно) Для защиты сервера от засыпания телефона:"
    echo "      pkg install termux-api   +  приложение Termux:API из F-Droid"
fi

echo "==> Создаём виртуальное окружение .venv..."
python -m venv .venv 2>/dev/null || {
    # на некоторых устройствах без ensurepip
    pkg install -y python-pip 2>/dev/null || true
    python -m venv --without-pip .venv && \
        curl -sS https://bootstrap.pypa.io/get-pip.py | .venv/bin/python
}

echo "==> Ставим зависимости (Flask, openpyxl)..."
# lxml ставим только если соберётся за разумное время — на слабых телефонах
# это самая долгая позиция; без него всё работает (у openpyxl он опционален)
.venv/bin/pip install --upgrade pip wheel setuptools
echo "==> Flask + openpyxl (lxml опционально, пропуск при ошибке)..."
.venv/bin/pip install -r requirements.txt || .venv/bin/pip install Flask openpyxl

echo ""
echo "============================================================"
echo " Готово! Запуск сервера:   bash run_termux.sh"
echo " В браузере телефона:     http://localhost:5000"
echo "============================================================"
