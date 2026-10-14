#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
# Запуск сервера калькуляции в Termux (Android)
# Использование:  bash run_termux.sh            — запуск в этой сессии
#                 bash run_termux.sh --daemon   — в фоне (tmux), Ctrl+C не нужен
# ============================================================
cd "$(dirname "$0")"

VENV=".venv/bin/python"
[ -x "$VENV" ] || VENV="python"   # если setup_termux.sh не запускался

# wakelock: не даём телефону убить сервер при блокировке экрана
command -v termux-wake-lock >/dev/null 2>&1 && termux-wake-lock

if [ "$1" = "--daemon" ]; then
    if command -v tmux >/dev/null 2>&1; then
        tmux new-session -d -s calc "bash '$PWD/run_termux.sh'" 2>/dev/null || \
            tmux attach -t calc
        echo "Сервер запущен в фоне (tmux-сессия 'calc')."
        echo "Вернуться к логам:  tmux attach -t calc"
        echo "Остановить:         tmux kill-session -t calc"
        exit 0
    else
        echo "Нет tmux (pkg install tmux) — запускаю в фоне через nohup..."
        nohup "$VENV" app.py > server.log 2>&1 &
        echo "PID: $!  |  Логи: tail -f server.log  |  Стоп: kill $!"
        exit 0
    fi
fi

echo "Запуск. Адрес — на экране ниже. Остановка: Ctrl+C"
exec "$VENV" app.py
