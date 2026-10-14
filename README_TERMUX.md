# Калькуляция — запуск сервера на Android через Termux

Веб-сервер калькуляции (Flask) оптимизирован для работы прямо на телефоне
в [Termux](https://termux.dev). Данные (SQLite + справочник материалов)
автоматически размещаются в доступном для записи каталоге.

## Быстрый старт

```bash
pkg update -y
pkg install git python -y
git clone <репозиторий> && cd <репозиторий>
bash setup_termux.sh      # установка зависимостей в .venv
bash run_termux.sh        # запуск сервера
```

В браузере телефона открыть: **http://localhost:5000**
С других устройств в Wi-Fi-сети: адрес, который сервер печатает при старте
(например `http://192.168.1.47:5000`).

## Что сделано для Termux

| Проблема на Android | Решение |
|---|---|
| `$HOME` недоступен на запись без `termux-setup-storage` → SQLite падал с `attempt to write a readonly database` | Приложение само определяет Termux (`$PREFIX`) и хранит `calc.db` / `materials.xlsx` в private-каталоге `/data/data/com.termux/.../app_private_files/calc_data`, куда можно писать всегда; существующие файлы из `data/` копируются туда автоматически |
| `debug=True` — перезагрузчик файлов, консоль отладки с `exec()` (дыра безопасности), лишний расход памяти/заряда | debug выключен по умолчанию; включается явно: `DEBUG=1 python app.py`. `PORT`/`HOST` задаются через переменные окружения |
| Android убивает фоновый процесс при блокировке экрана | `run_termux.sh` вызывает `termux-wake-lock`; сам сервер тоже пытается его включить при старте (нужны пакет `termux-api` + приложение Termux:API из F-Droid — опционально) |
| Нет tmux-сессии → сервер умирает при закрытии терминала | `bash run_termux.sh --daemon` — запуск в фоне через tmux (или nohup + `server.log`) |
| Долгая сборка нативных модулей на телефоне | Минимальные зависимости: только Flask и openpyxl (`requirements.txt`); lxml не требуется |
| Неясно, какой адрес открыть | Сервер при старте печатает готовый URL с локальным IP и подсказку `http://localhost:PORT` |

## Полезные команды

```bash
# фоновый режим
bash run_termux.sh --daemon          # старт в tmux-сессии "calc"
tmux attach -t calc                  # посмотреть логи
tmux kill-session -t calc            # остановить

# другой порт (например, 8080)
PORT=8080 bash run_termux.sh

# автозапуск при подключении зарядки (через Termux:Boot):
#   pkg install termux-services  либо добавить в ~/.termux/boot/start-calc:
mkdir -p ~/.termux/boot
cat > ~/.termux/boot/start-calc <<'EOF'
#!/data/data/com.termux/files/usr/bin/bash
termux-wake-lock
cd ~/калькуляция-проект && bash run_termux.sh --daemon
EOF
chmod +x ~/.termux/boot/start-calc
```

Требования: Termux из **F-Droid/GitHub** (не из Google Play — там устаревшая
и сломанная версия), Android 7+, ~100 МБ свободного места.
