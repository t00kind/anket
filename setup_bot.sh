#!/bin/bash

# Переход в директорию проекта
cd /home/satoru/anketTgBot || exit

# Проверка и установка python3-venv при необходимости
if ! dpkg -s python3-venv >/dev/null 2>&1; then
  echo "📦 Устанавливается python3-venv..."
  sudo apt update && sudo apt install -y python3-venv
fi

# Создание виртуального окружения, если не существует
if [ ! -d "venv" ]; then
  echo "🌀 Создаётся виртуальное окружение..."
  python3 -m venv venv
fi

# Активация окружения и установка зависимостей
echo "📦 Устанавливаются зависимости из requirements.txt..."
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
deactivate

# Запуск через PM2
echo "🚀 Запускается бот через pm2..."
pm2 start main.py --interpreter ./venv/bin/python --name anketbot
pm2 save

# Настройка автозапуска
echo "🔧 Настраивается автозапуск pm2..."
pm2 startup | tail -n 1 | bash

echo "✅ Всё готово. Бот должен работать и перезапускаться автоматически."