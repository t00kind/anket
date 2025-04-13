# Используем минимальный образ Python
FROM python:3.12-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем зависимости
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Копируем все остальные файлы
COPY . .

# Указываем переменные среды (рекомендуется задавать через .env или docker-compose)
ENV PYTHONUNBUFFERED=1

# Запуск скрипта
CMD ["python", "main.py"]
