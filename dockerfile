# Use the minimal Python image
FROM python:3.12-slim

# Set the working directory
WORKDIR /app

# Copy dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy all other files
COPY . .

# Set environment variables (it's recommended to specify them via .env or docker-compose)
ENV PYTHONUNBUFFERED=1

# Run the script
CMD ["python", "main.py"]
