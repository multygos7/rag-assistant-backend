FROM python:3.11-slim

WORKDIR /app

# Системные зависимости для pdfplumber/lxml и т.п.
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 40010

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:40010", "--timeout", "120"]
