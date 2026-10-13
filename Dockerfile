FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py telegram_alerts.py telegram_setup.py container_entrypoint.py ./
RUN mkdir -p /app/input /app/output /app/templates

ENTRYPOINT ["python", "container_entrypoint.py"]
