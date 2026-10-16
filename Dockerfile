
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt \
    && playwright install --with-deps chromium

COPY . .

CMD ["python", "app.py"]
