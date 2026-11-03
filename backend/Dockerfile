FROM python:3.12-slim-bookworm

# LibreOffice is the essential part. The final Word is saved by LibreOffice,
# NOT downloaded directly from the browser's OOXML composer.
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    SAL_USE_VCLPLUGIN=svp

RUN apt-get update \
 && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
       libreoffice-writer libreoffice-core fonts-dejavu-core fonts-liberation \
       ca-certificates \
 && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY backend/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt
COPY backend/main.py /app/main.py
RUN useradd --create-home --uid 10001 zainnet
USER zainnet
EXPOSE 10000
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000} --workers 1"]
