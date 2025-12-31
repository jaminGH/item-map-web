# syntax=docker/dockerfile:1
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=on \
    PIP_NO_CACHE_DIR=1 \
    ITEMMAP_DATA_DIR=/app/data \
    FLASK_SECRET_KEY=change-me \
    PORT=8000

WORKDIR /app

# System deps (if any)
RUN apt-get update -y && apt-get install -y --no-install-recommends \
    build-essential \
  && rm -rf /var/lib/apt/lists/*

# Requirements
COPY requirements.txt ./
RUN pip install -r requirements.txt

# App code
COPY webtool ./webtool
COPY wsgi.py ./

# Data directories (bind-mount recommended)
RUN mkdir -p ${ITEMMAP_DATA_DIR}/uploads ${ITEMMAP_DATA_DIR}/outputs
VOLUME ["/app/data"]

EXPOSE ${PORT}

# Start via gunicorn
CMD ["/bin/sh", "-c", "gunicorn -w 2 -b 0.0.0.0:${PORT} wsgi:app"]
