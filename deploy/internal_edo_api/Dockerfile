FROM python:3.10-slim

WORKDIR /opt/app

ENV PYTHONDONTWRITECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    TZ="Asia/Atyrau" \
    POETRY_NO_INTERACTION=1 \
    POETRY_VIRTUALENVS_CREATE=false

COPY internal_edo_api/pyproject.toml pyproject.toml
COPY internal_edo_api/poetry.lock poetry.lock

RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir poetry==2.1.4 \
    && poetry install --only main --no-root \
    && rm -rf /root/.cache/pip

COPY deploy/internal_edo_api/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

COPY internal_edo_api/ .

ENTRYPOINT ["/entrypoint.sh"]