FROM python:3.12-slim AS builder

COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

WORKDIR /app
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

COPY backend/ backend/
COPY frontend/ frontend/

FROM python:3.12-slim

WORKDIR /app
COPY --from=builder /app /app

EXPOSE 8321

CMD ["/app/.venv/bin/uvicorn", "app:app", "--app-dir", "backend", "--host", "0.0.0.0", "--port", "8321"]
