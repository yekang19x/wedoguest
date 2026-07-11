FROM python:3.12-slim AS builder

# 不从 ghcr.io 拉 uv 镜像，国内拉取会卡到几 KB/s
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple uv==0.11.28

WORKDIR /app
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev --no-install-project

COPY backend/ backend/
COPY frontend/ frontend/

FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1

WORKDIR /app
COPY --from=builder /app /app

EXPOSE 8321

CMD ["/app/.venv/bin/uvicorn", "app:app", "--app-dir", "backend", "--host", "0.0.0.0", "--port", "8321"]
