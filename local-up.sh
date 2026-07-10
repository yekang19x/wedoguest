#!/usr/bin/env bash
# 一键启动婚礼宾客统计系统
set -e
cd "$(dirname "$0")"

PORT="${1:-8321}"

echo "启动服务：http://127.0.0.1:${PORT}/"
exec uv run uvicorn app:app --app-dir backend --host 127.0.0.1 --port "${PORT}"
