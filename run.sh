#!/usr/bin/env bash
# 一键启动婚礼宾客统计系统
set -e
cd "$(dirname "$0")"

PORT="${1:-8321}"

python3 -c "import fastapi, uvicorn, openpyxl" 2>/dev/null || {
  echo "缺少依赖，正在安装 ..."
  python3 -m pip install -r requirements.txt
}

echo "启动服务：http://127.0.0.1:${PORT}/"
exec python3 -m uvicorn app:app --app-dir backend --host 127.0.0.1 --port "${PORT}"
