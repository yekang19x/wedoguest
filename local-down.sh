#!/usr/bin/env bash
# 停止婚礼宾客统计系统
set -e

PORT="${1:-8321}"

PID=$(lsof -ti :"${PORT}" 2>/dev/null || true)

if [ -z "$PID" ]; then
  echo "端口 ${PORT} 上没有运行中的服务"
  exit 0
fi

kill "$PID"
echo "已停止端口 ${PORT} 上的服务 (PID: ${PID})"
