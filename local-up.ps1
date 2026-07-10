# 一键启动婚礼宾客统计系统
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$Port = if ($args[0]) { $args[0] } else { 8321 }

Write-Host "启动服务：http://127.0.0.1:${Port}/"
uv run uvicorn app:app --app-dir backend --host 127.0.0.1 --port $Port
