# 一键启动婚礼宾客统计系统
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$Port = if ($args[0]) { $args[0] } else { 8321 }

try {
    python -c "import fastapi, uvicorn, openpyxl" 2>$null
} catch {
    Write-Host "缺少依赖，正在安装 ..."
    python -m pip install -r requirements.txt
}

Write-Host "启动服务：http://127.0.0.1:${Port}/"
python -m uvicorn app:app --app-dir backend --host 127.0.0.1 --port $Port
