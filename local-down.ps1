# 停止婚礼宾客统计系统
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Port = if ($args[0]) { $args[0] } else { 8321 }

$Proc = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue |
    Where-Object State -eq Listen |
    Select-Object -ExpandProperty OwningProcess -First 1

if (-not $Proc) {
    Write-Host "端口 ${Port} 上没有运行中的服务"
    exit 0
}

Stop-Process -Id $Proc -Force
Write-Host "已停止端口 ${Port} 上的服务 (PID: ${Proc})"
