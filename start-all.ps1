# start-all.ps1 - OpenCode 全栈一键启动 (Windows 侧)
# 用法: powershell -ExecutionPolicy Bypass -File F:\scripts\start-all.ps1

$ErrorActionPreference = "Continue"
Write-Host "=== OpenCode 全栈启动 (Windows) ===" -ForegroundColor Cyan
Write-Host "时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# 1. 检查 WSL 是否可用
Write-Host "`n[1/3] 检查 WSL..." -ForegroundColor Yellow
$wslCheck = wsl --list --running 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] WSL 不可用" -ForegroundColor Red
    exit 1
}
Write-Host "  WSL 运行中" -ForegroundColor Green

# 2. 启动 OpenCode 双实例 (WSL 内)
Write-Host "`n[2/3] 启动 OpenCode 双实例..." -ForegroundColor Yellow
wsl -d Ubuntu -- bash -c "bash /mnt/f/scripts/start-all.sh"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[WARN] OpenCode 启动可能有错误" -ForegroundColor Yellow
}

# 3. 启动端口转发
Write-Host "`n[3/3] 检查端口转发..." -ForegroundColor Yellow
$fwdRunning = Get-Process -Name python -ErrorAction SilentlyContinue | 
    Where-Object { $_.CommandLine -like '*port-fwd*' }

if ($fwdRunning) {
    Write-Host "  端口转发已在运行 (PID: $($fwdRunning.Id))" -ForegroundColor Green
} else {
    Write-Host "  启动端口转发..." -ForegroundColor Yellow
    Start-Process -FilePath "python" -ArgumentList "F:\scripts\port-fwd-win.py" `
        -WindowStyle Hidden -PassThru | ForEach-Object {
        Write-Host "  已启动 (PID: $($_.Id))" -ForegroundColor Green
    }
}

# 最终状态
Write-Host "`n=== 状态 ===" -ForegroundColor Cyan
Write-Host "Code Learning:  http://100.84.60.105:9080" -ForegroundColor White
Write-Host "KB Teacher:     http://100.84.60.105:9081" -ForegroundColor White
Write-Host "`n全部完成!" -ForegroundColor Green
