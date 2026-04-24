# start-all.ps1 - OpenCode 全栈一键启动 (Windows 侧)
# 用法: powershell -ExecutionPolicy Bypass -File F:\scripts\start-all.ps1

$ErrorActionPreference = "Continue"
Write-Host "=== OpenCode 全栈启动 (Windows) ===" -ForegroundColor Cyan
Write-Host "时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# 1. 确保 WSL 已启动
Write-Host "`n[1/4] 启动 WSL..." -ForegroundColor Yellow
wsl -d Ubuntu -- echo "WSL ready" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] WSL 不可用" -ForegroundColor Red
    exit 1
}
Write-Host "  WSL 运行中" -ForegroundColor Green

# 2. 启动 OpenCode 双实例 (WSL 内)
Write-Host "`n[2/4] 启动 OpenCode 双实例..." -ForegroundColor Yellow
wsl -d Ubuntu -- bash -c "bash /mnt/f/scripts/start-all.sh"

# 3. 启动端口转发
Write-Host "`n[3/4] 检查端口转发..." -ForegroundColor Yellow
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

# 4. 验证手机端可达性
Write-Host "`n[4/4] 验证 Tailscale 端口..." -ForegroundColor Yellow
Start-Sleep -Seconds 2
foreach ($port in @(9080, 9081)) {
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.Connect("127.0.0.1", $port)
        $tcp.Close()
        Write-Host "  端口 $port 可达" -ForegroundColor Green
    } catch {
        Write-Host "  端口 $port 不可达" -ForegroundColor Red
    }
}

# 最终状态
Write-Host "`n=== 状态 ===" -ForegroundColor Cyan
Write-Host "Code Learning:  http://100.84.60.105:9080" -ForegroundColor White
Write-Host "KB Teacher:     http://100.84.60.105:9081" -ForegroundColor White
Write-Host "`n全部完成!" -ForegroundColor Green
