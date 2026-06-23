$ErrorActionPreference = "Continue"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "      AI BeiKe Assistant - 启动中..." -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "检查 Python 环境..." -ForegroundColor Gray
try {
    python --version | Out-Null
} catch {
    Write-Host "ERROR: 未找到 Python！请先安装 Python。" -ForegroundColor Red
    Write-Host "下载地址: https://www.python.org/downloads/" -ForegroundColor Red
    Read-Host -Prompt "`n按 Enter 键退出"
    exit 1
}

Write-Host "检查依赖库..." -ForegroundColor Gray
$requiredModules = @("flask", "openai", "pyyaml", "python-docx")
foreach ($module in $requiredModules) {
    try {
        python -c "import $module" | Out-Null
        Write-Host "  ✓ $module" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ $module 未安装" -ForegroundColor Yellow
        Write-Host "    正在安装 $module..." -ForegroundColor Gray
        pip install $module --quiet | Out-Null
        python -c "import $module" | Out-Null
        Write-Host "    ✓ 安装成功" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "启动 Flask 服务器..." -ForegroundColor Cyan

Start-Process -FilePath "python" -ArgumentList "web_app.py" -WorkingDirectory $scriptDir -WindowStyle Minimized

Write-Host "等待服务器启动..." -ForegroundColor Gray
Start-Sleep -Seconds 3

Write-Host "打开浏览器..." -ForegroundColor Cyan
try {
    Start-Process "http://127.0.0.1:5000"
} catch {
    Write-Host "提示: 请手动打开浏览器访问 http://localhost:5000" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "         启动成功！" -ForegroundColor Green
Write-Host "         访问地址: http://localhost:5000" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "提示: 关闭此窗口将同时停止服务" -ForegroundColor Gray
Write-Host ""
Read-Host -Prompt "按 Enter 键最小化此窗口"