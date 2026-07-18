@echo off
chcp 65001 >nul
echo ============================================
echo    AI 备课助手 5.0 - 依赖重装脚本(便携版)
echo ============================================
echo.
echo 注意: 便携版已预装所有依赖,通常无需运行此脚本。
echo 仅在依赖损坏时使用。
echo.

cd /d "%~dp0"

if not exist ".venv\python.exe" (
    echo [错误] 未找到便携版 Python (.venv\python.exe)
    pause
    exit /b 1
)

echo [1/2] 正在升级 pip...
".venv\python.exe" -m pip install --upgrade pip
echo.

echo [2/2] 正在重新安装依赖包...
".venv\python.exe" -m pip install -r requirements.txt
echo.

echo ============================================
echo    依赖安装完成!
echo    双击 "启动备课助手.bat" 即可启动
echo ============================================
pause
