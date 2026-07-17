@echo off
chcp 65001 >nul
echo ============================================
echo    AI 备课助手 5.0 - 依赖安装脚本
echo ============================================
echo.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境，请先创建虚拟环境
    echo 可以运行: python -m venv .venv
    pause
    exit /b 1
)

echo [1/2] 正在升级 pip...
.\.venv\Scripts\python.exe -m pip install --upgrade pip
echo.

echo [2/2] 正在安装依赖包...
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
echo.

echo ============================================
echo    依赖安装完成！
echo    双击 start.bat 即可启动程序
echo ============================================
pause
