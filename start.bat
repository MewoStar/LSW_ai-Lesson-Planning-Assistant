@echo off
cd /d "%~dp0"

set PYTHON=.venv\Scripts\python.exe
set APP=web_app.py

if not exist "%PYTHON%" (
    echo ERROR: Python environment not found!
    echo Please check if .venv folder exists.
    pause
    exit /b 1
)

echo Starting AI BeiKe Assistant...
echo.

start "BeiKe Assistant" /min "%PYTHON%" "%APP%"

echo Waiting for server...
ping 127.0.0.1 -n 3 >nul

echo Opening browser...
start http://127.0.0.1:5000

echo.
echo =============================================
echo         STARTED SUCCESSFULLY!
echo         URL: http://localhost:5000
echo         Press any key to close...
echo =============================================
pause >nul