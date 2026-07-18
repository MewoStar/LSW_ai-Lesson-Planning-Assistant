@echo off
title Stop AI BeiKe ZhuShou
powershell -Command "Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"
echo Service stopped.
timeout /t 2 >nul
