@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Probando http://127.0.0.1:8000/health ...
powershell -NoProfile -Command "try { Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8000/health -TimeoutSec 5 | Select-Object StatusCode,Content } catch { Write-Host $_.Exception.Message; exit 1 }"
echo.
pause
