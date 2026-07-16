@echo off
cd /d %~dp0
echo Restableciendo usuarios de Distribuidora RIMEL...
if exist .venv\Scripts\python.exe (
  .venv\Scripts\python.exe reset_users.py
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 reset_users.py
  ) else (
    python reset_users.py
  )
)
echo.
echo Credenciales:
echo gerencia / Rimel2026!
echo gerson / Gerson2026!
echo eduardo / Eduardo2026!
echo victoria / Victoria2026!
echo.
pause
