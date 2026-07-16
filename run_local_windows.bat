@echo off
cd /d %~dp0
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
%PY% -m venv .venv
call .venv\Scripts\activate
python -m pip install -r requirements.txt
set SECRET_KEY=CAMBIAR_ESTA_CLAVE_ANTES_DE_USAR_EN_PRODUCCION
python app.py
pause
