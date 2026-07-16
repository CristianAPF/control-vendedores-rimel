@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Se eliminara solamente el entorno tecnico .venv.
echo La base de datos rimel.db y sus registros NO se eliminaran.
if exist .venv rmdir /s /q .venv
call INICIAR_APLICACION.bat
