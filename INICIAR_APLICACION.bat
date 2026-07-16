@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

echo =============================================
echo   CONTROL DE VENDEDORES RIMEL
echo =============================================
echo.

if exist ".venv\Scripts\python.exe" goto INSTALL

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 goto NOPYTHON
  set "PY=python"
)

echo Creando entorno de la aplicacion...
%PY% -m venv .venv
if errorlevel 1 goto ERROR

:INSTALL
echo Instalando o verificando componentes locales...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto ERROR
".venv\Scripts\python.exe" -m pip install -r requirements-local.txt
if errorlevel 1 goto ERROR

echo Restableciendo usuarios iniciales...
".venv\Scripts\python.exe" reset_users.py
if errorlevel 1 goto ERROR

echo.
echo Aplicacion disponible en: http://127.0.0.1:8000/login
echo No cierre esta ventana mientras use la aplicacion.
echo.
start "" http://127.0.0.1:8000/login
".venv\Scripts\python.exe" app.py
goto END

:NOPYTHON
echo No se encontro Python 3 en este equipo.
echo Instale Python desde python.org y marque Add Python to PATH.
pause
goto END

:ERROR
echo.
echo Ocurrio un error. Revise el mensaje mostrado arriba.
echo Puede borrar la carpeta .venv y volver a ejecutar este archivo.
pause

:END
endlocal
