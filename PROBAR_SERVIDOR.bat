@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal EnableExtensions
set "LOG=diagnostico_inicio.txt"
echo ===== INICIO %date% %time% ===== > "%LOG%"
echo Carpeta: %cd% >> "%LOG%"

echo =============================================
echo   CONTROL DE VENDEDORES RIMEL
echo =============================================
echo.

if exist ".venv\Scripts\python.exe" goto INSTALL

where py >nul 2>nul
if not errorlevel 1 (
  set "PY=py -3"
  goto CREATE
)
where python >nul 2>nul
if not errorlevel 1 (
  set "PY=python"
  goto CREATE
)
goto NOPYTHON

:CREATE
echo Creando entorno...
%PY% --version >> "%LOG%" 2>&1
%PY% -m venv .venv >> "%LOG%" 2>&1
if errorlevel 1 goto ERROR

:INSTALL
echo Verificando componentes...
".venv\Scripts\python.exe" --version >> "%LOG%" 2>&1
".venv\Scripts\python.exe" -m pip install -r requirements-local.txt >> "%LOG%" 2>&1
if errorlevel 1 goto ERROR

echo Preparando usuarios...
".venv\Scripts\python.exe" reset_users.py >> "%LOG%" 2>&1
if errorlevel 1 goto ERROR

echo Iniciando servidor...
start "RIMEL SERVIDOR" /min cmd /c ""%cd%\.venv\Scripts\python.exe" "%cd%\app.py" >> "%cd%\servidor.log" 2>&1"

echo Esperando que la aplicacion quede disponible...
for /L %%i in (1,1,30) do (
  powershell -NoProfile -Command "try { $r=Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8000/health -TimeoutSec 2; if($r.StatusCode -eq 200){exit 0}else{exit 1} } catch { exit 1 }" >nul 2>nul
  if not errorlevel 1 goto OPEN
  timeout /t 1 /nobreak >nul
)

echo.
echo La aplicacion no pudo iniciar.
echo Abra servidor.log y diagnostico_inicio.txt para ver el error.
echo.
type servidor.log
pause
goto END

:OPEN
echo.
echo Aplicacion iniciada correctamente.
echo Direccion: http://127.0.0.1:8000/login
echo.
start "" http://127.0.0.1:8000/login
echo Puede cerrar esta ventana. Para detener el servidor, cierre la ventana llamada RIMEL SERVIDOR.
pause
goto END

:NOPYTHON
echo No se encontro Python 3. >> "%LOG%"
echo No se encontro Python 3 en este equipo.
echo Instale Python desde python.org y marque Add Python to PATH.
pause
goto END

:ERROR
echo ERROR %errorlevel% >> "%LOG%"
echo.
echo Ocurrio un error durante la preparacion.
echo Abra diagnostico_inicio.txt para ver el detalle.
echo.
type "%LOG%"
pause

:END
endlocal
