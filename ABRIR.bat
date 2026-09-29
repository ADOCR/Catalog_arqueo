@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
  py -3 catalogador_pozos.py
  if errorlevel 1 pause
  exit /b
)

where python >nul 2>nul
if not errorlevel 1 (
  python catalogador_pozos.py
  if errorlevel 1 pause
  exit /b
)

echo No se encontro Python 3.
echo Instala Python desde python.org y activa la opcion de agregarlo al PATH.
pause
exit /b 1

