@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m pip install --upgrade -r requirements.txt
  goto :resultado
)

where python >nul 2>nul
if not errorlevel 1 (
  python -m pip install --upgrade -r requirements.txt
  goto :resultado
)

echo No se encontro Python 3.
echo Instala Python desde python.org y activa la opcion de agregarlo al PATH.
pause
exit /b 1

:resultado
if errorlevel 1 (
  echo.
  echo No se pudo instalar. Comprueba la conexion a Internet y vuelve a intentarlo.
  pause
  exit /b 1
)
echo.
echo Instalacion completa. Ya puedes ejecutar ABRIR.bat.
pause
