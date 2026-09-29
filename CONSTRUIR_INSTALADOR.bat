@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=py -3"
  goto :python_listo
)
where python >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=python"
  goto :python_listo
)
echo No se encontro Python 3.
pause
exit /b 1

:python_listo
echo Instalando dependencias de construccion...
%PYTHON_CMD% -m pip install --upgrade -r requirements-build.txt
if errorlevel 1 goto :error

echo Construyendo aplicacion autonoma...
%PYTHON_CMD% -m PyInstaller --noconfirm --clean catalogador_pozos.spec
if errorlevel 1 goto :error

set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ISCC%" goto :inno_listo
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ISCC%" goto :inno_listo
set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%ISCC%" goto :inno_listo

echo No se encontro Inno Setup 6.
echo Instalalo desde https://jrsoftware.org/isinfo.php o con:
echo winget install --id JRSoftware.InnoSetup --exact
pause
exit /b 1

:inno_listo
echo Compilando instalador...
"%ISCC%" packaging\catalogador_pozos.iss
if errorlevel 1 goto :error

echo.
echo Instalador creado en instalador_generado.
pause
exit /b 0

:error
echo.
echo La construccion fallo. Revisa los mensajes anteriores.
pause
exit /b 1

