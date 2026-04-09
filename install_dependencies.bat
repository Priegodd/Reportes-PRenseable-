@echo off
setlocal

set "PYTHON_EXE=C:\Users\dsala\AppData\Local\Python\bin\python.exe"

if not exist "%PYTHON_EXE%" (
  echo No se encontro Python en:
  echo %PYTHON_EXE%
  echo.
  echo Edita este archivo y actualiza la variable PYTHON_EXE con tu ruta real.
  pause
  exit /b 1
)

echo Instalando dependencias...
"%PYTHON_EXE%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo.
  echo Ocurrio un error al instalar dependencias.
  pause
  exit /b 1
)

echo.
echo Dependencias instaladas correctamente.
pause
