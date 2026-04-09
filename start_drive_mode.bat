@echo off
setlocal

set "PYTHON_EXE=C:\Users\dsala\AppData\Local\Python\bin\python.exe"
set "SCRIPT=%~dp0process_drive_folder.py"

if not exist "%PYTHON_EXE%" (
  echo No se encontro Python en:
  echo %PYTHON_EXE%
  pause
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo No se encontro el script:
  echo %SCRIPT%
  pause
  exit /b 1
)

echo Iniciando modo Drive...
"%PYTHON_EXE%" "%SCRIPT%"

echo.
echo Proceso finalizado.
pause
