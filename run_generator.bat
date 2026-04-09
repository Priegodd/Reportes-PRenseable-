@echo off
setlocal

set "PYTHON_EXE=C:\Users\dsala\AppData\Local\Python\bin\python.exe"
set "SCRIPT=%~dp0generate_plan.py"

if not exist "%PYTHON_EXE%" (
  echo No se encontro Python en:
  echo %PYTHON_EXE%
  echo.
  echo Edita run_generator.bat y actualiza la variable PYTHON_EXE.
  pause
  exit /b 1
)

if not exist "%SCRIPT%" (
  echo No se encontro el script principal:
  echo %SCRIPT%
  pause
  exit /b 1
)

if not "%~1"=="" (
  "%PYTHON_EXE%" "%SCRIPT%" "%~1"
  goto :end
)

echo Buscando archivos dentro de la carpeta input...
"%PYTHON_EXE%" "%SCRIPT%" --auto

:end
echo.
echo Proceso finalizado.
pause
