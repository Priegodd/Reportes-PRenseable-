@echo off
setlocal

set "PYTHON_EXE=C:\Users\dsala\AppData\Local\Python\bin\python.exe"
set "APP_MODULE=web_app:app"

if not exist "%PYTHON_EXE%" (
  echo No se encontro Python en:
  echo %PYTHON_EXE%
  pause
  exit /b 1
)

echo Iniciando aplicacion web en http://127.0.0.1:8000
"%PYTHON_EXE%" -m uvicorn %APP_MODULE% --host 127.0.0.1 --port 8000

echo.
echo Proceso finalizado.
pause
