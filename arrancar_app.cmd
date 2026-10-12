@echo off
setlocal
title Mapa de potencial de farmacias 2027
cd /d "%~dp0"
if not exist "%~dp0pharmacy\.venv\Scripts\python.exe" (
  echo No se encuentra el entorno Python del proyecto.
  echo Consulta visualizacion\README.md para prepararlo.
  pause
  exit /b 1
)
"%~dp0pharmacy\.venv\Scripts\python.exe" "%~dp0visualizacion\server.py" --launch
if errorlevel 1 (
  echo No se ha podido abrir la aplicacion.
  pause
  exit /b 1
)
endlocal
