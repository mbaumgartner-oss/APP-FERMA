@echo off
rem Instala las macros de exportar-fase en todas las versiones de Tekla de esta PC. Doble clic y listo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_macros.ps1"
echo.
pause
