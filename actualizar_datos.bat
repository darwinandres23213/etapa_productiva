@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Actualizando data.json desde el Excel...
python exportar_datos.py
echo.
pause
