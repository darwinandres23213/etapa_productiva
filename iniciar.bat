@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  SENA - Dashboard Etapa Productiva
echo  ---------------------------------
echo  Abriendo servidor local en http://127.0.0.1:8765
echo  No cierres esta ventana mientras uses el dashboard.
echo.
start "" "http://127.0.0.1:8765/index.html"
python -m http.server 8765
pause
