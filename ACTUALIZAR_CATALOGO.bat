@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   NUMKO - Actualizando catalogo desde Excel
echo ============================================
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo No encuentro Python. Instalalo gratis desde https://www.python.org/downloads/
  echo ^(marca la casilla "Add Python to PATH" al instalar^) y vuelve a hacer doble clic aqui.
  pause
  exit /b 1
)
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Instalando openpyxl ^(solo la primera vez^)...
  python -m pip install openpyxl
)
python herramientas\actualizar_catalogo.py
echo.
echo Si dice "Listo", abre index.html y recarga la pagina.
echo Revisa reporte_actualizacion.txt para ver perfumes sin foto y avisos.
pause
