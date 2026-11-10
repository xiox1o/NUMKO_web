#!/bin/bash
cd "$(dirname "$0")"
echo "NUMKO - Actualizando catalogo desde Excel"
python3 -c "import openpyxl" 2>/dev/null || python3 -m pip install openpyxl
python3 herramientas/actualizar_catalogo.py
echo
echo "Si dice 'Listo', recarga la pagina. Detalle en reporte_actualizacion.txt"
read -n 1 -s -r -p "Presiona una tecla para cerrar..."
