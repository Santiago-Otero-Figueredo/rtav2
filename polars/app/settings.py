import json
import os

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent  # Carpeta base de la app
SECRETS = {
    "ruta_datos": str(BASE_DIR / "datos"),  # Carpeta "datos" dentro de la app
    "ruta_logs": str(BASE_DIR / "logs"),  # Carpeta "logs" dentro de la app
}

RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
RUTA_INSUMOS = os.path.join(RUTA_RAIZ, 'insumos')

with open(os.path.join(RUTA_RAIZ, 'secrets.json'), 'r') as f:
    variables_ambiente = json.load(f)

if variables_ambiente['RUTA_RESULTADOS'] == "":
    RUTA_RESULTADOS = os.path.join(RUTA_RAIZ, 'resultados')
else:
    RUTA_RESULTADOS = variables_ambiente['RUTA_RESULTADOS']


RUTA_INSUMOS_ADDI =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_ADDI']
RUTA_INSUMOS_ERP =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_ERP']
RUTA_INSUMOS_MP =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_MP']
RUTA_INSUMOS_ML =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_ML']
RUTA_INSUMOS_OMS =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_OMS']
RUTA_INSUMOS_SISTECREDITO =  BASE_DIR / variables_ambiente['RUTA_INSUMOS_SISTECREDITO']
RUTA_INSUMOS_FACTURA_SISTECREDITO = RUTA_INSUMOS_SISTECREDITO / variables_ambiente['RUTA_INSUMOS_FACTURA_SISTECREDITO']
RUTA_INSUMOS_PAGARE_SISTECREDITO = RUTA_INSUMOS_SISTECREDITO / variables_ambiente['RUTA_INSUMOS_PAGARE_SISTECREDITO']
RUTA_RESULTADOS =  BASE_DIR / variables_ambiente['RUTA_RESULTADOS']

print('BASE_DIR: ', RUTA_INSUMOS_FACTURA_SISTECREDITO)