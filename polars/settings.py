import json
import os

RUTA_RAIZ = os.path.dirname(os.path.abspath(__file__))
RUTA_INSUMOS = os.path.join(RUTA_RAIZ, 'insumos')

with open(os.path.join(RUTA_RAIZ, 'secrets.json'), 'r') as f:
    variables_ambiente = json.load(f)

if variables_ambiente['RUTA_RESULTADOS'] == "":
    RUTA_RESULTADOS = os.path.join(RUTA_RAIZ, 'resultados')
else:
    RUTA_RESULTADOS = variables_ambiente['RUTA_RESULTADOS']
