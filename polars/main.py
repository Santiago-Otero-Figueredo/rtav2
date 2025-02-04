import os
from pathlib import Path
from typing import List, Dict
import polars as pl
from lectores.lector_excel import LectorExcel
from lectores.modelos import ConfiguracionLector
import time
import psutil


def obtener_archivos_oms(ruta_carpeta: str) -> Dict[str, List[str]]:
    """
    Obtiene todos los archivos Excel y CSV de la carpeta especificada.

    Args:
        ruta_carpeta (str): Ruta de la carpeta a escanear

    Returns:
        Dict[str, List[str]]: Diccionario con listas de rutas de archivos por tipo
    """
    archivos = {
        'excel': [],
        'csv': []
    }

    for archivo in Path(ruta_carpeta).glob('*'):
        if archivo.suffix.lower() in ['.xlsx', '.xls']:
            archivos['excel'].append(str(archivo))
        elif archivo.suffix.lower() == '.csv':
            archivos['csv'].append(str(archivo))

    return archivos

def main():
    """
    Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
    """
    ruta_carpeta_oms = 'insumos/oms'  # Ajusta esta ruta según tu estructura
    archivos = obtener_archivos_oms(ruta_carpeta_oms)
    dataframes = []

    resultado = pl.DataFrame()  # Iniciar un DataFrame vacío

    # Procesar archivos Excel
    inicio_tiempo = time.time()
    inicio_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    # Procesar archivos Excel
    for ruta_excel in archivos['excel']:
        config = ConfiguracionLector(ruta_archivo=ruta_excel)
        lector = LectorExcel(config)
        df_temp = lector.dataframe
        resultado = resultado.vstack(df_temp)

    fin_tiempo = time.time()
    fin_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    tiempo_total = fin_tiempo - inicio_tiempo
    memoria_usada = fin_memoria - inicio_memoria

    print(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
    print(f"Memoria RAM usada: {memoria_usada:.2f} GB")

    # Unir todos los DataFrames
    if resultado.is_empty() is False:
        print("DataFrame final unificado:")
        print(f"Total de registros: {resultado.height}")

        columnas_mostrar = [
            'consecutivo',
            'orden externa',
            'orden externa dupliacada',
            'costo excl imp',
            'precio vta excl imp',
            'imp vta',
            'imp costo',
            'sub total costo exl imp',
            'sub total vta exl imp',
            'forma pago 1 valor',
            'forma pago 2 valor',
            'forma pago 3 valor',
            'cantidad'
        ]

        #resultado = resultado.filter(pl.col("orden externa").str.contains(r"^\d+\s\d+$"))


        print(resultado.select(columnas_mostrar).slice(0, 5).to_pandas().to_string())  # Muestra las primeras 5 filas
        return resultado
    else:
        print("No se encontraron archivos para procesar")
        return None

if __name__ == "__main__":
    main()
