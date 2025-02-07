from lectores.lector_oms import LectorOMS
from lectores.lector_mercadopago import LectorMercadoPago
from lectores.lector_erp import LectorERP

from lectores.modelos import ConfiguracionLector

from functools import wraps
from typing import Callable, Any
from datetime import datetime

import polars as pl
import time
import psutil
import os
import glob
import pandas as pd



def medir_rendimiento(funcion: Callable) -> Callable:
    """
    Decorador que mide el tiempo de ejecución y uso de memoria de una función.

    Args:
        funcion (Callable): Función a decorar

    Returns:
        Callable: Función decorada con medición de rendimiento
    """
    @wraps(funcion)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        inicio_tiempo = time.time()
        inicio_memoria = psutil.Process().memory_info().rss / (1024 ** 3)

        resultado = funcion(*args, **kwargs)

        fin_tiempo = time.time()
        fin_memoria = psutil.Process().memory_info().rss / (1024 ** 3)

        tiempo_total = fin_tiempo - inicio_tiempo
        memoria_usada = fin_memoria - inicio_memoria

        print(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
        print(f"Memoria RAM usada: {memoria_usada:.2f} GB")

        return resultado

    return wrapper


def exportar_resultados_excel(df: pl.DataFrame, nombre_base: str, ruta_carpeta:str='resultados') -> None:
    """
    Exporta un DataFrame de pandas a Excel con timestamp en el nombre.

    Args:
        df (pd.DataFrame): DataFrame de pandas a exportar
        nombre_base (str): Nombre base del archivo

    Returns:
        None: Guarda el archivo en disco
    """
    try:
        # Crear carpeta si no existe
        carpeta = ruta_carpeta
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)

        # Generar nombre con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"{carpeta}/{nombre_base}_{timestamp}.xlsx"

        # Exportar a Excel usando pandas
        df.to_excel(
            nombre_archivo,
            index=False,  # No incluir índice
            engine='openpyxl'  # Motor para archivos xlsx
        )
        print(f"Archivo exportado: {nombre_archivo}")

    except Exception as e:
        print(f"Error al exportar el archivo: {str(e)}")



def buscar_archivo_procesado(ruta: str, patron: str) -> str:
    """
    Busca y retorna el nombre completo del primer archivo que coincida con el patrón.

    Args:
        patron (str): Patrón inicial del nombre del archivo
        ruta (str): Ruta donde buscar el archivo

    Returns:
        str: Ruta completa del archivo encontrado o None si no existe

    Raises:
        FileNotFoundError: Si no se encuentra ningún archivo que coincida con el patrón
    """
    ruta_busqueda = os.path.join(ruta, f"{patron}*.xlsx")
    archivos_encontrados = glob.glob(ruta_busqueda)

    if not archivos_encontrados:
       return None

    # Retorna el archivo más reciente si hay varios
    return max(archivos_encontrados, key=os.path.getctime)


def main():
    """
    Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
    """
    #prueba_lectura_directa()
    prueba_cruce_oms_mercado_pago()
    #prueba_oms()
    #prueba_mercadopago()
    #prueba_erp()




@medir_rendimiento
def prueba_cruce_oms_mercado_pago():

    archivo_oms = buscar_archivo_procesado('insumos/oms/', 'df_oms_procesado')
    config_oms = ConfiguracionLector(ruta_carpeta='insumos/oms')
    if archivo_oms:
        config_oms = ConfiguracionLector(ruta_archivo=archivo_oms, cargue_inicial=False)
    lector_oms = LectorOMS(config_oms)
    df_oms = lector_oms.dataframe()

    if archivo_oms is None:
        exportar_resultados_excel(df_oms, "df_oms_procesado", ruta_carpeta='insumos/oms/')

    archivo_mp = buscar_archivo_procesado('insumos/mercadopago/', 'df_mp_procesado')
    config_mp = ConfiguracionLector(ruta_archivo='insumos/mercadopago/reserve-release-686352448-2025-02-06-072601.xlsx')
    if archivo_mp:
        config_mp = ConfiguracionLector(ruta_archivo=archivo_mp, cargue_inicial=False)
    lector_mp = LectorMercadoPago(config_mp)
    df_mp = lector_mp.dataframe()

    if archivo_mp is None:
        exportar_resultados_excel(df_mp, "df_mp_procesado", ruta_carpeta='insumos/mercadopago/')


    archivo_erp = buscar_archivo_procesado('insumos/erp/', 'df_erp_procesado')
    config_erp = ConfiguracionLector(ruta_archivo='insumos/erp/ZOMAC - 2025-02-06T142211.293.xls')
    if archivo_erp:
        config_erp = ConfiguracionLector(ruta_archivo=archivo_erp, cargue_inicial=False)

    lector = LectorERP(config_erp)
    df_erp = lector.dataframe()

    if archivo_erp is None:
        exportar_resultados_excel(df_erp, "df_erp_procesado", ruta_carpeta='insumos/erp/')


    df_cruce = df_mp.copy()

    print("\nResultados ANTES del cruce:", len(df_cruce))

    # Primero, limpiar los DataFrames antes del merge
    df_oms_sin_ordenes_na = df_oms[
        # Eliminar NaN
        df_oms["orden_externa_str_limpio"].notna() &
        # Eliminar strings vacíos o solo espacios
        (df_oms["orden_externa_str_limpio"].str.strip() != "")
    ].copy()

    # Realizar el cruce de datos
    df_cruce = df_cruce.merge(
        df_oms_sin_ordenes_na[["orden_externa_str_limpio", "consecutivo"]],  # Seleccionamos solo las columnas necesarias
        left_on="numero_identificacion_str_limpio",  # Columna en df_cruce
        right_on="orden_externa_str_limpio",  # Columna en df_oms_sin_ordenes_na
        how="left"  # Mantener todos los registros de df_cruce
    ).rename(columns={"consecutivo": "num_oms"})

    # Mostrar resultados del cruce
    print("\nResultados del cruce:", len(df_cruce))


  # Primero, limpiar los DataFrames antes del merge
    df_erp_sin_n_comercial_na = df_erp[
        # Eliminar NaN
        df_erp["numero_oc_comercial"].notna() &
        # Eliminar strings vacíos o solo espacios
        (df_erp["numero_oc_comercial"].str.strip() != "")
    ].copy()

    df_cruce["num_oms"] = pd.to_numeric(df_cruce["num_oms"], errors="coerce").fillna(0).astype(int).astype(str)
    df_erp_sin_n_comercial_na["numero_oc_comercial"] = df_erp_sin_n_comercial_na["numero_oc_comercial"].astype(str)
    # Realizar el cruce de datos
    df_cruce = df_cruce.merge(
        # Seleccionar columnas necesarias del DataFrame ERP
        df_erp_sin_n_comercial_na[["numero_oc_comercial", "numero_documento_cruce", "cedula_cliente", "total_cop", "auxiliar"]],
        left_on="num_oms",  # Columna en df_cruce
        right_on="numero_oc_comercial",  # Columna en df_erp_sin_n_comercial_na
        how="left"  # Mantener todos los registros de df_cruce
    ).rename(columns={
        "numero_documento_cruce": "factura_erp",
        "total_cop": "valor_fv_erp",
        "cedula_cliente": "cc_erp",
        "auxiliar": "aux_erp"
    })



    # Mostrar resultados del cruce
    print("\nResultados del cruce:", len(df_cruce))


    # print(df_cruce)
    # print(f"Registros cruzados: {df_cruce.filter(pl.col('factura_erp').is_not_null()).height}")
    # print(f"Registros sin cruzar: {df_cruce.filter(pl.col('factura_erp').is_null())}")


    # Calcular la diferencia y crear nueva columna
    df_cruce['diferencia_valores'] = df_cruce['valor_fv_erp'] - df_cruce['monto_bruto_operacion']

    # Mostrar resultados del cruce y la nueva columna

    # Definir orden específico de columnas
    columnas_ordenadas = [
        "numero_identificacion_str_limpio",
        "num_oms",
        "factura_erp",
        "valor_fv_erp",
        "cc_erp",
        "aux_erp",
        "diferencia_valores",
        "monto_bruto_operacion"
    ]

    # Reordenar columnas manteniendo el resto
    columnas_restantes = [col for col in df_cruce.columns if col not in columnas_ordenadas]
    df_cruce = df_cruce[columnas_ordenadas + columnas_restantes]


    exportar_resultados_excel(df_cruce, "cruce_oms_mercadopago_erp")


if __name__ == "__main__":
    main()
