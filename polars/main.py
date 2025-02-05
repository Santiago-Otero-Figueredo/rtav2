from lectores.lector_oms import LectorOMS
from lectores.lector_mercadopago import LectorMercadoPago
from lectores.lector_erp import LectorERP
from lectores.modelos import ConfiguracionLector

from functools import wraps
from typing import Callable, Any
from datetime import datetime

import pandas as pd
import time
import psutil
import os

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


def exportar_resultados_excel(df: pd.DataFrame, nombre_base: str) -> None:
    """
    Exporta un DataFrame a Excel con timestamp en el nombre.

    Args:
        df (pd.DataFrame): DataFrame a exportar
        nombre_base (str): Nombre base del archivo
    """
    carpeta_resultados = "resultados"
    if not os.path.exists(carpeta_resultados):
        os.makedirs(carpeta_resultados)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"{carpeta_resultados}/{nombre_base}_{timestamp}.xlsx"

    df.to_excel(nombre_archivo, index=False)
    print(f"Archivo exportado: {nombre_archivo}")

def main():
    """
    Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
    """
    prueba_cruce_oms_mercado_pago()



@medir_rendimiento
def prueba_cruce_oms_mercado_pago():

    ruta_carpeta_oms = 'insumos/oms'
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()

    ruta_archivo_mercaopago = 'insumos/mercadopago/ORIGINAL MERCADOPAGO.xlsx'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_mercaopago)
    lector_mp = LectorMercadoPago(config)
    df_mp = lector_mp.dataframe()

    ruta_archivo_erp = 'insumos/erp/CARTERA ERP.xls'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    df_erp = lector.dataframe()

    exportar_resultados_excel(df_oms, "df_oms")


    raise Exception("Fin de la prueba")

    df_cruce = df_mp.copy()

    print(df_oms[['orden_externa_limpio', 'consecutivo']])

    print("\nResultados del cruce:", len(df_cruce))

    # Realizar el cruce de datos con pandas
    df_cruce = df_cruce.merge(
        df_oms[["orden_externa_limpio", "consecutivo"]],
        left_on="numero_identificacion_limpio",
        right_on="orden_externa_limpio",
        how="left"
    ).rename(columns={"consecutivo": "num_oms"})

    print("\nResultados del cruce:", len(df_cruce))

    # Realizar el segundo cruce
    df_cruce = df_cruce.merge(
        df_erp[["numero_oc_comercial", "numero_documento_cruce", "cedula_cliente", "total_cop", "auxiliar"]],
        left_on="num_oms",
        right_on="numero_oc_comercial",
        how="left"
    ).rename(columns={
        "numero_documento_cruce": "factura_erp",
        "total_cop": "valor_fv_erp",
        "cedula_cliente": "cc_erp",
        "auxiliar": "aux_erp"
    })

    print("\nResultados del cruce:", len(df_cruce))

    # Calcular diferencia de valores
    df_cruce['diferencia_valores'] = df_cruce['valor_fv_erp'] - df_cruce['monto_bruto_operacion']

    # Reordenar columnas
    columnas_ordenadas = [
        "numero_identificacion_limpio",
        "num_oms",
        "factura_erp",
        "valor_fv_erp",
        "cc_erp",
        "aux_erp",
        "diferencia_valores",
        "monto_bruto_operacion",
    ]

    # Reordenar columnas manteniendo el resto
    otras_columnas = [col for col in df_cruce.columns if col not in columnas_ordenadas]
    df_cruce = df_cruce[columnas_ordenadas + otras_columnas]

    exportar_resultados_excel(df_cruce, "cruce_oms_mercadopago_erp")

if __name__ == "__main__":
    main()
