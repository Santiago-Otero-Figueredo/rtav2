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


def exportar_resultados_excel(df: pl.DataFrame, nombre_base: str) -> None:
    """
    Exporta un DataFrame a Excel con timestamp en el nombre.

    Args:
        df (pl.DataFrame): DataFrame a exportar
        nombre_base (str): Nombre base del archivo

    Returns:
        None: Guarda el archivo en disco
    """
    # Crear carpeta si no existe
    carpeta_resultados = "resultados"
    if not os.path.exists(carpeta_resultados):
        os.makedirs(carpeta_resultados)

    # Generar nombre con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"{carpeta_resultados}/{nombre_base}_{timestamp}.xlsx"

    # Exportar a Excel
    df.write_excel(nombre_archivo)
    print(f"Archivo exportado: {nombre_archivo}")

def main():
    """
    Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
    """
    prueba_cruce_oms_mercado_pago()
    #prueba_oms()
    #prueba_mercadopago()
    #prueba_erp()


def prueba_oms():
    ruta_carpeta_oms = 'insumos/oms'  # Ajusta esta ruta según tu estructura

    # Procesar archivos Excel
    inicio_tiempo = time.time()
    inicio_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    # Procesar archivos Excel
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
    lector = LectorOMS(config)
    resultado = lector.dataframe()

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
            'orden_externa_limpio',
            'orden_externa_duplicada_limpio',
            'costo excl imp',
            'precio vta excl imp',
            'imp vta',
            'imp costo',
            'sub total costo exl imp',
            'sub total vta exl imp',
            'forma_pago_1_valor',
            'forma_pago_2_valor',
            'forma_pago_3_valor',
            'cantidad',
            'monto'
        ]

        #resultado = resultado.filter(pl.col("orden externa").str.contains(r"^\d+\s\d+$"))


        print(resultado.select(columnas_mostrar).slice(0, 5).to_pandas().to_string())  # Muestra las primeras 5 filas
        return resultado
    else:
        print("No se encontraron archivos para procesar")
        return None


def prueba_mercadopago():
    ruta_archivo_mercaopago = 'insumos/mercadopago/ORIGINAL MERCADOPAGO.xlsx'  # Ajusta esta ruta según tu estructura

    # Procesar archivos Excel
    inicio_tiempo = time.time()
    inicio_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    # Procesar archivos Excel
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_mercaopago)
    lector = LectorMercadoPago(config)
    resultado = lector.dataframe()

    fin_tiempo = time.time()
    fin_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    tiempo_total = fin_tiempo - inicio_tiempo
    memoria_usada = fin_memoria - inicio_memoria

    print(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
    print(f"Memoria RAM usada: {memoria_usada:.2f} GB")

    # Unir todos los DataFrames
    if resultado is not None:
        columnas_mostrar = [
            'numero_identificacion_limpio',
            'numero_identificacion'
        ]

        print(resultado.select(columnas_mostrar).slice(0, 5).to_pandas().to_string())  # Muestra las primeras 5 filas

    else:
        print("No se encontraron archivos para procesar")
        return None



def prueba_erp():
    ruta_archivo_erp = 'insumos/erp/CARTERA ERP.xls'  # Ajusta esta ruta según tu estructura

    # Procesar archivos Excel
    inicio_tiempo = time.time()
    inicio_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    # Procesar archivos Excel
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    resultado = lector.dataframe

    fin_tiempo = time.time()
    fin_memoria = psutil.Process().memory_info().rss / (1024 ** 3)  # Convertir a GB

    tiempo_total = fin_tiempo - inicio_tiempo
    memoria_usada = fin_memoria - inicio_memoria

    print(f"Tiempo total de procesamiento: {tiempo_total:.2f} segundos")
    print(f"Memoria RAM usada: {memoria_usada:.2f} GB")

    # Unir todos los DataFrames
    if resultado.is_empty() is False:
        print(resultado.slice(0, 5))  # Muestra las primeras 5 filas
    else:
        print("No se encontraron archivos para procesar")
        return None


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



    df_cruce = df_mp.clone()

    print(df_oms.select(['orden_externa_limpio', 'consecutivo']))

    print("\nResultados del cruce:", df_cruce.height)

    # Realizar el cruce de datos
    df_cruce = df_cruce.join(
        df_oms.select(["orden_externa_limpio", "consecutivo"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_identificacion_limpio",  # Columna en df_cruce
        right_on="orden_externa_limpio",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    ).rename({"consecutivo": "num_oms"})


    # Mostrar resultados del cruce
    print("\nResultados del cruce:", df_cruce.height)
    # print(df_cruce)
    # print(f"Registros cruzados: {df_cruce.filter(pl.col('num_oms').is_not_null()).height}")
    # print(f"Registros sin cruzar: {df_cruce.filter(pl.col('num_oms').is_null())}")


     # Realizar el cruce de datos
    df_cruce = df_cruce.join(
        df_erp.select(["numero_oc_comercial", "numero_documento_cruce", "cedula_cliente", "total_cop", "auxiliar"]),  # Seleccionamos solo las columnas necesarias
        left_on="num_oms",  # Columna en df_cruce
        right_on="numero_oc_comercial",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    ).rename(
        {
            "numero_documento_cruce": "factura_erp",
            "total_cop": "valor_fv_erp",
            "cedula_cliente": "cc_erp",
            "auxiliar": "aux_erp"
        }
    )


    # Mostrar resultados del cruce
    print("\nResultados del cruce:", df_cruce.height)

    # print(df_cruce)
    # print(f"Registros cruzados: {df_cruce.filter(pl.col('factura_erp').is_not_null()).height}")
    # print(f"Registros sin cruzar: {df_cruce.filter(pl.col('factura_erp').is_null())}")


    # Realizar el cruce y crear columna de diferencia
    df_cruce = df_cruce.with_columns([
        (pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion"))
        .alias("diferencia_valores")
    ])

    # Mostrar resultados del cruce y la nueva columna

    # Definir orden específico de columnas
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
    df_cruce = df_cruce.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )


    exportar_resultados_excel(df_cruce, "cruce_oms_mercadopago_erp")


if __name__ == "__main__":
    main()
