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
import pandas as pd


import json

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

def unir_dataframes_cruce(df_principal: pl.DataFrame, df_adicional: pl.DataFrame) -> pl.DataFrame:
    """
    Une dos DataFrames verticalmente asegurando compatibilidad de columnas.

    Args:
        df_principal (pl.DataFrame): DataFrame principal con los cruces originales
        df_adicional (pl.DataFrame): DataFrame con cruces adicionales por cédula

    Returns:
        pl.DataFrame: DataFrame unificado con todos los cruces
    """
    try:
        # Verificar que ambos DataFrames no estén vacíos
        if df_principal.height == 0 or df_adicional.height == 0:
            print("Advertencia: Uno de los DataFrames está vacío")
            return df_principal

        # Imprimir información antes de la unión
        print("\nEstadísticas antes de la unión:")
        print(f"Registros en df_principal: {df_principal.height}")
        print(f"Registros en df_adicional: {df_adicional.height}")

        # Realizar la unión vertical
        df_unificado = pl.concat([df_principal, df_adicional])

        # Imprimir información después de la unión
        print("\nEstadísticas después de la unión:")
        print(f"Total de registros unidos: {df_unificado.height}")

        return df_unificado

    except Exception as e:
        print(f"Error al unir DataFrames: {str(e)}")
        raise


def exportar_multiples_dataframes_excel(
    dataframes: dict[str, pl.DataFrame],
    nombre_base: str
) -> None:
    """
    Exporta múltiples DataFrames a diferentes hojas de un mismo archivo Excel.

    Args:
        dataframes (dict[str, pl.DataFrame]): Diccionario con nombre_hoja:DataFrame
        nombre_base (str): Nombre base para el archivo Excel

    Returns:
        None: Guarda el archivo en disco
    """
    try:
        # Crear carpeta si no existe
        carpeta_resultados = "resultados"
        if not os.path.exists(carpeta_resultados):
            os.makedirs(carpeta_resultados)

        # Generar nombre con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"{carpeta_resultados}/{nombre_base}_{timestamp}.xlsx"

        # Crear el escritor de Excel
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            # Iterar sobre cada DataFrame y escribirlo en su hoja
            for nombre_hoja, df in dataframes.items():
                # Convertir DataFrame de polars a pandas
                df_pandas = df.to_pandas()
                # Escribir en la hoja específica
                df_pandas.to_excel(writer, sheet_name=nombre_hoja, index=False)

        print(f"Archivo exportado: {nombre_archivo}")
        print(f"Hojas creadas: {list(dataframes.keys())}")

    except Exception as e:
        print(f"Error al exportar DataFrames: {str(e)}")
        raise

def obtener_facturas_canceladas(df: pl.DataFrame) -> pl.DataFrame:
    """
    Identifica las facturas que han sido canceladas basándose en la suma de montos
    y la presencia de registros tipo 'refund'.

    Args:
        df (pl.DataFrame): DataFrame con las columnas factura_erp, monto_bruto_operacion y descripcion

    Returns:
        pl.DataFrame: DataFrame con la marca de cancelación por factura
    """
    # Realizar el análisis por factura
    df_cancelaciones = df.group_by("factura_erp").agg([
        # Suma total de montos
        pl.col("monto_bruto_operacion").sum().alias("suma_montos"),
        # Verificar si existe algún refund
        pl.col("descripcion").str.to_lowercase().eq("refund").any().alias("tiene_refund")
    ]).with_columns([
        # Marcar como cancelado si suma es 0 y tiene refund
        pl.when(
            (pl.col("suma_montos").abs() < 1) & pl.col("tiene_refund")
        ).then(
            pl.lit(1)
        ).otherwise(
            pl.lit(0)
        ).alias("cancelado")
    ])

    # Unir el resultado con el DataFrame original
    df_resultado = df.join(
        df_cancelaciones.select(["factura_erp", "cancelado"]),
        on="factura_erp",
        how="left"
    ).filter(
        pl.col("cancelado") == 1
    ).drop(
        'cancelado'
    )

    return df_resultado

def obtener_facturas_devueltas(df: pl.DataFrame) -> pl.DataFrame:
    """
    Identifica las facturas que tienen devoluciones basándose en la suma negativa
    de montos y la presencia de registros tipo 'refund'.

    Args:
        df (pl.DataFrame): DataFrame con las columnas factura_erp, monto_bruto_operacion y descripcion

    Returns:
        pl.DataFrame: DataFrame con las facturas que tienen devoluciones
    """
    # Realizar el análisis por factura
    df_devoluciones = df.group_by("factura_erp").agg([
        # Suma total de montos
        pl.col("monto_bruto_operacion").sum().alias("suma_montos"),
        # Verificar si existe algún refund
        pl.col("descripcion").str.to_lowercase().eq("refund").any().alias("tiene_refund")
    ]).with_columns([
        # Marcar como devuelto si suma es negativa y tiene refund
        pl.when(
            (pl.col("suma_montos") < 0) & pl.col("tiene_refund")
        ).then(
            pl.lit(1)
        ).otherwise(
            pl.lit(0)
        ).alias("devuelto")
    ])

    # Unir el resultado con el DataFrame original y filtrar devoluciones
    df_resultado = df.join(
        df_devoluciones.select(["factura_erp", "devuelto"]),
        on="factura_erp",
        how="left"
    ).filter(
        pl.col("devuelto") == 1
    ).drop(
        'devuelto'
    )

    return df_resultado

# Función auxiliar para comparar valores con margen de error
def son_valores_similares(valor1: float, valores: list[float], margen: float = 0.01) -> bool:
    """
    Compara un valor con una lista de valores permitiendo un margen de error.

    Args:
        valor1 (float): Valor a comparar
        valores (list[float]): Lista de valores para comparar
        margen (float): Margen de error permitido (por defecto 0.01)

    Returns:
        bool: True si encuentra algún valor dentro del margen permitido
    """
    return any(abs(valor1 - valor2) <= margen for valor2 in valores)

def procesar_facturas_agrupadas(df: pl.DataFrame) -> pl.DataFrame:
    """
    Procesa las facturas agrupadas según los criterios especificados.

    Args:
        df (pl.DataFrame): DataFrame con las columnas factura_erp, diferencia_valores,
                          valor_fv_erp y monto_bruto_operacion

    Returns:
        pl.DataFrame: DataFrame procesado con las facturas agrupadas
    """
    # Paso 1: Análisis inicial por factura
    analisis_facturas = df.group_by("factura_erp").agg([
        pl.col("diferencia_valores").cast(pl.Decimal(20,2)).alias("valores_diferencia"),
        pl.count().alias("cantidad_registros"),
        pl.col("diferencia_valores").eq(0).any().alias("tiene_cero"),
        pl.col("valor_fv_erp").cast(pl.Decimal(20,2)).unique().alias("valores_fv_erp")
    ])

    # Paso 2: Identificar facturas que cumplen los criterios
    facturas_a_procesar = analisis_facturas.filter(
        (pl.col("cantidad_registros") > 1) &  # Más de un registro
        ~pl.col("tiene_cero")  # No tiene ceros
    )

    # Paso 3: Para cada factura que cumple los criterios
    facturas_procesadas = []
    facturas_a_mantener = []


    for factura in facturas_a_procesar["factura_erp"]:
        registros_factura = df.filter(pl.col("factura_erp") == factura)
        suma_diferencias = registros_factura["diferencia_valores"].cast(pl.Decimal(20,2)).sum()
        valores_fv_erp = registros_factura["valor_fv_erp"].cast(pl.Decimal(20,2)).to_list()

        # Usar la nueva función de comparación con margen
        if son_valores_similares(float(suma_diferencias), [float(x) for x in valores_fv_erp]):
            nuevo_registro = registros_factura.select([
                pl.all().first()
            ]).with_columns([
                pl.lit(suma_diferencias).cast(pl.Decimal(20,2)).alias("monto_bruto_operacion"),
                pl.lit(0).cast(pl.Decimal(20,2)).alias("diferencia_valores")
            ])
            facturas_procesadas.append(nuevo_registro)
        else:
            facturas_a_mantener.append(factura)

    # Paso 4: Construir DataFrame resultante
    df_procesadas = pl.concat(facturas_procesadas) if facturas_procesadas else pl.DataFrame()

    # Añadir columna es_multiple con valor por defecto 0
    df_procesadas = df_procesadas.with_columns([
        pl.lit(1).cast(pl.Int64).alias("es_multiple")
    ])



    return df_procesadas

def extraer_valores_impuestos(impuestos: str, tipo_impuesto: str) -> float:
    """
    Extrae el valor del impuesto especificado de una cadena JSON.

    Args:
        impuestos (str): Cadena JSON con los detalles de los impuestos.
        tipo_impuesto (str): Tipo de impuesto a extraer (e.g., "iva", "fuente", "ica_bogota").

    Returns:
        float: Valor del impuesto extraído o 0 si no se encuentra.
    """
    try:
        lista_impuestos = json.loads(impuestos)
        for impuesto in lista_impuestos:
            if impuesto["financial_entity"].startswith(tipo_impuesto):
                return impuesto["amount"]
    except (json.JSONDecodeError, KeyError, TypeError):
        return 0.0
    return 0.0


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

    ruta_carpeta_oms = 'insumos/oms/'
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()

    ruta_archivo_mercaopago = 'insumos/mercadopago/reserve-release-686352448-2025-02-06-072601.xlsx'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_mercaopago)
    lector_mp = LectorMercadoPago(config)
    df_mp = lector_mp.dataframe()

    ruta_archivo_erp = 'insumos/erp/ZOMAC - 2025-02-06T142211.293.xls'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    df_erp = lector.dataframe()

    df_cruce = df_mp.clone()

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

    # Realizar el cruce de datos
    df_cruce = df_cruce.join(
        df_erp.select(["numero_oc_comercial", "factura_erp", "cc_erp", "valor_fv_erp", "aux_erp"]),  # Seleccionamos solo las columnas necesarias
        left_on="num_oms",  # Columna en df_cruce
        right_on="numero_oc_comercial",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    )

    # Mostrar resultados del cruce
    print("\nResultados del cruce:", df_cruce.height)

    df_cruce = df_cruce.with_columns([
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "iva")).alias("iva"),
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "fuente")).alias("fuente"),
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "ica_bogota")).alias("ica")
    ])

    # Realizar el cruce y crear columna de diferencia con ajuste de valores cercanos a 0
    df_cruce = df_cruce.with_columns([
        pl.when(
            pl.col("valor_fv_erp").sub(pl.col("monto_bruto_operacion")).abs() < 1
        ).then(
            pl.lit(0)  # Si la diferencia absoluta es menor a 1, asignar 0
        ).otherwise(
            pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion")  # Mantener la diferencia original
        ).alias("diferencia_valores")
    ])

     # Definir orden específico de columnas
    columnas_ordenadas = [
        'numero_identificacion',
        "numero_identificacion_limpio",
        "num_oms",
        "factura_erp",
        "valor_fv_erp",
        "cc_erp",
        "aux_erp",
        "diferencia_valores",
        "monto_bruto_operacion",
        "iva",
        "fuente",
        "ica",
    ]

    # Reordenar columnas manteniendo el resto
    df_cruce = df_cruce.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    df_cruce = df_cruce.sort([
        "factura_erp",  # Orden por defecto (ascendente)
        "cc_erp",
        pl.col("fecha_aprobacion").sort_by("fecha_aprobacion", descending=False)  # Orden explícito
    ])


    # Filtrar registros con diferencia negativa
    df_cruce_negativos = df_cruce.filter(
        pl.col("diferencia_valores") < 0
    )

    # Obtener lista de facturas existentes
    facturas_existentes = df_cruce.select("factura_erp").unique().to_series().to_list()


    # Obtener todas las columnas que no vienen del select de df_erp
    columnas_a_nulificar = [
        col for col in df_cruce_negativos.columns
        if col not in ["cc_erp", "factura_erp", "valor_fv_erp", "aux_erp"]
    ]
    # Realizar el cruce y filtrado para traer todos los registros coincidentes
    df_cruce_cedulas = df_cruce_negativos.join(
        df_erp.select([
            "cc_erp",
            "factura_erp",
            "valor_fv_erp",
            "aux_erp"
        ]),
        left_on="cc_erp",
        right_on="cc_erp",
        how="inner"  # Inner join para mantener solo coincidencias
    ).filter(
        # Solo mantener registros donde la factura_erp no exista en el cruce original
        ~pl.col("factura_erp_right").is_in(facturas_existentes)
    ).with_columns([
        # Usar los valores nuevos cuando existan, mantener los originales si no
        # Establecer todas las columnas no incluidas en el select a None
        *[pl.lit(None).alias(col) for col in columnas_a_nulificar],
        # Usar los valores nuevos del join
        pl.col("factura_erp_right").alias("factura_erp"),
        pl.col("valor_fv_erp_right").alias("valor_fv_erp"),
        pl.col("aux_erp_right").alias("aux_erp")
    ]).drop([
        # Eliminar columnas duplicadas del join
        "factura_erp_right",
        "valor_fv_erp_right",
        "aux_erp_right"
    ])

    df_cruce_cedulas_facturas = unir_dataframes_cruce(df_cruce, df_cruce_cedulas)


    df_canceladas = obtener_facturas_canceladas(df_cruce_cedulas_facturas)
    listados_facturas_canceladas = df_canceladas.select("factura_erp").unique().to_series().to_list()
    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        ~pl.col("factura_erp").is_in(listados_facturas_canceladas)
    )

    df_devoluciones = obtener_facturas_devueltas(df_cruce_cedulas_facturas)
    listados_facturas_devueltas = df_devoluciones.select("factura_erp").unique().to_series().to_list()
    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        ~pl.col("factura_erp").is_in(listados_facturas_devueltas)
    )


    # Filtrar registros con diferencia negativa
    listados_facturas_negativas = df_cruce_cedulas_facturas.filter(
        pl.col("diferencia_valores") < 0
    ).select("factura_erp").unique().to_series().to_list()

    print('listados_facturas_negativas: ', listados_facturas_negativas)
    df_facturas_negativas = df_cruce.filter(
        # Solo mantener registros donde la factura_erp no exista en el cruce original
        pl.col("factura_erp").is_in(listados_facturas_negativas)
    )

    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        # Solo mantener registros donde la diferecnia sea mayor a 0
        (~pl.col("factura_erp").is_in(listados_facturas_negativas)) &
        ((pl.col("descripcion").str.to_lowercase().eq("payment")))
    )

    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.with_columns([
        pl.lit(0).cast(pl.Int64).alias("es_multiple")
    ])

    df_facturas_dobles = procesar_facturas_agrupadas(df_cruce_cedulas_facturas)
    listados_facturas_dobles = df_facturas_dobles.select("factura_erp").unique().to_series().to_list()
    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        ~pl.col("factura_erp").is_in(listados_facturas_dobles)
    )

    # Unir df_cruce_cedulas_facturas con df_facturas_dobles
    df_cruce_cedulas_facturas = unir_dataframes_cruce(df_cruce_cedulas_facturas, df_facturas_dobles)

    # Obtener registros sin factura_erp (nulos o vacíos)
    df_sin_factura = df_cruce.filter(
        pl.col("factura_erp").is_null() |
        (pl.col("factura_erp").cast(pl.Utf8).fill_null("").str.strip_chars().eq(""))
    )

    dataframes_a_exportar = {
        "Informacion original": df_cruce,
        "Cruce Principal": df_cruce_cedulas_facturas,
        "Facturas devolucion": df_devoluciones,
        "Facturas Canceladas": df_canceladas,
        "Facturas dobles": df_facturas_dobles,
        "Facturas sin cruzar": df_sin_factura,
    }

    exportar_multiples_dataframes_excel(dataframes_a_exportar, "reporte_completo")




if __name__ == "__main__":
    main()