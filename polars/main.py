from lectores.lector_oms import LectorOMS
from lectores.lector_mercadopago import LectorMercadoPago
from lectores.lector_erp import LectorERP
from lectores.lector_addi import LectorADDI
from lectores.lector_mercadolibre import LectorMercadoLibre

from lectores.sistecredito.lector_facturas import LectorSisCredFacturas
from lectores.sistecredito.lector_pagare import LectorSisCredPagare

from cruces.cruce_oms_erp_mp_ml import CruceOmsErpMpMl

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

                if df is None:
                    continue

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


    if facturas_procesadas:
        # Paso 4: Construir DataFrame resultante
        df_procesadas = pl.concat(facturas_procesadas) if facturas_procesadas else pl.DataFrame()

        # Añadir columna es_multiple con valor por defecto 0
        df_procesadas = df_procesadas.with_columns([
            pl.lit(1).cast(pl.Int64).alias("es_multiple")
        ])


        return df_procesadas

    return None

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

def obtener_cedula_asociada(factura: str, d_referencias: dict[str, str]) -> str:
    """
    Función que busca en el diccionario de referencias si alguna clave se encuentra contenida
    en el valor de factura. Si se encuentra, retorna la cédula asociada;
    en caso contrario retorna una cadena vacía.

    Args:
        factura (str): Valor de la factura (factura_codigo).
        d_referencias (dict[str, str]): Diccionario con referencia_sistecredito y cédula_cliente

    Returns:
        str: Cédula asociada o cadena vacía si no se encuentra coincidencia.
    """
    for referencia, cedula in d_referencias.items():
        if factura in referencia:
            return cedula
    return ""

def main():
    """
    Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
    """
    #prueba_cruce_oms_mercado_pago_clase()

    prueba_cruce_addi_erp()

    #prueba_sistecredito()


@medir_rendimiento
def prueba_sistecredito():

    ruta_carpeta_oms = 'insumos/oms/'
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()

    ruta_carpeta_factura = 'insumos/sistecredito/facturas/Facturas_pagadas SISTECREDITO.xlsx'
    config = ConfiguracionLector(ruta_archivo=ruta_carpeta_factura)
    lector_factura = LectorSisCredFacturas(config)
    df_factura = lector_factura.dataframe()

    ruta_carpeta_pagare = 'insumos/sistecredito/pagares/PAGARES SISTECREDITO.xlsx'
    config = ConfiguracionLector(ruta_archivo=ruta_carpeta_pagare)
    lector_pagare = LectorSisCredPagare(config)
    df_pagare = lector_pagare.dataframe()

    ruta_archivo_erp = 'insumos/erp/ERP.xls'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    df_erp = lector.dataframe()

    df_filtrado_oms = df_oms.filter(
        (pl.col("forma_pago_1") == "SISTECREDITO") |
        (pl.col("forma_pago_2") == "SISTECREDITO") |
        (pl.col("forma_pago_3") == "SISTECREDITO")
    )

    # Crear la nueva columna 'referencia_sistecredito' que contenga el valor de la columna de referencia correspondiente
    # dependiendo de cuál de las formas de pago es igual a "SISTECREDITO"
    df_filtrado_oms = df_filtrado_oms.with_columns(
        pl.when(pl.col("forma_pago_1") == "SISTECREDITO")
        .then(pl.col("forma_pago_1_referencia"))
        .when(pl.col("forma_pago_2") == "SISTECREDITO")
        .then(pl.col("forma_pago_2_referencia"))
        .when(pl.col("forma_pago_3") == "SISTECREDITO")
        .then(pl.col("forma_pago_3_referencia"))
        .otherwise(None)
        .alias("referencia_sistecredito")
    ).select(['cedula_cliente', 'referencia_sistecredito'])


    # 🔹 Realizar un LEFT JOIN en base a 'almacen' y 'consecutivo_pagare'
    df_cruce_factura_pagare = df_factura.join(
        df_pagare.select(["almacen", "consecutivo_pagare", "documento_identidad"]),
        on=["almacen", "consecutivo_pagare"],
        how="left"
    )

    df_cruce_erp = df_cruce_factura_pagare.join(
        df_erp.select(["cc_erp", "aux_erp", "factura_erp", "valor_fv_erp"]),
        left_on="documento_identidad", # Columna en df_cruce_factura_pagare
        right_on="cc_erp", # Columna en df_cruce_erp
        how="left"
    ).with_columns(
        pl.lit("").alias("cedula_oms")
    )

    listado_faturas_sin_cruzar = [
        registro for registro in df_cruce_erp.filter(pl.col('factura_erp').is_null()).select("factura_codigo").unique().to_series().to_list()
        if registro is not None and len(registro) > 0
    ]

    df_facturas_sin_cruce = df_cruce_erp.filter(
        pl.col("factura_codigo").is_in(listado_faturas_sin_cruzar)
    )

    # Primero, creamos un diccionario que relacione cada referencia_sistecredito (clave) con la cédula asociada (valor)
    # Esto asume que df_filtrado_oms contiene ambas columnas: 'referencia_sistecredito' y 'cedula_cliente'
    diccionario_referencias: dict[str, str] = {
        fila["referencia_sistecredito"]: fila["cedula_cliente"]
        for fila in df_filtrado_oms.to_dicts()
        if fila["referencia_sistecredito"] is not None
    }

    # Aplicamos la función en df_facturas_sin_cruce usando pl.apply para crear la nueva columna 'cedula_oms'
    df_facturas_sin_cruce = df_facturas_sin_cruce.with_columns(
        pl.col("factura_codigo")
        .map_elements(lambda codigo: obtener_cedula_asociada(codigo, diccionario_referencias), return_dtype=pl.Utf8)
        .alias("cedula_oms")
    ).filter(pl.col('cedula_oms') != '').select(['factura_codigo', 'cedula_oms'])



    df_cruce_erp = df_cruce_erp.join(
        df_facturas_sin_cruce.select(["factura_codigo", "cedula_oms"]),
        left_on="factura_codigo", # Columna en df_cruce_erp
        right_on="factura_codigo", # Columna en df_facturas_sin_cruce
        how="left"
    ).with_columns(
        pl.col('cedula_oms_right').alias('cedula_oms')
    )

    # Filtrar los registros que tienen un valor en "cedula_oms" (no nulo y no vacío)
    df_con_cedula = df_cruce_erp.filter(
        (pl.col("cedula_oms").is_not_null()) & (pl.col("cedula_oms") != "")
    )

    # Filtrar los registros que NO tienen valor en "cedula_oms"
    df_sin_cedula = df_cruce_erp.filter(
        (pl.col("cedula_oms").is_null()) | (pl.col("cedula_oms") == "")
    )

    # Realizar el join solo en los registros que tienen "cedula_oms" para actualizar sus valores
    df_con_cedula_actualizada = df_con_cedula.join(
        df_erp.select(["cc_erp", "aux_erp", "factura_erp", "valor_fv_erp"]),
        left_on="cedula_oms",    # Clave en df_con_cedula
        right_on="cc_erp",        # Clave en df_erp
        how="left"
    ).with_columns([
        # Usamos coalesce para que, en caso de no obtener un valor, se mantenga el original
        pl.coalesce([pl.col("aux_erp_right"), pl.col("aux_erp")]).alias("aux_erp"),
        pl.coalesce([pl.col("factura_erp_right"), pl.col("factura_erp")]).alias("factura_erp"),
        pl.coalesce([pl.col("valor_fv_erp_right"), pl.col("valor_fv_erp")]).alias("valor_fv_erp")
    ]).drop(["aux_erp_right", "factura_erp_right", "valor_fv_erp_right"])


    # Combinar nuevamente los registros actualizados con aquellos que no tenían "cedula_oms"
    df_cruce_erp = unir_dataframes_cruce(df_con_cedula_actualizada, df_sin_cedula)

    # Separar la columna en dos partes usando `.struct`
    df_cruce_erp = df_cruce_erp.with_columns(
        df_cruce_erp["factura_erp"]
        .str.split_exact("-", 1)
        .alias("factura_erp_split")  # Se convierte en un Struct con dos campos
    )

    # Acceder a los campos del Struct correctamente
    df_cruce_erp = df_cruce_erp.with_columns([
        pl.col("factura_erp_split").struct.field("field_0").alias("tipo_factura"),
        pl.col("factura_erp_split").struct.field("field_1").alias("numero_factura")
    ]).drop(['factura_erp_split'])



    df_agrupado = df_cruce_erp.group_by(["almacen", "consecutivo_pagare", "fecha_creacion", "documento_identidad"]).agg(
       pl.col("valor_fv_erp").sum().alias("total_valor_factura")
    )

    # Agrupar por las columnas del sort y sumar la columna 'valor'
    df_cruce_erp = df_cruce_erp.join(
        df_agrupado.select(["almacen", "consecutivo_pagare", "fecha_creacion", "documento_identidad", "total_valor_factura"]),
        on=["almacen", "consecutivo_pagare", "fecha_creacion", "documento_identidad"],
        how="left"
    )

    df_cruce_erp = df_cruce_erp.with_columns(
         pl.when((pl.col("valor_factura") - pl.col("total_valor_factura")).abs() < 1)
            .then(pl.lit(0))
            .otherwise(pl.col("valor_factura") - pl.col("total_valor_factura"))
            .cast(pl.Int64)  # Castear a entero
            .alias("diferencia")
    )

    # Ordenar por "almacen", "aux_erp", "consecutivo_pagare", "documento_identidad"
    df_cruce_erp = df_cruce_erp.select([
        'almacen',
        'consecutivo_pagare',
        'factura_codigo',
        'fecha_creacion',
        'valor_factura',
        'valor_neto_pagar',
        'documento_identidad',
        'cedula_oms',
        'aux_erp',
        'factura_erp',
        'tipo_factura',
        'numero_factura',
        'valor_fv_erp',
        'total_valor_factura',
        'diferencia',

    ])
    df_cruce_erp = df_cruce_erp.sort(["almacen", "consecutivo_pagare", "aux_erp", "fecha_creacion", "documento_identidad"])


    # Separar facturas diferentes de 0
    df_facturas_diferentes_cero = df_cruce_erp.filter(pl.col("diferencia") != 0)


    df_cruce_erp = df_cruce_erp.filter(pl.col("diferencia") == 0)


    #print(df_facturas_diferentes_cero.filter(pl.col("diferencia") == 0))
    #raise NotImplemented('STOP')



    dataframes_a_exportar = {
        "Factura": df_factura,
        "Pagare": df_pagare,
        "Cruce factura pagare": df_cruce_factura_pagare,
        "Facturas devolucion": df_facturas_diferentes_cero,
        "Cruce erp": df_cruce_erp,
    }




    exportar_multiples_dataframes_excel(dataframes_a_exportar, "reporte_completo_siscredito")


@medir_rendimiento
def prueba_cruce_oms_mercado_pago_clase():

    ruta_carpeta_oms = 'insumos/oms/'
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()

    ruta_carpeta_mercadolibre = 'insumos/mercadolibre/'
    config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_mercadolibre)
    lector_mercadolibre = LectorMercadoLibre(config)
    df_mercadolibre = lector_mercadolibre.dataframe()

    ruta_archivo_mercdaopago = 'insumos/mercadopago/MERCADOPAGO 172.xlsx'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_mercdaopago)
    lector_mp = LectorMercadoPago(config)
    df_mp = lector_mp.dataframe()

    ruta_archivo_erp = 'insumos/erp/ERP.xls'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    df_erp = lector.dataframe()

    df_cruce = df_mp.clone()

    print("\nCantidad elementos antes cruce oms:", df_cruce.height)

    # 1 Realizar el cruce de datos de mercadopago con la OMS
    df_cruce = df_cruce.join(
        df_oms.select(["orden_externa_limpio", "consecutivo"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_identificacion_limpio",  # Columna en df_cruce
        right_on="orden_externa_limpio",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    ).rename({"consecutivo": "num_oms"})

    print("\nCantidad elementos despues cruce oms:", df_cruce.height)

    print(df_cruce.select(["numero_identificacion_limpio"]))

    # 1.1 Validar que si se ecnuntre información en las formas de pago de la oms
    df_sin_cruzar = df_cruce.filter(
        pl.col("num_oms").is_null() |
        (pl.col("num_oms").cast(pl.Utf8).fill_null("").str.strip_chars().eq(""))
    )
    listados_registros_sin_cruzar = [
        registro for registro in df_sin_cruzar.select("numero_identificacion_limpio").unique().to_series().to_list()
        if registro is not None
    ]

    df_cruce = df_cruce.filter(
        ~pl.col("numero_identificacion_limpio").is_in(listados_registros_sin_cruzar)
    )


    df_cruce_referencia = df_sin_cruzar.join(
        df_oms.select(["referencia_mercadopago", "consecutivo"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_identificacion_limpio",  # Columna en df_sin_cruzar
        right_on="referencia_mercadopago",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    ).with_columns(
        pl.col("consecutivo").alias("num_oms")
    ).drop("consecutivo")

    df_cruce = unir_dataframes_cruce(df_cruce, df_cruce_referencia)


    # Mostrar resultados del cruce
    print("\nResultados del cruce:", df_cruce.height)

    # 2 Realizar el cruce de datos mercadopago-OMS con ERP
    df_cruce = df_cruce.join(
        df_erp.select(["numero_oc_comercial", "factura_erp", "cc_erp", "valor_fv_erp", "aux_erp"]),  # Seleccionamos solo las columnas necesarias
        left_on="num_oms",  # Columna en df_cruce
        right_on="numero_oc_comercial",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    )


    # Mostrar resultados del cruce
    print("\nResultados del cruce:", df_cruce.height)

    # Desagregando la columna impuestos_desagregados en iva, fuente, ica
    df_cruce = df_cruce.with_columns([
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "iva")).alias("iva"),
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "fuente")).alias("fuente"),
        pl.col("impuestos_desagregados").map_elements(lambda x: extraer_valores_impuestos(x, "ica")).alias("ica")
    ])

    # Realizar el cruce y crear columna de diferencia con ajuste de valores cercanos a 0
    df_cruce = df_cruce.with_columns([
        pl.when(
            (pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion")).abs() < 0.02
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
    ]).with_columns([
        pl.lit(0).cast(pl.Int64).alias("cruzado_en_mercadolibre")
    ])

    # 3 Obtener registros sin factura_erp (nulos o vacíos)
    df_sin_factura = df_cruce.filter(
        pl.col("factura_erp").is_null() |
        (pl.col("factura_erp").cast(pl.Utf8).fill_null("").str.strip_chars().eq(""))
    )

    df_cruce_mercado_libre = df_sin_factura.join(
        df_mercadolibre.select([
            "numero_identificacion",
            "numero_identificacion_limpio",
            pl.col("cedula")
        ]),
        left_on="numero_identificacion_limpio",  # Columna en df_mercadolibre
        right_on="numero_identificacion_limpio",  # Columna en df_erp
        how="left"  # Mantener todos los registros de df_sin_factura
    ).with_columns([
        pl.coalesce([pl.col("cedula"), pl.col("cc_erp")]).alias("cc_erp"),
        pl.lit(1).cast(pl.Int64).alias("cruzado_en_mercadolibre")
    ]).drop("cedula")# Eliminamos la columna temporal


    df_cruce_mercado_libre = df_cruce_mercado_libre.join(
        df_erp.select(["factura_erp", "cc_erp", "valor_fv_erp", "aux_erp"]),  # Seleccionamos solo las columnas necesarias
        left_on="cc_erp",  # Columna en df_cruce
        right_on="cc_erp",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    ).with_columns([
        pl.coalesce([pl.col("factura_erp_right"), pl.col("factura_erp")]).alias("factura_erp"),
        pl.coalesce([pl.col("valor_fv_erp_right"), pl.col("valor_fv_erp")]).alias("valor_fv_erp"),
        pl.coalesce([pl.col("aux_erp_right"), pl.col("aux_erp")]).alias("aux_erp"),
    ]).drop(
        [
            'factura_erp_right',
            'valor_fv_erp_right',
            'aux_erp_right',
            'numero_identificacion_right'
        ]
    )

    # Filtrar registros donde "nombre" NO es nulo o vacío
    # df_con_factura = df_cruce_mercado_libre.filter(pl.col("factura_erp").is_not_null() & (pl.col("factura_erp") != ""))

    # Filtrar registros donde "nombre" ES nulo o vacío
    df_sin_facturas_asociadas = df_cruce_mercado_libre.filter(pl.col("factura_erp").is_null() | (pl.col("factura_erp") == ""))

    print("\nResultados del cruce antes mercado libre:", df_cruce.height)


    listados_facturas_mercadolibre = [
        factura for factura in df_cruce_mercado_libre.select("factura_erp").unique().to_series().to_list()
        if factura is not None
    ]
    df_cruce = df_cruce.filter(
        ~pl.col("factura_erp").is_in(listados_facturas_mercadolibre)
    )

    df_cruce = unir_dataframes_cruce(df_cruce, df_cruce_mercado_libre)

    print("\nResultados del cruce despues mercado libre:", df_cruce.height)


    df_cruce = df_cruce.with_columns([
        pl.when(
             (pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion")).abs() < 0.02
        ).then(
            pl.lit(0)  # Si la diferencia absoluta es menor a 0,02, asignar 0
        ).otherwise(
            pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion")  # Mantener la diferencia original
        ).alias("diferencia_valores")
    ])

            # Reordenar columnas manteniendo el resto
    df_cruce = df_cruce.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    columnas_ordenadas.remove('diferencia_valores')
    # Reordenar columnas manteniendo el resto
    df_sin_facturas_asociadas = df_sin_facturas_asociadas.select(
        columnas_ordenadas + [
            col for col in df_sin_facturas_asociadas.columns
            if col not in columnas_ordenadas
        ]
    )


    listados_facturas_reseve = [
         factura for factura in df_sin_facturas_asociadas.filter(pl.col("descripcion").str.contains('reserve')).select("numero_identificacion_limpio").unique().to_series().to_list()
         if factura is not None and factura != ''
    ]
    df_facturas_reservadas = df_cruce.filter(
        pl.col("numero_identificacion_limpio").is_in(listados_facturas_reseve)
    )

    # Agrupar por "categoria" y sumar las columnas "ventas" y "descuentos"
    df_facturas_reservadas_sumatoria = df_facturas_reservadas.group_by("numero_identificacion_limpio").agg(
        pl.col("iva").sum().alias("total_iva"),
        pl.col("fuente").sum().alias("total_fuente"),
        pl.col("ica").sum().alias("total_ica"),
        pl.col("comision_mercado_pago_incluye_iva").sum().alias("total_comision_mercado_pago_incluye_iva"),
    ).with_columns(
        (pl.sum_horizontal(["total_iva", "total_fuente", "total_ica", "total_comision_mercado_pago_incluye_iva"])).alias("diferencia_final")
    )

    listados_sin_facturas_pendientes = [
         factura for factura in df_facturas_reservadas_sumatoria.filter(pl.col("diferencia_final") == 0).select("numero_identificacion_limpio").unique().to_series().to_list()
         if factura is not None and factura != ''
    ]

    df_facturas_pendientes = df_sin_facturas_asociadas.filter(
        pl.col("numero_identificacion_limpio").is_in(listados_sin_facturas_pendientes)
    )

    df_sin_facturas_asociadas = df_sin_facturas_asociadas.filter(
        ~pl.col("numero_identificacion_limpio").is_in(listados_sin_facturas_pendientes)
    )

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

    df_negativas = df_cruce_cedulas_facturas.filter(pl.col("diferencia_valores") < 0)
    listados_facturas_negativas = df_negativas.select("factura_erp").unique().to_series().to_list()
    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        ~pl.col("factura_erp").is_in(listados_facturas_negativas)
    )

    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
        # Solo mantener registros donde la diferecnia sea mayor a 0
        (~pl.col("factura_erp").is_in(listados_facturas_negativas)) &
        ((pl.col("descripcion").str.to_lowercase().eq("payment")))
    )

    df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.with_columns([
        pl.lit(0).cast(pl.Int64).alias("es_multiple")
    ])


    # df_facturas_dobles = procesar_facturas_agrupadas(df_cruce_cedulas_facturas)
    # listados_facturas_dobles = df_facturas_dobles.select("factura_erp").unique().to_series().to_list()
    # df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
    #     ~pl.col("factura_erp").is_in(listados_facturas_dobles)
    # )

    # # Unir df_cruce_cedulas_facturas con df_facturas_dobles
    # df_cruce_cedulas_facturas = unir_dataframes_cruce(df_cruce_cedulas_facturas, df_facturas_dobles)

    dataframes_a_exportar = {
        "Informacion original": df_cruce,
        "Cruce principal": df_cruce_cedulas_facturas,
        "Facturas devolucion": df_devoluciones,
        "Facturas canceladas": df_canceladas,
        #"Facturas duplicadas": df_facturas_dobles,
        "Facturas negativas": df_negativas,
        "Facturas reservas canceladas": df_facturas_pendientes,
        "Facturas sin cruzar": df_sin_facturas_asociadas,
    }

    exportar_multiples_dataframes_excel(dataframes_a_exportar, "reporte_completo")


@medir_rendimiento
def prueba_cruce_addi_erp():
    ruta_archivo_erp = 'insumos/erp/ERP.xls'  # Ajusta esta ruta según tu estructura
    config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    lector = LectorERP(config)
    df_erp = lector.dataframe()

    ruta_carpeta_addi = 'insumos/addi/'  # Ajusta esta ruta según tu estructura
    config_addi = ConfiguracionLector(ruta_carpeta=ruta_carpeta_addi)
    lector_addi = LectorADDI(config_addi)
    df_addi = lector_addi.dataframe()

    df_cruce = df_addi.clone()

        # Realizar el cruce de datos
    df_cruce = df_cruce.join(
        df_erp.select(["cc_erp", "aux_erp", "factura_erp", "valor_fv_erp"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_documento",  # Columna en df_cruce
        right_on="cc_erp",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    )

    # Realizar el cruce y crear columna de diferencia con ajuste de valores cercanos a 0
    df_cruce = df_cruce.with_columns([
        pl.when(
            pl.col("valor_fv_erp").sub(pl.col("total_ventas")).abs() < 1
        ).then(
            pl.lit(0)  # Si la diferencia absoluta es menor a 1, asignar 0
        ).otherwise(
            pl.col("valor_fv_erp") - pl.col("total_ventas")  # Mantener la diferencia original
        ).alias("diferencia_valores")
    ])



   # Definir orden específico de columnas
    columnas_ordenadas = [
        "numero_documento",
        "total_ventas",
        "factura_erp",
        "aux_erp",
        "valor_fv_erp",
        "diferencia_valores"
    ]

    # Reordenar columnas manteniendo el resto
    df_cruce = df_cruce.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    dataframes_a_exportar = {
        "Cruce": df_cruce
    }

    exportar_multiples_dataframes_excel(dataframes_a_exportar, "addi_cruce")


if __name__ == "__main__":
    main()