from .lectores.lector_oms import LectorOMS
from .lectores.lector_mercadopago import LectorMercadoPago
from .lectores.lector_erp import LectorERP
from .lectores.lector_addi import LectorADDI
from .lectores.lector_mercadolibre import LectorMercadoLibre

from .lectores.sistecredito.lector_facturas import LectorSisCredFacturas
from .lectores.sistecredito.lector_pagare import LectorSisCredPagare

from .lectores.modelos import ConfiguracionLector

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
    Decorador que mide el rendimiento de una función, calculando el tiempo de ejecución
    y el uso de memoria RAM.

    Características:
    - Captura el tiempo inicial y final de ejecución
    - Mide el uso de memoria RAM antes y después de la ejecución
    - Imprime estadísticas de rendimiento

    Args:
        funcion (Callable): Función a decorar

    Returns:
        Callable: Función decorada que incluye medición de rendimiento

    Ejemplo:
        @medir_rendimiento
        def mi_funcion():
            # código de la función
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
    Exporta un DataFrame de Polars a un archivo Excel con nombre timestampeado.

    Proceso:
    1. Crea una carpeta 'resultados' si no existe
    2. Genera un nombre de archivo único con timestamp
    3. Exporta el DataFrame a Excel

    Args:
        df (pl.DataFrame): DataFrame a exportar
        nombre_base (str): Nombre base para el archivo (se añadirá timestamp)

    Returns:
        None: El archivo se guarda en disco en la carpeta 'resultados'

    Ejemplo:
        exportar_resultados_excel(df, "reporte_ventas")
        # Crea: resultados/reporte_ventas_20231201_143022.xlsx
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
    Une verticalmente dos DataFrames asegurando compatibilidad de columnas y reportando estadísticas.

    Proceso:
    1. Valida que los DataFrames no estén vacíos
    2. Muestra estadísticas pre-unión
    3. Realiza la concatenación vertical
    4. Muestra estadísticas post-unión

    Args:
        df_principal (pl.DataFrame): DataFrame base para la unión
        df_adicional (pl.DataFrame): DataFrame a añadir

    Returns:
        pl.DataFrame: DataFrame unificado con todos los registros

    Raises:
        Exception: Si hay error durante la unión
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
    nombre_base: str,
    carpeta_resultados: str = "resultados"
) -> None:
    """
    Exporta múltiples DataFrames a diferentes hojas de un archivo Excel.

    Proceso:
    1. Crea carpeta de resultados si no existe
    2. Genera nombre de archivo con timestamp
    3. Crea archivo Excel con múltiples hojas
    4. Convierte cada DataFrame de Polars a Pandas
    5. Exporta cada DataFrame a su hoja correspondiente

    Args:
        dataframes (dict[str, pl.DataFrame]): Diccionario de {nombre_hoja: DataFrame}
        nombre_base (str): Nombre base para el archivo Excel
        carpeta_resultados (str): Ruta de la carpeta de salida

    Returns:
        None: Guarda el archivo en disco

    Raises:
        Exception: Si hay error durante la exportación
    """
    try:
        # Crear carpeta si no existe
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
    Identifica y filtra las facturas canceladas basándose en la suma de montos, la comisión y
    la presencia de refunds.

    Proceso:
      1. Para cada factura (agrupada por 'factura_erp'):
         - Suma el valor de 'monto_bruto_operacion' y se obtiene el total en 'suma_montos'
         - Suma el valor de 'comision_mercado_pago_incluye_iva' y se obtiene el total en 'suma_comision'
         - Verifica si en 'descripcion' aparece la palabra 'refund' (en minúsculas).
      2. Se marca la factura como cancelada (columna 'cancelado' = 1) si se cumple alguna de las siguientes condiciones:
         a) El valor absoluto de 'suma_montos' es menor a 2.5 y existe refund; o
         b) El valor absoluto de la suma de 'suma_montos' y 'suma_comision' es menor a 2.5.
      3. Se une el resultado con el DataFrame original filtrando solo las facturas canceladas.

    Args:
        df (pl.DataFrame): DataFrame que contiene las columnas 'factura_erp',
                           'monto_bruto_operacion', 'comision_mercado_pago_incluye_iva' y 'descripcion'.

    Returns:
        pl.DataFrame: DataFrame con solo las facturas canceladas.
    """
    df_cancelaciones = df.group_by("factura_erp").agg([
        # Suma total de montos
        pl.col("monto_bruto_operacion").sum().alias("suma_montos"),
        # Suma total de comisión
        pl.col("comision_mercado_pago_incluye_iva").sum().alias("suma_comision"),
        # Verificar si existe algún refund en la descripción
        pl.col("descripcion").str.to_lowercase().eq("refund").any().alias("tiene_refund")
    ]).with_columns([
        # Marcar como cancelado si se cumple alguna de las condiciones:
        # a) absolute(suma_montos) < 2.5 y tiene refund, O
        # b) absolute(suma_montos + suma_comision) < 2.5
        pl.when(
            (
                (pl.col("suma_montos").abs() < 2.5) & pl.col("tiene_refund")
            ) | ((pl.col("suma_montos") + pl.col("suma_comision")).abs() < 2.5)
        ).then(
            pl.lit(1)
        ).otherwise(
            pl.lit(0)
        ).alias("cancelado")
    ])

    # Unir el resultado con el DataFrame original y filtrar solo las facturas canceladas
    df_resultado = df.join(
        df_cancelaciones.select(["factura_erp", "cancelado"]),
        on="factura_erp",
        how="left"
    ).filter(
        pl.col("cancelado") == 1
    ).drop("cancelado")

    return df_resultado

def obtener_facturas_devueltas(df: pl.DataFrame) -> pl.DataFrame:
    """
    Identifica y filtra las facturas con devoluciones basándose en montos negativos
    y presencia de refunds.

    Proceso:
    1. Agrupa por factura_erp
    2. Suma montos y verifica refunds
    3. Marca facturas con suma negativa y refund como devueltas
    4. Filtra solo facturas devueltas

    Args:
        df (pl.DataFrame): DataFrame con columnas factura_erp, monto_bruto_operacion y descripcion

    Returns:
        pl.DataFrame: DataFrame con solo las facturas devueltas
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

    Proceso:
    - Itera sobre la lista de valores
    - Compara cada valor con el valor1 usando el margen de error
    - Retorna True si encuentra alguna coincidencia dentro del margen

    Args:
        valor1 (float): Valor a comparar
        valores (list[float]): Lista de valores para comparar
        margen (float): Margen de error permitido (default 0.01)

    Returns:
        bool: True si encuentra algún valor similar dentro del margen
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

@medir_rendimiento
def cruce_sistecredito(df_oms, df_factura, df_pagare, df_erp, carpeta_resultados):

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




    if not df_facturas_sin_cruce.is_empty():
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

    if not df_con_cedula.is_empty():
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

    df_sin_erp = df_facturas_diferentes_cero.filter(pl.col('factura_erp').is_null())
    df_facturas_diferentes_cero = df_facturas_diferentes_cero.filter(pl.col('factura_erp').is_not_null())


    df_aux_facturas_revision = df_facturas_diferentes_cero.with_columns(
         pl.when((pl.col("valor_factura") - pl.col("valor_fv_erp")).abs() < 2)
            .then(pl.lit(0))
            .otherwise(pl.col("valor_factura") - pl.col("valor_fv_erp"))
            .cast(pl.Int64)  # Castear a entero
            .alias("diferencia")
    )

    listado_faturas_revision = [
        registro for registro in df_aux_facturas_revision.filter(pl.col('diferencia')==0).select("documento_identidad").unique().to_series().to_list()
        if registro is not None and len(registro) > 0
    ]

    df_facturas_revision = df_aux_facturas_revision.filter(
        pl.col("documento_identidad").is_in(listado_faturas_revision)
    )

    df_facturas_diferentes_cero = df_aux_facturas_revision.filter(
        ~pl.col("documento_identidad").is_in(listado_faturas_revision)
    )

    saldo_a_favor = df_facturas_diferentes_cero.filter(pl.col('diferencia')>=0)
    saldo_por_cobrar = df_facturas_diferentes_cero.filter(pl.col('diferencia')<0)

    dataframes_a_exportar = {
        "Factura": df_factura,
        "Pagare": df_pagare,
        "Cruce factura pagare": df_cruce_factura_pagare,
        "Sin ERP": df_sin_erp,
        "Facts revision": df_facturas_revision,

        "Saldo a favor": saldo_a_favor,
        "Saldo por cobrar": saldo_por_cobrar,

        "Cruce erp": df_cruce_erp,
    }




    exportar_multiples_dataframes_excel(dataframes_a_exportar, "reporte_completo_siscredito", carpeta_resultados)


@medir_rendimiento
def cruce_oms_mercado_pago_clase(df_oms, df_mercadolibre, df_mp, df_erp, carpeta_resultados):
    """
    Realiza el cruce de información entre las tablas de OMS, Mercado Pago, Mercado Libre y ERP.
    Este proceso permite conciliar las transacciones y validar la información financiera entre sistemas.

    El proceso se divide en las siguientes etapas:

    1. Carga de Datos:
       - Lee datos de OMS (Sistema de Gestión de Órdenes)
       - Lee datos de Mercado Libre
       - Lee datos de Mercado Pago
       - Lee datos del ERP (Sistema de Planificación de Recursos Empresariales)

    2. Proceso de Cruce Principal (OMS-MercadoPago):
       - Une datos de Mercado Pago con OMS usando número de identificación
       - Filtra registros sin cruzar para procesamiento posterior
       - Valida información en formas de pago de OMS

    3. Cruce con ERP:
       - Une datos del cruce anterior con ERP usando número de orden comercial
       - Calcula diferencias entre valores facturados y montos de operación
       - Desagrega impuestos (IVA, Fuente, ICA)

    4. Procesamiento de Casos Especiales:
       - Maneja facturas canceladas
       - Procesa devoluciones
       - Identifica y procesa reservas
       - Maneja facturas con valores negativos
       - Procesa facturas duplicadas

    5. Generación de Reportes:
       - Crea múltiples hojas de Excel con diferentes perspectivas:
         * Información original
         * Cruce principal
         * Facturas con devolución
         * Facturas canceladas
         * Facturas negativas
         * Facturas de reservas canceladas
         * Facturas sin cruzar

    Aspectos Importantes:
    - Usa tolerancia de 0.02 para comparaciones de valores monetarios
    - Maneja casos especiales de MercadoLibre separadamente
    - Implementa lógica para detectar facturas duplicadas y canceladas
    - Genera reportes detallados para análisis posterior

    Returns:
        None: Los resultados se exportan a un archivo Excel con múltiples hojas

    Raises:
        Exception: Cualquier error durante el proceso es capturado y logged
    """
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

    # 1.1 Validar que si se ecuentre información en las formas de pago de la oms
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

    if not df_mercadolibre.is_empty():
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

    else:
        df_sin_facturas_asociadas = df_sin_factura
        df_cruce = df_cruce.filter(~(pl.col("factura_erp").is_null() & (pl.col("factura_erp") == "")))


    print("\nResultados del cruce despues mercado libre:", df_cruce.height)


    df_cruce = df_cruce.with_columns([
        pl.when(
             (pl.col("valor_fv_erp") - pl.col("monto_bruto_operacion")).abs() < 2.5
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

    # Agrupar el DataFrame df_negativas por la columna cc_erp y calcular la diferencia
    # entre la suma de valor_fv_erp y la suma de monto_bruto_operacion para cada grupo.
    df_negativas_agrupadas = df_negativas.group_by("cc_erp").agg([
        pl.col("valor_fv_erp").sum().alias("suma_valor_fv_erp"),
        pl.col("monto_bruto_operacion").cast(pl.Float64).drop_nulls().mean().alias("promedio_monto_bruto_operacion")
    ]).with_columns([
        # Se resta la suma de monto_bruto_operacion de la suma de valor_fv_erp,
        # asignando el resultado a la columna diferencia_negativa
        (pl.col("suma_valor_fv_erp") - pl.col("promedio_monto_bruto_operacion")).alias("diferencia_valores")
    ]).with_columns([
        pl.col("diferencia_valores").cast(pl.Decimal(20, 2)).alias("diferencia_valores")
    ])

    # Se realiza un join entre df_negativas y df_negativas_agrupadas utilizando la columna "valor_fv_erp" como llave.
    # Solo se actualizarán los valores de la columna "diferencia_valores" en df_negativas para aquellos registros
    # que tengan correspondencia en df_negativas_agrupadas, reemplazando el valor original por el calculado en el agrupamiento.

    df_negativas = df_negativas.join(
        # Seleccionar únicamente las columnas necesarias del DataFrame agrupado
        df_negativas_agrupadas.select(["cc_erp", "diferencia_valores"]),
        left_on="cc_erp",   # Llave de unión en df_negativas
        right_on="cc_erp",  # Llave de unión en df_negativas_agrupadas
        how="left",
        suffix="_nuevo"         # Para evitar duplicidad de nombres de columnas
    ).with_columns(
        # Actualizar la columna "diferencia_valores" en df_negativas:
        # Si se encontró una coincidencia en el join (valor no nulo en la columna "diferencia_valores_nuevo"),
        # se reemplaza el valor original por el de "diferencia_valores_nuevo"; en caso contrario se conserva el original.
        pl.coalesce([pl.col("diferencia_valores_nuevo"), pl.col("diferencia_valores")])
        .alias("diferencia_valores")
    ).drop("diferencia_valores_nuevo")

    # Se realiza una validacion adicional para quitar aquellas facturas donde su monto si es cero pero agrupandola por la cedula
    df_crue_cero_negativas = df_negativas.filter(pl.col('diferencia_valores') == 0)
    df_cruce = unir_dataframes_cruce(df_cruce, df_crue_cero_negativas)

    df_negativas = df_negativas.filter(pl.col('diferencia_valores') != 0)

    # Calculamos la suma de "valor_fv_erp" para cada grupo identificado por "cc_erp"
    # sin usar group_by, utilizando funciones de ventana "over". Luego, seleccionamos
    # las columnas únicas y se castea el resultado a Decimal(20,2) asignándolo a "diferencia_valores".
    df_facturas_pendientes_agrupadas = df_facturas_pendientes.group_by("numero_identificacion_limpio").agg([
        pl.col("monto_bruto_operacion").cast(pl.Float64).sum().alias("suma_monto_bruto_operacion"),
    ]).with_columns([
        pl.col("suma_monto_bruto_operacion").cast(pl.Decimal(20, 2)).alias("diferencia_valores")
    ]).drop(
        'suma_monto_bruto_operacion'
    )

    df_facturas_pendientes = df_facturas_pendientes.join(
        # Seleccionar únicamente las columnas necesarias del DataFrame agrupado
        df_facturas_pendientes_agrupadas.select(["numero_identificacion_limpio", "diferencia_valores"]),
        left_on="numero_identificacion_limpio",   # Llave de unión en df_negativas
        right_on="numero_identificacion_limpio",  # Llave de unión en df_negativas_agrupadas
        how="left",
        suffix="_nuevo"         # Para evitar duplicidad de nombres de columnas
    ).with_columns(
        # Actualizar la columna "diferencia_valores" en df_negativas:
        # Si se encontró una coincidencia en el join (valor no nulo en la columna "diferencia_valores_nuevo"),
        # se reemplaza el valor original por el de "diferencia_valores_nuevo"; en caso contrario se conserva el original.
        pl.coalesce([pl.col("diferencia_valores_nuevo"), pl.col("diferencia_valores")]).alias("diferencia_valores")
    ).drop("diferencia_valores_nuevo").filter(pl.col('numero_identificacion_limpio').is_not_null())


    # Se valida que en facturas pendientes la diferecnia por nuemro_identifiacion sea cero
    df_facturas_pendientes_efecto_0 = df_facturas_pendientes.filter(pl.col('diferencia_valores') == 0)

    # Reordenar columnas manteniendo el resto
    df_canceladas = df_canceladas.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    # Reordenar columnas manteniendo el resto
    df_facturas_pendientes_efecto_0 = df_facturas_pendientes_efecto_0.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    # Reordenar columnas manteniendo el resto
    df_negativas = df_negativas.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    df_canceladas = unir_dataframes_cruce(df_canceladas, df_facturas_pendientes_efecto_0)

    df_facturas_a_favor = df_facturas_pendientes.filter(pl.col('diferencia_valores') != 0)

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


    # Filtrar los registros de df_de_din_facturas_asociadas donde:
    # - La columna "descripcion" contenga la palabra "shipping" (ignorando mayúsculas/minúsculas)
    # - O la columna "numero_identificacion" comience con "444"
    df_facturas_shipping = df_sin_facturas_asociadas.filter(
        (pl.col("descripcion").str.to_lowercase().str.contains("shipping")) |
        (pl.col("numero_identificacion").str.starts_with("444"))
    )

    # Filtrar los registros de df_sin_facturas_asociadas que NO tengan:
    # - La palabra "shipping" en la columna "descripcion" (ignorando mayúsculas/mínusculas)
    # - Y que su "numero_identificacion" NO comience con "444"
    df_sin_facturas_asociadas = df_sin_facturas_asociadas.filter(
        (~pl.col("descripcion").str.to_lowercase().str.contains("shipping")) &
        (~pl.col("numero_identificacion").str.starts_with("444"))
    )

    # Calculamos la suma de "monto_bruto_operacion" para cada grupo identificado por "cc_erp"
    # sin usar group_by, utilizando funciones de ventana "over". Luego, seleccionamos
    # las columnas únicas y se castea el resultado a Decimal(20,2) asignándolo a "diferencia_valores".
    df_sin_facturas_asociadas_agrupadas = df_sin_facturas_asociadas.group_by("numero_identificacion_limpio").agg([
        pl.col("monto_bruto_operacion").cast(pl.Float64).sum().alias("suma_monto_bruto_operacion"),
    ]).with_columns([
        pl.col("suma_monto_bruto_operacion").cast(pl.Decimal(20, 2)).alias("diferencia_valores")
    ]).drop(
        'suma_monto_bruto_operacion'
    )

    df_sin_facturas_asociadas = df_sin_facturas_asociadas.join(
        # Seleccionar únicamente las columnas necesarias del DataFrame agrupado
        df_sin_facturas_asociadas_agrupadas.select(["numero_identificacion_limpio", "diferencia_valores"]),
        left_on="numero_identificacion_limpio",   # Llave de unión en df_negativas
        right_on="numero_identificacion_limpio",  # Llave de unión en df_negativas_agrupadas
        how="left",
        suffix="_nuevo"         # Para evitar duplicidad de nombres de columnas
    ).with_columns(
        # Actualizar la columna "diferencia_valores" en df_negativas:
        # Si se encontró una coincidencia en el join (valor no nulo en la columna "diferencia_valores_nuevo"),
        # se reemplaza el valor original por el de "diferencia_valores_nuevo"; en caso contrario se conserva el original.
        pl.coalesce([pl.col("diferencia_valores_nuevo"), pl.col("diferencia_valores")]).alias("diferencia_valores")
    ).drop("diferencia_valores_nuevo").filter(pl.col('numero_identificacion_limpio').is_not_null())

    # Reordenar columnas manteniendo el resto
    df_sin_facturas_asociadas = df_sin_facturas_asociadas.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    df_sin_facturas_saldo_a_favor= df_sin_facturas_asociadas.filter(pl.col('diferencia_valores') > 0)
    df_sin_facturas_saldo_a_favor_con_cedula = df_sin_facturas_saldo_a_favor.filter(pl.col('cc_erp').is_not_null())

    df_sin_facturas_saldo_a_favor = df_sin_facturas_saldo_a_favor.filter(pl.col('cc_erp').is_null())
    df_sin_facturas_asociadas = df_sin_facturas_asociadas.filter(pl.col('diferencia_valores') <= 0)

    df_sin_facturas_asociadas_sin_cedula_negativa = df_sin_facturas_asociadas.filter((pl.col('diferencia_valores') < 0) & (pl.col('numero_identificacion').is_not_null()) & (pl.col('numero_identificacion') != ''))
    df_negativas = unir_dataframes_cruce(df_negativas, df_sin_facturas_asociadas_sin_cedula_negativa)

    df_sin_facturas_asociadas = df_sin_facturas_asociadas.filter((pl.col('diferencia_valores') == 0) | (pl.col('numero_identificacion').is_null()) | (pl.col('numero_identificacion') == ''))

    # Calculamos la suma de "monto_bruto_operacion" para cada grupo identificado por "cc_erp"
    # sin usar group_by, utilizando funciones de ventana "over". Luego, seleccionamos
    # las columnas únicas y se castea el resultado a Decimal(20,2) asignándolo a "diferencia_valores".
    df_sin_facturas_asociadas_agrupadas = df_sin_facturas_asociadas.group_by("id_operacion_mercado_pago").agg([
        pl.col("monto_bruto_operacion").cast(pl.Float64).sum().alias("suma_monto_bruto_operacion"),
    ]).with_columns([
        pl.col("suma_monto_bruto_operacion").cast(pl.Decimal(20, 2)).alias("diferencia_valores")
    ]).drop(
        'suma_monto_bruto_operacion'
    )

    df_sin_facturas_asociadas = df_sin_facturas_asociadas.join(
        # Seleccionar únicamente las columnas necesarias del DataFrame agrupado
        df_sin_facturas_asociadas_agrupadas.select(["id_operacion_mercado_pago", "diferencia_valores"]),
        left_on="id_operacion_mercado_pago",   # Llave de unión en df_negativas
        right_on="id_operacion_mercado_pago",  # Llave de unión en df_negativas_agrupadas
        how="left",
        suffix="_nuevo"         # Para evitar duplicidad de nombres de columnas
    ).with_columns(
        # Actualizar la columna "diferencia_valores" en df_negativas:
        # Si se encontró una coincidencia en el join (valor no nulo en la columna "diferencia_valores_nuevo"),
        # se reemplaza el valor original por el de "diferencia_valores_nuevo"; en caso contrario se conserva el original.
        pl.coalesce([pl.col("diferencia_valores_nuevo"), pl.col("diferencia_valores")]).alias("diferencia_valores")
    ).drop("diferencia_valores_nuevo").filter(pl.col('id_operacion_mercado_pago').is_not_null())


    df_saldo_negativo = df_sin_facturas_asociadas.filter(pl.col('diferencia_valores') < 0)
    df_negativas = unir_dataframes_cruce(df_negativas, df_saldo_negativo)

    df_saldo_positivo_sin_cc = df_sin_facturas_asociadas.filter(pl.col('diferencia_valores') >= 0)
    df_sin_facturas_saldo_a_favor = unir_dataframes_cruce(df_sin_facturas_saldo_a_favor, df_saldo_positivo_sin_cc)

    # Reordenar columnas manteniendo el resto
    df_devoluciones = df_devoluciones.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )

    df_negativas = unir_dataframes_cruce(df_negativas, df_devoluciones)

    dataframes_a_exportar = {
        "Info original": df_cruce,
        "Cruce fact 0": df_cruce_cedulas_facturas,
        #"Facts devolucion": df_devoluciones,
        "Facts canceladas": df_canceladas,
        "Saldo negativo": df_negativas,
        "Saldo a favor SIN CC": df_sin_facturas_saldo_a_favor,
        "Saldo a favor CON CC": df_sin_facturas_saldo_a_favor_con_cedula,
        "Shipping": df_facturas_shipping
    }


    exportar_multiples_dataframes_excel(dataframes_a_exportar, "reporte_completo", carpeta_resultados)


@medir_rendimiento
def cruce_addi_erp(df_erp, df_addi, carpeta_resultados):
    # ruta_archivo_erp = 'insumos/erp/ZOMAC de ADDD.xls'  # Ajusta esta ruta según tu estructura
    # config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
    # lector = LectorERP(config)
    # df_erp = lector.dataframe()

    # ruta_carpeta_addi = 'insumos/addi/'  # Ajusta esta ruta según tu estructura
    # config_addi = ConfiguracionLector(ruta_carpeta=ruta_carpeta_addi)
    # lector_addi = LectorADDI(config_addi)
    # df_addi = lector_addi.dataframe()

    df_cruce = df_addi.clone()

    # Realizar el cruce de datos
    df_cruce = df_cruce.join(
        df_erp.select(["cc_erp", "aux_erp", "factura_erp", "valor_fv_erp"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_documento",  # Columna en df_cruce
        right_on="cc_erp",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    )

    df_total_facturas = df_cruce.group_by(["numero_documento"]).agg(
       pl.col("valor_fv_erp").sum().alias("total_valor_factura")
    )

    df_cruce = df_cruce.join(
        df_total_facturas.select(["numero_documento", "total_valor_factura"]),  # Seleccionamos solo las columnas necesarias
        left_on="numero_documento",  # Columna en df_cruce
        right_on="numero_documento",  # Columna en df_oms
        how="left"  # Mantener todos los registros de df_cruce
    )


    # Realizar el cruce y crear columna de diferencia con ajuste de valores cercanos a 0
    df_cruce_transacciones = df_cruce.with_columns([
        pl.when(
            pl.col("total_valor_factura").sub(pl.col("total_ventas")).abs() < 1
        ).then(
            pl.lit(0)  # Si la diferencia absoluta es menor a 1, asignar 0
        ).otherwise(
            pl.col("total_ventas")  - pl.col("total_valor_factura")  # Mantener la diferencia original
        ).alias("diferencia_valores")
    ])



    # Actualizamos el valor de la columna "diferencia_valores" sin crear una nueva columna.
    # Si la diferencia absoluta entre "valor_fv_erp" y "total_ventas" es menor a 1,
    # entonces se asigna 0; en caso contrario, se mantiene el valor original de "diferencia_valores".
    df_cruce = df_cruce_transacciones.with_columns(
        pl.when((pl.col("valor_fv_erp") - pl.col("total_ventas")).abs() < 1)
        .then(pl.lit(0))
        .otherwise(pl.col("diferencia_valores"))
        .alias("diferencia_valores")
    )

    # Actualización adicional de la columna "diferencia_valores":
    # Si "total_ventas" es 0 y "total_cancelaciones" contiene algún valor,
    # se suma "total_cancelaciones" con "diferencia_valores".
    # Además, si el resultado de la suma tiene una diferencia (valor absoluto) menor a 0.01,
    # se asigna 0; de lo contrario, se mantiene el valor de la suma.
    df_cruce = df_cruce.with_columns(
        pl.when(
            (pl.col("total_ventas") == 0) &
            (pl.col("total_cancelaciones").is_not_null()) &
            (pl.col("total_cancelaciones") != 0)
        ).then(
            pl.when(
                (pl.col("total_cancelaciones") + pl.col("diferencia_valores")).abs() < 0.01
            ).then(
                pl.lit(0)
            ).otherwise(
                pl.col("total_cancelaciones") + pl.col("diferencia_valores")
            )
        ).otherwise(
            pl.col("diferencia_valores")
        ).alias("diferencia_valores")
    )

    # Definir orden específico de columnas
    columnas_ordenadas = [
        "numero_documento",
        "total_ventas",
        "factura_erp",
        "aux_erp",
        "valor_fv_erp",
        "total_valor_factura",
        "diferencia_valores"
    ]

    # Reordenar columnas manteniendo el resto
    df_cruce = df_cruce.select(
        columnas_ordenadas + [
            col for col in df_cruce.columns
            if col not in columnas_ordenadas
        ]
    )


    df_cruce_facturas_cero = df_cruce.filter(pl.col("diferencia_valores") == 0)

    df_cruce_facturas_distintas_cero = df_cruce.filter(pl.col("diferencia_valores") != 0)

    # Extraer los números de documento de cada DataFrame
    docs_cero = set(df_cruce_facturas_cero.select("numero_documento").to_series().to_list())
    docs_distintas = set(df_cruce_facturas_distintas_cero.select("numero_documento").to_series().to_list())

    # Calcular la intersección de números de documento
    docs_interseccion = list(docs_cero.intersection(docs_distintas))

    # Filtrar el DataFrame original para obtener solo los registros cuyo número de documento esté en la intersección
    df_cruce_interseccion = df_cruce.filter(
        pl.col("numero_documento").is_in(docs_interseccion)
    )

    df_cruce_facturas_cero = df_cruce_facturas_cero.filter(
        ~pl.col("numero_documento").is_in(docs_interseccion)
    )
    df_cruce_facturas_distintas_cero = df_cruce_facturas_distintas_cero.filter(
        ~pl.col("numero_documento").is_in(docs_interseccion)
    )



    df_cruce_facturas_distintas_cero = calcular_diferencia_cancelaciones_efecto_0(df_cruce_facturas_distintas_cero)

    df_cruce_cancelaciones_efecto_0 = df_cruce_facturas_distintas_cero.filter(pl.col('diferencia_valores')==0)
    df_cruce_facturas_distintas_cero = df_cruce_facturas_distintas_cero.filter(pl.col('diferencia_valores')!=0)

    df_cruce_cancelaciones_efecto_0_sin_erp = df_cruce_cancelaciones_efecto_0.filter(pl.col('factura_erp').is_null())
    df_cruce_cancelaciones_efecto_0 = df_cruce_cancelaciones_efecto_0.filter(pl.col('factura_erp').is_not_null())


    df_cruce_saldos_por_cobrar = df_cruce_facturas_distintas_cero.filter(pl.col('diferencia_valores') < 0)
    df_cruce_saldos_a_favor = df_cruce_facturas_distintas_cero.filter(pl.col('diferencia_valores') >= 0)



    dataframes_a_exportar = {
        "Cruce principal": df_cruce_facturas_cero,
        "Saldos a favor": df_cruce_saldos_a_favor,
        "Cancelaciones 0": df_cruce_cancelaciones_efecto_0,
        "Cancelaciones sin erp": df_cruce_cancelaciones_efecto_0_sin_erp,
        "Saldos por cobrar": df_cruce_saldos_por_cobrar,
        "Facts revision": df_cruce_interseccion

    }

    exportar_multiples_dataframes_excel(dataframes_a_exportar, "addi_cruce", carpeta_resultados)


def calcular_diferencia_cancelaciones_efecto_0(df: pl.DataFrame) -> pl.DataFrame:
    """
    Esto es para quellas facturas que con su total cancelacion geenran efecto 0 en la diferebncuia
    Calcula la diferencia entre la suma de 'valor' y el promedio del valor absoluto de 'total_cancelaciones'
    para los registros en los que 'estado_transaccion' es "Transaccion", agrupando por 'numero_documento'.

    En este cálculo:
      - La suma se realiza únicamente sobre los registros donde 'estado_transaccion' es "Transaccion".
      - El promedio se calcula tomando el valor absoluto de 'total_cancelaciones' de los mismos registros.

    Args:
        df (pl.DataFrame): DataFrame que contiene las columnas
                           'numero_documento', 'estado_transaccion', 'valor' y 'total_cancelaciones'.

    Returns:
        pl.DataFrame: DataFrame agrupado por 'numero_documento' con la nueva columna 'diferencia_valores'.
    """

    # Agrupar por 'numero_documento' y realizar las agregaciones condicionales
    df_agrupado = df.group_by("numero_documento").agg([
        # Sumar 'total_valor_factura' solo para los registros donde 'estado_transaccion' es "Transacción"
        pl.col("valor_fv_erp")
        .filter(pl.col("estado_transaccion") == "Transacción")
        .cast(pl.Decimal(20, 2))
        .sum()
        .alias("suma_valor"),

        # Calcular el promedio de 'total_ventas' (convertido a Float64) solo para los registros donde 'estado_transaccion' es "Transacción"
        pl.col("total_ventas")
        .cast(pl.Float64)
        .filter(pl.col("estado_transaccion") == "Transacción")
        .mean()
        .alias("suma_ventas"),

        # Calcular el promedio de los valores absolutos de 'total_cancelaciones' (convertido a Float64)
        # solo para los registros donde 'estado_transaccion' contiene "Cancelación"
        pl.col("total_cancelaciones")
        .cast(pl.Float64)
        .abs()
        .filter(pl.col("estado_transaccion").str.contains("Cancelación"))
        .mean()
        .alias("promedio_total_cancelaciones")
    ])


    # Calcular la diferencia: suma_valor + promedio_total_cancelaciones - suma_ventas
    df_agrupado = df_agrupado.with_columns([
        (pl.col("suma_valor") + pl.col("promedio_total_cancelaciones") - pl.col("suma_ventas"))
        .alias("diferencia_valores")
    ]).with_columns([
        pl.col("diferencia_valores").cast(pl.Decimal(20, 2)).alias("diferencia_valores")
    ])


    # Realizar un join entre el DataFrame original y el DataFrame agrupado usando 'numero_documento' como llave
    df_actualizado = df.join(
        df_agrupado.select(["numero_documento", "diferencia_valores"]),
        on="numero_documento",
        how="left",
        suffix="_nuevo"  # Para diferenciar la columna traída del join
    ).with_columns(
        # Actualizar la columna 'diferencia_valores':
        # Si existe un valor calculado (no nulo) en 'diferencia_valores_nuevo', se usa; de lo contrario se conserva el original.
        pl.coalesce([pl.col("diferencia_valores_nuevo"), pl.col("diferencia_valores")])
         .alias("diferencia_valores")
    ).drop("diferencia_valores_nuevo")

    return df_actualizado


def filtrar_grupos_cancelacion(df: pl.DataFrame) -> pl.DataFrame:
    """
    Filtra los registros del DataFrame agrupados por "numero_identificacion" y conserva solo aquellos
    grupos donde todos los registros en "estado_transaccion" contienen la subcadena "Cancelación".

    Procedimiento:
      1. Se filtran los registros donde 'factura_erp' es nula.
      2. Se agrupa por "numero_identificacion" y se convierte la columna "estado_transaccion" en una lista.
         Luego, se evalúa (usando funciones de arreglo) si cada elemento contiene "Cancelación".
         Finalmente, se aplica arr.all() para obtener una bandera booleana que es True solo si
         todos los elementos cumplen la condición.
      3. Se une este DataFrame (con la bandera) con el DataFrame filtrado original.
      4. Se filtran los registros manteniendo únicamente aquellos donde la bandera es True.

    Args:
        df (pl.DataFrame): DataFrame de entrada que debe contener las columnas 'numero_identificacion',
                           'estado_transaccion' y 'factura_erp'.

    Returns:
        pl.DataFrame: DataFrame filtrado que contiene solo los grupos donde todos los registros en
                      "estado_transaccion" incluyen la cadena "Cancelación".
    """
    # Filtrar registros donde 'factura_erp' es nula
    df_filtrado = df.filter(pl.col("factura_erp").is_null())

    # Agrupar por "numero_identificacion" y calcular la bandera booleana
    df_bandera = df_filtrado.group_by("numero_identificacion").agg([
        # Convertir los valores del grupo en una lista, evaluar si cada elemento contiene "Cancelación"
        # y luego verificar que todos los resultados sean True.
        pl.col("estado_transaccion").list()
          .arr.eval(pl.element().str.contains("Cancelación"), parallel=True)
          .arr.all()
          .alias("todos_cancelacion")
    ])

    # Unir el DataFrame de bandera con el filtrado original
    df_unido = df_filtrado.join(df_bandera, on="numero_identificacion", how="left")

    # Filtrar solo los registros cuyo grupo cumpla la condición: todos los elementos contienen "Cancelación"
    df_resultado = df_unido.filter(pl.col("todos_cancelacion"))

    return df_resultado
