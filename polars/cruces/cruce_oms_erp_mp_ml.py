import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING, List, Dict

from lectores.lector_oms import LectorOMS
from lectores.lector_mercadopago import LectorMercadoPago
from lectores.lector_erp import LectorERP
from lectores.lector_mercadolibre import LectorMercadoLibre

from lectores.modelos import ConfiguracionLector

from .utils import exportar_multiples_dataframes_excel, unir_dataframes_cruce, son_valores_similares

import json


class CruceOmsErpMpMl():
    """
        Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self):

        ruta_carpeta_oms = 'insumos/oms/'
        config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
        lector_oms = LectorOMS(config)
        self.df_oms = lector_oms.dataframe()

        ruta_carpeta_mercadolibre = 'insumos/mercadolibre/'
        config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_mercadolibre)
        lector_mercadolibre = LectorMercadoLibre(config)
        self.df_mercadolibre = lector_mercadolibre.dataframe()

        ruta_archivo_mercaopago = 'insumos/mercadopago/reserve-release-686352448-2025-02-06-072601.xlsx'  # Ajusta esta ruta según tu estructura
        config = ConfiguracionLector(ruta_archivo=ruta_archivo_mercaopago)
        lector_mp = LectorMercadoPago(config)
        self.df_mp = lector_mp.dataframe()

        ruta_archivo_erp = 'insumos/erp/ZOMAC - 2025-02-06T142211.293.xls'  # Ajusta esta ruta según tu estructura
        config = ConfiguracionLector(ruta_archivo=ruta_archivo_erp)
        lector = LectorERP(config)
        self.df_erp = lector.dataframe()

        self.df_cruce = self.cruce_mp_oms_erp()
        self.df_sin_factura = self.obtener_facturas_sin_cruce_mp_oms_erp_mp()

        self.df_cruce_cedulas_facturas, self.df_canceladas, self.df_devoluciones, self.df_facturas_negativas, self.df_facturas_dobles= self.facturas_canceladas_devoluciones()

        dataframes_a_exportar = {
            "Informacion original": self.df_cruce,
            "Cruce principal": self.df_cruce_cedulas_facturas,
            "Facturas devolucion": self.df_devoluciones,
            "Facturas canceladas": self.df_canceladas,
            "Facturas duplicadas": self.df_facturas_dobles,
            "Facturas sin cruzar": self.df_sin_factura,
            "Facturas negativas": self.df_facturas_negativas,

        }

        exportar_multiples_dataframes_excel(dataframes_a_exportar, "self_df_cruce")

    def cruce_mp_oms_erp(self):
         # Realizar el cruce de datos
        df_cruce = self.df_mp.join(
            self.df_oms.select(["orden_externa_limpio", "consecutivo"]),  # Seleccionamos solo las columnas necesarias
            left_on="numero_identificacion_limpio",  # Columna en df_cruce
            right_on="orden_externa_limpio",  # Columna en df_oms
            how="left"  # Mantener todos los registros de df_cruce
        ).rename({"consecutivo": "num_oms"})

        # Realizar el cruce de datos
        df_cruce = df_cruce.join(
            self.df_erp.select(["numero_oc_comercial", "factura_erp", "cc_erp", "valor_fv_erp", "aux_erp"]),  # Seleccionamos solo las columnas necesarias
            left_on="num_oms",  # Columna en df_cruce
            right_on="numero_oc_comercial",  # Columna en df_oms
            how="left"  # Mantener todos los registros de df_cruce
        ).with_columns([
            pl.lit(0).cast(pl.Int64).alias("cruzado_en_mercadolibre")
        ])

        return self.configuraciones_cruce(df_cruce)


    def configuraciones_cruce(self, df_cruce):
          # Mostrar resultados del cruce
        print("\nResultados del cruce:", df_cruce.height)

        df_cruce = df_cruce.with_columns([
            pl.col("impuestos_desagregados").map_elements(lambda x: self.__extraer_valores_impuestos(x, "iva")).alias("iva"),
            pl.col("impuestos_desagregados").map_elements(lambda x: self.__extraer_valores_impuestos(x, "fuente")).alias("fuente"),
            pl.col("impuestos_desagregados").map_elements(lambda x: self.__extraer_valores_impuestos(x, "ica_bogota")).alias("ica")
        ])

        # Realizar el cruce y crear columna de diferencia con ajuste de valores cercanos a 0




        df_cruce = df_cruce.sort([
            "factura_erp",  # Orden por defecto (ascendente)
            "cc_erp",
            pl.col("fecha_aprobacion").sort_by("fecha_aprobacion", descending=False)  # Orden explícito
        ])

        return df_cruce


    def obtener_facturas_sin_cruce_mp_oms_erp_mp(self):
        # Obtener registros sin factura_erp (nulos o vacíos)
        df_sin_factura = self.df_cruce.filter(
            pl.col("factura_erp").is_null() |
            (pl.col("factura_erp").cast(pl.Utf8).fill_null("").str.strip_chars().eq(""))
        )

        df_cruce_mercado_libre = df_sin_factura.join(
            self.df_mercadolibre.select([
                "numero_identificacion",
                "numero_identificacion_limpio",
                pl.col("cedula")  # Renombrar 'cedula' a 'cc_erp'
            ]),
            left_on="numero_identificacion_limpio",  # Columna en df_mercadolibre
            right_on="numero_identificacion_limpio",  # Columna en df_erp
            how="left"  # Mantener todos los registros de df_sin_factura
        ).with_columns([
            pl.coalesce([pl.col("cedula"), pl.col("cc_erp")]).alias("cc_erp"),
            pl.lit(1).cast(pl.Int64).alias("cruzado_en_mercadolibre")
        ]).drop("cedula")# Eliminamos la columna temporal


        df_cruce_mercado_libre = df_cruce_mercado_libre.join(
            self.df_erp.select(["factura_erp", "cc_erp", "valor_fv_erp", "aux_erp"]),  # Seleccionamos solo las columnas necesarias
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
        df_con_factura = df_cruce_mercado_libre.filter(pl.col("factura_erp").is_not_null() & (pl.col("factura_erp") != ""))

        # Filtrar registros donde "nombre" ES nulo o vacío
        df_sin_facturas_asociadas = df_cruce_mercado_libre.filter(pl.col("factura_erp").is_null() | (pl.col("factura_erp") == ""))

        self.df_cruce = unir_dataframes_cruce(self.df_cruce, df_con_factura)

        self.df_cruce = self.df_cruce.with_columns([
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
        self.df_cruce = self.df_cruce.select(
            columnas_ordenadas + [
                col for col in self.df_cruce.columns
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


        return df_sin_facturas_asociadas


    def cruce_facturas_por_cedula(self):
        df_cruce_negativos = self.df_cruce.filter(
            pl.col("diferencia_valores") < 0
        )

        # Obtener lista de facturas existentes
        facturas_existentes = self.df_cruce.select("factura_erp").unique().to_series().to_list()


        # Obtener todas las columnas que no vienen del select de df_erp
        columnas_a_nulificar = [
            col for col in df_cruce_negativos.columns
            if col not in ["cc_erp", "factura_erp", "valor_fv_erp", "aux_erp"]
        ]
        # Realizar el cruce y filtrado para traer todos los registros coincidentes
        df_cruce_cedulas = df_cruce_negativos.join(
            self.df_erp.select([
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

        return unir_dataframes_cruce(self.df_cruce, df_cruce_cedulas)


    def facturas_canceladas_devoluciones(self):

        df_cruce_cedulas_facturas = self.cruce_facturas_por_cedula()

        df_canceladas = self.__obtener_facturas_canceladas(df_cruce_cedulas_facturas)
        listados_facturas_canceladas = df_canceladas.select("factura_erp").unique().to_series().to_list()
        df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
            ~pl.col("factura_erp").is_in(listados_facturas_canceladas)
        )

        df_devoluciones = self.__obtener_facturas_devueltas(df_cruce_cedulas_facturas)
        listados_facturas_devueltas = df_devoluciones.select("factura_erp").unique().to_series().to_list()
        df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
            ~pl.col("factura_erp").is_in(listados_facturas_devueltas)
        )


        # Filtrar registros con diferencia negativa
        listados_facturas_negativas = df_cruce_cedulas_facturas.filter(
            pl.col("diferencia_valores") < 0
        ).select("factura_erp").unique().to_series().to_list()

        df_facturas_negativas = self.df_cruce.filter(
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

        df_facturas_dobles = self.__procesar_facturas_agrupadas(df_cruce_cedulas_facturas)
        listados_facturas_dobles = df_facturas_dobles.select("factura_erp").unique().to_series().to_list()
        df_cruce_cedulas_facturas = df_cruce_cedulas_facturas.filter(
            ~pl.col("factura_erp").is_in(listados_facturas_dobles)
        )

        # Unir df_cruce_cedulas_facturas con df_facturas_dobles
        df_cruce_cedulas_facturas = unir_dataframes_cruce(df_cruce_cedulas_facturas, df_facturas_dobles)

        return df_cruce_cedulas_facturas, df_canceladas, df_devoluciones, df_facturas_negativas, df_facturas_dobles


    def obtener_df_oms(self):
        return self.df_oms

    def obtener_df_mercadolibre(self):
        return self.df_mercadolibre

    def obtener_df_mercadopago(self):
        return self.df_mp

    def obtener_df_erp(self):
        return self.df_erp


    def __extraer_valores_impuestos(self, impuestos: str, tipo_impuesto: str) -> float:
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


    def __obtener_facturas_canceladas(self, df: pl.DataFrame) -> pl.DataFrame:
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

    def __obtener_facturas_devueltas(self, df: pl.DataFrame) -> pl.DataFrame:
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


    def __procesar_facturas_agrupadas(self, df: pl.DataFrame) -> pl.DataFrame:
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

