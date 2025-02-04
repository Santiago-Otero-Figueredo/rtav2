import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorMercadoPago(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            1:'id_operacion_mercado_pago', #ID DE OPERACIÓN EN MERCADO PAGO
            2:'numero_identificacion', # NÚMERO DE IDENTIFICACIÓN
            3:'tipo_registro', # TIPO DE REGISTRO
            7:'monto_bruto_operacion', # MONTO BRUTO DE LA OPERACIÓN
        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion

        self.leer_archivo()


    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Polars.
        """

        self._dataframe = pl.read_excel(
            self.configuracion.ruta_archivo,
            infer_schema_length=False
        )

        self._cambiar_nombres_columnas()
        self._limpieza_datos()


    def _limpieza_datos(self) -> None:
        """
        Realiza la limpieza de datos en el DataFrame.

        Pasos:
        1. Mantiene la columna original numero_identificacion
        2. Crea una nueva columna numero_identificacion_limpio
        3. En la nueva columna remueve el prefijo '20000' si existe

        Returns:
            None: Modifica el DataFrame internamente
        """


        self._dataframe = self._dataframe.with_columns([
            pl.col("numero_identificacion")
            .str.strip_chars()
            .alias("numero_identificacion")
        ])

        self._dataframe = self._dataframe.with_columns([
            # Crear nueva columna con la limpieza del prefijo '20000'
            pl.when(pl.col('numero_identificacion') == "20000")
            .then(pl.col('numero_identificacion'))
            .otherwise(pl.col('numero_identificacion').str.replace_all("^20000", ""))
            .alias("numero_identificacion_limpio")
        ])

        self._dataframe = self._dataframe.filter(
            ~pl.col("numero_identificacion").str.contains(r"(?i)total")
        )

        columnas_decimales = [
            'monto_bruto_operacion'
        ]

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]
        )
