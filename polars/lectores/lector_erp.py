import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorERP(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            3:'cc_erp', # Cliente
            5:'aux_erp', # Auxiliar
            6:'numero_oc_comercial', # Número O.C. comercial
            7:'documento_causación', # Docto. causación
            8:'factura_erp', # Nro. docto. cruce
            14:'valor_fv_erp', # Total COP
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
        1. Quita los puntos al final del texto en la columna numero_oc_comercial
        2. Elimina espacios en blanco después de quitar los puntos
        """

        columnas_decimales = [
            'valor_fv_erp'
        ]

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]
        )

        columnas_str = [
            'cc_erp',
            'aux_erp',
            'factura_erp'
        ]

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).str.strip_chars())
                for col in columnas_str
            ]
        )

        self._dataframe = self._dataframe.with_columns([
            pl.col("numero_oc_comercial")
            .str.replace_all(r"\.+$", "")  # Quita uno o más puntos al final
            .str.strip_chars()  # Elimina espacios en blanco
            .alias("numero_oc_comercial")
        ])
