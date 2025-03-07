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
            4:'descripcion', # DESCRIPCIÓN
            7:'monto_bruto_operacion', # MONTO BRUTO DE LA OPERACIÓN
            9:'comision_mercado_pago_incluye_iva', #COMISIÓN DE MERCADO PAGO O MERCADO LIBRE (INCLUYE IVA)
            18:'fecha_aprobacion', # FECHA DE APROBACIÓN
            26:'impuestos_desagregados' # IMPUESTOS DESAGREGADOS

        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion

        self.leer_archivo()


    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Polars.
        """

        self._dataframe = pl.read_excel(
            self.configuracion.ruta_archivo
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

        columnas_a_limpiar = ["numero_identificacion"]

        for columna in columnas_a_limpiar:
            self._dataframe = self._dataframe.with_columns([
                pl.when(pl.col(columna).str.contains("^2[0]+$"))  # Verifica si el valor COMPLETO es un 2 seguido de solo ceros
                .then(pl.col(columna))
                .otherwise(
                    pl.col(columna)
                    .str.replace_all("^2[0]+", "")  # Quita el 2 inicial seguido de cualquier cantidad de ceros
                )
                .alias(f"{columna}_limpio")
            ])




        self._dataframe = self._dataframe.filter(
            ~pl.col("numero_identificacion").str.contains(r"(?i)total")
        )

        print(self._dataframe.filter(
            pl.col("numero_identificacion").is_not_null()
        ))


        columnas_decimales = [
            'monto_bruto_operacion',
            'comision_mercado_pago_incluye_iva'
        ]

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]
        )


        # Sustituir las comillas dobles de inicio y final por comillas simples en la columna impuestos_desagregados
        # Se encadenan dos reemplazos: uno para la comilla inicial y otro para la comilla final
        self._dataframe = self._dataframe.with_columns([
            pl.when(pl.col("impuestos_desagregados").is_not_null())
            .then(
                pl.col("impuestos_desagregados")
                .str.replace(r'^"', "")
                .str.replace(r'"$', "")
            )
            .otherwise(pl.col("impuestos_desagregados"))
            .alias("impuestos_desagregados")
        ])