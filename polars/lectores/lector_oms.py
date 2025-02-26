import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING, List, Dict

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorOMS(LectorArchivos):
    """
        Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            0:'consecutivo', # consecutivo
            2: 'orden_externa', # orden externa
            29:'cliente_nombre', # cliente nombre
            31:'cedula_cliente', # documento
            43:'estado', # estado
            44:'forma_pago_1', # forma pago 1
            45:'forma_pago_1_referencia', # forma pago 1 referencia
            46:'forma_pago_1_valor', # forma pago 1 valor
            47:'forma_pago_2', # forma pago 2
            48:'forma_pago_2_referencia', # forma pago 2 referencia
            49:'forma_pago_2_valor', # forma pago 2 valor
            50:'forma_pago_3', # forma pago 3
            51:'forma_pago_3_referencia', # forma pago 3 referencia
            52:'forma_pago_3_valor', # forma pago 3 valor
            53:'monto' # pvp
        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion
        self.archivos = []

        self.__obtener_archivos_oms()

        self.leer_archivo()



    def leer_archivo(self) -> pl.DataFrame:
        """
            Lee un archivo Excel y devuelve un DataFrame de Polars.
        """
        self._dataframe = pl.DataFrame()  # Iniciar un DataFrame vacío

        for archivo in self.archivos:

            if not (archivo.endswith('.xlsx') or archivo.endswith('.xls')):
                raise ValueError("Formato no soportado. Solo .xlsx y .xls")

            df = pl.read_excel(
                archivo,
                infer_schema_length=False
            )

            self._dataframe = self._dataframe.vstack(df)

        self._cambiar_nombres_columnas()
        self._limpieza_datos()


    def _limpieza_datos(self) -> None:
        columnas_decimales = [
            'costo excl imp',
            'precio vta excl imp',
            'imp vta',
            'imp costo',
            'sub total costo exl imp',
            'sub total vta exl imp',
            'forma_pago_1_valor',
            'forma_pago_2_valor',
            'forma_pago_3_valor',
            'monto'
        ]

        columnas_int = [
            'cantidad'
        ]

        columans_str = [
            'estado'
        ]

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]+ [
                # Conversión a int (solo si es necesario) en otras columnas
                (pl.col(col).cast(pl.Int64))  # Convertir a Int64
                for col in columnas_int  # Otras columnas que deseas convertir a int
            ]+ [
                # Conversión a str (solo si es necesario) en otras columnas
                (pl.col(col).str.strip_chars())  # Convertir a str
                for col in columans_str  # Otras columnas que deseas convertir a str
            ]
        )

        # Filtrar registros anulados (no distingue mayúsculas/minúsculas)
        self._dataframe = self._dataframe.filter(
            ~pl.col("estado").str.to_lowercase().eq("anulado")
        )

        self._dataframe = self._dataframe.with_columns([
            pl.col("orden_externa")
            .str.strip_chars()
            .alias("orden_externa")
        ])

        # Crear una nueva columna que contenga el valor de la columna original solo si coincide con el patrón
        self._dataframe = self._dataframe.with_columns(
            pl.when(pl.col("orden_externa").str.contains(r"^\d+\s\d+$"))  # Verifica si la cadena coincide con el patrón
            .then(pl.col("orden_externa").str.replace(r"\s\d+$", ""))  # Elimina el número final y el espacio
            .otherwise(pl.lit(None))  # Si no coincide, el valor será None
            .alias("orden_externa_duplicada")
        )

        columnas_a_limpiar = ["orden_externa", "orden_externa_duplicada"]

        for columna in columnas_a_limpiar:
            self._dataframe = self._dataframe.with_columns([
                pl.when(pl.col(columna).str.contains("^2[0]+$"))  # Verifica si el valor COMPLETO es un 2 seguido de solo ceros usando $ al final
                .then(pl.col(columna))
                .otherwise(
                    pl.col(columna)
                    .str.replace_all("^2[0]+", "")  # Quita el 2 inicial seguido de cualquier cantidad de ceros
                )
                .alias(f"{columna}_limpio")
            ])

        # Evaluar si alguna de las formas de pago es "MERCADOPAGO" y traer la referencia correspondiente
        self._dataframe = self._dataframe.with_columns(
            pl.when(pl.col("forma_pago_1") == "MERCADOPAGO")
            .then(pl.col("forma_pago_1_referencia"))
            .when(pl.col("forma_pago_2") == "MERCADOPAGO")
            .then(pl.col("forma_pago_2_referencia"))
            .when(pl.col("forma_pago_3") == "MERCADOPAGO")
            .then(pl.col("forma_pago_3_referencia"))
            .otherwise(None)  # Si no hay "MERCADOPAGO", deja el valor como nulo
            .alias("referencia_mercadopago")
        )


    def __obtener_archivos_oms(self) -> List[str]:
        """
        Obtiene todos los archivos Excel y CSV de la carpeta especificada.

        Returns:
            List[str]: Listas de rutas de archivos por tipo
        """

        if self.configuracion.ruta_carpeta:

            for archivo in Path(self.configuracion.ruta_carpeta).glob('*'):
                if archivo.suffix.lower() in ['.xlsx', '.xls']:
                    self.archivos.append(str(archivo))
        else:
            self.archivos = [self.configuracion.ruta_archivo]

        return self.archivos