import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING, List, Dict

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorADDI(LectorArchivos):
    """
        Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            0:'estado_transaccion', #Estado de la transacción
            7:'numero_documento', # Número de documento
            11:'total_ventas', # Total Ventas (1)
            12:'total_cancelaciones' # Total Cancelaciones

        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion
        self.archivos = []

        self.__obtener_archivos_addi()

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
                infer_schema_length=False,
                sheet_id=2
            )

             # Extraer la fila 6 (índice 5) como nombres de columnas
            column_names = [self.__limpiar_str(name) if name is not None else f"col_{i}" for i, name in enumerate(df.row(4))]

            # Filtrar desde la fila 7 en adelante (índice 6)
            df = df.slice(5)

            # Asignar los nuevos nombres de columna
            df = df.rename(dict(zip(df.columns, column_names)))

            self._dataframe = self._dataframe.vstack(df)

        self._cambiar_nombres_columnas()
        self._limpieza_datos()


    def _limpieza_datos(self) -> None:

        columnas_decimales = [
            'total_ventas',
            'total_cancelaciones'
        ]

        # Primero limpiamos los espacios en blanco de la columna total_ventas
        self._dataframe = self._dataframe.with_columns([
            pl.col("total_ventas").str.replace_all(" ", "")
        ])

        self._dataframe = self._dataframe.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]
        )





    def __obtener_archivos_addi(self) -> List[str]:
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

    def __limpiar_str(self, texto):
        return texto.strip().replace("\t", " ").replace("\n", " ").replace("\r", " ")
