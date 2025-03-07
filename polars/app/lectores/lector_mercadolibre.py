import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING, List, Dict

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorMercadoLibre(LectorArchivos):
    """
        Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            0: 'numero_identificacion', # # de venta
            24: 'cedula' # Tipo y número de documento
        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion
        self.archivos = []

        self.__obtener_archivos_mercadolibe()
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

              # Extraer la fila 6 (índice 5) como nombres de columnas
            column_names = [self.__limpiar_str(name) if name is not None else f"col_{i}" for i, name in enumerate(df.row(3))]

            # Hacer que los nombres sean únicos
            column_names = self.__hacer_nombres_unicos(column_names)

            # Filtrar desde la fila 7 en adelante (índice 6)
            df = df.slice(4)

            # Asignar los nuevos nombres de columna
            df = df.rename(dict(zip(df.columns, column_names)))

            self._dataframe = self._dataframe.vstack(df)



        self._cambiar_nombres_columnas()
        self._limpieza_datos()


    def _limpieza_datos(self) -> None:


        columnas_a_limpiar = ["numero_identificacion"]

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

        self._dataframe = self._dataframe.with_columns(
            pl.col("cedula").str.extract(r"(\d+)$").alias("cedula")
        )



    def __obtener_archivos_mercadolibe(self) -> List[str]:
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

    def __hacer_nombres_unicos(self, column_names):
        seen = {}
        unique_names = []

        for name in column_names:
            if name in seen:
                seen[name] += 1
                unique_names.append(f"{name}_{seen[name]}")
            else:
                seen[name] = 0
                unique_names.append(name)

        return unique_names