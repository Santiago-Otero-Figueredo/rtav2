from typing import List, Dict, Any, Tuple
import polars as pl
from .modelos import ConfiguracionLector


class LectorArchivos:
    """
    Clase padre para la lectura de archivos usando Polars.
    """
    def __init__(self, configuracion: ConfiguracionLector, mapeo_indices_nombres_columnas: dict={}):
        self.configuracion = configuracion
        self._dataframe = None
        self._metadata = {}
        self._mapeo_indices_nombres_columnas = mapeo_indices_nombres_columnas



    def leer_archivo(self) -> None:
        """
        Método para leer el archivo. Debe ser implementado por las clases hijas.
        """
        raise NotImplementedError("Este método debe ser implementado por las clases hijas.")

    def _limpieza_datos(self) -> None:
        """
        Realiza la limpieza de los datos del DataFrame cargado.
        """
        raise NotImplementedError("Este método debe ser implementado por las clases hijas.")

    def dataframe(self) -> pl.DataFrame:
        """
        Obtiene el DataFrame cargado
        """
        return self._dataframe

    def _cambiar_nombres_columnas(self) -> pl.DataFrame:
        """
        Cambia los nombres de las columnas según su índice.

        Returns:
            pl.DataFrame: DataFrame con los nombres de columnas actualizados
        """
        if self._dataframe is None:
            raise ValueError("No se ha cargado el DataFrame")

        nombres_columnas = list(self._dataframe.columns)
        for indice, nuevo_nombre in self._mapeo_indices_nombres_columnas.items():
            if 0 <= indice < len(nombres_columnas):
                nombres_columnas[indice] = nuevo_nombre

        self._dataframe = self._dataframe.select(pl.col("*")).rename(dict(zip(self._dataframe.columns, nombres_columnas)))

    def procesar_datos_en_chunks(self, tamano_chunk: int = 10000) -> pl.DataFrame:
        """
        Procesa el archivo en chunks para manejar grandes volúmenes de datos.

        Args:
            tamano_chunk (int): Tamaño de cada chunk a procesar

        Returns:
            pl.DataFrame: DataFrame procesado
        """
        return pl.scan_csv(self.configuracion.ruta_archivo).collect()

