from typing import List, Dict, Any, Tuple
import polars as pl
from .modelos import ConfiguracionLector

class LectorArchivos:
    """
    Clase padre para la lectura de archivos usando Polars.
    """
    def __init__(self, configuracion: ConfiguracionLector):
        self.configuracion = configuracion
        self._dataframe = None
        self._metadata = {}

    def leer_archivo(self) -> pl.DataFrame:
        """
        Método para leer el archivo. Debe ser implementado por las clases hijas.
        """
        raise NotImplementedError("Este método debe ser implementado por las clases hijas.")

    @property
    def dataframe(self) -> pl.DataFrame:
        """
        Obtiene el DataFrame cargado o lo carga si no existe.
        """
        if self._dataframe is None:
            self._dataframe = self.leer_archivo()
        return self._dataframe

    def procesar_datos_en_chunks(self, tamano_chunk: int = 10000) -> pl.DataFrame:
        """
        Procesa el archivo en chunks para manejar grandes volúmenes de datos.

        Args:
            tamano_chunk (int): Tamaño de cada chunk a procesar

        Returns:
            pl.DataFrame: DataFrame procesado
        """
        return pl.scan_csv(self.configuracion.ruta_archivo).collect()

