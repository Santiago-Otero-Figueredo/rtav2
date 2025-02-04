from typing import List, Dict, Any, Tuple
import pandas as pd
from .modelos import ConfiguracionLector

class LectorArchivos:
    """
    Clase padre para la lectura de archivos.
    """
    def __init__(self, configuracion: ConfiguracionLector):
        self.configuracion = configuracion
        self._dataframe = None
        self._metadata = {}

    def leer_archivo(self) -> pd.DataFrame:
        """
        Método para leer el archivo. Debe ser implementado por las clases hijas.
        """
        raise NotImplementedError("Este método debe ser implementado por las clases hijas.")

    @property
    def dataframe(self) -> pd.DataFrame:
        """
        Obtiene el DataFrame cargado o lo carga si no existe.
        """
        if self._dataframe is None:
            self._dataframe = self.leer_archivo()
        return self._dataframe

    def procesar_datos_en_chunks(self, tamano_chunk: int = 10000) -> pd.DataFrame:
        """
        Procesa el archivo en chunks para manejar grandes volúmenes de datos.

        :param tamano_chunk: Tamaño de cada chunk a procesar
        :return: DataFrame procesado
        """
        chunks = []
        for chunk in pd.read_csv(self.configuracion.ruta_archivo, chunksize=tamano_chunk):
            # Procesar cada chunk según sea necesario
            chunks.append(chunk)
        return pd.concat(chunks, ignore_index=True)

