from typing import List, Dict, Any, Tuple
import pandas as pd
from .modelos import ConfiguracionLector


class LectorArchivos:
    """
    Clase padre para la lectura de archivos usando Pandas.
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

    def dataframe(self) -> pd.DataFrame:
        """
        Obtiene el DataFrame cargado
        """
        return self._dataframe

    def _cambiar_nombres_columnas(self) -> pd.DataFrame:
        """
        Cambia los nombres de las columnas según su índice.

        Returns:
            pd.DataFrame: DataFrame con los nombres de columnas actualizados
        """
        if self._dataframe is None:
            raise ValueError("No se ha cargado el DataFrame")

        nombres_columnas = list(self._dataframe.columns)
        for indice, nuevo_nombre in self._mapeo_indices_nombres_columnas.items():
            if 0 <= indice < len(nombres_columnas):
                nombres_columnas[indice] = nuevo_nombre

        self._dataframe.columns = nombres_columnas
        return self._dataframe

    def procesar_datos_en_chunks(self, tamano_chunk: int = 10000) -> pd.DataFrame:
        """
        Procesa el archivo en chunks para manejar grandes volúmenes de datos.

        Args:
            tamano_chunk (int): Tamaño de cada chunk a procesar

        Returns:
            pd.DataFrame: DataFrame procesado
        """
        chunks = []
        # Utilizamos el método read_csv con chunks de pandas
        for chunk in pd.read_csv(self.configuracion.ruta_archivo, chunksize=tamano_chunk):
            chunks.append(chunk)
        return pd.concat(chunks, ignore_index=True)

