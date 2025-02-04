import polars as pl
from .lector_archivos import LectorArchivos

class LectorCSV(LectorArchivos):
    """
    Clase para la lectura de archivos CSV usando Polars.
    """
    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo CSV y devuelve un DataFrame de Polars.
        """
        return pl.read_csv(
            self.configuracion.ruta_archivo,
            separator=self.configuracion.separador,
            encoding=self.configuracion.encoding
        )
