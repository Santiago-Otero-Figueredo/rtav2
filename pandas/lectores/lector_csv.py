import pandas as pd
from .lector_archivos import LectorArchivos

class LectorCSV(LectorArchivos):
    """
    Clase para la lectura de archivos CSV usando Pydantic.
    """
    def leer_archivo(self) -> pd.DataFrame:
        """
        Lee un archivo CSV y devuelve un ResultadoLectura.
        """
        self.resultado.datos = pd.read_csv(
            self.configuracion.ruta_archivo,
            encoding=self.configuracion.encoding,
            sep=self.configuracion.separador
        )
        self.resultado.metadata = {
            "tipo_archivo": "csv",
            "encoding": self.configuracion.encoding,
            "separador": self.configuracion.separador
        }
        return self.resultado
