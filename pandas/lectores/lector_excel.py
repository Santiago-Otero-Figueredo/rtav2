import pandas as pd
from .lector_archivos import LectorArchivos

class LectorExcel(LectorArchivos):
    """
    Clase para la lectura de archivos Excel.
    """
    def leer_archivo(self) -> pd.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame.
        """
        if self.configuracion.ruta_archivo.endswith('.xlsx'):
            engine = 'openpyxl'
        elif self.configuracion.ruta_archivo.endswith('.xls'):
            engine = 'xlrd'
        else:
            raise ValueError("Formato no soportado. Solo .xlsx y .xls")

        return pd.read_excel(
            self.configuracion.ruta_archivo,
            engine=engine
        )
