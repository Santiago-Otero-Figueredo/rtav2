import pandas as pd
from pathlib import Path
from decimal import Decimal
from typing import TYPE_CHECKING

from .lector_archivos import LectorArchivos

if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorERP(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Pandas.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {}
        if configuracion.cargue_inicial is True:
            mapeo_indices_nombres_columnas = {
                3:'cedula_cliente', # Cliente
                5:'auxiliar', # Auxiliar
                6:'numero_oc_comercial', # Número O.C. comercial
                7:'documento_causación', # Docto. causación
                8:'numero_documento_cruce', # Nro. docto. cruce
                14:'total_cop', # Total COP
            }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion
        self.leer_archivo()

    def leer_archivo(self) -> pd.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Pandas.

        El método detecta automáticamente si el archivo es .xlsx o .xls y usa
        el engine apropiado para cada formato:
        - .xlsx: usa 'openpyxl'
        - .xls: usa 'xlrd'

        Raises:
            ValueError: Si el archivo no es un formato Excel válido
        """
        ruta_archivo = Path(self.configuracion.ruta_archivo)
        extension = ruta_archivo.suffix.lower()

        if extension not in ['.xlsx', '.xls']:
            raise ValueError(f"Formato de archivo no soportado: {extension}. Use .xlsx o .xls")

        try:
            engine = 'openpyxl' if extension == '.xlsx' else 'xlrd'
            self._dataframe = pd.read_excel(
                self.configuracion.ruta_archivo,
                engine=engine
            )
        except Exception as e:
            raise ValueError(f"Error al leer el archivo Excel: {str(e)}")

        if self.configuracion.cargue_inicial is True:
            self._cambiar_nombres_columnas()
            self._limpieza_datos()

    def _limpieza_datos(self) -> None:
        """
        Realiza la limpieza de datos en el DataFrame.

        Pasos:
        1. Quita los puntos al final del texto en la columna numero_oc_comercial
        2. Elimina espacios en blanco después de quitar los puntos
        3. Convierte las columnas numéricas a decimal
        """
        # Convertir columnas decimales
        columnas_decimales = ['total_cop']
        for col in columnas_decimales:
            self._dataframe[col] = self._dataframe[col].round(2).apply(Decimal)

        # Limpiar numero_oc_comercial
        self._dataframe['numero_oc_comercial'] = (
            self._dataframe['numero_oc_comercial']
            .str.replace(r'\.+$', '', regex=True)  # Quita uno o más puntos al final
            .str.strip()  # Elimina espacios en blanco
        )

