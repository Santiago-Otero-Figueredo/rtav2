import pandas as pd
from pathlib import Path
from decimal import Decimal
from typing import TYPE_CHECKING

from .lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorMercadoPago(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Pandas.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            1:'id_operacion_mercado_pago', #ID DE OPERACIÓN EN MERCADO PAGO
            2:'numero_identificacion', # NÚMERO DE IDENTIFICACIÓN
            3:'tipo_registro', # TIPO DE REGISTRO
            7:'monto_bruto_operacion', # MONTO BRUTO DE LA OPERACIÓN
        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion

        self.leer_archivo()

    def leer_archivo(self) -> pd.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Pandas.
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

        self._cambiar_nombres_columnas()
        self._limpieza_datos()

    def _limpieza_datos(self) -> None:
        """
        Realiza la limpieza de datos en el DataFrame.

        Pasos:
        1. Mantiene la columna original numero_identificacion
        2. Crea una nueva columna numero_identificacion_limpio
        3. En la nueva columna remueve el prefijo '20000' si existe

        Returns:
            None: Modifica el DataFrame internamente
        """
        # Limpiar espacios en numero_identificacion
        self._dataframe['numero_identificacion'] = self._dataframe['numero_identificacion'].str.strip()

        # Crear nueva columna con la limpieza del prefijo '20000'
        self._dataframe['numero_identificacion_limpio'] = (
            self._dataframe['numero_identificacion'].apply(
                lambda x: x if x == "20000" else str(x).replace("^20000", "", regex=True)
            )
        )

        # Convertir columnas decimales
        columnas_decimales = [
            'monto_bruto_operacion'
        ]

        for col in columnas_decimales:
            self._dataframe[col] = (
                self._dataframe[col]
                .round(2)
                .apply(Decimal)
            )
