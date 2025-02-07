import pandas as pd
from pathlib import Path
from decimal import Decimal
from typing import TYPE_CHECKING, List, Dict

import re
from .lector_archivos import LectorArchivos

if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorOMS(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Pandas.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {}
        if configuracion.cargue_inicial is True:
            mapeo_indices_nombres_columnas = {
                2: 'orden_externa', # orden externa
                29:'cliente_nombre', # cliente nombre
                44:'forma_pago_1', # forma pago 1
                45:'forma_pago_1_referencia', # forma pago 1 referencia
                46:'forma_pago_1_valor', # forma pago 1 valor
                47:'forma_pago_2', # forma pago 2
                48:'forma_pago_2_referencia', # forma pago 2 referencia
                49:'forma_pago_2_valor', # forma pago 2 valor
                50:'forma_pago_3', # forma pago 3
                51:'forma_pago_3_referencia', # forma pago 3 referencia
                52:'forma_pago_3_valor', # forma pago 3 valor
                53:'monto' # pvp
            }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion
        self.archivos = []

        self.__obtener_archivos_oms()

        self.leer_archivo()

    def leer_archivo(self) -> pd.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Pandas.
        """

        self._dataframe = pd.DataFrame()  # Iniciar un DataFrame vacío

        for archivo in self.archivos:
            if not (archivo.endswith('.xlsx') or archivo.endswith('.xls')):
                raise ValueError("Formato no soportado. Solo .xlsx y .xls")

            engine = 'openpyxl' if archivo.endswith('.xlsx') else 'xlrd'
            df = pd.read_excel(
                archivo,
                engine=engine
            )

            # Asegurar que 'orden externa' sea string
            #df['orden externa'] = df['orden externa'].astype(str)

            self._dataframe = pd.concat([self._dataframe, df], ignore_index=True)

        if self.configuracion.cargue_inicial is True:
            self._dataframe['orden_externa_str'] = self._dataframe['orden externa']
            self._dataframe['orden externa'] = self._dataframe['orden externa'].astype(str)

            self._cambiar_nombres_columnas()
            self._limpieza_datos()

    def _limpieza_datos(self) -> None:
        """
        Realiza la limpieza de datos en el DataFrame.
        """

        columnas_evaluar = [
            'consecutivo',
            'orden_externa', # orden externa
            'orden_externa_str',
            'cliente_nombre', # cliente nombre
            'forma_pago_1', # forma pago 1
            'forma_pago_1_referencia', # forma pago 1 referencia
            'forma_pago_1_valor', # forma pago 1 valor
            'forma_pago_2', # forma pago 2
            'forma_pago_2_referencia', # forma pago 2 referencia
            'forma_pago_2_valor', # forma pago 2 valor
            'forma_pago_3', # forma pago 3
            'forma_pago_3_referencia', # forma pago 3 referencia
            'forma_pago_3_valor', # forma pago 3 valor
            'monto' # pvp
        ]

        self._dataframe = self._dataframe[columnas_evaluar]



        columnas_decimales = [
            #'costo excl imp',
            #'precio vta excl imp',
            #'imp vta',
            #'imp costo',
            #'sub total costo exl imp',
            #'sub total vta exl imp',
            'forma_pago_1_valor',
            'forma_pago_2_valor',
            'forma_pago_3_valor',
            'monto'
        ]

        columnas_int = [
            'cantidad'
        ]

        columnas_str = [
            'consecutivo'
        ]

        # Convertir columnas decimales
        for col in columnas_decimales:
            if col in self._dataframe.columns:
                self._dataframe[col] = (
                    self._dataframe[col]
                    .round(2)
                    .apply(Decimal)
                )

        # Convertir columnas enteras
        for col in columnas_int:
            if col in self._dataframe.columns:
                self._dataframe[col] = self._dataframe[col].astype('Int64')

        # Convertir columnas str
        for col in columnas_str:
            if col in self._dataframe.columns:
                self._dataframe[col] = self._dataframe[col].astype(str)


        print('self._dataframe>>>>>>>>>>> ', self._dataframe.dtypes)

        # Crear orden_externa_duplicada
        patron = r'^\d+\s\d+$'
        self._dataframe['orden_externa_duplicada'] = self._dataframe['orden_externa_str'].apply(
            lambda x: x.replace(x.split()[-1], '').strip() if pd.notna(x) and bool(re.match(patron, str(x))) else None
        )

        # Limpiar prefijo '20000'
        columnas_a_limpiar = ["orden_externa_str", "orden_externa_duplicada"]
        for columna in columnas_a_limpiar:
            self._dataframe[f'{columna}_limpio'] = self._dataframe[columna].apply(
                lambda x: x if x == "20000" else str(x).replace("20000", "", 1) if pd.notna(x) else None
            )


        columnas_ordenadas = [
            'consecutivo',
            'orden_externa', # orden externa
            'orden_externa_str',
            'orden_externa_str_limpio',
            'cliente_nombre', # cliente nombre
            'forma_pago_1', # forma pago 1
            'forma_pago_1_referencia', # forma pago 1 referencia
            'forma_pago_1_valor', # forma pago 1 valor
            'forma_pago_2', # forma pago 2
            'forma_pago_2_referencia', # forma pago 2 referencia
            'forma_pago_2_valor', # forma pago 2 valor
            'forma_pago_3', # forma pago 3
            'forma_pago_3_referencia', # forma pago 3 referencia
            'forma_pago_3_valor', # forma pago 3 valor
            'monto' # pvp
        ]

        # Reordenar columnas manteniendo el resto
        columnas_restantes = [col for col in self._dataframe.columns if col not in columnas_ordenadas]
        self._dataframe = self._dataframe[columnas_ordenadas + columnas_restantes]

    def __obtener_archivos_oms(self) -> List[str]:
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
            self.archivos.append(self.configuracion.ruta_archivo)

        return self.archivos