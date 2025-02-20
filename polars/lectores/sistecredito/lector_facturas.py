import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING

from ..lector_archivos import LectorArchivos


if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorSisCredFacturas(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            1:'pagare', # Pagaré
            2:'factura_codigo', # Factura Codigo
            4:'fecha_creacion', #  Fecha Creación
            5:'valor_factura', # Valor Factura
            6:'valor_neto_pagar' #  Valor Neto Pagar

        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion

        self.leer_archivo()



    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Polars.
        """

        df = pl.read_excel(
            self.configuracion.ruta_archivo,
            infer_schema_length=False
        )

        # Convertir todas las columnas a texto para búsqueda
        df_str = df.with_columns(df.select(pl.all().cast(pl.Utf8)).columns)

        # Buscar la primera fila donde aparece "Almacen" o "Almacén" en cualquier columna
        idx_inicio = (
            df_str
            .select(pl.any_horizontal(pl.col(pl.Utf8).str.contains(r"(?i)Almac[eé]n", strict=False)))
            .to_series()
            .arg_max()
        )

        # Buscar la primera fila donde aparece "Total" después de idx_inicio
        idx_fin = (
            df_str.slice(idx_inicio + 1, None)  # Buscar "Total" solo en filas posteriores a "Almacén"
            .select(pl.any_horizontal(pl.col(pl.Utf8).str.contains(r"^Total$", strict=False)))
            .to_series()
            .arg_max()
        )


        # Ajustar idx_fin para que sea relativo al DataFrame completo
        idx_fin = idx_inicio + 1 + idx_fin


        # Obtener los nombres de columnas desde la fila encontrada
        nuevas_columnas = df.row(idx_inicio)

        # Reemplazar valores `None` en los nombres de columnas
        nuevas_columnas = [col if col is not None else f"Columna_{i}" for i, col in enumerate(nuevas_columnas)]

        # Filtrar el DataFrame entre las filas encontradas (desde "Almacén" hasta "Total")
        df_filtrado = df.slice(idx_inicio + 1, idx_fin - idx_inicio - 1)

        # Renombrar las columnas
        df_filtrado = df_filtrado.rename({df_filtrado.columns[i]: nuevas_columnas[i] for i in range(len(nuevas_columnas))})

        # Buscar la columna que corresponde a "Almacen" o "Almacén"
        for col in df_filtrado.columns:
            if col.lower() in ["almacen", "almacén"]:
                df_filtrado = df_filtrado.rename({col: "almacen"})
                break  # Salimos del loop una vez encontrada

        # Rellenar valores nulos en la columna "Almacén" usando forward fill (ffill)
        columna_almacen = "Almacen"  # Ajustar según el nombre final de la columna en el DataFrame
        if columna_almacen in df_filtrado.columns:
            df_filtrado = df_filtrado.with_columns(
                df_filtrado[columna_almacen].fill_null(strategy="forward")
            )

        # Guardarlo en el atributo _dataframe
        self._dataframe = df_filtrado

        self._cambiar_nombres_columnas()

        # Filtrar las filas donde la columna "pagare" NO contenga "Total". Esto es para quitar los registros de subtotal que tiene cada almacen
        self._dataframe = self._dataframe.filter(~self._dataframe["pagare"].cast(pl.Utf8).str.contains(r"(?i)Total", strict=False))
        self._dataframe = self._dataframe.select(['almacen', 'pagare', 'factura_codigo', 'fecha_creacion', 'valor_factura', 'valor_neto_pagar'])


        self._limpieza_datos()

    def _limpieza_datos(self) -> None:

        """
        Realiza la limpieza de datos en el DataFrame.

        Pasos:
        1. Quita los puntos al final del texto en la columna numero_oc_comercial
        2. Elimina espacios en blanco después de quitar los puntos
        """

        pass
