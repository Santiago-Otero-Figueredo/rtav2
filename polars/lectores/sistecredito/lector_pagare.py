import polars as pl
from pathlib import Path

from typing import TYPE_CHECKING

from ..lector_archivos import LectorArchivos

import openpyxl

if TYPE_CHECKING:
    from lectores.modelos import ConfiguracionLector


class LectorSisCredPagare(LectorArchivos):
    """
    Clase para la lectura de archivos Excel específicos con nombre OMS usando Polars.
    """
    def __init__(self, configuracion: 'ConfiguracionLector'):

        mapeo_indices_nombres_columnas = {
            0:'asesor', # Asesor
            1:'fecha_consulta', # Fecha de consulta
            2:'consecutivo_pagare', # Consecutivo pagare
            3:'documento_identidad', #  Documento de identidad
            4:'factura', #  Factura
            5:'valor_credito', # Valor Credito
            6:'metodo_pago' # Asesor
        }

        super().__init__(configuracion=configuracion, mapeo_indices_nombres_columnas=mapeo_indices_nombres_columnas)
        self.configuracion = configuracion

        self.leer_archivo()



    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Polars.
        """

        # Obtener nombres de todas las hojas
        wb = openpyxl.load_workbook(self.configuracion.ruta_archivo, read_only=True)
        hojas = wb.sheetnames  # Lista de nombres de las hojas
        wb.close()

        # Lista para almacenar los DataFrames de cada hoja
        dfs = []

        # Leer cada hoja y agregar columna "almacen"
        for hoja in hojas:
            df = pl.read_excel(self.configuracion.ruta_archivo, sheet_name=hoja, infer_schema_length=None)


            idx_asesor = 0
            # ✅ Verificar si "Asesor" NO está en las columnas
            if not "Asesor" in df.columns:
                # Convertir todas las columnas a texto para búsqueda
                df_str = df.with_columns(df.select(pl.all().cast(pl.Utf8)).columns)

                # Buscar la primera fila donde aparece "Asesor" en cualquier columna
                idx_asesor = (
                    df_str.select(
                        pl.struct(df.columns).map_elements(
                            lambda row: any("Asesor" in str(v) for v in row.values()),
                            return_dtype=pl.Boolean  # ✅ Evita el warning
                        )
                    )
                    .to_series()
                    .arg_max()
                )

                # Si no se encuentra "Asesor", saltar la hoja
                if idx_asesor == 0 and "Asesor" not in df.row(0):
                    print(f"⚠️ No se encontró 'Asesor' en la hoja {hoja}, se omite.")
                    continue

                # Obtener la fila con los nombres de las columnas
                nuevas_columnas = df.row(idx_asesor)

                # Recortar el DataFrame desde la fila de "Asesor"
                df = df.slice(idx_asesor + 1)

                # Asignar los nuevos nombres de columna
                df = df.rename({df.columns[i]: nuevas_columnas[i] for i in range(len(nuevas_columnas))})

            print(idx_asesor, hoja)
            # Agregar la columna "almacen" con el nombre de la hoja
            df = df.with_columns(pl.lit(hoja).alias("almacen"))

            # Agregar el DataFrame procesado a la lista
            dfs.append(df)

        # Unir todos los DataFrames en uno solo
        df_consolidado = pl.concat(dfs, how="diagonal_relaxed")
        df_consolidado = df_consolidado.filter(~df_consolidado["Asesor"].cast(pl.Utf8).str.contains(r"(?i)Asesor$", strict=False))

        # Guardar en el atributo _dataframe
        self._dataframe = df_consolidado

        self._cambiar_nombres_columnas()

        self._dataframe = self._dataframe.select([
            'almacen',
            'asesor',
            'fecha_consulta',
            'consecutivo_pagare',
            'documento_identidad',
            'factura',
            'valor_credito',
            'metodo_pago',
        ])
        self._limpieza_datos()

    def _limpieza_datos(self) -> None:

        """
        Realiza la limpieza de datos en el DataFrame.

        Pasos:
        1. Quita los puntos al final del texto en la columna numero_oc_comercial
        2. Elimina espacios en blanco después de quitar los puntos
        """

        self._dataframe = self._dataframe.with_columns(
            pl.col("documento_identidad")
            .str.replace(r"^\D+", "", literal=False)  # 🔹 Quita solo los caracteres no numéricos al inicio
            .alias("documento_identidad")
        )