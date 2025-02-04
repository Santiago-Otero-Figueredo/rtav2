import polars as pl
from .lector_archivos import LectorArchivos

class LectorExcel(LectorArchivos):
    """
    Clase para la lectura de archivos Excel usando Polars.
    """
    def leer_archivo(self) -> pl.DataFrame:
        """
        Lee un archivo Excel y devuelve un DataFrame de Polars.
        """
        if not (self.configuracion.ruta_archivo.endswith('.xlsx') or
                self.configuracion.ruta_archivo.endswith('.xls')):
            raise ValueError("Formato no soportado. Solo .xlsx y .xls")


        df = pl.read_excel(
            self.configuracion.ruta_archivo,
            infer_schema_length=False
        )


        columnas_decimales = [
            'costo excl imp',
            'precio vta excl imp',
            'imp vta',
            'imp costo',
            'sub total costo exl imp',
            'sub total vta exl imp',
            'forma pago 1 valor',
            'forma pago 2 valor',
            'forma pago 3 valor'
        ]

        columnas_int = [
            'cantidad'
        ]

        df = df.with_columns(
            [
                (pl.col(col).cast(pl.Float64).round(2))  # Redondear a 2 decimales
                .cast(pl.Decimal(20, 2))  # Convertir a Decimal(10, 2)
                for col in columnas_decimales
            ]+ [
                # Conversión a int (solo si es necesario) en otras columnas
                (pl.col(col).cast(pl.Int64))  # Convertir a Int64
                for col in columnas_int  # Otras columnas que deseas convertir a int
            ]
        )

        # Crear una nueva columna que contenga el valor de la columna original solo si coincide con el patrón
        df = df.with_columns(
            pl.when(pl.col("orden externa").str.contains(r"^\d+\s\d+$"))  # Verifica si la cadena coincide con el patrón
            .then(pl.col("orden externa").str.replace(r"\s\d+$", ""))  # Elimina el número final y el espacio
            .otherwise(pl.lit(None))  # Si no coincide, el valor será None
            .alias("orden externa dupliacada")
        )


        return df
