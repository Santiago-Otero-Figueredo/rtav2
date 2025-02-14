
from datetime import datetime
import polars as pl
import pandas as pd
import os


def exportar_multiples_dataframes_excel(
    dataframes: dict[str, pl.DataFrame],
    nombre_base: str
) -> None:
    """
    Exporta múltiples DataFrames a diferentes hojas de un mismo archivo Excel.

    Args:
        dataframes (dict[str, pl.DataFrame]): Diccionario con nombre_hoja:DataFrame
        nombre_base (str): Nombre base para el archivo Excel

    Returns:
        None: Guarda el archivo en disco
    """
    try:
        # Filtrar DataFrames que no sean None
        dataframes = {k: v for k, v in dataframes.items() if v is not None}

        # Si no hay DataFrames válidos, agregar una hoja vacía
        if not dataframes:
            dataframes["Hoja_Vacia"] = pl.DataFrame()

        # Crear carpeta si no existe
        carpeta_resultados = "resultados"
        os.makedirs(carpeta_resultados, exist_ok=True)

        # Generar nombre con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"{carpeta_resultados}/{nombre_base}_{timestamp}.xlsx"

        # Crear el escritor de Excel
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            # Iterar sobre cada DataFrame y escribirlo en su hoja
            for nombre_hoja, df in dataframes.items():
                # Convertir DataFrame de Polars a Pandas
                df_pandas = df.to_pandas()
                # Escribir en la hoja específica
                df_pandas.to_excel(writer, sheet_name=nombre_hoja, index=False)

        print(f"Archivo exportado: {nombre_archivo}")
        print(f"Hojas creadas: {list(dataframes.keys())}")

    except Exception as e:
        print(f"Error al exportar DataFrames: {str(e)}")
        raise



def unir_dataframes_cruce(df_principal: pl.DataFrame, df_adicional: pl.DataFrame) -> pl.DataFrame:
    """
    Une dos DataFrames verticalmente asegurando compatibilidad de columnas.

    Args:
        df_principal (pl.DataFrame): DataFrame principal con los cruces originales
        df_adicional (pl.DataFrame): DataFrame con cruces adicionales por cédula

    Returns:
        pl.DataFrame: DataFrame unificado con todos los cruces
    """
    try:
        # Verificar que ambos DataFrames no estén vacíos
        if df_principal.height == 0 or df_adicional.height == 0:
            print("Advertencia: Uno de los DataFrames está vacío")
            return df_principal

        # Imprimir información antes de la unión
        print("\nEstadísticas antes de la unión:")
        print(f"Registros en df_principal: {df_principal.height}")
        print(f"Registros en df_adicional: {df_adicional.height}")

        # Realizar la unión vertical
        df_unificado = pl.concat([df_principal, df_adicional])

        # Imprimir información después de la unión
        print("\nEstadísticas después de la unión:")
        print(f"Total de registros unidos: {df_unificado.height}")

        return df_unificado

    except Exception as e:
        print(f"Error al unir DataFrames: {str(e)}")
        raise



def son_valores_similares(valor1: float, valores: list[float], margen: float = 0.01) -> bool:
    """
    Compara un valor con una lista de valores permitiendo un margen de error.

    Args:
        valor1 (float): Valor a comparar
        valores (list[float]): Lista de valores para comparar
        margen (float): Margen de error permitido (por defecto 0.01)

    Returns:
        bool: True si encuentra algún valor dentro del margen permitido
    """
    return any(abs(valor1 - valor2) <= margen for valor2 in valores)
