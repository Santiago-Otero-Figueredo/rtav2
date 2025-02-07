from pydantic import BaseModel, Field
from typing import Optional

class ConfiguracionLector(BaseModel):
    """
    Modelo base para la configuración de lectores de archivos
    """
    ruta_carpeta: Optional[str] = Field(default=None, description="Ruta a la carpeta que contiene los archivos que se van a leer")
    ruta_archivo: Optional[str] = Field(default=None, description="Nombre del archivo específico a leer (opcional)")
    encoding: str = Field(default="utf-8", description="Codificación del archivo")
    separador: str = Field(default=",", description="Separador para archivos CSV")
    cargue_inicial: bool = Field(default=True, description="Se carga el archivo por primera vez")
    hoja: Optional[str] = Field(default=None, description="Nombre de la hoja para archivos Excel")
    mapeo_indices_nombres_columnas: Optional[dict] = Field(default=None, description="Diccionario con los cambios de nombres de las columnas según su índice.")

    class Config:
        arbitrary_types_allowed = True
