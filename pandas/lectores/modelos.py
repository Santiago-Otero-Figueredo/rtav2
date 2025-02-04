from pydantic import BaseModel, Field
from typing import Optional
import pandas as pd

class ConfiguracionLector(BaseModel):
    """
    Modelo base para la configuración de lectores de archivos
    """
    ruta_archivo: str = Field(..., description="Ruta al archivo que se va a leer")
    encoding: str = Field(default="utf-8", description="Codificación del archivo")
    separador: str = Field(default=",", description="Separador para archivos CSV")
    hoja: Optional[str] = Field(default=None, description="Nombre de la hoja para archivos Excel")

    class Config:
        arbitrary_types_allowed = True
