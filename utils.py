import pandas as pd
import re


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def valor_visible(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in ["", "no", "n/a", "no aplica"]:
        return ""
    return texto


def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).strip()


def clave_unidad(region, deprov, modalidad, nombre_asesoria):
    """
    Genera clave normalizada para cruce registro-planificación.
    """
    return (
        normalizar_texto(region).upper(),
        normalizar_texto(deprov).upper(),
        normalizar_texto(modalidad).upper(),
        normalizar_texto(nombre_asesoria).upper()
    )