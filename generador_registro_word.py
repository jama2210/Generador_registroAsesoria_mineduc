from docx import Document
import os
import pandas as pd
import re

from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml import OxmlElement, ns

def aplicar_color_fondo(celda, color_hex="D9E1F2"):
    """
    Aplica color de fondo a una celda de tabla Word.
    """
    tc_pr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(ns.qn('w:val'), 'clear')
    shd.set(ns.qn('w:color'), 'auto')
    shd.set(ns.qn('w:fill'), color_hex)
    tc_pr.append(shd)


def agregar_titulo_y_logo(doc, ruta_logo="logo_mineduc.png"):
    """
    Agrega el logo institucional y el título del informe
    al inicio del documento.
    """

    # Logo
    if os.path.exists(ruta_logo):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_logo = p_logo.add_run()
        run_logo.add_picture(ruta_logo, width=Cm(4))

    # Espacio
    doc.add_paragraph("")

    # Título
    titulo = doc.add_heading(
        "Informe Individual de Registro de Asesoría MINEDUC",
        level=0
    )
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Separación visual
    doc.add_paragraph("")

def limpiar_documento(doc):
    """
    Limpia el contenido del documento Word (párrafos y tablas),
    pero conserva la sección para evitar errores de python-docx.
    """
    body = doc._element.body
    for element in list(body):
        # Mantener la definición de sección (<w:sectPr>)
        if element.tag.endswith('sectPr'):
            continue
        body.remove(element)

# =====================================================
# UTILIDADES
# =====================================================
def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def valor_visible(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto == "":
        return ""
    return texto

def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).strip()


# =====================================================
# PLANIFICACIÓN
# =====================================================
def obtener_objetivos_planificacion(
    df_planificacion, region, deprov, modalidad, nombre_asesoria
):

    filtro = (
        (df_planificacion["Indique su región"].str.upper().str.strip() == region.upper().strip()) &
        (df_planificacion["Deprov"].str.upper().str.strip() == deprov.upper().strip()) &
        (df_planificacion["Tipo Asesoría"].str.upper().str.strip() == modalidad.upper().strip()) &
        (df_planificacion["Nombre Asesoría"].str.upper().str.strip() == nombre_asesoria.upper().strip())
    )

    fila = df_planificacion[filtro]

    if fila.empty:
        return "No informado", "No informado"

    return (
        normalizar_texto(fila.iloc[0].get("Objetivo estratégico de la asesoría")),
        normalizar_texto(fila.iloc[0].get("Objetivo anual de la asesoría")),
    )


# =====================================================
# SECCIONES DEL INFORME
# =====================================================
def agregar_encabezado(doc, region, deprov, modalidad, nombre_asesoria,
                       objetivo_estrategico, objetivo_anual):

    tabla = doc.add_table(rows=6, cols=2)
    tabla.style = "Table Grid"

    filas = [
        ("Región", region),
        ("DEPROV", deprov),
        ("Modalidad", modalidad),
        ("RBD / Nombre de la Asesoría", nombre_asesoria),
        ("Objetivo estratégico de la asesoría", objetivo_estrategico),
        ("Objetivo anual de la asesoría", objetivo_anual),
    ]

    for i, (k, v) in enumerate(filas):
        tabla.cell(i, 0).text = k
        tabla.cell(i, 1).text = v


def agregar_antecedentes_generales(doc, datos):

    doc.add_heading("ANTECEDENTES GENERALES", level=1)

    tabla = doc.add_table(rows=1, cols=8)
    tabla.style = "Table Grid"

    encabezados = [
        "N° de sesión", "Supervisor", "Fecha de la reunión",
        "Hora inicio", "Hora término", "Tipo de encuentro",
        "Participantes", "Objetivo de la sesión"
    ]

    
    for i, h in enumerate(encabezados):
        celda = tabla.rows[0].cells[i]
        celda.text = h
        aplicar_color_fondo(celda)

    datos = datos.sort_values("NUM SESIÓN")

    for _, fila in datos.iterrows():
        r = tabla.add_row().cells
        r[0].text = str(fila.get("NUM SESIÓN", ""))
        r[1].text = valor_visible(fila.get("Supervisor"))
        r[2].text = valor_visible(fila.get("Fecha de la reunión"))       
        r[3].text = valor_visible(
            fila.get("Indique hora aproximada de inicio de la reunión")
            or fila.get("Hora de inicio")
        )
        r[4].text = valor_visible(
            fila.get("Indique hora aproximada de término de la reunión")
            or fila.get("Hora de finalización")
        )
        r[5].text = valor_visible(fila.get("Tipo de encuentro"))
        r[6].text = valor_visible(fila.get("Participantes de la reunión"))
        r[7].text = valor_visible(fila.get("Objetivo de la sesión de asesoría"))


def agregar_eid_capacidades_practicas(doc, datos):

    columnas = [
        "Estándares Indicativos de Desempeño asociado",
        "Capacidad abordada en la sesión de asesoría",
        "Dimensión asociada al EID seleccionado",
        "Indique qué práctica se está abordando en el establecimiento a partir del EID trabajado."
    ]

    if not any(datos[c].notna().any() for c in columnas if c in datos.columns):
        return

    doc.add_heading("EID, CAPACIDADES Y PRÁCTICAS ABORDADAS", level=1)

    tabla = doc.add_table(rows=1, cols=2)
    tabla.style = "Table Grid"

    tabla.rows[0].cells[0].text = "N° de sesión"
    tabla.rows[0].cells[1].text = "Detalle"

    for _, fila in datos.iterrows():
        r = tabla.add_row().cells
        r[0].text = str(fila.get("NUM SESIÓN", ""))

        partes = []
        for c in columnas:
            texto = valor_visible(fila.get(c))
            if texto:
                partes.append(texto)

        r[1].text = "\n".join(partes)


def agregar_otras_indicaciones(doc, datos):

    columnas = [
        ("Estrategia /acciones de acompañamiento realizadas", "Estrategia / acciones"),
        ("Indique si surgió algún tema no planificado que impacta a la asesoría directa / trabajo en red.",
         "Tema no planificado"),
        ("¿Se evidencian cambios o progresos en relación a la práctica abordada?", "Cambios evidenciados"),
        ("¿Qué dificultades están limitando el avance de la(s) práctica(s) abordada(s)?", "Dificultades"),
        ("Acuerdos concretos de la reunión", "Acuerdos"),
        ("Próximos pasos que se realizarán antes de la próxima sesión y responsables de cada acción",
         "Próximos pasos"),
        ("Comentarios u observaciones", "Comentarios")
    ]

    if not any(datos[c].notna().any() for c, _ in columnas if c in datos.columns):
        return

    doc.add_heading("OTRAS INDICACIONES", level=1)

    tabla = doc.add_table(rows=1, cols=len(columnas) + 1)
    tabla.style = "Table Grid"

    tabla.rows[0].cells[0].text = "N° sesión"
    for i, (_, titulo) in enumerate(columnas):
        tabla.rows[0].cells[i + 1].text = titulo

    for _, fila in datos.iterrows():
        r = tabla.add_row().cells
        r[0].text = str(fila.get("NUM SESIÓN", ""))

        for i, (c, _) in enumerate(columnas):
            r[i + 1].text = valor_visible(fila.get(c))


# =====================================================
# FUNCIÓN PRINCIPAL
# =====================================================
def generar_informes_registro(df_registros, df_planificacion,
                              carpeta_salida, plantilla="plantilla_registro.docx"):

        # Normalizar nombres de columnas del registro
    df_registros.columns = (
        df_registros.columns
        .astype(str)
        .str.strip()
        .str.replace('\u00a0', ' ')   # espacios duros
    )
    
    df_planificacion.columns = (
        df_planificacion.columns
        .astype(str)
        .str.strip()
        .str.replace('\u00a0', ' ')
    )

    os.makedirs(carpeta_salida, exist_ok=True)

    modalidad_mapa = {
        "Directa EE": "Directa a Establecimientos",
        "Directa a Sostenedor": "Directa a Sostenedor",
        "Red EE": "Red de Establecimientos",
        "Red de Sostenedor": "Red de Sostenedor"
    }

    grupos = df_registros.groupby(
        ["INDIQUE SU REGIÓN", "Deprov", "Modalidad Asesoría", "Nombre Asesoría"]
    )

    for (region, deprov, modalidad, nombre), datos in grupos:

        region = normalizar_texto(region)
        deprov = normalizar_texto(deprov)
        modalidad = normalizar_texto(modalidad)
        nombre = normalizar_texto(nombre)

        obj_est, obj_anual = obtener_objetivos_planificacion(
            df_planificacion, region, deprov, modalidad, nombre
        )

        doc = Document(plantilla)
        limpiar_documento(doc)
        
        # Título institucional y logo
        agregar_titulo_y_logo(doc)
        
        # Contenido del informe
        agregar_encabezado(doc, region, deprov, modalidad, nombre, obj_est, obj_anual)

        agregar_antecedentes_generales(doc, datos)
        agregar_eid_capacidades_practicas(doc, datos)
        agregar_otras_indicaciones(doc, datos)

        carpeta_region = limpiar_nombre_archivo(region)
        carpeta_deprov = limpiar_nombre_archivo(deprov)
        carpeta_modalidad = modalidad_mapa.get(modalidad, limpiar_nombre_archivo(modalidad))

        ruta_final = os.path.join(
            carpeta_salida,
            carpeta_region,
            carpeta_deprov,
            carpeta_modalidad
        )

        os.makedirs(ruta_final, exist_ok=True)

        archivo = limpiar_nombre_archivo(
            f"Informe_Registro_{region}_{deprov}_{modalidad}_{nombre}.docx"
        )

        doc.save(os.path.join(ruta_final, archivo))
