import streamlit as st
import pandas as pd
import os
from generador_registro_word import generar_informes_registro
import shutil
import zipfile

# Ejemplo de uso: doc = Document()

# --- Imports para openpyxl ---
import openpyxl
# =====================================================
# CONFIGURACIÓN GENERAL
# =====================================================
st.set_page_config(
    page_title="Generador de Informes de Registro de Asesoría MINEDUC",
    layout="wide"
)

st.markdown("""
<style>
.stApp { background-color:#f4f6f8; }
h1,h2,h3 { color:#003366; }
.stButton>button{
    background-color:#d52b1e;
    color:white;
    font-weight:bold;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# ENCABEZADO
# =====================================================
col1, col2 = st.columns([1, 5])

with col1:
    st.image("logo_mineduc.png", width=140)

with col2:
    st.title("Generador de Informes de Registro de Asesoría")
    st.write("Implementación del apoyo directo y en red – MINEDUC")

st.divider()

# =====================================================
# SUBIDA DE ARCHIVOS
# =====================================================
st.subheader("1. Carga de archivos")

col1, col2 = st.columns(2)

with col1:
    archivo_registros = st.file_uploader(
        "📄 Suba el Excel de REGISTROS de Asesoría",
        type=["xlsx"]
    )

with col2:
    archivo_planificacion = st.file_uploader(
        "📄 Suba el Excel de PLANIFICACIÓN de Asesoría",
        type=["xlsx"]
    )

if not archivo_registros or not archivo_planificacion:
    st.info("Debe cargar ambos archivos para continuar.")
    st.stop()

# =====================================================
# CARGA DE DATAFRAMES
# =====================================================
df_registros = pd.read_excel(archivo_registros)
df_planificacion = pd.read_excel(archivo_planificacion)

# =====================================================
# VALIDACIÓN DE COLUMNAS
# =====================================================
st.subheader("2. Validación de archivos")

columnas_registro_obligatorias = [
    "Nombre",
    "Correo electrónico",
    "INDIQUE SU REGIÓN",
    "Deprov",
    "Modalidad Asesoría",
    "Nombre Asesoría",
    "Fecha de la reunión",
    "NUM SESIÓN"
]

columnas_planificacion_obligatorias = [
    "Indique su región",
    "Deprov",
    "Tipo Asesoría",
    "Nombre Asesoría",
    "Objetivo estratégico de la asesoría",
    "Objetivo anual de la asesoría"
]

faltantes_registros = [
    c for c in columnas_registro_obligatorias
    if c not in df_registros.columns
]

faltantes_planificacion = [
    c for c in columnas_planificacion_obligatorias
    if c not in df_planificacion.columns
]

if faltantes_registros:
    st.error("❌ El Excel de REGISTROS no contiene las siguientes columnas:")
    for c in faltantes_registros:
        st.write("-", c)
    st.stop()

if faltantes_planificacion:
    st.error("❌ El Excel de PLANIFICACIÓN no contiene las siguientes columnas:")
    for c in faltantes_planificacion:
        st.write("-", c)
    st.stop()

st.success("✅ Ambos archivos contienen las columnas mínimas requeridas.")

# =====================================================
# VISTAS PREVIAS
# =====================================================
st.subheader("3. Vista previa de los datos")

with st.expander("🔍 Ver primeras filas – REGISTROS"):
    st.dataframe(df_registros.head())

with st.expander("🔍 Ver primeras filas – PLANIFICACIÓN"):
    st.dataframe(df_planificacion.head())

st.divider()

# =====================================================
# RESUMEN AUTOMÁTICO
# =====================================================
st.subheader("4. Resumen general")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Total de registros", len(df_registros))

with c2:
    st.metric("Unidades asesoradas", df_registros["Nombre Asesoría"].nunique())

with c3:
    st.metric("Regiones", df_registros["INDIQUE SU REGIÓN"].nunique())

with c4:
    st.metric("Excel planificación", "Cargado")

st.info(
    "ℹ️ En el siguiente paso se generará un informe Word por cada "
    "Unidad Asesorada (Región + DEPROV + Modalidad + Nombre Asesoría)."
)

st.divider()

# =====================================================
# SIGUIENTE ETAPA
# =====================================================
st.success(
    "Base del generador creada correctamente. "
    "El siguiente paso es implementar la generación de los documentos Word."
)

# =====================================================
# GENERACIÓN DE INFORMES
# =====================================================
st.subheader("5. Generación de informes")

if st.button("📄 Generar Informes de Registro"):

    carpeta_salida = "informes_registro"

    # Limpiar carpeta previa
    if os.path.exists(carpeta_salida):
        shutil.rmtree(carpeta_salida)
    os.makedirs(carpeta_salida, exist_ok=True)

    with st.spinner("Generando informes de registro…"):

        generar_informes_registro(
            df_registros=df_registros,
            df_planificacion=df_planificacion,
            carpeta_salida=carpeta_salida
        )

    # Crear ZIP
    zip_nombre = "Informes_Registro_Asesoria.zip"

    
    with zipfile.ZipFile(zip_nombre, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(carpeta_salida):
            for file in files:
                ruta_completa = os.path.join(root, file)
                ruta_en_zip = os.path.relpath(ruta_completa, carpeta_salida)
                zipf.write(ruta_completa, ruta_en_zip)


    st.success("✅ Informes generados correctamente")

    with open(zip_nombre, "rb") as f:
        st.download_button(
            "⬇️ Descargar Informes de Registro (ZIP)",
            f,
            file_name=zip_nombre,
            mime="application/zip"
        )
