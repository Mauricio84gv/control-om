import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Border, Side, Alignment
from streamlit_gsheets import GSheetsConnection

# 1. Configuración de pantalla en modo ancho
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

# ENLACE DIRECTO DE TU HOJA DE GOOGLE
URL_HOJA_DIRECTA = "https://docs.google.com/spreadsheets/d/1veXDuDIPG7KhUM3drp034Y7NfdY-3I0NyAqhRtRiFt4/edit"

# INYECCIÓN CSS AVANZADA: ELIMINAR BARRA DE DESARROLLO COMPLETAMENTE
st.markdown("""
    <style>
        div[data-testid="stStatusWidget"] { display: none !important; }
        [data-testid="stDecoration"] { display: none !important; }
        [data-testid="stActionButton"] { display: none !important; }
        header { visibility: hidden !important; display: none !important; height: 0px !important; }
        #MainMenu { visibility: hidden !important; }
        footer { visibility: hidden !important; }
        .main .block-container { padding-top: 1rem !important; }
        [data-testid="stSidebarCollapseButton"] { display: none !important; visibility: hidden !important; }
    </style>
""", unsafe_allow_html=True)

if "mostrar_globos" not in st.session_state:
    st.session_state.mostrar_globos = False
if st.session_state.mostrar_globos:
    st.balloons()
    st.session_state.mostrar_globos = False

# CONEXIÓN DIRECTA A GOOGLE SHEETS UTILIZANDO LOS SECRETS DEL SISTEMA
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

# ACCESO DIRECTO DE ADMINISTRACIÓN
with st.expander("🛠️ Panel de Acceso - Administrador (Dar clic aquí para cargar Excel)"):
    clave_directa = st.text_input("Introduce la clave de acceso de O&M", type="password", key="clave_main")
    
    if clave_directa == "admin123":
        st.success("👨‍💻 Modo Administrator Activo")
        st.write("---")
        st.subheader("Cargar Nueva Semana Activa")
        nuevo_excel_main = st.file_uploader("Subir archivo Excel (.xlsx)", type=["xlsx"], key="uploader_main")
        
        if nuevo_excel_main is not None and conn is not None:
            if st.button("🔄 Inicializar Nueva Semana", use_container_width=True):
                try:
                    # Guardar la plantilla en el servidor local
                    with open("plantilla_original.xlsx", "wb") as f:
                        f.write(nuevo_excel_main.getbuffer())

                    df_t = pd.read_excel(nuevo_excel_main, header=None, dtype=str).fillna("")
                    fila_h = 0
                    for i, r in df_t.iterrows():
                        vals = [str(v).strip().upper() for v in r.values]
                        if any("SITIO" in x or "SEMANA" in x for x in vals):
                            fila_h = i
                            break
                    
                    df_c = pd.read_excel(nuevo_excel_main, skiprows=fila_h, dtype=str).fillna("")
                    df_c.columns = df_c.columns.str.strip()
                    columnas_reales = list(df_c.columns)
                    
                    if len(columnas_reales) >= 8: columnas_reales[7] = "Se realizó (Si/No)"
                    if len(columnas_reales) >= 9: columnas_reales[8] = "Si no se realizó detallar el por qué."
                    df_c.columns = columnas_reales
                    
                    df_c
