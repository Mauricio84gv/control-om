import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment

# 1. Configuración de la pantalla
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

# -------------------------------------------------------------------------
# INYECCIÓN CSS: OCULTAR DE RAÍZ MENÚS, BOTONES DE GITHUB, EDITORES Y SHARE
# -------------------------------------------------------------------------
st.markdown("""
    <style>
        /* Ocultar la barra superior completa de desarrollo (Share, Edit, GitHub) */
        stDecoration {
            display: none !important;
        }
        /* Ocultar el menú de tres puntos de arriba a la derecha */
        #MainMenu {
            visibility: hidden !important;
        }
        header {
            visibility: hidden !important;
            height: 0px !important;
        }
        /* Eliminar espacios muertos superiores creados por la barra */
        .main .block-container {
            padding-top: 2rem !important;
        }
        /* Ocultar leyendas de "Made with Streamlit" al final */
        footer {
            visibility: hidden !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo manteniendo el formato oficial.")

# Archivos de control en el servidor
ARCHIVO_CSV = "datos_semana_activa.csv"
PLANTILLA_EXCEL = "plantilla_original.xlsx"

# -------------------------------------------------------------------------
# BARRA LATERAL CON CONTROL DE ACCESO REAL
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Opciones")
    
    # Checkbox para activar el modo admin de forma discreta
    modo_admin = st.checkbox("🔑 Modo Administrador")
    
    if modo_admin:
        # Solo si marcan la casilla, se les pide la clave
        clave = st.text_input("Introduce la clave de acceso", type="password")
        
        # Si la clave es correcta, se liberan tus herramientas exclusivas
        if clave == "admin123": # Podés cambiar "admin123" por la contraseña que quieras
            st.success("👨‍💻 Modo Administrador Activo")
            st.write("---")
            st.subheader("Cargar Nueva Semana")
            nuevo_excel = st.file_uploader("Subir Excel oficial (.xlsx)", type=["xlsx"])
            
            if nuevo_excel is not None:
                if st.button("🔄 Inicializar Semana", use_container_width=True):
                    try:
                        with open(PLANTILLA_EXCEL, "wb") as f:
                            f.write(nuevo_excel.getbuffer())

                        df_temp = pd.read_excel(nuevo_excel, header=None, dtype=str)
                        df_temp = df_temp.fillna("")

                        fila_encabezado = 0
                        for i, row in df_temp.iterrows():
                            valores_fila = [str(val).strip().upper() for val in row.values]
                            if any("SITIO" in v or "SEMANA" in v for v in valores_fila):
                                fila_encabezado = i
                                break
                        
                        df_cargado = pd.read_excel(nuevo_excel, skiprows=fila_encabezado, dtype=str)
                        df_cargado = df_cargado.fillna("")
                        df_cargado.columns = df_cargado.columns.str.strip()
                        
                        columnas_limpias = []
                        for col in df_cargado.columns:
                            nombre_limpio = str(col).strip().upper()
                            if "SITIO" in nombre_limpio: nombre_limpio = "SITIO"
                            elif "SEMANA" in nombre_limpio: nombre_limpio = "SEMANA"
                            columnas_limpias.append(nombre_limpio)
                        
                        df_cargado.columns = columnas_limpias
                        
                        col_realizo = [c for c in df_cargado.columns if "REALIZÓ" in c or "REALIZO" in c]
                        col_justif = [c for c in df_cargado.columns if "POR QUÉ" in c or "POR QUE" in c or "DETALLAR" in c or "JUSTIF" in c or "COMENT" in c]
                        
                        if col_realizo:
                            df_cargado = df_cargado.rename(columns={col_realizo[0]: "Se realizó (Si/No)"})
                        else:
                            df_cargado["Se realizó (Si/No)"] = "Pendiente"
                            
                        if col_justif:
                            df_cargado = df_cargado.rename(columns={col_justif[0]: "Si no se realizó detallar el por qué."})
                        else:
                            df_cargado["Si no se realizó detallar el por qué."] = ""
                        
                        df_cargado.loc[df_cargado["Se realizó (Si/No)"] == "", "Se realizó (Si/No)"] = "Pendiente"

                        if "SITIO" in df_cargado.columns:
                            df_cargado = df_cargado[df_cargado["SITIO"] != ""]

                        df_cargado = df_cargado.loc[:, ~df_cargado.columns.duplicated()]
                        columnas_validas = [c for c in df_cargado.columns if "JUSTIFICACI" not in c.upper() and "COMENTARIO" not in c.upper() or c == "Si no se realizó detallar el por qué."]
                        df_cargado = df_cargado[columnas_validas]

                        for col in df_cargado.columns:
                            df_cargado[col] = df_cargado[col].astype(str).str.strip()

                        df_cargado.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
                        st.success("¡Estructura oficial inicializada con éxito!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al procesar el archivo: {e}")
            
            if os.path.exists(ARCHIVO_CSV):
                st.write("---")
                st.subheader("🚨 Zona de Peligro")
                if st.button("🗑️ Borrar Semana Actual", type="primary", use_container_width=True):
                    if os.path.exists(ARCHIVO_CSV): 
                        os.remove(ARCHIVO_CSV)
                    if os.path.exists(PLANTILLA_EXCEL): 
                        os.remove(PLANTILLA_EXCEL)
                    st.success("Semana borrada exitosamente.")
                    st.rerun()
        else:
            if clave != "":
                st.error("❌ Contraseña incorrecta.")
    else:
        st.info("Panel de llenado para Técnicos de Campo. Selecciona tus sitios en la tabla derecha.")

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (LO QUE VEN LOS MUCHACHOS COMPARTIDO)
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_CSV) and os.path.exists(PLANTILLA_EXCEL):
    try:
        df_actual = pd.read_csv(ARCHIVO_CSV, dtype=str, encoding='utf-8-sig')
        df_actual = df_actual.fillna("")
        
        num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
        st.info("📱 **Muchachos:** Busquen su fila, marquen el estado y denle al botón **Guardar Cambios** al final.")

        df_editado = st.data_editor(
            df_actual,
            column_config={
                "Se realizó (Si/No)": st.column_config.SelectboxColumn(
                    "Se realizó (Si/No)",
                    options=["Pendiente", "Si", "No", "Cancelado"],
                    required=True,
                    width="medium"
                ),
                "Si no se realizó detallar el por qué.": st.column_config.TextColumn(
                    "Si no se
