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
        clave = st.text_input("Introduce la clave de acceso", type="password")
        
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
                    if os.path.exists(ARCHIVO_CSV): os.remove(ARCHIVO_CSV)
                    if os.path.exists(PLANTILLA_EXCEL): os.remove(PLANTILLA_EXCEL)
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
                    "Si no se realizó detallar el por qué.",
                    width="large"
                )
            },
            disabled=[col for col in df_actual.columns if col not in ["Se realizó (Si/No)", "Si no se realizó detallar el por qué."]],
            hide_index=True,
            use_container_width=True
        )

        st.write("---")

        if st.button("💾 Guardar Mis Cambios del Día", use_container_width=True):
            df_editado.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados en el sistema!")
            st.balloons()

        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado["Se realizó (Si/No)"] == "Si"])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
            porcentaje = hechos / total if total > 0 else 0.0
            st.progress(porcentaje)
            
        with col2:
            if st.button("🚀 Generar Reporte Oficial (.xlsx)", use_container_width=True):
                try:
                    wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                    ws = wb.active

                    idx_sitio, idx_realizo, idx_justif, fila_titulos = None, None, None, None

                    # Encontrar la fila real de títulos en la plantilla
                    for r in range(1, 16):
                        valores = [str(ws.cell(row=r, column=c).value).strip().upper() for c in range(1, ws.max_column + 1)]
                        if any("SITIO" in v or "SEMANA" in v for v in valores):
                            fila_titulos = r
                            for c in range(1, ws.max_column + 1):
                                txt = str(ws.cell(row=r, column=c).value).strip().upper()
                                if "SITIO" in txt:
                                    idx_sitio = c
                                elif "REALIZÓ" in txt or "REALIZO" in txt:
                                    idx_realizo = c
                                elif "POR QUÉ" in txt or "POR QUE" in txt or "DETALLAR" in txt or "JUSTIF" in txt or "COMENT" in txt:
                                    idx_justif = c
                            break

                    if idx_sitio is None:
                        st.error("No se pudo encontrar la columna SITIO en la plantilla.")
                    else:
                        if idx_realizo is None:
                            idx_realizo = idx_sitio + 1 if idx_justif != idx_sitio + 1 else idx_sitio + 2
                        if idx_justif is None:
                            idx_justif = idx_realizo + 1

                        # Forzar los encabezados oficiales en la fila correcta
                        ws.cell(row=fila_titulos, column=idx_realizo, value="Se realizó (Si/No)")
                        ws.cell(row=fila_titulos, column=idx_justif, value="Si no se realizó detallar el por qué.")

                        # Mapear los datos ingresados por los muchachos
                        mapeo_realizo = dict(zip(df_editado["SITIO"], df_editado["Se realizó (Si/No)"]))
                        mapeo_justif = dict(zip(df_editado["SITIO"], df_editado["Si no se realizó detallar el por qué."]))

                        borde_cuadrícula = Border(
                            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                            top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
                        )

                        # Inyectar las respuestas celda por celda evitando filas muertas o vacías
                        for f in range(fila_titulos + 1, ws.max_row + 1):
                            val_sitio = ws.cell(row=f, column=idx_sitio).value
                            if val_sitio is None:
                                continue
                                
                            cod_sitio = str(val_sitio).strip()
                            if cod_sitio in mapeo_
