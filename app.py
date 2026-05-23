import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment

# 1. Configuración de pantalla en modo ancho
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

# -------------------------------------------------------------------------
# INYECCIÓN CSS AVANZADA: ELIMINAR BARRA DE DESARROLLO COMPLETAMENTE
# -------------------------------------------------------------------------
st.markdown("""
    <style>
        div[data-testid="stStatusWidget"] { display: none !important; }
        [data-testid="stDecoration"] { display: none !important; }
        [data-testid="stActionButton"] { display: none !important; }
        header { 
            visibility: hidden !important; 
            display: none !important;
            height: 0px !important; 
        }
        #MainMenu { visibility: hidden !important; }
        footer { visibility: hidden !important; }
        .main .block-container { padding-top: 1rem !important; }
        [data-testid="stSidebarCollapseButton"] {
            display: none !important;
            visibility: hidden !important;
        }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# ACCESO DIRECTO DE ADMINISTRACIÓN EN LA PÁGINA PRINCIPAL
# -------------------------------------------------------------------------
with st.expander("🛠️ Panel de Acceso - Administrador (Dar clic aquí para cargar Excel)"):
    clave_directa = st.text_input("Introduce la clave de acceso de O&M", type="password", key="clave_main")
    
    if clave_directa == "admin123":
        st.success("👨‍💻 Modo Administrador Activo")
        st.write("---")
        st.subheader("Cargar Nueva Semana Activa")
        nuevo_excel_main = st.file_uploader("Subir archivo Excel (.xlsx)", type=["xlsx"], key="uploader_main")
        
        if nuevo_excel_main is not None:
            if st.button("🔄 Inicializar Nueva Semana", use_container_width=True, key="btn_init_main"):
                try:
                    with open("plantilla_original.xlsx", "wb") as f:
                        f.write(nuevo_excel_main.getbuffer())

                    # Leer la plantilla original
                    df_t = pd.read_excel(nuevo_excel_main, header=None, dtype=str).fillna("")
                    fila_h = 0
                    for i, r in df_t.iterrows():
                        vals = [str(v).strip().upper() for v in r.values]
                        if any("SITIO" in x or "SEMANA" in x for x in vals):
                            fila_h = i
                            break
                    
                    # Forzar lectura limpia de columnas basadas en la plantilla del supervisor
                    df_c = pd.read_excel(nuevo_excel_main, skiprows=fila_h, dtype=str).fillna("")
                    df_c.columns = df_c.columns.str.strip()
                    
                    # Renombrar columnas fijas por su posición exacta para que la app no se confunda
                    columnas_reales = list(df_c.columns)
                    
                    # La columna 8 (Índice 7) es la del Estado (letra blanca o vacía arriba)
                    if len(columnas_reales) >= 8:
                        columnas_reales[7] = "Se realizó (Si/No)"
                    # La columna 9 (Índice 8) es la de Justificación
                    if len(columnas_reales) >= 9:
                        columnas_reales[8] = "Si no se realizó detallar el por qué."
                        
                    df_c.columns = columnas_reales
                    
                    # Filtrar filas vacías o basura
                    df_c = df_c[df_c["SITIO"] != ""]
                    df_c = df_c[~df_c["SITIO"].str.contains("Total", na=False, case=False)]
                    
                    # Forzar estado por defecto
                    df_c.loc[df_c["Se realizó (Si/No)"] == "", "Se realizó (Si/No)"] = "Pendiente"
                    
                    # Guardar base de datos limpia para los muchachos
                    df_c.to_csv("datos_semana_activa.csv", index=False, encoding='utf-8-sig')
                    st.success("¡Estructura oficial inicializada!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
        
        if os.path.exists("datos_semana_activa.csv"):
            st.write("---")
            st.subheader("🚨 Zona de Peligro")
            if st.button("🗑️ Borrar Semana Actual", type="primary", use_container_width=True, key="btn_delete_main"):
                if os.path.exists("datos_semana_activa.csv"): os.remove("datos_semana_activa.csv")
                if os.path.exists("plantilla_original.xlsx"): os.remove("plantilla_original.xlsx")
                st.success("Semana borrada.")
                st.rerun()
    elif clave_directa != "":
        st.error("❌ Contraseña incorrecta.")

st.write("---")

# Título de la aplicación
st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

ARCHIVO_CSV = "datos_semana_activa.csv"
PLANTILLA_EXCEL = "plantilla_original.xlsx"

# -------------------------------------------------------------------------
# PANEL DE TRABAJO - CAMPOS COMPARTIDOS
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_CSV) and os.path.exists(PLANTILLA_EXCEL):
    try:
        df_base = pd.read_csv(ARCHIVO_CSV, dtype=str, encoding='utf-8-sig').fillna("")
        df_actual = df_base[df_base["SITIO"] != ""]
        
        num_sem = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        
        st.subheader(f"📋 Sitios Activos - Semana {num_sem}")
        st.info("📱 **Muchachos:** Marquen el estado de sus sitios y denle al botón **Guardar Cambios** al final.")

        df_editado = st.data_editor(
            df_actual,
            column_config={
                "Se realizó (Si/No)": st.column_config.SelectboxColumn(
                    "Se realizó (Si/No)", options=["Pendiente", "Si", "No", "Cancelado"], required=True, width="medium"
                ),
                "Si no se realizó detallar el por qué.": st.column_config.TextColumn(
                    "Si no se realizó detallar el por qué.", width="large"
                )
            },
            disabled=[c for c in df_actual.columns if c not in ["Se realizó (Si/No)", "Si no se realizó detallar el por qué."]],
            hide_index=True,
            use_container_width=True
        )

        st.write("---")
        if st.button("💾 Guardar Mis Cambios del Día", use_container_width=True):
            df_editado.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados en el sistema!")
            st.rerun()

        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado["Se realizó (Si/No)"] == "Si"])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
            st.progress(hechos / total if total > 0 else 0.0)
            
        with col2:
            try:
                wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                ws = wb.active
                idx_s = None
                fila_t = None

                # 1. Encontrar la fila de encabezados y la columna de SITIO
                for r in range(1, 15):
                    vals = [str(ws.cell(row=r, column=c).value).strip().upper() for c in range(1, ws.max_column + 1)]
                    if any("SITIO" in v or "SEMANA" in v for v in vals):
                        fila_t = r
                        for c in range(1, ws.max_column + 1):
                            txt = str(ws.cell(row=r, column=c).value).strip().upper()
                            if "SITIO" in txt: 
                                idx_s = c
                                break
                        break

                if idx_s is None:
                    st.error("No se encontró la columna SITIO en la plantilla.")
                else:
                    # FORZADO DE POSICIONES REALES DEL SUPERVISOR (Columna 8 y Columna 9)
                    idx_r = 8  # Columna H
                    idx_j = 9  # Columna I

                    # Mapeos de datos editados por los muchachos
                    map_r = dict(zip(df_editado["SITIO"], df_editado["Se realizó (Si/No)"]))
                    map_j = dict(zip(df_editado["SITIO"], df_editado["Si no se realizó detallar el por qué."]))

                    borde = Border(
                        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
                    )

                    # 2. Inyectar datos directamente sobre las celdas exactas sin mover nada
                    for f in range(fila_t + 1, ws.max_row + 1):
                        val_s = ws.cell(row=f, column=idx_s).value
                        if val_s is None: continue
                        
                        cod_s = str(val_s).strip()
                        if cod_s in map_r:
                            c_r = ws.cell(row=f, column=idx_r, value=map_r[cod_s])
                            c_j = ws.cell(row=f, column=idx_j, value=map_j[cod_s])
                            
                            c_r.border, c_j.border = borde, borde
                            
                            # Mantener formato de letra original
                            c_base = ws.cell(row=f, column=idx_s)
                            if c_base.has_style and c_base.font:
                                font_oficial = Font(
                                    name=c_base.font.name, 
                                    size=c_base.font.size, 
                                    bold=c_base.font.bold, 
                                    color=c_base.font.color if hasattr(c_base.font, 'color') else None
                                )
                                c_r.font = font_oficial
                                c_j.font = font_oficial
                            
                            c_r.alignment = Alignment(horizontal="center", vertical="center")
                            c_j.alignment = Alignment(horizontal="left", vertical="center")

                    buf = io.BytesIO()
                    wb.save(buf)
                    
                    # Botón verde oficial de descarga en Excel (.xlsx)
                    st.download_button(
                        label=f"📥 Descargar Reporte Semana {num_sem} (.xlsx)",
                        data=buf.getvalue(),
                        file_name=f"Reporte_O&M_Semana_{num_sem}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as ex:
                st.error(f"Error al clonar formatos: {ex}")
    except Exception as e:
        st.error(f"Error de base de datos: {e}")
else:
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel (.xlsx) en el panel superior.")
