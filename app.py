import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment

# 1. Configuración de pantalla y estilos profesionales limpios
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.markdown("""
    <style>
        stDecoration { display: none !important; }
        #MainMenu { visibility: hidden !important; }
        header { visibility: hidden !important; height: 0px !important; }
        footer { visibility: hidden !important; }
        .main .block-container { padding-top: 2rem !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

ARCHIVO_CSV = "datos_semana_activa.csv"
PLANTILLA_EXCEL = "plantilla_original.xlsx"

# -------------------------------------------------------------------------
# BARRA LATERAL - CONTROL DE ACCESO
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Opciones")
    modo_admin = st.checkbox("🔑 Modo Administrador")
    
    if modo_admin:
        clave = st.text_input("Introduce la clave de acceso", type="password")
        if clave == "admin123":
            st.success("👨‍💻 Modo Administrador Activo")
            st.write("---")
            st.subheader("Cargar Nueva Semana")
            nuevo_excel = st.file_uploader("Subir Excel (.xlsx)", type=["xlsx"])
            
            if nuevo_excel is not None:
                if st.button("🔄 Inicializar Semana", use_container_width=True):
                    try:
                        with open(PLANTILLA_EXCEL, "wb") as f:
                            f.write(nuevo_excel.getbuffer())

                        df_t = pd.read_excel(nuevo_excel, header=None, dtype=str).fillna("")
                        fila_h = 0
                        for i, r in df_t.iterrows():
                            vals = [str(v).strip().upper() for v in r.values]
                            if any("SITIO" in x or "SEMANA" in x for x in vals):
                                fila_h = i
                                break
                        
                        df_c = pd.read_excel(nuevo_excel, skiprows=fila_h, dtype=str).fillna("")
                        df_c.columns = df_c.columns.str.strip()
                        
                        cols_l = []
                        for col in df_c.columns:
                            c_up = str(col).strip().upper()
                            if "SITIO" in c_up: cols_l.append("SITIO")
                            elif "SEMANA" in c_up: cols_l.append("SEMANA")
                            else: cols_l.append(c_up)
                        df_c.columns = cols_l
                        
                        c_re = [c for c in df_c.columns if "REALIZ" in c]
                        c_ju = [c for c in df_c.columns if any(x in c for x in ["POR", "QUE", "JUSTI", "COME"])]
                        
                        if c_re: df_c = df_c.rename(columns={c_re[0]: "Se realizó (Si/No)"})
                        else: df_c["Se realizó (Si/No)"] = "Pendiente"
                            
                        if c_ju: df_c = df_c.rename(columns={c_ju[0]: "Si no se realizó detallar el por qué."})
                        else: df_c["Si no se realizó detallar el por qué."] = ""
                        
                        df_c.loc[df_c["Se realizó (Si/No)"] == "", "Se realizó (Si/No)"] = "Pendiente"
                        if "SITIO" in df_c.columns: df_c = df_c[df_c["SITIO"] != ""]

                        df_c = df_c.loc[:, ~df_c.columns.duplicated()]
                        v_cols = [c for c in df_c.columns if "JUSTIF" not in c and "COMENT" not in c or c == "Si no se realizó detallar el por qué."]
                        df_c = df_c[v_cols]

                        for col in df_c.columns: df_c[col] = df_c[col].astype(str).str.strip()
                        df_c.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
                        st.success("¡Estructura oficial inicializada!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")
            
            if os.path.exists(ARCHIVO_CSV):
                st.write("---")
                st.subheader("🚨 Zona de Peligro")
                if st.button("🗑️ Borrar Semana Actual", type="primary", use_container_width=True):
                    if os.path.exists(ARCHIVO_CSV): os.remove(ARCHIVO_CSV)
                    if os.path.exists(PLANTILLA_EXCEL): os.remove(PLANTILLA_EXCEL)
                    st.success("Semana borrada.")
                    st.rerun()
        elif clave != "":
            st.error("❌ Contraseña incorrecta.")
    else:
        st.info("Panel de Técnicos de Campo. Busque y llene sus datos a la derecha.")

# -------------------------------------------------------------------------
# PANEL DE TRABAJO - CAMPOS COMPARTIDOS
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_CSV) and os.path.exists(PLANTILLA_EXCEL):
    try:
        df_actual = pd.read_csv(ARCHIVO_CSV, dtype=str, encoding='utf-8-sig').fillna("")
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
            st.balloons()

        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado["Se realizó (Si/No)"] == "Si"])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
            st.progress(hechos / total if total > 0 else 0.0)
            
        with col2:
            if st.button("🚀 Generar Reporte Oficial (.xlsx)", use_container_width=True):
                try:
                    wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                    ws = wb.active
                    idx_s, idx_r, idx_j, fila_t = None, None, None, None

                    for r in range(1, 16):
                        vals = [str(ws.cell(row=r, column=c).value).strip().upper() for c in range(1, ws.max_column + 1)]
                        if any("SITIO" in v or "SEMANA" in v for v in vals):
                            fila_t = r
                            for c in range(1, ws.max_column + 1):
                                txt = str(ws.cell(row=r, column=c).value).strip().upper()
                                if "SITIO" in txt: idx_s = c
                                elif "REALIZ" in txt: idx_r = c
                                elif any(x in txt for x in ["POR", "QUE", "JUSTI", "COME"]): idx_j = c
                            break

                    if idx_s is None:
                        st.error("No se encontró la columna SITIO en el Excel.")
                    else:
                        if idx_r is None: idx_r = idx_s + 1 if idx_j != idx_s + 1 else idx_s + 2
                        if idx_j is None: idx_j = idx_r + 1

                        ws.cell(row=fila_t, column=idx_r, value="Se realizó (Si/No)")
                        ws.cell(row=fila_t, column=idx_j, value="Si no se realizó detallar el por qué.")

                        # Mapeos compactos para evitar cortes de línea bruscos en producción
                        map_r = dict(zip(df_editado["SITIO"], df_editado["Se realizó (Si/No)"]))
                        map_j = dict(zip(df_editado["SITIO"], df_editado["Si no se realizó detallar el por qué."]))

                        borde = Border(
                            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                            top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
                        )

                        for f in range(fila_t + 1, ws.max_row + 1):
                            val_s = ws.cell(row=f, column=idx_s).value
                            if val_s is None: continue
                            
                            cod_s = str(val_s).strip()
                            if cod_s in map_r:
                                c_r = ws.cell(row=f, column=idx_r, value=map_r[cod_s])
                                c_j = ws.cell(row=f, column=idx_j, value=map_j[cod_s])
                                c_r.border, c_j.border = borde, borde
                                
                                c_b = ws.cell(row=f, column=idx_s)
                                if c_b.has_style and c_b.font:
                                    c_r.font = Font(name=c_b.font.name, size=c_b.font.size, bold=c_b.font.bold, color=c_b.font.color)
                                    c_j.font = Font(name=c_b.font.name, size=c_b.font.size, bold=c_b.font.bold, color=c_b.font.color)
                                
                                c_r.alignment = Alignment(horizontal="center", vertical="center")
                                c_j.alignment = Alignment(horizontal="left", vertical="center")

                        buf = io.BytesIO()
                        wb.save(buf)
                        st.download_button(
                            label="📥 Descargar Archivo Oficial Clón",
                            data=buf.getvalue(),
                            file_name=f"Reporte_O&M_Semana_{num_sem}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.success("¡Reporte listo!")
                except Exception as ex:
                    st.error(f"Error al clonar: {ex}")
    except Exception as e:
        st.error(f"Error de base de datos: {e}")
else:
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra izquierda.")
