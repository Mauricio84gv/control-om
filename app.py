import streamlit as st
import pandas as pd
import os
import io
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment

# 1. Configuración de la pantalla
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo manteniendo el formato oficial.")

# Archivos de control en el servidor
ARCHIVO_CSV = "datos_semana_activa.csv"
PLANTILLA_EXCEL = "plantilla_original.xlsx"

# -------------------------------------------------------------------------
# SÓLO TU PANEL (BARRA LATERAL) - PARA CARGAR LA NUEVA SEMANA
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Administrador")
    st.write("Carga el archivo Excel oficial aquí.")
    
    nuevo_excel = st.file_uploader("Subir Excel de la semana", type=["xlsx"])
    
    if nuevo_excel is not None:
        if st.button("🔄 Inicializar Semana", use_container_width=True):
            try:
                # Guardar una copia exacta del archivo original con sus formatos, colores y celdas
                with open(PLANTILLA_EXCEL, "wb") as f:
                    f.write(nuevo_excel.getbuffer())

                # Leer el Excel temporalmente para la vista interactiva de los muchachos
                df_temp = pd.read_excel(nuevo_excel, header=None, dtype=str)
                df_temp = df_temp.fillna("")

                # ESCANEO: Buscar la fila real de encabezados
                fila_encabezado = 0
                for i, row in df_temp.iterrows():
                    valores_fila = [str(val).strip().upper() for val in row.values]
                    if any("SITIO" in v or "SEMANA" in v for v in valores_fila):
                        fila_encabezado = i
                        break
                
                # Leer la tabla de datos saltando la basura de arriba
                df_cargado = pd.read_excel(nuevo_excel, skiprows=fila_encabezado, dtype=str)
                df_cargado = df_cargado.fillna("")
                
                df_cargado.columns = df_cargado.columns.str.strip()
                
                # Unificar encabezados a mayúsculas para el control interno de la app
                columnas_limpias = []
                for col in df_cargado.columns:
                    nombre_limpio = str(col).strip().upper()
                    if "SITIO" in nombre_limpio:
                        nombre_limpio = "SITIO"
                    elif "SEMANA" in nombre_limpio:
                        nombre_limpio = "SEMANA"
                    columnas_limpias.append(nombre_limpio)
                
                df_cargado.columns = columnas_limpias
                
                # TEXTOS REALES DEL JEFE: Mapeamos los nombres exactos del Excel original
                col_realizo = [c for c in df_cargado.columns if "REALIZÓ" in c or "REALIZO" in c]
                col_justif = [c for c in df_cargado.columns if "POR QUÉ" in c or "POR QUE" in c or "DETALLAR" in c]
                
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

                for col in df_cargado.columns:
                    df_cargado[col] = df_cargado[col].astype(str).str.strip()

                # Guardar los datos limpios para la interfaz web
                df_cargado.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
                
                st.success("¡Estructura guardada y formato original respaldado con éxito!")
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
            
    if os.path.exists(ARCHIVO_CSV):
        st.write("---")
        if st.button("🗑️ Borrar Semana Actual", type="primary", use_container_width=True):
            if os.path.exists(ARCHIVO_CSV): os.remove(ARCHIVO_CSV)
            if os.path.exists(PLANTILLA_EXCEL): os.remove(PLANTILLA_EXCEL)
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (COMPARTIDO POR TODO EL EQUIPO)
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_CSV) and os.path.exists(PLANTILLA_EXCEL):
    try:
        df_actual = pd.read_csv(ARCHIVO_CSV, dtype=str, encoding='utf-8-sig')
        df_actual = df_actual.fillna("")
        
        num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
        
        st.info("📱 **Muchachos:** Busquen su fila, marquen el estado y denle al botón **Guardar Cambios**.")

        # Tabla interactiva en la web con los nombres reales que pide la oficina
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

        # BOTÓN GUARDAR GENERAL
        if st.button("💾 Guardar Mis Cambios del Día", use_container_width=True):
            df_editado.to_csv(ARCHIVO_CSV, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados en el sistema!")
            st.balloons()

        # Progreso
        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado["Se realizó (Si/No)"] == "Si"])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
            porcentaje = hechos / total if total > 0 else 0.0
            st.progress(porcentaje)
            
        with col2:
            # PROCESO DE INYECCIÓN CON COPIA DE BORDES Y TEXTO OFICIAL
            if st.button("🚀 Generar Reporte Oficial (.xlsx)", use_container_width=True):
                try:
                    wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                    ws = wb.active

                    idx_sitio = None
                    idx_realizo = None
                    idx_justif = None
                    fila_titulos = None

                    # Escanear encabezados en el archivo original
                    for r in range(1, 16):
                        valores = [str(ws.cell(row=r, column=c).value).strip().upper() for c in range(1, ws.max_column + 1)]
                        if any("SITIO" in v or "SEMANA" in v for v in valores):
                            fila_titulos = r
                            for c in range(1, ws.max_column + 1):
                                txt = str(ws.cell(row=r, column=c).value).strip().upper()
                                if "SITIO" in txt: idx_sitio = c
                                elif "REALIZÓ" in txt or "REALIZO" in txt: idx_realizo = c
                                elif "POR QUÉ" in txt or "POR QUE" in txt or "DETALLAR" in txt: idx_justif = c
                            break

                    if idx_sitio is None:
                        st.error("No se pudo mapear la columna SITIO en el formato original.")
                    else:
                        # Si las columnas de control no venían de la oficina, las creamos manteniendo el texto exacto
                        if idx_realizo is None:
                            idx_realizo = ws.max_column + 1
                            ws.cell(row=fila_titulos, column=idx_realizo, value="Se realizó (Si/No)")
                        if idx_justif is None:
                            idx_justif = ws.max_column + 1
                            ws.cell(row=fila_titulos, column=idx_justif, value="Si no se realizó detallar el por qué.")

                        # Mapeos de las respuestas de los muchachos
                        mapeo_realizo = dict(zip(df_editado["SITIO"], df_editado["Se realizó (Si/No)"]))
                        mapeo_justif = dict(zip(df_editado["SITIO"], df_editado["Si no se realizó detallar el por qué."]))

                        # Recorrer filas e inyectar copiando estilos (bordes, fuentes, alineación) de la columna SITIO
                        for f in range(fila_titulos + 1, ws.max_row + 1):
                            cod_sitio = str(ws.cell(row=f, column=idx_sitio).value).strip()
                            if cod_sitio in mapeo_realizo:
                                # Celda de Estado
                                celda_r = ws.cell(row=f, column=idx_realizo, value=mapeo_realizo[cod_sitio])
                                # Celda de Comentario
                                celda_j = ws.cell(row=f, column=idx_justif, value=mapeo_justif[cod_sitio])
                                
                                # COPIA DE ESTILOS DE LA CELDA BASE (Bordes, Letra y Centrado)
                                celda_base = ws.cell(row=f, column=idx_sitio)
                                if celda_base.has_style:
                                    # Bordes idénticos
                                    celda_r.border = Border(left=celda_base.border.left, right=celda_base.border.right,
                                                            top=celda_base.border.top, bottom=celda_base.border.bottom)
                                    celda_j.border = Border(left=celda_base.border.left, right=celda_base.border.right,
                                                            top=celda_base.border.top, bottom=celda_base.border.bottom)
                                    # Misma letra
                                    celda_r.font = Font(name=celda_base.font.name, size=celda_base.font.size, 
                                                        bold=celda_base.font.bold, color=celda_base.font.color)
                                    celda_j.font = Font(name=celda_base.font.name, size=celda_base.font.size, 
                                                        bold=celda_base.font.bold, color=celda_base.font.color)
                                    # Misma alineación
                                    celda_r.alignment = Alignment(horizontal="center", vertical="center")
                                    celda_j.alignment = Alignment(horizontal="left", vertical="center")

                    # Exportación del archivo clonado
                    buffer_salida = io.BytesIO()
                    wb.save(buffer_salida)
                    data_final = buffer_salida.getvalue()

                    st.download_button(
                        label="📥 Descargar Archivo Oficial Clón",
                        data=data_final,
                        file_name=f"Reporte_O&M_Semana_{num_semana}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.success("¡Archivo oficial generado con bordes y textos idénticos! Dale clic al botón de descargar arriba.")
                except Exception as ex:
                    st.error(f"Error al clonar formatos: {ex}")

    except Exception as e:
        st.error(f"Error al leer la base de datos actual: {e}")
else:
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra lateral izquierda para activar la tabla de todo el equipo.")
