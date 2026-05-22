import streamlit as st
import pandas as pd
import os

# 1. Configuración de la pantalla (Debe ser la primera línea de todas)
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

# Nombre del archivo físico compartido en el servidor
ARCHIVO_SERVIDOR = "datos_semana_activa.csv"

# Columnas técnicas de origen indispensables
COLUMNAS_ORIGEN = ['SEMANA', 'ZONA', 'SITIO', 'RIR', 'RDA', 'NOMBRE', 'TORRERA']

# -------------------------------------------------------------------------
# SÓLO TU PANEL (BARRA LATERAL) - PARA CARGAR LA NUEVA SEMANA
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Administrador")
    st.write("Carga el archivo Excel del lunes aquí.")
    
    nuevo_excel = st.file_uploader("Subir Excel de la semana", type=["xlsx"])
    
    if nuevo_excel is not None:
        if st.button("🔄 Inicializar Semana", use_container_width=True):
            try:
                # Leer el archivo Excel nuevo desde la barra lateral
                df_cargado = pd.read_excel(nuevo_excel)
                
                # Limpiar espacios en blanco invisibles en los encabezados
                df_cargado.columns = df_cargado.columns.str.strip()
                
                # Validar que al menos las 7 columnas base existan
                columnas_faltantes = [col for col in COLUMNAS_ORIGEN if col not in df_cargado.columns]
                if columnas_faltantes:
                    st.error(f"❌ Al Excel le faltan columnas base obligatorias: {columnas_faltantes}")
                    st.stop()
                
                # Conservar únicamente la información técnica real del Excel
                df_recortado = df_cargado[COLUMNAS_ORIGEN].copy()
                
                # Eliminar filas fantasma donde el SITIO esté completamente vacío
                df_recortado = df_recortado.dropna(subset=['SITIO'])
                
                # Crear de forma limpia y nativa las dos columnas de trabajo en campo
                df_recortado['Se realizó (Si/No)'] = "Pendiente"
                df_recortado['Justificación / Comentarios'] = ""
                
                # Limpiar formatos de datos de texto
                df_recortado = df_recortado.fillna("")
                for col in df_recortado.columns:
                    df_recortado[col] = df_recortado[col].astype(str).str.strip()

                # Guardar el archivo limpio unificado en el servidor central (Crea el CSV por primera vez)
                df_recortado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
                
                st.success("¡Lista semanal filtrada y guardada con éxito!")
                st.rerun()
            except Exception as e:
                st.error(f"Error crítico al procesar la estructura del archivo: {e}")
            
    if os.path.exists(ARCHIVO_SERVIDOR):
        st.write("---")
        if st.button("🗑️ Borrar Semana Actual", color="red", use_container_width=True):
            os.remove(ARCHIVO_SERVIDOR)
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (COMPARTIDO POR TODO EL EQUIPO)
# -------------------------------------------------------------------------
# BLINDAJE CRÍTICO: Solo intentamos leer el archivo si REALMENTE existe en el servidor
if os.path.exists(ARCHIVO_SERVIDOR):
    try:
        # Cargar los datos guardados en el disco del servidor
        df_actual = pd.read_csv(ARCHIVO_SERVIDOR, encoding='utf-8-sig')
        df_actual = df_actual.fillna("")
        
        # Identificar la semana actual
        num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
        
        st.info("📱 **Muchachos:** Busquen su fila, cambien el estado en 'Se realizó (Si/No)', escriban el motivo si aplica y denle al botón **Guardar Cambios**.")

        # Renderizado de la tabla interactiva
        df_editado = st.data_editor(
            df_actual,
            column_config={
                "Se realizó (Si/No)": st.column_config.SelectboxColumn(
                    "Se realizó (Si/No)",
                    options=["Pendiente", "Si", "No", "Cancelado"],
                    required=True,
                    width="medium"
                ),
                "Justificación / Comentarios": st.column_config.TextColumn(
                    "Justificación / Comentarios",
                    placeholder="Escribir motivo si no se realizó...",
                    width="large"
                )
            },
            # Tus compañeros sólo pueden modificar los estados y comentarios
            disabled=[col for col in df_actual.columns if col not in ['Se realizó (Si/No)', 'Justificación / Comentarios']],
            hide_index=True,
            use_container_width=True
        )

        st.write("---")

        # BOTÓN GUARDAR GENERAL
        if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
            df_editado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados en el servidor central!")
            st.balloons()

        # Progreso del grupo y exportación
        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado['Se realizó (Si/No)'] == 'Si'])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
            
            # Control por si la tabla llegara a estar en 0 filas
            if total > 0:
                porcentaje = hechos / total
            else:
                porcentaje = 0.0
            st.progress(porcentaje)
            
        with col2:
            csv_final = df_editado.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Reporte Final para Jefatura",
                data=csv_final,
                file_name=f"Reporte_Final_Semana_{num_semana}.csv",
                mime="text/csv",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Error al leer la base de datos actual: {e}")
else:
    # Si el archivo no existe en el servidor, mostramos este aviso limpio en lugar del error rojo
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra lateral izquierda para activar la tabla de todo el equipo.")
