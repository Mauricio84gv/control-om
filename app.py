import streamlit as st
import pandas as pd
import os

# 1. Configuración de la pantalla (Debe ser la primera línea de todas)
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

# Nombre del archivo físico compartido en el servidor
ARCHIVO_SERVIDOR = "datos_semana_activa.csv"

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
                # Leer el archivo Excel tal cual venga, forzando todo a texto para evitar errores de formato
                df_cargado = pd.read_excel(nuevo_excel, dtype=str)
                
                # Limpiar espacios en blanco invisibles en los encabezados
                df_cargado.columns = df_cargado.columns.str.strip()
                
                # Validar la columna fundamental para que la app no se abra vacía
                if 'SITIO' not in df_cargado.columns:
                    st.error("❌ El archivo Excel debe contener obligatoriamente una columna llamada 'SITIO'.")
                    st.stop()
                
                # Eliminar filas fantasmas donde el SITIO esté completamente vacío
                df_cargado = df_cargado.dropna(subset=['SITIO'])
                
                # Crear de forma limpia las dos columnas de trabajo si no existen
                if 'Se realizó (Si/No)' not in df_cargado.columns:
                    df_cargado['Se realizó (Si/No)'] = "Pendiente"
                if 'Justificación / Comentarios' not in df_cargado.columns:
                    df_cargado['Justificación / Comentarios'] = ""
                
                # Rellenar celdas vacías para evitar errores visuales
                df_cargado = df_cargado.fillna("")
                
                # Limpiar espacios fantasmas en los textos de las celdas
                for col in df_cargado.columns:
                    df_cargado[col] = df_cargado[col].str.strip()

                # Guardar el archivo unificado en el servidor central
                df_cargado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
                
                st.success("¡Lista semanal guardada con éxito en el servidor!")
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar la estructura del archivo: {e}")
            
    if os.path.exists(ARCHIVO_SERVIDOR):
        st.write("---")
        if st.button("🗑️ Borrar Semana Actual", color="red", use_container_width=True):
            os.remove(ARCHIVO_SERVIDOR)
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (COMPARTIDO POR TODO EL EQUIPO)
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_SERVIDOR):
    try:
        # Cargar los datos guardados en el disco del servidor (Todo como texto)
        df_actual = pd.read_csv(ARCHIVO_SERVIDOR, dtype=str, encoding='utf-8-sig')
        df_actual = df_actual.fillna("")
        
        # Identificar la semana actual
        num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
        
        st.info("📱 **Muchachos:** Busquen su fila, cambien el estado en 'Se realizó (Si/No)', escriban el motivo si aplica y denle al botón **Guardar Cambios**.")

        # Renderizado de la tabla interactiva sin restricciones de columnas fijas
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
            # Solo permitimos modificar las dos columnas que usan los técnicos
            disabled=[col for col in df_actual.columns if col not in ['Se realizó (Si/No)', 'Justificación / Comentarios']],
            hide_index=True,
            use_container_width=True
        )

        st.write("---")

        # BOTÓN GUARDAR GENERAL
        if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
            df_editado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados con éxito!")
            st.balloons()

        # Progreso del grupo y exportación
        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado['Se realizó (Si/No)'] == 'Si'])
        
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
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
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra lateral izquierda para activar la tabla de todo el equipo.")
