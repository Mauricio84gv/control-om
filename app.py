import streamlit as st
import pandas as pd
import os

# 1. Configuración de la pantalla
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

# Nombre del archivo físico compartido en el servidor
ARCHIVO_SERVIDOR = "datos_semana_activa.csv"

# Definimos exactamente las columnas que tu archivo debe tener
COLUMNAS_VALIDAS = [
    'SEMANA', 'ZONA', 'SITIO', 'RIR', 'RDA', 'NOMBRE', 'TORRERA', 
    'Se realizó (Si/No)', 'Si no se realizó detallar el por qué.'
]

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
                # Leer el archivo Excel nuevo
                df_cargado = pd.read_excel(nuevo_excel)
                
                # REGLA DE LIMPIEZA CLAVE: Asegurar nombres de columnas exactos (quitar espacios fantasmas)
                df_cargado.columns = df_cargado.columns.str.strip()
                
                # Inyectar las columnas de control si el Excel viene limpio de la oficina
                if 'Se realizó (Si/No)' not in df_cargado.columns:
                    df_cargado['Se realizó (Si/No)'] = "Pendiente"
                if 'Si no se realizó detallar el por qué.' not in df_cargado.columns:
                    df_cargado['Si no se realizó detallar el por qué.'] = ""
                
                # BLINDAJE CONTRA EL VALUEERROR: Filtramos para quedarnos SOLO con tus columnas reales
                # Esto descarta filas deformes, espacios vacíos al final y notas extrañas
                df_cargado = df_cargado[COLUMNAS_VALIDAS]
                
                # Limpiar cualquier fila que esté completamente en blanco
                df_cargado = df_cargado.dropna(subset=['SITIO'])
                df_cargado = df_cargado.fillna("")
                
                # Asegurar que los textos de las celdas estén limpios
                for col in df_cargado.columns:
                    df_cargado[col] = df_cargado[col].astype(str).str.strip()

                # GUARDADO REAL EN EL SERVIDOR
                df_cargado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
                
                st.success("¡Lista semanal filtrada y guardada con éxito!")
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar la estructura del Excel: {e}")
            
    if os.path.exists(ARCHIVO_SERVIDOR):
        st.write("---")
        if st.button("🗑️ Borrar Semana Actual", color="red", use_container_width=True):
            os.remove(ARCHIVO_SERVIDOR)
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (COMPARTIDO POR TODO EL EQUIPO)
# -------------------------------------------------------------------------
if os.path.exists(ARCHIVO_SERVIDOR):
    # Todos leen exactamente el mismo archivo limpio del disco duro
    df_actual = pd.read_csv(ARCHIVO_SERVIDOR, encoding='utf-8-sig')
    df_actual = df_actual.fillna("")
    
    # Detectar el número de semana
    num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
    st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
    
    st.info("📱 **Muchachos:** Busquen su fila, cambien el estado en 'Se realizó (Si/No)', escriban el motivo si aplica y denle al botón **Guardar Cambios**.")

    # Renderizado de la tabla interactiva sin peligro de longitudes variables
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
                placeholder="Escribir justificación aquí...",
                width="large"
            )
        },
        # Tus compañeros sólo pueden tocar las últimas dos columnas de control
        disabled=[col for col in df_actual.columns if col not in ['Se realizó (Si/No)', 'Si no se realizó detallar el por qué.']],
        hide_index=True,
        use_container_width=True
    )

    st.write("---")

    # BOTÓN GUARDAR
    if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
        df_editado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
        st.success("✅ ¡Cambios guardados en el servidor central!")
        st.balloons()

    # Progreso e informe final
    st.write("---")
    total = len(df_editado)
    hechos = len(df_editado[df_editado['Se realizó (Si/No)'] == 'Si'])
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
        st.progress(hechos / total if total > 0 else 0)
    with col2:
        csv_final = df_editado.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Descargar Reporte Final para Jefatura",
            data=csv_final,
            file_name=f"Reporte_Final_Semana_{num_semana}.csv",
            mime="text/csv",
            use_container_width=True
        )
else:
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra lateral izquierda para activar la tabla de todo el equipo.")
