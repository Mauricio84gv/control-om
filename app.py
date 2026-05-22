import streamlit as st
import pandas as pd
import os

# 1. Configuración de la pantalla
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

# Nombre del archivo físico que compartirán todos en el servidor
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
            # Leer el archivo nuevo
            df_cargado = pd.read_excel(nuevo_excel)
            df_cargado = df_cargado.fillna("")
            
            # Forzar la creación de las columnas si vienen vacías
            if 'Se realizó (Si/No)' not in df_cargado.columns:
                df_cargado['Se realizó (Si/No)'] = "Pendiente"
            if 'Si no se realizó detallar el por qué.' not in df_cargado.columns:
                df_cargado['Si no se realizó detallar el por qué.'] = ""
                
            # GUARDADO REAL EN EL SERVIDOR: Esto hace que guarde el archivo en el disco duro del hosting
            df_cargado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
            
            st.success("¡Lista semanal guardada en el servidor para todo el equipo!")
            st.rerun()
            
    # Botón de reinicio por si quieres borrar la semana actual
    if os.path.exists(ARCHIVO_SERVIDOR):
        if st.button("🗑️ Borrar Semana Actual", color="red"):
            os.remove(ARCHIVO_SERVIDOR)
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (COMPARTIDO POR TODO EL EQUIPO)
# -------------------------------------------------------------------------
# Verificamos si el archivo físico existe en el servidor
if os.path.exists(ARCHIVO_SERVIDOR):
    # Todos leen exactamente el mismo archivo del disco duro
    df_actual = pd.read_csv(ARCHIVO_SERVIDOR, encoding='utf-8-sig')
    df_actual = df_actual.fillna("")
    
    # Detectar el número de semana
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
            "Si no se realizó detallar el por qué.": st.column_config.TextColumn(
                "Si no se realizó detallar el por qué.",
                placeholder="Escribir justificación aquí...",
                width="large"
            ),
            "RIR": st.column_config.NumberColumn("RIR", format="%d"),
            "RDA": st.column_config.NumberColumn("RDA", format="%d"),
            "SEMANA": st.column_config.NumberColumn("SEMANA", format="%d")
        },
        disabled=[col for col in df_actual.columns if col not in ['Se realizó (Si/No)', 'Si no se realizó detallar el por qué.']],
        hide_index=True,
        use_container_width=True
    )

    st.write("---")

    # BOTÓN GUARDAR REAL: Sobreescribe el archivo compartido en el servidor
    if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
        df_editado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
        st.success("✅ ¡Cambios guardados en el servidor central! Tus compañeros los verán al refrescar la página.")
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
