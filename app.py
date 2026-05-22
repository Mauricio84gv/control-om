import streamlit as st
import pandas as pd

# Configuración de pantalla ancha para comodidad en campo
st.set_page_config(page_title="Control O&M Semanal", layout="wide")

st.title("📊 RIR/RDA Status Tracker - O&M")
st.write("Panel centralizado para actualización de trabajos en el campo.")

# Mantener los datos en la memoria del servidor mientras la app esté abierta
if "df_semana" not in st.session_state:
    st.session_state.df_semana = None

# -------------------------------------------------------------------------
# SÓLO TU PANEL (BARRA LATERAL) - PARA CARGAR LA NUEVA SEMANA
# -------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Administrador")
    st.write("Carga el archivo Excel del lunes aquí.")
    
    nuevo_excel = st.file_uploader("Subir Excel de la semana", type=["xlsx"])
    
    if nuevo_excel is not None:
        if st.button("🔄 Inicializar Semana", use_container_width=True):
            # Leer el archivo respetando la estructura exacta
            df_cargado = pd.read_excel(nuevo_excel)
            df_cargado = df_cargado.fillna("")
            
            # Verificar si existen tus columnas de control, si no, se aseguran
            if 'Se realizó (Si/No)' not in df_cargado.columns:
                df_cargado['Se realizó (Si/No)'] = "Pendiente"
            if 'Si no se realizó detallar el por qué.' not in df_cargado.columns:
                df_cargado['Si no se realizó detallar el por qué.'] = ""
                
            st.session_state.df_semana = df_cargado
            st.success("¡Lista semanal activada!")
            st.rerun()

# -------------------------------------------------------------------------
# PANEL DE TRABAJO (TÚ Y TUS COMPAÑEROS)
# -------------------------------------------------------------------------
if st.session_state.df_semana is not None:
    df_actual = st.session_state.df_semana
    
    # Detectar el número de semana del archivo
    num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
    st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
    
    st.info("📱 **Muchachos:** Busquen su fila, cambien el estado en 'Se realizó (Si/No)', escriban el motivo si aplica y denle al botón **Guardar Cambios**.")

    # Renderizado de tu tabla
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
            # Formatear números para que no muestren comas innecesarias (ej: 8255647)
            "RIR": st.column_config.NumberColumn("RIR", format="%d"),
            "RDA": st.column_config.NumberColumn("RDA", format="%d"),
            "SEMANA": st.column_config.NumberColumn("SEMANA", format="%d")
        },
        # Tus compañeros sólo pueden tocar las últimas dos columnas de control
        disabled=[col for col in df_actual.columns if col not in ['Se realizó (Si/No)', 'Si no se realizó detallar el por qué.']],
        hide_index=True,
        use_container_width=True
    )

    st.write("---")

    # Botón para salvar las actualizaciones en vivo
    if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
        st.session_state.df_semana = df_editado
        st.success("✅ Cambios guardados. Tus compañeros ya los pueden ver reflejados.")
        st.balloons()

    # Progress bar e informe final
    st.write("---")
    total = len(df_editado)
    hechos = len(df_editado[df_editado['Se realizó (Si/No)'] == 'Si'])
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.write(f"**Avance del Equipo:** {hechos} de {total} completados.")
        st.progress(hechos / total if total > 0 else 0)
    with col2:
        # Descarga final del archivo corregido
        csv_final = df_editado.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Descargar Reporte Final para Jefatura",
            data=csv_final,
            file_name=f"Reporte_Final_Semana_{num_semana}.csv",
            mime="text/csv",
            use_container_width=True
        )
else:
    st.warning("⏳ Esperando que el administrador cargue el archivo Excel en la barra lateral izquierda para activar la tabla.")