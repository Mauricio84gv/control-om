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
                # 1. Leer el Excel completo de forma temporal para inspeccionarlo
                df_temp = pd.read_excel(nuevo_excel, header=None, dtype=str)
                df_temp = df_temp.fillna("")

                # 2. ESCANEO INTELIGENTE: Buscar en cuál fila de verdad están los encabezados reales
                fila_encabezado = 0
                for i, row in df_temp.iterrows():
                    valores_fila = [str(val).strip().upper() for val in row.values]
                    # Si encontramos "SITIO" o "SEMANA", encontramos la tabla real
                    if any("SITIO" in v or "SEMANA" in v for v in valores_fila):
                        fila_encabezado = i
                        break
                
                # 3. Volver a leer el archivo pero saltándose las filas muertas de arriba
                df_cargado = pd.read_excel(nuevo_excel, skiprows=fila_encabezado, dtype=str)
                df_cargado = df_cargado.fillna("")
                
                # Limpiar espacios en blanco invisibles en los encabezados nuevos
                df_cargado.columns = df_cargado.columns.str.strip()
                
                # Pasar todos los encabezados a MAYÚSCULAS para evitar fallas
                columnas_limpias = []
                for col in df_cargado.columns:
                    nombre_limpio = str(col).strip().upper()
                    if "SITIO" in nombre_limpio:
                        nombre_limpio = "SITIO"
                    elif "SEMANA" in nombre_limpio:
                        nombre_limpio = "SEMANA"
                    columnas_limpias.append(nombre_limpio)
                
                df_cargado.columns = columnas_limpias
                
                # 4. Crear de forma limpia las dos columnas que usarán tus compañeros
                col_realizo = [c for c in df_cargado.columns if "REALIZÓ" in c or "REALIZO" in c]
                col_justif = [c for c in df_cargado.columns if "POR QUÉ" in c or "POR QUE" in c or "JUSTIFICACI" in c or "COMENTARIO" in c]
                
                if col_realizo:
                    df_cargado = df_cargado.rename(columns={col_realizo[0]: "Se realizó (Si/No)"})
                else:
                    df_cargado["Se realizó (Si/No)"] = "Pendiente"
                    
                if col_justif:
                    df_cargado = df_cargado.rename(columns={col_justif[0]: "Justificación / Comentarios"})
                else:
                    df_cargado["Justificación / Comentarios"] = ""
                
                # Forzar que si el estado está vacío empiece en "Pendiente"
                df_cargado.loc[df_cargado["Se realizó (Si/No)"] == "", "Se realizó (Si/No)"] = "Pendiente"

                # Quitar filas que queden completamente vacías abajo
                if "SITIO" in df_cargado.columns:
                    df_cargado = df_cargado[df_cargado["SITIO"] != ""]

                # Limpieza de espacios fantasmas dentro de las celdas
                for col in df_cargado.columns:
                    df_cargado[col] = df_cargado[col].astype(str).str.strip()

                # Guardar el archivo final unificado en el servidor central
                df_cargado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
                
                st.success("¡Lista semanal procesada e inicializada con éxito!")
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
        # Cargar los datos compartidos
        df_actual = pd.read_csv(ARCHIVO_SERVIDOR, dtype=str, encoding='utf-8-sig')
        df_actual = df_actual.fillna("")
        
        # Obtener la semana de forma segura
        num_semana = df_actual['SEMANA'].iloc[0] if 'SEMANA' in df_actual.columns else ""
        st.subheader(f"📋 Sitios Activos - Semana {num_semana}")
        
        st.info("📱 **Muchachos:** Busquen su fila, cambien el estado en 'Se realizó (Si/No)', escriban el motivo si aplica y denle al botón **Guardar Cambios**.")

        # Renderizado flexible de la tabla interactiva
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
            disabled=[col for col in df_actual.columns if col not in ["Se realizó (Si/No)", "Justificación / Comentarios"]],
            hide_index=True,
            use_container_width=True
        )

        st.write("---")

        # BOTÓN GUARDAR GENERAL
        if st.button("💾 Guardar Mis Cambios del Día", type="primary", use_container_width=True):
            df_editado.to_csv(ARCHIVO_SERVIDOR, index=False, encoding='utf-8-sig')
            st.success("✅ ¡Cambios guardados con éxito!")
            st.balloons()

        # Barra de progreso e informe de exportación
        st.write("---")
        total = len(df_editado)
        hechos = len(df_editado[df_editado["Se realizó (Si/No)"] == "Si"])
        
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
