import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Configuración de página limpia y minimalista
st.set_page_config(
    page_title="Unificador de Comentarios",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS personalizados para una interfaz elegante
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; hgight: 3em; }
    .upload-box { border: 2px dashed #cccccc; padding: 20px; border-radius: 10px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

st.title("📊 Unificador y Limpiador de Comentarios para Power BI")
st.write("Sube todos los archivos Excel que necesites (sin límite). El sistema detectará automáticamente las columnas y los combinará.")

# Palabras clave para mapeo inteligente
FECHA_KEYWORDS = ['date', 'fecha', 'created', 'time', 'timestamp', 'publicado', 'enviado']
TEXTO_KEYWORDS = ['comment', 'comentario', 'text', 'body', 'mensaje', 'review', 'contenido']

def detectar_columna(df_columns, keywords):
    for col in df_columns:
        if any(kw in str(col).lower() for kw in keywords):
            return col
    return None

# Contenedor principal de carga
st.subheader("1. Cargar Archivos Excel")
uploaded_files = st.file_uploader(
    "Arrastra y suelta aquí tus archivos exportados (.xlsx, .xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True,
    key="file_uploader"
)

if uploaded_files:
    st.write("---")
    st.subheader("2. Procesamiento de Archivos")
    
    dfs_procesados = []
    
    # Barra de progreso dinámica para no congelar la pantalla con muchos archivos
    progreso_barra = st.progress(0)
    status_text = st.empty()
    
    total_archivos = len(uploaded_files)
    
    for idx, file in enumerate(uploaded_files):
        # Actualizar estado de la barra
        porcentaje = int((idx + 1) / total_archivos * 100)
        progreso_barra.progress(porcentaje)
        status_text.text(f"Procesando archivo {idx + 1} de {total_archivos}: {file.name}")
        
        try:
            # Leer el archivo optimizando el uso de memoria
            df = pd.read_excel(file)
            
            # Detectar plataforma por nombre de archivo para armar el ID personalizado
            nombre_archivo = file.name.lower()
            plataforma = "rrss"
            for p in ["instagram", "facebook", "youtube", "tiktok", "twitter", "linkedin", "threads"]:
                if p in nombre_archivo:
                    plataforma = p
                    break
            
            # Mapeo automático de columnas
            col_fecha = detectar_columna(df.columns, FECHA_KEYWORDS)
            col_texto = detectar_columna(df.columns, TEXTO_KEYWORDS)
            
            if not col_fecha or not col_texto:
                st.error(f"❌ '{file.name}' omitido: No se halló columna de Fecha o Comentario.")
                continue
            
            # Estructurar la información de salida obligatoria
            df_limpio = pd.DataFrame()
            
            # ID Único: combinación de plataforma, marca de tiempo del archivo e índice de fila
            timestamp_id = int(datetime.now().timestamp())
            df_limpio['id'] = [f"{plataforma}_{timestamp_id}_{i+1}" for i in range(len(df))]
            
            # Estandarizar Fechas uniformemente a YYYY-MM-DD
            df_limpio['date'] = pd.to_datetime(df[col_fecha], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # Texto limpio de comentarios
            df_limpio['comment'] = df[col_texto].astype(str).str.strip()
            
            # Sentimiento (columna vacía solicitada para Power BI)
            df_limpio['sentiment'] = ""
            
            # Limpieza: Remover filas donde la fecha o el comentario falten por completo
            df_limpio.dropna(subset=['date', 'comment'], inplace=True)
            df_limpio = df_limpio[df_limpio['comment'] != "nan"] # Filtrar conversiones de null a string
            
            dfs_procesados.append(df_limpio)
            
        except Exception as e:
            st.error(f"💥 Error crítico al abrir {file.name}: {str(e)}")
            
    # Finalizar barra de progreso
    status_text.text("¡Procesamiento completo!")
    progreso_barra.empty()

    if dfs_procesados:
        # Combinación masiva eficiente
        df_final = pd.concat(dfs_procesados, ignore_index=True)
        
        st.write("---")
        st.subheader("3. Resultado de la Combinación Unificada")
        
        # Métricas rápidas en pantalla
        col1, col2 = st.columns(2)
        col1.metric("Archivos combinados con éxito", len(dfs_procesados))
        col2.metric("Total de registros generados", f"{len(df_final):,}")
        
        # Tabla de vista previa inteligente (solo muestra las primeras 10 filas para no saturar el navegador)
        st.write("**Vista previa del set de datos final (Primeras 10 filas):**")
        st.dataframe(df_final.head(10), use_container_width=True)
        
        # Generar la descarga del Excel consolidado en memoria intermedia
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Data_PowerBI')
        
        st.write("---")
        # Botón de descarga destacado
        st.download_button(
            label="🚀 DESCARGAR EXCEL COMBINADO PARA POWER BI",
            data=output.getvalue(),
            file_name=f"reporte_sentimientos_consolidado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.warning("⚠️ No se pudo extraer información válida de ninguno de los archivos subidos.")
