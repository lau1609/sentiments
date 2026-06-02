import streamlit as st
import pandas as pd
from datetime import datetime
import io
import re

# Configuración de página limpia y minimalista
st.set_page_config(
    page_title="Unificador de Comentarios",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("📊 Unificador y Limpiador de Comentarios para Power BI")
st.write("Sube todos los archivos Excel que necesites (sin límite). El sistema detectará automáticamente las columnas y los combinará.")

# Diccionario expandido de palabras clave (en minúsculas y limpias)
FECHA_KEYWORDS = [
    'date', 'fecha', 'created', 'time', 'timestamp', 'publicado', 
    'enviado', 'at', 'post date', 'comment date'
]
TEXTO_KEYWORDS = [
    'comment', 'comentario', 'text', 'body', 'mensaje', 'review', 
    'contenido', 'comment text', 'texto', 'caption', 'description'
]

def limpiar_texto(texto):
    """Normaliza el texto para facilitar la búsqueda de coincidencias"""
    texto = str(texto).lower().strip()
    # Quitar acentos básicos
    texto = re.sub(r'[áäâà]', 'a', texto)
    texto = re.sub(r'[éëêè]', 'e', texto)
    texto = re.sub(r'[íïîì]', 'i', texto)
    texto = re.sub(r'[óöôò]', 'o', texto)
    texto = re.sub(r'[úüûù]', 'u', texto)
    # Quitar caracteres especiales como guiones bajos o espacios extra
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

def detectar_columna(df_columns, keywords):
    # Intentar coincidencia exacta o contenida en el texto normalizado
    for col in df_columns:
        col_limpia = limpiar_texto(col)
        for kw in keywords:
            kw_limpia = limpiar_texto(kw)
            if kw_limpia in col_limpia:
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
    progreso_barra = st.progress(0)
    status_text = st.empty()
    total_archivos = len(uploaded_files)
    
    for idx, file in enumerate(uploaded_files):
        porcentaje = int((idx + 1) / total_archivos * 100)
        progreso_barra.progress(porcentaje)
        status_text.text(f"Procesando archivo {idx + 1} de {total_archivos}: {file.name}")
        
        try:
            df = pd.read_excel(file)
            
            # Detectar plataforma por nombre de archivo
            nombre_archivo = file.name.lower()
            plataforma = "rrss"
            for p in ["instagram", "ig", "facebook", "fb", "youtube", "yt", "tiktok", "twitter", "linkedin"]:
                if p in nombre_archivo:
                    plataforma = "instagram" if p == "ig" else ("youtube" if p == "yt" else ("facebook" if p == "fb" else p))
                    break
            
            # Mapeo inteligente con limpieza previa
            col_fecha = detectar_columna(df.columns, FECHA_KEYWORDS)
            col_texto = detectar_columna(df.columns, TEXTO_KEYWORDS)
            
            # Si falla la detección, te mostramos qué columnas traía el archivo para auditarlo
            if not col_fecha or not col_texto:
                st.error(f"⚠️ '{file.name}' omitido. Columnas encontradas: `{list(df.columns)}`. Asegúrate de que tenga campos de fecha y texto.")
                continue
            
            # Estructurar la información de salida obligatoria (Aquí se renombran las columnas)
            df_limpio = pd.DataFrame()
            
            timestamp_id = int(datetime.now().timestamp())
            df_limpio['id'] = [f"{plataforma}_{timestamp_id}_{i+1}" for i in range(len(df))]
            
            # Estandarizar Fechas uniformemente a YYYY-MM-DD
            df_limpio['date'] = pd.to_datetime(df[col_fecha], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # Texto limpio de comentarios
            df_limpio['comment'] = df[col_texto].astype(str).str.strip()
            
            # Sentimiento (columna vacía solicitada)
            df_limpio['sentiment'] = ""
            
            # Limpieza básica
            df_limpio.dropna(subset=['date', 'comment'], inplace=True)
            df_limpio = df_limpio[df_limpio['comment'] != "nan"]
            
            dfs_procesados.append(df_limpio)
            
        except Exception as e:
            st.error(f"💥 Error crítico al abrir {file.name}: {str(e)}")
            
    status_text.text("¡Procesamiento completo!")
    progreso_barra.empty()

    if dfs_procesados:
        df_final = pd.concat(dfs_procesados, ignore_index=True)
        
        st.write("---")
        st.subheader("3. Resultado de la Combinación Unificada")
        
        col1, col2 = st.columns(2)
        col1.metric("Archivos combinados con éxito", len(dfs_procesados))
        col2.metric("Total de registros generados", f"{len(df_final):,}")
        
        st.write("**Vista previa del set de datos final (Tus nuevas 4 columnas estandarizadas):**")
        st.dataframe(df_final.head(10), use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Data_PowerBI')
        
        st.write("---")
        st.download_button(
            label="🚀 DESCARGAR EXCEL COMBINADO PARA POWER BI",
            data=output.getvalue(),
            file_name=f"reporte_sentimientos_consolidado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
