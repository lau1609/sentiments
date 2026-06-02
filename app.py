import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil import parser
import io
import re

st.set_page_config(
    page_title="Unificador de Comentarios",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("📊 Unificador y Limpiador de Comentarios para Power BI")
st.write("Sube todos los archivos Excel que necesites (sin límite).")

# Palabras clave ultra-expandidas basados en Export Comments y TripAdvisor
FECHA_KEYWORDS = ['date', 'fecha', 'created', 'time', 'timestamp', 'publicado', 'enviado', 'at', 'published']
TEXTO_KEYWORDS = ['comment', 'comentario', 'text', 'body', 'mensaje', 'review', 'contenido', 'texto', 'caption']

def limpiar_texto(texto):
    texto = str(texto).lower().strip()
    texto = re.sub(r'[áäâà]', 'a', texto)
    texto = re.sub(r'[éëêè]', 'e', texto)
    texto = re.sub(r'[íïîì]', 'i', texto)
    texto = re.sub(r'[óöôò]', 'o', texto)
    texto = re.sub(r'[úüûù]', 'u', texto)
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

def detectar_columna(df_columns, keywords):
    for col in df_columns:
        col_limpia = limpiar_texto(col)
        for kw in keywords:
            if limpiar_texto(kw) in col_limpia:
                return col
    return None

def leer_excel_detectando_cabecera(file):
    # Aumentamos el escaneo hasta 30 filas por si Export Comments metió más espacio
    for filas_a_saltar in range(0, 30):
        try:
            df_intento = pd.read_excel(file, skiprows=filas_a_saltar)
            col_fecha = detectar_columna(df_intento.columns, FECHA_KEYWORDS)
            col_texto = detectar_columna(df_intento.columns, TEXTO_KEYWORDS)
            
            if col_fecha and col_texto:
                return df_intento, col_fecha, col_texto, filas_a_saltar
        except Exception:
            continue
    return None, None, None, None

def normalizar_fecha_loca(val):
    if pd.isna(val) or str(val).strip() == "":
        return datetime.today().strftime('%Y-%m-%d')
    try:
        fecha_parseada = parser.parse(str(val), dayfirst=True)
        return fecha_parseada.strftime('%Y-%m-%d')
    except Exception:
        return datetime.today().strftime('%Y-%m-%d')

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
    total_archivos = len(uploaded_files)
    
    for idx, file in enumerate(uploaded_files):
        try:
            df, col_fecha, col_texto, fila_inicio = leer_excel_detectando_cabecera(file)
            
            if df is None:
                st.warning(f"⚠️ '{file.name}' omitido. No se detectaron las columnas de fecha/texto de manera automática.")
                
                # MODO DIAGNÓSTICO: Leemos las primeras filas sin saltar nada para enseñarte qué hay
                df_diag = pd.read_excel(file, nrows=10)
                with st.expander(f"🔍 Ver estructura interna de diagnóstico para: {file.name}"):
                    st.write("Así se ven las primeras 10 filas de tu archivo tal como las recibe el servidor:")
                    st.dataframe(df_diag.dropna(how='all').head(10))
                    st.write("Columnas detectadas en la fila 0 (sin saltar nada):", list(df_diag.columns))
                continue
            
            # Identificar plataforma
            nombre_archivo = file.name.lower()
            plataforma = "rrss"
            for p in ["instagram", "ig", "facebook", "fb", "youtube", "yt", "tiktok", "twitter", "linkedin", "tripadvisor"]:
                if p in nombre_archivo:
                    plataforma = "instagram" if p == "ig" else ("youtube" if p == "yt" else ("facebook" if p == "fb" else p))
                    break
            
            df_limpio = pd.DataFrame()
            timestamp_id = int(datetime.now().timestamp())
            df_limpio['id'] = [f"{plataforma}_{timestamp_id}_{i+1}" for i in range(len(df))]
            
            # Aplicar conversión de fecha inteligente
            df_limpio['date'] = df[col_fecha].apply(normalizar_fecha_loca)
            
            # Texto de comentarios
            df_limpio['comment'] = df[col_texto].astype(str).str.strip()
            df_limpio['sentiment'] = ""
            
            # Limpieza básica
            df_limpio = df_limpio[df_limpio['comment'] != "nan"]
            df_limpio = df_limpio[df_limpio['comment'] != ""]
            df_limpio.dropna(subset=['comment'], inplace=True)
            
            if not df_limpio.empty:
                dfs_procesados.append(df_limpio)
                st.success(f"✔️ '{file.name}' procesado con éxito (Inició en fila {fila_inicio + 1}, {len(df_limpio)} filas).")
            
        except Exception as e:
            st.error(f"💥 Error crítico al abrir {file.name}: {str(e)}")

    if dfs_procesados:
        df_final = pd.concat(dfs_procesados, ignore_index=True)
        st.write("---")
        st.subheader("3. Resultado de la Combinación Unificada")
        
        col1, col2 = st.columns(2)
        col1.metric("Archivos combinados con éxito", len(dfs_procesados))
        col2.metric("Total de registros generados", f"{len(df_final):,}")
        
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
