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

# LISTAS DE PRIORIDAD STRICTA (Para evitar confundir números con textos)
# Se busca de izquierda a derecha. Si encuentra 'caption', se queda con ese antes de buscar 'comment'
TEXTO_PRIORIDADES = [
    'caption', 'review', 'comment text', 'texto del comentario', 
    'body', 'mensaje', 'contenido', 'text', 'comment', 'texto'
]

FECHA_PRIORIDADES = [
    'date created', 'date', 'fecha', 'created at', 'created', 'time', 
    'timestamp', 'publicado', 'published', 'post date'
]

def normalizar_nombre(txt):
    """Limpia los nombres de las columnas para poder compararlos exactamente"""
    txt = str(txt).lower().strip()
    # Quitar acentos
    txt = re.sub(r'[áäâà]', 'a', txt)
    txt = re.sub(r'[éëêè]', 'e', txt)
    txt = re.sub(r'[íïîì]', 'i', txt)
    txt = re.sub(r'[óöôò]', 'o', txt)
    txt = re.sub(r'[úüûù]', 'u', txt)
    # Dejar solo caracteres alfanuméricos y espacios simples
    txt = re.sub(r'[^a-z0-9 ]', '', txt)
    return ' '.join(txt.split())

def detectar_columna_inteligente(df_columns, lista_prioridades):
    """Busca la columna óptima basándose en el orden de prioridad estricto"""
    columnas_limpias = {normalizar_nombre(col): col for col in df_columns}
    
    for keyword in lista_prioridades:
        kw_limpia = normalizar_nombre(keyword)
        # Buscar coincidencia exacta en los nombres normalizados
        if kw_limpia in columnas_limpias:
            return columnas_limpias[kw_limpia]
            
    # Si no hay coincidencia exacta, buscar coincidencia parcial por si viene algo como "review_body"
    for keyword in lista_prioridades:
        kw_limpia = normalizar_nombre(keyword)
        for col_limpia, col_original in columnas_limpias.items():
            if kw_limpia in col_limpia:
                # Evitar falsos positivos peligrosos (como confundir 'comments' count con texto)
                if kw_limpia == 'comment' and ('count' in col_limpia or 'likes' in col_limpia or 'num' in col_limpia):
                    continue
                return col_original
    return None

def leer_excel_detectando_cabecera(file):
    """Busca en las primeras 20 filas el encabezado real que tenga fecha y texto de valor"""
    for filas_a_saltar in range(0, 20):
        try:
            df_intento = pd.read_excel(file, skiprows=filas_a_saltar)
            
            # Forzar a que la fila leída realmente tenga datos válidos y no metadatos cortos
            if len(df_intento.columns) < 2:
                continue
                
            col_fecha = detectar_columna_inteligente(df_intento.columns, FECHA_PRIORIDADES)
            col_texto = detectar_columna_inteligente(df_intento.columns, TEXTO_PRIORIDADES)
            
            # Condición estricta: Para que sea la cabecera real, 'col_texto' NO puede ser una columna numérica de recuento
            if col_fecha and col_texto:
                # Verificar si los primeros registros no son strings vacíos o metadatos de configuración
                primer_valor = str(df_intento[col_texto].iloc[0]).lower()
                if "exportcomments" in primer_valor or "timezone" in primer_valor:
                    continue
                return df_intento, col_fecha, col_texto, filas_a_saltar
        except Exception:
            continue
    return None, None, None, None

def normalizar_fecha_loca(val):
    if pd.isna(val) or str(val).strip() == "" or "timezone" in str(val).lower():
        return None
    try:
        # Extraer solo la fecha ignorando la hora
        fecha_parseada = parser.parse(str(val), dayfirst=True)
        return fecha_parseada.strftime('%Y-%m-%d')
    except Exception:
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
    
    for idx, file in enumerate(uploaded_files):
        try:
            df, col_fecha, col_texto, fila_inicio = leer_excel_detectando_cabecera(file)
            
            if df is None:
                st.warning(f"⚠️ '{file.name}' omitido. No se detectaron las columnas de fecha/texto.")
                continue
            
            # Identificar plataforma para construir el prefijo del ID
            nombre_archivo = file.name.lower()
            plataforma = "rrss"
            for p in ["instagram", "ig", "facebook", "fb", "youtube", "yt", "tiktok", "twitter", "linkedin", "tripadvisor"]:
                if p in nombre_archivo:
                    plataforma = "instagram" if p == "ig" else ("youtube" if p == "yt" else ("facebook" if p == "fb" else p))
                    break
            
            df_limpio = pd.DataFrame()
            
            # Aplicar filtro estricto de texto para comentarios
            df_limpio['comment'] = df[col_texto].astype(str).str.strip()
            df_limpio['date'] = df[col_fecha].apply(normalizar_fecha_loca)
            df_limpio['sentiment'] = ""
            
            # Limpieza drástica de filas basura
            df_limpio = df_limpio[df_limpio['comment'] != "nan"]
            df_limpio = df_limpio[df_limpio['comment'] != ""]
            df_limpio = df_limpio[~df_limpio['comment'].str.contains("america/", case=False, na=False)]
            df_limpio = df_limpio[~df_limpio['comment'].str.contains("exported by", case=False, na=False)]
            
            # Quitar filas donde la fecha falló por completo tras la limpieza
            df_limpio.dropna(subset=['date', 'comment'], inplace=True)
            
            # Generar el ID Único limpio basado en el número total de comentarios reales válidos
            timestamp_id = int(datetime.now().timestamp())
            df_limpio['id'] = [f"{plataforma}_{timestamp_id}_{i+1}" for i in range(len(df_limpio))]
            
            # Reordenar columnas exactas solicitadas
            df_limpio = df_limpio[['id', 'date', 'comment', 'sentiment']]
            
            if not df_limpio.empty:
                dfs_procesados.append(df_limpio)
                st.success(f"✔️ '{file.name}' unificado con éxito ({len(df_limpio)} filas reales).")
            else:
                st.warning(f"⚠️ '{file.name}' se quedó sin registros válidos tras la depuración.")
            
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
