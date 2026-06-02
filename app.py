import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="conector de comentarios", layout="wide")
st.title("📊 limpiador y combinador de comentarios")
st.write("Sube los archivos Excel exportados de tus redes sociales para unificarlos.")

# Palabras clave para detectar columnas dinámicamente
FECHA_KEYWORDS = ['date', 'fecha', 'created', 'time', 'timestamp', 'publicado']
TEXTO_KEYWORDS = ['comment', 'comentario', 'text', 'body', 'mensaje', 'review']

def detectar_columna(df_columns, keywords):
    for col in df_columns:
        if any(kw in str(col).lower() for kw in keywords):
            return col
    return None

uploaded_files = st.file_uploader("Selecciona uno o varios archivos Excel", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    dfs_procesados = []
    
    for file in uploaded_files:
        try:
            # Leer el excel
            df = pd.read_excel(file)
            
            # Identificar la plataforma según el nombre del archivo para el ID
            nombre_archivo = file.name.lower()
            plataforma = "social"
            for p in ["instagram", "facebook", "youtube", "tiktok", "twitter", "linkedin"]:
                if p in nombre_archivo:
                    plataforma = p
                    break
            
            # Detectar columnas de fecha y comentario
            col_fecha = detectar_columna(df.columns, FECHA_KEYWORDS)
            col_texto = detectar_columna(df.columns, TEXTO_KEYWORDS)
            
            if not col_fecha or not col_texto:
                st.warning(f"⚠️ No se pudieron mapear las columnas en '{file.name}'. Asegúrate de que existan columnas de fecha y texto.")
                continue
                
            # Crear el dataframe limpio con las columnas solicitadas
            df_limpio = pd.DataFrame()
            
            # 1. Asignar ID Único (Plataforma + Índice)
            df_limpio['id'] = [f"{plataforma}_{i+1}" for i in range(len(df))]
            
            # 2. Estandarizar Fecha a formato YYYY-MM-DD
            df_limpio['date'] = pd.to_datetime(df[col_fecha], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 3. Texto del comentario
            df_limpio['comment'] = df[col_texto].astype(str).str.strip()
            
            # 4. Sentimiento (Vacia para que la llenes/analices en Power BI)
            df_limpio['sentiment'] = ""
            
            # Eliminar filas donde la fecha o el comentario queden completamente nulos
            df_limpio.dropna(subset=['date', 'comment'], inplace=True)
            
            dfs_procesados.append(df_limpio)
            st.success(f"✅ Procesado con éxito: {file.name} ({len(df_limpio)} filas)")
            
        except Exception as e:
            st.error(f"❌ Error al procesar {file.name}: {str(e)}")

    if dfs_procesados:
        # Combinar todos los dataframes
        df_final = pd.concat(dfs_procesados, ignore_index=True)
        
        st.write("---")
        st.subheader("📋 Vista previa de los datos combinados")
        st.dataframe(df_final.head(10), use_container_width=True)
        st.info(f"Total de comentarios unificados: {len(df_final)}")
        
        # Generar descarga en Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Comentarios_Limpios')
        
        st.download_button(
            label="📥 Descargar Excel Combinado para Power BI",
            data=output.getvalue(),
            file_name=f"comentarios_unificados_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
