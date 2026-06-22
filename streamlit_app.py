import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Extractor de Facturas Ada", layout="wide")

st.title("🧾 Extractor Automático de Facturas")
st.subheader("Subí tus tickets y armá la base de datos al toque")

# 1. Función para simular o procesar la extracción
def procesar_factura(archivo_subido):
    # TODO: Acá adentro conectaríamos la lógica de EasyOCR o la API de Gemini
    # Por ahora, simulamos la devolución de datos asegurando formato de TEXTO
    datos_extraidos = {
        "Fecha de compra": "20-05-2026",
        "Razón Social": "S.A. IMPORTADORA Y EXPORTADORA DE LA PATAGONIA",
        "Tipo": "B",
        # Guardamos el número explícitamente como string para no perder los ceros
        "Número": "0564801509107", 
        "Monto total": 60073.86,
        "Observaciones": "Compra de supermercado (Carnes, fiambres, verduras, almacén)"
    }
    return datos_extraidos

# 2. Componente para subir múltiples archivos (Fotos o PDFs)
archivos = st.file_uploader("Arrastrá tus facturas acá", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)

if archivos:
    lista_facturas = []
    
    with st.spinner("Procesando facturas... ⏳"):
        for archivo in archivos:
            datos = procesar_factura(archivo)
            lista_facturas.append(datos)
            
    # 3. Crear el DataFrame
    df = pd.DataFrame(lista_facturas)
    
    # IMPORTANTE: Forzamos a que las columnas de identificación sean tratadas como texto
    df["Tipo"] = df["Tipo"].astype(str)
    df["Número"] = df["Número"].astype(str)
    
    st.success("¡Procesamiento completado!")
    
    # 4. Mostrar la tabla en la interfaz web de Streamlit
    # st.dataframe() renderiza los strings con ceros a la izquierda correctamente
    st.dataframe(df, use_container_width=True)
    
    # 5. Botón de descarga a Excel manteniendo el formato de texto
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Facturas')
        
        # Truco para openpyxl: Forzar a Excel a que interprete la columna "Número" como texto plano
        workbook  = writer.book
        worksheet = writer.sheets['Facturas']
        # Asumiendo que "Número" es la cuarta columna (columna D)
        for cell in worksheet['D']:
            cell.number_format = '@' # Formato de texto en Excel
            
    st.download_button(
        label="📥 Descargar Excel con ceros protegidos",
        data=buffer.getvalue(),
        file_name="facturas_procesadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
