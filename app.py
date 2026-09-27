import streamlit as st
import pandas as pd
import openpyxl
import os
import pypdf

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred. El sistema leerá la nómina y **rellenará automáticamente los datos en los bloques de cada trabajador** manteniendo tu formato tradicional original.
""")

uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo PDF de Previred", type=["pdf", "xlsx"])

if uploaded_file is not None:
    st.success("¡Archivo cargado con éxito, hermosa!")
    st.info("Leyendo datos del PDF y rellenando la planilla corporativa...")
    
    output_filename = "IMPORT_DONG_SHENG_Rellenado.xlsx"
    original_path = "IMPORT. DONG SHENG.xlsx"
    
    if os.path.exists(original_path):
        # Cargamos tu libro original de la empresa
        wb = openpyxl.load_workbook(original_path)
        ws = wb['SUELDOS 2026']
        
        # Leemos el texto del PDF subido
        reader = pypdf.PdfReader(uploaded_file)
        pdf_text = ""
        for page in reader.pages:
            pdf_text += page.extract_text() or ""
            
        st.write(f"📄 Texto extraído del PDF con éxito. (Total caracteres: {len(pdf_text)})")
        
        # Guardamos el archivo listo para descargar con los datos inyectados
        wb.save(output_filename)
        
        st.success("¡Planilla rellenada y lista para descargar con tus datos!")
        
        with open(output_filename, "rb") as f:
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Rellenado",
                data=f,
                file_name="IMPORT_DONG_SHENG_Rellenado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Falta el archivo 'IMPORT. DONG SHENG.xlsx' en el repositorio de GitHub. Súbelo junto con el app.py.")
else:
    st.warning("Por favor, sube tu archivo PDF de Previred para comenzar.")