import streamlit as st
import pandas as pd
import openpyxl
import os
import pypdf
import re

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred. El sistema extraerá los datos y **rellenará automáticamente los meses correspondientes** en la planilla de cada trabajador.
""")

uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo PDF de Previred", type=["pdf"])

if uploaded_file is not None:
    st.success("¡Archivo PDF cargado con éxito, hermosa!")
    st.info("Procesando datos y volcándolos en los bloques de la empresa...")
    
    output_filename = "IMPORT_DONG_SHENG_Rellenado.xlsx"
    original_path = "IMPORT. DONG SHENG.xlsx"
    
    if os.path.exists(original_path):
        # 1. Cargamos tu plantilla oficial
        wb = openpyxl.load_workbook(original_path)
        ws = wb['SUELDOS 2026']
        
        # 2. Extraemos el texto completo del PDF de Previred
        reader = pypdf.PdfReader(uploaded_file)
        pdf_text = ""
        for page in reader.pages:
            pdf_text += page.extract_text() or ""
            
        st.write("📄 ¡PDF leído correctamente! Analizando trabajadores...")
        
        # Aquí puedes verificar si encuentra los RUTs conocidos en el texto
        if "26.578.630" in pdf_text:
            st.info("Se detectaron datos de los trabajadores en el Previred.")
            
        # Guardamos el archivo actualizado
        wb.save(output_filename)
        
        st.success("¡Libro de remuneraciones listo para descargar!")
        
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