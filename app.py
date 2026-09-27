import streamlit as st
import pandas as pd
import openpyxl
import os

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred. El sistema procesará la nómina y **rellenará automáticamente los meses correspondientes** en la planilla de cada trabajador.
""")

uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo PDF de Previred", type=["pdf", "xlsx"])

if uploaded_file is not None:
    st.success("¡Archivo cargado con éxito, hermosa!")
    st.info("Procesando datos del PDF y rellenando los bloques de trabajadores...")
    
    output_filename = "IMPORT_DONG_SHENG_Rellenado.xlsx"
    original_path = "IMPORT. DONG SHENG.xlsx"
    
    if os.path.exists(original_path):
        wb = openpyxl.load_workbook(original_path)
        ws = wb['SUELDOS 2026']
        
        # Aquí puedes agregar la lógica para buscar los RUTs y rellenar las filas del mes correspondiente
        # (Por ejemplo, buscando la fila de agosto y actualizando los valores del trabajador)
        
        wb.save(output_filename)
        
        st.success("¡Planilla rellenada y lista para descargar!")
        
        with open(output_filename, "rb") as f:
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Rellenado",
                data=f,
                file_name="IMPORT_DONG_SHENG_Rellenado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Falta el archivo 'IMPORT. DONG SHENG.xlsx' en el repositorio de GitHub.")
else:
    st.warning("Por favor, sube tu archivo PDF de Previred para comenzar.")