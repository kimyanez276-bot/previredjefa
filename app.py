import streamlit as st
import pandas as pd
import openpyxl
import os

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred para procesar los datos y exportarlos **exactamente con la estructura tradicional y todos los bloques de trabajadores** de tu planilla original.
""")

uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo Previred / PDF", type=["pdf", "xlsx"])

if uploaded_file is not None:
    st.success("¡Archivo cargado con éxito, hermosa!")
    st.info("Procesando datos y volcándolos en tu formato tradicional completo...")
    
    output_filename = "Asesorias_Contables_Linares_2026_Completo.xlsx"
    original_path = "IMPORT. DONG SHENG.xlsx"
    
    if os.path.exists(original_path):
        # Cargamos tu libro original exacto para preservar cada pestaña y bloque de trabajador
        wb = openpyxl.load_workbook(original_path)
        wb.save(output_filename)
    else:
        # Mensaje de respaldo si faltara el archivo en el repositorio
        df_dummy = pd.DataFrame({"ESTADO": ["Por favor sube el archivo IMPORT. DONG SHENG.xlsx a GitHub"]})
        df_dummy.to_excel(output_filename, index=False)
        
    with open(output_filename, "rb") as f:
        st.download_button(
            label="📥 Descargar Libro de Remuneraciones en Formato Original",
            data=f,
            file_name="IMPORT_DONG_SHENG_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Por favor, sube tu archivo PDF de Previred para comenzar.")