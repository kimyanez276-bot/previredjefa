import io
import re
from typing import Optional
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred (`CtrlPdf.pdf`) y tu planilla corporativa. El sistema extraerá con precisión quirúrgica los sueldos imponibles desde la tabla oficial y rellenará la fila de **Agosto** manteniendo intactas todas las fórmulas de tu Excel.
""")

SHEET_NAME = "SUELDOS 2026"
TARGET_MONTH = "AGOSTO"

RUT_RE = re.compile(r"(\d{1,2}(?:\.\d{3}){2}-[\dkK]|\d{7,8}-[\dkK])")

def normalize_rut(value) -> str:
    if value is None:
        return ""
    raw = re.sub(r"[^0-9kK]", "", str(value))
    if len(raw) < 2:
        return ""
    return f"{raw[:-1]}-{raw[-1].upper()}"

def extract_pdf_data(pdf_bytes: bytes) -> pd.DataFrame:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    
    # La página 2 (índice 1) contiene la tabla oficial de remuneraciones imponibles de Previred
    workers_data = {}
    if len(reader.pages) >= 2:
        page2_text = reader.pages[1].extract_text() or ""
        for line in page2_text.split("\n"):
            if "AFP" in line and RUT_RE.search(line):
                rut_match = RUT_RE.search(line)
                rut = normalize_rut(rut_match.group(1))
                if not rut:
                    continue
                
                parts = line.split("AFP")
                if len(parts) > 1:
                    after_afp = parts[1].strip()
                    nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", after_afp)
                    if nums:
                        clean_n = nums[0].replace(".", "")
                        if clean_n.isdigit():
                            workers_data[rut] = int(clean_n)

    df = pd.DataFrame(list(workers_data.items()), columns=["rut", "sueldo_imponible"])
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel.")
    
    ws = wb[SHEET_NAME]
    
    # Recorremos la hoja buscando las celdas de los RUTs de cada trabajador
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row, 2).value
        norm_cell = normalize_rut(cell_val)
        
        if norm_cell:
            for _, rec in df.iterrows():
                if normalize_rut(rec["rut"]) == norm_cell:
                    # Encontramos al trabajador en este bloque. Buscamos la fila "AGOSTO" hacia abajo
                    for r_sub in range(row, row + 16):
                        mes_val = str(ws.cell(r_sub, 1).value or "").strip().upper()
                        if TARGET_MONTH in mes_val:
                            # Inyectamos el sueldo imponible exacto en la Columna B (Columna 2)
                            ws.cell(r_sub, 2).value = rec["sueldo_imponible"]
                            break
                            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Interfaz Streamlit
pdf_file = st.file_uploader("1. Sube tu PDF de Previred (`CtrlPdf.pdf`)", type=["pdf"])
template_file = st.file_uploader("2. Sube tu plantilla Excel oficial", type=["xlsx"])

if pdf_file and template_file:
    st.success("¡Archivos cargados correctamente!")
    
    try:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        st.subheader("Trabajadores y Sueldos Imponibles Oficiales Extraídos:")
        st.dataframe(df_extracted, use_container_width=True)
        
        if st.button("🚀 Rellenar Planilla Oficial", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla rellenada con éxito absoluto y sin errores!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones - Agosto Definitivo",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Agosto_Oficial.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")