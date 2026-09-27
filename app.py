import io
import re
from typing import Optional, Tuple
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred (`CtrlPdf.pdf`) y tu planilla corporativa. El sistema extraerá correctamente los sueldos imponibles por RUT y rellenará con precisión la fila de **Agosto** en la pestaña **SUELDOS 2026**.
""")

SHEET_NAME = "SUELDOS 2026"
TARGET_MONTH = "AGOSTO"

RUT_RE = re.compile(r"(?<!\d)(\d{1,2}(?:\.\d{3}){2}-[\dkK]|\d{7,8}-[\dkK])")

def normalize_rut(value) -> str:
    if value is None:
        return ""
    raw = re.sub(r"[^0-9kK]", "", str(value))
    if len(raw) < 2:
        return ""
    return f"{raw[:-1]}-{raw[-1].upper()}"

def parse_clp(value) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(round(value))
    text = str(value).strip().replace("$", "").replace(" ", "")
    if not text or "%" in text:
        return None
    text = text.replace(".", "")
    if text.isdigit():
        return int(text)
    return None

def extract_pdf_data(pdf_bytes: bytes) -> pd.DataFrame:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    full_text = ""
    for page in reader.pages:
        full_text += (page.extract_text() or "") + "\n"
    
    lines = full_text.split("\n")
    records = {}
    
    # Recorremos línea por línea buscando patrones de RUT y montos válidos de remuneración
    for i, line in enumerate(lines):
        rut_match = RUT_RE.search(line)
        if rut_match:
            rut_raw = rut_match.group(1)
            rut = normalize_rut(rut_raw)
            if not rut:
                continue
            
            # Buscamos en las líneas cercanas un monto imponible válido (ej entre 300.000 y 10.000.000)
            window = " ".join(lines[max(0, i-2):min(len(lines), i+3)])
            tokens = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d{6,8}\b", window)
            
            amounts = []
            for t in tokens:
                val = parse_clp(t)
                if val and 300000 <= val <= 15000000: # Rango normal de sueldos imponibles
                    amounts.append(val)
            
            if amounts:
                # Tomamos el primer monto coherente como sueldo imponible
                records[rut] = {
                    "rut": rut,
                    "sueldo_imponible": amounts[0]
                }

    df = pd.DataFrame(list(records.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel.")
    
    ws = wb[SHEET_NAME]
    
    # Recorremos el Excel buscando los bloques de cada trabajador por RUT
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row, 2).value
        norm_cell = normalize_rut(cell_val)
        if norm_cell:
            for _, rec in df.iterrows():
                if normalize_rut(rec["rut"]) == norm_cell:
                    # Encontramos al trabajador. Buscamos la fila de AGOSTO en su bloque hacia abajo
                    for r_sub in range(row, row + 15):
                        mes_val = str(ws.cell(r_sub, 1).value or "").strip().upper()
                        if "AGOSTO" in mes_val:
                            # Columna 2 (B) es el Sueldo Bruto/Imponible en la plantilla de sueldos
                            ws.cell(r_sub, 2).value = rec["sueldo_imponible"]
                            break
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Interfaz Streamlit
pdf_file = st.file_uploader("1. Sube tu PDF de Previred (`CtrlPdf.pdf`)", type=["pdf"])
template_file = st.file_uploader("2. Sube tu plantilla Excel (`IMPORT. DONG SHENG.xlsx`)", type=["xlsx"])

if pdf_file and template_file:
    st.success("¡Archivos cargados correctamente, hermosa!")
    
    try:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        st.subheader("Datos extraídos del Previred:")
        st.dataframe(df_extracted, use_container_width=True)
        
        if st.button("🚀 Rellenar Planilla con Montos Exactos", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla rellenada y corregida con éxito!")
            
            st.download_button(
                label="📥 Descargar Libro Corregido",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Agosto_Corregido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")