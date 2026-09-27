import io
import re
import unicodedata
from typing import Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pypdf import PdfReader

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Import. Dong Sheng Ltda.")

st.markdown("""
Sube tu archivo PDF de Previred (`CtrlPdf.pdf`) y tu planilla corporativa. El sistema extraerá los sueldos y rellenará automáticamente la fila de **Agosto** en la pestaña **SUELDOS 2026**.
""")

# Definición de columnas y expresiones regulares
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
    
    # Buscamos todos los RUTs únicos en el texto y extraemos montos cercanos
    matches = list(RUT_RE.finditer(full_text))
    records = {}

    for match in matches:
        rut_raw = match.group(1)
        rut = normalize_rut(rut_raw)
        if not rut:
            continue
        
        # Extraer texto alrededor del RUT (ventana de 400 caracteres)
        start = max(0, match.start() - 100)
        end = min(len(full_text), match.end() + 300)
        context = full_text[start:end]
        
        # Buscar números grandes que correspondan a remuneración imponible
        numbers = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d{5,7}\b", context)
        amounts = [parse_clp(n) for n in numbers if parse_clp(n) and parse_clp(n) > 50000]
        
        if rut not in records:
            records[rut] = {
                "rut": rut,
                "sueldo_imponible": amounts[0] if amounts else 0,
                "afp": amounts[1] if len(amounts) > 1 else 0
            }
        else:
            if amounts and records[rut]["sueldo_imponible"] == 0:
                records[rut]["sueldo_imponible"] = amounts[0]

    df = pd.DataFrame(list(records.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "afp"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel.")
    
    ws = wb[SHEET_NAME]
    
    # Recorremos la hoja buscando RUTs para actualizar la fila de Agosto
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row, 2).value # Columna típica de RUT o nombres
        norm_cell = normalize_rut(cell_val)
        if norm_cell:
            for _, rec in df.iterrows():
                if normalize_rut(rec["rut"]) == norm_cell:
                    # Buscamos la fila de Agosto cerca de este RUT
                    for r_sub in range(row, row + 15):
                        mes_val = str(ws.cell(r_sub, 1).value or "").strip().upper()
                        if "AGOSTO" in mes_val:
                            # Inyectamos el sueldo imponible en la columna correspondiente (ej. columna B o C según tu formato)
                            val_imp = rec["sueldo_imponible"]
                            if val_imp > 0:
                                ws.cell(r_sub, 2).value = val_imp
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
        
        if st.button("🚀 Rellenar Planilla y Descargar", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla rellenada con éxito!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Actualizado",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Agosto_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")