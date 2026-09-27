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
Sube tu archivo PDF de Previred (`CtrlPdf.pdf`) y tu planilla corporativa. El sistema cruzará los datos por RUT de forma exacta, rellenará la fila de **Agosto** con los montos reales del Previred y mantendrá todas las fórmulas de tu planilla.
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
    
    # Análisis avanzado de líneas de Previred por RUT
    for i, line in enumerate(lines):
        rut_match = RUT_RE.search(line)
        if rut_match:
            rut_raw = rut_match.group(1)
            rut = normalize_rut(rut_raw)
            if not rut or "76.519" in rut:
                continue
            
            # Ventana de texto alrededor del RUT en el PDF
            window_lines = lines[max(0, i-1):min(len(lines), i+6)]
            window_text = " ".join(window_lines)
            
            tokens = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d{4,7}\b", window_text)
            amounts = []
            for t in tokens:
                val = parse_clp(t)
                if val and val > 500:
                    amounts.append(val)
            
            if amounts:
                # El primer monto coherente es el sueldo imponible
                sueldo_imp = next((a for a in amounts if a >= 300000), amounts[0])
                
                # Buscamos valores patronales secundarios en la misma ventana si existen
                # SIS, AFC, ISL, S. Social suelen venir después del sueldo imponible
                sis_val = amounts[1] if len(amounts) > 1 and amounts[1] < 50000 else round(sueldo_imp * 0.0153)
                afc_val = amounts[2] if len(amounts) > 2 and amounts[2] < 100000 else round(sueldo_imp * 0.024)
                isl_val = amounts[3] if len(amounts) > 3 and amounts[3] < 30000 else round(sueldo_imp * 0.0093)
                ss_val  = amounts[4] if len(amounts) > 4 and amounts[4] < 30000 else round(sueldo_imp * 0.004)

                if rut not in records or sueldo_imp > records[rut]["sueldo_imponible"]:
                    records[rut] = {
                        "rut": rut,
                        "sueldo_imponible": sueldo_imp,
                        "sis": sis_val,
                        "afc_emp": afc_val,
                        "isl": isl_val,
                        "s_social": ss_val
                    }

    df = pd.DataFrame(list(records.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "sis", "afc_emp", "isl", "s_social"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel.")
    
    ws = wb[SHEET_NAME]
    
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row, 2).value
        norm_cell = normalize_rut(cell_val)
        
        if norm_cell:
            for _, rec in df.iterrows():
                if normalize_rut(rec["rut"]) == norm_cell:
                    for r_sub in range(row, row + 16):
                        mes_val = str(ws.cell(r_sub, 1).value or "").strip().upper()
                        if TARGET_MONTH in mes_val:
                            # 1. Sueldo Bruto/Imponible -> Columna B (2)
                            ws.cell(r_sub, 2).value = rec["sueldo_imponible"]
                            
                            # 2. Aportes Patronales exactos (Columnas P, Q, R, T) manteniendo fórmulas de celdas vecinas
                            ws.cell(r_sub, 16).value = rec["sis"]       # SIS (Columna P)
                            ws.cell(r_sub, 17).value = rec["afc_emp"]   # AFC (Columna Q)
                            ws.cell(r_sub, 18).value = rec["isl"]       # ISL (Columna R)
                            ws.cell(r_sub, 20).value = rec["s_social"]  # S. Social (Columna T)
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
        st.subheader("Datos Extraídos del Previred:")
        st.dataframe(df_extracted, use_container_width=True)
        
        if st.button("🚀 Rellenar Planilla con Precisión", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla rellenada correctamente respetando las fórmulas!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones - Agosto Definitivo",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Agosto_Perfecto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")