import streamlit as st
import pandas as pd
import pdfplumber
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import io

# ================================
# CONFIG
# ================================
SHEET_NAME = "SUELDOS 2026"
TARGET_MONTH = "AGOSTO"

# ================================
# UTILIDADES
# ================================

def clean_rut(rut):
    rut = re.sub(r"[^0-9kK]", "", str(rut))
    if len(rut) < 2:
        return ""
    return rut[:-1] + "-" + rut[-1].upper()

def parse_money(value):
    if not value:
        return None
    value = value.replace(".", "").replace("$", "").strip()
    if value.isdigit():
        return int(value)
    return None

# ================================
# EXTRAER PDF COMPLETO
# ================================

def extract_pdf_data(pdf_file):
    workers = {}
    
    with pdfplumber.open(pdf_file) as pdf:
        full_text = ""

        for page in pdf.pages:
            text = page.extract_text() or ""
            full_text += "\n" + text

    # Detectar RUTs
    rut_matches = list(re.finditer(r"\d{1,2}\.\d{3}\.\d{3}-[\dkK]", full_text))

    for i, match in enumerate(rut_matches):
        rut = clean_rut(match.group())

        start = match.start()
        end = rut_matches[i+1].start() if i+1 < len(rut_matches) else start + 2000

        block = full_text[start:end]

        # Buscar valores
        sueldo = re.search(r"(RENTA|SUELDO).*?([\d\.]{5,})", block, re.IGNORECASE)
        afp = re.search(r"AFP.*?([\d\.]{3,})", block, re.IGNORECASE)
        salud = re.search(r"(SALUD|FONASA|ISAPRE).*?([\d\.]{3,})", block, re.IGNORECASE)
        afc = re.search(r"(AFC|CESANTIA).*?([\d\.]{3,})", block, re.IGNORECASE)

        workers[rut] = {
            "rut": rut,
            "sueldo_imponible": parse_money(sueldo.group(2)) if sueldo else None,
            "afp": parse_money(afp.group(1)) if afp else None,
            "salud": parse_money(salud.group(2)) if salud else None,
            "afc": parse_money(afc.group(2)) if afc else None
        }

    return pd.DataFrame(workers.values())

# ================================
# BUSCAR RUT EN EXCEL
# ================================

def find_rut(ws, rut):
    for row in ws.iter_rows():
        for cell in row:
            if clean_rut(cell.value) == rut:
                return cell.row, cell.column
    return None, None

# ================================
# BUSCAR FILA AGOSTO
# ================================

def find_august_row(ws, start_row):
    for i in range(start_row, start_row + 15):
        for j in range(1, ws.max_column + 1):
            val = str(ws.cell(i, j).value).upper()
            if "AGOSTO" in val:
                return i
    return None

# ================================
# DETECTAR COLUMNAS
# ================================

def detect_columns(ws, header_row):
    mapping = {}

    for col in range(1, ws.max_column + 1):
        val = str(ws.cell(header_row, col).value).upper()

        if "IMPONIBLE" in val:
            mapping["sueldo_imponible"] = col
        elif "AFP" in val:
            mapping["afp"] = col
        elif "SALUD" in val:
            mapping["salud"] = col
        elif "AFC" in val or "CESANTIA" in val:
            mapping["afc"] = col

    return mapping

# ================================
# ESCRIBIR EN EXCEL
# ================================

def write_excel(template, df):
    wb = load_workbook(template)
    ws = wb[SHEET_NAME]

    report = []

    for _, row in df.iterrows():
        rut = row["rut"]

        r, c = find_rut(ws, rut)

        if not r:
            report.append([rut, "NO ENCONTRADO"])
            continue

        august_row = find_august_row(ws, r)

        if not august_row:
            report.append([rut, "SIN AGOSTO"])
            continue

        headers = detect_columns(ws, august_row - 1)

        for field, col in headers.items():
            if pd.notna(row[field]):
                ws.cell(august_row, col).value = int(row[field])

        report.append([rut, "OK"])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output, pd.DataFrame(report, columns=["RUT", "Estado"])

# ================================
# STREAMLIT APP
# ================================

st.title("📊 PREVIRED → AGOSTO automático")

pdf_file = st.file_uploader("Sube PDF PREVIRED", type="pdf")
excel_file = st.file_uploader("Sube plantilla Excel", type="xlsx")

if pdf_file and excel_file:

    st.info("Procesando PDF...")

    df = extract_pdf_data(pdf_file)

    st.subheader("Datos detectados")
    df = st.data_editor(df)

    if st.button("Generar Excel"):

        output, report = write_excel(excel_file, df)

        st.subheader("Resultado")
        st.dataframe(report)

        st.download_button(
            "Descargar archivo final",
            data=output,
            file_name="Previred_Agosto.xlsx"
        )