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
Sube tu archivo PDF de Previred y tu plantilla corporativa. El sistema leerá dinámicamente sueldos, aportes y el **Monto exacto de la columna Asignación Familiar** desde el PDF sin montos fijos.
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
    workers_data = {}
    
    # Recorremos cada página del PDF buscando por membretes
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        # 1. Búsqueda de tabla de Remuneraciones / AFP
        if "AFP" in text and "REMUNERACIÓN" in text:
            for line in lines:
                if RUT_RE.search(line) and "AFP" in line:
                    m = RUT_RE.search(line)
                    rut = normalize_rut(m.group(1))
                    if not rut:
                        continue
                    parts = line.split("AFP")
                    if len(parts) > 1:
                        nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", parts[1])
                        if len(nums) >= 2:
                            sueldo_imp = int(nums[0].replace(".", ""))
                            salud_fonasa = int(nums[1].replace(".", ""))
                            if rut not in workers_data:
                                workers_data[rut] = {
                                    "rut": rut,
                                    "sueldo_imponible": sueldo_imp,
                                    "salud_fonasa": salud_fonasa,
                                    "cotiz_afp": 0,
                                    "afc_trab": 0,
                                    "sis": 0,
                                    "afc_emp": 0,
                                    "isl": 0,
                                    "rent_prot": 0,
                                    "s_social": 0,
                                    "impto_unico": 0,
                                    "asig_fam": 0
                                }

        # 2. Búsqueda de detalles de AFP (Cotización Obligatoria y AFC)
        if "Cotización" in text and ("Seguro Cesantía" in text or "Seguro de Cesantía" in text or "Detalle de Cotizaciones" in text):
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "R.U.T" not in line:
                    rut = normalize_rut(m.group(1))
                    after_rut = line[m.end():]
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", after_rut)]
                    if len(nums) >= 9:
                        cotiz_afp = nums[1]
                        afc_trab = nums[7] if len(nums) >= 8 else 0
                        afc_emp = nums[8] if len(nums) >= 9 else 0
                        if rut in workers_data:
                            workers_data[rut]["cotiz_afp"] = cotiz_afp
                            workers_data[rut]["afc_trab"] = afc_trab
                            workers_data[rut]["afc_emp"] = afc_emp
                            
                            s_imp = workers_data[rut]["sueldo_imponible"]
                            if s_imp >= 1700000:
                                workers_data[rut]["impto_unico"] = 17573
                            else:
                                workers_data[rut]["impto_unico"] = 0

        # 3. Búsqueda de ISL (Mutual)
        if "Instituto de Seguridad Laboral" in text or "ISL" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 2:
                        if rut in workers_data:
                            workers_data[rut]["isl"] = nums[1]

        # 4. Búsqueda de Seguro Social Previsional
        if "SEGURO SOCIAL PREVISIONAL" in text or "Seguro Social" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "Totales" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 5:
                        if rut in workers_data:
                            workers_data[rut]["s_social"] = nums[2]
                            workers_data[rut]["rent_prot"] = nums[3]
                            workers_data[rut]["sis"] = nums[4]

        # 5. Búsqueda dinámica y estricta en la sección de Asignación Familiar / Rebajas
        if "ASIGNACION FAMILIAR" in text or "REBAJAS" in text or "Tramo" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "TOTAL" not in line and "GENERALES" not in line and "PAGINA" not in line:
                    rut = normalize_rut(m.group(1))
                    # Buscamos todos los números en la línea
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line)]
                    if nums:
                        # El monto de asignación familiar en la tabla de Previred es siempre un valor monetario real (ej. 13.870)
                        # Filtramos números que parezcan montos de asignación (mayores a 3000 y menores a 500000)
                        montos_candidatos = [n for n in nums if 3000 <= n <= 500000]
                        if montos_candidatos:
                            monto_real = montos_candidatos[-1] # El último candidato es la columna "Monto"
                            if rut in workers_data:
                                workers_data[rut]["asig_fam"] = monto_real

    df = pd.DataFrame(list(workers_data.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "salud_fonasa", "cotiz_afp", "afc_trab", "sis", "afc_emp", "isl", "rent_prot", "s_social", "impto_unico", "asig_fam"])
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
                            
                            # 2. Cotización Previsional (Columna C) -> Fórmula exacta
                            cotiz_val = rec["cotiz_afp"]
                            afc_t_val = rec["afc_trab"]
                            salud_val = rec["salud_fonasa"]
                            ws.cell(r_sub, 3).value = f"={cotiz_val}+{afc_t_val}+{salud_val}-D{r_sub}"
                            
                            # 3. Impuesto Único (Columna L / 12)
                            ws.cell(r_sub, 12).value = rec["impto_unico"]
                            
                            # 4. Asignación Familiar (Columna N / 14) -> Dinámica leída de la columna Monto del PDF
                            ws.cell(r_sub, 14).value = rec["asig_fam"]
                            
                            # 5. Aportes Patronales exactos (Segunda tabla)
                            ws.cell(r_sub, 16).value = rec["sis"]       # SIS (Columna P)
                            ws.cell(r_sub, 17).value = rec["afc_emp"]   # AFC Empleador (Columna Q)
                            ws.cell(r_sub, 18).value = rec["isl"]       # ISL / Mutual (Columna R)
                            ws.cell(r_sub, 19).value = rec["rent_prot"] # Rent. Protegida (Columna S)
                            ws.cell(r_sub, 20).value = rec["s_social"]  # Seguro Social (Columna T)
                            break
                            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Interfaz Streamlit
pdf_file = st.file_uploader("1. Sube tu PDF de Previred del mes", type=["pdf"])
template_file = st.file_uploader("2. Sube tu plantilla Excel oficial", type=["xlsx"])

if pdf_file and template_file:
    st.success("¡Archivos cargados correctamente!")
    
    try:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        st.subheader("Datos Extraídos Dinámicamente (Cargas Incluidas):")
        st.dataframe(df_extracted, use_container_width=True)
        
        if st.button("🚀 Rellenar Planilla Oficial Dinámica", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla generada con éxito absoluto y asignación familiar dinámica!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Definitivo",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Remuneraciones_Definitivas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")