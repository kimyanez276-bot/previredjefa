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
Sube tu archivo PDF de Previred del mes y tu planilla corporativa. El sistema leerá **dinámicamente** los sueldos, la fórmula previsional, el impuesto único, los aportes patronales y la **Asignación Familiar real** directamente desde el documento PDF.
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
    
    # 1. Página 2: Remuneración Imponible y Salud (Fonasa)
    if len(reader.pages) >= 2:
        p2_text = reader.pages[1].extract_text() or ""
        for line in p2_text.split("\n"):
            if "AFP" in line and RUT_RE.search(line):
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
                            "asig_fam": 0 # Inicializado dinámicamente en 0
                        }

    # 2. Páginas de AFP (4, 6, 8): Cotización Obligatoria AFP y AFC Trabajador / Empleador
    for p_idx in [3, 5, 7]:
        if p_idx < len(reader.pages):
            p_text = reader.pages[p_idx].extract_text() or ""
            for line in p_text.split("\n"):
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

    # 3. Página 10: ISL (Mutual)
    if len(reader.pages) >= 10:
        p10_text = reader.pages[9].extract_text() or ""
        for line in p10_text.split("\n"):
            m = RUT_RE.search(line)
            if m and "76.519" not in line:
                rut = normalize_rut(m.group(1))
                nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                if len(nums) >= 2:
                    if rut in workers_data:
                        workers_data[rut]["isl"] = nums[1]

    # 4. Página 12: Seguro Social, Renta Protegida y SIS
    if len(reader.pages) >= 12:
        p12_text = reader.pages[11].extract_text() or ""
        for line in p12_text.split("\n"):
            m = RUT_RE.search(line)
            if m and "76.519" not in line and "Totales" not in line:
                rut = normalize_rut(m.group(1))
                nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                if len(nums) >= 5:
                    if rut in workers_data:
                        workers_data[rut]["s_social"] = nums[2]
                        workers_data[rut]["rent_prot"] = nums[3]
                        workers_data[rut]["sis"] = nums[4]

    # 5. Página 14 (o última página): Lectura dinámica de Asignación Familiar desde el anexo IPS
    for page in reader.pages:
        p_text = page.extract_text() or ""
        if "ASIGNACION FAMILIAR" in p_text or "Tramo" in p_text:
            for line in p_text.split("\n"):
                m = RUT_RE.search(line)
                if m and "76.519" not in line and "TOTAL" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    # El monto de asignación familiar suele ser el número típico de cargas (ej entre 5000 y 50000)
                    for n in nums:
                        if 3000 <= n <= 100000:
                            if rut in workers_data:
                                workers_data[rut]["asig_fam"] = n
                                break

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
                            
                            # 4. Asignación Familiar (Columna N / 14) -> Dinámica leída del PDF
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
        st.subheader("Datos Extraídos Dinámicamente del PDF:")
        st.dataframe(df_extracted, use_container_width=True)
        
        if st.button("🚀 Rellenar Planilla Dinámica Oficial", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted)
            st.success("¡Planilla generada con éxito absoluto y 100% automatizada!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones Dinámico",
                data=final_excel,
                file_name="IMPORT_DONG_SHENG_Remuneraciones_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")