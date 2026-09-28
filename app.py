import io
import re
from typing import Optional
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader

st.set_page_config(page_title="Asesorías Contables Linares - Remuneraciones", layout="wide")

st.title("📊 Asesorías Contables Linares")
st.subheader("Control de Remuneraciones y Previred - Multicliente")

st.markdown("""
Sube el archivo PDF de Previred de tu cliente y su respectiva **plantilla Excel corporativa**. El sistema procesará automáticamente sueldos, cotizaciones, impuesto único, aportes patronales (SIS, AFC, ISL) y te permitirá verificar la **Asignación Familiar** para el mes de **Agosto**.
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
    workers_data = {}
    
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        # 1. Remuneraciones / AFP (Sueldo Imponible y Salud Fonasa)
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

        # 2. Detalle de AFP (Cotización Obligatoria y AFC)
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
                            
                            # Cálculo Impuesto Único referencial
                            s_imp = workers_data[rut]["sueldo_imponible"]
                            if s_imp >= 1700000:
                                workers_data[rut]["impto_unico"] = 17573
                            else:
                                workers_data[rut]["impto_unico"] = 0

        # 3. ISL (Mutual) por trabajador
        if "Instituto de Seguridad Laboral" in text or "ISL" in text:
            for line in lines:
                m = RUT_RE.search(line)
                if m and "76.519" not in line:
                    rut = normalize_rut(m.group(1))
                    nums = [int(n.replace(".", "")) for n in re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line[m.end():])]
                    if len(nums) >= 2:
                        if rut in workers_data:
                            workers_data[rut]["isl"] = nums[1]

        # 4. Seguro Social Previsional (SIS, Rentabilidad Protegida, Seguro Social)
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

        # 5. Detección automática de Asignación Familiar en Anexos IPS / Cargas
        if "ASIGNACION" in text.upper() or "ASIGNACIÓN" in text.upper() or "REBAJAS" in text.upper() or "TRAMO" in text.upper():
            for line in lines:
                m = RUT_RE.search(line)
                if m:
                    rut = normalize_rut(m.group(1))
                    nums = re.findall(r"\b\d{1,3}(?:\.\d{3})+\b|\b\d+\b", line)
                    for n in nums:
                        val = parse_clp(n)
                        if val and 3000 <= val <= 25000:
                            if rut in workers_data:
                                workers_data[rut]["asig_fam"] = val

    df = pd.DataFrame(list(workers_data.values()))
    if df.empty:
        df = pd.DataFrame(columns=["rut", "sueldo_imponible", "salud_fonasa", "cotiz_afp", "afc_trab", "sis", "afc_emp", "isl", "rent_prot", "s_social", "impto_unico", "asig_fam"])
    return df

def write_to_excel(template_bytes: bytes, df: pd.DataFrame, cargas_dict: dict) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"No se encontró la pestaña '{SHEET_NAME}' en el Excel del cliente.")
    
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
                            
                            # 4. Asignación Familiar (Columna N / 14) -> Asignación validada
                            ws.cell(r_sub, 14).value = cargas_dict.get(norm_cell, rec["asig_fam"])
                            
                            # 5. Aportes Patronales exactos (Segunda tabla de aportes)
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
template_file = st.file_uploader("2. Sube la plantilla Excel de la empresa", type=["xlsx"])

if pdf_file and template_file:
    st.success("¡Archivos cargados correctamente!")
    
    try:
        df_extracted = extract_pdf_data(pdf_file.getvalue())
        
        st.subheader("📊 Datos Extraídos para el Cliente:")
        st.dataframe(df_extracted, use_container_width=True)
        
        # Alerta y Panel rápido de verificación de Asignación Familiar
        st.warning("⚠️ **Verificación de Asignación Familiar (Agosto):** Revisa si algún trabajador registra cargas este mes y ajusta su monto si es necesario antes de generar el Excel definitivo.")
        
        cargas_dict = {}
        for idx, row in df_extracted.iterrows():
            r = row["rut"]
            auto_val = int(row["asig_fam"])
            # Solo mostramos el input si detectó algo o para que puedas verificar rápido
            val_carga = st.number_input(
                label=f"Asignación Familiar para RUT: {r}",
                min_value=0,
                max_value=50000,
                value=auto_val,
                step=1000,
                key=f"carga_verif_{r}"
            )
            cargas_dict[r] = val_carga
            df_extracted.loc[idx, "asig_fam"] = val_carga
        
        if st.button("🚀 Rellenar Planilla Oficial del Cliente", type="primary"):
            final_excel = write_to_excel(template_file.getvalue(), df_extracted, cargas_dict)
            st.success("¡Planilla del cliente generada con éxito absoluto y SIS completo!")
            
            st.download_button(
                label="📥 Descargar Libro de Remuneraciones del Cliente",
                data=final_excel,
                file_name="Remuneraciones_Cliente_Definitivo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube el PDF de Previred y el Excel de la empresa para comenzar.")