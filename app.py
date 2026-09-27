import io
import re
import unicodedata
from copy import copy
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


# ============================================================
# CONFIGURACIÓN
# ============================================================

SHEET_NAME = "SUELDOS 2026"
TARGET_MONTH = "AGOSTO"
DEFAULT_TEMPLATE = "IMPORT. DONG SHENG.xlsx"

# Campos que intentaremos obtener desde PREVIRED y escribir en Excel.
# Puedes ampliar esta lista si tu plantilla usa más conceptos.
FIELDS = [
    "sueldo_imponible",
    "afp",
    "salud",
    "afc",
    "total_cotizaciones",
]

DISPLAY_NAMES = {
    "rut": "RUT",
    "nombre": "Nombre",
    "sueldo_imponible": "Sueldo imponible",
    "afp": "AFP",
    "salud": "Salud",
    "afc": "AFC",
    "total_cotizaciones": "Total cotizaciones",
}

# Sinónimos para encontrar encabezados en la plantilla Excel.
HEADER_SYNONYMS = {
    "sueldo_imponible": [
        "SUELDO IMPONIBLE",
        "RENTA IMPONIBLE",
        "REMUNERACION IMPONIBLE",
        "REMUNERACIÓN IMPONIBLE",
        "TOTAL IMPONIBLE",
        "IMPONIBLE",
    ],
    "afp": [
        "AFP",
        "COTIZACION AFP",
        "COTIZACIÓN AFP",
        "PREVISION AFP",
        "PREVISIÓN AFP",
    ],
    "salud": [
        "SALUD",
        "COTIZACION SALUD",
        "COTIZACIÓN SALUD",
        "FONASA",
        "ISAPRE",
    ],
    "afc": [
        "AFC",
        "SEGURO CESANTIA",
        "SEGURO DE CESANTIA",
        "SEGURO CESANTÍA",
        "SEGURO DE CESANTÍA",
        "CESANTIA",
        "CESANTÍA",
    ],
    "total_cotizaciones": [
        "TOTAL COTIZACIONES",
        "TOTAL COTIZACION",
        "TOTAL COTIZACIÓN",
        "TOTAL PREVISIONAL",
        "TOTAL PREVISION",
        "TOTAL PREVISIÓN",
    ],
}

RUT_RE = re.compile(
    r"(?<!\d)(\d{1,2}(?:\.\d{3}){2}-[\dkK]|\d{7,8}-[\dkK])"
)

MONEY_TOKEN_RE = re.compile(
    r"(?<![\d.,])(?:\$?\s*)?(\d{1,3}(?:\.\d{3})+|\d+)(?![\d.,])"
)


# ============================================================
# UTILIDADES DE TEXTO / RUT / NÚMEROS
# ============================================================

def strip_accents(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return "".join(c for c in value if not unicodedata.combining(c))


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = strip_accents(str(value)).upper()
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_rut(value) -> str:
    """
    Devuelve RUT sin puntos, con guion y DV en mayúscula.
    Ej.: 12.345.678-9 -> 12345678-9
    """
    if value is None:
        return ""
    raw = re.sub(r"[^0-9kK]", "", str(value))
    if len(raw) < 2:
        return ""
    return f"{raw[:-1]}-{raw[-1].upper()}"


def format_rut(value) -> str:
    rut = normalize_rut(value)
    if not rut or "-" not in rut:
        return str(value or "")
    body, dv = rut.split("-")
    groups = []
    while body:
        groups.append(body[-3:])
        body = body[:-3]
    return f"{'.'.join(reversed(groups))}-{dv}"


def rut_is_valid(value) -> bool:
    rut = normalize_rut(value)
    if "-" not in rut:
        return False
    body, dv = rut.split("-")
    if not body.isdigit():
        return False

    s = 0
    multiplier = 2
    for digit in reversed(body):
        s += int(digit) * multiplier
        multiplier += 1
        if multiplier > 7:
            multiplier = 2

    expected = 11 - (s % 11)
    if expected == 11:
        expected_dv = "0"
    elif expected == 10:
        expected_dv = "K"
    else:
        expected_dv = str(expected)

    return expected_dv == dv.upper()


def parse_clp(value) -> Optional[int]:
    """
    Convierte valores típicos chilenos a entero:
    '$ 1.234.567' -> 1234567
    '1234567' -> 1234567
    """
    if value is None:
        return None

    if isinstance(value, (int, float)) and not pd.isna(value):
        return int(round(value))

    text = str(value).strip()
    if not text:
        return None

    # Evitar porcentajes como 7,00 o 10,77 %
    if "%" in text:
        return None

    text = text.replace("$", "").replace(" ", "")

    # Formato CLP: puntos como separador de miles.
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
        return int(text.replace(".", ""))

    # Número entero simple.
    if re.fullmatch(r"\d+", text):
        return int(text)

    return None


# ============================================================
# EXTRACCIÓN DE PDF CON PYPDF
# ============================================================

def extract_pdf_text(pdf_bytes: bytes) -> Tuple[str, List[str]]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_texts = []

    for page in reader.pages:
        txt = page.extract_text() or ""
        page_texts.append(txt)

    full_text = "\n".join(page_texts)
    return full_text, page_texts


def find_labeled_amount(text: str, label_patterns: List[str]) -> Optional[int]:
    """
    Busca montos asociados a etiquetas dentro de un bloque de texto.
    Ej.: 'Renta imponible $ 850.000'
    """
    normalized = strip_accents(text)

    for label in label_patterns:
        label_norm = strip_accents(label)
        patterns = [
            rf"{label_norm}\s*[:\-]?\s*\$?\s*([\d.]+)",
            rf"{label_norm}.*?\$?\s*([\d]{{1,3}}(?:\.\d{{3}})+|\d+)",
        ]
        for pattern in patterns:
            m = re.search(pattern, normalized, flags=re.IGNORECASE | re.DOTALL)
            if m:
                val = parse_clp(m.group(1))
                if val is not None:
                    return val
    return None


PDF_LABELS = {
    "sueldo_imponible": [
        r"RENTA\s+IMPONIBLE",
        r"SUELDO\s+IMPONIBLE",
        r"REMUNERACION\s+IMPONIBLE",
        r"TOTAL\s+IMPONIBLE",
    ],
    "afp": [
        r"COTIZACION\s+AFP",
        r"AFP",
        r"PREVISION",
    ],
    "salud": [
        r"COTIZACION\s+SALUD",
        r"SALUD",
        r"FONASA",
        r"ISAPRE",
    ],
    "afc": [
        r"SEGURO\s+(?:DE\s+)?CESANTIA",
        r"AFC",
        r"CESANTIA",
    ],
    "total_cotizaciones": [
        r"TOTAL\s+COTIZACIONES",
        r"TOTAL\s+PREVISIONAL",
    ],
}


def extract_name_from_context(context: str, rut_text: str) -> str:
    """
    Intenta obtener un nombre desde líneas cercanas al RUT.
    Es deliberadamente conservador para no confundir títulos del PDF con nombres.
    """
    lines = [x.strip() for x in context.splitlines() if x.strip()]
    rut_norm = normalize_rut(rut_text)

    for i, line in enumerate(lines):
        if rut_norm and rut_norm in normalize_rut(line):
            candidates = []
            if i > 0:
                candidates.append(lines[i - 1])
            if i + 1 < len(lines):
                candidates.append(lines[i + 1])

            for cand in candidates:
                c = normalize_text(cand)
                if (
                    len(cand) >= 5
                    and not RUT_RE.search(cand)
                    and not re.search(r"\d{4,}", cand)
                    and not any(word in c for word in [
                        "PREVIRED", "AFP", "SALUD", "COTIZ", "EMPRESA",
                        "PERIODO", "TOTAL", "PAGINA", "REMUNERACION",
                    ])
                ):
                    return cand.strip()
    return ""


def parse_previred_pdf(full_text: str) -> pd.DataFrame:
    """
    Parser tolerante:
    1) detecta todos los RUT;
    2) crea un bloque de texto alrededor de cada RUT;
    3) busca montos etiquetados;
    4) si pypdf desordena la tabla, deja el registro disponible para revisión manual
       en la interfaz de Streamlit.
    """
    matches = list(RUT_RE.finditer(full_text))
    records = []

    if not matches:
        return pd.DataFrame(columns=["rut", "nombre", *FIELDS, "rut_valido"])

    for idx, match in enumerate(matches):
        rut_raw = match.group(1)
        rut = normalize_rut(rut_raw)

        # Evitar duplicados del mismo RUT si aparece varias veces en el PDF.
        if any(r["rut"] == rut for r in records):
            continue

        start = max(0, match.start() - 500)
        if idx + 1 < len(matches):
            end = min(len(full_text), matches[idx + 1].start() + 200)
        else:
            end = min(len(full_text), match.end() + 1500)

        context = full_text[start:end]

        rec = {
            "rut": rut,
            "nombre": extract_name_from_context(context, rut_raw),
            "rut_valido": rut_is_valid(rut),
        }

        for field in FIELDS:
            rec[field] = find_labeled_amount(context, PDF_LABELS[field])

        records.append(rec)

    df = pd.DataFrame(records)

    # Orden de columnas.
    ordered = ["rut", "nombre", *FIELDS, "rut_valido"]
    for col in ordered:
        if col not in df.columns:
            df[col] = None

    return df[ordered]


# ============================================================
# DETECCIÓN DE BLOQUES EN EXCEL
# ============================================================

def cell_has_rut(cell_value, target_rut: str) -> bool:
    if cell_value is None:
        return False
    return normalize_rut(cell_value) == normalize_rut(target_rut)


def find_rut_cell(ws, rut: str) -> Optional[Tuple[int, int]]:
    """
    Busca el RUT en toda la hoja.
    Retorna (fila, columna).
    """
    target = normalize_rut(rut)

    for row in ws.iter_rows():
        for cell in row:
            if cell_has_rut(cell.value, target):
                return cell.row, cell.column

    return None


def find_month_row_near_rut(
    ws,
    rut_row: int,
    rut_col: int,
    month: str = TARGET_MONTH,
    max_rows_down: int = 20,
    cols_radius: int = 18,
) -> Optional[Tuple[int, int]]:
    """
    Busca 'AGOSTO' cerca del RUT, normalmente dentro del bloque Enero-Diciembre.
    """
    month_norm = normalize_text(month)

    r1 = max(1, rut_row - 3)
    r2 = min(ws.max_row, rut_row + max_rows_down)
    c1 = max(1, rut_col - cols_radius)
    c2 = min(ws.max_column, rut_col + cols_radius)

    candidates = []

    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            val = normalize_text(ws.cell(r, c).value)
            if val == month_norm or month_norm in val:
                # Penaliza filas anteriores al RUT y columnas muy lejanas.
                distance = abs(r - rut_row) * 2 + abs(c - rut_col)
                if r < rut_row:
                    distance += 20
                candidates.append((distance, r, c))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])
    _, r, c = candidates[0]
    return r, c


def header_match_score(cell_text: str, synonyms: List[str]) -> int:
    txt = normalize_text(cell_text)
    if not txt:
        return 0

    best = 0
    for syn in synonyms:
        syn_n = normalize_text(syn)
        if txt == syn_n:
            best = max(best, 100 + len(syn_n))
        elif syn_n in txt:
            best = max(best, 50 + len(syn_n))
    return best


def detect_field_columns(
    ws,
    rut_row: int,
    month_row: int,
    rut_col: int,
) -> Dict[str, Optional[int]]:
    """
    Encuentra columnas de los conceptos buscando encabezados alrededor del bloque.
    Busca por encima de la fila de Agosto y alrededor del RUT.
    """
    result = {field: None for field in FIELDS}

    search_top = max(1, rut_row - 12)
    search_bottom = min(ws.max_row, month_row + 2)
    search_left = max(1, rut_col - 20)
    search_right = min(ws.max_column, rut_col + 35)

    for field, synonyms in HEADER_SYNONYMS.items():
        best = None  # (score, distance, col)

        for r in range(search_top, search_bottom + 1):
            for c in range(search_left, search_right + 1):
                score = header_match_score(ws.cell(r, c).value, synonyms)
                if score <= 0:
                    continue

                # Preferimos encabezados próximos y por encima de la fila del mes.
                distance = abs(month_row - r) + abs(rut_col - c) * 0.1
                if r > month_row:
                    distance += 50

                candidate = (score, -distance, c)
                if best is None or candidate > best:
                    best = candidate

        if best:
            result[field] = best[2]

    return result


def copy_style_from_nearby_month(ws, target_row: int, column: int):
    """
    Si la celda de Agosto no tiene estilo particular, copia estilo desde Julio
    (fila anterior) como respaldo. Normalmente la plantilla ya tiene formato.
    """
    if target_row <= 1:
        return

    src = ws.cell(target_row - 1, column)
    dst = ws.cell(target_row, column)

    if src.has_style and not dst.has_style:
        dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)


def write_records_to_workbook(
    template_bytes: bytes,
    df: pd.DataFrame,
    sheet_name: str = SHEET_NAME,
    month: str = TARGET_MONTH,
    manual_columns: Optional[Dict[str, str]] = None,
) -> Tuple[bytes, pd.DataFrame]:
    """
    Escribe los datos en la fila de Agosto de cada trabajador, identificado por RUT.
    Devuelve:
      - Excel final como bytes
      - reporte de escritura
    """
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"No existe la pestaña '{sheet_name}'. "
            f"Pestañas disponibles: {', '.join(wb.sheetnames)}"
        )

    ws = wb[sheet_name]
    report = []

    # Convertir letras de columna manuales a índices.
    manual_col_indexes = {}
    if manual_columns:
        from openpyxl.utils.cell import column_index_from_string

        for field, col_letter in manual_columns.items():
            col_letter = (col_letter or "").strip().upper()
            if col_letter:
                manual_col_indexes[field] = column_index_from_string(col_letter)

    for _, rec in df.iterrows():
        rut = normalize_rut(rec.get("rut"))

        status = {
            "RUT": format_rut(rut),
            "Nombre": str(rec.get("nombre") or ""),
            "Fila Agosto": "",
            "Resultado": "",
            "Detalle": "",
        }

        if not rut:
            status["Resultado"] = "ERROR"
            status["Detalle"] = "Registro sin RUT."
            report.append(status)
            continue

        rut_pos = find_rut_cell(ws, rut)
        if not rut_pos:
            status["Resultado"] = "NO ENCONTRADO"
            status["Detalle"] = "El RUT no aparece en la plantilla."
            report.append(status)
            continue

        rut_row, rut_col = rut_pos

        month_pos = find_month_row_near_rut(ws, rut_row, rut_col, month=month)
        if not month_pos:
            status["Resultado"] = "ERROR"
            status["Detalle"] = f"No se encontró la fila de {month} cerca del RUT."
            report.append(status)
            continue

        month_row, month_col = month_pos
        status["Fila Agosto"] = month_row

        detected = detect_field_columns(ws, rut_row, month_row, rut_col)

        # Las columnas manuales, si se indicaron, tienen prioridad.
        for field, col_idx in manual_col_indexes.items():
            detected[field] = col_idx

        written = []
        skipped = []

        for field in FIELDS:
            value = rec.get(field)

            if pd.isna(value) or value is None or str(value).strip() == "":
                skipped.append(f"{DISPLAY_NAMES[field]}: sin dato")
                continue

            try:
                numeric_value = int(float(value))
            except (TypeError, ValueError):
                skipped.append(f"{DISPLAY_NAMES[field]}: valor inválido")
                continue

            col_idx = detected.get(field)
            if not col_idx:
                skipped.append(f"{DISPLAY_NAMES[field]}: columna no detectada")
                continue

            copy_style_from_nearby_month(ws, month_row, col_idx)
            ws.cell(month_row, col_idx).value = numeric_value
            written.append(
                f"{DISPLAY_NAMES[field]} -> {get_column_letter(col_idx)}{month_row}"
            )

        if written:
            status["Resultado"] = "OK"
            status["Detalle"] = "; ".join(written)
            if skipped:
                status["Detalle"] += " | Revisar: " + "; ".join(skipped)
        else:
            status["Resultado"] = "SIN DATOS"
            status["Detalle"] = "; ".join(skipped) or "No se escribió ningún valor."

        report.append(status)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), pd.DataFrame(report)


# ============================================================
# APP STREAMLIT
# ============================================================

st.set_page_config(
    page_title="PREVIRED → Sueldos 2026",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PREVIRED → plantilla de Sueldos 2026")
st.caption(
    "Sube el PDF de PREVIRED, revisa los datos extraídos y genera una copia "
    "de la plantilla con la fila de Agosto completada por RUT."
)

st.info(
    "La app trabaja sobre una copia del Excel y no modifica el archivo original. "
    "Como los PDF de PREVIRED pueden cambiar de estructura, siempre revisa la "
    "tabla de datos antes de generar el archivo final."
)

with st.sidebar:
    st.header("Configuración")
    sheet_name = st.text_input("Pestaña Excel", value=SHEET_NAME)
    target_month = st.text_input("Mes a completar", value=TARGET_MONTH)

    st.subheader("Mapeo manual opcional")
    st.caption(
        "Déjalo vacío para detección automática. Si la plantilla usa encabezados "
        "especiales, escribe solo la letra de la columna (por ejemplo: H, J, K)."
    )

    manual_columns = {
        "sueldo_imponible": st.text_input("Columna Sueldo imponible", value=""),
        "afp": st.text_input("Columna AFP", value=""),
        "salud": st.text_input("Columna Salud", value=""),
        "afc": st.text_input("Columna AFC", value=""),
        "total_cotizaciones": st.text_input("Columna Total cotizaciones", value=""),
    }


pdf_file = st.file_uploader(
    "1. Sube el PDF de PREVIRED",
    type=["pdf"],
    accept_multiple_files=False,
)

template_file = st.file_uploader(
    "2. Sube la plantilla IMPORT. DONG SHENG.xlsx",
    type=["xlsx"],
    accept_multiple_files=False,
)

if pdf_file and template_file:
    pdf_bytes = pdf_file.getvalue()
    template_bytes = template_file.getvalue()

    try:
        full_text, page_texts = extract_pdf_text(pdf_bytes)
    except Exception as exc:
        st.error(f"No fue posible leer el PDF: {exc}")
        st.stop()

    if not full_text.strip():
        st.error(
            "El PDF no contiene texto extraíble. Probablemente es un PDF escaneado. "
            "pypdf no hace OCR; en ese caso necesitarás una versión PDF con texto "
            "o agregar OCR al flujo."
        )
        st.stop()

    parsed_df = parse_previred_pdf(full_text)

    st.subheader("3. Revisión de datos extraídos")

    if parsed_df.empty:
        st.warning(
            "No se detectaron RUT en el PDF. Puedes revisar el texto extraído "
            "para adaptar las expresiones regulares."
        )
    else:
        invalid_count = int((~parsed_df["rut_valido"].fillna(False)).sum())
        if invalid_count:
            st.warning(
                f"Se detectaron {invalid_count} RUT que no pasan la validación "
                "del dígito verificador. Revísalos antes de continuar."
            )

        edited_df = st.data_editor(
            parsed_df,
            width="stretch",
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "rut": st.column_config.TextColumn("RUT"),
                "nombre": st.column_config.TextColumn("Nombre"),
                "sueldo_imponible": st.column_config.NumberColumn(
                    "Sueldo imponible", min_value=0, step=1, format="%d"
                ),
                "afp": st.column_config.NumberColumn(
                    "AFP", min_value=0, step=1, format="%d"
                ),
                "salud": st.column_config.NumberColumn(
                    "Salud", min_value=0, step=1, format="%d"
                ),
                "afc": st.column_config.NumberColumn(
                    "AFC", min_value=0, step=1, format="%d"
                ),
                "total_cotizaciones": st.column_config.NumberColumn(
                    "Total cotizaciones", min_value=0, step=1, format="%d"
                ),
                "rut_valido": st.column_config.CheckboxColumn(
                    "RUT válido", disabled=True
                ),
            },
            key="previred_editor",
        )

        col1, col2 = st.columns([1, 1])

        with col1:
            st.metric("Trabajadores detectados", len(edited_df))

        with col2:
            valid_amounts = (
                edited_df[FIELDS]
                .notna()
                .sum()
                .sum()
            )
            st.metric("Montos detectados", int(valid_amounts))

        if st.button(
            "4. Generar Excel con Agosto",
            type="primary",
            use_container_width=True,
        ):
            try:
                output_bytes, report_df = write_records_to_workbook(
                    template_bytes=template_bytes,
                    df=edited_df,
                    sheet_name=sheet_name,
                    month=target_month,
                    manual_columns=manual_columns,
                )
            except Exception as exc:
                st.error(f"No fue posible generar el Excel: {exc}")
                st.stop()

            st.subheader("Resultado de la carga")
            st.dataframe(report_df, width="stretch", hide_index=True)

            ok_count = int((report_df["Resultado"] == "OK").sum())
            total_count = len(report_df)

            if ok_count == total_count:
                st.success(
                    f"Listo: se procesaron correctamente los {ok_count} trabajadores."
                )
            else:
                st.warning(
                    f"Se escribieron datos para {ok_count} de {total_count} trabajadores. "
                    "Revisa el reporte antes de usar el archivo."
                )

            st.download_button(
                label="⬇️ Descargar IMPORT. DONG SHENG - Agosto.xlsx",
                data=output_bytes,
                file_name="IMPORT. DONG SHENG - Agosto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with st.expander("Diagnóstico: texto extraído del PDF"):
        st.caption(
            "Útil si PREVIRED cambia el formato o si algún monto no fue reconocido."
        )
        st.text_area(
            "Texto extraído",
            value=full_text[:100000],
            height=350,
        )

else:
    st.write(
        "Para comenzar, sube el PDF de PREVIRED y la plantilla Excel."
    )
