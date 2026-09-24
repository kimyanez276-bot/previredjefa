"""
Módulo Generador de Reportes Oficiales:
1. Exportador a Excel (.xlsx) con la plantilla anual exacta (Resumen Empresa + Bloques individuales).
2. Generador del archivo plano (.txt) con la estructura técnica del SII para la DJ 1887.
3. Generador de Certificados Nº 6 de Sueldos y Rentas en PDF para los trabajadores (ReportLab).
"""

import io
import os
import zipfile
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from modules.tax_engine import (
    clean_rut, format_rut_dots, get_factor_actualizacion,
    MESES_MAP, MESES_NOMBRE
)


# --- 1. GENERADOR DE EXCEL CON PLANTILLA ANUAL EXACTA ---

def generate_excel_libro_sueldos_anual(
    empresa_info: Dict[str, Any],
    anio: int,
    df_anual: pd.DataFrame,
    custom_factors: Optional[Dict[int, float]] = None
) -> io.BytesIO:
    """
    Genera el archivo Excel idéntico a la plantilla contable anual chilena:
    - Hoja Resumen Empresa con los 12 meses consolidados y factores.
    - Bloques individuales para cada trabajador con sus 12 meses y factores.
    - Hoja adicional de Base de Datos para análisis.
    """
    wb = openpyxl.Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Resumen DJ 1887 y Sueldos"
    ws_resumen.views.sheetView[0].showGridLines = True
    
    # Fuentes y estilos
    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_section = Font(name="Calibri", size=11, bold=True, color="000000")
    font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=9, bold=True)
    font_normal = Font(name="Calibri", size=9)
    font_small = Font(name="Calibri", size=8)
    
    fill_header_main = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header_worker = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    fill_total = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_top_bottom_bold = Border(
        top=Side(style='thin', color='000000'),
        bottom=Side(style='double', color='000000')
    )
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    curr_row = 1
    
    # Encabezado Empresa
    ws_resumen.cell(row=curr_row, column=1, value="RESUMEN DECLARACION JURADA Nº 1887").font = font_title
    curr_row += 2
    ws_resumen.cell(row=curr_row, column=1, value="NOMBRE EMPLEADOR").font = font_section
    ws_resumen.cell(row=curr_row, column=3, value=empresa_info.get("razon_social", "")).font = font_section
    curr_row += 1
    ws_resumen.cell(row=curr_row, column=1, value="RUT").font = font_section
    ws_resumen.cell(row=curr_row, column=3, value=format_rut_dots(empresa_info.get("rut", ""))).font = font_section
    curr_row += 1
    ws_resumen.cell(row=curr_row, column=1, value="DIRECCION").font = font_section
    ws_resumen.cell(row=curr_row, column=3, value=empresa_info.get("direccion", "")).font = font_section
    curr_row += 1
    ws_resumen.cell(row=curr_row, column=1, value="AÑO TRIBUTARIO").font = font_section
    ws_resumen.cell(row=curr_row, column=3, value=f"{anio} (Año Comercial {anio-1})").font = font_section
    curr_row += 2

    # Columnas exactas solicitadas
    columnas_cabecera = [
        "PERÍODO", "SUELDO BRUTO", "COTIZACION PREVISIONAL", "SEGURO SOCIAL (0,1%)",
        "RENTA NETA PAGADA", "FACTOR ACTUALIZAC.", "SUELDO BRUTO ACTUALIZADO",
        "MONTOS ACTUALIZADOS RENTA NETA PAGADA", "GASTOS PATRONALES", "IMP. PAGADAS SIN ACTUALIZAR",
        "IMPUESTO ÚNICO", "IMPTO. ÚNICO ACTUALIZADO", "ASIGNACIÓN FAMILIAR",
        "SIS", "AFC EMPLEADOR", "ISL MUTUAL", "S. SOCIAL PATRONAL", "S. SOCIAL (0,1 %)"
    ]

    def render_12_months_block(start_row: int, df_subset: Optional[pd.DataFrame] = None, header_fill = fill_header_main):
        r = start_row
        # Escribir cabecera
        for col_idx, col_name in enumerate(columnas_cabecera, start=1):
            cell = ws_resumen.cell(row=r, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_thin
        ws_resumen.row_dimensions[r].height = 28
        r += 1

        first_data_row = r
        for m in range(1, 13):
            m_name = MESES_MAP[m]
            factor = get_factor_actualizacion(m, anio, custom_factors)
            
            # Buscar datos para el mes m en df_subset
            if df_subset is not None and not df_subset.empty:
                m_rows = df_subset[df_subset["mes"] == m]
            else:
                m_rows = pd.DataFrame()

            bruto = float(m_rows["sueldo_bruto"].sum()) if not m_rows.empty else 0.0
            cotiz = float(m_rows["cotizacion_previsional_total"].sum()) if not m_rows.empty else 0.0
            ss_01 = float(m_rows["seguro_social_01"].sum()) if not m_rows.empty else 0.0
            r_neta = float(m_rows["renta_neta_pagada"].sum()) if not m_rows.empty else 0.0
            gastos_pat = float(m_rows["gastos_patronales_total"].sum()) if not m_rows.empty else 0.0
            iu = float(m_rows["impuesto_unico"].sum()) if not m_rows.empty else 0.0
            asig_fam = float(m_rows["asignacion_familiar"].sum()) if not m_rows.empty else 0.0
            sis = float(m_rows["sis"].sum()) if not m_rows.empty else 0.0
            afc_emp = float(m_rows["afc_empleador"].sum()) if not m_rows.empty else 0.0
            isl = float(m_rows["isl_mutual"].sum()) if not m_rows.empty else 0.0
            ss_pat = float(m_rows["seguro_social_patronal"].sum()) if not m_rows.empty else 0.0

            # Celdas y Fórmulas
            ws_resumen.cell(row=r, column=1, value=m_name).alignment = align_left
            ws_resumen.cell(row=r, column=2, value=bruto).number_format = '#,##0'
            ws_resumen.cell(row=r, column=3, value=cotiz).number_format = '#,##0'
            ws_resumen.cell(row=r, column=4, value=ss_01).number_format = '#,##0'
            ws_resumen.cell(row=r, column=5, value=r_neta).number_format = '#,##0'
            ws_resumen.cell(row=r, column=6, value=factor).number_format = '0.000'
            
            # Fórmulas Excel de montos actualizados: Bruto * Factor y Renta Neta * Factor
            ws_resumen.cell(row=r, column=7, value=f"=B{r}*F{r}").number_format = '#,##0'
            ws_resumen.cell(row=r, column=8, value=f"=E{r}*F{r}").number_format = '#,##0'
            ws_resumen.cell(row=r, column=9, value=gastos_pat).number_format = '#,##0'
            ws_resumen.cell(row=r, column=10, value=f"=C{r}+I{r}").number_format = '#,##0'
            ws_resumen.cell(row=r, column=11, value=iu).number_format = '#,##0'
            ws_resumen.cell(row=r, column=12, value=f"=K{r}*F{r}").number_format = '#,##0'
            ws_resumen.cell(row=r, column=13, value=asig_fam).number_format = '#,##0'
            ws_resumen.cell(row=r, column=14, value=sis).number_format = '#,##0'
            ws_resumen.cell(row=r, column=15, value=afc_emp).number_format = '#,##0'
            ws_resumen.cell(row=r, column=16, value=isl).number_format = '#,##0'
            ws_resumen.cell(row=r, column=17, value=ss_pat).number_format = '#,##0'
            ws_resumen.cell(row=r, column=18, value=ss_01).number_format = '#,##0'

            for c in range(1, 19):
                cell = ws_resumen.cell(row=r, column=c)
                cell.font = font_normal
                cell.border = border_thin
                if c > 1 and c != 6:
                    cell.alignment = align_right
            r += 1

        last_data_row = r - 1
        # Fila TOTALES
        total_row = r
        ws_resumen.cell(row=total_row, column=1, value="TOTALES").font = font_bold
        ws_resumen.cell(row=total_row, column=1).alignment = align_left
        ws_resumen.cell(row=total_row, column=1).fill = fill_total
        ws_resumen.cell(row=total_row, column=1).border = border_top_bottom_bold

        for col_idx, col_letter in enumerate(["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"], start=2):
            cell = ws_resumen.cell(row=total_row, column=col_idx)
            cell.font = font_bold
            cell.fill = fill_total
            cell.border = border_top_bottom_bold
            cell.alignment = align_right
            if col_letter == "F":
                cell.value = ""
            else:
                cell.value = f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})"
                cell.number_format = '#,##0'
                
        return total_row + 2

    # 1. Bloque Consolidado Empresa
    ws_resumen.cell(row=curr_row, column=1, value="1. CONSOLIDADO ANUAL GENERAL DE LA EMPRESA").font = font_section
    curr_row += 1
    curr_row = render_12_months_block(curr_row, df_anual, header_fill=fill_header_main)
    curr_row += 1

    # 2. Bloques Individuales por Trabajador
    if not df_anual.empty:
        ruts_unicos = df_anual["rut"].unique()
        ws_resumen.cell(row=curr_row, column=1, value="2. DETALLE ANUAL INDIVIDUAL POR TRABAJADOR").font = font_title
        curr_row += 2

        for idx_w, w_rut in enumerate(ruts_unicos, start=1):
            df_w = df_anual[df_anual["rut"] == w_rut]
            w_nombre = df_w["nombre"].iloc[0] if not df_w.empty else "TRABAJADOR"
            
            # Encabezado del trabajador
            ws_resumen.cell(row=curr_row, column=1, value="NOMBRE").font = font_bold
            ws_resumen.cell(row=curr_row, column=3, value=w_nombre).font = font_bold
            ws_resumen.cell(row=curr_row, column=18, value=idx_w).font = font_bold
            curr_row += 1
            ws_resumen.cell(row=curr_row, column=1, value="RUT").font = font_bold
            ws_resumen.cell(row=curr_row, column=3, value=format_rut_dots(w_rut)).font = font_bold
            curr_row += 1
            
            curr_row = render_12_months_block(curr_row, df_w, header_fill=fill_header_worker)
            curr_row += 1

    # Autoajustar ancho de columnas
    for col in ws_resumen.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and not val_str.startswith("="):
                max_len = len(val_str)
        ws_resumen.column_dimensions[col_letter].width = max(max_len + 3, 14)
    ws_resumen.column_dimensions['A'].width = 16
    ws_resumen.column_dimensions['B'].width = 17
    ws_resumen.column_dimensions['C'].width = 16

    # Hoja 2: Base de Datos Plana para Tablas Dinámicas
    ws_db = wb.create_sheet(title="Base_Datos_Sueldos")
    ws_db.views.sheetView[0].showGridLines = True
    if not df_anual.empty:
        df_export = df_anual.copy()
        df_export["mes_nombre"] = df_export["mes"].map(MESES_MAP)
        df_export["factor_actualiz"] = df_export["mes"].apply(lambda m: get_factor_actualizacion(m, anio, custom_factors))
        df_export["sueldo_bruto_actualizado"] = df_export["sueldo_bruto"] * df_export["factor_actualiz"]
        df_export["renta_neta_actualizada"] = df_export["renta_neta_pagada"] * df_export["factor_actualiz"]
        df_export["impuesto_unico_actualizado"] = df_export["impuesto_unico"] * df_export["factor_actualiz"]
        
        # Escribir cabecera
        headers_db = list(df_export.columns)
        ws_db.append(headers_db)
        for cell in ws_db[1]:
            cell.font = font_header
            cell.fill = fill_header_main
            cell.alignment = align_center

        for row in df_export.itertuples(index=False):
            ws_db.append(list(row))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- 2. GENERADOR ARCHIVO PLANO TXT DJ 1887 PARA EL SII ---

def generate_txt_dj1887(
    empresa_info: Dict[str, Any],
    anio: int,
    df_anual: pd.DataFrame,
    custom_factors: Optional[Dict[int, float]] = None
) -> str:
    """
    Genera el archivo plano (.TXT) con el estándar oficial exigido por el SII
    para la Declaración Jurada Nº 1887 (Sueldos y Rentas del Art. 42 Nº 1).
    Estructura:
    - Registro Tipo 1: Carátula del Empleador
    - Registro Tipo 2: Detalle por RUT de Trabajador con Rentas e Impuestos Actualizados y marcas de meses
    - Registro Tipo 3: Resumen y Totales de Control
    """
    c_rut_emp = clean_rut(empresa_info.get("rut", "")).replace("-", "")
    lines = []
    
    if df_anual.empty:
        return "NO_DATA"

    # Agrupar por trabajador y calcular valores anuales actualizados
    workers_records = []
    tot_renta_neta_act = 0.0
    tot_imp_unico_act = 0.0
    tot_leyes_sociales_act = 0.0
    
    for rut, df_w in df_anual.groupby("rut"):
        r_clean = clean_rut(rut).replace("-", "")
        renta_neta_act = 0.0
        imp_unico_act = 0.0
        leyes_sociales_act = 0.0
        
        # Array de meses 1 al 12
        meses_activos = ["0"] * 12
        for _, row in df_w.iterrows():
            m = int(row["mes"])
            if 1 <= m <= 12 and row.get("sueldo_bruto", 0.0) > 0:
                meses_activos[m - 1] = "1"
                factor = get_factor_actualizacion(m, anio, custom_factors)
                renta_neta_act += round(row.get("renta_neta_pagada", 0.0) * factor)
                imp_unico_act += round(row.get("impuesto_unico", 0.0) * factor)
                leyes_sociales_act += round(row.get("cotizacion_previsional_total", 0.0) * factor)

        tot_renta_neta_act += renta_neta_act
        tot_imp_unico_act += imp_unico_act
        tot_leyes_sociales_act += leyes_sociales_act

        workers_records.append({
            "rut": r_clean,
            "renta_neta_act": int(renta_neta_act),
            "imp_unico_act": int(imp_unico_act),
            "mayor_retencion": 0,
            "renta_exenta": 0,
            "rebaja_zonas": 0,
            "leyes_sociales_act": int(leyes_sociales_act),
            "meses_str": "".join(meses_activos)
        })

    cant_registros = len(workers_records)

    # 1. REGISTRO TIPO 1 (Carátula)
    # Formato: 1;RUT_DECLARANTE;AÑO_TRIBUTARIO;CANT_REGISTROS;FOLIO
    reg_1 = f"1;{c_rut_emp};{anio};{cant_registros};1887"
    lines.append(reg_1)

    # 2. REGISTROS TIPO 2 (Detalle por Trabajador)
    # Formato: 2;RUT_TRABAJADOR;RENTA_NETA_ACT;IMP_UNICO_ACT;MAYOR_RET;RENTA_EXENTA;REBAJA_ZONAS;LEYES_SOCIALES;MESES_01_12
    for w in workers_records:
        reg_2 = (
            f"2;{w['rut']};{w['renta_neta_act']};{w['imp_unico_act']};"
            f"{w['mayor_retencion']};{w['renta_exenta']};{w['rebaja_zonas']};"
            f"{w['leyes_sociales_act']};{w['meses_str']}"
        )
        lines.append(reg_2)

    # 3. REGISTRO TIPO 3 (Cuadro Resumen)
    # Formato: 3;CANT_REGISTROS;TOTAL_RENTA_NETA_ACT;TOTAL_IMP_UNICO_ACT;TOTAL_LEYES_SOCIALES
    reg_3 = f"3;{cant_registros};{int(tot_renta_neta_act)};{int(tot_imp_unico_act)};{int(tot_leyes_sociales_act)}"
    lines.append(reg_3)

    return "\r\n".join(lines)


# --- 3. GENERADOR DE CERTIFICADO Nº 6 DE SUELDOS Y RENTAS EN PDF ---

def generate_pdf_certificado_6(
    empresa_info: Dict[str, Any],
    anio: int,
    worker_rut: str,
    df_worker_anual: pd.DataFrame,
    custom_factors: Optional[Dict[int, float]] = None
) -> io.BytesIO:
    """
    Genera el Certificado Nº 6 oficial en PDF timbrado y firmado para un trabajador individual.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(
        'CertTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=16,
        alignment=1, # Centrado
        textColor=colors.HexColor('#1F497D')
    )
    style_subtitle = ParagraphStyle(
        'CertSub',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        alignment=1,
        textColor=colors.HexColor('#333333')
    )
    style_body = ParagraphStyle(
        'CertBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13
    )
    style_bold = ParagraphStyle(
        'CertBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=13
    )
    style_cell = ParagraphStyle(
        'CertCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=1
    )
    style_cell_r = ParagraphStyle(
        'CertCellR',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=2 # Derecha
    )
    style_cell_bold_r = ParagraphStyle(
        'CertCellBoldR',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        alignment=2
    )

    story = []

    # Título Principal
    story.append(Paragraph("CERTIFICADO Nº 6", style_title))
    story.append(Paragraph("SOBRE SUELDOS, PENSIONES O JUBILACIONES Y OTRAS RENTAS SIMILARES", style_title))
    story.append(Paragraph(f"AÑO TRIBUTARIO {anio} (Rentas Percibidas en el Año Comercial {anio-1})", style_subtitle))
    story.append(Paragraph("Emitido en cumplimiento del Art. 101 de la Ley sobre Impuesto a la Renta (Resolución Ex. SII)", style_subtitle))
    story.append(Spacer(1, 14))

    # Identificación Empleador y Trabajador
    w_nombre = df_worker_anual["nombre"].iloc[0] if not df_worker_anual.empty else "TRABAJADOR"
    
    info_table_data = [
        [
            Paragraph(f"<b>Razón Social Empleador:</b> {empresa_info.get('razon_social', '')}", style_body),
            Paragraph(f"<b>RUT Empleador:</b> {format_rut_dots(empresa_info.get('rut', ''))}", style_body)
        ],
        [
            Paragraph(f"<b>Dirección:</b> {empresa_info.get('direccion', 'Linares, VII Región')}", style_body),
            Paragraph(f"<b>Giro Comercial:</b> Comercio y Servicios", style_body)
        ],
        [
            Paragraph(f"<b>Nombre del Trabajador:</b> {w_nombre}", style_body),
            Paragraph(f"<b>RUT Trabajador:</b> {format_rut_dots(worker_rut)}", style_body)
        ]
    ]
    
    t_info = Table(info_table_data, colWidths=[330, 210])
    t_info.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#1F497D')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F2F5F9')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 12))

    # Tabla Mensualizada de Rentas e Impuestos
    headers_table = [
        Paragraph("<b>Mes</b>", style_cell),
        Paragraph("<b>Renta Bruta ($)</b>", style_cell),
        Paragraph("<b>Cotiz. Previs. ($)</b>", style_cell),
        Paragraph("<b>Renta Neta ($)</b>", style_cell),
        Paragraph("<b>Factor SII</b>", style_cell),
        Paragraph("<b>Renta Actualizada ($)</b>", style_cell),
        Paragraph("<b>Impto. Único ($)</b>", style_cell),
        Paragraph("<b>Impto. Actualizado ($)</b>", style_cell)
    ]
    
    data_rows = [headers_table]

    tot_bruto = 0.0
    tot_cotiz = 0.0
    tot_neta = 0.0
    tot_neta_act = 0.0
    tot_iu = 0.0
    tot_iu_act = 0.0

    for m in range(1, 13):
        m_name = MESES_NOMBRE[m - 1]
        factor = get_factor_actualizacion(m, anio, custom_factors)
        
        m_row = df_worker_anual[df_worker_anual["mes"] == m] if not df_worker_anual.empty else pd.DataFrame()
        bruto = float(m_row["sueldo_bruto"].sum()) if not m_row.empty else 0.0
        cotiz = float(m_row["cotizacion_previsional_total"].sum()) if not m_row.empty else 0.0
        neta = float(m_row["renta_neta_pagada"].sum()) if not m_row.empty else 0.0
        iu = float(m_row["impuesto_unico"].sum()) if not m_row.empty else 0.0
        
        neta_act = round(neta * factor)
        iu_act = round(iu * factor)

        tot_bruto += bruto
        tot_cotiz += cotiz
        tot_neta += neta
        tot_neta_act += neta_act
        tot_iu += iu
        tot_iu_act += iu_act

        data_rows.append([
            Paragraph(m_name, style_cell),
            Paragraph(f"{int(bruto):,}".replace(",", "."), style_cell_r),
            Paragraph(f"{int(cotiz):,}".replace(",", "."), style_cell_r),
            Paragraph(f"{int(neta):,}".replace(",", "."), style_cell_r),
            Paragraph(f"{factor:.3f}", style_cell),
            Paragraph(f"{int(neta_act):,}".replace(",", "."), style_cell_r),
            Paragraph(f"{int(iu):,}".replace(",", "."), style_cell_r),
            Paragraph(f"{int(iu_act):,}".replace(",", "."), style_cell_r),
        ])

    # Fila de Totales
    data_rows.append([
        Paragraph("<b>TOTALES</b>", style_cell),
        Paragraph(f"<b>{int(tot_bruto):,}</b>".replace(",", "."), style_cell_bold_r),
        Paragraph(f"<b>{int(tot_cotiz):,}</b>".replace(",", "."), style_cell_bold_r),
        Paragraph(f"<b>{int(tot_neta):,}</b>".replace(",", "."), style_cell_bold_r),
        Paragraph("-", style_cell),
        Paragraph(f"<b>{int(tot_neta_act):,}</b>".replace(",", "."), style_cell_bold_r),
        Paragraph(f"<b>{int(tot_iu):,}</b>".replace(",", "."), style_cell_bold_r),
        Paragraph(f"<b>{int(tot_iu_act):,}</b>".replace(",", "."), style_cell_bold_r),
    ])

    t_rentas = Table(data_rows, colWidths=[65, 68, 68, 68, 45, 80, 68, 78])
    t_rentas.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F497D')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D9D9D9')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#DCE6F1')),
        ('LINEABOVE', (0,-1), (-1,-1), 1, colors.HexColor('#1F497D')),
        ('LINEBELOW', (0,-1), (-1,-1), 1.5, colors.HexColor('#1F497D')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_rentas)
    story.append(Spacer(1, 14))

    # Declaración Legal y Firma
    declaracion_texto = (
        "Se extiende el presente certificado a petición del interesado para los fines tributarios y legales pertinentes, "
        "en especial para su Declaración Anual de Impuestos a la Renta (Formulario 22) ante el Servicio de Impuestos Internos (SII)."
    )
    story.append(Paragraph(declaracion_texto, style_body))
    story.append(Spacer(1, 40))

    # Bloque de Firmas
    fecha_actual = datetime.now().strftime("%d de %B de %Y")
    firma_data = [
        [
            Paragraph("____________________________________________<br/><b>REPRESENTANTE LEGAL / EMPLEADOR</b><br/>Firma y Timbre", style_cell),
            Paragraph(f"Fecha de emisión: {fecha_actual}<br/>Linares, Chile", style_cell)
        ]
    ]
    t_firma = Table(firma_data, colWidths=[270, 270])
    story.append(KeepTogether(t_firma))

    doc.build(story)
    buffer.seek(0)
    return buffer


def generate_zip_all_certificados(
    empresa_info: Dict[str, Any],
    anio: int,
    df_anual: pd.DataFrame,
    custom_factors: Optional[Dict[int, float]] = None
) -> io.BytesIO:
    """
    Genera un archivo ZIP conteniendo los Certificados Nº 6 en PDF de todos los trabajadores del año.
    """
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        if not df_anual.empty:
            for rut, df_w in df_anual.groupby("rut"):
                w_nombre = df_w["nombre"].iloc[0].replace(" ", "_")
                clean_r = clean_rut(rut)
                pdf_buffer = generate_pdf_certificado_6(empresa_info, anio, clean_r, df_w, custom_factors)
                filename = f"Certificado_6_{clean_r}_{w_nombre}.pdf"
                zip_file.writestr(filename, pdf_buffer.getvalue())
                
    zip_buffer.seek(0)
    return zip_buffer
