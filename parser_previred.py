"""
Módulo de Extracción y Conciliación de Comprobantes Previred con pdfplumber.
Procesa archivos PDF de Fonasa, IPS, Mutual/ISL, Seguro Social Previsional y AFPs
(UNO, Modelo, PlanVital, Habitat, Cuprum, Capital, ProVida).
Consolida la información mensual por trabajador y calcula los totales para cuadratura.
"""

import io
import re
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import pdfplumber

from modules.tax_engine import clean_rut, format_rut_dots, parse_chilean_number, CODIGOS_MOVIMIENTO


def identify_previred_doc(text: str) -> str:
    """Identifica el tipo de comprobante previsional según su encabezado y contenido."""
    text_upper = text.upper()
    if "FONASA" in text_upper or "FONDO NACIONAL DE SALUD" in text_upper:
        return "FONASA"
    elif "IPS (EX INP)" in text_upper or "INSTITUTO DE PREVISIÓN SOCIAL" in text_upper:
        return "IPS"
    elif "SEGURO SOCIAL PREVISIONAL" in text_upper or "SEGURO SOCIAL" in text_upper and "RENTABILIDAD PROTEGIDA" in text_upper:
        return "SEGURO_SOCIAL"
    elif "INSTITUTO DE SEGURIDAD LABORAL" in text_upper or "ISL" in text_upper or "MUTUAL" in text_upper or "ACHS" in text_upper:
        return "MUTUAL"
    elif "AFP UNO" in text_upper:
        return "AFP_UNO"
    elif "AFP MODELO" in text_upper:
        return "AFP_MODELO"
    elif "AFP PLANVITAL" in text_upper or "PLANVITAL" in text_upper:
        return "AFP_PLANVITAL"
    elif "AFP HABITAT" in text_upper or "HABITAT" in text_upper:
        return "AFP_HABITAT"
    elif "AFP CUPRUM" in text_upper or "CUPRUM" in text_upper:
        return "AFP_CUPRUM"
    elif "AFP CAPITAL" in text_upper or "CAPITAL" in text_upper:
        return "AFP_CAPITAL"
    elif "AFP PROVIDA" in text_upper or "PROVIDA" in text_upper:
        return "AFP_PROVIDA"
    elif "AFP" in text_upper:
        return "AFP_GENERICA"
    return "DESCONOCIDO"


def extract_periodo_and_empresa(text: str) -> Dict[str, Any]:
    """Extrae el período (mes, año), RUT y Razón Social del empleador."""
    res = {
        "rut_empleador": "",
        "razon_social": "",
        "mes": 0,
        "anio": 0,
        "folio": ""
    }
    
    # Folio / Serie
    folio_match = re.search(r'(?:N°\s*SERIE\s*RESUMEN|NÚMERO\s*DE\s*SERIE|NÚMERO\s*DE\s*FOLIO)[:\s]*([0-9]+)', text, re.IGNORECASE)
    if folio_match:
        res["folio"] = folio_match.group(1).strip()
        
    # RUT Empleador
    rut_match = re.search(r'(?:R\.?U\.?T\.?(?:\s*EMPLEADOR|\s*ENTE\s*PAGADOR)?[:\s]*)([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK])', text, re.IGNORECASE)
    if rut_match:
        res["rut_empleador"] = clean_rut(rut_match.group(1))
    else:
        # Fallback RUT con formato simple
        ruts = re.findall(r'([0-9]{7,8}\s*-\s*[0-9kK])', text)
        if ruts:
            res["rut_empleador"] = clean_rut(ruts[0])
            
    # Razón Social
    rz_match = re.search(r'(?:RAZÓN\s*SOCIAL\s*O\s*NOMBRE|NOMBRE\s*O\s*RAZÓN\s*SOCIAL)[:\s]*([A-Z0-9\.\,\s\-]+?)(?:R\.?U\.?T|DIRECCIÓN|TELEFONO|\n)', text, re.IGNORECASE)
    if rz_match:
        res["razon_social"] = rz_match.group(1).strip()
        
    # Período: ej "REMUNERACION 07 2026" o "07/2026" o "Julio 2026"
    per_match = re.search(r'(?:REMUNERACI[OÓ]N|PER[IÍ]ODO(?:\s*DE\s*REMUNERACIONES)?)[^\n\d]*([0-9]{1,2})[\s\/]+([0-9]{4})', text, re.IGNORECASE)
    if per_match:
        res["mes"] = int(per_match.group(1))
        res["anio"] = int(per_match.group(2))
    else:
        # Buscar "Julio 2026"
        meses_nom = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        for idx, m_name in enumerate(meses_nom, start=1):
            if re.search(rf'\b{m_name}\s+([0-9]{{4}})\b', text, re.IGNORECASE):
                m_yr = re.search(rf'\b{m_name}\s+([0-9]{{4}})\b', text, re.IGNORECASE).group(1)
                res["mes"] = idx
                res["anio"] = int(m_yr)
                break
                
    return res


def parse_previred_pdf(file_bytes: bytes) -> Dict[str, Any]:
    """
    Parsea un archivo PDF de Previred en memoria extrayendo resumen y detalle de trabajadores.
    """
    doc_info = {
        "tipo_doc": "DESCONOCIDO",
        "institucion": "",
        "cabecera": {},
        "resumen_totales": {},
        "trabajadores": [],
        "alertas": []
    }
    
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_text = ""
        pages_text = []
        for p in pdf.pages:
            t = p.extract_text() or ""
            pages_text.append(t)
            all_text += "\n" + t
            
        doc_type = identify_previred_doc(all_text)
        doc_info["tipo_doc"] = doc_type
        doc_info["cabecera"] = extract_periodo_and_empresa(all_text)
        
        # --- PARSEO FONASA ---
        if doc_type == "FONASA":
            doc_info["institucion"] = "FONASA (Fondo Nacional de Salud)"
            # Resumen
            rem_imp = re.search(r'TOTAL\s*REMUNERACI[OÓ]N\s*IMPONIBLE\s*DECLARADA\s*EN\s*\$?\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            cotiz_legal = re.search(r'Cotiz\.?\s*Legal\s*\+?\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            asig_fam = re.search(r'Compensaci[oó]n\s*Asignac[ií][oó]n\s*Familiar\s*[-–]?\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            total_pagado = re.search(r'MONTO\s*PAGADO\s*=\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            
            doc_info["resumen_totales"] = {
                "remuneracion_imponible": parse_chilean_number(rem_imp.group(1)) if rem_imp else 0.0,
                "cotizacion_legal": parse_chilean_number(cotiz_legal.group(1)) if cotiz_legal else 0.0,
                "asignacion_familiar": parse_chilean_number(asig_fam.group(1)) if asig_fam else 0.0,
                "total_pagado": parse_chilean_number(total_pagado.group(1)) if total_pagado else 0.0,
            }
            
            # Detalle trabajadores: Formato estándar Fonasa
            # N° | RUT | Nombre completo | Días | Entidad | Rem Imponible | Cotiz 7% | Cod Mov | Fechas
            lineas = all_text.split("\n")
            for line in lineas:
                # Patrón: [N°]? [RUT] [NOMBRES] [DIAS] AFP/IPS [REM] [COTIZ] [COD_MOV]?
                rut_match = re.search(r'([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK])\s+([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)\s+([0-9]{1,2})\s+(AFP|IPS)\s+([0-9\.\,]+)\s+([0-9\.\,]+)(?:\s+([0-9]{1,2}))?', line)
                if rut_match:
                    r_raw = rut_match.group(1)
                    n_raw = rut_match.group(2).strip()
                    dias = int(rut_match.group(3))
                    entidad = rut_match.group(4)
                    rem = parse_chilean_number(rut_match.group(5))
                    cotiz = parse_chilean_number(rut_match.group(6))
                    cod_mov = int(rut_match.group(7)) if rut_match.group(7) else 0
                    
                    doc_info["trabajadores"].append({
                        "rut": clean_rut(r_raw),
                        "nombre": n_raw,
                        "dias_trabajados": dias,
                        "salud_entidad": "FONASA",
                        "remuneracion_salud": rem,
                        "cotizacion_salud": cotiz,
                        "cod_movimiento": cod_mov
                    })
                    
        # --- PARSEO IPS ---
        elif doc_type == "IPS":
            doc_info["institucion"] = "IPS (Instituto de Previsión Social)"
            asig_fam = re.search(r'Asignaci[oó]n\s*Familiar\s*\+?\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            total_pagado = re.search(r'A\s*FAVOR\s*DE\s*INSTITUCION\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            
            doc_info["resumen_totales"] = {
                "asignacion_familiar": parse_chilean_number(asig_fam.group(1)) if asig_fam else 0.0,
                "total_pagado": parse_chilean_number(total_pagado.group(1)) if total_pagado else 0.0,
            }
            
            # Anexo trabajadores IPS
            lineas = all_text.split("\n")
            for line in lineas:
                rut_match = re.search(r'([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK]|[0-9]{7,8}\s+[0-9kK])\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+([0-9]{1,2})\s+([0-9\.\,]+)', line)
                if rut_match and "TOTAL" not in line.upper():
                    r_str = rut_match.group(1).replace(" ", "-") if " " in rut_match.group(1) and "-" not in rut_match.group(1) else rut_match.group(1)
                    asig_val = 0.0
                    m_asig = re.search(r'([0-9\.\,]+)\s+0$', line)
                    if m_asig:
                        asig_val = parse_chilean_number(m_asig.group(1))
                    doc_info["trabajadores"].append({
                        "rut": clean_rut(r_str),
                        "nombre": rut_match.group(2).strip(),
                        "dias_trabajados": int(rut_match.group(3)),
                        "asignacion_familiar": asig_val
                    })
                    
        # --- PARSEO SEGURO SOCIAL PREVISIONAL ---
        elif doc_type == "SEGURO_SOCIAL":
            doc_info["institucion"] = "Seguro Social Previsional"
            rem_imp = re.search(r'TOTAL\s*REMUNERACI[OÓ]N\s*IMPONIBLE\s*DECLARADA\s*EN\s*\$?\s*\$?([0-9\.\,]+)', all_text, re.IGNORECASE)
            ss_prev = re.search(r'Seguro\s*Social\s*Previsional\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            tot_pag = re.search(r'MONTO\s*PAGADO\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            
            doc_info["resumen_totales"] = {
                "remuneracion_imponible": parse_chilean_number(rem_imp.group(1)) if rem_imp else 0.0,
                "seguro_social_previsional": parse_chilean_number(ss_prev.group(1)) if ss_prev else 0.0,
                "total_pagado": parse_chilean_number(tot_pag.group(1)) if tot_pag else 0.0,
            }
            
            lineas = all_text.split("\n")
            for line in lineas:
                # RUT | Nombre | Renta Imponible | Renta Lic | Seguro Social | Rentab Prot | SIS | JC | Días | Cod Mov
                rut_match = re.search(r'([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK])\s+([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+(JC|JP)\s+([0-9]{1,2})\s+([0-9]{1,2})?', line)
                if rut_match:
                    r_raw = rut_match.group(1)
                    n_raw = rut_match.group(2).strip()
                    rem = parse_chilean_number(rut_match.group(3))
                    ss_prev_val = parse_chilean_number(rut_match.group(5))
                    dias = int(rut_match.group(9))
                    cod_mov = int(rut_match.group(10)) if rut_match.group(10) else 0
                    
                    doc_info["trabajadores"].append({
                        "rut": clean_rut(r_raw),
                        "nombre": n_raw,
                        "remuneracion_imponible_ss": rem,
                        "seguro_social_patronal": ss_prev_val,
                        "seguro_social_01": round(rem * 0.001, 3), # 0.1% aporte trabajador
                        "dias_trabajados": dias,
                        "cod_movimiento": cod_mov
                    })
                    
        # --- PARSEO MUTUAL / ISL ---
        elif doc_type == "MUTUAL":
            doc_info["institucion"] = "Instituto de Seguridad Laboral (ISL) / Mutual"
            rem_imp = re.search(r'TOTAL\s*REMUNERACIONES\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            tot_pag = re.search(r'TOTAL\s*A\s*PAGAR\s*A\s*LA\s*MUTUAL\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            tasa = re.search(r'TASA\s*COTIZACI[OÓ]N\s*([0-9\.\,]+)%?', all_text, re.IGNORECASE)
            
            doc_info["resumen_totales"] = {
                "remuneracion_imponible": parse_chilean_number(rem_imp.group(1)) if rem_imp else 0.0,
                "tasa": parse_chilean_number(tasa.group(1)) if tasa else 0.93,
                "total_pagado": parse_chilean_number(tot_pag.group(1)) if tot_pag else 0.0,
            }
            
            lineas = all_text.split("\n")
            for line in lineas:
                # RUT | Apellido Paterno | Materno | Nombres | Remuneración | Cotización | Movimiento
                rut_match = re.search(r'([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK])\s+([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)\s+([0-9\.\,]+)\s+([0-9\.\,]+)(?:\s+([0-9]{1,2}))?', line)
                if rut_match and "TOTAL" not in line.upper():
                    r_raw = rut_match.group(1)
                    n_raw = rut_match.group(2).strip()
                    rem = parse_chilean_number(rut_match.group(3))
                    cotiz_isl = parse_chilean_number(rut_match.group(4))
                    cod_mov = int(rut_match.group(5)) if rut_match.group(5) else 0
                    
                    doc_info["trabajadores"].append({
                        "rut": clean_rut(r_raw),
                        "nombre": n_raw,
                        "remuneracion_isl": rem,
                        "cotizacion_isl": cotiz_isl,
                        "cod_movimiento": cod_mov
                    })
                    
        # --- PARSEO AFPs (UNO, MODELO, PLANVITAL, HABITAT, CUPRUM, CAPITAL, PROVIDA) ---
        elif doc_type.startswith("AFP_"):
            nom_afp = doc_type.replace("AFP_", "").replace("_", " ").title()
            doc_info["institucion"] = f"AFP {nom_afp}"
            
            # Resumen Fondo Pensiones
            rem_pen = re.search(r'FONDO\s*DE\s*PENSIONES\s*\n\s*Renta\s*Imponible\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            cot_oblig = re.search(r'Cotizaci[oó]n\s*Obligatoria\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            sis_tot = re.search(r'Seguro\s*Invalidez\s*y\s*Sobrevivencia\s*\(SIS\)\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            tot_pension = re.search(r'TOTAL\s*A\s*PAGAR\s*FONDO\s*DE\s*PENSIONES[^\n\d]*([0-9\.\,]+)', all_text, re.IGNORECASE)
            
            # Resumen Fondo Cesantía (AFC)
            rem_afc = re.search(r'FONDO\s*DE\s*CESANTIA\s*\n\s*Renta\s*Imponible\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            if not rem_afc:
                rem_afc = re.search(r'TOTAL\s*REMUNERACIONES\s*O\s*GRATIFICACIONES\s*\n?\s*([0-9\.\,]+)\s*NUMERO\s*AFILIADOS\s*INFORMADOS\s*FDO\.\s*CESANTIA', all_text, re.IGNORECASE)
            cot_afc_afil = re.search(r'Cotizaci[oó]n\s*Afiliados\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            cot_afc_emp = re.search(r'Cotizaci[oó]n\s*Empleador\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            tot_afc = re.search(r'TOTAL\s*A\s*PAGAR\s*AL\s*FONDO\s*DE\s*CESANTIA\s*([0-9\.\,]+)', all_text, re.IGNORECASE)
            
            doc_info["resumen_totales"] = {
                "remuneracion_pension": parse_chilean_number(rem_pen.group(1)) if rem_pen else 0.0,
                "cotizacion_obligatoria": parse_chilean_number(cot_oblig.group(1)) if cot_oblig else 0.0,
                "sis_total": parse_chilean_number(sis_tot.group(1)) if sis_tot else 0.0,
                "total_pension": parse_chilean_number(tot_pension.group(1)) if tot_pension else 0.0,
                "remuneracion_afc": parse_chilean_number(rem_afc.group(1)) if rem_afc else 0.0,
                "cotizacion_afc_trabajador": parse_chilean_number(cot_afc_afil.group(1)) if cot_afc_afil else 0.0,
                "cotizacion_afc_empleador": parse_chilean_number(cot_afc_emp.group(1)) if cot_afc_emp else 0.0,
                "total_afc": parse_chilean_number(tot_afc.group(1)) if tot_afc else 0.0,
                "total_pagado": (parse_chilean_number(tot_pension.group(1)) if tot_pension else 0.0) + (parse_chilean_number(tot_afc.group(1)) if tot_afc else 0.0)
            }
            
            lineas = all_text.split("\n")
            for line in lineas:
                # Patrón detalle AFP:
                # RUT | Nombre | Rem Imponible Pension | Cotiz Oblig | SIS | APVI ... Rem AFC | Cotiz Afil | Cotiz Emp | Cod Mov
                rut_match = re.search(r'([0-9]{1,2}\.?[0-9]{3}\.?[0-9]{3}\s*-\s*[0-9kK])\s+([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+([0-9\.\,]+)\s+[0-9\.\,\s]+?([0-9\.\,]+)\s+([0-9\.\,]+)\s+([0-9\.\,]+)(?:\s+([0-9]{1,2}))?', line)
                if rut_match and "TOTAL" not in line.upper():
                    r_raw = rut_match.group(1)
                    n_raw = rut_match.group(2).strip()
                    rem_pen_w = parse_chilean_number(rut_match.group(3))
                    cot_afp_w = parse_chilean_number(rut_match.group(4))
                    sis_w = parse_chilean_number(rut_match.group(5))
                    rem_afc_w = parse_chilean_number(rut_match.group(6))
                    cot_afc_w = parse_chilean_number(rut_match.group(7))
                    cot_afc_emp_w = parse_chilean_number(rut_match.group(8))
                    cod_mov = int(rut_match.group(9)) if rut_match.group(9) else 0
                    
                    doc_info["trabajadores"].append({
                        "rut": clean_rut(r_raw),
                        "nombre": n_raw,
                        "afp": f"AFP {nom_afp}",
                        "remuneracion_pension": rem_pen_w,
                        "cotizacion_afp": cot_afp_w,
                        "sis": sis_w,
                        "remuneracion_afc": rem_afc_w,
                        "cotizacion_afc_trabajador": cot_afc_w,
                        "cotizacion_afc_empleador": cot_afc_emp_w,
                        "cod_movimiento": cod_mov
                    })
                    
    return doc_info


def consolidate_previred_batch(parsed_docs: List[Dict[str, Any]]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Consolida un conjunto de comprobantes PDF de Previred de un mismo mes.
    Agrupa por RUT de trabajador unificando FONASA, IPS, AFPs, Mutual y Seguro Social.
    Retorna (DataFrame de trabajadores consolidado, Diccionario de cuadratura/resumen mensual).
    """
    workers_dict: Dict[str, Dict[str, Any]] = {}
    
    # Acumuladores de cuadratura Previred (Totales declarados/pagados según carátulas)
    resumen_previred = {
        "total_fonasa_pagado": 0.0,
        "total_ips_asig_fam": 0.0,
        "total_seguro_social_pagado": 0.0,
        "total_isl_pagado": 0.0,
        "total_afp_pension_pagado": 0.0,
        "total_afp_afc_pagado": 0.0,
        "total_general_previred_pagado": 0.0,
        "docs_procesados": len(parsed_docs),
        "instituciones": []
    }
    
    for doc in parsed_docs:
        doc_type = doc["tipo_doc"]
        res_tot = doc["resumen_totales"]
        inst = doc["institucion"]
        if inst and inst not in resumen_previred["instituciones"]:
            resumen_previred["instituciones"].append(inst)
            
        if doc_type == "FONASA":
            resumen_previred["total_fonasa_pagado"] += res_tot.get("total_pagado", 0.0)
        elif doc_type == "IPS":
            resumen_previred["total_ips_asig_fam"] += res_tot.get("asignacion_familiar", 0.0)
        elif doc_type == "SEGURO_SOCIAL":
            resumen_previred["total_seguro_social_pagado"] += res_tot.get("total_pagado", 0.0)
        elif doc_type == "MUTUAL":
            resumen_previred["total_isl_pagado"] += res_tot.get("total_pagado", 0.0)
        elif doc_type.startswith("AFP_"):
            resumen_previred["total_afp_pension_pagado"] += res_tot.get("total_pension", 0.0)
            resumen_previred["total_afp_afc_pagado"] += res_tot.get("total_afc", 0.0)
            
        # Procesar detalle de cada trabajador
        for w in doc["trabajadores"]:
            rut = clean_rut(w.get("rut", ""))
            if not rut:
                continue
                
            if rut not in workers_dict:
                workers_dict[rut] = {
                    "rut": rut,
                    "rut_formateado": format_rut_dots(rut),
                    "nombre": w.get("nombre", ""),
                    "dias_trabajados": 30,
                    "sueldo_base_previred": 0.0,
                    "bono_adicional": 0.0,
                    "sueldo_bruto": 0.0,
                    "afp": "AFP MODELO",
                    "cotizacion_afp": 0.0,
                    "salud_entidad": "FONASA",
                    "cotizacion_salud": 0.0,
                    "afc_trabajador": 0.0,
                    "cotizacion_previsional_total": 0.0,
                    "seguro_social_01": 0.0,
                    "impuesto_unico": 0.0,
                    "renta_neta_pagada": 0.0,
                    "asignacion_familiar": 0.0,
                    "sis": 0.0,
                    "afc_empleador": 0.0,
                    "isl_mutual": 0.0,
                    "seguro_social_patronal": 0.0,
                    "gastos_patronales_total": 0.0,
                    "cod_movimiento": 0,
                    "movimiento_desc": "Sin movimientos",
                    "alerta_movimiento": ""
                }
                
            entry = workers_dict[rut]
            
            # Actualizar nombre si está más completo
            if len(w.get("nombre", "")) > len(entry["nombre"]):
                entry["nombre"] = w.get("nombre", "")
                
            if "dias_trabajados" in w and w["dias_trabajados"] is not None:
                entry["dias_trabajados"] = w["dias_trabajados"]
                
            if "cod_movimiento" in w and w["cod_movimiento"] > 0:
                entry["cod_movimiento"] = w["cod_movimiento"]
                info_mov = CODIGOS_MOVIMIENTO.get(w["cod_movimiento"], {"desc": f"Código {w['cod_movimiento']}", "alerta": "Movimiento registrado"})
                entry["movimiento_desc"] = info_mov["desc"]
                entry["alerta_movimiento"] = info_mov.get("alerta") or ""
                
            # Datos FONASA
            if "cotizacion_salud" in w:
                entry["cotizacion_salud"] = w["cotizacion_salud"]
                if entry["sueldo_base_previred"] == 0.0 and w.get("remuneracion_salud", 0.0) > 0:
                    entry["sueldo_base_previred"] = w["remuneracion_salud"]
                    
            # Datos IPS
            if "asignacion_familiar" in w:
                entry["asignacion_familiar"] = w["asignacion_familiar"]
                
            # Datos Seguro Social Previsional
            if "seguro_social_patronal" in w:
                entry["seguro_social_patronal"] = w["seguro_social_patronal"]
            if "seguro_social_01" in w:
                entry["seguro_social_01"] = w["seguro_social_01"]
            if entry["sueldo_base_previred"] == 0.0 and w.get("remuneracion_imponible_ss", 0.0) > 0:
                entry["sueldo_base_previred"] = w["remuneracion_imponible_ss"]
                
            # Datos Mutual / ISL
            if "cotizacion_isl" in w:
                entry["isl_mutual"] = w["cotizacion_isl"]
            if entry["sueldo_base_previred"] == 0.0 and w.get("remuneracion_isl", 0.0) > 0:
                entry["sueldo_base_previred"] = w["remuneracion_isl"]
                
            # Datos AFP
            if "afp" in w:
                entry["afp"] = w["afp"]
            if "cotizacion_afp" in w:
                entry["cotizacion_afp"] = w["cotizacion_afp"]
            if "sis" in w:
                entry["sis"] = w["sis"]
            if "cotizacion_afc_trabajador" in w:
                entry["afc_trabajador"] = w["cotizacion_afc_trabajador"]
            if "cotizacion_afc_empleador" in w:
                entry["afc_empleador"] = w["cotizacion_afc_empleador"]
            if entry["sueldo_base_previred"] == 0.0 and w.get("remuneracion_pension", 0.0) > 0:
                entry["sueldo_base_previred"] = w["remuneracion_pension"]
                
            # Total Imponible Bruto = Base Previred + Bono Adicional
            entry["sueldo_bruto"] = round(entry["sueldo_base_previred"] + entry["bono_adicional"], 3)
                
    # Cálculo derivado final por trabajador
    for rut, e in workers_dict.items():
        # Cotización previsional = AFP + Salud (Fonasa 7%) + AFC trabajador
        e["cotizacion_previsional_total"] = round(e["cotizacion_afp"] + e["cotizacion_salud"] + e["afc_trabajador"], 3)
        
        # Seguro Social 0.1% (si no vino del formulario, es el 0.1% del sueldo imponible)
        if e["seguro_social_01"] == 0.0 and e["sueldo_bruto"] > 0:
            e["seguro_social_01"] = round(e["sueldo_bruto"] * 0.001, 3)
            
        # Renta Neta Pagada = Sueldo Bruto - Cotiz Previs - Seguro Social 0.1% - Impuesto Único
        e["renta_neta_pagada"] = round(e["sueldo_bruto"] - e["cotizacion_previsional_total"] - e["seguro_social_01"] - e["impuesto_unico"], 3)
        
        # Gastos Patronales Totales = SIS + AFC Empleador + ISL/Mutual + Seguro Social Patronal + SS 0.1%
        e["gastos_patronales_total"] = round(e["sis"] + e["afc_empleador"] + e["isl_mutual"] + e["seguro_social_patronal"] + e["seguro_social_01"], 3)

    # Calcular Total General Previred Pagado
    resumen_previred["total_general_previred_pagado"] = (
        resumen_previred["total_fonasa_pagado"] +
        resumen_previred["total_seguro_social_pagado"] +
        resumen_previred["total_isl_pagado"] +
        resumen_previred["total_afp_pension_pagado"] +
        resumen_previred["total_afp_afc_pagado"]
    )

    df_result = pd.DataFrame(list(workers_dict.values()))
    return df_result, resumen_previred


def generate_sample_previred_july_2026() -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Genera la nómina consolidada exacta de ejemplo basada en los comprobantes
    reales de julio 2026 de IMPORTADORA DONG SHENG LTDA (RUT 76.519.206-4).
    Permite probar la aplicación de inmediato sin requerir subida inicial de archivos.
    """
    raw_data = [
        {"rut": "26578630-8", "nombre": "SHEN HUI", "dias": 30, "bruto": 691941, "afp_nom": "AFP PLANVITAL", "afp_cot": 77913, "salud": 48436, "afc_trab": 4152, "ss_01": 691.941, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 16607, "isl": 6435, "ss_pat": 6227, "mov": 0},
        {"rut": "18441583-6", "nombre": "FUENTES IBARRA JENIFFER VALESKA", "dias": 30, "bruto": 1719115, "afp_nom": "AFP MODELO", "afp_cot": 183601, "salud": 120338, "afc_trab": 10315, "ss_01": 1719.115, "iu": 17573, "asig_fam": 0, "sis": 34382, "afc_emp": 41259, "isl": 15988, "ss_pat": 15472, "mov": 0},
        {"rut": "24364035-0", "nombre": "SHEN XUAN", "dias": 30, "bruto": 1719115, "afp_nom": "AFP MODELO", "afp_cot": 183601, "salud": 120338, "afc_trab": 10315, "ss_01": 1719.115, "iu": 17573, "asig_fam": 0, "sis": 34382, "afc_emp": 41259, "isl": 15988, "ss_pat": 15472, "mov": 0},
        {"rut": "26879752-1", "nombre": "SANTANA COLMENARES LIZ MARBELLA", "dias": 30, "bruto": 691941, "afp_nom": "AFP MODELO", "afp_cot": 73899, "salud": 48436, "afc_trab": 4152, "ss_01": 691.941, "iu": 0, "asig_fam": 13870, "sis": 13839, "afc_emp": 16607, "isl": 6435, "ss_pat": 6227, "mov": 0},
        {"rut": "19842504-4", "nombre": "JARAMILLO DINAMARCA ESTEFANIA ANDREA", "dias": 30, "bruto": 1000000, "afp_nom": "AFP PLANVITAL", "afp_cot": 112600, "salud": 70000, "afc_trab": 6000, "ss_01": 1000.0, "iu": 0, "asig_fam": 0, "sis": 20000, "afc_emp": 24000, "isl": 9300, "ss_pat": 9000, "mov": 0},
        {"rut": "17054567-2", "nombre": "FUENTES IBARRA MARÍA JOSE", "dias": 0, "bruto": 0, "afp_nom": "AFP UNO", "afp_cot": 0, "salud": 0, "afc_trab": 0, "ss_01": 0.0, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 16607, "isl": 208, "ss_pat": 6227, "mov": 3},
        {"rut": "44336047-6", "nombre": "AMAYA ACURERO ASHLEY PAOLA", "dias": 0, "bruto": 0, "afp_nom": "AFP UNO", "afp_cot": 0, "salud": 0, "afc_trab": 0, "ss_01": 0.0, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 16607, "isl": 208, "ss_pat": 6227, "mov": 3},
        {"rut": "44336008-5", "nombre": "BLANCO PEROZA DIANA MARILY", "dias": 30, "bruto": 691941, "afp_nom": "AFP UNO", "afp_cot": 73069, "salud": 48436, "afc_trab": 4152, "ss_01": 691.941, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 16607, "isl": 6435, "ss_pat": 6227, "mov": 0},
        {"rut": "41450884-7", "nombre": "OJEDA BRACHO KEIDIMAR", "dias": 29, "bruto": 668876, "afp_nom": "AFP UNO", "afp_cot": 70633, "salud": 46821, "afc_trab": 0, "ss_01": 668.876, "iu": 0, "asig_fam": 0, "sis": 13378, "afc_emp": 20066, "isl": 6221, "ss_pat": 6020, "mov": 11},
        {"rut": "29117027-7", "nombre": "SANTANA COLMENARES EDGAR ALEXANDER", "dias": 30, "bruto": 691942, "afp_nom": "AFP UNO", "afp_cot": 73069, "salud": 48436, "afc_trab": 4152, "ss_01": 691.942, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 16607, "isl": 6435, "ss_pat": 6227, "mov": 0},
        {"rut": "28853911-1", "nombre": "SEQUERA PEÑA WILMARY ALEJANDRA", "dias": 30, "bruto": 691941, "afp_nom": "AFP UNO", "afp_cot": 73069, "salud": 48436, "afc_trab": 0, "ss_01": 691.941, "iu": 0, "asig_fam": 0, "sis": 13839, "afc_emp": 20758, "isl": 6435, "ss_pat": 6227, "mov": 0},
        {"rut": "26412680-0", "nombre": "PEREZ RONDON YNES NATALYS", "dias": 29, "bruto": 668876, "afp_nom": "AFP MODELO", "afp_cot": 71436, "salud": 46821, "afc_trab": 0, "ss_01": 668.876, "iu": 0, "asig_fam": 0, "sis": 13378, "afc_emp": 20066, "isl": 6221, "ss_pat": 6020, "mov": 11},
    ]
    
    rows = []
    for r in raw_data:
        cot_prev = round(r["afp_cot"] + r["salud"] + r["afc_trab"], 3)
        renta_neta = round(r["bruto"] - cot_prev - r["ss_01"] - r["iu"], 3)
        gastos_pat = round(r["sis"] + r["afc_emp"] + r["isl"] + r["ss_pat"] + r["ss_01"], 3)
        info_mov = CODIGOS_MOVIMIENTO.get(r["mov"], {"desc": "Normal", "alerta": ""})
        
        rows.append({
            "rut": clean_rut(r["rut"]),
            "rut_formateado": format_rut_dots(r["rut"]),
            "nombre": r["nombre"],
            "dias_trabajados": r["dias"],
            "sueldo_base_previred": float(r["bruto"]),
            "bono_adicional": 0.0,
            "sueldo_bruto": float(r["bruto"]),
            "afp": r["afp_nom"],
            "cotizacion_afp": float(r["afp_cot"]),
            "salud_entidad": "FONASA",
            "cotizacion_salud": float(r["salud"]),
            "afc_trabajador": float(r["afc_trab"]),
            "cotizacion_previsional_total": cot_prev,
            "seguro_social_01": float(r["ss_01"]),
            "impuesto_unico": float(r["iu"]),
            "renta_neta_pagada": renta_neta,
            "asignacion_familiar": float(r["asig_fam"]),
            "sis": float(r["sis"]),
            "afc_empleador": float(r["afc_emp"]),
            "isl_mutual": float(r["isl"]),
            "seguro_social_patronal": float(r["ss_pat"]),
            "gastos_patronales_total": gastos_pat,
            "cod_movimiento": r["mov"],
            "movimiento_desc": info_mov["desc"],
            "alerta_movimiento": info_mov.get("alerta") or ""
        })
        
    resumen = {
        "total_fonasa_pagado": 632628.0,
        "total_ips_asig_fam": 13870.0,
        "total_seguro_social_pagado": 95573.0,
        "total_isl_pagado": 86309.0,
        "total_afp_pension_pagado": 372413.0 + 608518.0 + 224352.0, # UNO + Modelo + PlanVital
        "total_afp_afc_pagado": 115556.0 + 143973.0 + 50759.0,
        "total_general_previred_pagado": 2249569.0,
        "docs_procesados": 6,
        "instituciones": ["FONASA", "IPS", "ISL Mutual", "Seguro Social", "AFP UNO", "AFP Modelo", "AFP PlanVital"]
    }
    
    return pd.DataFrame(rows), resumen
