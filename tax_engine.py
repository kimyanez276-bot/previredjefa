"""
Módulo del Motor Tributario y Normativo Chileno (Previred y SII)
Soporte para Factores de Actualización (IPC), Topes Imponibles UF,
Códigos de Movimientos de Personal y Estructura DJ 1887.
"""

from typing import Dict, List, Tuple, Optional
import re

# Factores de Actualización Oficiales SII por defecto
# Para Año Tributario 2026 (Año Comercial 2025)
FACTORES_ACTUALIZACION_2026 = {
    1: 1.036,  # Enero
    2: 1.026,  # Febrero
    3: 1.022,  # Marzo
    4: 1.016,  # Abril
    5: 1.014,  # Mayo
    6: 1.012,  # Junio
    7: 1.017,  # Julio
    8: 1.008,  # Agosto
    9: 1.007,  # Septiembre
    10: 1.003, # Octubre
    11: 1.003, # Noviembre
    12: 1.000  # Diciembre
}

# Factores para Año Tributario 2025 (Año Comercial 2024)
FACTORES_ACTUALIZACION_2025 = {
    1: 1.042,
    2: 1.038,
    3: 1.032,
    4: 1.028,
    5: 1.024,
    6: 1.020,
    7: 1.016,
    8: 1.012,
    9: 1.009,
    10: 1.005,
    11: 1.002,
    12: 1.000
}

MESES_NOMBRE = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

MESES_MAP = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# Topes Imponibles en UF (Superintendencia de Pensiones)
TOPES_UF = {
    2024: {"afp_salud": 84.3, "afc": 126.6, "uf_promedio": 37200.0},
    2025: {"afp_salud": 87.8, "afc": 131.7, "uf_promedio": 38400.0},
    2026: {"afp_salud": 87.8, "afc": 131.7, "uf_promedio": 39500.0}
}

# Diccionario Oficial de Códigos de Movimientos de Personal Previred
CODIGOS_MOVIMIENTO = {
    0: {"desc": "Sin movimientos en el mes", "alerta": None, "tipo": "normal"},
    1: {"desc": "Iniciación de servicios (Contrato indefinido)", "alerta": "🎉 Nueva Contratación (Indefinido)", "tipo": "alta"},
    2: {"desc": "Cesación de servicios prestados (Finiquito)", "alerta": "⚠️ Finiquito / Desvinculación", "tipo": "baja"},
    3: {"desc": "Subsidios (Licencia médica)", "alerta": "🏥 Subsidio Licencia Médica", "tipo": "licencia"},
    4: {"desc": "Permiso sin goce de remuneraciones", "alerta": "📋 Permiso Sin Goce", "tipo": "permiso"},
    5: {"desc": "Incorporación en el lugar de trabajo", "alerta": "🏢 Incorporación Lugar Trabajo", "tipo": "normal"},
    6: {"desc": "Accidente del trabajo", "alerta": "🚨 Accidente Laboral", "tipo": "licencia"},
    7: {"desc": "Iniciación de servicios (Contrato a plazo fijo)", "alerta": "🎉 Nueva Contratación (Plazo Fijo)", "tipo": "alta"},
    8: {"desc": "Cambio de contrato plazo fijo a indefinido", "alerta": "📝 Pase a Indefinido", "tipo": "cambio"},
    9: {"desc": "Trabajador Part-Time", "alerta": "⏱️ Jornada Parcial", "tipo": "normal"},
    11: {"desc": "Otros movimientos (Ausentismo)", "alerta": "⚠️ Ausentismo / Días no trabajados", "tipo": "ausentismo"},
    12: {"desc": "Reliquidación, premio o bono posterior al finiquito", "alerta": "💰 Bono Posterior Finiquito", "tipo": "reliquidacion"},
    13: {"desc": "Suspensión de contrato por acto de autoridad", "alerta": "⏸️ Suspensión Autoridad", "tipo": "suspension"},
    14: {"desc": "Suspensión de contrato por pacto", "alerta": "⏸️ Suspensión Pactada", "tipo": "suspension"},
    15: {"desc": "Reducción de jornada", "alerta": "⏱️ Reducción Jornada", "tipo": "cambio"},
    20: {"desc": "Inicio relación laboral", "alerta": "🎉 Inicio Relación Laboral", "tipo": "alta"}
}


def clean_rut(rut_str: str) -> str:
    """Limpia y estandariza un RUT chileno (ej: '76.519.206-4' -> '76519206-4')"""
    if not rut_str:
        return ""
    rut_clean = re.sub(r'[^0-9kK]', '', str(rut_str)).upper()
    if len(rut_clean) > 1:
        cuerpo = rut_clean[:-1]
        dv = rut_clean[-1]
        return f"{cuerpo}-{dv}"
    return rut_clean


def format_rut_dots(rut_str: str) -> str:
    """Formatea un RUT chileno con puntos y guion (ej: '76519206-4' -> '76.519.206-4')"""
    rut = clean_rut(rut_str)
    if "-" not in rut:
        return rut
    cuerpo, dv = rut.split("-")
    try:
        cuerpo_num = int(cuerpo)
        cuerpo_format = f"{cuerpo_num:,}".replace(",", ".")
        return f"{cuerpo_format}-{dv}"
    except ValueError:
        return rut


def get_factor_actualizacion(mes: int, anio_tributario: int = 2026, custom_factors: Optional[Dict[int, float]] = None) -> float:
    """Obtiene el factor de corrección monetaria para el mes especificado."""
    if custom_factors and mes in custom_factors:
        return float(custom_factors[mes])
    if anio_tributario == 2026:
        return FACTORES_ACTUALIZACION_2026.get(mes, 1.0)
    elif anio_tributario == 2025:
        return FACTORES_ACTUALIZACION_2025.get(mes, 1.0)
    return 1.0


def check_topes_imponibles(remuneracion_imponible: float, anio: int = 2026, valor_uf: Optional[float] = None) -> Dict:
    """
    Verifica si una remuneración imponible sobrepasa los topes legales de AFP y AFC.
    """
    info_topes = TOPES_UF.get(anio, TOPES_UF[2026])
    uf = valor_uf if valor_uf and valor_uf > 0 else info_topes["uf_promedio"]
    
    tope_afp_pesos = round(info_topes["afp_salud"] * uf)
    tope_afc_pesos = round(info_topes["afc"] * uf)
    
    excede_afp = remuneracion_imponible > tope_afp_pesos
    excede_afc = remuneracion_imponible > tope_afc_pesos
    
    return {
        "valor_uf": uf,
        "tope_afp_pesos": tope_afp_pesos,
        "tope_afc_pesos": tope_afc_pesos,
        "excede_afp": excede_afp,
        "excede_afc": excede_afc,
        "exceso_afp": max(0, remuneracion_imponible - tope_afp_pesos) if excede_afp else 0,
        "exceso_afc": max(0, remuneracion_imponible - tope_afc_pesos) if excede_afc else 0
    }


def parse_chilean_number(val) -> float:
    """Parsea números en formato chileno con separadores de miles y decimales."""
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("$", "").replace(" ", "")
    if not val_str:
        return 0.0
    # Caso 1: Formato con punto de miles y coma decimal '1.719.115,50'
    if "." in val_str and "," in val_str:
        val_str = val_str.replace(".", "").replace(",", ".")
    # Caso 2: Formato Previred con punto decimal '129809.059'
    elif val_str.count(".") == 1 and "," not in val_str:
        # Podría ser decimal o miles (ej. 691.941 vs 129809.059)
        partes = val_str.split(".")
        if len(partes[1]) == 3 and len(partes[0]) <= 3:
            # Muy probablemente miles: ej. 691.941
            val_str = val_str.replace(".", "")
        else:
            # Decimal: 129809.059 o 7235.415
            pass
    elif val_str.count(".") > 1:
        # Separadores de miles: 9.235.687
        val_str = val_str.replace(".", "")
    elif "," in val_str:
        val_str = val_str.replace(",", ".")
    
    try:
        return float(val_str)
    except ValueError:
        return 0.0
