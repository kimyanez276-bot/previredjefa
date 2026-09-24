"""
Módulo de Almacenamiento Local (SQLite) y Sincronización en la Nube (Google Sheets).
Gestión de Multi-Empresas, Libro de Remuneraciones, Registro de Auditoría de Ediciones
y Respaldos Portátiles en JSON.
"""

import sqlite3
import os
import json
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd

from modules.tax_engine import clean_rut, format_rut_dots

DB_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
DB_PATH = os.path.join(DB_DIR, "conta_marlene.db")


def get_connection() -> sqlite3.Connection:
    """Retorna una conexión a la base de datos SQLite."""
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Crea las tablas necesarias si no existen."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Tabla de Empresas Clientes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS empresas (
        rut TEXT PRIMARY KEY,
        razon_social TEXT NOT NULL,
        direccion TEXT,
        contacto TEXT,
        creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # 2. Tabla de Remuneraciones Mensuales Detalladas
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS remuneraciones_mensuales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_rut TEXT NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        rut_trabajador TEXT NOT NULL,
        nombre_trabajador TEXT NOT NULL,
        dias_trabajados INTEGER DEFAULT 30,
        sueldo_base_previred REAL DEFAULT 0.0,
        bono_adicional REAL DEFAULT 0.0,
        sueldo_bruto REAL DEFAULT 0.0,
        afp_nombre TEXT DEFAULT 'AFP MODELO',
        cotizacion_afp REAL DEFAULT 0.0,
        salud_entidad TEXT DEFAULT 'FONASA',
        cotizacion_salud REAL DEFAULT 0.0,
        afc_trabajador REAL DEFAULT 0.0,
        cotizacion_previsional_total REAL DEFAULT 0.0,
        seguro_social_01 REAL DEFAULT 0.0,
        impuesto_unico REAL DEFAULT 0.0,
        renta_neta_pagada REAL DEFAULT 0.0,
        asignacion_familiar REAL DEFAULT 0.0,
        sis REAL DEFAULT 0.0,
        afc_empleador REAL DEFAULT 0.0,
        isl_mutual REAL DEFAULT 0.0,
        seguro_social_patronal REAL DEFAULT 0.0,
        gastos_patronales_total REAL DEFAULT 0.0,
        cod_movimiento INTEGER DEFAULT 0,
        movimiento_desc TEXT DEFAULT 'Sin movimientos',
        alerta_movimiento TEXT DEFAULT '',
        actualizado_por TEXT DEFAULT 'Sistema',
        actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(empresa_rut, anio, mes, rut_trabajador)
    )
    """)

    # Migraciones en caliente para bases de datos existentes
    try:
        cursor.execute("ALTER TABLE remuneraciones_mensuales ADD COLUMN sueldo_base_previred REAL DEFAULT 0.0")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE remuneraciones_mensuales ADD COLUMN bono_adicional REAL DEFAULT 0.0")
    except sqlite3.OperationalError:
        pass
    
    # 3. Tabla de Auditoría de Modificaciones
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS auditoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        usuario TEXT NOT NULL,
        empresa_rut TEXT NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        rut_trabajador TEXT NOT NULL,
        campo_modificado TEXT NOT NULL,
        valor_anterior TEXT,
        valor_nuevo TEXT
    )
    """)
    
    # 4. Tabla de Resumen y Cuadratura Previred por Mes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cuadratura_previred (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        empresa_rut TEXT NOT NULL,
        anio INTEGER NOT NULL,
        mes INTEGER NOT NULL,
        total_previred_comprobantes REAL DEFAULT 0.0,
        total_nomina_calculada REAL DEFAULT 0.0,
        diferencia REAL DEFAULT 0.0,
        estado_cuadratura TEXT DEFAULT 'CUADRADO',
        instituciones TEXT DEFAULT '',
        actualizado_por TEXT DEFAULT 'Sistema',
        actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(empresa_rut, anio, mes)
    )
    """)
    
    # Insertar empresa de prueba si no existe
    cursor.execute("SELECT COUNT(*) FROM empresas")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
        INSERT INTO empresas (rut, razon_social, direccion, contacto)
        VALUES ('76519206-4', 'IMPORTADORA DONG SHENG LTDA', 'MAIPU 616, LINARES', 'serviciocontablelinares@gmail.com')
        """)
        cursor.execute("""
        INSERT INTO empresas (rut, razon_social, direccion, contacto)
        VALUES ('15151945-8', 'ALEX ALEJANDRO HERNANDEZ ROJAS', 'BERNARDO O''HIGGINS 968', 'contacto@hernandez.cl')
        """)
        
    conn.commit()
    conn.close()


def get_empresas() -> List[Dict[str, Any]]:
    """Retorna la lista de empresas registradas."""
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT rut, razon_social, direccion, contacto FROM empresas ORDER BY razon_social ASC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def add_empresa(rut: str, razon_social: str, direccion: str = "", contacto: str = "") -> bool:
    """Registra una nueva empresa."""
    init_db()
    c_rut = clean_rut(rut)
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO empresas (rut, razon_social, direccion, contacto)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(rut) DO UPDATE SET
            razon_social=excluded.razon_social,
            direccion=excluded.direccion,
            contacto=excluded.contacto
        """, (c_rut, razon_social.strip().upper(), direccion.strip(), contacto.strip()))
        conn.commit()
        return True
    finally:
        conn.close()


def get_remuneraciones_mes(empresa_rut: str, anio: int, mes: int) -> pd.DataFrame:
    """Obtiene el DataFrame de remuneraciones de un mes específico."""
    init_db()
    c_rut = clean_rut(empresa_rut)
    conn = get_connection()
    query = """
    SELECT 
        rut_trabajador as rut,
        nombre_trabajador as nombre,
        dias_trabajados,
        COALESCE(sueldo_base_previred, sueldo_bruto) as sueldo_base_previred,
        COALESCE(bono_adicional, 0.0) as bono_adicional,
        sueldo_bruto,
        afp_nombre as afp,
        cotizacion_afp,
        salud_entidad,
        cotizacion_salud,
        afc_trabajador,
        cotizacion_previsional_total,
        seguro_social_01,
        impuesto_unico,
        renta_neta_pagada,
        asignacion_familiar,
        sis,
        afc_empleador,
        isl_mutual,
        seguro_social_patronal,
        gastos_patronales_total,
        cod_movimiento,
        movimiento_desc,
        alerta_movimiento
    FROM remuneraciones_mensuales
    WHERE empresa_rut = ? AND anio = ? AND mes = ?
    ORDER BY nombre_trabajador ASC
    """
    df = pd.read_sql_query(query, conn, params=(c_rut, anio, mes))
    conn.close()
    if not df.empty:
        df["rut_formateado"] = df["rut"].apply(format_rut_dots)
    return df


def save_remuneraciones_mes(
    empresa_rut: str, 
    anio: int, 
    mes: int, 
    df: pd.DataFrame, 
    usuario: str = "Analista",
    resumen_previred: Optional[Dict[str, Any]] = None
) -> int:
    """
    Guarda o actualiza la nómina de trabajadores del mes en SQLite.
    Registra en la tabla de auditoría los campos que hayan sido modificados por las usuarias.
    """
    init_db()
    c_rut = clean_rut(empresa_rut)
    conn = get_connection()
    cursor = conn.cursor()
    
    # Obtener valores existentes para auditoría
    existing_df = get_remuneraciones_mes(c_rut, anio, mes)
    existing_dict = {}
    if not existing_df.empty:
        for _, row in existing_df.iterrows():
            existing_dict[clean_rut(row["rut"])] = row.to_dict()
            
    saved_count = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for _, row in df.iterrows():
        w_rut = clean_rut(str(row.get("rut", "")))
        if not w_rut:
            continue
            
        nombre = str(row.get("nombre", "")).strip().upper()
        dias = int(row.get("dias_trabajados", 30))
        bono_adicional = float(row.get("bono_adicional", 0.0))
        sueldo_base = float(row.get("sueldo_base_previred", row.get("sueldo_bruto", 0.0)))
        
        # El sueldo bruto total es la base de Previred + Bono / Haber Adicional
        sueldo_bruto = float(row.get("sueldo_bruto", sueldo_base + bono_adicional))
        if bono_adicional > 0 and abs(sueldo_bruto - (sueldo_base + bono_adicional)) > 0.01:
            sueldo_bruto = round(sueldo_base + bono_adicional, 3)
            
        afp = str(row.get("afp", "AFP MODELO"))
        cot_afp = float(row.get("cotizacion_afp", 0.0))
        salud = str(row.get("salud_entidad", "FONASA"))
        cot_salud = float(row.get("cotizacion_salud", 0.0))
        afc_trab = float(row.get("afc_trabajador", 0.0))
        cot_prev = float(row.get("cotizacion_previsional_total", cot_afp + cot_salud + afc_trab))
        ss_01 = float(row.get("seguro_social_01", round(sueldo_bruto * 0.001, 3)))
        iu = float(row.get("impuesto_unico", 0.0))
        renta_neta = float(row.get("renta_neta_pagada", sueldo_bruto - cot_prev - ss_01 - iu))
        asig_fam = float(row.get("asignacion_familiar", 0.0))
        sis = float(row.get("sis", 0.0))
        afc_emp = float(row.get("afc_empleador", 0.0))
        isl = float(row.get("isl_mutual", 0.0))
        ss_pat = float(row.get("seguro_social_patronal", 0.0))
        gastos_pat = float(row.get("gastos_patronales_total", sis + afc_emp + isl + ss_pat + ss_01))
        cod_mov = int(row.get("cod_movimiento", 0))
        mov_desc = str(row.get("movimiento_desc", "Sin movimientos"))
        alerta_mov = str(row.get("alerta_movimiento", ""))
        
        # Comparar con existente para auditoría
        if w_rut in existing_dict:
            old = existing_dict[w_rut]
            campos_auditar = [
                ("bono_adicional", bono_adicional),
                ("sueldo_bruto", sueldo_bruto),
                ("impuesto_unico", iu),
                ("dias_trabajados", dias),
                ("renta_neta_pagada", renta_neta),
                ("asignacion_familiar", asig_fam),
                ("cotizacion_previsional_total", cot_prev)
            ]
            for col_name, new_val in campos_auditar:
                old_val = old.get(col_name)
                if old_val is not None and abs(float(old_val) - float(new_val)) > 0.01:
                    cursor.execute("""
                    INSERT INTO auditoria (fecha_hora, usuario, empresa_rut, anio, mes, rut_trabajador, campo_modificado, valor_anterior, valor_nuevo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (now_str, usuario, c_rut, anio, mes, w_rut, col_name, str(old_val), str(new_val)))
        else:
            # Nuevo registro ingresado
            cursor.execute("""
            INSERT INTO auditoria (fecha_hora, usuario, empresa_rut, anio, mes, rut_trabajador, campo_modificado, valor_anterior, valor_nuevo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (now_str, usuario, c_rut, anio, mes, w_rut, "NUEVO_TRABAJADOR", "No existía", f"{nombre} ({sueldo_bruto})"))
            
        cursor.execute("""
        INSERT INTO remuneraciones_mensuales (
            empresa_rut, anio, mes, rut_trabajador, nombre_trabajador, dias_trabajados,
            sueldo_base_previred, bono_adicional, sueldo_bruto, afp_nombre, cotizacion_afp, salud_entidad, cotizacion_salud, afc_trabajador,
            cotizacion_previsional_total, seguro_social_01, impuesto_unico, renta_neta_pagada, asignacion_familiar,
            sis, afc_empleador, isl_mutual, seguro_social_patronal, gastos_patronales_total,
            cod_movimiento, movimiento_desc, alerta_movimiento, actualizado_por, actualizado_en
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(empresa_rut, anio, mes, rut_trabajador) DO UPDATE SET
            nombre_trabajador=excluded.nombre_trabajador,
            dias_trabajados=excluded.dias_trabajados,
            sueldo_base_previred=excluded.sueldo_base_previred,
            bono_adicional=excluded.bono_adicional,
            sueldo_bruto=excluded.sueldo_bruto,
            afp_nombre=excluded.afp_nombre,
            cotizacion_afp=excluded.cotizacion_afp,
            salud_entidad=excluded.salud_entidad,
            cotizacion_salud=excluded.cotizacion_salud,
            afc_trabajador=excluded.afc_trabajador,
            cotizacion_previsional_total=excluded.cotizacion_previsional_total,
            seguro_social_01=excluded.seguro_social_01,
            impuesto_unico=excluded.impuesto_unico,
            renta_neta_pagada=excluded.renta_neta_pagada,
            asignacion_familiar=excluded.asignacion_familiar,
            sis=excluded.sis,
            afc_empleador=excluded.afc_empleador,
            isl_mutual=excluded.isl_mutual,
            seguro_social_patronal=excluded.seguro_social_patronal,
            gastos_patronales_total=excluded.gastos_patronales_total,
            cod_movimiento=excluded.cod_movimiento,
            movimiento_desc=excluded.movimiento_desc,
            alerta_movimiento=excluded.alerta_movimiento,
            actualizado_por=excluded.actualizado_por,
            actualizado_en=excluded.actualizado_en
        """, (
            c_rut, anio, mes, w_rut, nombre, dias,
            sueldo_base, bono_adicional, sueldo_bruto, afp, cot_afp, salud, cot_salud, afc_trab,
            cot_prev, ss_01, iu, renta_neta, asig_fam,
            sis, afc_emp, isl, ss_pat, gastos_pat,
            cod_mov, mov_desc, alerta_mov, usuario, now_str
        ))
        saved_count += 1
        
    # Guardar resumen de cuadratura Previred si fue provisto
    if resumen_previred:
        tot_prev = resumen_previred.get("total_general_previred_pagado", 0.0)
        tot_nom = float(df["cotizacion_previsional_total"].sum() + df["gastos_patronales_total"].sum()) if "cotizacion_previsional_total" in df else 0.0
        dif = round(tot_prev - tot_nom, 2)
        estado = "CUADRADO" if abs(dif) < 10.0 else "DESCUADRADO"
        insts = ", ".join(resumen_previred.get("instituciones", []))
        
        cursor.execute("""
        INSERT INTO cuadratura_previred (
            empresa_rut, anio, mes, total_previred_comprobantes, total_nomina_calculada, diferencia, estado_cuadratura, instituciones, actualizado_por, actualizado_en
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(empresa_rut, anio, mes) DO UPDATE SET
            total_previred_comprobantes=excluded.total_previred_comprobantes,
            total_nomina_calculada=excluded.total_nomina_calculada,
            diferencia=excluded.diferencia,
            estado_cuadratura=excluded.estado_cuadratura,
            instituciones=excluded.instituciones,
            actualizado_por=excluded.actualizado_por,
            actualizado_en=excluded.actualizado_en
        """, (c_rut, anio, mes, tot_prev, tot_nom, dif, estado, insts, usuario, now_str))

    conn.commit()
    conn.close()
    return saved_count


def get_remuneraciones_anuales(empresa_rut: str, anio: int) -> pd.DataFrame:
    """Obtiene todas las remuneraciones del año para la empresa seleccionada."""
    init_db()
    c_rut = clean_rut(empresa_rut)
    conn = get_connection()
    query = """
    SELECT 
        rut_trabajador as rut,
        nombre_trabajador as nombre,
        mes,
        anio,
        dias_trabajados,
        COALESCE(sueldo_base_previred, sueldo_bruto) as sueldo_base_previred,
        COALESCE(bono_adicional, 0.0) as bono_adicional,
        sueldo_bruto,
        afp_nombre as afp,
        cotizacion_afp,
        salud_entidad,
        cotizacion_salud,
        afc_trabajador,
        cotizacion_previsional_total,
        seguro_social_01,
        impuesto_unico,
        renta_neta_pagada,
        asignacion_familiar,
        sis,
        afc_empleador,
        isl_mutual,
        seguro_social_patronal,
        gastos_patronales_total,
        cod_movimiento,
        movimiento_desc,
        alerta_movimiento
    FROM remuneraciones_mensuales
    WHERE empresa_rut = ? AND anio = ?
    ORDER BY mes ASC, nombre_trabajador ASC
    """
    df = pd.read_sql_query(query, conn, params=(c_rut, anio))
    conn.close()
    if not df.empty and "rut" in df.columns:
        df["rut_formateado"] = df["rut"].apply(format_rut_dots)
    return df


def get_auditoria(empresa_rut: Optional[str] = None, anio: Optional[int] = None, mes: Optional[int] = None) -> pd.DataFrame:
    """Consulta el historial de auditoría de modificaciones."""
    init_db()
    conn = get_connection()
    query = "SELECT fecha_hora, usuario, empresa_rut, anio, mes, rut_trabajador, campo_modificado, valor_anterior, valor_nuevo FROM auditoria WHERE 1=1"
    params = []
    if empresa_rut:
        query += " AND empresa_rut = ?"
        params.append(clean_rut(empresa_rut))
    if anio:
        query += " AND anio = ?"
        params.append(anio)
    if mes:
        query += " AND mes = ?"
        params.append(mes)
    query += " ORDER BY id DESC LIMIT 500"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_cuadratura(empresa_rut: str, anio: int, mes: int) -> Optional[Dict[str, Any]]:
    """Consulta los datos de cuadratura Previred para el mes."""
    init_db()
    c_rut = clean_rut(empresa_rut)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT total_previred_comprobantes, total_nomina_calculada, diferencia, estado_cuadratura, instituciones, actualizado_en
    FROM cuadratura_previred
    WHERE empresa_rut = ? AND anio = ? AND mes = ?
    """, (c_rut, anio, mes))
    row = cursor.fetchone()
    conn.close()
    if row:
        return dict(row)
    return None


def export_backup_json() -> str:
    """Genera un archivo JSON con el respaldo total de la base de datos."""
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    
    backup = {
        "timestamp": datetime.now().isoformat(),
        "empresas": [],
        "remuneraciones": [],
        "auditoria": []
    }
    
    cursor.execute("SELECT * FROM empresas")
    backup["empresas"] = [dict(r) for r in cursor.fetchall()]
    
    cursor.execute("SELECT * FROM remuneraciones_mensuales")
    backup["remuneraciones"] = [dict(r) for r in cursor.fetchall()]
    
    cursor.execute("SELECT * FROM auditoria")
    backup["auditoria"] = [dict(r) for r in cursor.fetchall()]
    
    conn.close()
    return json.dumps(backup, ensure_ascii=False, indent=2)


def import_backup_json(json_content: str) -> Tuple[bool, str]:
    """Restaura la base de datos a partir de un respaldo JSON."""
    try:
        data = json.loads(json_content)
        init_db()
        conn = get_connection()
        cursor = conn.cursor()
        
        emp_count = 0
        rem_count = 0
        
        for e in data.get("empresas", []):
            cursor.execute("""
            INSERT INTO empresas (rut, razon_social, direccion, contacto)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(rut) DO UPDATE SET razon_social=excluded.razon_social
            """, (e["rut"], e["razon_social"], e.get("direccion", ""), e.get("contacto", "")))
            emp_count += 1
            
        for r in data.get("remuneraciones", []):
            cursor.execute("""
            INSERT INTO remuneraciones_mensuales (
                empresa_rut, anio, mes, rut_trabajador, nombre_trabajador, dias_trabajados,
                sueldo_bruto, afp_nombre, cotizacion_afp, salud_entidad, cotizacion_salud, afc_trabajador,
                cotizacion_previsional_total, seguro_social_01, impuesto_unico, renta_neta_pagada, asignacion_familiar,
                sis, afc_empleador, isl_mutual, seguro_social_patronal, gastos_patronales_total,
                cod_movimiento, movimiento_desc, alerta_movimiento, actualizado_por, actualizado_en
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(empresa_rut, anio, mes, rut_trabajador) DO UPDATE SET
                sueldo_bruto=excluded.sueldo_bruto,
                renta_neta_pagada=excluded.renta_neta_pagada
            """, (
                r["empresa_rut"], r["anio"], r["mes"], r["rut_trabajador"], r["nombre_trabajador"], r.get("dias_trabajados", 30),
                r.get("sueldo_bruto", 0.0), r.get("afp_nombre", "AFP MODELO"), r.get("cotizacion_afp", 0.0),
                r.get("salud_entidad", "FONASA"), r.get("cotizacion_salud", 0.0), r.get("afc_trabajador", 0.0),
                r.get("cotizacion_previsional_total", 0.0), r.get("seguro_social_01", 0.0), r.get("impuesto_unico", 0.0),
                r.get("renta_neta_pagada", 0.0), r.get("asignacion_familiar", 0.0), r.get("sis", 0.0),
                r.get("afc_empleador", 0.0), r.get("isl_mutual", 0.0), r.get("seguro_social_patronal", 0.0),
                r.get("gastos_patronales_total", 0.0), r.get("cod_movimiento", 0), r.get("movimiento_desc", "Sin movimientos"),
                r.get("alerta_movimiento", ""), r.get("actualizado_por", "Importación"), r.get("actualizado_en", datetime.now().isoformat())
            ))
            rem_count += 1
            
        conn.commit()
        conn.close()
        return True, f"Respaldo cargado con éxito: {emp_count} empresas y {rem_count} registros de sueldos."
    except Exception as ex:
        return False, f"Error al restaurar respaldo: {str(ex)}"


def sync_to_google_sheet(service_account_info: Dict[str, Any], spreadsheet_title_or_id: str, df: pd.DataFrame, sheet_name: str = "Nómina_Consolidada") -> Tuple[bool, str]:
    """
    Sincroniza el DataFrame consolidado directamente con una hoja de cálculo en Google Sheets
    usando la librería gspread y credenciales de Service Account (st.secrets).
    """
    try:
        import gspread
        gc = gspread.service_account_from_dict(service_account_info)
        
        try:
            sh = gc.open_by_key(spreadsheet_title_or_id)
        except Exception:
            sh = gc.open(spreadsheet_title_or_id)
            
        try:
            ws = sh.worksheet(sheet_name)
        except Exception:
            ws = sh.add_worksheet(title=sheet_name, rows=100, cols=30)
            
        # Limpiar y escribir cabeceras + filas
        ws.clear()
        data_to_write = [df.columns.values.tolist()] + df.astype(str).values.tolist()
        ws.update(values=data_to_write, range_name="A1")
        return True, f"¡Sincronizado con Google Sheets con éxito! ({len(df)} filas)"
    except Exception as e:
        return False, f"Error en sincronización Google Sheets: {str(e)}"
