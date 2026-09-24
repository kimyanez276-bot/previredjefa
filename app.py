"""
=============================================================================
DASHBOARD WEB COLABORATIVO: PREVIRED, LIBRO DE SUELDOS & DJ 1887 DEL SII
Oficina Contable Conta Marlene - Diseñado para 4 usuarias (Jefa y Analistas)
=============================================================================
"""

import os
import io
import json
from datetime import datetime
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Conta Marlene | Previred & DJ 1887 SII",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Importar módulos locales
from modules.tax_engine import (
    clean_rut, format_rut_dots, get_factor_actualizacion,
    check_topes_imponibles, MESES_MAP, MESES_NOMBRE, CODIGOS_MOVIMIENTO
)
from modules.parser_previred import (
    parse_previred_pdf, consolidate_previred_batch, generate_sample_previred_july_2026
)
from modules.storage import (
    init_db, get_empresas, add_empresa, get_remuneraciones_mes,
    save_remuneraciones_mes, get_remuneraciones_anuales, get_auditoria,
    get_cuadratura, export_backup_json, import_backup_json, sync_to_google_sheet
)
from modules.report_generator import (
    generate_excel_libro_sueldos_anual, generate_txt_dj1887,
    generate_pdf_certificado_6, generate_zip_all_certificados
)

# Inicializar Base de Datos SQLite
init_db()

# =============================================================================
# 1. GESTIÓN DE SESIÓN Y AUTENTICACIÓN
# =============================================================================

USUARIOS_AUTORIZADOS = {
    "marlene": {"nombre": "Marlene (Jefa / Remota)", "rol": "Administradora", "password": "123"},
    "analista1": {"nombre": "Camila (Analista 1)", "rol": "Analista", "password": "123"},
    "analista2": {"nombre": "Andrea (Analista 2)", "rol": "Analista", "password": "123"},
    "analista3": {"nombre": "Daniela (Analista 3)", "rol": "Analista", "password": "123"},
}

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.session_state["user_fullname"] = ""
    st.session_state["role"] = ""

def login_form():
    st.markdown("""
    <div style='text-align: center; padding: 25px 0 10px 0;'>
        <h1 style='color: #1F497D; margin-bottom: 5px;'>💼 Conta Marlene</h1>
        <h3 style='color: #555; font-weight: normal; margin-top: 0;'>Sistema Colaborativo de Previred, Libro de Sueldos y DJ 1887</h3>
        <p style='color: #777; font-size: 14px;'>Bienvenidas al panel de control contable-tributario.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.container(border=True):
            st.subheader("🔐 Iniciar Sesión")
            user_input = st.selectbox(
                "Seleccione su usuario:",
                options=list(USUARIOS_AUTORIZADOS.keys()),
                format_func=lambda u: f"{USUARIOS_AUTORIZADOS[u]['nombre']} [{USUARIOS_AUTORIZADOS[u]['rol']}]"
            )
            pass_input = st.text_input("Contraseña:", type="password", value="123")
            
            btn_login = st.button("🚀 Ingresar al Dashboard", use_container_width=True, type="primary")
            if btn_login:
                if pass_input == USUARIOS_AUTORIZADOS[user_input]["password"]:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = user_input
                    st.session_state["user_fullname"] = USUARIOS_AUTORIZADOS[user_input]["nombre"]
                    st.session_state["role"] = USUARIOS_AUTORIZADOS[user_input]["rol"]
                    st.success(f"¡Bienvenida, {st.session_state['user_fullname']}!")
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
        
        st.info("💡 **Nota para el equipo:** La contraseña por defecto es `123`. Marlene tiene permisos completos de Administradora para cierres anuales.")

if not st.session_state["authenticated"]:
    login_form()
    st.stop()


# =============================================================================
# 2. BARRA LATERAL: EMPRESA, AÑO, PERÍODO Y SINCRONIZACIÓN
# =============================================================================

with st.sidebar:
    st.markdown(f"""
    <div style='background: #F0F4F8; padding: 12px; border-radius: 8px; margin-bottom: 15px;'>
        <b style='color: #1F497D;'>👤 {st.session_state['user_fullname']}</b><br/>
        <small style='color: #555;'>Rol: <b>{st.session_state['role']}</b></small>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()
        
    st.markdown("---")
    st.subheader("🏢 Selección de Empresa")
    
    lista_empresas = get_empresas()
    if not lista_empresas:
        add_empresa("76519206-4", "IMPORTADORA DONG SHENG LTDA", "MAIPU 616, LINARES", "serviciocontablelinares@gmail.com")
        lista_empresas = get_empresas()
        
    empresa_opciones = {f"{format_rut_dots(e['rut'])} - {e['razon_social']}": e for e in lista_empresas}
    empresa_sel_label = st.selectbox("Empresa / Cliente:", options=list(empresa_opciones.keys()), index=0)
    empresa_actual = empresa_opciones[empresa_sel_label]
    
    with st.expander("➕ Registrar Nueva Empresa"):
        with st.form("form_nueva_empresa", clear_on_submit=True):
            nuevo_rut = st.text_input("RUT Empresa (ej. 76.519.206-4):")
            nueva_rz = st.text_input("Razón Social:")
            nueva_dir = st.text_input("Dirección:")
            nuevo_contacto = st.text_input("Email / Teléfono:")
            if st.form_submit_button("Guardar Empresa"):
                if nuevo_rut and nueva_rz:
                    add_empresa(nuevo_rut, nueva_rz, nueva_dir, nuevo_contacto)
                    st.success("Empresa registrada.")
                    st.rerun()
                else:
                    st.warning("Ingrese RUT y Razón Social.")

    st.markdown("---")
    st.subheader("📅 Período Contable")
    
    col_y, col_m = st.columns(2)
    with col_y:
        anio_sel = st.selectbox("Año Tributario:", options=[2026, 2025, 2024], index=0)
    with col_m:
        mes_sel = st.selectbox(
            "Mes de Proceso:", 
            options=list(MESES_MAP.keys()), 
            format_func=lambda m: f"{m:02d} - {MESES_NOMBRE[m-1]}",
            index=6 # Julio por defecto (07)
        )
        
    st.caption(f"🗓️ Año Comercial: **{anio_sel - 1}** | Período: **{MESES_NOMBRE[mes_sel-1]} {anio_sel}**")
    
    st.markdown("---")
    st.subheader("☁️ Nube y Respaldos")
    
    # Exportar / Importar Respaldo
    backup_data = export_backup_json()
    st.download_button(
        label="💾 Descargar Respaldo Total (JSON)",
        data=backup_data,
        file_name=f"Respaldo_Conta_Marlene_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        use_container_width=True
    )
    
    uploaded_backup = st.file_uploader("Restaurar Respaldo:", type=["json"], key="uploader_backup")
    if uploaded_backup is not None:
        if st.button("Restaurar Datos", use_container_width=True):
            content = uploaded_backup.getvalue().decode("utf-8")
            ok, msg = import_backup_json(content)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)


# =============================================================================
# 3. ENCABEZADO PRINCIPAL DEL DASHBOARD
# =============================================================================

st.markdown(f"""
<div style='background: linear-gradient(90deg, #1F497D 0%, #295C99 100%); padding: 18px 25px; border-radius: 10px; color: white; margin-bottom: 20px;'>
    <div style='display: flex; justify-content: space-between; align-items: center;'>
        <div>
            <h2 style='margin: 0; color: white;'>🏢 {empresa_actual['razon_social']}</h2>
            <p style='margin: 3px 0 0 0; opacity: 0.9; font-size: 15px;'>
                RUT: <b>{format_rut_dots(empresa_actual['rut'])}</b> &nbsp;|&nbsp; 
                Período: <b>{MESES_NOMBRE[mes_sel-1]} {anio_sel}</b> &nbsp;|&nbsp; 
                DJ Nº 1887 SII
            </p>
        </div>
        <div style='text-align: right;'>
            <span style='background: rgba(255,255,255,0.2); padding: 6px 14px; border-radius: 20px; font-size: 13px;'>
                Sesión: <b>{st.session_state['user_fullname']}</b>
            </span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# =============================================================================
# 4. PESTAÑAS PRINCIPALES DEL SISTEMA
# =============================================================================

tab_carga, tab_tabla, tab_kpi, tab_anual, tab_descargas, tab_auditoria = st.tabs([
    "📥 Carga y Parseo Previred",
    "📝 Nómina en Vivo (st.data_editor)",
    "🚦 Panel de Control y Alertas",
    "📚 Libro de Sueldos Anual",
    "📦 Descargas y DJ 1887 SII",
    "🕒 Historial de Auditoría"
])


# =============================================================================
# TAB 1: MÓDULO DE CARGA Y PARSEO PREVIRED (pdfplumber)
# =============================================================================

with tab_carga:
    st.subheader("📄 Carga Inteligente de Comprobantes de Pago Previred")
    st.markdown("""
    Sube aquí los comprobantes mensuales en formato PDF emitidos por **Previred** (Fonasa, IPS, Mutual/ISL, Seguro Social, AFP UNO, Modelo, PlanVital, etc.).
    El motor **`pdfplumber`** extraerá y conciliará de forma unificada la nómina de trabajadores del mes.
    """)
    
    col_up1, col_up2 = st.columns([2, 1])
    
    with col_up1:
        uploaded_pdfs = st.file_uploader(
            "Arrastra o selecciona los archivos PDF del mes:",
            type=["pdf"],
            accept_multiple_files=True,
            help="Puedes seleccionar varios PDFs simultáneamente."
        )
        
        btn_procesar = st.button("⚡ Procesar y Conciliar Comprobantes PDF", type="primary", use_container_width=True)

    with col_up2:
        with st.container(border=True):
            st.markdown("##### 🚀 Demostración Rápida")
            st.caption("¿No tienes los archivos PDF a mano? Carga los datos reales de **Julio 2026 de Dong Sheng** para probar el flujo completo en 1 clic:")
            if st.button("🧪 Cargar Muestra Real Previred (Julio 2026)", use_container_width=True):
                df_sample, res_sample = generate_sample_previred_july_2026()
                saved_n = save_remuneraciones_mes(
                    empresa_actual["rut"], anio_sel, mes_sel, df_sample, 
                    usuario=st.session_state["user_fullname"], 
                    resumen_previred=res_sample
                )
                st.session_state["cached_df_mes"] = df_sample
                st.session_state["cached_resumen_prev"] = res_sample
                st.success(f"¡Cargados exitosamente {saved_n} trabajadores con cuadratura Previred!")
                st.rerun()

    if btn_procesar and uploaded_pdfs:
        with st.spinner("Procesando archivos PDF con pdfplumber y conciliando información..."):
            parsed_docs = []
            log_mensajes = []
            
            for f in uploaded_pdfs:
                bytes_data = f.read()
                try:
                    doc_res = parse_previred_pdf(bytes_data)
                    parsed_docs.append(doc_res)
                    log_mensajes.append(f"✅ Archivo **{f.name}** detectado como: `{doc_res['institucion'] or doc_res['tipo_doc']}` ({len(doc_res['trabajadores'])} trabajadores)")
                except Exception as err:
                    log_mensajes.append(f"❌ Error al procesar **{f.name}**: {str(err)}")
                    
            if parsed_docs:
                df_mes_consolidado, resumen_prev = consolidate_previred_batch(parsed_docs)
                
                # Guardar en base de datos SQLite
                saved_count = save_remuneraciones_mes(
                    empresa_actual["rut"], anio_sel, mes_sel, df_mes_consolidado,
                    usuario=st.session_state["user_fullname"],
                    resumen_previred=resumen_prev
                )
                st.session_state["cached_df_mes"] = df_mes_consolidado
                st.session_state["cached_resumen_prev"] = resumen_prev
                
                st.success(f"¡Procesamiento finalizado! Se consolidaron {saved_count} trabajadores para el período {MESES_NOMBRE[mes_sel-1]} {anio_sel}.")
                
                with st.expander("🔍 Ver Detalle de Detección de Archivos"):
                    for msg in log_mensajes:
                        st.markdown(msg)
                st.rerun()
            else:
                st.warning("No se pudo extraer información válida de los archivos subidos.")


# =============================================================================
# CARGAR DATOS DEL MES ACTIVO PARA VISTAS
# =============================================================================

df_mes_actual = get_remuneraciones_mes(empresa_actual["rut"], anio_sel, mes_sel)
cuadratura_actual = get_cuadratura(empresa_actual["rut"], anio_sel, mes_sel)


# =============================================================================
# TAB 2: TABLA EDITABLE EN VIVO (st.data_editor) CON AUDITORÍA
# =============================================================================

with tab_tabla:
    st.subheader(f"📝 Nómina Editable en Vivo - {MESES_NOMBRE[mes_sel-1]} {anio_sel}")
    st.caption("Modifica en vivo cualquier valor (Impuesto Único, Renta Neta, Días, Bonos). Cada edición quedará registrada en el historial con tu nombre y fecha.")
    
    if df_mes_actual.empty:
        st.info("💡 Aún no hay datos cargados para este mes. Puedes subir los comprobantes en la pestaña **'Carga y Parseo Previred'** o presionar el botón de **'Cargar Muestra Real'**.")
    else:
        # Configurar columnas para st.data_editor
        column_order = [
            "rut_formateado", "nombre", "dias_trabajados",
            "sueldo_base_previred", "bono_adicional", "sueldo_bruto",
            "afp", "cotizacion_afp", "salud_entidad", "cotizacion_salud",
            "afc_trabajador", "cotizacion_previsional_total", "seguro_social_01",
            "impuesto_unico", "renta_neta_pagada", "asignacion_familiar",
            "sis", "afc_empleador", "isl_mutual", "seguro_social_patronal",
            "gastos_patronales_total", "movimiento_desc"
        ]
        
        column_config = {
            "rut_formateado": st.column_config.TextColumn("RUT Trabajador", disabled=True),
            "nombre": st.column_config.TextColumn("Nombre y Apellidos", width="medium"),
            "dias_trabajados": st.column_config.NumberColumn("Días", min_value=0, max_value=31, step=1),
            "sueldo_base_previred": st.column_config.NumberColumn("Base Previred ($)", format="$%d", disabled=True, help="Renta imponible original extraída de los comprobantes Previred"),
            "bono_adicional": st.column_config.NumberColumn("Bono / Haber Adicional ($)", format="$%d", min_value=0, step=1000, help="Ingreso manual de bonos o haberes imponibles. Se sumará automáticamente al Total Imponible / Bruto"),
            "sueldo_bruto": st.column_config.NumberColumn("Total Imponible / Bruto ($)", format="$%d", help="Suma de Base Previred + Bono / Haber Adicional"),
            "afp": st.column_config.SelectboxColumn("AFP", options=["AFP UNO", "AFP MODELO", "AFP PLANVITAL", "AFP HABITAT", "AFP CUPRUM", "AFP CAPITAL", "AFP PROVIDA"]),
            "cotizacion_afp": st.column_config.NumberColumn("Cotiz. AFP ($)", format="$%d"),
            "salud_entidad": st.column_config.SelectboxColumn("Salud", options=["FONASA", "ISAPRE", "PARTICULAR"]),
            "cotizacion_salud": st.column_config.NumberColumn("Cotiz. Salud ($)", format="$%d"),
            "afc_trabajador": st.column_config.NumberColumn("AFC Trabajador ($)", format="$%d"),
            "cotizacion_previsional_total": st.column_config.NumberColumn("Cotiz. Previs. Total ($)", format="$%d", disabled=True),
            "seguro_social_01": st.column_config.NumberColumn("Seguro Social 0.1% ($)", format="$%.1f"),
            "impuesto_unico": st.column_config.NumberColumn("Impuesto Único 2ª Cat ($)", format="$%d"),
            "renta_neta_pagada": st.column_config.NumberColumn("Renta Neta Pagada ($)", format="$%d"),
            "asignacion_familiar": st.column_config.NumberColumn("Asig. Familiar ($)", format="$%d"),
            "sis": st.column_config.NumberColumn("SIS ($)", format="$%d"),
            "afc_empleador": st.column_config.NumberColumn("AFC Empleador ($)", format="$%d"),
            "isl_mutual": st.column_config.NumberColumn("ISL / Mutual ($)", format="$%d"),
            "seguro_social_patronal": st.column_config.NumberColumn("Seguro Social Patronal ($)", format="$%d"),
            "gastos_patronales_total": st.column_config.NumberColumn("Gastos Patronales ($)", format="$%d", disabled=True),
            "movimiento_desc": st.column_config.TextColumn("Movimiento Personal", disabled=True)
        }
        
        # Filtrar solo columnas existentes
        valid_cols = [c for c in column_order if c in df_mes_actual.columns]
        
        edited_df = st.data_editor(
            df_mes_actual[valid_cols],
            column_config=column_config,
            use_container_width=True,
            num_rows="dynamic",
            key=f"editor_sueldos_{empresa_actual['rut']}_{anio_sel}_{mes_sel}"
        )
        
        col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
        
        with col_btn1:
            if st.button("💾 Guardar Cambios y Registrar Auditoría", type="primary", use_container_width=True):
                # Recalcular columnas derivadas en base a ediciones
                df_to_save = edited_df.copy()
                df_to_save["rut"] = df_to_save["rut_formateado"].apply(clean_rut)
                
                # Auto-sumar Bono / Haber Adicional a Base Previred para Total Imponible / Bruto
                if "bono_adicional" in df_to_save.columns:
                    base_vals = df_to_save["sueldo_base_previred"] if "sueldo_base_previred" in df_to_save.columns else df_to_save["sueldo_bruto"]
                    df_to_save["sueldo_bruto"] = base_vals.fillna(0.0) + df_to_save["bono_adicional"].fillna(0.0)
                
                df_to_save["cotizacion_previsional_total"] = df_to_save["cotizacion_afp"] + df_to_save["cotizacion_salud"] + df_to_save["afc_trabajador"]
                df_to_save["seguro_social_01"] = (df_to_save["sueldo_bruto"] * 0.001).round(3)
                df_to_save["renta_neta_pagada"] = (df_to_save["sueldo_bruto"] - df_to_save["cotizacion_previsional_total"] - df_to_save["seguro_social_01"] - df_to_save["impuesto_unico"]).round(3)
                df_to_save["gastos_patronales_total"] = (df_to_save["sis"] + df_to_save["afc_empleador"] + df_to_save["isl_mutual"] + df_to_save["seguro_social_patronal"] + df_to_save["seguro_social_01"]).round(3)
                
                n_saved = save_remuneraciones_mes(
                    empresa_actual["rut"], anio_sel, mes_sel, df_to_save,
                    usuario=st.session_state["user_fullname"]
                )
                st.success(f"¡Cambios guardados con éxito por {st.session_state['user_fullname']}! Se recalculó el Total Imponible y la Renta Neta ({n_saved} registros).")
                st.rerun()

        with col_btn2:
            # Sincronización Google Sheets
            if st.button("☁️ Sincronizar a Google Sheets", use_container_width=True):
                # Intentar leer service account de secrets de Streamlit
                if "gcp_service_account" in st.secrets:
                    service_account_info = dict(st.secrets["gcp_service_account"])
                    sheet_target = st.secrets.get("google_sheet_title", f"Conta_Marlene_{empresa_actual['rut']}")
                    ok_sync, msg_sync = sync_to_google_sheet(service_account_info, sheet_target, edited_df, sheet_name=f"{mes_sel:02d}_{anio_sel}")
                    if ok_sync:
                        st.success(msg_sync)
                    else:
                        st.error(msg_sync)
                else:
                    st.info("ℹ️ Para sincronizar automáticamente a Google Sheets en Streamlit Cloud, configure sus credenciales de Service Account en los `st.secrets`. Consulte la guía adjunta.")


# =============================================================================
# TAB 3: PANEL DE CONTROL, KPIS Y SEMÁFOROS DE CUADRATURA
# =============================================================================

with tab_kpi:
    st.subheader(f"🚦 Semáforos de Cuadratura y Control de Inconsistencias ({MESES_NOMBRE[mes_sel-1]} {anio_sel})")
    
    if df_mes_actual.empty:
        st.warning("No hay datos cargados para este período. Cargue los comprobantes para ver el análisis de control.")
    else:
        # Métricas Principales (KPIs)
        total_imponible = float(df_mes_actual["sueldo_bruto"].sum())
        total_cotiz_previs = float(df_mes_actual["cotizacion_previsional_total"].sum())
        total_gastos_patronales = float(df_mes_actual["gastos_patronales_total"].sum())
        total_impuesto_unico = float(df_mes_actual["impuesto_unico"].sum())
        total_afiliados = len(df_mes_actual)

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("👥 Total Trabajadores", f"{total_afiliados}")
        k2.metric("💰 Renta Imponible Total", f"${int(total_imponible):,}".replace(",", "."))
        k3.metric("🛡️ Cotiz. Previsional Trabajadores", f"${int(total_cotiz_previs):,}".replace(",", "."))
        k4.metric("🏢 Aportes Patronales (Coste Emp.)", f"${int(total_gastos_patronales):,}".replace(",", "."))
        k5.metric("🏛️ Impuesto Único Retenido", f"${int(total_impuesto_unico):,}".replace(",", "."))

        st.markdown("---")
        
        # 1. SEMÁFORO DE CUADRATURA PREVIRED
        st.markdown("#### 1. 🚦 Semáforo de Cuadratura con Comprobantes Previred")
        
        col_sem1, col_sem2 = st.columns([1.5, 2])
        
        with col_sem1:
            total_nomina_cotiz = total_cotiz_previs + total_gastos_patronales
            total_comprobantes = cuadratura_actual["total_previred_comprobantes"] if cuadratura_actual else total_nomina_cotiz
            diferencia = round(total_comprobantes - total_nomina_cotiz, 2)
            
            if abs(diferencia) < 5.0:
                st.markdown(f"""
                <div style='background: #D4EDDA; border: 2px solid #28A745; border-radius: 8px; padding: 18px; text-align: center;'>
                    <h3 style='color: #155724; margin: 0;'>🟢 CUADRATURA PERFECTA</h3>
                    <p style='color: #155724; margin: 8px 0 0 0; font-size: 15px;'>
                        El total pagado en comprobantes Previred coincide al 100% con la nómina detallada.<br/>
                        Diferencia detectada: <b>$0</b>
                    </p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style='background: #F8D7DA; border: 2px solid #DC3545; border-radius: 8px; padding: 18px; text-align: center;'>
                    <h3 style='color: #721C24; margin: 0;'>🔴 ALERTA DE DESCUADRE</h3>
                    <p style='color: #721C24; margin: 8px 0 0 0; font-size: 15px;'>
                        Diferencia detectada: <b>${int(diferencia):,}</b> entre los comprobantes Previred y la nómina calculada.<br/>
                        Revise ajustes manuales o cotizaciones no registradas.
                    </p>
                </div>
                """, unsafe_allow_html=True)

        with col_sem2:
            st.markdown(f"""
            - **Total Pagado según Carátulas Previred:** `${int(total_comprobantes):,}`.replace(",", ".")
            - **Total Calculado en Nómina Detallada:** `${int(total_nomina_cotiz):,}`.replace(",", ".")
            - **Instituciones Verificadas:** {cuadratura_actual.get('instituciones', 'Fonasa, IPS, ISL, AFPs') if cuadratura_actual else 'Fonasa, IPS, ISL, AFPs'}
            """)
            
        st.markdown("---")
        
        # 2. CONTROL DE TOPES IMPONIBLES EN UF
        st.markdown("#### 2. 🛡️ Verificación de Topes Imponibles en UF (Superintendencia de Pensiones)")
        
        trabajadores_sobre_tope = []
        for _, row in df_mes_actual.iterrows():
            chk = check_topes_imponibles(row["sueldo_bruto"], anio=anio_sel)
            if chk["excede_afp"] or chk["excede_afc"]:
                trabajadores_sobre_tope.append({
                    "RUT": format_rut_dots(row["rut"]),
                    "Nombre": row["nombre"],
                    "Sueldo Imponible": f"${int(row['sueldo_bruto']):,}".replace(",", "."),
                    "Tope AFP": f"${int(chk['tope_afp_pesos']):,}".replace(",", "."),
                    "Exceso AFP": f"${int(chk['exceso_afp']):,}".replace(",", "."),
                    "Tope AFC": f"${int(chk['tope_afc_pesos']):,}".replace(",", "."),
                    "Exceso AFC": f"${int(chk['exceso_afc']):,}".replace(",", "."),
                })
                
        if trabajadores_sobre_tope:
            st.warning(f"⚠️ Se detectaron **{len(trabajadores_sobre_tope)}** trabajadores con remuneración imponible superior al tope legal de AFP o AFC:")
            st.dataframe(pd.DataFrame(trabajadores_sobre_tope), use_container_width=True)
        else:
            st.success("✅ Todos los trabajadores se encuentran dentro o topados adecuadamente en los límites legales de UF.")

        st.markdown("---")
        
        # 3. DETECCIÓN AUTOMÁTICA DE MOVIMIENTOS DE PERSONAL
        st.markdown("#### 3. 👥 Movimientos de Personal Detectados en el Mes")
        
        movimientos_destacados = []
        for _, row in df_mes_actual.iterrows():
            c_mov = row.get("cod_movimiento", 0)
            if c_mov and c_mov > 0:
                movimientos_destacados.append({
                    "RUT": format_rut_dots(row["rut"]),
                    "Nombre": row["nombre"],
                    "Código": c_mov,
                    "Descripción Previred": row.get("movimiento_desc", ""),
                    "Alerta": row.get("alerta_movimiento", "Movimiento")
                })
                
        if movimientos_destacados:
            st.info(f"📢 Se registraron **{len(movimientos_destacados)}** movimientos de personal en Previred durante este mes:")
            st.dataframe(pd.DataFrame(movimientos_destacados), use_container_width=True)
        else:
            st.caption("ℹ️ No se registraron altas, bajas ni licencias médicas en este período (Código 0).")


# =============================================================================
# TAB 4: CONSOLIDADOR ANUAL Y LIBRO DE SUELDOS
# =============================================================================

with tab_anual:
    st.subheader(f"📚 Consolidación del Libro de Sueldos Anual - Año Tributario {anio_sel}")
    st.caption("Resumen consolidado de los 12 meses de remuneraciones para la empresa con aplicación de factores de actualización del SII.")
    
    df_anual = get_remuneraciones_anuales(empresa_actual["rut"], anio_sel)
    
    if df_anual.empty:
        st.info("Aún no hay meses guardados para este año tributario. Cargue meses adicionales para armar el consolidado anual.")
    else:
        # Agrupar por mes
        resumen_anual_rows = []
        for m in range(1, 13):
            m_df = df_anual[df_anual["mes"] == m]
            factor = get_factor_actualizacion(m, anio_sel)
            
            base_p = float(m_df["sueldo_base_previred"].sum()) if not m_df.empty and "sueldo_base_previred" in m_df.columns else (float(m_df["sueldo_bruto"].sum()) if not m_df.empty else 0.0)
            bono = float(m_df["bono_adicional"].sum()) if not m_df.empty and "bono_adicional" in m_df.columns else 0.0
            bruto = float(m_df["sueldo_bruto"].sum()) if not m_df.empty else 0.0
            cotiz = float(m_df["cotizacion_previsional_total"].sum()) if not m_df.empty else 0.0
            ss_01 = float(m_df["seguro_social_01"].sum()) if not m_df.empty else 0.0
            neta = float(m_df["renta_neta_pagada"].sum()) if not m_df.empty else 0.0
            iu = float(m_df["impuesto_unico"].sum()) if not m_df.empty else 0.0
            gastos_pat = float(m_df["gastos_patronales_total"].sum()) if not m_df.empty else 0.0
            asig_fam = float(m_df["asignacion_familiar"].sum()) if not m_df.empty else 0.0
            
            resumen_anual_rows.append({
                "Mes": MESES_MAP[m],
                "N° Trab.": len(m_df),
                "Base Previred ($)": base_p,
                "Bono Adicional ($)": bono,
                "Total Imponible / Bruto ($)": bruto,
                "Cotiz. Previsional ($)": cotiz,
                "Seguro Social 0.1% ($)": ss_01,
                "Renta Neta Pagada ($)": neta,
                "Factor SII": factor,
                "Bruto Actualizado ($)": round(bruto * factor),
                "Renta Neta Actualizada ($)": round(neta * factor),
                "Impuesto Único ($)": iu,
                "Imp. Único Actualizado ($)": round(iu * factor),
                "Gastos Patronales ($)": gastos_pat,
                "Asignación Familiar ($)": asig_fam
            })
            
        df_resumen_anual = pd.DataFrame(resumen_anual_rows)
        
        st.dataframe(
            df_resumen_anual.style.format({
                "Base Previred ($)": "${:,.0f}",
                "Bono Adicional ($)": "${:,.0f}",
                "Total Imponible / Bruto ($)": "${:,.0f}",
                "Cotiz. Previsional ($)": "${:,.0f}",
                "Seguro Social 0.1% ($)": "${:,.1f}",
                "Renta Neta Pagada ($)": "${:,.0f}",
                "Factor SII": "{:.3f}",
                "Bruto Actualizado ($)": "${:,.0f}",
                "Renta Neta Actualizada ($)": "${:,.0f}",
                "Impuesto Único ($)": "${:,.0f}",
                "Imp. Único Actualizado ($)": "${:,.0f}",
                "Gastos Patronales ($)": "${:,.0f}",
                "Asignación Familiar ($)": "${:,.0f}"
            }),
            use_container_width=True
        )


# =============================================================================
# TAB 5: CENTRO DE DESCARGAS OFICIALES (EXCEL, TXT DJ 1887 Y CERTIFICADO 6)
# =============================================================================

with tab_descargas:
    st.subheader("📦 Generador y Descargas Oficiales para el SII y la Oficina")
    st.markdown("Genera con un solo clic los reportes en los formatos técnicos exactos solicitados por la normativa chilena:")
    
    df_para_reportes = get_remuneraciones_anuales(empresa_actual["rut"], anio_sel)
    
    col_d1, col_d2, col_d3 = st.columns(3)
    
    # 1. EXCEL PLANTILLA ANUAL
    with col_d1:
        with st.container(border=True):
            st.markdown("#### 📊 Libro de Sueldos Anual (.XLSX)")
            st.write("Plantilla contable idéntica a la oficial con **hoja resumen de empresa** + **bloques individuales por trabajador** y fórmulas de actualización.")
            if not df_para_reportes.empty:
                excel_bytes = generate_excel_libro_sueldos_anual(empresa_actual, anio_sel, df_para_reportes)
                st.download_button(
                    label="📥 Descargar Excel Libro de Sueldos",
                    data=excel_bytes,
                    file_name=f"Libro_Sueldos_Anual_{empresa_actual['rut']}_{anio_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.warning("Sin datos anuales.")

    # 2. ARCHIVO PLANO TXT DJ 1887 PARA EL SII
    with col_d2:
        with st.container(border=True):
            st.markdown("#### 📄 Archivo Plano DJ 1887 (.TXT)")
            st.write("Estructura técnica oficial del SII (Registros Tipo 1, Tipo 2 y Tipo 3) para importación directa en la plataforma de Declaraciones Juradas.")
            if not df_para_reportes.empty:
                txt_content = generate_txt_dj1887(empresa_actual, anio_sel, df_para_reportes)
                st.download_button(
                    label="📥 Descargar TXT DJ 1887 (SII)",
                    data=txt_content,
                    file_name=f"DJ1887_{clean_rut(empresa_actual['rut'])}_{anio_sel}.txt",
                    mime="text/plain",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.warning("Sin datos anuales.")

    # 3. CERTIFICADOS Nº 6 DE SUELDOS (PDF)
    with col_d3:
        with st.container(border=True):
            st.markdown("#### 📜 Certificados Nº 6 de Sueldos (PDF)")
            st.write("Certificados anuales exigidos por el Art. 101 de la Ley de la Renta para entregar a los trabajadores para su Operación Renta (F22).")
            if not df_para_reportes.empty:
                # Selector individual o ZIP masivo
                ruts_unicos = df_para_reportes["rut"].unique()
                trabajador_sel_rut = st.selectbox(
                    "Seleccione Trabajador:",
                    options=ruts_unicos,
                    format_func=lambda r: f"{format_rut_dots(r)} - {df_para_reportes[df_para_reportes['rut']==r]['nombre'].iloc[0]}"
                )
                
                df_w = df_para_reportes[df_para_reportes["rut"] == trabajador_sel_rut]
                pdf_cert = generate_pdf_certificado_6(empresa_actual, anio_sel, trabajador_sel_rut, df_w)
                
                st.download_button(
                    label=f"📄 Descargar Certificado Nº 6",
                    data=pdf_cert,
                    file_name=f"Certificado_6_{clean_rut(trabajador_sel_rut)}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # Descarga masiva ZIP
                zip_certs = generate_zip_all_certificados(empresa_actual, anio_sel, df_para_reportes)
                st.download_button(
                    label="🗂️ Descargar Todos en ZIP",
                    data=zip_certs,
                    file_name=f"Certificados_6_Completos_{empresa_actual['rut']}_{anio_sel}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.warning("Sin datos anuales.")


# =============================================================================
# TAB 6: HISTORIAL DE AUDITORÍA (QUIÉN CAMBIÓ QUÉ Y CUÁNDO)
# =============================================================================

with tab_auditoria:
    st.subheader("🕒 Registro de Auditoría de Modificaciones")
    st.caption("Trazabilidad completa de cada modificación realizada por las usuarias (Marlene y analistas) para seguridad contable.")
    
    df_audit = get_auditoria(empresa_rut=empresa_actual["rut"], anio=anio_sel)
    
    if df_audit.empty:
        st.info("Aún no hay modificaciones registradas en el historial de auditoría.")
    else:
        st.dataframe(
            df_audit.rename(columns={
                "fecha_hora": "Fecha / Hora",
                "usuario": "Usuaria",
                "rut_trabajador": "RUT Trabajador",
                "campo_modificado": "Campo Editado",
                "valor_anterior": "Valor Anterior",
                "valor_nuevo": "Valor Nuevo"
            }),
            use_container_width=True
        )


# =============================================================================
# PIE DE PÁGINA
# =============================================================================

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #888; font-size: 13px;'>
    💼 <b>Conta Marlene</b> - Sistema Colaborativo de Previred, Libro de Sueldos y Declaración Jurada Nº 1887<br/>
    Diseñado para el equipo de Marlene (Oficina y Remoto) | Cumplimiento Dirección del Trabajo y SII Chile
</div>
""", unsafe_allow_html=True)
