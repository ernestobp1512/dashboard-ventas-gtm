import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import io
import calendar

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Go To Market - Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CUSTOM CSS (IDENTIDAD GTM + ESTILIZACIÓN) ──────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main { background-color: #F8FAFC; }
    
    .stSelectbox label p, .stMultiSelect label p {
        font-weight: bold !important;
        color: #1E293B !important;
        font-size: 14px !important;
    }

    .header-container {
        display: flex; align-items: center; justify-content: space-between;
        background: white; padding: 25px 35px; border-radius: 15px;
        margin-bottom: 25px; border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }

    .logo-bars {
        display: flex; align-items: flex-end; gap: 6px; margin-right: 25px;
    }
    .bar-small { width: 14px; height: 25px; background-color: #334155; border-radius: 2px; }
    .bar-med { width: 14px; height: 45px; background-color: #334155; border-radius: 2px; }
    .bar-large { width: 14px; height: 60px; background-color: #334155; border-radius: 2px; }

    .header-text h1 { 
        color: #0F172A; font-size: 52px; font-weight: 900; margin: 0; 
        line-height: 1; letter-spacing: -2px;
    }
    .brand-sub { 
        font-size: 24px; font-weight: 800; color: #0F172A; margin-top: 5px; 
    }
    .red-span { color: #DC2626; }

    .kpi-card {
        background: white; border-radius: 12px; padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04); border: 1px solid #E2E8F0;
        border-left: 6px solid #DC2626; height: 100%;
    }
    .kpi-label { font-size: 11px; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 30px; font-weight: 700; color: #1E293B; margin-top: 5px; }
    .kpi-card.green { border-left-color: #16A34A; }
    .kpi-card.green .kpi-value { color: #16A34A; }
    
    .kpi-desglose {
        font-size: 13px;
        margin-top: 10px;
        padding-top: 8px;
        border-top: 1px solid #E2E8F0;
        color: #1E293B;
        font-weight: 500;
    }
    
    .section-title {
        font-size: 17px; font-weight: 700; color: #1E293B;
        margin: 25px 0 15px; padding-bottom: 8px;
        border-bottom: 2px solid #F1F5F9;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .zona-badge {
        background-color: #DC2626;
        border-radius: 25px;
        padding: 8px 20px;
        font-size: 14px;
        font-weight: 600;
        color: white;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    .alerta-mensaje {
        border-radius: 12px;
        padding: 15px;
        margin: 15px 0;
        border-left: 5px solid;
        font-size: 13px;
        line-height: 1.4;
    }
    .alerta-mensaje.alerta {
        background-color: #FEE2E2;
        border-left-color: #DC2626;
    }
    .alerta-mensaje.recomendacion {
        background-color: #DCFCE7;
        border-left-color: #16A34A;
    }
    
    .ranking-header {
        background: white;
        padding: 20px 25px;
        border-radius: 15px;
        margin-bottom: 25px;
        border: 1px solid #E2E8F0;
    }
    .ranking-title {
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 5px;
    }
    .ranking-sub {
        color: #64748B;
        font-size: 13px;
    }
    
    /* Estilos para el embudo */
    .funnel-card {
        transition: all 0.2s ease;
        cursor: help;
    }
    .funnel-card:hover {
        transform: translateX(5px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# ─── HEADER PROFESIONAL ─────────────────────────────────────────────────────
st.markdown("""
<div class="header-container">
    <div style="display: flex; align-items: center;">
        <div class="logo-bars">
            <div class="bar-small"></div>
            <div class="bar-med"></div>
            <div class="bar-large"></div>
        </div>
        <div class="header-text">
            <h1>Go To Market SAC</h1>
            <div class="brand-sub">go<span class="red-span">to</span>market</div>
        </div>
    </div>
    <div style="text-align: right; color: #64748B; font-size: 12px;">
        <b>REPORTE EJECUTIVO</b><br>Gestión Comercial 2026
    </div>
</div>
""", unsafe_allow_html=True)

# ─── CONSTANTS ──────────────────────────────────────────────────────────────
META_SEMANAL = 25
META_DIARIA = 5
COL_POSITIONS = {"fecha": 1, "zona": 3, "cliente": 6, "tipo": 8, "giro": 14, "tipo_visita": 20}

# ─── CLASIFICACIÓN DE ZONAS (LIMA vs PROVINCIA) ─────────────────────────────
CLASIFICACION_ZONAS = {
    "MAYORISTAS ABARROTES": "LIMA",
    "LIMA NORTE 1": "LIMA",
    "SUR CHICO 1": "LIMA",
    "LIMA SUR 1": "LIMA",
    "MAYORISTAS": "LIMA",
    "HUNTER 1": "LIMA",
    "MAYORISTAS 2": "LIMA",
    "CENTRO 1": "PROVINCIA",
    "SUR 1": "PROVINCIA",
    "NORTE 2": "PROVINCIA",
    "ORIENTE 1": "PROVINCIA",
    "SUR 2": "PROVINCIA",
    "NORTE 1": "PROVINCIA",
    "NORTE 3": "PROVINCIA",
    "CENTRO 2": "PROVINCIA",
    "SUR 3": "PROVINCIA",
}

# ─── ZONAS ACTIVAS FIJAS ────────────────────────────────────────────────────
ZONAS_LIMA = ["MAYORISTAS ABARROTES", "LIMA NORTE 1", "SUR CHICO 1", "LIMA SUR 1", "MAYORISTAS", "HUNTER 1", "MAYORISTAS 2"]
ZONAS_PROVINCIA = ["CENTRO 1", "SUR 1", "NORTE 2", "ORIENTE 1", "SUR 2", "NORTE 1", "NORTE 3", "CENTRO 2", "SUR 3"]
ZONAS_EXCLUIR = ["BROKER", "NORTE CHICO 1", "OFICINA", "MODERNO", "MARCAS PROPIAS"]

NUM_LIMA = len(ZONAS_LIMA)
NUM_PROVINCIA = len(ZONAS_PROVINCIA)
NUM_LIMA = len(ZONAS_LIMA)
NUM_PROVINCIA = len(ZONAS_PROVINCIA)

# ─── METAS SEGÚN ZONA ────────────────────────────────────────────────────────
def get_metas_zona(zona):
    """Retorna (meta_diaria, meta_semanal) según el tipo de zona"""
    zona_upper = zona.upper().strip()
    if zona_upper in ["MAYORISTAS", "MAYORISTAS 2"]:
        return 8, 40  # Mayoristas: 8 visitas/día, 40/semana
    else:
        return 5, 25  # Zonas normales: 5 visitas/día, 25/semana

# ─── FUNCIÓN PARA CONTAR SEMANAS (SÁBADOS) EN UN MES ─────────────────────────
def contar_semanas_en_mes(mes_str):
    ...

# ─── FUNCIÓN PARA CONTAR SEMANAS (SÁBADOS) EN UN MES ─────────────────────────
def contar_semanas_en_mes(mes_str):
    try:
        fecha = datetime.strptime(mes_str, "%B %Y")
        año = fecha.year
        mes = fecha.month
        cal = calendar.monthcalendar(año, mes)
        num_sabados = 0
        for semana in cal:
            if semana[5] != 0:
                num_sabados += 1
        return num_sabados
    except:
        return 4

# ─── HELPERS ────────────────────────────────────────────────────────────────
def obtener_ultimos_3_meses():
    hoy = datetime.now()
    meses = []
    for i in range(3, 0, -1):
        mes = hoy.month - i
        año = hoy.year
        if mes <= 0:
            mes += 12
            año -= 1
        meses.append(f"{datetime(año, mes, 1).strftime('%B')} {año}")
    return meses

def obtener_datos_semana(fecha):
    lunes = fecha - pd.to_timedelta(fecha.weekday(), unit='D')
    sabado = lunes + pd.Timedelta(days=5)
    rango = f"{lunes.strftime('%d/%m')} - {sabado.strftime('%d/%m/%Y')}"
    return lunes, rango


# ============================================================
# FUNCIÓN DE ALERTAS Y RECOMENDACIONES (VERSIÓN CON SEMANA CERRADA)
# ============================================================
def generar_alertas_simple(df_zona, nombre_zona, total_visitas, meta_semanal=25):
    """Genera mensajes de alerta con metas dinámicas según la zona"""
    
    if df_zona.empty:
        return None
    
    # Obtener metas según la zona
    meta_diaria, meta_semanal_zona = get_metas_zona(nombre_zona)
    
    hoy = datetime.now()
    dia_actual_num = hoy.weekday()
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
    dia_actual_nombre = dias_semana[dia_actual_num] if dia_actual_num < 6 else "Sábado"
    
    # META ACUMULADA SEGÚN EL DÍA (con meta_diaria personalizada)
    if dia_actual_num == 0:  # Lunes
        meta_acumulada = meta_diaria
        dias_a_contar = ["Lunes"]
        es_viernes_o_sabado = False
    elif dia_actual_num == 1:  # Martes
        meta_acumulada = meta_diaria * 2
        dias_a_contar = ["Lunes", "Martes"]
        es_viernes_o_sabado = False
    elif dia_actual_num == 2:  # Miércoles
        meta_acumulada = meta_diaria * 3
        dias_a_contar = ["Lunes", "Martes", "Miércoles"]
        es_viernes_o_sabado = False
    elif dia_actual_num == 3:  # Jueves
        meta_acumulada = meta_diaria * 4
        dias_a_contar = ["Lunes", "Martes", "Miércoles", "Jueves"]
        es_viernes_o_sabado = False
    else:  # Viernes o Sábado
        meta_acumulada = meta_semanal_zona
        dias_a_contar = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
        es_viernes_o_sabado = True
    
    # CALCULAR VISITAS POR DÍA Y ACUMULADAS
    lunes_actual = hoy - timedelta(days=dia_actual_num)
    visitas_acumuladas = 0
    visitas_por_dia = {}
    
    for i, dia in enumerate(dias_semana[:dia_actual_num + 1]):
        fecha_dia = lunes_actual + timedelta(days=i)
        visitas = len(df_zona[df_zona["fecha"].dt.date == fecha_dia.date()])
        visitas_por_dia[dia] = visitas
        if dia in dias_a_contar:
            visitas_acumuladas += visitas
    
    pct_acumulado = (visitas_acumuladas / meta_acumulada * 100) if meta_acumulada > 0 else 0
    pct_semana = (total_visitas / meta_semanal_zona * 100) if meta_semanal_zona > 0 else 0
    
    restantes_acumulado = meta_acumulada - visitas_acumuladas
    restantes_semana = meta_semanal_zona - total_visitas
    
    # IDENTIFICAR DÍAS QUE FALLARON Y DÍAS EXCELENTES
    dias_fallados = []
    dias_excelentes = []
    
    for dia, visitas in visitas_por_dia.items():
        if visitas < meta_diaria:
            if visitas == 0:
                dias_fallados.append(f"{dia} (0)")
            else:
                dias_fallados.append(f"{dia} ({visitas})")
        elif visitas >= meta_diaria + 2:
            dias_excelentes.append(f"{dia} ({visitas})")
    
    # CALCULAR MÉTRICAS DE PROSPECCIÓN
    prospecciones = len(df_zona[df_zona["tipo"] == "PROSPECCIÓN"]) if "tipo" in df_zona.columns else 0
    total_visitas_zona = len(df_zona)
    es_primera_o_segunda_semana = hoy.day <= 14
    semana_actual = 1 if hoy.day <= 7 else 2 if hoy.day <= 14 else 3
    
    # Prospección recomendada según zona
    if meta_diaria == 8:
        prospeccion_recomendada = 3
    else:
        prospeccion_recomendada = 5
    
    # DETERMINAR TIPO DE MENSAJE
    if pct_acumulado >= 100:
        tipo = "recomendacion"
        emoji = "🏆"
        titulo = "¡EXCELENTE!"
    elif pct_acumulado >= 80:
        tipo = "recomendacion"
        emoji = "✅"
        titulo = "RECOMENDACIÓN"
    elif pct_acumulado >= 60:
        tipo = "recomendacion"
        emoji = "⚠️"
        titulo = "ATENCIÓN"
    else:
        tipo = "alerta"
        emoji = "🔴"
        titulo = "ALERTA"
    
    # BLOQUE 1: ESTADO GENERAL
    if es_viernes_o_sabado:
        mensaje = f"{emoji} {titulo} - {nombre_zona}: Llevas {total_visitas} de {meta_semanal_zona} visitas semanales ({pct_semana:.0f}%). "
    else:
        mensaje = f"{emoji} {titulo} - {nombre_zona}: Llevas {visitas_acumuladas} de {meta_acumulada} visitas acumuladas ({pct_acumulado:.0f}%). "
    
    # BLOQUE 2: DÍAS QUE FALLÓ
    if dias_fallados:
        if len(dias_fallados) == 1:
            mensaje += f"Fallaste el {dias_fallados[0]}. "
        else:
            dias_texto = ", ".join(dias_fallados[:-1]) + f" y {dias_fallados[-1]}"
            mensaje += f"Fallaste el {dias_texto}. "
    
    # BLOQUE 3: DÍAS EXCELENTES
    if dias_excelentes:
        if len(dias_excelentes) == 1:
            mensaje += f"Solo el {dias_excelentes[0]} cumpliste. "
        else:
            mensaje += f"Los días {', '.join(dias_excelentes)} fueron excelentes. "
    
    # BLOQUE 4: META RESTANTE
    if es_viernes_o_sabado:
        if restantes_semana > 0:
            mensaje += f"Hoy es {dia_actual_nombre}, te faltan {restantes_semana} visitas para cerrar la semana. "
        else:
            mensaje += f"¡Excelente! Cumpliste la meta semanal. "
    else:
        if restantes_acumulado > 0:
            mensaje += f"Hoy es {dia_actual_nombre}, te faltan {restantes_acumulado} visitas para cumplir la meta del día. "
        else:
            mensaje += f"¡Excelente! Cumpliste la meta del día. "
    
    # BLOQUE 5: PROSPECCIÓN (solo semanas 1 y 2)
    if es_primera_o_segunda_semana:
        if prospecciones == 0:
            mensaje += f"⚠️ PROSPECCIÓN CRÍTICA: Estamos en semana {semana_actual}. NO TIENES prospecciones. Necesitas al menos {prospeccion_recomendada} esta semana. ¡Sal a prospectar URGENTE! "
        elif prospecciones < prospeccion_recomendada:
            mensaje += f"⚠️ PROSPECCIÓN: Estamos en semana {semana_actual}. Solo tienes {prospecciones} prospección(es). Necesitas al menos {prospeccion_recomendada - prospecciones} más esta semana. "
    
    # CIERRE ESPECIAL PARA VIERNES
    if dia_actual_num == 4 and restantes_semana > 0:
        mensaje += "¡Es VIERNES! Último día hábil. ¡Dale con todo!"
    elif dia_actual_num == 5:
        mensaje += "Es SÁBADO. Último día de la semana. ¡Cierra fuerte!"
    
    return f'<div class="alerta-mensaje {tipo}">{mensaje}</div>'

@st.cache_data
def load_excel(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=0)
    cols = list(df_raw.columns)
    rename_map = {cols[1]: "fecha", cols[6]: "Cliente o Prospecto", cols[14]: "Giro"} 
    for name, idx in COL_POSITIONS.items():
        if idx < len(cols) and idx not in [1, 6, 14]:
            rename_map[cols[idx]] = name
    df = df_raw.rename(columns=rename_map)
    df = df.dropna(subset=["Cliente o Prospecto"])
    df = df[df["Cliente o Prospecto"].astype(str).str.strip() != ""]
    df = df[df["Cliente o Prospecto"].astype(str).str.upper().str.strip() != "GO TO MARKET"]
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df = df.dropna(subset=["fecha"])
    res = df['fecha'].apply(lambda x: pd.Series(obtener_datos_semana(x)))
    df["semana_inicio"] = res[0]
    df["semana_rango"] = res[1]
    semanas_ordenadas_fecha = sorted(df["semana_inicio"].unique())
    mapa_semanas = {fecha: f"Semana {i+1}" for i, fecha in enumerate(semanas_ordenadas_fecha)}
    df["semana"] = df["semana_inicio"].map(mapa_semanas)
    df["mes"] = df["fecha"].dt.strftime("%B %Y")
    df["dia_semana"] = df["fecha"].dt.day_name()
    return df, None

# ============================================================
# FUNCIÓN PARA GENERAR REPORTE HTML DEL DASHBOARD
# ============================================================
def generar_reporte_html(df_zona, nombre_zona, total_visitas, prospeccion_visitas, mantenimiento_visitas, meta_periodo, data_alertas, resumen_giro, mensaje_alerta, sel_sema, cierres=0, prospectos_unicos=0):
    """Genera reporte HTML con métricas clave, cierres y alertas de cumplimiento con colores"""
    
    # Calcular métricas adicionales
    tasa_conversion = (cierres / prospectos_unicos * 100) if prospectos_unicos > 0 else 0
    cobertura = (prospectos_unicos / prospeccion_visitas * 100) if prospeccion_visitas > 0 else 0
    pct_cumplimiento = (total_visitas / meta_periodo * 100) if meta_periodo > 0 else 0
    
    # Determinar colores según rendimiento
    if pct_cumplimiento >= 80:
        pct_color = "#16A34A"
        pct_bg = "#DCFCE7"
    elif pct_cumplimiento >= 50:
        pct_color = "#F59E0B"
        pct_bg = "#FEF3C7"
    else:
        pct_color = "#DC2626"
        pct_bg = "#FEE2E2"
    
    if tasa_conversion >= 30:
        conversion_color = "#16A34A"
    elif tasa_conversion >= 15:
        conversion_color = "#F59E0B"
    else:
        conversion_color = "#DC2626"
    
    if cobertura >= 70:
        cobertura_color = "#16A34A"
    elif cobertura >= 40:
        cobertura_color = "#F59E0B"
    else:
        cobertura_color = "#DC2626"
    
    # ============================================================
    # ALERTAS DE CUMPLIMIENTO CON COLORES (ROJO/VERDE)
    # ============================================================
    alertas_html = ""
    if not data_alertas.empty:
        if sel_sema == "Todas":
            # Para vista semanal
            for _, row in data_alertas.iterrows():
                visitas = int(row['visitas'])
                cumple = visitas >= 25  # Meta semanal
                bg_color = "#DCFCE7" if cumple else "#FEE2E2"
                text_color = "#16A34A" if cumple else "#DC2626"
                
                # Obtener prospección y mantenimiento de la semana si están disponibles
                prospeccion_sem = row.get('prospeccion', 0)
                mantenimiento_sem = row.get('mantenimiento', 0)
                
                alertas_html += f"""
                <div style="background: {bg_color}; border-radius: 12px; padding: 15px; margin: 8px; text-align: center; min-width: 140px;">
                    <strong style="font-size: 13px; color: #1E293B;">{row['semana']}</strong><br>
                    <div style="font-size: 28px; font-weight: 800; color: {text_color};">{visitas}</div>
                    <div style="font-size: 10px; color: #64748B; margin-top: 5px;">
                        🎯 {prospeccion_sem} | 🔧 {mantenimiento_sem}
                    </div>
                </div>
                """
        else:
            # Para vista diaria
            for _, row in data_alertas.iterrows():
                visitas = int(row['visitas'])
                cumple = visitas >= 5  # Meta diaria
                bg_color = "#DCFCE7" if cumple else "#FEE2E2"
                text_color = "#16A34A" if cumple else "#DC2626"
                
                # Obtener prospección y mantenimiento del día
                prospeccion_dia = row.get('prospeccion', 0)
                mantenimiento_dia = row.get('mantenimiento', 0)
                
                alertas_html += f"""
                <div style="background: {bg_color}; border-radius: 12px; padding: 15px; margin: 8px; text-align: center; min-width: 120px;">
                    <strong style="font-size: 13px; color: #1E293B;">{row['label']}</strong><br>
                    <div style="font-size: 28px; font-weight: 800; color: {text_color};">{visitas}</div>
                    <div style="font-size: 10px; color: #64748B; margin-top: 5px;">
                        🎯 {prospeccion_dia} | 🔧 {mantenimiento_dia}
                    </div>
                </div>
                """
    
    # ============================================================
    # RESUMEN POR GIRO (con cierres)
    # ============================================================
    resumen_giro_html = ""
    if not resumen_giro.empty:
        resumen_giro_html = '<table style="width: 100%; border-collapse: collapse; margin-top: 10px;">'
        resumen_giro_html += '<tr style="background-color: #DC2626; color: white;">'
        resumen_giro_html += '<th style="padding: 10px; border: 1px solid #ddd;">Giro</th>'
        resumen_giro_html += '<th style="padding: 10px; border: 1px solid #ddd;">Clientes</th>'
        resumen_giro_html += '<th style="padding: 10px; border: 1px solid #ddd;">Visitas</th>'
        resumen_giro_html += '<th style="padding: 10px; border: 1px solid #ddd;">Frecuencia</th>'
        resumen_giro_html += '<th style="padding: 10px; border: 1px solid #ddd;">Cierres</th>'
        resumen_giro_html += '</tr>'
        
        for _, row in resumen_giro.iterrows():
            resumen_giro_html += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">{row['Giro']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{row['Clientes']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{row['Visitas_Totales']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{row['Frecuencia']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{row.get('Cierres', 0)}</td>
            </tr>
            """
        resumen_giro_html += '</table>'
    
    # Determinar tipo de alerta para el color
    es_alerta = "ALERTA" in mensaje_alerta or "ATENCIÓN" in mensaje_alerta
    bg_alerta = "#FEE2E2" if es_alerta else "#DCFCE7"
    color_borde = "#DC2626" if es_alerta else "#16A34A"
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GTM SAC - Reporte {nombre_zona}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 40px; background: #F8FAFC; }}
        .report-container {{ max-width: 1000px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ font-size: 28px; margin: 0; }}
        .content {{ padding: 30px; }}
        .alerta {{ background-color: {bg_alerta}; border-left: 5px solid {color_borde}; padding: 15px 20px; border-radius: 12px; margin-bottom: 25px; font-size: 13px; line-height: 1.4; }}
        
        .metrics-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }}
        .metric-card {{ background: #F8FAFC; border-radius: 12px; padding: 15px; text-align: center; border-top: 4px solid #DC2626; }}
        .metric-label {{ font-size: 11px; color: #64748B; text-transform: uppercase; font-weight: 700; }}
        .metric-value {{ font-size: 28px; font-weight: 800; color: #1E293B; margin: 5px 0; }}
        .metric-sub {{ font-size: 10px; color: #94A3B8; }}
        .metric-badge {{ display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; margin-top: 5px; }}
        
        .section-title {{ font-size: 18px; font-weight: 700; margin: 25px 0 15px; padding-bottom: 8px; border-bottom: 2px solid #E2E8F0; }}
        .alertas-container {{ display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }}
        
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
        th, td {{ padding: 10px; border: 1px solid #E2E8F0; text-align: left; }}
        th {{ background-color: #DC2626; color: white; }}
        
        .footer {{ background: #F1F5F9; padding: 15px; text-align: center; font-size: 11px; color: #94A3B8; }}
    </style>
</head>
<body>
    <div class="report-container">
        <div class="header">
            <h1>Go To Market SAC</h1>
            <p>Reporte de Gestión Comercial</p>
            <p style="font-size: 14px;">{nombre_zona} | {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </div>
        <div class="content">
            <!-- Alerta -->
            <div class="alerta">{mensaje_alerta}</div>
            
            <!-- Métricas Clave (8 indicadores) -->
            <div class="metrics-grid">
                <div class="metric-card">
                    <div class="metric-label">TOTAL VISITAS</div>
                    <div class="metric-value">{total_visitas}</div>
                    <div class="metric-sub">Meta: {meta_periodo}</div>
                    <div class="metric-badge" style="background: {pct_bg}; color: {pct_color};">{pct_cumplimiento:.0f}%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">🎯 PROSPECCIÓN</div>
                    <div class="metric-value">{prospeccion_visitas}</div>
                    <div class="metric-sub">Meta: 5/semana</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">🔧 MANTENIMIENTO</div>
                    <div class="metric-value">{mantenimiento_visitas}</div>
                    <div class="metric-sub">-</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">✅ CIERRES</div>
                    <div class="metric-value">{cierres}</div>
                    <div class="metric-sub">Meta: 2/semana</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">👥 PROSPECTOS ÚNICOS</div>
                    <div class="metric-value">{prospectos_unicos}</div>
                    <div class="metric-sub">-</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">📈 TASA CONVERSIÓN</div>
                    <div class="metric-value" style="color: {conversion_color};">{tasa_conversion:.0f}%</div>
                    <div class="metric-sub">Meta: 20%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">📊 COBERTURA</div>
                    <div class="metric-value" style="color: {cobertura_color};">{cobertura:.0f}%</div>
                    <div class="metric-sub">Meta: 70%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">📅 META SEMANAL</div>
                    <div class="metric-value">{total_visitas}/{meta_periodo}</div>
                    <div class="metric-sub">{pct_cumplimiento:.0f}% cumplido</div>
                </div>
            </div>
            
            <!-- Alertas de Cumplimiento CON COLORES -->
            <div class="section-title">🚦 Alertas de Cumplimiento</div>
            <div class="alertas-container">
                {alertas_html if alertas_html else '<p>No hay alertas disponibles</p>'}
            </div>
            
            <!-- Resumen por Giro (con cierres) -->
            <div class="section-title">📋 Resumen Estratégico por Giro</div>
            {resumen_giro_html if resumen_giro_html else '<p>No hay datos disponibles</p>'}
            
            <!-- Nota de cierre -->
            <div style="margin-top: 30px; padding: 15px; background: #F1F5F9; border-radius: 12px; text-align: center; font-size: 12px; color: #64748B;">
                <strong>💡 ¿Qué significan estos indicadores?</strong><br>
                • <strong>Tasa Conversión</strong> = Cierres / Prospectos Únicos (ideal >20%)<br>
                • <strong>Cobertura</strong> = Prospectos Únicos / Visitas Prospección (ideal <70% = más de 1 visita por prospecto)<br>
                • <strong>Cierres</strong> = Negocios cerrados en la semana
            </div>
        </div>
        <div class="footer">
            <p>Go To Market SAC · Reporte generado automáticamente</p>
            <p>Para guardar como PDF: Ctrl+P → "Guardar como PDF"</p>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ============================================================
# FUNCIÓN PARA GENERAR REPORTE HTML COMPLETO DEL RANKING
# ============================================================
def generar_reporte_ranking_completo_html(df_ranking, mes_seleccionado, mostrar_opcion, meta_periodo, semanas_transcurridas, alertas_data, total_cierres, cobertura_promedio, conversion_promedio, zonas_alta, zonas_media, zonas_baja, total_zonas, zonas_sin_cierre, zonas_sin_prospeccion, zonas_bajo_avance, zonas_baja_prospeccion, zonas_exceso_mantenimiento, zonas_alta_cobertura):
    """Genera un reporte HTML completo del ranking con tabla y alertas"""
    
    traduccion_meses = {
        "January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril",
        "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto",
        "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"
    }
    
    mes_sel_esp = mes_seleccionado
    for eng, esp in traduccion_meses.items():
        if eng in mes_seleccionado:
            mes_sel_esp = mes_seleccionado.replace(eng, esp)
            break
    
    # Tabla principal
    columnas = df_ranking.columns.tolist()
    headers_html = ""
    for col in columnas:
        headers_html += f'<th style="background-color: #DC2626; color: white; padding: 12px 15px; text-align: center; border: 1px solid #E2E8F0;">{col}</th>'
    
    filas_html = ""
    for _, row in df_ranking.iterrows():
        filas_html += '<tr>'
        for col in columnas:
            valor = row[col]
            if col == "POS":
                filas_html += f'<td style="background-color: #1E293B; color: white; font-weight: 800; text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            elif col == "ZONA":
                filas_html += f'<td style="text-align: left; font-weight: 600; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            elif col in ["% AVANCE", "COBERTURA", "TASA CONV."]:
                try:
                    num_val = float(str(valor).replace('%', ''))
                    filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{num_val:.0f}%</td>'
                except:
                    filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            else:
                filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
        filas_html += '</table>'
    
    # Alertas HTML
    alertas_html = ""
    for a in alertas_data[:15]:
        alertas_html += f"""
        <tr>
            <td style="padding: 10px; border: 1px solid #E2E8F0;">{a['prioridad']}</td>
            <td style="padding: 10px; border: 1px solid #E2E8F0; font-weight: 600;">{a['zona']}</td>
            <td style="padding: 10px; border: 1px solid #E2E8F0;">{a['alerta']}</td>
            <td style="padding: 10px; border: 1px solid #E2E8F0;">{a['insight']}</td>
            <td style="padding: 10px; border: 1px solid #E2E8F0;">{a['accion']}</td>
        </tr>
        """
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GTM SAC - Ranking de Zonas</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Inter', sans-serif; background: #F8FAFC; padding: 40px; }}
        .report-container {{ max-width: 1400px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ font-size: 28px; }}
        .content {{ padding: 30px; }}
        .info {{ background: #F1F5F9; padding: 15px; border-radius: 12px; margin-bottom: 25px; }}
        .insights-card {{ background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border-radius: 16px; padding: 20px; margin: 20px 0; color: white; }}
        .insights-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px; }}
        .insight-box {{ background: rgba(255,255,255,0.1); padding: 12px; border-radius: 10px; text-align: center; }}
        .table-container {{ overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        th {{ background-color: #DC2626; color: white; padding: 12px; }}
        td {{ padding: 10px; border: 1px solid #E2E8F0; }}
        .alertas-table {{ margin-top: 30px; }}
        .footer {{ background: #F1F5F9; padding: 15px; text-align: center; font-size: 11px; color: #94A3B8; }}
    </style>
</head>
<body>
    <div class="report-container">
        <div class="header">
            <h1>Go To Market SAC</h1>
            <p>Ranking de Gestión de Calidad GTM - {mes_sel_esp}</p>
            <p>Mostrar: {mostrar_opcion} | {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </div>
        <div class="content">
            <div class="info">
                <strong>📊 DATOS GENERALES</strong><br>
                • Ranking INDEPENDIENTE del dashboard<br>
                • ✅ Solo visitas FÍSICAS (regla comercial)<br>
                • Basado en CIERRES, no en tasas<br>
                • Meta del periodo: {meta_periodo} visitas ({semanas_transcurridas} semanas × 25)
            </div>
            
            <div class="insights-card">
                <h3 style="color: #DC2626; margin-bottom: 15px;">📈 INSIGHTS ESTRATÉGICOS - CIERRE DE SEMANA</h3>
                <div class="insights-grid">
                    <div class="insight-box">
                        <strong>🔴 PRIORIDAD ALTA</strong><br>
                        <span style="font-size: 24px; font-weight: 800;">{zonas_alta}</span><br>
                        <span style="font-size: 11px;">requieren one-to-one URGENTE</span>
                    </div>
                    <div class="insight-box">
                        <strong>🟡 SEGUIMIENTO</strong><br>
                        <span style="font-size: 24px; font-weight: 800;">{zonas_media}</span><br>
                        <span style="font-size: 11px;">requieren monitoreo</span>
                    </div>
                    <div class="insight-box">
                        <strong>🟢 RECONOCIMIENTO</strong><br>
                        <span style="font-size: 24px; font-weight: 800;">{zonas_baja}</span><br>
                        <span style="font-size: 11px;">desempeño destacado</span>
                    </div>
                    <div class="insight-box">
                        <strong>🎯 TOTAL CIERRES</strong><br>
                        <span style="font-size: 24px; font-weight: 800;">{total_cierres}</span><br>
                        <span style="font-size: 11px;">negocios cerrados</span>
                    </div>
                </div>
                
                <div style="margin-bottom: 15px;">
                    <div style="background: rgba(220,38,38,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🔴 BAJO AVANCE EN VISITAS (&lt;80% de meta)</strong><br>
                        {', '.join(zonas_bajo_avance) if zonas_bajo_avance else '✅ Todas las zonas superan el 80% de avance'}<br>
                        <span style="font-size: 11px;">→ Revisar planificación de ruta en one-to-one</span>
                    </div>
                    <div style="background: rgba(220,38,38,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🔴 BAJAS VISITAS DE PROSPECCIÓN (&lt;5 en semana 1-2)</strong><br>
                        {', '.join(zonas_baja_prospeccion) if zonas_baja_prospeccion else '✅ Todas las zonas prospectan adecuadamente'}<br>
                        <span style="font-size: 11px;">→ Salir a buscar nuevos clientes</span>
                    </div>
                    <div style="background: rgba(220,38,38,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🔴 EXCESO DE MANTENIMIENTO (más mantenimiento que prospección)</strong><br>
                        {', '.join(zonas_exceso_mantenimiento) if zonas_exceso_mantenimiento else '✅ Balance adecuado entre prospección y mantenimiento'}<br>
                        <span style="font-size: 11px;">→ Las primeras semanas son para PROSPECTAR</span>
                    </div>
                    <div style="background: rgba(245,158,11,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🟡 PROSPECTOS VISITADOS SOLO UNA VEZ (cobertura &gt;80%)</strong><br>
                        {', '.join(zonas_alta_cobertura) if zonas_alta_cobertura else '✅ Buena re-visita a prospectos'}<br>
                        <span style="font-size: 11px;">→ Dar segunda visita para poder cerrar</span>
                    </div>
                    <div style="background: rgba(220,38,38,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🔴 ZONAS SIN CIERRES</strong><br>
                        {', '.join(zonas_sin_cierre) if zonas_sin_cierre else '✅ Todas las zonas tienen cierres'}<br>
                        <span style="font-size: 11px;">→ Revisar técnica de cierre en one-to-one</span>
                    </div>
                    <div style="background: rgba(245,158,11,0.15); padding: 12px; border-radius: 10px; margin-bottom: 10px;">
                        <strong>🟡 ZONAS SIN PROSPECCIÓN</strong><br>
                        {', '.join(zonas_sin_prospeccion) if zonas_sin_prospeccion else '✅ Todas las zonas prospectan'}<br>
                        <span style="font-size: 11px;">→ Activar prospección URGENTE</span>
                    </div>
                </div>
                
                <div style="background: rgba(255,255,255,0.1); padding: 12px; border-radius: 10px;">
                    <strong>📊 MÉTRICAS CLAVE DEL PERIODO</strong><br>
                    • Cobertura promedio de prospección: {cobertura_promedio:.0f}%<br>
                    • Tasa de conversión promedio: {conversion_promedio:.0f}%<br>
                    • Meta del periodo: {meta_periodo} visitas ({semanas_transcurridas} semanas × 25)<br>
                    • Zonas que superan 80% de avance: {len(zonas_bajo_avance) if zonas_bajo_avance else 0}/{total_zonas}
                </div>
            </div>
            
            <div class="table-container">
                <h3>📊 Ranking de Productividad por Zona</h3>
                <tr>
                    <thead>{headers_html}</thead>
                    <tbody>{filas_html}</tbody>
                </table>
            </div>
            
            <div class="alertas-table">
                <h3>🚦 Alertas Estratégicas para One-to-One</h3>
                <p>Priorizadas por nivel de urgencia. Basadas en momento del mes (Semana {semanas_transcurridas}).</p>
                <table>
                    <thead>
                        <tr>
                            <th style="background-color: #DC2626; color: white; padding: 10px;">PRIORIDAD</th>
                            <th style="background-color: #DC2626; color: white; padding: 10px;">ZONA</th>
                            <th style="background-color: #DC2626; color: white; padding: 10px;">ALERTA</th>
                            <th style="background-color: #DC2626; color: white; padding: 10px;">INSIGHT</th>
                            <th style="background-color: #DC2626; color: white; padding: 10px;">ACCIÓN PARA ONE-TO-ONE</th>
                        </tr>
                    </thead>
                    <tbody>{alertas_html}</tbody>
                </table>
            </div>
        </div>
        <div class="footer">
            <p>Go To Market SAC · Reporte generado automáticamente</p>
            <p>Para guardar como PDF: Ctrl+P → "Guardar como PDF"</p>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ============================================================
# FUNCIÓN PARA GENERAR REPORTE HTML SOLO TABLA DEL RANKING
# ============================================================
def generar_reporte_solo_tabla_html(df_ranking, mes_seleccionado, mostrar_opcion, meta_periodo, semanas_transcurridas):
    """Genera un reporte HTML solo con la tabla de datos del ranking"""
    
    traduccion_meses = {
        "January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril",
        "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto",
        "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"
    }
    
    mes_sel_esp = mes_seleccionado
    for eng, esp in traduccion_meses.items():
        if eng in mes_seleccionado:
            mes_sel_esp = mes_seleccionado.replace(eng, esp)
            break
    
    columnas = df_ranking.columns.tolist()
    headers_html = ""
    for col in columnas:
        headers_html += f'<th style="background-color: #DC2626; color: white; padding: 12px 15px; text-align: center; border: 1px solid #E2E8F0;">{col}</th>'
    
    filas_html = ""
    for _, row in df_ranking.iterrows():
        filas_html += '<tr>'
        for col in columnas:
            valor = row[col]
            if col == "POS":
                filas_html += f'<td style="background-color: #1E293B; color: white; font-weight: 800; text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            elif col == "ZONA":
                filas_html += f'<td style="text-align: left; font-weight: 600; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            elif col in ["% AVANCE", "COBERTURA", "TASA CONV."]:
                try:
                    num_val = float(str(valor).replace('%', ''))
                    filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{num_val:.0f}%</td>'
                except:
                    filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
            else:
                filas_html += f'<td style="text-align: center; padding: 10px; border: 1px solid #E2E8F0;">{valor}</td>'
        filas_html += '</tr>'
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GTM SAC - Tabla de Ranking</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Inter', sans-serif; background: #F8FAFC; padding: 40px; }}
        .report-container {{ max-width: 1400px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ font-size: 28px; }}
        .content {{ padding: 30px; }}
        .info {{ background: #F1F5F9; padding: 15px; border-radius: 12px; margin-bottom: 25px; }}
        .table-container {{ overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        th {{ background-color: #DC2626; color: white; padding: 12px; }}
        td {{ padding: 10px; border: 1px solid #E2E8F0; }}
        .footer {{ background: #F1F5F9; padding: 15px; text-align: center; font-size: 11px; color: #94A3B8; }}
    </style>
</head>
<body>
    <div class="report-container">
        <div class="header">
            <h1>Go To Market SAC</h1>
            <p>Ranking de Productividad - {mes_sel_esp}</p>
            <p>Mostrar: {mostrar_opcion} | {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </div>
        <div class="content">
            <div class="info">
                <strong>📊 DATOS GENERALES</strong><br>
                • Solo visitas FÍSICAS (regla comercial)<br>
                • Basado en CIERRES, no en tasas<br>
                • Meta del periodo: {meta_periodo} visitas ({semanas_transcurridas} semanas × 25)
            </div>
            <div class="table-container">
                <table>
                    <thead>{headers_html}</thead>
                    <tbody>{filas_html}</tbody>
                </table>
            </div>
        </div>
        <div class="footer">
            <p>Go To Market SAC · Reporte generado automáticamente</p>
            <p>Para guardar como PDF: Ctrl+P → "Guardar como PDF"</p>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ============================================================
# FUNCIÓN PARA GENERAR REPORTE HTML SOLO ALERTAS
# ============================================================
def generar_reporte_solo_alertas_html(alertas_data, mes_seleccionado, mostrar_opcion):
    """Genera un reporte HTML solo con las alertas para one-to-one"""
    
    traduccion_meses = {
        "January": "Enero", "February": "Febrero", "March": "Marzo", "April": "Abril",
        "May": "Mayo", "June": "Junio", "July": "Julio", "August": "Agosto",
        "September": "Septiembre", "October": "Octubre", "November": "Noviembre", "December": "Diciembre"
    }
    
    mes_sel_esp = mes_seleccionado
    for eng, esp in traduccion_meses.items():
        if eng in mes_seleccionado:
            mes_sel_esp = mes_seleccionado.replace(eng, esp)
            break
    
    alertas_html = ""
    for a in alertas_data:
        alertas_html += f"""
        <tr>
            <td style="padding: 12px; border: 1px solid #E2E8F0;">{a['prioridad']}</td>
            <td style="padding: 12px; border: 1px solid #E2E8F0; font-weight: 600;">{a['zona']}</td>
            <td style="padding: 12px; border: 1px solid #E2E8F0;">{a['alerta']}</td>
            <td style="padding: 12px; border: 1px solid #E2E8F0;">{a['insight']}</td>
            <td style="padding: 12px; border: 1px solid #E2E8F0;">{a['accion']}</td>
        </tr>
        """
    
    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GTM SAC - Alertas One-to-One</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Inter', sans-serif; background: #F8FAFC; padding: 40px; }}
        .report-container {{ max-width: 1200px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); color: white; padding: 30px; text-align: center; }}
        .header h1 {{ font-size: 28px; }}
        .content {{ padding: 30px; }}
        .info {{ background: #F1F5F9; padding: 15px; border-radius: 12px; margin-bottom: 25px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background-color: #DC2626; color: white; padding: 12px; text-align: center; }}
        td {{ padding: 12px; border: 1px solid #E2E8F0; }}
        .footer {{ background: #F1F5F9; padding: 15px; text-align: center; font-size: 11px; color: #94A3B8; }}
    </style>
</head>
<body>
    <div class="report-container">
        <div class="header">
            <h1>Go To Market SAC</h1>
            <p>Alertas Estratégicas - {mes_sel_esp}</p>
            <p>Mostrar: {mostrar_opcion} | {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        </div>
        <div class="content">
            <div class="info">
                <strong>🚦 ALERTAS PARA ONE-TO-ONE</strong><br>
                • Priorizadas por nivel de urgencia<br>
                • Basadas en momento del mes<br>
                • Usar en reuniones semanales con cada responsable de zona
            </div>
            <table>
                <thead>
                    <tr>
                        <th>PRIORIDAD</th>
                        <th>ZONA</th>
                        <th>ALERTA</th>
                        <th>INSIGHT</th>
                        <th>ACCIÓN PARA ONE-TO-ONE</th>
                    </tr>
                </thead>
                <tbody>
                    {alertas_html}
                </tbody>
            </table>
        </div>
        <div class="footer">
            <p>Go To Market SAC · Reporte generado automáticamente</p>
            <p>Para guardar como PDF: Ctrl+P → "Guardar como PDF"</p>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ============================================================
# PESTAÑA RANKING - EMBUDO DE PRODUCTIVIDAD CON ALERTAS ONE-TO-ONE
# ============================================================
def mostrar_pagina_ranking(df_datos):
    """Ranking con estructura de embudo de productividad y alertas para one-to-one"""
    
    st.markdown("""
    <div class="ranking-header">
        <div class="ranking-title">🏆 Ranking de Gestión de Calidad GTM</div>
        <div class="ranking-sub">Embudo de Productividad | Ordenado por visitas totales | ✅ Solo visitas FÍSICAS | Basado en CIERRES</div>
    </div>
    """, unsafe_allow_html=True)
    
    # ============================================================
    # 1. FILTROS PROPIOS DEL RANKING
    # ============================================================
    col_filtros1, col_filtros2 = st.columns(2)
    
    with col_filtros1:
        mostrar_opcion = st.radio(
            "📍 Mostrar:",
            options=["Nacional", "Lima", "Provincia"],
            horizontal=True,
            key="ranking_mostrar"
        )
    
    df_temp_meses = df_datos.copy()
    df_temp_meses['sabado'] = df_temp_meses['semana_inicio'] + pd.Timedelta(days=5)
    df_temp_meses['mes_sabado'] = df_temp_meses['sabado'].dt.strftime("%B %Y")
    df_temp_meses['fecha_sabado'] = df_temp_meses['sabado']
    
    df_ordenado = df_temp_meses.groupby('mes_sabado')['fecha_sabado'].min().reset_index()
    df_ordenado = df_ordenado.sort_values('fecha_sabado')
    meses_disponibles = df_ordenado['mes_sabado'].tolist()
    
    if not meses_disponibles:
        st.warning("No hay meses disponibles")
        return
    
    with col_filtros2:
        mes_seleccionado = st.selectbox(
            "📅 Periodo:",
            options=meses_disponibles,
            format_func=lambda x: x,
            index=len(meses_disponibles) - 1 if meses_disponibles else 0,
            key="ranking_periodo"
        )
    
    # ============================================================
    # 2. DETECTAR SEMANAS TRANSCURRIDAS
    # ============================================================
    df_temp_semanas = df_datos.copy()
    df_temp_semanas['sabado'] = df_temp_semanas['semana_inicio'] + pd.Timedelta(days=5)
    df_temp_semanas['mes_sabado'] = df_temp_semanas['sabado'].dt.strftime("%B %Y")
    
    semanas_del_mes = df_temp_semanas[df_temp_semanas['mes_sabado'] == mes_seleccionado]['semana'].unique()
    semanas_del_mes = sorted(semanas_del_mes, key=lambda x: int(x.split()[-1]))
    
    hoy = datetime.now()
    semanas_transcurridas = 0
    for semana in semanas_del_mes:
        df_temp_sab = df_temp_semanas[(df_temp_semanas['mes_sabado'] == mes_seleccionado) & 
                                       (df_temp_semanas['semana'] == semana)]
        if not df_temp_sab.empty:
            sabado_semana = df_temp_sab['sabado'].max()
            if sabado_semana <= hoy:
                semanas_transcurridas += 1
    
    if semanas_transcurridas == 0:
        semanas_transcurridas = len(semanas_del_mes)
    
    # ============================================================
    # 3. CALCULAR META DINÁMICA
    # ============================================================
    meta_periodo = META_SEMANAL * semanas_transcurridas
    
    # ============================================================
    # 4. OBTENER MES ANTERIOR
    # ============================================================
    idx_actual = meses_disponibles.index(mes_seleccionado)
    mes_anterior = meses_disponibles[idx_actual - 1] if idx_actual > 0 else None
    
    semanas_mes_anterior = []
    if mes_anterior:
        semanas_mes_anterior = df_temp_semanas[df_temp_semanas['mes_sabado'] == mes_anterior]['semana'].unique()
        semanas_mes_anterior = sorted(semanas_mes_anterior, key=lambda x: int(x.split()[-1]))[:semanas_transcurridas]
    
    # ============================================================
    # 5. CLASIFICACIÓN DE ZONAS
    # ============================================================
    todas_las_zonas = df_datos["zona"].unique().tolist()
    zona_clasificacion = {}
    for zona in todas_las_zonas:
        zona_upper = zona.upper().strip()
        if zona_upper in [z.upper().strip() for z in ZONAS_LIMA]:
            zona_clasificacion[zona] = "LIMA"
        elif zona_upper in [z.upper().strip() for z in ZONAS_PROVINCIA]:
            zona_clasificacion[zona] = "PROVINCIA"
        else:
            zona_clasificacion[zona] = "OTROS"
    
    # ============================================================
    # 6. CALCULAR DATOS POR ZONA
    # ============================================================
    datos_ranking = []
    
    for zona in todas_las_zonas:
        if mostrar_opcion == "Lima" and zona_clasificacion.get(zona) != "LIMA":
            continue
        if mostrar_opcion == "Provincia" and zona_clasificacion.get(zona) != "PROVINCIA":
            continue
        
        df_temp_zona = df_datos.copy()
        df_temp_zona['sabado'] = df_temp_zona['semana_inicio'] + pd.Timedelta(days=5)
        df_temp_zona['mes_sabado'] = df_temp_zona['sabado'].dt.strftime("%B %Y")
        df_mes_zona = df_temp_zona[(df_temp_zona['mes_sabado'] == mes_seleccionado) & (df_temp_zona["zona"] == zona)]
        
        df_periodo = df_mes_zona[df_mes_zona['semana'].isin(semanas_del_mes[:semanas_transcurridas])]
        visitas_actual = len(df_periodo)
        
        if visitas_actual == 0:
            continue
        
        pct_avance = (visitas_actual / meta_periodo * 100) if meta_periodo > 0 else 0
        
        visitas_por_semana = {}
        for i, semana in enumerate(semanas_del_mes[:semanas_transcurridas], 1):
            df_semana = df_periodo[df_periodo['semana'] == semana]
            visitas_por_semana[f"S{i}"] = len(df_semana)
        
        df_prospeccion = df_periodo[df_periodo["tipo"] == "PROSPECCIÓN"] if "tipo" in df_periodo.columns else pd.DataFrame()
        prospeccion_visitas = len(df_prospeccion)
        prospectos_unicos = df_prospeccion["Cliente o Prospecto"].nunique() if not df_prospeccion.empty else 0
        
        if "Task" in df_periodo.columns:
            cierres = len(df_periodo[(df_periodo["tipo"] == "PROSPECCIÓN") & 
                                      (df_periodo["Task"].str.upper().str.contains("CIERRE", na=False))])
        else:
            cierres = 0
        
        tasa_conversion = (cierres / prospectos_unicos * 100) if prospectos_unicos > 0 else 0
        cobertura = (prospectos_unicos / prospeccion_visitas * 100) if prospeccion_visitas > 0 else 0
        mantenimiento_visitas = len(df_periodo[df_periodo["tipo"] == "MANTENIMIENTO"]) if "tipo" in df_periodo.columns else 0
        
        visitas_anterior = 0
        crecimiento = 0
        if mes_anterior and semanas_mes_anterior:
            df_temp_ant = df_datos.copy()
            df_temp_ant['sabado'] = df_temp_ant['semana_inicio'] + pd.Timedelta(days=5)
            df_temp_ant['mes_sabado'] = df_temp_ant['sabado'].dt.strftime("%B %Y")
            df_mes_anterior_zona = df_temp_ant[(df_temp_ant['mes_sabado'] == mes_anterior) & 
                                                (df_temp_ant["zona"] == zona)]
            df_periodo_anterior = df_mes_anterior_zona[df_mes_anterior_zona['semana'].isin(semanas_mes_anterior)]
            visitas_anterior = len(df_periodo_anterior)
            
            if visitas_anterior > 0:
                crecimiento = ((visitas_actual - visitas_anterior) / visitas_anterior) * 100
        
        datos_ranking.append({
            "zona": zona,
            "visitas_actual": visitas_actual,
            "pct_avance": pct_avance,
            "visitas_por_semana": visitas_por_semana,
            "prospeccion_visitas": prospeccion_visitas,
            "prospectos_unicos": prospectos_unicos,
            "cierres": cierres,
            "cobertura": cobertura,
            "tasa_conversion": tasa_conversion,
            "mantenimiento_visitas": mantenimiento_visitas,
            "visitas_anterior": visitas_anterior,
            "crecimiento": crecimiento
        })
    
    if not datos_ranking:
        st.warning(f"No hay datos para la categoría seleccionada")
        return
    
    datos_ranking.sort(key=lambda x: x["visitas_actual"], reverse=True)
    for i, item in enumerate(datos_ranking, 1):
        item["posicion"] = i
    
        # ============================================================
    # 7. CONSTRUIR TABLA PRINCIPAL
    # ============================================================
    datos_para_df = []
    mes_anterior_nombre = mes_anterior.split()[0] if mes_anterior else "N/A"
    
    for d in datos_ranking:
        fila = {
            "POS": d['posicion'],
            "ZONA": d['zona'],
            "% CUMPLIMIENTO": d['pct_avance'],  # ← CAMBIADO
            "VISITAS": d['visitas_actual'],
        }
        for i in range(1, semanas_transcurridas + 1):
            fila[f"SEM {i}"] = d['visitas_por_semana'].get(f"S{i}", 0)  # ← CAMBIADO
        
        fila["PROSPECCIÓN"] = d['prospeccion_visitas']
        fila["PROSPECTOS ÚNICOS"] = d['prospectos_unicos']
        fila["CIERRES"] = d['cierres']
        fila["COBERTURA"] = d['cobertura']
        fila["MANTENIMIENTO"] = d['mantenimiento_visitas']
        fila["TASA CONV."] = d['tasa_conversion']
        fila[f"{mes_anterior_nombre} (mismo rango)"] = d['visitas_anterior']
        crecimiento_str = f"▲ {d['crecimiento']:.1f}%" if d['crecimiento'] >= 0 else f"▼ {abs(d['crecimiento']):.1f}%"
        fila["MES ANTERIOR"] = crecimiento_str if d['crecimiento'] != 0 else "0%"  # ← CAMBIADO
        datos_para_df.append(fila)
    
    df_ranking_show = pd.DataFrame(datos_para_df)
    
    st.subheader("📊 Ranking de Productividad por Zona")
    
    column_config = {
        "POS": st.column_config.TextColumn("POS", width="small"),
        "ZONA": st.column_config.TextColumn("ZONA", width="medium"),
        "% CUMPLIMIENTO": st.column_config.ProgressColumn("% CUMPLIMIENTO", format="%.0f%%", min_value=0, max_value=100, width="small"),  # ← CAMBIADO
        "VISITAS": st.column_config.NumberColumn("VISITAS", width="small"),
    }
    for i in range(1, semanas_transcurridas + 1):
        column_config[f"SEM {i}"] = st.column_config.NumberColumn(f"SEM {i}", width="small")  # ← CAMBIADO
    column_config.update({
        "PROSPECCIÓN": st.column_config.NumberColumn("PROSPECCIÓN", width="medium"),
        "PROSPECTOS ÚNICOS": st.column_config.NumberColumn("PROSPECTOS ÚNICOS", width="medium"),
        "CIERRES": st.column_config.NumberColumn("CIERRES", width="small"),
        "COBERTURA": st.column_config.ProgressColumn("COBERTURA", format="%.0f%%", min_value=0, max_value=100, width="small"),
        "MANTENIMIENTO": st.column_config.NumberColumn("MANTENIMIENTO", width="medium"),
        "TASA CONV.": st.column_config.ProgressColumn("TASA CONV.", format="%.0f%%", min_value=0, max_value=100, width="small"),
        f"{mes_anterior_nombre} (mismo rango)": st.column_config.NumberColumn(f"{mes_anterior_nombre} (mismo rango)", width="medium"),
        "MES ANTERIOR": st.column_config.TextColumn("MES ANTERIOR", width="small"),  # ← CAMBIADO
    })
    
    
    st.dataframe(df_ranking_show, use_container_width=True, hide_index=True, column_config=column_config)
    
    # ============================================================
    # 8. ALERTAS ESTRATÉGICAS (CON FÓRMULA PONDERADA)
    # ============================================================
    st.subheader("🚦 Alertas Estratégicas para One-to-One")
    st.caption(f"📅 Semana {semanas_transcurridas} de {len(semanas_del_mes)} | Prioridad calculada con fórmula ponderada")
    
    # Función para calcular factor según rango
    def get_factor_avance(pct):
        if pct < 60:
            return 3  # 🔴 ALTA
        elif pct < 80:
            return 2  # 🟡 MEDIA
        else:
            return 1  # 🟢 BAJA
    
    def get_factor_gestion(mantenimiento, prospeccion, semana):
        total = mantenimiento + prospeccion
        if total == 0:
            return 2
        pct_mantenimiento = (mantenimiento / total) * 100
        
        if semana <= 2:  # Semanas 1-2: enfoque en prospección
            if pct_mantenimiento >= 80:
                return 3  # 🔴 ALTA
            elif pct_mantenimiento >= 60:
                return 2  # 🟡 MEDIA
            else:
                return 1  # 🟢 BAJA
        else:  # Semanas 3-5: enfoque en mantenimiento
            if pct_mantenimiento >= 90:
                return 1  # 🟢 BAJA (bueno, está haciendo mantenimiento)
            elif pct_mantenimiento >= 80:
                return 2  # 🟡 MEDIA
            else:
                return 1  # 🟢 BAJA
    
    def get_factor_cobertura(cobertura):
        if cobertura >= 80:
            return 3  # 🔴 ALTA (visita solo una vez)
        elif cobertura >= 50:
            return 2  # 🟡 MEDIA
        else:
            return 1  # 🟢 BAJA
    
    def get_factor_cierres(cierres, semana):
        if semana <= 2:  # Semanas 1-2: no se evalúa cierres
            return 1
        if cierres < 2:
            return 3  # 🔴 ALTA
        elif cierres >= 3:
            return 1  # 🟢 BAJA
        else:
            return 2  # 🟡 MEDIA
    
    # Pesos
    PESO_AVANCE = 0.40
    PESO_GESTION = 0.25
    PESO_COBERTURA = 0.25
    PESO_CIERRES = 0.10
    
    alertas_data = []
    
    for d in datos_ranking:
        # Calcular factores
        factor_avance = get_factor_avance(d['pct_avance'])
        factor_gestion = get_factor_gestion(d['mantenimiento_visitas'], d['prospeccion_visitas'], semanas_transcurridas)
        factor_cobertura = get_factor_cobertura(d['cobertura'])
        factor_cierres = get_factor_cierres(d['cierres'], semanas_transcurridas)
        
        # Calcular peso total
        peso_total = (PESO_AVANCE * factor_avance) + (PESO_GESTION * factor_gestion) + (PESO_COBERTURA * factor_cobertura) + (PESO_CIERRES * factor_cierres)
        
        # Determinar prioridad final
        if peso_total >= 2.3:
            prioridad_final = "🔴 ALTA"
        elif peso_total >= 1.5:
            prioridad_final = "🟡 MEDIA"
        else:
            prioridad_final = "🟢 BAJA"
        
        # Determinar la alerta principal según el factor más crítico
        factores = {
            "avance": factor_avance,
            "gestion": factor_gestion,
            "cobertura": factor_cobertura,
            "cierres": factor_cierres
        }
        factor_max = max(factores.values())
        
        if factor_max == factor_avance and factor_avance >= 2:
            alerta_nombre = "Bajo avance en visitas"
            insight_texto = f"{d['pct_avance']:.0f}% de avance ({d['visitas_actual']}/{meta_periodo} visitas)"
            accion_texto = f"Vamos por la semana {semanas_transcurridas} y llevas solo {d['pct_avance']:.0f}% de la meta. Necesitas activar más visitas."
        elif factor_max == factor_gestion and factor_gestion >= 2:
            alerta_nombre = "Distribución incorrecta de visitas"
            insight_texto = f"{d['mantenimiento_visitas']} mantenimiento vs {d['prospeccion_visitas']} prospección"
            if semanas_transcurridas <= 2:
                accion_texto = "Estás haciendo más visitas de mantenimiento que de prospección. Las primeras semanas son para PROSPECTAR."
            else:
                accion_texto = "Revisa tu distribución de visitas entre prospección y mantenimiento."
        elif factor_max == factor_cobertura and factor_cobertura >= 2:
            alerta_nombre = "Prospectos visitados solo una vez"
            insight_texto = f"{d['prospeccion_visitas']} visitas a {d['prospectos_unicos']} prospectos únicos ({d['cobertura']:.0f}% cobertura)"
            accion_texto = "Estás visitando a tus prospectos solo una vez. Dale una segunda visita para poder cerrar."
        elif factor_max == factor_cierres and factor_cierres >= 2 and semanas_transcurridas > 2:
            alerta_nombre = "Sin cierres suficientes"
            insight_texto = f"{d['cierres']} cierres en el periodo"
            accion_texto = "Necesitas cerrar más prospectos. Revisemos tu técnica de cierre."
        else:
            # Si todos los factores son bajos, es una zona con buen desempeño
            if d['cierres'] >= 3:
                alerta_nombre = "¡Cierres destacados!"
                insight_texto = f"{d['cierres']} cierres en el periodo"
                accion_texto = "Reconocimiento público. Comparte tu método con el equipo."
                prioridad_final = "🟢 BAJA"
            else:
                continue  # No mostrar alerta si no hay nada crítico
        
        alertas_data.append({
            "prioridad": prioridad_final,
            "zona": d['zona'],
            "alerta": alerta_nombre,
            "insight": insight_texto,
            "accion": accion_texto,
            "peso_total": peso_total
        })
    
    if alertas_data:
        # Ordenar por prioridad (ALTA → MEDIA → BAJA) y luego por peso_total descendente
        orden_prioridad = {"🔴 ALTA": 0, "🟡 MEDIA": 1, "🟢 BAJA": 2}
        alertas_data.sort(key=lambda x: (orden_prioridad.get(x["prioridad"], 3), -x["peso_total"]))
        
        df_alertas = pd.DataFrame(alertas_data)
        st.dataframe(
            df_alertas[["prioridad", "zona", "alerta", "insight", "accion"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "prioridad": st.column_config.TextColumn("PRIORIDAD", width="small"),
                "zona": st.column_config.TextColumn("ZONA", width="medium"),
                "alerta": st.column_config.TextColumn("ALERTA", width="medium"),
                "insight": st.column_config.TextColumn("INSIGHT", width="large"),
                "accion": st.column_config.TextColumn("ACCIÓN PARA ONE-TO-ONE", width="large"),
            }
        )
        
        # Mostrar resumen de prioridades
        st.markdown(f"""
        <div style="background: #F1F5F9; padding: 12px; border-radius: 10px; margin-top: 10px;">
            <strong>📊 RESUMEN DE PRIORIDADES:</strong><br>
            🔴 ALTA: {len([a for a in alertas_data if a['prioridad'] == '🔴 ALTA'])} zonas | 
            🟡 MEDIA: {len([a for a in alertas_data if a['prioridad'] == '🟡 MEDIA'])} zonas | 
            🟢 BAJA: {len([a for a in alertas_data if a['prioridad'] == '🟢 BAJA'])} zonas
        </div>
        """, unsafe_allow_html=True)
    else:
        st.success("✅ No hay alertas activas. ¡Excelente desempeño general!")
    
    # ============================================================
    # 9. BOTONES DE DESCARGA
    # ============================================================
    st.markdown("---")
    st.subheader("📥 Descargar Reportes")
    st.caption("Los reportes se descargan en formato HTML (abrir en navegador y Ctrl+P para PDF)")
    
    col_boton1, col_boton2, col_boton3 = st.columns(3)
    
    with col_boton1:
        ranking_html_completo = generar_reporte_ranking_completo_html(
            df_ranking_show, mes_seleccionado, mostrar_opcion, meta_periodo, semanas_transcurridas,
            alertas_data, 0, 0, 0,  # total_cierres, cobertura_promedio, conversion_promedio (no se usan)
            0, 0, 0, len(datos_ranking), [], [], [], [], [], []
        )
        st.download_button(
            label="📊 Ranking de Productividad (HTML)",
            data=ranking_html_completo,
            file_name=f"ranking_productividad_{mes_seleccionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True
        )
    
    with col_boton2:
        if alertas_data:
            alertas_html = generar_reporte_solo_alertas_html(
                alertas_data, mes_seleccionado, mostrar_opcion
            )
            st.download_button(
                label="⚠️ Alertas One-to-One (HTML)",
                data=alertas_html,
                file_name=f"alertas_one_to_one_{mes_seleccionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                mime="text/html",
                use_container_width=True
            )
        else:
            st.info("ℹ️ No hay alertas para descargar")
    
    with col_boton3:
        solo_tabla_html = generar_reporte_solo_tabla_html(
            df_ranking_show, mes_seleccionado, mostrar_opcion, meta_periodo, semanas_transcurridas
        )
        st.download_button(
            label="📋 Solo Tabla de Datos (HTML)",
            data=solo_tabla_html,
            file_name=f"tabla_ranking_{mes_seleccionado.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True
        )

# ─── SIDEBAR - FILE UPLOADER Y FILTROS ─────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Panel de Control")
    uploaded = st.file_uploader("📂 Carga tu archivo Excel", type=["xlsx", "xls"])
    
    if uploaded is None:
        st.info("👈 Carga un archivo Excel para comenzar")
        st.stop()
    
    st.markdown("---")
    st.markdown("### 🎯 Filtros")
    
    df, err = load_excel(uploaded.read())
    if err: 
        st.error(err)
        st.stop()
        
    
    # ✅ FILTRO DE VISITAS FÍSICAS (REGLAS DE NEGOCIO)
    if "tipo_visita" in df.columns:
        df = df[df["tipo_visita"] == "FÍSICA"].copy()
        st.success("✅ Solo se consideran visitas FÍSICAS (regla comercial)")
    else:
        st.warning("⚠️ Columna 'tipo_visita' no encontrada. Verifica el archivo.")
    
    sel_mes = st.selectbox("📅 Periodo Mensual", ["Todos"] + sorted(df["mes"].unique().tolist()))
    
    df_temp = df.copy()
    
    if sel_mes != "Todos":
        df_temp['sabado'] = df_temp['semana_inicio'] + pd.Timedelta(days=5)
        df_temp['mes_sabado'] = df_temp['sabado'].dt.strftime("%B %Y")
        df_f = df_temp[df_temp['mes_sabado'] == sel_mes].copy()
        df_f = df_f.drop(columns=['sabado', 'mes_sabado'])
    else:
        df_f = df_temp.copy()
    
    semanas_temp = sorted(df_f["semana"].unique(), key=lambda x: int(x.split()[-1]))
    lista_semanas = semanas_temp if len(semanas_temp) > 0 else []
    sel_sem = st.selectbox("🗓️ Vista por Semana", ["Todas"] + lista_semanas)
    
    opciones_zona = sorted(df_f["zona"].dropna().unique().tolist())
    if opciones_zona:
        sel_zona = st.selectbox("🌎 Zona o Territorio", opciones_zona, index=0)
        df_f = df_f[df_f["zona"] == sel_zona]
    else:
        sel_zona = None
        st.warning("⚠️ No hay zonas disponibles en los datos")
    
    if sel_sem != "Todas": 
        df_f = df_f[df_f["semana"] == sel_sem]

if df_f.empty:
    st.warning("⚠️ No hay datos con los filtros seleccionados. Ajusta los criterios.")
    st.stop()

# ─── SELECTOR DE PESTAÑAS ─────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 DASHBOARD", "🏆 RANKING"])

with tab1:
    # ============================================================
    # DASHBOARD CON METAS PERSONALIZADAS POR ZONA
    # ============================================================
    if sel_mes != "Todos" and sel_sem == "Todas":
        num_semanas_visibles = len(df_f["semana"].unique())
    else:
        num_semanas_visibles = df_f["semana"].nunique()

    zonas_activas = df_f["zona"].nunique()
    
    # ============================================================
    # METAS SEGÚN ZONA (PERSONALIZADO)
    # ============================================================
    if sel_zona:
        meta_diaria_zona, meta_semanal_zona = get_metas_zona(sel_zona)
    else:
        meta_diaria_zona, meta_semanal_zona = 5, 25
    
    if sel_sem == "Todas":
        meta_actual = meta_semanal_zona * num_semanas_visibles * zonas_activas
    else:
        meta_actual = meta_semanal_zona

    zona_badge = ""
    if sel_zona:
        tipo_zona_sel = CLASIFICACION_ZONAS.get(sel_zona.upper().strip(), "SIN CLASIFICAR")
        zona_badge = f'<span class="zona-badge">📍 {tipo_zona_sel} - {sel_zona}</span>'

    st.markdown(f'<div class="section-title">📊 Indicadores de Gestión {zona_badge}</div>', unsafe_allow_html=True)

    total_v = len(df_f)
    visitas_prospeccion = len(df_f[df_f["tipo"] == "PROSPECCIÓN"]) if "tipo" in df_f.columns else 0
    visitas_mantenimiento = len(df_f[df_f["tipo"] == "MANTENIMIENTO"]) if "tipo" in df_f.columns else 0

    # ============================================================
    # CALCULAR META ACUMULADA SEGÚN EL DÍA (con meta personalizada)
    # ============================================================
    hoy = datetime.now()
    dia_actual_num = hoy.weekday()
    
    if dia_actual_num == 1:  # Martes
        meta_acumulada = meta_diaria_zona
    elif dia_actual_num == 2:  # Miércoles
        meta_acumulada = meta_diaria_zona * 2
    elif dia_actual_num == 3:  # Jueves
        meta_acumulada = meta_diaria_zona * 3
    elif dia_actual_num == 4:  # Viernes
        meta_acumulada = meta_diaria_zona * 4
    else:
        meta_acumulada = meta_semanal_zona
    
    # Calcular visitas acumuladas del día
    lunes_actual = hoy - timedelta(days=dia_actual_num)
    visitas_acumuladas = 0
    for i in range(dia_actual_num + 1):
        fecha_dia = lunes_actual + timedelta(days=i)
        visitas = len(df_f[df_f["fecha"].dt.date == fecha_dia.date()])
        visitas_acumuladas += visitas
    
    pct_avance = (visitas_acumuladas / meta_acumulada * 100) if meta_acumulada > 0 else 0
    
    # Calcular prospectos únicos para tasa de conversión
    df_prospeccion_dash = df_f[df_f["tipo"] == "PROSPECCIÓN"] if "tipo" in df_f.columns else pd.DataFrame()
    prospectos_unicos_dash = df_prospeccion_dash["Cliente o Prospecto"].nunique() if not df_prospeccion_dash.empty else 0
    
    if "Task" in df_f.columns:
        cierres_dash = len(df_f[(df_f["tipo"] == "PROSPECCIÓN") & 
                                  (df_f["Task"].str.upper().str.contains("CIERRE", na=False))])
    else:
        cierres_dash = 0
    
    tasa_conversion = (cierres_dash / prospectos_unicos_dash * 100) if prospectos_unicos_dash > 0 else 0
    
    # Colores según rendimiento
    if pct_avance >= 80:
        avance_color = "#16A34A"
    elif pct_avance >= 50:
        avance_color = "#F59E0B"
    else:
        avance_color = "#DC2626"
    
    # ============================================================
    # 5 KPIs
    # ============================================================
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1: 
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Visitas</div>
            <div class="kpi-value">{total_v}</div>
            <div class="kpi-desglose">
                🎯 Prospección = {visitas_prospeccion}<br>
                🔧 Mantenimiento = {visitas_mantenimiento}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with c2: 
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Clientes Únicos</div><div class="kpi-value">{df_f["Cliente o Prospecto"].nunique()}</div></div>', unsafe_allow_html=True)

    color_clase = "green" if total_v >= meta_actual else "red"
    with c3: 
        st.markdown(f'<div class="kpi-card {color_clase}"><div class="kpi-label">Meta del Periodo</div><div class="kpi-value">{total_v}/{meta_actual}</div></div>', unsafe_allow_html=True)
    
    with c4:
        if visitas_acumuladas >= meta_acumulada:
            color_avance_kpi = "#16A34A"
        else:
            color_avance_kpi = "#DC2626"
        
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">📈 % de Avance</div>
            <div class="kpi-value" style="color: {color_avance_kpi};">{pct_avance:.1f}%</div>
            <div class="kpi-desglose">
                Meta: {meta_acumulada} visitas (acumulado)
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with c5:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">🎯 Cierres en la semana</div>
            <div class="kpi-value">{cierres_dash}</div>
        </div>
        """, unsafe_allow_html=True)

    # ============================================================
    # GRÁFICO + ALERTA
    # ============================================================
    col_grafico, col_alerta = st.columns([2, 1])
    
    with col_grafico:
        st.markdown('<p class="section-title">📈 Análisis de Tendencia Semanal</p>', unsafe_allow_html=True)

        if sel_sem == "Todas":
            df_filtrada = df_f
            df_graph_detalle = df_filtrada.groupby(["semana", "tipo"]).size().reset_index(name="visitas")
            df_graph_pivot = df_graph_detalle.pivot(index="semana", columns="tipo", values="visitas").fillna(0).reset_index()
            
            if "PROSPECCIÓN" not in df_graph_pivot.columns:
                df_graph_pivot["PROSPECCIÓN"] = 0
            if "MANTENIMIENTO" not in df_graph_pivot.columns:
                df_graph_pivot["MANTENIMIENTO"] = 0
            
            df_graph_pivot["Total"] = df_graph_pivot["PROSPECCIÓN"] + df_graph_pivot["MANTENIMIENTO"]
            df_graph_pivot["hover_text"] = df_graph_pivot.apply(
                lambda r: f"Total: {int(r['Total'])}<br>🎯 Prospección: {int(r['PROSPECCIÓN'])}<br>🔧 Mantenimiento: {int(r['MANTENIMIENTO'])}", axis=1
            )
            
            semanas_rango = df_filtrada.groupby("semana")["semana_rango"].first().to_dict()
            df_graph_pivot["semana_rango"] = df_graph_pivot["semana"].map(semanas_rango)
            
            fig = px.bar(df_graph_pivot, x="semana", y="Total", text="Total",
                         color_discrete_sequence=["#CBD5E1"],
                         hover_data={"hover_text": True, "semana_rango": True})
            fig.add_hline(y=meta_semanal_zona, line_dash="dash", line_color="#1E293B", annotation_text=f"Meta {meta_semanal_zona}")
            fig.update_traces(textposition='outside', hovertemplate="%{customdata[0]}<br>📅 %{customdata[1]}<extra></extra>",
                              texttemplate='%{text:.0f}')
            
            data_alerts = df_graph_pivot.rename(columns={"semana": "semana", "Total": "visitas", "semana_rango": "semana_rango"})
            meta_ref = meta_semanal_zona
        else:
            dias_es = {"Monday":"Lun", "Tuesday":"Mar", "Wednesday":"Mie", "Thursday":"Jue", "Friday":"Vie", "Saturday":"Sab", "Sunday":"Dom"}
            
            df_dia_detalle = df_f.groupby(["fecha", "dia_semana", "tipo"]).size().reset_index(name="visitas").sort_values("fecha")
            df_dia_pivot = df_dia_detalle.pivot(index=["fecha", "dia_semana"], columns="tipo", values="visitas").fillna(0).reset_index()
            
            if "PROSPECCIÓN" not in df_dia_pivot.columns:
                df_dia_pivot["PROSPECCIÓN"] = 0
            if "MANTENIMIENTO" not in df_dia_pivot.columns:
                df_dia_pivot["MANTENIMIENTO"] = 0
            
            df_dia_pivot["Total"] = df_dia_pivot["PROSPECCIÓN"] + df_dia_pivot["MANTENIMIENTO"]
            df_dia_pivot["label"] = df_dia_pivot.apply(lambda r: f"{dias_es.get(r['dia_semana'])} {r['fecha'].strftime('%d/%m')}", axis=1)
            df_dia_pivot["hover_text"] = df_dia_pivot.apply(
                lambda r: f"Total: {int(r['Total'])}<br>🎯 Prospección: {int(r['PROSPECCIÓN'])}<br>🔧 Mantenimiento: {int(r['MANTENIMIENTO'])}", axis=1
            )
            
            fig = px.bar(df_dia_pivot.sort_values("fecha"), x="label", y="Total", text="Total",
                         color_discrete_sequence=["#CBD5E1"],
                         hover_data={"hover_text": True})
            fig.add_hline(y=meta_diaria_zona, line_dash="dash", line_color="#1E293B", annotation_text=f"Meta {meta_diaria_zona}")
            fig.update_traces(textposition='outside', hovertemplate="%{customdata[0]}<extra></extra>",
                              texttemplate='%{text:.0f}')
            
            data_alerts = df_dia_pivot.rename(columns={"label": "label", "Total": "visitas"})
            meta_ref = meta_diaria_zona

        fig.update_layout(height=380, bargap=0.5, margin=dict(t=20, b=20, l=20, r=20), 
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with col_alerta:
        if sel_zona and sel_sem != "Todas":
            alerta_html = generar_alertas_simple(df_f, sel_zona, total_v)
            if alerta_html:
                st.markdown(alerta_html, unsafe_allow_html=True)
        else:
            st.markdown('<div style="height: 100%; display: flex; align-items: center; justify-content: center; padding: 20px; background: #F8FAFC; border-radius: 12px;"><span style="color: #94A3B8;">Selecciona una zona y semana para ver alertas</span></div>', unsafe_allow_html=True)

    # ============================================================
    # ALERTAS DE CUMPLIMIENTO (CON COLORES PERSONALIZADOS)
    # ============================================================
    st.markdown('<p class="section-title">🚦 Alertas de Cumplimiento</p>', unsafe_allow_html=True)
    if not data_alerts.empty:
        if sel_sem == "Todas":
            num_alerts = len(data_alerts)
            cols_a = st.columns(min(num_alerts, 4))
            for i, (idx, row) in enumerate(data_alerts.iterrows()):
                if i >= 4: break
                cumple = row["visitas"] >= meta_semanal_zona
                bg_color = "#DCFCE7" if cumple else "#FEE2E2"
                color_text = "#16A34A" if cumple else "#DC2626"
                
                semana = row["semana"]
                df_semana = df_filtrada[df_filtrada["semana"] == semana]
                prospeccion_sem = len(df_semana[df_semana["tipo"] == "PROSPECCIÓN"])
                mantenimiento_sem = len(df_semana[df_semana["tipo"] == "MANTENIMIENTO"])
                
                texto_semana = row["semana"]
                rango_fechas = row["semana_rango"]
                with cols_a[i]:
                    st.markdown(f"""
                    <div style="background:{bg_color}; border-radius:12px; padding:12px; text-align:center;">
                        <div style="font-weight:700; font-size:11px;">{texto_semana}</div>
                        <div style="font-size:11px; color:#64748B; margin-bottom:5px;">{rango_fechas}</div>
                        <div style="font-size:28px; font-weight:800; color:{color_text};">{int(row['visitas'])}</div>
                        <div style="font-size:11px; margin-top:8px; padding-top:5px; border-top:1px solid #E2E8F0;">
                            🎯 Prospección = {prospeccion_sem}<br>
                            🔧 Mantenimiento = {mantenimiento_sem}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            num_alerts = len(data_alerts)
            cols_a = st.columns(num_alerts)
            for i, (idx, row) in enumerate(data_alerts.iterrows()):
                cumple = row["visitas"] >= meta_diaria_zona
                bg_color = "#DCFCE7" if cumple else "#FEE2E2"
                color_text = "#16A34A" if cumple else "#DC2626"
                fecha = row["fecha"]
                df_dia = df_f[df_f["fecha"] == fecha]
                prospeccion_dia = len(df_dia[df_dia["tipo"] == "PROSPECCIÓN"])
                mantenimiento_dia = len(df_dia[df_dia["tipo"] == "MANTENIMIENTO"])
                with cols_a[i]:
                    st.markdown(f"""
                    <div style="background:{bg_color}; border-radius:12px; padding:12px; text-align:center;">
                        <div style="font-weight:700; font-size:12px;">{row['label']}</div>
                        <div style="font-size:28px; font-weight:800; color:{color_text};">{int(row['visitas'])}</div>
                        <div style="font-size:11px; margin-top:8px; padding-top:5px; border-top:1px solid #E2E8F0;">
                            🎯 Prospección = {prospeccion_dia}<br>
                            🔧 Mantenimiento = {mantenimiento_dia}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)    

    st.markdown('<p class="section-title">📋 Resumen Estratégico por Giro</p>', unsafe_allow_html=True)
    df_resumen_giro = df_f.groupby("Giro").agg(
        Clientes=("Cliente o Prospecto", "nunique"),
        Visitas_Totales=("Cliente o Prospecto", "count")
    ).reset_index()
    df_resumen_giro["Frecuencia"] = (df_resumen_giro["Visitas_Totales"] / df_resumen_giro["Clientes"]).round(2)
    st.dataframe(df_resumen_giro, use_container_width=True, hide_index=True)

    st.markdown('<p class="section-title">🍩 Distribución de Visitas por Giro</p>', unsafe_allow_html=True)
    visitas_por_giro = df_f.groupby("Giro").size().reset_index(name="Visitas").sort_values("Visitas", ascending=False)
    colores_dona = ["#DC2626", "#F59E0B", "#10B981", "#3B82F6", "#8B5CF6", "#EC4899", "#06B6D4", "#84CC16"]
    fig_dona_giro = go.Figure(data=[go.Pie(labels=visitas_por_giro["Giro"], values=visitas_por_giro["Visitas"], hole=0.45,
                                            marker_colors=colores_dona[:len(visitas_por_giro)], textinfo='label+percent', textposition='auto')])
    fig_dona_giro.update_layout(height=420, margin=dict(t=20, b=20, l=20, r=20), paper_bgcolor='white',
                                annotations=[dict(text=f"Total<br>{len(df_f)}", x=0.5, y=0.5, font_size=16, 
                                                  font_weight='bold', showarrow=False, font_color="#1E293B")])
    st.plotly_chart(fig_dona_giro, use_container_width=True)

    giro_top = visitas_por_giro.iloc[0]["Giro"]
    visitas_top = visitas_por_giro.iloc[0]["Visitas"]
    porcentaje_top = (visitas_top / len(df_f) * 100).round(1)
    st.markdown(f"""
    <div style="background-color: #F1F5F9; border-radius: 10px; padding: 12px; margin-top: 10px; text-align: center;">
        <span style="font-size: 13px;">🎯 Giro con más visitas:</span>
        <span style="font-size: 16px; font-weight: 700; color: #DC2626;">{giro_top}</span>
        <span style="font-size: 13px;"> con {visitas_top} visitas ({porcentaje_top}% del total)</span>
    </div>
    """, unsafe_allow_html=True)

    if sel_zona:
        st.markdown('<p class="section-title">📂 Detalle de Clientes por Giro</p>', unsafe_allow_html=True)
        giros = [g for g in df_f["Giro"].unique() if pd.notna(g)]
        if giros:
            for i in range(0, len(giros), 2):
                trozo = giros[i : i + 2]
                cols = st.columns(2)
                for j, giro_nombre in enumerate(trozo):
                    with cols[j]:
                        st.markdown(f'<div style="background:#F1F5F9; padding:8px 15px; border-radius:8px; border-left:4px solid #DC2626; margin-bottom:10px; font-weight:700;">📍 {giro_nombre}</div>', unsafe_allow_html=True)
                        df_det = df_f[df_f["Giro"] == giro_nombre].groupby("Cliente o Prospecto").size().reset_index(name="Visitas").sort_values("Visitas", ascending=False)
                        st.dataframe(df_det, use_container_width=True, hide_index=True, height=200)

        # ============================================================
        # EMBUDO DE VENTAS
        # ============================================================
        st.markdown('<p class="section-title">📈 EMBUDO DE VENTAS - Conversión y Velocidad</p>', unsafe_allow_html=True)
        
        etapas = ["PROSPECCIÓN", "CALIFICACIÓN DE LEADS", "VISITA", "PROPUESTA", "NEGOCIACIÓN", "CIERRE"]
        
        orden_etapas = {
            "PROSPECCIÓN": 1,
            "CALIFICACIÓN DE LEADS": 2,
            "VISITA": 3,
            "PROPUESTA": 4,
            "NEGOCIACIÓN": 5,
            "CIERRE": 6
        }
        
        colores_embudo = {
            "PROSPECCIÓN": "#DC2626",
            "CALIFICACIÓN DE LEADS": "#F59E0B",
            "VISITA": "#3B82F6",
            "PROPUESTA": "#10B981",
            "NEGOCIACIÓN": "#8B5CF6",
            "CIERRE": "#16A34A"
        }
        
        df_embudo = df.copy()
        
        if sel_zona:
            df_embudo = df_embudo[df_embudo["zona"] == sel_zona]
        
        fecha_inicio = datetime(2026, 4, 1)
        if "fecha" in df_embudo.columns:
            df_embudo["fecha_dt"] = pd.to_datetime(df_embudo["fecha"])
            df_embudo = df_embudo[df_embudo["fecha_dt"] >= fecha_inicio]
        
        semanas_filtro = []
        if sel_sem != "Todas":
            semanas_filtro = [sel_sem]
        elif sel_mes != "Todos":
            semanas_filtro = df_f["semana"].unique().tolist() if not df_f.empty else []
        
        fecha_actual = datetime.now()
        df_semanas_cliente = df_embudo.groupby("Cliente o Prospecto")["semana"].apply(list).to_dict()
        
        df_cliente_etapas = df_embudo.groupby("Cliente o Prospecto").agg({
            "Task": lambda x: list(x),
            "fecha_dt": ["min", "max"]
        }).reset_index()
        df_cliente_etapas.columns = ["Cliente o Prospecto", "etapas", "fecha_registro", "fecha_ultima"]
        
        cliente_etapa_final = {}
        for _, row in df_cliente_etapas.iterrows():
            nombre = row["Cliente o Prospecto"]
            etapas_cliente = row["etapas"]
            etapa_max = max(etapas_cliente, key=lambda x: orden_etapas.get(x, 0))
            cliente_etapa_final[nombre] = etapa_max
        
        visitas_por_cliente = df_embudo.groupby("Cliente o Prospecto").size().to_dict()
        
        clientes_data = {}
        for _, row in df_cliente_etapas.iterrows():
            nombre = row["Cliente o Prospecto"]
            fecha_reg = row["fecha_registro"]
            fecha_ult = row["fecha_ultima"]
            
            clientes_data[nombre] = {
                "fecha_registro": fecha_reg,
                "fecha_ultima": fecha_ult,
                "etapa_final": cliente_etapa_final[nombre],
                "visitas": visitas_por_cliente.get(nombre, 1)
            }
        
        datos_etapas = []
        
        for etapa in etapas:
            clientes_en_etapa = []
            
            for nombre, data in clientes_data.items():
                if data["etapa_final"] == etapa:
                    dias_en_etapa = (fecha_actual - data["fecha_registro"]).days
                    
                    if pd.notna(data["fecha_ultima"]):
                        dias_sin_visita = (fecha_actual - data["fecha_ultima"]).days
                        fecha_ultima_str = data["fecha_ultima"].strftime("%d/%m/%Y")
                    else:
                        dias_sin_visita = 999
                        fecha_ultima_str = "N/A"
                    
                    fecha_reg_str = data["fecha_registro"].strftime("%d/%m/%Y")
                    visitas = data["visitas"]
                    
                    semanas_cliente = df_semanas_cliente.get(nombre, [])
                    visitado_semana = "✅ Sí" if any(sem in semanas_filtro for sem in semanas_cliente) else "❌ No"
                    
                    if dias_en_etapa <= 30:
                        temperatura = "🔥 Caliente"
                    elif dias_en_etapa <= 60:
                        temperatura = "🟡 Tibio"
                    else:
                        temperatura = "❄️ Frío"
                    
                    clientes_en_etapa.append({
                        "nombre": nombre,
                        "fecha_reg": fecha_reg_str,
                        "dias": dias_en_etapa,
                        "visitas": visitas,
                        "fecha_ultima": fecha_ultima_str,
                        "dias_sin_visita": dias_sin_visita,
                        "visitado_semana": visitado_semana,
                        "temperatura": temperatura
                    })
            
            if clientes_en_etapa:
                dias_prom = round(sum(c["dias"] for c in clientes_en_etapa) / len(clientes_en_etapa))
                clientes_en_etapa.sort(key=lambda x: x["dias"], reverse=True)
                top5 = clientes_en_etapa[:5]
                total_rest = len(clientes_en_etapa) - 5
                urgentes = sum(1 for c in clientes_en_etapa if c["dias_sin_visita"] > 15)
                
                tooltip_lines = [
                    f"📊 Total: {len(clientes_en_etapa)} prospectos | ⏱️ Promedio: {dias_prom} días | 🔴 Urgentes: {urgentes}",
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
                    "🔥 Los más antiguos (prioridad):"
                ]
                for c in top5:
                    tooltip_lines.append(f"• {c['nombre']} - {c['dias']} días | {c['visitas']} visitas | {c['visitado_semana']}")
                if total_rest > 0:
                    tooltip_lines.append(f"... y {total_rest} más. Ver detalle completo abajo.")
                
                tooltip_text = "\n".join(tooltip_lines)
                clientes_vista = ", ".join([f"{c['nombre']} ({c['fecha_reg']})" for c in clientes_en_etapa[:3]])
                if len(clientes_en_etapa) > 3:
                    clientes_vista += f" (+{len(clientes_en_etapa)-3})"
            else:
                dias_prom = 0
                tooltip_text = "Sin prospectos en esta etapa"
                clientes_vista = "Sin prospectos"
                clientes_en_etapa = []
            
            datos_etapas.append({
                "etapa": etapa,
                "cantidad": len(clientes_en_etapa),
                "dias_prom": dias_prom,
                "tooltip_text": tooltip_text,
                "clientes_vista": clientes_vista,
                "clientes_detalle": clientes_en_etapa
            })
        
        total_prospectos = sum(d["cantidad"] for d in datos_etapas) if datos_etapas else 1
        
        for d in datos_etapas:
            d["porcentaje"] = (d["cantidad"] / total_prospectos * 100) if total_prospectos > 0 else 0
        
        for i, d in enumerate(datos_etapas):
            color = colores_embudo[d["etapa"]]
            anchos = [90, 80, 70, 60, 50, 40]
            ancho = anchos[i] if i < len(anchos) else 50
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown(f"""
                <div style="width: {ancho}%; margin: 0 auto 15px auto;">
                    <div title="{d['tooltip_text']}" style="background: white; border-left: 5px solid {color}; border-radius: 12px; padding: 15px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); cursor: help;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span style="font-size: 16px; font-weight: 700; color: {color};">{d['etapa']}</span>
                            <span style="font-size: 24px; font-weight: 800; color: {color};">{d['cantidad']}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 12px;">
                            <span style="font-size: 12px; color: #64748B;">{d['porcentaje']:.0f}% del total</span>
                            <span style="font-size: 12px; color: #94A3B8;">⏱️ {d['dias_prom']} días</span>
                        </div>
                        <div style="padding-top: 8px; border-top: 1px solid #E2E8F0; font-size: 12px; color: #475569;">
                            <strong>📋 Clientes:</strong><br>
                            <span>{d['clientes_vista']}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                if i < len(etapas) - 1:
                    st.markdown('<div style="text-align: center; font-size: 24px; color: #CBD5E1; margin: -5px 0 5px 0;">▼</div>', unsafe_allow_html=True)
        
        with st.expander("📋 Ver detalle completo de todos los prospectos por etapa"):
            for d in datos_etapas:
                if d["clientes_detalle"]:
                    st.markdown(f"### {d['etapa']} - {len(d['clientes_detalle'])} prospectos")
                    df_detalle = pd.DataFrame([
                        {
                            "Prospecto": c["nombre"],
                            "Fecha registro": c["fecha_reg"],
                            "Días en etapa": c["dias"],
                            "Total visitas": c["visitas"],
                            "Última visita": c["fecha_ultima"],
                            "Días sin visita": c["dias_sin_visita"],
                            "Visitado esta semana": c["visitado_semana"],
                            "Temperatura": c["temperatura"]
                        }
                        for c in d["clientes_detalle"]
                    ])
                    st.dataframe(df_detalle, use_container_width=True, hide_index=True)
                    st.markdown("---")
                else:
                    st.markdown(f"### {d['etapa']} - Sin prospectos")
                    st.markdown("---")
        
        st.markdown("### 🚦 Alertas por Etapa")
        alertas_embudo = []
        for i in range(len(etapas) - 1):
            cant_actual = datos_etapas[i]["cantidad"]
            cant_siguiente = datos_etapas[i + 1]["cantidad"]
            if cant_actual > 0 and cant_siguiente == 0:
                alertas_embudo.append(f"🔴 **{etapas[i]} → {etapas[i+1]}**: {cant_actual} prospectos no han avanzado")
            elif cant_actual > 0 and cant_siguiente < (cant_actual * 0.3):
                pct = (cant_siguiente / cant_actual * 100) if cant_actual > 0 else 0
                alertas_embudo.append(f"🟡 **{etapas[i]} → {etapas[i+1]}**: Solo {cant_siguiente} de {cant_actual} avanzaron ({pct:.0f}%)")
        
        for d in datos_etapas:
            if d["dias_prom"] > 15 and d["cantidad"] > 0:
                alertas_embudo.append(f"🟡 **{d['etapa']}**: Prospectos llevan promedio {d['dias_prom']} días (recomendado <15 días)")
        
        total_prospectos_general = sum(d["cantidad"] for d in datos_etapas)
        if datos_etapas[-1]["cantidad"] == 0 and total_prospectos_general > 0:
            alertas_embudo.append(f"🔴 **CIERRE**: {total_prospectos_general} prospectos en el embudo, 0 cierres")
        
        if alertas_embudo:
            for a in alertas_embudo[:8]:
                st.markdown(f"- {a}")
        else:
            st.success("✅ No hay alertas activas en el embudo de ventas")

    # ============================================================
    # BOTÓN DE DESCARGA DE REPORTE (dentro del sidebar)
    # ============================================================
    if sel_zona and sel_sem != "Todas" and 'total_v' in dir():
        st.markdown("---")
        st.markdown("### 📥 Descargar Reporte")
        
        # Calcular cierres y prospectos únicos
        cierres_zona = 0
        prospectos_unicos_zona = 0
        if "Task" in df_f.columns:
            cierres_zona = len(df_f[(df_f["tipo"] == "PROSPECCIÓN") & 
                                    (df_f["Task"].str.upper().str.contains("CIERRE", na=False))])
        if "tipo" in df_f.columns:
            df_prospeccion_zona = df_f[df_f["tipo"] == "PROSPECCIÓN"]
            prospectos_unicos_zona = df_prospeccion_zona["Cliente o Prospecto"].nunique() if not df_prospeccion_zona.empty else 0
        
        # Generar mensaje de alerta
        mensaje_reporte = f"Reporte de gestión comercial para la zona {sel_zona}"
        if 'alerta_html' in locals() and alerta_html:
            mensaje_reporte = alerta_html
        
        # Generar HTML (asegúrate de tener df_resumen_giro, data_alerts, etc.)
        reporte_html = generar_reporte_html(
            df_f, sel_zona, total_v, 
            visitas_prospeccion, visitas_mantenimiento, 
            meta_actual, data_alerts, df_resumen_giro, 
            mensaje_reporte, sel_sem,
            cierres_zona, prospectos_unicos_zona
        )
        
        st.download_button(
            label="📥 Descargar Reporte (HTML)",
            data=reporte_html,
            file_name=f"reporte_{sel_zona}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True
        )    
    

with tab2:
    mostrar_pagina_ranking(df)

st.markdown('<div style="text-align:center; color:#94A3B8; font-size:12px; margin-top:40px; padding:20px; border-top:1px solid #E2E8F0;">Go To Market SAC · Dashboard Comercial · ' + str(datetime.now().year) + '</div>', unsafe_allow_html=True)