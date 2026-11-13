# =============================================================================
# ANÁLISIS DE RENTABILIDAD DE PROYECTOS DE INVERSIÓN (Streamlit)
# Para ejecutar, escribe en la terminal (las dos líneas, sin el símbolo #):
# cd "c:\Users\vale1\OneDrive\Cursos\Analisis de datos con python IA_TecMTY\Proyecto Final EVAM"
# python -m streamlit run AnalisisRentabilidadProyectos.py
# =============================================================================

import ast  # Módulo para analizar fórmulas de forma segura (árbol de sintaxis)
import json  # Módulo para guardar / cargar la configuración en formato JSON
import keyword  # Módulo para validar que un nombre de variable no sea palabra reservada
from datetime import date, datetime  # Clases para manejar fechas (inicio del proyecto y fecha de guardado)
from pathlib import Path  # Manejo de rutas de archivos (biblioteca de proyectos guardados)
from functools import lru_cache  # Decorador para memorizar fórmulas ya compiladas
import html as html_lib  # Módulo para escapar texto del usuario dentro de HTML
import io  # Módulo para crear archivos en memoria (Excel)
import re  # Módulo de expresiones regulares (limpiar texto)
import urllib.parse  # Módulo para armar el enlace de correo (mailto)

import numpy as np  # Librería para cálculos numéricos con arreglos
import pandas as pd  # Librería para manejo de tablas (DataFrames)
import plotly.graph_objects as go  # Librería para gráficas interactivas
import streamlit as st  # Librería para construir la aplicación web

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA Y ESTILO VISUAL (minimalista)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Rentabilidad de Proyectos", page_icon="◆", layout="wide")  # Título, ícono y diseño ancho

NAVY = "#0B2545"  # Azul marino ejecutivo: color institucional principal
ACENTO = "#0E7C7B"  # Verde petróleo: valores positivos y utilidades
NEGATIVO = "#B42318"  # Rojo vino: valores negativos
AZUL = "#1D4E89"  # Azul acero: ingresos y flujos
MORADO = "#3E5C76"  # Azul pizarra: volumetría y recuperación
DORADO = "#C9A227"  # Dorado: acentos ejecutivos
AMBAR = "#B8892B"  # Dorado oscuro: egresos e índices
NARANJA = "#9C6B30"  # Bronce: gastos corporativos
INDIGO = "#13315C"  # Marino medio: financiamiento
TINTA = NAVY  # Color para líneas principales
SUAVE = "#94A3B8"  # Gris claro para series secundarias
COLORES = {"teal": ACENTO, "azul": AZUL, "morado": MORADO, "ambar": AMBAR, "naranja": NARANJA,  # Colores de tarjetas
           "indigo": INDIGO, "rosa": NEGATIVO, "dorado": DORADO, "navy": NAVY}  # (continuación)


def rgba(hex_color, alfa):  # Convierte un color hexadecimal a rgba con transparencia
    h = hex_color.lstrip("#")  # Quita el símbolo #
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{alfa})"  # Arma el texto rgba


css_tarjetas = "".join(  # Genera una clase CSS por cada color de tarjeta
    f".kpi.{n} {{ background:linear-gradient(160deg,{rgba(c, .10)},{rgba(c, .02)}); border-top:3px solid {c}; }}\n"  # Fondo tenue y borde superior
    f".kpi.{n} .l {{ color:{c}; opacity:1; }}\n"  # Etiqueta del color de la tarjeta
    for n, c in COLORES.items())  # Una regla por color

# Hoja de estilo del resumen ejecutivo (se usa en la app y en el archivo HTML que se envía por correo)
CSS_REPORTE = f"""
.rep {{ background:#FFFFFF; color:#1F2937; font-family:'Inter',Arial,sans-serif; max-width:1020px; margin:0 auto; padding:34px 40px; border-radius:14px; box-shadow:0 10px 30px rgba(11,37,69,.12); border:1px solid #E5E7EB; }}
.rep .head {{ background:linear-gradient(135deg,{NAVY},{INDIGO} 70%,{AZUL}); color:#FFFFFF; padding:24px 28px; border-radius:12px; border-bottom:4px solid {DORADO}; }}
.rep .eyebrow {{ color:{DORADO}; font-size:11px; letter-spacing:.18em; text-transform:uppercase; font-weight:700; }}
.rep .titulo {{ color:#FFFFFF; font-size:26px; margin:6px 0 4px; font-weight:800; letter-spacing:-.02em; }}
.rep .meta {{ color:rgba(255,255,255,.82); font-size:13px; }}
.rep .h2 {{ color:{NAVY}; font-size:13px; text-transform:uppercase; letter-spacing:.1em; border-bottom:2px solid {DORADO}; padding:0 0 5px; margin:26px 0 10px; font-weight:800; }}
.rep p {{ font-size:13.5px; line-height:1.6; margin:6px 0; color:#1F2937; }}
.rep table {{ width:100%; border-collapse:collapse; font-size:12.5px; margin:4px 0; }}
.rep th {{ background:{NAVY}; color:#FFFFFF; text-align:left; padding:7px 9px; font-weight:600; }}
.rep td {{ padding:6px 9px; border-bottom:1px solid #E5E7EB; color:#1F2937; vertical-align:top; }}
.rep tr:nth-child(even) td {{ background:#F8FAFC; }}
.rep .n {{ text-align:right; white-space:nowrap; }}
.rep td.neg {{ color:{NEGATIVO}; }}
.rep table.kv td:first-child {{ color:#64748B; width:40%; }}
.rep .cols {{ display:grid; grid-template-columns:1fr 1fr; gap:22px; }}
.rep .kpis {{ display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }}
.rep .k {{ border:1px solid #E5E7EB; border-top:3px solid {NAVY}; border-radius:10px; padding:10px 12px; background:#FFFFFF; }}
.rep .k .l {{ font-size:10px; text-transform:uppercase; letter-spacing:.08em; color:#64748B; font-weight:700; }}
.rep .k .v {{ font-size:19px; font-weight:800; color:{NAVY}; margin-top:3px; }}
.rep .k.pos {{ border-top-color:{ACENTO}; }}
.rep .k.pos .v {{ color:{ACENTO}; }}
.rep .k.neg {{ border-top-color:{NEGATIVO}; }}
.rep .k.neg .v {{ color:{NEGATIVO}; }}
.rep .veredicto {{ padding:14px 18px; border-radius:10px; font-size:13.5px; line-height:1.6; margin:12px 0 4px; color:#1F2937; }}
.rep .veredicto.ok {{ background:rgba(14,124,123,.08); border-left:5px solid {ACENTO}; }}
.rep .veredicto.no {{ background:rgba(180,35,24,.07); border-left:5px solid {NEGATIVO}; }}
.rep .pie {{ margin-top:28px; padding-top:10px; border-top:1px solid #E5E7EB; font-size:11px; color:#64748B; }}
@media (max-width:700px) {{ .rep .cols {{ grid-template-columns:1fr; }} .rep .kpis {{ grid-template-columns:1fr 1fr; }} .rep {{ padding:20px; }} }}
@media print {{ .rep {{ box-shadow:none; border:none; }} }}
"""

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
.stApp :is(h1,h2,h3,h4,p,label,li,.kpi,.hero,.sec,.cover) {{ font-family:'Inter',sans-serif; }}
#MainMenu, footer {{ visibility:hidden; }}
.block-container {{ padding-top:2rem; max-width:1440px; }}
.hero {{ background:linear-gradient(135deg,{NAVY} 0%,{INDIGO} 70%,{AZUL} 100%); color:#fff; border-radius:18px; padding:24px 32px; margin-bottom:1.2rem; border-bottom:4px solid {DORADO}; box-shadow:0 10px 28px {rgba(NAVY, .22)}; }}
.hero .eyebrow {{ color:{DORADO}; font-size:.7rem; letter-spacing:.18em; text-transform:uppercase; font-weight:700; }}
.hero .titulo {{ color:#fff; font-size:2.2rem; font-weight:800; letter-spacing:-0.03em; line-height:1.15; margin-top:.25rem; }}
.hero .sub {{ color:rgba(255,255,255,.8); font-size:.9rem; margin-top:.4rem; }}
.badge {{ display:inline-block; padding:6px 16px; border-radius:999px; font-size:.8rem; font-weight:700; letter-spacing:.02em; margin-top:.9rem; }}
.badge.ok {{ background:#fff; color:{ACENTO}; }}
.badge.no {{ background:#fff; color:{NEGATIVO}; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin:.4rem 0 1.2rem; }}
.kpi {{ border:1px solid rgba(128,128,128,.16); border-radius:14px; padding:16px 18px; }}
.kpi .l {{ font-size:.68rem; text-transform:uppercase; letter-spacing:.09em; font-weight:700; opacity:.65; }}
.kpi .v {{ font-size:1.6rem; font-weight:800; letter-spacing:-0.02em; margin-top:6px; }}
.kpi .s {{ font-size:.74rem; opacity:.6; margin-top:2px; }}
.kpi.pos .v {{ color:{ACENTO}; }}
.kpi.neg .v {{ color:{NEGATIVO}; }}
{css_tarjetas}
.sec {{ font-size:.74rem; text-transform:uppercase; letter-spacing:.14em; font-weight:800; margin:2.2rem 0 .7rem; padding-left:12px; border-left:4px solid {DORADO}; opacity:.85; }}
.cover {{ border-radius:20px; overflow:hidden; border:1px solid rgba(128,128,128,.18); box-shadow:0 14px 36px {rgba(NAVY, .14)}; margin:.6rem 0 1.4rem; }}
.cover .top {{ background:linear-gradient(135deg,{NAVY} 0%,{INDIGO} 65%,{AZUL} 100%); color:#fff; padding:40px 44px 34px; border-bottom:5px solid {DORADO}; position:relative; }}
.cover .eyebrow {{ color:{DORADO}; font-size:.74rem; letter-spacing:.2em; text-transform:uppercase; font-weight:700; }}
.cover .titulo {{ color:#fff; font-size:2.7rem; font-weight:800; letter-spacing:-0.035em; line-height:1.1; margin:.5rem 0 .7rem; max-width:80%; }}
.cover .tipo {{ display:inline-block; border:1px solid {rgba(DORADO, .7)}; color:{DORADO}; border-radius:999px; padding:3px 14px; font-size:.78rem; font-weight:600; }}
.cover .desc {{ color:rgba(255,255,255,.86); max-width:900px; margin-top:1.1rem; line-height:1.6; font-size:.98rem; }}
.cover .folio {{ position:absolute; top:30px; right:40px; text-align:right; color:rgba(255,255,255,.75); font-size:.8rem; line-height:1.5; }}
.cover .folio b {{ color:#fff; font-size:.95rem; }}
.cover .body {{ padding:30px 44px 34px; }}
.cover .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:20px 30px; }}
.cover .k {{ font-size:.67rem; text-transform:uppercase; letter-spacing:.11em; opacity:.55; font-weight:700; }}
.cover .v {{ font-size:1rem; font-weight:600; margin-top:3px; }}
.cover .obj {{ margin-top:26px; padding:16px 20px; border-left:4px solid {DORADO}; background:{rgba(DORADO, .08)}; border-radius:0 12px 12px 0; line-height:1.55; }}
.stTabs [data-baseweb="tab-list"] {{ gap:1.5rem; }}
.stTabs [data-baseweb="tab"] {{ font-weight:600; }}
.stTabs [aria-selected="true"] {{ font-weight:800; }}
.stTabs [data-baseweb="tab-highlight"] {{ background:{DORADO} !important; height:3px; }}
[data-testid="stSidebar"] {{ background:linear-gradient(180deg,{rgba(NAVY, .07)} 0%,{rgba(DORADO, .05)} 100%); }}
.stDownloadButton button, .stLinkButton a {{ border:1px solid {rgba(AZUL, .45)}; border-radius:10px; }}
.grupo {{ border-radius:14px; padding:14px 18px; margin:.3rem 0 .7rem; color:#fff; }}
.grupo.sin {{ background:linear-gradient(135deg,{ACENTO} 0%,#0A5E5D 100%); border-bottom:3px solid {rgba(DORADO, .9)}; }}
.grupo.con {{ background:linear-gradient(135deg,{NAVY} 0%,{AZUL} 100%); border-bottom:3px solid {DORADO}; }}
.grupo .t {{ font-weight:800; font-size:1rem; letter-spacing:-.01em; color:#fff; }}
.grupo .d {{ font-size:.8rem; color:rgba(255,255,255,.85); margin-top:3px; }}
.subg {{ font-size:.78rem; font-weight:800; text-transform:uppercase; letter-spacing:.1em; margin:1.2rem 0 .35rem; padding-left:10px; }}
.subg.sin {{ border-left:4px solid {ACENTO}; color:{ACENTO}; }}
.subg.con {{ border-left:4px solid {NAVY}; opacity:.85; }}
.pcard .banda {{ height:4px; border-radius:4px; background:linear-gradient(90deg,{NAVY},{AZUL} 60%,{DORADO}); margin-bottom:.8rem; }}
.pcard .chip {{ display:inline-block; font-size:.64rem; letter-spacing:.09em; text-transform:uppercase; font-weight:700; color:{DORADO}; border:1px solid {rgba(DORADO, .6)}; border-radius:999px; padding:2px 10px; }}
.pcard .nom {{ font-size:1.08rem; font-weight:800; margin:.55rem 0 .3rem; line-height:1.25; letter-spacing:-.01em; }}
.pcard .des {{ font-size:.82rem; opacity:.7; line-height:1.45; min-height:4.4em; }}
.pcard .met {{ display:flex; gap:18px; margin-top:.8rem; flex-wrap:wrap; }}
.pcard .met .k {{ font-size:.6rem; text-transform:uppercase; letter-spacing:.09em; opacity:.55; font-weight:700; }}
.pcard .met .v {{ font-size:1.05rem; font-weight:800; }}
.pcard .met .v.pos {{ color:{ACENTO}; }}
.pcard .met .v.neg {{ color:{NEGATIVO}; }}
.pcard .pie {{ font-size:.72rem; opacity:.55; margin:.6rem 0 .3rem; }}
{CSS_REPORTE}
</style>
""", unsafe_allow_html=True)  # Inyecta la hoja de estilos (encabezado, carátula, tarjetas, pestañas, resumen)

# -----------------------------------------------------------------------------
# CONSTANTES
# -----------------------------------------------------------------------------
MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]  # Abreviaturas de meses en español
MONEDAS = ["MXN", "USD", "—"]  # Monedas posibles de una variable ("—" = no es dinero, p. ej. volumen o %)
FUNCIONES = {"min": min, "max": max, "abs": abs, "round": round}  # Funciones que se permiten dentro de las fórmulas
RESERVADAS = {"mes", "anio", "mes_del_anio", "mes_proyecto", "ingresos", "egresos", "utilidad_bruta"} | set(FUNCIONES)  # Nombres que el usuario no puede usar
NODOS_PERMITIDOS = (  # Tipos de elementos que puede contener una fórmula (todo lo demás se rechaza por seguridad)
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name, ast.Load,  # Expresión, operaciones, números y nombres
    ast.Call, ast.IfExp, ast.Compare, ast.BoolOp,  # Llamadas a funciones, condicionales y comparaciones
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.Mod, ast.FloorDiv,  # Operadores aritméticos
    ast.USub, ast.UAdd, ast.Gt, ast.GtE, ast.Lt, ast.LtE, ast.Eq, ast.NotEq,  # Signos y comparadores
    ast.And, ast.Or,  # Operadores lógicos
)  # Fin de la tupla de nodos permitidos
COLS_ULTIMO = {"Saldo de deuda", "FE sin fin. acumulado", "FE con fin. acumulado", "Tipo de cambio"}  # Columnas que al agrupar por año toman el último valor

# Parámetros generales con sus valores por defecto (la llave es el nombre del widget)
PARAMS_DEF = {
    "p_nombre": "Proyecto de inversión",  # Nombre del proyecto
    "p_fecha": date(2027, 1, 1),  # Fecha de inicio del proyecto (inicio del desarrollo)
    "p_dev_meses": 12,  # Duración de la etapa de desarrollo (Año 0) en meses
    "p_anios": 10,  # Años de operación a evaluar
    "p_moneda": "MXN",  # Moneda en la que se reportan los resultados
    "p_tc": 18.50,  # Tipo de cambio inicial (MXN por USD)
    "p_tc_var": 2.0,  # Variación anual esperada del tipo de cambio en %
    "p_inf_mxn": 4.0,  # Inflación anual en México en %
    "p_inf_usd": 2.5,  # Inflación anual en EE.UU. en %
    "p_tasa_desc": 12.0,  # Tasa de descuento del proyecto (WACC) en % anual
    "p_ke": 15.0,  # Costo de capital del accionista (Ke) en % anual
    "p_isr": 30.0,  # Tasa de ISR en %
    "p_ptu": 10.0,  # Tasa de PTU en %
    "p_rescate": True,  # Incluir valor de rescate (valor en libros) al final del horizonte
    "p_fin_pct": 60.0,  # Porcentaje del CAPEX que se financia con deuda
    "p_fin_tasa": 12.0,  # Tasa de interés anual del crédito en %
    "p_fin_comision": 1.0,  # Comisión por financiamiento en % sobre el total del crédito (se paga en la primera disposición)
    "p_fin_plazo": 6,  # Plazo del crédito en años
    "p_fin_gracia": 12,  # Meses de gracia (solo se pagan intereses)
    "p_fin_mes": 1,  # Mes del proyecto en que se dispone el crédito
}  # Fin del diccionario de parámetros

# Datos informativos de la carátula (no afectan los cálculos)
CARATULA_DEF = {
    "p_folio": "EVAL-2026-001",  # Folio o número de estudio
    "p_tipo": "Nuevo negocio",  # Tipo de proyecto
    "p_fecha_estudio": date.today(),  # Fecha de elaboración del estudio
    "p_empresa": "Grupo Ejemplo S.A. de C.V.",  # Empresa que invierte
    "p_area": "Dirección de Finanzas",  # Área solicitante
    "p_solicitante": "Nombre del solicitante",  # Persona que pidió el estudio
    "p_correo_sol": "",  # Correo del solicitante (para enviar el resumen)
    "p_responsable": "Analista financiero",  # Responsable del análisis
    "p_ubicacion": "Monterrey, N.L.",  # Ubicación del proyecto
    "p_descripcion": "Construcción y operación de una estación de venta de combustible con capacidad aproximada "  # Descripción
                     "de 150 mil litros mensuales, con una etapa de desarrollo de 12 meses.",  # (continuación)
    "p_objetivo": "Determinar la viabilidad financiera del proyecto, su capacidad para generar valor al accionista "  # Objetivo
                  "y su sensibilidad ante cambios en precio, volumen, tipo de cambio e inversión.",  # (continuación)
}  # Fin del diccionario de carátula
PARAMS_DEF.update(CARATULA_DEF)  # Se agregan a los parámetros para guardarlos y cargarlos juntos
FECHAS = {"p_fecha", "p_fecha_estudio"}  # Parámetros que son fechas (se guardan como texto ISO)
TIPOS_PROYECTO = ["Nuevo negocio", "Expansión", "Reemplazo de activos", "Eficiencia / ahorro", "Cumplimiento regulatorio", "Otro"]  # Tipos de proyecto
MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",  # Nombres completos de los meses
                "septiembre", "octubre", "noviembre", "diciembre"]  # (continuación)
INDICADORES = {"VPN proyecto (sin fin.)": "VPN proyecto", "TIR proyecto (anual)": "TIR proyecto",  # Indicadores para la sensibilidad
               "VPN accionista (con fin.)": "VPN accionista", "TIR accionista (anual)": "TIR accionista",  # (continuación)
               "Periodo de recuperación (años)": "Recuperación proyecto",  # Payback del proyecto (sin financiamiento)
               "Recuperación accionista (años)": "Recuperación accionista"}  # Payback del accionista (con financiamiento)
RAZONES_SIN_FIN = {"VPN proyecto (sin fin.)", "TIR proyecto (anual)", "Periodo de recuperación (años)",  # Razones del proyecto
                   "Recuperación descontada (años)", "Índice de rentabilidad", "Relación beneficio/costo",  # (continuación)
                   "Margen bruto", "Margen EBITDA", "Margen neto (sin fin.)", "ROI anual (sin fin.)"}  # sin financiamiento
RAZONES_GENERALES = {"Inversión total (CAPEX)", "Monto del crédito", "Capital propio"}  # Datos generales de la inversión
REGLAS_SEMAFORO = {  # (nivel sano, nivel de alerta, ¿mayor es mejor?) para el semáforo de razones por periodo
    "DSCR (veces)": (1.2, 1.0, True),  # Cobertura del servicio de deuda
    "Cobertura de intereses (veces)": (3.0, 1.5, True),  # EBITDA / intereses
    "Deuda / EBITDA (veces)": (2.5, 4.0, False),  # Apalancamiento (menor es mejor)
}  # Los márgenes usan (5%, 0%, mayor es mejor)


def es_recuperacion(indicador):  # Indica si el indicador es un periodo de recuperación
    return "ecuperación" in indicador  # Coincide con 'Recuperación' y 'recuperación'


# -----------------------------------------------------------------------------
# DATOS DE EJEMPLO (el usuario puede modificarlos libremente)
# -----------------------------------------------------------------------------
def config_default():  # Función que regresa las tablas de ejemplo
    variables = pd.DataFrame([  # Tabla de variables declaradas por el usuario
        {"Variable": "litros_vendidos", "Descripción": "Volumen de venta mensual", "Tipo": "Volumen", "Unidad": "litros", "Moneda": "—", "Valor": 150000.0, "Crec. anual %": 3.0, "Inflación": False},  # Volumen vendido
        {"Variable": "precio_venta", "Descripción": "Precio de venta por litro", "Tipo": "Precio", "Unidad": "por litro", "Moneda": "MXN", "Valor": 25.6, "Crec. anual %": 0.5, "Inflación": True},  # Precio en pesos
        {"Variable": "costo_compra", "Descripción": "Costo de compra por litro", "Tipo": "Costo", "Unidad": "por litro", "Moneda": "USD", "Valor": 1.12, "Crec. anual %": 0.0, "Inflación": True},  # Costo en dólares
        {"Variable": "energia_kwh", "Descripción": "Consumo de energía mensual", "Tipo": "Volumen", "Unidad": "kWh", "Moneda": "—", "Valor": 12000.0, "Crec. anual %": 2.0, "Inflación": False},  # Insumo energía
        {"Variable": "tarifa_kwh", "Descripción": "Tarifa eléctrica", "Tipo": "Costo", "Unidad": "por kWh", "Moneda": "MXN", "Valor": 3.2, "Crec. anual %": 1.0, "Inflación": True},  # Tarifa energía
        {"Variable": "nomina", "Descripción": "Nómina mensual de operación", "Tipo": "Costo", "Unidad": "por mes", "Moneda": "MXN", "Valor": 180000.0, "Crec. anual %": 1.0, "Inflación": True},  # Nómina
        {"Variable": "mantenimiento", "Descripción": "Mantenimiento mensual", "Tipo": "Costo", "Unidad": "por mes", "Moneda": "MXN", "Valor": 25000.0, "Crec. anual %": 0.0, "Inflación": True},  # Mantenimiento
        {"Variable": "fee_corporativo", "Descripción": "Fee corporativo sobre ingresos", "Tipo": "Otro", "Unidad": "% decimal", "Moneda": "—", "Valor": 0.03, "Crec. anual %": 0.0, "Inflación": False},  # Porcentaje (no es dinero)
    ])  # Fin de la tabla de variables
    ingresos = pd.DataFrame([  # Tabla de conceptos de ingreso (solo existen en la etapa de operación)
        {"Concepto": "Venta de combustible", "Fórmula": "litros_vendidos * precio_venta", "Mes inicio": 1, "Mes fin": None},  # Ingreso principal
        {"Concepto": "Servicios adicionales", "Fórmula": "35000", "Mes inicio": 4, "Mes fin": None},  # Ingreso fijo desde el mes 4 de operación
    ])  # Fin de la tabla de ingresos
    egresos = pd.DataFrame([  # Tabla de conceptos de egreso (costos)
        {"Concepto": "Estudios y permisos", "Etapa": "Desarrollo", "Fórmula": "150000", "Mes inicio": 1, "Mes fin": 6},  # Gasto preoperativo
        {"Concepto": "Equipo de proyecto", "Etapa": "Desarrollo", "Fórmula": "nomina * 0.4", "Mes inicio": 1, "Mes fin": None},  # Nómina reducida en desarrollo
        {"Concepto": "Costo de producto", "Etapa": "Operación", "Fórmula": "litros_vendidos * costo_compra", "Mes inicio": 1, "Mes fin": None},  # Costo de venta
        {"Concepto": "Energía eléctrica", "Etapa": "Operación", "Fórmula": "energia_kwh * tarifa_kwh", "Mes inicio": 1, "Mes fin": None},  # Costo de energía
        {"Concepto": "Nómina", "Etapa": "Operación", "Fórmula": "nomina", "Mes inicio": 1, "Mes fin": None},  # Costo de personal
        {"Concepto": "Mantenimiento", "Etapa": "Operación", "Fórmula": "mantenimiento", "Mes inicio": 1, "Mes fin": None},  # Costo de mantenimiento
    ])  # Fin de la tabla de egresos
    corporativos = pd.DataFrame([  # Gastos corporativos (pueden usar 'ingresos', 'egresos', 'utilidad_bruta')
        {"Concepto": "Gestión del proyecto", "Etapa": "Desarrollo", "Fórmula": "30000", "Mes inicio": 1, "Mes fin": None},  # Gasto corporativo en desarrollo
        {"Concepto": "Fee corporativo", "Etapa": "Operación", "Fórmula": "fee_corporativo * ingresos", "Mes inicio": 1, "Mes fin": None},  # % de ingresos
        {"Concepto": "Servicios administrativos", "Etapa": "Operación", "Fórmula": "45000", "Mes inicio": 1, "Mes fin": None},  # Cuota fija mensual
    ])  # Fin de la tabla de gastos corporativos
    capex = pd.DataFrame([  # Tabla de inversiones (CAPEX); Año 0 = etapa de desarrollo
        {"Concepto": "Terreno", "Moneda": "MXN", "Monto": 2500000.0, "Año": 0, "Mes": 1, "Vida útil (años)": 0.0, "Inflación": False},  # Terreno (no se deprecia)
        {"Concepto": "Obra civil", "Moneda": "MXN", "Monto": 3500000.0, "Año": 0, "Mes": 4, "Vida útil (años)": 20.0, "Inflación": False},  # Construcción
        {"Concepto": "Equipo", "Moneda": "USD", "Monto": 190000.0, "Año": 0, "Mes": 9, "Vida útil (años)": 10.0, "Inflación": False},  # Equipo importado
        {"Concepto": "Reposición de equipo menor", "Moneda": "MXN", "Monto": 800000.0, "Año": 6, "Mes": 1, "Vida útil (años)": 5.0, "Inflación": True},  # Reinversión inflada
    ])  # Fin de la tabla de CAPEX
    escenarios = pd.DataFrame([  # Ajustes porcentuales sobre variables, CAPEX o tipo de cambio
        {"Escenario": "Pesimista", "Variable": "litros_vendidos", "Ajuste %": -10.0},  # Menor volumen
        {"Escenario": "Pesimista", "Variable": "precio_venta", "Ajuste %": -3.0},  # Menor precio
        {"Escenario": "Pesimista", "Variable": "Tipo de cambio", "Ajuste %": 8.0},  # Peso más débil (encarece insumos en USD)
        {"Escenario": "Pesimista", "Variable": "CAPEX", "Ajuste %": 10.0},  # Mayor inversión
        {"Escenario": "Optimista", "Variable": "litros_vendidos", "Ajuste %": 10.0},  # Mayor volumen
        {"Escenario": "Optimista", "Variable": "precio_venta", "Ajuste %": 2.0},  # Mayor precio
        {"Escenario": "Optimista", "Variable": "CAPEX", "Ajuste %": -5.0},  # Menor inversión
    ])  # Fin de la tabla de escenarios
    return {"variables": variables, "ingresos": ingresos, "egresos": egresos,  # Regresa las tablas en un diccionario
            "corporativos": corporativos, "capex": capex, "escenarios": escenarios}  # (continuación del diccionario)


# -----------------------------------------------------------------------------
# PLANTILLAS DE PROYECTOS PRELLENADAS (base para nuevos análisis)
# -----------------------------------------------------------------------------
def _v(nombre, desc, tipo, unidad, moneda, valor, crec=0.0, infl=False):  # Crea un renglón de variable
    return {"Variable": nombre, "Descripción": desc, "Tipo": tipo, "Unidad": unidad, "Moneda": moneda,  # Datos de la variable
            "Valor": float(valor), "Crec. anual %": float(crec), "Inflación": infl}  # Valor, crecimiento e inflación


def _i(concepto, formula, ini=1, fin=None):  # Crea un renglón de ingreso
    return {"Concepto": concepto, "Fórmula": formula, "Mes inicio": ini, "Mes fin": fin}  # Datos del ingreso


def _l(concepto, formula, etapa="Operación", ini=1, fin=None):  # Crea un renglón de egreso o gasto corporativo
    return {"Concepto": concepto, "Etapa": etapa, "Fórmula": formula, "Mes inicio": ini, "Mes fin": fin}  # Datos del concepto


def _k(concepto, moneda, monto, anio, mes, vida, infl=False):  # Crea un renglón de CAPEX
    return {"Concepto": concepto, "Moneda": moneda, "Monto": float(monto), "Año": anio, "Mes": mes,  # Datos de la inversión
            "Vida útil (años)": float(vida), "Inflación": infl}  # Vida útil e inflación


def _e(escenario, variable, ajuste):  # Crea un renglón de escenario
    return {"Escenario": escenario, "Variable": variable, "Ajuste %": float(ajuste)}  # Datos del ajuste


def _registros(df):  # Convierte un DataFrame a lista de diccionarios (sin NaN)
    return df.astype(object).where(df.notna(), None).to_dict("records")  # Registros limpios


def _plantilla(meta, parametros, variables, ingresos, egresos, corporativos, capex, escenarios):  # Arma el paquete de una plantilla
    return {"meta": meta, "parametros": parametros,  # Datos generales y parámetros
            "tablas": {"variables": variables, "ingresos": ingresos, "egresos": egresos,  # Tablas
                       "corporativos": corporativos, "capex": capex, "escenarios": escenarios}}  # (continuación)


_base_ej = config_default()  # Tablas del ejemplo original (estación de combustible)
PLANTILLAS = {  # Biblioteca de plantillas
    "estacion": _plantilla(  # Plantilla 1: estación de servicio
        {"nombre": "Estación de servicio de combustible", "tipo": "Nuevo negocio",  # Datos de la tarjeta
         "descripcion": "Venta de combustible con costo de producto en USD, 12 meses de desarrollo y 10 años de operación."},  # Descripción
        {"p_nombre": "Estación de servicio de combustible", "p_folio": "PLT-EST-01", "p_tipo": "Nuevo negocio"},  # Parámetros (resto por defecto)
        *[_registros(_base_ej[k]) for k in ("variables", "ingresos", "egresos", "corporativos", "capex", "escenarios")]),  # Tablas del ejemplo

    "solar": _plantilla(  # Plantilla 2: parque solar
        {"nombre": "Parque solar fotovoltaico 5 MW", "tipo": "Nuevo negocio",  # Datos de la tarjeta
         "descripcion": "Generación de energía con contrato de compraventa (PPA) en USD, degradación anual de paneles y 25 años de operación."},  # Descripción
        {"p_nombre": "Parque solar fotovoltaico 5 MW", "p_folio": "PLT-SOL-01", "p_tipo": "Nuevo negocio",  # Identificación
         "p_ubicacion": "Hermosillo, Son.", "p_dev_meses": 18, "p_anios": 25, "p_tasa_desc": 10.0, "p_ke": 13.0,  # Horizonte y tasas
         "p_fin_pct": 70.0, "p_fin_tasa": 10.0, "p_fin_plazo": 12, "p_fin_gracia": 18, "p_fin_mes": 6,  # Financiamiento
         "p_descripcion": "Construcción y operación de un parque solar de 5 MW con venta de energía bajo contrato de largo plazo "  # Descripción
                          "y venta de certificados de energía limpia.",  # (continuación)
         "p_objetivo": "Evaluar la rentabilidad del parque y su capacidad de pago de deuda de largo plazo."},  # Objetivo
        [_v("generacion_mwh", "Energía generada al mes", "Volumen", "MWh", "—", 850, -0.5),  # Generación (degradación 0.5% anual)
         _v("precio_mwh", "Precio PPA de la energía", "Precio", "por MWh", "USD", 68, 0, True),  # Precio de la energía
         _v("precio_cel", "Precio del certificado de energía limpia", "Precio", "por MWh", "USD", 6, 0, False),  # Precio del CEL
         _v("om_fijo", "Operación y mantenimiento mensual", "Costo", "por mes", "USD", 9000, 0, True),  # O&M
         _v("seguro", "Seguro de la planta", "Costo", "por mes", "MXN", 60000, 0, True),  # Seguro
         _v("renta_terreno", "Arrendamiento del terreno", "Costo", "por mes", "MXN", 45000, 0, True),  # Terreno
         _v("fee_gestion", "Fee de gestión sobre ingresos", "Otro", "% decimal", "—", 0.02)],  # Fee
        [_i("Venta de energía (PPA)", "generacion_mwh * precio_mwh"),  # Ingreso por energía
         _i("Certificados de energía limpia", "generacion_mwh * precio_cel")],  # Ingreso por CEL
        [_l("Estudios, permisos e interconexión", "250000", "Desarrollo", 1, 12),  # Gasto de desarrollo
         _l("Gestión del desarrollo", "120000", "Desarrollo"),  # Gasto de desarrollo
         _l("Operación y mantenimiento", "om_fijo"), _l("Seguros", "seguro"),  # Costos de operación
         _l("Arrendamiento de terreno", "renta_terreno")],  # Costo de terreno
        [_l("Fee de gestión", "fee_gestion * ingresos"), _l("Administración", "35000")],  # Gastos corporativos
        [_k("Paneles e inversores", "USD", 2200000, 0, 10, 25),  # Equipo principal
         _k("Estructura y obra civil", "MXN", 14000000, 0, 8, 25),  # Obra civil
         _k("Subestación e interconexión", "MXN", 9000000, 0, 14, 25),  # Interconexión
         _k("Reemplazo de inversores", "USD", 300000, 12, 1, 10)],  # Reinversión
        [_e("Pesimista", "generacion_mwh", -8), _e("Pesimista", "precio_mwh", -10), _e("Pesimista", "CAPEX", 10),  # Escenario pesimista
         _e("Optimista", "generacion_mwh", 5), _e("Optimista", "precio_mwh", 5)]),  # Escenario optimista

    "manufactura": _plantilla(  # Plantilla 3: línea de manufactura
        {"nombre": "Nueva línea de manufactura", "tipo": "Expansión",  # Datos de la tarjeta
         "descripcion": "Línea de producción con precio y materia prima en USD, rampa de arranque de 6 meses y merma de producción."},  # Descripción
        {"p_nombre": "Nueva línea de manufactura", "p_folio": "PLT-MAN-01", "p_tipo": "Expansión",  # Identificación
         "p_ubicacion": "Apodaca, N.L.", "p_dev_meses": 9, "p_anios": 10, "p_tasa_desc": 13.0, "p_ke": 16.0,  # Horizonte y tasas
         "p_fin_pct": 50.0, "p_fin_tasa": 11.5, "p_fin_plazo": 7, "p_fin_gracia": 9, "p_fin_mes": 2,  # Financiamiento
         "p_descripcion": "Instalación de una nueva línea de producción para atender la demanda de clientes de exportación.",  # Descripción
         "p_objetivo": "Determinar si la expansión genera valor y en cuánto tiempo se recupera la inversión."},  # Objetivo
        [_v("unidades_producidas", "Producción mensual a plena capacidad", "Volumen", "piezas", "—", 40000, 4),  # Volumen
         _v("precio_unitario", "Precio de venta por pieza", "Precio", "por pieza", "USD", 4.8, 0, True),  # Precio
         _v("materia_prima", "Materia prima por pieza", "Costo", "por pieza", "USD", 2.1, 0, True),  # Materia prima
         _v("mano_obra", "Nómina de la línea", "Costo", "por mes", "MXN", 650000, 1, True),  # Mano de obra
         _v("energia_kwh", "Consumo de energía", "Volumen", "kWh", "—", 90000, 3),  # Energía
         _v("tarifa_kwh", "Tarifa eléctrica", "Costo", "por kWh", "MXN", 2.9, 1, True),  # Tarifa
         _v("mantenimiento", "Mantenimiento mensual", "Costo", "por mes", "MXN", 80000, 0, True),  # Mantenimiento
         _v("merma", "Porcentaje de merma", "Otro", "% decimal", "—", 0.03)],  # Merma
        [_i("Venta de piezas", "unidades_producidas * (1 - merma) * precio_unitario * min(1, mes / 6)")],  # Ingreso con rampa de 6 meses
        [_l("Ingeniería y puesta en marcha", "300000", "Desarrollo"),  # Gasto de desarrollo
         _l("Capacitación de personal", "mano_obra * 0.3", "Desarrollo", 7, 9),  # Capacitación previa al arranque
         _l("Materia prima", "unidades_producidas * materia_prima * min(1, mes / 6)"),  # Costo de materia prima
         _l("Mano de obra", "mano_obra"), _l("Energía eléctrica", "energia_kwh * tarifa_kwh"),  # Personal y energía
         _l("Mantenimiento", "mantenimiento")],  # Mantenimiento
        [_l("Regalías corporativas", "0.02 * ingresos"), _l("Servicios compartidos", "90000")],  # Gastos corporativos
        [_k("Maquinaria", "USD", 1800000, 0, 3, 10), _k("Adecuación de nave", "MXN", 9000000, 0, 2, 20),  # Equipo y nave
         _k("Herramentales", "USD", 250000, 0, 8, 5), _k("Reposición de herramentales", "USD", 250000, 6, 1, 5, True)],  # Herramentales
        [_e("Pesimista", "unidades_producidas", -15), _e("Pesimista", "materia_prima", 8), _e("Pesimista", "Tipo de cambio", -8),  # Pesimista
         _e("Optimista", "unidades_producidas", 10), _e("Optimista", "precio_unitario", 3)]),  # Optimista

    "cedis": _plantilla(  # Plantilla 4: centro de distribución
        {"nombre": "Centro de distribución en renta", "tipo": "Nuevo negocio",  # Datos de la tarjeta
         "descripcion": "Naves industriales en renta (USD por m²) con ocupación gradual, terreno no depreciable y 15 años de operación."},  # Descripción
        {"p_nombre": "Centro de distribución en renta", "p_folio": "PLT-CED-01", "p_tipo": "Nuevo negocio",  # Identificación
         "p_ubicacion": "Escobedo, N.L.", "p_dev_meses": 15, "p_anios": 15, "p_tasa_desc": 11.0, "p_ke": 14.0,  # Horizonte y tasas
         "p_fin_pct": 65.0, "p_fin_tasa": 10.5, "p_fin_plazo": 10, "p_fin_gracia": 15, "p_fin_mes": 3,  # Financiamiento
         "p_descripcion": "Desarrollo de un parque logístico de 18,000 m² rentables para operadores de comercio electrónico.",  # Descripción
         "p_objetivo": "Evaluar el retorno del desarrollo inmobiliario considerando la curva de ocupación."},  # Objetivo
        [_v("m2_rentables", "Superficie rentable", "Volumen", "m²", "—", 18000),  # Superficie
         _v("ocupacion", "Ocupación estabilizada", "Otro", "% decimal", "—", 0.9),  # Ocupación
         _v("renta_m2", "Renta mensual por m²", "Precio", "por m²/mes", "USD", 8.2, 0, True),  # Renta
         _v("mantenimiento_m2", "Mantenimiento por m²", "Costo", "por m²/mes", "MXN", 9, 0, True),  # Mantenimiento
         _v("predial", "Predial y derechos", "Costo", "por mes", "MXN", 85000, 0, True),  # Predial
         _v("seguridad", "Seguridad y vigilancia", "Costo", "por mes", "MXN", 140000, 0, True),  # Seguridad
         _v("comision_admin", "Administración inmobiliaria", "Otro", "% decimal", "—", 0.03)],  # Comisión
        [_i("Renta de naves", "m2_rentables * min(ocupacion, 0.5 + 0.15 * anio) * renta_m2"),  # Renta con ocupación gradual
         _i("Servicios a inquilinos", "m2_rentables * 0.4")],  # Cuotas de servicios
        [_l("Permisos y licencias", "180000", "Desarrollo", 1, 8), _l("Comercialización", "90000", "Desarrollo", 7),  # Desarrollo
         _l("Mantenimiento", "m2_rentables * mantenimiento_m2"), _l("Predial", "predial"), _l("Seguridad", "seguridad")],  # Operación
        [_l("Administración inmobiliaria", "comision_admin * ingresos"), _l("Corporativo", "60000")],  # Gastos corporativos
        [_k("Terreno", "MXN", 45000000, 0, 1, 0), _k("Construcción de naves", "MXN", 95000000, 0, 4, 30),  # Terreno y construcción
         _k("Equipamiento (andenes y racks)", "USD", 900000, 0, 12, 10)],  # Equipamiento
        [_e("Pesimista", "renta_m2", -10), _e("Pesimista", "ocupacion", -15), _e("Pesimista", "CAPEX", 8),  # Pesimista
         _e("Optimista", "renta_m2", 6), _e("Optimista", "ocupacion", 5)]),  # Optimista
}  # Fin de la biblioteca de plantillas

CARPETA_PROYECTOS = Path(__file__).parent / "proyectos_guardados"  # Carpeta donde se guardan los proyectos del usuario


# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES
# -----------------------------------------------------------------------------
def num(valor, defecto=0.0):  # Convierte un valor a número, usando un valor por defecto si está vacío
    try:  # Intenta la conversión
        v = float(valor)  # Convierte a flotante
        return defecto if np.isnan(v) else v  # Si es NaN regresa el valor por defecto
    except (TypeError, ValueError):  # Si no se puede convertir (None, texto, etc.)
        return defecto  # Regresa el valor por defecto


def si(valor):  # Interpreta una casilla de verificación (True solo si está marcada)
    return isinstance(valor, (bool, np.bool_)) and bool(valor)  # Evita que NaN o None cuenten como verdadero


def bonito(nombre):  # Convierte 'litros_vendidos' en 'Litros vendidos' para mostrarlo
    return str(nombre).replace("_", " ").capitalize()  # Reemplaza guiones bajos y pone mayúscula inicial


def finito(x):  # Indica si un valor es un número válido
    return x is not None and np.isfinite(x)  # Falso para None, NaN o infinito


def dinero(x):  # Da formato compacto de moneda (ej. $3.3 M)
    if not finito(x):  # Si no hay dato
        return "N/D"  # Regresa 'no disponible'
    signo = "-" if x < 0 else ""  # Signo negativo antes del símbolo de pesos
    x = abs(x)  # Trabaja con el valor absoluto
    if x >= 1e6:  # Si son millones
        return f"{signo}${x / 1e6:,.1f} M"  # Formato en millones
    if x >= 1e3:  # Si son miles
        return f"{signo}${x / 1e3:,.0f} k"  # Formato en miles
    return f"{signo}${x:,.0f}"  # Formato normal


def pct(x):  # Da formato de porcentaje a un número (0.12 -> 12.0%)
    return f"{x:.1%}" if finito(x) else "N/D"  # Regresa 'N/D' si no hay dato


def dinero_full(x):  # Da formato de moneda completo (ej. -$1,234,567)
    return f"{'-' if x < 0 else ''}${abs(x):,.0f}" if finito(x) else "N/D"  # Signo antes del símbolo de pesos


def fecha_larga(d):  # Convierte una fecha a texto (ej. 10 de noviembre de 2026)
    return f"{d.day} de {MESES_LARGOS[d.month - 1]} de {d.year}" if isinstance(d, date) else "—"  # Texto en español


def fmt_razon(nombre, x):  # Da formato a una razón financiera según su tipo
    if nombre in RAZONES_GENERALES or nombre in ("Intereses totales", "Comisión por financiamiento", "Costo financiero total") or nombre.startswith("VPN"):  # Montos
        return dinero_full(x)  # Formato moneda
    if "años" in nombre:  # Periodos de recuperación
        return f"{x:.1f} años" if finito(x) else "No recupera"  # Formato en años
    if nombre.startswith(("Índice", "Relación", "DSCR", "Múltiplo", "Cobertura")):  # Razones en veces
        return f"{x:.2f}x" if finito(x) else "N/D"  # Formato en veces
    return pct(x)  # TIR, márgenes y ROI en porcentaje


def pintar(texto_html, destino=None):  # Dibuja HTML en la app (escapa '$' para que no se lea como fórmula matemática)
    (destino or st).markdown(texto_html.replace("$", "&#36;"), unsafe_allow_html=True)  # Inserta el HTML


def factor_moneda(moneda, transcurrido, p):  # Factor para convertir de la moneda de la variable a la moneda de reporte
    t = np.asarray(transcurrido, dtype=float)  # Años transcurridos desde el inicio
    if moneda not in ("MXN", "USD") or moneda == p["p_moneda"]:  # Si no es dinero o ya está en la moneda de reporte
        return np.ones_like(t)  # No hay conversión
    tc = p["p_tc"] * (1 + p["p_tc_var"] / 100) ** t  # Tipo de cambio proyectado de cada periodo
    return tc if moneda == "USD" else 1 / tc  # USD→MXN multiplica; MXN→USD divide


def factor_inflacion(moneda, transcurrido, p, aplica):  # Factor de inflación acumulada según la moneda
    t = np.asarray(transcurrido, dtype=float)  # Años transcurridos desde el inicio
    if not aplica:  # Si la variable no se ve afectada por la inflación
        return np.ones_like(t)  # Factor 1
    tasa = p["p_inf_usd"] if moneda == "USD" else p["p_inf_mxn"]  # Inflación del país de la moneda
    return (1 + tasa / 100) ** t  # Inflación acumulada


@lru_cache(maxsize=2048)  # Memoriza el resultado para no recompilar la misma fórmula
def compilar_formula(texto, nombres):  # Valida y compila una fórmula; 'nombres' son las variables permitidas
    texto = str(texto).strip()  # Quita espacios al inicio y al final
    if not texto:  # Si la fórmula está vacía
        return None, "fórmula vacía"  # Regresa error
    try:  # Intenta interpretar la fórmula
        arbol = ast.parse(texto, mode="eval")  # Convierte el texto en árbol de sintaxis
    except SyntaxError:  # Si la fórmula está mal escrita
        return None, f"error de sintaxis en «{texto}»"  # Regresa error
    for nodo in ast.walk(arbol):  # Recorre cada elemento de la fórmula
        if not isinstance(nodo, NODOS_PERMITIDOS):  # Si el elemento no está permitido
            return None, f"elemento no permitido ({type(nodo).__name__}) en «{texto}»"  # Regresa error
        if isinstance(nodo, ast.Call) and not (isinstance(nodo.func, ast.Name) and nodo.func.id in FUNCIONES):  # Función no permitida
            return None, f"solo se permiten las funciones {', '.join(FUNCIONES)}"  # Regresa error
        if isinstance(nodo, ast.Name) and nodo.id not in nombres and nodo.id not in FUNCIONES:  # Variable no declarada
            return None, f"variable desconocida «{nodo.id}»"  # Regresa error
        if isinstance(nodo, ast.Constant) and not isinstance(nodo.value, (int, float)):  # Texto u otro tipo de constante
            return None, f"solo se permiten números en «{texto}»"  # Regresa error
    return compile(arbol, "<formula>", "eval"), None  # Regresa la fórmula compilada y sin error


def tir(flujos):  # Calcula la Tasa Interna de Retorno por periodo (búsqueda por bisección)
    f = np.asarray(flujos, dtype=float)  # Convierte los flujos a arreglo numérico
    if not ((f > 0).any() and (f < 0).any()):  # Sin flujos positivos y negativos no existe TIR
        return np.nan  # Regresa 'no disponible'
    t = np.arange(len(f))  # Periodos 0, 1, 2, ...

    def vpn(r):  # Valor presente neto a una tasa r
        with np.errstate(all="ignore"):  # Ignora advertencias por números muy grandes
            return float(np.sum(f / (1 + r) ** t))  # Suma de flujos descontados

    r_ant, v_ant = None, None  # Tasa y VPN del punto anterior de la búsqueda
    for r in np.linspace(-0.9, 1.0, 381):  # Recorre tasas de -90% a 100% por periodo
        v = vpn(r)  # Calcula el VPN a esa tasa
        if not np.isfinite(v):  # Si el resultado no es válido
            continue  # Pasa a la siguiente tasa
        if v_ant is not None and np.sign(v) != np.sign(v_ant):  # Si el VPN cambió de signo hay una raíz
            a, b, va = r_ant, r, v_ant  # Intervalo donde está la raíz
            for _ in range(100):  # Bisección con 100 iteraciones
                c = (a + b) / 2  # Punto medio
                vc = vpn(c)  # VPN en el punto medio
                if np.sign(vc) == np.sign(va):  # Mismo signo que el extremo izquierdo
                    a, va = c, vc  # Mueve el extremo izquierdo
                else:  # Signo distinto
                    b = c  # Mueve el extremo derecho
            return (a + b) / 2  # Regresa la TIR encontrada
        r_ant, v_ant = r, v  # Guarda el punto actual como anterior
    return np.nan  # Si no encontró cambio de signo, no hay TIR


def periodo_recuperacion(flujos):  # Calcula en cuántos meses se recupera la inversión
    acumulado = np.cumsum(flujos)  # Flujo acumulado
    if len(acumulado) == 0 or acumulado[-1] < 0:  # Si al final sigue negativo
        return np.nan  # No se recupera dentro del horizonte
    negativos = np.where(acumulado < 0)[0]  # Posiciones donde el acumulado es negativo
    if len(negativos) == 0:  # Si nunca fue negativo
        return 0.0  # Se recupera de inmediato
    i = negativos[-1]  # Último mes con acumulado negativo
    return i + 1 + (-acumulado[i] / flujos[i + 1])  # Meses + fracción interpolada del mes siguiente


# -----------------------------------------------------------------------------
# MOTOR DE CÁLCULO DEL MODELO FINANCIERO (MENSUAL)
# -----------------------------------------------------------------------------
def calcular_impuestos(uai, anio, tasa_isr, tasa_ptu):  # Calcula PTU e ISR anuales y los reparte por mes
    ptu = np.zeros(len(uai))  # Arreglo mensual de PTU
    isr = np.zeros(len(uai))  # Arreglo mensual de ISR
    perdidas = 0.0  # Pérdidas fiscales acumuladas por amortizar
    for a in np.unique(anio):  # Recorre cada año del proyecto (incluye el Año 0)
        idx = np.where(anio == a)[0]  # Meses que pertenecen al año
        uai_anual = uai[idx].sum()  # Utilidad antes de impuestos del año
        ptu_anual = max(uai_anual, 0) * tasa_ptu  # PTU sobre la utilidad positiva
        base = uai_anual - ptu_anual  # Base para ISR (la PTU se deduce)
        if base <= 0:  # Si hay pérdida
            perdidas += -base  # Se acumula para años futuros
            isr_anual = 0.0  # No se paga ISR
        else:  # Si hay utilidad
            uso = min(perdidas, base)  # Pérdidas que se pueden amortizar
            perdidas -= uso  # Se descuentan de las pendientes
            isr_anual = (base - uso) * tasa_isr  # ISR sobre la base restante
        positivos = np.clip(uai[idx], 0, None)  # Utilidad positiva de cada mes del año
        pesos = positivos / positivos.sum() if positivos.sum() > 0 else np.full(len(idx), 1 / len(idx))  # Proporción para repartir
        ptu[idx] = ptu_anual * pesos  # Reparte la PTU anual por mes
        isr[idx] = isr_anual * pesos  # Reparte el ISR anual por mes
    return ptu, isr  # Regresa ambos arreglos


def evaluar_lineas(tabla, nombres, contextos, seccion, operando, errores, etapa_fija=None):  # Evalúa conceptos con fórmula mes por mes
    n = len(contextos)  # Número total de meses
    resultado = {}  # Diccionario concepto -> serie mensual
    for _, fila in tabla.iterrows():  # Recorre cada concepto de la tabla
        concepto = str(fila.get("Concepto") or "").strip()  # Nombre del concepto
        formula = fila.get("Fórmula")  # Texto de la fórmula
        if not concepto or concepto == "nan" or formula is None or str(formula).strip() in ("", "nan", "None"):  # Fila incompleta
            continue  # Se ignora
        codigo, error = compilar_formula(str(formula), nombres)  # Valida y compila la fórmula
        if error:  # Si la fórmula tiene error
            errores.append(f"**{seccion} → {concepto}**: {error}")  # Registra el error para mostrarlo
            continue  # Pasa al siguiente concepto
        etapa = etapa_fija or str(fila.get("Etapa") or "Operación")  # Etapa del concepto (Desarrollo u Operación)
        en_desarrollo = etapa == "Desarrollo"  # Indica si el concepto es de la etapa de desarrollo
        inicio = int(num(fila.get("Mes inicio"), 1))  # Mes inicial dentro de su etapa (por defecto el 1)
        fin = int(num(fila.get("Mes fin"), n))  # Mes final dentro de su etapa (por defecto el último)
        serie = np.zeros(n)  # Serie mensual en ceros
        for i in range(n):  # Recorre cada mes
            if operando[i] == en_desarrollo:  # Si el mes no corresponde a la etapa del concepto
                continue  # Se salta
            mes_etapa = contextos[i]["mes_proyecto"] if en_desarrollo else contextos[i]["mes"]  # Mes contado dentro de la etapa
            if not (inicio <= mes_etapa <= fin):  # Si está fuera de la vigencia del concepto
                continue  # Se salta
            try:  # Intenta evaluar la fórmula
                serie[i] = float(eval(codigo, {"__builtins__": {}, **FUNCIONES}, contextos[i]))  # Evalúa con las variables del mes
            except ZeroDivisionError:  # Si hay división entre cero
                serie[i] = 0.0  # Se toma como cero
            except Exception as e:  # Cualquier otro error
                errores.append(f"**{seccion} → {concepto}**: {e}")  # Registra el error
                break  # Deja de evaluar este concepto
        resultado[concepto] = resultado.get(concepto, 0) + serie  # Suma (por si hay conceptos repetidos)
    return resultado  # Regresa las series por concepto


def suma(series, n):  # Suma un diccionario de series; si está vacío regresa ceros
    return np.sum(list(series.values()), axis=0) if series else np.zeros(n)  # Suma elemento a elemento


def correr_modelo(p, variables, ingresos, egresos, corporativos, capex):  # Corre el modelo completo y regresa resultados mensuales
    dev = int(p["p_dev_meses"])  # Meses de la etapa de desarrollo (Año 0)
    n = dev + int(p["p_anios"]) * 12  # Total de meses = desarrollo + operación
    meses = np.arange(1, n + 1)  # Mes del proyecto 1..n (cuenta desde el inicio del desarrollo)
    operando = meses > dev  # Verdadero en los meses de operación
    mes_op = np.where(operando, meses - dev, 0)  # Mes de operación (0 durante el desarrollo)
    anio = np.where(operando, (mes_op - 1) // 12 + 1, 0)  # Año del proyecto (0 = desarrollo)
    transcurrido = (meses - 1) // 12  # Años completos transcurridos desde el inicio (para inflación y tipo de cambio)
    mes_del_anio = (meses - 1) % 12 + 1  # Número de mes dentro de cada bloque de 12 meses
    errores = []  # Lista de errores encontrados

    # --- 1) Variables: valor base × crecimiento real × inflación × tipo de cambio ---
    series_var = {}  # Diccionario nombre -> serie mensual de la variable
    volumetria = []  # Lista de (nombre, unidad) de variables tipo Volumen
    for _, fila in variables.iterrows():  # Recorre cada variable declarada
        nombre = str(fila.get("Variable") or "").strip()  # Nombre de la variable
        if not nombre or nombre == "nan":  # Si está vacía
            continue  # Se ignora
        if not nombre.isidentifier() or keyword.iskeyword(nombre) or nombre in RESERVADAS:  # Si no es un nombre válido
            errores.append(f"**Variable «{nombre}»**: nombre no válido (usa letras, números y _, sin espacios, y que no sea reservado)")  # Error
            continue  # Se ignora
        moneda = str(fila.get("Moneda") or p["p_moneda"])  # Moneda de la variable (si falta, la de reporte)
        valor = num(fila.get("Valor"))  # Valor base (año 1, a precios del inicio)
        crecimiento = num(fila.get("Crec. anual %")) / 100  # Crecimiento real anual en decimal
        serie = valor * (1 + crecimiento) ** np.maximum(anio - 1, 0)  # Crecimiento real por año de operación
        serie = serie * factor_inflacion(moneda, transcurrido, p, si(fila.get("Inflación")))  # Aplica inflación si se marcó
        serie = serie * factor_moneda(moneda, transcurrido, p)  # Convierte a la moneda de reporte
        if str(fila.get("Tipo")) == "Volumen":  # Si es una variable de volumen / insumo
            serie = np.where(operando, serie, 0.0)  # No hay volumen durante el desarrollo
            volumetria.append((nombre, str(fila.get("Unidad") or "")))  # Se agrega a la volumetría
        series_var[nombre] = serie  # Guarda la serie mensual

    nombres_base = frozenset(series_var) | {"mes", "anio", "mes_del_anio", "mes_proyecto"}  # Nombres disponibles en fórmulas
    contextos = [{**{k: v[i] for k, v in series_var.items()},  # Valores de las variables en el mes i
                  "mes": int(mes_op[i]), "anio": int(anio[i]),  # Mes y año de operación
                  "mes_del_anio": int(mes_del_anio[i]), "mes_proyecto": int(meses[i])}  # Mes del año y mes del proyecto
                 for i in range(n)]  # Un contexto por mes

    # --- 2) Ingresos (solo en operación) y egresos (desarrollo u operación) ---
    ing = evaluar_lineas(ingresos, nombres_base, contextos, "Ingresos", operando, errores, etapa_fija="Operación")  # Series de ingresos
    egr = evaluar_lineas(egresos, nombres_base, contextos, "Egresos", operando, errores)  # Series de egresos
    tot_ing = suma(ing, n)  # Total de ingresos por mes
    tot_egr = suma(egr, n)  # Total de egresos por mes
    utilidad_bruta = tot_ing - tot_egr  # Utilidad bruta por mes

    # --- 3) Gastos corporativos (pueden depender de ingresos, egresos y utilidad bruta) ---
    nombres_corp = nombres_base | {"ingresos", "egresos", "utilidad_bruta"}  # Nombres extra disponibles
    for i in range(n):  # Agrega los totales al contexto de cada mes
        contextos[i].update(ingresos=tot_ing[i], egresos=tot_egr[i], utilidad_bruta=utilidad_bruta[i])  # Totales del mes
    corp = evaluar_lineas(corporativos, nombres_corp, contextos, "Gastos corporativos", operando, errores)  # Series corporativas
    tot_corp = suma(corp, n)  # Total de gastos corporativos por mes
    ebitda = utilidad_bruta - tot_corp  # EBITDA por mes

    # --- 4) CAPEX y depreciación (línea recta; inicia al arrancar la operación o el mes siguiente a la compra) ---
    capex_m = np.zeros(n)  # CAPEX por mes
    depreciacion = np.zeros(n)  # Depreciación por mes
    valor_libros = 0.0  # Valor en libros de los activos al final del horizonte
    for _, fila in capex.iterrows():  # Recorre cada inversión
        a_inv = int(num(fila.get("Año"), 0))  # Año de la inversión (0 = desarrollo)
        m_inv = int(num(fila.get("Mes"), 1))  # Mes dentro de ese año
        mes = m_inv if a_inv == 0 else dev + (a_inv - 1) * 12 + m_inv  # Mes del proyecto en que ocurre
        mes = max(mes, 1)  # Nunca antes del mes 1
        monto = num(fila.get("Monto"))  # Monto en su moneda original
        if monto <= 0 or mes > n:  # Si no es válida o cae fuera del horizonte
            continue  # Se ignora
        moneda = str(fila.get("Moneda") or p["p_moneda"])  # Moneda de la inversión
        t = (mes - 1) // 12  # Años transcurridos al momento de invertir
        monto = monto * float(factor_inflacion(moneda, t, p, si(fila.get("Inflación")))) * float(factor_moneda(moneda, t, p))  # Monto en moneda de reporte
        capex_m[mes - 1] += monto  # Registra la inversión en su mes
        vida = num(fila.get("Vida útil (años)"))  # Vida útil en años
        if vida > 0:  # Si el activo se deprecia
            meses_vida = max(int(round(vida * 12)), 1)  # Vida útil en meses
            cuota = monto / meses_vida  # Depreciación mensual
            inicio = max(mes, dev)  # Índice donde empieza a depreciarse (no antes de operar)
            fin = min(n, inicio + meses_vida)  # Índice donde termina dentro del horizonte
            depreciacion[inicio:fin] += cuota  # Aplica la depreciación mensual
            valor_libros += monto - cuota * max(fin - inicio, 0)  # Valor no depreciado al final
        else:  # Si no se deprecia (ej. terreno)
            valor_libros += monto  # Conserva todo su valor

    # --- 5) Financiamiento (pagos fijos tipo francés con periodo de gracia) ---
    disposicion = np.zeros(n)  # Dinero recibido del crédito
    intereses = np.zeros(n)  # Intereses pagados
    amortizacion = np.zeros(n)  # Pagos a capital
    saldo = np.zeros(n)  # Saldo insoluto de la deuda
    comision = np.zeros(n)  # Comisión por financiamiento
    monto_credito = p["p_fin_pct"] / 100 * capex_m.sum()  # Monto del crédito = % del CAPEX total
    d = int(p["p_fin_mes"])  # Mes de disposición
    if monto_credito > 0 and 1 <= d <= n:  # Solo si hay crédito y el mes es válido
        r = p["p_fin_tasa"] / 100 / 12  # Tasa de interés mensual
        total = max(int(p["p_fin_plazo"] * 12), 1)  # Plazo total en meses
        gracia = min(int(p["p_fin_gracia"]), total)  # Meses de gracia (no mayores al plazo)
        k = total - gracia  # Meses en que se amortiza capital
        pago = (monto_credito * r / (1 - (1 + r) ** -k) if r > 0 else monto_credito / k) if k > 0 else 0.0  # Pago fijo mensual
        s = 0.0  # Saldo vigente
        for i in range(d - 1, n):  # Recorre los meses desde la disposición
            j = i - (d - 1)  # Meses transcurridos desde la disposición
            if j == 0:  # Mes de la disposición
                disposicion[i] = monto_credito  # Entra el dinero del crédito
                comision[i] = monto_credito * num(p.get("p_fin_comision")) / 100  # Comisión sobre el total del crédito, pagada en la primera disposición
                s = monto_credito  # El saldo es el monto total
                saldo[i] = s  # Registra el saldo
                continue  # Pasa al siguiente mes
            if s <= 1e-6:  # Si la deuda ya se pagó
                break  # Termina el calendario
            interes = s * r  # Interés del mes sobre saldo
            capital = 0.0 if j <= gracia else min(pago - interes, s)  # En gracia no se paga capital
            if j >= total or i == n - 1:  # Último pago del plazo o fin del horizonte
                capital = s  # Se liquida el saldo restante
            intereses[i] = interes  # Registra el interés
            amortizacion[i] = capital  # Registra el pago a capital
            s -= capital  # Disminuye el saldo
            saldo[i] = s  # Registra el saldo

    # --- 6) Utilidades e impuestos ---
    uai = ebitda - depreciacion - intereses - comision  # Utilidad antes de impuestos (la comisión es gasto financiero deducible)
    ptu, isr = calcular_impuestos(uai, anio, p["p_isr"] / 100, p["p_ptu"] / 100)  # PTU e ISR con financiamiento
    utilidad_neta = uai - ptu - isr  # Utilidad neta
    ptu_sf, isr_sf = calcular_impuestos(ebitda - depreciacion, anio, p["p_isr"] / 100, p["p_ptu"] / 100)  # Impuestos sin financiamiento

    # --- 7) Flujos de efectivo ---
    rescate = np.zeros(n)  # Valor de rescate por mes
    if p["p_rescate"] and n > 0:  # Si se incluye valor de rescate
        rescate[-1] = valor_libros  # Se recupera el valor en libros en el último mes
    fe_sin = ebitda - ptu_sf - isr_sf - capex_m + rescate  # Flujo del proyecto (sin deuda)
    fe_con = utilidad_neta + depreciacion - capex_m + disposicion - amortizacion + rescate  # Flujo del accionista (con deuda)

    m = pd.DataFrame({"mes": meses, "anio": anio, "mes_del_anio": mes_del_anio})  # Tabla mensual base
    for nombre, _ in volumetria:  # Agrega la volumetría
        m[f"vol::{nombre}"] = series_var[nombre]  # Serie de la variable de volumen
    for c, s_ in ing.items():  # Agrega cada concepto de ingreso
        m[f"ing::{c}"] = s_  # Serie del concepto
    for c, s_ in egr.items():  # Agrega cada concepto de egreso
        m[f"egr::{c}"] = s_  # Serie del concepto
    for c, s_ in corp.items():  # Agrega cada gasto corporativo
        m[f"corp::{c}"] = s_  # Serie del concepto
    columnas = {  # Columnas de totales y resultados
        "Tipo de cambio": p["p_tc"] * (1 + p["p_tc_var"] / 100) ** transcurrido,  # Tipo de cambio proyectado (referencia)
        "Ingresos": tot_ing, "Egresos": tot_egr, "Utilidad bruta": utilidad_bruta,  # Ingresos, egresos y utilidad bruta
        "Gastos corporativos": tot_corp, "EBITDA": ebitda, "CAPEX": capex_m, "Depreciación": depreciacion,  # Gastos, EBITDA, inversión
        "Disposición de crédito": disposicion, "Comisión por financiamiento": comision,  # Crédito recibido y su comisión
        "Intereses": intereses, "Amortización de capital": amortizacion,  # Financiamiento
        "Saldo de deuda": saldo, "Utilidad antes de impuestos": uai, "PTU": ptu, "ISR": isr,  # Impuestos
        "Utilidad neta": utilidad_neta, "Valor de rescate": rescate,  # Utilidad neta y rescate
        "FE sin financiamiento": fe_sin, "FE con financiamiento": fe_con,  # Flujos de efectivo
        "FE sin fin. acumulado": np.cumsum(fe_sin), "FE con fin. acumulado": np.cumsum(fe_con),  # Flujos acumulados
        "PTU sin fin": ptu_sf, "ISR sin fin": isr_sf,  # Impuestos del escenario sin deuda (para razones)
    }  # Fin del diccionario de columnas
    m = pd.concat([m, pd.DataFrame(columnas)], axis=1)  # Une todo en una sola tabla mensual
    return {"m": m, "vol": volumetria, "ing": list(ing), "egr": list(egr), "corp": list(corp),  # Regresa resultados
            "errores": errores, "monto_credito": monto_credito, "valor_libros": valor_libros}  # (continuación)


# -----------------------------------------------------------------------------
# RAZONES FINANCIERAS
# -----------------------------------------------------------------------------
def anualizar(m):  # Agrupa la tabla mensual por año del proyecto
    reglas = {c: ("last" if c in COLS_ULTIMO else "sum") for c in m.columns if c not in ("mes", "anio", "mes_del_anio")}  # Suma o último valor
    return m.groupby("anio").agg(reglas)  # Tabla anual (Año 0, 1, 2, ...)


def razones_globales(res, p):  # Calcula las razones financieras de todo el proyecto
    m = res["m"]  # Tabla mensual
    t = np.arange(len(m))  # Periodos mensuales 0..n-1
    r_wacc = (1 + p["p_tasa_desc"] / 100) ** (1 / 12) - 1  # Tasa de descuento mensual equivalente (proyecto)
    r_ke = (1 + p["p_ke"] / 100) ** (1 / 12) - 1  # Tasa de descuento mensual equivalente (accionista)
    vp = lambda serie, r: float(np.sum(np.asarray(serie) / (1 + r) ** t))  # Función de valor presente
    fs = m["FE sin financiamiento"].values  # Flujo del proyecto
    fc = m["FE con financiamiento"].values  # Flujo del accionista
    tir_s, tir_c = tir(fs), tir(fc)  # TIR mensuales
    vp_capex = vp(m["CAPEX"], r_wacc)  # Valor presente de las inversiones
    vpn_s = vp(fs, r_wacc)  # VPN del proyecto
    costos = m["Egresos"] + m["Gastos corporativos"] + m["CAPEX"] + m["PTU sin fin"] + m["ISR sin fin"]  # Salidas de efectivo
    vp_costos = vp(costos, r_wacc)  # Valor presente de las salidas
    anual = anualizar(m)  # Tabla anual para la cobertura de deuda
    servicio = anual["Intereses"] + anual["Amortización de capital"]  # Servicio de la deuda por año
    disponible = anual["EBITDA"] - anual["PTU"] - anual["ISR"]  # Flujo disponible para pagar deuda
    dscr = (disponible / servicio)[(servicio > 0) & (anual.index > 0)]  # Cobertura en años de operación con pagos
    ing_tot = m["Ingresos"].sum()  # Ingresos totales del horizonte
    capex_tot = m["CAPEX"].sum()  # Inversión total
    capital_propio = max(capex_tot - res["monto_credito"], 0.0)  # Inversión pagada con recursos del accionista
    un_sf = (m["EBITDA"] - m["Depreciación"] - m["PTU sin fin"] - m["ISR sin fin"]).sum()  # Utilidad neta como si no hubiera deuda
    un_cf = m["Utilidad neta"].sum()  # Utilidad neta real (con intereses)
    anios_op = max(int(p["p_anios"]), 1)  # Años de operación (para promediar ROI y ROE)
    descontados = fs / (1 + r_wacc) ** t  # Flujos del proyecto descontados al WACC
    descontados_c = fc / (1 + r_ke) ** t  # Flujos del accionista descontados al Ke
    cobertura = (anual["EBITDA"] / anual["Intereses"])[(anual["Intereses"] > 0) & (anual.index > 0)]  # EBITDA / intereses por año
    entradas, salidas = fc[fc > 0].sum(), -fc[fc < 0].sum()  # Lo que recibe y lo que aporta el accionista
    return {  # Diccionario de razones (ordenado: general, sin financiamiento, con financiamiento)
        # --- Datos generales de la inversión ---
        "Inversión total (CAPEX)": capex_tot,  # Monto total invertido
        "Monto del crédito": res["monto_credito"],  # Monto financiado
        "Capital propio": capital_propio,  # Monto aportado por el accionista
        # --- Evaluación SIN financiamiento (como si toda la inversión fuera capital propio) ---
        "VPN proyecto (sin fin.)": vpn_s,  # Valor presente neto del proyecto al WACC
        "TIR proyecto (anual)": (1 + tir_s) ** 12 - 1 if np.isfinite(tir_s) else np.nan,  # TIR anualizada del proyecto
        "Periodo de recuperación (años)": periodo_recuperacion(fs) / 12,  # Payback simple del proyecto
        "Recuperación descontada (años)": periodo_recuperacion(descontados) / 12,  # Payback descontado del proyecto
        "Índice de rentabilidad": (vpn_s + vp_capex) / vp_capex if vp_capex > 0 else np.nan,  # VP de beneficios / VP de inversión
        "Relación beneficio/costo": vp(m["Ingresos"], r_wacc) / vp_costos if vp_costos > 0 else np.nan,  # B/C
        "Margen bruto": m["Utilidad bruta"].sum() / ing_tot if ing_tot else np.nan,  # Utilidad bruta / ingresos
        "Margen EBITDA": m["EBITDA"].sum() / ing_tot if ing_tot else np.nan,  # EBITDA / ingresos
        "Margen neto (sin fin.)": un_sf / ing_tot if ing_tot else np.nan,  # Utilidad neta sin intereses / ingresos
        "ROI anual (sin fin.)": un_sf / anios_op / capex_tot if capex_tot else np.nan,  # Utilidad neta anual promedio sin intereses / inversión
        # --- Evaluación CON financiamiento (punto de vista del accionista) ---
        "VPN accionista (con fin.)": vp(fc, r_ke),  # Valor presente neto del accionista al Ke
        "TIR accionista (anual)": (1 + tir_c) ** 12 - 1 if np.isfinite(tir_c) else np.nan,  # TIR anualizada del accionista
        "Recuperación accionista (años)": periodo_recuperacion(fc) / 12,  # Payback del capital propio
        "Recuperación accionista descontada (años)": periodo_recuperacion(descontados_c) / 12,  # Payback descontado al Ke
        "Margen neto (con fin.)": un_cf / ing_tot if ing_tot else np.nan,  # Utilidad neta con intereses / ingresos
        "ROE anual": un_cf / anios_op / capital_propio if capital_propio > 0 else np.nan,  # Utilidad neta anual promedio / capital propio
        "Múltiplo sobre capital": entradas / salidas if salidas > 0 else np.nan,  # Veces que el accionista recupera lo aportado
        "Intereses totales": m["Intereses"].sum(),  # Intereses pagados en todo el horizonte
        "Comisión por financiamiento": m["Comisión por financiamiento"].sum(),  # Comisión pagada al disponer el crédito
        "Costo financiero total": m["Intereses"].sum() + m["Comisión por financiamiento"].sum(),  # Intereses + comisión
        "DSCR mínimo": dscr.min() if len(dscr) else np.nan,  # Peor año de cobertura de deuda
        "DSCR promedio": dscr.mean() if len(dscr) else np.nan,  # Cobertura promedio
        "Cobertura de intereses mínima": cobertura.min() if len(cobertura) else np.nan,  # Peor año de EBITDA / intereses
    }  # Fin del diccionario


def razones_periodo(datos, columnas, meses=12):  # Razones por periodo: regresa (sin financiamiento, con financiamiento)
    ing = datos["Ingresos"].replace(0, np.nan)  # Ingresos (ceros a NaN para no dividir entre cero)
    servicio = (datos["Intereses"] + datos["Amortización de capital"]).replace(0, np.nan)  # Servicio de deuda
    intereses = datos["Intereses"].replace(0, np.nan)  # Intereses (ceros a NaN)
    ebitda_anual = (datos["EBITDA"] * 12 / meses).where(datos["EBITDA"] > 0)  # EBITDA anualizado (solo si es positivo)
    un_sf = datos["EBITDA"] - datos["Depreciación"] - datos["PTU sin fin"] - datos["ISR sin fin"]  # Utilidad neta sin deuda
    sin = pd.DataFrame({  # Razones del proyecto sin financiamiento
        "Margen bruto": datos["Utilidad bruta"] / ing,  # Utilidad bruta / ingresos
        "Margen EBITDA": datos["EBITDA"] / ing,  # EBITDA / ingresos
        "Margen neto (sin fin.)": un_sf / ing,  # Utilidad neta sin intereses / ingresos
        "Flujo libre / ingresos": datos["FE sin financiamiento"] / ing,  # Conversión de ingresos en flujo del proyecto
    })  # Fin de la tabla
    con = pd.DataFrame({  # Razones con financiamiento
        "Margen neto (con fin.)": datos["Utilidad neta"] / ing,  # Utilidad neta con intereses / ingresos
        "DSCR (veces)": (datos["EBITDA"] - datos["PTU"] - datos["ISR"]) / servicio,  # Cobertura del servicio de deuda
        "Cobertura de intereses (veces)": datos["EBITDA"] / intereses,  # EBITDA / intereses
        "Deuda / EBITDA (veces)": (datos["Saldo de deuda"] / ebitda_anual).where(datos["Saldo de deuda"] > 0),  # Apalancamiento
    })  # Fin de la tabla
    sin.index = columnas  # Etiquetas de los periodos
    con.index = columnas  # Etiquetas de los periodos
    return sin.T, con.T  # Transpone: filas = razones, columnas = periodos


# -----------------------------------------------------------------------------
# PRESENTACIÓN: TABLAS, TARJETAS Y GRÁFICAS
# -----------------------------------------------------------------------------
def seccion(texto):  # Escribe un título de sección discreto en mayúsculas
    st.markdown(f'<div class="sec">{texto}</div>', unsafe_allow_html=True)  # Usa la clase CSS 'sec'


def filas_estado(res):  # Define el orden y tipo de cada renglón del reporte
    f = [("VOLUMETRÍA (insumos)", None, "seccion")]  # Encabezado de volumetría
    f += [(f"   {bonito(n)} ({u})", f"vol::{n}", "linea") for n, u in res["vol"]]  # Una fila por variable de volumen
    f += [("INGRESOS", None, "seccion")]  # Encabezado de ingresos
    f += [(f"   {c}", f"ing::{c}", "linea") for c in res["ing"]]  # Conceptos de ingreso
    f += [("Total ingresos", "Ingresos", "total")]  # Total de ingresos
    f += [("EGRESOS", None, "seccion")]  # Encabezado de egresos
    f += [(f"   {c} ", f"egr::{c}", "linea") for c in res["egr"]]  # Conceptos de egreso
    f += [("Total egresos", "Egresos", "total")]  # Total de egresos
    f += [("UTILIDAD BRUTA", "Utilidad bruta", "total")]  # Utilidad bruta
    f += [("GASTOS CORPORATIVOS", None, "seccion")]  # Encabezado de gastos corporativos
    f += [(f"   {c}  ", f"corp::{c}", "linea") for c in res["corp"]]  # Conceptos corporativos
    f += [("Total gastos corporativos", "Gastos corporativos", "total")]  # Total corporativo
    f += [("EBITDA", "EBITDA", "total")]  # EBITDA
    f += [("CAPEX", None, "seccion")]  # Encabezado de inversión
    f += [("   Inversión (CAPEX)", "CAPEX", "linea"), ("   Depreciación", "Depreciación", "linea")]  # CAPEX y depreciación
    f += [("FINANCIAMIENTO", None, "seccion")]  # Encabezado de financiamiento
    f += [("   Disposición de crédito", "Disposición de crédito", "linea"),  # Dinero recibido
          ("   Comisión por financiamiento", "Comisión por financiamiento", "linea"),  # Comisión pagada en la primera disposición
          ("   Intereses", "Intereses", "linea"),  # Intereses pagados
          ("   Amortización de capital", "Amortización de capital", "linea"),  # Pagos a capital
          ("   Saldo de deuda (cierre)", "Saldo de deuda", "linea")]  # Saldo al cierre del periodo
    f += [("Utilidad antes de impuestos", "Utilidad antes de impuestos", "total")]  # UAI
    f += [("IMPUESTOS", None, "seccion")]  # Encabezado de impuestos
    f += [("   PTU", "PTU", "linea"), ("   ISR", "ISR", "linea")]  # PTU e ISR
    f += [("UTILIDAD NETA", "Utilidad neta", "total")]  # Utilidad neta
    f += [("FLUJO DE EFECTIVO", None, "seccion")]  # Encabezado de flujos
    f += [("   Valor de rescate", "Valor de rescate", "linea")]  # Valor de rescate
    f += [("Flujo de efectivo CON financiamiento", "FE con financiamiento", "total"),  # Flujo del accionista
          ("   Acumulado con financiamiento", "FE con fin. acumulado", "linea"),  # Acumulado del accionista
          ("Flujo de efectivo SIN financiamiento", "FE sin financiamiento", "total"),  # Flujo del proyecto
          ("   Acumulado sin financiamiento", "FE sin fin. acumulado", "linea")]  # Acumulado del proyecto
    return f  # Regresa la lista de filas


def construir_tabla(res, datos, etiquetas):  # Arma la tabla del reporte: filas = conceptos, columnas = periodos
    filas = []  # Lista de renglones
    for etiqueta, col, tipo in filas_estado(res):  # Recorre la definición de filas
        valores = datos[col].values if col else [np.nan] * len(datos)  # Valores del periodo (vacío en encabezados)
        if col is None:  # Si es encabezado
            total = np.nan  # No lleva total
        elif col in COLS_ULTIMO:  # Si es saldo o acumulado
            total = datos[col].iloc[-1]  # El total es el último valor
        else:  # Si es un concepto que se suma
            total = datos[col].sum()  # El total es la suma
        filas.append({"Concepto": etiqueta, "Tipo": tipo, **dict(zip(etiquetas, valores)), "Total": total})  # Agrega el renglón
    return pd.DataFrame(filas)  # Regresa la tabla


COLOR_SECCION = {  # Color de cada sección del reporte
    "VOLUMETRÍA (insumos)": MORADO, "INGRESOS": AZUL, "EGRESOS": AMBAR, "GASTOS CORPORATIVOS": NARANJA,  # Operación
    "CAPEX": "#64748B", "FINANCIAMIENTO": INDIGO, "IMPUESTOS": NEGATIVO, "FLUJO DE EFECTIVO": ACENTO,  # Inversión, deuda, impuestos y flujo
}  # Fin del diccionario
COLOR_TOTAL = {  # Color de los totales (los resultados clave llevan más color)
    "Total ingresos": AZUL, "Total egresos": AMBAR, "Total gastos corporativos": NARANJA,  # Totales de sección
    "UTILIDAD BRUTA": ACENTO, "EBITDA": ACENTO, "Utilidad antes de impuestos": INDIGO, "UTILIDAD NETA": ACENTO,  # Utilidades
    "Flujo de efectivo CON financiamiento": AZUL, "Flujo de efectivo SIN financiamiento": MORADO,  # Flujos
}  # Fin del diccionario
TOTALES_FUERTES = {"UTILIDAD BRUTA", "EBITDA", "UTILIDAD NETA"}  # Totales que se resaltan con fondo sólido


def mostrar_tabla(tabla, columnas):  # Muestra la tabla con colores por sección, negativos en rojo y miles
    def estilo_fila(fila):  # Estilo según el tipo de renglón
        concepto = fila["Concepto"]  # Texto del renglón
        if fila["Tipo"] == "seccion":  # Encabezados de sección
            c = COLOR_SECCION.get(concepto, ACENTO)  # Color de la sección
            return [f"font-weight:800; color:#FFFFFF; background-color:{c}"] * len(fila)  # Banda de color con texto blanco
        if fila["Tipo"] == "total":  # Totales
            c = COLOR_TOTAL.get(concepto, SUAVE)  # Color del total
            if concepto in TOTALES_FUERTES:  # Resultados clave
                return [f"font-weight:800; color:{c}; background-color:{rgba(c, .18)}"] * len(fila)  # Fondo más intenso
            return [f"font-weight:700; color:{c}; background-color:{rgba(c, .08)}"] * len(fila)  # Fondo tenue del color
        return [""] * len(fila)  # Líneas normales sin estilo
    cols_num = columnas + ["Total"]  # Columnas numéricas
    estilo = (tabla.style.apply(estilo_fila, axis=1)  # Aplica estilo por fila
              .format("{:,.0f}", subset=cols_num, na_rep="")  # Formato con separador de miles
              .map(lambda v: f"color:{NEGATIVO}" if isinstance(v, (int, float)) and v < 0 else "", subset=cols_num))  # Negativos en rojo
    st.dataframe(estilo, column_order=["Concepto"] + cols_num, hide_index=True,  # Muestra solo concepto y periodos
                 height=min(35 * (len(tabla) + 1) + 5, 1200), width="stretch",  # Altura según número de filas
                 column_config={"Concepto": st.column_config.TextColumn(pinned=True, width="medium")})  # Fija la columna de concepto


def mostrar_razones_periodo(tabla):  # Muestra la tabla de razones por periodo con formato
    def semaforo(fila):  # Colorea cada celda: verde = sano, ámbar = alerta, rojo = problema
        sano, alerta, mayor_mejor = REGLAS_SEMAFORO.get(fila.name, (0.05, 0.0, True))  # Regla de la razón
        estilos = []  # Lista de estilos de la fila
        for v in fila:  # Recorre cada valor
            if not finito(v):  # Sin dato
                estilos.append("")  # Sin color
                continue  # Siguiente celda
            if mayor_mejor:  # Razones donde más es mejor
                c = ACENTO if v >= sano else (AMBAR if v >= alerta else NEGATIVO)  # Verde, ámbar o rojo
            else:  # Razones donde menos es mejor (apalancamiento)
                c = ACENTO if v <= sano else (AMBAR if v <= alerta else NEGATIVO)  # Verde, ámbar o rojo
            estilos.append(f"color:{c}; background-color:{rgba(c, .12)}; font-weight:700")  # Estilo de la celda
        return estilos  # Regresa los estilos
    filas_veces = [i for i in tabla.index if "(veces)" in i]  # Razones expresadas en veces
    estilo = tabla.style.apply(semaforo, axis=1).format("{:.1%}", na_rep="–")  # Semáforo y formato porcentaje
    if filas_veces:  # Si hay razones en veces
        estilo = estilo.format("{:.2f}x", subset=pd.IndexSlice[filas_veces, :], na_rep="–")  # Formato en veces
    st.dataframe(estilo, width="stretch")  # Muestra la tabla


def tarjeta(etiqueta, valor, sub="", tono="", color="azul"):  # Regresa el HTML de una tarjeta KPI con su color
    return f'<div class="kpi {color} {tono}"><div class="l">{etiqueta}</div><div class="v">{valor}</div><div class="s">{sub}</div></div>'  # Estructura de la tarjeta


def mostrar_kpis(raz, p):  # Muestra las razones globales en dos bloques: sin financiamiento y con financiamiento
    tono = lambda ok: "pos" if ok else "neg"  # Verde si cumple, rojo si no
    anios_txt = lambda x: f"{x:.1f} años" if finito(x) else "No recupera"  # Formato de periodos de recuperación
    veces = lambda x: f"{x:.2f}x" if finito(x) else "N/D"  # Formato en veces
    wacc, ke = p["p_tasa_desc"], p["p_ke"]  # Tasas de descuento
    vpn_s, tir_s = raz["VPN proyecto (sin fin.)"], raz["TIR proyecto (anual)"]  # Indicadores del proyecto
    vpn_c, tir_c = raz["VPN accionista (con fin.)"], raz["TIR accionista (anual)"]  # Indicadores del accionista
    ir, bc = raz["Índice de rentabilidad"], raz["Relación beneficio/costo"]  # Índices del proyecto
    roe, dscr, cob = raz["ROE anual"], raz["DSCR mínimo"], raz["Cobertura de intereses mínima"]  # Indicadores de deuda
    sin = [  # Tarjetas SIN financiamiento (el proyecto por sí mismo)
        tarjeta("VPN proyecto", dinero(vpn_s), f"descontado al WACC {wacc:.1f}%", tono(finito(vpn_s) and vpn_s > 0), "teal"),  # VPN
        tarjeta("TIR proyecto", pct(tir_s), f"vs WACC {wacc:.1f}%", tono(finito(tir_s) and tir_s * 100 > wacc), "teal"),  # TIR
        tarjeta("Recuperación", anios_txt(raz["Periodo de recuperación (años)"]),  # Payback del proyecto
                f"descontada: {anios_txt(raz['Recuperación descontada (años)'])}", color="morado"),  # Payback descontado
        tarjeta("Índice de rentabilidad", veces(ir), "VP beneficios / VP inversión", tono(ir > 1) if finito(ir) else "", "ambar"),  # IR
        tarjeta("Beneficio / costo", veces(bc), "VP ingresos / VP salidas", tono(bc > 1) if finito(bc) else "", "ambar"),  # B/C
        tarjeta("Margen EBITDA", pct(raz["Margen EBITDA"]), f"margen bruto {pct(raz['Margen bruto'])}", color="naranja"),  # Márgenes operativos
        tarjeta("Margen neto", pct(raz["Margen neto (sin fin.)"]), "sin gastos financieros", color="naranja"),  # Margen neto sin intereses
        tarjeta("ROI anual", pct(raz["ROI anual (sin fin.)"]), f"promedio sobre inversión de {dinero(raz['Inversión total (CAPEX)'])}", color="indigo"),  # ROI
    ]  # Fin de las tarjetas sin financiamiento
    con = [  # Tarjetas CON financiamiento (punto de vista del accionista)
        tarjeta("VPN accionista", dinero(vpn_c), f"descontado al Ke {ke:.1f}%", tono(finito(vpn_c) and vpn_c > 0), "azul"),  # VPN
        tarjeta("TIR accionista", pct(tir_c), f"vs Ke {ke:.1f}%", tono(finito(tir_c) and tir_c * 100 > ke), "azul"),  # TIR
        tarjeta("Recuperación del capital", anios_txt(raz["Recuperación accionista (años)"]),  # Payback del accionista
                f"descontada: {anios_txt(raz['Recuperación accionista descontada (años)'])}", color="morado"),  # Payback descontado
        tarjeta("Capital propio", dinero(raz["Capital propio"]),  # Aportación del accionista
                f"crédito {dinero(raz['Monto del crédito'])} ({p['p_fin_pct']:.0f}% del CAPEX)", color="navy"),  # Monto del crédito
        tarjeta("ROE anual", pct(roe), f"promedio · múltiplo sobre capital {veces(raz['Múltiplo sobre capital'])}", tono(roe > 0) if finito(roe) else "", "indigo"),  # ROE
        tarjeta("Margen neto", pct(raz["Margen neto (con fin.)"]), f"costo financiero {dinero(raz['Costo financiero total'])} (comisión {dinero(raz['Comisión por financiamiento'])})", color="naranja"),  # Margen neto
        tarjeta("DSCR mínimo", veces(dscr), f"promedio {veces(raz['DSCR promedio'])}", tono(dscr >= 1.2) if finito(dscr) else "", "dorado"),  # DSCR
        tarjeta("Cobertura de intereses", veces(cob), "EBITDA / intereses (año mínimo)", tono(cob >= 3) if finito(cob) else "", "dorado"),  # Cobertura
    ]  # Fin de las tarjetas con financiamiento
    nota_con = (f"{p['p_fin_pct']:.0f}% del CAPEX con deuda al {p['p_fin_tasa']:.1f}% + comisión {num(p.get('p_fin_comision')):.2f}% · "  # Descripción del bloque
                f"flujo del accionista al Ke {ke:.1f}%"  # (continuación)
                if p["p_fin_pct"] > 0 else "Sin crédito capturado: los resultados coinciden con el proyecto")  # Caso sin deuda
    c1, c2 = st.columns(2, gap="large")  # Dos bloques lado a lado
    pintar(f'<div class="grupo sin"><div class="t">Sin financiamiento · evaluación del proyecto</div>'  # Encabezado del bloque
           f'<div class="d">Como si toda la inversión fuera capital propio · flujo del proyecto al WACC {wacc:.1f}%</div></div>'  # Descripción
           f'<div class="kpi-grid">{"".join(sin)}</div>', c1)  # Tarjetas
    pintar(f'<div class="grupo con"><div class="t">Con financiamiento · evaluación del accionista</div>'  # Encabezado del bloque
           f'<div class="d">{nota_con}</div></div>'  # Descripción
           f'<div class="kpi-grid">{"".join(con)}</div>', c2)  # Tarjetas


def bloque_razones(raz, p, datos, etiquetas, meses, periodo):  # Sección completa de razones financieras (tarjetas y tablas por periodo)
    seccion("Razones financieras")  # Título
    mostrar_kpis(raz, p)  # Tarjetas globales en dos bloques
    t_sin, t_con = razones_periodo(datos, etiquetas, meses)  # Razones por periodo
    pintar(f'<div class="subg sin">Por {periodo} · sin financiamiento</div>')  # Subtítulo
    mostrar_razones_periodo(t_sin)  # Tabla sin financiamiento
    pintar(f'<div class="subg con">Por {periodo} · con financiamiento</div>')  # Subtítulo
    mostrar_razones_periodo(t_con)  # Tabla con financiamiento


def estilo_fig(fig, titulo):  # Aplica un estilo limpio a una gráfica
    fig.update_layout(title=dict(text=titulo, font=dict(size=14)), height=360, barmode="group", bargap=0.25,  # Título y barras
                      legend=dict(orientation="h", y=-0.18, x=0), margin=dict(t=48, b=10, l=10, r=10),  # Leyenda y márgenes
                      font=dict(family="Inter, sans-serif"))  # Tipografía
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.15)", zerolinecolor="rgba(128,128,128,0.35)", tickformat="~s")  # Ejes discretos
    fig.update_xaxes(showgrid=False)  # Sin cuadrícula vertical
    return fig  # Regresa la gráfica


def graficas(datos, etiquetas, clave):  # Dibuja las gráficas de resultados y flujos
    c1, c2 = st.columns(2)  # Dos gráficas lado a lado
    fig = go.Figure()  # Gráfica 1: resultados
    fig.add_bar(x=etiquetas, y=datos["Ingresos"], name="Ingresos", marker_color=AZUL)  # Barras de ingresos
    fig.add_bar(x=etiquetas, y=datos["Egresos"] + datos["Gastos corporativos"], name="Egresos + corporativos", marker_color=AMBAR)  # Egresos
    fig.add_bar(x=etiquetas, y=datos["Utilidad neta"], name="Utilidad neta",  # Utilidad neta
                marker_color=[ACENTO if v >= 0 else NEGATIVO for v in datos["Utilidad neta"]])  # Verde o rojo según el signo
    c1.plotly_chart(estilo_fig(fig, "Resultados"), width="stretch", key=f"g1_{clave}")  # Muestra la gráfica 1
    fig = go.Figure()  # Gráfica 2: flujos de efectivo
    fig.add_bar(x=etiquetas, y=datos["FE con financiamiento"], name="Flujo con financiamiento",  # Flujo del accionista
                marker_color=[ACENTO if v >= 0 else NEGATIVO for v in datos["FE con financiamiento"]])  # Verde o rojo según el signo
    fig.add_scatter(x=etiquetas, y=datos["FE con fin. acumulado"], name="Acumulado con fin.", mode="lines+markers",  # Acumulado accionista
                    line=dict(color=AZUL, width=3), marker=dict(size=6))  # Línea azul
    fig.add_scatter(x=etiquetas, y=datos["FE sin fin. acumulado"], name="Acumulado sin fin.", mode="lines+markers",  # Acumulado proyecto
                    line=dict(color=MORADO, width=2.5, dash="dot"), marker=dict(size=5))  # Línea morada punteada
    c2.plotly_chart(estilo_fig(fig, "Flujo de efectivo"), width="stretch", key=f"g2_{clave}")  # Muestra la gráfica 2


def descargar(tabla, columnas, nombre, clave):  # Botón para descargar la tabla a CSV (se abre en Excel)
    csv = tabla[["Concepto"] + columnas + ["Total"]].to_csv(index=False).encode("utf-8-sig")  # CSV con acentos compatibles
    st.download_button("Descargar CSV", csv, file_name=nombre, mime="text/csv", key=f"dl_{clave}")  # Botón de descarga


def etiqueta_anio(a):  # Etiqueta de columna para cada año
    return "Año 0 · Dev" if a == 0 else f"Año {a}"  # El Año 0 se marca como desarrollo


def vista_anual(res, raz, p, clave):  # Vista completa de resultados por año
    anual = anualizar(res["m"])  # Agrupa por año
    etiquetas = [etiqueta_anio(a) for a in anual.index]  # Etiquetas de columnas
    seccion(f"Estado de resultados y flujo · cifras en {p['p_moneda']}")  # Título de la sección
    tabla = construir_tabla(res, anual, etiquetas)  # Arma la tabla del reporte
    mostrar_tabla(tabla, etiquetas)  # Muestra el reporte
    descargar(tabla, etiquetas, "resultados_anuales.csv", clave)  # Botón de descarga
    bloque_razones(raz, p, anual, etiquetas, 12, "año")  # Razones financieras sin y con financiamiento
    seccion("Gráficas")  # Título de gráficas
    graficas(anual, etiquetas, clave)  # Gráficas


# -----------------------------------------------------------------------------
# ESTADO DE LA SESIÓN (para conservar lo capturado entre interacciones)
# -----------------------------------------------------------------------------
def params_de(paquete):  # Parámetros completos de un proyecto guardado o plantilla
    p = dict(PARAMS_DEF)  # Empieza con los valores por defecto
    for k, v in paquete.get("parametros", {}).items():  # Recorre los parámetros del paquete
        if k in PARAMS_DEF and v is not None:  # Solo parámetros conocidos y con valor
            p[k] = date.fromisoformat(v) if k in FECHAS and isinstance(v, str) else v  # Convierte fechas guardadas como texto
    return p  # Parámetros completos


def tablas_de(paquete):  # Tablas (DataFrames) de un proyecto guardado o plantilla
    base = config_default()  # Tablas por defecto (por si falta alguna)
    t = paquete.get("tablas", {})  # Tablas del paquete
    return {k: (pd.DataFrame(t[k]) if t.get(k) else pd.DataFrame(columns=base[k].columns)) if k in t else base[k] for k in base}  # Una tabla por sección


def cargar_paquete(paquete, aviso=None):  # Carga un proyecto o plantilla en la app (se usa como callback de botones)
    for k, v in params_de(paquete).items():  # Recorre los parámetros
        st.session_state[k] = v  # Actualiza cada campo de la app
    st.session_state.cfg = tablas_de(paquete)  # Reemplaza las tablas
    st.session_state.ver = st.session_state.get("ver", 0) + 1  # Cambia la versión para refrescar los editores
    st.session_state.aviso = aviso  # Mensaje a mostrar después de cargar


if "cfg" not in st.session_state:  # Si es la primera vez que se abre la app
    st.session_state.cfg = config_default()  # Carga las tablas de ejemplo
    st.session_state.ver = 0  # Versión de las tablas (sirve para refrescarlas al cargar un archivo)
    st.session_state.archivo_cargado = None  # Identificador del último archivo cargado
for llave, valor in PARAMS_DEF.items():  # Recorre los parámetros por defecto
    st.session_state.setdefault(llave, valor)  # Los inicializa si no existen
if st.session_state.get("aviso"):  # Si hay un mensaje pendiente (por ejemplo, proyecto cargado)
    st.toast(st.session_state.aviso)  # Muestra una notificación breve
    st.session_state.aviso = None  # Limpia el mensaje

# -----------------------------------------------------------------------------
# BARRA LATERAL: PARÁMETROS GENERALES Y CARGA DE CONFIGURACIÓN
# -----------------------------------------------------------------------------
with st.sidebar:  # Todo lo siguiente va en la barra lateral
    archivo = st.file_uploader("Cargar configuración (.json)", type="json")  # Permite cargar una configuración guardada
    if archivo is not None and st.session_state.archivo_cargado != archivo.file_id:  # Si es un archivo nuevo
        cargar_paquete(json.load(archivo), "Configuración cargada desde archivo")  # Carga parámetros y tablas del JSON
        st.session_state.archivo_cargado = archivo.file_id  # Recuerda que ya se cargó este archivo
        st.rerun()  # Vuelve a ejecutar la app con los nuevos datos

    seccion("Proyecto")  # Grupo de parámetros del proyecto
    st.text_input("Nombre", key="p_nombre")  # Nombre del proyecto
    st.date_input("Fecha de inicio (desarrollo)", key="p_fecha")  # Fecha de inicio
    c1, c2 = st.columns(2)  # Dos campos en una fila
    c1.number_input("Desarrollo (meses)", min_value=0, max_value=60, step=1, key="p_dev_meses", help="Año 0: etapa sin ingresos")  # Etapa de desarrollo
    c2.number_input("Operación (años)", min_value=1, max_value=30, step=1, key="p_anios")  # Años de operación

    seccion("Moneda y tipo de cambio")  # Grupo de moneda
    st.segmented_control("Moneda de reporte", ["MXN", "USD"], key="p_moneda")  # Moneda de los resultados
    c1, c2 = st.columns(2)  # Dos campos en una fila
    c1.number_input("TC inicial (MXN/USD)", min_value=0.01, step=0.25, format="%.2f", key="p_tc")  # Tipo de cambio inicial
    c2.number_input("Variación TC anual %", step=0.5, key="p_tc_var", help="Positivo = depreciación del peso")  # Variación anual

    seccion("Inflación anual")  # Grupo de inflación
    c1, c2 = st.columns(2)  # Dos campos en una fila
    c1.number_input("México %", step=0.25, key="p_inf_mxn")  # Inflación MXN
    c2.number_input("EE.UU. %", step=0.25, key="p_inf_usd")  # Inflación USD

    seccion("Evaluación e impuestos")  # Grupo de tasas
    c1, c2 = st.columns(2)  # Dos campos en una fila
    c1.number_input("WACC %", min_value=0.0, step=0.5, key="p_tasa_desc")  # Tasa de descuento del proyecto
    c2.number_input("Ke %", min_value=0.0, step=0.5, key="p_ke")  # Costo del accionista
    c1, c2 = st.columns(2)  # Dos campos en una fila
    c1.number_input("ISR %", min_value=0.0, max_value=100.0, step=1.0, key="p_isr")  # ISR
    c2.number_input("PTU %", min_value=0.0, max_value=100.0, step=1.0, key="p_ptu")  # PTU
    st.toggle("Valor de rescate al final", key="p_rescate", help="Suma el valor en libros de los activos en el último mes.")  # Rescate

params = {k: st.session_state[k] for k in PARAMS_DEF}  # Diccionario de parámetros actuales
params["p_moneda"] = params["p_moneda"] or "MXN"  # Si se deselecciona la moneda, usa MXN

# -----------------------------------------------------------------------------
# PESTAÑAS PRINCIPALES
# -----------------------------------------------------------------------------
hero = st.empty()  # Espacio reservado para el encabezado (se llena después de calcular)
tab_proy, tab_car, tab_cfg, tab_anual, tab_mensual, tab_esc, tab_sens, tab_res = st.tabs(  # Pestañas principales
    ["Proyectos", "Carátula", "Supuestos", "Resultados anuales", "Resultados mensuales", "Escenarios", "Sensibilidad", "Resumen ejecutivo"])  # Nombres

# -----------------------------------------------------------------------------
# PESTAÑA 0: CAPTURA DE SUPUESTOS (variables, fórmulas, CAPEX, crédito, escenarios)
# -----------------------------------------------------------------------------
cfg = st.session_state.cfg  # Atajo a las tablas guardadas
ver = st.session_state.ver  # Versión actual de las tablas
col_formula = {  # Configuración común de columnas para tablas con fórmula
    "Concepto": st.column_config.TextColumn("Concepto", required=True),  # Nombre del concepto
    "Etapa": st.column_config.SelectboxColumn("Etapa", options=["Desarrollo", "Operación"], required=True, default="Operación"),  # Etapa
    "Fórmula": st.column_config.TextColumn("Fórmula", required=True, width="large", help="Ej.: litros_vendidos * precio_venta"),  # Fórmula
    "Mes inicio": st.column_config.NumberColumn("Mes inicio", min_value=1, step=1, help="Mes dentro de su etapa. Vacío = 1"),  # Mes inicial
    "Mes fin": st.column_config.NumberColumn("Mes fin", min_value=1, step=1, help="Mes dentro de su etapa. Vacío = hasta el final"),  # Mes final
}  # Fin de la configuración

with tab_cfg:  # Contenido de la pestaña de supuestos
    s_var, s_ing, s_egr, s_corp, s_cap, s_fin, s_esc = st.tabs(  # Sub-pestañas de captura
        ["Variables", "Ingresos", "Egresos", "Gastos corporativos", "CAPEX", "Financiamiento", "Escenarios"])  # Nombres

    with s_var:  # Sub-pestaña de variables
        st.caption("Valores a precios del inicio y en su **moneda original**. El modelo aplica crecimiento real, "  # Ayuda
                   "**inflación** (si se marca) y **tipo de cambio** para llevarlos a la moneda de reporte. "  # (continuación)
                   "Tipo *Volumen* = aparece en Volumetría y vale 0 durante el desarrollo. Moneda «—» = no es dinero.")  # (continuación)
        df_vars = st.data_editor(cfg["variables"], num_rows="dynamic", width="stretch", key=f"vars_{ver}",  # Editor de variables
                                 column_config={  # Configuración de columnas
                                     "Variable": st.column_config.TextColumn("Variable", required=True, help="Sin espacios, ej. precio_venta"),  # Nombre
                                     "Tipo": st.column_config.SelectboxColumn("Tipo", options=["Volumen", "Precio", "Costo", "Otro"], required=True),  # Tipo
                                     "Moneda": st.column_config.SelectboxColumn("Moneda", options=MONEDAS, required=True, default="MXN"),  # Moneda
                                     "Valor": st.column_config.NumberColumn("Valor", format="localized"),  # Valor base
                                     "Crec. anual %": st.column_config.NumberColumn("Crec. real %", format="%.2f %%", help="Crecimiento real anual, sin inflación"),  # Crecimiento
                                     "Inflación": st.column_config.CheckboxColumn("Inflación", default=False, help="Marca si el valor se actualiza con la inflación de su moneda"),  # Inflación
                                 })  # Fin de la configuración
    nombres_vars = [str(v).strip() for v in df_vars["Variable"].dropna() if str(v).strip()]  # Nombres de variables capturadas

    ayuda = (f"**Variables:** `{'`, `'.join(nombres_vars)}`  \n"  # Lista de variables
             "**Tiempo:** `mes` (mes de operación), `anio` (0 = desarrollo), `mes_del_anio`, `mes_proyecto`  \n"  # Variables de tiempo
             "**Funciones:** `min()` `max()` `abs()` `round()` · condicional: `valor if anio >= 3 else 0`")  # Funciones permitidas
    with s_ing:  # Sub-pestaña de ingresos
        st.caption("Los ingresos solo existen en la etapa de **operación**; los meses se cuentan desde el arranque.")  # Aclaración
        st.markdown(ayuda)  # Muestra la ayuda de fórmulas
        df_ing = st.data_editor(cfg["ingresos"], num_rows="dynamic", width="stretch", key=f"ing_{ver}",  # Editor de ingresos
                                column_config={k: v for k, v in col_formula.items() if k != "Etapa"})  # Sin columna de etapa
    with s_egr:  # Sub-pestaña de egresos
        st.markdown(ayuda)  # Muestra la ayuda de fórmulas
        df_egr = st.data_editor(cfg["egresos"], num_rows="dynamic", width="stretch", key=f"egr_{ver}", column_config=col_formula)  # Editor de egresos
    with s_corp:  # Sub-pestaña de gastos corporativos
        st.markdown(ayuda + "  \n**Además:** `ingresos`, `egresos`, `utilidad_bruta` (totales del mes)")  # Ayuda con variables extra
        df_corp = st.data_editor(cfg["corporativos"], num_rows="dynamic", width="stretch", key=f"corp_{ver}", column_config=col_formula)  # Editor corporativo
    with s_cap:  # Sub-pestaña de CAPEX
        st.caption("Año 0 = desarrollo. Vida útil 0 = no se deprecia (terreno). La depreciación inicia al arrancar la operación.")  # Ayuda
        df_capex = st.data_editor(cfg["capex"], num_rows="dynamic", width="stretch", key=f"capex_{ver}",  # Editor de CAPEX
                                  column_config={  # Configuración de columnas
                                      "Moneda": st.column_config.SelectboxColumn("Moneda", options=["MXN", "USD"], required=True, default="MXN"),  # Moneda
                                      "Monto": st.column_config.NumberColumn("Monto", min_value=0.0, format="localized"),  # Monto
                                      "Año": st.column_config.NumberColumn("Año", min_value=0, step=1, default=0),  # Año de inversión
                                      "Mes": st.column_config.NumberColumn("Mes", min_value=1, max_value=60, step=1, default=1),  # Mes dentro del año
                                      "Vida útil (años)": st.column_config.NumberColumn("Vida útil (años)", min_value=0.0),  # Vida útil
                                      "Inflación": st.column_config.CheckboxColumn("Inflación", default=False),  # Inflación
                                  })  # Fin de la configuración
    with s_fin:  # Sub-pestaña de financiamiento
        c = st.columns(6)  # Seis campos en una fila
        c[0].number_input("% del CAPEX financiado", min_value=0.0, max_value=100.0, step=5.0, key="p_fin_pct")  # Porcentaje financiado
        c[1].number_input("Tasa anual %", min_value=0.0, step=0.5, key="p_fin_tasa")  # Tasa del crédito
        c[2].number_input("Comisión %", min_value=0.0, max_value=20.0, step=0.25, format="%.2f", key="p_fin_comision",  # Comisión
                          help="Porcentaje sobre el total del crédito; se paga una sola vez en la primera disposición")  # Ayuda
        c[3].number_input("Plazo (años)", min_value=1, max_value=30, step=1, key="p_fin_plazo")  # Plazo
        c[4].number_input("Meses de gracia", min_value=0, step=1, key="p_fin_gracia")  # Gracia
        c[5].number_input("Mes de disposición", min_value=1, step=1, key="p_fin_mes", help="Mes del proyecto (1 = inicio del desarrollo)")  # Disposición
        st.caption("Pagos fijos mensuales (sistema francés), en la moneda de reporte. En la gracia solo se pagan intereses. "  # Explicación
                   "Si el crédito rebasa el horizonte, el saldo se liquida en el último mes. "  # (continuación)
                   "La **comisión** se calcula sobre el total del crédito y se paga en el mes de la primera disposición; "  # Comisión
                   "se registra como gasto financiero (deducible) y solo afecta el flujo **con** financiamiento.")  # (continuación)
    with s_esc:  # Sub-pestaña de escenarios
        st.caption("Cada renglón ajusta en % una variable, el **CAPEX**, el **Tipo de cambio** o el **WACC** respecto al escenario Base. "  # Explicación
                   "Repite el nombre del escenario para combinar ajustes.")  # (continuación)
        df_esc = st.data_editor(cfg["escenarios"], num_rows="dynamic", width="stretch", key=f"esc_{ver}",  # Editor de escenarios
                                column_config={  # Configuración de columnas
                                    "Escenario": st.column_config.TextColumn("Escenario", required=True),  # Nombre del escenario
                                    "Variable": st.column_config.SelectboxColumn("Ajusta a", options=nombres_vars + ["CAPEX", "Tipo de cambio", "WACC"], required=True),  # Qué se ajusta
                                    "Ajuste %": st.column_config.NumberColumn("Ajuste %", format="%.1f %%"),  # Porcentaje de ajuste
                                })  # Fin de la configuración

    tablas_actuales = {"variables": df_vars, "ingresos": df_ing, "egresos": df_egr,  # Tablas capturadas
                       "corporativos": df_corp, "capex": df_capex, "escenarios": df_esc}  # (continuación)
    paquete = {  # Contenido del archivo a guardar
        "parametros": {k: (st.session_state[k].isoformat() if k in FECHAS else st.session_state[k]) for k in PARAMS_DEF},  # Parámetros
        "tablas": {k: v.astype(object).where(v.notna(), None).to_dict("records") for k, v in tablas_actuales.items()},  # Tablas (sin NaN)
    }  # Fin del paquete
    st.download_button("Guardar configuración (.json)", json.dumps(paquete, ensure_ascii=False, indent=2, default=str),  # Botón de guardado
                       file_name="config_proyecto.json", mime="application/json")  # Nombre del archivo


# -----------------------------------------------------------------------------
# CÁLCULO DEL ESCENARIO BASE Y DE LOS ESCENARIOS
# -----------------------------------------------------------------------------
def aplicar_escenario(ajustes):  # Regresa variables, CAPEX y parámetros ajustados para un escenario
    v = df_vars.copy()  # Copia de variables
    k = df_capex.copy()  # Copia de CAPEX
    p = dict(params)  # Copia de parámetros
    for _, fila in ajustes.iterrows():  # Recorre cada ajuste
        factor = 1 + num(fila.get("Ajuste %")) / 100  # Factor multiplicador
        objetivo = str(fila.get("Variable"))  # Qué se ajusta
        if objetivo == "CAPEX":  # Ajuste a la inversión
            k["Monto"] = pd.to_numeric(k["Monto"], errors="coerce") * factor  # Ajusta todos los montos
        elif objetivo == "Tipo de cambio":  # Ajuste al tipo de cambio
            p["p_tc"] = p["p_tc"] * factor  # Ajusta el tipo de cambio inicial
        elif objetivo == "WACC":  # Ajuste a la tasa de descuento
            p["p_tasa_desc"] = p["p_tasa_desc"] * factor  # Ajusta el WACC
        else:  # Ajuste a una variable
            mascara = v["Variable"].astype(str).str.strip() == objetivo  # Ubica la variable
            v.loc[mascara, "Valor"] = pd.to_numeric(v.loc[mascara, "Valor"], errors="coerce") * factor  # Ajusta su valor
    return v, k, p  # Regresa las tablas y parámetros ajustados


@st.cache_data(show_spinner=False, max_entries=400)  # Guarda resultados ya calculados para no repetirlos
def evaluar(p, variables, ingresos, egresos, corporativos, capex):  # Corre el modelo y calcula sus razones
    res = correr_modelo(p, variables, ingresos, egresos, corporativos, capex)  # Modelo mensual
    return res, razones_globales(res, p)  # Resultados y razones


def solo_modelo(p):  # Quita los datos de la carátula (no afectan el cálculo y así se aprovecha la caché)
    return {k: v for k, v in p.items() if k not in CARATULA_DEF}  # Solo parámetros del modelo


resultados = {}  # Diccionario escenario -> (resultados, razones, parámetros)
res_base, raz_base = evaluar(solo_modelo(params), df_vars, df_ing, df_egr, df_corp, df_capex)  # Corre el escenario base
resultados["Base"] = (res_base, raz_base, params)  # Guarda el escenario base
for nombre_esc, grupo in df_esc.dropna(subset=["Escenario"]).groupby("Escenario", sort=False):  # Recorre cada escenario capturado
    if str(nombre_esc).strip() and str(nombre_esc) != "Base":  # Ignora nombres vacíos o 'Base'
        v_esc, k_esc, p_esc = aplicar_escenario(grupo)  # Aplica los ajustes
        r_esc, raz_esc = evaluar(solo_modelo(p_esc), v_esc, df_ing, df_egr, df_corp, k_esc)  # Corre el modelo
        resultados[str(nombre_esc)] = (r_esc, raz_esc, p_esc)  # Guarda el resultado

with tab_cfg:  # Muestra el estado de la captura en la pestaña de supuestos
    if res_base["errores"]:  # Si hubo errores en fórmulas o variables
        st.error("Hay conceptos con problemas (se calculan como 0):\n\n" + "\n".join(f"- {e}" for e in dict.fromkeys(res_base["errores"])))  # Errores
    else:  # Si no hay errores
        st.success("Todas las fórmulas se calcularon correctamente.")  # Mensaje de éxito

# -----------------------------------------------------------------------------
# ENCABEZADO DE IMPACTO (veredicto del escenario base)
# -----------------------------------------------------------------------------
rb = resultados["Base"][1]  # Razones del escenario base
rentable = finito(rb["VPN proyecto (sin fin.)"]) and rb["VPN proyecto (sin fin.)"] > 0  # El proyecto crea valor si VPN > 0
fecha_op = pd.Period(params["p_fecha"], freq="M") + int(params["p_dev_meses"])  # Mes de arranque de la operación
pintar(  # Dibuja el encabezado
    f'<div class="hero"><div class="eyebrow">Análisis de rentabilidad · {html_lib.escape(str(params["p_folio"]))}</div>'  # Folio
    f'<div class="titulo">{html_lib.escape(str(params["p_nombre"]))}</div>'  # Nombre del proyecto
    f'<div class="sub">{int(params["p_dev_meses"])} meses de desarrollo · operación desde {MESES_ES[fecha_op.month - 1]}-{fecha_op.year} '  # Etapas
    f'· {int(params["p_anios"])} años · cifras en {params["p_moneda"]} · TC {params["p_tc"]:.2f} MXN/USD</div>'  # Moneda y tipo de cambio
    f'<span class="badge {"ok" if rentable else "no"}">{"● Crea valor" if rentable else "● No crea valor"} · '  # Veredicto
    f'VPN {dinero(rb["VPN proyecto (sin fin.)"])} · TIR {pct(rb["TIR proyecto (anual)"])}</span></div>',  # VPN y TIR
    hero)  # Se dibuja en el espacio reservado arriba de las pestañas

# -----------------------------------------------------------------------------
# PESTAÑA 1: RESULTADOS POR AÑO
# -----------------------------------------------------------------------------
with tab_anual:  # Contenido de la pestaña anual
    vista_anual(res_base, rb, params, "anual")  # Reporte anual completo del escenario base

# -----------------------------------------------------------------------------
# PESTAÑA 2: RESULTADOS MENSUALES
# -----------------------------------------------------------------------------
with tab_mensual:  # Contenido de la pestaña mensual
    m = res_base["m"]  # Tabla mensual del escenario base
    inicio = pd.Period(params["p_fecha"], freq="M")  # Mes calendario de inicio
    etiquetas_m = [f"{MESES_ES[(inicio + i).month - 1]}-{(inicio + i).year}" for i in range(len(m))]  # Etiquetas tipo 'ene-2027'
    anios_disp = sorted(m["anio"].unique())  # Años disponibles (incluye el 0)
    opciones = ["Todos"] + [etiqueta_anio(a) for a in anios_disp]  # Opciones del filtro
    eleccion = st.segmented_control("Periodo", opciones, default=opciones[1] if len(opciones) > 1 else "Todos", key="filtro_mes")  # Filtro
    eleccion = eleccion if eleccion in opciones else "Todos"  # Si se deselecciona o ya no existe, muestra todos
    filtro = np.full(len(m), True) if eleccion == "Todos" else (m["anio"] == anios_disp[opciones.index(eleccion) - 1]).values  # Meses a mostrar
    datos_m = m[filtro]  # Datos filtrados
    etiquetas_f = [e for e, ok in zip(etiquetas_m, filtro) if ok]  # Etiquetas filtradas
    seccion(f"Estado de resultados y flujo mensual · cifras en {params['p_moneda']}")  # Título
    tabla_m = construir_tabla(res_base, datos_m, etiquetas_f)  # Arma la tabla mensual
    mostrar_tabla(tabla_m, etiquetas_f)  # Muestra el reporte mensual
    descargar(tabla_m, etiquetas_f, "resultados_mensuales.csv", "mensual")  # Botón de descarga
    bloque_razones(rb, params, datos_m, etiquetas_f, 1, "mes")  # Razones financieras sin y con financiamiento
    seccion("Gráficas")  # Título de gráficas
    graficas(datos_m, etiquetas_f, "mensual")  # Gráficas mensuales

# -----------------------------------------------------------------------------
# PESTAÑA 3: ESCENARIOS
# -----------------------------------------------------------------------------
with tab_esc:  # Contenido de la pestaña de escenarios
    seccion("Comparativo")  # Título del comparativo
    cols = st.columns(len(resultados))  # Una columna por escenario
    colores_esc = ["azul", "rosa", "teal", "morado", "ambar", "indigo"]  # Color de tarjeta para cada escenario
    paleta = [AZUL, NEGATIVO, ACENTO, MORADO, AMBAR, INDIGO]  # Mismos colores para las gráficas
    for i, (col, (nombre, (_, raz, _))) in enumerate(zip(cols, resultados.items())):  # Recorre cada escenario
        v = raz["VPN proyecto (sin fin.)"]  # VPN del escenario
        pintar(tarjeta(nombre, dinero(v), f"TIR {pct(raz['TIR proyecto (anual)'])} · TIR acc. {pct(raz['TIR accionista (anual)'])}",  # Tarjeta resumen
                       "pos" if finito(v) and v > 0 else "neg", colores_esc[i % len(colores_esc)]), col)  # Color del escenario

    comparativo = pd.DataFrame({nombre: raz for nombre, (_, raz, _) in resultados.items()})  # Filas = razones, columnas = escenarios
    vista_comp = comparativo.apply(lambda col: [fmt_razon(i, v) for i, v in col.items()])  # Da formato a cada celda
    vista_comp.insert(0, "Evaluación", ["General" if i in RAZONES_GENERALES else  # Clasifica cada razón
                                        "Sin financiamiento" if i in RAZONES_SIN_FIN else "Con financiamiento"  # (continuación)
                                        for i in vista_comp.index])  # Columna de grupo
    color_grupo = {"General": SUAVE, "Sin financiamiento": ACENTO, "Con financiamiento": NAVY}  # Color de cada grupo
    st.dataframe(vista_comp.style.apply(lambda f: [f"border-left:4px solid {color_grupo[f['Evaluación']]}; "  # Barra de color por grupo
                                                   f"background-color:{rgba(color_grupo[f['Evaluación']], .06)}"] * len(f), axis=1),  # Fondo tenue
                 width="stretch", height=35 * (len(vista_comp) + 1) + 5)  # Muestra el comparativo completo

    c1, c2 = st.columns(2)  # Dos gráficas lado a lado
    fig = go.Figure()  # Gráfica de VPN por escenario
    fig.add_bar(x=list(resultados), y=[r[1]["VPN proyecto (sin fin.)"] for r in resultados.values()], name="VPN proyecto", marker_color=MORADO)  # VPN proyecto
    fig.add_bar(x=list(resultados), y=[r[1]["VPN accionista (con fin.)"] for r in resultados.values()], name="VPN accionista", marker_color=ACENTO)  # VPN accionista
    c1.plotly_chart(estilo_fig(fig, "VPN por escenario"), width="stretch", key="esc_vpn")  # Muestra la gráfica
    fig = go.Figure()  # Gráfica de flujo acumulado por escenario
    for i, (nombre, (r, _, _)) in enumerate(resultados.items()):  # Recorre cada escenario
        anual_e = anualizar(r["m"])  # Datos anuales del escenario
        fig.add_scatter(x=[etiqueta_anio(a) for a in anual_e.index], y=anual_e["FE con fin. acumulado"], name=nombre,  # Línea del acumulado
                        mode="lines", line=dict(color=paleta[i % len(paleta)], width=2.5))  # Color y grosor
    c2.plotly_chart(estilo_fig(fig, "Flujo acumulado con financiamiento"), width="stretch", key="esc_acum")  # Muestra la gráfica

    st.divider()  # Línea separadora
    nombre_sel = st.segmented_control("Detalle del escenario", list(resultados), default="Base", key="esc_sel")  # Selector de escenario
    nombre_sel = nombre_sel if nombre_sel in resultados else "Base"  # Si se deselecciona o ya no existe, usa Base
    ajustes_sel = df_esc[df_esc["Escenario"] == nombre_sel]  # Ajustes del escenario elegido
    if len(ajustes_sel):  # Si el escenario tiene ajustes
        st.caption("Ajustes: " + " · ".join(f"{r['Variable']} {num(r['Ajuste %']):+.1f}%" for _, r in ajustes_sel.iterrows()))  # Muestra los ajustes
    r_sel, raz_sel, p_sel = resultados[nombre_sel]  # Resultados del escenario elegido
    vista_anual(r_sel, raz_sel, p_sel, f"esc_{nombre_sel}")  # Reporte anual del escenario elegido


# -----------------------------------------------------------------------------
# PESTAÑA 5: ANÁLISIS DE SENSIBILIDAD
# -----------------------------------------------------------------------------
def nombre_objetivo(o):  # Nombre legible de un supuesto
    return bonito(o) if o in nombres_vars else o  # Variables con formato; CAPEX, TC y WACC tal cual


def correr_ajuste(objetivo, pct_aj):  # Corre el modelo moviendo un solo supuesto en un porcentaje
    if abs(pct_aj) < 1e-9:  # Sin cambio
        return rb  # Regresa las razones del escenario base
    v, k, p = aplicar_escenario(pd.DataFrame([{"Variable": objetivo, "Ajuste %": pct_aj}]))  # Aplica el ajuste
    return evaluar(solo_modelo(p), v, df_ing, df_egr, df_corp, k)[1]  # Razones del caso ajustado


def valor_base_de(objetivo):  # Valor original del supuesto y su unidad
    if objetivo == "CAPEX":  # Inversión total
        return rb["Inversión total (CAPEX)"], params["p_moneda"]  # Monto en moneda de reporte
    if objetivo == "Tipo de cambio":  # Tipo de cambio
        return params["p_tc"], "MXN/USD"  # Tipo de cambio inicial
    if objetivo == "WACC":  # Tasa de descuento
        return params["p_tasa_desc"], "%"  # WACC en %
    fila = df_vars[df_vars["Variable"].astype(str).str.strip() == objetivo]  # Busca la variable
    if fila.empty:  # Si no existe
        return np.nan, ""  # Sin dato
    f = fila.iloc[0]  # Primera coincidencia
    partes = [str(x) for x in (f.get("Moneda"), f.get("Unidad")) if x is not None and str(x) not in ("nan", "—", "")]  # Moneda y unidad
    return num(f.get("Valor"), np.nan), " ".join(partes)  # Valor y unidad


def fmt_ind(indicador, x):  # Formato de un indicador (dinero, porcentaje o años)
    if indicador.startswith("TIR"):  # TIR
        return pct(x)  # En porcentaje
    if es_recuperacion(indicador):  # Periodo de recuperación
        return f"{x:.1f} años" if finito(x) else "No recupera"  # En años
    return dinero(x)  # VPN en dinero


def formato_eje(indicador):  # Formato del eje de las gráficas según el indicador
    return ".0%" if indicador.startswith("TIR") else (".1f" if es_recuperacion(indicador) else "~s")  # %, años o montos


sens = None  # Tabla de sensibilidad (se usa también en el Excel)
top_sensibles = []  # Supuestos con mayor impacto (se usan en el resumen)
with tab_sens:  # Contenido de la pestaña de sensibilidad
    seccion("Análisis de sensibilidad · un supuesto a la vez")  # Título
    objetivos = nombres_vars + ["CAPEX", "Tipo de cambio", "WACC"]  # Supuestos que se pueden sensibilizar
    c1, c2, c3, c4 = st.columns([2, 2, 1, 2])  # Controles en una fila
    var_sens = c1.selectbox("Supuesto a sensibilizar", objetivos, format_func=nombre_objetivo,  # Selector del supuesto
                            index=objetivos.index("precio_venta") if "precio_venta" in objetivos else 0, key="sens_var")  # Por defecto el precio
    indicador = c2.selectbox("Indicador a observar", list(INDICADORES), format_func=lambda i: INDICADORES[i], key="sens_ind")  # Indicador
    pasos = c3.selectbox("Puntos", [5, 7, 9, 11], index=2, key="sens_pasos")  # Número de puntos a evaluar
    rango = c4.slider("Rango de variación ±%", 5, 50, 20, 5, key="sens_rango")  # Rango de variación

    base_val, unidad = valor_base_de(var_sens)  # Valor original del supuesto
    variaciones = np.linspace(-rango, rango, pasos)  # Porcentajes de variación a evaluar
    with st.spinner("Calculando sensibilidad..."):  # Indicador de progreso
        filas = []  # Renglones de la tabla
        for pc in variaciones:  # Recorre cada variación
            raz = correr_ajuste(var_sens, float(pc))  # Corre el modelo con el ajuste
            filas.append({"Variación %": pc, f"Valor del supuesto ({unidad})" if unidad else "Valor del supuesto": base_val * (1 + pc / 100),  # Variación y valor
                          **{INDICADORES[i]: raz[i] for i in INDICADORES}})  # Indicadores resultantes
        sens = pd.DataFrame(filas)  # Tabla de sensibilidad

    y = sens[INDICADORES[indicador]].values.astype(float)  # Valores del indicador elegido
    recuperacion = es_recuperacion(indicador)  # ¿El indicador es un periodo de recuperación?
    if indicador.startswith("VPN"):  # Para el VPN
        umbral, txt_umbral = 0.0, "VPN = 0"  # El equilibrio es VPN igual a cero
    elif recuperacion:  # Para la recuperación no hay umbral fijo: interesa cuándo deja de recuperarse
        umbral, txt_umbral = None, f"recuperación dentro de {int(params['p_dev_meses']) / 12 + int(params['p_anios']):.1f} años"  # Horizonte
    elif "proyecto" in indicador:  # Para la TIR del proyecto
        umbral, txt_umbral = params["p_tasa_desc"] / 100, f"TIR = WACC ({params['p_tasa_desc']:.1f}%)"  # Equilibrio contra el WACC
    else:  # Para la TIR del accionista
        umbral, txt_umbral = params["p_ke"] / 100, f"TIR = Ke ({params['p_ke']:.1f}%)"  # Equilibrio contra el Ke
    equilibrio = np.nan  # Variación en la que se alcanza el equilibrio (o se deja de recuperar)
    if recuperacion:  # Busca la variación más cercana a 0% en la que ya no se recupera la inversión
        sin_recuperar = [v for v, val in zip(variaciones, y) if not finito(val)]  # Variaciones sin recuperación
        equilibrio = min(sin_recuperar, key=abs) if sin_recuperar else np.nan  # La más cercana al escenario base
    else:  # Para VPN y TIR busca el cruce con el umbral
        for j in range(len(y) - 1):  # Recorre puntos consecutivos
            a, b = y[j] - umbral, y[j + 1] - umbral  # Distancia al umbral en ambos puntos
            if finito(a) and finito(b) and a * b <= 0 and a != b:  # Si hay cambio de signo
                equilibrio = variaciones[j] + (variaciones[j + 1] - variaciones[j]) * (-a) / (b - a)  # Interpolación lineal
                break  # Toma el primer cruce
    validos = [(v, val) for v, val in zip(variaciones, y) if finito(val)]  # Puntos con valor (la recuperación puede faltar)
    pendiente = ((validos[-1][1] - validos[0][1]) / (validos[-1][0] - validos[0][0])  # Cambio del indicador por cada 1%
                 if len(validos) >= 2 and validos[-1][0] != validos[0][0] else np.nan)  # Solo si hay dos puntos válidos
    base_ind = rb[indicador]  # Valor del indicador en el escenario base
    if indicador.startswith("TIR"):  # Texto del impacto para la TIR
        txt_pend = f"{pendiente * 100:+.2f} pp" if finito(pendiente) else "N/D"  # Puntos porcentuales
    elif recuperacion:  # Texto del impacto para la recuperación
        txt_pend = f"{pendiente * 12:+.1f} meses" if finito(pendiente) else "N/D"  # Meses por cada 1%
    else:  # Texto del impacto para el VPN
        txt_pend = ("+" if pendiente >= 0 else "") + dinero(pendiente) if finito(pendiente) else "N/D"  # Monto
    if recuperacion:  # Tono de las tarjetas de recuperación: verde si recupera y no tarda más que el base
        bueno = lambda v: "pos" if finito(v) and (not finito(base_ind) or v <= base_ind + 1e-9) else "neg"  # Menos años es mejor
        eq_valor = f"{equilibrio:+.0f}%" if finito(equilibrio) else "Siempre recupera"  # Valor de la tarjeta
        eq_sub = "primer punto evaluado sin recuperación" if finito(equilibrio) else f"en todo el rango ±{rango}%"  # Subtítulo
        eq_titulo = "Límite de recuperación"  # Título de la tarjeta
    else:  # VPN y TIR
        bueno = lambda v: "pos" if finito(v) and v > umbral else "neg"  # Verde si supera el umbral
        eq_valor = f"{equilibrio:+.1f}%" if finito(equilibrio) else f"Fuera de ±{rango}%"  # Valor de la tarjeta
        eq_sub, eq_titulo = txt_umbral, "Punto de equilibrio"  # Subtítulo y título
    pintar('<div class="kpi-grid">' +  # Tarjetas del análisis
           tarjeta(f"{INDICADORES[indicador]} base", fmt_ind(indicador, base_ind),  # Valor base
                   f"{nombre_objetivo(var_sens)} = {base_val:,.2f} {unidad}" if finito(base_val) else "", bueno(base_ind), "navy") +  # Supuesto base
           tarjeta(f"Supuesto −{rango}%", fmt_ind(indicador, y[0]), "escenario a la baja", bueno(y[0]), "morado") +  # Caso a la baja
           tarjeta(f"Supuesto +{rango}%", fmt_ind(indicador, y[-1]), "escenario al alza", bueno(y[-1]), "dorado") +  # Caso al alza
           tarjeta("Impacto por cada 1%", txt_pend, f"cambio en {INDICADORES[indicador]}", color="azul") +  # Pendiente
           tarjeta(eq_titulo, eq_valor, eq_sub, color="teal") +  # Equilibrio o límite de recuperación
           '</div>')  # Cierre de la cuadrícula

    c1, c2 = st.columns([3, 2])  # Gráfica y tabla
    fig = go.Figure()  # Gráfica de sensibilidad
    fig.add_scatter(x=variaciones, y=y, mode="lines+markers", name=INDICADORES[indicador],  # Curva del indicador
                    line=dict(color=AZUL, width=3), marker=dict(size=8, color=AZUL))  # Estilo de la curva
    if umbral is not None:  # Solo VPN y TIR tienen línea de equilibrio
        fig.add_hline(y=umbral, line_dash="dash", line_color=NEGATIVO, annotation_text=txt_umbral, annotation_position="bottom right")  # Línea de equilibrio
    elif finito(equilibrio):  # En recuperación se marca dónde deja de recuperarse
        fig.add_vline(x=equilibrio, line_dash="dash", line_color=NEGATIVO, annotation_text="No recupera", annotation_position="top")  # Límite
    fig.add_scatter(x=[0], y=[base_ind], mode="markers", name="Base",  # Punto del escenario base
                    marker=dict(size=15, color=DORADO, line=dict(color=NAVY, width=2)))  # Marcador dorado
    estilo_fig(fig, f"{INDICADORES[indicador]} vs. variación de {nombre_objetivo(var_sens)}")  # Estilo
    fig.update_xaxes(ticksuffix="%", title="Variación del supuesto")  # Eje X en porcentaje
    fig.update_yaxes(tickformat=formato_eje(indicador), ticksuffix=" años" if recuperacion else "")  # Formato del eje Y
    c1.plotly_chart(fig, width="stretch", key="sens_linea")  # Muestra la gráfica

    col_valor = [c for c in sens.columns if c.startswith("Valor del supuesto")][0]  # Nombre de la columna del valor
    estilo_sens = (sens.style  # Formato de la tabla de sensibilidad
                   .format({"Variación %": "{:+.0f}%", col_valor: "{:,.2f}",  # Variación y valor
                            "VPN proyecto": dinero_full, "VPN accionista": dinero_full,  # VPN en dinero
                            "TIR proyecto": pct, "TIR accionista": pct,  # TIR en porcentaje
                            "Recuperación proyecto": lambda x: fmt_ind("Recuperación", x),  # Recuperación en años
                            "Recuperación accionista": lambda x: fmt_ind("Recuperación", x)})  # Recuperación en años
                   .apply(lambda f: [f"background-color:{rgba(DORADO, .18)}; font-weight:700" if abs(f["Variación %"]) < 1e-9 else ""] * len(f), axis=1))  # Resalta la base
    c2.dataframe(estilo_sens, hide_index=True, width="stretch", height=38 * (len(sens) + 1))  # Muestra la tabla

    seccion(f"Gráfico tornado · ¿qué supuesto mueve más el {INDICADORES[indicador]}?")  # Título del tornado
    st.caption(f"Cada barra muestra el {INDICADORES[indicador]} al mover un supuesto −{rango}% y +{rango}%, dejando todo lo demás igual.")  # Explicación
    with st.spinner("Calculando tornado..."):  # Indicador de progreso
        torn = []  # Lista (supuesto, valor bajo, valor alto)
        for o in objetivos:  # Recorre cada supuesto
            lo = correr_ajuste(o, -rango)[indicador]  # Indicador con el supuesto a la baja
            hi = correr_ajuste(o, rango)[indicador]  # Indicador con el supuesto al alza
            if finito(lo) and finito(hi) and abs(hi - lo) > 1e-9:  # Solo supuestos con impacto
                torn.append((o, lo, hi))  # Agrega a la lista
    torn.sort(key=lambda t: abs(t[2] - t[1]))  # Ordena de menor a mayor impacto (el mayor queda arriba)
    top_sensibles = list(reversed(torn))[:3]  # Los tres supuestos de mayor impacto
    if torn:  # Si hay supuestos con impacto
        etiquetas_t = [nombre_objetivo(o) for o, _, _ in torn]  # Nombres de los supuestos
        fig = go.Figure()  # Gráfica tornado
        fig.add_bar(y=etiquetas_t, x=[lo - base_ind for _, lo, _ in torn], base=base_ind, orientation="h",  # Barra del caso a la baja
                    name=f"Supuesto −{rango}%", marker_color=MORADO)  # Color pizarra
        fig.add_bar(y=etiquetas_t, x=[hi - base_ind for _, _, hi in torn], base=base_ind, orientation="h",  # Barra del caso al alza
                    name=f"Supuesto +{rango}%", marker_color=DORADO)  # Color dorado
        fig.add_vline(x=base_ind, line_color=NAVY, line_width=2)  # Línea del valor base
        estilo_fig(fig, f"Impacto en {INDICADORES[indicador]} (base {fmt_ind(indicador, base_ind)})")  # Estilo
        fig.update_layout(barmode="overlay", height=max(320, 42 * len(torn) + 120))  # Barras superpuestas y altura
        fig.update_xaxes(tickformat=formato_eje(indicador), showgrid=True, gridcolor="rgba(128,128,128,0.15)")  # Eje X
        st.plotly_chart(fig, width="stretch", key="sens_tornado")  # Muestra el tornado


# -----------------------------------------------------------------------------
# TEXTOS COMUNES PARA CARÁTULA, RESUMEN Y CORREO
# -----------------------------------------------------------------------------
def esc(k):  # Texto de la carátula escapado para HTML
    v = params.get(k)  # Valor del parámetro
    return html_lib.escape(str(v)) if v not in (None, "") else "—"  # Guion si está vacío


def supuestos_generales(p):  # Lista (supuesto, valor) con los parámetros principales
    fecha_op = pd.Period(p["p_fecha"], freq="M") + int(p["p_dev_meses"])  # Mes de arranque de operación
    return [  # Lista de supuestos
        ("Inicio del proyecto", fecha_larga(p["p_fecha"])),  # Fecha de inicio
        ("Etapa de desarrollo (Año 0)", f"{int(p['p_dev_meses'])} meses, sin ingresos"),  # Desarrollo
        ("Arranque de operación", f"{MESES_LARGOS[fecha_op.month - 1]} de {fecha_op.year}"),  # Arranque
        ("Horizonte de operación", f"{int(p['p_anios'])} años"),  # Operación
        ("Moneda de reporte", p["p_moneda"]),  # Moneda
        ("Tipo de cambio", f"{p['p_tc']:.2f} MXN/USD inicial, {p['p_tc_var']:+.1f}% anual"),  # Tipo de cambio
        ("Inflación anual", f"México {p['p_inf_mxn']:.1f}% · EE.UU. {p['p_inf_usd']:.1f}%"),  # Inflación
        ("Tasa de descuento", f"WACC {p['p_tasa_desc']:.1f}% · Ke {p['p_ke']:.1f}%"),  # Tasas
        ("Impuestos", f"ISR {p['p_isr']:.0f}% · PTU {p['p_ptu']:.0f}%"),  # Impuestos
        ("Financiamiento", f"{p['p_fin_pct']:.0f}% del CAPEX a {p['p_fin_tasa']:.1f}% anual, {int(p['p_fin_plazo'])} años, "  # Crédito
                           f"{int(p['p_fin_gracia'])} meses de gracia"),  # (continuación)
        ("Comisión por financiamiento", f"{num(p.get('p_fin_comision')):.2f}% sobre el total del crédito, pagada en la primera disposición"),  # Comisión
        ("Valor de rescate", "Sí (valor en libros al final)" if p["p_rescate"] else "No"),  # Rescate
    ]  # Fin de la lista


def conclusion(p, raz):  # Redacta la conclusión del estudio (HTML) y si el proyecto crea valor
    vpn, tir_p, tir_c = raz["VPN proyecto (sin fin.)"], raz["TIR proyecto (anual)"], raz["TIR accionista (anual)"]  # Indicadores
    crea = finito(vpn) and vpn > 0  # Crea valor si el VPN es positivo
    t = (f"Con los supuestos utilizados, el proyecto <b>{'crea valor' if crea else 'no crea valor'}</b>: "  # Veredicto
         f"el VPN es de {dinero_full(vpn)} {p['p_moneda']} descontado al {p['p_tasa_desc']:.1f}% (WACC)")  # VPN
    if finito(tir_p):  # Si hay TIR
        t += f" y la TIR es de {pct(tir_p)}, {'superior' if tir_p * 100 > p['p_tasa_desc'] else 'inferior'} al costo de capital"  # TIR vs WACC
    t += ". "  # Fin de la oración
    pb = raz["Periodo de recuperación (años)"]  # Payback
    t += (f"La inversión de {dinero_full(raz['Inversión total (CAPEX)'])} se recupera en {pb:.1f} años desde el inicio del desarrollo. "  # Recuperación
          if finito(pb) else "La inversión no se recupera dentro del horizonte evaluado. ")  # Sin recuperación
    if p["p_fin_pct"] > 0 and finito(tir_c):  # Si hay financiamiento
        t += f"Para el accionista, con {p['p_fin_pct']:.0f}% de deuda, la TIR es de {pct(tir_c)} frente a un Ke de {p['p_ke']:.1f}%. "  # Accionista
    otros = [(n, r[1]["VPN proyecto (sin fin.)"]) for n, r in resultados.items() if n != "Base"]  # Escenarios alternativos
    if otros:  # Si hay escenarios
        positivos = sum(1 for _, v in otros if finito(v) and v > 0)  # Escenarios con VPN positivo
        peor = min(otros, key=lambda x: x[1] if finito(x[1]) else -1e30)  # Escenario más adverso
        t += (f"De {len(otros)} escenario(s) alternativo(s), {positivos} mantiene(n) VPN positivo; "  # Escenarios
              f"el más adverso («{html_lib.escape(peor[0])}») arroja un VPN de {dinero_full(peor[1])}. ")  # Peor caso
    if top_sensibles:  # Si se calculó la sensibilidad
        t += ("Los supuestos con mayor impacto en el resultado son: "  # Sensibilidad
              + ", ".join(html_lib.escape(nombre_objetivo(o)) for o, _, _ in top_sensibles) + ".")  # Lista
    return t, crea  # Texto y veredicto


texto_conclusion, crea_valor = conclusion(params, rb)  # Conclusión del escenario base
texto_conclusion_plano = re.sub("<[^>]+>", "", html_lib.unescape(texto_conclusion))  # Versión sin HTML (Excel y correo)

# -----------------------------------------------------------------------------
# PESTAÑA 0: CARÁTULA DEL PROYECTO
# -----------------------------------------------------------------------------
with tab_car:  # Contenido de la carátula
    with st.expander("Editar datos de la carátula", expanded=False):  # Formulario de datos informativos
        c1, c2, c3 = st.columns(3)  # Tres columnas de campos
        c1.text_input("Folio del estudio", key="p_folio")  # Folio
        c2.selectbox("Tipo de proyecto", TIPOS_PROYECTO, key="p_tipo")  # Tipo
        c3.date_input("Fecha del estudio", key="p_fecha_estudio")  # Fecha del estudio
        c1.text_input("Empresa", key="p_empresa")  # Empresa
        c2.text_input("Área solicitante", key="p_area")  # Área
        c3.text_input("Ubicación", key="p_ubicacion")  # Ubicación
        c1.text_input("Solicitante", key="p_solicitante")  # Solicitante
        c2.text_input("Correo del solicitante", key="p_correo_sol")  # Correo
        c3.text_input("Responsable del análisis", key="p_responsable")  # Responsable
        st.text_area("Descripción del proyecto", key="p_descripcion", height=90)  # Descripción
        st.text_area("Objetivo del estudio", key="p_objetivo", height=90)  # Objetivo
        st.caption("El nombre del proyecto se edita en la barra lateral.")  # Aclaración

    fecha_op = pd.Period(params["p_fecha"], freq="M") + int(params["p_dev_meses"])  # Mes de arranque de operación
    datos_cover = [  # Datos que aparecen en la carátula
        ("Empresa", esc("p_empresa")), ("Área solicitante", esc("p_area")), ("Solicitante", esc("p_solicitante")),  # Quién
        ("Responsable del análisis", esc("p_responsable")), ("Ubicación", esc("p_ubicacion")),  # Responsable y lugar
        ("Inicio del desarrollo", fecha_larga(params["p_fecha"])),  # Inicio
        ("Arranque de operación", f"{MESES_LARGOS[fecha_op.month - 1]} de {fecha_op.year}"),  # Arranque
        ("Horizonte", f"{int(params['p_dev_meses'])} meses de desarrollo + {int(params['p_anios'])} años"),  # Horizonte
        ("Moneda de reporte", f"{params['p_moneda']} · TC {params['p_tc']:.2f}"),  # Moneda
        ("Inversión total", dinero_full(rb["Inversión total (CAPEX)"])),  # CAPEX
        ("Financiamiento", f"{params['p_fin_pct']:.0f}% del CAPEX · {params['p_fin_tasa']:.1f}% anual"),  # Deuda
        ("Escenarios evaluados", f"{len(resultados)} (incluye Base)"),  # Escenarios
    ]  # Fin de la lista
    rejilla = "".join(f'<div><div class="k">{k}</div><div class="v">{v}</div></div>' for k, v in datos_cover)  # Celdas de la rejilla
    pintar(  # Dibuja la carátula
        f'<div class="cover"><div class="top">'  # Bloque superior
        f'<div class="folio">Folio<br><b>{esc("p_folio")}</b><br>{fecha_larga(params["p_fecha_estudio"])}</div>'  # Folio y fecha
        f'<div class="eyebrow">Análisis de rentabilidad · Proyecto de inversión</div>'  # Antetítulo
        f'<div class="titulo">{esc("p_nombre")}</div><span class="tipo">{esc("p_tipo")}</span>'  # Nombre y tipo
        f'<div class="desc">{esc("p_descripcion")}</div></div>'  # Descripción
        f'<div class="body"><div class="grid">{rejilla}</div>'  # Datos del estudio
        f'<div class="obj"><div class="k">Objetivo del estudio</div>{esc("p_objetivo")}</div></div></div>')  # Objetivo
    seccion("Resultado del escenario base")  # Título de resultados
    pb = rb["Periodo de recuperación (años)"]  # Payback
    pintar('<div class="kpi-grid">' +  # Tarjetas principales
           tarjeta("Veredicto", "Crea valor" if crea_valor else "No crea valor", "VPN del proyecto > 0" if crea_valor else "VPN del proyecto ≤ 0",  # Veredicto
                   "pos" if crea_valor else "neg", "navy") +  # Color
           tarjeta("VPN proyecto", dinero(rb["VPN proyecto (sin fin.)"]), f"WACC {params['p_tasa_desc']:.1f}%",  # VPN
                   "pos" if crea_valor else "neg", "teal") +  # Color
           tarjeta("TIR proyecto", pct(rb["TIR proyecto (anual)"]), f"TIR accionista {pct(rb['TIR accionista (anual)'])}", color="azul") +  # TIR
           tarjeta("Recuperación", f"{pb:.1f} años" if finito(pb) else "No recupera", "desde el inicio del desarrollo", color="dorado") +  # Payback
           '</div>')  # Cierre

# -----------------------------------------------------------------------------
# PESTAÑA 6: RESUMEN EJECUTIVO (para enviar por correo al solicitante)
# -----------------------------------------------------------------------------
def tabla_html(encabezados, filas, numericas=()):  # Arma una tabla HTML simple
    th = "".join(f'<th class="{"n" if i in numericas else ""}">{h}</th>' for i, h in enumerate(encabezados))  # Encabezados
    cuerpo = ""  # Renglones
    for fila in filas:  # Recorre cada renglón
        celdas = ""  # Celdas del renglón
        for i, v in enumerate(fila):  # Recorre cada celda
            clase = "n" if i in numericas else ""  # Alineación a la derecha para números
            if i in numericas and str(v).startswith("-"):  # Si es negativo
                clase += " neg"  # Se pinta en rojo
            celdas += f'<td class="{clase}">{v}</td>'  # Agrega la celda
        cuerpo += f"<tr>{celdas}</tr>"  # Agrega el renglón
    return f"<table><thead><tr>{th}</tr></thead><tbody>{cuerpo}</tbody></table>"  # Tabla completa


def kv_html(pares):  # Tabla de dos columnas (concepto: valor)
    return '<table class="kv"><tbody>' + "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in pares) + "</tbody></table>"  # Tabla


def limpio(x):  # Convierte un valor de tabla a texto seguro (vacío si es nulo)
    return "" if x is None or (isinstance(x, float) and np.isnan(x)) else html_lib.escape(str(x))  # Texto escapado


def generar_resumen():  # Genera el HTML del resumen ejecutivo
    p = params  # Parámetros actuales
    anual = anualizar(res_base["m"])  # Resultados anuales del escenario base
    etiq = [etiqueta_anio(a) for a in anual.index]  # Etiquetas de años
    claves = [("Ingresos", "Ingresos"), ("Egresos", "Egresos"), ("Gastos corporativos", "Gastos corporativos"),  # Renglones clave
              ("EBITDA", "EBITDA"), ("Utilidad neta", "Utilidad neta"), ("CAPEX", "CAPEX"),  # (continuación)
              ("Flujo sin financiamiento", "FE sin financiamiento"), ("Flujo con financiamiento", "FE con financiamiento")]  # Flujos
    filas_res = [[n] + [f"{v / 1e6:,.1f}" for v in anual[c]] + [f"{anual[c].sum() / 1e6:,.1f}"] for n, c in claves]  # Cifras en millones
    tabla_res = tabla_html(["Concepto"] + etiq + ["Total"], filas_res, numericas=range(1, len(etiq) + 2))  # Tabla de resultados

    filas_var = [[html_lib.escape(bonito(f.get("Variable"))), limpio(f.get("Descripción")), limpio(f.get("Unidad")), limpio(f.get("Moneda")),  # Variables
                  f"{num(f.get('Valor')):,.2f}", f"{num(f.get('Crec. anual %')):.1f}%", "Sí" if si(f.get("Inflación")) else "No"]  # (continuación)
                 for _, f in df_vars.iterrows() if str(f.get("Variable") or "").strip() not in ("", "nan")]  # Solo filas con nombre
    tabla_var = tabla_html(["Variable", "Descripción", "Unidad", "Moneda", "Valor", "Crec. real", "Inflación"], filas_var, numericas=(4, 5))  # Tabla

    filas_form = []  # Conceptos con su fórmula
    for tipo, df in (("Ingreso", df_ing), ("Egreso", df_egr), ("Corporativo", df_corp)):  # Recorre cada tabla
        for _, f in df.iterrows():  # Recorre cada concepto
            if str(f.get("Concepto") or "").strip() not in ("", "nan"):  # Solo conceptos con nombre
                filas_form.append([tipo, limpio(f.get("Concepto")), limpio(f.get("Etapa")) or "Operación", f"<code>{limpio(f.get('Fórmula'))}</code>"])  # Renglón
    tabla_form = tabla_html(["Tipo", "Concepto", "Etapa", "Fórmula"], filas_form)  # Tabla de fórmulas

    filas_cap = [[limpio(f.get("Concepto")), limpio(f.get("Moneda")), f"{num(f.get('Monto')):,.0f}", f"{int(num(f.get('Año')))}",  # CAPEX
                  f"{int(num(f.get('Mes'), 1))}", f"{num(f.get('Vida útil (años)')):.0f}"]  # (continuación)
                 for _, f in df_capex.iterrows() if num(f.get("Monto")) > 0]  # Solo inversiones con monto
    tabla_cap = tabla_html(["Concepto", "Moneda", "Monto", "Año", "Mes", "Vida útil (años)"], filas_cap, numericas=(2, 3, 4, 5))  # Tabla

    filas_esc = []  # Comparativo de escenarios
    for n, (_, raz, _) in resultados.items():  # Recorre cada escenario
        aj = df_esc[df_esc["Escenario"] == n]  # Ajustes del escenario
        txt_aj = "—" if aj.empty else ", ".join(f"{html_lib.escape(nombre_objetivo(str(r['Variable'])))} {num(r['Ajuste %']):+.0f}%" for _, r in aj.iterrows())  # Texto
        filas_esc.append([html_lib.escape(n), txt_aj, dinero_full(raz["VPN proyecto (sin fin.)"]), pct(raz["TIR proyecto (anual)"]),  # Indicadores
                          pct(raz["TIR accionista (anual)"]), fmt_razon("Periodo de recuperación (años)", raz["Periodo de recuperación (años)"])])  # Payback
    tabla_esc = tabla_html(["Escenario", "Ajustes", "VPN proyecto", "TIR proyecto", "TIR accionista", "Recuperación"], filas_esc, numericas=(2, 3, 4, 5))  # Tabla

    kpis = [("VPN proyecto", dinero(rb["VPN proyecto (sin fin.)"]), rb["VPN proyecto (sin fin.)"] > 0),  # Lista de indicadores (nombre, valor, ¿bueno?)
            ("TIR proyecto", pct(rb["TIR proyecto (anual)"]), finito(rb["TIR proyecto (anual)"]) and rb["TIR proyecto (anual)"] * 100 > p["p_tasa_desc"]),  # TIR
            ("VPN accionista", dinero(rb["VPN accionista (con fin.)"]), rb["VPN accionista (con fin.)"] > 0),  # VPN accionista
            ("TIR accionista", pct(rb["TIR accionista (anual)"]), finito(rb["TIR accionista (anual)"]) and rb["TIR accionista (anual)"] * 100 > p["p_ke"]),  # TIR acc.
            ("Inversión total", dinero(rb["Inversión total (CAPEX)"]), None),  # CAPEX
            ("Recuperación", fmt_razon("Periodo de recuperación (años)", rb["Periodo de recuperación (años)"]), None),  # Payback
            ("Beneficio / costo", fmt_razon("Relación beneficio/costo", rb["Relación beneficio/costo"]), None),  # B/C
            ("DSCR mínimo", fmt_razon("DSCR mínimo", rb["DSCR mínimo"]), finito(rb["DSCR mínimo"]) and rb["DSCR mínimo"] >= 1.2 if finito(rb["DSCR mínimo"]) else None)]  # DSCR
    caja_kpis = '<div class="kpis">' + "".join(  # Cuadrícula de indicadores
        f'<div class="k {"" if b is None else ("pos" if b else "neg")}"><div class="l">{n}</div><div class="v">{v}</div></div>' for n, v, b in kpis) + "</div>"  # Cajas

    datos = [("Empresa", esc("p_empresa")), ("Área solicitante", esc("p_area")), ("Solicitante", esc("p_solicitante")),  # Datos del estudio
             ("Responsable", esc("p_responsable")), ("Ubicación", esc("p_ubicacion")), ("Tipo de proyecto", esc("p_tipo")),  # (continuación)
             ("Fecha del estudio", fecha_larga(p["p_fecha_estudio"]))]  # Fecha
    return (  # HTML del resumen
        '<div class="rep">'  # Contenedor
        f'<div class="head"><div class="eyebrow">Resumen ejecutivo · Análisis de rentabilidad</div>'  # Encabezado
        f'<div class="titulo">{esc("p_nombre")}</div><div class="meta">Folio {esc("p_folio")} · {fecha_larga(p["p_fecha_estudio"])} · cifras en {p["p_moneda"]}</div></div>'  # Título
        f'<div class="h2">Resultado</div>{caja_kpis}'  # Indicadores
        f'<div class="veredicto {"ok" if crea_valor else "no"}">{texto_conclusion}</div>'  # Conclusión
        f'<div class="h2">Proyecto</div><p>{esc("p_descripcion")}</p><p><b>Objetivo:</b> {esc("p_objetivo")}</p>'  # Descripción
        f'<div class="cols"><div><div class="h2">Datos del estudio</div>{kv_html(datos)}</div>'  # Datos del estudio
        f'<div><div class="h2">Supuestos generales</div>{kv_html(supuestos_generales(p))}</div></div>'  # Supuestos
        f'<div class="h2">Variables del modelo</div>{tabla_var}'  # Variables
        f'<div class="h2">Conceptos de ingreso y costo</div>{tabla_form}'  # Fórmulas
        f'<div class="h2">Inversiones (CAPEX, moneda original)</div>{tabla_cap}'  # CAPEX
        f'<div class="h2">Resultados anuales · millones de {p["p_moneda"]}</div>{tabla_res}'  # Resultados
        f'<div class="h2">Escenarios</div>{tabla_esc}'  # Escenarios
        f'<div class="pie">Elaborado por {esc("p_responsable")} · {esc("p_area")} · Documento generado el {fecha_larga(date.today())} '  # Pie
        f'con el modelo de Análisis de Rentabilidad de Proyectos.</div></div>')  # (continuación)


def generar_excel():  # Genera un libro de Excel con el resumen, los supuestos y los resultados
    from openpyxl.styles import Alignment, Font, PatternFill  # Estilos de Excel
    buf = io.BytesIO()  # Archivo en memoria
    resumen_xl = ([("Proyecto", params["p_nombre"]), ("Folio", params["p_folio"]), ("Tipo de proyecto", params["p_tipo"]),  # Datos del estudio
                   ("Empresa", params["p_empresa"]), ("Área solicitante", params["p_area"]), ("Solicitante", params["p_solicitante"]),  # (continuación)
                   ("Responsable", params["p_responsable"]), ("Ubicación", params["p_ubicacion"]),  # (continuación)
                   ("Fecha del estudio", fecha_larga(params["p_fecha_estudio"])), ("Descripción", params["p_descripcion"]),  # (continuación)
                   ("Objetivo", params["p_objetivo"]), ("", ""), ("INDICADORES (escenario base)", "")]  # Encabezado de indicadores
                  + [(k, fmt_razon(k, v)) for k, v in rb.items()]  # Indicadores con formato
                  + [("", ""), ("CONCLUSIÓN", texto_conclusion_plano)])  # Conclusión
    anual = anualizar(res_base["m"])  # Resultados anuales
    tabla_xl = construir_tabla(res_base, anual, [etiqueta_anio(a) for a in anual.index]).drop(columns="Tipo")  # Estado de resultados
    with pd.ExcelWriter(buf, engine="openpyxl") as w:  # Escritor de Excel
        pd.DataFrame(resumen_xl, columns=["Concepto", "Valor"]).to_excel(w, sheet_name="Resumen", index=False)  # Hoja de resumen
        pd.DataFrame(supuestos_generales(params), columns=["Supuesto", "Valor"]).to_excel(w, sheet_name="Supuestos generales", index=False)  # Supuestos
        for hoja, df in (("Variables", df_vars), ("Ingresos", df_ing), ("Egresos", df_egr),  # Tablas capturadas
                         ("Gastos corporativos", df_corp), ("CAPEX", df_capex), ("Ajustes de escenarios", df_esc)):  # (continuación)
            df.to_excel(w, sheet_name=hoja, index=False)  # Una hoja por tabla
        tabla_xl.to_excel(w, sheet_name="Resultados anuales", index=False)  # Estado de resultados anual
        pd.DataFrame({n: r[1] for n, r in resultados.items()}).to_excel(w, sheet_name="Comparativo escenarios")  # Comparativo
        if sens is not None:  # Si hay sensibilidad
            sens.to_excel(w, sheet_name=f"Sensibilidad {var_sens}"[:31], index=False)  # Tabla de sensibilidad
        for ws in w.book.worksheets:  # Da formato a cada hoja
            for celda in ws[1]:  # Encabezados
                celda.font = Font(bold=True, color="FFFFFF")  # Letra blanca en negrita
                celda.fill = PatternFill("solid", fgColor=NAVY[1:])  # Fondo azul marino
            for col in ws.columns:  # Ancho de columnas
                largo = max((len(str(c.value)) for c in col if c.value is not None), default=8)  # Texto más largo
                ws.column_dimensions[col[0].column_letter].width = min(largo + 2, 60)  # Ancho limitado
            for fila in ws.iter_rows(min_row=2):  # Formato numérico
                for c in fila:  # Recorre cada celda
                    if isinstance(c.value, float):  # Si es número decimal
                        c.number_format = "#,##0.00"  # Separador de miles con dos decimales
        hoja_res = w.book["Resumen"]  # Hoja de resumen
        hoja_res.column_dimensions["B"].width = 100  # Columna de valores más ancha
        for c in hoja_res["B"]:  # Ajusta el texto largo
            c.alignment = Alignment(wrap_text=True, vertical="top")  # Texto en varias líneas
        for c in hoja_res["A"]:  # Resalta los títulos de sección
            if c.value in ("INDICADORES (escenario base)", "CONCLUSIÓN"):  # Títulos
                c.font = Font(bold=True, color=DORADO[1:])  # Dorado en negrita
    return buf.getvalue()  # Contenido del archivo


with tab_res:  # Contenido del resumen ejecutivo
    cuerpo_resumen = generar_resumen()  # HTML del resumen
    documento = (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'  # Documento completo
                 f'<title>Resumen ejecutivo – {esc("p_nombre")}</title>'  # Título de la pestaña
                 f'<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">'  # Tipografía
                 f'<style>body{{background:#EEF2F7;margin:0;padding:24px;}}{CSS_REPORTE}</style></head><body>{cuerpo_resumen}</body></html>')  # Estilos y contenido
    nombre_arch = re.sub(r"[^A-Za-z0-9]+", "_", str(params["p_nombre"])).strip("_") or "proyecto"  # Nombre de archivo seguro

    seccion("Compartir con el solicitante")  # Título
    c1, c2 = st.columns(2)  # Dos botones
    c1.download_button("Descargar resumen (HTML → PDF)", documento.encode("utf-8"), file_name=f"Resumen_{nombre_arch}.html",  # Resumen HTML
                       mime="text/html", width="stretch", help="Ábrelo en el navegador y usa Imprimir → Guardar como PDF.")  # Ayuda
    c2.download_button("Descargar resumen y detalle (Excel)", generar_excel(), file_name=f"Analisis_{nombre_arch}.xlsx",  # Excel
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")  # Tipo de archivo

    with st.expander("Redactar correo al solicitante", expanded=True):  # Borrador del correo
        correo = str(params["p_correo_sol"] or "").strip()  # Correo del solicitante
        if not correo:  # Si no se capturó correo
            st.info("Captura el correo del solicitante en la pestaña **Carátula** para llenarlo automáticamente.")  # Aviso
        asunto = f"Resultados del análisis de rentabilidad – {params['p_nombre']} ({params['p_folio']})"  # Asunto
        cuerpo_correo = (  # Texto del correo
            f"Estimado(a) {params['p_solicitante']}:\n\n"  # Saludo
            f"Comparto los resultados del análisis de rentabilidad del proyecto «{params['p_nombre']}» (folio {params['p_folio']}).\n\n"  # Introducción
            f"Resultados principales (escenario base, cifras en {params['p_moneda']}):\n"  # Encabezado
            f"• VPN del proyecto: {dinero_full(rb['VPN proyecto (sin fin.)'])} (WACC {params['p_tasa_desc']:.1f}%)\n"  # VPN
            f"• TIR del proyecto: {pct(rb['TIR proyecto (anual)'])}\n"  # TIR
            f"• TIR del accionista: {pct(rb['TIR accionista (anual)'])} (Ke {params['p_ke']:.1f}%)\n"  # TIR accionista
            f"• Inversión total: {dinero_full(rb['Inversión total (CAPEX)'])}\n"  # CAPEX
            f"• Periodo de recuperación: {fmt_razon('Periodo de recuperación (años)', rb['Periodo de recuperación (años)'])}\n\n"  # Payback
            f"Conclusión: {texto_conclusion_plano}\n\n"  # Conclusión
            f"Adjunto el resumen ejecutivo con los supuestos utilizados y el detalle de cálculo en Excel.\n\n"  # Adjuntos
            f"Quedo atento(a) a cualquier comentario.\n\n"  # Cierre
            f"Saludos cordiales,\n{params['p_responsable']}\n{params['p_area']}")  # Firma
        st.text_input("Para", correo, disabled=True)  # Destinatario
        st.text_input("Asunto", asunto, disabled=True)  # Asunto
        mensaje = st.text_area("Mensaje (puedes editarlo)", cuerpo_correo, height=330)  # Mensaje editable
        enlace = (f"mailto:{urllib.parse.quote(correo)}?subject={urllib.parse.quote(asunto)}"  # Enlace de correo
                  f"&body={urllib.parse.quote(mensaje)}")  # Cuerpo del mensaje
        st.link_button("Abrir en mi aplicación de correo", enlace, type="primary")  # Abre Outlook / Gmail / etc.
        st.caption("Los archivos no se adjuntan automáticamente: descarga el HTML/PDF y el Excel y adjúntalos al correo.")  # Aclaración

    seccion("Vista previa del resumen")  # Título
    pintar(cuerpo_resumen)  # Muestra el resumen dentro de la app


# -----------------------------------------------------------------------------
# PESTAÑA DE PROYECTOS: PLANTILLAS PRELLENADAS Y BIBLIOTECA DE PROYECTOS GUARDADOS
# -----------------------------------------------------------------------------
def nombre_archivo(texto):  # Convierte un nombre de proyecto en un nombre de archivo válido
    return re.sub(r"[^A-Za-z0-9áéíóúñÁÉÍÓÚÑ]+", "_", str(texto)).strip("_")[:80] or "proyecto"  # Sin caracteres especiales


def listar_proyectos():  # Lee los proyectos guardados en la carpeta (más recientes primero)
    if not CARPETA_PROYECTOS.exists():  # Si la carpeta aún no existe
        return []  # No hay proyectos
    lista = []  # Lista de (ruta, contenido)
    for ruta in sorted(CARPETA_PROYECTOS.glob("*.json"), key=lambda r: r.stat().st_mtime, reverse=True):  # Recorre los archivos
        try:  # Intenta leer el archivo
            lista.append((ruta, json.loads(ruta.read_text(encoding="utf-8"))))  # Agrega el proyecto
        except (OSError, ValueError):  # Si el archivo está dañado
            continue  # Se ignora
    return lista  # Lista de proyectos


def tarjeta_proyecto(tipo, nombre, descripcion, metricas, pie=""):  # HTML de la tarjeta de un proyecto
    desc = html_lib.escape(str(descripcion or ""))  # Descripción escapada
    desc = desc[:170] + "…" if len(desc) > 170 else desc  # Recorta descripciones largas
    met = "".join(f'<div><div class="k">{k}</div><div class="v {c}">{v}</div></div>' for k, v, c in metricas)  # Indicadores
    return (f'<div class="pcard"><div class="banda"></div><span class="chip">{html_lib.escape(str(tipo or "Proyecto"))}</span>'  # Banda y tipo
            f'<div class="nom">{html_lib.escape(str(nombre))}</div><div class="des">{desc}</div>'  # Nombre y descripción
            f'<div class="met">{met}</div><div class="pie">{pie}</div></div>')  # Indicadores y pie


def tono_vpn(v):  # Clase de color según el VPN
    return "pos" if finito(v) and v > 0 else "neg"  # Verde si es positivo


with tab_proy:  # Contenido de la pestaña de proyectos
    seccion("Plantillas prellenadas · punto de partida para un nuevo análisis")  # Título
    st.caption("Al usar una plantilla se reemplazan **todos** los supuestos actuales. Si necesitas conservar el proyecto abierto, guárdalo primero.")  # Aviso
    cols = st.columns(len(PLANTILLAS))  # Una columna por plantilla
    for col, (clave, paq) in zip(cols, PLANTILLAS.items()):  # Recorre cada plantilla
        p_t = params_de(paq)  # Parámetros de la plantilla
        t_t = tablas_de(paq)  # Tablas de la plantilla
        _, raz_t = evaluar(solo_modelo(p_t), t_t["variables"], t_t["ingresos"], t_t["egresos"],  # Corre la plantilla (resultado en caché)
                           t_t["corporativos"], t_t["capex"])  # (continuación)
        with col.container(border=True):  # Tarjeta con borde
            pintar(tarjeta_proyecto(paq["meta"]["tipo"], paq["meta"]["nombre"], paq["meta"]["descripcion"],  # Tarjeta
                                    [("VPN referencia", dinero(raz_t["VPN proyecto (sin fin.)"]), tono_vpn(raz_t["VPN proyecto (sin fin.)"])),  # VPN
                                     ("TIR", pct(raz_t["TIR proyecto (anual)"]), ""),  # TIR
                                     ("Horizonte", f"{int(p_t['p_dev_meses'])}m + {int(p_t['p_anios'])}a", "")],  # Horizonte
                                    f"Inversión {dinero(raz_t['Inversión total (CAPEX)'])} · cifras en {p_t['p_moneda']}"))  # Pie
            st.button("Usar como base", key=f"tpl_{clave}", type="primary", width="stretch",  # Botón para cargar la plantilla
                      on_click=cargar_paquete, args=(paq, f"Plantilla «{paq['meta']['nombre']}» cargada. Ajusta los supuestos y guárdala como proyecto."))  # Acción

    seccion("Guardar el proyecto actual en la biblioteca")  # Título
    c1, c2 = st.columns([3, 1], vertical_alignment="bottom")  # Campo y botón
    nombre_guardar = c1.text_input("Nombre con el que se guardará", params["p_nombre"])  # Nombre del proyecto a guardar
    ruta_destino = CARPETA_PROYECTOS / f"{nombre_archivo(nombre_guardar)}.json"  # Archivo destino
    existe = ruta_destino.exists()  # ¿Ya hay un proyecto con ese nombre?
    sobrescribir = st.checkbox("Ya existe un proyecto con ese nombre: sobrescribirlo", key="sobrescribir") if existe else True  # Confirmación
    if c2.button("Guardar proyecto", type="primary", width="stretch", disabled=not nombre_guardar.strip() or not sobrescribir):  # Botón guardar
        CARPETA_PROYECTOS.mkdir(parents=True, exist_ok=True)  # Crea la carpeta si no existe
        datos_guardar = dict(paquete)  # Parámetros y tablas actuales
        datos_guardar["meta"] = {  # Datos para la tarjeta del proyecto
            "nombre": nombre_guardar.strip(), "tipo": params["p_tipo"], "descripcion": params["p_descripcion"],  # Identificación
            "folio": params["p_folio"], "guardado": datetime.now().isoformat(timespec="minutes"),  # Folio y fecha
            "moneda": params["p_moneda"], "vpn": rb["VPN proyecto (sin fin.)"], "tir": rb["TIR proyecto (anual)"],  # Resultados
            "capex": rb["Inversión total (CAPEX)"],  # Inversión
        }  # Fin de los metadatos
        ruta_destino.write_text(json.dumps(datos_guardar, ensure_ascii=False, indent=2, default=str), encoding="utf-8")  # Escribe el archivo
        st.success(f"Proyecto «{nombre_guardar.strip()}» guardado.")  # Confirmación
    st.caption(f"Los proyectos se guardan como archivos .json en: `{CARPETA_PROYECTOS}`")  # Ubicación de la biblioteca

    seccion("Mis proyectos guardados")  # Título
    guardados = listar_proyectos()  # Proyectos de la biblioteca
    if not guardados:  # Si no hay proyectos
        st.info("Aún no hay proyectos guardados. Llena un análisis (o parte de una plantilla) y usa **Guardar proyecto**.")  # Aviso
    else:  # Si hay proyectos
        buscar = st.text_input("Buscar", placeholder="Nombre, folio, tipo o descripción", key="buscar_proy").strip().lower()  # Filtro
        visibles = [(r, d) for r, d in guardados  # Proyectos que coinciden con la búsqueda
                    if buscar in (r.stem + " " + json.dumps(d.get("meta", {}), ensure_ascii=False)).lower()]  # Busca en nombre y datos
        st.caption(f"{len(visibles)} de {len(guardados)} proyecto(s)")  # Conteo
        cols = st.columns(3)  # Tres tarjetas por fila
        for i, (ruta, datos) in enumerate(visibles):  # Recorre cada proyecto
            meta = datos.get("meta", {})  # Metadatos del proyecto
            p_g = datos.get("parametros", {})  # Parámetros guardados
            try:  # Formatea la fecha de guardado
                guardado = datetime.fromisoformat(meta.get("guardado", "")).strftime("%d/%m/%Y %H:%M")  # Fecha legible
            except ValueError:  # Si no hay fecha válida
                guardado = "—"  # Sin fecha
            vpn_g, tir_g = num(meta.get("vpn"), np.nan), num(meta.get("tir"), np.nan)  # Resultados al momento de guardar
            with cols[i % 3].container(border=True):  # Tarjeta con borde
                pintar(tarjeta_proyecto(meta.get("tipo") or p_g.get("p_tipo"), meta.get("nombre") or ruta.stem,  # Tarjeta
                                        meta.get("descripcion") or p_g.get("p_descripcion"),  # Descripción
                                        [("VPN", dinero(vpn_g), tono_vpn(vpn_g)), ("TIR", pct(tir_g), ""),  # Indicadores
                                         ("Inversión", dinero(num(meta.get("capex"), np.nan)), "")],  # Inversión
                                        f"Folio {html_lib.escape(str(meta.get('folio') or p_g.get('p_folio') or '—'))} · guardado {guardado}"))  # Pie
                b1, b2, b3 = st.columns(3)  # Botones de la tarjeta
                b1.button("Abrir", key=f"abrir_{ruta.stem}", type="primary", width="stretch",  # Botón para abrir el proyecto
                          on_click=cargar_paquete, args=(datos, f"Proyecto «{meta.get('nombre') or ruta.stem}» abierto."))  # Acción
                b2.download_button("JSON", ruta.read_bytes(), file_name=ruta.name, mime="application/json",  # Descarga del archivo
                                   key=f"dl_proy_{ruta.stem}", width="stretch")  # Botón
                with b3.popover("Eliminar"):  # Confirmación para eliminar
                    st.write(f"¿Eliminar «{meta.get('nombre') or ruta.stem}»? Esta acción no se puede deshacer.")  # Pregunta
                    if st.button("Sí, eliminar", key=f"del_{ruta.stem}", type="primary"):  # Confirmar
                        ruta.unlink(missing_ok=True)  # Borra el archivo
                        st.rerun()  # Actualiza la lista
