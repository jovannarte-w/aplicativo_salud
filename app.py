# -*- coding: utf-8 -*-
"""
ANALÍTICA DE FACTURACIÓN EN SALUD
Facturación, radicación, glosas, recaudo y cartera.

Aplicación Streamlit de un solo archivo para el análisis exploratorio de archivos
de facturación del sector salud. Trabaja exclusivamente con el archivo que carga
el usuario, lo procesa en memoria y no envía información a servicios externos.

Estructura del código:
    1. Configuración, constantes y estilos.
    2. Utilidades de formato y cálculo seguro.
    3. Lectura de archivos (Excel / CSV).
    4. Normalización y mapeo de columnas.
    5. Conversión de tipos, validación de calidad y limpieza.
    6. Variables derivadas y filtros globales.
    7. Funciones analíticas (KPI, cierre, entidades, servicios, glosas, cartera...).
    8. Estadística, correlaciones, pruebas y atípicos.
    9. Insights, alertas y descargas.
   10. Páginas de la interfaz y función main().
"""

import hashlib
import io
import re
import unicodedata
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from scipy import stats

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning)

# =============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# =============================================================================

APP_TITULO = "ANALÍTICA DE FACTURACIÓN EN SALUD"
APP_SUBTITULO = "Facturación, radicación, glosas, recaudo y cartera"

NO_DISPONIBLE = "Columna no disponible"
SIN_INFO = "Sin información"
SIN_DATO = "(Sin dato)"

AVISO_EXPLORATORIO = (
    "Los resultados dependen de la calidad y estructura del archivo cargado. "
    "Esta herramienta facilita el análisis exploratorio y no reemplaza una auditoría "
    "financiera, contable, médica o contractual."
)
AVISO_PRIVACIDAD = (
    "Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, "
    "historias clínicas, diagnósticos u otra información personal o sensible. Utilice "
    "archivos anonimizados y siga las políticas de tratamiento de datos de su organización."
)
AVISO_CALIDAD = (
    "El puntaje es un indicador exploratorio de calidad y no reemplaza una auditoría de datos."
)
AVISO_CORRELACION = (
    "La correlación no implica causalidad. Algunas relaciones pueden surgir directamente "
    "de las fórmulas contables entre las variables."
)

# Paleta institucional (salud y analítica)
COLOR_AZUL_OSCURO = "#0B3C5D"
COLOR_AZUL_MEDIO = "#328CC1"
COLOR_VERDE_AZULADO = "#1B998B"
COLOR_GRIS_CLARO = "#E8EEF2"
COLOR_NARANJA = "#E07A00"
COLOR_ROJO = "#C0392B"
COLOR_VERDE = "#2E8B57"
PALETA = [
    "#0B3C5D", "#328CC1", "#1B998B", "#E07A00", "#6C5B7B", "#8FB8DE",
    "#2E4057", "#5FAD56", "#C0392B", "#A3A9B0", "#4D9DE0", "#7FB7BE",
]
ESCALA_CONTINUA = ["#E8EEF2", "#8FB8DE", "#328CC1", "#0B3C5D"]
ESCALA_ALERTA = ["#E8F5E9", "#FFE0B2", "#E07A00", "#C0392B"]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
MESES_ABREV = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}
DIAS_ES = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes", 5: "Sábado", 6: "Domingo"}

MONEDAS = {"COP": "$ ", "USD": "US$ ", "EUR": "€ ", "Sin símbolo": ""}

# Catálogo de campos esperados: etiqueta, tipo de dato y sinónimos frecuentes.
# Los sinónimos se comparan después de normalizar (sin tildes, minúsculas, sin conectores).
CAMPOS = {
    "ID_Factura": {"etiqueta": "Identificador de factura", "tipo": "id",
                   "sinonimos": ["factura", "no_factura", "num_factura", "numero_factura", "nro_factura",
                                 "n_factura", "factura_id", "id", "consecutivo_factura", "prefijo_factura"]},
    "Numero_Radicado": {"etiqueta": "Número de radicado", "tipo": "id",
                        "sinonimos": ["radicado", "no_radicado", "nro_radicado", "num_radicado", "n_radicado",
                                      "radicacion_numero"]},
    "Fecha_Servicio": {"etiqueta": "Fecha de servicio", "tipo": "fecha",
                       "sinonimos": ["fecha_atencion", "fecha_prestacion", "fecha_procedimiento"]},
    "Fecha_Emision": {"etiqueta": "Fecha de emisión", "tipo": "fecha",
                      "sinonimos": ["fecha_factura", "fecha_facturacion", "fecha_expedicion", "fecha_emision_factura"]},
    "Fecha_Radicacion": {"etiqueta": "Fecha de radicación", "tipo": "fecha",
                         "sinonimos": ["fecha_radicado", "fecha_rad", "fecha_radicacion_factura"]},
    "Fecha_Aceptacion": {"etiqueta": "Fecha de aceptación", "tipo": "fecha",
                         "sinonimos": ["fecha_aceptado"]},
    "Fecha_Pago": {"etiqueta": "Fecha de pago", "tipo": "fecha",
                   "sinonimos": ["fecha_recaudo", "fecha_de_pago"]},
    "Periodo_Cierre": {"etiqueta": "Periodo de cierre", "tipo": "periodo",
                       "sinonimos": ["periodo", "mes_cierre", "periodo_facturacion", "periodo_contable", "cierre"]},
    "Codigo_Entidad": {"etiqueta": "Código de entidad", "tipo": "categoria",
                       "sinonimos": ["cod_entidad", "codigo_eps", "nit_entidad", "codigo_pagador", "nit"]},
    "Entidad_Pagadora": {"etiqueta": "Entidad pagadora", "tipo": "categoria",
                         "sinonimos": ["entidad", "pagador", "eps", "aseguradora", "asegurador", "eapb", "erp",
                                       "cliente", "nombre_entidad", "entidad_responsable_pago"]},
    "Regimen": {"etiqueta": "Régimen", "tipo": "categoria",
                "sinonimos": ["tipo_regimen", "regimen_afiliacion"]},
    "Sede": {"etiqueta": "Sede", "tipo": "categoria",
             "sinonimos": ["sucursal", "sede_atencion", "centro_atencion"]},
    "Codigo_Servicio": {"etiqueta": "Código de servicio", "tipo": "categoria",
                        "sinonimos": ["cod_servicio", "cups", "codigo_cups", "codigo_procedimiento"]},
    "Servicio": {"etiqueta": "Servicio", "tipo": "categoria",
                 "sinonimos": ["descripcion_servicio", "procedimiento", "nombre_servicio", "descripcion_procedimiento"]},
    "Linea_Servicio": {"etiqueta": "Línea de servicio", "tipo": "categoria",
                       "sinonimos": ["linea", "unidad_negocio", "linea_negocio"]},
    "Cantidad": {"etiqueta": "Cantidad", "tipo": "numero",
                 "sinonimos": ["cant", "unidades", "cantidad_servicios"]},
    "Canal_Radicacion": {"etiqueta": "Canal de radicación", "tipo": "categoria",
                         "sinonimos": ["canal", "medio_radicacion", "canal_de_radicacion"]},
    "Estado_Factura": {"etiqueta": "Estado de factura", "tipo": "categoria",
                       "sinonimos": ["estado", "estado_cartera", "estado_actual"]},
    "Dias_Emision_Radicacion": {"etiqueta": "Días emisión-radicación", "tipo": "numero",
                                "sinonimos": ["dias_radicacion", "dias_emision_a_radicacion", "dias_para_radicar",
                                              "tiempo_radicacion"]},
    "Soporte_Completo": {"etiqueta": "Soporte completo", "tipo": "booleano",
                         "sinonimos": ["soportes_completos", "soporte", "soportes", "documentacion_completa"]},
    "Valor_Bruto": {"etiqueta": "Valor bruto", "tipo": "moneda",
                    "sinonimos": ["bruto", "valor_total_bruto", "subtotal", "valor_servicio"]},
    "Descuento": {"etiqueta": "Descuento", "tipo": "moneda",
                  "sinonimos": ["descuentos", "valor_descuento"]},
    "Copago_Cuota": {"etiqueta": "Copago / cuota moderadora", "tipo": "moneda",
                     "sinonimos": ["copago", "cuota_moderadora", "copago_cuota_moderadora", "copagos", "copago_cuotas"]},
    "Valor_Neto": {"etiqueta": "Valor neto", "tipo": "moneda",
                   "sinonimos": ["neto", "total_neto", "valor_factura", "valor_a_pagar", "valor_total"]},
    "Glosa_Inicial": {"etiqueta": "Glosa inicial", "tipo": "moneda",
                      "sinonimos": ["glosa", "valor_glosa", "valor_glosado", "glosas", "glosa_valor"]},
    "Valor_Pagado": {"etiqueta": "Valor pagado", "tipo": "moneda",
                     "sinonimos": ["pagado", "pago", "valor_pago", "recaudo", "valor_recaudado", "pagos"]},
    "Saldo_Cartera": {"etiqueta": "Saldo de cartera", "tipo": "moneda",
                      "sinonimos": ["saldo", "cartera", "saldo_pendiente", "saldo_factura"]},
}

COLUMNAS_MINIMAS = ["ID_Factura", "Fecha_Radicacion o Periodo_Cierre", "Valor_Neto"]
COLUMNAS_RECOMENDADAS = ["Entidad_Pagadora", "Estado_Factura", "Valor_Pagado", "Saldo_Cartera", "Glosa_Inicial"]
CAMPOS_MONEDA = [c for c, v in CAMPOS.items() if v["tipo"] == "moneda"]
CAMPOS_FECHA = [c for c, v in CAMPOS.items() if v["tipo"] == "fecha"]
CAMPOS_CATEGORIA = [c for c, v in CAMPOS.items() if v["tipo"] in ("categoria", "id")]
CAMPOS_NUMERO = [c for c, v in CAMPOS.items() if v["tipo"] == "numero"]

VARIABLES_NUMERICAS_ANALISIS = [
    "Cantidad", "Valor_Bruto", "Descuento", "Copago_Cuota", "Valor_Neto", "Dias_Emision_Radicacion",
    "Dias_Radicacion_Pago", "Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera", "Porcentaje_Glosa",
    "Porcentaje_Recaudo", "Porcentaje_Cartera", "Porcentaje_Descuento", "Porcentaje_Copago",
]
COLUMNAS_PORCENTAJE = ["Porcentaje_Descuento", "Porcentaje_Copago", "Porcentaje_Glosa",
                       "Porcentaje_Recaudo", "Porcentaje_Cartera"]

PAGINAS = [
    "1. Inicio y carga de datos",
    "2. Calidad de datos",
    "3. Resumen ejecutivo",
    "4. Cierre mensual",
    "5. Entidades pagadoras",
    "6. Servicios y líneas",
    "7. Sedes y regímenes",
    "8. Estados de facturación",
    "9. Glosas",
    "10. Recaudo y cartera",
    "11. Tiempos de radicación",
    "12. Análisis estadístico",
    "13. Valores atípicos",
    "14. Insights y alertas",
    "15. Explorador de datos",
    "16. Descargas",
    "17. Metodología",
]
PAGINAS_SIN_MINIMAS = {PAGINAS[0], PAGINAS[1], PAGINAS[14], PAGINAS[15], PAGINAS[16]}

UMBRALES_DEFECTO = {
    "glosa_verde": 5.0, "glosa_amarillo": 10.0,
    "cartera_verde": 30.0, "cartera_amarillo": 50.0,
    "recaudo_verde": 80.0, "recaudo_amarillo": 60.0,
    "dias_verde": 7.0, "dias_amarillo": 15.0,
    "soporte_verde": 5.0, "soporte_amarillo": 10.0,
    "crec_verde": 0.0, "crec_amarillo": -10.0,
}


def configurar_pagina():
    """Configura título, ícono y diseño general de la página."""
    st.set_page_config(
        page_title="Analítica de Facturación en Salud",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def aplicar_estilos():
    """Aplica estilos CSS moderados (tarjetas, encabezados, alertas y tablas)."""
    st.markdown(
        f"""
        <style>
        .encabezado-app {{
            background: linear-gradient(90deg, {COLOR_AZUL_OSCURO} 0%, {COLOR_AZUL_MEDIO} 70%, {COLOR_VERDE_AZULADO} 100%);
            padding: 1.1rem 1.4rem; border-radius: 10px; margin-bottom: 0.8rem;
        }}
        .encabezado-app h1 {{ color: #FFFFFF; font-size: 1.65rem; margin: 0; padding: 0; letter-spacing: 0.5px; }}
        .encabezado-app p {{ color: #EAF4FB; font-size: 1.0rem; margin: 0.2rem 0 0 0; }}
        .kpi-card {{
            background: #FFFFFF; border: 1px solid #D6DEE5; border-left: 5px solid {COLOR_AZUL_MEDIO};
            border-radius: 8px; padding: 0.7rem 0.9rem; margin-bottom: 0.7rem; min-height: 112px;
        }}
        .kpi-card .kpi-titulo {{ color: #44515C; font-size: 0.82rem; font-weight: 600; text-transform: uppercase; }}
        .kpi-card .kpi-valor {{ color: {COLOR_AZUL_OSCURO}; font-size: 1.45rem; font-weight: 700; margin: 0.15rem 0; }}
        .kpi-card .kpi-sub {{ color: #4A5560; font-size: 0.78rem; line-height: 1.25; }}
        .kpi-verde {{ border-left-color: {COLOR_VERDE}; }}
        .kpi-amarillo {{ border-left-color: {COLOR_NARANJA}; }}
        .kpi-rojo {{ border-left-color: {COLOR_ROJO}; }}
        .kpi-gris {{ border-left-color: #8A96A3; }}
        .caja-aviso {{
            background: {COLOR_GRIS_CLARO}; border-left: 4px solid {COLOR_AZUL_OSCURO}; color: #1F2D3A;
            padding: 0.6rem 0.9rem; border-radius: 6px; font-size: 0.88rem; margin: 0.4rem 0 0.8rem 0;
        }}
        .insight-card {{
            background: #FFFFFF; border: 1px solid #D6DEE5; border-radius: 8px; padding: 0.7rem 0.9rem;
            margin-bottom: 0.6rem; color: #1F2D3A;
        }}
        .insight-card b {{ color: {COLOR_AZUL_OSCURO}; }}
        .etiqueta-semaforo {{ font-weight: 700; padding: 0.1rem 0.45rem; border-radius: 4px; }}
        .sem-verde {{ background: #E3F4EA; color: #1E6B41; }}
        .sem-amarillo {{ background: #FFF0DB; color: #8A4B00; }}
        .sem-rojo {{ background: #FBE3E0; color: #8E2318; }}
        .sem-gris {{ background: #ECEFF1; color: #44515C; }}
        h2, h3 {{ color: {COLOR_AZUL_OSCURO}; }}
        div[data-testid="stExpander"] details summary p {{ font-weight: 600; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# 2. UTILIDADES DE FORMATO Y CÁLCULO SEGURO
# =============================================================================

def _simbolo():
    """Devuelve el prefijo visual de moneda configurado en la sesión."""
    return MONEDAS.get(st.session_state.get("moneda", "COP"), "$ ")


def _miles(valor, decimales=0):
    """Formatea un número con punto de miles y coma decimal (convención colombiana)."""
    texto = f"{abs(valor):,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def es_valor_valido(valor):
    """True si el valor es numérico, finito y no nulo."""
    try:
        return valor is not None and not pd.isna(valor) and np.isfinite(float(valor))
    except (TypeError, ValueError):
        return False


def formatear_moneda(valor, simbolo=None):
    """Formatea un valor monetario: $ 1.234.567. Devuelve 'No calculable' si no es válido."""
    if not es_valor_valido(valor):
        return "No calculable"
    prefijo = _simbolo() if simbolo is None else simbolo
    signo = "-" if float(valor) < 0 else ""
    return f"{signo}{prefijo}{_miles(float(valor), 0)}"


def formatear_moneda_corta(valor):
    """Formato abreviado para tarjetas: $ 1.234,5 M."""
    if not es_valor_valido(valor):
        return "No calculable"
    valor = float(valor)
    absoluto = abs(valor)
    signo = "-" if valor < 0 else ""
    if absoluto >= 1e12:
        return f"{signo}{_simbolo()}{_miles(absoluto / 1e12, 2)} B"
    if absoluto >= 1e9:
        return f"{signo}{_simbolo()}{_miles(absoluto / 1e6, 0)} M"
    if absoluto >= 1e6:
        return f"{signo}{_simbolo()}{_miles(absoluto / 1e6, 1)} M"
    return formatear_moneda(valor)


def formatear_porcentaje(valor, decimales=1):
    """Formatea un porcentaje ya expresado en escala 0-100: 12,3 %."""
    if not es_valor_valido(valor):
        return "No calculable"
    signo = "-" if float(valor) < 0 else ""
    return f"{signo}{_miles(float(valor), decimales)} %"


def formatear_numero(valor, decimales=0):
    """Formatea un número con separador de miles."""
    if not es_valor_valido(valor):
        return "No calculable"
    signo = "-" if float(valor) < 0 else ""
    return f"{signo}{_miles(float(valor), decimales)}"


def dividir_seguro(numerador, denominador, escala=1.0):
    """División que devuelve NaN cuando el denominador es cero, nulo o el resultado es infinito.
    Funciona con escalares, Series y arreglos."""
    if isinstance(numerador, (pd.Series, np.ndarray)) or isinstance(denominador, (pd.Series, np.ndarray)):
        num = pd.to_numeric(pd.Series(numerador), errors="coerce") if not isinstance(numerador, pd.Series) \
            else pd.to_numeric(numerador, errors="coerce")
        den = pd.to_numeric(pd.Series(denominador), errors="coerce") if not isinstance(denominador, pd.Series) \
            else pd.to_numeric(denominador, errors="coerce")
        with np.errstate(divide="ignore", invalid="ignore"):
            resultado = num.astype(float) / den.astype(float).where(den != 0) * escala
        return resultado.replace([np.inf, -np.inf], np.nan)
    try:
        if not es_valor_valido(numerador) or not es_valor_valido(denominador) or float(denominador) == 0:
            return np.nan
        resultado = float(numerador) / float(denominador) * escala
        return resultado if np.isfinite(resultado) else np.nan
    except (TypeError, ValueError, ZeroDivisionError):
        return np.nan


def normalizar_texto(texto):
    """Minúsculas, sin tildes y sin espacios extremos (para comparaciones de categorías)."""
    if texto is None or (isinstance(texto, float) and np.isnan(texto)):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto.strip().lower())


def normalizar_serie_texto(serie):
    """Versión vectorizada de normalizar_texto para una Serie."""
    s = serie.astype("string").str.strip().str.lower()
    s = s.str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("ascii")
    return s.str.replace(r"\s+", " ", regex=True)


def columnas_visibles(df):
    """Excluye columnas internas (prefijo '_') de las vistas y descargas."""
    return [c for c in df.columns if not str(c).startswith("_")]


def formatear_tabla(df, columnas_moneda=None, columnas_pct=None, columnas_num=None, columnas_fecha=None, decimales_num=1):
    """Devuelve una copia con valores formateados como texto para presentación.
    Detecta automáticamente columnas monetarias, porcentuales y fechas por su nombre."""
    if df is None or df.empty:
        return df
    salida = df.copy()
    columnas_moneda = set(columnas_moneda or [])
    columnas_pct = set(columnas_pct or [])
    columnas_num = set(columnas_num or [])
    columnas_fecha = set(columnas_fecha or [])
    claves_moneda = ("valor", "bruto", "neto", "descuento", "copago", "glosa", "pagado", "saldo",
                     "ticket", "facturacion", "cartera", "recaudo", "limite")
    for col in salida.columns:
        nombre = normalizar_texto(col)
        serie = salida[col]
        if col in columnas_fecha or pd.api.types.is_datetime64_any_dtype(serie):
            salida[col] = pd.to_datetime(serie, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        elif not pd.api.types.is_numeric_dtype(serie) or pd.api.types.is_bool_dtype(serie):
            continue
        elif col in columnas_num:
            entero = bool(np.all(np.isclose(serie.dropna().astype(float), np.round(serie.dropna().astype(float)))))
            decimales = 0 if entero else decimales_num
            salida[col] = serie.map(lambda v, d=decimales: formatear_numero(v, d) if es_valor_valido(v) else "")
        elif col in columnas_pct or "%" in nombre or "porcentaje" in nombre or "particip" in nombre \
                or nombre.startswith("pct"):
            salida[col] = serie.map(lambda v: formatear_porcentaje(v, 1) if es_valor_valido(v) else "")
        elif col in columnas_moneda or (any(k in nombre for k in claves_moneda) and "dias" not in nombre
                                        and "facturas" not in nombre and "registros" not in nombre
                                        and "con " not in nombre and "cantidad" not in nombre):
            salida[col] = serie.map(lambda v: formatear_moneda(v) if es_valor_valido(v) else "")
        elif col in columnas_num or pd.api.types.is_float_dtype(serie):
            salida[col] = serie.map(lambda v: formatear_numero(v, decimales_num) if es_valor_valido(v) else "")
        else:
            salida[col] = serie.map(lambda v: formatear_numero(v, 0) if es_valor_valido(v) else "")
    return salida


def mostrar_tabla(df, altura=None, **kwargs):
    """Muestra una tabla formateada con ancho adaptable y altura razonable."""
    if df is None or df.empty:
        st.info("No hay registros para mostrar con la selección actual.")
        return
    tabla = formatear_tabla(df, **kwargs)
    if altura is None:
        altura = "auto" if len(tabla) <= 12 else 460
    st.dataframe(tabla, hide_index=True, height=altura)


def estilizar_figura(fig, titulo=None, altura=420, moneda_y=False, pct_y=False, leyenda=True):
    """Aplica estilo coherente a una figura Plotly: separadores colombianos, márgenes y paleta."""
    fig.update_layout(
        title={"text": titulo, "x": 0.0, "font": {"size": 15, "color": COLOR_AZUL_OSCURO}} if titulo else None,
        height=altura,
        separators=",.",
        margin={"l": 10, "r": 10, "t": 50 if titulo else 20, "b": 10},
        colorway=PALETA,
        legend={"orientation": "h", "yanchor": "bottom", "y": -0.28, "x": 0, "title": None},
        showlegend=leyenda,
        hoverlabel={"font_size": 12},
        font={"size": 12},
    )
    # Los periodos AAAA-MM / AAAA-Tn se muestran como categorías ordenadas (evita meses en inglés)
    try:
        for traza in fig.data:
            valores_x = getattr(traza, "x", None)
            if valores_x is not None and len(valores_x) and isinstance(valores_x[0], str) \
                    and re.fullmatch(r"\d{4}(-\d{2}|-T\d)?", valores_x[0]):
                fig.update_xaxes(type="category", categoryorder="category ascending")
                break
    except (TypeError, IndexError):
        pass
    if moneda_y:
        fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f")
    if pct_y:
        fig.update_yaxes(ticksuffix=" %", tickformat=",.1f")
    return fig


def mostrar_grafico(fig, clave=None):
    """Renderiza una figura Plotly dentro de un bloque protegido."""
    try:
        if fig is None:
            st.info("No hay datos suficientes para construir este gráfico.")
            return
        st.plotly_chart(fig, key=clave, config={"displaylogo": False, "locale": "es"})
    except Exception as error:  # noqa: BLE001
        st.warning(f"No fue posible generar el gráfico: {type(error).__name__}.")


def tarjeta_kpi(titulo, valor, subtitulo="", estado="", ayuda=""):
    """Renderiza una tarjeta KPI en HTML. 'estado' admite verde, amarillo, rojo o gris."""
    clase = f"kpi-{estado}" if estado else ""
    titulo_html = f'<span title="{ayuda}">{titulo}</span>' if ayuda else titulo
    st.markdown(
        f'<div class="kpi-card {clase}"><div class="kpi-titulo">{titulo_html}</div>'
        f'<div class="kpi-valor">{valor}</div><div class="kpi-sub">{subtitulo}</div></div>',
        unsafe_allow_html=True,
    )


def caja_aviso(texto):
    """Caja informativa con estilo institucional."""
    st.markdown(f'<div class="caja-aviso">{texto}</div>', unsafe_allow_html=True)


def clave_hash(*partes):
    """Genera una clave corta y estable para caché a partir de varios elementos."""
    texto = "|".join(str(p) for p in partes)
    return hashlib.md5(texto.encode("utf-8")).hexdigest()[:16]


# =============================================================================
# 3. LECTURA DE ARCHIVOS
# =============================================================================

def firma_archivo(nombre, contenido):
    """Firma única del archivo cargado (nombre + tamaño + hash del contenido)."""
    return f"{nombre}|{len(contenido)}|{hashlib.md5(contenido).hexdigest()[:12]}"


def _limpiar_encabezados(df):
    """Convierte encabezados a texto, elimina espacios y resuelve nombres duplicados o vacíos."""
    nuevos, vistos = [], {}
    for i, col in enumerate(df.columns):
        nombre = str(col).strip() if col is not None else ""
        if nombre == "" or nombre.lower().startswith("unnamed:"):
            nombre = f"Columna_{i + 1}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}_{vistos[nombre]}"
        else:
            vistos[nombre] = 0
        nuevos.append(nombre)
    df.columns = nuevos
    return df


@st.cache_data(show_spinner=False, max_entries=8)
def listar_hojas(firma, _contenido, extension):
    """Devuelve las hojas disponibles de un libro Excel. Usa la firma como clave de caché."""
    motor = "xlrd" if extension == "xls" else "openpyxl"
    try:
        libro = pd.ExcelFile(io.BytesIO(_contenido), engine=motor)
        return list(libro.sheet_names), None
    except ImportError:
        return [], "El formato .xls requiere la dependencia 'xlrd'. Guarde el archivo como .xlsx o .csv."
    except Exception as error:  # noqa: BLE001
        return [], f"No fue posible leer el archivo. Verifique que sea un Excel válido ({type(error).__name__})."


@st.cache_data(show_spinner=False, max_entries=8)
def leer_excel(firma, _contenido, extension, hoja):
    """Lee una hoja de Excel en memoria. Devuelve (DataFrame, mensaje_error)."""
    motor = "xlrd" if extension == "xls" else "openpyxl"
    try:
        df = pd.read_excel(io.BytesIO(_contenido), sheet_name=hoja, engine=motor)
    except Exception as error:  # noqa: BLE001
        return None, f"No fue posible leer la hoja '{hoja}' ({type(error).__name__})."
    df = _limpiar_encabezados(df)
    df = df.dropna(axis=1, how="all") if df.shape[1] > 0 else df
    if df.empty:
        return None, "La hoja seleccionada no contiene datos."
    return df, None


def _detectar_codificacion(contenido):
    """Intenta decodificar el CSV con codificaciones habituales en Colombia."""
    for codificacion in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            contenido[:200000].decode(codificacion)
            return codificacion
        except UnicodeDecodeError:
            continue
    return "latin-1"


def _detectar_separador(texto_muestra):
    """Detecta el separador más probable contando ocurrencias consistentes en las primeras líneas."""
    lineas = [l for l in texto_muestra.splitlines()[:30] if l.strip()]
    if not lineas:
        return ","
    mejor, mejor_puntaje = ",", -1
    for sep in (";", ",", "\t", "|"):
        conteos = [l.count(sep) for l in lineas]
        if conteos[0] == 0:
            continue
        consistentes = sum(1 for c in conteos if c == conteos[0])
        puntaje = consistentes * 10 + conteos[0]
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = sep, puntaje
    return mejor


@st.cache_data(show_spinner=False, max_entries=8)
def leer_csv(firma, _contenido, separador="auto", codificacion="auto"):
    """Lee un CSV en memoria detectando codificación y separador. Devuelve (df, error, info)."""
    info = {}
    try:
        cod = _detectar_codificacion(_contenido) if codificacion == "auto" else codificacion
        texto = _contenido.decode(cod, errors="replace")
        sep = _detectar_separador(texto[:50000]) if separador == "auto" else separador
        info = {"codificacion": cod, "separador": sep}
        df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, keep_default_na=True,
                         skip_blank_lines=True, on_bad_lines="warn", engine="python")
    except pd.errors.EmptyDataError:
        return None, "El archivo está vacío.", info
    except Exception as error:  # noqa: BLE001
        return None, f"No fue posible leer el archivo CSV ({type(error).__name__}). Pruebe otro separador.", info
    df = _limpiar_encabezados(df)
    if df.empty or df.shape[1] == 0:
        return None, "El archivo no contiene registros.", info
    if df.shape[1] == 1 and separador == "auto":
        info["advertencia"] = ("Solo se detectó una columna. Es posible que el separador no sea el correcto; "
                               "seleccione otro separador.")
    return df, None, info


# =============================================================================
# 4. NORMALIZACIÓN Y MAPEO DE COLUMNAS
# =============================================================================

_CONECTORES = {"de", "del", "la", "el", "los", "las", "y", "a", "en", "por"}


def normalizar_nombre_columna(nombre):
    """Normaliza un encabezado: sin tildes, minúsculas, espacios y guiones como '_', sin dobles '_'."""
    texto = unicodedata.normalize("NFKD", str(nombre)).encode("ascii", "ignore").decode("ascii")
    texto = texto.strip().lower()
    texto = re.sub(r"[\s\-\.\/]+", "_", texto)
    texto = re.sub(r"[^a-z0-9_]", "", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def clave_compacta(nombre):
    """Clave de comparación: nombre normalizado sin conectores y sin guiones bajos."""
    partes = [p for p in normalizar_nombre_columna(nombre).split("_") if p and p not in _CONECTORES]
    return "".join(partes)


def detectar_columnas(columnas):
    """Asocia automáticamente columnas del archivo con los campos esperados.
    Devuelve (mapeo, detalle, ambiguedades):
      mapeo: {campo: columna_origen o None}
      detalle: lista de equivalencias para mostrar al usuario
      ambiguedades: {campo: [candidatas]} cuando varias columnas coinciden."""
    compactas = {col: clave_compacta(col) for col in columnas}
    usadas, mapeo, detalle, ambiguedades = set(), {}, [], {}

    # Paso 1: coincidencias exactas con el nombre canónico (tras normalizar)
    for campo in CAMPOS:
        objetivo = clave_compacta(campo)
        candidatas = [c for c in columnas if compactas[c] == objetivo and c not in usadas]
        if candidatas:
            mapeo[campo] = candidatas[0]
            usadas.add(candidatas[0])
            tipo = "Exacta" if candidatas[0] == campo else "Normalizada"
            detalle.append({"Columna del archivo": candidatas[0], "Campo esperado": campo, "Coincidencia": tipo})
            if len(candidatas) > 1:
                ambiguedades[campo] = candidatas

    # Paso 2: sinónimos, en orden de prioridad
    for campo, meta in CAMPOS.items():
        if campo in mapeo:
            continue
        candidatas = []
        for sinonimo in meta["sinonimos"]:
            objetivo = clave_compacta(sinonimo)
            for col in columnas:
                if col not in usadas and compactas[col] == objetivo and col not in candidatas:
                    candidatas.append(col)
        if candidatas:
            mapeo[campo] = candidatas[0]
            usadas.add(candidatas[0])
            detalle.append({"Columna del archivo": candidatas[0], "Campo esperado": campo,
                            "Coincidencia": "Sinónimo"})
            if len(candidatas) > 1:
                ambiguedades[campo] = candidatas
        else:
            mapeo[campo] = None
    return mapeo, detalle, ambiguedades


def validar_mapeo(mapeo):
    """Detecta columnas de origen asignadas a más de un campo. Devuelve lista de conflictos."""
    conteo = {}
    for campo, origen in mapeo.items():
        if origen:
            conteo.setdefault(origen, []).append(campo)
    return {origen: campos for origen, campos in conteo.items() if len(campos) > 1}


def mapear_columnas(df, mapeo):
    """Renombra las columnas del archivo a los nombres esperados según el mapeo.
    Si una columna se asignó a varios campos solo se usa en el primero. Las columnas no mapeadas
    se conservan; si su nombre coincide con un campo esperado no asignado se renombran con '_original'."""
    renombres, usadas = {}, set()
    for campo, origen in mapeo.items():
        if origen and origen in df.columns and origen not in usadas:
            renombres[origen] = campo
            usadas.add(origen)
    resultado = df.copy()
    conflictos = {}
    for col in resultado.columns:
        if col not in renombres and col in CAMPOS:
            conflictos[col] = f"{col}_original"
    resultado = resultado.rename(columns={**conflictos, **renombres})
    return resultado


def evaluar_columnas_minimas(df):
    """Verifica columnas mínimas utilizables. Devuelve (ok, faltantes)."""
    faltantes = []
    if not columna_util(df, "ID_Factura"):
        faltantes.append("ID_Factura")
    if not (columna_util(df, "Fecha_Radicacion") or columna_util(df, "Periodo_Cierre")):
        faltantes.append("Fecha_Radicacion o Periodo_Cierre")
    if not columna_util(df, "Valor_Neto"):
        faltantes.append("Valor_Neto")
    return len(faltantes) == 0, faltantes


def columna_util(df, columna):
    """True si la columna existe, tiene el tipo adecuado y al menos un valor no nulo."""
    if df is None or columna not in df.columns:
        return False
    serie = df[columna]
    tipo = CAMPOS.get(columna, {}).get("tipo")
    if tipo in ("moneda", "numero") or columna in VARIABLES_NUMERICAS_ANALISIS:
        if not pd.api.types.is_numeric_dtype(serie):
            return False
    if tipo == "fecha" and not pd.api.types.is_datetime64_any_dtype(serie):
        return False
    return bool(serie.notna().any())


def disponibles(df, columnas):
    """Filtra la lista y devuelve solo las columnas utilizables."""
    return [c for c in columnas if columna_util(df, c)]


# =============================================================================
# 5. CONVERSIÓN DE TIPOS, VALIDACIÓN DE CALIDAD Y LIMPIEZA
# =============================================================================

def convertir_fechas(serie):
    """Convierte una serie a fecha con errors='coerce'. Acepta fechas ISO, día/mes/año y
    números de serie de Excel. Devuelve (serie_convertida, n_no_convertidos)."""
    no_vacios = serie.notna()
    if pd.api.types.is_datetime64_any_dtype(serie):
        convertida = pd.to_datetime(serie, errors="coerce")
        if getattr(convertida.dt, "tz", None) is not None:
            convertida = convertida.dt.tz_localize(None)
        return convertida.astype("datetime64[ns]"), 0
    if pd.api.types.is_numeric_dtype(serie):
        numerica = pd.to_numeric(serie, errors="coerce")
        es_serial = numerica.between(1000, 80000)
        convertida = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
        if es_serial.any():
            convertida.loc[es_serial] = pd.to_datetime(numerica[es_serial], unit="D", origin="1899-12-30",
                                                       errors="coerce")
        fallidos = int((no_vacios & convertida.isna()).sum())
        return convertida, fallidos
    texto = serie.astype("string").str.strip()
    texto = texto.where(texto != "", pd.NA)
    no_vacios = texto.notna()
    convertida = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    # 1) Formato ISO (AAAA-MM-DD con o sin hora)
    iso = texto.str.match(r"^\d{4}-\d{1,2}-\d{1,2}", na=False)
    if iso.any():
        convertida.loc[iso] = pd.to_datetime(texto[iso].str.slice(0, 19), errors="coerce", format="mixed")
    # 2) Resto: formato colombiano día/mes/año
    resto = no_vacios & ~iso
    if resto.any():
        try:
            convertida.loc[resto] = pd.to_datetime(texto[resto], errors="coerce", dayfirst=True, format="mixed")
        except (ValueError, TypeError):
            convertida.loc[resto] = pd.to_datetime(texto[resto], errors="coerce", dayfirst=True)
    # 3) Números de serie de Excel guardados como texto
    pendientes = no_vacios & convertida.isna() & texto.str.fullmatch(r"\d{4,5}(\.0+)?", na=False)
    if pendientes.any():
        numeros = pd.to_numeric(texto[pendientes], errors="coerce")
        convertida.loc[pendientes] = pd.to_datetime(numeros, unit="D", origin="1899-12-30", errors="coerce")
    fallidos = int((no_vacios & convertida.isna()).sum())
    return convertida.astype("datetime64[ns]"), fallidos


def convertir_valores_monetarios(serie):
    """Convierte valores monetarios con símbolos, separadores de miles y coma o punto decimal.
    Determina la convención predominante de la columna (colombiana 1.234,56 o anglosajona 1,234.56).
    Devuelve (serie_numerica, reporte)."""
    reporte = {"no_convertidos": 0, "convencion": "Numérica", "ambiguos": 0, "advertencia": ""}
    if pd.api.types.is_bool_dtype(serie):
        return serie.astype(float), reporte
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").astype(float), reporte

    texto = serie.astype("string").str.strip()
    texto = texto.where(texto != "", pd.NA)
    no_vacios = texto.notna()
    negativo = texto.str.contains(r"^\(.*\)$|^-|-$|^[^\d]*-", regex=True, na=False)
    limpio = texto.str.replace(r"[^\d,\.]", "", regex=True)
    limpio = limpio.where(limpio != "", pd.NA)
    tiene_punto = limpio.str.contains(".", regex=False, na=False)
    tiene_coma = limpio.str.contains(",", regex=False, na=False)
    ambos = tiene_punto & tiene_coma
    ultima_coma = limpio.str.rfind(",") > limpio.str.rfind(".")
    miles_punto = limpio.str.fullmatch(r"\d{1,3}(\.\d{3})+", na=False)
    miles_coma = limpio.str.fullmatch(r"\d{1,3}(,\d{3})+", na=False)
    miles_punto_multi = limpio.str.fullmatch(r"\d{1,3}(\.\d{3}){2,}", na=False)
    miles_coma_multi = limpio.str.fullmatch(r"\d{1,3}(,\d{3}){2,}", na=False)

    evidencia_co = int((ambos & ultima_coma).sum() + (tiene_coma & ~tiene_punto & ~miles_coma).sum()
                       + miles_punto_multi.sum())
    evidencia_us = int((ambos & ~ultima_coma).sum() + (tiene_punto & ~tiene_coma & ~miles_punto).sum()
                       + miles_coma_multi.sum())
    ambiguos = int(((miles_punto & ~miles_punto_multi) | (miles_coma & ~miles_coma_multi)).sum())

    if evidencia_us > evidencia_co:
        convencion = "Anglosajona (1,234.56)"
        normalizado = limpio.str.replace(",", "", regex=False)
    else:
        convencion = "Colombiana (1.234,56)"
        normalizado = limpio.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    numeros = pd.to_numeric(normalizado, errors="coerce").astype(float)
    numeros = numeros.where(~negativo, -numeros)
    reporte["convencion"] = convencion
    reporte["no_convertidos"] = int((no_vacios & numeros.isna()).sum())
    reporte["ambiguos"] = ambiguos
    if evidencia_co > 0 and evidencia_us > 0:
        reporte["advertencia"] = ("La columna mezcla convenciones de separadores decimales; "
                                  f"se aplicó la convención {convencion.lower()}. Revise los valores.")
    elif ambiguos > 0 and evidencia_co == 0 and evidencia_us == 0:
        reporte["advertencia"] = (f"{ambiguos} valores tienen formato ambiguo (por ejemplo 1.234); se "
                                  "interpretaron con la convención colombiana (punto como separador de miles).")
    return numeros, reporte


def normalizar_periodo(serie):
    """Normaliza el periodo de cierre al formato AAAA-MM. Devuelve (serie_texto, n_invalidos)."""
    resultado = pd.Series(pd.NA, index=serie.index, dtype="string")
    no_vacios = serie.notna()
    if pd.api.types.is_datetime64_any_dtype(serie):
        resultado = serie.dt.strftime("%Y-%m").astype("string")
        return resultado, int((no_vacios & resultado.isna()).sum())
    if pd.api.types.is_numeric_dtype(serie):
        numeros = pd.to_numeric(serie, errors="coerce")
        aaaamm = numeros.between(190001, 210012) & (numeros % 100).between(1, 12)
        resultado.loc[aaaamm] = (numeros[aaaamm] // 100).astype(int).astype(str) + "-" + \
            (numeros[aaaamm] % 100).astype(int).astype(str).str.zfill(2)
        serial = numeros.between(1000, 80000) & ~aaaamm
        if serial.any():
            fechas = pd.to_datetime(numeros[serial], unit="D", origin="1899-12-30", errors="coerce")
            resultado.loc[serial] = fechas.dt.strftime("%Y-%m").astype("string")
        return resultado, int((no_vacios & resultado.isna()).sum())

    texto = normalizar_serie_texto(serie)
    texto = texto.where(texto != "", pd.NA)
    no_vacios = texto.notna()
    anio = pd.Series(np.nan, index=serie.index)
    mes = pd.Series(np.nan, index=serie.index)

    patron_am = texto.str.extract(r"^(\d{4})[-/_\. ]?(\d{1,2})$")
    m1 = patron_am[0].notna()
    anio[m1], mes[m1] = pd.to_numeric(patron_am.loc[m1, 0]), pd.to_numeric(patron_am.loc[m1, 1])

    patron_ma = texto.str.extract(r"^(\d{1,2})[-/_\. ](\d{4})$")
    m2 = patron_ma[0].notna() & anio.isna()
    anio[m2], mes[m2] = pd.to_numeric(patron_ma.loc[m2, 1]), pd.to_numeric(patron_ma.loc[m2, 0])

    patron_fecha = texto.str.extract(r"^(\d{4})[-/](\d{1,2})[-/]\d{1,2}")
    m3 = patron_fecha[0].notna() & anio.isna()
    anio[m3], mes[m3] = pd.to_numeric(patron_fecha.loc[m3, 0]), pd.to_numeric(patron_fecha.loc[m3, 1])

    patron_nombre = texto.str.extract(r"^([a-z]{3,})[-/_\. ]+(?:de[ ]+)?(\d{2,4})$")
    m4 = patron_nombre[0].notna() & anio.isna()
    if m4.any():
        meses_txt = patron_nombre.loc[m4, 0].str.slice(0, 3).map(MESES_ABREV)
        anios_txt = pd.to_numeric(patron_nombre.loc[m4, 1], errors="coerce")
        anios_txt = anios_txt.where(anios_txt >= 100, anios_txt + 2000)
        anio[m4], mes[m4] = anios_txt, pd.to_numeric(meses_txt, errors="coerce")

    restantes = no_vacios & anio.isna()
    if restantes.any():
        fechas, _ = convertir_fechas(serie[restantes])
        anio[restantes] = fechas.dt.year
        mes[restantes] = fechas.dt.month

    validos = anio.between(1900, 2100) & mes.between(1, 12)
    resultado.loc[validos] = anio[validos].astype(int).astype(str) + "-" + \
        mes[validos].astype(int).astype(str).str.zfill(2)
    return resultado, int((no_vacios & resultado.isna()).sum())


def normalizar_booleano(serie):
    """Normaliza valores tipo Sí/No a 'Sí' y 'No'. Devuelve (serie, n_no_reconocidos)."""
    verdaderos = {"si", "s", "yes", "y", "true", "verdadero", "1", "1.0", "x", "completo", "completa", "ok"}
    falsos = {"no", "n", "false", "falso", "0", "0.0", "incompleto", "incompleta", "pendiente"}
    if pd.api.types.is_bool_dtype(serie):
        return serie.map({True: "Sí", False: "No"}).astype("string"), 0
    texto = normalizar_serie_texto(serie)
    resultado = pd.Series(pd.NA, index=serie.index, dtype="string")
    resultado[texto.isin(verdaderos)] = "Sí"
    resultado[texto.isin(falsos)] = "No"
    no_reconocidos = int((texto.notna() & (texto != "") & resultado.isna()).sum())
    return resultado, no_reconocidos


def convertir_tipos(df_entrada, convertir_fecha=True, convertir_moneda=True):
    """Convierte fechas, valores monetarios, numéricos, periodo y booleanos de una copia del DataFrame.
    Informa cada conversión. Agrega la columna interna _Fecha_Invalida."""
    df = df_entrada.copy()
    reporte = []
    fecha_invalida = pd.Series(False, index=df.index)

    # Celdas de texto vacías o con solo espacios se tratan como nulas
    for col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_datetime64_any_dtype(df[col]):
            texto = df[col].astype("string").str.strip()
            df[col] = texto.where(texto != "", pd.NA)

    for col in [c for c in CAMPOS_FECHA if c in df.columns]:
        no_vacios = int(df[col].notna().sum())
        if not convertir_fecha and not pd.api.types.is_datetime64_any_dtype(df[col]):
            reporte.append({"Columna": col, "Tipo": "Fecha", "Valores no vacíos": no_vacios,
                            "No convertidos": None, "Nulos resultantes": int(df[col].isna().sum()),
                            "Observación": "Conversión desactivada; la columna no se usará en análisis temporales."})
            continue
        original_no_nulo = df[col].notna()
        try:
            df[col], fallidos = convertir_fechas(df[col])
        except Exception as error:  # noqa: BLE001
            reporte.append({"Columna": col, "Tipo": "Fecha", "Valores no vacíos": no_vacios, "No convertidos": None,
                            "Nulos resultantes": None, "Observación": f"Error de conversión ({type(error).__name__})."})
            continue
        fecha_invalida = fecha_invalida | (original_no_nulo & df[col].isna())
        reporte.append({"Columna": col, "Tipo": "Fecha", "Valores no vacíos": no_vacios, "No convertidos": fallidos,
                        "Nulos resultantes": int(df[col].isna().sum()),
                        "Observación": "Convertida con errors='coerce'." if fallidos == 0 else
                        f"{fallidos} valores no pudieron interpretarse como fecha y quedaron nulos."})

    if "Periodo_Cierre" in df.columns:
        no_vacios = int(df["Periodo_Cierre"].notna().sum())
        df["Periodo_Cierre"], invalidos = normalizar_periodo(df["Periodo_Cierre"])
        reporte.append({"Columna": "Periodo_Cierre", "Tipo": "Periodo AAAA-MM", "Valores no vacíos": no_vacios,
                        "No convertidos": invalidos, "Nulos resultantes": int(df["Periodo_Cierre"].isna().sum()),
                        "Observación": "Formato validado como AAAA-MM." if invalidos == 0 else
                        f"{invalidos} periodos con formato no reconocido quedaron nulos."})

    for col in [c for c in CAMPOS_MONEDA + CAMPOS_NUMERO if c in df.columns]:
        no_vacios = int(df[col].notna().sum())
        if not convertir_moneda and not pd.api.types.is_numeric_dtype(df[col]):
            reporte.append({"Columna": col, "Tipo": "Numérica", "Valores no vacíos": no_vacios,
                            "No convertidos": None, "Nulos resultantes": int(df[col].isna().sum()),
                            "Observación": "Conversión desactivada; la columna no se usará en cálculos."})
            continue
        try:
            df[col], info = convertir_valores_monetarios(df[col])
        except Exception as error:  # noqa: BLE001
            reporte.append({"Columna": col, "Tipo": "Numérica", "Valores no vacíos": no_vacios, "No convertidos": None,
                            "Nulos resultantes": None, "Observación": f"Error de conversión ({type(error).__name__})."})
            continue
        observacion = f"Convención: {info['convencion']}."
        if info["no_convertidos"]:
            observacion += f" {info['no_convertidos']} valores no numéricos quedaron nulos."
        if info["advertencia"]:
            observacion += " " + info["advertencia"]
        reporte.append({"Columna": col, "Tipo": "Monetaria" if col in CAMPOS_MONEDA else "Numérica",
                        "Valores no vacíos": no_vacios, "No convertidos": info["no_convertidos"],
                        "Nulos resultantes": int(df[col].isna().sum()), "Observación": observacion})

    if "Soporte_Completo" in df.columns:
        no_vacios = int(df["Soporte_Completo"].notna().sum())
        df["Soporte_Completo"], no_reconocidos = normalizar_booleano(df["Soporte_Completo"])
        reporte.append({"Columna": "Soporte_Completo", "Tipo": "Sí/No", "Valores no vacíos": no_vacios,
                        "No convertidos": no_reconocidos, "Nulos resultantes": int(df["Soporte_Completo"].isna().sum()),
                        "Observación": "Normalizada a 'Sí' / 'No'." if no_reconocidos == 0 else
                        f"{no_reconocidos} valores no reconocidos como Sí/No quedaron nulos."})

    for col in [c for c in CAMPOS_CATEGORIA if c in df.columns]:
        if not pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].astype("string")
        df[col] = df[col].str.strip()

    df["_Fecha_Invalida"] = fecha_invalida
    return df, pd.DataFrame(reporte)


def calcular_banderas(df, tolerancia=1.0):
    """Calcula banderas de inconsistencia por registro. Solo evalúa reglas cuyas columnas existen."""
    banderas = pd.DataFrame(index=df.index)
    util = lambda c: columna_util(df, c)  # noqa: E731
    neto = df["Valor_Neto"] if util("Valor_Neto") else None

    if util("Fecha_Emision") and util("Fecha_Radicacion"):
        banderas["Radicación anterior a emisión"] = (df["Fecha_Radicacion"] < df["Fecha_Emision"]).fillna(False)
    if util("Fecha_Pago") and util("Fecha_Radicacion"):
        banderas["Pago anterior a radicación"] = (df["Fecha_Pago"] < df["Fecha_Radicacion"]).fillna(False)
    if neto is not None:
        if util("Saldo_Cartera"):
            banderas["Saldo mayor al valor neto"] = (df["Saldo_Cartera"] > neto + tolerancia).fillna(False)
        if util("Valor_Pagado"):
            banderas["Pago mayor al valor neto"] = (df["Valor_Pagado"] > neto + tolerancia).fillna(False)
        if util("Glosa_Inicial"):
            banderas["Glosa mayor al valor neto"] = (df["Glosa_Inicial"] > neto + tolerancia).fillna(False)
        if all(util(c) for c in ("Valor_Bruto", "Descuento", "Copago_Cuota")):
            esperado = df["Valor_Bruto"] - df["Descuento"].fillna(0) - df["Copago_Cuota"].fillna(0)
            banderas["Identidad valor neto"] = ((neto - esperado).abs() > tolerancia).fillna(False) & \
                neto.notna() & df["Valor_Bruto"].notna()
        if all(util(c) for c in ("Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera")):
            esperado = neto - df["Glosa_Inicial"].fillna(0) - df["Valor_Pagado"].fillna(0)
            banderas["Identidad saldo de cartera"] = ((df["Saldo_Cartera"] - esperado).abs() > tolerancia) \
                .fillna(False) & neto.notna() & df["Saldo_Cartera"].notna()
    negativos = pd.Series(False, index=df.index)
    for col in CAMPOS_MONEDA:
        if util(col):
            negativos = negativos | (df[col] < 0).fillna(False)
    banderas["Valor monetario negativo"] = negativos
    return banderas.astype(bool)


BANDERAS_FINANCIERAS = ["Saldo mayor al valor neto", "Pago mayor al valor neto", "Glosa mayor al valor neto",
                        "Identidad valor neto", "Identidad saldo de cartera", "Valor monetario negativo"]


def validar_datos(df_original, df, reporte_conversion, tolerancia=1.0):
    """Diagnóstico de calidad de datos sobre el archivo mapeado y convertido (antes de la limpieza)."""
    n = len(df)
    resultado = {"filas": n, "columnas": df_original.shape[1]}
    campos_encontrados = [c for c in CAMPOS if c in df.columns]
    resultado["encontradas"] = campos_encontrados
    resultado["ausentes"] = [c for c in CAMPOS if c not in df.columns]
    resultado["filas_duplicadas"] = int(df_original.duplicated().sum())
    resultado["ids_duplicados"] = int(df["ID_Factura"].dropna().duplicated().sum()) if "ID_Factura" in df else None
    resultado["radicados_duplicados"] = int(df["Numero_Radicado"].dropna().duplicated().sum()) \
        if "Numero_Radicado" in df else None

    columnas_visibles_df = [c for c in df.columns if not str(c).startswith("_")]
    nulos = df[columnas_visibles_df].isna().sum()
    tabla_nulos = pd.DataFrame({"Columna": nulos.index, "Nulos": nulos.values,
                                "% nulos": dividir_seguro(nulos.values, n, 100).values})
    resultado["tabla_nulos"] = tabla_nulos.sort_values("Nulos", ascending=False)

    rep = reporte_conversion if reporte_conversion is not None else pd.DataFrame()
    if not rep.empty:
        fechas = rep[rep["Tipo"] == "Fecha"]["No convertidos"]
        monedas = rep[rep["Tipo"].isin(["Monetaria", "Numérica"])]["No convertidos"]
        resultado["fechas_invalidas"] = int(pd.to_numeric(fechas, errors="coerce").fillna(0).sum())
        resultado["no_numericos"] = int(pd.to_numeric(monedas, errors="coerce").fillna(0).sum())
    else:
        resultado["fechas_invalidas"], resultado["no_numericos"] = 0, 0

    negativos = {}
    for col in CAMPOS_MONEDA:
        if columna_util(df, col):
            negativos[col] = int((df[col] < 0).sum())
    resultado["negativos"] = negativos
    resultado["neto_cero"] = int((df["Valor_Neto"] == 0).sum()) if columna_util(df, "Valor_Neto") else None
    resultado["sin_entidad"] = int(df["Entidad_Pagadora"].isna().sum()) if "Entidad_Pagadora" in df else None
    if "Periodo_Cierre" in df.columns:
        resultado["sin_periodo"] = int(df["Periodo_Cierre"].isna().sum())
    elif columna_util(df, "Fecha_Radicacion"):
        resultado["sin_periodo"] = int(df["Fecha_Radicacion"].isna().sum())
    else:
        resultado["sin_periodo"] = None

    banderas = calcular_banderas(df, tolerancia)
    resultado["banderas"] = {col: int(banderas[col].sum()) for col in banderas.columns}
    resultado["registros_inconsistentes"] = int(banderas.any(axis=1).sum()) if not banderas.empty else 0

    # ---- Puntaje orientativo de calidad ----
    claves = [c for c in ["ID_Factura", "Fecha_Radicacion", "Periodo_Cierre", "Valor_Neto"] + COLUMNAS_RECOMENDADAS
              if c in df.columns]
    completitud = 100 * (1 - float(df[claves].isna().mean().mean())) if claves and n else 0.0
    ids_dup = resultado["ids_duplicados"] or 0
    unicidad = 100 * (1 - min(1.0, (resultado["filas_duplicadas"] + 0.5 * ids_dup) / n)) if n else 0.0
    celdas_validables = sum(int(df[c].notna().sum()) for c in CAMPOS_MONEDA + CAMPOS_FECHA if c in df.columns)
    errores_validez = resultado["fechas_invalidas"] + resultado["no_numericos"] + sum(negativos.values())
    validez = 100 * (1 - min(1.0, dividir_seguro(errores_validez, celdas_validables + errores_validez) or 0)) \
        if celdas_validables + errores_validez > 0 else 100.0
    consistencia = 100 * (1 - min(1.0, resultado["registros_inconsistentes"] / n)) if n else 0.0
    componentes = {"Completitud": completitud, "Unicidad": unicidad, "Validez": validez, "Consistencia": consistencia}
    resultado["componentes"] = {k: (v if es_valor_valido(v) else 0.0) for k, v in componentes.items()}
    resultado["puntaje"] = float(np.mean(list(resultado["componentes"].values())))
    return resultado


def limpiar_datos(df, opciones):
    """Aplica la limpieza controlada sobre una copia. Devuelve (df_limpio, bitácora)."""
    resultado = df.copy()
    bitacora = []

    def registrar(paso, antes, despues, razon):
        bitacora.append({"Paso": paso, "Registros antes": antes, "Registros excluidos": antes - despues,
                         "Registros después": despues, "Razón": razon})

    if not opciones.get("aplicar", False):
        return resultado, bitacora
    visibles = [c for c in resultado.columns if not str(c).startswith("_")]
    if opciones.get("vacias", False):
        antes = len(resultado)
        resultado = resultado.dropna(how="all", subset=visibles)
        registrar("Eliminar filas vacías", antes, len(resultado), "Todas las columnas vacías.")
    if opciones.get("duplicados", False):
        antes = len(resultado)
        resultado = resultado.drop_duplicates(subset=visibles)
        registrar("Eliminar duplicados exactos", antes, len(resultado), "Registro idéntico a uno anterior.")
    if opciones.get("neto_invalido", False) and "Valor_Neto" in resultado.columns:
        antes = len(resultado)
        if pd.api.types.is_numeric_dtype(resultado["Valor_Neto"]):
            resultado = resultado[resultado["Valor_Neto"].notna()]
        registrar("Excluir Valor_Neto inválido", antes, len(resultado), "Valor neto vacío o no numérico.")
    if opciones.get("fechas_invalidas", False) and "_Fecha_Invalida" in resultado.columns:
        antes = len(resultado)
        resultado = resultado[~resultado["_Fecha_Invalida"]]
        registrar("Excluir fechas inválidas", antes, len(resultado),
                  "Al menos una fecha con valor no interpretable.")
    if opciones.get("rellenar", False):
        rellenadas = 0
        for col in ["Entidad_Pagadora", "Regimen", "Sede", "Servicio", "Linea_Servicio", "Canal_Radicacion",
                    "Estado_Factura", "Codigo_Entidad", "Codigo_Servicio"]:
            if col in resultado.columns:
                rellenadas += int(resultado[col].isna().sum())
                resultado[col] = resultado[col].fillna(SIN_INFO)
        bitacora.append({"Paso": "Rellenar categorías vacías", "Registros antes": len(resultado),
                         "Registros excluidos": 0, "Registros después": len(resultado),
                         "Razón": f"{rellenadas} celdas categóricas vacías reemplazadas por '{SIN_INFO}'."})
    return resultado, bitacora


# =============================================================================
# 6. VARIABLES DERIVADAS Y FILTROS
# =============================================================================

def etiquetas_rangos(limites, unidad="días"):
    """Construye etiquetas para rangos: 0 a 3, 4 a 7, ..., Más de N."""
    etiquetas, inicio = [], 0
    for limite in limites:
        etiquetas.append(f"{inicio} a {int(limite)} {unidad}")
        inicio = int(limite) + 1
    etiquetas.append(f"Más de {int(limites[-1])} {unidad}")
    return etiquetas


def clasificar_rangos(serie, limites, unidad="días"):
    """Clasifica valores numéricos en rangos ordenados (categoría ordenada)."""
    limites = sorted(set(int(x) for x in limites))
    bordes = [-np.inf] + [l + 0.5 for l in limites] + [np.inf]
    etiquetas = etiquetas_rangos(limites, unidad)
    return pd.cut(serie, bins=bordes, labels=etiquetas, right=False, ordered=True)


def crear_variables_derivadas(df_entrada, limites_tiempo=(3, 7, 15), tolerancia=1.0):
    """Crea periodos, variables temporales, porcentajes, banderas y rangos (solo si existen las columnas)."""
    df = df_entrada.copy()
    util = lambda c: columna_util(df, c)  # noqa: E731

    # ---- Periodo de análisis ----
    if "Periodo_Cierre" in df.columns:
        df["Periodo"] = df["Periodo_Cierre"].astype("string")
        df["_Origen_Periodo"] = "Periodo_Cierre"
    elif util("Fecha_Radicacion"):
        df["Periodo"] = df["Fecha_Radicacion"].dt.strftime("%Y-%m").astype("string")
        df["_Origen_Periodo"] = "Fecha_Radicacion"
    if "Periodo" in df.columns:
        df["Periodo_Fecha"] = pd.to_datetime(df["Periodo"].astype("string") + "-01", errors="coerce",
                                             format="%Y-%m-%d")
        df["Año_Periodo"] = df["Periodo_Fecha"].dt.year.astype("Int64")
        df["Numero_Mes_Periodo"] = df["Periodo_Fecha"].dt.month.astype("Int64")
        df["Mes_Periodo"] = df["Numero_Mes_Periodo"].map(MESES_ES).astype("string")
        df["Trimestre_Periodo"] = (df["Año_Periodo"].astype("string") + "-T" +
                                   df["Periodo_Fecha"].dt.quarter.astype("Int64").astype("string"))

    # ---- Variables de la fecha de radicación ----
    if util("Fecha_Radicacion"):
        fr = df["Fecha_Radicacion"]
        df["Año_Radicacion"] = fr.dt.year.astype("Int64")
        df["Numero_Mes"] = fr.dt.month.astype("Int64")
        df["Mes_Radicacion"] = df["Numero_Mes"].map(MESES_ES).astype("string")
        df["Trimestre_Radicacion"] = (df["Año_Radicacion"].astype("string") + "-T" +
                                      fr.dt.quarter.astype("Int64").astype("string"))
        df["Año_Mes"] = fr.dt.strftime("%Y-%m").astype("string")
        df["Día_Semana"] = fr.dt.dayofweek.map(DIAS_ES).astype("string")
        df["_Fecha_Referencia"] = fr
    elif "Periodo_Fecha" in df.columns:
        df["_Fecha_Referencia"] = df["Periodo_Fecha"]

    # ---- Tiempos ----
    if util("Dias_Emision_Radicacion"):
        df["_Origen_Dias"] = "Columna del archivo"
    elif util("Fecha_Emision") and util("Fecha_Radicacion"):
        df["Dias_Emision_Radicacion"] = (df["Fecha_Radicacion"] - df["Fecha_Emision"]).dt.days.astype(float)
        df["_Origen_Dias"] = "Calculado: Fecha_Radicacion - Fecha_Emision"
    if util("Fecha_Pago") and util("Fecha_Radicacion"):
        df["Dias_Radicacion_Pago"] = (df["Fecha_Pago"] - df["Fecha_Radicacion"]).dt.days.astype(float)
    if util("Dias_Emision_Radicacion"):
        df["Rango_Tiempo_Radicacion"] = clasificar_rangos(df["Dias_Emision_Radicacion"], limites_tiempo)

    # ---- Porcentajes por registro ----
    if util("Valor_Bruto"):
        if util("Descuento"):
            df["Porcentaje_Descuento"] = dividir_seguro(df["Descuento"], df["Valor_Bruto"], 100).values
        if util("Copago_Cuota"):
            df["Porcentaje_Copago"] = dividir_seguro(df["Copago_Cuota"], df["Valor_Bruto"], 100).values
    if util("Valor_Neto"):
        if util("Glosa_Inicial"):
            df["Porcentaje_Glosa"] = dividir_seguro(df["Glosa_Inicial"], df["Valor_Neto"], 100).values
        if util("Valor_Pagado"):
            df["Porcentaje_Recaudo"] = dividir_seguro(df["Valor_Pagado"], df["Valor_Neto"], 100).values
        if util("Saldo_Cartera"):
            df["Porcentaje_Cartera"] = dividir_seguro(df["Saldo_Cartera"], df["Valor_Neto"], 100).values

    # ---- Indicadores binarios ----
    if util("Glosa_Inicial"):
        df["Tiene_Glosa"] = (df["Glosa_Inicial"] > 0).fillna(False)
    if util("Saldo_Cartera"):
        df["Tiene_Saldo"] = (df["Saldo_Cartera"] > 0).fillna(False)
    if "Estado_Factura" in df.columns:
        estado_norm = normalizar_serie_texto(df["Estado_Factura"])
        df["Es_Pagada"] = (estado_norm == "pagada").fillna(False)
        df["_Es_Devuelta"] = estado_norm.str.contains("devuel", na=False)
    if "Soporte_Completo" in df.columns:
        df["_Soporte_Incompleto"] = (df["Soporte_Completo"] == "No").astype(float)
        df.loc[df["Soporte_Completo"].isna(), "_Soporte_Incompleto"] = np.nan

    # ---- Banderas de inconsistencia ----
    banderas = calcular_banderas(df, tolerancia)
    financieras = [c for c in BANDERAS_FINANCIERAS if c in banderas.columns]
    df["Inconsistencia_Financiera"] = banderas[financieras].any(axis=1) if financieras else False
    df["_Posible_Inconsistencia"] = banderas.any(axis=1) if not banderas.empty else False
    if not banderas.empty:
        detalle = pd.Series("", index=df.index, dtype="string")
        for col in banderas.columns:
            detalle = detalle.where(~banderas[col], detalle + col + "; ")
        df["Detalle_Inconsistencia"] = detalle.str.rstrip("; ").replace("", pd.NA)
    return df


FILTROS_CATEGORICOS = [
    ("Entidad_Pagadora", "Entidad pagadora"), ("Regimen", "Régimen"), ("Sede", "Sede"),
    ("Servicio", "Servicio"), ("Linea_Servicio", "Línea de servicio"), ("Estado_Factura", "Estado de factura"),
    ("Canal_Radicacion", "Canal de radicación"), ("Soporte_Completo", "Soporte completo"),
]


def opciones_categoria(df, columna):
    """Opciones ordenadas de una columna categórica, incluyendo '(Sin dato)' si existen nulos."""
    serie = df[columna]
    valores = sorted(serie.dropna().astype(str).unique().tolist())
    if serie.isna().any():
        valores.append(SIN_DATO)
    return valores


def aplicar_filtros(df, filtros):
    """Aplica los filtros globales de forma vectorizada. Todas las páginas usan este resultado."""
    if df is None or df.empty:
        return df
    mascara = pd.Series(True, index=df.index)
    rango = filtros.get("fechas")
    if rango and "_Fecha_Referencia" in df.columns:
        inicio = pd.Timestamp(rango[0])
        fin = pd.Timestamp(rango[1]) + pd.Timedelta(days=1)
        mascara &= (df["_Fecha_Referencia"] >= inicio) & (df["_Fecha_Referencia"] < fin)
    if filtros.get("anios") and "Año_Periodo" in df.columns:
        mascara &= df["Año_Periodo"].astype("string").isin([str(a) for a in filtros["anios"]]).fillna(False)
    if filtros.get("meses") and "Mes_Periodo" in df.columns:
        mascara &= df["Mes_Periodo"].isin(filtros["meses"]).fillna(False)
    for columna, _ in FILTROS_CATEGORICOS:
        seleccion = filtros.get(columna)
        if seleccion and columna in df.columns:
            serie = df[columna]
            parcial = serie.astype("string").isin([s for s in seleccion if s != SIN_DATO]).fillna(False)
            if SIN_DATO in seleccion:
                parcial = parcial | serie.isna()
            mascara &= parcial
    glosa = filtros.get("glosa", "Todas")
    if glosa != "Todas" and "Tiene_Glosa" in df.columns:
        mascara &= df["Tiene_Glosa"] if glosa == "Con glosa" else ~df["Tiene_Glosa"]
    saldo = filtros.get("saldo", "Todas")
    if saldo != "Todas" and "Tiene_Saldo" in df.columns:
        mascara &= df["Tiene_Saldo"] if saldo == "Con saldo" else ~df["Tiene_Saldo"]
    rango_neto = filtros.get("neto")
    if rango_neto and columna_util(df, "Valor_Neto"):
        mascara &= df["Valor_Neto"].between(rango_neto[0], rango_neto[1]).fillna(False)
    return df[mascara]


# =============================================================================
# 7. FUNCIONES ANALÍTICAS
# =============================================================================

def _serie_grupo(df, columna):
    """Serie de agrupación con nulos etiquetados como 'Sin información'."""
    serie = df[columna]
    if pd.api.types.is_numeric_dtype(serie) and not pd.api.types.is_bool_dtype(serie):
        return serie.astype("Int64").astype("string").fillna(SIN_INFO).rename(columna)
    return serie.astype("string").fillna(SIN_INFO).rename(columna)


def agregar_por(df, columnas_grupo):
    """Tabla agregada estándar de facturación por una o varias dimensiones.
    Calcula facturas únicas, valores, porcentajes, ticket, tiempos y participación."""
    if isinstance(columnas_grupo, str):
        columnas_grupo = [columnas_grupo]
    columnas_grupo = [c for c in columnas_grupo if c in df.columns]
    if df is None or df.empty or not columnas_grupo:
        return pd.DataFrame()
    util = lambda c: columna_util(df, c)  # noqa: E731
    agregaciones = {}
    if util("ID_Factura"):
        agregaciones["Facturas"] = ("ID_Factura", "nunique")
    agregaciones["Registros"] = (df.columns[0], "size")
    for col, nombre in [("Cantidad", "Cantidad"), ("Valor_Bruto", "Valor bruto"), ("Descuento", "Descuentos"),
                        ("Copago_Cuota", "Copagos"), ("Valor_Neto", "Valor neto"), ("Glosa_Inicial", "Valor glosado"),
                        ("Valor_Pagado", "Valor pagado"), ("Saldo_Cartera", "Saldo de cartera")]:
        if util(col):
            agregaciones[nombre] = (col, "sum")
    if util("Dias_Emision_Radicacion"):
        agregaciones["Días promedio radicación"] = ("Dias_Emision_Radicacion", "mean")
    if "_Soporte_Incompleto" in df.columns:
        agregaciones["% soportes incompletos"] = ("_Soporte_Incompleto", "mean")
    if "Tiene_Glosa" in df.columns:
        agregaciones["Registros con glosa"] = ("Tiene_Glosa", "sum")
    if "Tiene_Saldo" in df.columns:
        agregaciones["Registros con saldo"] = ("Tiene_Saldo", "sum")
    if "_Es_Devuelta" in df.columns:
        agregaciones["Registros devueltos"] = ("_Es_Devuelta", "sum")
    grupos = [_serie_grupo(df, c) for c in columnas_grupo]
    tabla = df.groupby(grupos, sort=False, dropna=False).agg(**agregaciones).reset_index()
    if "% soportes incompletos" in tabla.columns:
        tabla["% soportes incompletos"] = tabla["% soportes incompletos"] * 100
    base_facturas = tabla["Facturas"] if "Facturas" in tabla.columns else tabla["Registros"]
    if "Valor neto" in tabla.columns:
        tabla["Ticket promedio"] = dividir_seguro(tabla["Valor neto"], base_facturas).values
        if "Valor glosado" in tabla.columns:
            tabla["% glosa"] = dividir_seguro(tabla["Valor glosado"], tabla["Valor neto"], 100).values
        if "Valor pagado" in tabla.columns:
            tabla["% recaudo"] = dividir_seguro(tabla["Valor pagado"], tabla["Valor neto"], 100).values
        if "Saldo de cartera" in tabla.columns:
            tabla["% cartera"] = dividir_seguro(tabla["Saldo de cartera"], tabla["Valor neto"], 100).values
        total = tabla["Valor neto"].sum()
        tabla["Participación %"] = dividir_seguro(tabla["Valor neto"], pd.Series([total] * len(tabla)), 100).values
        tabla = tabla.sort_values("Valor neto", ascending=False)
    return tabla.reset_index(drop=True)


def calcular_kpis(df):
    """KPI principales sobre el conjunto recibido. Los no calculables quedan como NaN."""
    util = lambda c: columna_util(df, c)  # noqa: E731
    k = {}
    k["bruto"] = df["Valor_Bruto"].sum() if util("Valor_Bruto") else np.nan
    k["neto"] = df["Valor_Neto"].sum() if util("Valor_Neto") else np.nan
    k["facturas"] = df["ID_Factura"].nunique() if util("ID_Factura") else len(df)
    k["cantidad"] = df["Cantidad"].sum() if util("Cantidad") else np.nan
    k["pagado"] = df["Valor_Pagado"].sum() if util("Valor_Pagado") else np.nan
    k["saldo"] = df["Saldo_Cartera"].sum() if util("Saldo_Cartera") else np.nan
    k["glosa"] = df["Glosa_Inicial"].sum() if util("Glosa_Inicial") else np.nan
    k["descuento"] = df["Descuento"].sum() if util("Descuento") else np.nan
    k["copago"] = df["Copago_Cuota"].sum() if util("Copago_Cuota") else np.nan
    k["pct_recaudo"] = dividir_seguro(k["pagado"], k["neto"], 100)
    k["pct_glosa"] = dividir_seguro(k["glosa"], k["neto"], 100)
    k["pct_cartera"] = dividir_seguro(k["saldo"], k["neto"], 100)
    k["ticket"] = dividir_seguro(k["neto"], k["facturas"])
    k["dias"] = df["Dias_Emision_Radicacion"].mean() if util("Dias_Emision_Radicacion") else np.nan
    k["pct_soporte_incompleto"] = df["_Soporte_Incompleto"].mean() * 100 \
        if "_Soporte_Incompleto" in df.columns and df["_Soporte_Incompleto"].notna().any() else np.nan
    k["registros"] = len(df)
    return k


def _ordinal_periodo(texto, periodicidad="Mensual"):
    """Convierte la etiqueta de periodo en un ordinal para detectar periodos consecutivos."""
    texto = pd.Series(texto).astype("string")
    if periodicidad == "Mensual":
        partes = texto.str.extract(r"^(\d{4})-(\d{2})$").astype(float)
        return partes[0] * 12 + partes[1]
    if periodicidad == "Trimestral":
        partes = texto.str.extract(r"^(\d{4})-T(\d)$").astype(float)
        return partes[0] * 4 + partes[1]
    return pd.to_numeric(texto, errors="coerce")


def periodo_anterior(periodo):
    """Periodo AAAA-MM inmediatamente anterior."""
    try:
        fecha = pd.Timestamp(f"{periodo}-01") - pd.DateOffset(months=1)
        return fecha.strftime("%Y-%m")
    except (ValueError, TypeError):
        return None


def construir_cierre_mensual(df, periodicidad="Mensual"):
    """Tabla de cierre por periodo con crecimiento, acumulado anual y promedio móvil de 3 periodos.
    El crecimiento y el promedio móvil solo se calculan entre periodos consecutivos."""
    columna = {"Mensual": "Periodo", "Trimestral": "Trimestre_Periodo", "Anual": "Año_Periodo"}[periodicidad]
    if columna not in df.columns:
        return pd.DataFrame()
    base = df[df[columna].notna()]
    if base.empty:
        return pd.DataFrame()
    tabla = agregar_por(base, columna).rename(columns={columna: "Periodo"})
    tabla = tabla[tabla["Periodo"] != SIN_INFO].copy()
    tabla["_ordinal"] = _ordinal_periodo(tabla["Periodo"], periodicidad).values
    tabla = tabla.sort_values("_ordinal").reset_index(drop=True)
    consecutivo = tabla["_ordinal"].diff() == 1
    if "Valor neto" in tabla.columns:
        crecimiento = dividir_seguro(tabla["Valor neto"] - tabla["Valor neto"].shift(1),
                                     tabla["Valor neto"].shift(1).abs(), 100)
        tabla["Crecimiento %"] = crecimiento.where(consecutivo).values
        anio = tabla["Periodo"].astype(str).str.slice(0, 4)
        tabla["Acumulado anual"] = tabla.groupby(anio)["Valor neto"].cumsum()
        ventana_valida = (tabla["_ordinal"] - tabla["_ordinal"].shift(2)) == 2
        tabla["Promedio móvil 3"] = tabla["Valor neto"].rolling(3, min_periods=3).mean().where(ventana_valida)
    tabla = tabla.drop(columns=["Participación %"], errors="ignore")
    return tabla


def comparar_ultimo_periodo(df):
    """Compara KPI del último periodo con el inmediatamente anterior, si existe.
    Devuelve (ultimo, anterior, kpis_ultimo, kpis_anterior) o None."""
    if "Periodo" not in df.columns:
        return None
    periodos = sorted(df["Periodo"].dropna().astype(str).unique())
    if len(periodos) < 2:
        return None
    ultimo = periodos[-1]
    anterior = periodo_anterior(ultimo)
    if anterior not in periodos:
        return {"ultimo": ultimo, "anterior": None}
    return {"ultimo": ultimo, "anterior": anterior,
            "k_ultimo": calcular_kpis(df[df["Periodo"] == ultimo]),
            "k_anterior": calcular_kpis(df[df["Periodo"] == anterior])}


def calcular_pareto(df, dimension, valor):
    """Análisis de Pareto: participación individual y acumulada, clase 80 % e índices de concentración."""
    if dimension not in df.columns or not columna_util(df, valor):
        return pd.DataFrame(), {}
    tabla = df.groupby(_serie_grupo(df, dimension), dropna=False)[valor].sum().reset_index()
    tabla.columns = [dimension, "Valor"]
    tabla = tabla[tabla["Valor"] > 0].sort_values("Valor", ascending=False).reset_index(drop=True)
    total = tabla["Valor"].sum()
    if tabla.empty or total <= 0:
        return pd.DataFrame(), {}
    tabla["Participación %"] = tabla["Valor"] / total * 100
    tabla["Participación acumulada %"] = tabla["Participación %"].cumsum()
    anterior = tabla["Participación acumulada %"].shift(1).fillna(0)
    tabla["Grupo Pareto"] = np.where(anterior < 80, "Explica ~80 %", "Resto")
    cuotas = tabla["Participación %"] / 100
    resumen = {
        "n_categorias": len(tabla),
        "n_80": int((tabla["Grupo Pareto"] == "Explica ~80 %").sum()),
        "top1": float(tabla["Participación %"].iloc[0]),
        "top3": float(tabla["Participación %"].head(3).sum()),
        "top5": float(tabla["Participación %"].head(5).sum()),
        "hhi": float((cuotas ** 2).sum() * 10000),
        "lider": str(tabla[dimension].iloc[0]),
    }
    resumen["pct_categorias_80"] = resumen["n_80"] / resumen["n_categorias"] * 100
    return tabla, resumen


def interpretar_hhi(hhi):
    """Lectura orientativa del índice Herfindahl-Hirschman (umbrales de uso común en análisis de mercado)."""
    if not es_valor_valido(hhi):
        return "No calculable"
    if hhi < 1500:
        return "concentración baja"
    if hhi < 2500:
        return "concentración moderada"
    return "concentración alta"


def calcular_antiguedad(df, fecha_corte, limites=(30, 60, 90, 180)):
    """Antigüedad aproximada de la cartera (días desde radicación hasta la fecha de corte)."""
    if not (columna_util(df, "Fecha_Radicacion") and columna_util(df, "Saldo_Cartera")):
        return pd.DataFrame(), pd.DataFrame()
    base = df[(df["Saldo_Cartera"] > 0) & df["Fecha_Radicacion"].notna()].copy()
    if base.empty:
        return pd.DataFrame(), pd.DataFrame()
    base["Días de antigüedad"] = (pd.Timestamp(fecha_corte) - base["Fecha_Radicacion"]).dt.days
    base = base[base["Días de antigüedad"] >= 0]
    base["Rango de antigüedad"] = clasificar_rangos(base["Días de antigüedad"], limites)
    agregaciones = {"Registros": ("Saldo_Cartera", "size"), "Saldo de cartera": ("Saldo_Cartera", "sum")}
    resumen = base.groupby("Rango de antigüedad", observed=False).agg(**agregaciones).reset_index()
    total = resumen["Saldo de cartera"].sum()
    resumen["Participación %"] = dividir_seguro(resumen["Saldo de cartera"],
                                                pd.Series([total] * len(resumen)), 100).values
    resumen["Rango de antigüedad"] = resumen["Rango de antigüedad"].astype(str)
    return resumen, base


def calcular_estadisticos(serie):
    """Estadística descriptiva completa de una variable numérica."""
    datos = pd.to_numeric(serie, errors="coerce").dropna()
    datos = datos[np.isfinite(datos)]
    n = len(datos)
    if n == 0:
        return None
    media = float(datos.mean())
    desviacion = float(datos.std(ddof=1)) if n > 1 else np.nan
    q1, q3 = float(datos.quantile(0.25)), float(datos.quantile(0.75))
    resultado = {
        "Conteo": n, "Media": media, "Mediana": float(datos.median()),
        "Moda": np.nan, "Desviación estándar": desviacion,
        "Varianza": float(datos.var(ddof=1)) if n > 1 else np.nan,
        "Mínimo": float(datos.min()), "Máximo": float(datos.max()),
        "Rango": float(datos.max() - datos.min()), "Cuartil 1": q1, "Cuartil 3": q3,
        "Rango intercuartílico": q3 - q1, "Percentil 90": float(datos.quantile(0.90)),
        "Percentil 95": float(datos.quantile(0.95)),
        "Coeficiente de variación %": dividir_seguro(desviacion, abs(media), 100),
        "Asimetría": float(stats.skew(datos, bias=False)) if n > 2 and desviacion and desviacion > 0 else np.nan,
        "Curtosis (exceso)": float(stats.kurtosis(datos, bias=False)) if n > 3 and desviacion and desviacion > 0
        else np.nan,
    }
    # La moda solo se reporta para variables discretas (enteras con repetición relevante)
    es_entera = bool(np.all(np.isclose(datos, np.round(datos))))
    if es_entera and datos.nunique() <= max(50, n * 0.2):
        modas = datos.mode()
        if not modas.empty:
            resultado["Moda"] = float(modas.iloc[0])
    return resultado


def interpretar_estadisticos(e, nombre):
    """Interpretación textual basada exclusivamente en los estadísticos calculados."""
    if not e:
        return []
    textos = []
    textos.append(f"Se analizaron {formatear_numero(e['Conteo'])} valores válidos de {nombre}.")
    if es_valor_valido(e["Media"]) and es_valor_valido(e["Mediana"]) and e["Mediana"] != 0:
        diferencia = (e["Media"] - e["Mediana"]) / abs(e["Mediana"]) * 100
        if diferencia > 10:
            textos.append(f"La media supera a la mediana en {formatear_porcentaje(diferencia)}: se observa una "
                          "cola hacia valores altos que eleva el promedio.")
        elif diferencia < -10:
            textos.append(f"La media es inferior a la mediana en {formatear_porcentaje(abs(diferencia))}: se "
                          "observa una cola hacia valores bajos.")
        else:
            textos.append("La media y la mediana son cercanas (diferencia menor al 10 %).")
    cv = e.get("Coeficiente de variación %")
    if es_valor_valido(cv):
        nivel = "baja" if cv < 15 else "moderada" if cv < 50 else "alta"
        textos.append(f"El coeficiente de variación es {formatear_porcentaje(cv)}, lo que indica dispersión {nivel} "
                      "respecto a la media.")
    asim = e.get("Asimetría")
    if es_valor_valido(asim):
        if abs(asim) < 0.5:
            textos.append(f"La asimetría ({formatear_numero(asim, 2)}) es baja.")
        else:
            lado = "derecha (valores altos)" if asim > 0 else "izquierda (valores bajos)"
            textos.append(f"La asimetría ({formatear_numero(asim, 2)}) indica sesgo hacia la {lado}.")
    curt = e.get("Curtosis (exceso)")
    if es_valor_valido(curt) and curt > 1:
        textos.append(f"La curtosis en exceso ({formatear_numero(curt, 2)}) sugiere colas más pesadas que una "
                      "distribución normal de referencia; conviene revisar valores extremos.")
    textos.append("Esta descripción no evalúa la normalidad. Para ello utilice la prueba de Shapiro-Wilk en la "
                  "pestaña de pruebas estadísticas.")
    return textos


def calcular_correlaciones(df, variables, metodo="pearson"):
    """Matriz de correlación entre variables numéricas disponibles."""
    datos = df[variables].apply(pd.to_numeric, errors="coerce")
    datos = datos.loc[:, datos.notna().sum() >= 3]
    datos = datos.loc[:, datos.std(numeric_only=True) > 0]
    if datos.shape[1] < 2:
        return pd.DataFrame()
    return datos.corr(method=metodo, min_periods=3)


def clasificar_correlaciones(matriz, umbral_fuerte=0.7, umbral_debil=0.3):
    """Lista de pares con su coeficiente y clasificación (asociación, no causalidad)."""
    filas = []
    columnas = list(matriz.columns)
    for i, a in enumerate(columnas):
        for b in columnas[i + 1:]:
            r = matriz.loc[a, b]
            if not es_valor_valido(r):
                continue
            if r >= umbral_fuerte:
                clase = "Positiva fuerte"
            elif r <= -umbral_fuerte:
                clase = "Negativa fuerte"
            elif abs(r) < umbral_debil:
                clase = "Débil"
            else:
                clase = "Moderada positiva" if r > 0 else "Moderada negativa"
            filas.append({"Variable 1": a, "Variable 2": b, "Coeficiente": round(float(r), 3),
                          "Clasificación": clase})
    tabla = pd.DataFrame(filas)
    if not tabla.empty:
        tabla = tabla.reindex(tabla["Coeficiente"].abs().sort_values(ascending=False).index)
    return tabla


def detectar_atipicos(df, variable, metodo="IQR", multiplicador=1.5, umbral_z=3.0,
                      percentiles=(1.0, 99.0), alcance=None):
    """Detecta valores atípicos de forma global o por grupo. Devuelve (máscara, tabla_límites)."""
    valores = pd.to_numeric(df[variable], errors="coerce")
    if alcance and alcance in df.columns:
        grupos = _serie_grupo(df, alcance)
    else:
        grupos = pd.Series("Global", index=df.index)
    agrupado = valores.groupby(grupos)
    def por_grupo(estadistico):
        """Expande un estadístico por grupo a cada registro (vectorizado)."""
        return grupos.map(estadistico).astype(float)

    if metodo == "IQR":
        q1, q3 = por_grupo(agrupado.quantile(0.25)), por_grupo(agrupado.quantile(0.75))
        iqr = q3 - q1
        inferior, superior = q1 - multiplicador * iqr, q3 + multiplicador * iqr
    elif metodo == "Puntaje Z":
        media, desviacion = por_grupo(agrupado.mean()), por_grupo(agrupado.std())
        inferior = media - umbral_z * desviacion
        superior = media + umbral_z * desviacion
    else:
        inferior = por_grupo(agrupado.quantile(percentiles[0] / 100))
        superior = por_grupo(agrupado.quantile(percentiles[1] / 100))
    conteo = por_grupo(agrupado.count())
    mascara = ((valores < inferior) | (valores > superior)).fillna(False) & (conteo >= 4)
    limites = pd.DataFrame({"Grupo": grupos, "Límite inferior": inferior, "Límite superior": superior,
                            "Registros válidos": conteo, "Atípicos": mascara})
    limites = limites.groupby("Grupo").agg({"Límite inferior": "first", "Límite superior": "first",
                                            "Registros válidos": "first", "Atípicos": "sum"}).reset_index()
    limites["% atípicos"] = dividir_seguro(limites["Atípicos"], limites["Registros válidos"], 100).values
    return mascara, limites.sort_values("Atípicos", ascending=False)


def ejecutar_prueba_estadistica(tipo, datos, alfa=0.05, **kwargs):
    """Ejecuta pruebas de hipótesis con scipy y devuelve un diccionario con resultados e interpretación."""
    resultado = {"alfa": alfa}
    if tipo == "dos_grupos":
        g1, g2 = [pd.to_numeric(g, errors="coerce").dropna() for g in datos]
        if len(g1) < 3 or len(g2) < 3:
            return {"error": "Cada grupo debe tener al menos 3 observaciones válidas."}
        prueba = kwargs.get("prueba", "Automática")
        if prueba == "Automática":
            normal = all(stats.shapiro(g.sample(min(len(g), 5000), random_state=42))[1] > alfa
                         for g in (g1, g2) if len(g) >= 3)
            prueba = "t de Student (Welch)" if normal else "Mann-Whitney U"
            resultado["seleccion"] = ("Se eligió t de Student (Welch) porque Shapiro-Wilk no rechazó normalidad en "
                                      "ambos grupos." if normal else
                                      "Se eligió Mann-Whitney U porque al menos un grupo no muestra evidencia "
                                      "compatible con normalidad (Shapiro-Wilk).")
        if prueba.startswith("t de Student"):
            estadistico, p = stats.ttest_ind(g1, g2, equal_var=False)
            resultado.update({"prueba": "t de Student (Welch, varianzas no iguales)",
                              "h0": "Las medias de ambos grupos son iguales.",
                              "h1": "Las medias de ambos grupos son diferentes."})
        else:
            estadistico, p = stats.mannwhitneyu(g1, g2, alternative="two-sided")
            resultado.update({"prueba": "Mann-Whitney U",
                              "h0": "Las distribuciones de ambos grupos son iguales (sin desplazamiento).",
                              "h1": "Las distribuciones de ambos grupos difieren."})
        resultado["muestras"] = {"Grupo 1": len(g1), "Grupo 2": len(g2)}
    elif tipo == "varios_grupos":
        grupos = [pd.to_numeric(g, errors="coerce").dropna() for g in datos]
        grupos = [g for g in grupos if len(g) >= 3]
        if len(grupos) < 3:
            return {"error": "No existen suficientes grupos para realizar la prueba (mínimo 3 con 3 o más datos)."}
        prueba = kwargs.get("prueba", "Automática")
        if prueba == "Automática":
            normal = all(stats.shapiro(g.sample(min(len(g), 5000), random_state=42))[1] > alfa for g in grupos)
            homogenea = stats.levene(*grupos)[1] > alfa
            prueba = "ANOVA de una vía" if normal and homogenea else "Kruskal-Wallis"
            resultado["seleccion"] = (f"Se eligió {prueba}: normalidad por grupo "
                                      f"{'no rechazada' if normal else 'rechazada en al menos un grupo'} y "
                                      f"homogeneidad de varianzas (Levene) {'no rechazada' if homogenea else 'rechazada'}.")
        if prueba == "ANOVA de una vía":
            estadistico, p = stats.f_oneway(*grupos)
            resultado.update({"prueba": "ANOVA de una vía", "h0": "Las medias de todos los grupos son iguales.",
                              "h1": "Al menos un grupo tiene una media diferente."})
        else:
            estadistico, p = stats.kruskal(*grupos)
            resultado.update({"prueba": "Kruskal-Wallis",
                              "h0": "Las distribuciones de todos los grupos son iguales.",
                              "h1": "Al menos un grupo tiene una distribución diferente."})
        resultado["muestras"] = {f"Grupo {i + 1}": len(g) for i, g in enumerate(grupos)}
    elif tipo == "chi_cuadrado":
        tabla = datos
        if tabla.shape[0] < 2 or tabla.shape[1] < 2:
            return {"error": "Se requieren al menos dos categorías en cada variable."}
        estadistico, p, gl, esperados = stats.chi2_contingency(tabla)
        n = tabla.values.sum()
        v_cramer = np.sqrt(estadistico / (n * (min(tabla.shape) - 1))) if n > 0 else np.nan
        pct_bajos = float((esperados < 5).mean() * 100)
        resultado.update({"prueba": "Chi-cuadrado de independencia",
                          "h0": "Las dos variables son independientes.",
                          "h1": "Existe asociación entre las dos variables.",
                          "muestras": {"Observaciones": int(n), "Grados de libertad": int(gl)},
                          "extra": f"V de Cramér: {formatear_numero(v_cramer, 3)} (magnitud de la asociación)."})
        if pct_bajos > 20:
            resultado["advertencia"] = (f"El {formatear_porcentaje(pct_bajos)} de las celdas tiene frecuencia esperada "
                                        "menor a 5; el resultado puede no ser confiable.")
    elif tipo == "normalidad":
        serie = pd.to_numeric(datos, errors="coerce").dropna()
        if len(serie) < 3:
            return {"error": "La variable seleccionada no contiene suficientes datos numéricos."}
        if serie.std() == 0:
            return {"error": "La variable no tiene variabilidad; no es posible evaluar normalidad."}
        muestra = serie.sample(min(len(serie), 5000), random_state=42)
        estadistico, p = stats.shapiro(muestra)
        resultado.update({"prueba": "Shapiro-Wilk", "h0": "Los datos provienen de una distribución normal.",
                          "h1": "Los datos no provienen de una distribución normal.",
                          "muestras": {"Observaciones evaluadas": len(muestra), "Observaciones totales": len(serie)}})
        if len(serie) > 5000:
            resultado["advertencia"] = ("El conjunto supera 5.000 registros; se evaluó una muestra aleatoria "
                                        "reproducible de 5.000 (semilla 42). Con muestras grandes la prueba detecta "
                                        "desviaciones pequeñas.")
    else:
        return {"error": "Prueba no reconocida."}
    resultado["estadistico"] = float(estadistico)
    resultado["p"] = float(p)
    rechaza = p < alfa
    resultado["decision"] = "Se rechaza la hipótesis nula." if rechaza else "No se rechaza la hipótesis nula."
    if rechaza:
        resultado["interpretacion"] = (f"Con un nivel de significancia de {formatear_numero(alfa, 2)}, los datos "
                                       "aportan evidencia estadística a favor de la hipótesis alternativa. Esto indica "
                                       "una diferencia o asociación estadística, no una relación de causa y efecto.")
    else:
        resultado["interpretacion"] = (f"Con un nivel de significancia de {formatear_numero(alfa, 2)}, los datos no "
                                       "aportan evidencia suficiente para rechazar la hipótesis nula. Esto no "
                                       "demuestra que la hipótesis nula sea verdadera.")
    return resultado


def evaluar_semaforo(valor, verde, amarillo, mayor_es_mejor=False):
    """Clasifica un indicador en Verde / Amarillo / Rojo según umbrales definidos por el usuario."""
    if not es_valor_valido(valor):
        return "gris", "Sin dato"
    if mayor_es_mejor:
        if valor >= verde:
            return "verde", "Verde"
        if valor >= amarillo:
            return "amarillo", "Amarillo"
        return "rojo", "Rojo"
    if valor <= verde:
        return "verde", "Verde"
    if valor <= amarillo:
        return "amarillo", "Amarillo"
    return "rojo", "Rojo"


ICONO_SEMAFORO = {"verde": "🟢", "amarillo": "🟡", "rojo": "🔴", "gris": "⚪"}


# =============================================================================
# 9. INSIGHTS, ALERTAS Y DESCARGAS
# =============================================================================

def _insight(categoria, hallazgo, valor, contexto, implicacion, accion):
    """Estructura estándar de un insight: hallazgo calculado, interpretación y recomendación."""
    return {"Categoría": categoria, "Hallazgo calculado": hallazgo, "Valor observado": valor,
            "Contexto": contexto, "Interpretación": implicacion, "Recomendación para investigar": accion}


def _grupo_extremo(tabla, columna_valor, minimo_registros=5, mayor=True):
    """Fila con el valor máximo o mínimo, considerando solo grupos con suficientes registros."""
    if tabla.empty or columna_valor not in tabla.columns:
        return None
    base = tabla[(tabla["Registros"] >= minimo_registros) & tabla[columna_valor].notna()]
    base = base[base.iloc[:, 0] != SIN_INFO]
    if len(base) < 2:
        return None
    return base.loc[base[columna_valor].idxmax() if mayor else base[columna_valor].idxmin()]


def generar_insights(df, calidad=None):
    """Genera insights mediante reglas programadas sobre el DataFrame filtrado. No usa modelos externos."""
    insights = []
    if df is None or df.empty or not columna_util(df, "Valor_Neto"):
        return insights
    k = calcular_kpis(df)

    # ---- Tendencia y periodos ----
    cierre = construir_cierre_mensual(df, "Mensual")
    if len(cierre) >= 2:
        fila_max = cierre.loc[cierre["Valor neto"].idxmax()]
        insights.append(_insight(
            "Facturación", f"El periodo con mayor facturación neta es {fila_max['Periodo']}.",
            formatear_moneda(fila_max["Valor neto"]),
            f"Promedio mensual del conjunto: {formatear_moneda(cierre['Valor neto'].mean())} en {len(cierre)} periodos.",
            "Se observa un pico de facturación frente al promedio del periodo analizado.",
            "Conviene revisar si el pico corresponde a mayor actividad, a radicación acumulada o a cambios de tarifa."))
        crec = cierre["Crecimiento %"].dropna() if "Crecimiento %" in cierre else pd.Series(dtype=float)
        if not crec.empty and crec.min() < 0:
            fila = cierre.loc[crec.idxmin()]
            insights.append(_insight(
                "Facturación", f"La mayor caída mensual de facturación neta se presentó en {fila['Periodo']}.",
                formatear_porcentaje(fila["Crecimiento %"]),
                "Variación frente al periodo inmediatamente anterior (solo periodos consecutivos).",
                "Los datos muestran una reducción relevante de la facturación en ese periodo.",
                "Este resultado podría justificar una revisión de la producción, la radicación y el cierre de ese mes."))
        if len(cierre) >= 2 and "Crecimiento %" in cierre and es_valor_valido(cierre["Crecimiento %"].iloc[-1]):
            ultimo = cierre.iloc[-1]
            insights.append(_insight(
                "Facturación", f"Variación del último periodo ({ultimo['Periodo']}) frente al anterior.",
                formatear_porcentaje(ultimo["Crecimiento %"]),
                f"Valor neto del último periodo: {formatear_moneda(ultimo['Valor neto'])}.",
                "Se observa " + ("crecimiento" if ultimo["Crecimiento %"] >= 0 else "disminución") +
                " de la facturación en el último periodo disponible. El último periodo puede estar incompleto.",
                "Conviene confirmar que el último periodo esté cerrado antes de concluir."))
    if len(cierre) >= 4:
        try:
            x = np.arange(len(cierre))
            regresion = stats.linregress(x, cierre["Valor neto"].astype(float))
            if regresion.pvalue < 0.05:
                sentido = "creciente" if regresion.slope > 0 else "decreciente"
                insights.append(_insight(
                    "Tendencia", f"La facturación neta mensual muestra una tendencia lineal {sentido}.",
                    f"{formatear_moneda(regresion.slope)} por periodo (p = {formatear_numero(regresion.pvalue, 4)})",
                    f"Regresión lineal simple sobre {len(cierre)} periodos ordenados.",
                    "Se observa una tendencia estadísticamente significativa; la información disponible no permite "
                    "establecer la causa.", "Conviene contrastar la tendencia con la capacidad instalada y los contratos."))
            else:
                insights.append(_insight(
                    "Tendencia", "No se identifica una tendencia lineal estadísticamente clara en la facturación.",
                    f"p = {formatear_numero(regresion.pvalue, 4)}",
                    f"Regresión lineal simple sobre {len(cierre)} periodos.",
                    "Las variaciones observadas podrían corresponder a fluctuaciones del periodo.",
                    "Conviene analizar la estacionalidad con una serie más larga."))
        except (ValueError, FloatingPointError):
            pass

    # ---- Entidades ----
    if "Entidad_Pagadora" in df.columns:
        entidades = agregar_por(df, "Entidad_Pagadora")
        if not entidades.empty and len(entidades) >= 2:
            lider = entidades.iloc[0]
            insights.append(_insight(
                "Entidades", f"{lider['Entidad_Pagadora']} es la entidad con mayor facturación neta.",
                f"{formatear_moneda(lider['Valor neto'])} ({formatear_porcentaje(lider['Participación %'])})",
                f"Total de entidades en el conjunto: {len(entidades)}.",
                "Los datos muestran dependencia relevante de este pagador en los ingresos facturados.",
                "Conviene revisar condiciones contractuales y riesgo de concentración con esta entidad."))
            if "Saldo de cartera" in entidades:
                fila = entidades.loc[entidades["Saldo de cartera"].idxmax()]
                part = dividir_seguro(fila["Saldo de cartera"], entidades["Saldo de cartera"].sum(), 100)
                insights.append(_insight(
                    "Cartera", f"{fila['Entidad_Pagadora']} concentra el mayor saldo de cartera.",
                    f"{formatear_moneda(fila['Saldo de cartera'])} ({formatear_porcentaje(part)} de la cartera)",
                    f"Porcentaje de cartera de la entidad: {formatear_porcentaje(fila.get('% cartera'))}.",
                    "Se observa una acumulación de saldo pendiente en este pagador.",
                    "Conviene revisar el estado de conciliación y los compromisos de pago con la entidad."))
            fila = _grupo_extremo(entidades, "% recaudo", mayor=False)
            if fila is not None:
                insights.append(_insight(
                    "Recaudo", f"{fila['Entidad_Pagadora']} presenta el menor porcentaje de recaudo.",
                    formatear_porcentaje(fila["% recaudo"]),
                    f"Recaudo global del conjunto: {formatear_porcentaje(k['pct_recaudo'])}. "
                    "Solo se comparan entidades con 5 o más registros.",
                    "Los datos muestran un recaudo inferior al de las demás entidades; la información disponible "
                    "no permite establecer la causa.", "Este resultado podría justificar una revisión de la gestión "
                    "de cobro y del estado de las facturas de esta entidad."))
            fila = _grupo_extremo(entidades, "% glosa", mayor=True)
            if fila is not None and fila["% glosa"] > 0:
                insights.append(_insight(
                    "Glosas", f"{fila['Entidad_Pagadora']} presenta el mayor porcentaje de glosa.",
                    formatear_porcentaje(fila["% glosa"]),
                    f"Porcentaje de glosa global: {formatear_porcentaje(k['pct_glosa'])}.",
                    "Se observa una severidad de glosa superior al promedio.",
                    "Conviene revisar las causales de glosa y los soportes enviados a esta entidad."))
        pareto, resumen = calcular_pareto(df, "Entidad_Pagadora", "Valor_Neto")
        if resumen and resumen["n_categorias"] >= 3:
            insights.append(_insight(
                "Concentración", f"{resumen['n_80']} de {resumen['n_categorias']} entidades explican cerca del "
                "80 % de la facturación neta.",
                f"Top 3: {formatear_porcentaje(resumen['top3'])} | HHI: {formatear_numero(resumen['hhi'])}",
                f"Lectura orientativa del índice HHI: {interpretar_hhi(resumen['hhi'])}.",
                "La concentración de ingresos en pocos pagadores aumenta la exposición a cambios en esos contratos.",
                "Conviene evaluar la diversificación de pagadores. El 80 % es una referencia analítica de Pareto."))
        if columna_util(df, "Saldo_Cartera"):
            _, resumen_c = calcular_pareto(df, "Entidad_Pagadora", "Saldo_Cartera")
            if resumen_c and resumen_c["n_categorias"] >= 3:
                insights.append(_insight(
                    "Concentración", f"{resumen_c['n_80']} de {resumen_c['n_categorias']} entidades concentran "
                    "cerca del 80 % de la cartera.",
                    f"Top 3: {formatear_porcentaje(resumen_c['top3'])} | HHI: {formatear_numero(resumen_c['hhi'])}",
                    f"Entidad con mayor cartera: {resumen_c['lider']}.",
                    "La cartera está concentrada; la gestión de pocos pagadores define gran parte del saldo.",
                    "Conviene priorizar la gestión de cobro en las entidades del grupo de concentración."))

    # ---- Servicios ----
    if "Servicio" in df.columns:
        servicios = agregar_por(df, "Servicio")
        if len(servicios) >= 2:
            fila = servicios.iloc[0]
            insights.append(_insight(
                "Servicios", f"{fila['Servicio']} es el servicio con mayor facturación neta.",
                f"{formatear_moneda(fila['Valor neto'])} ({formatear_porcentaje(fila['Participación %'])})",
                f"Servicios distintos en el conjunto: {len(servicios)}.",
                "Se observa que este servicio aporta la mayor proporción de la facturación.",
                "Conviene monitorear su tarifa, volumen y glosa de forma prioritaria."))
            if "Saldo de cartera" in servicios:
                fila = servicios.loc[servicios["Saldo de cartera"].idxmax()]
                insights.append(_insight(
                    "Servicios", f"{fila['Servicio']} es el servicio con mayor saldo de cartera.",
                    formatear_moneda(fila["Saldo de cartera"]),
                    f"Porcentaje de cartera del servicio: {formatear_porcentaje(fila.get('% cartera'))}.",
                    "Los datos muestran saldo pendiente concentrado en este servicio.",
                    "Conviene revisar facturas pendientes y glosas asociadas a este servicio."))

    # ---- Tiempos ----
    if columna_util(df, "Dias_Emision_Radicacion"):
        for dimension, nombre in [("Sede", "La sede"), ("Canal_Radicacion", "El canal de radicación")]:
            if dimension in df.columns:
                tabla = agregar_por(df, dimension)
                fila = _grupo_extremo(tabla, "Días promedio radicación", mayor=True)
                if fila is not None:
                    insights.append(_insight(
                        "Tiempos", f"{nombre} con mayor tiempo promedio de radicación es {fila[dimension]}.",
                        f"{formatear_numero(fila['Días promedio radicación'], 1)} días",
                        f"Promedio global: {formatear_numero(k['dias'], 1)} días.",
                        "Se observa un tiempo de radicación superior al de los demás grupos.",
                        "Conviene revisar el flujo de emisión a radicación en este grupo."))

    # ---- Soportes ----
    if es_valor_valido(k["pct_soporte_incompleto"]):
        insights.append(_insight(
            "Soportes", "Proporción de registros con soporte incompleto.",
            formatear_porcentaje(k["pct_soporte_incompleto"]),
            "Calculado sobre registros con dato de soporte.",
            "Un soporte incompleto puede afectar la radicación y la aceptación de las facturas.",
            "Conviene revisar los procesos de alistamiento documental."))
        if "Estado_Factura" in df.columns:
            base = df[df["Soporte_Completo"].notna()]
            if base["Soporte_Completo"].nunique() == 2:
                estado = normalizar_serie_texto(base["Estado_Factura"])
                problema = estado.str.contains("glosad|devuel", na=False)
                tasa = problema.groupby(base["Soporte_Completo"]).mean() * 100
                if {"Sí", "No"}.issubset(tasa.index):
                    insights.append(_insight(
                        "Soportes", "Relación entre soporte incompleto y estados glosada/devuelta.",
                        f"Incompleto: {formatear_porcentaje(tasa['No'])} | Completo: {formatear_porcentaje(tasa['Sí'])}",
                        "Porcentaje de registros en estado glosada o devuelta según el soporte.",
                        "Se observa una asociación entre soporte y estado; una asociación no demuestra causalidad.",
                        "Conviene validar con una muestra de casos si el soporte incompleto explica las devoluciones."))

    # ---- Glosas, inconsistencias y atípicos ----
    if "Tiene_Glosa" in df.columns:
        frecuencia = df["Tiene_Glosa"].mean() * 100
        insights.append(_insight(
            "Glosas", "Proporción de registros con glosa (frecuencia).", formatear_porcentaje(frecuencia),
            f"Severidad de glosa (valor glosado / valor neto): {formatear_porcentaje(k['pct_glosa'])}.",
            "La frecuencia mide cuántas facturas se glosan; la severidad, cuánto valor se glosa.",
            "Conviene analizar ambas dimensiones por entidad y servicio."))
    if "Inconsistencia_Financiera" in df.columns:
        n_inc = int(df["Inconsistencia_Financiera"].sum())
        insights.append(_insight(
            "Calidad", "Registros con posibles inconsistencias financieras.",
            f"{formatear_numero(n_inc)} ({formatear_porcentaje(dividir_seguro(n_inc, len(df), 100))})",
            "Reglas: identidades de valor neto y saldo, pagos/glosas/saldos mayores al neto y valores negativos.",
            "Las diferencias pueden deberse a glosas conciliadas, notas crédito o errores de registro.",
            "Conviene revisar los registros en el Explorador de datos con el filtro de inconsistencias."))
    mascara, _ = detectar_atipicos(df, "Valor_Neto", "IQR", 1.5)
    n_atip = int(mascara.sum())
    insights.append(_insight(
        "Atípicos", "Registros con valor neto atípico (método IQR, multiplicador 1,5).",
        f"{formatear_numero(n_atip)} ({formatear_porcentaje(dividir_seguro(n_atip, len(df), 100))})",
        "Un valor atípico no es necesariamente un error.",
        "Puede representar facturas legítimas de mayor complejidad o valor.",
        "Conviene revisar los casos de mayor valor en la sección Valores atípicos."))
    return insights


def generar_alertas(df, umbrales):
    """Semáforos de indicadores con umbrales definidos por el usuario (parámetros demostrativos editables)."""
    k = calcular_kpis(df)
    cierre = construir_cierre_mensual(df, "Mensual")
    crecimiento = cierre["Crecimiento %"].iloc[-1] if not cierre.empty and "Crecimiento %" in cierre else np.nan
    u = umbrales
    definiciones = [
        ("Porcentaje de glosa", k["pct_glosa"], "glosa", False, "%"),
        ("Porcentaje de cartera", k["pct_cartera"], "cartera", False, "%"),
        ("Porcentaje de recaudo", k["pct_recaudo"], "recaudo", True, "%"),
        ("Días promedio de radicación", k["dias"], "dias", False, "días"),
        ("Porcentaje de soportes incompletos", k["pct_soporte_incompleto"], "soporte", False, "%"),
        ("Crecimiento del último periodo", crecimiento, "crec", True, "%"),
    ]
    filas = []
    for nombre, valor, clave, mayor_mejor, unidad in definiciones:
        verde, amarillo = u[f"{clave}_verde"], u[f"{clave}_amarillo"]
        color, etiqueta = evaluar_semaforo(valor, verde, amarillo, mayor_mejor)
        if mayor_mejor:
            regla = f"Verde ≥ {formatear_numero(verde, 1)}; Amarillo ≥ {formatear_numero(amarillo, 1)}; Rojo < " \
                    f"{formatear_numero(amarillo, 1)}"
        else:
            regla = f"Verde ≤ {formatear_numero(verde, 1)}; Amarillo ≤ {formatear_numero(amarillo, 1)}; Rojo > " \
                    f"{formatear_numero(amarillo, 1)}"
        valor_txt = formatear_porcentaje(valor) if unidad == "%" else (
            f"{formatear_numero(valor, 1)} días" if es_valor_valido(valor) else "No calculable")
        lectura = {"verde": "Dentro del rango verde definido por el usuario.",
                   "amarillo": "En zona de atención según los umbrales definidos.",
                   "rojo": "Fuera del rango definido; conviene revisar.",
                   "gris": "No es posible calcular el indicador con los datos disponibles."}[color]
        filas.append({"Indicador": nombre, "Valor": valor_txt, "Estado": f"{ICONO_SEMAFORO[color]} {etiqueta}",
                      "Umbrales configurados": regla, "Lectura": lectura, "_color": color})
    return filas


def _preparar_para_exportar(df):
    """Elimina columnas internas y convierte tipos no compatibles con Excel."""
    salida = df[columnas_visibles(df)].copy()
    for col in salida.columns:
        if isinstance(salida[col].dtype, pd.CategoricalDtype):
            salida[col] = salida[col].astype("string")
        elif pd.api.types.is_datetime64_any_dtype(salida[col]):
            if getattr(salida[col].dt, "tz", None) is not None:
                salida[col] = salida[col].dt.tz_localize(None)
    return salida


def convertir_dataframe_csv(df):
    """Convierte un DataFrame a CSV (UTF-8 con BOM para abrir correctamente en Excel)."""
    buffer = io.StringIO()
    _preparar_para_exportar(df).to_csv(buffer, index=False, sep=";", decimal=",")
    return buffer.getvalue().encode("utf-8-sig")


LIMITE_FILAS_EXCEL = 1_048_575


def convertir_dataframe_excel(hojas):
    """Convierte uno o varios DataFrames a un libro Excel en memoria (BytesIO)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        for nombre, tabla in hojas.items():
            datos = _preparar_para_exportar(tabla) if tabla is not None else pd.DataFrame()
            if len(datos) > LIMITE_FILAS_EXCEL:
                datos = datos.head(LIMITE_FILAS_EXCEL)
            nombre_hoja = re.sub(r"[\[\]\*\?/\\:]", "_", str(nombre))[:31] or "Hoja"
            datos.to_excel(escritor, sheet_name=nombre_hoja, index=False)
    return buffer.getvalue()


@st.cache_data(show_spinner=False, max_entries=24)
def preparar_descarga(clave, tipo, _hojas):
    """Prepara en caché el contenido binario de una descarga (Excel o CSV)."""
    if tipo == "excel":
        return convertir_dataframe_excel(_hojas)
    primera = next(iter(_hojas.values()))
    return convertir_dataframe_csv(primera)


def texto_insights(insights, alertas, info):
    """Construye el resumen de insights en texto plano para descarga."""
    lineas = [APP_TITULO, APP_SUBTITULO, "=" * 70,
              f"Archivo: {info.get('nombre', '')}", f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
              f"Registros analizados (filtrados): {info.get('registros', '')}", "", AVISO_EXPLORATORIO, "",
              "SEMÁFOROS (umbrales definidos por el usuario)", "-" * 70]
    for alerta in alertas:
        lineas.append(f"- {alerta['Indicador']}: {alerta['Valor']} -> {alerta['Estado']} "
                      f"({alerta['Umbrales configurados']})")
    lineas += ["", "INSIGHTS", "-" * 70]
    for i, ins in enumerate(insights, 1):
        lineas += [f"{i}. [{ins['Categoría']}] {ins['Hallazgo calculado']}",
                   f"   Valor observado: {ins['Valor observado']}",
                   f"   Contexto: {ins['Contexto']}",
                   f"   Interpretación: {ins['Interpretación']}",
                   f"   Recomendación para investigar: {ins['Recomendación para investigar']}", ""]
    return "\n".join(lineas).encode("utf-8")


# =============================================================================
# 10. ESTADO DE SESIÓN, PROCESAMIENTO Y BARRA LATERAL
# =============================================================================

ESTADO_INICIAL = {
    "archivo": None, "pagina": PAGINAS[0], "moneda": "COP", "hoja_sel": None, "sep_sel": "auto",
    "cod_sel": "auto", "mapeo": None, "mapeo_origen": None, "mapeo_detalle": [], "ambiguedades": {},
    "uploader_n": 0, "conv_fechas": True, "conv_moneda": True, "limp_aplicar": True, "limp_vacias": True,
    "limp_duplicados": False, "limp_rellenar": True, "limp_neto": True, "limp_fechas": False,
    "tolerancia": 1.0, "rt_1": 3, "rt_2": 7, "rt_3": 15, "ant_1": 30, "ant_2": 60, "ant_3": 90, "ant_4": 180,
    "vista_tabla_filas": 100,
}
ESTADO_INICIAL.update({f"u_{k}": v for k, v in UMBRALES_DEFECTO.items()})
PREFIJOS_REINICIO_ARCHIVO = ("flt_", "map_", "w_")


def inicializar_estado():
    """Crea las claves de sesión con valores por defecto si no existen."""
    for clave, valor in ESTADO_INICIAL.items():
        if clave not in st.session_state:
            st.session_state[clave] = valor


def _borrar_claves(prefijos):
    """Elimina claves de sesión que comienzan con alguno de los prefijos indicados."""
    for clave in list(st.session_state.keys()):
        if str(clave).startswith(prefijos):
            del st.session_state[clave]


def cb_cargar_archivo():
    """Callback del cargador: guarda el archivo en memoria de sesión y reinicia el estado si es nuevo."""
    clave = f"uploader_{st.session_state.get('uploader_n', 0)}"
    archivo = st.session_state.get(clave)
    if archivo is None:
        return
    contenido = archivo.getvalue()
    firma = firma_archivo(archivo.name, contenido)
    actual = st.session_state.get("archivo")
    if actual and actual.get("firma") == firma:
        return
    _borrar_claves(PREFIJOS_REINICIO_ARCHIVO)
    for clave_estado in ("mapeo", "mapeo_origen", "hoja_sel"):
        st.session_state[clave_estado] = None
    st.session_state["sep_sel"], st.session_state["cod_sel"] = "auto", "auto"
    st.session_state["mapeo_detalle"], st.session_state["ambiguedades"] = [], {}
    extension = archivo.name.rsplit(".", 1)[-1].lower() if "." in archivo.name else ""
    st.session_state["archivo"] = {"nombre": archivo.name, "bytes": contenido, "firma": firma,
                                   "tamano": len(contenido), "extension": extension}


def cb_reiniciar():
    """Limpia todos los datos de la sesión y vuelve a la pantalla inicial."""
    siguiente = st.session_state.get("uploader_n", 0) + 1
    for clave in list(st.session_state.keys()):
        del st.session_state[clave]
    st.session_state["uploader_n"] = siguiente


def cb_copiar(origen, destino, reiniciar_mapeo=False):
    """Copia el valor de un widget a una clave persistente (antes de la siguiente ejecución)."""
    st.session_state[destino] = st.session_state.get(origen)
    if reiniciar_mapeo:
        st.session_state["mapeo"] = None
        _borrar_claves(("map_", "flt_"))


def cb_mapeo(campo):
    """Actualiza el mapeo de un campo a partir de su selector."""
    valor = st.session_state.get(f"map_{campo}")
    mapeo = dict(st.session_state.get("mapeo") or {})
    mapeo[campo] = None if valor in (None, NO_DISPONIBLE) else valor
    st.session_state["mapeo"] = mapeo
    _borrar_claves(("flt_",))


def cb_restablecer_mapeo():
    """Vuelve a ejecutar la detección automática de columnas."""
    st.session_state["mapeo"] = None
    _borrar_claves(("map_", "flt_"))


def cb_restaurar_originales():
    """Desactiva la limpieza para volver a los datos originales (conservados en memoria)."""
    st.session_state["limp_aplicar"] = False
    _borrar_claves(("flt_",))


def cb_limpiar_filtros():
    """Restablece todos los filtros globales."""
    _borrar_claves(("flt_",))


def cb_seleccionar_todos(opciones):
    """Marca explícitamente todas las opciones de cada filtro categórico."""
    for clave, valores in opciones.items():
        st.session_state[clave] = list(valores)


def elegir_hoja_defecto(hojas):
    """Selecciona la hoja 'Facturacion' si existe (sin importar tildes o mayúsculas); si no, la primera."""
    for hoja in hojas:
        if clave_compacta(hoja) == "facturacion":
            return hoja
    return hojas[0] if hojas else None


@st.cache_data(show_spinner=False, max_entries=6)
def procesar_datos(clave, _crudo, mapeo_items, conv_fechas, conv_moneda, opciones_limpieza, limites_tiempo,
                   tolerancia):
    """Canal completo de preparación (en caché): mapeo, conversión, diagnóstico, limpieza y derivadas."""
    mapeo = dict(mapeo_items)
    mapeado = mapear_columnas(_crudo, mapeo)
    convertido, reporte = convertir_tipos(mapeado, conv_fechas, conv_moneda)
    calidad = validar_datos(_crudo, convertido, reporte, tolerancia)
    limpio, bitacora = limpiar_datos(convertido, dict(opciones_limpieza))
    final = crear_variables_derivadas(limpio, limites_tiempo, tolerancia)
    return {"df": final, "reporte": reporte, "calidad": calidad, "bitacora": bitacora,
            "n_inicial": len(_crudo), "n_final": len(final)}


def obtener_datos():
    """Lee el archivo de la sesión y ejecuta el procesamiento. Devuelve un diccionario de contexto."""
    archivo = st.session_state.get("archivo")
    if not archivo:
        return None
    ctx = {"archivo": archivo, "hojas": [], "error": None, "info_csv": {}}
    extension = archivo["extension"]
    if extension not in ("xlsx", "xls", "csv", "txt"):
        ctx["error"] = "Formato no admitido. Cargue un archivo .xlsx, .xls o .csv."
        return ctx
    if archivo["tamano"] == 0:
        ctx["error"] = "El archivo está vacío."
        return ctx
    if extension in ("xlsx", "xls"):
        hojas, error = listar_hojas(archivo["firma"], archivo["bytes"], extension)
        ctx["hojas"] = hojas
        if error or not hojas:
            ctx["error"] = error or "El libro no contiene hojas."
            return ctx
        hoja = st.session_state.get("hoja_sel")
        if hoja not in hojas:
            hoja = elegir_hoja_defecto(hojas)
            st.session_state["hoja_sel"] = hoja
        ctx["hoja"] = hoja
        crudo, error = leer_excel(archivo["firma"], archivo["bytes"], extension, hoja)
        origen = (archivo["firma"], hoja)
    else:
        crudo, error, info = leer_csv(archivo["firma"], archivo["bytes"], st.session_state["sep_sel"],
                                      st.session_state["cod_sel"])
        ctx["info_csv"], ctx["hoja"] = info, None
        origen = (archivo["firma"], st.session_state["sep_sel"], st.session_state["cod_sel"])
    if error or crudo is None:
        ctx["error"] = error or "No fue posible leer el archivo."
        return ctx
    ctx["crudo"] = crudo

    if st.session_state.get("mapeo") is None or st.session_state.get("mapeo_origen") != origen:
        mapeo, detalle, ambiguedades = detectar_columnas(list(crudo.columns))
        st.session_state["mapeo"], st.session_state["mapeo_origen"] = mapeo, origen
        st.session_state["mapeo_detalle"], st.session_state["ambiguedades"] = detalle, ambiguedades
        _borrar_claves(("map_",))
    mapeo = {c: st.session_state["mapeo"].get(c) for c in CAMPOS}
    ctx["conflictos_mapeo"] = validar_mapeo(mapeo)

    opciones_limpieza = (("aplicar", st.session_state["limp_aplicar"]), ("vacias", st.session_state["limp_vacias"]),
                         ("duplicados", st.session_state["limp_duplicados"]),
                         ("rellenar", st.session_state["limp_rellenar"]), ("neto_invalido", st.session_state["limp_neto"]),
                         ("fechas_invalidas", st.session_state["limp_fechas"]))
    limites_tiempo = tuple(sorted({int(st.session_state["rt_1"]), int(st.session_state["rt_2"]),
                                   int(st.session_state["rt_3"])}))
    tolerancia = float(st.session_state["tolerancia"])
    mapeo_items = tuple(sorted((k, v) for k, v in mapeo.items() if v))
    clave = clave_hash(origen, mapeo_items, st.session_state["conv_fechas"], st.session_state["conv_moneda"],
                       opciones_limpieza, limites_tiempo, tolerancia)
    try:
        with st.spinner("Procesando el archivo..."):
            resultado = procesar_datos(clave, crudo, mapeo_items, st.session_state["conv_fechas"],
                                       st.session_state["conv_moneda"], opciones_limpieza, limites_tiempo, tolerancia)
    except Exception as error:  # noqa: BLE001
        ctx["error"] = f"No fue posible procesar el archivo ({type(error).__name__}: {error})."
        return ctx
    ctx.update(resultado)
    ctx["clave"] = clave
    ctx["minimas_ok"], ctx["faltantes"] = evaluar_columnas_minimas(resultado["df"])
    return ctx


def render_navegacion():
    """Menú lateral de navegación entre secciones."""
    st.sidebar.markdown("### 📊 Analítica de Facturación en Salud")
    st.sidebar.caption(APP_SUBTITULO)
    st.sidebar.radio("Navegación", PAGINAS, key="pagina", label_visibility="collapsed")


def render_configuracion_lateral(hay_archivo):
    """Opciones de moneda, limpieza controlada, parámetros analíticos y umbrales en la barra lateral."""
    with st.sidebar.expander("💱 Moneda (solo visual)", expanded=False):
        st.selectbox("Etiqueta de moneda", list(MONEDAS.keys()), key="moneda",
                     help="Cambia solo el símbolo mostrado. No convierte tasas de cambio.")
    if not hay_archivo:
        return
    with st.sidebar.expander("🧹 Limpieza controlada", expanded=False):
        st.checkbox("Convertir columnas de fecha", key="conv_fechas",
                    help="Convierte con pd.to_datetime(errors='coerce') e informa los valores no convertidos.")
        st.checkbox("Convertir columnas monetarias", key="conv_moneda",
                    help="Reconoce símbolos de moneda, separadores de miles y coma o punto decimal.")
        st.divider()
        st.checkbox("Aplicar transformaciones de limpieza", key="limp_aplicar")
        deshabilitado = not st.session_state["limp_aplicar"]
        st.checkbox("Eliminar filas completamente vacías", key="limp_vacias", disabled=deshabilitado)
        st.checkbox("Eliminar duplicados exactos", key="limp_duplicados", disabled=deshabilitado)
        st.checkbox(f"Reemplazar categorías vacías por '{SIN_INFO}'", key="limp_rellenar", disabled=deshabilitado)
        st.checkbox("Excluir registros con Valor_Neto inválido", key="limp_neto", disabled=deshabilitado)
        st.checkbox("Excluir registros con fechas inválidas", key="limp_fechas", disabled=deshabilitado)
        st.button("Restaurar datos originales", on_click=cb_restaurar_originales, width="stretch",
                  help="Desactiva la limpieza. Los datos originales se conservan en memoria durante la sesión.")
    with st.sidebar.expander("⚙️ Parámetros analíticos", expanded=False):
        st.number_input("Tolerancia de redondeo (valor monetario)", min_value=0.0, step=1.0, key="tolerancia",
                        help="Diferencia máxima aceptada en las validaciones de identidad financiera.")
        st.caption("Rangos de tiempo de radicación (días). Parámetros analíticos, no estándares normativos.")
        c1, c2, c3 = st.columns(3)
        c1.number_input("Límite 1", min_value=0, step=1, key="rt_1")
        c2.number_input("Límite 2", min_value=0, step=1, key="rt_2")
        c3.number_input("Límite 3", min_value=0, step=1, key="rt_3")
        st.caption("Rangos de antigüedad de cartera (días). Analíticos y configurables.")
        a1, a2, a3, a4 = st.columns(4)
        a1.number_input("L1", min_value=0, step=1, key="ant_1")
        a2.number_input("L2", min_value=0, step=1, key="ant_2")
        a3.number_input("L3", min_value=0, step=1, key="ant_3")
        a4.number_input("L4", min_value=0, step=1, key="ant_4")
    with st.sidebar.expander("🚦 Umbrales de alertas", expanded=False):
        st.caption("Valores demostrativos y editables. No son estándares oficiales ni metas contractuales.")
        for clave, nombre, mayor_mejor in [("glosa", "% glosa", False), ("cartera", "% cartera", False),
                                           ("recaudo", "% recaudo", True), ("dias", "Días de radicación", False),
                                           ("soporte", "% soportes incompletos", False),
                                           ("crec", "Crecimiento %", True)]:
            st.markdown(f"**{nombre}** {'(mayor es mejor)' if mayor_mejor else '(menor es mejor)'}")
            c1, c2 = st.columns(2)
            c1.number_input("Verde", step=1.0, key=f"u_{clave}_verde", label_visibility="visible")
            c2.number_input("Amarillo", step=1.0, key=f"u_{clave}_amarillo")


def obtener_umbrales():
    """Lee los umbrales configurados en la barra lateral."""
    return {k: float(st.session_state.get(f"u_{k}", v)) for k, v in UMBRALES_DEFECTO.items()}


def render_filtros(contenedor, df):
    """Filtros globales en la barra lateral. Devuelve (df_filtrado, filtros)."""
    filtros = {}
    with contenedor.expander("🔎 Filtros globales", expanded=True):
        opciones_cat = {}
        # Rango de fechas (fecha de radicación o, en su defecto, el periodo)
        if "_Fecha_Referencia" in df.columns and df["_Fecha_Referencia"].notna().any():
            minimo = df["_Fecha_Referencia"].min().date()
            maximo = df["_Fecha_Referencia"].max().date()
            valor = st.session_state.get("flt_fechas")
            if not (isinstance(valor, (tuple, list)) and len(valor) == 2 and minimo <= valor[0] <= valor[1] <= maximo):
                st.session_state["flt_fechas"] = (minimo, maximo)
            etiqueta = "Rango de fechas de radicación" if columna_util(df, "Fecha_Radicacion") else \
                "Rango de fechas del periodo"
            rango = st.date_input(etiqueta, min_value=minimo, max_value=maximo, key="flt_fechas",
                                  format="DD/MM/YYYY")
            if isinstance(rango, (tuple, list)) and len(rango) == 2 and (rango[0] > minimo or rango[1] < maximo):
                filtros["fechas"] = (rango[0], rango[1])
        if "Año_Periodo" in df.columns:
            anios = sorted(df["Año_Periodo"].dropna().astype(int).astype(str).unique().tolist())
            st.session_state["flt_anios"] = [a for a in st.session_state.get("flt_anios", []) if a in anios]
            filtros["anios"] = st.multiselect("Año", anios, key="flt_anios", placeholder="Todos")
            opciones_cat["flt_anios"] = anios
        if "Mes_Periodo" in df.columns:
            numeros = sorted(df["Numero_Mes_Periodo"].dropna().astype(int).unique().tolist())
            meses = [MESES_ES[n] for n in numeros]
            st.session_state["flt_meses"] = [m for m in st.session_state.get("flt_meses", []) if m in meses]
            filtros["meses"] = st.multiselect("Mes", meses, key="flt_meses", placeholder="Todos")
            opciones_cat["flt_meses"] = meses
        for columna, etiqueta in FILTROS_CATEGORICOS:
            if columna in df.columns and df[columna].notna().any():
                opciones = opciones_categoria(df, columna)
                clave = f"flt_{columna}"
                st.session_state[clave] = [v for v in st.session_state.get(clave, []) if v in opciones]
                filtros[columna] = st.multiselect(etiqueta, opciones, key=clave, placeholder="Todos")
                opciones_cat[clave] = opciones
        if "Tiene_Glosa" in df.columns:
            filtros["glosa"] = st.radio("Facturas con glosa", ["Todas", "Con glosa", "Sin glosa"], key="flt_glosa",
                                        horizontal=True)
        if "Tiene_Saldo" in df.columns:
            filtros["saldo"] = st.radio("Facturas con saldo", ["Todas", "Con saldo", "Sin saldo"], key="flt_saldo",
                                        horizontal=True)
        if columna_util(df, "Valor_Neto"):
            minimo_n, maximo_n = float(df["Valor_Neto"].min()), float(df["Valor_Neto"].max())
            if maximo_n > minimo_n:
                valor = st.session_state.get("flt_neto")
                if not (isinstance(valor, (tuple, list)) and len(valor) == 2
                        and minimo_n <= valor[0] <= valor[1] <= maximo_n):
                    st.session_state["flt_neto"] = (minimo_n, maximo_n)
                rango_neto = st.slider("Rango de valor neto", min_value=minimo_n, max_value=maximo_n,
                                       key="flt_neto", format="%.0f")
                if rango_neto[0] > minimo_n or rango_neto[1] < maximo_n:
                    filtros["neto"] = rango_neto
        c1, c2 = st.columns(2)
        c1.button("Seleccionar todos", on_click=cb_seleccionar_todos, args=(opciones_cat,), width="stretch",
                  help="Marca explícitamente todas las categorías. Un filtro vacío también equivale a 'todos'.")
        c2.button("Limpiar filtros", on_click=cb_limpiar_filtros, width="stretch")
    # Si el usuario seleccionó todas las opciones, el filtro no excluye registros
    for clave_opc, valores in opciones_cat.items():
        nombre = clave_opc.replace("flt_", "")
        clave_filtro = {"anios": "anios", "meses": "meses"}.get(nombre, nombre)
        if filtros.get(clave_filtro) and len(filtros[clave_filtro]) == len(valores):
            filtros[clave_filtro] = []
    try:
        filtrado = aplicar_filtros(df, filtros)
    except Exception as error:  # noqa: BLE001
        contenedor.error(f"No fue posible aplicar los filtros ({type(error).__name__}). Se muestran todos los datos.")
        filtrado = df
    total = len(df)
    contenedor.metric("Registros filtrados", formatear_numero(len(filtrado)),
                      help="Registros que cumplen todos los filtros activos.")
    contenedor.caption(f"Visibles: {formatear_porcentaje(dividir_seguro(len(filtrado), total, 100))} de "
                       f"{formatear_numero(total)} registros.")
    if filtrado.empty:
        contenedor.warning("Los filtros actuales no tienen registros.")
    return filtrado, filtros


# =============================================================================
# 11. PÁGINAS DE LA INTERFAZ
# =============================================================================

def encabezado():
    """Encabezado profesional y aviso de uso exploratorio."""
    st.markdown(f'<div class="encabezado-app"><h1>{APP_TITULO}</h1><p>{APP_SUBTITULO}</p></div>',
                unsafe_allow_html=True)
    caja_aviso(f"ℹ️ {AVISO_EXPLORATORIO}")


def requiere_columnas(df, columnas, analisis):
    """Verifica columnas necesarias para un análisis. Muestra advertencia y devuelve False si faltan."""
    faltantes = [c for c in columnas if not columna_util(df, c)]
    if faltantes:
        st.warning(f"No se encontró la columna necesaria para {analisis}: {', '.join(faltantes)}. "
                   "Asígnela en 'Configurar correspondencia de columnas' (página de inicio) si existe con otro nombre.")
        return False
    return True


def selector_top(etiqueta="Mostrar", clave=None, defecto=1):
    """Selector Top 5 / 10 / 15 / Todos. Devuelve el número de categorías o None para todas."""
    opcion = st.radio(etiqueta, ["Top 5", "Top 10", "Top 15", "Todos"], index=defecto, horizontal=True, key=clave)
    return None if opcion == "Todos" else int(opcion.split()[1])


def _limitar_categorias(tabla, n):
    """Limita las categorías a mostrar para mantener los gráficos legibles."""
    if n is None:
        return tabla.head(40) if len(tabla) > 40 else tabla
    return tabla.head(n)


def mostrar_inicio(ctx):
    """Página 1: carga del archivo, selección de hoja, mapeo de columnas y vista previa."""
    st.subheader("Inicio y carga de datos")
    st.warning(f"🔒 {AVISO_PRIVACIDAD}")
    st.file_uploader("Cargue el archivo de facturación que desea analizar", type=["xlsx", "xls", "csv"],
                     key=f"uploader_{st.session_state.get('uploader_n', 0)}", on_change=cb_cargar_archivo,
                     help="El archivo se procesa en memoria durante la sesión; no se almacena ni se envía a terceros. "
                          "Límite de carga por defecto de Streamlit: 200 MB.")
    if ctx is None:
        c1, c2 = st.columns([3, 2])
        with c1:
            st.markdown("#### Objetivo")
            st.write("Cargar, validar, explorar y analizar archivos de facturación del sector salud: facturación, "
                     "radicación, glosas, recaudo, cartera, tiempos, concentración y valores atípicos, con insights "
                     "calculados y resultados descargables.")
            st.markdown("#### Instrucciones")
            st.markdown("1. Cargue un archivo **.xlsx**, **.xls** o **.csv**.\n"
                        "2. Si es Excel, confirme la hoja (se selecciona *Facturacion* si existe).\n"
                        "3. Revise la correspondencia de columnas y ajústela si es necesario.\n"
                        "4. Consulte el diagnóstico de calidad y configure la limpieza en la barra lateral.\n"
                        "5. Aplique filtros globales y navegue por las secciones de análisis.")
        with c2:
            st.markdown("#### Columnas mínimas")
            st.markdown("\n".join(f"- `{c}`" for c in COLUMNAS_MINIMAS))
            st.markdown("#### Columnas recomendadas")
            st.markdown("\n".join(f"- `{c}`" for c in COLUMNAS_RECOMENDADAS))
            st.caption("Las demás columnas son opcionales. Los análisis que dependan de columnas ausentes se "
                       "desactivan sin afectar el resto de la aplicación.")
        return

    archivo = ctx["archivo"]
    if ctx.get("hojas"):
        indice = ctx["hojas"].index(ctx["hoja"]) if ctx.get("hoja") in ctx["hojas"] else 0
        st.selectbox("Hoja a analizar", ctx["hojas"], index=indice, key="w_hoja",
                     on_change=cb_copiar, args=("w_hoja", "hoja_sel", True),
                     help="Se selecciona automáticamente 'Facturacion' si existe.")
    if archivo["extension"] in ("csv", "txt"):
        separadores = {"Detección automática": "auto", "Coma (,)": ",", "Punto y coma (;)": ";",
                       "Tabulación": "\t", "Barra vertical (|)": "|"}
        codificaciones = {"Detección automática": "auto", "UTF-8": "utf-8-sig", "Windows-1252": "cp1252",
                          "Latin-1": "latin-1"}
        c1, c2 = st.columns(2)
        actual_sep = [k for k, v in separadores.items() if v == st.session_state["sep_sel"]][0]
        actual_cod = [k for k, v in codificaciones.items() if v == st.session_state["cod_sel"]][0]
        sep = c1.selectbox("Separador", list(separadores), index=list(separadores).index(actual_sep))
        cod = c2.selectbox("Codificación", list(codificaciones), index=list(codificaciones).index(actual_cod))
        if separadores[sep] != st.session_state["sep_sel"] or codificaciones[cod] != st.session_state["cod_sel"]:
            st.session_state["sep_sel"], st.session_state["cod_sel"] = separadores[sep], codificaciones[cod]
            st.session_state["mapeo"] = None
            _borrar_claves(("map_", "flt_"))
            st.rerun()
        info = ctx.get("info_csv") or {}
        if info:
            st.caption(f"Codificación utilizada: {info.get('codificacion')} · Separador detectado: "
                       f"{repr(info.get('separador'))}")
            if info.get("advertencia"):
                st.warning(info["advertencia"])

    if ctx.get("error"):
        st.error(ctx["error"])
        st.info("Verifique el archivo o seleccione otra hoja / separador. Puede usar 'Reiniciar aplicación' en la "
                "barra lateral para comenzar de nuevo.")
        return

    df = ctx["df"]
    tamano = archivo["tamano"]
    if tamano > 50 * 1024 * 1024:
        st.info("El archivo es grande; algunas secciones pueden tardar más. Los cálculos usan caché para evitar "
                "repeticiones.")
    st.markdown("#### Resumen del archivo")
    tamano_txt = f"{formatear_numero(tamano / 1024, 1)} KB" if tamano < 1024 ** 2 \
        else f"{formatear_numero(tamano / 1024 ** 2, 1)} MB"
    if "Periodo" in df.columns and df["Periodo"].notna().any():
        periodos = df["Periodo"].dropna()
        periodo_txt = f"{periodos.min()} a {periodos.max()}"
    else:
        periodo_txt = "No disponible"
    datos_resumen = [("Archivo", archivo["nombre"], ""), ("Tamaño", tamano_txt, ""),
                     ("Hoja utilizada", ctx.get("hoja") or "CSV", ""),
                     ("Filas", formatear_numero(ctx["n_inicial"]), f"{formatear_numero(ctx['n_final'])} tras limpieza"),
                     ("Columnas", formatear_numero(ctx["crudo"].shape[1]), ""), ("Periodo identificado", periodo_txt, "")]
    for col, nombre in [("Entidad_Pagadora", "Entidades"), ("Sede", "Sedes"), ("Servicio", "Servicios")]:
        datos_resumen.append((nombre, formatear_numero(df[col].nunique()) if col in df.columns else "No disponible", ""))
    columnas_resumen = st.columns(5)
    for i, (titulo, valor, sub) in enumerate(datos_resumen):
        with columnas_resumen[i % 5]:
            st.markdown(f'<div class="kpi-card" style="min-height:86px"><div class="kpi-titulo">{titulo}</div>'
                        f'<div class="kpi-valor" style="font-size:1.05rem;word-break:break-word">{valor}</div>'
                        f'<div class="kpi-sub">{sub}</div></div>', unsafe_allow_html=True)

    # ---- Estado de columnas mínimas ----
    if ctx["minimas_ok"]:
        st.success("Las columnas mínimas están disponibles. Los análisis financieros están habilitados.")
    else:
        st.error("Faltan columnas mínimas. Los análisis financieros principales están desactivados.")
        st.markdown(f"**Columnas faltantes o no utilizables:** {', '.join(ctx['faltantes'])}")
        st.markdown(f"**Nombres esperados:** {', '.join(COLUMNAS_MINIMAS)}")
        st.markdown(f"**Columnas encontradas en el archivo:** {', '.join(map(str, ctx['crudo'].columns))}")
        st.info("Cómo corregirlo: asigne la columna equivalente en 'Configurar correspondencia de columnas' o "
                "renombre el encabezado en el archivo. Verifique que Valor_Neto sea numérico y que las fechas o "
                "periodos tengan un formato reconocible.")
    faltan_recomendadas = [c for c in COLUMNAS_RECOMENDADAS if not columna_util(df, c)]
    if faltan_recomendadas:
        st.warning(f"Columnas recomendadas ausentes: {', '.join(faltan_recomendadas)}. Los análisis que dependen de "
                   "ellas se desactivarán.")

    # ---- Equivalencias y mapeo ----
    detalle = pd.DataFrame(st.session_state.get("mapeo_detalle") or [])
    with st.expander("Equivalencias detectadas automáticamente", expanded=False):
        if detalle.empty:
            st.info("No se detectaron equivalencias automáticas.")
        else:
            st.dataframe(detalle, hide_index=True)
        for campo, candidatas in (st.session_state.get("ambiguedades") or {}).items():
            st.warning(f"Ambigüedad para {campo}: varias columnas coinciden ({', '.join(candidatas)}). "
                       f"Se asignó '{st.session_state['mapeo'].get(campo)}'. Puede cambiarla manualmente.")
    expandir = (not ctx["minimas_ok"]) or bool(st.session_state.get("ambiguedades")) or \
        bool(ctx.get("conflictos_mapeo"))
    with st.expander("Configurar correspondencia de columnas", expanded=expandir):
        st.caption("Asocie cada campo esperado con una columna del archivo. El mapeo se conserva solo durante la "
                   "sesión.")
        opciones = [NO_DISPONIBLE] + list(map(str, ctx["crudo"].columns))
        columnas_ui = st.columns(3)
        for i, (campo, meta) in enumerate(CAMPOS.items()):
            clave = f"map_{campo}"
            actual = st.session_state["mapeo"].get(campo) or NO_DISPONIBLE
            if clave not in st.session_state or st.session_state[clave] not in opciones:
                st.session_state[clave] = actual if actual in opciones else NO_DISPONIBLE
            columnas_ui[i % 3].selectbox(f"{meta['etiqueta']} ({campo})", opciones, key=clave,
                                         on_change=cb_mapeo, args=(campo,))
        for origen, campos in (ctx.get("conflictos_mapeo") or {}).items():
            st.error(f"La columna '{origen}' está asignada a varios campos ({', '.join(campos)}). Solo se usará para "
                     f"'{campos[0]}'. Ajuste el mapeo para evitar duplicados.")
        st.button("Restablecer detección automática", on_click=cb_restablecer_mapeo)

    # ---- Vista previa ----
    st.markdown("#### Vista previa")
    modo = st.radio("Registros a visualizar", ["Primeros registros", "Muestra aleatoria"], horizontal=True)
    n_filas = st.slider("Número de filas", 5, 100, 10, step=5)
    vista = df[columnas_visibles(df)]
    if modo == "Muestra aleatoria":
        vista = vista.sample(min(n_filas, len(vista)), random_state=None) if len(vista) else vista
    else:
        vista = vista.head(n_filas)
    st.dataframe(formatear_tabla(vista, columnas_moneda=CAMPOS_MONEDA), hide_index=True, height=360)


def mostrar_calidad(ctx):
    """Página 2: diagnóstico de calidad de datos, conversiones y limpieza aplicada."""
    st.subheader("Diagnóstico de calidad de datos")
    cal = ctx["calidad"]
    c1, c2 = st.columns([1, 2])
    with c1:
        puntaje = cal["puntaje"]
        color = COLOR_VERDE if puntaje >= 85 else COLOR_NARANJA if puntaje >= 70 else COLOR_ROJO
        fig = go.Figure(go.Indicator(mode="gauge+number", value=puntaje, number={"suffix": " / 100"},
                                     gauge={"axis": {"range": [0, 100]}, "bar": {"color": color},
                                            "steps": [{"range": [0, 70], "color": "#FBE3E0"},
                                                      {"range": [70, 85], "color": "#FFF0DB"},
                                                      {"range": [85, 100], "color": "#E3F4EA"}]}))
        mostrar_grafico(estilizar_figura(fig, "Puntaje orientativo de calidad", 280))
        st.caption(AVISO_CALIDAD)
    with c2:
        comp = pd.DataFrame({"Dimensión": list(cal["componentes"].keys()),
                             "Puntaje": list(cal["componentes"].values())})
        fig = px.bar(comp, x="Puntaje", y="Dimensión", orientation="h", text=comp["Puntaje"].round(1),
                     color_discrete_sequence=[COLOR_AZUL_MEDIO], range_x=[0, 100])
        mostrar_grafico(estilizar_figura(fig, "Componentes del puntaje (0-100)", 280, leyenda=False))

    st.markdown("#### Indicadores de calidad")
    filas = [
        ("Número de filas", cal["filas"]), ("Número de columnas", cal["columnas"]),
        ("Campos esperados encontrados", len(cal["encontradas"])), ("Campos esperados ausentes", len(cal["ausentes"])),
        ("Filas duplicadas exactas", cal["filas_duplicadas"]), ("ID de factura duplicados", cal["ids_duplicados"]),
        ("Radicados duplicados", cal["radicados_duplicados"]), ("Fechas inválidas", cal["fechas_invalidas"]),
        ("Valores financieros no numéricos", cal["no_numericos"]),
        ("Valores negativos (monetarios)", sum(cal["negativos"].values()) if cal["negativos"] else 0),
        ("Valores netos iguales a cero", cal["neto_cero"]), ("Facturas sin entidad pagadora", cal["sin_entidad"]),
        ("Facturas sin periodo", cal["sin_periodo"]),
    ]
    etiquetas_banderas = {
        "Radicación anterior a emisión": "Fechas de radicación anteriores a la emisión",
        "Pago anterior a radicación": "Fechas de pago anteriores a la radicación",
        "Saldo mayor al valor neto": "Saldos de cartera superiores al valor neto",
        "Pago mayor al valor neto": "Pagos superiores al valor neto",
        "Glosa mayor al valor neto": "Glosas superiores al valor neto",
        "Identidad valor neto": "Inconsistencias: Valor_Neto ≈ Bruto − Descuento − Copago",
        "Identidad saldo de cartera": "Inconsistencias: Saldo ≈ Neto − Glosa − Pagado",
        "Valor monetario negativo": "Registros con algún valor monetario negativo",
    }
    for bandera, valor in cal["banderas"].items():
        filas.append((etiquetas_banderas.get(bandera, bandera), valor))
    tabla = pd.DataFrame(filas, columns=["Indicador", "Resultado"])
    tabla["Resultado"] = tabla["Resultado"].map(lambda v: "No evaluable (columna ausente)" if v is None
                                                else formatear_numero(v))
    tabla["% de filas"] = [formatear_porcentaje(dividir_seguro(v, cal["filas"], 100))
                           if isinstance(v, (int, np.integer)) and i >= 4 else "" for i, (_, v) in enumerate(filas)]
    st.dataframe(tabla, hide_index=True, height=min(38 + 35 * len(tabla), 700))
    if "Identidad saldo de cartera" in cal["banderas"]:
        st.caption("Las diferencias en la identidad del saldo pueden reflejar glosas levantadas o conciliadas, notas "
                   "crédito o pagos posteriores, y no necesariamente errores. Tolerancia aplicada: "
                   f"{formatear_moneda(st.session_state['tolerancia'])}.")
    if cal["ids_duplicados"]:
        st.caption("Los ID de factura repetidos pueden ser legítimos si el archivo contiene una fila por servicio. "
                   "Los conteos de facturas usan valores únicos (nunique).")

    t1, t2, t3, t4 = st.tabs(["Columnas", "Valores nulos", "Conversiones de tipo", "Limpieza aplicada"])
    with t1:
        c1, c2 = st.columns(2)
        c1.markdown("**Campos esperados encontrados**")
        c1.write(", ".join(cal["encontradas"]) or "Ninguno")
        c2.markdown("**Campos esperados ausentes**")
        c2.write(", ".join(cal["ausentes"]) or "Ninguno")
        minimas = pd.DataFrame({"Columna": COLUMNAS_MINIMAS,
                                "Estado": ["Faltante" if c in ctx["faltantes"] else "Disponible"
                                           for c in COLUMNAS_MINIMAS]})
        recomendadas = pd.DataFrame({"Columna": COLUMNAS_RECOMENDADAS,
                                     "Estado": ["Disponible" if columna_util(ctx["df"], c) else "Faltante"
                                                for c in COLUMNAS_RECOMENDADAS]})
        c1.markdown("**Columnas mínimas**")
        c1.dataframe(minimas, hide_index=True)
        c2.markdown("**Columnas recomendadas**")
        c2.dataframe(recomendadas, hide_index=True)
    with t2:
        nulos = cal["tabla_nulos"]
        con_nulos = nulos[nulos["Nulos"] > 0]
        if con_nulos.empty:
            st.success("No se encontraron valores nulos.")
        else:
            fig = px.bar(con_nulos.head(30), x="% nulos", y="Columna", orientation="h",
                         color_discrete_sequence=[COLOR_NARANJA])
            fig.update_yaxes(autorange="reversed")
            mostrar_grafico(estilizar_figura(fig, "Porcentaje de nulos por columna", 420, leyenda=False))
        mostrar_tabla(nulos, columnas_num=["Nulos"])
    with t3:
        if ctx["reporte"].empty:
            st.info("No se realizaron conversiones de tipo.")
        else:
            st.dataframe(ctx["reporte"], hide_index=True)
        if cal["negativos"]:
            st.markdown("**Valores negativos por columna monetaria**")
            st.dataframe(pd.DataFrame({"Columna": list(cal["negativos"].keys()),
                                       "Valores negativos": list(cal["negativos"].values())}), hide_index=True)
    with t4:
        c = st.columns(3)
        c[0].metric("Registros iniciales", formatear_numero(ctx["n_inicial"]))
        c[1].metric("Registros posteriores a la limpieza", formatear_numero(ctx["n_final"]))
        c[2].metric("Registros excluidos", formatear_numero(ctx["n_inicial"] - ctx["n_final"]))
        if not st.session_state["limp_aplicar"]:
            st.info("La limpieza está desactivada: se analizan los datos originales (con conversión de tipos).")
        if ctx["bitacora"]:
            st.dataframe(pd.DataFrame(ctx["bitacora"]), hide_index=True)
        st.caption("La limpieza se aplica sobre una copia. Los datos originales permanecen en memoria durante la "
                   "sesión y pueden restaurarse desde la barra lateral.")


def _subtexto_comparacion(comp, clave, tipo):
    """Texto de comparación del último periodo frente al anterior para una tarjeta KPI."""
    if comp is None:
        return "Sin periodo anterior para comparar."
    if comp.get("anterior") is None:
        return f"Último periodo {comp['ultimo']}: sin periodo anterior consecutivo."
    actual, previo = comp["k_ultimo"].get(clave), comp["k_anterior"].get(clave)
    if not es_valor_valido(actual):
        return "No calculable en el último periodo."
    if tipo == "moneda":
        valor = formatear_moneda_corta(actual)
    elif tipo == "pct":
        valor = formatear_porcentaje(actual)
    elif tipo == "dias":
        valor = f"{formatear_numero(actual, 1)} días"
    else:
        valor = formatear_numero(actual)
    if not es_valor_valido(previo):
        return f"{comp['ultimo']}: {valor} (anterior no calculable)."
    if tipo in ("pct",):
        diferencia = actual - previo
        flecha = "▲" if diferencia > 0 else "▼" if diferencia < 0 else "="
        return f"{comp['ultimo']}: {valor} ({flecha} {formatear_numero(abs(diferencia), 1)} p.p. vs {comp['anterior']})"
    if tipo == "dias":
        diferencia = actual - previo
        flecha = "▲" if diferencia > 0 else "▼" if diferencia < 0 else "="
        return f"{comp['ultimo']}: {valor} ({flecha} {formatear_numero(abs(diferencia), 1)} días vs {comp['anterior']})"
    variacion = dividir_seguro(actual - previo, abs(previo), 100)
    if not es_valor_valido(variacion):
        return f"{comp['ultimo']}: {valor} (variación no calculable por división entre cero)."
    flecha = "▲" if variacion > 0 else "▼" if variacion < 0 else "="
    return f"{comp['ultimo']}: {valor} ({flecha} {formatear_porcentaje(abs(variacion))} vs {comp['anterior']})"


def mostrar_resumen(ctx):
    """Página 3: tarjetas KPI y gráficos ejecutivos."""
    st.subheader("Resumen ejecutivo")
    df = ctx["df"]
    k = calcular_kpis(df)
    comp = comparar_ultimo_periodo(df)
    u = obtener_umbrales()
    sem_glosa = evaluar_semaforo(k["pct_glosa"], u["glosa_verde"], u["glosa_amarillo"])[0]
    sem_recaudo = evaluar_semaforo(k["pct_recaudo"], u["recaudo_verde"], u["recaudo_amarillo"], True)[0]
    sem_dias = evaluar_semaforo(k["dias"], u["dias_verde"], u["dias_amarillo"])[0]
    tarjetas = [
        ("Facturación bruta", formatear_moneda_corta(k["bruto"]), "bruto", "moneda", "", "Suma de Valor_Bruto."),
        ("Facturación neta", formatear_moneda_corta(k["neto"]), "neto", "moneda", "", "Suma de Valor_Neto."),
        ("Número de facturas", formatear_numero(k["facturas"]), "facturas", "num", "", "ID_Factura únicos."),
        ("Cantidad de servicios", formatear_numero(k["cantidad"]), "cantidad", "num", "", "Suma de Cantidad."),
        ("Valor pagado", formatear_moneda_corta(k["pagado"]), "pagado", "moneda", "", "Suma de Valor_Pagado."),
        ("Saldo de cartera", formatear_moneda_corta(k["saldo"]), "saldo", "moneda", "", "Suma de Saldo_Cartera."),
        ("Valor glosado", formatear_moneda_corta(k["glosa"]), "glosa", "moneda", "", "Suma de Glosa_Inicial."),
        ("% de recaudo", formatear_porcentaje(k["pct_recaudo"]), "pct_recaudo", "pct", sem_recaudo,
         "Valor_Pagado / Valor_Neto × 100."),
        ("% de glosa", formatear_porcentaje(k["pct_glosa"]), "pct_glosa", "pct", sem_glosa,
         "Glosa_Inicial / Valor_Neto × 100."),
        ("Ticket promedio", formatear_moneda_corta(k["ticket"]), "ticket", "moneda", "",
         "Valor_Neto / número de facturas."),
        ("Tiempo promedio de radicación", f"{formatear_numero(k['dias'], 1)} días" if es_valor_valido(k["dias"])
         else "No calculable", "dias", "dias", sem_dias, "Promedio de Dias_Emision_Radicacion."),
    ]
    columnas = st.columns(4)
    for i, (titulo, valor, clave, tipo, estado, ayuda) in enumerate(tarjetas):
        with columnas[i % 4]:
            sub = _subtexto_comparacion(comp, clave, tipo) if valor != "No calculable" else \
                "No es posible calcular el indicador con las columnas disponibles."
            if estado:
                sub += f" · Semáforo: {ICONO_SEMAFORO[estado]} {estado.capitalize()}"
            tarjeta_kpi(titulo, valor, sub, estado if estado else ("gris" if valor == "No calculable" else ""), ayuda)
    st.caption("Las tarjetas usan el conjunto filtrado. La comparación se hace entre el último periodo y el "
               "inmediatamente anterior; el último periodo puede estar incompleto.")

    cierre = construir_cierre_mensual(df, "Mensual")
    if cierre.empty:
        st.info("No hay periodos disponibles para los gráficos temporales.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.line(cierre, x="Periodo", y="Valor neto", markers=True,
                          color_discrete_sequence=[COLOR_AZUL_OSCURO])
            mostrar_grafico(estilizar_figura(fig, "1. Facturación neta mensual", moneda_y=True, leyenda=False))
        with c2:
            series = [c for c in ["Valor neto", "Valor pagado", "Saldo de cartera"] if c in cierre.columns]
            largo = cierre.melt(id_vars="Periodo", value_vars=series, var_name="Concepto", value_name="Valor")
            fig = px.bar(largo, x="Periodo", y="Valor", color="Concepto", barmode="group",
                         color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_VERDE_AZULADO, COLOR_NARANJA])
            mostrar_grafico(estilizar_figura(fig, "2. Facturación, pagos y cartera por periodo", moneda_y=True))
    c1, c2 = st.columns(2)
    with c1:
        if "Estado_Factura" in df.columns:
            estados = df["Estado_Factura"].astype("string").fillna(SIN_INFO).value_counts().reset_index()
            estados.columns = ["Estado", "Registros"]
            fig = px.pie(estados, names="Estado", values="Registros", hole=0.45, color_discrete_sequence=PALETA)
            fig.update_traces(textinfo="percent+label")
            mostrar_grafico(estilizar_figura(fig, "3. Distribución de estados"))
        else:
            st.info("Distribución de estados no disponible: falta Estado_Factura.")
    with c2:
        if "Entidad_Pagadora" in df.columns:
            ent = agregar_por(df, "Entidad_Pagadora").head(10)
            fig = px.bar(ent, x="Valor neto", y="Entidad_Pagadora", orientation="h",
                         color_discrete_sequence=[COLOR_AZUL_MEDIO], hover_data={"Participación %": ":.1f"})
            fig.update_yaxes(autorange="reversed", title=None)
            fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            mostrar_grafico(estilizar_figura(fig, "4. Facturación por entidad pagadora (top 10)", leyenda=False))
        else:
            st.info("Facturación por entidad no disponible: falta Entidad_Pagadora.")
    c1, c2 = st.columns(2)
    with c1:
        if "Linea_Servicio" in df.columns:
            lineas = agregar_por(df, "Linea_Servicio")
            fig = px.pie(lineas.head(12), names="Linea_Servicio", values="Valor neto", hole=0.45,
                         color_discrete_sequence=PALETA)
            mostrar_grafico(estilizar_figura(fig, "5. Participación por línea de servicio"))
        else:
            st.info("Participación por línea no disponible: falta Linea_Servicio.")
    with c2:
        if not cierre.empty and "Valor glosado" in cierre.columns:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_bar(x=cierre["Periodo"], y=cierre["Valor glosado"], name="Valor glosado",
                        marker_color=COLOR_NARANJA)
            fig.add_scatter(x=cierre["Periodo"], y=cierre["% glosa"], name="% glosa", mode="lines+markers",
                            line_color=COLOR_ROJO, secondary_y=True)
            fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
            fig.update_yaxes(ticksuffix=" %", secondary_y=True)
            mostrar_grafico(estilizar_figura(fig, "6. Glosa por periodo"))
        else:
            st.info("Glosa por periodo no disponible: falta Glosa_Inicial o periodo.")
    if not cierre.empty and "Valor pagado" in cierre.columns:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=cierre["Periodo"], y=cierre["Valor pagado"], name="Valor pagado", marker_color=COLOR_VERDE_AZULADO)
        fig.add_scatter(x=cierre["Periodo"], y=cierre["% recaudo"], name="% recaudo", mode="lines+markers",
                        line_color=COLOR_AZUL_OSCURO, secondary_y=True)
        fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
        fig.update_yaxes(ticksuffix=" %", secondary_y=True)
        mostrar_grafico(estilizar_figura(fig, "7. Recaudo por periodo"))
    else:
        st.info("Recaudo por periodo no disponible: falta Valor_Pagado o periodo.")


def mostrar_cierre(ctx):
    """Página 4: cierre mensual, trimestral o anual con crecimiento, acumulado y promedio móvil."""
    st.subheader("Cierre mensual")
    df = ctx["df"]
    periodicidad = st.radio("Periodicidad", ["Mensual", "Trimestral", "Anual"], horizontal=True)
    cierre = construir_cierre_mensual(df, periodicidad)
    if cierre.empty:
        st.warning("No hay periodos válidos para construir el cierre.")
        return
    visibles = [c for c in cierre.columns if c not in ("_ordinal", "Registros con saldo", "Registros devueltos")]
    tabla = cierre[visibles]
    mostrar_tabla(tabla, columnas_num=["Facturas", "Registros", "Cantidad", "Registros con glosa"],
                  columnas_pct=["% soportes incompletos"])
    st.caption("El crecimiento y el promedio móvil solo se calculan entre periodos consecutivos. El acumulado anual "
               "se reinicia cada año.")

    if len(cierre) >= 2:
        st.markdown("#### Periodos destacados")
        destacados = []
        for columna, texto, mayor in [("Valor neto", "Mayor facturación", True),
                                      ("Valor neto", "Menor facturación", False),
                                      ("Valor pagado", "Mayor recaudo", True),
                                      ("Saldo de cartera", "Mayor cartera", True),
                                      ("Valor glosado", "Mayor glosa", True)]:
            if columna in cierre.columns and cierre[columna].notna().any():
                fila = cierre.loc[cierre[columna].idxmax() if mayor else cierre[columna].idxmin()]
                destacados.append((texto, fila["Periodo"], formatear_moneda_corta(fila[columna])))
        cols = st.columns(len(destacados) + 1)
        for i, (texto, periodo, valor) in enumerate(destacados):
            with cols[i]:
                tarjeta_kpi(texto, str(periodo), valor)
        with cols[-1]:
            ultimo = cierre.iloc[-1]
            crec = ultimo.get("Crecimiento %")
            if es_valor_valido(crec):
                flecha = "▲" if crec > 0 else "▼" if crec < 0 else "="
                tarjeta_kpi("Último vs anterior", f"{flecha} {formatear_porcentaje(abs(crec))}",
                            f"{cierre.iloc[-2]['Periodo']} → {ultimo['Periodo']}")
            else:
                tarjeta_kpi("Último vs anterior", "No comparable",
                            "No existe un periodo anterior consecutivo.", "gris")

    t1, t2, t3, t4 = st.tabs(["Tendencias", "Facturación y crecimiento", "Comparación entre años", "Promedio móvil"])
    with t1:
        c1, c2, c3 = st.columns(3)
        for contenedor, columna, titulo, color in [(c1, "Valor neto", "Facturación neta", COLOR_AZUL_OSCURO),
                                                   (c2, "Valor pagado", "Recaudo", COLOR_VERDE_AZULADO),
                                                   (c3, "Saldo de cartera", "Cartera", COLOR_NARANJA)]:
            with contenedor:
                if columna in cierre.columns:
                    fig = px.line(cierre, x="Periodo", y=columna, markers=True, color_discrete_sequence=[color])
                    mostrar_grafico(estilizar_figura(fig, titulo, 340, moneda_y=True, leyenda=False))
                else:
                    st.info(f"{titulo}: columna no disponible.")
    with t2:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=cierre["Periodo"], y=cierre["Valor neto"], name="Valor neto", marker_color=COLOR_AZUL_MEDIO)
        if "Crecimiento %" in cierre:
            fig.add_scatter(x=cierre["Periodo"], y=cierre["Crecimiento %"], name="Crecimiento %",
                            mode="lines+markers", line_color=COLOR_NARANJA, secondary_y=True)
        fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
        fig.update_yaxes(ticksuffix=" %", secondary_y=True)
        mostrar_grafico(estilizar_figura(fig, "Facturación neta y crecimiento"))
    with t3:
        if periodicidad != "Mensual":
            st.info("La comparación entre años se presenta con periodicidad mensual.")
        mensual = construir_cierre_mensual(df, "Mensual")
        if mensual.empty or mensual["Periodo"].str.slice(0, 4).nunique() < 2:
            st.info("Se requieren al menos dos años con datos para comparar.")
        else:
            mensual["Año"] = mensual["Periodo"].str.slice(0, 4)
            mensual["Número de mes"] = mensual["Periodo"].str.slice(5, 7).astype(int)
            mensual["Mes"] = mensual["Número de mes"].map(MESES_ES)
            mensual = mensual.sort_values(["Año", "Número de mes"])
            fig = px.line(mensual, x="Mes", y="Valor neto", color="Año", markers=True,
                          category_orders={"Mes": list(MESES_ES.values())}, color_discrete_sequence=PALETA)
            mostrar_grafico(estilizar_figura(fig, "Facturación neta por mes y año", moneda_y=True))
    with t4:
        if "Promedio móvil 3" in cierre and cierre["Promedio móvil 3"].notna().any():
            fig = go.Figure()
            fig.add_scatter(x=cierre["Periodo"], y=cierre["Valor neto"], name="Valor neto", mode="lines+markers",
                            line_color=COLOR_AZUL_MEDIO)
            fig.add_scatter(x=cierre["Periodo"], y=cierre["Promedio móvil 3"], name="Promedio móvil 3 periodos",
                            mode="lines", line={"color": COLOR_AZUL_OSCURO, "dash": "dash", "width": 3})
            mostrar_grafico(estilizar_figura(fig, "Promedio móvil de tres periodos", moneda_y=True))
        else:
            st.info("Se requieren al menos tres periodos consecutivos para el promedio móvil.")


def grafico_pareto(tabla, dimension, titulo):
    """Gráfico combinado de Pareto: barras de valor, línea acumulada y referencia del 80 %."""
    datos = tabla.head(30)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    colores = [COLOR_AZUL_OSCURO if g == "Explica ~80 %" else "#A3A9B0" for g in datos["Grupo Pareto"]]
    fig.add_bar(x=datos[dimension], y=datos["Valor"], name="Valor", marker_color=colores)
    fig.add_scatter(x=datos[dimension], y=datos["Participación acumulada %"], name="% acumulado",
                    mode="lines+markers", line_color=COLOR_NARANJA, secondary_y=True)
    fig.add_hline(y=80, line_dash="dot", line_color=COLOR_ROJO, secondary_y=True,
                  annotation_text="Referencia 80 %", annotation_position="top left")
    fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
    fig.update_yaxes(range=[0, 105], ticksuffix=" %", secondary_y=True)
    fig.update_xaxes(tickangle=-35)
    return estilizar_figura(fig, titulo, 460)


def mostrar_pareto(df, clave="pareto"):
    """Bloque de análisis de concentración y Pareto (facturación, cartera, glosas y servicios)."""
    opciones = {}
    if "Entidad_Pagadora" in df.columns:
        opciones["Facturación por entidad"] = ("Entidad_Pagadora", "Valor_Neto")
        if columna_util(df, "Saldo_Cartera"):
            opciones["Cartera por entidad"] = ("Entidad_Pagadora", "Saldo_Cartera")
        if columna_util(df, "Glosa_Inicial"):
            opciones["Glosas por entidad"] = ("Entidad_Pagadora", "Glosa_Inicial")
    if "Servicio" in df.columns:
        opciones["Facturación por servicio"] = ("Servicio", "Valor_Neto")
    if not opciones:
        st.info("El análisis de Pareto requiere Entidad_Pagadora o Servicio.")
        return
    eleccion = st.selectbox("Análisis de concentración", list(opciones), key=clave)
    dimension, valor = opciones[eleccion]
    tabla, resumen = calcular_pareto(df, dimension, valor)
    if tabla.empty:
        st.info("No hay valores positivos para construir el Pareto.")
        return
    c = st.columns(5)
    c[0].metric("Categorías con valor", formatear_numero(resumen["n_categorias"]))
    c[1].metric("Explican ~80 %", f"{resumen['n_80']} ({formatear_porcentaje(resumen['pct_categorias_80'])})")
    c[2].metric("Participación top 1", formatear_porcentaje(resumen["top1"]))
    c[3].metric("Participación top 3", formatear_porcentaje(resumen["top3"]))
    c[4].metric("Índice HHI", formatear_numero(resumen["hhi"]), help="Herfindahl-Hirschman: suma de cuotas al "
                "cuadrado × 10.000. Lectura orientativa: < 1.500 baja, 1.500-2.500 moderada, > 2.500 alta.")
    mostrar_grafico(grafico_pareto(tabla, dimension, f"Pareto: {eleccion.lower()}"), clave=f"{clave}_fig")
    st.caption(f"Concentración {interpretar_hhi(resumen['hhi'])} según el HHI. El 80 % se usa como herramienta de "
               "análisis de Pareto, no como una regla obligatoria. La concentración de facturación, de cartera y de "
               "glosas son fenómenos distintos y deben leerse por separado.")
    with st.expander("Tabla de Pareto"):
        mostrar_tabla(tabla.rename(columns={"Valor": f"Valor ({valor})"}), columnas_moneda=[f"Valor ({valor})"])


def mostrar_entidades(ctx):
    """Página 5: análisis por entidad pagadora, rankings, dispersión, mapa de calor, detalle y Pareto."""
    st.subheader("Entidades pagadoras")
    df = ctx["df"]
    if "Entidad_Pagadora" not in df.columns:
        st.warning("No se encontró la columna necesaria para este análisis: Entidad_Pagadora.")
        return
    tabla = agregar_por(df, "Entidad_Pagadora")
    t1, t2, t3, t4, t5 = st.tabs(["Tabla comparativa", "Rankings", "Dispersión y mapa de calor",
                                  "Detalle de entidad", "Concentración y Pareto"])
    with t1:
        mostrar_tabla(tabla, columnas_num=["Facturas", "Registros", "Cantidad", "Registros con glosa",
                                           "Registros con saldo", "Registros devueltos",
                                           "Días promedio radicación"], altura=480)
        total_part = tabla["Participación %"].sum() if "Participación %" in tabla else np.nan
        st.caption(f"Suma de participaciones: {formatear_porcentaje(total_part, 2)} (puede diferir de 100 % por "
                   "redondeo o valores negativos).")
    with t2:
        n = selector_top(clave="top_entidades")
        metricas = [("Valor neto", "Ranking por facturación", COLOR_AZUL_OSCURO),
                    ("Saldo de cartera", "Ranking por cartera", COLOR_NARANJA),
                    ("Valor glosado", "Ranking por glosa", COLOR_ROJO),
                    ("% recaudo", "Ranking por recaudo (%)", COLOR_VERDE_AZULADO)]
        columnas = st.columns(2)
        for i, (metrica, titulo, color) in enumerate([m for m in metricas if m[0] in tabla.columns]):
            datos = _limitar_categorias(tabla.sort_values(metrica, ascending=False), n)
            fig = px.bar(datos, x=metrica, y="Entidad_Pagadora", orientation="h", color_discrete_sequence=[color],
                         hover_data={"Facturas": True} if "Facturas" in datos else None)
            fig.update_yaxes(autorange="reversed", title=None)
            if metrica.startswith("%"):
                fig.update_xaxes(ticksuffix=" %")
            else:
                fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            with columnas[i % 2]:
                mostrar_grafico(estilizar_figura(fig, titulo, max(320, 28 * len(datos) + 80), leyenda=False),
                                clave=f"rank_ent_{i}")
    with t3:
        if {"% recaudo", "Saldo de cartera"}.issubset(tabla.columns):
            datos = tabla[tabla["Entidad_Pagadora"] != SIN_INFO].copy()
            datos["Tamaño"] = datos["Saldo de cartera"].clip(lower=0).fillna(0) + 1
            fig = px.scatter(datos, x="Valor neto", y="% recaudo", size="Tamaño",
                             color="% glosa" if "% glosa" in datos else None,
                             color_continuous_scale=ESCALA_ALERTA, hover_name="Entidad_Pagadora",
                             hover_data={"Facturas": True if "Facturas" in datos else False,
                                         "Ticket promedio": ":,.0f", "Tamaño": False, "Saldo de cartera": ":,.0f"},
                             size_max=45)
            fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            fig.update_yaxes(ticksuffix=" %")
            mostrar_grafico(estilizar_figura(fig, "Facturación neta vs % de recaudo (tamaño: cartera; color: % glosa)",
                                             480))
        else:
            st.info("La dispersión requiere Valor_Pagado y Saldo_Cartera.")
        if "Periodo" in df.columns:
            principales = tabla.head(20)["Entidad_Pagadora"].tolist()
            base = df[_serie_grupo(df, "Entidad_Pagadora").isin(principales) & df["Periodo"].notna()]
            if not base.empty:
                matriz = base.pivot_table(index=_serie_grupo(base, "Entidad_Pagadora"), columns="Periodo",
                                          values="Valor_Neto", aggfunc="sum", fill_value=0)
                matriz = matriz.reindex(columns=sorted(matriz.columns))
                fig = px.imshow(matriz, aspect="auto", color_continuous_scale=ESCALA_CONTINUA,
                                labels={"color": "Valor neto", "x": "Periodo", "y": "Entidad"})
                mostrar_grafico(estilizar_figura(fig, "Mapa de calor: facturación neta por entidad y periodo (top 20)",
                                                 max(380, 26 * len(matriz) + 120)))
    with t4:
        entidades = tabla["Entidad_Pagadora"].tolist()
        seleccion = st.selectbox("Seleccione una entidad", entidades, key="det_entidad")
        sub = df[_serie_grupo(df, "Entidad_Pagadora") == seleccion]
        k = calcular_kpis(sub)
        c = st.columns(5)
        c[0].metric("Facturación neta", formatear_moneda_corta(k["neto"]))
        c[1].metric("Facturas", formatear_numero(k["facturas"]))
        c[2].metric("% recaudo", formatear_porcentaje(k["pct_recaudo"]))
        c[3].metric("% glosa", formatear_porcentaje(k["pct_glosa"]))
        c[4].metric("Saldo de cartera", formatear_moneda_corta(k["saldo"]))
        evolucion = construir_cierre_mensual(sub, "Mensual")
        c1, c2 = st.columns(2)
        with c1:
            if not evolucion.empty:
                series = [s for s in ["Valor neto", "Valor pagado", "Saldo de cartera"] if s in evolucion]
                fig = px.line(evolucion, x="Periodo", y=series, markers=True,
                              color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_VERDE_AZULADO, COLOR_NARANJA])
                mostrar_grafico(estilizar_figura(fig, f"Evolución de {seleccion}", moneda_y=True))
        with c2:
            if "Servicio" in sub.columns:
                serv = agregar_por(sub, "Servicio").head(10)
                fig = px.bar(serv, x="Valor neto", y="Servicio", orientation="h",
                             color_discrete_sequence=[COLOR_AZUL_MEDIO])
                fig.update_yaxes(autorange="reversed", title=None)
                fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
                mostrar_grafico(estilizar_figura(fig, "Servicios principales de la entidad", leyenda=False))
    with t5:
        mostrar_pareto(df, "pareto_entidades")


def mostrar_servicios(ctx):
    """Página 6: análisis por servicio y línea de servicio."""
    st.subheader("Servicios y líneas de servicio")
    df = ctx["df"]
    dimensiones = [d for d in ["Servicio", "Linea_Servicio"] if d in df.columns]
    if not dimensiones:
        st.warning("No se encontró la columna necesaria para este análisis: Servicio o Linea_Servicio.")
        return
    c1, c2 = st.columns([1, 2])
    dimension = c1.radio("Dimensión", dimensiones, horizontal=True,
                         format_func=lambda d: "Servicio" if d == "Servicio" else "Línea de servicio")
    with c2:
        n = selector_top(clave="top_servicios")
    tabla = agregar_por(df, dimension)
    if len(tabla) > 40 and n is None:
        st.info(f"Hay {len(tabla)} categorías; los gráficos muestran las 40 principales para mantener la legibilidad.")
    datos = _limitar_categorias(tabla, n)
    t1, t2, t3, t4, t5 = st.tabs(["Tabla y barras", "Treemap y burbujas", "Distribución (boxplot)",
                                  "Servicio vs entidad", "Evolución mensual"])
    with t1:
        fig = px.bar(datos, x="Valor neto", y=dimension, orientation="h", color_discrete_sequence=[COLOR_AZUL_MEDIO],
                     hover_data={"Participación %": ":.1f"})
        fig.update_yaxes(autorange="reversed", title=None)
        fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
        mostrar_grafico(estilizar_figura(fig, f"Facturación neta por {dimension.lower().replace('_', ' ')}",
                                         max(340, 26 * len(datos) + 80), leyenda=False))
        mostrar_tabla(tabla, columnas_num=["Facturas", "Registros", "Cantidad", "Registros con glosa",
                                           "Registros con saldo", "Registros devueltos",
                                           "Días promedio radicación"])
    with t2:
        if {"Linea_Servicio", "Servicio"}.issubset(df.columns):
            arbol = agregar_por(df, ["Linea_Servicio", "Servicio"])
            arbol = arbol[arbol["Valor neto"] > 0]
            if not arbol.empty:
                fig = px.treemap(arbol, path=[px.Constant("Total"), "Linea_Servicio", "Servicio"], values="Valor neto",
                                 color="% glosa" if "% glosa" in arbol else None,
                                 color_continuous_scale=ESCALA_ALERTA)
                fig.update_traces(hovertemplate="%{label}<br>Valor neto: " + _simbolo() + "%{value:,.0f}")
                mostrar_grafico(estilizar_figura(fig, "Treemap: línea y servicio (color: % glosa)", 480))
        else:
            positivo = datos[datos["Valor neto"] > 0]
            if not positivo.empty:
                fig = px.treemap(positivo, path=[dimension], values="Valor neto", color_discrete_sequence=PALETA)
                mostrar_grafico(estilizar_figura(fig, "Treemap de facturación neta", 460))
        if "Facturas" in datos.columns:
            burbujas = datos.copy()
            burbujas["Tamaño"] = burbujas["Valor neto"].clip(lower=0).fillna(0) + 1
            fig = px.scatter(burbujas, x="Facturas", y="Ticket promedio", size="Tamaño",
                             color="% glosa" if "% glosa" in burbujas else None, hover_name=dimension,
                             color_continuous_scale=ESCALA_ALERTA, size_max=50,
                             hover_data={"Tamaño": False, "Valor neto": ":,.0f"})
            fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f")
            mostrar_grafico(estilizar_figura(fig, "Burbujas: facturas vs ticket promedio (tamaño: valor neto)", 460))
    with t3:
        categorias = datos[dimension].tolist()
        base = df[_serie_grupo(df, dimension).isin(categorias)]
        if not base.empty:
            fig = px.box(base.assign(**{dimension: _serie_grupo(base, dimension)}), x=dimension, y="Valor_Neto",
                         color_discrete_sequence=[COLOR_AZUL_MEDIO], category_orders={dimension: categorias})
            fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f")
            fig.update_xaxes(tickangle=-35, title=None)
            mostrar_grafico(estilizar_figura(fig, "Distribución del valor neto por categoría", 480, leyenda=False))
    with t4:
        if "Entidad_Pagadora" in df.columns:
            principales_ent = agregar_por(df, "Entidad_Pagadora").head(15)["Entidad_Pagadora"].tolist()
            categorias = datos[dimension].head(20).tolist()
            base = df[_serie_grupo(df, dimension).isin(categorias) &
                      _serie_grupo(df, "Entidad_Pagadora").isin(principales_ent)]
            if not base.empty:
                matriz = base.pivot_table(index=_serie_grupo(base, dimension),
                                          columns=_serie_grupo(base, "Entidad_Pagadora"), values="Valor_Neto",
                                          aggfunc="sum", fill_value=0)
                fig = px.imshow(matriz, aspect="auto", color_continuous_scale=ESCALA_CONTINUA,
                                labels={"color": "Valor neto", "x": "Entidad", "y": dimension})
                mostrar_grafico(estilizar_figura(fig, "Mapa de calor: valor neto por categoría y entidad",
                                                 max(400, 26 * len(matriz) + 140)))
        else:
            st.info("Requiere Entidad_Pagadora.")
    with t5:
        seleccion = st.selectbox("Seleccione una categoría", tabla[dimension].tolist(), key="evol_servicio")
        sub = df[_serie_grupo(df, dimension) == seleccion]
        evolucion = construir_cierre_mensual(sub, "Mensual")
        if evolucion.empty:
            st.info("No hay periodos para la categoría seleccionada.")
        else:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_bar(x=evolucion["Periodo"], y=evolucion["Valor neto"], name="Valor neto",
                        marker_color=COLOR_AZUL_MEDIO)
            if "Cantidad" in evolucion:
                fig.add_scatter(x=evolucion["Periodo"], y=evolucion["Cantidad"], name="Cantidad",
                                mode="lines+markers", line_color=COLOR_NARANJA, secondary_y=True)
            fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
            mostrar_grafico(estilizar_figura(fig, f"Evolución mensual: {seleccion}"))


def mostrar_sedes_regimenes(ctx):
    """Página 7: análisis por sede y por régimen (solo si las columnas existen)."""
    st.subheader("Sedes y regímenes")
    df = ctx["df"]
    dimensiones = [d for d in ["Sede", "Regimen"] if d in df.columns and df[d].notna().any()]
    if not dimensiones:
        st.warning("No se encontraron las columnas Sede ni Regimen. Esta sección se desactiva.")
        return
    for faltante in [d for d in ["Sede", "Regimen"] if d not in dimensiones]:
        st.info(f"La columna {faltante} no está disponible; se omite su análisis.")
    pestanas = st.tabs(["Sede" if d == "Sede" else "Régimen" for d in dimensiones])
    for pestana, dimension in zip(pestanas, dimensiones):
        with pestana:
            tabla = agregar_por(df, dimension)
            mostrar_tabla(tabla, columnas_num=["Facturas", "Registros", "Cantidad", "Registros con glosa",
                                               "Registros con saldo", "Registros devueltos",
                                               "Días promedio radicación"])
            datos = tabla.head(25)
            c1, c2 = st.columns(2)
            with c1:
                series = [s for s in ["Valor neto", "Valor pagado", "Saldo de cartera", "Valor glosado"]
                          if s in datos.columns]
                largo = datos.melt(id_vars=dimension, value_vars=series, var_name="Concepto", value_name="Valor")
                fig = px.bar(largo, x=dimension, y="Valor", color="Concepto", barmode="group",
                             color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_VERDE_AZULADO, COLOR_NARANJA,
                                                      COLOR_ROJO])
                fig.update_xaxes(title=None)
                mostrar_grafico(estilizar_figura(fig, "Valores por grupo (barras agrupadas)", moneda_y=True),
                                clave=f"agr_{dimension}")
            with c2:
                if "Estado_Factura" in df.columns:
                    cruce = pd.crosstab(_serie_grupo(df, dimension), _serie_grupo(df, "Estado_Factura"),
                                        normalize="index") * 100
                    largo = cruce.reset_index().melt(id_vars=dimension, var_name="Estado", value_name="Porcentaje")
                    fig = px.bar(largo, x=dimension, y="Porcentaje", color="Estado", color_discrete_sequence=PALETA)
                    fig.update_layout(barmode="stack")
                    fig.update_xaxes(title=None)
                    mostrar_grafico(estilizar_figura(fig, "Composición por estado (barras al 100 %)", pct_y=True),
                                    clave=f"pct_{dimension}")
                elif {"Valor pagado", "Saldo de cartera"}.issubset(datos.columns):
                    comp = datos[[dimension, "Valor pagado", "Saldo de cartera"]].copy()
                    total = comp[["Valor pagado", "Saldo de cartera"]].sum(axis=1)
                    comp["Pagado %"] = dividir_seguro(comp["Valor pagado"], total, 100).values
                    comp["Saldo %"] = dividir_seguro(comp["Saldo de cartera"], total, 100).values
                    largo = comp.melt(id_vars=dimension, value_vars=["Pagado %", "Saldo %"], var_name="Concepto",
                                      value_name="Porcentaje")
                    fig = px.bar(largo, x=dimension, y="Porcentaje", color="Concepto",
                                 color_discrete_sequence=[COLOR_VERDE_AZULADO, COLOR_NARANJA])
                    fig.update_layout(barmode="stack")
                    mostrar_grafico(estilizar_figura(fig, "Composición pagado / saldo (barras al 100 %)", pct_y=True),
                                    clave=f"pct_{dimension}")
            c1, c2 = st.columns(2)
            with c1:
                principales = datos[dimension].tolist()
                base = df[_serie_grupo(df, dimension).isin(principales)]
                fig = px.box(base.assign(**{dimension: _serie_grupo(base, dimension)}), x=dimension, y="Valor_Neto",
                             color_discrete_sequence=[COLOR_AZUL_MEDIO])
                fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f")
                fig.update_xaxes(title=None)
                mostrar_grafico(estilizar_figura(fig, "Distribución del valor neto", leyenda=False),
                                clave=f"box_{dimension}")
            with c2:
                if "Periodo" in df.columns:
                    base = df[df["Periodo"].notna()]
                    matriz = base.pivot_table(index=_serie_grupo(base, dimension), columns="Periodo",
                                              values="Valor_Neto", aggfunc="sum", fill_value=0)
                    matriz = matriz.reindex(columns=sorted(matriz.columns))
                    fig = px.imshow(matriz, aspect="auto", color_continuous_scale=ESCALA_CONTINUA,
                                    labels={"color": "Valor neto", "x": "Periodo", "y": dimension})
                    mostrar_grafico(estilizar_figura(fig, "Mapa de calor por periodo"), clave=f"heat_{dimension}")


def mostrar_estados(ctx):
    """Página 8: estados de facturación detectados dinámicamente."""
    st.subheader("Estados de facturación")
    df = ctx["df"]
    if "Estado_Factura" not in df.columns:
        st.warning("No se encontró la columna necesaria para este análisis: Estado_Factura.")
        return
    estado = _serie_grupo(df, "Estado_Factura")
    norm = normalizar_serie_texto(estado)
    categorias = {
        "Pagadas": norm.str.contains("pagad", na=False) & ~norm.str.contains("parcial", na=False),
        "Parcialmente pagadas": norm.str.contains("parcial", na=False),
        "Radicadas": norm.str.contains("radicad", na=False),
        "Aceptadas": norm.str.contains("aceptad", na=False),
        "Glosadas": norm.str.contains("glosad", na=False),
        "Devueltas": norm.str.contains("devuel", na=False),
    }
    presentes = {k: v for k, v in categorias.items() if v.any()}
    if presentes:
        columnas = st.columns(len(presentes))
        for (nombre, mascara), contenedor in zip(presentes.items(), columnas):
            with contenedor:
                n_fact = df.loc[mascara, "ID_Factura"].nunique() if columna_util(df, "ID_Factura") else int(mascara.sum())
                tarjeta_kpi(nombre, formatear_numero(n_fact),
                            f"{formatear_porcentaje(mascara.mean() * 100)} de registros · "
                            f"{formatear_moneda_corta(df.loc[mascara, 'Valor_Neto'].sum())}")
    else:
        st.info("No se identificaron los estados típicos (pagada, radicada, glosada...). Se muestran las categorías "
                "presentes en el archivo.")
    tabla = agregar_por(df, "Estado_Factura")
    tabla["% registros"] = dividir_seguro(tabla["Registros"], pd.Series([tabla["Registros"].sum()] * len(tabla)),
                                          100).values
    columnas_tabla = [c for c in ["Estado_Factura", "Facturas", "Registros", "% registros", "Valor neto",
                                  "Valor pagado", "Saldo de cartera", "Valor glosado", "Participación %"]
                      if c in tabla.columns]
    mostrar_tabla(tabla[columnas_tabla], columnas_num=["Facturas", "Registros"])
    c1, c2 = st.columns(2)
    with c1:
        series = [s for s in ["Valor neto", "Valor pagado", "Saldo de cartera"] if s in tabla.columns]
        largo = tabla.melt(id_vars="Estado_Factura", value_vars=series, var_name="Concepto", value_name="Valor")
        fig = px.bar(largo, x="Estado_Factura", y="Valor", color="Concepto", barmode="group",
                     color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_VERDE_AZULADO, COLOR_NARANJA])
        fig.update_xaxes(title=None)
        mostrar_grafico(estilizar_figura(fig, "Valores por estado", moneda_y=True))
    with c2:
        fig = px.bar(tabla, x="Estado_Factura", y="Registros", color="Estado_Factura", color_discrete_sequence=PALETA,
                     text="Registros")
        fig.update_xaxes(title=None)
        mostrar_grafico(estilizar_figura(fig, "Registros por estado", leyenda=False))
    if "Periodo" in df.columns:
        base = df[df["Periodo"].notna()]
        cruce = pd.crosstab(base["Periodo"], _serie_grupo(base, "Estado_Factura")).sort_index()
        c1, c2 = st.columns(2)
        with c1:
            largo = cruce.reset_index().melt(id_vars="Periodo", var_name="Estado", value_name="Registros")
            fig = px.bar(largo, x="Periodo", y="Registros", color="Estado", color_discrete_sequence=PALETA)
            fig.update_layout(barmode="stack")
            mostrar_grafico(estilizar_figura(fig, "Estados por periodo (apiladas)"))
        with c2:
            porcentaje = cruce.div(cruce.sum(axis=1).replace(0, np.nan), axis=0) * 100
            largo = porcentaje.reset_index().melt(id_vars="Periodo", var_name="Estado", value_name="Porcentaje")
            fig = px.bar(largo, x="Periodo", y="Porcentaje", color="Estado", color_discrete_sequence=PALETA)
            fig.update_layout(barmode="stack")
            mostrar_grafico(estilizar_figura(fig, "Estados por periodo (100 %)", pct_y=True))
    opciones = [d for d in ["Entidad_Pagadora", "Sede", "Periodo", "Servicio", "Regimen"] if d in df.columns]
    if opciones:
        dimension = st.selectbox("Tabla cruzada y mapa de calor: estado por", opciones,
                                 format_func=lambda d: d.replace("_", " "))
        principales = agregar_por(df, dimension).head(25)[dimension].tolist()
        base = df[_serie_grupo(df, dimension).isin(principales)]
        cruce = pd.crosstab(_serie_grupo(base, dimension), _serie_grupo(base, "Estado_Factura"), margins=True,
                            margins_name="Total")
        fig = px.imshow(cruce.drop(index="Total", columns="Total"), aspect="auto", text_auto=True,
                        color_continuous_scale=ESCALA_CONTINUA, labels={"color": "Registros"})
        mostrar_grafico(estilizar_figura(fig, f"Mapa de calor: estado por {dimension.replace('_', ' ').lower()}",
                                         max(380, 26 * len(cruce) + 120)))
        st.dataframe(cruce.reset_index(), hide_index=True)


def mostrar_glosas(ctx):
    """Página 9: frecuencia y severidad de glosas, rankings, distribución y soporte."""
    st.subheader("Análisis de glosas")
    df = ctx["df"]
    if not requiere_columnas(df, ["Glosa_Inicial", "Valor_Neto"], "el análisis de glosas"):
        return
    glosadas = df[df["Tiene_Glosa"]]
    k = calcular_kpis(df)
    n_fact = k["facturas"]
    n_glosa = glosadas["ID_Factura"].nunique() if columna_util(df, "ID_Factura") else len(glosadas)
    frecuencia = dividir_seguro(n_glosa, n_fact, 100)
    c = st.columns(4)
    with c[0]:
        tarjeta_kpi("Valor total glosado", formatear_moneda_corta(k["glosa"]))
    with c[1]:
        tarjeta_kpi("Facturas con glosa", formatear_numero(n_glosa), f"de {formatear_numero(n_fact)} facturas")
    with c[2]:
        tarjeta_kpi("A. Frecuencia de glosa", formatear_porcentaje(frecuencia), "Facturas con glosa / total facturas")
    with c[3]:
        tarjeta_kpi("B. Severidad de glosa", formatear_porcentaje(k["pct_glosa"]), "Valor glosado / valor neto")
    c = st.columns(3)
    with c[0]:
        tarjeta_kpi("Glosa promedio", formatear_moneda(glosadas["Glosa_Inicial"].mean() if len(glosadas) else np.nan),
                    "Entre registros con glosa")
    with c[1]:
        tarjeta_kpi("Glosa mediana", formatear_moneda(glosadas["Glosa_Inicial"].median() if len(glosadas) else np.nan))
    with c[2]:
        tarjeta_kpi("Glosa máxima", formatear_moneda(glosadas["Glosa_Inicial"].max() if len(glosadas) else np.nan))
    if glosadas.empty:
        st.info("No hay registros con glosa en el conjunto filtrado.")
        return

    cierre = construir_cierre_mensual(df, "Mensual")
    if not cierre.empty:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_bar(x=cierre["Periodo"], y=cierre["Valor glosado"], name="Valor glosado", marker_color=COLOR_NARANJA)
        fig.add_scatter(x=cierre["Periodo"], y=cierre["% glosa"], name="Severidad %", mode="lines+markers",
                        line_color=COLOR_ROJO, secondary_y=True)
        if "Registros con glosa" in cierre:
            frecuencia_mes = dividir_seguro(cierre["Registros con glosa"], cierre["Registros"], 100)
            fig.add_scatter(x=cierre["Periodo"], y=frecuencia_mes, name="Frecuencia % (registros)",
                            mode="lines+markers", line={"color": COLOR_AZUL_OSCURO, "dash": "dot"}, secondary_y=True)
        fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f", secondary_y=False)
        fig.update_yaxes(ticksuffix=" %", secondary_y=True)
        mostrar_grafico(estilizar_figura(fig, "Tendencia mensual de glosas"))

    t1, t2, t3, t4 = st.tabs(["Rankings", "Distribución", "Soporte completo vs incompleto", "Glosas más altas"])
    with t1:
        for dimension in [d for d in ["Entidad_Pagadora", "Servicio", "Sede"] if d in df.columns]:
            tabla = agregar_por(df, dimension)
            facturas_glosa = glosadas.groupby(_serie_grupo(glosadas, dimension))["ID_Factura"].nunique() \
                if columna_util(df, "ID_Factura") else glosadas.groupby(_serie_grupo(glosadas, dimension)).size()
            tabla["Facturas con glosa"] = tabla[dimension].map(facturas_glosa).fillna(0)
            base = tabla["Facturas"] if "Facturas" in tabla else tabla["Registros"]
            tabla["Frecuencia %"] = dividir_seguro(tabla["Facturas con glosa"], base, 100).values
            tabla = tabla.sort_values("Valor glosado", ascending=False)
            c1, c2 = st.columns([3, 2])
            with c1:
                datos = tabla.head(12)
                fig = px.bar(datos, x="Valor glosado", y=dimension, orientation="h", color="% glosa",
                             color_continuous_scale=ESCALA_ALERTA)
                fig.update_yaxes(autorange="reversed", title=None)
                fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
                mostrar_grafico(estilizar_figura(fig, f"Glosa por {dimension.replace('_', ' ').lower()} "
                                                      "(color: severidad %)", max(320, 26 * len(datos) + 90)),
                                clave=f"glosa_{dimension}")
            with c2:
                mostrar_tabla(tabla[[dimension, "Valor glosado", "% glosa", "Facturas con glosa", "Frecuencia %"]]
                              .head(15), columnas_num=["Facturas con glosa"])
    with t2:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(glosadas, x="Glosa_Inicial", nbins=40, color_discrete_sequence=[COLOR_NARANJA])
            fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            mostrar_grafico(estilizar_figura(fig, "Histograma del valor glosado", leyenda=False))
        with c2:
            dimension = "Entidad_Pagadora" if "Entidad_Pagadora" in df.columns else None
            if dimension:
                principales = agregar_por(glosadas, dimension).head(10)[dimension].tolist()
                base = glosadas[_serie_grupo(glosadas, dimension).isin(principales)]
                fig = px.box(base.assign(**{dimension: _serie_grupo(base, dimension)}), x=dimension,
                             y="Porcentaje_Glosa", color_discrete_sequence=[COLOR_ROJO])
                fig.update_xaxes(tickangle=-35, title=None)
            else:
                fig = px.box(glosadas, y="Porcentaje_Glosa", color_discrete_sequence=[COLOR_ROJO])
            mostrar_grafico(estilizar_figura(fig, "Boxplot del % de glosa por registro", pct_y=True, leyenda=False))
    with t3:
        if "Soporte_Completo" in df.columns and df["Soporte_Completo"].notna().any():
            tabla = agregar_por(df[df["Soporte_Completo"].notna()], "Soporte_Completo")
            if "Registros con glosa" in tabla:
                tabla["Frecuencia % (registros)"] = dividir_seguro(tabla["Registros con glosa"], tabla["Registros"],
                                                                   100).values
            mostrar_tabla(tabla[[c for c in ["Soporte_Completo", "Registros", "Valor neto", "Valor glosado", "% glosa",
                                             "Registros con glosa", "Frecuencia % (registros)"] if c in tabla]],
                          columnas_num=["Registros", "Registros con glosa"])
            largo = tabla.melt(id_vars="Soporte_Completo",
                               value_vars=[c for c in ["% glosa", "Frecuencia % (registros)"] if c in tabla],
                               var_name="Indicador", value_name="Porcentaje")
            fig = px.bar(largo, x="Soporte_Completo", y="Porcentaje", color="Indicador", barmode="group",
                         color_discrete_sequence=[COLOR_ROJO, COLOR_AZUL_OSCURO])
            mostrar_grafico(estilizar_figura(fig, "Glosa según soporte completo", pct_y=True))
            st.caption("Una asociación estadística entre soporte y glosa no demuestra causalidad.")
        else:
            st.info("Requiere la columna Soporte_Completo.")
    with t4:
        columnas = [c for c in ["ID_Factura", "Numero_Radicado", "Periodo", "Entidad_Pagadora", "Servicio", "Sede",
                                "Estado_Factura", "Valor_Neto", "Glosa_Inicial", "Porcentaje_Glosa",
                                "Soporte_Completo"] if c in glosadas.columns]
        mostrar_tabla(glosadas.nlargest(50, "Glosa_Inicial")[columnas], altura=460)


def mostrar_cartera(ctx):
    """Página 10: recaudo, cartera, rankings y antigüedad aproximada."""
    st.subheader("Recaudo y cartera")
    df = ctx["df"]
    if not requiere_columnas(df, ["Valor_Neto", "Valor_Pagado", "Saldo_Cartera"], "el análisis de recaudo y cartera"):
        return
    k = calcular_kpis(df)
    con_saldo = df[df["Saldo_Cartera"] > 0]
    pagadas_total = df[(df["Saldo_Cartera"] <= 0) & (df["Valor_Pagado"] > 0)]
    parciales = df[(df["Saldo_Cartera"] > 0) & (df["Valor_Pagado"] > 0)]
    contar = lambda d: d["ID_Factura"].nunique() if columna_util(d, "ID_Factura") else len(d)  # noqa: E731
    c = st.columns(4)
    with c[0]:
        tarjeta_kpi("Valor neto", formatear_moneda_corta(k["neto"]))
    with c[1]:
        tarjeta_kpi("Valor pagado", formatear_moneda_corta(k["pagado"]), f"Recaudo: {formatear_porcentaje(k['pct_recaudo'])}")
    with c[2]:
        tarjeta_kpi("Saldo de cartera", formatear_moneda_corta(k["saldo"]), f"Cartera: {formatear_porcentaje(k['pct_cartera'])}")
    with c[3]:
        tarjeta_kpi("Facturas con saldo", formatear_numero(contar(con_saldo)))
    c = st.columns(4)
    with c[0]:
        tarjeta_kpi("Totalmente pagadas", formatear_numero(contar(pagadas_total)), "Saldo ≤ 0 y pago > 0")
    with c[1]:
        tarjeta_kpi("Parcialmente pagadas", formatear_numero(contar(parciales)), "Saldo > 0 y pago > 0")
    with c[2]:
        tarjeta_kpi("Saldo promedio", formatear_moneda(con_saldo["Saldo_Cartera"].mean() if len(con_saldo) else np.nan),
                    "Entre registros con saldo")
    with c[3]:
        tarjeta_kpi("Saldo mediano", formatear_moneda(con_saldo["Saldo_Cartera"].median() if len(con_saldo) else np.nan))

    cierre = construir_cierre_mensual(df, "Mensual")
    t1, t2, t3, t4 = st.tabs(["Evolución", "Cartera por dimensión", "Antigüedad aproximada", "Facturas con mayor saldo"])
    with t1:
        if not cierre.empty:
            c1, c2 = st.columns(2)
            with c1:
                largo = cierre.melt(id_vars="Periodo", value_vars=["Valor neto", "Valor pagado", "Saldo de cartera"],
                                    var_name="Concepto", value_name="Valor")
                fig = px.line(largo, x="Periodo", y="Valor", color="Concepto", markers=True,
                              color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_VERDE_AZULADO, COLOR_NARANJA])
                mostrar_grafico(estilizar_figura(fig, "Valor neto, pago y saldo por periodo", moneda_y=True))
            with c2:
                largo = cierre.melt(id_vars="Periodo", value_vars=["% recaudo", "% cartera"], var_name="Indicador",
                                    value_name="Porcentaje")
                fig = px.line(largo, x="Periodo", y="Porcentaje", color="Indicador", markers=True,
                              color_discrete_sequence=[COLOR_VERDE_AZULADO, COLOR_NARANJA])
                mostrar_grafico(estilizar_figura(fig, "Evolución de % recaudo y % cartera", pct_y=True))
            st.caption("Los periodos recientes suelen mostrar menor recaudo porque las facturas aún están dentro de "
                       "sus plazos de pago; compare con prudencia.")
    with t2:
        opciones = [d for d in ["Entidad_Pagadora", "Servicio", "Sede", "Periodo", "Regimen"] if d in df.columns]
        dimension = st.selectbox("Cartera por", opciones, format_func=lambda d: d.replace("_", " "))
        tabla = agregar_por(df, dimension).sort_values("Saldo de cartera", ascending=False)
        datos = tabla.head(15)
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(datos, x="Saldo de cartera", y=dimension, orientation="h", color="% cartera",
                         color_continuous_scale=ESCALA_ALERTA)
            fig.update_yaxes(autorange="reversed", title=None)
            fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            mostrar_grafico(estilizar_figura(fig, "Ranking de cartera (color: % cartera)", max(340, 26 * len(datos) + 90)))
        with c2:
            fig = px.scatter(tabla[tabla[dimension] != SIN_INFO], x="Valor neto", y="Valor pagado",
                             hover_name=dimension, size=tabla.loc[tabla[dimension] != SIN_INFO, "Saldo de cartera"]
                             .clip(lower=0).fillna(0) + 1, color_discrete_sequence=[COLOR_AZUL_MEDIO], size_max=40)
            maximo = float(tabla["Valor neto"].max()) if len(tabla) else 0
            fig.add_scatter(x=[0, maximo], y=[0, maximo], mode="lines", name="Recaudo 100 %",
                            line={"dash": "dot", "color": "#8A96A3"})
            fig.update_xaxes(tickprefix=_simbolo(), tickformat=",.0f")
            fig.update_yaxes(tickprefix=_simbolo(), tickformat=",.0f")
            mostrar_grafico(estilizar_figura(fig, "Relación facturación - recaudo (tamaño: saldo)"))
        mostrar_tabla(tabla[[c for c in [dimension, "Facturas", "Valor neto", "Valor pagado", "Saldo de cartera",
                                         "% recaudo", "% cartera"] if c in tabla]], columnas_num=["Facturas"])
        if dimension == "Entidad_Pagadora" and len(tabla) >= 3:
            mediana_neto = tabla["Valor neto"].median()
            alerta = tabla[(tabla["Valor neto"] >= mediana_neto) & (tabla["% recaudo"] < k["pct_recaudo"])]
            st.markdown("**Entidades con alto volumen y bajo recaudo**")
            st.caption(f"Criterio analítico: facturación neta ≥ mediana de entidades ({formatear_moneda(mediana_neto)}) "
                       f"y % de recaudo inferior al global ({formatear_porcentaje(k['pct_recaudo'])}).")
            if alerta.empty:
                st.success("Ninguna entidad cumple simultáneamente ambos criterios.")
            else:
                mostrar_tabla(alerta[[c for c in ["Entidad_Pagadora", "Valor neto", "Valor pagado", "% recaudo",
                                                  "Saldo de cartera"] if c in alerta]])
    with t3:
        if not columna_util(df, "Fecha_Radicacion"):
            st.info("La antigüedad requiere Fecha_Radicacion.")
        else:
            maximo = df["Fecha_Radicacion"].max().date()
            corte = st.date_input("Fecha de corte", value=maximo, format="DD/MM/YYYY",
                                  help="Por defecto, la fecha de radicación más reciente del conjunto filtrado.")
            limites = tuple(sorted({int(st.session_state[f"ant_{i}"]) for i in range(1, 5)}))
            resumen, base = calcular_antiguedad(df, corte, limites)
            if resumen.empty:
                st.info("No hay registros con saldo y fecha de radicación anteriores a la fecha de corte.")
            else:
                c1, c2 = st.columns([3, 2])
                with c1:
                    fig = px.bar(resumen, x="Rango de antigüedad", y="Saldo de cartera", text=resumen["Participación %"]
                                 .map(lambda v: formatear_porcentaje(v)), color_discrete_sequence=[COLOR_NARANJA])
                    mostrar_grafico(estilizar_figura(fig, "Saldo por rango de antigüedad", moneda_y=True, leyenda=False))
                with c2:
                    mostrar_tabla(resumen, columnas_num=["Registros"])
                if "Entidad_Pagadora" in base.columns:
                    cruce = base.pivot_table(index=_serie_grupo(base, "Entidad_Pagadora"),
                                             columns="Rango de antigüedad", values="Saldo_Cartera", aggfunc="sum",
                                             fill_value=0, observed=False)
                    cruce = cruce.loc[cruce.sum(axis=1).sort_values(ascending=False).index].head(20)
                    fig = px.imshow(cruce, aspect="auto", color_continuous_scale=ESCALA_ALERTA,
                                    labels={"color": "Saldo", "x": "Antigüedad", "y": "Entidad"})
                    mostrar_grafico(estilizar_figura(fig, "Saldo por entidad y antigüedad (top 20)",
                                                     max(360, 26 * len(cruce) + 120)))
            st.caption("Antigüedad aproximada = fecha de corte − fecha de radicación. Los rangos son analíticos y "
                       "configurables en la barra lateral; no constituyen una política contable o contractual.")
    with t4:
        columnas = [c for c in ["ID_Factura", "Numero_Radicado", "Periodo", "Entidad_Pagadora", "Servicio",
                                "Estado_Factura", "Fecha_Radicacion", "Valor_Neto", "Valor_Pagado", "Glosa_Inicial",
                                "Saldo_Cartera", "Porcentaje_Cartera"] if c in df.columns]
        mostrar_tabla(con_saldo.nlargest(50, "Saldo_Cartera")[columnas], altura=460)


def mostrar_tiempos(ctx):
    """Página 11: tiempos entre emisión y radicación."""
    st.subheader("Tiempos de radicación")
    df = ctx["df"]
    if not columna_util(df, "Dias_Emision_Radicacion"):
        st.warning("No se encontró Dias_Emision_Radicacion ni las fechas de emisión y radicación para calcularlo.")
        return
    origen = df["_Origen_Dias"].iloc[0] if "_Origen_Dias" in df.columns and len(df) else ""
    st.caption(f"Fuente de los días: {origen}.")
    dias = df["Dias_Emision_Radicacion"]
    e = calcular_estadisticos(dias)
    if e is None:
        st.info("La variable seleccionada no contiene suficientes datos numéricos.")
        return
    metricas = [("Promedio", e["Media"]), ("Mediana", e["Mediana"]), ("Desv. estándar", e["Desviación estándar"]),
                ("Mínimo", e["Mínimo"]), ("Máximo", e["Máximo"]), ("Percentil 75", e["Cuartil 3"]),
                ("Percentil 90", e["Percentil 90"]), ("Percentil 95", e["Percentil 95"])]
    columnas = st.columns(8)
    for contenedor, (nombre, valor) in zip(columnas, metricas):
        contenedor.metric(nombre, f"{formatear_numero(valor, 1)} d")
    negativos = int((dias < 0).sum())
    if negativos:
        st.warning(f"{negativos} registros tienen días negativos (radicación anterior a la emisión). Revíselos en "
                   "Calidad de datos.")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.histogram(df, x="Dias_Emision_Radicacion", nbins=40, color_discrete_sequence=[COLOR_AZUL_MEDIO])
        fig.add_vline(x=e["Mediana"], line_dash="dash", line_color=COLOR_NARANJA, annotation_text="Mediana")
        mostrar_grafico(estilizar_figura(fig, "Histograma de días emisión - radicación", leyenda=False))
    with c2:
        if "Rango_Tiempo_Radicacion" in df.columns:
            rangos = df["Rango_Tiempo_Radicacion"].value_counts(sort=False).reset_index()
            rangos.columns = ["Rango", "Registros"]
            rangos["%"] = dividir_seguro(rangos["Registros"], pd.Series([rangos["Registros"].sum()] * len(rangos)),
                                         100).values
            fig = px.bar(rangos, x="Rango", y="Registros", text=rangos["%"].map(formatear_porcentaje),
                         color_discrete_sequence=[COLOR_VERDE_AZULADO])
            mostrar_grafico(estilizar_figura(fig, "Registros por rango de días (configurable)", leyenda=False))
    st.caption("Los rangos de días son parámetros analíticos configurables en la barra lateral, no estándares "
               "normativos.")
    opciones = [d for d in ["Entidad_Pagadora", "Sede", "Canal_Radicacion", "Servicio", "Estado_Factura",
                            "Soporte_Completo", "Periodo"] if d in df.columns]
    if opciones:
        dimension = st.selectbox("Analizar por", opciones, format_func=lambda d: d.replace("_", " "))
        agrupado = df.groupby(_serie_grupo(df, dimension))["Dias_Emision_Radicacion"]
        tabla = agrupado.agg(["count", "mean", "median", "std", "min", "max"]).reset_index()
        tabla.columns = [dimension, "Registros", "Promedio días", "Mediana días", "Desv. estándar", "Mínimo", "Máximo"]
        tabla["Percentil 90"] = agrupado.quantile(0.9).values
        tabla = tabla.sort_values(dimension if dimension == "Periodo" else "Promedio días",
                                  ascending=dimension == "Periodo")
        c1, c2 = st.columns(2)
        with c1:
            if dimension == "Periodo":
                datos = tabla[tabla["Periodo"] != SIN_INFO]
                fig = px.line(datos, x="Periodo", y=["Promedio días", "Mediana días"], markers=True,
                              color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_NARANJA])
                mostrar_grafico(estilizar_figura(fig, "Tendencia mensual de días de radicación"))
            else:
                datos = tabla.head(20)
                fig = px.bar(datos, x="Promedio días", y=dimension, orientation="h",
                             color_discrete_sequence=[COLOR_AZUL_MEDIO], hover_data={"Mediana días": ":.1f"})
                fig.update_yaxes(autorange="reversed", title=None)
                mostrar_grafico(estilizar_figura(fig, "Promedio de días por grupo", max(320, 26 * len(datos) + 90),
                                                 leyenda=False))
        with c2:
            principales = tabla.head(15)[dimension].tolist()
            base = df[_serie_grupo(df, dimension).isin(principales)]
            fig = px.box(base.assign(**{dimension: _serie_grupo(base, dimension)}), x=dimension,
                         y="Dias_Emision_Radicacion", color_discrete_sequence=[COLOR_VERDE_AZULADO])
            fig.update_xaxes(tickangle=-35, title=None)
            mostrar_grafico(estilizar_figura(fig, "Boxplot de días por grupo", leyenda=False))
        mostrar_tabla(tabla, columnas_num=["Registros", "Promedio días", "Mediana días", "Desv. estándar", "Mínimo",
                                           "Máximo", "Percentil 90"])
        if "Periodo" in df.columns and dimension != "Periodo":
            mensual = df[df["Periodo"].notna()].groupby("Periodo")["Dias_Emision_Radicacion"].agg(["mean", "median"])
            mensual = mensual.reset_index().rename(columns={"mean": "Promedio", "median": "Mediana"})
            fig = px.line(mensual, x="Periodo", y=["Promedio", "Mediana"], markers=True,
                          color_discrete_sequence=[COLOR_AZUL_OSCURO, COLOR_NARANJA])
            mostrar_grafico(estilizar_figura(fig, "Tendencia mensual de días de radicación"))
    st.markdown("#### Casos extremos (mayor número de días)")
    columnas = [c for c in ["ID_Factura", "Entidad_Pagadora", "Sede", "Canal_Radicacion", "Servicio", "Fecha_Emision",
                            "Fecha_Radicacion", "Dias_Emision_Radicacion", "Valor_Neto", "Estado_Factura"]
                if c in df.columns]
    mostrar_tabla(df.nlargest(30, "Dias_Emision_Radicacion")[columnas], columnas_num=["Dias_Emision_Radicacion"])


def mostrar_estadistica(ctx):
    """Página 12: estadística descriptiva, correlaciones y pruebas estadísticas."""
    st.subheader("Análisis estadístico")
    df = ctx["df"]
    variables = disponibles(df, VARIABLES_NUMERICAS_ANALISIS)
    if not variables:
        st.warning("No hay variables numéricas disponibles.")
        return
    t1, t2, t3 = st.tabs(["Estadística descriptiva", "Correlaciones", "Pruebas estadísticas"])
    with t1:
        variable = st.selectbox("Variable numérica", variables, key="est_var")
        e = calcular_estadisticos(df[variable])
        if e is None:
            st.info("La variable seleccionada no contiene suficientes datos numéricos.")
        else:
            c1, c2 = st.columns([1, 2])
            with c1:
                tabla = pd.DataFrame({"Estadístico": list(e.keys()), "Valor": list(e.values())})
                tabla["Valor"] = [formatear_numero(v, 0 if k == "Conteo" else 2) if es_valor_valido(v)
                                  else ("No aplica" if k == "Moda" else "No calculable")
                                  for k, v in zip(tabla["Estadístico"], tabla["Valor"])]
                st.dataframe(tabla, hide_index=True, height=640)
            with c2:
                datos = pd.to_numeric(df[variable], errors="coerce").dropna()
                fig = px.histogram(datos, x=variable, nbins=50, histnorm="probability density", marginal="box",
                                   color_discrete_sequence=[COLOR_AZUL_MEDIO])
                if len(datos) > 5 and datos.std() > 0:
                    try:
                        muestra = datos.sample(min(len(datos), 5000), random_state=42)
                        kde = stats.gaussian_kde(muestra)
                        x = np.linspace(datos.min(), datos.max(), 200)
                        fig.add_scatter(x=x, y=kde(x), mode="lines", name="Distribución observada (KDE)",
                                        line_color=COLOR_AZUL_OSCURO)
                    except (np.linalg.LinAlgError, ValueError):
                        st.caption("No fue posible estimar la curva de densidad.")
                mostrar_grafico(estilizar_figura(fig, f"Distribución de {variable}", 520))
            st.markdown("**Interpretación automática**")
            for texto in interpretar_estadisticos(e, variable):
                st.markdown(f"- {texto}")
    with t2:
        c1, c2, c3 = st.columns(3)
        metodo = c1.radio("Método", ["Pearson", "Spearman"], horizontal=True)
        fuerte = c2.slider("Umbral de correlación fuerte", 0.5, 0.95, 0.7, 0.05)
        debil = c3.slider("Umbral de correlación débil", 0.05, 0.5, 0.3, 0.05)
        base_vars = [v for v in ["Cantidad", "Valor_Bruto", "Descuento", "Copago_Cuota", "Valor_Neto",
                                 "Dias_Emision_Radicacion", "Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera"]
                     if v in variables]
        matriz = calcular_correlaciones(df, base_vars, metodo.lower())
        if matriz.empty:
            st.info("Se requieren al menos dos variables numéricas con variabilidad.")
        else:
            fig = px.imshow(matriz, text_auto=".2f", color_continuous_scale="RdBu_r", zmin=-1, zmax=1, aspect="auto")
            mostrar_grafico(estilizar_figura(fig, f"Matriz de correlación ({metodo})", 520))
            pares = clasificar_correlaciones(matriz, fuerte, debil)
            if not pares.empty:
                c1, c2, c3 = st.columns(3)
                for contenedor, clase in zip([c1, c2, c3], ["Positiva fuerte", "Negativa fuerte", "Débil"]):
                    contenedor.metric(f"Pares: {clase.lower()}", int((pares["Clasificación"] == clase).sum()))
                st.dataframe(pares, hide_index=True)
            st.warning(AVISO_CORRELACION)
            c1, c2 = st.columns(2)
            var_x = c1.selectbox("Variable X", list(matriz.columns), index=0)
            var_y = c2.selectbox("Variable Y", list(matriz.columns), index=min(1, len(matriz.columns) - 1))
            pares_xy = df[[var_x, var_y]].apply(pd.to_numeric, errors="coerce").dropna() if var_x != var_y else None
            if pares_xy is None or len(pares_xy) < 3:
                st.info("Seleccione dos variables distintas con al menos 3 pares válidos.")
            else:
                muestra = pares_xy.sample(min(len(pares_xy), 5000), random_state=42)
                fig = px.scatter(muestra, x=var_x, y=var_y, opacity=0.55, color_discrete_sequence=[COLOR_AZUL_MEDIO])
                if muestra[var_x].std() > 0 and muestra[var_x].nunique() > 2:
                    try:
                        pendiente, intercepto = np.polyfit(muestra[var_x], muestra[var_y], 1)
                        x = np.linspace(muestra[var_x].min(), muestra[var_x].max(), 50)
                        fig.add_scatter(x=x, y=pendiente * x + intercepto, mode="lines", name="Tendencia lineal",
                                        line_color=COLOR_NARANJA)
                    except (np.linalg.LinAlgError, ValueError):
                        pass
                r = matriz.loc[var_x, var_y]
                mostrar_grafico(estilizar_figura(fig, f"{var_x} vs {var_y} (r = {formatear_numero(r, 3)})", 460))
                if len(pares_xy) > 5000:
                    st.caption("Se grafica una muestra reproducible de 5.000 puntos.")
    with t3:
        mostrar_pruebas(df, variables)


def _mostrar_resultado_prueba(resultado):
    """Presenta hipótesis, estadístico, p-valor, decisión e interpretación de una prueba."""
    if "error" in resultado:
        st.warning(resultado["error"])
        return
    if resultado.get("seleccion"):
        st.info(resultado["seleccion"])
    c1, c2 = st.columns(2)
    c1.markdown(f"**Prueba:** {resultado['prueba']}")
    c1.markdown(f"**Hipótesis nula (H0):** {resultado['h0']}")
    c1.markdown(f"**Hipótesis alternativa (H1):** {resultado['h1']}")
    c2.markdown(f"**Estadístico:** {formatear_numero(resultado['estadistico'], 4)}")
    c2.markdown(f"**p-valor:** {formatear_numero(resultado['p'], 4) if resultado['p'] >= 0.0001 else '< 0,0001'}")
    c2.markdown(f"**Nivel de significancia:** {formatear_numero(resultado['alfa'], 2)}")
    c2.markdown(f"**Decisión estadística:** {resultado['decision']}")
    st.markdown("**Tamaño de las muestras:** " + " · ".join(f"{k}: {formatear_numero(v)}"
                                                            for k, v in resultado.get("muestras", {}).items()))
    if resultado.get("extra"):
        st.markdown(resultado["extra"])
    if resultado.get("advertencia"):
        st.warning(resultado["advertencia"])
    st.success(f"Interpretación: {resultado['interpretacion']}")


def mostrar_pruebas(df, variables):
    """Sección opcional de pruebas de hipótesis."""
    st.caption("Sección opcional. Los resultados describen diferencias o asociaciones estadísticas; no establecen "
               "causalidad.")
    categoricas = [c for c in ["Entidad_Pagadora", "Regimen", "Sede", "Servicio", "Linea_Servicio", "Estado_Factura",
                               "Canal_Radicacion", "Soporte_Completo", "Rango_Tiempo_Radicacion"]
                   if c in df.columns and df[c].nunique() >= 2]
    tipo = st.radio("Tipo de prueba", ["A. Variable numérica vs dos grupos", "B. Variable numérica vs más de dos grupos",
                                       "C. Dos variables categóricas (chi-cuadrado)", "D. Normalidad (Shapiro-Wilk)"])
    alfa = st.select_slider("Nivel de significancia (α)", options=[0.01, 0.05, 0.10], value=0.05)
    try:
        if tipo.startswith("A") or tipo.startswith("B"):
            if not categoricas:
                st.info("No existen variables de agrupación con al menos dos categorías.")
                return
            c1, c2, c3 = st.columns(3)
            variable = c1.selectbox("Variable numérica", variables, key="prueba_var")
            agrupacion = c2.selectbox("Agrupación", categoricas, key="prueba_grupo")
            grupos_serie = _serie_grupo(df, agrupacion)
            conteo = grupos_serie.value_counts()
            if tipo.startswith("A"):
                prueba = c3.selectbox("Prueba", ["Automática", "t de Student (Welch)", "Mann-Whitney U"])
                opciones = conteo.index.tolist()
                g1 = st.selectbox("Grupo 1", opciones, index=0)
                g2 = st.selectbox("Grupo 2", opciones, index=1 if len(opciones) > 1 else 0)
                if g1 == g2:
                    st.info("Seleccione dos grupos diferentes.")
                    return
                resultado = ejecutar_prueba_estadistica(
                    "dos_grupos", [df.loc[grupos_serie == g1, variable], df.loc[grupos_serie == g2, variable]],
                    alfa, prueba=prueba)
            else:
                prueba = c3.selectbox("Prueba", ["Automática", "ANOVA de una vía", "Kruskal-Wallis"])
                principales = conteo[conteo >= 3].head(20).index.tolist()
                if len(conteo) > 20:
                    st.caption("Se usan los 20 grupos con más registros (mínimo 3 por grupo).")
                resultado = ejecutar_prueba_estadistica(
                    "varios_grupos", [df.loc[grupos_serie == g, variable] for g in principales], alfa, prueba=prueba)
            _mostrar_resultado_prueba(resultado)
        elif tipo.startswith("C"):
            if len(categoricas) < 2:
                st.info("Se requieren dos variables categóricas con al menos dos categorías.")
                return
            c1, c2 = st.columns(2)
            var1 = c1.selectbox("Variable 1", categoricas, index=0)
            var2 = c2.selectbox("Variable 2", categoricas, index=1)
            if var1 == var2:
                st.info("Seleccione dos variables diferentes.")
                return
            tabla = pd.crosstab(_serie_grupo(df, var1), _serie_grupo(df, var2))
            if tabla.shape[0] > 30 or tabla.shape[1] > 30:
                st.caption("Tabla con muchas categorías; se usan las 30 más frecuentes de cada variable.")
                tabla = tabla.loc[tabla.sum(axis=1).nlargest(30).index, tabla.sum(axis=0).nlargest(30).index]
            st.dataframe(tabla, height=min(38 + 35 * len(tabla), 380))
            _mostrar_resultado_prueba(ejecutar_prueba_estadistica("chi_cuadrado", tabla, alfa))
        else:
            variable = st.selectbox("Variable numérica", variables, key="prueba_normal")
            _mostrar_resultado_prueba(ejecutar_prueba_estadistica("normalidad", df[variable], alfa))
    except Exception as error:  # noqa: BLE001
        st.warning(f"No fue posible ejecutar la prueba ({type(error).__name__}). Verifique que existan suficientes "
                   "datos y grupos.")


def mostrar_atipicos(ctx):
    """Página 13: detección de valores atípicos por IQR, puntaje Z o percentiles."""
    st.subheader("Valores atípicos")
    df = ctx["df"]
    variables = disponibles(df, VARIABLES_NUMERICAS_ANALISIS)
    if not variables:
        st.warning("No hay variables numéricas disponibles.")
        return
    c1, c2, c3 = st.columns(3)
    variable = c1.selectbox("Variable", variables, index=variables.index("Valor_Neto") if "Valor_Neto" in variables else 0)
    metodo = c2.selectbox("Método", ["IQR", "Puntaje Z", "Percentiles"])
    alcances = {"Global": None}
    for col, nombre in [("Servicio", "Por servicio"), ("Linea_Servicio", "Por línea de servicio"),
                        ("Entidad_Pagadora", "Por entidad")]:
        if col in df.columns:
            alcances[nombre] = col
    alcance = c3.selectbox("Alcance", list(alcances))
    multiplicador, umbral_z, percentiles = 1.5, 3.0, (1.0, 99.0)
    if metodo == "IQR":
        multiplicador = st.slider("Multiplicador del IQR", 0.5, 5.0, 1.5, 0.1)
        st.caption("Límite inferior = Q1 − k × IQR; límite superior = Q3 + k × IQR.")
    elif metodo == "Puntaje Z":
        umbral_z = st.slider("Umbral |Z|", 1.5, 5.0, 3.0, 0.1)
        st.caption("El puntaje Z es más apropiado para distribuciones aproximadamente simétricas; en variables "
                   "monetarias con sesgo fuerte el IQR suele ser más robusto.")
    else:
        percentiles = st.slider("Percentiles inferior y superior", 0.0, 100.0, (1.0, 99.0), 0.5)
    mascara, limites = detectar_atipicos(df, variable, metodo, multiplicador, umbral_z, percentiles,
                                         alcances[alcance])
    n_atip = int(mascara.sum())
    validos = int(pd.to_numeric(df[variable], errors="coerce").notna().sum())
    c = st.columns(4)
    c[0].metric("Valores atípicos", formatear_numero(n_atip))
    c[1].metric("Porcentaje", formatear_porcentaje(dividir_seguro(n_atip, validos, 100)))
    es_moneda = variable in CAMPOS_MONEDA
    formato = formatear_moneda if es_moneda else (lambda v: formatear_numero(v, 2))
    if alcance == "Global" and not limites.empty:
        c[2].metric("Límite inferior", formato(limites["Límite inferior"].iloc[0]))
        c[3].metric("Límite superior", formato(limites["Límite superior"].iloc[0]))
    else:
        c[2].metric("Grupos evaluados", formatear_numero(len(limites)))
        c[3].metric("Grupos con atípicos", formatear_numero(int((limites["Atípicos"] > 0).sum())))
    st.info("Un valor atípico no es necesariamente un error. Puede representar una factura legítima de mayor "
            "complejidad o valor.")
    marcado = df.assign(Clasificación=np.where(mascara, "Atípico", "Normal"))
    c1, c2 = st.columns(2)
    with c1:
        fig = px.box(marcado, y=variable, points="outliers", color_discrete_sequence=[COLOR_AZUL_MEDIO])
        mostrar_grafico(estilizar_figura(fig, "Boxplot", moneda_y=es_moneda, leyenda=False))
    with c2:
        fig = px.histogram(marcado, x=variable, color="Clasificación", nbins=50,
                           color_discrete_map={"Normal": COLOR_AZUL_MEDIO, "Atípico": COLOR_ROJO})
        if alcance == "Global" and not limites.empty:
            for nombre, valor in [("Inferior", limites["Límite inferior"].iloc[0]),
                                  ("Superior", limites["Límite superior"].iloc[0])]:
                if es_valor_valido(valor):
                    fig.add_vline(x=valor, line_dash="dash", line_color=COLOR_NARANJA, annotation_text=nombre)
        mostrar_grafico(estilizar_figura(fig, "Histograma con límites"))
    if "_Fecha_Referencia" in df.columns:
        muestra = marcado[marcado["_Fecha_Referencia"].notna()]
        if len(muestra) > 8000:
            muestra = pd.concat([muestra[muestra["Clasificación"] == "Atípico"],
                                 muestra[muestra["Clasificación"] == "Normal"].sample(6000, random_state=42)])
            st.caption("La dispersión muestra todos los atípicos y una muestra reproducible de registros normales.")
        fig = px.scatter(muestra, x="_Fecha_Referencia", y=variable, color="Clasificación", opacity=0.6,
                         color_discrete_map={"Normal": COLOR_AZUL_MEDIO, "Atípico": COLOR_ROJO},
                         labels={"_Fecha_Referencia": "Fecha"})
        mostrar_grafico(estilizar_figura(fig, "Dispersión por periodo", moneda_y=es_moneda))
    if alcance != "Global":
        with st.expander("Límites por grupo"):
            mostrar_tabla(limites, columnas_num=["Registros válidos", "Atípicos"],
                          columnas_moneda=["Límite inferior", "Límite superior"] if es_moneda else None)
    st.markdown("#### Registros atípicos")
    columnas = [c for c in ["ID_Factura", "Periodo", "Entidad_Pagadora", "Servicio", "Linea_Servicio", "Sede",
                            "Estado_Factura", variable, "Valor_Neto"] if c in df.columns]
    columnas = list(dict.fromkeys(columnas))
    atipicos = df.loc[mascara, columnas].sort_values(variable, ascending=False)
    mostrar_tabla(atipicos.head(500), altura=420)
    if len(atipicos) > 500:
        st.caption(f"Se muestran 500 de {formatear_numero(len(atipicos))} registros. Descárguelos completos en "
                   "la sección Descargas.")
    st.session_state["_atipicos_config"] = {"variable": variable, "metodo": metodo, "multiplicador": multiplicador,
                                            "umbral_z": umbral_z, "percentiles": tuple(percentiles),
                                            "alcance": alcances[alcance]}


def mostrar_insights(ctx):
    """Página 14: insights calculados por reglas y semáforos configurables."""
    st.subheader("Insights y alertas")
    df = ctx["df"]
    st.markdown("#### Semáforos")
    st.caption("Umbrales demostrativos y editables en la barra lateral. No representan normas, contratos ni "
               "políticas institucionales salvo que el usuario los haya definido así.")
    alertas = generar_alertas(df, obtener_umbrales())
    columnas = st.columns(3)
    for i, alerta in enumerate(alertas):
        with columnas[i % 3]:
            tarjeta_kpi(alerta["Indicador"], alerta["Valor"], f"{alerta['Estado']} · {alerta['Lectura']}<br>"
                        f"<span style='font-size:0.72rem'>{alerta['Umbrales configurados']}</span>", alerta["_color"])
    if "Entidad_Pagadora" in df.columns:
        with st.expander("Semáforos por entidad"):
            u = obtener_umbrales()
            tabla = agregar_por(df, "Entidad_Pagadora")
            filas = []
            for _, fila in tabla.iterrows():  # tabla agregada (pocas filas), no el DataFrame original
                registro = {"Entidad": fila["Entidad_Pagadora"], "Valor neto": fila["Valor neto"]}
                for col, clave, mayor in [("% glosa", "glosa", False), ("% cartera", "cartera", False),
                                          ("% recaudo", "recaudo", True)]:
                    if col in tabla.columns:
                        color, etiqueta = evaluar_semaforo(fila[col], u[f"{clave}_verde"], u[f"{clave}_amarillo"], mayor)
                        registro[col] = f"{ICONO_SEMAFORO[color]} {formatear_porcentaje(fila[col])} ({etiqueta})"
                filas.append(registro)
            mostrar_tabla(pd.DataFrame(filas))
    st.markdown("#### Insights calculados")
    st.caption("Generados mediante reglas programadas sobre los datos filtrados. No se utilizan modelos de "
               "inteligencia artificial ni servicios externos.")
    insights = generar_insights(df, ctx.get("calidad"))
    if not insights:
        st.info("Los datos no son suficientes para generar insights.")
        return
    categorias = sorted({i["Categoría"] for i in insights})
    seleccion = st.multiselect("Categorías", categorias, default=categorias)
    for ins in [i for i in insights if i["Categoría"] in seleccion]:
        st.markdown(
            f'<div class="insight-card"><b>[{ins["Categoría"]}] {ins["Hallazgo calculado"]}</b><br>'
            f'<b>Hallazgo calculado:</b> {ins["Valor observado"]} — {ins["Contexto"]}<br>'
            f'<b>Interpretación:</b> {ins["Interpretación"]}<br>'
            f'<b>Recomendación para investigar:</b> {ins["Recomendación para investigar"]}</div>',
            unsafe_allow_html=True)


def mostrar_explorador(ctx):
    """Página 15: exploración controlada de registros."""
    st.subheader("Explorador de datos")
    df = ctx["df"]
    if df.empty:
        st.warning("Los filtros actuales no tienen registros.")
        return
    columnas_todas = columnas_visibles(df)
    c1, c2 = st.columns([2, 1])
    texto = c1.text_input("Buscar texto (en columnas de texto)", placeholder="Ej.: número de factura, entidad...")
    completos = c2.checkbox("Mostrar registros completos (todas las columnas)")
    defecto = [c for c in ["ID_Factura", "Periodo", "Entidad_Pagadora", "Servicio", "Sede", "Estado_Factura",
                           "Valor_Neto", "Glosa_Inicial", "Valor_Pagado", "Saldo_Cartera"] if c in columnas_todas]
    seleccion = columnas_todas if completos else st.multiselect("Columnas", columnas_todas,
                                                                default=defecto or columnas_todas[:10])
    st.markdown("**Filtros rápidos**")
    q = st.columns(5)
    solo_saldo = q[0].checkbox("Con saldo", disabled="Tiene_Saldo" not in df.columns)
    solo_glosa = q[1].checkbox("Con glosa", disabled="Tiene_Glosa" not in df.columns)
    solo_soporte = q[2].checkbox("Soporte incompleto", disabled="Soporte_Completo" not in df.columns)
    solo_incons = q[3].checkbox("Posibles inconsistencias", disabled="_Posible_Inconsistencia" not in df.columns)
    solo_atip = q[4].checkbox("Valores atípicos", help="Usa la configuración de la sección Valores atípicos o, por "
                              "defecto, Valor_Neto con IQR 1,5.")
    mascara = pd.Series(True, index=df.index)
    if solo_saldo and "Tiene_Saldo" in df.columns:
        mascara &= df["Tiene_Saldo"]
    if solo_glosa and "Tiene_Glosa" in df.columns:
        mascara &= df["Tiene_Glosa"]
    if solo_soporte and "Soporte_Completo" in df.columns:
        mascara &= (df["Soporte_Completo"] == "No").fillna(False)
    if solo_incons and "_Posible_Inconsistencia" in df.columns:
        mascara &= df["_Posible_Inconsistencia"]
    if solo_atip:
        mascara &= mascara_atipicos_sesion(df)
    if texto:
        columnas_texto = [c for c in columnas_todas if pd.api.types.is_string_dtype(df[c])
                          or pd.api.types.is_object_dtype(df[c])]
        coincide = pd.Series(False, index=df.index)
        for col in columnas_texto:
            coincide |= df[col].astype("string").str.contains(texto, case=False, regex=False, na=False)
        mascara &= coincide
    resultado = df[mascara]
    c1, c2, c3 = st.columns(3)
    orden = c1.selectbox("Ordenar por", ["(sin orden)"] + list(seleccion))
    ascendente = c2.radio("Sentido", ["Descendente", "Ascendente"], horizontal=True) == "Ascendente"
    filas = c3.select_slider("Filas a mostrar", options=[25, 50, 100, 250, 500, 1000, 2000], value=100)
    if orden != "(sin orden)" and orden in resultado.columns:
        resultado = resultado.sort_values(orden, ascending=ascendente, na_position="last")
    st.caption(f"{formatear_numero(len(resultado))} registros coinciden; se muestran hasta {formatear_numero(filas)}.")
    if not seleccion:
        st.info("Seleccione al menos una columna.")
        return
    st.dataframe(formatear_tabla(resultado[seleccion].head(filas), columnas_moneda=CAMPOS_MONEDA,
                                 columnas_num=["Cantidad", "Dias_Emision_Radicacion", "Dias_Radicacion_Pago"]),
                 hide_index=True, height=520)


def mascara_atipicos_sesion(df):
    """Máscara de atípicos según la última configuración usada en la sección de atípicos."""
    config = st.session_state.get("_atipicos_config") or {}
    variable = config.get("variable", "Valor_Neto")
    if not columna_util(df, variable):
        return pd.Series(False, index=df.index)
    mascara, _ = detectar_atipicos(df, variable, config.get("metodo", "IQR"), config.get("multiplicador", 1.5),
                                   config.get("umbral_z", 3.0), config.get("percentiles", (1.0, 99.0)),
                                   config.get("alcance"))
    return mascara


UMBRAL_DESCARGA_DIFERIDA = 20_000


def mostrar_descargas(ctx):
    """Página 16: descargas generadas en memoria."""
    st.subheader("Descargas")
    df = ctx["df"]
    st.caption("Los archivos se generan en memoria y no se almacenan en el servidor. Los CSV usan punto y coma como "
               "separador y coma decimal para abrir correctamente en Excel con configuración regional colombiana.")
    if len(df) > LIMITE_FILAS_EXCEL:
        st.warning(f"El conjunto supera el límite de filas de Excel ({formatear_numero(LIMITE_FILAS_EXCEL)}); el "
                   "archivo Excel se truncará. Use la descarga CSV para obtener todos los registros.")
    clave = clave_hash(ctx.get("clave"), len(df), df.index[:50].tolist() if len(df) else [], df.index[-50:].tolist()
                       if len(df) else [])
    minimas = ctx.get("minimas_ok")
    preparar_pesados = True
    if len(df) > UMBRAL_DESCARGA_DIFERIDA:
        st.info(f"El conjunto filtrado tiene {formatear_numero(len(df))} registros. Para mantener la navegación "
                "fluida, los archivos con registros individuales se generan solo cuando usted lo solicita.")
        preparar_pesados = st.toggle("Preparar archivos de registros", key="preparar_descargas")

    def boton(etiqueta, nombre, tipo, generador, descripcion, pesado=False):
        """Genera el contenido de forma protegida y muestra el botón de descarga."""
        try:
            if pesado and not preparar_pesados:
                st.button(etiqueta, disabled=True, key=f"pend_{nombre}", width="stretch",
                          help="Active 'Preparar archivos de registros' para generar este archivo.")
                return
            with st.spinner(f"Preparando {nombre}..."):
                hojas = generador()
            if hojas is None:
                st.button(etiqueta, disabled=True, help="No disponible con las columnas actuales.", key=f"dis_{nombre}")
                return
            if tipo == "txt":
                contenido, mime = hojas, "text/plain"
            else:
                contenido = preparar_descarga(clave_hash(clave, nombre), tipo, hojas)
                mime = "text/csv" if tipo == "csv" else \
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.download_button(etiqueta, data=contenido, file_name=nombre, mime=mime, key=f"dl_{nombre}",
                               help=descripcion, width="stretch", on_click="ignore")
        except Exception as error:  # noqa: BLE001
            st.warning(f"No fue posible generar {nombre} ({type(error).__name__}).")

    cierre = (lambda: {"Cierre mensual": construir_cierre_mensual(df, "Mensual").drop(columns=["_ordinal"],
                                                                                      errors="ignore")}) \
        if minimas else (lambda: None)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Datos**")
        boton("1. Datos filtrados (CSV)", "facturacion_filtrada.csv", "csv", lambda: {"Datos": df},
              "Registros que cumplen los filtros actuales.", pesado=True)
        boton("2. Datos filtrados (Excel)", "facturacion_filtrada.xlsx", "excel", lambda: {"Facturacion": df},
              "Registros que cumplen los filtros actuales.", pesado=True)
        boton("8. Registros con glosa (CSV)", "registros_glosados.csv", "csv",
              (lambda: {"Glosas": df[df["Tiene_Glosa"]]}) if "Tiene_Glosa" in df.columns else (lambda: None),
              "Registros con Glosa_Inicial > 0.", pesado=True)
        boton("9. Registros con cartera (CSV)", "registros_cartera.csv", "csv",
              (lambda: {"Cartera": df[df["Tiene_Saldo"]]}) if "Tiene_Saldo" in df.columns else (lambda: None),
              "Registros con Saldo_Cartera > 0.", pesado=True)
        boton("10. Registros atípicos (CSV)", "registros_atipicos.csv", "csv",
              lambda: {"Atipicos": df[mascara_atipicos_sesion(df)]},
              "Según la configuración de la sección Valores atípicos (por defecto Valor_Neto, IQR 1,5).", pesado=True)
    with c2:
        st.markdown("**Análisis**")
        boton("3. Cierre mensual (CSV)", "cierre_mensual.csv", "csv", cierre, "Tabla mensual del cierre.")
        boton("4. Cierre mensual (Excel)", "cierre_mensual.xlsx", "excel", cierre, "Tabla mensual del cierre.")
        boton("5. Análisis por entidades (Excel)", "analisis_entidades.xlsx", "excel",
              (lambda: {"Entidades": agregar_por(df, "Entidad_Pagadora"),
                        "Pareto": calcular_pareto(df, "Entidad_Pagadora", "Valor_Neto")[0]})
              if "Entidad_Pagadora" in df.columns and minimas else (lambda: None), "Tabla comparativa y Pareto.")
        boton("6. Análisis por servicios (Excel)", "analisis_servicios.xlsx", "excel",
              (lambda: {k: agregar_por(df, c) for k, c in [("Servicios", "Servicio"), ("Lineas", "Linea_Servicio")]
                        if c in df.columns}) if minimas and ("Servicio" in df.columns or "Linea_Servicio" in df.columns)
              else (lambda: None), "Servicios y líneas de servicio.")
        boton("7. Análisis por sedes (Excel)", "analisis_sedes.xlsx", "excel",
              (lambda: {k: agregar_por(df, c) for k, c in [("Sedes", "Sede"), ("Regimenes", "Regimen")]
                        if c in df.columns}) if minimas and ("Sede" in df.columns or "Regimen" in df.columns)
              else (lambda: None), "Sedes y regímenes.")
    with c3:
        st.markdown("**Informes**")
        cal = ctx["calidad"]

        def informe_calidad():
            resumen = pd.DataFrame([
                ("Puntaje orientativo de calidad", round(cal["puntaje"], 1)),
                *[(f"Componente: {k}", round(v, 1)) for k, v in cal["componentes"].items()],
                ("Filas", cal["filas"]), ("Columnas", cal["columnas"]), ("Filas duplicadas", cal["filas_duplicadas"]),
                ("ID de factura duplicados", cal["ids_duplicados"]), ("Radicados duplicados", cal["radicados_duplicados"]),
                ("Fechas inválidas", cal["fechas_invalidas"]), ("Valores no numéricos", cal["no_numericos"]),
                ("Valores netos iguales a cero", cal["neto_cero"]), ("Sin entidad pagadora", cal["sin_entidad"]),
                ("Sin periodo", cal["sin_periodo"]),
                *[(k, v) for k, v in cal["banderas"].items()],
                ("Nota", AVISO_CALIDAD),
            ], columns=["Indicador", "Valor"])
            resumen["Valor"] = resumen["Valor"].astype(str)
            return {"Resumen": resumen, "Nulos": cal["tabla_nulos"], "Conversiones": ctx["reporte"],
                    "Limpieza": pd.DataFrame(ctx["bitacora"]) if ctx["bitacora"] else
                    pd.DataFrame({"Paso": ["Sin limpieza aplicada"]}),
                    "Inconsistencias": df[df["_Posible_Inconsistencia"]] if "_Posible_Inconsistencia" in df.columns
                    else pd.DataFrame()}

        boton("11. Informe de calidad (Excel)", "informe_calidad.xlsx", "excel", informe_calidad,
              "Diagnóstico, nulos, conversiones, limpieza y registros con posibles inconsistencias.", pesado=True)

        def resumen_insights():
            if not minimas:
                return None
            info = {"nombre": ctx["archivo"]["nombre"], "registros": formatear_numero(len(df))}
            return texto_insights(generar_insights(df, cal), generar_alertas(df, obtener_umbrales()), info)

        boton("12. Resumen de insights (TXT)", "insights_facturacion.txt", "txt", resumen_insights,
              "Semáforos e insights calculados sobre los datos filtrados.")


def mostrar_metodologia(ctx=None):
    """Página 17: metodología, fórmulas y limitaciones."""
    st.subheader("Metodología y transparencia")
    st.markdown(f"""
**Indicadores principales**

| Indicador | Cálculo |
|---|---|
| Facturación bruta | Suma de `Valor_Bruto` |
| Facturación neta | Suma de `Valor_Neto` (valor esperado ≈ `Valor_Bruto − Descuento − Copago_Cuota`) |
| Número de facturas | Valores únicos de `ID_Factura` (nunique), para no contar una factura varias veces |
| Porcentaje de glosa (severidad) | `Glosa_Inicial / Valor_Neto × 100` |
| Frecuencia de glosa | `Facturas con glosa / total de facturas × 100` |
| Porcentaje de recaudo | `Valor_Pagado / Valor_Neto × 100` |
| Porcentaje de cartera | `Saldo_Cartera / Valor_Neto × 100` |
| Ticket promedio | `Valor_Neto / número de facturas` |
| Días emisión-radicación | `Dias_Emision_Radicacion` o, si no existe, `Fecha_Radicacion − Fecha_Emision` |
| Crecimiento | `(Valor neto del periodo − periodo anterior) / periodo anterior × 100`, solo entre periodos consecutivos |
| Promedio móvil | Promedio de tres periodos consecutivos |

Cuando el denominador es cero o nulo, el indicador se reporta como **No calculable** (nunca infinito).

**Validaciones de identidad financiera** (con tolerancia configurable):
- `Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota`
- `Saldo_Cartera ≈ Valor_Neto − Glosa_Inicial − Valor_Pagado`

Las diferencias en la segunda identidad pueden reflejar glosas levantadas o conciliadas, notas crédito o pagos
posteriores; se reportan como posibles inconsistencias, no como errores confirmados.

**Puntaje orientativo de calidad (0-100):** promedio simple de cuatro dimensiones:
completitud (1 − % de nulos en columnas mínimas y recomendadas), unicidad (penaliza duplicados exactos y,
con medio peso, ID de factura repetidos), validez (fechas no interpretables, valores no numéricos y negativos
sobre las celdas evaluadas) y consistencia (registros con alguna regla de inconsistencia). {AVISO_CALIDAD}

**Correlación:** mide la asociación lineal (Pearson) o monótona (Spearman) entre dos variables, entre −1 y 1.
{AVISO_CORRELACION}

**Valor atípico:** observación alejada del patrón del resto. Método IQR: `IQR = Q3 − Q1`; límite inferior
`Q1 − k × IQR`; límite superior `Q3 + k × IQR` (k = 1,5 por defecto, configurable). También se ofrecen puntaje Z y
percentiles. Un atípico no es necesariamente un error.

**p-valor:** probabilidad de observar un resultado igual o más extremo que el obtenido si la hipótesis nula fuera
cierta. Un p-valor menor que α lleva a rechazar H0; un p-valor mayor no demuestra que H0 sea verdadera.

**Pareto y concentración:** participación individual y acumulada; el 80 % es una referencia analítica.
El índice HHI (suma de cuotas² × 10.000) se interpreta de forma orientativa.

**Rangos y umbrales:** los rangos de tiempo de radicación, antigüedad de cartera y los umbrales de semáforos son
parámetros analíticos configurables, no estándares normativos, contractuales ni contables.

**Limitaciones**
- Los resultados dependen de la calidad, completitud y estructura del archivo cargado.
- La herramienta es exploratoria y no reemplaza una auditoría financiera, contable, médica o contractual, ni es un
  sistema oficial de facturación, contabilidad o historia clínica.
- Las asociaciones estadísticas no demuestran causalidad.
- El último periodo puede estar incompleto y los periodos recientes suelen tener menor recaudo.
- Los insights se generan con reglas programadas; no utilizan modelos de inteligencia artificial ni servicios externos.

**Privacidad:** {AVISO_PRIVACIDAD} Los datos se procesan en memoria durante la sesión y no se envían a servicios externos.
""")


# =============================================================================
# 12. FUNCIÓN PRINCIPAL
# =============================================================================

FUNCIONES_PAGINA = {
    PAGINAS[2]: mostrar_resumen, PAGINAS[3]: mostrar_cierre, PAGINAS[4]: mostrar_entidades,
    PAGINAS[5]: mostrar_servicios, PAGINAS[6]: mostrar_sedes_regimenes, PAGINAS[7]: mostrar_estados,
    PAGINAS[8]: mostrar_glosas, PAGINAS[9]: mostrar_cartera, PAGINAS[10]: mostrar_tiempos,
    PAGINAS[11]: mostrar_estadistica, PAGINAS[12]: mostrar_atipicos, PAGINAS[13]: mostrar_insights,
    PAGINAS[14]: mostrar_explorador, PAGINAS[15]: mostrar_descargas,
}


def main():
    """Orquesta la configuración, el procesamiento, los filtros y la navegación."""
    configurar_pagina()
    aplicar_estilos()
    inicializar_estado()
    render_navegacion()
    contenedor_filtros = st.sidebar.container()
    render_configuracion_lateral(st.session_state.get("archivo") is not None)
    encabezado()

    ctx = None
    try:
        ctx = obtener_datos()
    except Exception as error:  # noqa: BLE001
        st.error(f"No fue posible leer el archivo ({type(error).__name__}).")
    if ctx is not None and not ctx.get("error") and "df" in ctx and not ctx["df"].empty:
        ctx["df_completo"] = ctx["df"]
        ctx["df"], ctx["filtros"] = render_filtros(contenedor_filtros, ctx["df"])
    st.sidebar.divider()
    st.sidebar.button("Reiniciar aplicación", on_click=cb_reiniciar, type="primary", width="stretch",
                      help="Elimina los datos de la sesión y vuelve a la pantalla inicial.")
    st.sidebar.caption("Datos procesados en memoria. Sin conexión a servicios externos.")

    pagina = st.session_state.get("pagina", PAGINAS[0])
    try:
        if pagina == PAGINAS[0]:
            mostrar_inicio(ctx)
            return
        if pagina == PAGINAS[16]:
            mostrar_metodologia(ctx)
            return
        if ctx is None:
            st.info("Cargue un archivo en '1. Inicio y carga de datos' para habilitar esta sección.")
            return
        if ctx.get("error"):
            st.error(ctx["error"])
            return
        if ctx["df_completo"].empty if "df_completo" in ctx else True:
            st.warning("El archivo no contiene registros después de la limpieza. Revise las opciones de limpieza.")
            return
        if pagina == PAGINAS[1]:
            mostrar_calidad(ctx)
            return
        if not ctx["minimas_ok"] and pagina not in PAGINAS_SIN_MINIMAS:
            st.error(f"Faltan columnas mínimas ({', '.join(ctx['faltantes'])}). Los análisis financieros están "
                     "desactivados. Configure la correspondencia de columnas en la página de inicio.")
            return
        if ctx["df"].empty:
            st.warning("Los filtros actuales no tienen registros. Use 'Limpiar filtros' en la barra lateral.")
            st.button("Limpiar filtros", on_click=cb_limpiar_filtros, key="limpiar_principal")
            return
        if len(ctx["df"]) < len(ctx["df_completo"]):
            st.caption(f"Vista filtrada: {formatear_numero(len(ctx['df']))} de "
                       f"{formatear_numero(len(ctx['df_completo']))} registros "
                       f"({formatear_porcentaje(dividir_seguro(len(ctx['df']), len(ctx['df_completo']), 100))}).")
        FUNCIONES_PAGINA[pagina](ctx)
    except Exception as error:  # noqa: BLE001
        st.error(f"Ocurrió un error inesperado en esta sección ({type(error).__name__}). Las demás secciones siguen "
                 "disponibles.")
        with st.expander("Detalle técnico"):
            st.write(str(error)[:500])


if __name__ == "__main__":
    main()
