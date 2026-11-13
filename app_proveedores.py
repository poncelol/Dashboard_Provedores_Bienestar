"""
Programa de Abasto Rural 2026 · Tablero de Proveedores
Subgerencia de Abarrotes · Alimentación para el Bienestar

Tablero para exposición: cada proveedor entra con su código y ve su información (surtido, precios dentro del
rango del mercado, cobertura y vigencias). Las comparaciones son por rangos y nunca muestran el nombre de otro
proveedor. Con el código de administrador se ve todo el programa y se puede abrir el tablero de cualquier proveedor.

Ejecutar:   streamlit run app_proveedores.py

Secretos (.streamlit/secrets.toml en tu computadora, o "Settings > Secrets" en Streamlit Community Cloud):
  [onedrive]               vínculo del Excel (se vuelve a descargar cada `minutos`)
  [acceso]                 admin = "código"  y  [acceso.proveedores]  "CVE" = "código"
  Sin [acceso] el tablero abre directo en modo administrador (útil para pruebas en tu computadora).
  Los códigos se generan con:   python generar_codigos.py
"""
import base64
import hmac
import io
import json
import re
import time
import unicodedata
from contextlib import contextmanager
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# ──────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Proveedores · Abasto Rural 2026",
    page_icon=":material/storefront:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CARPETA = Path(__file__).parent
NOMBRES_EXCEL = ["PROGRAMA DE ABASTO RURAL 2026.xlsx", "PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx"]
ARCHIVOS = [d / n for d in (CARPETA, CARPETA.parent / "Dashboard") for n in NOMBRES_EXCEL]
ARCHIVO = next((a for a in ARCHIVOS if a.exists()), ARCHIVOS[0])
MAPA = CARPETA / "mexico_estados.json"
ADMIN = "__admin__"
MAX_INTENTOS, ESPERA_SEG = 5, 120   # tras 5 códigos equivocados se espera 2 minutos

try:
    ZONA = ZoneInfo("America/Mexico_City")
except Exception:  # noqa: BLE001  (Windows sin tzdata)
    ZONA = None


def ahora():
    return datetime.now(ZONA) if ZONA else datetime.now()


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Gerencias regionales: sigla -> (nombre, clave, estados que cubre)
REGIONES = {
    "NOE": ("Noroeste", 41, ["Sinaloa", "Baja California", "Baja California Sur", "Sonora"]),
    "NOR": ("Norte", 42, ["Chihuahua", "Durango", "Coahuila"]),
    "NC": ("Norte Centro", 43, ["Zacatecas", "San Luis Potosí", "Aguascalientes", "Nuevo León"]),
    "GLF": ("Golfo", 44, ["Veracruz", "Tamaulipas"]),
    "OCC": ("Occidente", 45, ["Jalisco", "Nayarit", "Colima"]),
    "CEN": ("Centro", 46, ["Hidalgo", "Querétaro", "Guanajuato"]),
    "OAX": ("Oaxaca", 47, ["Oaxaca"]),
    "MET": ("Metropolitana", 48, ["México", "Morelos"]),
    "PEN": ("Peninsular", 49, ["Yucatán", "Campeche", "Quintana Roo"]),
    "STE": ("Sureste", 50, ["Tabasco", "Chiapas"]),
    "PSR": ("Pacífico Sur", 51, ["Michoacán", "Guerrero"]),
    "PUE": ("Puebla", 52, ["Puebla", "Tlaxcala"]),
}
ESTADO_REGION = {e: s for s, (_, _, est) in REGIONES.items() for e in est}

# Rangos que usan las gráficas: (límite superior incluido, etiqueta)
INF = float("inf")
RANGOS_ART = [(5, "1 a 5"), (10, "6 a 10"), (20, "11 a 20"), (50, "21 a 50"), (INF, "Más de 50")]
RANGOS_REG = [(0, "Sin dato"), (1, "1 región"), (4, "2 a 4"), (8, "5 a 8"), (11, "9 a 11"), (12, "Las 12 (nacional)")]
RANGOS_COSTO = [(10, "Hasta $10"), (20, "$10 a $20"), (50, "$20 a $50"), (100, "$50 a $100"), (INF, "Más de $100")]
RANGOS_VIG = [(-1, "Vencida"), (30, "0 a 30 días"), (60, "31 a 60 días"), (90, "61 a 90 días"), (INF, "Más de 90 días")]

# Paleta institucional
GUINDA, GUINDA_MED = "#611232", "#9D2449"
VERDE_MED, VERDE_CLARO = "#235B4E", "#3E8E7E"
ORO, AZUL = "#A57F2C", "#2E6B9E"
P626, P7420, P1255_60, P1255 = "#1E5B4F", "#9B2247", "#CFB579", "#A57F2C"
P627, P7421 = "#002F2A", "#611232"
CLAROS = {P1255_60}
BANDA, DORADO = "#4A0E25", "#BC955C"

LOGO = f"""<svg width="46" height="46" viewBox="0 0 120 120" aria-label="Alimentación para el Bienestar" role="img">
<defs>
 <linearGradient id="lg-c" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#EAD39B"/><stop offset="1" stop-color="#C9A563"/></linearGradient>
 <linearGradient id="lg-o" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#C99A3F"/><stop offset="1" stop-color="#8F6A22"/></linearGradient>
 <clipPath id="lg-k"><rect x="28" y="8" width="64" height="84"/></clipPath>
</defs>
<g clip-path="url(#lg-k)">
 <rect x="28" y="8" width="64" height="84" fill="url(#lg-c)"/>
 <path stroke="{BANDA}" stroke-width="2.4"
  d="M60 60 L20 40 M60 60 L20 10 M60 60 L42 0 M60 60 L60 0 M60 60 L78 0 M60 60 L100 10 M60 60 L100 40"/>
 <circle cx="60" cy="58" r="17" fill="url(#lg-o)" stroke="{BANDA}" stroke-width="2.6"/>
 <g stroke="{BANDA}" stroke-width="2.6">
  <path d="M20 56 C40 58 54 70 60 96 L20 96 Z" fill="url(#lg-c)"/><path d="M100 56 C80 58 66 70 60 96 L100 96 Z" fill="url(#lg-c)"/>
  <path d="M20 66 C38 68 52 78 60 96 L20 96 Z" fill="#B88A3A"/><path d="M100 66 C82 68 68 78 60 96 L100 96 Z" fill="#B88A3A"/>
  <path d="M20 76 C36 78 50 86 60 96 L20 96 Z" fill="url(#lg-c)"/><path d="M100 76 C84 78 70 86 60 96 L100 96 Z" fill="url(#lg-c)"/>
  <path d="M20 85 C34 87 48 91 60 96 L20 96 Z" fill="#B88A3A"/><path d="M100 85 C86 87 72 91 60 96 L100 96 Z" fill="#B88A3A"/>
 </g>
</g>
<path d="M14 93 Q60 116 106 89 Q62 106 14 93 Z" fill="url(#lg-o)"/>
</svg>"""

# Íconos de línea (Lucide, licencia ISC): heredan el color del texto
ICONOS = {
    "proveedor": '<path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/>'
                 '<path d="M17 18h1"/><path d="M12 18h1"/><path d="M7 18h1"/>',
    "producto": '<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 '
                '2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/>',
    "articulo": '<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 '
                '0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>',
    "marca": '<path d="M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 '
             '2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z"/><circle cx="7.5" cy="7.5" r=".5" fill="currentColor"/>',
    "precio": '<line x1="12" x2="12" y1="2" y2="22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
    "estrella": '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
    "mapa": '<path d="M14.106 5.553a2 2 0 0 0 1.788 0l3.659-1.83A1 1 0 0 1 21 4.619v12.764a1 1 0 0 1-.553.894l-4.553 '
            '2.277a2 2 0 0 1-1.788 0l-4.212-2.106a2 2 0 0 0-1.788 0l-3.659 1.83A1 1 0 0 1 3 19.381V6.618a1 1 0 0 1 '
            '.553-.894l4.553-2.277a2 2 0 0 1 1.788 0z"/><path d="M15 5.764v15"/><path d="M9 3.236v15"/>',
    "idea": '<path d="M15 14c.2-1 .7-1.7 1.5-2.5 1-.9 1.5-2.2 1.5-3.5A6 6 0 0 0 6 8c0 1 .2 2.2 1.5 3.5.7.7 1.3 1.5 1.5 2.5"/>'
            '<path d="M9 18h6"/><path d="M10 22h4"/>',
    "balanza": '<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 '
               '1s-2.13-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>',
    "rangos": '<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/>',
    "calendario": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "hoja": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/>'
            '<path d="M16 13H8"/><path d="M16 17H8"/><path d="M10 9H8"/>',
    "candado": '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    "lista": '<path d="M3 12h.01"/><path d="M3 18h.01"/><path d="M3 6h.01"/><path d="M8 12h13"/><path d="M8 18h13"/>'
             '<path d="M8 6h13"/>',
    "buscar": '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    "ok": '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
    "alerta": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/>'
              '<path d="M12 17h.01"/>',
}


def svg(nombre, tam=18):
    return (f'<svg width="{tam}" height="{tam}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
            f'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICONOS[nombre]}</svg>')


TEMAS = {
    "claro": dict(
        fondo="#f3f4f6", card="#ffffff", sutil="#f9fafb", linea="#f3f4f6", borde="#e5e7eb",
        texto="#1f2937", texto2="#4b5563", texto3="#6b7280", tenue="#9ca3af",
        acento=GUINDA_MED, acento_rgb="157 36 73", ok="#239f52", alerta="#e63939", otros="#d1d5db",
        verde=VERDE_MED, oro=ORO, azul=AZUL, verde_claro=VERDE_CLARO, banda=BANDA, dorado=DORADO,
        sombra="0 1px 1px 0 rgb(0 0 0 / .05), 0 1px 2px 0 rgb(0 0 0 / .02)",
    ),
    "oscuro": dict(
        fondo="#111827", card="#1f2937", sutil="#263142", linea="#2d3748", borde="#374151",
        texto="#f3f4f6", texto2="#d1d5db", texto3="#9ca3af", tenue="#6b7280",
        acento="#E0698C", acento_rgb="224 105 140", ok="#4bd37d", alerta="#ff7474", otros="#4b5563",
        verde="#4FB39A", oro="#D9AE55", azul="#6AA6DE", verde_claro="#8FD3C2", banda=BANDA, dorado=DORADO,
        sombra="0 1px 2px 0 rgb(0 0 0 / .35)",
    ),
}

# ──────────────────────────────────────────────────────────────
# CARGA Y LIMPIEZA (misma lógica que app_mosaic.py)
# ──────────────────────────────────────────────────────────────
RENOMBRAR = {
    "NEGOCIACION": "NEGOCIACION", "CVE": "CVE", "PROVEEDOR": "PROVEEDOR",
    "PROD.": "PROD", "PRODUCTOS": "PRODUCTO", "ART.": "ART_ID", "ARTICULOS": "ARTICULO",
    "MARCA": "MARCA", "CAP": "CAP", "GRAM": "GRAM", "U/M": "UM",
    "COSTO DE ADQUISICIÓN 2026": "COSTO_CAJA",
    "COSTO UNITARIO DE ADQUISICIÓN 2026": "COSTO_UNIT",
    "TIPO DE INVENTARIO": "INVENTARIO", "COBERTURA": "COBERTURA", "VIGENCIA": "VIGENCIA",
}
NUMERICAS = ["CAP", "GRAM", "COSTO_CAJA", "COSTO_UNIT"]
INVENTARIO_LABEL = {"PAR": "PAR", "OFERTA": "Oferta"}
# Unidad de medida -> (factor a la unidad base, unidad base): así se comparan presentaciones distintas
UNIDAD_BASE = {"GR": (0.001, "kg"), "KG": (1, "kg"), "ML": (0.001, "L"), "LT": (1, "L"), "PZ": (1, "pza")}


def nombre_corto(nombre: str) -> str:
    corto_ = re.sub(r"[,.\s]+(S\.?\s?A\.?|S\.?\s?DE\s?R\.?\s?L\.?|S\.?\s?C\.?|S\.?\s?P\.?\s?R\.?)(?![A-Z]).*$", "",
                    str(nombre).strip(), flags=re.I)
    return corto_.strip(" ,.").title() or str(nombre)


def clave_texto(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return s.where(~s.isin(["", "nan", "NaN", "<NA>", "None"]), pd.NA)


def presentacion(cap, gram, um) -> str:
    if pd.isna(cap) or pd.isna(gram):
        return "–"
    return f"{cap:g} × {gram:g} {um}".strip()


@st.cache_data(show_spinner="Preparando el tablero…")
def cargar(contenido: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(contenido))
    hoja = "VIGENTES" if "VIGENTES" in xls.sheet_names else xls.sheet_names[0]
    vista = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=30)
    fila = next((i for i, r in vista.iterrows()
                 if r.astype(str).str.strip().str.upper().isin(["NEGOCIACION", "CVE"]).any()), None)
    if fila is None:
        raise ValueError(f"no encontré la fila de encabezados en la hoja {hoja}")
    raw = pd.read_excel(xls, sheet_name=hoja, header=fila)
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]
    raw = raw.rename(columns=RENOMBRAR)
    raw = raw.loc[:, ~raw.columns.str.startswith("Unnamed")]
    raw = raw.loc[:, ~raw.columns.duplicated()]
    faltan = [c for c in ["CVE", "PROVEEDOR", "PRODUCTO", "ART_ID", "ARTICULO"] if c not in raw]
    if faltan:
        raise ValueError("faltan las columnas " + ", ".join(faltan))

    df = raw[raw["ART_ID"].notna()].copy()
    for c in ["NEGOCIACION", "PROD", "MARCA", "UM", "INVENTARIO", "COBERTURA", "VIGENCIA", *NUMERICAS]:
        if c not in df:
            df[c] = pd.NA
    for c in NUMERICAS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["NEGOCIACION", "CVE", "PROD", "ART_ID"]:
        df[c] = clave_texto(df[c])
    for c in ["PROVEEDOR", "PRODUCTO", "ARTICULO", "MARCA", "UM"]:
        df[c] = df[c].astype(str).str.strip().str.upper().replace({"NAN": pd.NA, "": pd.NA})
    df = df[df["CVE"].notna()]
    df["MARCA"] = df["MARCA"].fillna("SIN MARCA").str.title()
    df["PRODUCTO"] = df["PRODUCTO"].fillna("SIN CLASIFICAR").str.capitalize()
    df["ARTICULO"] = df["ARTICULO"].fillna("SIN NOMBRE")
    df["UM"] = df["UM"].fillna("")
    df["PROV_CORTO"] = df["PROVEEDOR"].map(nombre_corto)
    df["PROD"] = df["PROD"].fillna(df["PRODUCTO"])
    df["PRODUCTO"] = df.groupby("PROD")["PRODUCTO"].transform(lambda s: s.mode().iat[0])

    inv = clave_texto(df["INVENTARIO"]).str.upper()
    df["INVENTARIO"] = inv.map(lambda x: INVENTARIO_LABEL.get(x, x.capitalize()), na_action="ignore").fillna("Sin dato")

    cob = df["COBERTURA"].astype(str).str.upper().str.replace(r"[^A-Z]+", " ", regex=True).str.strip()
    fichas = cob.where(~cob.isin(["", "NAN", "NA"]), "").str.split()
    df["ALCANCE"] = ["Nacional" if "NAL" in f else ("Regional" if f else "Sin dato") for f in fichas]
    df["REGIONES"] = [tuple(REGIONES) if "NAL" in f else tuple(r for r in REGIONES if r in f) for f in fichas]
    df["VIG_FECHA"] = pd.to_datetime(df["VIGENCIA"], errors="coerce")
    df["PRESENTACION"] = [presentacion(c, g, u) for c, g, u in zip(df["CAP"], df["GRAM"], df["UM"])]

    factor = pd.to_numeric(df["UM"].map(lambda u: UNIDAD_BASE.get(u, (None, None))[0]), errors="coerce")
    df["BASE"] = df["UM"].map(lambda u: UNIDAD_BASE.get(u, (None, None))[1])
    cantidad = df["GRAM"] * factor
    df["PRECIO_U"] = (df["COSTO_UNIT"] / cantidad).where(cantidad > 0)
    return df.reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# DATOS (OneDrive o archivo local)
# ──────────────────────────────────────────────────────────────
def secreto(seccion_: str) -> dict:
    try:
        return dict(st.secrets.get(seccion_, {}))
    except Exception:  # noqa: BLE001  (no existe secrets.toml)
        return {}


def codigo_share(link: str) -> str:
    return "u!" + base64.urlsafe_b64encode(link.strip().encode()).decode().rstrip("=")


def url_descarga(link: str) -> str:
    link = link.strip()
    if "sharepoint.com" in link:
        partes = urlsplit(link)
        query = [(k, v) for k, v in parse_qsl(partes.query) if k != "download"] + [("download", "1")]
        return urlunsplit(partes._replace(query=urlencode(query)))
    return f"https://api.onedrive.com/v1.0/shares/{codigo_share(link)}/root/content"


@st.cache_data(max_entries=2, show_spinner="Descargando el Excel de OneDrive…")
def descargar_excel(link: str, bloque: int):
    """`bloque` cambia cada N minutos, así el Excel se vuelve a bajar solo."""
    cfg = secreto("onedrive")
    encabezados = {"User-Agent": "Mozilla/5.0"}
    if all(cfg.get(k) for k in ("tenant_id", "client_id", "client_secret")):
        token = requests.post(
            f"https://login.microsoftonline.com/{cfg['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": cfg["client_id"], "client_secret": cfg["client_secret"],
                  "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"},
            timeout=30)
        token.raise_for_status()
        encabezados["Authorization"] = f"Bearer {token.json()['access_token']}"
        url = f"https://graph.microsoft.com/v1.0/shares/{codigo_share(link)}/driveItem/content"
    else:
        url = url_descarga(link)
    r = requests.get(url, headers=encabezados, timeout=60, allow_redirects=True)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):
        raise ValueError("OneDrive no devolvió un Excel: revisa que el vínculo permita ver el archivo.")
    return r.content, ahora()


def motivo(e: Exception) -> str:
    if isinstance(e, requests.HTTPError) and e.response is not None:
        return f"HTTP {e.response.status_code}"
    if isinstance(e, requests.RequestException):
        return "no hubo conexión con OneDrive"
    return str(e)


@st.cache_data
def geo_estados():
    return json.loads(MAPA.read_text(encoding="utf-8")) if MAPA.exists() else None


@st.cache_data
def centros_region():
    gj = geo_estados()
    if not gj:
        return {}
    cajas = {}
    for f in gj["features"]:
        g = f["geometry"]
        pol = g["coordinates"] if g["type"] == "MultiPolygon" else [g["coordinates"]]
        pts = [p for p_ in pol for p in p_[0]]
        lons, lats = [p[0] for p in pts], [p[1] for p in pts]
        cajas[f["properties"]["name"]] = ((min(lons) + max(lons)) / 2, (min(lats) + max(lats)) / 2)
    salida = {}
    for sig, (_, _, estados) in REGIONES.items():
        c = [cajas[e] for e in estados if e in cajas]
        if c:
            salida[sig] = (sum(x for x, _ in c) / len(c), sum(y for _, y in c) / len(c))
    return salida


# ──────────────────────────────────────────────────────────────
# TEMA Y ESTILO
# ──────────────────────────────────────────────────────────────
def md(texto, destino=st):
    destino.markdown(texto, unsafe_allow_html=True)


if "oscuro" not in st.session_state:
    st.session_state["oscuro"] = st.query_params.get("tema") == "oscuro"
oscuro = st.session_state["oscuro"]
st.query_params["tema"] = "oscuro" if oscuro else "claro"
T = TEMAS["oscuro" if oscuro else "claro"]
COLOR_VIG = {"Vencida": T["alerta"], "0 a 30 días": T["oro"], "31 a 60 días": T["azul"], "61 a 90 días": T["verde_claro"],
             "Más de 90 días": T["verde"], "Sin fecha": T["tenue"]}

VARIABLES = ":root{" + "".join(f"--{k.replace('_', '-')}:{v};" for k, v in T.items() if isinstance(v, str)) + "}"
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, .stApp, [class*="css"], button, input, textarea, select {
  font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif; }
.stApp { background: var(--fondo); color: var(--texto2); }
.block-container { padding-top: .5rem !important; padding-bottom: 3rem; max-width: 1440px; }
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent; pointer-events: none; }
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"] {
  display: none !important; }
h1, h2, h3, h4 { font-family: 'Inter', sans-serif !important; color: var(--texto); letter-spacing: -.01em; }
.stApp p, .stApp label { color: var(--texto2); }
[data-testid="stCaptionContainer"] { color: var(--texto3); }
[data-testid="stIconMaterial"] { color: var(--texto3); }

/* Controles de Streamlit siguiendo el tema */
[data-testid="stCheckbox"] label > div:not([data-testid]) { background: var(--borde) !important; }
[data-testid="stCheckbox"][data-selected="true"] label > div:not([data-testid]) { background: var(--acento) !important; }
[data-testid="stCheckbox"] label > div:not([data-testid]) > div { background: #fff !important; }
[data-testid="stCheckbox"] p, [data-testid="stCheckbox"] span[role="img"] { color: var(--texto2) !important; }
[data-testid="stSelectbox"] div[role="group"][data-rac], [data-testid="stTextInputRootElement"] {
  background: var(--sutil) !important; border: 1px solid var(--borde) !important; border-radius: 8px !important; }
[data-testid="stSelectbox"] div[role="group"][data-focus-within="true"],
[data-testid="stTextInputRootElement"]:focus-within { border-color: var(--acento) !important; }
[data-testid="stSelectbox"] input, [data-testid="stTextInputField"], [data-testid="stTextInputRootElement"] input {
  color: var(--texto) !important; background: transparent !important; caret-color: var(--acento); }
[data-testid="stTextInputRootElement"] input::placeholder { color: var(--tenue) !important; opacity: 1; }
[data-testid="stSelectbox"] button, [data-testid="stTextInputRootElement"] button {
  color: var(--texto3) !important; background: transparent !important; }
[data-testid$="Dropdown"] { background: var(--card) !important; border: 1px solid var(--borde) !important;
  border-radius: 8px !important; box-shadow: 0 8px 24px rgb(0 0 0 / .18) !important; }
[data-testid$="Dropdown"] [role="listbox"] { background: var(--card) !important; }
[data-testid$="Dropdown"] [role="option"] { color: var(--texto2) !important; background: transparent !important; }
[data-testid$="Dropdown"] [role="option"][data-focused="true"],
[data-testid$="Dropdown"] [role="option"]:hover { background: var(--sutil) !important; color: var(--texto) !important; }
[data-testid$="Dropdown"] [role="option"][aria-selected="true"] { color: var(--acento) !important; font-weight: 600; }

/* Encabezado institucional */
.st-key-hero { position: relative; background: var(--banda); padding: 22px 0 26px 0 !important; margin-bottom: 28px;
  box-shadow: 0 0 0 100vmax var(--banda); clip-path: inset(-100vmax -100vmax 0 -100vmax); flex-wrap: nowrap !important; }
.st-key-hero::after { content: ""; position: absolute; left: -100vmax; right: -100vmax; bottom: 0; height: 4px;
  background: var(--dorado); }
.st-key-hero > .stElementContainer { flex: 1 1 auto; min-width: 0; }
.hero .marca { display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }
.hero .marca svg { flex-shrink: 0; }
.hero .marca .sep { width: 1px; align-self: stretch; background: rgb(255 255 255 / .45); }
.hero .marca .gob { font-size: 2.4rem; color: #fff; line-height: 1; white-space: nowrap; font-weight: 800; }
.hero h1 { font-size: 2.3rem !important; font-weight: 700 !important; color: #fff !important; margin: 4px 0 0 0;
  padding: 0 !important; line-height: 1.2; }
.hero .gerencia { margin: 6px 0 0 0; font-size: 1.35rem; font-weight: 600; color: rgb(255 255 255 / .95); }
.hero .subgerencia { margin: 4px 0 0 0; font-size: 1.1rem; font-weight: 700; color: var(--dorado); }
.hero .fecha { margin: 14px 0 8px 0; font-size: .82rem; color: rgb(255 255 255 / .75); }
.st-key-controles { flex: 0 0 auto; align-items: flex-end; }
.st-key-controles [data-testid="stCheckbox"] p, .st-key-controles [data-testid="stCheckbox"] span[role="img"],
.st-key-controles [data-testid="stIconMaterial"] { color: #fff !important; }
.st-key-controles [data-testid="stCheckbox"] label > div:not([data-testid]) { background: rgb(255 255 255 / .3) !important; }
.st-key-controles [data-testid="stCheckbox"][data-selected="true"] label > div:not([data-testid]) {
  background: var(--dorado) !important; }
.st-key-controles button { background: transparent !important; border: 1px solid rgb(255 255 255 / .35) !important;
  color: #fff !important; border-radius: 999px !important; font-size: .8rem; }
.st-key-controles button:hover { border-color: var(--dorado) !important; color: var(--dorado) !important; }
.st-key-controles button * { color: inherit !important; }

/* Pantalla de acceso */
.st-key-login { max-width: 460px; margin: 10px auto 0 auto; background: var(--card); border-radius: 14px;
  box-shadow: 0 10px 30px rgb(0 0 0 / .08); padding: 28px 28px 20px 28px !important; border-top: 5px solid var(--acento); }
.login-cab { text-align: center; }
.login-cab .ico { width: 58px; height: 58px; border-radius: 16px; display: inline-grid; place-items: center;
  background: rgb(var(--acento-rgb) / .1); color: var(--acento); margin-bottom: 10px; }
.login-cab h2 { font-size: 1.35rem !important; margin: 0 !important; padding: 0 !important; }
.login-cab p { font-size: .86rem; color: var(--texto3); margin: 4px 0 6px 0; }
.st-key-login [data-testid="stFormSubmitButton"] button { background: var(--acento) !important; border: none !important;
  color: #fff !important; border-radius: 8px !important; font-weight: 600; }
.st-key-login [data-testid="stFormSubmitButton"] button * { color: #fff !important; }
.aviso { display: flex; gap: 8px; align-items: center; font-size: .84rem; border-radius: 8px; padding: 8px 12px;
  margin-top: 8px; }
.aviso.mal { color: var(--alerta); background: rgb(230 57 57 / .08); }
.aviso svg { flex-shrink: 0; }

/* Ficha del proveedor bajo el encabezado */
.ficha { display: flex; align-items: center; gap: 16px; padding: 18px 22px; border-left: 5px solid var(--acento) !important; }
.ficha .ico { width: 54px; height: 54px; border-radius: 14px; display: grid; place-items: center; flex-shrink: 0;
  background: rgb(var(--acento-rgb) / .12); color: var(--acento); }
.ficha .t { font-size: .74rem; font-weight: 600; color: var(--tenue); text-transform: uppercase; letter-spacing: .05em; }
.ficha b { display: block; font-size: 1.55rem; color: var(--texto); line-height: 1.2; }
.ficha span { font-size: .84rem; color: var(--texto3); }
.ficha .etq { margin-left: auto; display: flex; flex-wrap: wrap; gap: 6px; justify-content: flex-end; }
.etq-c { display: inline-flex; align-items: center; gap: 6px; font-size: .78rem; font-weight: 600; color: var(--texto2);
  border: 1px solid var(--borde); background: var(--sutil); border-radius: 999px; padding: 4px 12px; white-space: nowrap; }
.etq-c svg { color: var(--acento); }

/* Secciones e índice */
.seccion { display: flex; align-items: flex-start; gap: 12px; margin: 40px 0 16px 0; scroll-margin-top: 70px; }
.seccion h2 { font-size: 1.3rem !important; font-weight: 700 !important; margin: 0; padding: 0 !important; }
.seccion p { margin: 0; font-size: .84rem; color: var(--texto3); }
.indice { display: flex; flex-wrap: wrap; gap: 8px; margin: 18px 0 8px 0; }
.indice a { display: inline-flex; align-items: center; gap: 8px; text-decoration: none !important; font-size: .82rem;
  font-weight: 500; color: var(--texto2) !important; background: var(--card); border: 1px solid var(--borde);
  border-radius: 999px; padding: 6px 14px 6px 6px; transition: all .15s; }
.indice a:hover { border-color: var(--acento); color: var(--acento) !important; }
.indice a span, .seccion .paso { width: 22px; height: 22px; border-radius: 50%; display: grid; place-items: center;
  background: var(--acento); color: #fff; font-size: .72rem; font-weight: 700; flex-shrink: 0; }
.seccion .paso { width: 30px; height: 30px; font-size: .9rem; box-shadow: 0 0 0 2px var(--fondo), 0 0 0 3.5px var(--acento); }

/* Tarjetas */
.m-card, div[class*="st-key-card_"] { background: var(--card); border-radius: 12px; border: none !important;
  box-shadow: var(--sombra); }
div[class*="st-key-card_"] { padding: 0 20px 16px 20px !important; gap: .5rem; }
.card-h { margin: 0 -20px 8px -20px; padding: 16px 20px; border-bottom: 1px solid var(--linea); flex-wrap: wrap;
  display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
.card-h > div:first-child { flex: 1 1 180px; min-width: 0; }
.card-h h2 { font-size: 1rem !important; font-weight: 600 !important; margin: 0; padding: 0; color: var(--texto); }
.card-h p { font-size: .8rem; color: var(--tenue); margin: 2px 0 0 0; }
.card-h .extras { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 6px; }
.card-h .extra { font-size: .75rem; font-weight: 600; color: var(--acento); white-space: nowrap;
  background: rgb(var(--acento-rgb) / .1); border-radius: 999px; padding: 2px 10px; }
[data-testid="stPlotlyChart"] { background: transparent; }

/* KPIs */
.kpis { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 20px; margin-bottom: 20px; }
.kpi { padding: 18px; display: flex; gap: 14px; align-items: flex-start; }
.kpi > div:last-child { min-width: 0; flex: 1; }
.kpi .ico { width: 46px; height: 46px; border-radius: 12px; flex-shrink: 0; display: grid; place-items: center;
  background: var(--c); color: var(--ci, #fff); }
.kpi .l { font-size: .74rem; font-weight: 600; color: var(--tenue); text-transform: uppercase; letter-spacing: .04em; }
.kpi .v { font-size: 2rem; font-weight: 700; color: var(--texto); line-height: 1.15; margin-top: 2px; }
.kpi .s { font-size: .8rem; color: var(--texto3); margin-top: 4px; }

/* Lectura rápida */
.hallazgos { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
.hz { display: flex; gap: 12px; align-items: flex-start; padding: 16px; font-size: .9rem; color: var(--texto2);
  border: 1.5px solid var(--c) !important; border-left-width: 5px !important; line-height: 1.45; }
.hz .ico { width: 38px; height: 38px; border-radius: 50%; flex-shrink: 0; display: grid; place-items: center;
  color: var(--ci, #fff); background: var(--c); }
.hz b { color: var(--texto); font-weight: 700; }

/* Guía y leyenda de rangos */
.guia { display: flex; gap: 10px; align-items: flex-start; background: rgb(188 149 92 / .12);
  border-radius: 8px; padding: 10px 14px; font-size: .82rem; color: var(--texto2); margin: 0 0 6px 0; }
.guia svg { color: var(--oro); flex-shrink: 0; margin-top: 1px; }
.guia b { color: var(--texto); }
.leyenda { display: flex; flex-wrap: wrap; gap: 6px 14px; font-size: .78rem; color: var(--texto2); margin: 2px 0 4px 0; }
.leyenda span { display: inline-flex; align-items: center; gap: 6px; }
.leyenda i { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
.leyenda i.punto { border-radius: 50%; }
.leyenda i.rombo { transform: rotate(45deg) scale(.8); border-radius: 2px; }

/* Tablas */
.tabla-wrap { overflow-x: auto; }
.tabla-wrap.scroll { overflow-y: auto; }
.tabla { width: 100%; border-collapse: collapse; font-size: .86rem; }
.tabla thead th { font-size: .7rem; text-transform: uppercase; color: var(--tenue); font-weight: 600;
  background: var(--sutil); padding: 8px; text-align: left; border: none; white-space: nowrap; }
.tabla-wrap.scroll thead th { position: sticky; top: 0; z-index: 1; }
.tabla thead th:first-child { border-radius: 6px 0 0 6px; }
.tabla thead th:last-child { border-radius: 0 6px 6px 0; }
.tabla td { padding: 9px 8px; border: none; border-bottom: 1px solid var(--linea); color: var(--texto2); font-weight: 500; }
.tabla tr:last-child td { border-bottom: none; }
.tabla td.fuerte { color: var(--texto); }
.tabla .num { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
.tabla small { display: block; color: var(--tenue); font-size: .74rem; font-weight: 500; }
.chip { display: inline-flex; align-items: center; gap: 5px; font-size: .74rem; font-weight: 600; border-radius: 999px;
  padding: 1px 9px; white-space: nowrap; border: 1px solid var(--borde); color: var(--texto2); }
.chip i { width: 7px; height: 7px; border-radius: 50%; }
.chip.bajo { color: var(--ok); border-color: rgb(62 201 114 / .5); background: rgb(62 201 114 / .1); }
.chip.medio { color: var(--oro); border-color: rgb(188 149 92 / .55); background: rgb(188 149 92 / .1); }
.chip.alto { color: var(--alerta); border-color: rgb(230 57 57 / .45); background: rgb(230 57 57 / .08); }

/* Buscador de artículos */
.st-key-buscar_art { position: relative; gap: 0 !important; }
.st-key-buscar_art .stElementContainer:has(.lupa) { position: absolute; left: 12px; top: 0; bottom: 0; z-index: 1;
  display: flex; align-items: center; width: auto !important; pointer-events: none; }
.st-key-buscar_art .lupa { display: flex; color: var(--tenue); }
.st-key-buscar_art [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
.st-key-buscar_art input { padding-left: 34px !important; }

.lb-mas { font-size: .84rem; color: var(--tenue); padding: 10px 0; display: flex; gap: 8px; align-items: center; }
.lb-mas svg { color: var(--ok); }
.pie { text-align: center; margin-top: 40px; font-size: .75rem; color: var(--tenue); }

@media (max-width: 1100px) {
  .kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
  .hallazgos { grid-template-columns: repeat(2, 1fr); }
}
@media (max-width: 900px) {
  .hero h1 { font-size: 1.9rem !important; }
  .hero .marca .gob { font-size: 1.8rem; }
  .ficha { flex-wrap: wrap; }
  .ficha .etq { margin-left: 0; justify-content: flex-start; }
}
@media (max-width: 640px) {
  .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
  .st-key-hero { padding: 12px 0 20px 0 !important; flex-wrap: wrap !important; row-gap: 6px; }
  .hero h1 { font-size: 1.6rem !important; }
  .hero .marca .gob { font-size: 1.15rem; white-space: normal; }
  .hero .gerencia { font-size: 1.05rem; }
  .kpis, .hallazgos { grid-template-columns: minmax(0, 1fr); }
  .kpi .v { font-size: 1.5rem; }
  .ficha b { font-size: 1.2rem; }
  div[class*="st-key-card_"] { padding: 0 14px 12px 14px !important; }
  .card-h { margin: 0 -14px 8px -14px; padding: 14px; }
}
"""
md(f"<style>{VARIABLES}{CSS}</style>")


# ──────────────────────────────────────────────────────────────
# COMPONENTES
# ──────────────────────────────────────────────────────────────
def fecha_larga(d) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def dinero(x) -> str:
    return "–" if pd.isna(x) else f"${x:,.2f}"


def precio_u(v, base):
    return "–" if pd.isna(v) else f"${v:,.2f}/{base}"


def alfa(color, a):
    c = color.lstrip("#")
    return f"rgba({int(c[0:2], 16)},{int(c[2:4], 16)},{int(c[4:6], 16)},{a})"


def corto(texto, n=30):
    texto = str(texto)
    return texto if len(texto) <= n else texto[:n - 1].rstrip() + "…"


def norm_txt(t) -> str:
    return unicodedata.normalize("NFD", str(t)).encode("ascii", "ignore").decode().lower()


def en_rangos(valores: pd.Series, rangos) -> pd.Series:
    """Asigna a cada valor la etiqueta de su rango (límites superiores incluidos)."""
    bordes = [-INF] + [r[0] for r in rangos]
    return pd.cut(valores, bins=bordes, labels=[r[1] for r in rangos], right=True)


def conteo_rangos(valores: pd.Series, rangos) -> pd.Series:
    return en_rangos(valores, rangos).value_counts().reindex([r[1] for r in rangos], fill_value=0)


@contextmanager
def tarjeta(titulo, clave, sub="", extra=""):
    extras = [extra] if isinstance(extra, str) else extra
    chips = "".join(f"<span class=extra>{e}</span>" for e in extras if e)
    with st.container(border=True, key=f"card_{clave}"):
        md(f'<div class="card-h"><div><h2>{escape(titulo)}</h2>{f"<p>{sub}</p>" if sub else ""}</div>'
           f'{f"<div class=extras>{chips}</div>" if chips else ""}</div>')
        yield


def guia(texto):
    md(f'<div class="guia">{svg("idea", 16)}<span>{texto}</span></div>')


def tabla(encabezados, filas, numericas=(), alto=None):
    th = "".join(f'<th class="{"num" if i in numericas else ""}">{h}</th>' for i, h in enumerate(encabezados))
    tr = "".join("<tr>" + "".join(f'<td class="{"num" if i in numericas else ""} {"fuerte" if i == 0 else ""}">{v}</td>'
                                  for i, v in enumerate(fila)) + "</tr>" for fila in filas)
    marco = f' scroll" style="max-height:{alto}px' if alto else ""
    md(f'<div class="tabla-wrap{marco}"><table class="tabla"><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>')


def kpi_html(etiqueta, valor, sub, icono, color):
    return (f'<div class="m-card kpi" style="--c:{color};--ci:{P627 if color in CLAROS else "#fff"}">'
            f'<div class="ico">{svg(icono, 24)}</div><div><div class="l">{etiqueta}</div>'
            f'<div class="v">{valor}</div><div class="s">{sub}</div></div></div>')


def lectura_rapida(hallazgos):
    md('<div class="hallazgos">' + "".join(
        f'<div class="m-card hz" style="--c:{c};--ci:{P627 if c in CLAROS else "#fff"}"><div class="ico">{svg(i, 18)}'
        f'</div><div>{t}</div></div>' for c, i, t in hallazgos) + "</div>")


def estilo(fig, alto=380, leyenda=False):
    fig.update_layout(
        height=alto, margin=dict(l=8, r=24, t=20, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=T["texto3"], size=12),
        showlegend=leyenda,
        legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", title=None, font=dict(color=T["texto2"])),
        hoverlabel=dict(bgcolor=T["card"], bordercolor=T["borde"], font_size=12, font_family="Inter, sans-serif",
                        font_color=T["texto"]),
        bargap=0.3, barcornerradius=4,
    )
    fig.update_xaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["texto3"]))
    fig.update_yaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["texto3"]))
    return fig


def mostrar(fig, clave):
    st.plotly_chart(fig, width="stretch", theme=None, key=clave, config={"displayModeBar": False})


def barras_rangos(conteo: pd.Series, clave, unidad, resaltar=None, colores=None, nota="Aquí estás", alto=300):
    """Columnas por rango. Con `resaltar`, ese rango va en el color de acento y lleva la nota encima."""
    if colores is None:
        colores = [T["acento"] if k == resaltar else alfa(T["verde"], .55 if resaltar else .85) for k in conteo.index]
    total = max(int(conteo.sum()), 1)
    fig = go.Figure(go.Bar(
        x=list(conteo.index), y=conteo.values, marker_color=colores,
        text=[f"{v:,}" for v in conteo.values], textposition="outside", cliponaxis=False,
        textfont=dict(color=T["texto2"], size=12), customdata=[v / total for v in conteo.values],
        hovertemplate="<b>%{x}</b><br>%{y:,} " + unidad + " · %{customdata:.0%}<extra></extra>"))
    if resaltar in conteo.index:
        fig.add_annotation(x=resaltar, y=conteo[resaltar], yshift=30, text=f"<b>{nota}</b>", showarrow=False,
                           font=dict(color=T["acento"], size=12))
    estilo(fig, alto)
    fig.update_yaxes(visible=False, range=[0, max(conteo.max(), 1) * 1.35])
    fig.update_xaxes(tickfont=dict(color=T["texto2"], size=12), gridcolor="rgba(0,0,0,0)")
    fig.update_layout(margin=dict(t=10, b=8, l=8, r=8), bargap=.25)
    mostrar(fig, clave)


def barras_h(serie: pd.Series, clave, color, unidad="artículos", n=12):
    s = serie.head(n)
    fig = go.Figure(go.Bar(
        x=s.values, y=[corto(k, 32) for k in s.index], orientation="h", marker_color=color,
        text=[f"{v:,}" for v in s.values], textposition="outside", cliponaxis=False,
        textfont=dict(color=T["texto2"], size=11), customdata=list(s.index),
        hovertemplate="<b>%{customdata}</b><br>%{x:,} " + unidad + "<extra></extra>"))
    estilo(fig, max(240, 28 * len(s) + 40))
    fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)", tickfont=dict(size=11, color=T["texto2"]))
    fig.update_xaxes(visible=False, range=[0, s.max() * 1.2])
    fig.update_layout(margin=dict(t=4, b=4, r=8))
    mostrar(fig, clave)


def posicion_rango(precio, minimo, maximo):
    """Bajo / Medio / Alto según el tercio del rango del mercado en que cae el precio."""
    if pd.isna(precio):
        return None
    if maximo <= minimo:
        return "Igual"
    p = (precio - minimo) / (maximo - minimo)
    return "Bajo" if p < 1 / 3 else "Medio" if p <= 2 / 3 else "Alto"


def chip_pos(pos):
    if pos is None:
        return "–"
    clase = {"Bajo": "bajo", "Medio": "medio", "Alto": "alto"}.get(pos, "")
    return f'<span class="chip {clase}">{"Igual al mercado" if pos == "Igual" else f"Rango {pos.lower()}"}</span>'


def grafica_rangos_precio(filas: pd.DataFrame, clave, con_precio=True):
    """Cada producto es una barra del precio más bajo (izquierda) al más alto (derecha) del mercado, dividida en
    tercios bajo/medio/alto. El rombo es la mediana; el círculo, el precio del proveedor."""
    filas = filas.reset_index(drop=True)
    etiquetas = [f"{corto(p, 30)} · {b}" for p, b in zip(filas.PRODUCTO, filas.BASE)]

    def pos(v, mn, mx):
        return 50.0 if mx <= mn else (v - mn) / (mx - mn) * 100

    fig = go.Figure()
    for x0, x1, color in [(0, 33.3, T["ok"]), (33.3, 66.7, T["oro"]), (66.7, 100, T["alerta"])]:
        fig.add_trace(go.Bar(y=etiquetas, x=[x1 - x0] * len(filas), base=[x0] * len(filas), orientation="h",
                             marker_color=alfa(color, .18), marker_line_width=0, hoverinfo="skip", width=.42))
    fig.add_trace(go.Scatter(
        x=[pos(r.mediana, r.minimo, r.maximo) for r in filas.itertuples()], y=etiquetas, mode="markers",
        marker=dict(symbol="diamond", size=11, color=T["texto3"], line=dict(color=T["card"], width=1.5)),
        customdata=[[precio_u(r.mediana, r.BASE), int(r.n_prov)] for r in filas.itertuples()],
        hovertemplate="<b>%{y}</b><br>Mediana del mercado %{customdata[0]}<br>%{customdata[1]} proveedores"
                      "<extra></extra>"))
    if con_precio:
        fig.add_trace(go.Scatter(
            x=[pos(r.precio, r.minimo, r.maximo) for r in filas.itertuples()], y=etiquetas, mode="markers",
            marker=dict(symbol="circle", size=15, color=T["acento"], line=dict(color=T["card"], width=2)),
            customdata=[[precio_u(r.precio, r.BASE), precio_u(r.minimo, r.BASE), precio_u(r.maximo, r.BASE)]
                        for r in filas.itertuples()],
            hovertemplate="<b>%{y}</b><br>Tu precio %{customdata[0]}<br>Rango del mercado %{customdata[1]} a "
                          "%{customdata[2]}<extra></extra>"))
    # Precio más bajo y más alto a los lados de cada barra
    fig.add_trace(go.Scatter(x=[-1.5] * len(filas), y=etiquetas, mode="text", textposition="middle left",
                             text=[f"${v:,.2f}" for v in filas.minimo], hoverinfo="skip",
                             textfont=dict(size=11, color=T["texto3"])))
    fig.add_trace(go.Scatter(x=[101.5] * len(filas), y=etiquetas, mode="text", textposition="middle right",
                             text=[f"${v:,.2f}" for v in filas.maximo], hoverinfo="skip",
                             textfont=dict(size=11, color=T["texto3"])))
    estilo(fig, max(220, 38 * len(filas) + 60))
    fig.update_layout(barmode="overlay", margin=dict(t=30, b=8, l=8, r=8))
    fig.update_xaxes(range=[-22, 122], tickvals=[16.7, 50, 83.3], ticktext=["Bajo", "Medio", "Alto"], side="top",
                     tickfont=dict(color=T["texto2"], size=12), gridcolor="rgba(0,0,0,0)")
    fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)", tickfont=dict(size=11, color=T["texto2"]))
    mostrar(fig, clave)


def leyenda_rangos(con_precio=True):
    items = [f'<span><i style="background:{alfa(T["ok"], .35)}"></i>Rango bajo</span>',
             f'<span><i style="background:{alfa(T["oro"], .35)}"></i>Rango medio</span>',
             f'<span><i style="background:{alfa(T["alerta"], .35)}"></i>Rango alto</span>',
             f'<span><i class="rombo" style="background:{T["texto3"]}"></i>Mediana del mercado</span>']
    if con_precio:
        items.append(f'<span><i class="punto" style="background:{T["acento"]}"></i>Tu precio</span>')
    md('<div class="leyenda">' + "".join(items) + "</div>")


def mapa(valor: dict, textos: dict, binario: bool, clave, alto=430):
    """Estados por gerencia regional. binario: 1 = presente (acento), 0 = ausente (gris); si no, escala de intensidad."""
    gj = geo_estados()
    if not gj:
        st.caption(f"No encontré {MAPA.name}: el mapa no se puede dibujar.")
        return
    locs, z, hover = [], [], []
    for f in gj["features"]:
        e = f["properties"]["name"]
        s = ESTADO_REGION.get(e)
        locs.append(e)
        z.append(valor.get(s, 0) if s else (0 if binario else None))
        hover.append(f"<b>{s} · {REGIONES[s][0]}</b><br>{e}<br>{textos[s]}" if s else f"<b>{e}</b><br>Sin gerencia regional")
    escala = [[0, T["otros"]], [1, T["acento"]]] if binario else [[0, alfa(T["acento"], .12)], [1, T["acento"]]]
    fig = go.Figure(go.Choropleth(
        geojson=gj, featureidkey="properties.name", locations=locs, z=z, colorscale=escala, zmin=0,
        zmax=1 if binario else max(max(valor.values()), 1), showscale=False, marker_line_color=T["card"],
        marker_line_width=1, customdata=hover, hovertemplate="%{customdata}<extra></extra>"))
    centros = centros_region()
    sig = [s for s in REGIONES if s in centros]
    fig.add_trace(go.Scattergeo(lon=[centros[s][0] for s in sig], lat=[centros[s][1] for s in sig], text=sig,
                                mode="text", hoverinfo="skip",
                                textfont=dict(family="Inter, sans-serif", size=11, color=T["texto"])))
    fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)", projection_type="mercator")
    estilo(fig, alto)
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False)
    st.plotly_chart(fig, width="stretch", theme=None, key=clave, config={"displayModeBar": False, "scrollZoom": False})


# ──────────────────────────────────────────────────────────────
# DATOS
# ──────────────────────────────────────────────────────────────
cfg_od = secreto("onedrive")
if cfg_od.get("link"):
    minutos = max(1, int(cfg_od.get("minutos", 10)))
    try:
        contenido, descargado = descargar_excel(cfg_od["link"], int(ahora().timestamp() // (minutos * 60)))
        origen = f"OneDrive · {descargado:%H:%M}"
    except Exception as e:  # noqa: BLE001
        if not ARCHIVO.exists():
            st.error(f"No pude descargar el Excel de OneDrive ({motivo(e)}).")
            st.stop()
        st.warning(f"No pude descargar el Excel de OneDrive ({motivo(e)}). Muestro la copia local.",
                   icon=":material/cloud_off:")
        contenido, origen = ARCHIVO.read_bytes(), "Copia local"
elif ARCHIVO.exists():
    contenido, origen = ARCHIVO.read_bytes(), "Archivo local"
else:
    st.error(f"No hay datos: agrega el vínculo de OneDrive en los secretos (`[onedrive] link`) "
             f"o coloca **{NOMBRES_EXCEL[0]}** junto a app_proveedores.py.")
    st.stop()

try:
    datos = cargar(contenido)
except Exception as e:  # noqa: BLE001
    st.error(f"No pude leer el Excel con la estructura esperada: {e}")
    st.stop()
if datos.empty:
    st.warning("El Excel no tiene artículos capturados.")
    st.stop()

hoy = pd.Timestamp(ahora().date())
unicos = datos.drop_duplicates(["CVE", "ART_ID"]).copy()   # un artículo cuenta una vez por proveedor
unicos["DIAS"] = (unicos["VIG_FECHA"] - hoy).dt.days
unicos["RANGO_VIG"] = en_rangos(unicos["DIAS"], RANGOS_VIG).astype(object).where(unicos["DIAS"].notna(), "Sin fecha")
nombre_prod = datos.groupby("PROD")["PRODUCTO"].first()
prov_por_prod = unicos.groupby("PROD")["CVE"].nunique()

# Precio por kg / L / pieza de cada proveedor en cada producto (mediana de sus artículos)
precios = (unicos.dropna(subset=["PRECIO_U", "BASE"]).groupby(["PROD", "BASE", "CVE"])
           .agg(precio=("PRECIO_U", "median"), arts=("ART_ID", "nunique")).reset_index())
grp = precios.groupby(["PROD", "BASE"])
precios["n_prov"] = grp["CVE"].transform("nunique")
precios = precios[precios["n_prov"] >= 2].copy()     # el rango solo tiene sentido con dos o más proveedores
grp = precios.groupby(["PROD", "BASE"])
for c, f in [("mediana", "median"), ("minimo", "min"), ("maximo", "max")]:
    precios[c] = grp["precio"].transform(f)
precios["PRODUCTO"] = precios["PROD"].map(nombre_prod)
precios["posicion"] = [posicion_rango(p, a, b) for p, a, b in zip(precios.precio, precios.minimo, precios.maximo)]

# Indicadores por proveedor
m = unicos.groupby("CVE").agg(nombre=("PROV_CORTO", "first"), razon=("PROVEEDOR", "first"),
                              art=("ART_ID", "nunique"), prod=("PROD", "nunique"), marcas=("MARCA", "nunique"),
                              hojas=("NEGOCIACION", "nunique"))
m["regiones"] = unicos.groupby("CVE")["REGIONES"].agg(lambda s: len({r for t in s for r in t}))
m["nacional"] = unicos.groupby("CVE")["ALCANCE"].agg(lambda s: (s == "Nacional").any())
ofertas = unicos[(unicos["INVENTARIO"] == "Oferta") & ~(unicos["VIG_FECHA"] < hoy)]
m["ofertas"] = ofertas.groupby("CVE")["ART_ID"].nunique().reindex(m.index, fill_value=0)
m["rango_art"] = en_rangos(m["art"], RANGOS_ART).astype(str)
m["rango_reg"] = en_rangos(m["regiones"], RANGOS_REG).astype(str)
N = len(m)
exp_reg = unicos[["CVE", "ART_ID", "REGIONES"]].explode("REGIONES").dropna()
reg = exp_reg.groupby("REGIONES").agg(prov=("CVE", "nunique"), art=("ART_ID", "nunique")) \
    .reindex(list(REGIONES), fill_value=0)

# ──────────────────────────────────────────────────────────────
# ACCESO CON CÓDIGO
# ──────────────────────────────────────────────────────────────
def normaliza_codigo(c) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(c).upper())


def leer_accesos() -> dict:
    """{código normalizado: CVE o ADMIN} desde [acceso] de los secretos."""
    cfg = secreto("acceso")
    salida = {}
    admin = cfg.get("admin", [])
    for c in [admin] if isinstance(admin, str) else list(admin):
        if normaliza_codigo(c):
            salida[normaliza_codigo(c)] = ADMIN
    for cve, c in dict(cfg.get("proveedores", {})).items():
        if normaliza_codigo(c):
            salida[normaliza_codigo(c)] = str(cve).strip()
    return salida


ACCESOS = leer_accesos()


def entrar(codigo: str) -> bool:
    buscado = normaliza_codigo(codigo)
    for c, usuario in ACCESOS.items():
        if hmac.compare_digest(c, buscado):   # comparación de tiempo constante
            st.session_state["usuario"] = usuario
            st.session_state["intentos"] = 0
            return True
    return False


def salir():
    for k in ["usuario", "ver", "buscar_art_q"]:
        st.session_state.pop(k, None)


if not ACCESOS:
    st.session_state.setdefault("usuario", ADMIN)   # sin códigos configurados: modo administrador (pruebas)
usuario = st.session_state.get("usuario")

# ──────────────────────────────────────────────────────────────
# ENCABEZADO
# ──────────────────────────────────────────────────────────────
with st.container(key="hero", horizontal=True, vertical_alignment="top", gap="medium"):
    md(f'<div class="hero"><div class="marca">{LOGO}<span class="sep"></span>'
       f'<div class="gob">Alimentación para el Bienestar</div></div>'
       f'<h1>Tablero de Proveedores</h1>'
       f'<div class="gerencia">Programa de Abasto Rural 2026 · Dirección Comercial</div>'
       f'<div class="subgerencia">Subgerencia de Abarrotes</div>'
       f'<div class="fecha">Información al {fecha_larga(hoy)} · {origen}</div></div>')
    with st.container(key="controles", width="content", gap="small"):
        st.toggle(":material/dark_mode: Modo oscuro", key="oscuro")
        if usuario and ACCESOS:
            st.button("Salir", icon=":material/logout:", key="salir", on_click=salir)
        if usuario == ADMIN and cfg_od.get("link"):
            st.button("Actualizar datos", icon=":material/refresh:", key="actualizar", on_click=descargar_excel.clear)

if not usuario:
    bloqueo = st.session_state.get("bloqueo_hasta", 0) - time.time()
    with st.container(key="login"):
        md(f'<div class="login-cab"><div class="ico">{svg("candado", 30)}</div><h2>Acceso para proveedores</h2>'
           f'<p>Escribe el código que te entregó la Subgerencia de Abarrotes para ver la información de tu empresa.</p></div>')
        with st.form("f_login", border=False):
            codigo = st.text_input("Código de acceso", type="password", placeholder="XXXX-XXXX",
                                   disabled=bloqueo > 0)
            enviado = st.form_submit_button("Entrar", icon=":material/login:", width="stretch", disabled=bloqueo > 0)
        if bloqueo > 0:
            md(f'<div class="aviso mal">{svg("alerta", 16)}Demasiados intentos. Espera {int(bloqueo // 60) + 1} '
               f'min y vuelve a intentar.</div>')
        elif enviado:
            if entrar(codigo):
                st.rerun()
            st.session_state["intentos"] = st.session_state.get("intentos", 0) + 1
            if st.session_state["intentos"] >= MAX_INTENTOS:
                st.session_state["bloqueo_hasta"] = time.time() + ESPERA_SEG
                st.session_state["intentos"] = 0
            md(f'<div class="aviso mal">{svg("alerta", 16)}El código no es válido. Revísalo e intenta de nuevo.</div>')
    st.stop()

if usuario != ADMIN and usuario not in m.index:
    st.warning("Tu código es válido, pero tu clave de proveedor no aparece en la información vigente. "
               "Comunícate con la Subgerencia de Abarrotes.", icon=":material/info:")
    st.stop()

# El administrador puede ver el programa completo o abrir el tablero de un proveedor
if usuario == ADMIN:
    with st.container(border=True, key="card_ver"):
        c1, c2 = st.columns([2, 3], vertical_alignment="center")
        md('<div class="card-h" style="border:none;margin:0;padding:12px 0"><div><h2>Vista de administrador</h2>'
           '<p>Elige un proveedor para ver su tablero tal como él lo ve, o deja «Todo el programa».</p></div></div>', c1)
        opciones = [""] + sorted(m.index, key=lambda c: m.at[c, "nombre"])
        yo = c2.selectbox("Proveedor", opciones, key="ver", label_visibility="collapsed",
                          format_func=lambda c: "Todo el programa" if not c else f"{m.at[c, 'nombre']} · CVE {c}") or None
else:
    yo = usuario

_secciones = []


def seccion(titulo, sub, ancla):
    _secciones.append((ancla, titulo))
    md(f'<div class="seccion" id="{ancla}"><span class="paso">{len(_secciones)}</span><div><h2>{titulo}</h2>'
       f'<p>{sub}</p></div></div>')


def indice(items):
    md('<nav class="indice">' + "".join(f'<a href="#{k}"><span>{i}</span>{v}</a>'
                                        for i, (k, v) in enumerate(items, start=1)) + "</nav>")


# ──────────────────────────────────────────────────────────────
# VISTA DE UN PROVEEDOR
# ──────────────────────────────────────────────────────────────
def vista_proveedor(yo):
    r = m.loc[yo]
    mio = unicos[unicos.CVE == yo]
    mios_p = precios[precios.CVE == yo].sort_values("PRODUCTO")
    prods = mio.groupby("PRODUCTO")["ART_ID"].nunique().sort_values(ascending=False)
    regs_mias = {x for t in mio.REGIONES for x in t}
    otros_rango = int((m.rango_art == r.rango_art).sum()) - 1

    md(f'<div class="m-card ficha"><div class="ico">{svg("proveedor", 28)}</div><div><div class="t">Tablero de</div>'
       f'<b>{escape(r["nombre"])}</b><span>{escape(r.razon)} · CVE {escape(yo)}</span></div><div class="etq">'
       f'<span class="etq-c">{svg("mapa", 14)}{"Cobertura nacional" if r.nacional else "Cobertura regional"}</span>'
       f'<span class="etq-c">{svg("hoja", 14)}{r.hojas:,} {"hoja" if r.hojas == 1 else "hojas"} de negociación</span>'
       f'</div></div>')
    indice([("resumen", "Resumen"), ("surtido", "Tu surtido"), ("precios", "Tu precio en el mercado"),
            ("programa", "Tu lugar en el programa"), ("regiones", "Cobertura"), ("vigencia", "Vigencias"),
            ("articulos", "Tus artículos")])

    # 1. Resumen
    seccion("Resumen", "Lo principal de tu participación en el programa.", "resumen")
    med = m.median(numeric_only=True)
    md('<div class="kpis">' + "".join([
        kpi_html("Artículos", f"{r.art:,}", f"Mediana del programa: {med.art:.0f}", "articulo", P626),
        kpi_html("Productos", f"{r['prod']:,}", f"Mediana del programa: {med['prod']:.0f}", "producto", P7420),
        kpi_html("Marcas", f"{r.marcas:,}", f"Mediana del programa: {med.marcas:.0f}", "marca", P1255),
        kpi_html("Regiones", f"{r.regiones} de {len(REGIONES)}",
                 "Llegas a todo el país" if r.nacional else "Gerencias regionales que cubres", "mapa", P7421),
        kpi_html("Ofertas vigentes", f"{r.ofertas:,}", "Artículos en oferta con vigencia activa", "estrella", P627),
        kpi_html("Productos con más proveedores", f"{len(mios_p):,}",
                 f"de tus {r['prod']:,} productos tienen rango de precio", "balanza", P1255_60),
    ]) + "</div>")

    hall = [(P626, "articulo", f"Participas con <b>{r.art:,} artículos</b> de <b>{r['prod']:,} productos</b> y "
                               f"<b>{r.marcas:,} marcas</b>."),
            (P7420, "producto", f"Tu producto con más artículos es <b>{escape(prods.index[0])}</b> "
                                f"({prods.iloc[0]:,})."),
            (P1255, "rangos", f"Te ubicas en el rango de <b>{escape(r.rango_art)} artículos</b>"
                              + (f", junto con otros <b>{otros_rango}</b> proveedores." if otros_rango > 0 else "."))]
    if len(mios_p):
        bajos = int((mios_p.posicion == "Bajo").sum())
        hall.append((P7421, "balanza", f"En <b>{bajos} de {len(mios_p)}</b> productos que surten varios proveedores, "
                                       f"tu precio está en el <b>rango bajo</b> del mercado."))
    else:
        hall.append((P7421, "balanza", "Tus productos no tienen otro proveedor en el programa: no hay rango de "
                                       "precio para compararlos."))
    hall.append((P627, "mapa", f"Tienes <b>cobertura nacional</b>: llegas a las {len(REGIONES)} regiones." if r.nacional
                 else f"Llegas a <b>{len(regs_mias)} de {len(REGIONES)}</b> gerencias regionales: "
                      f"<b>{escape(', '.join(REGIONES[s][0] for s in REGIONES if s in regs_mias))}</b>."))
    pronto = mio[mio.DIAS.between(0, 30)]
    if len(pronto):
        hall.append((P7420, "calendario", f"<b>{len(pronto):,}</b> de tus artículos vencen en los próximos 30 días; el "
                                          f"primero el <b>{pronto.VIG_FECHA.min():%d/%m/%Y}</b>."))
    else:
        prox = mio.loc[mio.DIAS >= 0, "VIG_FECHA"].min()
        hall.append((P1255_60, "calendario", "Ninguno de tus artículos vence en los próximos 30 días"
                     + (f"; la vigencia más próxima es el <b>{prox:%d/%m/%Y}</b>." if pd.notna(prox) else ".")))
    md('<div style="height:4px"></div>')
    lectura_rapida(hall)

    # 2. Surtido
    seccion("Tu surtido", "Qué productos ofreces y en qué rangos de costo unitario están tus artículos.", "surtido")
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Artículos por producto", "s_prod", "Tus productos con más artículos.",
                     extra=f"{len(prods):,} productos"):
            barras_h(prods, "g_s_prod", T["verde"])
    with b:
        with tarjeta("Tus artículos por rango de costo unitario", "s_costo",
                     "Costo de adquisición por pieza, sin impuestos.", extra=f"{r.art:,} artículos"):
            c = conteo_rangos(mio.COSTO_UNIT, RANGOS_COSTO)
            barras_rangos(c, "g_s_costo", "artículos", resaltar=c.idxmax(), nota="La mayoría")
            inv = mio.INVENTARIO.value_counts()
            md('<div class="leyenda">' + "".join(
                f'<span><i style="background:{T["acento"] if k == "Oferta" else T["verde"]}"></i>{escape(k)}: '
                f'<b>{v:,}</b></span>' for k, v in inv.items()) + "</div>")

    # 3. Precio dentro del rango del mercado
    seccion("Tu precio dentro del rango del mercado",
            "Para los productos que surten dos o más proveedores: el rango va del precio más bajo al más alto por kilo, "
            "litro o pieza. No se muestra el nombre de ningún otro proveedor.", "precios")
    if mios_p.empty:
        with tarjeta("Rangos de precio", "p_vacio"):
            md(f'<div class="lb-mas">{svg("ok", 16)}Tus productos no tienen otro proveedor en el programa.</div>')
    else:
        a, b = st.columns([3, 2], gap="medium")
        with a:
            with tarjeta("Dónde cae tu precio", "p_graf", "Pasa el cursor para ver los precios.",
                         extra=f"{len(mios_p):,} productos"):
                leyenda_rangos()
                grafica_rangos_precio(mios_p, "g_p_rangos")
        with b:
            with tarjeta("Detalle por producto", "p_tabla", "Tu precio mediano frente al rango del mercado."):
                filas = [[escape(x.PRODUCTO) + f"<small>{int(x.n_prov)} proveedores</small>",
                          f"{precio_u(x.minimo, x.BASE)}<small>a {precio_u(x.maximo, x.BASE)}</small>",
                          precio_u(x.precio, x.BASE), chip_pos(x.posicion)] for x in mios_p.itertuples()]
                tabla(["Producto", "Rango del mercado", "Tu precio", "Posición"], filas, numericas=(1, 2), alto=460)
            pos = mios_p.posicion.value_counts()
            guia(" · ".join(f"<b>{pos.get(k, 0)}</b> en rango {k.lower()}" for k in ["Bajo", "Medio", "Alto"])
                 + (f" · <b>{pos['Igual']}</b> igual al mercado" if "Igual" in pos else "")
                 + ". El rango se divide en tres partes iguales.")

    # 4. Lugar en el programa (por rangos, sin nombres)
    seccion("Tu lugar en el programa", f"Cuántos de los {N} proveedores hay en cada rango; en color, el tuyo.",
            "programa")
    a, b = st.columns(2, gap="medium")
    with a:
        with tarjeta("Proveedores por rango de artículos", "r_art", "Artículos distintos que ofrece cada proveedor.",
                     extra=f"Tu rango: {r.rango_art}"):
            barras_rangos(conteo_rangos(m.art, RANGOS_ART), "g_r_art", "proveedores", resaltar=r.rango_art)
    with b:
        with tarjeta("Proveedores por rango de cobertura", "r_reg", "Gerencias regionales a las que llega cada uno.",
                     extra=f"Tu rango: {r.rango_reg}"):
            c = conteo_rangos(m.regiones, RANGOS_REG)
            barras_rangos(c[(c > 0) | (c.index == r.rango_reg)], "g_r_reg", "proveedores", resaltar=r.rango_reg)

    # 5. Cobertura
    seccion("Cobertura regional", "Las gerencias regionales a las que llegan tus artículos.", "regiones")
    art_r = exp_reg[exp_reg.CVE == yo].groupby("REGIONES")["ART_ID"].nunique()
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Mapa de tu cobertura", "c_mapa", "En color, las regiones donde estás.",
                     extra=f"{len(regs_mias)} de {len(REGIONES)} regiones"):
            mapa({s: int(s in regs_mias) for s in REGIONES},
                 {s: (f"<b>{int(art_r.get(s, 0))} de tus artículos</b>" if s in regs_mias else "<b>No llegas aquí</b>")
                  + f"<br>{int(reg.at[s, 'prov'])} proveedores en la región" for s in REGIONES}, True, "g_c_mapa")
    with b:
        with tarjeta("Tus artículos por región", "c_tabla", "Los nacionales cuentan en todas las regiones."):
            filas = [[f"<b>{s}</b> · {escape(REGIONES[s][0])}", f"{int(art_r.get(s, 0)):,}" if s in regs_mias else
                      '<span class="chip alto">Sin cobertura</span>', f"{int(reg.at[s, 'prov'])}"] for s in REGIONES]
            tabla(["Región", "Tus artículos", "Proveedores en la región"], filas, numericas=(1, 2))

    # 6. Vigencias
    seccion("Vigencias", "Cuánto falta para que venza la vigencia de tus artículos.", "vigencia")
    seccion_vigencias(mio, propio=True)

    # 7. Artículos
    seccion("Tus artículos", "Todos tus artículos vigentes; escribe para filtrar.", "articulos")
    with tarjeta("Listado de artículos", "a_tabla", extra=f"{r.art:,} artículos"):
        with st.container(key="buscar_art"):
            md(f'<span class="lupa">{svg("buscar", 16)}</span>')
            q = st.text_input("Buscar", key="buscar_art_q", placeholder="Buscar artículo, producto, marca o clave…",
                              label_visibility="collapsed")
        d = mio.sort_values(["PRODUCTO", "ARTICULO"])
        if q:
            palabras = norm_txt(q).split()
            texto = (d.ARTICULO + " " + d.PRODUCTO + " " + d.MARCA + " " + d.ART_ID.astype(str) + " " + d.INVENTARIO) \
                .map(norm_txt)
            d = d[[all(p in t for p in palabras) for t in texto]]
        filas = [[f"{escape(x.ARTICULO)}<small>ART. {escape(x.ART_ID)}</small>", escape(x.PRODUCTO), escape(x.MARCA),
                  escape(x.PRESENTACION), escape(x.INVENTARIO),
                  "Nacional" if x.ALCANCE == "Nacional" else escape(" ".join(x.REGIONES)) or "–",
                  dinero(x.COSTO_UNIT), "–" if pd.isna(x.VIG_FECHA) else f"{x.VIG_FECHA:%d/%m/%Y}"]
                 for x in d.itertuples()]
        if filas:
            tabla(["Artículo", "Producto", "Marca", "Presentación", "Inventario", "Cobertura", "Costo unitario",
                   "Vigencia"], filas, numericas=(6, 7), alto=480)
        else:
            md('<div class="lb-mas">Sin coincidencias.</div>')


def seccion_vigencias(d, propio):
    a, b = st.columns([2, 3], gap="medium")
    with a:
        with tarjeta("Artículos por rango de vigencia", f"v_rangos_{propio}", "Días que faltan para el vencimiento.",
                     extra=f"{len(d):,} artículos"):
            orden = [x[1] for x in RANGOS_VIG] + ["Sin fecha"]
            c = d.RANGO_VIG.value_counts().reindex(orden, fill_value=0)
            c = c[(c > 0) | c.index.isin(["0 a 30 días"])]
            barras_rangos(c, f"g_v_{propio}", "artículos", colores=[COLOR_VIG[k] for k in c.index])
    with b:
        prox = d[d.DIAS.between(0, 60)].sort_values(["VIG_FECHA", "ARTICULO"])
        with tarjeta("Próximos vencimientos", f"v_lista_{propio}", "Artículos que vencen en los próximos 60 días.",
                     extra=f"{len(prox):,} artículos"):
            if prox.empty:
                md(f'<div class="lb-mas">{svg("ok", 16)}Sin vencimientos en los próximos 60 días.</div>')
            else:
                def dias_chip(x):
                    clase = "alto" if x <= 15 else "medio" if x <= 30 else "bajo"
                    return f'<span class="chip {clase}">en {int(x)} días</span>'
                filas = [[f"{x.VIG_FECHA:%d/%m/%Y}", dias_chip(x.DIAS), escape(x.ARTICULO)]
                         + ([] if propio else [escape(x.PROV_CORTO)]) for x in prox.itertuples()]
                tabla(["Vence", "Faltan", "Artículo"] + ([] if propio else ["Proveedor"]), filas, alto=330)


# ──────────────────────────────────────────────────────────────
# VISTA DEL PROGRAMA COMPLETO (administrador)
# ──────────────────────────────────────────────────────────────
def vista_programa():
    indice([("resumen", "Resumen"), ("proveedores", "Proveedores por rangos"), ("precios", "Rangos de precio"),
            ("regiones", "Cobertura"), ("vigencia", "Vigencias")])
    seccion("Resumen", "El programa de un vistazo.", "resumen")
    n_comp = int((prov_por_prod >= 2).sum())
    md('<div class="kpis">' + "".join([
        kpi_html("Proveedores", f"{N:,}", f"{int(m.nacional.sum())} con cobertura nacional", "proveedor", P626),
        kpi_html("Productos", f"{unicos.PROD.nunique():,}", f"{n_comp} con dos o más proveedores", "producto", P7420),
        kpi_html("Artículos", f"{unicos.ART_ID.nunique():,}", f"Mediana por proveedor: {m.art.median():.0f}",
                 "articulo", P1255),
        kpi_html("Marcas", f"{unicos.MARCA.nunique():,}", f"Mediana por proveedor: {m.marcas.median():.0f}", "marca",
                 P7421),
        kpi_html("Ofertas vigentes", f"{ofertas.ART_ID.nunique():,}", f"de {int((m.ofertas > 0).sum())} proveedores",
                 "estrella", P627),
        kpi_html("Hojas de negociación", f"{unicos.NEGOCIACION.nunique():,}", "Negociaciones vigentes", "hoja", P1255_60),
    ]) + "</div>")
    rango_moda = m.rango_art.value_counts()
    top5 = m.art.nlargest(5).sum() / m.art.sum()
    pronto = unicos[unicos.DIAS.between(0, 30)]
    lectura_rapida([
        (P626, "rangos", f"La mayoría de los proveedores está en el rango de <b>{rango_moda.index[0]} artículos</b> "
                         f"(<b>{rango_moda.iloc[0]}</b> de {N})."),
        (P7420, "producto", f"<b>{n_comp / len(prov_por_prod):.0%}</b> de los productos tienen dos o más proveedores."),
        (P1255, "articulo", f"Los 5 proveedores con más artículos reúnen el <b>{top5:.0%}</b> del surtido."),
        (P7421, "mapa", f"<b>{int(m.nacional.sum())}</b> proveedores tienen cobertura nacional y "
                        f"<b>{int((~m.nacional).sum())}</b> son regionales."),
        (P627, "estrella", f"<b>{ofertas.ART_ID.nunique():,}</b> artículos en oferta vigente de "
                           f"<b>{int((m.ofertas > 0).sum())}</b> proveedores."),
        (P1255_60, "calendario", f"<b>{len(pronto):,}</b> artículos vencen en los próximos 30 días."),
    ])

    seccion("Proveedores por rangos", "Cuántos proveedores hay en cada rango de surtido y de cobertura.", "proveedores")
    a, b = st.columns(2, gap="medium")
    with a:
        with tarjeta("Proveedores por rango de artículos", "pa_art", extra=f"{N} proveedores"):
            barras_rangos(conteo_rangos(m.art, RANGOS_ART), "g_pa_art", "proveedores")
    with b:
        with tarjeta("Proveedores por rango de cobertura", "pa_reg", extra=f"{N} proveedores"):
            c = conteo_rangos(m.regiones, RANGOS_REG)
            barras_rangos(c[c > 0], "g_pa_reg", "proveedores")
    st.write("")
    with tarjeta("Proveedores", "pa_tabla", "Abre el tablero de cualquiera con el selector de arriba.",
                 extra=f"{N} proveedores"):
        filas = [[escape(x.nombre) + f"<small>CVE {escape(c)}</small>", f"{x.art:,}", escape(x.rango_art),
                  f"{x['prod']:,}", f"{x.marcas:,}", "Nacional" if x.nacional else f"{x.regiones} de {len(REGIONES)}",
                  f"{x.ofertas:,}"] for c, x in m.sort_values("nombre").iterrows()]
        tabla(["Proveedor", "Artículos", "Rango", "Productos", "Marcas", "Cobertura", "Ofertas"], filas,
              numericas=(1, 3, 4, 6), alto=460)

    seccion("Rangos de precio por producto", "Precio por kilo, litro o pieza de los productos que surten dos o más "
            "proveedores: del más bajo al más alto.", "precios")
    resumen = (precios.groupby(["PROD", "BASE"]).agg(PRODUCTO=("PRODUCTO", "first"), n_prov=("CVE", "nunique"),
                                                     minimo=("minimo", "first"), mediana=("mediana", "first"),
                                                     maximo=("maximo", "first")).reset_index())
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Productos más disputados", "pr_graf", "Los 15 con más proveedores; el rombo es la mediana.",
                     extra=f"{len(resumen):,} productos con rango"):
            leyenda_rangos(con_precio=False)
            top = resumen.sort_values(["n_prov", "PRODUCTO"], ascending=[False, True]).head(15)
            grafica_rangos_precio(top, "g_pr_rangos", con_precio=False)
    with b:
        with tarjeta("Artículos por rango de costo unitario", "pr_costo", "Todo el programa, por pieza sin impuestos.",
                     extra=f"{unicos.ART_ID.nunique():,} artículos"):
            barras_rangos(conteo_rangos(unicos.drop_duplicates("ART_ID").COSTO_UNIT, RANGOS_COSTO), "g_pr_costo",
                          "artículos")
    st.write("")
    with tarjeta("Todos los rangos de precio", "pr_tabla", "Amplitud = cuántas veces el precio más alto es el más bajo."):
        filas = [[escape(x.PRODUCTO), f"{x.n_prov}", precio_u(x.minimo, x.BASE), precio_u(x.mediana, x.BASE),
                  precio_u(x.maximo, x.BASE), f"{x.maximo / x.minimo:.1f}×"]
                 for x in resumen.sort_values("PRODUCTO").itertuples()]
        tabla(["Producto", "Proveedores", "Más bajo", "Mediana", "Más alto", "Amplitud"], filas,
              numericas=(1, 2, 3, 4, 5), alto=420)

    seccion("Cobertura regional", "Proveedores y artículos que llegan a cada gerencia regional.", "regiones")
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Proveedores por región", "pc_mapa", "Más intenso = más proveedores."):
            mapa({s: int(reg.at[s, "prov"]) for s in REGIONES},
                 {s: f"{int(reg.at[s, 'prov'])} proveedores · {int(reg.at[s, 'art']):,} artículos" for s in REGIONES},
                 False, "g_pc_mapa")
    with b:
        with tarjeta("Detalle por región", "pc_tabla", "Los nacionales cuentan en todas."):
            filas = [[f"<b>{s}</b> · {escape(REGIONES[s][0])}", f"{int(x.prov)}", f"{int(x.art):,}"]
                     for s, x in reg.sort_values("prov", ascending=False).iterrows()]
            tabla(["Región", "Proveedores", "Artículos"], filas, numericas=(1, 2))

    seccion("Vigencias", "Cuánto falta para que venzan las vigencias de los artículos del programa.", "vigencia")
    seccion_vigencias(unicos, propio=False)


if yo:
    vista_proveedor(yo)
else:
    vista_programa()

md('<div class="pie">Subgerencia de Abarrotes · Dirección Comercial · Alimentación para el Bienestar<br>'
   'Información de la hoja VIGENTES del Programa de Abasto Rural 2026.</div>')
