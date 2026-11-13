"""
Programa de Abasto Rural 2026 · Tablero de Proveedores
Subgerencia de Abarrotes · Alimentación para el Bienestar

Tablero comparativo pensado para mostrarse a los proveedores: cada uno ve su lugar frente a los demás
(ranking, precio por kilo/litro/pieza contra el mercado, cobertura regional y un cara a cara).

Ejecutar:   streamlit run app_proveedores.py

Origen de los datos (en este orden):
  1. OneDrive / SharePoint: [onedrive] link = "..." en .streamlit/secrets.toml (igual que app_mosaic.py).
  2. "PROGRAMA DE ABASTO RURAL 2026.xlsx" junto a este archivo o en la carpeta ../Dashboard.

Vínculo fijo para un proveedor:   ?vista=<CVE>
  Abre el tablero ya en ese proveedor, sin selector y con los nombres de la competencia ocultos.
  (No es un control de acceso: quien cambie la clave en la URL ve otro proveedor.)

Diseño: el mismo de app_mosaic.py (Mosaic + paleta institucional, modo claro/oscuro).
"""
import base64
import io
import json
import re
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
    page_icon=":material/emoji_events:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CARPETA = Path(__file__).parent
NOMBRES_EXCEL = ["PROGRAMA DE ABASTO RURAL 2026.xlsx", "PROGRAMA_DE_ABASTO_RURAL_2026_.xlsx"]
ARCHIVOS = [d / n for d in (CARPETA, CARPETA.parent / "Dashboard") for n in NOMBRES_EXCEL]
ARCHIVO = next((a for a in ARCHIVOS if a.exists()), ARCHIVOS[0])
MAPA = CARPETA / "mexico_estados.json"

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

# Íconos de línea (Lucide, licencia ISC)
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
    "trofeo": '<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6"/><path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18"/><path d="M4 22h16"/>'
              '<path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/>'
              '<path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2Z"/>',
    "objetivo": '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
    "balanza": '<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 '
               '1s-2.13-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>',
    "duelo": '<path d="M14.5 17.5 3 6V3h3l11.5 11.5"/><path d="m13 19 6-6"/><path d="m16 16 4 4"/><path d="m19 21 2-2"/>'
             '<path d="M14.5 6.5 18 3h3v3l-3.5 3.5"/><path d="m5 14 4 4"/><path d="m7 17-3 3"/><path d="m3 19 2 2"/>',
    "tendencia": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
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
# Unidad de medida -> (factor a la unidad base, unidad base) para comparar precios entre presentaciones
UNIDAD_BASE = {"GR": (0.001, "kg"), "KG": (1, "kg"), "ML": (0.001, "L"), "LT": (1, "L"), "PZ": (1, "pza")}


def nombre_corto(nombre: str) -> str:
    corto = re.sub(r"[,.\s]+(S\.?\s?A\.?|S\.?\s?DE\s?R\.?\s?L\.?|S\.?\s?C\.?|S\.?\s?P\.?\s?R\.?)(?![A-Z]).*$", "",
                   str(nombre).strip(), flags=re.I)
    return corto.strip(" ,.").title() or str(nombre)


def clave_texto(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return s.where(~s.isin(["", "nan", "NaN", "<NA>", "None"]), pd.NA)


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
    df["INVENTARIO"] = inv.map(lambda x: INVENTARIO_LABEL.get(x, x), na_action="ignore").fillna("Sin dato")

    cob = df["COBERTURA"].astype(str).str.upper().str.replace(r"[^A-Z]+", " ", regex=True).str.strip()
    fichas = cob.where(~cob.isin(["", "NAN", "NA"]), "").str.split()
    df["ALCANCE"] = ["Nacional" if "NAL" in f else ("Regional" if f else "Sin dato") for f in fichas]
    df["REGIONES"] = [tuple(REGIONES) if "NAL" in f else tuple(r for r in REGIONES if r in f) for f in fichas]
    df["VIG_FECHA"] = pd.to_datetime(df["VIGENCIA"], errors="coerce")

    # Precio comparable: costo unitario entre el contenido de la pieza en kg, L o piezas
    factor = df["UM"].map(lambda u: UNIDAD_BASE.get(u, (None, None))[0])
    df["BASE"] = df["UM"].map(lambda u: UNIDAD_BASE.get(u, (None, None))[1])
    cantidad = df["GRAM"] * pd.to_numeric(factor, errors="coerce")
    df["PRECIO_U"] = (df["COSTO_UNIT"] / cantidad).where(cantidad > 0)
    return df.reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# DATOS (OneDrive o archivo local)
# ──────────────────────────────────────────────────────────────
def codigo_share(link: str) -> str:
    return "u!" + base64.urlsafe_b64encode(link.strip().encode()).decode().rstrip("=")


def url_descarga(link: str) -> str:
    link = link.strip()
    if "sharepoint.com" in link:
        partes = urlsplit(link)
        query = [(k, v) for k, v in parse_qsl(partes.query) if k != "download"] + [("download", "1")]
        return urlunsplit(partes._replace(query=urlencode(query)))
    return f"https://api.onedrive.com/v1.0/shares/{codigo_share(link)}/root/content"


def config_onedrive() -> dict:
    try:
        return dict(st.secrets.get("onedrive", {}))
    except Exception:  # noqa: BLE001
        return {}


@st.cache_data(max_entries=2, show_spinner="Descargando el Excel de OneDrive…")
def descargar_excel(link: str, bloque: int):
    cfg = config_onedrive()
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

/* Controles de Streamlit siguiendo el tema */
[data-testid="stCheckbox"] label > div:not([data-testid]) { background: var(--borde) !important; }
[data-testid="stCheckbox"][data-selected="true"] label > div:not([data-testid]) { background: var(--acento) !important; }
[data-testid="stCheckbox"] label > div:not([data-testid]) > div { background: #fff !important; }
[data-testid="stCheckbox"] p, [data-testid="stCheckbox"] span[role="img"] { color: var(--texto2) !important; }
[data-testid="stSelectbox"] div[role="group"][data-rac] {
  background: var(--sutil) !important; border: 1px solid var(--borde) !important; border-radius: 8px !important; }
[data-testid="stSelectbox"] div[role="group"][data-focus-within="true"] { border-color: var(--acento) !important; }
[data-testid="stSelectbox"] input { color: var(--texto) !important; background: transparent !important; }
[data-testid="stSelectbox"] button { color: var(--texto3) !important; background: transparent !important; }
[data-testid$="Dropdown"] { background: var(--card) !important; border: 1px solid var(--borde) !important;
  border-radius: 8px !important; box-shadow: 0 8px 24px rgb(0 0 0 / .18) !important; }
[data-testid$="Dropdown"] [role="listbox"] { background: var(--card) !important; }
[data-testid$="Dropdown"] [role="option"] { color: var(--texto2) !important; background: transparent !important; }
[data-testid$="Dropdown"] [role="option"][data-focused="true"],
[data-testid$="Dropdown"] [role="option"]:hover { background: var(--sutil) !important; color: var(--texto) !important; }
[data-testid$="Dropdown"] [role="option"][aria-selected="true"] { color: var(--acento) !important; font-weight: 600; }
/* Botones segmentados (métrica del ranking) */
[data-testid="stButtonGroup"] button { background: var(--card) !important; border-color: var(--borde) !important;
  color: var(--texto2) !important; }
[data-testid="stButtonGroup"] button[aria-checked="true"], [data-testid="stButtonGroup"] button[aria-pressed="true"],
[data-testid="stButtonGroup"] button[data-selected="true"] {
  background: rgb(var(--acento-rgb) / .12) !important; border-color: var(--acento) !important; color: var(--acento) !important; }
[data-testid="stButtonGroup"] button * { color: inherit !important; }

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
.st-key-actualizar button { background: transparent !important; border: 1px solid rgb(255 255 255 / .35) !important;
  color: #fff !important; border-radius: 999px !important; font-size: .8rem; }
.st-key-actualizar button * { color: inherit !important; }

/* Selector "¿Quién eres?" */
.st-key-quien { background: var(--card); border-radius: 12px; box-shadow: var(--sombra); padding: 16px 20px !important;
  border-left: 5px solid var(--acento); margin-bottom: 8px; }
.quien h3 { font-size: 1rem !important; margin: 0 !important; padding: 0 !important; }
.quien p { font-size: .82rem; color: var(--texto3); margin: 2px 0 0 0; }
.vista-fija { display: flex; align-items: center; gap: 14px; }
.vista-fija .ico { width: 48px; height: 48px; border-radius: 12px; display: grid; place-items: center;
  background: rgb(var(--acento-rgb) / .12); color: var(--acento); flex-shrink: 0; }
.vista-fija b { font-size: 1.4rem; color: var(--texto); display: block; line-height: 1.2; }
.vista-fija span { font-size: .82rem; color: var(--texto3); }

/* Secciones e índice */
.seccion { display: flex; align-items: flex-start; gap: 12px; margin: 40px 0 16px 0; scroll-margin-top: 70px; }
.seccion h2 { font-size: 1.3rem !important; font-weight: 700 !important; margin: 0; padding: 0 !important; }
.seccion p { margin: 0; font-size: .84rem; color: var(--texto3); }
.indice { display: flex; flex-wrap: wrap; gap: 8px; margin: 16px 0 8px 0; }
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
.kpis { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 20px; margin-bottom: 20px; }
.kpis.cinco { grid-template-columns: repeat(5, minmax(0, 1fr)); }
.kpi { padding: 18px; display: flex; gap: 14px; align-items: flex-start; }
.kpi > div:last-child { min-width: 0; flex: 1; }
.kpi .ico { width: 44px; height: 44px; border-radius: 10px; flex-shrink: 0; display: grid; place-items: center;
  background: var(--c, rgb(var(--acento-rgb) / .1)); color: var(--ci, var(--acento)); }
.kpi .l { font-size: .72rem; font-weight: 600; color: var(--tenue); text-transform: uppercase; letter-spacing: .04em; }
.kpi .row { display: flex; align-items: flex-start; flex-wrap: wrap; gap: 4px 8px; margin-top: 2px; }
.kpi .v { font-size: 1.9rem; font-weight: 700; color: var(--texto); line-height: 1.15; }
.kpi .s { font-size: .78rem; color: var(--texto3); margin-top: 4px; }
.badge { font-size: .75rem; font-weight: 700; border-radius: 999px; padding: 2px 9px; white-space: nowrap; margin-top: 8px; }
.badge.oro { color: #7a5a12; background: #f6e7bf; }
.badge.ok { color: var(--ok); background: rgb(62 201 114 / .16); }
.badge.medio { color: var(--azul); background: rgb(46 107 158 / .14); }
.badge.bajo { color: var(--texto3); background: var(--sutil); border: 1px solid var(--borde); }
/* Barra de percentil bajo cada KPI */
.pbar { height: 6px; border-radius: 999px; background: var(--sutil); margin-top: 8px; position: relative; overflow: hidden; }
.pbar i { position: absolute; inset: 0 auto 0 0; border-radius: 999px; background: var(--acento); }

/* Lectura rápida */
.hallazgos { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; }
.hz { display: flex; gap: 12px; align-items: flex-start; padding: 16px; font-size: .88rem; color: var(--texto2);
  border: 1.5px solid var(--c) !important; border-left-width: 5px !important; }
.hz .ico { width: 36px; height: 36px; border-radius: 50%; flex-shrink: 0; display: grid; place-items: center;
  color: var(--ci, #fff); background: var(--c); }
.hz b { color: var(--texto); font-weight: 700; }

/* Guía */
.guia { display: flex; gap: 10px; align-items: flex-start; background: rgb(188 149 92 / .12);
  border-radius: 8px; padding: 10px 14px; font-size: .82rem; color: var(--texto2); margin: 0 0 6px 0; }
.guia svg { color: var(--oro); flex-shrink: 0; margin-top: 1px; }
.guia b { color: var(--texto); }

/* Podio */
.podio { display: grid; grid-template-columns: 1fr 1.15fr 1fr; gap: 14px; align-items: end; margin: 6px 0 4px 0; }
.podio .p { text-align: center; border-radius: 12px; padding: 14px 10px; background: var(--sutil);
  border: 1px solid var(--borde); position: relative; }
.podio .p.yo { border: 2px solid var(--acento); background: rgb(var(--acento-rgb) / .07); }
.podio .medalla { font-size: 2rem; line-height: 1; }
.podio .n { font-weight: 700; color: var(--texto); font-size: .92rem; margin-top: 6px; overflow: hidden;
  text-overflow: ellipsis; white-space: nowrap; }
.podio .v { font-size: .8rem; color: var(--texto3); margin-top: 2px; }
.podio .base { height: 10px; border-radius: 0 0 8px 8px; margin: 10px -10px -14px -10px; }
.podio .p1 { padding-top: 26px; } .podio .p1 .base { background: #D4A72C; height: 18px; }
.podio .p2 .base { background: #A9B1BA; } .podio .p3 .base { background: #B87333; }

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
.tabla tr.yo td { background: rgb(var(--acento-rgb) / .09); color: var(--texto); font-weight: 700; }
.tabla tr.yo td:first-child { box-shadow: inset 4px 0 0 var(--acento); }
.lugar { display: inline-grid; place-items: center; min-width: 30px; height: 24px; border-radius: 999px; padding: 0 6px;
  font-size: .76rem; font-weight: 700; background: var(--sutil); color: var(--texto2); border: 1px solid var(--borde); }
.lugar.l1 { background: #F6E3A6; color: #6b4f00; border-color: #D4A72C; }
.lugar.l2 { background: #E5E8EC; color: #3f4650; border-color: #A9B1BA; }
.lugar.l3 { background: #F1D3B8; color: #6b3a10; border-color: #B87333; }
.chip { display: inline-flex; align-items: center; gap: 5px; font-size: .74rem; font-weight: 600; border-radius: 999px;
  padding: 1px 9px; white-space: nowrap; border: 1px solid var(--borde); color: var(--texto2); }
.chip.bien { color: var(--ok); border-color: rgb(62 201 114 / .5); background: rgb(62 201 114 / .1); }
.chip.medio { color: var(--oro); border-color: rgb(188 149 92 / .55); background: rgb(188 149 92 / .1); }
.chip.mal { color: var(--alerta); border-color: rgb(230 57 57 / .45); background: rgb(230 57 57 / .08); }
.chip i { width: 7px; height: 7px; border-radius: 50%; }

/* Cara a cara */
.duelo-cab { display: grid; grid-template-columns: 1fr auto 1fr; gap: 12px; align-items: center; margin: 4px 0 10px 0; }
.duelo-cab .lado { font-weight: 700; color: var(--texto); font-size: 1rem; overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; }
.duelo-cab .lado.izq { text-align: right; color: var(--acento); }
.duelo-cab .vs { font-size: .75rem; font-weight: 800; color: #fff; background: var(--texto3); border-radius: 999px;
  padding: 4px 10px; }
.duelo-cab .marcador { grid-column: 1 / -1; text-align: center; font-size: .84rem; color: var(--texto3); }
.duelo-cab .marcador b { color: var(--texto); font-size: 1.05rem; }
.duelo { display: grid; grid-template-columns: 1fr 150px 1fr; gap: 6px 12px; align-items: center; }
.duelo .m { text-align: center; font-size: .76rem; font-weight: 600; color: var(--texto3); text-transform: uppercase;
  letter-spacing: .03em; line-height: 1.2; }
.duelo .b { display: flex; align-items: center; gap: 8px; min-height: 30px; }
.duelo .b.izq { flex-direction: row-reverse; }
.duelo .b .barra { height: 14px; border-radius: 4px; background: var(--otros); min-width: 3px; }
.duelo .b.izq .barra { background: rgb(var(--acento-rgb) / .45); }
.duelo .b.gana .barra { background: var(--acento); }
.duelo .b.der.gana .barra { background: var(--verde); }
.duelo .b span { font-size: .84rem; font-weight: 600; color: var(--texto2); white-space: nowrap; font-variant-numeric: tabular-nums; }
.duelo .b.gana span { color: var(--texto); font-weight: 800; }

.lb-mas { font-size: .82rem; color: var(--tenue); padding: 10px 0; }
.pie { text-align: center; margin-top: 40px; font-size: .75rem; color: var(--tenue); }

@media (max-width: 1100px) {
  .kpis, .kpis.cinco { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
  .hallazgos { grid-template-columns: repeat(2, 1fr); }
}
@media (max-width: 900px) {
  .hero h1 { font-size: 1.9rem !important; }
  .hero .marca .gob { font-size: 1.8rem; }
}
@media (max-width: 640px) {
  .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
  .st-key-hero { padding: 12px 0 20px 0 !important; flex-wrap: wrap !important; row-gap: 6px; }
  .hero h1 { font-size: 1.6rem !important; }
  .hero .marca .gob { font-size: 1.15rem; white-space: normal; }
  .hero .gerencia { font-size: 1.05rem; }
  .kpis, .kpis.cinco, .hallazgos { grid-template-columns: minmax(0, 1fr); }
  .kpi .v { font-size: 1.5rem; }
  div[class*="st-key-card_"] { padding: 0 14px 12px 14px !important; }
  .card-h { margin: 0 -14px 8px -14px; padding: 14px; }
  .duelo { grid-template-columns: 1fr 90px 1fr; }
  .podio .n { white-space: normal; }
}
"""
md(f"<style>{VARIABLES}{CSS}</style>")


# ──────────────────────────────────────────────────────────────
# COMPONENTES
# ──────────────────────────────────────────────────────────────
def fecha_larga(d) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def alfa(color, a):
    c = color.lstrip("#")
    return f"rgba({int(c[0:2], 16)},{int(c[2:4], 16)},{int(c[4:6], 16)},{a})"


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


def lugar(n):
    if pd.isna(n):
        return '<span class="lugar">–</span>'
    n = int(n)
    return f'<span class="lugar{f" l{n}" if n <= 3 else ""}">#{n}</span>'


def tabla(encabezados, filas, numericas=(), alto=None, clases=None):
    th = "".join(f'<th class="{"num" if i in numericas else ""}">{h}</th>' for i, h in enumerate(encabezados))
    tr = "".join(
        f'<tr class="{c or ""}">' + "".join(
            f'<td class="{"num" if i in numericas else ""} {"fuerte" if i == 1 else ""}">{v}</td>'
            for i, v in enumerate(fila)) + "</tr>"
        for fila, c in zip(filas, clases or [""] * len(filas)))
    marco = f' scroll" style="max-height:{alto}px' if alto else ""
    md(f'<div class="tabla-wrap{marco}"><table class="tabla"><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>')


def estilo(fig, alto=380, leyenda=False):
    fig.update_layout(
        height=alto, margin=dict(l=8, r=24, t=20, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=T["texto3"], size=12),
        showlegend=leyenda,
        legend=dict(orientation="h", y=-0.1, x=0.5, xanchor="center", title=None, font=dict(color=T["texto2"])),
        hoverlabel=dict(bgcolor=T["card"], bordercolor=T["borde"], font_size=12, font_family="Inter, sans-serif",
                        font_color=T["texto"]),
        bargap=0.3, barcornerradius=4,
    )
    fig.update_xaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["tenue"]))
    fig.update_yaxes(gridcolor=T["linea"], zeroline=False, title=None, showline=False, automargin=True,
                     tickfont=dict(color=T["texto3"]))
    return fig


def mostrar(fig, clave):
    st.plotly_chart(fig, width="stretch", theme=None, key=clave, config={"displayModeBar": False})


def corto(texto, n=30):
    texto = str(texto)
    return texto if len(texto) <= n else texto[:n - 1].rstrip() + "…"


def precio_u(v, base):
    return "–" if pd.isna(v) else f"${v:,.2f}/{base}"


# ──────────────────────────────────────────────────────────────
# DATOS
# ──────────────────────────────────────────────────────────────
cfg_od = config_onedrive()
if cfg_od.get("link"):
    minutos = max(1, int(cfg_od.get("minutos", 10)))
    try:
        contenido, descargado = descargar_excel(cfg_od["link"], int(ahora().timestamp() // (minutos * 60)))
        origen = f"OneDrive · {descargado:%H:%M}"
    except Exception as e:  # noqa: BLE001
        if not ARCHIVO.exists():
            st.error(f"No pude descargar el Excel de OneDrive ({motivo(e)}).")
            st.stop()
        st.warning(f"No pude descargar el Excel de OneDrive ({motivo(e)}). Muestro la copia local.")
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
unicos = datos.drop_duplicates(["CVE", "ART_ID"])   # un artículo cuenta una vez por proveedor

# ── Precios comparables por producto: cada proveedor con su precio mediano por kg / L / pieza ──
precios = (unicos.dropna(subset=["PRECIO_U", "BASE"]).groupby(["PROD", "BASE", "CVE"])
           .agg(precio=("PRECIO_U", "median"), arts=("ART_ID", "nunique")).reset_index())
grp = precios.groupby(["PROD", "BASE"])
precios["n_prov"] = grp["CVE"].transform("nunique")
precios = precios[precios["n_prov"] >= 2].copy()          # solo donde hay competencia
grp = precios.groupby(["PROD", "BASE"])
precios["mediana"] = grp["precio"].transform("median")
precios["minimo"] = grp["precio"].transform("min")
precios["maximo"] = grp["precio"].transform("max")
precios["indice"] = precios["precio"] / precios["mediana"] * 100
precios["lugar"] = grp["precio"].rank(method="min")
nombre_prod = datos.groupby("PROD")["PRODUCTO"].first()
precios["PRODUCTO"] = precios["PROD"].map(nombre_prod)

# ── Indicadores por proveedor ──
m = unicos.groupby("CVE").agg(nombre=("PROV_CORTO", "first"), razon=("PROVEEDOR", "first"),
                              art=("ART_ID", "nunique"), prod=("PROD", "nunique"), marcas=("MARCA", "nunique"),
                              hojas=("NEGOCIACION", "nunique"))
m["regiones"] = unicos.groupby("CVE")["REGIONES"].agg(lambda s: len({r for t in s for r in t}))
m["nacional"] = unicos.groupby("CVE")["ALCANCE"].agg(lambda s: (s == "Nacional").any())
ofertas = unicos[(unicos["INVENTARIO"] == "Oferta") & ~(unicos["VIG_FECHA"] < hoy)]
m["ofertas"] = ofertas.groupby("CVE")["ART_ID"].nunique().reindex(m.index, fill_value=0)
pp = precios.groupby("CVE")
m["precio"] = pp["indice"].median()                            # índice de precio: 100 = mediana del mercado
m["competidos"] = pp.size().reindex(m.index, fill_value=0)
m["lider_precio"] = precios[precios["lugar"] == 1].groupby("CVE").size().reindex(m.index, fill_value=0)
prov_por_prod = unicos.groupby("PROD")["CVE"].nunique()
m["exclusivos"] = unicos[unicos["PROD"].map(prov_por_prod) == 1].groupby("CVE")["PROD"].nunique() \
    .reindex(m.index, fill_value=0)

# Puntaje general (0–100): cada indicador contra el mejor del programa, con pesos.
# El de precio solo cuenta para quien tiene productos en competencia (los pesos se reparten si no).
PESOS = {"art": 30, "prod": 20, "regiones": 20, "precio": 20, "ofertas": 10}
norm = pd.DataFrame({c: m[c] / m[c].max() for c in ["art", "prod", "regiones", "ofertas"]})
norm["precio"] = m["precio"].min() / m["precio"]
norm = norm.clip(0, 1)
peso = pd.DataFrame({c: norm[c].notna() * w for c, w in PESOS.items()})
m["puntaje"] = (norm.fillna(0) * pd.Series(PESOS)).sum(axis=1) / peso.sum(axis=1) * 100

N = len(m)
# Alias estable para ocultar nombres: número según el lugar en artículos
m = m.sort_values(["art", "nombre"], ascending=[False, True])
m["alias"] = [f"Proveedor {i:02d}" for i in range(1, N + 1)]

# Indicadores del ranking: columna, ¿menor es mejor?, unidad, formato
METRICAS = {
    "Puntaje general": ("puntaje", False, "pts", "{:.0f}"),
    "Artículos": ("art", False, "artículos", "{:,.0f}"),
    "Productos": ("prod", False, "productos", "{:,.0f}"),
    "Marcas": ("marcas", False, "marcas", "{:,.0f}"),
    "Regiones": ("regiones", False, "regiones", "{:,.0f}"),
    "Ofertas vigentes": ("ofertas", False, "ofertas", "{:,.0f}"),
    "Precio competitivo": ("precio", True, "índice", "{:.0f}"),
}
for _, (col, asc, _, _) in METRICAS.items():
    m[f"rk_{col}"] = m[col].rank(ascending=asc, method="min")
m["rk_lider_precio"] = m["lider_precio"].rank(ascending=False, method="min")

# ──────────────────────────────────────────────────────────────
# ENCABEZADO
# ──────────────────────────────────────────────────────────────
vista = st.query_params.get("vista")
vista = vista if vista in m.index else None   # vínculo fijo para un proveedor

with st.container(key="hero", horizontal=True, vertical_alignment="top", gap="medium"):
    md(f'<div class="hero"><div class="marca">{LOGO}<span class="sep"></span>'
       f'<div class="gob">Alimentación para el Bienestar</div></div>'
       f'<h1>Tablero de Proveedores</h1>'
       f'<div class="gerencia">Programa de Abasto Rural 2026 · Dirección Comercial</div>'
       f'<div class="subgerencia">Subgerencia de Abarrotes</div>'
       f'<div class="fecha">Información al {fecha_larga(hoy)} · {origen}</div></div>')
    with st.container(key="controles", width="content", gap="small"):
        st.toggle(":material/dark_mode: Modo oscuro", key="oscuro")
        if not vista:
            st.toggle(":material/visibility_off: Ocultar nombres", key="anonimo", value=True,
                      help="Muestra a los demás proveedores como «Proveedor 01, 02…». Úsalo al presentar a un proveedor.")
        if cfg_od.get("link"):
            st.button("Actualizar datos", icon=":material/refresh:", key="actualizar", on_click=descargar_excel.clear)

anonimo = True if vista else st.session_state.get("anonimo", True)

# ── ¿Qué proveedor eres? ──
with st.container(key="quien"):
    if vista:
        yo = vista
        md(f'<div class="vista-fija"><div class="ico">{svg("proveedor", 26)}</div><div>'
           f'<span>Tablero de</span><b>{escape(m.at[yo, "nombre"])}</b>'
           f'<span>Los demás proveedores aparecen sin nombre.</span></div></div>')
    else:
        c1, c2 = st.columns([2, 3], vertical_alignment="center")
        md('<div class="quien"><h3>¿Qué proveedor quieres ver?</h3>'
           '<p>Elige uno para ver su lugar frente a los demás; sin elegir, ves el panorama de todo el programa.</p></div>',
           c1)
        opciones = [""] + sorted(m.index, key=lambda c: m.at[c, "nombre"])
        yo = c2.selectbox("Proveedor", opciones, key="yo", label_visibility="collapsed",
                          format_func=lambda c: "— Todo el programa —" if not c else m.at[c, "nombre"]) or None


def alias(cve):
    """Nombre a mostrar: el real para el proveedor elegido o si los nombres no están ocultos."""
    return m.at[cve, "nombre"] if (not anonimo or cve == yo) else m.at[cve, "alias"]


SECCIONES = {"posicion": "Tu posición" if yo else "Panorama", "ranking": "Ranking de proveedores",
             "precios": "Competencia en precio", "regiones": "Cobertura regional", "duelo": "Cara a cara"}


def seccion(ancla, sub):
    n = list(SECCIONES).index(ancla) + 1
    md(f'<div class="seccion" id="{ancla}"><span class="paso">{n}</span><div><h2>{SECCIONES[ancla]}</h2>'
       f'<p>{sub}</p></div></div>')


md('<nav class="indice">' + "".join(f'<a href="#{k}"><span>{i}</span>{v}</a>'
                                    for i, (k, v) in enumerate(SECCIONES.items(), start=1)) + "</nav>")


def badge_lugar(rk, total):
    """Etiqueta del lugar con color según el tercio en que cae."""
    if pd.isna(rk):
        return '<span class="badge bajo">sin dato</span>'
    rk = int(rk)
    clase = "oro" if rk <= 3 else "ok" if rk <= total / 3 else "medio" if rk <= 2 * total / 3 else "bajo"
    return f'<span class="badge {clase}">#{rk} de {total}</span>'


def kpi_html(etiqueta, valor, sub="", badge="", icono="", pct=None, color=None):
    estilo_c = f' style="--c:{color};--ci:{P627 if color in CLAROS else "#fff"}"' if color else ""
    barra = f'<div class="pbar"><i style="width:{pct:.0f}%"></i></div>' if pct is not None else ""
    return (f'<div class="m-card kpi"{estilo_c}><div class="ico">{svg(icono, 22)}</div><div>'
            f'<div class="l">{etiqueta}</div><div class="row"><div class="v">{valor}</div>{badge}</div>'
            f'{barra}<div class="s">{sub}</div></div></div>')


# ──────────────────────────────────────────────────────────────
# 1. TU POSICIÓN / PANORAMA
# ──────────────────────────────────────────────────────────────
n_compet = int((prov_por_prod >= 2).sum())
if yo:
    r = m.loc[yo]
    seccion("posicion", f"Cómo está <b>{escape(r['nombre'])}</b> frente a los otros {N - 1} proveedores del programa.")

    def pct_de(col, asc=False):
        """Qué tan arriba está: 100 = primer lugar."""
        s = m[col].dropna()
        if pd.isna(r[col]) or len(s) < 2:
            return None
        return float((s > r[col]).mean() * 100 if asc else (s < r[col]).mean() * 100)

    items = [
        ("Puntaje general", f"{r.puntaje:.0f}", f"Promedio {m.puntaje.mean():.0f} · líder {m.puntaje.max():.0f}",
         "puntaje", "trofeo", P7421, False),
        ("Artículos", f"{r.art:,}", f"Promedio {m.art.mean():.0f} · líder {m.art.max():,}", "art", "articulo", P626, False),
        ("Productos", f"{r['prod']:,}", f"Promedio {m['prod'].mean():.0f} · líder {m['prod'].max():,}", "prod",
         "producto", P7420, False),
        ("Regiones", f"{r.regiones} / {len(REGIONES)}",
         "Cobertura nacional" if r.nacional else f"Promedio {m.regiones.mean():.1f} regiones", "regiones", "mapa",
         P1255, False),
        ("Índice de precio", "–" if pd.isna(r.precio) else f"{r.precio:.0f}",
         "Sin productos en competencia" if pd.isna(r.precio) else
         ("Más barato que la mediana" if r.precio < 100 else "Más caro que la mediana" if r.precio > 100
          else "Igual a la mediana") + " (100)", "precio", "balanza", P1255_60, True),
    ]
    md('<div class="kpis cinco">' + "".join(
        kpi_html(et, val, sub, badge_lugar(r[f"rk_{col}"], int(m[col].notna().sum())), ic, pct_de(col, asc), c)
        for et, val, sub, col, ic, c, asc in items) + "</div>")

    # Lectura rápida personalizada
    hall = []
    rk_art = int(r.rk_art)
    if rk_art == 1:
        hall.append((P7421, "trofeo", f"<b>Eres el proveedor con más artículos</b> del programa ({r.art:,})."))
    else:
        arriba = m[m.rk_art < rk_art].sort_values("art").iloc[0]
        hall.append((P7421, "objetivo", f"Estás en el lugar <b>#{rk_art} de {N}</b> en artículos. Te faltan "
                                        f"<b>{arriba.art - r.art + 1:,}</b> para pasar al lugar <b>#{int(arriba.rk_art)}</b>."))
    if r.competidos:
        hall.append((P626, "balanza", f"Tienes el <b>mejor precio</b> en <b>{r.lider_precio:,}</b> de tus "
                                      f"<b>{r.competidos:,}</b> productos con competencia."))
    else:
        hall.append((P626, "balanza", "Ninguno de tus productos tiene otro proveedor: aún no compites en precio."))
    if r.nacional:
        hall.append((P1255, "mapa", f"Tienes <b>cobertura nacional</b>: llegas a las {len(REGIONES)} regiones."))
    else:
        mias = {x for t in unicos.loc[unicos.CVE == yo, "REGIONES"] for x in t}
        faltan = [s for s in REGIONES if s not in mias]
        hall.append((P1255, "mapa", f"Llegas a <b>{len(mias)} de {len(REGIONES)}</b> regiones. Te faltan: "
                                    f"<b>{', '.join(REGIONES[s][0] for s in faltan[:4])}</b>"
                                    f"{'…' if len(faltan) > 4 else ''}."))
    hall.append((P7420, "producto", f"<b>{r['prod']:,}</b> productos con <b>{r.marcas:,}</b> marcas; en "
                                    f"<b>{r.exclusivos:,}</b> de ellos eres el único proveedor."))
    hall.append((P627, "estrella", f"Tienes <b>{r.ofertas:,}</b> ofertas vigentes; el promedio del programa es "
                                   f"<b>{m.ofertas.mean():.1f}</b>."))
    pt_rk = int(r.rk_puntaje)
    hall.append((P1255_60, "tendencia", f"Tu puntaje general es <b>{r.puntaje:.0f}</b>: lugar <b>#{pt_rk} de {N}</b>"
                                        f"{' — ¡estás en el podio!' if pt_rk <= 3 else ''}."))
else:
    seccion("posicion", "El programa de un vistazo. Elige un proveedor arriba para ver su lugar frente a los demás.")
    md('<div class="kpis">' + "".join([
        kpi_html("Proveedores", f"{N:,}", "participan en el programa", icono="proveedor", color=P626),
        kpi_html("Productos", f"{unicos.PROD.nunique():,}", f"{n_compet:,} con dos o más proveedores", icono="producto",
                 color=P7420),
        kpi_html("Artículos", f"{unicos.ART_ID.nunique():,}", f"Promedio por proveedor: {m.art.mean():.0f}",
                 icono="articulo", color=P1255),
        kpi_html("Ofertas vigentes", f"{ofertas.ART_ID.nunique():,}", f"de {int((m.ofertas > 0).sum())} proveedores",
                 icono="estrella", color=P7421),
    ]) + "</div>")
    lider = m.sort_values("puntaje", ascending=False).iloc[0]
    lp = m.sort_values(["lider_precio", "precio"], ascending=[False, True]).iloc[0]
    mas_comp = prov_por_prod.sort_values(ascending=False)
    top5 = m.art.nlargest(5).sum() / m.art.sum()
    hall = [
        (P7421, "trofeo", f"<b>{escape(alias(lider.name))}</b> encabeza el puntaje general con <b>{lider.puntaje:.0f}</b>."),
        (P626, "balanza", f"<b>{escape(alias(lp.name))}</b> tiene el mejor precio en <b>{lp.lider_precio:,}</b> productos."),
        (P1255, "producto", f"<b>{escape(nombre_prod[mas_comp.index[0]])}</b> es el producto más disputado: "
                            f"<b>{mas_comp.iloc[0]}</b> proveedores."),
        (P7420, "objetivo", f"<b>{n_compet / len(prov_por_prod):.0%}</b> de los productos tienen competencia "
                            f"(dos o más proveedores)."),
        (P627, "articulo", f"Los <b>5</b> proveedores con más artículos concentran el <b>{top5:.0%}</b> del surtido."),
        (P1255_60, "mapa", f"<b>{int(m.nacional.sum())}</b> proveedores tienen cobertura nacional; "
                           f"<b>{int((~m.nacional).sum())}</b> son regionales."),
    ]
md('<div class="hallazgos">' + "".join(
    f'<div class="m-card hz" style="--c:{c};--ci:{P627 if c in CLAROS else "#fff"}"><div class="ico">{svg(i, 17)}</div>'
    f'<div>{t}</div></div>' for c, i, t in hall) + "</div>")

if yo:
    st.write("")
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Tu perfil contra el programa", "radar",
                     "Cada eje compara contra el mejor proveedor en ese indicador (100 = el mejor).",
                     extra=f"Puntaje {m.at[yo, 'puntaje']:.0f}"):
            ejes = {"Artículos": "art", "Productos": "prod", "Marcas": "marcas", "Regiones": "regiones",
                    "Ofertas": "ofertas", "Precio": "precio"}
            nrm = pd.DataFrame({k: (m[c].min() / m[c] if c == "precio" else m[c] / m[c].max()) * 100
                                for k, c in ejes.items()}).clip(0, 100)
            th = list(ejes) + [list(ejes)[0]]
            fig = go.Figure()
            for nombre_t, vals, color, relleno, linea in [
                ("Promedio del programa", nrm.mean(), T["tenue"], .08, "dot"),
                (m.at[yo, "nombre"], nrm.loc[yo].fillna(0), T["acento"], .25, "solid")]:
                v = list(vals[list(ejes)]) + [vals[list(ejes)[0]]]
                fig.add_trace(go.Scatterpolar(r=v, theta=th, name=corto(nombre_t, 34), fill="toself",
                                              fillcolor=alfa(color, relleno), line=dict(color=color, width=2, dash=linea),
                                              hovertemplate="%{theta}: %{r:.0f}<extra></extra>"))
            estilo(fig, 400, leyenda=True)
            fig.update_layout(polar=dict(bgcolor="rgba(0,0,0,0)",
                                         radialaxis=dict(range=[0, 100], gridcolor=T["borde"], tickfont=dict(size=9),
                                                         showline=False, tickvals=[25, 50, 75, 100]),
                                         angularaxis=dict(gridcolor=T["borde"], tickfont=dict(color=T["texto2"], size=12))),
                              margin=dict(l=40, r=40, t=20, b=20))
            mostrar(fig, "g_radar")
    with b:
        with tarjeta("¿Quién está cerca de ti?", "vecinos", "Los proveedores justo arriba y abajo en el puntaje general.",
                     extra=f"#{int(m.at[yo, 'rk_puntaje'])} de {N}"):
            orden_p = m.sort_values("puntaje", ascending=False)
            i = list(orden_p.index).index(yo)
            ventana = orden_p.iloc[max(0, i - 3): i + 4]
            filas = [[lugar(x.rk_puntaje), escape(alias(c)), f"{x.puntaje:.0f}",
                      ("" if c == yo else f"{x.puntaje - m.at[yo, 'puntaje']:+.0f}")]
                     for c, x in ventana.iterrows()]
            tabla(["Lugar", "Proveedor", "Puntaje", "Diferencia"], filas, numericas=(2, 3),
                  clases=["yo" if c == yo else "" for c in ventana.index])
            guia("El <b>puntaje general</b> combina artículos (30%), productos (20%), regiones (20%), precio (20%) y "
                 "ofertas vigentes (10%), cada uno contra el mejor del programa.")

# ──────────────────────────────────────────────────────────────
# 2. RANKING
# ──────────────────────────────────────────────────────────────
seccion("ranking", "Elige el indicador para ordenar a todos los proveedores. "
        + ("Tu renglón está resaltado." if yo else ""))
met = st.segmented_control("Indicador", list(METRICAS), default="Puntaje general", key="metrica",
                           label_visibility="collapsed") or "Puntaje general"
col, asc, unidad, fmt = METRICAS[met]
serie = m[col].dropna().sort_values(ascending=asc, kind="stable")
rk = m[f"rk_{col}"]

if met == "Precio competitivo":
    guia("El <b>índice de precio</b> compara el costo por kilo, litro o pieza de cada producto contra la mediana de "
         "todos los proveedores que lo surten (<b>100 = mediana</b>; 90 = 10% más barato). Solo cuentan los productos "
         "con dos o más proveedores. <b>Menor es mejor.</b>")

a, b = st.columns([2, 3], gap="medium")
with a:
    with tarjeta(f"Podio · {met}", "podio", "Los tres primeros lugares.", extra=f"{len(serie):,} proveedores"):
        top3 = list(serie.index[:3])
        orden_podio = [(top3[1], 2, "🥈"), (top3[0], 1, "🥇"), (top3[2], 3, "🥉")] if len(top3) == 3 else \
            [(c, k + 1, "🥇🥈🥉"[k]) for k, c in enumerate(top3)]
        md('<div class="podio">' + "".join(
            f'<div class="p p{k}{" yo" if c == yo else ""}"><div class="medalla">{med}</div>'
            f'<div class="n" title="{escape(alias(c))}">{escape(alias(c))}</div>'
            f'<div class="v">{fmt.format(m.at[c, col])} {unidad}</div><div class="base"></div></div>'
            for c, k, med in orden_podio) + "</div>")
        if yo:
            md("<br>")
            if yo in serie.index:
                pos = int(rk[yo])
                if pos <= 3:
                    texto = f"¡<b>{escape(m.at[yo, 'nombre'])}</b> está en el podio de {met.lower()}!"
                else:
                    tercero = serie.iloc[2]
                    falta = abs(tercero - m.at[yo, col])
                    texto = (f"<b>{escape(m.at[yo, 'nombre'])}</b> está en el lugar <b>#{pos}</b>. Para subir al podio "
                             f"le faltan <b>{fmt.format(falta)}</b> {unidad}"
                             f"{' de diferencia' if asc else ''}.")
            else:
                texto = "Sin dato en este indicador: no tiene productos que otros proveedores también surtan."
            guia(texto)
with b:
    with tarjeta(f"Primeros 15 · {met}", "top", "Barras ordenadas del primer lugar hacia abajo"
                 + (" · tu lugar aparece al final si no estás entre los 15." if yo else "."), extra=unidad):
        top = list(serie.index[:15])
        if yo and yo in serie.index and yo not in top:
            top.append(yo)
        vals = serie[top]
        colores = [T["acento"] if c == yo else T["oro"] if rk[c] <= 3 else alfa(T["verde"], .7) for c in top]
        etiquetas = [f"#{int(rk[c])} · {corto(alias(c), 26)}" for c in top]
        fig = go.Figure(go.Bar(
            x=vals.values, y=etiquetas, orientation="h", marker_color=colores,
            text=[fmt.format(v) for v in vals.values], textposition="outside", cliponaxis=False,
            textfont=dict(color=T["texto2"], size=11),
            hovertemplate="<b>%{y}</b><br>%{x:,.0f} " + unidad + "<extra></extra>"))
        if met == "Precio competitivo":
            fig.add_vline(x=100, line_dash="dot", line_color=T["tenue"], annotation_text="mediana",
                          annotation_font_color=T["tenue"])
        estilo(fig, max(300, 26 * len(top) + 40))
        fig.update_layout(margin=dict(t=4, b=4, r=8))
        fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)", tickfont=dict(size=11, color=T["texto2"]))
        fig.update_xaxes(visible=False, range=[0, vals.max() * 1.18])
        mostrar(fig, "g_top")

st.write("")
with tarjeta("Tabla completa de posiciones", "tabla_rank",
             f"Todos los proveedores ordenados por {met.lower()}. Los lugares de cada columna son independientes.",
             extra=f"{N:,} proveedores"):
    orden_t = list(serie.index) + [c for c in m.index if c not in serie.index]
    filas = []
    for c in orden_t:
        x = m.loc[c]
        filas.append([
            lugar(rk[c]), escape(alias(c)) + ("" if anonimo and c != yo else f"<small>CVE {escape(c)}</small>"),
            f"{x.puntaje:.0f}", f"{x.art:,}", f"{x['prod']:,}", f"{x.marcas:,}",
            "Nacional" if x.nacional else f"{x.regiones} / {len(REGIONES)}", f"{x.ofertas:,}",
            "–" if pd.isna(x.precio) else f"{x.precio:.0f}",
            f"{x.lider_precio:,} de {x.competidos:,}" if x.competidos else "–"])
    tabla(["Lugar", "Proveedor", "Puntaje", "Artículos", "Productos", "Marcas", "Regiones", "Ofertas",
           "Índice precio", "Mejor precio en"], filas, numericas=(2, 3, 4, 5, 6, 7, 8, 9), alto=480,
          clases=["yo" if c == yo else "" for c in orden_t])

# ──────────────────────────────────────────────────────────────
# 3. COMPETENCIA EN PRECIO
# ──────────────────────────────────────────────────────────────
seccion("precios", "Precio por kilo, litro o pieza de cada proveedor en los productos que surten dos o más. "
        "Así se comparan presentaciones distintas del mismo producto.")


def chip_indice(v):
    if pd.isna(v):
        return "–"
    clase = "bien" if v <= 95 else "medio" if v <= 105 else "mal"
    return f'<span class="chip {clase}">{v:.0f}</span>'


if yo:
    mios = precios[precios.CVE == yo].sort_values(["lugar", "indice"])
    solos = unicos.loc[(unicos.CVE == yo) & (unicos.PROD.map(prov_por_prod) == 1), "PRODUCTO"].unique()
    a, b = st.columns([3, 2], gap="medium")
    with a:
        with tarjeta("Tus productos frente a la competencia", "p_mios",
                     "Tu precio mediano contra la mediana y el mejor precio del mercado.",
                     extra=[f"{len(mios):,} en competencia", f"{int((mios.lugar == 1).sum())} con el mejor precio"]):
            if mios.empty:
                md('<div class="lb-mas">Ninguno de tus productos tiene otro proveedor todavía.</div>')
            else:
                filas = [[lugar(x.lugar) + f'<small>de {int(x.n_prov)}</small>', escape(x.PRODUCTO),
                          precio_u(x.precio, x.BASE), precio_u(x.mediana, x.BASE), precio_u(x.minimo, x.BASE),
                          chip_indice(x.indice)] for x in mios.itertuples()]
                tabla(["Tu lugar", "Producto", "Tu precio", "Mediana", "Mejor precio", "Índice"], filas,
                      numericas=(2, 3, 4, 5), alto=420)
            if len(solos):
                guia(f"Eres el <b>único proveedor</b> de: {escape(', '.join(sorted(solos)[:8]))}"
                     f"{'…' if len(solos) > 8 else ''}.")
    with b:
        with tarjeta("Dónde puedes ganar", "p_oportunidad",
                     "Tus productos más caros que la mediana: cuánto bajar para igualarla.", extra="índice > 100"):
            caros = mios[mios.indice > 100].sort_values("indice", ascending=False).head(10)
            if caros.empty:
                md('<div class="lb-mas">✅ Todos tus productos están en la mediana del mercado o por debajo.</div>')
            else:
                fig = go.Figure(go.Bar(
                    x=caros.indice - 100, y=[corto(p, 26) for p in caros.PRODUCTO], orientation="h",
                    marker_color=T["alerta"], text=[f"+{v - 100:.0f}%" for v in caros.indice], textposition="outside",
                    cliponaxis=False, textfont=dict(color=T["texto2"], size=11),
                    customdata=[[precio_u(x.precio, x.BASE), precio_u(x.mediana, x.BASE)] for x in caros.itertuples()],
                    hovertemplate="<b>%{y}</b><br>Tu precio %{customdata[0]} · mediana %{customdata[1]}<extra></extra>"))
                estilo(fig, max(240, 30 * len(caros) + 40))
                fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")
                fig.update_xaxes(visible=False, range=[0, (caros.indice - 100).max() * 1.3])
                fig.update_layout(margin=dict(t=4, b=4))
                mostrar(fig, "g_caros")
    lista_prod = list(dict.fromkeys(mios.PROD)) or list(precios.PROD.unique())
else:
    resumen = (precios.groupby(["PROD", "BASE"])
               .agg(producto=("PRODUCTO", "first"), n=("CVE", "nunique"), minimo=("minimo", "first"),
                    mediana=("mediana", "first"), maximo=("maximo", "first")).reset_index()
               .sort_values(["n", "producto"], ascending=[False, True]))
    lider_p = precios[precios.lugar == 1].drop_duplicates(["PROD", "BASE"]).set_index(["PROD", "BASE"])["CVE"]
    with tarjeta("Productos con competencia", "p_todos",
                 "Cuántos proveedores surten cada producto y qué tan separados están sus precios.",
                 extra=f"{len(resumen):,} productos"):
        filas = [[f"{x.n}", escape(x.producto), escape(alias(lider_p[(x.PROD, x.BASE)])),
                  precio_u(x.minimo, x.BASE), precio_u(x.mediana, x.BASE), precio_u(x.maximo, x.BASE),
                  f"{x.maximo / x.minimo:.1f}×"] for x in resumen.itertuples()]
        tabla(["Proveedores", "Producto", "Mejor precio", "Precio más bajo", "Mediana", "Precio más alto",
               "Diferencia"], filas, numericas=(0, 3, 4, 5, 6), alto=420)
    lista_prod = list(dict.fromkeys(resumen.PROD))

st.write("")
with tarjeta("Detalle de un producto", "p_detalle", "Cada barra es un proveedor; la línea punteada es la mediana."):
    prod_sel = st.selectbox("Producto", lista_prod, key=f"prod_{yo or 'todos'}", label_visibility="collapsed",
                            format_func=lambda p: nombre_prod.get(p, p))
    d = precios[precios.PROD == prod_sel].sort_values("precio")
    for base, g in d.groupby("BASE", sort=False):
        etiquetas = [f"#{int(x.lugar)} · {corto(alias(x.CVE), 28)}" for x in g.itertuples()]
        fig = go.Figure(go.Bar(
            x=g.precio, y=etiquetas, orientation="h",
            marker_color=[T["acento"] if c == yo else T["oro"] if lg == 1 else alfa(T["verde"], .6)
                          for c, lg in zip(g.CVE, g.lugar)],
            text=[f"${v:,.2f}/{base}" for v in g.precio], textposition="outside", cliponaxis=False,
            textfont=dict(color=T["texto2"], size=11), customdata=g.arts,
            hovertemplate="<b>%{y}</b><br>$%{x:,.2f} por " + base + " · %{customdata} artículos<extra></extra>"))
        fig.add_vline(x=g.mediana.iloc[0], line_dash="dot", line_color=T["tenue"],
                      annotation_text=f"mediana ${g.mediana.iloc[0]:,.2f}", annotation_font_color=T["tenue"])
        estilo(fig, max(200, 32 * len(g) + 60))
        fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)", tickfont=dict(color=T["texto2"]))
        fig.update_xaxes(visible=False, range=[0, g.precio.max() * 1.3])
        fig.update_layout(margin=dict(t=24, b=4))
        mostrar(fig, f"g_det_{base}")

# ──────────────────────────────────────────────────────────────
# 4. COBERTURA REGIONAL
# ──────────────────────────────────────────────────────────────
seccion("regiones", "Cuántos proveedores compiten en cada gerencia regional"
        + (" y en cuáles estás tú." if yo else "."))
exp = unicos[["CVE", "ART_ID", "REGIONES"]].explode("REGIONES").dropna()
reg = exp.groupby("REGIONES").agg(prov=("CVE", "nunique"), art=("ART_ID", "nunique")).reindex(list(REGIONES), fill_value=0)
art_rp = exp.groupby(["REGIONES", "CVE"])["ART_ID"].nunique()
if yo:
    reg["mios"] = [int(art_rp.get((s, yo), 0)) for s in reg.index]
    reg["lugar"] = [int(art_rp[s].rank(ascending=False, method="min").get(yo)) if reg.at[s, "mios"] else None
                    for s in reg.index]

a, b = st.columns([3, 2], gap="medium")
with a:
    with tarjeta("Mapa de la competencia", "mapa",
                 ("En color, las regiones donde estás; pasa el cursor para ver tu lugar en cada una." if yo else
                  "Más oscuro = más proveedores compitiendo en la región.")):
        gj = geo_estados()
        if not gj:
            st.caption(f"No encontré {MAPA.name}: el mapa no se puede dibujar.")
        else:
            locs, z, textos = [], [], []
            for f in gj["features"]:
                e = f["properties"]["name"]
                s = ESTADO_REGION.get(e)
                locs.append(e)
                if not s:
                    z.append(None if not yo else 0)
                    textos.append(f"<b>{e}</b><br>Sin gerencia regional")
                    continue
                x = reg.loc[s]
                t = f"<b>{s} · {REGIONES[s][0]}</b><br>{e}<br>{int(x.prov)} proveedores · {int(x.art):,} artículos"
                if yo:
                    t += (f"<br><b>Tú: {int(x.mios)} artículos · lugar #{int(x.lugar)} de {int(x.prov)}</b>" if x.mios
                          else "<br><b>No llegas a esta región</b>")
                    z.append(1 if x.mios else 0)
                else:
                    z.append(int(x.prov))
                textos.append(t)
            escala = ([[0, T["otros"]], [1, T["acento"]]] if yo else
                      [[0, alfa(T["acento"], .12)], [1, T["acento"]]])
            fig = go.Figure(go.Choropleth(
                geojson=gj, featureidkey="properties.name", locations=locs, z=z, colorscale=escala,
                zmin=0, zmax=1 if yo else max(reg.prov.max(), 1), showscale=False,
                marker_line_color=T["card"], marker_line_width=1, customdata=textos,
                hovertemplate="%{customdata}<extra></extra>"))
            centros = centros_region()
            sig = [s for s in REGIONES if s in centros]
            fig.add_trace(go.Scattergeo(lon=[centros[s][0] for s in sig], lat=[centros[s][1] for s in sig],
                                        text=[f"{s}<br>{int(reg.at[s, 'prov'])}" for s in sig], mode="text",
                                        textfont=dict(family="Inter, sans-serif", size=10, color=T["texto"]),
                                        hoverinfo="skip"))
            fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)", projection_type="mercator")
            estilo(fig, 440)
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False)
            st.plotly_chart(fig, width="stretch", theme=None, key="g_mapa",
                            config={"displayModeBar": False, "scrollZoom": False})
with b:
    with tarjeta("Competencia por región", "reg_tabla", "Proveedores y artículos que llegan a cada región "
                 "(los nacionales cuentan en todas).", extra=f"{len(REGIONES)} regiones"):
        orden_r = reg.sort_values(["prov", "art"], ascending=False)
        if yo:
            filas = [[f"<b>{s}</b>", escape(REGIONES[s][0]), f"{int(x.prov)}", f"{int(x.art):,}",
                      f"{int(x.mios):,}" if x.mios else "–",
                      lugar(x.lugar) if x.mios else '<span class="chip mal">No llegas</span>']
                     for s, x in orden_r.iterrows()]
            tabla(["Sigla", "Región", "Proveedores", "Artículos", "Tuyos", "Tu lugar"], filas, numericas=(2, 3, 4, 5),
                  clases=["yo" if x.mios else "" for _, x in orden_r.iterrows()])
        else:
            top_r = art_rp.groupby(level=0).idxmax()
            filas = [[f"<b>{s}</b>", escape(REGIONES[s][0]), f"{int(x.prov)}", f"{int(x.art):,}",
                      escape(alias(top_r[s][1])) if s in top_r else "–"] for s, x in orden_r.iterrows()]
            tabla(["Sigla", "Región", "Proveedores", "Artículos", "Líder en artículos"], filas, numericas=(2, 3))

# ──────────────────────────────────────────────────────────────
# 5. CARA A CARA
# ──────────────────────────────────────────────────────────────
seccion("duelo", "Compara dos proveedores indicador por indicador; gana quien tenga más (o el menor índice de precio).")
DUELO = [("Puntaje general", "puntaje", False, "{:.0f}"), ("Artículos", "art", False, "{:,.0f}"),
         ("Productos", "prod", False, "{:,.0f}"), ("Marcas", "marcas", False, "{:,.0f}"),
         ("Regiones", "regiones", False, "{:,.0f}"), ("Ofertas vigentes", "ofertas", False, "{:,.0f}"),
         ("Hojas de negociación", "hojas", False, "{:,.0f}"), ("Índice de precio", "precio", True, "{:.0f}"),
         ("Productos con mejor precio", "lider_precio", False, "{:,.0f}"),
         ("Productos exclusivos", "exclusivos", False, "{:,.0f}")]
ESPECIALES = {"__mejor__": "El mejor en cada indicador", "__prom__": "Promedio del programa"}
por_puntaje = list(m.sort_values("puntaje", ascending=False).index)

with tarjeta("Cara a cara", "duelo", "Elige contra quién comparar."):
    c1, c2 = st.columns(2)
    if yo:
        izq = yo
        c1.markdown(f"**{escape(m.at[yo, 'nombre'])}**")
    else:
        izq = c1.selectbox("Proveedor A", por_puntaje, key="duelo_a", format_func=lambda c: alias(c))
    rivales = list(ESPECIALES) + [c for c in por_puntaje if c != izq]
    der = c2.selectbox("Comparar con", rivales, key="duelo_b",
                       format_func=lambda c: ESPECIALES.get(c) or f"#{int(m.at[c, 'rk_puntaje'])} · {alias(c)}")

    def valor_rival(col, asc):
        if der == "__prom__":
            return m[col].mean()
        if der == "__mejor__":
            return m[col].min() if asc else m[col].max()
        return m.at[der, col]

    filas_html, g_izq, g_der = [], 0, 0
    for et, col, asc, fmt_d in DUELO:
        vi, vd = m.at[izq, col], valor_rival(col, asc)
        tope = max(v for v in [vi, vd, 0] if pd.notna(v)) or 1
        gi = gd = False
        if pd.notna(vi) and pd.notna(vd) and vi != vd:
            gi = (vi < vd) if asc else (vi > vd)
            gd = not gi
        g_izq += gi
        g_der += gd

        def barra(v, lado, gana):
            ancho = 0 if pd.isna(v) else v / tope * 100
            txt = "–" if pd.isna(v) else fmt_d.format(v)
            return (f'<div class="b {lado}{" gana" if gana else ""}"><div class="barra" style="width:{ancho:.0f}%">'
                    f'</div><span>{txt}{" ✓" if gana else ""}</span></div>')
        filas_html.append(barra(vi, "izq", gi) + f'<div class="m">{et}</div>' + barra(vd, "der", gd))
    nombre_der = ESPECIALES.get(der) or alias(der)
    md(f'<div class="duelo-cab"><div class="lado izq">{escape(alias(izq))}</div><span class="vs">VS</span>'
       f'<div class="lado">{escape(nombre_der)}</div>'
       f'<div class="marcador">Indicadores ganados: <b>{g_izq}</b> – <b>{g_der}</b></div></div>'
       f'<div class="duelo">{"".join(filas_html)}</div>')

md('<div class="pie">Subgerencia de Abarrotes · Dirección Comercial · Alimentación para el Bienestar<br>'
   'Cifras calculadas con la hoja VIGENTES del Programa de Abasto Rural 2026.</div>')
