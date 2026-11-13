"""
Genera los códigos de acceso del Tablero de Proveedores.

Uso:
    python generar_codigos.py                 # usa el Excel de OneDrive (secrets) o la copia local
    python generar_codigos.py ruta/al.xlsx    # usa ese Excel

Qué hace:
  - Lee los proveedores (CVE) de la hoja VIGENTES.
  - Conserva los códigos que ya existan en .streamlit/secrets.toml y crea uno nuevo solo para los proveedores
    que aún no tengan (así los códigos ya entregados no cambian).
  - Reescribe .streamlit/secrets.toml con [onedrive] y [acceso] (admin + un código por proveedor).
  - Escribe codigos_proveedores.csv (proveedor, CVE, código) para entregarlos. Ese archivo no se sube a GitHub.

Después copia el contenido de .streamlit/secrets.toml en Streamlit Community Cloud > Settings > Secrets.
"""
import csv
import io
import re
import secrets
import sys
import tomllib
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import pandas as pd
import requests

CARPETA = Path(__file__).parent
SECRETOS = CARPETA / ".streamlit" / "secrets.toml"
CSV = CARPETA / "codigos_proveedores.csv"
LOCALES = [d / "PROGRAMA DE ABASTO RURAL 2026.xlsx" for d in (CARPETA, CARPETA.parent / "Dashboard")]
ALFABETO = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # sin 0/O ni 1/I para que no se confundan al dictarlos


def nuevo_codigo() -> str:
    c = "".join(secrets.choice(ALFABETO) for _ in range(8))
    return f"{c[:4]}-{c[4:]}"


def toml_texto(v: str) -> str:
    return '"' + str(v).replace("\\", "\\\\").replace('"', '\\"') + '"'


def leer_excel(cfg: dict) -> bytes:
    if len(sys.argv) > 1:
        return Path(sys.argv[1]).read_bytes()
    link = cfg.get("onedrive", {}).get("link")
    if link and "sharepoint.com" in link:
        p = urlsplit(link)
        q = [(k, v) for k, v in parse_qsl(p.query) if k != "download"] + [("download", "1")]
        try:
            r = requests.get(urlunsplit(p._replace(query=urlencode(q))), headers={"User-Agent": "Mozilla/5.0"},
                             timeout=60)
            if r.ok and r.content.startswith(b"PK"):
                print("Excel descargado de OneDrive.")
                return r.content
        except requests.RequestException as e:
            print(f"Error de conexión con OneDrive ({type(e).__name__}).")
        print("No pude descargar de OneDrive; uso la copia local.")
    local = next((a for a in LOCALES if a.exists()), None)
    if not local:
        sys.exit("No encontré el Excel. Pásalo como argumento: python generar_codigos.py ruta/al.xlsx")
    return local.read_bytes()


def proveedores(contenido: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(contenido))
    hoja = "VIGENTES" if "VIGENTES" in xls.sheet_names else xls.sheet_names[0]
    vista = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=30)
    fila = next(i for i, r in vista.iterrows() if r.astype(str).str.strip().str.upper().isin(["NEGOCIACION", "CVE"]).any())
    df = pd.read_excel(xls, sheet_name=hoja, header=fila)
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    df = df[df["ART."].notna() & df["CVE"].notna()]
    df["CVE"] = df["CVE"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return df.groupby("CVE")["PROVEEDOR"].first().str.strip().reset_index().sort_values("PROVEEDOR")


def main():
    cfg = tomllib.loads(SECRETOS.read_text(encoding="utf-8")) if SECRETOS.exists() else {}
    acceso = cfg.get("acceso", {})
    previos = {str(k): v for k, v in acceso.get("proveedores", {}).items()}
    admin = acceso.get("admin") or nuevo_codigo()
    provs = proveedores(leer_excel(cfg))

    codigos, nuevos = {}, 0
    for cve in provs.CVE:
        if cve in previos:
            codigos[cve] = previos[cve]
        else:
            codigos[cve] = nuevo_codigo()
            nuevos += 1
    for cve, c in previos.items():   # proveedores que ya no están en el Excel: se conservan por si regresan
        codigos.setdefault(cve, c)
    nombres = dict(zip(provs.CVE, provs.PROVEEDOR))

    od = cfg.get("onedrive", {})
    lineas = ["# Secretos del Tablero de Proveedores. NO se sube a GitHub (está en .gitignore).",
              "# Copia todo este contenido en Streamlit Community Cloud > Settings > Secrets.", "", "[onedrive]"]
    lineas += [f"{k} = {toml_texto(v) if isinstance(v, str) else v}" for k, v in od.items()] or \
              ['# link = "https://...vínculo para compartir del Excel..."', "# minutos = 10"]
    lineas += ["", "[acceso]", "# Código del administrador: ve todo el programa y puede abrir el tablero de cualquier proveedor",
               f"admin = {toml_texto(admin)}", "", "[acceso.proveedores]", "# CVE = código de acceso"]
    for cve in sorted(codigos, key=lambda c: nombres.get(c, "~")):
        lineas.append(f"{toml_texto(cve)} = {toml_texto(codigos[cve])}  # {nombres.get(cve, 'ya no está en el Excel')}")
    SECRETOS.parent.mkdir(exist_ok=True)
    SECRETOS.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    with CSV.open("w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig: Excel abre bien los acentos
        w = csv.writer(f)
        w.writerow(["PROVEEDOR", "CVE", "CODIGO"])
        w.writerow(["ADMINISTRADOR", "", admin])
        for cve in sorted(codigos, key=lambda c: nombres.get(c, "~")):
            w.writerow([nombres.get(cve, "(ya no está en el Excel)"), cve, codigos[cve]])

    print(f"{len(codigos)} proveedores ({nuevos} códigos nuevos). Escribí {SECRETOS} y {CSV.name}.")


if __name__ == "__main__":
    main()
