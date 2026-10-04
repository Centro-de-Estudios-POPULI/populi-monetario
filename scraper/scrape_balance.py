"""
Scraper: Balance del BCB (Activo y Pasivo)
Fuente: BCB — cuadro «4. Banco Central.xlsx» (Balances Consolidados), en miles de bolivianos.
Genera: data/activo_bcb.json, data/pasivo_bcb.json (en millones de Bs)

Cada partida se ubica por su RÓTULO, no por número de columna (lección del scraper de RIN:
el BCB inserta columnas y un mapa fijo lee partidas equivocadas sin avisar). Y antes de escribir
se comprueba la identidad del balance mes por mes, con todas las partidas del cuadro:

  ACTIVO = Reservas internacionales brutas + Aportes a organismos internacionales
         + Otros activos externos de mediano y largo plazo + Crédito al sector público
         + Crédito al sector financiero + Otras cuentas de activo
  PASIVO = Emisión + Depósitos bancarios + Obligaciones externas a corto plazo
         + Depósitos del sector público + Depósito de organismos internacionales
         + Obligaciones externas a mediano y largo plazo + Otras cuentas de pasivo
         + CDD (certificados de devolución de depósitos) + Capital y reservas
  y ACTIVO = PASIVO (la columna «TOTAL ACTIVO Y PASIVO»).

Si una identidad no cierra, el scraper aborta SIN escribir: el dato publicado no cambia.
"""

import datetime
import json
import re
import sys
import time
import unicodedata
from pathlib import Path

import openpyxl
import requests

URL = "https://www.bcb.gob.bo/webdocs/sector_monetario/Balances%20Consolidados/4.%20Banco%20Central.xlsx"
OUT_DIR = Path(__file__).resolve().parent.parent / "data"
XLSX_PATH = Path(__file__).resolve().parent / "bcb_raw" / "balance_bcb.xlsx"
TOLERANCIA = 2.0   # miles de Bs: el cuadro redondea cada partida por separado

# El WAF del BCB puede rechazar (403) el User-Agent por defecto de python-requests.
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
              "application/vnd.ms-excel,*/*",
    "Referer": "https://www.bcb.gob.bo/",
}

# Partida → condiciones sobre el rótulo normalizado (sin tildes, sin espacios ni signos).
# El rótulo de una columna junta todas sus filas de encabezado, con las celdas combinadas
# propagadas: «CRÉDITO AL SECTOR PÚBLICO» + «Gobierno Central» → creditoalsectorpublico…gobiernocentral.
ACTIVO = {
    "rib":          ["reservasinternacionalesbrutas", "total"],
    "oro":          ["reservasinternacionalesbrutas", "oro"],
    "divisas":      ["reservasinternacionalesbrutas", "divisas"],
    "otros_rib":    ["reservasinternacionalesbrutas", "otrosactivos"],
    "aportes_oi":   ["aportesaorganismosinternacionales"],
    "otros_ext":    ["otrosactivosext"],
    "credito_sp":   ["creditoalsectorpublico", "total"],
    "credito_gc":   ["creditoalsectorpublico", "gobiernocentral"],
    "credito_ss":   ["creditoalsectorpublico", "seguridadsocial"],
    "credito_gl":   ["creditoalsectorpublico", "locales"],
    "credito_ep":   ["creditoalsectorpublico", "empresaspublicas"],
    "credito_sf":   ["creditoalsectorfinanciero", "total"],
    "credito_bancos": ["creditoalsectorfinanciero", "bancoscomerciales"],
    "credito_oef":  ["creditoalsectorfinanciero", "especializados"],
    "otras_ctas":   ["otrascuentasdeactivo"],
    "total_activo": ["totalactivoypasivo"],
}
PASIVO = {
    "emision":       ["emision"],
    "dep_bancarios": ["depositosbancarios", "total"],
    "oblig_ext_cp":  ["obligacionesexternasacortoplazo", "total"],
    "oblig_fmi":     ["obligacionesexternasacortoplazo", "fmi"],
    "oblig_bancos_oi": ["obligacionesexternasacortoplazo", "bancos"],
    "dep_gc":        ["depositosdelsectorpublico", "gobiernocentral", "total"],
    "dep_ss":        ["depositosdelsectorpublico", "seguridadsocial", "total"],
    "dep_gl":        ["depositosdelsectorpublico", "gobiernoslocales", "total"],
    "dep_ep":        ["depositosdelsectorpublico", "empresaspublicas", "total"],
    "dep_oi":        ["depositodeorganismosinternacionales"],
    "oblig_ext_mlp": ["obligacionesexternasamedianoylargoplazo"],
    "otras_ctas":    ["otrascuentaspasivo"],
    "cdd":           ["cdd", "total"],
    "capital":       ["capitalyreservas"],
}


def download():
    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    print(f"Descargando {URL}")
    for attempt in range(3):
        try:
            r = requests.get(URL, headers=HEADERS, timeout=90)
            r.raise_for_status()
            XLSX_PATH.write_bytes(r.content)
            print(f"  -> {XLSX_PATH} ({len(r.content)//1024} KB)")
            return
        except Exception as e:
            delay = [10, 30, 60][attempt]
            if attempt < 2:
                print(f"  Intento {attempt+1}/3: {e} — reintentando en {delay}s...")
                time.sleep(delay)
            else:
                print(f"  Error tras 3 intentos: {e}")
                sys.exit(1)


def norm(s):
    s = unicodedata.normalize("NFKD", str(s).lower()).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s)


def parse_date(val):
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m")
    if isinstance(val, (int, float)):
        return f"{int(val)}-12"
    return None


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def rotulos(ws):
    """Rótulo normalizado de cada columna y fila donde empiezan los datos."""
    primera = next(r for r in range(1, 40) if parse_date(ws.cell(r, 1).value))
    filas = range(1, primera)
    celda = {(r, c): ws.cell(r, c).value for r in filas for c in range(1, ws.max_column + 1)}
    for rg in ws.merged_cells.ranges:          # propagar los encabezados combinados
        v = ws.cell(rg.min_row, rg.min_col).value
        for r in range(rg.min_row, rg.max_row + 1):
            for c in range(rg.min_col, rg.max_col + 1):
                if r in filas:
                    celda[(r, c)] = v
    lab = {c: limpio(norm(" ".join(str(celda[(r, c)]) for r in filas if celda[(r, c)] not in (None, ""))))
           for c in range(2, ws.max_column + 1)}
    return lab, primera


# El título del cuadro, «A C T I V O»/«P A S I V O» y «(En miles de bolivianos)» están combinados
# sobre muchas columnas: se quitan del comienzo del rótulo para que quede sólo la partida.
PREFIJOS = ("balancedelbancocentraldebolivia", "activo", "pasivo", "enmilesdebolivianos")


def limpio(t):
    sigue = True
    while sigue:
        sigue = False
        for p in PREFIJOS:
            if t.startswith(p) and len(t) > len(p):
                t, sigue = t[len(p):], True
    return t


def ubicar(lab, mapa, hoja):
    cols, faltan = {}, []
    for clave, partes in mapa.items():
        hallado = [c for c, t in lab.items() if all(p in t for p in partes)]
        if clave == "dep_bancarios":          # el total, no «comerciales» ni «especializados»
            hallado = [c for c in hallado if "comerciales" not in lab[c] and "especializados" not in lab[c]]
        if clave in ("rib", "credito_sp", "credito_sf", "oblig_ext_cp"):
            hallado = [c for c in hallado if lab[c].endswith("total")]
        if clave == "cdd":
            hallado = [c for c in hallado if lab[c].endswith("total")]
        if len(hallado) != 1:
            faltan.append(f"{clave} → {hallado}")
        else:
            cols[clave] = hallado[0]
    if faltan:
        raise ValueError(f"{hoja}: no se ubicaron sin ambigüedad las columnas: {faltan}")
    return cols


def leer(ws, cols, primera, clave_obligatoria):
    filas = {}
    for r in range(primera, ws.max_row + 1):
        f = parse_date(ws.cell(r, 1).value)
        if not f or not isinstance(ws.cell(r, cols[clave_obligatoria]).value, (int, float)):
            continue
        filas[f] = {k: num(ws.cell(r, c).value) for k, c in cols.items()}
    return filas


def verificar(nombre, filas, total, partes):
    malos = []
    for f, x in filas.items():
        d = total(f, x) - sum(x[p] for p in partes)
        if abs(d) > TOLERANCIA:
            malos.append((f, round(d, 1)))
    if malos:
        raise ValueError(f"{nombre}: la identidad no cierra en {len(malos)} meses (miles de Bs): {malos[:6]}")
    print(f"  ✓ {nombre}: cierra en {len(filas)} meses")


def mm(v):
    return round(v / 1000, 2)


def main():
    if "--no-download" not in sys.argv:
        download()
    wb = openpyxl.load_workbook(str(XLSX_PATH), data_only=True)

    wa, wp = wb["Activo"], wb["Pasivo"]
    lab_a, ini_a = rotulos(wa)
    lab_p, ini_p = rotulos(wp)
    ca = ubicar(lab_a, ACTIVO, "Activo")
    cp = ubicar(lab_p, PASIVO, "Pasivo")
    # «Depósitos del sector público»: el BCB rotula su total sólo como «TOTAL»; es la columna
    # inmediatamente anterior al bloque, y se comprueba que sea la suma de sus cuatro subtotales.
    cp["dep_sp_total"] = min(cp["dep_gc"], cp["dep_ss"], cp["dep_gl"], cp["dep_ep"]) - 1
    if lab_p.get(cp["dep_sp_total"]) != "total":
        raise ValueError(f"Pasivo: la columna del total de depósitos del sector público dice «{lab_p.get(cp['dep_sp_total'])}»")
    print("Columnas Activo:", ca)
    print("Columnas Pasivo:", cp)

    A = leer(wa, ca, ini_a, "rib")
    P = leer(wp, cp, ini_p, "emision")
    P = {f: x for f, x in P.items() if f in A}            # el total del pasivo es la columna del activo

    print("Identidades (miles de Bs, tolerancia ±%.0f):" % TOLERANCIA)
    verificar("Activo = 6 partidas", A, lambda f, x: x["total_activo"],
              ["rib", "aportes_oi", "otros_ext", "credito_sp", "credito_sf", "otras_ctas"])
    verificar("Pasivo = 9 partidas", P, lambda f, x: A[f]["total_activo"],
              ["emision", "dep_bancarios", "oblig_ext_cp", "dep_sp_total", "dep_oi", "oblig_ext_mlp", "otras_ctas", "cdd", "capital"])
    verificar("RIB = oro + divisas + otros", A, lambda f, x: x["rib"], ["oro", "divisas", "otros_rib"])
    verificar("Crédito SP = GC + SS + GL + EP", A, lambda f, x: x["credito_sp"], ["credito_gc", "credito_ss", "credito_gl", "credito_ep"])
    verificar("Crédito SF = bancos + otras entidades", A, lambda f, x: x["credito_sf"], ["credito_bancos", "credito_oef"])
    verificar("Depósitos SP = GC + SS + GL + EP", P, lambda f, x: x["dep_sp_total"], ["dep_gc", "dep_ss", "dep_gl", "dep_ep"])

    activo = [{"date": f, **{k: mm(v) for k, v in x.items()}} for f, x in A.items()]
    pasivo = [{"date": f, **{k: mm(v) for k, v in x.items()}, "total_pasivo": mm(A[f]["total_activo"])} for f, x in P.items()]
    ua, up = activo[-1], pasivo[-1]

    salida_a = {"metadata": {
        "titulo": "Activo del Banco Central de Bolivia",
        "subtitulo": "Composición del activo: reservas, otros activos externos, crédito al sector público y financiero",
        "fuente": "Banco Central de Bolivia (BCB) — 4. Banco Central.xlsx",
        "unidad": "Millones de Bs", "frecuencia": "Mensual",
        "identidad": "total_activo = rib + aportes_oi + otros_ext + credito_sp + credito_sf + otras_ctas",
        "ultimo_dato": ua["date"], "primer_dato": activo[0]["date"], "observaciones": len(activo),
        "total_activo_mm": ua["total_activo"], "rib_mm": ua["rib"], "credito_sp_mm": ua["credito_sp"], "credito_sf_mm": ua["credito_sf"],
    }, "series": activo}
    salida_p = {"metadata": {
        "titulo": "Pasivo del Banco Central de Bolivia",
        "subtitulo": "Composición del pasivo: emisión, depósitos, pasivos externos, capital y otras cuentas",
        "fuente": "Banco Central de Bolivia (BCB) — 4. Banco Central.xlsx",
        "unidad": "Millones de Bs", "frecuencia": "Mensual",
        "identidad": "total_pasivo = emision + dep_bancarios + oblig_ext_cp + dep_sp_total + dep_oi + oblig_ext_mlp + otras_ctas + cdd + capital",
        "ultimo_dato": up["date"], "primer_dato": pasivo[0]["date"], "observaciones": len(pasivo),
        "emision_mm": up["emision"], "dep_sp_mm": up["dep_sp_total"], "dep_bancarios_mm": up["dep_bancarios"],
        "pasivos_externos_mm": round(up["oblig_ext_cp"] + up["oblig_ext_mlp"] + up["dep_oi"], 2),
        "total_pasivo_mm": up["total_pasivo"],
    }, "series": pasivo}

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for nombre, datos in (("activo_bcb.json", salida_a), ("pasivo_bcb.json", salida_p)):
        with open(OUT_DIR / nombre, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    print(f"ACTIVO: {len(activo)} obs, {activo[0]['date']} a {ua['date']} · total {ua['total_activo']:,.0f} MM")
    print(f"PASIVO: {len(pasivo)} obs, {pasivo[0]['date']} a {up['date']} · pasivos externos {salida_p['metadata']['pasivos_externos_mm']:,.0f} MM")


if __name__ == "__main__":
    main()
