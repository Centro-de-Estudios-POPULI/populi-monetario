"""
Parser: ARA-FMI (Assessing Reserve Adequacy)
Fuente: FMI — https://www.imf.org/external/datamapper/datasets/ARA
Genera: data/ara_fmi.json
Datos anuales, actualización manual.
"""

import xlrd, os, json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent / "data" / "ara_raw"
OUT = Path(__file__).resolve().parent.parent / "data" / "ara_fmi.json"

COUNTRIES = [
    "Bolivia", "Chile", "Peru", "Colombia", "Brazil",
    "Paraguay", "Argentina", "Uruguay", "Ecuador",
]

INDICATOR_META = {
    "Reserves_ARA": {
        "name": "Reservas / Métrica ARA",
        "threshold": 1.0,
        "desc": "Ratio compuesto de adecuación de reservas del FMI",
    },
    "Reserves_M2": {
        "name": "Reservas / Dinero Amplio (M2)",
        "threshold": 0.2,
        "desc": "Reservas como proporción de M2 (estándar mínimo: 20%)",
    },
    "Reserves_STD": {
        "name": "Reservas / Deuda Externa CP",
        "threshold": 1.0,
        "desc": "Cobertura de deuda externa de corto plazo (regla Guidotti: ≥1)",
    },
    "Reserves_M": {
        "name": "Reservas / Meses de Importación",
        "threshold": 3.0,
        "desc": "Meses de importación cubiertos por reservas (prudente: ≥3)",
    },
}

def parse():
    all_data = {}

    for f in sorted(os.listdir(BASE)):
        if not f.endswith(".xls"):
            continue
        wb = xlrd.open_workbook(str(BASE / f), ignore_workbook_corruption=True)
        ws = wb.sheet_by_index(0)
        indicator_key = ws.name

        years = []
        for c in range(1, ws.ncols):
            v = ws.cell_value(0, c)
            if isinstance(v, float):
                years.append(int(v))

        for r in range(2, ws.nrows):
            name = ws.cell_value(r, 0).strip()
            if name not in COUNTRIES:
                continue

            if name not in all_data:
                all_data[name] = {}
            if indicator_key not in all_data[name]:
                all_data[name][indicator_key] = {}

            for ci, yr in enumerate(years):
                v = ws.cell_value(r, ci + 1)
                if isinstance(v, (int, float)):
                    all_data[name][indicator_key][yr] = round(float(v), 4)

    all_years = sorted(set(
        y for c in all_data.values()
        for ind in c.values()
        for y in ind.keys()
    ))

    series_by_indicator = {}
    for ind_key in INDICATOR_META:
        series = {}
        for country in COUNTRIES:
            if country in all_data and ind_key in all_data[country]:
                pts = []
                for yr in all_years:
                    val = all_data[country][ind_key].get(yr)
                    if val is not None:
                        pts.append({"year": yr, "value": val})
                series[country] = pts
        series_by_indicator[ind_key] = series

    bol = all_data.get("Bolivia", {})
    latest_year = max(all_years)
    bol_latest = {}
    for ind_key in INDICATOR_META:
        vals = bol.get(ind_key, {})
        for y in range(latest_year, 1999, -1):
            if y in vals:
                bol_latest[ind_key] = {"year": y, "value": vals[y]}
                break

    output = {
        "metadata": {
            "titulo": "Suficiencia de Reservas (ARA-FMI)",
            "subtitulo": "Assessing Reserve Adequacy",
            "fuente": "Fondo Monetario Internacional (FMI)",
            "frecuencia": "Anual",
            "paises": COUNTRIES,
            "primer_dato": min(all_years),
            "ultimo_dato": max(all_years),
            "bolivia_ultimo": bol_latest,
        },
        "indicators": INDICATOR_META,
        "series": series_by_indicator,
    }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"OK: {OUT.name}")
    print(f"Años: {min(all_years)}-{max(all_years)}")
    print(f"Países: {len(COUNTRIES)}")
    for k, v in bol_latest.items():
        meta = INDICATOR_META[k]
        status = "OK" if v["value"] >= meta["threshold"] else "BAJO"
        print(f"  {meta['name']}: {v['value']:.2f} ({v['year']}) — umbral: {meta['threshold']} [{status}]")


if __name__ == "__main__":
    parse()
