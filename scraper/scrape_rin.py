"""
Scraper: Reservas Internacionales (RIN/RIB)
Fuente: BCB — Estadísticas Semanales
Genera: data/reservas_rin.json
"""

import json, os, sys, datetime, time
from pathlib import Path
import openpyxl
import requests

URL = "https://www.bcb.gob.bo/webdocs/publicacionesbcb/estadisticasemanales/Estad%C3%ADstica%20Semanal.xlsx"
OUT_DIR = Path(__file__).resolve().parent.parent / "data"
XLSX_PATH = Path(__file__).resolve().parent / "bcb_raw" / "semanal.xlsx"

ROW_MAP = {
    "rib": 72,
    "deg": 73,
    "oro": 74,
    "divisas": 79,
    "fmi": 80,
    "rin": 81,
}


def download():
    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    print(f"Descargando {URL}")
    for attempt in range(3):
        try:
            r = requests.get(URL, timeout=90)
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


def parse():
    wb = openpyxl.load_workbook(str(XLSX_PATH), data_only=True)
    ws = wb[wb.sheetnames[0]]

    series = []
    last_month_date = None

    for c in range(5, ws.max_column + 1):
        date_val = ws.cell(3, c).value
        if date_val is None:
            continue

        if isinstance(date_val, datetime.datetime):
            date_str = date_val.strftime("%Y-%m")
            last_month_date = date_val
        elif isinstance(date_val, str) and "semana" in date_val.lower():
            if last_month_date is None:
                continue
            next_m = last_month_date.month + 1
            next_y = last_month_date.year
            if next_m > 12:
                next_m = 1
                next_y += 1
            week_num = int("".join(filter(str.isdigit, date_val)) or "1")
            date_str = f"{next_y}-{next_m:02d}-S{week_num}"
        else:
            continue

        row_data = {}
        all_none = True
        for key, row in ROW_MAP.items():
            v = ws.cell(row, c).value
            if v is not None and isinstance(v, (int, float)):
                row_data[key] = round(float(v), 2)
                all_none = False
            else:
                row_data[key] = None

        if all_none:
            continue

        entry = {"date": date_str}
        entry.update(row_data)
        series.append(entry)

    monthly = [s for s in series if "-S" not in s["date"]]
    weekly = [s for s in series if "-S" in s["date"]]

    last = monthly[-1] if monthly else series[-1]
    prev_12 = monthly[-13] if len(monthly) >= 13 else monthly[0]

    def safe_var(cur, prev):
        if cur and prev and prev > 0:
            return round((cur / prev - 1) * 100, 1)
        return None

    latest_point = weekly[-1] if weekly else last

    metadata = {
        "titulo": "Reservas Internacionales",
        "subtitulo": "Reservas Internacionales Netas y Brutas del BCB (en millones de USD)",
        "fuente": "Banco Central de Bolivia (BCB) — Estadísticas Semanales",
        "unidad": "Millones de USD",
        "frecuencia": "Mensual + Semanal",
        "ultimo_dato": latest_point["date"],
        "ultimo_mensual": last["date"],
        "primer_dato": series[0]["date"],
        "observaciones": len(series),
        "observaciones_mensuales": len(monthly),
        "ultimo_rin": latest_point.get("rin"),
        "ultimo_rib": latest_point.get("rib"),
        "var_12m_rin_pct": safe_var(last.get("rin"), prev_12.get("rin")),
    }

    return {"metadata": metadata, "series": series}


def main():
    if "--no-download" not in sys.argv:
        download()
    data = parse()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / "reservas_rin.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    m = data["metadata"]
    print(f"OK: {m['observaciones']} obs ({m['observaciones_mensuales']} mensuales), {m['primer_dato']} a {m['ultimo_dato']}")
    print(f"   RIN: ${m['ultimo_rin']:,.1f} MM USD | Var 12m: {m['var_12m_rin_pct']}%")


if __name__ == "__main__":
    main()
