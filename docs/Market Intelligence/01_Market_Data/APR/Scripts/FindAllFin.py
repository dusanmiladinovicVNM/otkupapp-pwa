import requests
import pandas as pd
import urllib3
from pathlib import Path

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://openapi.apr.gov.rs/api/opendata/companies/financial-statements"
INPUT_FILE = "Financial Dataset - Elderly care.xlsx"
OUTPUT_FILE = "Financial Dataset - Elderly care - with APR finansije.xlsx"

# putanja do foldera gde je .py fajl
BASE_DIR = Path(__file__).resolve().parent
input_path = BASE_DIR / INPUT_FILE
output_path = BASE_DIR / OUTPUT_FILE

# 1) Učitaj Excel
# kolona A = prva kolona, pa je uzimamo kao string
df_input = pd.read_excel(input_path, dtype={0: str})

# preimenuj prvu kolonu u maticni_broj ako već nije tako nazvana
first_col = df_input.columns[0]
if first_col != "maticni_broj":
    df_input = df_input.rename(columns={first_col: "maticni_broj"})

# očisti i sačuvaj vodeće nule
df_input["maticni_broj"] = (
    df_input["maticni_broj"]
    .astype(str)
    .str.strip()
    .str.replace(".0", "", regex=False)
    .str.zfill(8)
)

maticni_set = set(df_input["maticni_broj"])

# 2) Povuci APR finansije
print("Downloading financial statements dataset...")
response = requests.get(URL, timeout=120, verify=False)
response.raise_for_status()
data_api = response.json()

podaci = data_api.get("Podaci", {})

rows = []

for mb in maticni_set:
    item = podaci.get(mb)

    if item is None:
        rows.append({
            "maticni_broj": mb,
            "apr_naziv": None,
            "godina": None,
            "opstina_sifra": None,
            "opstina_apr": None,
            "poslovna_imovina": None,
            "kapital": None,
            "gubitak": None,
            "ukupni_prihodi": None,
            "neto_dobitak": None,
            "neto_gubitak": None,
            "prosecan_broj_zaposlenih": None,
            "found_in_api": False
        })
    else:
        rows.append({
            "maticni_broj": mb,
            "apr_naziv": item.get("PoslovnoIme"),
            "godina": item.get("GodinaFi"),
            "opstina_sifra": item.get("SifraOpstine"),
            "opstina_apr": item.get("NazivOpstine"),
            "poslovna_imovina": item.get("PoslovnaImovina"),
            "kapital": item.get("Kapital"),
            "gubitak": item.get("Gubitak"),
            "ukupni_prihodi": item.get("UkupniPrihodi"),
            "neto_dobitak": item.get("NetoDobitak"),
            "neto_gubitak": item.get("NetoGubitak"),
            "prosecan_broj_zaposlenih": item.get("ProsecanBrojZaposlenih"),
            "found_in_api": True
        })

df_fin = pd.DataFrame(rows)

# 3) Spoji sa originalnim podacima
df_final = df_input.merge(df_fin, on="maticni_broj", how="left")

# 4) Snimi rezultat
df_final.to_excel(output_path, index=False)

print(f"Saved: {output_path}")
print(f"Ukupno redova: {len(df_input)}")
print(f"Nađeno u API: {df_final['found_in_api'].fillna(False).sum()}")
print(f"Nije nađeno: {(~df_final['found_in_api'].fillna(False)).sum()}")