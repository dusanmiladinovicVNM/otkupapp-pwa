import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://openapi.apr.gov.rs/api/opendata/companies"
output_file = "Financial dataset - Elderly care.xlsx"
target_codes = {"1039", "4631"}

print("Downloading APR dataset...")
response = requests.get(url, timeout=120, verify=False)
response.raise_for_status()
data = response.json()

companies = []

for maticni_broj, company in data["Podaci"].items():
    sifra = company.get("SifraDelatnosti")

    if sifra in target_codes:
        companies.append({
            "maticni_broj": maticni_broj,
            "naziv": company.get("PoslovnoIme"),
            "sifra_delatnosti": sifra,
            "opstina": company.get("NazivOpstine"),
            "status": company.get("NazivStatus"),
            "datum_osnivanja": company.get("DatumOsnivanja"),
        })

df = pd.DataFrame(companies)
df.to_excel(output_file, index=False)

print(f"Saved: {output_file}")
print(f"Companies found: {len(df)}")