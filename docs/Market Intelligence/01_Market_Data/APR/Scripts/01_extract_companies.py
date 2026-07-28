from __future__ import annotations

import argparse
import re
from typing import Any

import pandas as pd

from apr_pipeline_common import (
    COMPANIES_URL,
    DEFAULT_ACTIVITY_CODES,
    RAW_DIR,
    date_stamp,
    ensure_directories,
    normalize_activity_code,
    normalize_company_id,
    normalize_pib,
    request_json,
    sha256_file,
    utc_timestamp,
    write_json,
)

# APR Open Data šema se može menjati bez promene URL-a. Kandidati pokrivaju
# dosad korišćene i očekivane nazive polja, a normalizovani fallback pronalazi
# i ekvivalentno polje koje u nazivu sadrži PIB ili poreski identifikacioni broj.
PIB_FIELD_CANDIDATES = (
    "PIB",
    "Pib",
    "pib",
    "PoreskiIdentifikacioniBroj",
    "PoreskiBroj",
    "TaxIdentificationNumber",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Izvlači APR firme po šiframa delatnosti.")
    parser.add_argument("--codes", nargs="+", default=list(DEFAULT_ACTIVITY_CODES))
    parser.add_argument("--timeout", type=int, default=120)
    tls_group = parser.add_mutually_exclusive_group()
    tls_group.add_argument(
        "--insecure",
        action="store_true",
        help="Odmah isključuje TLS proveru. Koristiti samo za dijagnostiku.",
    )
    tls_group.add_argument(
        "--strict-tls",
        action="store_true",
        help="Zabranjuje automatski APR-only fallback ako sertifikat ne može da se validira.",
    )
    return parser.parse_args()


def _normalized_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def extract_pib(company: dict[str, Any]) -> tuple[str | None, str | None]:
    """Vraća normalizovan PIB i naziv izvornog APR polja.

    Prvo proverava eksplicitne kandidate, zatim radi oprezan fallback po nazivu
    ključa. Ne pokušava da izvede PIB iz MB-a jer između ta dva identifikatora
    ne postoji računska konverzija.
    """
    for field_name in PIB_FIELD_CANDIDATES:
        if field_name in company:
            return normalize_pib(company.get(field_name)), field_name

    for field_name, value in company.items():
        normalized = _normalized_key(field_name)
        if normalized == "pib" or "poreskiidentifikacionibroj" in normalized:
            return normalize_pib(value), str(field_name)

    return None, None


def main() -> None:
    args = parse_args()
    ensure_directories()

    target_codes = {normalize_activity_code(code) for code in args.codes}
    if not target_codes or "" in target_codes:
        raise ValueError("Mora postojati najmanje jedna validna šifra delatnosti.")

    print("Preuzimanje APR registra firmi...")
    payload, request_meta = request_json(
        COMPANIES_URL,
        timeout=args.timeout,
        force_insecure=args.insecure,
        strict_tls=args.strict_tls,
    )
    source_data = payload["Podaci"]

    rows: list[dict[str, object]] = []
    invalid_company_ids = 0
    valid_pibs = 0
    invalid_or_missing_pibs = 0
    pib_source_fields: dict[str, int] = {}

    for raw_company_id, company in source_data.items():
        if not isinstance(company, dict):
            continue

        activity_code = normalize_activity_code(company.get("SifraDelatnosti"))
        if activity_code not in target_codes:
            continue

        company_id = normalize_company_id(raw_company_id)
        if company_id is None:
            invalid_company_ids += 1

        pib, pib_source_field = extract_pib(company)
        if pib is None:
            invalid_or_missing_pibs += 1
        else:
            valid_pibs += 1
        if pib_source_field:
            pib_source_fields[pib_source_field] = pib_source_fields.get(pib_source_field, 0) + 1

        rows.append(
            {
                "maticni_broj": company_id,
                "pib": pib,
                "naziv": company.get("PoslovnoIme"),
                "sifra_delatnosti": activity_code,
                "opstina": company.get("NazivOpstine"),
                "status": company.get("NazivStatus"),
                "datum_osnivanja": company.get("DatumOsnivanja"),
            }
        )

    if not rows:
        raise RuntimeError("APR nije vratio nijednu firmu za izabrane šifre delatnosti.")

    df = pd.DataFrame(rows)
    # Identifikatori moraju ostati tekst i u Excel izlazu.
    df["maticni_broj"] = df["maticni_broj"].astype("string")
    df["pib"] = df["pib"].astype("string")
    df = df.sort_values(["sifra_delatnosti", "naziv"], na_position="last").reset_index(drop=True)

    stamp = date_stamp()
    codes_slug = "_".join(sorted(target_codes))
    output_path = RAW_DIR / f"apr_companies_{codes_slug}_{stamp}.xlsx"
    metadata_path = RAW_DIR / f"apr_companies_{codes_slug}_{stamp}.metadata.json"

    df.to_excel(output_path, index=False)

    metadata = {
        "dataset": "APR companies",
        "source_url": COMPANIES_URL,
        "extracted_at_utc": utc_timestamp(),
        "activity_codes": sorted(target_codes),
        "api_total_records": len(source_data),
        "selected_records": len(df),
        "invalid_company_ids": invalid_company_ids,
        "valid_pibs": valid_pibs,
        "invalid_or_missing_pibs": invalid_or_missing_pibs,
        "pib_coverage_ratio": round(valid_pibs / len(df), 6),
        "pib_source_fields": pib_source_fields,
        "request": request_meta,
        "output_file": output_path.name,
        "output_sha256": sha256_file(output_path),
    }
    write_json(metadata_path, metadata)

    print(f"Sačuvano: {output_path}")
    print(f"Pronađeno firmi: {len(df)}")
    print(f"Nevalidni matični brojevi: {invalid_company_ids}")
    print(f"Validni PIB-ovi: {valid_pibs}")
    print(f"Nedostajući ili nevalidni PIB-ovi: {invalid_or_missing_pibs}")
    if valid_pibs == 0:
        print(
            "NAPOMENA: APR companies endpoint u ovom izvršavanju nije izložio PIB. "
            "Kolona je ipak kreirana radi stabilne izlazne šeme; potreban je dodatni "
            "autorizovani APR izvor ili drugi dozvoljeni registar za MB→PIB mapiranje."
        )
    if request_meta["tls_fallback_used"]:
        print("NAPOMENA: TLS fallback je korišćen i zabeležen u metadata JSON-u.")


if __name__ == "__main__":
    main()
