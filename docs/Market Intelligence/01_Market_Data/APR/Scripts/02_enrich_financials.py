from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from apr_pipeline_common import (
    CLEAN_DIR,
    FINANCIALS_URL,
    RAW_DIR,
    ensure_directories,
    latest_file,
    normalize_company_id,
    numeric_series,
    request_json,
    sha256_file,
    utc_timestamp,
    write_json,
)

FINANCIAL_COLUMNS = (
    "poslovna_imovina",
    "kapital",
    "gubitak",
    "ukupni_prihodi",
    "neto_dobitak",
    "neto_gubitak",
    "prosecan_broj_zaposlenih",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Dodaje APR finansijske podatke firmama.")
    parser.add_argument("--input", type=Path, help="Ulazni Excel iz Raw Data foldera.")
    parser.add_argument("--timeout", type=int, default=120)
    parser.add_argument("--insecure", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    ensure_directories()

    input_path = args.input or latest_file(RAW_DIR, "apr_companies_*.xlsx")
    if not input_path.is_absolute():
        input_path = input_path.resolve()

    df_input = pd.read_excel(input_path, dtype={"maticni_broj": "string"})
    if "maticni_broj" not in df_input.columns:
        raise ValueError("Ulazni fajl nema kolonu 'maticni_broj'.")

    df_input["maticni_broj"] = df_input["maticni_broj"].map(normalize_company_id)
    invalid_rows = int(df_input["maticni_broj"].isna().sum())
    if invalid_rows:
        print(f"Upozorenje: {invalid_rows} redova ima nevalidan matični broj.")

    duplicate_count = int(df_input["maticni_broj"].dropna().duplicated().sum())
    unique_ids = df_input["maticni_broj"].dropna().drop_duplicates().tolist()

    print("Preuzimanje APR finansijskih podataka...")
    payload = request_json(
        FINANCIALS_URL,
        timeout=args.timeout,
        verify_tls=not args.insecure,
    )
    financial_data = payload["Podaci"]

    rows: list[dict[str, object]] = []
    for company_id in unique_ids:
        item = financial_data.get(company_id)
        if not isinstance(item, dict):
            rows.append(
                {
                    "maticni_broj": company_id,
                    "apr_naziv": None,
                    "godina": None,
                    "opstina_sifra": None,
                    "opstina_apr": None,
                    **{column: None for column in FINANCIAL_COLUMNS},
                    "found_in_financial_api": False,
                }
            )
            continue

        rows.append(
            {
                "maticni_broj": company_id,
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
                "found_in_financial_api": True,
            }
        )

    df_fin = pd.DataFrame(rows)
    if df_fin["maticni_broj"].duplicated().any():
        raise RuntimeError("Finansijski rezultat sadrži više redova po matičnom broju.")

    df_final = df_input.merge(
        df_fin,
        on="maticni_broj",
        how="left",
        validate="many_to_one",
    )

    for column in ("godina",) + FINANCIAL_COLUMNS:
        df_final[column] = numeric_series(df_final[column])

    df_final["has_financials"] = (
        df_final["found_in_financial_api"].fillna(False)
        & df_final["ukupni_prihodi"].notna()
        & df_final["godina"].notna()
    )

    output_path = CLEAN_DIR / input_path.name.replace("apr_companies_", "apr_companies_financials_")
    metadata_path = output_path.with_suffix(".metadata.json")
    df_final.to_excel(output_path, index=False)

    metadata = {
        "dataset": "APR companies enriched with latest financial statement exposed by endpoint",
        "source_url": FINANCIALS_URL,
        "generated_at_utc": utc_timestamp(),
        "input_file": input_path.name,
        "input_sha256": sha256_file(input_path),
        "output_file": output_path.name,
        "output_sha256": sha256_file(output_path),
        "input_rows": len(df_input),
        "unique_company_ids": len(unique_ids),
        "invalid_company_ids": invalid_rows,
        "duplicate_company_ids_in_input": duplicate_count,
        "found_in_financial_api": int(df_final["found_in_financial_api"].fillna(False).sum()),
        "rows_with_usable_financials": int(df_final["has_financials"].sum()),
        "tls_verification": not args.insecure,
        "important_limitation": (
            "Endpoint se trenutno tretira kao jedan finansijski zapis po firmi. "
            "Ovaj korak ne dokazuje višegodišnju istoriju dok se API struktura posebno ne potvrdi."
        ),
    }
    write_json(metadata_path, metadata)

    print(f"Sačuvano: {output_path}")
    print(f"Ukupno redova: {len(df_final)}")
    print(f"Upotrebljive finansije: {int(df_final['has_financials'].sum())}")


if __name__ == "__main__":
    main()
