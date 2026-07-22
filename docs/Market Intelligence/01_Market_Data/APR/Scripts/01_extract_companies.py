from __future__ import annotations

import argparse

import pandas as pd

from apr_pipeline_common import (
    COMPANIES_URL,
    DEFAULT_ACTIVITY_CODES,
    RAW_DIR,
    date_stamp,
    ensure_directories,
    normalize_activity_code,
    normalize_company_id,
    request_json,
    sha256_file,
    utc_timestamp,
    write_json,
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

    for raw_company_id, company in source_data.items():
        if not isinstance(company, dict):
            continue

        activity_code = normalize_activity_code(company.get("SifraDelatnosti"))
        if activity_code not in target_codes:
            continue

        company_id = normalize_company_id(raw_company_id)
        if company_id is None:
            invalid_company_ids += 1

        rows.append(
            {
                "maticni_broj": company_id,
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
        "request": request_meta,
        "output_file": output_path.name,
        "output_sha256": sha256_file(output_path),
    }
    write_json(metadata_path, metadata)

    print(f"Sačuvano: {output_path}")
    print(f"Pronađeno firmi: {len(df)}")
    print(f"Nevalidni matični brojevi: {invalid_company_ids}")
    if request_meta["tls_fallback_used"]:
        print("NAPOMENA: TLS fallback je korišćen i zabeležen u metadata JSON-u.")


if __name__ == "__main__":
    main()
