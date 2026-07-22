from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

import pandas as pd

from apr_pipeline_common import (
    CLEAN_DIR,
    PROCESSED_DIR,
    ensure_directories,
    latest_file,
    normalize_activity_code,
    normalize_company_id,
    numeric_series,
    sha256_file,
    utc_timestamp,
    write_json,
)

NUMERIC_COLUMNS = (
    "godina",
    "poslovna_imovina",
    "kapital",
    "gubitak",
    "ukupni_prihodi",
    "neto_dobitak",
    "neto_gubitak",
    "prosecan_broj_zaposlenih",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Čisti i validira APR tržišni dataset.")
    parser.add_argument("--input", type=Path, help="Obogaćeni Excel iz Clean Data foldera.")
    return parser.parse_args()


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip()
    return " ".join(text.split())


def status_category(value: object) -> str:
    text = normalize_text(value).casefold()
    if not text:
        return "unknown"
    if "stečaj" in text or "stecaj" in text:
        return "bankruptcy"
    if "likvid" in text:
        return "liquidation"
    if "neaktivan" in text or "brisan" in text or "ugašen" in text or "ugasen" in text:
        return "inactive"
    if "aktivan" in text:
        return "active"
    return "other"


def main() -> None:
    args = parse_args()
    ensure_directories()

    input_path = args.input or latest_file(CLEAN_DIR, "apr_companies_financials_*.xlsx")
    if not input_path.is_absolute():
        input_path = input_path.resolve()

    df = pd.read_excel(input_path, dtype={"maticni_broj": "string"})
    required = {"maticni_broj", "naziv", "sifra_delatnosti", "status"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"Nedostaju obavezne kolone: {', '.join(missing)}")

    df["maticni_broj"] = df["maticni_broj"].map(normalize_company_id)
    df["naziv"] = df["naziv"].map(normalize_text)
    df["opstina"] = df.get("opstina", pd.Series(index=df.index, dtype="object")).map(normalize_text)
    df["opstina_apr"] = df.get("opstina_apr", pd.Series(index=df.index, dtype="object")).map(normalize_text)
    df["status"] = df["status"].map(normalize_text)
    df["sifra_delatnosti"] = df["sifra_delatnosti"].map(normalize_activity_code)

    for column in NUMERIC_COLUMNS:
        if column in df.columns:
            df[column] = numeric_series(df[column])

    df["status_category"] = df["status"].map(status_category)
    df["is_active_market_candidate"] = df["status_category"].eq("active")
    df["valid_company_id"] = df["maticni_broj"].notna()
    df["valid_activity_code"] = df["sifra_delatnosti"].str.fullmatch(r"\d{4}", na=False)
    df["valid_revenue"] = df.get("ukupni_prihodi", pd.Series(index=df.index, dtype="float64")).ge(0)
    df["valid_employees"] = df.get(
        "prosecan_broj_zaposlenih", pd.Series(index=df.index, dtype="float64")
    ).ge(0)

    duplicate_mask = df["maticni_broj"].notna() & df["maticni_broj"].duplicated(keep=False)
    df["duplicate_company_id"] = duplicate_mask
    df["data_quality_issue"] = (
        ~df["valid_company_id"]
        | ~df["valid_activity_code"]
        | df["duplicate_company_id"]
        | df["naziv"].eq("")
    )

    has_financials = (
        df["has_financials"].fillna(False).astype(bool)
        if "has_financials" in df.columns
        else df.get("ukupni_prihodi", pd.Series(index=df.index, dtype="float64")).notna()
    )

    quality_rows = [
        ("rows_total", len(df)),
        ("valid_company_id", int(df["valid_company_id"].sum())),
        ("invalid_company_id", int((~df["valid_company_id"]).sum())),
        ("duplicate_company_rows", int(df["duplicate_company_id"].sum())),
        ("active_market_candidates", int(df["is_active_market_candidate"].sum())),
        ("unknown_status", int(df["status_category"].eq("unknown").sum())),
        ("other_status", int(df["status_category"].eq("other").sum())),
        ("rows_with_financials", int(has_financials.sum())),
        ("rows_with_valid_revenue", int(df["valid_revenue"].sum())),
        ("rows_with_valid_employees", int(df["valid_employees"].sum())),
        ("rows_with_quality_issue", int(df["data_quality_issue"].sum())),
    ]
    quality_df = pd.DataFrame(quality_rows, columns=["metric", "value"])

    output_path = PROCESSED_DIR / input_path.name.replace("apr_companies_financials_", "apr_market_validated_")
    quality_path = PROCESSED_DIR / output_path.name.replace("apr_market_validated_", "apr_data_quality_")
    metadata_path = output_path.with_suffix(".metadata.json")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Validated Data", index=False)
        df[df["data_quality_issue"]].to_excel(writer, sheet_name="Quality Issues", index=False)
        quality_df.to_excel(writer, sheet_name="Quality Summary", index=False)

    quality_df.to_excel(quality_path, index=False)
    write_json(
        metadata_path,
        {
            "dataset": "Validated APR market dataset",
            "generated_at_utc": utc_timestamp(),
            "input_file": input_path.name,
            "input_sha256": sha256_file(input_path),
            "output_file": output_path.name,
            "output_sha256": sha256_file(output_path),
            "quality_metrics": dict(quality_rows),
            "status_rule": (
                "Aktivna tržišna firma se trenutno prepoznaje tekstualnim statusom koji sadrži 'aktivan', "
                "nakon prethodne provere stečaja, likvidacije, neaktivnog i brisanog statusa. "
                "Pravilo treba proveriti prema svim statusima prisutnim u datasetu."
            ),
        },
    )

    print(f"Sačuvano: {output_path}")
    print(f"Aktivni kandidati: {int(df['is_active_market_candidate'].sum())}")
    print(f"Redovi sa problemom kvaliteta: {int(df['data_quality_issue'].sum())}")


if __name__ == "__main__":
    main()
