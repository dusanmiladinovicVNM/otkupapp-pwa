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

# Negativna stanja se proveravaju pre pozitivnih da, na primer,
# "neaktivan" ne bi bio pogrešno prepoznat kao "aktivan".
BANKRUPTCY_TOKENS = ("stečaj", "stecaj")
LIQUIDATION_TOKENS = ("likvid",)
INACTIVE_TOKENS = (
    "neaktivan",
    "neaktivno",
    "brisan",
    "obrisan",
    "ugašen",
    "ugasen",
    "prestao",
    "prestanak",
)
ACTIVE_TOKENS = (
    "aktivan",
    "aktivno",
    "registrovan",
    "registrovano",
    "registrovana",
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


def contains_any(text: str, tokens: tuple[str, ...]) -> bool:
    return any(token in text for token in tokens)


def status_category(value: object) -> str:
    text = normalize_text(value).casefold()
    if not text:
        return "unknown"
    if contains_any(text, BANKRUPTCY_TOKENS):
        return "bankruptcy"
    if contains_any(text, LIQUIDATION_TOKENS):
        return "liquidation"
    if contains_any(text, INACTIVE_TOKENS):
        return "inactive"
    if contains_any(text, ACTIVE_TOKENS):
        return "active"
    return "other"


def format_top_statuses(status_profile: pd.DataFrame, limit: int = 10) -> str:
    rows = status_profile.head(limit)
    return "; ".join(
        f"{row.status!r}={int(row.companies)} ({row.status_category})"
        for row in rows.itertuples(index=False)
    )


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
    df["unclassified_status"] = df["status_category"].isin(["unknown", "other"])
    df["data_quality_issue"] = (
        ~df["valid_company_id"]
        | ~df["valid_activity_code"]
        | df["duplicate_company_id"]
        | df["naziv"].eq("")
        | df["unclassified_status"]
    )

    has_financials = (
        df["has_financials"].fillna(False).astype(bool)
        if "has_financials" in df.columns
        else df.get("ukupni_prihodi", pd.Series(index=df.index, dtype="float64")).notna()
    )

    status_profile = (
        df.groupby(["status", "status_category"], dropna=False)
        .agg(companies=("maticni_broj", "size"))
        .reset_index()
        .sort_values(["companies", "status"], ascending=[False, True])
    )

    category_profile = (
        df.groupby("status_category", dropna=False)
        .agg(companies=("maticni_broj", "size"))
        .reset_index()
        .sort_values("companies", ascending=False)
    )

    active_count = int(df["is_active_market_candidate"].sum())
    quality_rows = [
        ("rows_total", len(df)),
        ("valid_company_id", int(df["valid_company_id"].sum())),
        ("invalid_company_id", int((~df["valid_company_id"]).sum())),
        ("duplicate_company_rows", int(df["duplicate_company_id"].sum())),
        ("active_market_candidates", active_count),
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
        status_profile.to_excel(writer, sheet_name="Status Values", index=False)
        category_profile.to_excel(writer, sheet_name="Status Categories", index=False)

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
            "status_rules": {
                "active_tokens": list(ACTIVE_TOKENS),
                "inactive_tokens": list(INACTIVE_TOKENS),
                "liquidation_tokens": list(LIQUIDATION_TOKENS),
                "bankruptcy_tokens": list(BANKRUPTCY_TOKENS),
                "evaluation_order": ["bankruptcy", "liquidation", "inactive", "active", "other"],
            },
            "top_status_values": status_profile.head(20).to_dict(orient="records"),
        },
    )

    print(f"Sačuvano: {output_path}")
    print(f"Aktivni kandidati: {active_count}")
    print(f"Redovi sa problemom kvaliteta: {int(df['data_quality_issue'].sum())}")
    print("Status kategorije:")
    for row in category_profile.itertuples(index=False):
        print(f"  - {row.status_category}: {int(row.companies)}")

    if active_count == 0 and len(df) > 0:
        raise RuntimeError(
            "Status klasifikacija je vratila 0 aktivnih firmi, pa analiza nije bezbedna. "
            f"Najčešće vrednosti: {format_top_statuses(status_profile)}. "
            f"Detalji su sačuvani u sheet-u 'Status Values' fajla {output_path.name}."
        )


if __name__ == "__main__":
    main()
