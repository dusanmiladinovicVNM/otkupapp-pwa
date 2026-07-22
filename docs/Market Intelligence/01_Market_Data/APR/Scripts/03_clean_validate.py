from __future__ import annotations

import argparse
import json
import unicodedata
from pathlib import Path

import pandas as pd

from apr_pipeline_common import (
    CLEAN_DIR,
    FINANCIALS_URL,
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

MONETARY_COLUMNS = (
    "poslovna_imovina",
    "kapital",
    "gubitak",
    "ukupni_prihodi",
    "neto_dobitak",
    "neto_gubitak",
)
NUMERIC_COLUMNS = ("godina",) + MONETARY_COLUMNS + ("prosecan_broj_zaposlenih",)
APR_MONETARY_MULTIPLIER = 1_000

# APR trenutno vraća statuse na srpskoj ćirilici. Za pouzdano poređenje
# svodimo srpsku ćirilicu i latinicu sa dijakriticima na isti ASCII oblik.
SERBIAN_CYRILLIC_TO_LATIN = str.maketrans(
    {
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "ђ": "dj",
        "е": "e",
        "ж": "z",
        "з": "z",
        "и": "i",
        "ј": "j",
        "к": "k",
        "л": "l",
        "љ": "lj",
        "м": "m",
        "н": "n",
        "њ": "nj",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "ћ": "c",
        "у": "u",
        "ф": "f",
        "х": "h",
        "ц": "c",
        "ч": "c",
        "џ": "dz",
        "ш": "s",
    }
)

# Negativna stanja se proveravaju pre pozitivnih da, na primer,
# "neaktivan" ne bi bio pogrešno prepoznat kao "aktivan".
BANKRUPTCY_TOKENS = ("stecaj",)
LIQUIDATION_TOKENS = ("likvid",)
INACTIVE_TOKENS = (
    "neaktivan",
    "neaktivno",
    "brisan",
    "obrisan",
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

STATUS_CLASSIFICATION_CASES = {
    "Активан": "active",
    "У ликвидацији": "liquidation",
    "У принудној ликвидацији": "liquidation",
    "У стечају": "bankruptcy",
    "Неактиван": "inactive",
    "Брисан": "inactive",
    "Aktivan": "active",
    "U likvidaciji": "liquidation",
    "U stečaju": "bankruptcy",
    "Registrovan": "active",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Čisti i validira APR tržišni dataset.")
    parser.add_argument("--input", type=Path, help="Obogaćeni Excel iz Clean Data foldera.")
    parser.add_argument(
        "--financial-unit",
        choices=("auto", "rsd", "thousand-rsd"),
        default="auto",
        help=(
            "Jedinica novčanih kolona u ulazu. 'auto' koristi metadata; "
            "za legacy APR fajlove prepoznaje hiljade RSD."
        ),
    )
    return parser.parse_args()


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip()
    return " ".join(text.split())


def fold_for_matching(value: object) -> str:
    """Normalizuje srpsku latinicu/ćirilicu u stabilan ASCII oblik za pravila."""
    text = normalize_text(value).casefold().translate(SERBIAN_CYRILLIC_TO_LATIN)
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def contains_any(text: str, tokens: tuple[str, ...]) -> bool:
    return any(token in text for token in tokens)


def status_category(value: object) -> str:
    text = fold_for_matching(value)
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


def validate_status_classifier() -> None:
    failures = []
    for value, expected in STATUS_CLASSIFICATION_CASES.items():
        actual = status_category(value)
        if actual != expected:
            failures.append(f"{value!r}: očekivano {expected}, dobijeno {actual}")
    if failures:
        raise RuntimeError("Status classifier regresija: " + "; ".join(failures))


def format_top_statuses(status_profile: pd.DataFrame, limit: int = 10) -> str:
    rows = status_profile.head(limit)
    return "; ".join(
        f"{row.status!r}={int(row.companies)} ({row.status_category})"
        for row in rows.itertuples(index=False)
    )


def read_input_metadata(input_path: Path) -> tuple[dict[str, object], Path]:
    metadata_path = input_path.with_suffix(".metadata.json")
    if not metadata_path.exists():
        return {}, metadata_path
    try:
        payload = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Ne mogu da pročitam metadata fajl {metadata_path}: {exc}") from exc
    return payload if isinstance(payload, dict) else {}, metadata_path


def determine_financial_multiplier(
    df: pd.DataFrame,
    input_path: Path,
    requested_unit: str,
) -> tuple[int, str, dict[str, object], Path]:
    input_metadata, input_metadata_path = read_input_metadata(input_path)

    if requested_unit == "rsd":
        return 1, "explicit --financial-unit rsd", input_metadata, input_metadata_path
    if requested_unit == "thousand-rsd":
        return APR_MONETARY_MULTIPLIER, "explicit --financial-unit thousand-rsd", input_metadata, input_metadata_path

    normalized_unit = normalize_text(input_metadata.get("normalized_financial_unit")).casefold()
    source_unit = normalize_text(input_metadata.get("source_financial_unit")).casefold()
    source_url = normalize_text(input_metadata.get("source_url"))

    if normalized_unit == "rsd":
        return 1, "metadata normalized_financial_unit=RSD", input_metadata, input_metadata_path
    if "thousand" in source_unit or "hiljad" in source_unit:
        return APR_MONETARY_MULTIPLIER, "metadata source_financial_unit=thousand RSD", input_metadata, input_metadata_path
    if source_url == FINANCIALS_URL:
        return (
            APR_MONETARY_MULTIPLIER,
            "legacy APR financial endpoint metadata without explicit unit",
            input_metadata,
            input_metadata_path,
        )

    # Novi obogaćeni fajl čuva izvorne *_apr_000_rsd kolone. Ako su glavne
    # vrednosti već 1.000 puta veće, ne primenjujemo multiplier ponovo.
    for column in MONETARY_COLUMNS:
        raw_column = f"{column}_apr_000_rsd"
        if column not in df.columns or raw_column not in df.columns:
            continue
        sample = df[[column, raw_column]].dropna()
        sample = sample[sample[raw_column] != 0].head(50)
        if sample.empty:
            continue
        ratio = (sample[column] / sample[raw_column]).median()
        if 999 <= ratio <= 1001:
            return 1, f"detected normalized RSD from {raw_column}", input_metadata, input_metadata_path
        if 0.999 <= ratio <= 1.001:
            return (
                APR_MONETARY_MULTIPLIER,
                f"detected unscaled thousand-RSD values from {raw_column}",
                input_metadata,
                input_metadata_path,
            )

    return 1, "unit metadata unavailable; conservatively assumed RSD", input_metadata, input_metadata_path


def main() -> None:
    args = parse_args()
    ensure_directories()
    validate_status_classifier()

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
    df["status_normalized"] = df["status"].map(fold_for_matching)
    df["sifra_delatnosti"] = df["sifra_delatnosti"].map(normalize_activity_code)

    for column in NUMERIC_COLUMNS:
        if column in df.columns:
            df[column] = numeric_series(df[column])
    for column in MONETARY_COLUMNS:
        raw_column = f"{column}_apr_000_rsd"
        if raw_column in df.columns:
            df[raw_column] = numeric_series(df[raw_column])

    multiplier, unit_reason, input_metadata, input_metadata_path = determine_financial_multiplier(
        df,
        input_path,
        args.financial_unit,
    )
    if multiplier == APR_MONETARY_MULTIPLIER:
        for column in MONETARY_COLUMNS:
            if column not in df.columns:
                continue
            raw_column = f"{column}_apr_000_rsd"
            if raw_column not in df.columns:
                df[raw_column] = df[column]
            df[column] = df[column] * multiplier

    df["financial_values_unit"] = "RSD"
    df["financial_multiplier_applied_in_validation"] = multiplier

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
        df.groupby(["status", "status_normalized", "status_category"], dropna=False)
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
        ("financial_multiplier_applied", multiplier),
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
            "input_metadata_file": input_metadata_path.name if input_metadata_path.exists() else None,
            "input_metadata": input_metadata,
            "output_file": output_path.name,
            "output_sha256": sha256_file(output_path),
            "source_financial_unit": "thousand RSD" if multiplier == APR_MONETARY_MULTIPLIER else input_metadata.get("source_financial_unit"),
            "normalized_financial_unit": "RSD",
            "financial_multiplier_applied_in_validation": multiplier,
            "financial_unit_resolution": unit_reason,
            "quality_metrics": dict(quality_rows),
            "status_matching_normalization": (
                "Unicode NFKC + casefold + Serbian Cyrillic-to-Latin transliteration + diacritic removal"
            ),
            "status_rules": {
                "active_tokens": list(ACTIVE_TOKENS),
                "inactive_tokens": list(INACTIVE_TOKENS),
                "liquidation_tokens": list(LIQUIDATION_TOKENS),
                "bankruptcy_tokens": list(BANKRUPTCY_TOKENS),
                "evaluation_order": ["bankruptcy", "liquidation", "inactive", "active", "other"],
            },
            "status_classifier_regression_cases": STATUS_CLASSIFICATION_CASES,
            "top_status_values": status_profile.head(20).to_dict(orient="records"),
        },
    )

    print(f"Sačuvano: {output_path}")
    print(f"Aktivni kandidati: {active_count}")
    print(f"Redovi sa problemom kvaliteta: {int(df['data_quality_issue'].sum())}")
    print(f"Finansijska jedinica: RSD; multiplier u validaciji: {multiplier} ({unit_reason})")
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
