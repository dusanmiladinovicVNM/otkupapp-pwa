from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path
from typing import Iterable

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
COMPETITION_DIR = SCRIPT_DIR.parent
MARKET_INTELLIGENCE_DIR = COMPETITION_DIR.parent
APR_DIR = MARKET_INTELLIGENCE_DIR / "01_Market_Data" / "APR"
APR_SCRIPTS_DIR = APR_DIR / "Scripts"

if str(APR_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(APR_SCRIPTS_DIR))

from apr_pipeline_common import (  # noqa: E402
    COMPANIES_URL,
    FINANCIALS_URL,
    normalize_activity_code,
    normalize_company_id,
    numeric_series,
    request_json,
    sha256_file,
    utc_timestamp,
    write_json,
)
from build_infosys_replacement_targets import (  # noqa: E402
    AMBIGUOUS_CORE_TOKENS,
    company_core_tokens,
    name_match_features,
    normalize_text,
)

DEFAULT_INPUT = COMPETITION_DIR / "infosys_replacement_targets.csv"
DEFAULT_OUTPUT_CSV = COMPETITION_DIR / "infosys_wide_enrichment.csv"
DEFAULT_OUTPUT_XLSX = COMPETITION_DIR / "infosys_wide_enrichment.xlsx"
DEFAULT_METADATA = COMPETITION_DIR / "infosys_wide_enrichment.metadata.json"
WIDE_MATCH_SCOPE = "APR all registered activity codes"
APR_MONETARY_MULTIPLIER = 1_000

MONETARY_FINANCIAL_FIELDS = {
    "PoslovnaImovina": "wide_poslovna_imovina_rsd",
    "Kapital": "wide_kapital_rsd",
    "Gubitak": "wide_gubitak_rsd",
    "UkupniPrihodi": "wide_prihod_rsd",
    "NetoDobitak": "wide_neto_dobitak_rsd",
    "NetoGubitak": "wide_neto_gubitak_rsd",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Povezuje Infosys replacement reference sa celim APR registrom, "
            "bez ograničenja na šifre 1039/4631, i dodaje finansije."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--output-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX)
    parser.add_argument("--metadata", type=Path, default=DEFAULT_METADATA)
    parser.add_argument("--timeout", type=int, default=180)
    tls_group = parser.add_mutually_exclusive_group()
    tls_group.add_argument("--insecure", action="store_true")
    tls_group.add_argument("--strict-tls", action="store_true")
    return parser.parse_args()


def registry_dataframe(payload: dict[str, object]) -> pd.DataFrame:
    source_data = payload.get("Podaci")
    if not isinstance(source_data, dict):
        raise ValueError("APR companies odgovor nema očekivani objekat Podaci.")

    rows: list[dict[str, object]] = []
    for raw_company_id, raw_company in source_data.items():
        if not isinstance(raw_company, dict):
            continue
        company_id = normalize_company_id(raw_company_id)
        if company_id is None:
            continue
        rows.append(
            {
                "maticni_broj": company_id,
                "naziv": raw_company.get("PoslovnoIme"),
                "sifra_delatnosti": normalize_activity_code(
                    raw_company.get("SifraDelatnosti")
                ),
                "opstina": raw_company.get("NazivOpstine"),
                "opstina_apr": raw_company.get("NazivOpstine"),
                "status": raw_company.get("NazivStatus"),
                "datum_osnivanja": raw_company.get("DatumOsnivanja"),
            }
        )

    registry = pd.DataFrame(rows)
    if registry.empty:
        raise RuntimeError("APR companies endpoint nije vratio nijednu validnu firmu.")
    if registry["maticni_broj"].duplicated().any():
        registry = registry.drop_duplicates("maticni_broj", keep="first").copy()
    return registry.reset_index(drop=True)


def core_tokens_for_registry_row(row: pd.Series) -> list[str]:
    return company_core_tokens(row.get("naziv"), [row.get("opstina")])


def build_token_index(registry: pd.DataFrame) -> tuple[dict[str, set[int]], dict[str, set[int]]]:
    token_index: dict[str, set[int]] = defaultdict(set)
    location_index: dict[str, set[int]] = defaultdict(set)

    for index, row in registry.iterrows():
        for token in core_tokens_for_registry_row(row):
            if len(token) >= 3:
                token_index[token].add(index)
        normalized_location = normalize_text(row.get("opstina"))
        if normalized_location:
            location_index[normalized_location].add(index)

    return dict(token_index), dict(location_index)


def reference_core_tokens(reference: pd.Series) -> list[str]:
    return company_core_tokens(
        reference.get("reference_name"), [reference.get("location")]
    )


def candidate_indices(
    reference: pd.Series,
    registry: pd.DataFrame,
    token_index: dict[str, set[int]],
    location_index: dict[str, set[int]],
) -> set[int]:
    tokens = reference_core_tokens(reference)
    distinctive = {
        token
        for token in tokens
        if token not in AMBIGUOUS_CORE_TOKENS and len(token) >= 4
    }
    fallback_tokens = {token for token in tokens if len(token) >= 4}

    indices: set[int] = set()
    for token in distinctive or fallback_tokens:
        indices.update(token_index.get(token, set()))

    location = normalize_text(reference.get("location"))
    location_candidates = location_index.get(location, set()) if location else set()

    # Lokacija proširuje kandidatni skup, ali ga ne zamenjuje kada postoje jaki name tokeni.
    if indices and location_candidates:
        indices.update(location_candidates)
    elif not indices and location_candidates:
        indices = set(location_candidates)

    prior_candidate_id = normalize_company_id(reference.get("candidate_apr_maticni_broj"))
    if prior_candidate_id:
        prior_matches = registry.index[registry["maticni_broj"].eq(prior_candidate_id)]
        indices.update(int(index) for index in prior_matches)

    if not indices:
        # Retki kratki ili generički nazivi: ograničena fallback pretraga umesto tihog odustajanja.
        if fallback_tokens:
            mask = registry["naziv"].astype("string").map(normalize_text).map(
                lambda value: any(token in value.split() for token in fallback_tokens)
            )
            indices.update(int(index) for index in registry.index[mask])

    return indices


def best_wide_match(
    reference: pd.Series,
    registry: pd.DataFrame,
    token_index: dict[str, set[int]],
    location_index: dict[str, set[int]],
) -> tuple[pd.Series | None, float, float, dict[str, object], int]:
    indices = candidate_indices(reference, registry, token_index, location_index)
    scored: list[tuple[float, int, dict[str, object]]] = []

    for index in indices:
        registry_row = registry.loc[index]
        features = name_match_features(reference, registry_row)
        scored.append((float(features["score"]), index, features))

    if not scored:
        return None, 0.0, 0.0, {}, 0

    scored.sort(key=lambda item: item[0], reverse=True)
    best_score, best_index, best_features = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    return (
        registry.loc[best_index],
        best_score,
        best_score - second_score,
        best_features,
        len(scored),
    )


def wide_match_method(
    score: float,
    margin: float,
    features: dict[str, object],
    reference_location: object,
) -> tuple[str, str]:
    exact_core = bool(features.get("exact_core"))
    subset_core = bool(features.get("subset_core"))
    distinctive_shared = int(features.get("distinctive_shared", 0))
    location_score = float(features.get("location_score", 0.0))
    has_reference_location = bool(normalize_text(reference_location))

    if exact_core and (location_score >= 0.7 or not has_reference_location):
        return "matched", "wide_exact_core_name_location"
    if subset_core and distinctive_shared >= 1 and location_score >= 0.7 and margin >= 0.04:
        return "matched", "wide_core_subset_location"
    if score >= 0.92 and margin >= 0.08 and (location_score >= 0.7 or not has_reference_location):
        return "matched", "wide_high_confidence_name_location"

    if exact_core and margin >= 0.15:
        return "manual_review", "wide_exact_core_location_conflict"
    if subset_core and distinctive_shared >= 1 and margin >= 0.16:
        return "manual_review", "wide_core_subset_unique_candidate"
    if score >= 0.76 and margin >= 0.05:
        return "manual_review", "wide_fuzzy_name_location"
    return "unmatched", "no_safe_match_in_full_apr_registry"


def financial_fields(item: object) -> dict[str, object]:
    if not isinstance(item, dict):
        return {
            "wide_financial_year": pd.NA,
            "wide_prihod_rsd": pd.NA,
            "wide_employees": pd.NA,
            "wide_found_financials": False,
        }

    result: dict[str, object] = {
        "wide_financial_year": item.get("GodinaFi"),
        "wide_employees": item.get("ProsecanBrojZaposlenih"),
        "wide_found_financials": True,
    }
    for source, target in MONETARY_FINANCIAL_FIELDS.items():
        value = pd.to_numeric(pd.Series([item.get(source)]), errors="coerce").iloc[0]
        result[target] = value * APR_MONETARY_MULTIPLIER if pd.notna(value) else pd.NA
    return result


def copy_registry_fields(row: dict[str, object], registry_row: pd.Series | None) -> None:
    if registry_row is None:
        return
    row.update(
        {
            "wide_maticni_broj": registry_row.get("maticni_broj"),
            "wide_naziv": registry_row.get("naziv"),
            "wide_sifra_delatnosti": registry_row.get("sifra_delatnosti"),
            "wide_status": registry_row.get("status"),
            "wide_opstina": registry_row.get("opstina"),
            "wide_datum_osnivanja": registry_row.get("datum_osnivanja"),
        }
    )


def preserved_match_registry_row(
    reference: pd.Series, registry_by_id: pd.DataFrame
) -> pd.Series | None:
    company_id = normalize_company_id(reference.get("apr_maticni_broj"))
    if not company_id or company_id not in registry_by_id.index:
        return None
    result = registry_by_id.loc[company_id]
    return result.iloc[0] if isinstance(result, pd.DataFrame) else result


def main() -> None:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Nije pronađen ulazni fajl: {input_path}")

    references = pd.read_csv(input_path, dtype="string")
    required_columns = {"reference_name", "apr_match_status"}
    missing = required_columns - set(references.columns)
    if missing:
        raise ValueError(f"Ulaz nema obavezne kolone: {sorted(missing)}")

    print("Preuzimanje kompletnog APR registra...")
    companies_payload, companies_request = request_json(
        COMPANIES_URL,
        timeout=args.timeout,
        force_insecure=args.insecure,
        strict_tls=args.strict_tls,
    )
    registry = registry_dataframe(companies_payload)
    registry_by_id = registry.set_index("maticni_broj", drop=False)
    token_index, location_index = build_token_index(registry)

    output_rows: list[dict[str, object]] = []
    for _, reference in references.iterrows():
        row = reference.to_dict()
        row["wide_match_scope"] = WIDE_MATCH_SCOPE

        registry_row = preserved_match_registry_row(reference, registry_by_id)
        if registry_row is not None and reference.get("apr_match_status") == "matched":
            status = "matched"
            method = "preserved_confirmed_narrow_match"
            score = 1.0
            margin = 1.0
            features = {
                "exact_core": True,
                "subset_core": True,
                "distinctive_shared": 1,
                "location_score": 1.0,
                "reference_core": "",
                "apr_core": "",
            }
            candidate_count = 1
        else:
            registry_row, score, margin, features, candidate_count = best_wide_match(
                reference, registry, token_index, location_index
            )
            status, method = wide_match_method(
                score, margin, features, reference.get("location")
            )

        row.update(
            {
                "wide_match_status": status,
                "wide_match_method": method,
                "wide_match_score": round(score, 4),
                "wide_match_margin": round(margin, 4),
                "wide_candidate_count": candidate_count,
                "wide_core_name_exact": bool(features.get("exact_core")),
                "wide_core_name_subset": bool(features.get("subset_core")),
                "wide_distinctive_shared_tokens": int(
                    features.get("distinctive_shared", 0)
                ),
                "wide_location_match_score": float(
                    features.get("location_score", 0.0)
                ),
                "wide_reference_core_name": features.get("reference_core", ""),
                "wide_candidate_core_name": features.get("apr_core", ""),
                "wide_unmatched_reason": (
                    "No safe identity match in the complete APR companies registry."
                    if status == "unmatched"
                    else pd.NA
                ),
            }
        )
        copy_registry_fields(row, registry_row)
        output_rows.append(row)

    output = pd.DataFrame(output_rows)

    print("Preuzimanje APR finansijskih podataka...")
    financial_payload, financial_request = request_json(
        FINANCIALS_URL,
        timeout=args.timeout,
        force_insecure=args.insecure,
        strict_tls=args.strict_tls,
    )
    financial_data = financial_payload.get("Podaci")
    if not isinstance(financial_data, dict):
        raise ValueError("APR financials odgovor nema očekivani objekat Podaci.")

    for index, output_row in output.iterrows():
        company_id = normalize_company_id(output_row.get("wide_maticni_broj"))
        fields = financial_fields(financial_data.get(company_id) if company_id else None)
        for field, value in fields.items():
            output.at[index, field] = value

    for column in (
        "wide_financial_year",
        "wide_employees",
        *MONETARY_FINANCIAL_FIELDS.values(),
    ):
        if column in output.columns:
            output[column] = numeric_series(output[column])

    status_order = pd.CategoricalDtype(
        categories=["matched", "manual_review", "unmatched"], ordered=True
    )
    output["_status_order"] = output["wide_match_status"].astype(status_order)
    output = output.sort_values(
        ["priority_tier", "_status_order", "wide_prihod_rsd", "reference_name"],
        ascending=[True, True, False, True],
        na_position="last",
    ).drop(columns=["_status_order"])

    matched_count = int(output["wide_match_status"].eq("matched").sum())
    manual_count = int(output["wide_match_status"].eq("manual_review").sum())
    unmatched_count = int(output["wide_match_status"].eq("unmatched").sum())
    unique_matched_ids = int(
        output.loc[output["wide_match_status"].eq("matched"), "wide_maticni_broj"]
        .dropna()
        .nunique()
    )
    reference_count = len(output)

    summary = pd.DataFrame(
        [
            ("reference_rows", reference_count),
            ("registry_records", len(registry)),
            ("matched_rows", matched_count),
            ("manual_review_rows", manual_count),
            ("unmatched_rows", unmatched_count),
            ("unique_matched_company_ids", unique_matched_ids),
            (
                "safe_match_rate_pct",
                round(100 * matched_count / reference_count, 1)
                if reference_count
                else 0.0,
            ),
            (
                "coverage_with_manual_pct",
                round(100 * (matched_count + manual_count) / reference_count, 1)
                if reference_count
                else 0.0,
            ),
        ],
        columns=["metric", "value"],
    )

    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(args.output_csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(args.output_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        output.to_excel(writer, sheet_name="All Targets", index=False)
        output[output["wide_match_status"].eq("matched")].to_excel(
            writer, sheet_name="Matched", index=False
        )
        output[output["wide_match_status"].eq("manual_review")].to_excel(
            writer, sheet_name="Manual Review", index=False
        )
        output[output["wide_match_status"].eq("unmatched")].to_excel(
            writer, sheet_name="Unmatched", index=False
        )

    metadata = {
        "dataset": "Infosys references enriched against complete APR company registry",
        "generated_at_utc": utc_timestamp(),
        "input_file": input_path.name,
        "input_sha256": sha256_file(input_path),
        "output_csv": args.output_csv.name,
        "output_csv_sha256": sha256_file(args.output_csv),
        "output_xlsx": args.output_xlsx.name,
        "output_xlsx_sha256": sha256_file(args.output_xlsx),
        "match_scope": WIDE_MATCH_SCOPE,
        "reference_rows": reference_count,
        "registry_records": len(registry),
        "matched_rows": matched_count,
        "manual_review_rows": manual_count,
        "unmatched_rows": unmatched_count,
        "unique_matched_company_ids": unique_matched_ids,
        "financial_source_unit": "thousand RSD",
        "financial_normalized_unit": "RSD",
        "financial_multiplier": APR_MONETARY_MULTIPLIER,
        "companies_request": companies_request,
        "financials_request": financial_request,
        "limitations": [
            "Public competitor reference does not prove an active current installation.",
            "Name matching never replaces company-ID confirmation for external claims.",
            "Manual-review rows require human verification before CRM activation.",
            "The financial endpoint is treated as one exposed financial record per company.",
        ],
    }
    write_json(args.metadata, metadata)

    print(f"CSV: {args.output_csv.resolve()}")
    print(f"Excel: {args.output_xlsx.resolve()}")
    print(f"Metadata: {args.metadata.resolve()}")
    print(f"APR registry records: {len(registry)}")
    for summary_row in summary.itertuples(index=False):
        print(f"  - {summary_row.metric}: {summary_row.value}")


if __name__ == "__main__":
    main()
