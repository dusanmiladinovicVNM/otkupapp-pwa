from __future__ import annotations

import argparse
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
COMPETITION_DIR = SCRIPT_DIR.parent
APR_DIR = COMPETITION_DIR.parent / "01_Market_Data" / "APR"
DEFAULT_REFERENCES = COMPETITION_DIR / "infosys_agro_references.csv"
DEFAULT_ADDITIONS = COMPETITION_DIR / "infosys_manual_reference_additions.csv"
DEFAULT_APR = APR_DIR / "Processed" / "apr_market_validated_1039_4631_2026-07-22.xlsx"
DEFAULT_OUTPUT_CSV = COMPETITION_DIR / "infosys_replacement_targets.csv"
DEFAULT_OUTPUT_XLSX = COMPETITION_DIR / "infosys_replacement_targets.xlsx"
APR_MATCH_SCOPE = "APR activity codes 1039/4631 only"

LEGAL_TOKENS = {
    "ad",
    "doo",
    "d",
    "o",
    "od",
    "pr",
    "szr",
    "str",
    "stur",
    "pp",
    "pppup",
    "stkr",
    "stkrp",
    "preduzetnik",
    "drustvo",
    "ogranicenom",
    "odgovornoscu",
}

GENERIC_NAME_TOKENS = LEGAL_TOKENS | {
    "privredno",
    "trgovinsko",
    "preduzece",
    "preduezce",
    "firma",
    "kompanija",
    "company",
    "proizvodnja",
    "proizvodnje",
    "proizvodnju",
    "prerada",
    "prerade",
    "preradu",
    "skladistenje",
    "skladistenja",
    "promet",
    "prometa",
    "trgovina",
    "trgovine",
    "trgovinu",
    "usluge",
    "usluga",
    "poljoprivreda",
    "poljoprivredi",
    "poljoprivredne",
    "poljoprivrednih",
    "radnja",
    "za",
    "i",
    "u",
    "na",
    "sa",
}

CITY_TOKENS = {
    "arilje",
    "beograd",
    "cacak",
    "cajetina",
    "kosjeric",
    "kragujevac",
    "novi",
    "sad",
    "pozega",
    "sjenica",
    "smederevo",
    "uzice",
    "zlatibor",
}

# Reči koje same po sebi nisu dovoljne da potvrde identitet firme.
AMBIGUOUS_CORE_TOKENS = {
    "agro",
    "eko",
    "food",
    "fruit",
    "fruits",
    "frigo",
    "trade",
    "mlekara",
    "malina",
    "kooperativa",
}

APR_EXPORT_FIELDS = (
    ("maticni_broj", "apr_maticni_broj"),
    ("naziv", "apr_naziv"),
    ("sifra_delatnosti", "apr_sifra_delatnosti"),
    ("status", "apr_status"),
    ("godina", "apr_financial_year"),
    ("ukupni_prihodi", "apr_prihod_rsd"),
    ("prosecan_broj_zaposlenih", "apr_zaposleni"),
    ("opstina_apr", "apr_opstina"),
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Spaja Infosys agro reference sa validiranim APR tržišnim skupom."
    )
    parser.add_argument("--references", type=Path, default=DEFAULT_REFERENCES)
    parser.add_argument("--manual-additions", type=Path, default=DEFAULT_ADDITIONS)
    parser.add_argument("--apr", type=Path, default=DEFAULT_APR)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--output-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX)
    parser.add_argument(
        "--include-medium",
        action="store_true",
        help="Uključuje i reference sa srednjim AgriX potencijalom.",
    )
    return parser.parse_args()


def normalize_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value).casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("đ", "dj")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def tokens_from_values(values: Iterable[object]) -> set[str]:
    tokens: set[str] = set()
    for value in values:
        tokens.update(normalize_text(value).split())
    return tokens


def company_core_tokens(value: object, extra_stop_values: Iterable[object] = ()) -> list[str]:
    stop_tokens = GENERIC_NAME_TOKENS | CITY_TOKENS | tokens_from_values(extra_stop_values)
    return [
        token
        for token in normalize_text(value).split()
        if token not in stop_tokens and len(token) > 1
    ]


def normalized_company_name(value: object, extra_stop_values: Iterable[object] = ()) -> str:
    return " ".join(company_core_tokens(value, extra_stop_values))


def jaccard(left: set[str], right: set[str]) -> float:
    if not left or not right:
        return 0.0
    return len(left & right) / len(left | right)


def location_similarity(reference_location: object, apr_row: pd.Series) -> float:
    ref = normalize_text(reference_location)
    if not ref:
        return 0.0

    candidates = {
        normalize_text(apr_row.get("opstina")),
        normalize_text(apr_row.get("opstina_apr")),
    }
    candidates.discard("")

    if ref in candidates:
        return 1.0
    if any(ref in candidate or candidate in ref for candidate in candidates):
        return 0.8

    # APR naziv često sadrži naselje, dok dataset čuva samo opštinu.
    apr_name = normalize_text(apr_row.get("naziv"))
    if ref and apr_name and ref in apr_name:
        return 0.7

    ref_tokens = set(ref.split())
    if any(ref_tokens and ref_tokens <= set(candidate.split()) for candidate in candidates):
        return 0.7
    return 0.0


def name_match_features(reference: pd.Series, apr_row: pd.Series) -> dict[str, object]:
    reference_locations = [reference.get("location")]
    apr_locations = [apr_row.get("opstina"), apr_row.get("opstina_apr")]
    combined_stop_values = [*reference_locations, *apr_locations]

    ref_core_tokens = company_core_tokens(
        reference.get("reference_name"), combined_stop_values
    )
    apr_core_tokens = company_core_tokens(apr_row.get("naziv"), combined_stop_values)
    ref_core = " ".join(ref_core_tokens)
    apr_core = " ".join(apr_core_tokens)

    if not ref_core or not apr_core:
        return {
            "score": 0.0,
            "exact_core": False,
            "subset_core": False,
            "distinctive_shared": 0,
            "location_score": 0.0,
            "reference_core": ref_core,
            "apr_core": apr_core,
        }

    ref_set = set(ref_core_tokens)
    apr_set = set(apr_core_tokens)
    shared = ref_set & apr_set
    exact_core = ref_set == apr_set
    subset_core = ref_set <= apr_set or apr_set <= ref_set
    distinctive_shared = sum(
        1
        for token in shared
        if token not in AMBIGUOUS_CORE_TOKENS and len(token) >= 4
    )

    sequence = SequenceMatcher(None, ref_core, apr_core).ratio()
    token_score = jaccard(ref_set, apr_set)
    containment = 1.0 if ref_core in apr_core or apr_core in ref_core else 0.0
    location_score = location_similarity(reference.get("location"), apr_row)

    score = 0.50 * sequence + 0.30 * token_score + 0.08 * containment
    if exact_core:
        score += 0.28
    elif subset_core and distinctive_shared:
        score += 0.16
    score += 0.12 * location_score

    return {
        "score": min(score, 1.0),
        "exact_core": exact_core,
        "subset_core": subset_core,
        "distinctive_shared": distinctive_shared,
        "location_score": location_score,
        "reference_core": ref_core,
        "apr_core": apr_core,
    }


def best_match(
    reference: pd.Series, apr: pd.DataFrame
) -> tuple[pd.Series | None, float, float, dict[str, object]]:
    scores: list[tuple[float, int, dict[str, object]]] = []
    for index, apr_row in apr.iterrows():
        features = name_match_features(reference, apr_row)
        scores.append((float(features["score"]), index, features))

    if not scores:
        return None, 0.0, 0.0, {}

    scores.sort(key=lambda item: item[0], reverse=True)
    best_score, best_index, best_features = scores[0]
    second_score = scores[1][0] if len(scores) > 1 else 0.0
    return apr.loc[best_index], best_score, best_score - second_score, best_features


def match_method(
    score: float,
    margin: float,
    exact_id: bool,
    features: dict[str, object],
    reference_location: object,
) -> tuple[str, str]:
    if exact_id:
        return "matched", "known_company_id"

    exact_core = bool(features.get("exact_core"))
    subset_core = bool(features.get("subset_core"))
    distinctive_shared = int(features.get("distinctive_shared", 0))
    location_score = float(features.get("location_score", 0.0))
    has_reference_location = bool(normalize_text(reference_location))

    if exact_core and (location_score >= 0.7 or not has_reference_location):
        return "matched", "exact_core_name_location"
    if subset_core and distinctive_shared >= 1 and location_score >= 0.7 and margin >= 0.04:
        return "matched", "core_name_subset_location"

    # Tačno jezgro sa konfliktnom ili nedovoljno preciznom lokacijom ide na ručnu proveru.
    if exact_core and margin >= 0.15:
        return "manual_review", "exact_core_name_location_conflict"
    if subset_core and distinctive_shared >= 1 and margin >= 0.16:
        return "manual_review", "core_name_subset_unique_candidate"

    if score >= 0.90 and margin >= 0.08:
        return "matched", "high_confidence_name_location"
    if score >= 0.76 and margin >= 0.05:
        return "manual_review", "fuzzy_name_location"
    return "unmatched", "no_safe_match_in_narrow_apr_scope"


def load_reference_universe(references_path: Path, additions_path: Path) -> pd.DataFrame:
    references = pd.read_csv(references_path, dtype="string")
    references["known_maticni_broj"] = pd.NA
    references["known_activity_code"] = pd.NA
    references["known_status"] = pd.NA
    references["external_financial_year"] = pd.NA
    references["external_revenue_rsd"] = pd.NA
    references["external_employees"] = pd.NA

    if additions_path.exists():
        additions = pd.read_csv(additions_path, dtype="string")
        all_columns = list(dict.fromkeys([*references.columns, *additions.columns]))
        references = references.reindex(columns=all_columns).astype("object")
        additions = additions.reindex(columns=all_columns).astype("object")
        references = pd.concat([references, additions], ignore_index=True).convert_dtypes()

    return references


def add_candidate_fields(row: dict[str, object], apr_row: pd.Series | None) -> None:
    if apr_row is None:
        return
    for source, target in APR_EXPORT_FIELDS:
        row[f"candidate_{target}"] = apr_row.get(source)


def add_confirmed_match_fields(row: dict[str, object], apr_row: pd.Series | None) -> None:
    if apr_row is None:
        return
    for source, target in APR_EXPORT_FIELDS:
        row[target] = apr_row.get(source)


def main() -> None:
    args = parse_args()
    references = load_reference_universe(args.references.resolve(), args.manual_additions.resolve())

    allowed_potential = {"Visok", "Srednji"} if args.include_medium else {"Visok"}
    targets = references[references["agrix_potential"].isin(allowed_potential)].copy()

    apr = pd.read_excel(
        args.apr.resolve(),
        sheet_name="Validated Data",
        dtype={"maticni_broj": "string", "sifra_delatnosti": "string"},
    )
    if "is_active_market_candidate" in apr.columns:
        apr = apr[apr["is_active_market_candidate"].fillna(False)].copy()

    apr["maticni_broj"] = apr["maticni_broj"].astype("string").str.zfill(8)
    apr_by_id = apr.set_index("maticni_broj", drop=False)

    output_rows: list[dict[str, object]] = []
    for _, reference in targets.iterrows():
        known_id = normalize_text(reference.get("known_maticni_broj")).replace(" ", "")
        known_id = known_id.zfill(8) if known_id.isdigit() else ""
        exact_id = bool(known_id and known_id in apr_by_id.index)

        if exact_id:
            apr_row = apr_by_id.loc[known_id]
            if isinstance(apr_row, pd.DataFrame):
                apr_row = apr_row.iloc[0]
            score = 1.0
            margin = 1.0
            features: dict[str, object] = {
                "exact_core": True,
                "subset_core": True,
                "distinctive_shared": 1,
                "location_score": 1.0,
                "reference_core": normalized_company_name(reference.get("reference_name")),
                "apr_core": normalized_company_name(apr_row.get("naziv")),
            }
        else:
            apr_row, score, margin, features = best_match(reference, apr)

        status, method = match_method(
            score,
            margin,
            exact_id,
            features,
            reference.get("location"),
        )
        row = reference.to_dict()
        row.update(
            {
                "priority_tier": (
                    "A"
                    if reference.get("category")
                    in {"Voće, povrće i hladnjače", "Poljoprivreda i kooperative", "Duvan"}
                    else "B"
                    if reference.get("category") == "Mlekarstvo"
                    else "C"
                ),
                "apr_match_scope": APR_MATCH_SCOPE,
                "apr_match_status": status,
                "match_method": method,
                "match_score": round(score, 4),
                "match_margin": round(margin, 4),
                "core_name_exact": bool(features.get("exact_core")),
                "core_name_subset": bool(features.get("subset_core")),
                "distinctive_shared_tokens": int(features.get("distinctive_shared", 0)),
                "location_match_score": round(float(features.get("location_score", 0.0)), 3),
                "reference_core_name": features.get("reference_core", ""),
                "candidate_core_name": features.get("apr_core", ""),
                "unmatched_reason": (
                    "No safe identity match in the current APR dataset limited to activity codes 1039/4631."
                    if status == "unmatched"
                    else pd.NA
                ),
            }
        )

        # Best candidate is kept only as an explicit diagnostic. It is not a confirmed identity.
        add_candidate_fields(row, apr_row)
        if status in {"matched", "manual_review"}:
            add_confirmed_match_fields(row, apr_row)

        output_rows.append(row)

    output = pd.DataFrame(output_rows)
    output = output.sort_values(
        ["priority_tier", "apr_match_status", "apr_prihod_rsd", "reference_name"],
        ascending=[True, True, False, True],
        na_position="last",
    )

    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(args.output_csv, index=False, encoding="utf-8-sig")

    matched_count = int(output["apr_match_status"].eq("matched").sum())
    manual_count = int(output["apr_match_status"].eq("manual_review").sum())
    unmatched_count = int(output["apr_match_status"].eq("unmatched").sum())
    reference_count = len(output)

    summary = pd.DataFrame(
        [
            ("reference_rows", reference_count),
            ("matched", matched_count),
            ("manual_review", manual_count),
            ("unmatched", unmatched_count),
            ("safe_match_rate_pct", round(100 * matched_count / reference_count, 1) if reference_count else 0.0),
            ("review_or_match_rate_pct", round(100 * (matched_count + manual_count) / reference_count, 1) if reference_count else 0.0),
            ("tier_a", int(output["priority_tier"].eq("A").sum())),
            ("tier_b", int(output["priority_tier"].eq("B").sum())),
            ("tier_c", int(output["priority_tier"].eq("C").sum())),
        ],
        columns=["metric", "value"],
    )

    with pd.ExcelWriter(args.output_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        output.to_excel(writer, sheet_name="Replacement Targets", index=False)
        output[output["apr_match_status"].eq("matched")].to_excel(
            writer, sheet_name="Matched", index=False
        )
        output[output["apr_match_status"].eq("manual_review")].to_excel(
            writer, sheet_name="Manual Review", index=False
        )
        output[output["apr_match_status"].eq("unmatched")].to_excel(
            writer, sheet_name="Unmatched", index=False
        )

    print(f"CSV: {args.output_csv.resolve()}")
    print(f"Excel: {args.output_xlsx.resolve()}")
    print(f"APR match scope: {APR_MATCH_SCOPE}")
    for summary_row in summary.itertuples(index=False):
        print(f"  - {summary_row.metric}: {summary_row.value}")


if __name__ == "__main__":
    main()
