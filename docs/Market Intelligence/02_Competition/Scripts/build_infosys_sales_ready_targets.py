from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parent
COMPETITION_DIR = SCRIPT_DIR.parent
DEFAULT_INPUT = COMPETITION_DIR / "infosys_wide_enrichment.csv"
DEFAULT_OUTPUT_CSV = COMPETITION_DIR / "infosys_sales_ready_targets.csv"
DEFAULT_OUTPUT_XLSX = COMPETITION_DIR / "infosys_sales_ready_targets.xlsx"
DEFAULT_OUTPUT_MD = COMPETITION_DIR / "infosys_sales_ready_summary.md"

DIRECT_FRUIT_CODES = {"1039", "4631"}
DIRECT_AGRI_TRADE_CODES = {"4621", "5210"}
PRIMARY_AGRI_PREFIXES = ("01",)
DAIRY_CODES = {"1051", "1052"}
TOBACCO_CODES = {"1200", "4621"}

PRIORITY_REGIONS = {
    "arilje",
    "bajina basta",
    "brus",
    "cacak",
    "cajetina",
    "kosjeric",
    "kraljevo",
    "lucani",
    "pozega",
    "uzice",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Deduplikuje Infosys APR match-eve i odvaja potvrđen identitet od "
            "stvarnog AgriX process fit-a."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--output-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX)
    parser.add_argument("--output-md", type=Path, default=DEFAULT_OUTPUT_MD)
    return parser.parse_args()


def fold_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKD", str(value).casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    transliteration = str.maketrans(
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
    text = text.translate(transliteration).replace("đ", "dj")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def normalize_company_id(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value))
    return digits.zfill(8) if digits else ""


def normalize_activity_code(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.zfill(4) if text.isdigit() else text


def is_active_status(value: object) -> bool:
    text = fold_text(value)
    if not text:
        return False
    negative = ("likvid", "stecaj", "neaktivan", "brisan", "ugasen", "prestao")
    if any(token in text for token in negative):
        return False
    return "aktivan" in text or "registrovan" in text


def category_family(value: object) -> str:
    text = fold_text(value)
    if "voce" in text or "povrce" in text or "hladnjace" in text:
        return "fruit_cold_storage"
    if "poljoprivreda" in text or "kooperative" in text:
        return "agri_cooperation"
    if "duvan" in text:
        return "tobacco"
    if "mlekar" in text:
        return "dairy"
    return "other"


def process_fit_status(row: pd.Series) -> tuple[str, str]:
    if not is_active_status(row.get("wide_status")):
        return "exclude_non_active", "APR status nije aktivan."

    family = category_family(row.get("category"))
    code = normalize_activity_code(row.get("wide_sifra_delatnosti"))
    confidence = fold_text(row.get("classification_confidence"))

    if family == "fruit_cold_storage":
        if code in DIRECT_FRUIT_CODES:
            return "direct_evidence", "Voćarski/hladnjačarski izvor i direktna APR šifra."
        if code == "5210":
            return "strong_candidate_needs_process_check", "Skladištenje je relevantno, ali otkupni tok treba potvrditi."
        if "potrebna provera" in confidence:
            return "conflicting_or_weak_evidence", "Izvorna klasifikacija je nesigurna, a APR šifra nije direktno relevantna."
        return "strong_candidate_needs_process_check", "Naziv/reference signal je jak, ali APR šifra ne dokazuje otkup."

    if family == "agri_cooperation":
        if code in DIRECT_AGRI_TRADE_CODES:
            return "direct_evidence", "APR šifra potvrđuje trgovinu poljoprivrednim proizvodima ili skladištenje."
        if code.startswith(PRIMARY_AGRI_PREFIXES):
            return "strong_candidate_needs_process_check", "Poljoprivredna proizvodnja je potvrđena; kooperantski/otkupni tok treba proveriti."
        return "conflicting_or_weak_evidence", "APR delatnost ne potvrđuje kooperaciju ili otkup."

    if family == "tobacco":
        if code in TOBACCO_CODES:
            return "direct_evidence", "Duvanski ili poljoprivredno-trgovinski tok je potvrđen šifrom."
        return "strong_candidate_needs_process_check", "Duvanski reference signal postoji, ali proces treba potvrditi."

    if family == "dairy":
        if code in DAIRY_CODES:
            return "adjacent_discovery", "Mlekarstvo je potvrđeno; AgriX fit zahteva zaseban discovery."
        return "adjacent_discovery", "Infosys mlekarstvo je adjacent segment, ne automatski Enterprise target."

    return "conflicting_or_weak_evidence", "Nema dovoljno dokaza za trenutni AgriX process fit."


def revenue_score(value: object) -> int:
    revenue = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(revenue) or revenue < 0:
        return 0
    if 100_000_000 <= revenue <= 2_000_000_000:
        return 18
    if revenue > 2_000_000_000:
        return 15
    if revenue >= 50_000_000:
        return 12
    if revenue >= 10_000_000:
        return 8
    return 3


def employee_score(value: object) -> int:
    employees = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(employees) or employees < 0:
        return 0
    if 3 <= employees <= 100:
        return 8
    if employees > 100:
        return 6
    if employees >= 1:
        return 4
    return 0


def fit_score(status: str) -> int:
    return {
        "direct_evidence": 45,
        "strong_candidate_needs_process_check": 28,
        "adjacent_discovery": 12,
        "conflicting_or_weak_evidence": 0,
        "exclude_non_active": -100,
    }.get(status, 0)


def readiness_rank(status: str) -> int:
    return {
        "direct_evidence": 1,
        "strong_candidate_needs_process_check": 2,
        "adjacent_discovery": 3,
        "conflicting_or_weak_evidence": 4,
        "exclude_non_active": 5,
    }.get(status, 9)


def region_score(value: object) -> int:
    return 8 if fold_text(value) in PRIORITY_REGIONS else 0


def aggregate_unique(values: pd.Series) -> str:
    clean = [str(value).strip() for value in values.dropna() if str(value).strip()]
    return " | ".join(dict.fromkeys(clean))


def build_deduplicated_accounts(matched: pd.DataFrame) -> pd.DataFrame:
    matched = matched.copy()
    matched["wide_maticni_broj"] = matched["wide_maticni_broj"].map(normalize_company_id)
    matched = matched[matched["wide_maticni_broj"].ne("")].copy()

    rows: list[dict[str, object]] = []
    for company_id, group in matched.groupby("wide_maticni_broj", sort=False):
        representative = group.iloc[0].copy()
        representative["source_reference_names"] = aggregate_unique(group["reference_name"])
        representative["source_reference_count"] = len(group)
        representative["source_categories"] = aggregate_unique(group["category"])
        representative["wide_maticni_broj"] = company_id
        rows.append(representative.to_dict())

    accounts = pd.DataFrame(rows)
    if accounts.empty:
        return accounts

    fit_results = accounts.apply(process_fit_status, axis=1)
    accounts["process_fit_status"] = [result[0] for result in fit_results]
    accounts["process_fit_reason"] = [result[1] for result in fit_results]
    accounts["active_legal_entity"] = accounts["wide_status"].map(is_active_status)
    accounts["revenue_score"] = accounts["wide_prihod_rsd"].map(revenue_score)
    accounts["employee_score"] = accounts["wide_employees"].map(employee_score)
    accounts["region_score"] = accounts["wide_opstina"].map(region_score)
    accounts["fit_score"] = accounts["process_fit_status"].map(fit_score)
    accounts["identity_score"] = 10
    accounts["target_score"] = (
        accounts["fit_score"]
        + accounts["identity_score"]
        + accounts["revenue_score"]
        + accounts["employee_score"]
        + accounts["region_score"]
    )
    accounts["readiness_rank"] = accounts["process_fit_status"].map(readiness_rank)
    accounts["sales_readiness"] = accounts["process_fit_status"].map(
        {
            "direct_evidence": "ready_for_account_research",
            "strong_candidate_needs_process_check": "process_validation_first",
            "adjacent_discovery": "adjacent_discovery_only",
            "conflicting_or_weak_evidence": "hold_until_process_proven",
            "exclude_non_active": "exclude",
        }
    )

    return accounts.sort_values(
        ["readiness_rank", "target_score", "wide_prihod_rsd", "wide_naziv"],
        ascending=[True, False, False, True],
        na_position="last",
    ).reset_index(drop=True)


def format_rsd(value: object) -> str:
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(number):
        return "—"
    return f"{int(number):,}".replace(",", ".")


def write_markdown(
    path: Path,
    accounts: pd.DataFrame,
    manual_review: pd.DataFrame,
    unmatched: pd.DataFrame,
) -> None:
    ready = accounts[accounts["sales_readiness"].eq("ready_for_account_research")]
    validation = accounts[accounts["sales_readiness"].eq("process_validation_first")]
    adjacent = accounts[accounts["sales_readiness"].eq("adjacent_discovery_only")]
    hold = accounts[accounts["sales_readiness"].eq("hold_until_process_proven")]
    excluded = accounts[accounts["sales_readiness"].eq("exclude")]
    top20 = accounts[accounts["sales_readiness"].isin(
        ["ready_for_account_research", "process_validation_first"]
    )].head(20)

    lines = [
        "# Infosys replacement — sales-ready target summary",
        "",
        "**Status:** Evidence-based account prioritization  ",
        "**Input:** `infosys_wide_enrichment.csv`  ",
        "**Važno:** tačan APR identitet nije isto što i potvrđen AgriX process fit.",
        "",
        "## Sažetak",
        "",
        "| Pokazatelj | Broj |",
        "|---|---:|",
        f"| Jedinstveni matched pravni subjekti | {len(accounts)} |",
        f"| Spremni za account research | {len(ready)} |",
        f"| Prvo potvrditi proces | {len(validation)} |",
        f"| Adjacent discovery | {len(adjacent)} |",
        f"| Hold — slab ili konfliktan process evidence | {len(hold)} |",
        f"| Isključeni zbog neaktivnog statusa | {len(excluded)} |",
        f"| Identiteti za ručnu proveru | {len(manual_review)} |",
        f"| Bez bezbednog identiteta | {len(unmatched)} |",
        "",
        "## Top 20 za istraživanje",
        "",
        "| # | Firma | MB | Prihod RSD | Zaposleni | Status prioriteta | Score |",
        "|---:|---|---|---:|---:|---|---:|",
    ]

    for position, row in enumerate(top20.itertuples(index=False), start=1):
        lines.append(
            "| "
            f"{position} | {row.wide_naziv} | {row.wide_maticni_broj} | "
            f"{format_rsd(row.wide_prihod_rsd)} | "
            f"{int(row.wide_employees) if pd.notna(row.wide_employees) else '—'} | "
            f"{row.sales_readiness} | {int(row.target_score)} |"
        )

    lines.extend(
        [
            "",
            "## Pravila korišćenja",
            "",
            "1. `ready_for_account_research` znači da su identitet, aktivan status i procesni signal dovoljno jaki za dublji research — ne za automatski hladni kontakt.",
            "2. `process_validation_first` zahteva potvrdu broja stanica, kooperanata, dokumenata i stvarnog otkupnog toka.",
            "3. `adjacent_discovery_only` ne ulazi u Enterprise outbound bez posebnog product discovery-ja.",
            "4. `hold_until_process_proven` čuva konkurentski dokaz, ali ga ne pretvara u prodajni target.",
            "5. Reference ne dokazuju da firma i dalje koristi Infosys.",
            "",
        ]
    )
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Nije pronađen ulaz: {input_path}")

    data = pd.read_csv(input_path, dtype="string")
    required = {
        "reference_name",
        "category",
        "classification_confidence",
        "wide_match_status",
        "wide_maticni_broj",
        "wide_naziv",
        "wide_sifra_delatnosti",
        "wide_status",
        "wide_opstina",
        "wide_prihod_rsd",
        "wide_employees",
    }
    missing = sorted(required - set(data.columns))
    if missing:
        raise ValueError(f"Nedostaju kolone: {', '.join(missing)}")

    for column in ("wide_prihod_rsd", "wide_employees", "wide_match_score"):
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce")

    matched = data[data["wide_match_status"].eq("matched")].copy()
    manual_review = data[data["wide_match_status"].eq("manual_review")].copy()
    unmatched = data[data["wide_match_status"].eq("unmatched")].copy()
    accounts = build_deduplicated_accounts(matched)

    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    accounts.to_csv(args.output_csv, index=False, encoding="utf-8-sig")

    top20 = accounts[accounts["sales_readiness"].isin(
        ["ready_for_account_research", "process_validation_first"]
    )].head(20)
    with pd.ExcelWriter(args.output_xlsx, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                ("matched_reference_rows", len(matched)),
                ("unique_matched_accounts", len(accounts)),
                ("ready_for_account_research", int(accounts["sales_readiness"].eq("ready_for_account_research").sum())),
                ("process_validation_first", int(accounts["sales_readiness"].eq("process_validation_first").sum())),
                ("adjacent_discovery_only", int(accounts["sales_readiness"].eq("adjacent_discovery_only").sum())),
                ("hold_until_process_proven", int(accounts["sales_readiness"].eq("hold_until_process_proven").sum())),
                ("excluded", int(accounts["sales_readiness"].eq("exclude").sum())),
                ("manual_identity_review", len(manual_review)),
                ("unmatched", len(unmatched)),
            ],
            columns=["metric", "value"],
        ).to_excel(writer, sheet_name="Summary", index=False)
        top20.to_excel(writer, sheet_name="Top 20", index=False)
        accounts.to_excel(writer, sheet_name="All Matched Accounts", index=False)
        manual_review.to_excel(writer, sheet_name="Identity Manual Review", index=False)
        unmatched.to_excel(writer, sheet_name="Unmatched", index=False)

    write_markdown(args.output_md, accounts, manual_review, unmatched)

    print(f"CSV: {args.output_csv.resolve()}")
    print(f"Excel: {args.output_xlsx.resolve()}")
    print(f"Markdown: {args.output_md.resolve()}")
    print(f"  - matched_reference_rows: {len(matched)}")
    print(f"  - unique_matched_accounts: {len(accounts)}")
    for status, count in accounts["sales_readiness"].value_counts(dropna=False).items():
        print(f"  - {status}: {int(count)}")
    print(f"  - manual_identity_review: {len(manual_review)}")
    print(f"  - unmatched: {len(unmatched)}")


if __name__ == "__main__":
    main()
