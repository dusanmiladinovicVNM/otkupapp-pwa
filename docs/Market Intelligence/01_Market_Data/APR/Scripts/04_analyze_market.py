from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from apr_pipeline_common import (
    PROCESSED_DIR,
    REPORTS_DIR,
    ensure_directories,
    latest_file,
    numeric_series,
    sha256_file,
    utc_timestamp,
    write_json,
)

REVENUE_BINS = [-1, 10_000_000, 50_000_000, 100_000_000, 500_000_000, 1_000_000_000, float("inf")]
REVENUE_LABELS = [
    "0–10m RSD",
    "10–50m RSD",
    "50–100m RSD",
    "100–500m RSD",
    "500m–1bn RSD",
    "1bn+ RSD",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Pravi tržišne analize iz validiranog APR dataseta.")
    parser.add_argument("--input", type=Path, help="Validirani Excel iz Processed foldera.")
    parser.add_argument(
        "--include-inactive",
        action="store_true",
        help="U tržišnu analizu uključuje i firme koje nisu označene kao aktivne.",
    )
    return parser.parse_args()


def concentration_table(df: pd.DataFrame) -> pd.DataFrame:
    revenue_df = df[df["ukupni_prihodi"].notna() & df["ukupni_prihodi"].ge(0)].copy()
    revenue_df = revenue_df.sort_values("ukupni_prihodi", ascending=False).reset_index(drop=True)
    total = revenue_df["ukupni_prihodi"].sum()
    if total <= 0:
        return pd.DataFrame(
            columns=["rank", "maticni_broj", "naziv", "ukupni_prihodi", "revenue_share", "cumulative_share"]
        )
    revenue_df["rank"] = revenue_df.index + 1
    revenue_df["revenue_share"] = revenue_df["ukupni_prihodi"] / total
    revenue_df["cumulative_share"] = revenue_df["revenue_share"].cumsum()
    return revenue_df[["rank", "maticni_broj", "naziv", "ukupni_prihodi", "revenue_share", "cumulative_share"]]


def percentile_value(series: pd.Series, percentile: float) -> float | None:
    clean = series.dropna()
    return None if clean.empty else float(clean.quantile(percentile))


def main() -> None:
    args = parse_args()
    ensure_directories()

    input_path = args.input or latest_file(PROCESSED_DIR, "apr_market_validated_*.xlsx")
    if not input_path.is_absolute():
        input_path = input_path.resolve()

    df = pd.read_excel(input_path, sheet_name="Validated Data", dtype={"maticni_broj": "string"})
    if "ukupni_prihodi" not in df.columns:
        raise ValueError("Dataset nema kolonu 'ukupni_prihodi'.")

    df["ukupni_prihodi"] = numeric_series(df["ukupni_prihodi"])
    if "prosecan_broj_zaposlenih" in df.columns:
        df["prosecan_broj_zaposlenih"] = numeric_series(df["prosecan_broj_zaposlenih"])

    status_distribution = (
        df.groupby(["status", "status_category"], dropna=False)
        .agg(companies=("maticni_broj", "size"))
        .reset_index()
        .sort_values(["companies", "status"], ascending=[False, True])
        if {"status", "status_category"}.issubset(df.columns)
        else pd.DataFrame(columns=["status", "status_category", "companies"])
    )

    market_df = df.copy()
    if not args.include_inactive:
        if "is_active_market_candidate" not in market_df.columns:
            raise ValueError("Dataset nema kolonu 'is_active_market_candidate'. Pokreni 03_clean_validate.py.")
        market_df = market_df[market_df["is_active_market_candidate"].fillna(False)].copy()

    if market_df.empty:
        status_preview = "; ".join(
            f"{row.status!r}={int(row.companies)} ({row.status_category})"
            for row in status_distribution.head(10).itertuples(index=False)
        )
        raise RuntimeError(
            "Tržišni scope je prazan. Izveštaj nije generisan jer bi dao lažne nulte pokazatelje. "
            "Proveri status klasifikaciju u 03_clean_validate.py ili koristi --include-inactive samo za dijagnostiku. "
            f"Najčešći statusi: {status_preview or 'nema dostupnih statusa'}"
        )

    market_df["revenue_band"] = pd.cut(
        market_df["ukupni_prihodi"],
        bins=REVENUE_BINS,
        labels=REVENUE_LABELS,
        include_lowest=True,
        right=True,
    )

    revenue_valid = market_df[market_df["ukupni_prihodi"].notna() & market_df["ukupni_prihodi"].ge(0)].copy()
    concentration = concentration_table(market_df)

    summary_rows = [
        ("market_rows", len(market_df)),
        ("rows_with_revenue", len(revenue_valid)),
        ("revenue_coverage_pct", len(revenue_valid) / len(market_df)),
        ("total_revenue_rsd", float(revenue_valid["ukupni_prihodi"].sum())),
        ("mean_revenue_rsd", float(revenue_valid["ukupni_prihodi"].mean()) if len(revenue_valid) else None),
        ("median_revenue_rsd", float(revenue_valid["ukupni_prihodi"].median()) if len(revenue_valid) else None),
        ("p25_revenue_rsd", percentile_value(revenue_valid["ukupni_prihodi"], 0.25)),
        ("p75_revenue_rsd", percentile_value(revenue_valid["ukupni_prihodi"], 0.75)),
        ("p90_revenue_rsd", percentile_value(revenue_valid["ukupni_prihodi"], 0.90)),
        ("p95_revenue_rsd", percentile_value(revenue_valid["ukupni_prihodi"], 0.95)),
    ]

    if not concentration.empty:
        for top_n in (10, 25, 50, 100):
            share = float(concentration.head(top_n)["revenue_share"].sum())
            summary_rows.append((f"top_{top_n}_revenue_share", share))

    summary_df = pd.DataFrame(summary_rows, columns=["metric", "value"])

    revenue_bands = (
        market_df.groupby("revenue_band", observed=False)
        .agg(
            companies=("maticni_broj", "count"),
            total_revenue_rsd=("ukupni_prihodi", "sum"),
            median_revenue_rsd=("ukupni_prihodi", "median"),
        )
        .reset_index()
    )
    revenue_bands["company_share"] = revenue_bands["companies"] / len(market_df)

    municipality_column = "opstina_apr" if "opstina_apr" in market_df.columns else "opstina"
    regions = (
        market_df.assign(region=market_df[municipality_column].fillna("").replace("", "Nepoznato"))
        .groupby("region", dropna=False)
        .agg(
            companies=("maticni_broj", "count"),
            companies_with_revenue=("ukupni_prihodi", "count"),
            total_revenue_rsd=("ukupni_prihodi", "sum"),
            median_revenue_rsd=("ukupni_prihodi", "median"),
        )
        .reset_index()
        .sort_values(["total_revenue_rsd", "companies"], ascending=False)
    )

    activities = (
        market_df.groupby("sifra_delatnosti", dropna=False)
        .agg(
            companies=("maticni_broj", "count"),
            total_revenue_rsd=("ukupni_prihodi", "sum"),
            median_revenue_rsd=("ukupni_prihodi", "median"),
        )
        .reset_index()
        .sort_values("total_revenue_rsd", ascending=False)
    )

    top_companies = market_df.sort_values("ukupni_prihodi", ascending=False).head(200).copy()
    top_columns = [
        column
        for column in (
            "maticni_broj",
            "naziv",
            "sifra_delatnosti",
            "opstina",
            "opstina_apr",
            "status",
            "status_category",
            "godina",
            "ukupni_prihodi",
            "prosecan_broj_zaposlenih",
            "neto_dobitak",
            "neto_gubitak",
            "revenue_band",
        )
        if column in top_companies.columns
    ]
    top_companies = top_companies[top_columns]

    suffix = input_path.stem.replace("apr_market_validated_", "")
    workbook_path = REPORTS_DIR / f"apr_market_analysis_{suffix}.xlsx"
    markdown_path = REPORTS_DIR / f"apr_market_summary_{suffix}.md"
    metadata_path = workbook_path.with_suffix(".metadata.json")

    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        status_distribution.to_excel(writer, sheet_name="Status Distribution", index=False)
        revenue_bands.to_excel(writer, sheet_name="Revenue Bands", index=False)
        regions.to_excel(writer, sheet_name="Regions", index=False)
        activities.to_excel(writer, sheet_name="Activities", index=False)
        top_companies.to_excel(writer, sheet_name="Top Companies", index=False)
        concentration.to_excel(writer, sheet_name="Concentration", index=False)
        market_df.to_excel(writer, sheet_name="Market Data", index=False)

    summary_map = dict(summary_rows)
    markdown = f"""# APR tržišni pregled

**Generisano:** {utc_timestamp()}  
**Ulaz:** `{input_path.name}`  
**Scope:** {'sve firme' if args.include_inactive else 'aktivni tržišni kandidati'}

## Osnovni pokazatelji

- Broj firmi u analiziranom scope-u: **{len(market_df):,}**
- Firme sa podatkom o prihodima: **{len(revenue_valid):,}**
- Pokrivenost prihoda: **{summary_map['revenue_coverage_pct']:.1%}**
- Ukupni prihodi: **{summary_map['total_revenue_rsd']:,.0f} RSD**
- Medijana prihoda: **{(summary_map['median_revenue_rsd'] or 0):,.0f} RSD**
- 90. percentil prihoda: **{(summary_map['p90_revenue_rsd'] or 0):,.0f} RSD**

## Koncentracija

"""
    for top_n in (10, 25, 50, 100):
        key = f"top_{top_n}_revenue_share"
        if key in summary_map:
            markdown += f"- Top {top_n} firmi: **{summary_map[key]:.1%}** prihoda analiziranog skupa\n"

    markdown += """
## Važna ograničenja

- APR šifre `1039` i `4631` ne obuhvataju nužno sve organizovane otkupljivače.
- Prihod nije direktna mera broja stanica, kooperanata, logističke složenosti ili GGAP potrebe.
- Finansijski endpoint se u trenutnom pipeline-u tretira kao jedan zapis po firmi; višegodišnji trend nije potvrđen ovim izveštajem.
- Status klasifikacija se čuva u posebnim sheet-ovima i mora se proveriti kada APR uvede novu status vrednost.
- Ovaj rezultat je tržišni dokaz i input za `04_MARKET.md`, ali nije automatski TAM/SAM/SOM bez dodatnih pravila segmentacije.
"""
    markdown_path.write_text(markdown, encoding="utf-8")

    write_json(
        metadata_path,
        {
            "dataset": "APR market analysis",
            "generated_at_utc": utc_timestamp(),
            "input_file": input_path.name,
            "input_sha256": sha256_file(input_path),
            "output_workbook": workbook_path.name,
            "output_workbook_sha256": sha256_file(workbook_path),
            "output_markdown": markdown_path.name,
            "include_inactive": args.include_inactive,
            "market_rows": len(market_df),
            "rows_with_revenue": len(revenue_valid),
            "revenue_bands_rsd": REVENUE_LABELS,
            "status_distribution": status_distribution.to_dict(orient="records"),
        },
    )

    print(f"Analiza: {workbook_path}")
    print(f"Sažetak: {markdown_path}")
    print(f"Broj firmi u scope-u: {len(market_df)}")


if __name__ == "__main__":
    main()
