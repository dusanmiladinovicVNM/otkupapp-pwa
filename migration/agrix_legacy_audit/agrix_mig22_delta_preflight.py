#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-22A: read-only source delta preflight against the frozen MIG-02 stage.

Compares the NEW legacy XLSM against the ORIGINAL legacy rows captured by
MigrationStage_v2. NEVER compares canonical target rows to legacy rows.
Produces PRIVATE local count/field summaries and a CSV of new document IDs.
Does not write to XLSM, create canonical rows, or run/import VBA.

Coverage: OTK, OTP, ZBR, PRJ and legacy ambalaza. This is a source-delta
preflight, NOT approval to import; other legacy tables, lifecycle, linkage,
and canonical ambalaza still require separate reconciliation.

Usage (PowerShell; repo checkout contains this file):
  python -u agrix_mig22_delta_preflight.py --new "source.xlsm" \
      --stage "E:\AgriX\AgriX DEV\MigrationStage_v2" \
      --out-dir "E:\AgriX\AgriX DEV\MIG22\DeltaPreflight"
"""
from __future__ import annotations

import argparse
import collections
import csv
import hashlib
import json
import sys
from pathlib import Path

from agrix_mig19_malina_readonly_audit import XlsmReadOnly

VERSION = "AGRIx-MIG22A-SOURCE-DELTA-READONLY-v1"
# The first three are unmodified legacy snapshots, not canonical headers.
BASELINES = (
    ("tblOtkup", "SOURCE__tblOtkup.jsonl", "OtkupID", 1702, True),
    ("tblOtpremnica", "SOURCE__tblOtpremnica.jsonl", "OtpremnicaID", 907, True),
    ("tblZbirna", "SOURCE__tblZbirna.jsonl", "ZbirnaID", 907, True),
    # Passthrough stage: might omit legacy-only fields.
    ("tblPrijemnica", "tblPrijemnica.jsonl", "PrijemnicaID", 893, False),
    ("tblAmbalaza", "UNCONVERTED__tblAmbalaza.jsonl", "AmbID", None, True),
)
EXPECTED_NEW_COUNTS = {
    "tblOtkup": 1798, "tblOtpremnica": 965, "tblZbirna": 965,
    "tblPrijemnica": 943, "tblAmbalaza": 8242,
}
IGNORED_KEYS = {"_excel_row"}
# These are textual comparisons, not domain conversions. Any difference is a
# candidate to inspect; no difference is only proof for SHARED retained fields.


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for part in iter(lambda: f.read(1024 * 1024), b""):
            h.update(part)
    return h.hexdigest()


def find_stage_file(stage: Path, filename: str) -> Path:
    direct = stage / filename
    if direct.is_file():
        return direct
    matches = sorted(p for p in stage.rglob(filename) if p.is_file())
    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one stage file {filename}; found {len(matches)}"
        )
    return matches[0]


def load_jsonl(path: Path) -> list[dict]:
    rows = []
    with path.open("r", encoding="utf-8-sig") as f:
        for lineno, line in enumerate(f, 1):
            if not line.strip():
                continue
            row = json.loads(line)
            if not isinstance(row, dict):
                raise ValueError(f"Invalid object in {path.name}, line {lineno}")
            rows.append(row)
    return rows


def norm(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def indexed(rows: list[dict], pk: str, where: str) -> dict[str, dict]:
    out = {}
    for row in rows:
        key = norm(row.get(pk))
        if not key:
            raise ValueError(f"{where}: blank {pk}; STOP")
        if key in out:
            raise ValueError(f"{where}: duplicate {pk}; STOP")
        out[key] = row
    return out


def compare(old: dict[str, dict], new: dict[str, dict]) -> tuple[dict, list[str]]:
    oids, nids = set(old), set(new)
    additions = sorted(nids - oids)
    removed = oids - nids
    modifications = 0
    changed_fields = collections.Counter()
    checked_fields = set()
    unverified_old_fields = set()
    unverified_new_fields = set()

    for oid in oids & nids:
        before = {k: v for k, v in old[oid].items() if k not in IGNORED_KEYS}
        after = {k: v for k, v in new[oid].items() if k not in IGNORED_KEYS}
        common = before.keys() & after.keys()
        checked_fields.update(common)
        unverified_old_fields.update(before.keys() - after.keys())
        unverified_new_fields.update(after.keys() - before.keys())
        changed = [k for k in common if norm(before[k]) != norm(after[k])]
        if changed:
            modifications += 1
            changed_fields.update(changed)

    details = {
        "old": len(old),
        "new": len(new),
        "same_id": len(oids & nids),
        "new_ids": len(additions),
        "removed_ids": len(removed),
        "changed_old_rows": modifications,
        "changed_columns_counts": dict(sorted(changed_fields.items())),
        "shared_columns_checked": len(checked_fields),
        "stage_only_columns_unverified": sorted(unverified_old_fields),
        "new_only_columns_unverified": sorted(unverified_new_fields),
    }
    return details, additions


def guard_private_destination(out_dir: Path) -> None:
    repo_root = Path(__file__).resolve().parents[2]
    try:
        out_dir.resolve().relative_to(repo_root)
    except ValueError:
        return
    raise ValueError("Private migration reports cannot be written inside the Git repo")


def run(source: Path, stage: Path, out_dir: Path) -> int:
    if not source.is_file() or not stage.is_dir():
        raise FileNotFoundError("New legacy XLSM or MigrationStage_v2 does not exist")
    if source.suffix.lower() not in {".xlsm", ".xlsx"}:
        raise ValueError("Source must be an Excel workbook")
    guard_private_destination(out_dir)
    source_sha = digest(source)
    stage_paths = {t: find_stage_file(stage, name)
                   for t, name, _, _, _ in BASELINES}
    baseline = {}
    for table, _, pk, expected, _ in BASELINES:
        old = load_jsonl(stage_paths[table])
        if expected is not None and len(old) != expected:
            raise ValueError(
                f"{table}: frozen stage has {len(old)} records instead of {expected}; STOP"
            )
        baseline[table] = indexed(old, pk, f"frozen stage {table}")

    book = XlsmReadOnly(source)
    try:
        if len(book.tables) != 38:
            raise ValueError(f"Expected 38 legacy Excel tables, found {len(book.tables)}")
        comparisons = {}
        ids_for_csv = []
        any_blocker = False
        for table, _, pk, expected, full_baseline in BASELINES:
            if table not in book.tables:
                raise ValueError(f"Missing new-source table {table}; STOP")
            new_rows = book.read_table(table, {pk})
            now = indexed(new_rows, pk, f"new source {table}")
            detail, added_ids = compare(baseline[table], now)
            detail["baseline_raw_all_columns"] = full_baseline
            detail["expected_new_count"] = EXPECTED_NEW_COUNTS[table]
            detail["count_as_expected"] = detail["new"] == EXPECTED_NEW_COUNTS[table]
            detail["baseline_file_sha256"] = digest(stage_paths[table])
            comparisons[table] = detail
            if (detail["removed_ids"] or detail["changed_old_rows"]
                    or not detail["count_as_expected"]):
                any_blocker = True
            for ident in added_ids:
                ids_for_csv.append({"table": table, "id": ident})
            print(
                f"{table}: OLD={detail['old']} NEW={detail['new']} "
                f"ADDED={detail['new_ids']} CHANGED_OLD={detail['changed_old_rows']} "
                f"REMOVED={detail['removed_ids']} "
                f"COMMON_FIELDS={detail['shared_columns_checked']}"
            )
    finally:
        book.close()

    if digest(source) != source_sha:
        raise ValueError("Source workbook changed during the read-only audit; STOP")

    result = {
        "format": VERSION,
        "status": ("REVIEW_REQUIRED" if any_blocker
                   else "CORE_DELTA_CANDIDATE_NOT_IMPORT_APPROVAL"),
        "source_filename": source.name,
        "source_sha256": source_sha,
        "scope": "five core legacy tables versus frozen source stage",
        "import_approved": False,
        "notes": [
            "No writes to either XLSM, stage, production VBA or database.",
            "New ID lists and local reports are private and must never be committed.",
            "Only shared source-stage columns are compared; unverified fields are listed.",
            "This is not all-38-table drift coverage; links, lifecycle and "
            "canonical packaging require MIG-22B/C."
        ],
        "tables": comparisons,
    }

    out_dir.mkdir(parents=True, exist_ok=True)
    with (out_dir / "MIG22A_SUMMARY_PRIVATE.json").open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    with (out_dir / "MIG22A_NEW_IDS_PRIVATE.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as f:
        writer = csv.DictWriter(f, fieldnames=("table", "id"))
        writer.writeheader()
        writer.writerows(ids_for_csv)
    print("STATUS:", result["status"])
    print("SUMMARY:", out_dir / "MIG22A_SUMMARY_PRIVATE.json")
    print("IDS (PRIVATE):", out_dir / "MIG22A_NEW_IDS_PRIVATE.csv")
    return 0 if not any_blocker else 2


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--new", type=Path, required=True, help="new ORIGINAL legacy XLSM")
    ap.add_argument("--stage", type=Path, required=True, help="frozen MigrationStage_v2")
    ap.add_argument("--out-dir", type=Path, required=True, help="PRIVATE local output")
    args = ap.parse_args()
    try:
        return run(args.new, args.stage, args.out_dir)
    except Exception as exc:
        print(f"MIG22A STOP: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
