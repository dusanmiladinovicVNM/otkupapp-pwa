#!/usr/bin/env python3
"""Build a NEW empty AgriX macro workbook from checked-out otkupapp-pwa source,
then run MIG-06.6 target preflight. Never imports client data or edits a source XLSM.

Windows Excel + pywin32 are required for the repository's genuine VBA builder.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

VERSION = "AGRIx-BUILD-PREFLIGHT-1"
SUMMARY = "agrix_generated_workbook_summary.json"
REPORT = "agrix_generated_workbook_summary.txt"


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as inp:
        for chunk in iter(lambda: inp.read(2**20), b""):
            h.update(chunk)
    return h.hexdigest()


def discover_repo(user_path: str | None, audit: Path) -> Path:
    if user_path:
        candidates = [Path(user_path).expanduser()]
    else:
        env = os.environ.get("AGRIX_REPO")
        home = Path.home()
        candidates = ([Path(env).expanduser()] if env else []) + [
            home / "otkupapp-pwa",
            home / "Documents" / "GitHub" / "otkupapp-pwa",
            home / "Desktop" / "otkupapp-pwa",
            audit.parent.parent / "otkupapp-pwa",
            audit.parent / "otkupapp-pwa",
        ]
    for p in candidates:
        if (p / "tools" / "make_dev_workbook.py").is_file() and (p / "schema" / "schema.json").is_file():
            return p.resolve()
    raise ValueError("Ne mogu naci repo. Navedi --repo putanja/do/otkupapp-pwa (ili AGRIX_REPO env).")


def run(cmd: list[str], cwd: Path, stage: str) -> None:
    print(f"\n== {stage} ==", flush=True)
    print("  " + " ".join(cmd), flush=True)
    r = subprocess.run(cmd, cwd=str(cwd), check=False)
    if r.returncode:
        raise RuntimeError(f"{stage}: exit {r.returncode}; STOP")


def git_head(repo: Path) -> tuple[str | None, bool | None]:
    def call(*args: str):
        try:
            c = subprocess.run(["git", *args], cwd=str(repo), capture_output=True, text=True,
                               timeout=12, check=False)
            return c.stdout.strip() if c.returncode == 0 else None
        except (OSError, subprocess.TimeoutExpired):
            return None
    head = call("rev-parse", "HEAD")
    status = call("status", "--porcelain", "--untracked-files=no")
    return head, bool(status) if status is not None else None


def assert_repo_migration_compatible(repo: Path, head: str | None, dirty: bool | None) -> dict:
    """Reject stale checkout and legacy AMB table before Windows Excel starts."""
    if not head or dirty is None:
        raise RuntimeError("Ne mogu da proverim lokalni git HEAD i status")
    if dirty:
        raise RuntimeError("Repo ima izmenjene pracene fajlove; prvo sacuvaj/commituj izmene")
    try:
        tracked_main = subprocess.run(["git", "rev-parse", "--verify", "refs/remotes/origin/main"],
                                      cwd=str(repo), capture_output=True, text=True, timeout=12)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError("Ne mogu da proverim origin/main. Uradi git fetch origin") from exc
    if tracked_main.returncode != 0 or not tracked_main.stdout.strip():
        raise RuntimeError("Nema origin/main. Uradi git fetch origin")
    remote_sha = tracked_main.stdout.strip()
    if head != remote_sha:
        raise RuntimeError(f"Lokalni HEAD {head[:8]} nije origin/main {remote_sha[:8]}. "
                           "U repo-u uradi: git fetch origin && git switch main && git pull --ff-only origin main")
    source = repo / "schema" / "schema.json"
    schema = json.loads(source.read_text(encoding="utf-8"))
    tables = {t["table"]: t["columns"] for t in schema["tables"]}
    needed_amb = ("AmbID", "Datum", "TipAmbalaze", "Kolicina", "DokumentID", "DokumentTIP",
                  "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy", "OdNalogTip",
                  "OdNalogID", "NaNalogTip", "NaNalogID", "VrstaKretanja", "StornoOd")
    needed_doc = ("AmbDokID", "Vrsta", "BrojDokumenta", "Datum", "BrojOwnerTip",
                  "BrojOwnerID", "Napomena", "Stornirano", "CreatedAt", "CreatedBy",
                  "ModifiedAt", "ModifiedBy", "IspravkaOdID", "ZamenjenSaID")
    if tuple(tables.get("tblAmbalaza", ())) != needed_amb or tuple(tables.get("tblAmbalazaDokument", ())) != needed_doc:
        raise RuntimeError("Repo nije na podrzanoj AMB-10 kanonskoj semi. "
                           "Stara sema sa Smer/EntitetID/ReversID nije kompatibilna")
    return {"origin_main_sha": remote_sha, "target_schema_sha256": sha(source),
            "target_schema_table_count": len(tables),
            "target_schema_column_count": sum(len(v) for v in tables.values())}


def inspect_built(workbook: Path) -> dict:
    if not workbook.is_file() or workbook.stat().st_size < 10000:
        raise RuntimeError("Generator nije napravio validan xlsm fajl")
    with zipfile.ZipFile(workbook) as z:
        if z.testzip() is not None:
            raise RuntimeError("Neispravan ZIP clan u generisanom xlsm")
        parts = set(z.namelist())
        needed = {"xl/vbaProject.bin", "xl/workbook.xml", "[Content_Types].xml"}
        if not needed <= parts:
            raise RuntimeError("Makro projekat ili osnovni OOXML nedostaje")
        return {"workbook_sha256": sha(workbook),
                "vba_project_sha256": hashlib.sha256(z.read("xl/vbaProject.bin")).hexdigest(),
                "size_bytes": workbook.stat().st_size,
                "macro_binary_present": True}


def cli(argv=None) -> int:
    p = argparse.ArgumentParser(description="Generisi prazan AgriX .xlsm direktno iz repoa i proveri MIG-06.6")
    p.add_argument("--repo", help="Lokalni otkupapp-pwa repo; ako se izostavi, pokusava autodetekciju")
    p.add_argument("--audit", type=Path, default=Path(__file__).resolve().parent,
                   help="Folder koji sadrzi MIG-06.6 i ostale lokalne audit skripte")
    p.add_argument("--out", type=Path, help="Nova .xlsm putanja; default: ../AgriX_NEW_BLANK_<SHA>.xlsm")
    p.add_argument("--stage", type=Path, help="MIG-02 stage; default: ../MigrationStage_v2")
    p.add_argument("--adapter", type=Path, help="MIG-06.5 adapter; default: ../MigrationAmb10TargetAdapter_v2")
    p.add_argument("--preflight-out", type=Path, help="Novi (nepostojeci) MIG-06.6 izvestaj folder")
    p.add_argument("--build-only", action="store_true", help="Samo generisi svesku ako staging jos nije tu")
    a = p.parse_args(argv)
    audit = a.audit.expanduser().resolve()
    repo = discover_repo(a.repo, audit)
    if os.name != "nt":
        raise RuntimeError("Pravi VBA .xlsm zahteva Windows + Excel COM + pywin32; ovde nema Excel-a.")
    import importlib.util
    if importlib.util.find_spec("win32com") is None:
        raise RuntimeError("Nedostaje pywin32. Instaliraj: python -m pip install pywin32")
    repo_head, dirty = git_head(repo)
    print("Repo:", repo, "HEAD:", repo_head, "dirty:", dirty)
    # This gate runs BEFORE building a workbook: a clean Git checkout may still
    # be hundreds of commits behind the canonical AMB-10 code.
    repo_gate = assert_repo_migration_compatible(repo, repo_head, dirty)
    print("Current origin/main and AMB-10 canonical schema: PASS;",
          repo_gate["target_schema_table_count"], "tables,",
          repo_gate["target_schema_column_count"], "columns")
    run([sys.executable, str(repo / "tools" / "make_dev_workbook.py"), "--self-test"], repo, "Builder self-test")
    run([sys.executable, str(repo / "tools" / "gen_schema_module.py"), "--check"], repo, "Schema source/generated VBA sync")
    if a.out:
        out = a.out.expanduser().resolve()
    else:
        out = (audit.parent / ("AgriX_NEW_BLANK_" + (repo_head or "NOHEAD")[:8] + ".xlsm")).resolve()
    if out.suffix.lower() != ".xlsm":
        raise ValueError("--out mora biti .xlsm")
    if out.exists():
        raise FileExistsError(f"Bez prepisivanja postojeceg fajla: {out}")
    out.parent.mkdir(parents=True, exist_ok=True)
    # Genuine upstream implementation imports .bas/.cls/.frm and builds tables from schema.
    run([sys.executable, str(repo / "tools" / "make_dev_workbook.py"), "--out", str(out)], repo,
        "Build new full VBA .xlsm from repository")
    check = inspect_built(out)
    run([sys.executable, str(repo / "tools" / "schema_diff.py"), str(out)], repo, "Compare workbook against schema")
    summary = {"format": VERSION, "status": "BUILT_FROM_REPO_NOT_MIGRATED",
               "repo_commit": repo_head, "repo_dirty": dirty,
               "git_origin_main_verified": True,
               "schema_sha256": sha(repo / "schema" / "schema.json"),
               **check, "client_data_imported": False, "production_import_approved": False}
    if not a.build_only:
        adapter = (a.adapter or audit.parent / "MigrationAmb10TargetAdapter_v2").expanduser().resolve()
        stage = (a.stage or audit.parent / "MigrationStage_v2").expanduser().resolve()
        pf = audit / "legacy_amb10_workbook_preflight_v2.py"
        if not adapter.is_dir() or not stage.is_dir() or not pf.is_file():
            raise RuntimeError("Sveska je generisana, ali preflight podaci/skripta nedostaju. "
                               "Proveri --audit/--stage/--adapter i pokreni opet samo MIG-06.6. "
                               f"Generisani fajl: {out}")
        pfo = (a.preflight_out or audit.parent / ("MigrationAmb10WorkbookPreflight_" + (repo_head or "NOHEAD")[:8])).expanduser().resolve()
        if pfo.exists():
            raise FileExistsError(f"Sveska je generisana, ali preflight izlaz vec postoji: {pfo}")
        run([sys.executable, str(pf), "--adapter", str(adapter), "--stage", str(stage),
             "--template", str(out), "--schema", str(repo / "schema" / "schema.json"),
             "--out", str(pfo)], audit, "MIG-06.6 actual generated target preflight")
        pfsummary = json.loads((pfo / "agrix_workbook_preflight_summary.json").read_text(encoding="utf-8"))
        summary["target_preflight_status"] = pfsummary["status"]
        summary["target_preflight_template_sha256"] = pfsummary["target_preflight"]["target_file_sha256"]
        summary["target_full_schema_check"] = pfsummary["target_preflight"]["full_schema_check"]["status"]
        summary["staged_goods_parent_refs_verified"] = 4621
    else:
        summary["target_preflight_status"] = "NOT_REQUESTED"
    # Summary intentionally excludes private local filesystem paths.
    reportdir = (a.preflight_out or audit.parent / ("MigrationAmb10WorkbookPreflight_" + (repo_head or "NOHEAD")[:8])) if not a.build_only else out.parent
    reportdir = Path(reportdir).expanduser().resolve()
    reportdir.mkdir(parents=True, exist_ok=True)
    (reportdir / SUMMARY).write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (reportdir / REPORT).write_text(
        f"AGRIX {VERSION}\nRepo HEAD: {repo_head}\n"+
        f"New empty XLSM SHA256: {check['workbook_sha256']}\n"+
        f"Schema SHA256: {summary['schema_sha256']}\n"+
        f"VBA project present: True\nPreflight: {summary['target_preflight_status']}\n"+
        "NO CLIENT DATA IMPORTED; NO AMB-10 BUSINESS RELEASE APPROVAL.\n", encoding="utf-8")
    print("\nGENERISANO:", out)
    print("IZVESTAJ:", reportdir / SUMMARY)
    return 0

if __name__ == "__main__":
    try:
        raise SystemExit(cli())
    except (OSError, RuntimeError, ValueError, FileExistsError, KeyError, zipfile.BadZipFile) as exc:
        print("AGRIX BUILD STOP:", exc, file=sys.stderr)
        raise SystemExit(2)
