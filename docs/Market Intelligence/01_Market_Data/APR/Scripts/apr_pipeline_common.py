from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import requests

SCRIPTS_DIR = Path(__file__).resolve().parent
APR_DIR = SCRIPTS_DIR.parent
RAW_DIR = APR_DIR / "Raw Data"
CLEAN_DIR = APR_DIR / "Clean Data"
PROCESSED_DIR = APR_DIR / "Processed"
REPORTS_DIR = APR_DIR / "Reports"

COMPANIES_URL = "https://openapi.apr.gov.rs/api/opendata/companies"
FINANCIALS_URL = "https://openapi.apr.gov.rs/api/opendata/companies/financial-statements"
DEFAULT_ACTIVITY_CODES = ("1039", "4631")
MB_PATTERN = re.compile(r"^\d{8}$")


def ensure_directories() -> None:
    for directory in (RAW_DIR, CLEAN_DIR, PROCESSED_DIR, REPORTS_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def date_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def normalize_activity_code(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.zfill(4) if text.isdigit() else text


def normalize_company_id(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    text = text.zfill(8)
    return text if MB_PATTERN.fullmatch(text) else None


def request_json(url: str, timeout: int = 120, verify_tls: bool = True) -> dict[str, Any]:
    response = requests.get(url, timeout=timeout, verify=verify_tls)
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise ValueError(f"APR odgovor sa {url} nije JSON objekat.")
    podaci = payload.get("Podaci")
    if not isinstance(podaci, dict):
        raise ValueError(f"APR odgovor sa {url} nema očekivani objekat 'Podaci'.")
    return payload


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def latest_file(directory: Path, pattern: str) -> Path:
    matches = sorted(directory.glob(pattern), key=lambda item: item.stat().st_mtime, reverse=True)
    if not matches:
        raise FileNotFoundError(f"Nije pronađen fajl: {directory / pattern}")
    return matches[0]


def numeric_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype("string")
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")
