from __future__ import annotations

import hashlib
import json
import re
import warnings
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import pandas as pd
import requests
from requests.exceptions import SSLError
from urllib3.exceptions import InsecureRequestWarning

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
PIB_PATTERN = re.compile(r"^\d{9}$")
APR_HOST = "openapi.apr.gov.rs"


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


def normalize_pib(value: Any) -> str | None:
    """Normalizuje srpski PIB kao tekst od tačno devet cifara.

    PIB se namerno čuva kao string da Excel/pandas ne bi promenili format,
    odbacili vodeću nulu ili ga zapisali u naučnoj notaciji.
    """
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    text = re.sub(r"\D", "", text)
    if not text:
        return None
    text = text.zfill(9)
    return text if PIB_PATTERN.fullmatch(text) else None


def _validate_payload(url: str, response: requests.Response) -> dict[str, Any]:
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise ValueError(f"APR odgovor sa {url} nije JSON objekat.")
    podaci = payload.get("Podaci")
    if not isinstance(podaci, dict):
        raise ValueError(f"APR odgovor sa {url} nema očekivani objekat 'Podaci'.")
    return payload


def request_json(
    url: str,
    timeout: int = 120,
    *,
    force_insecure: bool = False,
    strict_tls: bool = False,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Preuzima APR JSON uz kontrolisanu TLS politiku.

    Podrazumevani režim prvo proverava TLS. Ako APR host vrati poznatu grešku
    validacije sertifikata, zahtev se ponavlja bez verifikacije, uz vidljivo
    upozorenje i metadata zapis. Fallback nikada nije dozvoljen za drugi host.
    `strict_tls=True` zabranjuje fallback. `force_insecure=True` ga preskače.
    """
    host = (urlparse(url).hostname or "").lower()
    if host != APR_HOST:
        raise ValueError(f"Nedozvoljen host za APR pipeline: {host or url}")
    if force_insecure and strict_tls:
        raise ValueError("Ne mogu zajedno --insecure i --strict-tls.")

    request_meta: dict[str, Any] = {
        "url": url,
        "host": host,
        "tls_mode": "strict" if strict_tls else "secure-first-with-apr-fallback",
        "tls_verified": not force_insecure,
        "tls_fallback_used": False,
    }

    if force_insecure:
        print("UPOZORENJE: APR zahtev se izvršava bez TLS verifikacije (--insecure).")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", InsecureRequestWarning)
            response = requests.get(url, timeout=timeout, verify=False)
        request_meta["tls_mode"] = "forced-insecure"
        request_meta["tls_verified"] = False
        return _validate_payload(url, response), request_meta

    try:
        response = requests.get(url, timeout=timeout, verify=True)
        return _validate_payload(url, response), request_meta
    except SSLError as exc:
        if strict_tls:
            raise RuntimeError(
                "APR TLS sertifikat nije mogao da se validira. "
                "Strict TLS režim zabranjuje fallback."
            ) from exc

        print(
            "UPOZORENJE: APR TLS sertifikat nije validiran. "
            "Ponavljam zahtev bez verifikacije samo za openapi.apr.gov.rs."
        )
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", InsecureRequestWarning)
            response = requests.get(url, timeout=timeout, verify=False)

        request_meta["tls_verified"] = False
        request_meta["tls_fallback_used"] = True
        request_meta["tls_error"] = str(exc)
        return _validate_payload(url, response), request_meta


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
