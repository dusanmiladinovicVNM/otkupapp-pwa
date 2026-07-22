from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
APR_DIR = SCRIPTS_DIR.parent


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Pokreće APR market-intelligence pipeline.")
    parser.add_argument(
        "--mode",
        choices=("full", "offline", "report"),
        default="full",
        help=(
            "full: preuzimanje + finansije + validacija + analiza; "
            "offline: validacija i analiza postojećeg Excela; "
            "report: samo ponovna analiza najnovijeg validiranog fajla."
        ),
    )
    parser.add_argument("--codes", nargs="+", default=["1039", "4631"])
    parser.add_argument(
        "--input",
        type=Path,
        help="Excel za offline režim. Može biti apsolutna ili relativna putanja.",
    )
    tls_group = parser.add_mutually_exclusive_group()
    tls_group.add_argument("--insecure", action="store_true", help="Odmah gasi TLS proveru.")
    tls_group.add_argument(
        "--strict-tls",
        action="store_true",
        help="Zabranjuje automatski APR-only TLS fallback.",
    )
    parser.add_argument("--include-inactive", action="store_true")
    return parser.parse_args()


def run(script_name: str, *arguments: str) -> None:
    command = [sys.executable, str(SCRIPTS_DIR / script_name), *arguments]
    print("\n>", " ".join(command))
    subprocess.run(command, check=True, cwd=SCRIPTS_DIR)


def resolve_input(path: Path) -> Path:
    if path.is_absolute():
        return path

    candidates = [
        Path.cwd() / path,
        SCRIPTS_DIR / path,
        APR_DIR / path,
    ]
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.exists():
            return resolved

    return (Path.cwd() / path).resolve()


def main() -> None:
    args = parse_args()
    tls_args: list[str] = []
    if args.insecure:
        tls_args = ["--insecure"]
    elif args.strict_tls:
        tls_args = ["--strict-tls"]

    analysis_args = ["--include-inactive"] if args.include_inactive else []

    if args.mode == "full":
        run("01_extract_companies.py", "--codes", *args.codes, *tls_args)
        run("02_enrich_financials.py", *tls_args)
        run("03_clean_validate.py")
        run("04_analyze_market.py", *analysis_args)

    elif args.mode == "offline":
        if args.input is None:
            raise SystemExit("Za --mode offline obavezno navedi --input putanju do Excela.")
        input_path = resolve_input(args.input)
        if not input_path.exists():
            raise FileNotFoundError(f"Offline ulazni fajl ne postoji: {input_path}")
        run("03_clean_validate.py", "--input", str(input_path))
        run("04_analyze_market.py", *analysis_args)

    else:  # report
        run("04_analyze_market.py", *analysis_args)

    print(f"\nAPR pipeline je uspešno završen. Režim: {args.mode}")


if __name__ == "__main__":
    main()
