from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Pokreće kompletan APR market-intelligence pipeline.")
    parser.add_argument("--codes", nargs="+", default=["1039", "4631"])
    parser.add_argument("--insecure", action="store_true")
    parser.add_argument("--include-inactive", action="store_true")
    return parser.parse_args()


def run(script_name: str, *arguments: str) -> None:
    command = [sys.executable, str(SCRIPTS_DIR / script_name), *arguments]
    print("\n>", " ".join(command))
    subprocess.run(command, check=True, cwd=SCRIPTS_DIR)


def main() -> None:
    args = parse_args()
    insecure_args = ["--insecure"] if args.insecure else []

    run("01_extract_companies.py", "--codes", *args.codes, *insecure_args)
    run("02_enrich_financials.py", *insecure_args)
    run("03_clean_validate.py")

    analysis_args = ["--include-inactive"] if args.include_inactive else []
    run("04_analyze_market.py", *analysis_args)

    print("\nAPR pipeline je uspešno završen.")


if __name__ == "__main__":
    main()
