"""Ispis seme (.xlsm) -- SAMO citanje, sveska se nikad ne snima.

Batch/headless varijanta onoga sto modSetup.DebugKoloneTabele radi interaktivno
(InputBox + MsgBox, jedna tabela po pozivu). Sluzi kao ulaz za tools/make_fixture.py:
seme se ne pogadja iz koda nego se ocita iz sveske (CLAUDE.md S4 -- "Sema tabela je
izvor istine, ne kod").

Windows + Excel + pywin32. Konvencija otvaranja je ista kao u tools/run_vba.py:
EnableEvents=False (Workbook_Open / StartApp se NE pokrece), AutomationSecurity=1.

    python tools/dump_schema.py "C:/.../AgriX_2.28.4.xlsm"
    python tools/dump_schema.py <putanja> --json tests/schema_2.28.4.json
"""

import argparse
import json
import os
import sys

MSO_AUTOMATION_SECURITY_LOW = 1


def dump(path: str):
    try:
        import win32com.client as win32
    except ImportError:
        print("pywin32 nije instaliran: python -m pip install pywin32", file=sys.stderr)
        return None

    path = os.path.abspath(path)
    if not os.path.exists(path):
        print(f"Sveska ne postoji: {path}", file=sys.stderr)
        return None

    xl = win32.DispatchEx("Excel.Application")
    wb = None
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        xl.EnableEvents = False        # KLJUCNO: Workbook_Open (StartApp) se ne pokrece

        # ReadOnly + nikad Save -> original je netaknut i kad skripta pukne.
        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)

        result = {"workbook": path, "sheets": [], "tables": []}

        for ws in wb.Worksheets:
            result["sheets"].append({
                "name": str(ws.Name),
                "code_name": str(ws.CodeName),
                "visible": int(ws.Visible),
                "tables": int(ws.ListObjects.Count),
            })
            for lo in ws.ListObjects:
                headers = []
                try:
                    for col in lo.ListColumns:
                        headers.append(str(col.Name))
                except Exception as exc:            # tabela bez zaglavlja / ostecena
                    headers = [f"<greska: {exc}>"]
                try:
                    rows = int(lo.ListRows.Count)
                except Exception:
                    rows = -1
                result["tables"].append({
                    "table": str(lo.Name),
                    "sheet": str(ws.Name),
                    "code_name": str(ws.CodeName),
                    "rows": rows,
                    "headers": headers,
                })

        return result
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Ispis seme .xlsm (samo citanje).")
    ap.add_argument("workbook", help="putanja do .xlsm koja se cita")
    ap.add_argument("--json", help="upisi pun ispis kao JSON u ovaj fajl")
    args = ap.parse_args(argv)

    data = dump(args.workbook)
    if data is None:
        return 2

    sheets = data["sheets"]
    tables = sorted(data["tables"], key=lambda t: t["table"].lower())

    print(f"Sveska: {data['workbook']}")
    print(f"Sheetova: {len(sheets)}   Tabela: {len(tables)}")
    print()
    for t in tables:
        print(f"{t['table']}  [sheet={t['sheet']} code={t['code_name']} redova={t['rows']}]")
        print(f"    {' | '.join(t['headers'])}")
    print()

    empty = [t["table"] for t in tables if t["rows"] == 0]
    print(f"Praznih tabela: {len(empty)}")
    if empty:
        print("  " + ", ".join(sorted(empty)))

    if args.json:
        outpath = os.path.abspath(args.json)
        os.makedirs(os.path.dirname(outpath), exist_ok=True)
        with open(outpath, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2, sort_keys=True)
        print(f"JSON: {outpath}")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
