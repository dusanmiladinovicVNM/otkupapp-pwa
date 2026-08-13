"""Cita test-log sheet iz sveske i grupise padove -- trijaza bez rucnog scroll-a.

Suite koje pisu log u sheet (za sada modBusinessFlowProTests ->
BUSINESS_FLOW_PRO_TEST_LOG) ostavljaju red po proveri: Timestamp | RunID | Kind |
Name | Status | Details | Operator. Kad padne 147 od 310, pitanje nije "koliko"
nego "da li se padovi grupisu" -- ako svi diraju istu temu, uzrok je fixture; ako
su razbacani po svim grupama, uzrok je proizvod.

Pokrece se nad temp kopijom koju ostavi `run_vba.py --keep`:

    python tools/run_vba.py --suite RunBusinessFlowProSuite --keep
    python tools/read_test_log.py "C:\\...\\Temp\\vbatest_xxxx\\otkup_test.xlsm"

Samo citanje -- sveska se ne snima. Windows + Excel + pywin32.
"""

import argparse
import collections
import os
import re
import sys

MSO_AUTOMATION_SECURITY_LOW = 1
DEFAULT_SHEET = "BUSINESS_FLOW_PRO_TEST_LOG"


def read_log(path: str, sheet_name: str):
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
        xl.EnableEvents = False        # Workbook_Open se ne pokrece

        wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)

        ws = None
        for sheet in wb.Worksheets:
            if str(sheet.Name).strip().lower() == sheet_name.strip().lower():
                ws = sheet
                break
        if ws is None:
            names = ", ".join(str(s.Name) for s in wb.Worksheets)
            print(f"Nema sheeta '{sheet_name}'. Postojeci: {names}", file=sys.stderr)
            return None

        last = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)   # xlUp
        if last < 2:
            print(f"Sheet '{sheet_name}' je prazan.", file=sys.stderr)
            return []

        block = ws.Range(ws.Cells(2, 1), ws.Cells(last, 6)).Value
        rows = []
        for r in block:
            rows.append({
                "run_id": "" if r[1] is None else str(r[1]),
                "kind": "" if r[2] is None else str(r[2]),
                "name": "" if r[3] is None else str(r[3]),
                "status": "" if r[4] is None else str(r[4]),
                "details": "" if r[5] is None else str(r[5]),
            })
        return rows
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


def group_key(name: str) -> str:
    """Gruba tema testa: Test_RF28_Xxx -> RF28, Test_HladnjacaChainYyy -> Hladnjaca."""
    n = re.sub(r"^Test_", "", name)
    m = re.match(r"(RF\d+)", n)
    if m:
        return m.group(1)
    m = re.match(r"([A-Z][a-z]+)", n)
    return m.group(1) if m else (n[:12] or "?")


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Trijaza test-loga iz sveske (samo citanje).")
    ap.add_argument("workbook", help="putanja do .xlsm (npr. temp kopija iz --keep)")
    ap.add_argument("--sheet", default=DEFAULT_SHEET, help=f"ime log sheeta (podr. {DEFAULT_SHEET})")
    ap.add_argument("--run", help="samo ovaj RunID (podrazumevano: poslednji u logu)")
    ap.add_argument("--all-details", action="store_true", help="ispisi Details za svaki pad")
    args = ap.parse_args(argv)

    rows = read_log(args.workbook, args.sheet)
    if rows is None:
        return 2
    if not rows:
        return 0

    run_id = args.run or rows[-1]["run_id"]
    rows = [r for r in rows if r["run_id"] == run_id]
    print(f"RunID: {run_id}   redova: {len(rows)}")

    fails = [r for r in rows if r["status"].strip().upper() in ("FAIL", "FATAL")]
    passes = [r for r in rows if r["status"].strip().upper() == "PASS"]
    print(f"PASS: {len(passes)}   FAIL: {len(fails)}")

    if not fails:
        return 0

    by_group = collections.Counter(group_key(r["name"]) for r in fails)
    pass_by_group = collections.Counter(group_key(r["name"]) for r in passes)

    print("\n--- padovi po temi (pad / ukupno) ---")
    for grp, cnt in by_group.most_common():
        total = cnt + pass_by_group.get(grp, 0)
        print(f"  {grp:22} {cnt:>4} / {total}")

    print("\n--- najcesci razlog (prvih 120 znakova Details) ---")
    reasons = collections.Counter(r["details"][:120] for r in fails)
    for reason, cnt in reasons.most_common(12):
        print(f"  {cnt:>4}x  {reason}")

    if args.all_details:
        print("\n--- svi padovi ---")
        for r in fails:
            print(f"  {r['name']}: {r['details'][:200]}")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
