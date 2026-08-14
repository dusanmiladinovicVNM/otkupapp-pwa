"""Generise tests/fixtures/otkup_test.xlsm iz donor sveske.

Zasto donor a ne "od nule": osnovna sema (sheetovi + ListObject-i sa kolonama)
ne postoji nigde u kodu -- Ensure* rutine u modSetup samo DODAJU kolone na
postojece tabele, a spiskovi kolona osnovnih tabela zive iskljucivo u .xlsm.
Zakucavanje tih spiskova u Python napravilo bi drugi izvor istine koji konkurise
svesci (CLAUDE.md S4: "Sema tabela je izvor istine, ne kod"). Zato: struktura se
uzima iz donora, a podaci su 100% sinteticki.

Donor se NIKAD ne menja -- radi se nad kopijom.

Rezultat: sveska u kojoj su svi redovi obrisani (osim kataloga -- vidi KEEP_ROWS)
i posejani samo test unosi. Nijedan klijentski podatak ne moze da zavrsi u
tests/golden/*.txt koji idu u git.

    python tools/make_fixture.py --donor "C:/.../AgriX_2.28.4.xlsm"
    python tools/make_fixture.py --donor <put> --out tests/fixtures/otkup_test.xlsm --force

Windows + Excel + pywin32. Semu donora ispisuje tools/dump_schema.py.
"""

import argparse
import datetime
import os
import shutil
import sys

MSO_AUTOMATION_SECURITY_LOW = 1

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")

# Tabele kojima se redovi NE brisu: katalog poruka (Poruka("KLJUC") bez njega
# vraca prazno) i config tabele (iz njih citaju GetConfigValue/GetLocalConfigValue).
KEEP_ROWS = {"tblporuke", "tblsefconfig", "tblconfig", "tbllocalconfig"}

# FIKSAN datum, ne "danas": golden snapshot hvata txtDatum, pa bi fixture vezan
# za danasnji dan obarao golden fajlove svaki sledeci dan.
FIXTURE_DATE = datetime.date(2026, 3, 15)

STATUS_AKTIVAN = "Aktivan"          # modConfig.STATUS_AKTIVAN
AMB_12_1 = "12/1"                   # modConfig.AMB_12_1

STANICA = "STA-TEST-1"
VOZAC = "VOZ-TEST-1"
VRSTA = "TESTVOCE"
SORTA = "TESTSORTA"
ZBIRNA = "ZB-TEST-1"
ZBIRNA_U_BLOKU = "ZB-TEST-3"        # zbirnu nosi otkupni blok, ne otpremnica
# Kupac postoji SAMO kao ID na fakturi -- red u tblKupci ne treba: kapije koje
# ga koriste porede identifikatore, ne citaju karticu kupca.
KUPAC = "KUP-TEST-1"
FAKTURA = "FAK-TEST-1"
FAKTURA_IZNOS = 10000

# Sejanje ide PO IMENU KOLONE -- ako donor nema neku od ovih kolona, skripta
# pukne glasno umesto da tiho napravi fixture nad kojim testovi lazu.
SEED = {
    "tblStanice": [
        {"StanicaID": STANICA, "Naziv": "Test Otkupno Mesto", "Mesto": "Test Mesto",
         "Aktivan": STATUS_AKTIVAN, "JeHladnjaca": "NE"},
    ],
    "tblVozaci": [
        {"VozacID": VOZAC, "Ime": "Test", "Prezime": "Vozac",
         "Aktivan": STATUS_AKTIVAN, "KapacitetKG": 5000},
    ],
    "tblKulture": [
        {"KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "GajbicaPoPaleti": 100, "Aktivan": STATUS_AKTIVAN, "TipAmbalaze": AMB_12_1},
    ],
    "tblTipAmbalaze": [
        {"TipAmbalaze": AMB_12_1, "TezinaGajbiceKg": 1.0, "Aktivan": STATUS_AKTIVAN},
    ],
    "tblKooperanti": [
        {"KooperantID": "KOOP-TEST-1", "Ime": "Prvi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        {"KooperantID": "KOOP-TEST-2", "Ime": "Drugi", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
        {"KooperantID": "KOOP-TEST-3", "Ime": "Treci", "Prezime": "Testni",
         "Mesto": "Test Mesto", "StanicaID": STANICA, "Aktivan": STATUS_AKTIVAN},
    ],
    "tblParcele": [
        {"ParcelaID": "PAR-TEST-1", "KooperantID": "KOOP-TEST-1", "KatBroj": "1001",
         "KatOpstina": "Test Opstina", "Kultura": VRSTA, "PovrsinaHa": 1.5,
         "Aktivna": STATUS_AKTIVAN},
        {"ParcelaID": "PAR-TEST-2", "KooperantID": "KOOP-TEST-2", "KatBroj": "1002",
         "KatOpstina": "Test Opstina", "Kultura": VRSTA, "PovrsinaHa": 2.25,
         "Aktivna": STATUS_AKTIVAN},
    ],
    "tblZbirna": [
        {"ZbirnaID": "ZBI-TEST-1", "Datum": FIXTURE_DATE, "VozacID": VOZAC,
         "BrojZbirne": ZBIRNA, "VrstaVoca": VRSTA, "SortaVoca": SORTA,
         "UkupnoKolicina": 1000, "TipAmbalaze": AMB_12_1, "UkupnoAmbalaze": 100,
         "Klasa": "I"},
    ],
    # Tri slucaja koje zadatak trazi:
    #   OTP-TEST-1  datum iz proslosti + poznata zbirna + ostatak != 0 (1000 - 400)
    #   OTP-TEST-2  bez zbirne
    #   OTP-TEST-3  bez zbirne, ali blok u tblOtkup nosi zbirnu (ZB-TEST-3)
    "tblOtpremnica": [
        {"OtpremnicaID": "OTP-TEST-1", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "1/TEST", "BrojZbirne": ZBIRNA,
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 1000, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 100, "Klasa": "I"},
        {"OtpremnicaID": "OTP-TEST-2", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "2/TEST", "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 500, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 50, "Klasa": "I"},
        {"OtpremnicaID": "OTP-TEST-3", "Datum": FIXTURE_DATE, "StanicaID": STANICA,
         "VozacID": VOZAC, "BrojOtpremnice": "3/TEST", "BrojZbirne": "",
         "VrstaVoca": VRSTA, "SortaVoca": SORTA, "Kolicina": 800, "Cena": 50.0,
         "TipAmbalaze": AMB_12_1, "KolAmbalaze": 80, "Klasa": "I"},
    ],
    "tblOtkup": [
        {"OtkupID": "OTK-TEST-1", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-1",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 400, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 40, "VozacID": VOZAC, "BrojDokumenta": "1/TEST",
         "Klasa": "I", "BrojZbirne": ZBIRNA, "OtpremnicaID": "OTP-TEST-1",
         "BrojOtpremnice": "1/TEST", "ParcelaID": "PAR-TEST-1"},
        {"OtkupID": "OTK-TEST-2", "Datum": FIXTURE_DATE, "KooperantID": "KOOP-TEST-2",
         "StanicaID": STANICA, "KulturaID": "KUL-TEST-1", "VrstaVoca": VRSTA,
         "SortaVoca": SORTA, "Kolicina": 200, "Cena": 50.0, "TipAmbalaze": AMB_12_1,
         "KolAmbalaze": 20, "VozacID": VOZAC, "BrojDokumenta": "3/TEST",
         "Klasa": "I", "BrojZbirne": ZBIRNA_U_BLOKU, "OtpremnicaID": "OTP-TEST-3",
         "BrojOtpremnice": "3/TEST", "ParcelaID": "PAR-TEST-2"},
    ],
    # Jedna faktura, samo zato da kapija UplataFakturaProblem ima nad cim da
    # radi: vlasnistvo (KupacID), trenutni preostali iznos (Iznos - uplate) i
    # razlika "postoji / ne postoji". Namerno samo tri kolone -- sejanje ide PO
    # IMENU, pa svaka dodatna kolona koju donor nema obara generator.
    "tblFakture": [
        {"FakturaID": FAKTURA, "KupacID": KUPAC, "Iznos": FAKTURA_IZNOS},
    ],
}

# tblLocalConfig (Kljuc | Vrednost | Opis)
LOCAL_CONFIG = {
    "APP_SETUP_COMPLETED": "DA",
}

# tblSEFConfig -- licenca off. LICENSE_ENABLED=NO nije dovoljno: modLicense ima
# LATCH (vidi modLicense.bas:21) -- gate radi i bez YES ako postoje LICENSE_KEY i
# LICENSE_BOUND_PARTS. Zato se ti kljucevi prazne.
SEF_CONFIG = {
    "LICENSE_ENABLED": "NO",
    "LICENSE_KEY": "",
    "LICENSE_TOKEN": "",
    "LICENSE_BOUND_PARTS": "",
    "LICENSE_NEXT_CHECK": "",
    "LICENSE_STATUS": "",
    "LICENSE_HWM": "",
}

# DEFAULT_VRSTA_VOCA / DEFAULT_SORTA_VOCA se NAMERNO ne postavljaju: bez njih
# ApplyDefaultProizvod ostavlja combo-e prazne (frmOtkup ga zove pod
# On Error Resume Next), pa Initialize ne okida auto-cenu i stanje forme je
# deterministicno za golden snapshot.

EXCEL_EPOCH = datetime.date(1899, 12, 30)


class SchemaError(Exception):
    pass


def xl_serial(d: datetime.date) -> int:
    """Excel serijski broj datuma -- bez vremenske zone, za razliku od datetime preko COM-a."""
    return (d - EXCEL_EPOCH).days


def iter_tables(wb):
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            yield lo


def find_table(wb, name: str):
    target = name.strip().lower()
    for lo in iter_tables(wb):
        if str(lo.Name).strip().lower() == target:
            return lo
    return None


def header_index(lo) -> dict:
    return {str(c.Name).strip().lower(): int(c.Index) for c in lo.ListColumns}


def strip_vba(wb) -> list:
    """Izbaci SAV standardni/klasni/form kod iz donora.

    Kod u fixture-u je balast: run_vba.py na svakom pokretanju uveze svez
    src-vba/ preko njega. Ali uvozi samo ono sto repo IMA -- modul zaostao iz
    starijeg donora ostaje i izvrsava se. Ako nosi Public ime koje postoji i u
    svezem kodu, VBA to vidi kao "Ambiguous name" i odbija da pokrene makro iz
    njega, uz poruku "Cannot run the macro" koja ne lici na compile gresku.
    Tako je TestLicense_All bio mrtav dok je vba_check bio uredno zelen --
    duplikat nije bio u repou nego u svesci.

    Document moduli (listovi, ThisWorkbook) se NE mogu ukloniti; njihov kod
    run_vba merge-uje iz .doccls fajlova.

    Trazi "Trust access to the VBA project object model".
    """
    STD, CLS, FRM = 1, 2, 3
    removed = []
    try:
        proj = wb.VBProject
        comps = [c for c in proj.VBComponents]      # snapshot: brisemo iz kolekcije
    except Exception as exc:
        raise SchemaError(
            f"nema pristupa VBA projektu ({exc}). Ukljuci: File > Options > "
            "Trust Center > Trust Center Settings > Macro Settings > "
            "'Trust access to the VBA project object model'")

    for comp in comps:
        try:
            if int(comp.Type) not in (STD, CLS, FRM):
                continue
            name = str(comp.Name)
            proj.VBComponents.Remove(comp)
            removed.append(name)
        except Exception:
            pass                                    # zakljucan projekat/komponenta
    return sorted(removed)


def strip_rows(wb) -> list:
    cleared = []
    for lo in iter_tables(wb):
        name = str(lo.Name)
        if name.strip().lower() in KEEP_ROWS:
            continue
        try:
            if int(lo.ListRows.Count) > 0:
                n = int(lo.ListRows.Count)
                lo.DataBodyRange.Delete()
                cleared.append((name, n))
        except Exception as exc:
            raise SchemaError(f"{name}: brisanje redova nije uspelo ({exc})")
    return cleared


def add_row(lo, values: dict, table_name: str) -> None:
    idx = header_index(lo)
    missing = [k for k in values if k.strip().lower() not in idx]
    if missing:
        raise SchemaError(
            f"{table_name}: donor nema kolone {missing}. "
            f"Postojece: {sorted(idx)}"
        )
    row = lo.ListRows.Add()
    for key, val in values.items():
        cell = row.Range.Cells(1, idx[key.strip().lower()])
        if isinstance(val, datetime.date):
            cell.NumberFormat = "dd.mm.yyyy"
            cell.Value = xl_serial(val)
        else:
            cell.Value = val


def upsert_config(wb, table_name: str, pairs: dict,
                  key_col: str = "Kljuc", val_col: str = "Vrednost") -> int:
    """Kljuc/vrednost tabele: postojeci kljuc se azurira, novi se dodaje.

    Imena kolona nisu ista svuda -- tblConfig/tblLocalConfig imaju Kljuc|Vrednost,
    a tblSEFConfig ConfigKey|ConfigValue.
    """
    lo = find_table(wb, table_name)
    if lo is None:
        raise SchemaError(f"{table_name} ne postoji u donoru")
    idx = header_index(lo)
    for needed in (key_col, val_col):
        if needed.strip().lower() not in idx:
            raise SchemaError(f"{table_name}: nema kolonu '{needed}' ({sorted(idx)})")
    kcol, vcol = idx[key_col.strip().lower()], idx[val_col.strip().lower()]
    akt = idx.get("aktivan")          # tblSEFConfig ima Aktivan; nov red mora biti aktivan

    existing = {}
    if int(lo.ListRows.Count) > 0:
        for r in range(1, int(lo.ListRows.Count) + 1):
            key = lo.ListRows(r).Range.Cells(1, kcol).Value
            if key is not None:
                existing[str(key).strip().upper()] = r

    for key, val in pairs.items():
        r = existing.get(key.strip().upper())
        if r is None:
            row = lo.ListRows.Add()
            row.Range.Cells(1, kcol).Value = key
            row.Range.Cells(1, vcol).Value = val
            if akt:
                row.Range.Cells(1, akt).Value = STATUS_AKTIVAN
        else:
            lo.ListRows(r).Range.Cells(1, vcol).Value = val
    return len(pairs)


def build(donor: str, out: str, force: bool) -> int:
    try:
        import win32com.client as win32
    except ImportError:
        print("pywin32 nije instaliran: python -m pip install pywin32", file=sys.stderr)
        return 2

    donor = os.path.abspath(donor)
    out = os.path.abspath(out)

    if not os.path.exists(donor):
        print(f"Donor ne postoji: {donor}", file=sys.stderr)
        return 2
    if os.path.normcase(donor) == os.path.normcase(out):
        print("Donor i izlaz su ista putanja -- odbijam (donor se ne dira).", file=sys.stderr)
        return 2
    if os.path.exists(out) and not force:
        print(f"Izlaz vec postoji: {out}\nDodaj --force da ga prepisem.", file=sys.stderr)
        return 2

    os.makedirs(os.path.dirname(out), exist_ok=True)
    shutil.copy2(donor, out)          # radi se nad kopijom; donor ostaje netaknut

    xl = win32.DispatchEx("Excel.Application")
    wb = None
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        xl.EnableEvents = False       # KLJUCNO: Workbook_Open (StartApp) se ne pokrece

        wb = xl.Workbooks.Open(out, UpdateLinks=0)

        stripped = strip_vba(wb)
        if stripped:
            print(f"Uklonjeno {len(stripped)} VBA modula iz donora: "
                  + ", ".join(stripped[:8])
                  + (f" ... (+{len(stripped) - 8})" if len(stripped) > 8 else ""))

        cleared = strip_rows(wb)
        print(f"Obrisani redovi u {len(cleared)} tabela"
              + (": " + ", ".join(f"{n}({c})" for n, c in cleared) if cleared else ""))

        seeded = []
        for table_name, rows in SEED.items():
            lo = find_table(wb, table_name)
            if lo is None:
                raise SchemaError(f"{table_name} ne postoji u donoru")
            for values in rows:
                add_row(lo, values, table_name)
            seeded.append((table_name, len(rows)))
        print("Posejano: " + ", ".join(f"{n}({c})" for n, c in seeded))

        upsert_config(wb, "tblLocalConfig", LOCAL_CONFIG)
        upsert_config(wb, "tblSEFConfig", SEF_CONFIG,
                      key_col="ConfigKey", val_col="ConfigValue")
        print(f"Config: tblLocalConfig({len(LOCAL_CONFIG)}), tblSEFConfig({len(SEF_CONFIG)}) -- licenca OFF")

        wb.Save()
        print(f"\nFixture: {out}")
        return 0
    except SchemaError as exc:
        print(f"\nSEMA: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"\nGRESKA: {exc}", file=sys.stderr)
        return 2
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        xl.Quit()


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Pravi tests/fixtures/otkup_test.xlsm iz donor sveske.")
    ap.add_argument("--donor", required=True, help="putanja do .xlsm koja daje semu (ne menja se)")
    ap.add_argument("--out", default=DEFAULT_OUT, help=f"izlaz (podrazumevano {DEFAULT_OUT})")
    ap.add_argument("--force", action="store_true", help="prepisi postojeci izlaz")
    args = ap.parse_args(argv)
    return build(args.donor, args.out, args.force)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
