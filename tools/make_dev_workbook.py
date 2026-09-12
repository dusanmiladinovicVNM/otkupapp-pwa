"""Napravi RADNU DEV svesku od nule: prazan .xlsm -> sav VBA -> sve tabele.

    python tools/make_dev_workbook.py --out C:\\putanja\\AgriX_DEV.xlsm

NIJE isto sto i `make_fixture.py`, i ne treba ih spajati:

    make_fixture.py       test fixture IZ DONOR sveske; STRIPUJE VBA, seje TEST
                          podatke (TST-*), pravi .sig potpis. Za `run_vba`.
    make_dev_workbook.py  radna sveska OD NULE; uvozi VBA, pravi tabele iz
                          kanona, BEZ ijednog poslovnog podatka. Za operatera.

Zasto postoji: `ImportAllVBA` radi merge nad ZATECENOM sveskom. Kad se VBA
projekat te sveske ostesti -- posle vise self-update ciklusa ili padova Excela --
merge pukne sa `AddFromString failed`, i pukne i ROLLBACK, tj. ne primi ni STARI
kod koji je do maloprde radio. Tada sveska nije popravljiva uvozom; treba nova.
Ovaj alat je pravi za par sekundi, iz gita, bez rucnog klikanja.

TRAZI Windows + Excel + pywin32 + "Trust access to the VBA project object model"
(isto kao `run_vba.py`). U web sesiji se ne izvrsava.

Redosled je bitan:

  1. prazna .xlsm, sa EnableEvents = False    Workbook_Open se NE sme okinuti nad
                                              sveskom koja jos nema tabele
  2. Import .bas / .cls / .frm                pravi NOVE komponente
  3. ThisWorkbook preko AddFromString         dokument-modul se ne moze Import-ovati
  4. EnsureAllTables                          tabele iz schema/schema.json
  5. tblLocalConfig + tblSEFConfig            licenca off, da app moze da se digne

Lista "Pregled listova" (dev dugmad: Pokreni program / Otvori VBA / Migracije /
Ocisti tabele / Uvezi VBA) NEMA u svezoj svesci i alat je NE pravi:
`NapraviPregledListova` zavrsava `MsgBox`-om, koji bi obesio automatski poziv, i
sam modul kaze da se pokrece rucno. Posle prvog otvaranja: Alt+F8 ->
`NapraviPregledListova`. Nije potrebna za rad -- `Workbook_Open` sam podize app.

Posle toga proveri i potvrdi:

    python tools/schema_diff.py <sveska>                      ocekivano: poklapaju se
    python tools/run_vba.py --workbook <sveska> --suite RunBusinessFlowProSuite

Druga komanda radi nad TEMP KOPIJOM, pa ne prlja isporucenu svesku. Ako suite
prodje, projekat se kompajlira -- suite se ne bi ni pokrenula da ne moze.

SHEET CODENAME: EnsureAllTables pravi listove sa Excel-ovim imenima (Sheet2,
Sheet3...), ne semantickim (sOtkup, sZbirna...). Zato `run_vba` prijavi
"SKIP N .doccls (nema komponente u svesci)". To je bezopasno i mereno:

  * ti .doccls fajlovi NEMAJU nijednu liniju koda, samo zaglavlje -- self-test
    ispod to i cuva, da tiho ne izgubimo kod ako ga neko tamo doda;
  * konstante SHT_OTKUP / SHT_NOVAC / SHT_KOOPERANTI / SHT_CONFIG / SHT_FAKTURE
    postoje u modConfig ali ih NIKO ne koristi (nula pojava u src-vba).
"""
import argparse
import io
import os
import shutil
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src-vba")
XL_OPENXML_MACRO = 52


# --- citanje izvora ----------------------------------------------------------

def telo_bez_zaglavlja(putanja):
    """Kod dokument-modula bez VERSION/BEGIN/Attribute zaglavlja.

    AddFromString prima samo telo; zaglavlje bi zavrsilo kao kod i oborilo
    kompajliranje.
    """
    s = io.open(putanja, encoding="ascii", errors="replace", newline="").read()
    redovi = s.split("\r\n")
    for i, r in enumerate(redovi):
        if r.startswith("Attribute VB_Exposed"):
            return "\r\n".join(redovi[i + 1:])
    return s


ZAGLAVLJE = ("VERSION", "BEGIN", "END", "Attribute", "MultiUse")


def linije_koda(putanja):
    """Linije koje su stvarno kod -- bez zaglavlja, komentara i praznih.

    Prefiks se poredi nad STRIP-ovanom linijom: zaglavlje .doccls-a je uvuceno
    ("  MultiUse = -1  'True"), pa je poredjenje nad sirovom linijom prijavljivalo
    po jednu "liniju koda" u svakom praznom listu.
    """
    s = io.open(putanja, encoding="ascii", errors="replace", newline="").read()
    out = []
    for r in s.split("\r\n"):
        t = r.strip()
        if not t or t.startswith("'") or t.startswith(ZAGLAVLJE):
            continue
        out.append(r)
    return out


# --- self-test (bez Excela) --------------------------------------------------

PRAZAN_DOCCLS = ("VERSION 1.0 CLASS\r\n"
                 "BEGIN\r\n"
                 "  MultiUse = -1  'True\r\n"
                 "END\r\n"
                 "Attribute VB_Name = \"sProba\"\r\n"
                 "Attribute VB_Exposed = True\r\n")

PUN_DOCCLS = PRAZAN_DOCCLS + ("Option Explicit\r\n"
                              "\r\n"
                              "' komentar se ne broji\r\n"
                              "Private Sub Worksheet_Change(ByVal Target As Range)\r\n"
                              "End Sub\r\n")


def _provera_mere(tmpdir):
    """Da li `linije_koda` uopste MERI -- ista funkcija koju zove self-test.

    Zelen self-test nad cistim repoom ne razlikuje "listovi su prazni" od
    "brojac uvek vraca nulu". Prva verzija je imala obrnutu gresku: poredila je
    prefiks nad neuvucenom linijom, pa je svaki prazan list izgledao kao da nosi
    kod. Zato oba smera idu kroz isti brojac.
    """
    nalazi = []
    for ime, sadrzaj, ocekivano_kod in (("prazan.doccls", PRAZAN_DOCCLS, False),
                                        ("pun.doccls", PUN_DOCCLS, True)):
        put = os.path.join(tmpdir, ime)
        with io.open(put, "w", encoding="ascii", newline="") as fh:
            fh.write(sadrzaj)
        n = len(linije_koda(put))
        if ocekivano_kod and n == 0:
            nalazi.append("brojac ne vidi kod u %s -- self-test nista ne meri" % ime)
        if not ocekivano_kod and n > 0:
            nalazi.append("brojac broji zaglavlje u %s kao kod (%d linija)" % (ime, n))
    return nalazi


def self_test():
    """Cuva pretpostavku na kojoj alat stoji: listovi nemaju kod.

    Ako neko doda kod u sOtkup.doccls, ovaj alat bi ga TIHO izgubio -- novi
    listovi imaju druge CodeName-ove, pa se taj modul nikad ne bi spojio.
    Bolje da padne ovde nego da se otkrije kao "dugme ne radi".
    """
    if not os.path.isdir(SRC):
        print("self-test: nema %s" % SRC, file=sys.stderr)
        return 2

    tmpdir = tempfile.mkdtemp(prefix="mkdevwb_")
    try:
        nalazi = _provera_mere(tmpdir)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    video_thisworkbook = False

    for f in sorted(os.listdir(SRC)):
        if not f.endswith(".doccls"):
            continue
        n = len(linije_koda(os.path.join(SRC, f)))
        if f == "ThisWorkbook.doccls":
            video_thisworkbook = True
            if n == 0:
                nalazi.append("ThisWorkbook.doccls je PRAZAN -- sveska bez Workbook_Open")
        elif n > 0:
            nalazi.append("%s ima %d linija koda -- ovaj alat bi ih izgubio "
                          "(list dobija drugi CodeName)" % (f, n))

    if not video_thisworkbook:
        nalazi.append("nema ThisWorkbook.doccls")

    if nalazi:
        print("make_dev_workbook --self-test: PALO", file=sys.stderr)
        for x in nalazi:
            print("  " + x, file=sys.stderr)
        return 2

    print("make_dev_workbook --self-test: cisto "
          "(brojac meri u oba smera; samo ThisWorkbook nosi kod)")
    return 0


# --- gradnja -----------------------------------------------------------------

def upisi_config(wb, tabela, parovi, kolona_kljuc, kolona_vrednost, aktivan=False):
    lo = None
    for ws in wb.Worksheets:
        for x in ws.ListObjects:
            if x.Name.lower() == tabela.lower():
                lo = x
                break
        if lo is not None:
            break
    if lo is None:
        return 0

    zaglavlje = {str(c.Name).strip().lower(): int(c.Index) for c in lo.ListColumns}
    n = 0
    for k, v in parovi.items():
        red = lo.ListRows.Add()
        red.Range.Cells(1, zaglavlje[kolona_kljuc.lower()]).Value = k
        red.Range.Cells(1, zaglavlje[kolona_vrednost.lower()]).Value = v
        if aktivan and "aktivan" in zaglavlje:
            red.Range.Cells(1, zaglavlje["aktivan"]).Value = "Aktivan"
        n += 1
    return n


def napravi(izlaz):
    import win32com.client as w32

    sys.path.insert(0, os.path.join(ROOT, "tools"))
    import make_fixture as mf            # LOCAL_CONFIG / SEF_CONFIG: licenca off

    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False

    wb = None
    try:
        wb = xl.Workbooks.Add()
        wb.SaveAs(izlaz, FileFormat=XL_OPENXML_MACRO)

        try:
            proj = wb.VBProject
        except Exception:
            print("VBProject nije dostupan -- ukljuci 'Trust access to the VBA "
                  "project object model' (File > Options > Trust Center > Macro "
                  "Settings).", file=sys.stderr)
            return 2

        uvezeno, pali = 0, []
        for f in sorted(os.listdir(SRC)):
            if not f.endswith((".bas", ".cls", ".frm")):
                continue
            try:
                proj.VBComponents.Import(os.path.join(SRC, f))
                uvezeno += 1
            except Exception as e:
                pali.append("%s -> %r" % (f, e))

        if pali:
            print("UVOZ PAO:", file=sys.stderr)
            for p in pali:
                print("  " + p, file=sys.stderr)
            return 2
        print("uvezeno komponenti: %d" % uvezeno)

        telo = telo_bez_zaglavlja(os.path.join(SRC, "ThisWorkbook.doccls"))
        cm = proj.VBComponents("ThisWorkbook").CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(telo)
        print("ThisWorkbook: %d linija" % cm.CountOfLines)

        wb.Save()

        t0 = time.time()
        xl.Run("EnsureAllTables")
        koliko = sum(ws.ListObjects.Count for ws in wb.Worksheets)
        print("EnsureAllTables: %d tabela za %.1fs" % (koliko, time.time() - t0))

        n1 = upisi_config(wb, "tblLocalConfig", mf.LOCAL_CONFIG, "Kljuc", "Vrednost")
        n2 = upisi_config(wb, "tblSEFConfig", mf.SEF_CONFIG, "ConfigKey",
                          "ConfigValue", aktivan=True)
        print("config: tblLocalConfig(%d), tblSEFConfig(%d) -- licenca off" % (n1, n2))

        wb.Save()
        print("\nGOTOVO: %s  (%.2f MB)"
              % (izlaz, os.path.getsize(izlaz) / 1024.0 / 1024.0))
        print("Sveska NEMA poslovne podatke -- maticne unesi kroz ekran "
              "'Maticni podaci'.")
        print("Nema jos lista 'Pregled listova' (dev dugmad: Pokreni program / "
              "Otvori VBA / Uvezi VBA):")
        print("  otvori svesku -> Alt+F8 -> NapraviPregledListova")
        return 0

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Radna DEV sveska od nule (VBA iz src-vba + tabele iz kanona).")
    ap.add_argument("--out", help="putanja do .xlsm koji se pravi")
    ap.add_argument("--force", action="store_true",
                    help="pregazi postojeci fajl")
    ap.add_argument("--self-test", action="store_true",
                    help="provera pretpostavki, bez Excela")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test()

    if not args.out:
        ap.error("--out je obavezan (ili koristi --self-test)")

    # Excel trazi backslash; "C:/x/y.xlsm" mu je nevalidna putanja.
    izlaz = os.path.abspath(args.out).replace("/", "\\")

    if os.path.exists(izlaz) and not args.force:
        print("Fajl vec postoji (--force da se pregazi): " + izlaz, file=sys.stderr)
        return 2
    if os.path.exists(izlaz):
        os.remove(izlaz)

    if os.name != "nt":
        print("Trazi Windows + Excel + pywin32.", file=sys.stderr)
        return 2

    return napravi(izlaz)


if __name__ == "__main__":
    sys.exit(main())
