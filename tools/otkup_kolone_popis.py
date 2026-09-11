"""Popis citalaca kolona tblOtkup koje ciljna sema BRISE.

Korak 7 refaktora (docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md S14.6) je podeljen po
MERENJU, ne po proceni: alat daje broj pojava po koloni i po modulu, pa se vidi
koliko posla pripada kom PR-u. Komentari se ne broje -- pokazivac na kolonu nije
citanje.

Cist Python: radi bez Excela, i u web sesiji.
"""
import collections
import io
import os
import re

ROOT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src-vba")

# Kolone koje po DOCUMENT_HEADER_LINES S4.1 na zaglavlju NE POSTOJE.
ODLAZE = {
    "COL_OTK_KOLICINA": "stavka",
    "COL_OTK_CENA": "stavka",
    "COL_OTK_KLASA": "stavka",
    "COL_OTK_KOL_AMB": "stavka",
    "COL_OTK_BRUTO_KG": "stavka",
    "COL_OTK_VOZAC": "otpremnica",
    "COL_OTK_ISPLACENO": "izvedeno",
    "COL_OTK_DATUM_ISPLATE": "izvedeno",
    "COL_OTK_BROJ_ZBIRNE": "labela tudjeg dokumenta (A2)",
    "COL_OTK_OTPREMNICA_ID": "clanstvo (A15)",
    "COL_OTK_BROJ_OTPREMNICE": "labela tudjeg dokumenta (A2)",
    "COL_OTK_NOVAC": "tblNovac",
    "COL_OTK_PRIMALAC": "tblNovac",
    "COL_GENERACIJA_ID": "identitet (zamenjuje ga OtkupID)",
}

PO_MODULU = collections.defaultdict(lambda: collections.defaultdict(int))
PO_KOLONI = collections.Counter()

for ime in sorted(os.listdir(ROOT)):
    if not ime.lower().endswith((".bas", ".cls", ".frm")):
        continue
    tekst = io.open(os.path.join(ROOT, ime), encoding="ascii",
                    errors="replace", newline="").read()
    # komentari se ne broje -- pokazivac na kolonu nije citanje
    kod = "\n".join(red.split("'")[0] for red in tekst.split("\r\n"))
    for kol in ODLAZE:
        n = len(re.findall(r"\b%s\b" % kol, kod))
        if n:
            PO_MODULU[ime][kol] = n
            PO_KOLONI[kol] += n

TEST = lambda m: m.lower().startswith("modtest") or "tests." in m.lower()

print("=== PO KOLONI ===")
for kol, n in PO_KOLONI.most_common():
    moduli = sorted(m for m in PO_MODULU if kol in PO_MODULU[m])
    prod = [m for m in moduli if not TEST(m)]
    print("%-28s %3d pojava | %2d modula (%d produkcionih) | %s"
          % (kol, n, len(moduli), len(prod), ODLAZE[kol]))

print()
print("=== PRODUKCIONI MODULI (bez testova) ===")
red = sorted(((sum(v.values()), m) for m, v in PO_MODULU.items() if not TEST(m)),
             reverse=True)
for n, m in red:
    print("%-34s %3d  %s" % (m, n, ", ".join(sorted(PO_MODULU[m]))))

print()
print("ukupno pojava: %d | produkcionih modula: %d | test modula: %d"
      % (sum(PO_KOLONI.values()),
         len([m for m in PO_MODULU if not TEST(m)]),
         len([m for m in PO_MODULU if TEST(m)])))
