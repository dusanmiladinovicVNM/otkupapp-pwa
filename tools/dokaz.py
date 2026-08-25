#!/usr/bin/env python3
"""Pun dvosmerni dokaz: svaka sabotaza iz kataloga mora da POKAZE crveno.

CLAUDE.md paragraf 5 trazi da se posle izmene pusti ceo dvosmerni dokaz i tvrdi
da je broj CRVENIH jednak broju sabotaza. Do sada je to bio rucni ritual --
skripta iz scratchpada, pa se u praksi vrteo samo podskup. Rezultat: deset sidara
je istrunulo neprimeceno, jedna sabotaza je prestala da obara ista, a "36 od 39"
se citalo kao zeleno (v. docs/UI_MIGRACIJA_KATALOG.md 13.11).

    python tools/dokaz.py                      # ceo katalog (dugo!)
    python tools/dokaz.py modOtkupUI.bas       # samo sabotaze nad tim fajlom
    python tools/dokaz.py mreza-podnozje       # samo sabotaze sa tim prefiksom

Sta se tvrdi za svaku sabotazu:
  * primeni se i vrati se (izvor je na kraju BIT-IDENTICAN pocetnom);
  * suite pocrveni, i medju palima je BAS njen test.

Sabotaza sme da obori i vise testova -- siroka izmena to i radi. Ne sme da ne
obori svoj. Tekst tvrdnje u katalogu je dokumentacija, cesto parafraza, pa se
razlika prijavljuje ali ne obara dokaz.

JEFTINA POLOVINA ovoga je `python tools/sabotaza.py --proveri-sidra` (ide i kroz
vba_check): hvata zastarela sidra i pogresna imena testova za sekundu. Ovaj alat
je jedini koji zna da li sabotaza STVARNO nesto obara -- i traje satima nad celim
katalogom, pa se pusta nad onim sto je izmena dirala.
"""
import argparse
import hashlib
import importlib.util
import io
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")


def _katalog() -> dict:
    put = os.path.join(ROOT, "tools", "sabotaza.py")
    spec = importlib.util.spec_from_file_location("_sab_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul.SABOTAZE


def _otisak() -> str:
    """Potpis celog src-vba -- da se vidi da je posle svih vracanja isti."""
    h = hashlib.sha256()
    for ime in sorted(os.listdir(SRC_VBA)):
        h.update(ime.encode())
        with open(os.path.join(SRC_VBA, ime), "rb") as fh:
            h.update(fh.read())
    return h.hexdigest()[:16]


def _pusti(*a, timeout=1200):
    return subprocess.run(a, cwd=ROOT, capture_output=True, text=True,
                          timeout=timeout)


def _pali(izlaz: str, suite: str, ocekTest: str) -> list:
    if suite == "RunAllTests":
        return re.findall(r"^\s*FAIL (\S+) -- (.*)$", izlaz, re.M)
    # banka-suite ispisuje "PAO   <tvrdnja>" i nema ime testa u redu
    return [(ocekTest, m) for m in re.findall(r"^PAO\s+(.*)$", izlaz, re.M)]


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("filter", nargs="*",
                    help="ime fajla (modX.bas) ili prefiks imena sabotaze")
    args = ap.parse_args(argv)

    katalog = _katalog()
    stavke = []
    for ime, (fajl, _staro, _novo, test, tvrdnja) in katalog.items():
        if args.filter and not (fajl in args.filter or
                                any(ime.startswith(f) for f in args.filter)):
            continue
        stavke.append((ime, fajl, test, tvrdnja))

    if not stavke:
        print("filter ne pogadja nijednu sabotazu", file=sys.stderr)
        return 2

    print("sabotaza: %d" % len(stavke), flush=True)
    pre = _otisak()

    r = _pusti(sys.executable, "tools/run_vba.py", "--suite", "RunAllTests")
    m = re.search(r"TESTS.*", r.stdout + r.stderr)
    print("BAZNO: %s" % (m.group(0) if m else "?"), flush=True)

    crvenih, lose = 0, []
    for ime, fajl, ocekTest, ocekTvrdnja in stavke:
        p = _pusti(sys.executable, "tools/sabotaza.py", ime)
        if p.returncode != 0:
            lose.append((ime, "APPLY-FAIL -- v. sabotaza.py --proveri-sidra"))
            print("%-46s APPLY-FAIL" % ime, flush=True)
            continue

        # Pisac ide u banka-suite: RunAllTests je nemutirajuca, pa tvrdnji o
        # upisu u njoj nema.
        suite = ("RunAllTests" if ocekTest.startswith("T_")
                 else "RunBankaImportTestSuite")
        run = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
        pali = _pali(run.stdout + run.stderr, suite, ocekTest)

        v = _pusti(sys.executable, "tools/sabotaza.py", "--vrati")
        if v.returncode != 0:
            lose.append((ime, "REVERT-FAIL"))
            print("%-46s REVERT-FAIL" % ime, flush=True)
            continue

        if not pali:
            lose.append((ime, "NE OBARA NISTA"))
            stanje = "NE OBARA NISTA"
        else:
            crvenih += 1
            imena = sorted({p0 for p0, _ in pali})
            poruke = " | ".join(p1 for _, p1 in pali)
            if ocekTest not in imena:
                stanje = "NE OBARA SVOJ TEST, nego: " + ", ".join(imena)
                lose.append((ime, stanje))
            elif ocekTvrdnja and ocekTvrdnja.lower() not in poruke.lower():
                stanje = "OK (tekst tvrdnje u katalogu je parafraza)"
            elif len(imena) > 1:
                stanje = "OK (uz jos %d testa)" % (len(imena) - 1)
            else:
                stanje = "OK"
        print("%-46s %s" % (ime, stanje), flush=True)

    posle = _otisak()
    print("\ncrvenih: %d / sabotaza: %d" % (crvenih, len(stavke)))
    print("izvor pre/posle: %s / %s -> %s"
          % (pre, posle, "IDENTICAN" if pre == posle else "RAZLIKA!"))
    for ime, sta in lose:
        print(" PROBLEM: %s -> %s" % (ime, sta))
    ok = not lose and pre == posle and crvenih == len(stavke)
    print("=== %s ===" % ("DOKAZANO" if ok else "NIJE DOKAZANO"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
