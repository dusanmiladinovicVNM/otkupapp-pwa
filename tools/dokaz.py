#!/usr/bin/env python3
"""Pun dvosmerni dokaz: svaka sabotaza iz kataloga mora da POKAZE crveno.

CLAUDE.md paragraf 5 trazi da se posle izmene pusti ceo dvosmerni dokaz i tvrdi
da je broj CRVENIH jednak broju sabotaza. Do sada je to bio rucni ritual --
skripta iz scratchpada, pa se u praksi vrteo samo podskup. Rezultat: deset sidara
je istrunulo neprimeceno, jedna sabotaza je prestala da obara ista, a "36 od 39"
se citalo kao zeleno (v. docs/engineering/postmortems/2026-08-verifikacija.md 10).

    python tools/dokaz.py                      # ceo katalog (satima!)
    python tools/dokaz.py modOtkupUI.bas       # samo sabotaze nad tim fajlom
    python tools/dokaz.py mreza-podnozje       # samo sabotaze sa tim prefiksom

TRI STVARI KOJE OVAJ ALAT TVRDI, a ne samo gleda:

1. BAZA MORA BITI ZELENA PRE PRVE MUTACIJE. Bez toga alat ne dokazuje "mutacija
   je izazvala crveno" nego samo "posle mutacije postoji crveno" -- a to nisu
   iste tvrdnje. Test koji vec pada iz nekog treceg razloga proglasio bi svaku
   sabotazu nad sobom dokazanom, ukljucujuci onu koja ne radi nista.

2. IZVOR SE VRACA I KAD RUN PUKNE. Mutacija namerno kvari radni izvor, pa
   ciscenje ide kroz `finally`: timeout Excela ili Ctrl+C usred prolaza inace
   ostavlja namerno pokvaren `src-vba/`. Posle svakog vracanja se poredi potpis
   celog `src-vba` -- ako se ne poklopi, dokaz STAJE odmah, jer bi sve merene
   posle toga islo nad pokvarenim kodom.

3. PAO JE BAS NJEN TEST. Sabotaza sme da obori i vise testova -- siroka izmena
   to i radi. Ne sme da ne obori svoj. Tekst tvrdnje u katalogu je dokumentacija,
   cesto parafraza, pa se razlika prijavljuje ali ne obara dokaz.

Banka-suite ne ispisuje ime testa uz pad, ali svaka njena tvrdnja nosi stabilan
prefiks ("T21 izabran placen blok: ..."), pa se identitet vadi iz njega. Bez toga
bi za te sabotaze tvrdnja bila samo "nesto je palo".

JEFTINA POLOVINA ovoga je `python tools/sabotaza.py --proveri-sidra` (ide i kroz
vba_check, dakle kroz PostToolUse hook): hvata zastarela sidra i pogresna imena
testova za sekundu. Ovaj alat je jedini koji zna da li sabotaza STVARNO nesto
obara -- i traje satima nad celim katalogom, pa se pusta nad onim sto je izmena
dirala.
"""
import argparse
import hashlib
import importlib.util
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
SUITE_BANKA = "RunBankaImportTestSuite"
SUITE_ALL = "RunAllTests"


def _katalog() -> dict:
    put = os.path.join(ROOT, "tools", "sabotaza.py")
    spec = importlib.util.spec_from_file_location("_sab_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul.SABOTAZE


def _otisak() -> str:
    """Potpis celog src-vba -- da se vidi da je posle svakog vracanja isti."""
    h = hashlib.sha256()
    for ime in sorted(os.listdir(SRC_VBA)):
        h.update(ime.encode())
        with open(os.path.join(SRC_VBA, ime), "rb") as fh:
            h.update(fh.read())
    return h.hexdigest()[:16]


def _pusti(*a, timeout=1200):
    return subprocess.run(a, cwd=ROOT, capture_output=True, text=True,
                          timeout=timeout)


def _suite_za(test: str) -> str:
    # Pisac ide u banka-suite: RunAllTests je nemutirajuca, pa tvrdnji o upisu
    # u njoj nema.
    return SUITE_ALL if test.startswith("T_") else SUITE_BANKA


def _tokeni_banke(izlaz: str) -> list:
    """Imena testova iz banka-suite izlaza, preko stabilnog prefiksa tvrdnje.

    ReportResults ispisuje "PAO   T21 izabran placen blok: ...", jer svaka
    tvrdnja u toj suite nosi `Const S As String = "T21 ..."`. Ime testa se ne
    ispisuje, ali se broj vadi -- a on je isti onaj iz T21_....
    """
    out = []
    for red in re.findall(r"^\s*PAO\s+(.*)$", izlaz, re.M):
        m = re.match(r"\s*(T\d+)\b", red)
        out.append((m.group(1) if m else "?", red))
    return out


def _pali(izlaz: str, suite: str) -> list:
    if suite == SUITE_ALL:
        return re.findall(r"^\s*FAIL (\S+) -- (.*)$", izlaz, re.M)
    return _tokeni_banke(izlaz)


def _kljuc_testa(test: str, suite: str) -> str:
    """Sta se poredi sa onim sto je palo."""
    if suite == SUITE_ALL:
        return test
    m = re.match(r"(T\d+)_", test)          # T21_IzabranPlacenBlok... -> T21
    return m.group(1) if m else test


def _baza_zelena(suite: str) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se."""
    try:
        r = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
    except subprocess.TimeoutExpired:
        return False, f"{suite}: timeout"
    izlaz = r.stdout + r.stderr

    if suite == SUITE_ALL:
        m = re.search(r"TESTS\s+(\d+) ukupno, (\d+) palo", izlaz)
        if not m:
            return False, f"{suite}: nema reda TESTS u izlazu"
        if int(m.group(2)) != 0:
            return False, f"{suite}: {m.group(0)}"
        return True, m.group(0)

    m = re.search(r"PASS=(\d+)\s+FAIL=(\d+)", izlaz)
    if not m:
        return False, f"{suite}: nema PASS/FAIL u izlazu"
    if int(m.group(2)) != 0:
        return False, f"{suite}: {m.group(0)}"
    return True, m.group(0)


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

    # --- KAPIJA: baza mora biti zelena pre prve mutacije --------------------
    potrebne = sorted({_suite_za(t) for _, _, t, _ in stavke} | {SUITE_ALL})
    for suite in potrebne:
        ok, opis = _baza_zelena(suite)
        print("BAZNO: %s" % opis, flush=True)
        if not ok:
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala.", file=sys.stderr)
            return 2

    pre = _otisak()
    print("potpis izvora: %s" % pre, flush=True)

    crvenih, lose = 0, []
    for ime, fajl, ocekTest, ocekTvrdnja in stavke:
        p = _pusti(sys.executable, "tools/sabotaza.py", ime)
        if p.returncode != 0:
            lose.append((ime, "APPLY-FAIL -- v. sabotaza.py --proveri-sidra"))
            print("%-46s APPLY-FAIL" % ime, flush=True)
            continue

        suite = _suite_za(ocekTest)
        pali, greska = [], ""
        try:
            run = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
            pali = _pali(run.stdout + run.stderr, suite)
        except subprocess.TimeoutExpired:
            greska = "TIMEOUT suite"
        except KeyboardInterrupt:
            greska = "PREKID"
        finally:
            # Izvor se vraca i kad je run pukao -- inace radni tree ostaje
            # namerno pokvaren.
            v = _pusti(sys.executable, "tools/sabotaza.py", "--vrati")

        sada = _otisak()
        if v.returncode != 0 or sada != pre:
            print("%-46s REVERT-FAIL (potpis %s)" % (ime, sada), flush=True)
            print("STOP: izvor nije vracen u pocetno stanje. Sve mereno posle "
                  "ovoga islo bi nad pokvarenim kodom.", file=sys.stderr)
            lose.append((ime, "REVERT-FAIL"))
            break

        if greska:
            lose.append((ime, greska))
            print("%-46s %s" % (ime, greska), flush=True)
            if greska == "PREKID":
                break
            continue

        if not pali:
            lose.append((ime, "NE OBARA NISTA"))
            stanje = "NE OBARA NISTA"
        else:
            crvenih += 1
            kljuc = _kljuc_testa(ocekTest, suite)
            imena = sorted({p0 for p0, _ in pali})
            poruke = " | ".join(p1 for _, p1 in pali)
            if kljuc not in imena:
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
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        # Poslednja mreza: prekid izmedju dve mutacije ne sme da ostavi
        # pokvaren izvor.
        subprocess.run([sys.executable, "tools/sabotaza.py", "--vrati"], cwd=ROOT)
        print("\nprekinuto -- izvor vracen", file=sys.stderr)
        sys.exit(130)
