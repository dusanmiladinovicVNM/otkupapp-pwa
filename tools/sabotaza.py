#!/usr/bin/env python3
"""Namerno kvarenje koda -- druga polovina dokaza iz CLAUDE.md paragraf 5.

Suite koja je zelena nad ispravnim kodom, a nije POKAZANA crvena nad pokvarenim,
ne dokazuje da isla sta meri (PR #181: cetiri puta zeleno-ali-nedokazano-crveno).
Za svaku proveru zato postoji sabotaza koja bas nju obara, po imenu.

    python tools/sabotaza.py --lista
    python tools/sabotaza.py clear-datum
    python tools/run_vba.py --suite RunAllTests      # ocekuj FAIL po imenu
    python tools/sabotaza.py --vrati

TRI ZAMKE koje su ovde vec pokupljene, da ih ne pokupi operater:

1. KRAJ REDA. `src-vba` se na Windows-u checkout-uje kao CRLF, a na Linuxu kao
   LF. Sidro sa zakucanim `\\n` ne pogodi nista, skripta tiho ne uradi nista, run
   prodje nad NEIZMENJENIM fajlom -- i izgleda kao da sabotaza "nije oborila"
   suite. Zato se kraj reda detektuje, a pogodak se TVRDI (tacno jednom).

2. UVLACENJE. Sidro se poredi od POCETKA REDA. Bez toga je
   `    mFrm...cbKupac.value = ""` (4 razmaka) podniz istog reda uvucenog za 8,
   pa je isto sidro pogadjalo dva razlicita mesta.

3. VRACANJE. `git checkout --` vraca fajl na HEAD, pa BRISE i nesnimljene izmene
   koje sa sabotazom nemaju veze (jednom vec pojelo test seam-ove). Zato se
   vraca obrnutom zamenom -- dira se tacno ono sto je i pokvareno.
"""

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")

# ime -> (fajl, sidro, zamena, test koji MORA da padne, sta tvrdnja kaze)
# Sidro i zamena se porede od POCETKA REDA (v. zamka 2) -- ne pisati vodece \n.
SABOTAZE = {
    # --- ParseDatum ---------------------------------------------------------
    "parse-tacka": (
        "modOtkupUI.bas",
        '    Do While Right$(t, 1) = "."\n'
        "        t = Left$(t, Len(t) - 1)\n"
        "    Loop\n",
        "    ' SABOTAZA: trailing tacka se vise ne skida\n",
        "T_ParseDatum_Ugovor",
        "trailing tacka se skida, ne obara unos",
    ),
    "parse-cdate": (
        "modOtkupUI.bas",
        "    If TryParseDateValue(t, d) Then ParseDatum = CDbl(d)\n",
        "    If IsDate(t) Then ParseDatum = CDbl(CDate(t))   ' SABOTAZA\n",
        "T_ParseDatum_Ugovor",
        "godina van poslovnog opsega",
    ),
    # --- ParcelaID ----------------------------------------------------------
    "parcela-tekst": (
        "modOtkupUI.bas",
        "    If CB.ListIndex >= 0 Then ParcelaID = Trim$(CStr(CB.List(CB.ListIndex, 1)))\n",
        "    If CB.ListIndex >= 0 Then ParcelaID = Trim$(CStr(CB.text))   ' SABOTAZA\n",
        "T_ParcelaID_IzSkriveneKolone",
        "ID parcele dolazi iz skrivene kolone, ne iz prikaznog teksta",
    ),
    "parcela-vidljivost": (
        "modOtkupUI.bas",
        '    If Not mFrm.Controls("zForm").Controls("fgParcela").Visible Then Exit Function\n',
        "    ' SABOTAZA: provera vidljivosti polja uklonjena\n",
        "T_ParcelaID_IzSkriveneKolone",
        "sakriveno polje ne salje parcelu u dokument",
    ),
    # --- ClearForm ----------------------------------------------------------
    "clear-datum": (
        "modOtkupUI.bas",
        '    If Not imaOtp Then SetDatumDanas mFrm.Controls("zForm")\n',
        '    SetDatumDanas mFrm.Controls("zForm")   \' SABOTAZA\n',
        "T_ClearForm_Ugovor",
        "dok je otpremnica aktivna datum se NE vraca na danas",
    ),
    "clear-zbirna": (
        "modOtkupUI.bas",
        '    nmv = Array("fgBrOtpr", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr")\n',
        '    nmv = Array("fgBrOtpr", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr", "fgBrZbir")\n',
        "T_ClearForm_Ugovor",
        "broj zbirne je kontekst -- ne brise se posle snimanja",
    ),
    "clear-partner": (
        "modOtkupUI.bas",
        '    mFrm.Controls("zCtx").Controls("cbKupac").value = ""\n',
        "    ' SABOTAZA: partner se vise ne brise\n",
        "T_ClearForm_Ugovor",
        "partner mora da bude obrisan posle snimanja",
    ),
    # --- kontekst otpremnice ------------------------------------------------
    "otp-izlaz-f1": (
        "modOtkupUI.bas",
        "    If modeKey(staraKey) = \"OTKUP\" And modeKey(key) <> \"OTKUP\" Then "
        "OtpustiOtpremnicu False\n",
        "    ' SABOTAZA: izlazak iz F1 vise ne otpusta otpremnicu\n",
        "T_OtpremnicaKontekst_PustaSeIzlaskomIzF1",
        "izlazak iz F1 otpusta otpremnicu",
    ),
    "otp-datum-rezim": (
        "modOtkupUI.bas",
        "    If staraKey <> key Then SetDatumPoRezimu\n",
        "    ' SABOTAZA: datum se vise ne racuna po rezimu\n",
        "T_OtpremnicaKontekst_PustaSeIzlaskomIzF1",
        "promena rezima bez otpremnice vraca datum na danas",
    ),
}
# NIJE ovde: otpustanje na izlasku sa EKRANA (ActivateScreen, Palete/Agrohemija).
# Rutinu OtpustiOtpremnicu pokriva otp-izlaz-f1, ali samo POZIVNO MESTO u
# SelectModeCore. Poziv iz ActivateScreen nema test -- trazio bi izgradjen
# zScr_PALETE ekran, a kad ScrBuild padne ActivateScreen izlazi PRE otpustanja,
# pa bi test padao iz tudjeg razloga. Ostaje na operaterskoj checklisti; sabotaza
# koja nema test ne ide u ovaj katalog jer bi razvodnila njegov ugovor.


def _procitaj(path: str) -> tuple[str, str]:
    """Sadrzaj sa LF krajevima + kraj reda kakav je zatecen na disku."""
    with open(path, "r", encoding="ascii", errors="strict", newline="") as fh:
        raw = fh.read()
    nl = "\r\n" if "\r\n" in raw else "\n"
    return raw.replace("\r\n", "\n"), nl


def _upisi(path: str, tekst: str, nl: str) -> None:
    with open(path, "w", encoding="ascii", newline="") as fh:
        fh.write(tekst.replace("\n", nl))


def _zameni(path: str, staro: str, novo: str) -> tuple[bool, int]:
    """Zameni sidro vezano za pocetak reda. Vraca (uspeh, broj pogodaka)."""
    tekst, nl = _procitaj(path)
    staro, novo = "\n" + staro, "\n" + novo      # zamka 2: sidro od pocetka reda
    pogodaka = tekst.count(staro)
    if pogodaka != 1:
        return False, pogodaka
    _upisi(path, tekst.replace(staro, novo), nl)
    return True, 1


def primeni(ime: str) -> int:
    fajl, staro, novo, test, tvrdnja = SABOTAZE[ime]
    path = os.path.join(SRC_VBA, fajl)

    ok, pogodaka = _zameni(path, staro, novo)
    if not ok:
        razlog = ("sabotaza je vec primenjena" if pogodaka == 0
                  else "sidro nije jednoznacno")
        print(f"sabotaza '{ime}': sidro nadjeno {pogodaka} puta u {fajl}, a mora "
              f"tacno jednom ({razlog}) -- proveri src-vba/{fajl} i sidro u "
              f"tools/sabotaza.py", file=sys.stderr)
        return 2

    print(f"sabotaza '{ime}' primenjena u src-vba/{fajl}")
    print(f"  ocekuj:  FAIL {test}")
    print(f"  tvrdnja: {tvrdnja}")
    print("  pokreni: python tools/run_vba.py --suite RunAllTests")
    print("  vrati:   python tools/sabotaza.py --vrati")
    return 0


def vrati() -> int:
    """Obrnuta zamena, ne git checkout (v. zamka 3)."""
    vraceno = []
    for ime, (fajl, staro, novo, _, _) in SABOTAZE.items():
        ok, _ = _zameni(os.path.join(SRC_VBA, fajl), novo, staro)
        if ok:
            vraceno.append(ime)

    if not vraceno:
        print("nema sta da se vrati -- nijedna sabotaza nije zatecena u src-vba/")
        return 0
    print("vraceno: " + ", ".join(vraceno))
    return 0


def lista() -> int:
    print("Sabotaze (svaka obara TACNO jedan test, po imenu):\n")
    sirina = max(len(k) for k in SABOTAZE)
    for ime, (_, _, _, test, tvrdnja) in SABOTAZE.items():
        print(f"  {ime.ljust(sirina)}  ->  FAIL {test}")
        print(f"  {' ' * sirina}      {tvrdnja}")
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Namerno kvarenje koda za dokaz u crvenom smeru")
    ap.add_argument("ime", nargs="?", help="koju sabotazu primeniti")
    ap.add_argument("--lista", action="store_true", help="ispisi sve sabotaze")
    ap.add_argument("--vrati", action="store_true", help="vrati sve zatecene sabotaze")
    args = ap.parse_args(argv)

    if args.lista:
        return lista()
    if args.vrati:
        return vrati()
    if not args.ime:
        ap.print_help()
        return 2
    if args.ime not in SABOTAZE:
        print(f"nepoznata sabotaza '{args.ime}'. Poznate: {', '.join(SABOTAZE)}",
              file=sys.stderr)
        return 2
    return primeni(args.ime)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
