"""Popis tabela i identitet dokumenata stare sveske -- SAMO CITA.

1) Tabele naspram kanona (schema/schema.json): koja fali, koja je visak, koje
   kolone fale/viskom, i da li se REDOSLED razilazi pre kraja (pozicioni hazard,
   isto pitanje kao tools/schema_diff.py, ali bez Excela).

2) Identitet dokumenta. Stari pisci (SaveOtkupMulti_TX, SaveOtpremnicaMulti_TX,
   SaveZbirnaMulti_TX, SavePrijemnicaMulti_TX, 4d3e2caa) pisu JEDAN RED PO KLASI,
   svaki sa svojim ID-jem. Main ima zaglavlje + stavke (DOCUMENT_HEADER_LINES.md).
   Ovde se meri koliko logickih dokumenata ima vise redova -- to je ulaz za
   odluku 1 (grupisati ili ne). Grupa je `GeneracijaID` kad ga red nosi
   (ZBR_IDENTITET.md: GeneracijaID je identitet, broj je labela), a inace
   heuristika vlasnik + datum + poslovni broj -- KANDIDAT, ne dokaz.
"""

import json
import os
from collections import Counter, defaultdict

from citac import datum, je_storno, kljuc, polje, tabela, tekst

KOREN = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
KANON = os.path.join(KOREN, "schema", "schema.json")

GRUPE = {
    # tabela: (PK, heuristicki kljuc kad nema GeneracijaID, poslovni broj)
    "tblOtkup": ("OtkupID", ("KooperantID", "StanicaID", "Datum", "BrojDokumenta"), "BrojDokumenta"),
    "tblOtpremnica": ("OtpremnicaID", ("StanicaID", "VozacID", "Datum", "BrojOtpremnice"), "BrojOtpremnice"),
    "tblZbirna": ("ZbirnaID", ("VozacID", "Datum", "BrojZbirne"), "BrojZbirne"),
    "tblPrijemnica": ("PrijemnicaID", ("KupacID", "Datum", "BrojPrijemnice"), "BrojPrijemnice"),
}
KANONSKE_STAVKE = ("tblOtkupStavke", "tblOtpremnicaStavke", "tblZbirnaStavke", "tblPrijemnicaStavke")


def ucitaj_kanon(putanja=KANON):
    with open(putanja, encoding="utf-8") as f:
        return {t["table"]: t["columns"] for t in json.load(f)["tables"]}


def popis_tabela(tabele, kanon):
    out = []
    for ime, kolone in kanon.items():
        t = tabela(tabele, ime)
        if t is None:
            out.append(dict(tabela=ime, u_svesci="NE", redova=0, fali="", visak="",
                            redosled="", napomena="nova u kanonu (nova sveska je pravi prazno)"))
            continue
        imena = [k for k in t.kolone]
        kf = [k.casefold() for k in imena]
        kanon_kf = [k.casefold() for k in kolone]
        fali = [k for k in kolone if k.casefold() not in kf]
        visak = [k for k in imena if k.casefold() not in kanon_kf]
        zajednicke_svesci = [k for k in kf if k in kanon_kf]
        zajednicke_kanonu = [k for k in kanon_kf if k in kf]
        redosled = "OK" if zajednicke_svesci == zajednicke_kanonu else "RAZLIKUJE SE"
        napomena = ""
        if visak:
            napomena = "kolone kojih vise nema u kanonu (stari model)"
        out.append(dict(tabela=ime, u_svesci="DA", redova=len(t.redovi), fali=", ".join(fali),
                        visak=", ".join(visak), redosled=redosled, napomena=napomena))
    for kljuc_t, t in sorted(tabele.items()):
        if t.ime not in kanon and kljuc_t not in {k.casefold() for k in kanon}:
            out.append(dict(tabela=t.ime, u_svesci="DA", redova=len(t.redovi), fali="",
                            visak=", ".join(t.kolone), redosled="",
                            napomena="NEMA JE U KANONU -- podaci bez cilja ako se ne odluci"))
    return out


def oblik_dokumenata(tabele):
    """'STARI' (red po klasi) ili 'KANONSKI' (zaglavlje + stavke sa redovima)."""
    for ime in KANONSKE_STAVKE:
        t = tabela(tabele, ime)
        if t is not None and t.redovi:
            return "KANONSKI"
    return "STARI"


def identitet_dokumenata(tabele):
    sazetak, grupe_vise = [], []
    for ime, (pk, heur, broj_kol) in GRUPE.items():
        t = tabela(tabele, ime)
        if t is None:
            continue
        ima_gen = t.ima_kolonu("GeneracijaID")
        ima_klasu = t.ima_kolonu("Klasa")
        grupe = defaultdict(list)
        bez_broja = 0
        for r in t.redovi:
            gen = tekst(polje(r, "GeneracijaID")) if ima_gen else ""
            if gen:
                k = ("GEN", kljuc(gen))
            elif tekst(polje(r, broj_kol)):
                k = ("HEUR",) + tuple(
                    (datum(polje(r, c)).isoformat() if datum(polje(r, c)) else kljuc(polje(r, c)))
                    if c == "Datum" else kljuc(polje(r, c)) for c in heur)
            else:
                bez_broja += 1
                k = ("RED", r["_redni"])
            grupe[k].append(r)
        vise = dupla = mesano = gen_grupa = 0
        for k, redovi in grupe.items():
            if k[0] == "GEN":
                gen_grupa += 1
            if len(redovi) < 2:
                continue
            vise += 1
            aktivni = [r for r in redovi if not je_storno(r)]
            klase = [kljuc(polje(r, "Klasa")) for r in aktivni] if ima_klasu else []
            ima_duplu = len(klase) != len(set(klase))
            ima_mesano = 0 < len(aktivni) < len(redovi)
            dupla += ima_duplu
            mesano += ima_mesano
            grupe_vise.append(dict(
                tabela=ime, osnov="GeneracijaID" if k[0] == "GEN" else "vlasnik+datum+broj",
                broj=tekst(polje(redovi[0], broj_kol)), datum=datum(polje(redovi[0], "Datum")),
                redova=len(redovi), aktivnih=len(aktivni),
                id=", ".join(tekst(polje(r, pk)) for r in redovi),
                klase=", ".join(tekst(polje(r, "Klasa")) for r in redovi) if ima_klasu else "",
                napomena="; ".join(x for x in (
                    "DUPLA KLASA medju aktivnim" if ima_duplu else "",
                    "deo redova storniran" if ima_mesano else "") if x)))
        sazetak.append(dict(
            tabela=ime, redova=len(t.redovi), aktivnih=sum(not je_storno(r) for r in t.redovi),
            logickih=len(grupe), sa_vise_redova=vise, po_generaciji=gen_grupa,
            dupla_klasa=dupla, mesano_storno=mesano, bez_broja=bez_broja,
            ima_generaciju="DA" if ima_gen else "NE"))
    return sazetak, grupe_vise


def brojevi_tabela(tabele):
    return Counter({t.ime: len(t.redovi) for t in tabele.values()})
