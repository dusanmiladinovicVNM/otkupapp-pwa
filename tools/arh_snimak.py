"""Snimak arhitektonskih mera nad src-vba -- brojke koje citira plan.

    python3 tools/arh_snimak.py                              # na stdout
    python3 tools/arh_snimak.py --out docs/Architecture/ARH_SNIMAK.md

Zasto postoji: `docs/Architecture/ARHITEKTURA_PLAN_OCENA.md` je jednom vec
zastareo za jedan dan -- repo je za 166 commita obrisao legacy forme i dodao
dva nova pisca poslovnim tabelama, a dokument je i dalje tvrdio suprotno, i to
bas na dve tacke koje vode odluku. Rucno odrzavanje tih brojki je traka za
trcanje.

Podela koja iz toga sledi: zakljucci ostaju u planu i pisu se rukom, BROJKE se
generisu ovde. Plan tvrdi "sta ovo znaci", snimak tvrdi "koliko ih ima danas".

Meri se samo ono sto plan zaista citira. Sve mere su staticke (nema Excela),
pa rade i u Linux sesiji.

JEDINICA: `TBL_`/`COL_` se broje kao POJAVE, ne kao linije koje ih sadrze. Rucni
`grep -c` daje linije i zato uvek manji broj (jedna linija zna da nosi i tabelu
i tri kolone). Raniji rucni nalazi u planu su bili linijski; ovde je pojava, jer
posao nizvodno nije "obidji linije" nego "zameni svaku referencu".
"""

from __future__ import annotations

import argparse
import collections
import glob
import os
import re
import subprocess
import sys
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src-vba")

# `AppendRow`/`UpdateCell` su Function, pa se zovu i sa zagradom i bez nje.
# Izraz koji trazi razmak posle imena gutao je 43% mesta upisa -- ista greska je
# godinu dana stajala u who_writes.py i pravila skoro prazan signal `direct`.
UPIS_RE = re.compile(r"\b(?:AppendRow|UpdateCell)\s*\(?\s*(TBL_\w+)", re.I)
TEST_RE = re.compile(r"(Test|Tests|HealthCheck|E2E)", re.I)
PROC_SPLIT = r"^(?=\s*(?:Public |Private |Friend )?(?:Sub|Function)\s+\w+)"


def _izvori(*ext: str) -> list[str]:
    out: list[str] = []
    for e in ext:
        out += sorted(glob.glob(os.path.join(SRC, f"*{e}")))
    return out


def _tekst(p: str) -> str:
    return open(p, encoding="latin-1").read()


def _ime(p: str) -> str:
    return os.path.basename(p).rsplit(".", 1)[0]


def pisci_po_tabeli() -> dict[str, set[str]]:
    """Moduli koji FIZICKI pisu tabelu. Ne racuna AddTableSnapshot: on kaze
    'ova operacija menja tabelu u svojoj transakciji', sto je druga tvrdnja."""
    d: dict[str, set[str]] = collections.defaultdict(set)
    for p in _izvori(".bas", ".cls", ".frm"):
        n = _ime(p)
        if TEST_RE.search(n):
            continue
        for m in UPIS_RE.finditer(_tekst(p)):
            d[m.group(1)].add(n)
    return d


def tx_klasifikacija() -> tuple[int, int, int, int, int]:
    """(ukupno, test, produkcija, omotaca, Public blizanaca).

    `_TX` oznacava vlasnistvo nad transaction boundary-jem, ne poslovni
    use-case: omotac je `X_TX` koji samo otvara transakciju i zove `X`."""
    svi = _izvori(".bas")
    src = "".join(_tekst(p) for p in svi)
    tx, test = set(), 0
    for p in svi:
        for m in re.finditer(r"^(?:Public |Private )?(?:Sub|Function) +(\w+_TX)\b",
                             _tekst(p), re.M):
            if m.group(1).startswith("Test"):
                test += 1
            else:
                tx.add(m.group(1))
    om = pub = 0
    for t in tx:
        sig = re.search(r"^((?:Public |Private )?(?:Sub|Function)) +"
                        + re.escape(t[:-3]) + r"\b", src, re.M)
        if sig:
            om += 1
            if not sig.group(1).startswith("Private"):
                pub += 1
    return len(tx) + test, test, len(tx), om, pub


def tx_invarijanta() -> tuple[int, int, int]:
    """(BeginTx, i snapshot u istoj proceduri, snapshot bez BeginTx).

    Treci broj mora ostati 0 -- to je invarijanta iz ADR-0003 tacke B."""
    nb = ns = sam = 0
    for p in _izvori(".bas", ".frm"):
        if TEST_RE.search(_ime(p)):
            continue
        for b in re.split(PROC_SPLIT, _tekst(p), flags=re.I | re.M):
            hb, hs = "BeginTx" in b, "AddTableSnapshot" in b
            nb += hb
            ns += hb and hs
            sam += hs and not hb
    return nb, ns, sam


def ekrani() -> list[tuple[str, int, int, int, int]]:
    """(modul, LOC, TBL_, COL_, upis) za sloj prikaza."""
    out = []
    for p in _izvori(".bas"):
        n = _ime(p)
        if not (n.startswith("modScr") or n == "modOtkupUI"):
            continue
        t = _tekst(p)
        out.append((n, t.count("\n"),
                    len(re.findall(r"TBL_[A-Z_]+", t)),
                    len(re.findall(r"COL_[A-Z_]+", t)),
                    len(re.findall(r"\b(?:AppendRow|UpdateCell|GetNextID)\s*[( ]", t))))
    return sorted(out, key=lambda r: -(r[2] + r[3]))


def obim() -> dict[str, int]:
    bas = _izvori(".bas")
    return {
        "loc": sum(_tekst(p).count("\n") for p in _izvori(".bas", ".cls", ".frm", ".doccls")),
        "moduli": len(_izvori(".bas", ".cls", ".frm", ".doccls")),
        "forme": len(_izvori(".frm")),
        "klase": len(_izvori(".cls")),
        "public": sum(len(re.findall(r"^Public (?:Sub|Function|Const|Type|Enum) +\w+",
                                     _tekst(p), re.M)) for p in bas),
        "implements": sum(len(re.findall(r"^Implements ", _tekst(p), re.M))
                          for p in _izvori(".bas", ".cls", ".frm")),
        "synccontrol": sum(1 for p in bas if "synccontrol" in _tekst(p).lower()),
    }


def _commit() -> str:
    try:
        return subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT,
                              capture_output=True, text=True, check=True).stdout.strip()
    except Exception:
        return "(nepoznat)"


def izvestaj() -> str:
    o = obim()
    uk, test, prod, om, pub = tx_klasifikacija()
    nb, ns, sam = tx_invarijanta()
    pisci = pisci_po_tabeli()
    vise = sorted(((t, s) for t, s in pisci.items() if len(s) > 1),
                  key=lambda x: (-len(x[1]), x[0]))
    ekr = ekrani()
    baseline = sum(r[2] + r[3] for r in ekr)

    L = [
        "# Snimak arhitektonskih mera",
        "",
        "> **Generisan fajl -- ne menjaj rukom.**",
        "> `python3 tools/arh_snimak.py --out docs/Architecture/ARH_SNIMAK.md`",
        "",
        f"Mereno: **{date.today().isoformat()}**, commit **{_commit()}**.",
        "",
        "Zakljucci i obrazlozenja su u `ARHITEKTURA_PLAN_OCENA.md` -- ovde su samo",
        "brojke. Kad se razidju, vazi ovaj fajl: plan se pise rukom, snimak se meri.",
        "",
        "## Obim",
        "",
        "| Mera | Vrednost |",
        "|---|---|",
        f"| Linija u `src-vba/` | {o['loc']:,} |".replace(",", "."),
        f"| Fajlova | {o['moduli']} |",
        f"| Formi (`.frm`) | {o['forme']} |",
        f"| Klasa (`.cls`) | {o['klase']} |",
        f"| `Public` simbola u `.bas` | {o['public']:,} |".replace(",", "."),
        f"| `Implements` | {o['implements']} |",
        f"| Modula koji pominju `SyncControl` | {o['synccontrol']} |",
        "",
        "## Fizicki pisci po tabeli",
        "",
        "Moduli koji zovu `AppendRow`/`UpdateCell` nad tabelom. Ovo je metrika",
        "Repository faze -- ne broj modula koji tabelu poslovno menjaju.",
        "",
    ]
    if vise:
        L += ["| Tabela | Pisaca | Moduli |", "|---|---|---|"]
        L += [f"| `{t}` | **{len(s)}** | " + ", ".join(f"`{m}`" for m in sorted(s)) + " |"
              for t, s in vise]
    else:
        L.append("Nijedna tabela nema vise od jednog fizickog pisca.")
    L += [
        "",
        f"**Tabela sa vise od jednog pisca: {len(vise)} od {len(pisci)}.**",
        "",
        "## Sloj prikaza",
        "",
        "`upis` mora ostati 0 u svakom redu -- to nadgleda `SLOJ_UPIS` u `vba_check`.",
        "`TBL_`/`COL_` je preostali dug: ekran zna imena kolona. Broje se POJAVE,",
        "ne linije -- jedna linija zna da nosi tabelu i tri kolone.",
        "",
        "| Modul | LOC | `TBL_` | `COL_` | upis |",
        "|---|---|---|---|---|",
    ]
    L += [f"| `{n}` | {loc} | {tb} | {co} | {up} |" for n, loc, tb, co, up in ekr]
    L += [
        "",
        f"**Zbir `TBL_`+`COL_` u sloju prikaza: {baseline}.**",
        "",
        "## Transakcije",
        "",
        "| Mera | Vrednost |",
        "|---|---|",
        f"| `*_TX` ukupno | {uk} |",
        f"| ...test helperi | {test} |",
        f"| ...produkcionih | {prod} |",
        f"| ...od toga transakcioni omotac oko blizanca bez `_TX` | **{om}** ({100 * om // max(prod, 1)}%) |",
        f"| ...od toga blizanac je `Public` (vrata pored granice) | {pub} |",
        "",
        "| Invarijanta ADR-0003 B | Vrednost |",
        "|---|---|",
        f"| Procedura sa `BeginTx` | {nb} |",
        f"| ...deklarise `AddTableSnapshot` u istoj proceduri | {ns} |",
        f"| **`AddTableSnapshot` bez `BeginTx` -- mora biti 0** | **{sam}** |",
        "",
    ]
    return "\n".join(L) + "\n"


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", help="putanja izlaznog .md (podrazumevano stdout)")
    a = ap.parse_args(argv)
    tekst = izvestaj()
    if a.out:
        with open(a.out, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(tekst)
        print(f"Upisano: {a.out}")
    else:
        sys.stdout.write(tekst)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
