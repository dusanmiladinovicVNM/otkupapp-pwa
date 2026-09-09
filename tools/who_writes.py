"""Ko pise koju tabelu -- mapa vlasnistva nad podacima, izvedena iz koda.

Rucno pisana mapa istruli za mesec dana. Ova se generise, pa je verifikovana po
konstrukciji: sto nije u kodu, nije ni u mapi.

Dva izvora, oba mehanicka -- ali NE znace isto:

  1. MUTATE: AppendRow / UpdateCell / RequireUpdateCell TBL_X
     Modul stvarno MENJA redove. Samo ovo je vlasnistvo (ugovor A11), i samo
     ovo meri --check-ownership.
  2. TX: clsTransaction.AddTableSnapshot TBL_X
     Operacija deklarise koje tabele njena transakcija mora da ume da VRATI.
     To je ucesce u transakciji, NE vlasnistvo: koordinator sme da snapshotuje
     tudju tabelu i pozove API njenog vlasnika.

Ranija verzija ovog fajla je tvrdila da je snapshot "najpouzdaniji signal: ako
modul snapshot-uje tabelu, on je i menja". Nije tacno i vodilo je do lazne slike
vlasnistva -- npr. modStorno snapshotuje tblOtkup a ne pise ga direktno.

Cemu sluzi: kad isto polje pise vise mesta po razlicitim pravilima, to je klasa
buga koju test hvata tek posle nastanka (v. CLAUDE.md S5). Mapa to cini vidljivim
PRE izmene -- ako menjas pravilo upisa, ovde vidis ko jos pise istu tabelu.

    python3 tools/who_writes.py                    # ispis na stdout
    python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md
    python3 tools/who_writes.py --check            # exit 2 ako je fajl zastareo

Radi svuda (ne treba Excel).
"""

import argparse
import collections
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src-vba")
DEFAULT_OUT = os.path.join(ROOT, "docs", "DOMEN", "WHO_WRITES.md")
OWNERSHIP_PATH = os.path.join(ROOT, "docs", "DOMEN", "WRITE_OWNERSHIP.json")

VBA_EXT = (".bas", ".cls", ".frm", ".doccls")

# DVA razlicita pojma, koja su do PR1 bila pomesana:
#
#   MUTATOR   -- modul koji stvarno MENJA redove tabele. Samo on je vlasnik u
#                smislu A11.
#   UCESNIK   -- modul cija transakcija snapshotuje tabelu da bi RollbackTx
#                umeo da je vrati. To je ucesce u transakciji, NE vlasnistvo:
#                ciljna arhitektura ima koordinatora koji snapshotuje tudju
#                tabelu i zove API njenog vlasnika.
#
# RequireUpdateCell je do PR1 bio NEVIDLJIV kapiji: stari regex je trazio
# \bUpdateCell, a u "RequireUpdateCell" pre "UpdateCell" nema granice reci.
# Time je 220 poziva -- 29 nad tblFakture, 26 nad tblPaleta, 12 nad tblOtkup --
# prolazilo kroz mapu neopazeno.
SNAPSHOT_RE = re.compile(r'AddTableSnapshot\s+(TBL_\w+|"(\w+)")', re.I)
MUTATE_RE = re.compile(
    r'\b(?:Require)?(?:AppendRow|UpdateCell)\s+(TBL_\w+|"(\w+)")', re.I)

# Test moduli se prikazuju odvojeno: oni pisu tabele namerno i uvek uz rollback,
# pa nisu vlasnici podataka i ne treba da zamagle pravu sliku.
TEST_MODULE_RE = re.compile(r"(Tests?$|^modTest)", re.I)


def table_constants() -> dict:
    path = os.path.join(SRC, "modConfig.bas")
    with open(path, encoding="utf-8", errors="replace") as fh:
        text = fh.read()
    return dict(re.findall(r'Public Const (TBL_\w+)\s+As String = "(\w+)"', text))


def scan() -> dict:
    const2tbl = table_constants()
    writers = collections.defaultdict(lambda: collections.defaultdict(set))

    for name in sorted(os.listdir(SRC)):
        if not name.endswith(VBA_EXT):
            continue
        module = name.rsplit(".", 1)[0]
        with open(os.path.join(SRC, name), encoding="utf-8", errors="replace") as fh:
            for line in fh:
                stripped = line.strip()
                if stripped.startswith("'"):        # komentar
                    continue
                for regex, kind in ((SNAPSHOT_RE, "tx"), (MUTATE_RE, "mutate")):
                    for m in regex.finditer(stripped):
                        token = m.group(1)
                        table = const2tbl.get(token, m.group(2) or token)
                        writers[table][kind].add(module)
    return writers


def render(writers: dict) -> str:
    lines = [
        "# Ko pise koju tabelu",
        "",
        "> **Generisan fajl -- ne menjaj rukom.**",
        "> `python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md`",
        "",
        "Izvedeno iz dva mehanicka signala u `src-vba/`:",
        "",
        "- **mutate** -- `AppendRow` / `UpdateCell` / `RequireUpdateCell`:",
        "  modul stvarno MENJA redove. Samo ovo je vlasnistvo (ugovor A11).",
        "- **tx** -- `clsTransaction.AddTableSnapshot TBL_X`: operacija",
        "  snapshotuje tabelu da bi `RollbackTx` umeo da je vrati. To je",
        "  UCESCE u transakciji, ne vlasnistvo -- koordinator sme da",
        "  snapshotuje tudju tabelu i zove API njenog vlasnika.",
        "",
        "Test moduli su odvojeni: pisu uz rollback i nisu vlasnici podataka.",
        "",
        "**Cemu sluzi:** kad isto polje pise vise mesta po razlicitim pravilima,",
        "to je klasa buga koju test hvata tek posle nastanka. Pre nego sto",
        "promenis pravilo upisa, ovde vidis ko jos pise istu tabelu.",
        "",
    ]

    def is_test(mod: str) -> bool:
        return bool(TEST_MODULE_RE.search(mod))

    rows = []
    for table, kinds in writers.items():
        mut = sorted({m for m in kinds.get("mutate", set()) if not is_test(m)})
        tx = sorted({m for m in kinds.get("tx", set()) if not is_test(m)})
        test = sorted({m for s in kinds.values() for m in s if is_test(m)})
        rows.append((table, mut, test, tx))

    rows.sort(key=lambda r: (-len(r[1]), r[0]))

    lines += ["| Tabela | Mutatora | Moduli koji MENJAJU redove |", "|---|---|---|"]
    for table, mut, _, _ in rows:
        mods = ", ".join(f"`{m}`" for m in mut) if mut else "_(samo testovi)_"
        lines.append(f"| `{table}` | {len(mut)} | {mods} |")

    lines += ["", "## Ucesnici transakcije (snapshot, NE vlasnistvo)", ""]
    for table, _, _, tx in rows:
        if tx:
            lines.append(f"- `{table}`: " + ", ".join(f"`{m}`" for m in tx))

    lines += ["", "## Test moduli po tabeli", ""]
    for table, _, test, _ in rows:
        if test:
            lines.append(f"- `{table}`: " + ", ".join(f"`{m}`" for m in test))

    lines += [
        "",
        "## Sta ovo NE pokriva",
        "",
        "- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad",
        "  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --",
        "  ako ga nadjes, to je nalaz, ne rupa u mapi.",
        "- Granularnost je tabela, ne kolona.",
        "",
    ]
    return "\n".join(lines) + "\n"


def production_writers(writers: dict, kind: str = "mutate") -> dict:
    """tabela -> sortirani produkcioni moduli date vrste.

    kind="mutate" su stvarni pisci (A11 gate meri njih).
    kind="tx" su ucesnici transakcije -- korisno za citanje, ali NE vlasnistvo.
    """
    out = {}
    for table, kinds in writers.items():
        prod = sorted(m for m in kinds.get(kind, set())
                      if not TEST_MODULE_RE.search(m))
        if prod:
            out[table] = prod
    return out


def check_ownership(writers: dict, path: str) -> int:
    """Architecture Contract A11: nov pisac van liste obara CI.

    Racna, ne cilj: 'dozvoljeni' je zamrznuto zateceno stanje, pa gate hvata
    SIRENJE vlasnistva od danas. Skracivanje ka 'cilj' je posao kasnijih PR-ova
    i svaki korak je vidljiva izmena ovog fajla.
    """
    if not os.path.exists(path):
        print(f"Ne postoji: {path}", file=sys.stderr)
        return 2

    with open(path, encoding="utf-8") as fh:
        reg = json.load(fh)

    stvarno = production_writers(writers, "mutate")
    greske = []

    for table, prod in sorted(stvarno.items()):
        if table not in reg:
            greske.append(
                f"  {table}: tabela nije u registru vlasnistva. "
                f"Pisci: {', '.join(prod)}")
            continue
        # ISKLJUCIVO row_owner. schema_owner NE ucestvuje: ugovor kaze da
        # modSetup sme da NAPRAVI tblOtkup ali ne i da upise otkup, pa bi unija
        # dva spiska bila poznat bypass -- kapija bi propustila bas ono sto
        # ugovor zabranjuje. Ko sme da menja SEMU je druga provera, ne ova.
        dozvoljeni = set(reg[table].get("row_owner", []))
        novi = [m for m in prod if m not in dozvoljeni]
        if novi:
            greske.append(
                f"  {table}: nov pisac van liste -> {', '.join(novi)}. "
                f"Ili zovi API vlasnika ({', '.join(sorted(dozvoljeni)) or 'nema'}), "
                f"ili svesno prosiri {os.path.basename(path)}.")

    if greske:
        print("A11 -- vlasnistvo nad upisom prekrseno:", file=sys.stderr)
        for g in greske:
            print(g, file=sys.stderr)
        return 2

    # napredak ka cilju -- informativno, ne obara
    otvoreno = []
    for table in sorted(reg):
        if table == "_o_fajlu":
            continue
        cilj = reg[table].get("cilj") or []
        if not cilj:
            continue
        viska = sorted(set(reg[table].get("row_owner", [])) - set(cilj))
        if viska:
            otvoreno.append(f"  {table}: jos {len(viska)} -> {', '.join(viska)}")

    print(f"{os.path.basename(path)}: nema novih pisaca "
          f"({len(stvarno)} tabela provereno)")
    if otvoreno:
        print("Do A11 cilja jos:")
        for o in otvoreno:
            print(o)
    return 0


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Mapa vlasnistva nad tabelama, iz koda.")
    ap.add_argument("--out", nargs="?", const=DEFAULT_OUT,
                    help=f"upisi u fajl (podrazumevano {DEFAULT_OUT})")
    ap.add_argument("--check", action="store_true",
                    help="exit 2 ako se generisan sadrzaj razlikuje od fajla")
    ap.add_argument("--check-ownership", action="store_true",
                    help="exit 2 ako tabelu pise modul van WRITE_OWNERSHIP.json")
    args = ap.parse_args(argv)

    writers = scan()

    if args.check_ownership:
        return check_ownership(writers, OWNERSHIP_PATH)

    text = render(writers)

    if args.check:
        path = args.out or DEFAULT_OUT
        if not os.path.exists(path):
            print(f"Ne postoji: {path} -- pokreni bez --check", file=sys.stderr)
            return 2
        with open(path, encoding="utf-8") as fh:
            if fh.read() != text:
                print(f"{path} je zastareo -- regenerisi ga "
                      "(python3 tools/who_writes.py --out)", file=sys.stderr)
                return 2
        print(f"{path}: azuran")
        return 0

    if args.out:
        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"Upisano: {args.out}")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
