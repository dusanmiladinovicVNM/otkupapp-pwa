"""Ko pise koju tabelu -- mapa vlasnistva nad podacima, izvedena iz koda.

Rucno pisana mapa istruli za mesec dana. Ova se generise, pa je verifikovana po
konstrukciji: sto nije u kodu, nije ni u mapi.

Dva izvora, oba mehanicka:

  1. clsTransaction.AddTableSnapshot TBL_X  -- svaka _TX operacija SAMA deklarise
     koje tabele menja (da bi RollbackTx umeo da ih vrati). To je najpouzdaniji
     signal: ako modul snapshot-uje tabelu, on je i menja.
  2. AppendRow / UpdateCell TBL_X           -- direktan upis kroz modDataAccess.

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
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src-vba")
DEFAULT_OUT = os.path.join(ROOT, "docs", "DOMEN", "WHO_WRITES.md")

VBA_EXT = (".bas", ".cls", ".frm", ".doccls")

SNAPSHOT_RE = re.compile(r'AddTableSnapshot\s+(TBL_\w+|"(\w+)")', re.I)
DIRECT_RE = re.compile(r'\b(?:AppendRow|UpdateCell)\s+(TBL_\w+|"(\w+)")', re.I)

# Test moduli se prikazuju odvojeno: oni pisu tabele namerno i uvek uz rollback,
# pa nisu vlasnici podataka i ne treba da zamagle pravu sliku.
TEST_MODULE_RE = re.compile(r"(Tests?$|^modTest)", re.I)


def table_constants() -> dict:
    path = os.path.join(SRC, "modConfig.bas")
    with open(path, encoding="utf-8", errors="replace") as fh:
        text = fh.read()
    return dict(re.findall(r'Public Const (TBL_\w+) As String = "(\w+)"', text))


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
                for regex, kind in ((SNAPSHOT_RE, "tx"), (DIRECT_RE, "direct")):
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
        "- **tx** -- `clsTransaction.AddTableSnapshot TBL_X`: operacija sama",
        "  deklarise koje tabele menja, da bi `RollbackTx` umeo da ih vrati.",
        "- **direct** -- `AppendRow` / `UpdateCell` kroz `modDataAccess`.",
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
        prod = sorted({m for s in kinds.values() for m in s if not is_test(m)})
        test = sorted({m for s in kinds.values() for m in s if is_test(m)})
        rows.append((table, prod, test))

    rows.sort(key=lambda r: (-len(r[1]), r[0]))

    lines += ["| Tabela | Pisaca | Produkcioni moduli |", "|---|---|---|"]
    for table, prod, _ in rows:
        mods = ", ".join(f"`{m}`" for m in prod) if prod else "_(samo testovi)_"
        lines.append(f"| `{table}` | {len(prod)} | {mods} |")

    lines += ["", "## Test moduli po tabeli", ""]
    for table, _, test in rows:
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


def main(argv) -> int:
    ap = argparse.ArgumentParser(description="Mapa vlasnistva nad tabelama, iz koda.")
    ap.add_argument("--out", nargs="?", const=DEFAULT_OUT,
                    help=f"upisi u fajl (podrazumevano {DEFAULT_OUT})")
    ap.add_argument("--check", action="store_true",
                    help="exit 2 ako se generisan sadrzaj razlikuje od fajla")
    args = ap.parse_args(argv)

    text = render(scan())

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
