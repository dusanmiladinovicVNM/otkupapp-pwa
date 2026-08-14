"""Uporedi EFEKTIVNU semu koju modSetup.bas pravi, izmedju dve verzije fajla.

Zasto postoji: "sema tabela je izvor istine, ne kod" (.claude/rules/podaci-i-config.md)
vazi za zatecenu instalaciju, ali kod odlucuje sta se kreira na NOVOJ. Refaktor
modSetup-a se zato ne sme oceniti "na oko" -- ispusten COL_* u nekoj Array listi
je tiha izmena seme koja se vidi tek kao prazna kolona na dokumentu kod klijenta.

Alat izvlaci, po ulaznoj tacki (Ensure*Schema...), UREDJENU listu DDL operacija i
diff-uje dve verzije. Redosled je deo semantike: EnsureColumnOnTable dodaje kolonu
na KRAJ tabele, a pozicijski AppendRow zavisi od redosleda kolona.

Razume OBA zapisa, pa radi i preko refaktora koji je uveo registar (F5):
  - inline pozivi  : EnsureDataTable TBL_X, "Sheet", Array(...)
  - schema registar: AddTableSpec c, SG_GRUPA, TBL_X, "Sheet", Array(...)

    git show <ref>:src-vba/modSetup.bas > /tmp/old.bas
    python3 tools/schema_diff.py /tmp/old.bas src-vba/modSetup.bas

Izlazni kod: 0 = sema identicna, 1 = razlika (ispisana red po red).

Ogranicenje: cita SAMO modSetup.bas i samo cetiri DDL primitiva. Ne vidi kolone
koje neki drugi modul doda sam, niti sta se desava sa POSTOJECIM podacima.
"""
from __future__ import annotations

import re
import sys

# Grupa registra -> ulazna tacka koja je primenjuje. Dve dorade grupe idu u istu
# proceduru, ovim redom -- izmedju njih se zove EnsureRuntimeSchema.
GROUP_PROC = {
    "SG_PALETNI": "EnsurePaletniListSchemaCore",
    "SG_CENOVNIK": "EnsureCenovnikSchema",
    "SG_STORNO_VEZE": "EnsureStornoVezeSchemaCore",
    "SG_STORNO_ZURNAL": "EnsureStornoZurnalSchemaCore",
    "SG_PORUKE": "EnsurePoruke",
    "SG_KORISNICI": "EnsureKorisniciSchema",
    "SG_RUNTIME": "EnsureRuntimeSchema",
    "SG_DORADE_SIFARNICI": "EnsureDoradeSchemaCore",
    "SG_DORADE_DOKUMENTI": "EnsureDoradeSchemaCore",
}

# Procedure cija se sema poredi. Sve ostalo u modSetup-u (checks, folderi,
# logovanje) nije DDL i ne ulazi u poredjenje.
TRACKED = set(GROUP_PROC.values())

PROC_START = re.compile(r"^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)", re.I)
PROC_END = re.compile(r"^\s*End (?:Sub|Function)\s*$", re.I)
REGISTRY_FN = re.compile(r"^\s*Private Function (SchemaTables|SchemaOps)\(", re.I)

OP_KIND = {"OP_COLUMN": "COLUMN", "OP_FORMAT": "FORMAT", "OP_BACKFILL": "BACKFILL"}


def strip_comment(text: str) -> str:
    """Odbaci prateci ' komentar, ali ne apostrof unutar stringa.

    Bez ovoga bi `EnsureColumnOnTable TBL_PALETA, COL_PAL_ISTORIJA   ' audit trag`
    bio drugacija operacija od istog poziva bez komentara -- lazan diff.
    """
    in_str = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return text[:i].rstrip()
    return text


def join_continuations(lines: list[str]) -> list[str]:
    """VBA ' _' nastavak reda -> jedna logicka linija."""
    out: list[str] = []
    buf = ""
    for ln in lines:
        s = ln.rstrip()
        if s.endswith(" _"):
            buf += s[:-2].strip() + " "
        else:
            out.append((buf + s.strip()).strip())
            buf = ""
    if buf:
        out.append(buf.strip())
    return out


def norm_cols(arr_text: str) -> str:
    """Array(a, b, c) -> 'a|b|c'"""
    inner = re.sub(r"^Array\(|\)$", "", arr_text.strip())
    return "|".join(p.strip() for p in inner.split(","))


def aktivan_pair(tbl: str) -> list[str]:
    """Kolona Aktivan + backfill -- isti par u oba zapisa."""
    return [f'COLUMN {tbl} "Aktivan"', f'BACKFILL {tbl} "Aktivan" STATUS_AKTIVAN']


def parse(path: str) -> dict[str, list[str]]:
    lines = join_continuations(
        open(path, encoding="latin-1").read().replace("\r\n", "\n").split("\n"))

    per_proc: dict[str, list[str]] = {}
    per_group: dict[str, list[str]] = {g: [] for g in GROUP_PROC}

    proc = ""
    in_registry = False

    for ln in lines:
        s = strip_comment(ln).strip()
        if not s:
            continue

        if REGISTRY_FN.match(ln):
            in_registry, proc = True, ""
            continue
        if PROC_END.match(ln):
            in_registry, proc = False, ""
            continue
        m = PROC_START.match(ln)
        if m:
            proc = m.group(1)
            continue

        # --- zapis A: schema registar ---
        if in_registry:
            m = re.match(
                r'AddTableSpec c,\s*(\w+),\s*([^,]+),\s*("?[^,]+"?),\s*(Array\(.*\))$', s)
            if m:
                per_group[m.group(1)].append(
                    f"TABLE {m.group(2).strip()} {m.group(3).strip()} "
                    f"[{norm_cols(m.group(4))}]")
                continue

            m = re.match(r"AddOp c,\s*(\w+),\s*(\w+),\s*([^,]+),\s*(.+?),\s*(.+)$", s)
            if m:
                grp, kind, tbl, col, arg = (x.strip() for x in m.groups())
                kind = OP_KIND[kind]
                per_group[grp].append(
                    f"{kind} {tbl} {col}" if kind == "COLUMN"
                    else f"{kind} {tbl} {col} {arg}")
                continue

            m = re.match(r"AddAktivanOps c,\s*(\w+)$", s)
            if m:
                per_group["SG_DORADE_SIFARNICI"].extend(aktivan_pair(m.group(1)))
            continue

        # --- zapis B: inline pozivi u telu Ensure*Schema ---
        if proc not in TRACKED:
            continue

        m = re.match(r'EnsureDataTable\s+([^,]+),\s*("?[^,]+"?),\s*(Array\(.*\))$', s)
        if m:
            per_proc.setdefault(proc, []).append(
                f"TABLE {m.group(1).strip()} {m.group(2).strip()} "
                f"[{norm_cols(m.group(3))}]")
            continue

        m = re.match(r"EnsureColumnOnTable\s+([^,]+),\s*(.+)$", s)
        if m:
            per_proc.setdefault(proc, []).append(
                f"COLUMN {m.group(1).strip()} {m.group(2).strip()}")
            continue

        m = re.match(r"SetColumnNumberFormat\s+([^,]+),\s*([^,]+),\s*(.+)$", s)
        if m:
            per_proc.setdefault(proc, []).append(
                f"FORMAT {m.group(1).strip()} {m.group(2).strip()} {m.group(3).strip()}")
            continue

        m = re.match(r"BackfillColumn\s+([^,]+),\s*([^,]+),\s*(.+)$", s)
        if m:
            per_proc.setdefault(proc, []).append(
                f"BACKFILL {m.group(1).strip()} {m.group(2).strip()} {m.group(3).strip()}")
            continue

        m = re.match(r"EnsureAktivanColumn\s+(\w+)$", s)
        if m:
            per_proc.setdefault(proc, []).extend(aktivan_pair(m.group(1)))

    for grp, ops in ((g, per_group[g]) for g in GROUP_PROC):
        if ops:
            per_proc.setdefault(GROUP_PROC[grp], []).extend(ops)

    return per_proc


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print(__doc__, file=sys.stderr)
        print("upotreba: schema_diff.py <stari modSetup.bas> <novi modSetup.bas>",
              file=sys.stderr)
        return 2

    old, new = parse(argv[0]), parse(argv[1])
    bad = 0

    for proc in sorted(set(old) | set(new)):
        a, b = old.get(proc, []), new.get(proc, [])
        if a == b:
            print(f"  OK   {proc:<32} {len(a)} operacija identicno")
            continue

        bad += 1
        print(f"  DIFF {proc}")
        for i in range(max(len(a), len(b))):
            x = a[i] if i < len(a) else "<nema>"
            y = b[i] if i < len(b) else "<nema>"
            if x != y:
                print(f"       [{i}] stara: {x}")
                print(f"           nova : {y}")

    print()
    if bad:
        print(f"schema_diff: RAZLIKA u {bad} procedura -- sema NIJE ista.")
        return 1

    total = sum(len(v) for v in new.values())
    print(f"schema_diff: sema identicna ({len(new)} ulaznih tacaka, {total} operacija).")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
