"""Identitet onoga sto je testirano: kanonski hash src-vba, last-green marker, manifest.

Radi svuda (nema Excela, nema COM-a). Ovo je sloj koji odgovara na pitanje koje
`run_vba.py` ne moze da postavi sam sebi:

    "da li je BAS OVAJ izvor prosao testove, ili neki drugi?"

Do sada je odgovor bio posredan i netacan: Stop hook je gledao `git diff HEAD~1`,
pa je commit koji dira samo docs "resetovao" potrebu za testom, a rebase/amend
je umeo da je sakrije. Ovde se umesto toga hesira sam izvor i pamti hash
poslednjeg zelenog run-a.

    python3 tools/vba_gate.py --hash              # kanonski SHA256 src-vba
    python3 tools/vba_gate.py --status            # citljiv rezime (izvor vs poslednji zeleni)
    python3 tools/vba_gate.py --require-green     # exit 0 samo ako je OVAJ izvor zelen
    python3 tools/vba_gate.py --mark-green        # zove run_vba.py posle zelenog run-a
    python3 tools/vba_gate.py --manifest-check    # manifest <-> stvarne procedure u kodu
    python3 tools/vba_gate.py --provenance        # JSON otisak run-a (git, hash, platforma)
    python3 tools/vba_gate.py --self-test         # provere samog alata

Izlazni kod: 0 = ok, 2 = nalaz.

KANONSKI HASH -- zasto normalizacija prelomaka
    `src-vba/` se na Windows-u checkout-uje kao CRLF, a na Linuxu kao LF (nema
    `.gitattributes` za `.bas`). Sirov hash bajtova bi zato na dve masine dao dva
    razlicita broja nad ISTIM commitom, pa bi "zeleno na build masini" izgledalo
    kao "nepoznat izvor" svuda drugde. Zato se tekstualni moduli normalizuju na
    LF pre hesiranja; `.frx` je binaran i ide sirov.

    `modBuildInfo.bas` je IZUZET: `tools/stamp-build` ga prepisuje pri svakom
    build-u (BUILD_SHA/BUILD_VERSION/BUILD_DATE) i ta izmena se ne commit-uje.
    Da je u hesu, stamp bi obarao zelen marker bas u trenutku release-a, a modul
    ne nosi nikakvo ponasanje -- tri konstante koje niko ne grana.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import locale
import os
import platform
import re
import subprocess
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vba_check  # noqa: E402  -- isti folder; reuse regex-a i strip-a komentara

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
MANIFEST = os.path.join(ROOT, "tests", "suite_manifest.json")
DEFAULT_FIXTURE = os.path.join(ROOT, "tests", "fixtures", "otkup_test.xlsm")

GATE_VERSION = 1                 # menja se kad se promeni sta hash pokriva
GREEN_FILE = "agrix-vba-last-green"

# Marker pamti i KOJA je kapija bila zelena, ne samo da je nesto bilo zeleno.
# Bez toga bi zelen `dev` run (bez RunNovacSmokeSuite i bez external ugovora)
# release kapiji rekao "testirano" -- a nije testirano celo.
GATE_RANK = {"dev": 1, "pr": 2, "release": 3}

TEXT_EXT = (".bas", ".cls", ".frm", ".doccls")
BINARY_EXT = (".frx",)
HASH_EXCLUDE = {"modbuildinfo.bas"}

# Sta se smatra ulaznom tackom test suite-a. Namerno usko: sve ostalo
# (`Test_XyzAuto`, `TestHook_*`, pojedinacni testovi) su komponente koje zove
# roditeljska suite, i one se NE ocekuju u manifestu.
ENTRY_POINT = re.compile(r"^(?:Run.+(?:Suite|Tests)|Test.*_All)$")


# --- Kanonski hash ------------------------------------------------------------

def _hash_files(paths: list[str], base: str) -> str:
    h = hashlib.sha256()
    for path in sorted(paths, key=lambda p: os.path.relpath(p, base).lower()):
        rel = os.path.relpath(path, base).replace("\\", "/").lower()
        with open(path, "rb") as fh:
            data = fh.read()
        if path.lower().endswith(TEXT_EXT):
            data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
        h.update(rel.encode("ascii", "replace"))
        h.update(b"\0")
        h.update(str(len(data)).encode("ascii"))
        h.update(b"\0")
        h.update(data)
        h.update(b"\0")
    return h.hexdigest()


def source_hash(src_dir: str = SRC_VBA) -> str:
    """SHA256 nad celim VBA izvorom -- isti na Windows-u i na Linuxu."""
    paths = [os.path.join(src_dir, n) for n in os.listdir(src_dir)
             if n.lower().endswith(TEXT_EXT + BINARY_EXT)
             and n.lower() not in HASH_EXCLUDE]
    return _hash_files(paths, src_dir)


def file_hash(path: str) -> str | None:
    """Sirov SHA256 jednog fajla (fixture, donor) -- None ako ga nema."""
    if not path or not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --- Marker poslednjeg zelenog run-a -------------------------------------------

def green_path(root: str = ROOT) -> str:
    """`.git/agrix-vba-last-green` -- u .git namerno: nikad ne moze da se commit-uje
    i ne prezivljava fresh klon (na novoj masini gate mora da se dokaze iznova)."""
    try:
        out = subprocess.run(["git", "-C", root, "rev-parse", "--absolute-git-dir"],
                             capture_output=True, text=True, timeout=15)
        if out.returncode == 0 and out.stdout.strip():
            return os.path.join(out.stdout.strip(), GREEN_FILE)
    except Exception:                       # noqa: BLE001
        pass
    return os.path.join(root, "." + GREEN_FILE)


def read_green(root: str = ROOT) -> dict | None:
    path = green_path(root)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:                       # noqa: BLE001
        return None
    return data if isinstance(data, dict) else None


def write_green(payload: dict, root: str = ROOT) -> str:
    path = green_path(root)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2, sort_keys=True)
    return path


def is_green(root: str = ROOT, src_dir: str = SRC_VBA,
             need_gate: str = "dev") -> tuple[bool, str, dict | None]:
    """(zelen?, tekuci hash, zapis poslednjeg zelenog).

    `need_gate` je minimum: marker iz `release` run-a zadovoljava i `dev` i `pr`,
    ali marker iz `dev` run-a NE zadovoljava `pr`.
    """
    cur = source_hash(src_dir)
    rec = read_green(root)
    if rec is None:
        return False, cur, None
    ok = (rec.get("source_hash") == cur
          and int(rec.get("gate_version", -1)) == GATE_VERSION
          and GATE_RANK.get(str(rec.get("gate", "dev")), 0) >= GATE_RANK.get(need_gate, 1))
    return ok, cur, rec


# --- Manifest ------------------------------------------------------------------

class ManifestError(Exception):
    pass


def load_manifest(path: str = MANIFEST) -> dict:
    if not os.path.exists(path):
        raise ManifestError(f"manifest ne postoji: {path}")
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if "suites" not in data:
        raise ManifestError(f"{path}: nema kljuc 'suites'")
    ids = [s["id"] for s in data["suites"]]
    dup = {i for i in ids if ids.count(i) > 1}
    if dup:
        raise ManifestError(f"{path}: dupli id: {', '.join(sorted(dup))}")
    return data


def suites_for_gate(manifest: dict, gate: str) -> list[str]:
    return [s["id"] for s in manifest["suites"] if gate in (s.get("gates") or [])]


def suite_meta(manifest: dict, suite_id: str) -> dict:
    for s in manifest["suites"]:
        if s["id"] == suite_id:
            return s
    raise ManifestError(f"suite nije u manifestu: {suite_id}")


def public_subs(src_dir: str = SRC_VBA) -> dict[str, str]:
    """Ime -> fajl, za svaki `Public Sub` u `.bas` modulima."""
    found: dict[str, str] = {}
    for name in sorted(os.listdir(src_dir)):
        if not name.lower().endswith(".bas"):
            continue
        with open(os.path.join(src_dir, name), "r", encoding="ascii",
                  errors="replace") as fh:
            for line in fh:
                m = vba_check.PUBLIC_PROC.match(line)
                if m and line.split()[1].lower() == "sub":
                    found.setdefault(m.group(1), name)
    return found


def called_in_statement_position(src_dir: str = SRC_VBA) -> set[str]:
    """Imena pozvana kao NAREDBA (`Foo`, `Call Foo`) bilo gde u `.bas` izvoru.

    Namerno ne broji pominjanje u stringu (`LogFatal "RunSEFOfflineSuite"`) ni u
    komentaru -- oba postoje u ovom kodu bas oko suite-ova koji se nigde ne
    pokrecu, pa bi ih sakrila.
    """
    called: set[str] = set()
    for name in sorted(os.listdir(src_dir)):
        if not name.lower().endswith(".bas"):
            continue
        with open(os.path.join(src_dir, name), "r", encoding="ascii",
                  errors="replace") as fh:
            for raw in fh:
                line = raw.rstrip()
                if vba_check.PROC_DEF.match(line) or vba_check.DECLARE_DEF.match(line):
                    continue
                stmt = vba_check._strip_comment(line.strip())
                if not stmt:
                    continue
                m = vba_check.CALL_STMT.match(stmt)
                if not m:
                    continue
                nm, rest = m.group(1), m.group(2).lstrip()
                if nm.lower() in vba_check.STMT_WORDS:
                    continue
                if rest.startswith(("=", ".", "!")):
                    continue
                called.add(nm)
    return called


def manifest_check(manifest: dict | None = None, src_dir: str = SRC_VBA) -> list[str]:
    """Nalazi u oba smera izmedju manifesta i stvarnog koda.

    Smer 1 (manifest -> kod): upisana suite koja vise ne postoji ili je u drugom
    modulu. Bez ovoga preimenovanje suite-a tiho izbaci ~200 provera iz kapije.

    Smer 2 (kod -> manifest): ulazna tacka koju NIKO ne poziva i koja nije
    upisana. To je suite koju niko nikad nece pokrenuti -- napisana pa
    zaboravljena. Priznata takva mesta idu u `unlisted` sa razlogom; cutanje nije
    opcija.
    """
    out: list[str] = []
    if manifest is None:
        manifest = load_manifest()

    subs = public_subs(src_dir)
    listed = {s["id"] for s in manifest["suites"]}
    unlisted = {u["id"] for u in manifest.get("unlisted", [])}

    for s in manifest["suites"]:
        sid = s["id"]
        if sid not in subs:
            out.append(f"MANIFEST  '{sid}' je u manifestu, a nema ga kao Public Sub u src-vba/ "
                       f"(preimenovan ili obrisan -- suite tiho ispada iz kapije)")
            continue
        want = (s.get("module") or "").lower()
        if want and subs[sid].lower() != want:
            out.append(f"MANIFEST  '{sid}': manifest kaze {s['module']}, "
                       f"a definisan je u {subs[sid]}")

    for u in manifest.get("unlisted", []):
        if u["id"] not in subs:
            out.append(f"MANIFEST  unlisted '{u['id']}' vise ne postoji u src-vba/ "
                       f"-- izbaci ga iz manifesta")
        if not (u.get("reason") or "").strip():
            out.append(f"MANIFEST  unlisted '{u['id']}' nema 'reason' "
                       f"-- nepokretana suite bez zapisanog razloga je tiho preskakanje")

    called = called_in_statement_position(src_dir)
    for name, module in sorted(subs.items()):
        if not ENTRY_POINT.match(name):
            continue
        if name in listed or name in unlisted or name in called:
            continue
        out.append(f"MANIFEST  '{name}' ({module}) izgleda kao ulazna tacka suite-a, "
                   f"niko je ne poziva i nije u manifestu -- upisi je u 'suites' "
                   f"ili u 'unlisted' sa razlogom")

    for s in manifest["suites"]:
        for key in ("category", "gates", "raises", "timeout_s"):
            if key not in s:
                out.append(f"MANIFEST  '{s['id']}': nedostaje polje '{key}'")
    return out


# --- Provenance ----------------------------------------------------------------

def _git(*args: str) -> str:
    try:
        out = subprocess.run(["git", "-C", ROOT, *args],
                             capture_output=True, text=True, timeout=30)
        return out.stdout.strip() if out.returncode == 0 else ""
    except Exception:                       # noqa: BLE001
        return ""


def provenance(fixture: str | None = None, src_dir: str = SRC_VBA) -> dict:
    """Otisak onoga NAD CIM je run izvrsen. Bez ovoga "zeleno" je apstraktna
    tvrdnja; sa ovim je: ovaj izvor, nad ovim fixture-om, na ovoj masini."""
    fixture = fixture or DEFAULT_FIXTURE
    dirty = _git("status", "--porcelain", "--", "src-vba")
    return {
        "gate_version": GATE_VERSION,
        "timestamp_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "git_sha": _git("rev-parse", "HEAD"),
        "git_branch": _git("rev-parse", "--abbrev-ref", "HEAD"),
        "git_describe": _git("describe", "--tags", "--always"),
        "src_vba_dirty": bool(dirty),
        "source_hash": source_hash(src_dir),
        "fixture": os.path.basename(fixture),
        "fixture_hash": file_hash(fixture),
        "platform": platform.platform(),
        "python": platform.python_version(),
        "locale": ".".join(x for x in locale.getlocale() if x) or "C",
    }


# --- Ispis ----------------------------------------------------------------------

def status_lines(root: str = ROOT, src_dir: str = SRC_VBA,
                 need_gate: str = "dev") -> list[str]:
    ok, cur, rec = is_green(root, src_dir, need_gate)
    lines = [f"IZVOR       {cur[:16]}...  ({len(os.listdir(src_dir))} fajlova u src-vba/)"]
    if rec is None:
        lines.append("ZELENO      nema zapisa -- ovaj klon nikad nije zavrsio zelen behavior run")
    else:
        lines.append(f"ZELENO      {str(rec.get('source_hash', ''))[:16]}...  "
                     f"{rec.get('timestamp_utc', '?')}  "
                     f"kapija={rec.get('gate', 'dev')}  "
                     f"suites={len(rec.get('suites') or [])}")
        if GATE_RANK.get(str(rec.get("gate", "dev")), 0) < GATE_RANK.get(need_gate, 1):
            lines.append(f"            zapis je iz kapije '{rec.get('gate', 'dev')}', "
                         f"a trazi se bar '{need_gate}'")
        if int(rec.get("gate_version", -1)) != GATE_VERSION:
            lines.append(f"            zapis je iz gate_version={rec.get('gate_version')}, "
                         f"tekuci je {GATE_VERSION} -- ne racuna se")
    lines.append("STANJE      " + (f"ZELEN za kapiju '{need_gate}' -- izvor je nepromenjen "
                                   f"od poslednjeg prolaza" if ok else
                                   f"TREBA TEST -- nema zelenog '{need_gate}' run-a nad OVIM izvorom"))
    return lines


# --- Self-test -------------------------------------------------------------------

def self_test() -> int:
    """Provere samog alata. Rade svuda; hash i manifest logika se ne oslanjaju na Excel."""
    import shutil
    import tempfile

    fails: list[str] = []

    def chk(cond: bool, label: str) -> None:
        if cond:
            print(f"  PASS  {label}")
        else:
            fails.append(label)
            print(f"  FAIL  {label}")

    tmp = tempfile.mkdtemp(prefix="vbagate_")
    try:
        a = os.path.join(tmp, "a")
        b = os.path.join(tmp, "b")
        os.makedirs(a)
        os.makedirs(b)

        # 1) CRLF vs LF nad istim sadrzajem mora dati ISTI hash -- inace bi
        #    "zeleno" sa Windows build masine bilo nevidljivo svuda drugde.
        with open(os.path.join(a, "modX.bas"), "wb") as fh:
            fh.write(b"Attribute VB_Name = \"modX\"\nOption Explicit\n")
        with open(os.path.join(b, "modX.bas"), "wb") as fh:
            fh.write(b"Attribute VB_Name = \"modX\"\r\nOption Explicit\r\n")
        chk(source_hash(a) == source_hash(b), "CRLF i LF daju isti kanonski hash")

        # 2) prava izmena sadrzaja MORA da promeni hash
        before = source_hash(a)
        with open(os.path.join(a, "modX.bas"), "ab") as fh:
            fh.write(b"' jedna linija vise\n")
        chk(source_hash(a) != before, "izmena sadrzaja menja hash")

        # 3) modBuildInfo.bas je izuzet (stamp ga prepisuje pred svaki build)
        before = source_hash(b)
        with open(os.path.join(b, "modBuildInfo.bas"), "w", encoding="ascii") as fh:
            fh.write('Public Const BUILD_SHA As String = "abc1234"\n')
        chk(source_hash(b) == before, "modBuildInfo.bas ne ulazi u hash")

        # 4) preimenovanje fajla menja hash (ime je deo otiska)
        before = source_hash(b)
        os.rename(os.path.join(b, "modX.bas"), os.path.join(b, "modY.bas"))
        chk(source_hash(b) != before, "preimenovanje modula menja hash")

        # 5) prazan .frx se hesira sirovo (bez normalizacije prelomaka)
        with open(os.path.join(a, "f.frx"), "wb") as fh:
            fh.write(b"\x00\x01\r\n\x02")
        h1 = source_hash(a)
        with open(os.path.join(a, "f.frx"), "wb") as fh:
            fh.write(b"\x00\x01\n\x02")
        chk(source_hash(a) != h1, ".frx se ne normalizuje (binaran je)")

        # 6) marker: zapis sa drugim hashom NIJE zelen
        fake_root = os.path.join(tmp, "repo")
        os.makedirs(fake_root)
        write_green({"source_hash": "deadbeef", "gate_version": GATE_VERSION}, fake_root)
        ok, _, _ = is_green(fake_root, a)
        chk(not ok, "marker sa tudjim hashom nije zelen")

        write_green({"source_hash": source_hash(a), "gate_version": GATE_VERSION}, fake_root)
        ok, _, _ = is_green(fake_root, a)
        chk(ok, "marker sa tacnim hashom je zelen")

        # 7) stara gate_version se ne racuna
        write_green({"source_hash": source_hash(a), "gate_version": GATE_VERSION - 1}, fake_root)
        ok, _, _ = is_green(fake_root, a)
        chk(not ok, "marker iz starije gate_version se ne racuna")

        # 8) rang kapija: dev marker ne moze da prodje kao pr/release
        write_green({"source_hash": source_hash(a), "gate_version": GATE_VERSION,
                     "gate": "dev"}, fake_root)
        chk(is_green(fake_root, a, "dev")[0], "dev marker zadovoljava dev kapiju")
        chk(not is_green(fake_root, a, "pr")[0], "dev marker NE zadovoljava pr kapiju")
        write_green({"source_hash": source_hash(a), "gate_version": GATE_VERSION,
                     "gate": "release"}, fake_root)
        chk(is_green(fake_root, a, "pr")[0], "release marker zadovoljava pr kapiju")

        # 8) manifest repoa mora da prodje sopstvenu proveru
        findings = manifest_check()
        chk(not findings, "manifest repoa je saglasan sa src-vba/ "
                          + ("" if not findings else f"({len(findings)} nalaza)"))
        for f in findings:
            print(f"        {f}")

        # 9) gate katalog mora da pokrije bar jednu suite po kapiji
        man = load_manifest()
        for gate in ("dev", "pr", "release"):
            chk(bool(suites_for_gate(man, gate)), f"kapija '{gate}' ima bar jednu suite")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print()
    if fails:
        print(f"self-test: {len(fails)} nalaza", file=sys.stderr)
        return 2
    print("self-test: cisto")
    return 0


# --- CLI ---------------------------------------------------------------------------

def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Kanonski hash src-vba, last-green marker, manifest.")
    ap.add_argument("--hash", action="store_true", help="ispisi kanonski SHA256 src-vba")
    ap.add_argument("--status", action="store_true", help="citljiv rezime izvor vs poslednji zeleni")
    ap.add_argument("--require-green", action="store_true",
                    help="exit 0 samo ako je BAS OVAJ izvor prosao behavior gate")
    ap.add_argument("--gate", choices=("dev", "pr", "release"), default="dev",
                    help="koja kapija se trazi uz --require-green / upisuje uz --mark-green")
    ap.add_argument("--mark-green", action="store_true", help="zapisi tekuci izvor kao zelen")
    ap.add_argument("--clear-green", action="store_true", help="obrisi marker")
    ap.add_argument("--suites", default="", help="uz --mark-green: lista suite-ova, zarezom")
    ap.add_argument("--note", default="", help="uz --mark-green: slobodan opis")
    ap.add_argument("--manifest-check", action="store_true", help="manifest <-> src-vba")
    ap.add_argument("--provenance", action="store_true", help="JSON otisak run-a")
    ap.add_argument("--fixture", default=None, help="putanja do fixture-a za --provenance")
    ap.add_argument("--json", action="store_true", help="masinski citljiv izlaz gde ima smisla")
    ap.add_argument("--self-test", action="store_true", help="provere samog alata")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test()

    if args.hash:
        print(source_hash())
        return 0

    if args.provenance:
        print(json.dumps(provenance(args.fixture), ensure_ascii=False, indent=2))
        return 0

    if args.clear_green:
        path = green_path()
        if os.path.exists(path):
            os.remove(path)
            print(f"Obrisan: {path}")
        else:
            print("Nema marker fajla.")
        return 0

    if args.mark_green:
        payload = provenance(args.fixture)
        payload["suites"] = [s for s in args.suites.split(",") if s]
        payload["gate"] = args.gate
        payload["note"] = args.note
        path = write_green(payload)
        print(f"ZELENO zapisano: {payload['source_hash'][:16]}...  -> {path}")
        return 0

    if args.manifest_check:
        try:
            findings = manifest_check()
        except ManifestError as exc:
            print(f"MANIFEST  {exc}", file=sys.stderr)
            return 2
        if args.json:
            print(json.dumps(findings, ensure_ascii=False, indent=2))
        else:
            for f in findings:
                print(f, file=sys.stderr)
        if findings:
            print(f"\n{len(findings)} nalaza u tests/suite_manifest.json.", file=sys.stderr)
            return 2
        man = load_manifest()
        print(f"manifest: cisto ({len(man['suites'])} suite-ova, "
              f"{len(man.get('unlisted') or [])} priznato nepokretanih)")
        return 0

    if args.require_green:
        ok, cur, rec = is_green(need_gate=args.gate)
        if ok:
            print(f"ZELENO  {cur[:16]}...  kapija={(rec or {}).get('gate', '?')}  "
                  f"({(rec or {}).get('timestamp_utc', '?')})")
            return 0
        for ln in status_lines(need_gate=args.gate):
            print(ln, file=sys.stderr)
        print(f"\nBEHAVIOR kapija '{args.gate}' nije prosla nad OVIM izvorom. Pokreni na "
              f"Windows masini sa Excelom:\n    python tools\\run_vba.py --gate {args.gate}",
              file=sys.stderr)
        return 2

    # bez argumenata = --status
    for ln in status_lines(need_gate=args.gate):
        print(ln)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
