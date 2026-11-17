"""P2b: predfilter u collect_definitions/collect_module_members + memo logicke_izjave."""
import sys
p = sys.argv[1]
s = open(p, encoding="utf-8").read()
def rep(old, new, n=1):
    global s
    assert s.count(old) == n, (s.count(old), old[:70])
    s = s.replace(old, new)

rep('''        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for line in fh:
                for rx in (DECLARE_DEF, PROC_DEF):
                    m = rx.match(line)
''', '''        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for line in fh:
                # PROTOTIP: oba izraza traze rec Sub/Function/Property
                if not _AR_KLJUCNA.search(line):
                    continue
                for rx in (DECLARE_DEF, PROC_DEF):
                    m = rx.match(line)
''')
rep('''        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for line in fh:
                for rx in (DECLARE_DEF, PROC_DEF, JAVNI_CLAN_DEF):
                    m = rx.match(line)
''', '''        with open(path, "r", encoding="ascii", errors="replace") as fh:
            for line in fh:
                # PROTOTIP: DECLARE/PROC traze Sub/Function/Property, JAVNI_CLAN
                # trazi Public/Global -- red bez ijedne od tih reci ne pogadja.
                if not _CLAN_KLJUCNA.search(line):
                    continue
                for rx in (DECLARE_DEF, PROC_DEF, JAVNI_CLAN_DEF):
                    m = rx.match(line)
''')
rep('''def collect_definitions(files: list[str]) -> set[str]:''',
    '''_AR_KLJUCNA = re.compile(r"sub|function|property", re.IGNORECASE)
_CLAN_KLJUCNA = re.compile(r"sub|function|property|public|global", re.IGNORECASE)


def collect_definitions(files: list[str]) -> set[str]:''')
rep('''_AR_KLJUCNA = re.compile(r"sub|function|property", re.IGNORECASE)


def collect_arities(''', '''def collect_arities(''')

# memo logicke_izjave po NORMALIZOVANOM tekstu (cista funkcija svog argumenta)
rep('''    out, bafer, prvi = [], "", 0
    for i, red in enumerate(tekst.replace("\\r\\n", "\\n").split("\\n"), 1):
        deo, nastavlja = _izjava_deo(red)
''', '''    tekst = tekst.replace("\\r\\n", "\\n")
    kes = _IZJAVE_KES.get(tekst)
    if kes is not None:
        return kes                       # PROTOTIP: isti tekst -> isti rezultat
    out, bafer, prvi = [], "", 0
    for i, red in enumerate(tekst.split("\\n"), 1):
        deo, nastavlja = _izjava_deo(red)
''')
rep('''    if bafer:
        out.append((bafer, prvi))
    return out


_DEKL_PROC = re.compile(''', '''    if bafer:
        out.append((bafer, prvi))
    _IZJAVE_KES[tekst] = out
    return out


_IZJAVE_KES: dict = {}

_DEKL_PROC = re.compile(''')
open(p, "w", encoding="utf-8").write(s)
print("patched", p)
