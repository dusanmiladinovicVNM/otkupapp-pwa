"""P2/P3: egzaktne brze putanje u vba_check.py (KOPIJA u proto/tools)."""
import sys
p = sys.argv[1]
s = open(p, encoding="utf-8").read()
def rep(old, new, n=1):
    global s
    assert s.count(old) == n, (s.count(old), old[:70])
    s = s.replace(old, new)

# (d) _izjava_deo: red bez apostrofa nema komentar; stanje stringa = parnost navodnika
rep('''    u_str = False
    kraj = len(red)
    for i, c in enumerate(red):
        if c == '"':
            u_str = not u_str
        elif c == "'" and not u_str:
            kraj = i
            break
    tekst = red[:kraj].rstrip()
''', '''    u_str = False
    kraj = len(red)
    ap = red.find("'")
    if ap < 0:
        # PROTOTIP: bez apostrofa nema komentara -- petlja bi samo brojala navodnike
        u_str = (red.count('"') & 1) == 1
    elif '"' not in red[:ap]:
        # prvi apostrof stoji pre svakog navodnika -> tu pocinje komentar
        kraj = ap
    else:
        for i, c in enumerate(red):
            if c == '"':
                u_str = not u_str
            elif c == "'" and not u_str:
                kraj = i
                break
    tekst = red[:kraj].rstrip()
''')

# (b) _strip_comment
rep('''    in_str = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return text[:i].rstrip()
    return text
''', '''    ap = text.find("'")
    if ap < 0:
        return text                      # PROTOTIP: nema apostrofa -> nema komentara
    if '"' not in text[:ap]:
        return text[:ap].rstrip()        # apostrof pre svakog navodnika
    in_str = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return text[:i].rstrip()
    return text
''')

# (a) _split_statements
rep('''    parts: list[tuple[str, bool]] = []
    depth, in_str, cur = 0, False, ""
    for i, ch in enumerate(text):
''', '''    if ":" not in text:
        # PROTOTIP: bez dvotacke nema razdvajanja -- cela linija je jedna naredba
        t = text.strip()
        return [(t, False)] if t else []
    parts: list[tuple[str, bool]] = []
    depth, in_str, cur = 0, False, ""
    for i, ch in enumerate(text):
''')

# (c) _strip_strings
rep('''    out, in_str = [], False
    for ch in text:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif not in_str:
            out.append(ch)
    return "".join(out)
''', '''    if '"' not in text:
        return text                      # PROTOTIP: nema literala
    delovi = text.split('"')
    # parni delovi su VAN stringa, neparni unutar; navodnici se svi cuvaju
    delovi[1::2] = [""] * len(delovi[1::2])
    return '"'.join(delovi)
''')

# (e) collect_arities: jedan poziv _split_statements + predfilter po kljucnoj reci
rep('''            glava = _split_statements(line)[0][0] if _split_statements(line) else line
''', '''            # PROTOTIP: PROC_DEF/DECLARE_DEF traze rec Sub/Function/Property, a
            # glava je deo ove linije -- linija bez tih reci ne moze da pogodi.
            if not _AR_KLJUCNA.search(line):
                i += 1
                continue
            _naredbe = _split_statements(line)
            glava = _naredbe[0][0] if _naredbe else line
''')
rep('''def collect_arities(files: list[str]) -> dict[str, tuple[int, float]]:''',
    '''_AR_KLJUCNA = re.compile(r"sub|function|property", re.IGNORECASE)


def collect_arities(files: list[str]) -> dict[str, tuple[int, float]]:''')

# (f) check_scalar_call: telo se cisti JEDNOM po proceduri, ne jednom po imenu
rep('''        for name, typ in _scalar_names(_strip_strings(header), body).items():
            hit = re.compile(r"(?<![\\w.$])" + re.escape(name) + r"\\s*\\(", re.IGNORECASE)
            for off, line in enumerate(body):
                if hit.search(_clean(line)):
''', '''        cisto = None
        for name, typ in _scalar_names(_strip_strings(header), body).items():
            hit = re.compile(r"(?<![\\w.$])" + re.escape(name) + r"\\s*\\(", re.IGNORECASE)
            if cisto is None:
                cisto = [_clean(line) for line in body]   # PROTOTIP: jednom po proceduri
            for off, line in enumerate(cisto):
                if hit.search(line):
''')
open(p, "w", encoding="utf-8").write(s)
print("patched", p)
