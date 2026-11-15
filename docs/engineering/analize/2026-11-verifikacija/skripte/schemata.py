import sys, os, importlib.util, collections, re
sys.dont_write_bytecode = True
ROOT = sys.argv[1]
spec = importlib.util.spec_from_file_location("_sab", os.path.join(ROOT, "tools", "sabotaza.py")); sab = importlib.util.module_from_spec(spec); spec.loader.exec_module(sab)
SRC = os.path.join(ROOT, "src-vba")
kes = {}
OPEN = [(r"^if\b.*\bthen\s*('.*)?$", "if"), (r"^for\b", "for"), (r"^do\b", "do"), (r"^while\b", "while"), (r"^select\s+case\b", "select"), (r"^with\b", "with")]
CLOSE = [(r"^end\s+if\b", "if"), (r"^next\b", "for"), (r"^loop\b", "do"), (r"^wend\b", "while"), (r"^end\s+select\b", "select"), (r"^end\s+with\b", "with")]
MID = r"^(else\b|elseif\b|case\b)"
DECL = r"^(dim|const|static|private|public|friend|declare|type|enum|option|#|sub|function|property|end\s+(sub|function|property|type|enum)|implements|attribute)\b"
LABEL = r"^[A-Za-z_]\w*:\s*('.*)?$"
def logicke(frag):
    redovi = frag.rstrip("\n").split("\n")
    return redovi
def ocena(frag):
    """None ako je fragment niz celih, uravnotezenih naredbi; inace razlog."""
    redovi = logicke(frag)
    if redovi and re.search(r"\s_\s*$", redovi[-1]): return "zavrsava nastavkom reda"
    stek = []
    nastavak = False
    for r in redovi:
        t = r.strip().lower()
        if nastavak:
            nastavak = bool(re.search(r"\s_\s*$", r)); 
            if not nastavak and prvi_if and re.search(r"\bthen\s*('.*)?$", t): stek.append("if")
            continue
        nastavak = bool(re.search(r"\s_\s*$", r))
        prvi_if = t.startswith("if ") 
        if not t or t.startswith("'"): continue
        if re.match(DECL, t): return "deklaracija/potpis"
        if re.match(LABEL, r.strip()) and not t.startswith("case") and not t.startswith("else"): return "labela"
        if re.match(MID, t):
            if not stek: return "Else/Case bez svog bloka"
            continue
        z = False
        for p, k in CLOSE:
            if re.match(p, t):
                if not stek or stek[-1] != k: return "zatvara tudji blok"
                stek.pop(); z = True; break
        if z: continue
        if nastavak:
            continue
        for p, k in OPEN:
            if re.match(p, t):
                stek.append(k); break
    if stek: return "otvara blok bez zatvaranja"
    return None
C = collections.Counter(); ok = 0
for ime,(f,staro,novo,t,tv) in sab.SABOTAZE.items():
    if f not in kes: kes[f] = sab._procitaj(os.path.join(SRC, f))[0]
    tekst = kes[f]
    poz = tekst.find("\n" + staro)
    razlog = None
    if poz >= 0:
        pret = tekst[:poz].rsplit("\n", 1)[-1]
        if re.search(r"\s_\s*$", pret): razlog = "sidro pocinje usred nastavka reda"
    razlog = razlog or ocena(staro) and ("sidro: " + ocena(staro)) or ocena(novo) and ("zamena: " + ocena(novo))
    if razlog: C[razlog] += 1
    else: ok += 1
N = len(sab.SABOTAZE)
print("mehanicki obmotljivo u If-prekidac (grubo): %d / %d (%.0f%%)" % (ok, N, 100*ok/N))
for r, n in C.most_common(): print("  %4d  %s" % (n, r))
