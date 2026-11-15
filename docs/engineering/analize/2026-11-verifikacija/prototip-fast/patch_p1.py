"""P1: _dodela_tudjoj_proceduri bez ponovnog parsiranja celog mutiranog fajla.
Primenjuje se na KOPIJU ($S/proto/tools/sabotaza.py)."""
import sys, re
p = sys.argv[1]
s = open(p, encoding="utf-8").read()

old_head = '''    idx = tekst.find("\\n" + staro)
    if idx < 0:
        return None
    red = tekst[:idx + 1].count("\\n")
'''
assert s.count(old_head) == 1
old_body_start = s.index(old_head)
old_body_end = s.index("    # Property se NE skuplja -- time je i izuzeta.", old_body_start)
new_body = '''    idx = tekst.find("\\n" + staro)
    if idx < 0:
        return None

    # PROTOTIP P1a: zamena bez ijedne linije-dodele ne moze dati nalaz -- petlja
    # ispod bi presla `novo` i vratila None, pa se parsiranje fajla ne placa.
    if not any(_ZAMENA_DODELA.match(l) for l in novo.split("\\n")):
        return None
    red = tekst[:idx + 1].count("\\n")

    # PROTOTIP P1b: dogadjaji (otvaranje/zatvaranje procedure) zdravog fajla se
    # racunaju JEDNOM po fajlu; za mutirani tekst se ponovo klasifikuju samo
    # redovi koje zamena dodiruje. Automat nad dogadjajima je isti kao u
    # _procedure_po_redu, pa je rezultat identican parsiranju mutiranog teksta.
    redovi0, dog0 = _dogadjaji_fajla(tekst)
    n_staro = staro.count("\\n") + 1          # redova koje sidro dodiruje
    n_novo = novo.count("\\n") + 1
    kraj0 = red + n_staro                    # prvi nedirnut red u zdravom
    pomak = n_novo - n_staro
    # mutirani region = novo + rep poslednjeg dodirnutog reda zdravog teksta
    rep_poc = idx + 1 + len(staro)
    rep_kraj = tekst.find("\\n", rep_poc)
    rep = tekst[rep_poc:] if rep_kraj < 0 else tekst[rep_poc:rep_kraj]
    region = (novo + rep).split("\\n")
    redovi = redovi0[:red] + region + redovi0[kraj0:]
    dog = [d for d in dog0 if d[0] < red]
    for k, r in enumerate(region):
        d = _dogadjaj_reda(r)
        if d:
            dog.append((red + k,) + d)
    dog += [(d[0] + pomak,) + d[1:] for d in dog0 if d[0] >= kraj0]
    procs = _procedure_iz_dogadjaja(dog)
    unutar, lokalna = None, set()
    for pime, _v, a, b in procs:
        if a <= red <= b:
            unutar = pime.lower()
            lokalna = _lokalna_imena(redovi, a, b)
            break

'''
s = s[:old_body_start] + new_body + s[old_body_end:]

helpers = '''
_dog_kes = {}


def _dogadjaj_reda(red: str):
    """None, ili (True, vrsta, ime) za otvaranje, (False, None, None) za zatvaranje."""
    t = red.strip()
    m = _PROC_OTVARA.match(t)
    if m:
        return (True, m.group(1).lower(), m.group(2))
    if _PROC_ZATVARA.match(t):
        return (False, None, None)
    return None


def _dogadjaji_fajla(tekst: str):
    """(redovi, [(red, otvara, vrsta, ime)]) -- jednom po SADRZAJU fajla."""
    k = _dog_kes.get(id(tekst))
    if k is not None and k[0] is tekst:
        return k[1], k[2]
    redovi = tekst.split("\\n")
    dog = []
    for i, r in enumerate(redovi):
        d = _dogadjaj_reda(r)
        if d:
            dog.append((i,) + d)
    _dog_kes[id(tekst)] = (tekst, redovi, dog)
    return redovi, dog


def _procedure_iz_dogadjaja(dog: list) -> list:
    """Isti automat kao _procedure_po_redu, nad dogadjajima umesto nad redovima."""
    out, ime, vrsta, poc = [], None, None, 0
    for i, otvara, v, n in dog:
        if otvara:
            vrsta, ime, poc = v, n, i
        elif ime:
            out.append((ime, vrsta, poc, i))
            ime, vrsta = None, None
    return out


def _dodela_tudjoj_proceduri('''
assert s.count("\ndef _dodela_tudjoj_proceduri(") == 1
s = s.replace("\ndef _dodela_tudjoj_proceduri(", helpers, 1)
open(p, "w", encoding="utf-8").write(s)
print("patched", p)
