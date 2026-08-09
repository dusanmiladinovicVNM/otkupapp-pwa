#!/usr/bin/env python3
"""
check-sef-asserts.py -- staticka provera da testovi ne tvrde nesto sto kod ne radi.

Povod (RF-22, review runda 9): promenjen je klasifikator SEF statusa
("Storno" izdvojen iz TERMINAL u zasebnu klasu), a stari assert u
`Test_SEFOfficialStatusEnumClassified` je ostao da ocekuje TERMINAL. Pošto
`RunSEFTestSuite` ima tvrd gate, suite bi sigurno pao -- a to se u repou ne vidi
jer CI ne pokrece Excel.

Skripta PARSIRA produkcioni izvor (ne kopira logiku rucno):
  * `ClassifySEFExternalStatus` iz modSEFStatusSync
  * tabelu dozvoljenih tranzicija iz `ValidateAllowedTransition` u modSEFValidator
  * `WF_*` konstante iz modConfig
  * mapu `klasa -> ciljno stanje` iz `SEFRefreshTargetState`
pa uporedjuje sa SVAKIM `AssertEquals` nad `ClassifySEFExternalStatus` i
`SEFRefreshTargetState` u modSEFTests.

OGRANICENJE (posteno receno): checker je **parcijalno** izveden iz izvora. Dva
pravila planera su ovde rucno preslikana -- grana za ERROR/UNKNOWN
(SENDING -> UNKNOWN, SENT -> SYNC_ERROR) i pravilo mosta preko SEF_SENT. Ako se
BAS ta dva pravila promene u VBA planeru, mora i ovde. Sve ostalo (spisak klasa,
ciljna stanja po klasi, tabela dozvoljenih tranzicija, WF_* konstante) cita se
iz koda, pa ne moze da zastari.

Pokretanje iz korena repoa:  python3 tools/check-sef-asserts.py
Izlaz 0 = nema neslaganja; 1 = ispisana lista neslaganja.
"""
import re, sys

# --- mirror of ClassifySEFExternalStatus (parsed from source, not hand-copied) ---
src = open('src-vba/modSEFStatusSync.bas').read()
body = src[src.index('Public Function ClassifySEFExternalStatus'):src.index('End Function', src.index('Public Function ClassifySEFExternalStatus'))]
mapping = {}
cur = None
for line in body.split('\n'):
    st = line.strip()
    m = re.match(r'^Case (.+)$', st)
    if m and 'Else' not in st:
        cur = [v.strip().strip('"').upper() for v in m.group(1).split(',')]
    m2 = re.match(r'^ClassifySEFExternalStatus = (SEF_CLS_\w+)$', st)
    if m2 and cur:
        for v in cur: mapping[v] = m2.group(1)
        cur = None
DEFAULT = 'SEF_CLS_UNKNOWN'
def classify(s): return mapping.get(s.strip().upper(), DEFAULT)

# --- mirror of the transition planner (state machine parsed from validator) ---
val = open('src-vba/modSEFValidator.bas').read()
vb = re.sub(r'_\n\s*', ' ', val[:val.index('InvalidTransition:')])
ALLOWED, cur = {}, None
for line in vb.split('\n'):
    st = line.strip()
    m = re.match(r'^Case (WF_\w+)$', st)
    if m: cur = m.group(1); ALLOWED.setdefault(cur, set()); continue
    if cur:
        m2 = re.match(r'^If newState <> (WF_\w+) Then GoTo InvalidTransition$', st)
        if m2: ALLOWED[cur] = {m2.group(1)}; cur = None; continue
        m3 = re.match(r'^Case ((?:WF_\w+\s*,?\s*)+)$', st)
        if m3:
            ALLOWED[cur] |= {x.strip() for x in m3.group(1).split(',') if x.strip()}
        if st == 'GoTo InvalidTransition' and not ALLOWED.get(cur): ALLOWED[cur] = set()
CONST = dict(re.findall(r'Public Const (WF_SEF_\w+|WF_LOCAL_\w+) As String = "(\w+)"', open('src-vba/modConfig.bas').read()))
def val_of(n): return CONST.get(n, n)
def allowed(old, new):
    for k, vs in ALLOWED.items():
        if val_of(k).upper() == old.strip().upper():
            return any(val_of(v).upper() == new.strip().upper() for v in vs)
    return False
# --- class -> desired-state map, PARSED from SEFRefreshTargetState -------------
# Ranije je ovaj mapping bio rucno prepisan ovde, pa je alat koji hvata
# duplirano znanje i sam bio duplirano znanje. Sada se cita iz izvora.
pl = src[src.index('Public Function SEFRefreshTargetState'):]
pl = pl[:pl.index('End Function')]
pl = re.sub(r'_\n\s*', ' ', pl)
DESIRED, pend = {}, []
for line in pl.split('\n'):
    st = line.strip()
    if st.startswith("'"):
        continue
    m = re.match(r'^Case ((?:SEF_CLS_\w+\s*,?\s*)+)$', st)
    if m:
        pend = [x.strip() for x in m.group(1).split(',') if x.strip()]
        continue
    m2 = re.match(r'^desired = (WF_\w+)$', st)
    if m2 and pend:
        for c in pend:
            DESIRED[c] = m2.group(1)
        pend = []
assert DESIRED, 'planer mapping nije isparsiran'

# HAND-MIRRORED (jedini deo koji nije izveden iz izvora): grana za ERROR/UNKNOWN
# (SENDING -> UNKNOWN, SENT -> SYNC_ERROR, ostalo bez promene) i pravilo mosta
# preko SEF_SENT. Ako se TA dva pravila promene u VBA planeru, mora i ovde --
# ostatak (klase, ciljna stanja, tabela tranzicija, konstante) se cita iz koda.
def plan(cur_state, cls):
    c = val_of(cur_state).upper()
    if cls in DESIRED:
        d = val_of(DESIRED[cls])
    else:
        if c == 'SEF_SENDING': d = 'SEF_UNKNOWN'
        elif c == 'SEF_SENT': d = 'SEF_SYNC_ERROR'
        else: return ''
    if not c: return d
    if c == d.upper(): return ''
    if allowed(c, d): return d
    if allowed(c, 'SEF_SENT') and allowed('SEF_SENT', d): return 'SEF_SENT'
    return ''

# --- check every AssertEquals over those two functions in the test module ---
tests = open('src-vba/modSEFTests.bas', encoding='latin-1').read().replace('\r\n', '\n')
tests = re.sub(r'_\n\s*', ' ', tests)          # join VBA line continuations
bad = 0

for expected, arg in re.findall(r'AssertEquals\s+(SEF_CLS_\w+),\s*ClassifySEFExternalStatus\("([^"]*)"\)', tests):
    got = classify(arg)
    if got != expected:
        print(f'MISMATCH classify("{arg}"): test expects {expected}, code gives {got}'); bad += 1

for expected, st, cls in re.findall(r'AssertEquals\s+(WF_\w+|""),\s*SEFRefreshTargetState\((WF_\w+),\s*(SEF_CLS_\w+)\)', tests):
    exp = '' if expected == '""' else val_of(expected).upper()
    got = plan(st, cls)
    if got.upper() != exp:
        print(f'MISMATCH plan({st}, {cls}): test expects "{exp}", code gives "{got}"'); bad += 1

print('assert-vs-code mismatches:', bad)
sys.exit(1 if bad else 0)
