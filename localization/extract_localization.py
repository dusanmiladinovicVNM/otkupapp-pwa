#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA lokalizacija - ekstraktor string literala sa ne-ASCII znakovima.
Citanje u cp1250, izlaz u UTF-8 CSV.
v2 - spajanje continuation linija, Style* kao DA, bolji kljucevi.
"""
import re
import csv
from pathlib import Path
from collections import defaultdict

SRC_DIR = "/home/user/otkupapp-pwa/src-vba"
OUT_DIR = "/tmp/claude-0/-home-user-otkupapp-pwa/e28b3512-2305-56f8-a36e-1867a245d9fb/scratchpad"

SERBIAN_CHARS = set('šžčćđŠŽČĆĐ')
GERMAN_CHARS  = set('äöüßÄÖÜ')

def has_serbian(s): return any(c in SERBIAN_CHARS for c in s)
def has_german(s):  return any(c in GERMAN_CHARS  for c in s)
def has_nonascii(s): return any(ord(c) > 127 for c in s)

# -----------------------------------------------------------------------
# Parser: logicke linije (spajanje VBA _ continuation)
# -----------------------------------------------------------------------
def build_logical_lines(raw_lines):
    """
    Vraca dict: original_lineno(1-based) -> logical_line (string)
    Continuation linije ( _ na kraju) se spajaju u jednu logicku liniju.
    """
    result = {}
    i = 0
    n = len(raw_lines)
    while i < n:
        parts = []
        start = i
        line = raw_lines[i].rstrip('\r\n')
        parts.append(line)
        # Nastavi dok linija zavrsava sa ' _'
        while line.rstrip().endswith(' _') or line.rstrip().endswith('\t_'):
            # Ukloni trailing _
            trimmed = line.rstrip()
            if trimmed.endswith(' _') or trimmed.endswith('\t_'):
                parts[-1] = trimmed[:-2]
            i += 1
            if i >= n:
                break
            line = raw_lines[i].rstrip('\r\n')
            parts.append(line)
        # Logicka linija = spojeni delovi (drugi i dalji se lstrip-uju za indentaciju)
        logical = parts[0] + ' ' + ' '.join(p.lstrip() for p in parts[1:]) if len(parts) > 1 else parts[0]
        # Svi originalni redovi mapiraju na istu logicku liniju
        for offset in range(len(parts)):
            result[start + offset + 1] = logical  # 1-indexed
        i += 1
    return result

# -----------------------------------------------------------------------
# Parser string literala u jednoj liniji
# -----------------------------------------------------------------------
def extract_strings(line):
    """
    Vraca listu (content, quote_start, quote_end) za svaki string literal.
    Staje na komentar (' izvan stringa).
    """
    result = []
    i, n = 0, len(line)
    in_str = False
    start = -1
    buf = []
    while i < n:
        c = line[i]
        if in_str:
            if c == '"':
                if i + 1 < n and line[i+1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                else:
                    result.append((''.join(buf), start, i))
                    in_str = False
                    buf = []
            else:
                buf.append(c)
        else:
            if c == '"':
                in_str = True
                start = i
                buf = []
            elif c == "'":
                break
        i += 1
    return result

# -----------------------------------------------------------------------
# Klasifikacija vidljivosti
# -----------------------------------------------------------------------

# Fonkcije koje postavljaju Caption na kontroli (drugi/treci arg = vidljiv tekst)
# SAMO one sa potvrdjenim text/caption parametrom.
_STYLE_FUNCS = re.compile(
    r'\bStyle(?:MenuButton|PrimaryButton|ExitButton|StornoButton'
    r'|SectionHeader|Subtitle|FrameTitleLabel|NavButton)\b'
    r'|\bSetColumnHeader\b'       # Private u frm - postavlja tekst labele kolone
    r'|\bOpenContentForm\b'       # 3. arg = sectionTitle -> lblStatus.caption
    r'|\bShowDashboardArea\b'     # statusText arg -> prikazuje se u labeli
    r'|\bSetGeoStatus\b'          # message arg -> prikazuje se u labeli forme
    r'|\bSetCaption\b|\bSetLabel\b|\bSetTitle\b',
    re.IGNORECASE
)

# Sigurno korisnicko-vidljivo
_VISIBLE_RE = re.compile(
    r'\bMsgBox\b'
    r'|\.(Caption|Text|ControlTipText|StatusBar|Title|HeaderText)\s*='
    r'|\.List\s*(?:\([^)]*\))?\s*='   # .List = "..." ili .List(r,c) = "..."
    r'|\.AddItem\b'
    r'|\bCells\s*\(.*\)\s*='
    r'|\bRange\s*\(.*\)\s*='
    r'|\.Cells\s*\(.*\)\s*='
    ,re.IGNORECASE
)

# Sigurno NEVIDLJIVO
_INVISIBLE_RE = re.compile(
    r'\bDebug\s*\.\s*Print\b'
    r'|\bLogError\b|\bLogWarning\b|\bJournalLog\b|\bWriteLog\b'
    r'|\bCall\s+(?:Log|Journal|Write)'
    r'|\bShell\b'
    r'|\bOpen\s+.+\s+For\b'   # file open
    ,re.IGNORECASE
)

# Err.Raise - Description arg (3. arg) = moze biti vidljiv kroz Err.Description
_ERR_RAISE_RE = re.compile(r'\bErr\s*\.\s*Raise\b', re.IGNORECASE)

# Poredi vrednost (ne prikazuje)
_COMPARE_RE = re.compile(
    r'(?:If|ElseIf|And|Or|Not)\s+\S+\s*=\s*$'
    r'|=\s*"[^"]*"\s*(?:Then|And|Or|$)',
    re.IGNORECASE
)

def classify(raw_line, logical_line, str_start, filename):
    """
    Vraca: 'DA' | 'NE' | 'SUMNJIVO'
    Koristimo logicku liniju za kontekst, raw za poziciju.
    """
    # Kontekst: SVE PRE STRINGA u logickoj liniji
    # Pozicija stringa u logickoj liniji je priblizna (string moze biti u continuation)
    # Koristimo before iz RAW linije za poziciju + logicku za opsti kontekst
    before_raw  = raw_line[:str_start]
    logical_up  = logical_line.upper()

    # 1. Cisti komentar na pocetku linije
    if raw_line.lstrip().startswith("'"):
        return 'NE'

    # 2. Sigurno nevidljivo u logickoj liniji
    if _INVISIBLE_RE.search(logical_line):
        return 'NE'

    # 3. Style* caption funkcije u logickoj liniji = sigurno vidljivo
    if _STYLE_FUNCS.search(logical_line):
        return 'DA'

    # 4. MsgBox u logickoj liniji = vidljivo
    if re.search(r'\bMsgBox\b', logical_line, re.IGNORECASE):
        return 'DA'

    # 5. Property assignment u logickoj liniji
    if _VISIBLE_RE.search(logical_line):
        return 'DA'

    # 6. Err.Raise description = SUMNJIVO (moze biti vidljivo)
    if _ERR_RAISE_RE.search(logical_line):
        return 'SUMNJIVO'

    # 7. Poredi vrednost (If x = "str") - NE
    if _COMPARE_RE.search(logical_line):
        return 'NE'

    # 8. Izvestaj / stampa moduli + celijska vrednost
    base = Path(filename).stem.lower()
    report_mods = ('modizvestaj', 'modprint', 'moddokumenta', 'modfaktura',
                   'modpaletnilist', 'modpaletnilistui', 'modsledljivost',
                   'frmizvestaj', 'frmpalete', 'frmdokumenta')
    if any(base == m for m in report_mods):
        if re.search(r'Cells|Range|\.Value\s*=|addItem', logical_line, re.IGNORECASE):
            return 'DA'

    return 'SUMNJIVO'

# -----------------------------------------------------------------------
# Generisanje kljuca
# -----------------------------------------------------------------------

_FILE_PREFIX_MAP = {
    'thisworkbook': 'APP',
    'frmagrohemija': 'AGRO',
    'frmbankaexportpregled': 'BANKA',
    'frmbankaimport': 'BANKA',
    'frmdokumenta': 'DOK',
    'frmexcelmini': 'EXCEL',
    'frmfakturisanje': 'FAK',
    'frmizvestaj': 'RPT',
    'frmmarza': 'MARZA',
    'frmmaticnipodaci': 'MATICNI',
    'frmotkup': 'OTKUP',
    'frmotkupapp': 'OTKUP',
    'frmpalete': 'PAL',
    'frmsef': 'SEF',
    'frmsledljivost': 'SLED',
    'frmsplash': 'SPLASH',
    'frmstammdaten': 'STM',
    'modagrohemija': 'AGRO',
    'modambalaza': 'AMBA',
    'modarrayutils': 'UTIL',
    'modautohlad': 'HLAD',
    'modbanka': 'BANKA',
    'modbankaimport': 'BANKA',
    'modbankaimportparserpdftotext': 'BANKA',
    'modbankaexportpregled': 'BANKA',
    'modbrojevi': 'BROJ',
    'modbuildguard': 'BUILD',
    'modcenovnik': 'CEN',
    'modclipboard': 'CLIP',
    'modcombobinding': 'UI',
    'modconfig': 'CFG',
    'moddataaccess': 'DA',
    'moddocstyle': 'STYLE',
    'moddokumenta': 'DOK',
    'mode2ereleasegate': 'BUILD',
    'modfaktura': 'FAK',
    'modgeoparcel': 'GEO',
    'modgoogleauth': 'GAUTH',
    'modgooglesheets': 'GS',
    'modgooglesyncorch': 'SYNC',
    'modhelpers': 'UTIL',
    'modhladnjaca': 'HLAD',
    'modhttputils': 'HTTP',
    'modizvestaj': 'RPT',
    'modjournaling': 'LOG',
    'modkarticadetalji': 'KART',
    'modkooperant': 'KOOP',
    'modkvalitet': 'KVAL',
    'modlicense': 'LIC',
    'modlicence': 'LIC',
    'modlogerror': 'LOG',
    'modml': 'ML',
    'modmain': 'MAIN',
    'modmalina': 'MAL',
    'modmarza': 'MARZA',
    'modmastersync': 'SYNC',
    'modmaticnilookups': 'MATICNI',
    'modmeteo': 'METEO',
    'modmigracija': 'MIG',
    'modmonitoring': 'MON',
    'modnovac': 'NOV',
    'modotkup': 'OTKUP',
    'modotkupblok': 'OTKUP',
    'modpaletnilist': 'PAL',
    'modpaletnilistui': 'PAL',
    'modparse': 'PARSE',
    'modpodesavanja': 'CFG',
    'modpregledlistova': 'RPT',
    'modprint': 'PRINT',
    'modprodhealthcheck': 'BUILD',
    'modsefclient': 'SEF',
    'modsefsvc': 'SEF',
    'modschema': 'SETUP',
    'modsetup': 'SETUP',
    'modsledljivost': 'SLED',
    'modstammdatensync': 'STM',
    'modstanica': 'STAN',
    'modstorno': 'STOR',
    'modtheme': 'UI',
    'modtrial': 'LIC',
    'modupdategate': 'BUILD',
    'modvbatools': 'UTIL',
    'modwindow': 'UI',
}

def _file_prefix(filename):
    stem = Path(filename).stem.lower()
    if stem in _FILE_PREFIX_MAP:
        return _FILE_PREFIX_MAP[stem]
    # Fallback: ukloni frm/mod/cls
    for p in ('frm', 'mod', 'cls'):
        if stem.startswith(p):
            stem = stem[len(p):]
            break
    return stem[:6].upper()

_TRANSLIT = str.maketrans(
    'šžčćđŠŽČĆĐäöüßÄÖÜéèêàùâî',
    'szccdSZCCDaouBAOUeeeauai'
)
def _ascii(s):
    return s.translate(_TRANSLIT)

# Specijalni stringovi koji dobijaju fiksne kljuceve
_SPECIAL_KEYS = {
    ' — ':    'UI_SEPARATOR_EMDASH_SPACE',
    '—':      'UI_SEPARATOR_EMDASH',
    '  •  ':  'UI_SEPARATOR_BULLET',
    ' • ':    'UI_SEPARATOR_BULLET_SPACE',
    '•':      'UI_SEPARATOR_BULLET',
    '—nedostaje—': 'UI_PLACEHOLDER_NEDOSTAJE',
}

def make_key(content, filename, logical_line):
    # Specijalni separatori
    if content in _SPECIAL_KEYS:
        return _SPECIAL_KEYS[content]

    mod = _file_prefix(filename)
    ctx = logical_line.upper()

    # Tip na osnovu konteksta
    if _STYLE_FUNCS.search(logical_line):
        typ = 'LBL'
    elif re.search(r'\bMsgBox\b', logical_line, re.IGNORECASE):
        # Pokusaj da razlikujemo greske od info poruka
        if re.search(r'(?:gresk|error|err\.|fault|fail)', logical_line, re.IGNORECASE):
            typ = 'ERR'
        else:
            typ = 'MSG'
    elif re.search(r'\.(Caption|Text|ControlTipText|StatusBar|Title)\s*=', logical_line, re.IGNORECASE):
        typ = 'LBL'
    elif re.search(r'\.AddItem\b', logical_line, re.IGNORECASE):
        typ = 'LST'
    elif re.search(r'\bErr\s*\.\s*Raise\b', logical_line, re.IGNORECASE):
        typ = 'ERR'
    elif re.search(r'\bCells\b|\bRange\b', logical_line, re.IGNORECASE):
        typ = 'RPT'
    else:
        typ = 'MSG'

    # Opis iz prvog dela teksta (3 reci, ascii)
    ascii_content = _ascii(content)
    words = re.sub(r'[^a-zA-Z0-9]+', ' ', ascii_content).split()
    sig_words = [w for w in words if len(w) >= 3]
    if not sig_words:
        sig_words = [w for w in words if len(w) >= 2]
    if not sig_words:
        sig_words = ['STR']
    desc = '_'.join(w.upper() for w in sig_words[:3])

    return f"{mod}_{typ}_{desc}"

# -----------------------------------------------------------------------
# Nemacki - predlog prevoda
# -----------------------------------------------------------------------
_DE_VOCAB = {
    'fehlgeschlagen': 'nije uspelo',
    'Fehler': 'greska',
    'Prüfe': 'proveri',
    'Einträge': 'stavki/redova',
    'Möglicher': 'moguc',
    'Datenverlust': 'gubitak podataka',
    'Absturz': 'pad aplikacije',
    'Austausch': 'razmena',
    'auswählen': 'izaberi',
    'PDF': 'PDF',
    'für': 'za',
    'nach': 'posle',
    'hat': 'ima',
    'Excel': 'Excel',
}

def propose_sr(german_text):
    hits = []
    for de, sr in _DE_VOCAB.items():
        if de.lower() in german_text.lower():
            hits.append(f"{de}→{sr}")
    if hits:
        return f"[PREVESTI ({'; '.join(hits[:4])})]"
    return "[PREVESTI - manuelno]"

# -----------------------------------------------------------------------
# Obrada fajlova
# -----------------------------------------------------------------------
def process():
    src = Path(SRC_DIR)
    EXTS = {'.bas', '.cls', '.frm', '.doccls'}

    rows   = []   # DA + SUMNJIVO
    ne_rows = []  # NE (za informaciju)

    text2key   = {}                  # tekst -> kljuc (dedup)
    key2texts  = defaultdict(set)    # kljuc -> {tekstovi} (za kolizije)

    def get_key(content, filename, logical_line):
        if content in text2key:
            return text2key[content]
        base = make_key(content, filename, logical_line)
        key = base
        n = 2
        while key in key2texts and content not in key2texts[key]:
            key = f"{base}_{n}"
            n += 1
        text2key[content] = key
        key2texts[key].add(content)
        return key

    files = sorted([f for f in src.iterdir()
                    if f.suffix in EXTS])

    for fpath in files:
        fname = fpath.name
        try:
            raw_lines = fpath.read_text(encoding='cp1250', errors='replace').splitlines()
        except Exception as e:
            print(f"  GRESKA {fname}: {e}")
            continue

        logical_map = build_logical_lines(raw_lines)

        for lineno, raw_line in enumerate(raw_lines, 1):
            # Preskoci Attribute header linije u .frm
            if raw_line.startswith('Attribute ') or raw_line.startswith('VERSION '):
                continue
            # Preskoci ciste komentare
            if raw_line.lstrip().startswith("'"):
                continue

            logical = logical_map.get(lineno, raw_line)
            literals = extract_strings(raw_line)

            for content, s_start, s_end in literals:
                if not has_nonascii(content):
                    continue

                vis = classify(raw_line, logical, s_start, fname)
                before = raw_line[:s_start].rstrip()
                key = get_key(content, fname, logical)

                german = has_german(content)
                row = {
                    'kljuc':       key,
                    'tekst':       content,
                    'fajl':        fname,
                    'linija':      lineno,
                    'tip':         vis,
                    'german':      german,
                    'before':      before,
                    'logical':     logical.strip()[:120],
                    'nemacki_orig': content if german else '',
                    'sr_predlog':  propose_sr(content) if german else '',
                }

                if vis == 'NE':
                    ne_rows.append(row)
                else:
                    rows.append(row)

    return rows, ne_rows, text2key

# -----------------------------------------------------------------------
# Pisanje CSV
# -----------------------------------------------------------------------
def write_poruke(rows, out_path):
    """Jedan red po jedinstvenom kljucu (dedup); DA pre SUMNJIVO."""
    seen = {}   # key -> {row, duplikati:[]}
    rows_sorted = sorted(rows, key=lambda r: (0 if r['tip']=='DA' else 1,
                                               r['fajl'], r['linija']))
    for row in rows_sorted:
        k = row['kljuc']
        if k not in seen:
            seen[k] = {'row': row, 'dup': []}
        else:
            seen[k]['dup'].append(f"{row['fajl']}:{row['linija']}")

    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['Kljuc', 'Tekst', 'Fajl', 'Linija', 'Tip', 'NemackiOriginal', 'Napomena'])
        for k, data in seen.items():
            row  = data['row']
            dups = data['dup']
            nap  = ''
            if row['german']:
                nap = row['sr_predlog']
            if dups:
                extra = 'Koristi se i u: ' + ', '.join(dups[:5])
                nap = (nap + ' | ' + extra) if nap else extra
            w.writerow([k, row['tekst'], row['fajl'], row['linija'],
                        row['tip'], row['nemacki_orig'], nap])
    return len(seen)

def write_zamene(rows, out_path):
    """Jedan red po svakom pojavljivanju."""
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['Fajl', 'Linija', 'OrigLiteral', 'Zamena', 'Napomena'])
        for row in sorted(rows, key=lambda r: (r['fajl'], r['linija'])):
            orig   = '"' + row['tekst'] + '"'
            zamena = f'Poruka("{row["kljuc"]}")'
            nap    = row['before'][-80:] if row['before'] else row['logical'][:80]
            # Detektuj konkatenaciju
            if '&' in row['logical']:
                nap = 'CONCAT: ' + row['logical'][:100]
            w.writerow([row['fajl'], row['linija'], orig, zamena, nap])
    return len(rows)

# -----------------------------------------------------------------------
# Dijagnostika
# -----------------------------------------------------------------------
def stats(rows, ne_rows):
    da       = [r for r in rows if r['tip'] == 'DA']
    sumnjivo = [r for r in rows if r['tip'] == 'SUMNJIVO']
    nemacki  = [r for r in rows if r['german']]

    print(f"\n{'='*60}")
    print(f"STATISTIKA")
    print(f"{'='*60}")
    print(f"Korisnicko-vidljivi (DA):      {len(da)}")
    print(f"Sumnjivo (SUMNJIVO):           {len(sumnjivo)}")
    print(f"Nevidljivi (NE, preskoceni):   {len(ne_rows)}")
    print(f"Sa nemackim znakovima:         {len(nemacki)}")

    print(f"\n{'='*60}")
    print("KORISNICKO-VIDLJIVI (DA) - prvih 20")
    print(f"{'='*60}")
    for r in da[:20]:
        print(f"  [{r['kljuc']:<48}] '{r['tekst'][:55]}'")
        print(f"   {r['fajl']}:{r['linija']}")

    print(f"\n{'='*60}")
    print("SUMNJIVO - prvih 25")
    print(f"{'='*60}")
    for r in sumnjivo[:25]:
        print(f"  [{r['kljuc']:<48}] '{r['tekst'][:55]}'")
        print(f"   {r['fajl']}:{r['linija']} | kontekst: {r['before'][-70:]}")

    print(f"\n{'='*60}")
    print("NEMACKI STRINGOVI (svi)")
    print(f"{'='*60}")
    for r in nemacki:
        print(f"  '{r['tekst']}'")
        print(f"   {r['fajl']}:{r['linija']} | TIP: {r['tip']}")
        print(f"   Predlog SR: {r['sr_predlog']}")
        print()

def main():
    print(f"Citam fajlove iz {SRC_DIR} ...")
    rows, ne_rows, text2key = process()

    stats(rows, ne_rows)

    pp = f"{OUT_DIR}/poruke.csv"
    zp = f"{OUT_DIR}/zamene.csv"
    u  = write_poruke(rows, pp)
    t  = write_zamene(rows, zp)

    print(f"\n{'='*60}")
    print(f"IZLAZ")
    print(f"{'='*60}")
    print(f"poruke.csv : {u} jedinstvenih kljuceva  -> {pp}")
    print(f"zamene.csv : {t} redova (sva pojav.)    -> {zp}")
    print()
    print(f"Ukupno stringova sa ne-ASCII: {len(rows)+len(ne_rows)}")
    print(f"  DA        : {sum(1 for r in rows if r['tip']=='DA')}")
    print(f"  SUMNJIVO  : {sum(1 for r in rows if r['tip']=='SUMNJIVO')}")
    print(f"  NE (skip) : {len(ne_rows)}")

if __name__ == '__main__':
    main()
