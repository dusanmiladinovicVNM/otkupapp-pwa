#!/usr/bin/env python3
"""
Fix all remaining un-diacriticized string literals in VBA source.
Uses the same word-map approach as scan_captions.py but scans ALL string literals,
not just .Caption assignments.

EXCLUDED (not touched):
  - JSON property names (lines containing json/JSON context)
  - Column lookup parameters (RequireColumnIndex / GetColumnIndex)
  - VBA comment lines (start with ')
  - Identifiers (words outside of string literals)
  - ALL_CAPS_UNDERSCORE string literals (Poruka keys, config keys)
  - Test data column arrays in modBusinessFlowProTests.bas
  - modConfig.bas (all strings are schema identifiers: column/table names)
"""
import re
import os
import csv
from collections import defaultdict

REPO = "/home/user/otkupapp-pwa"
SRC  = os.path.join(REPO, "src-vba")
PORUKE = os.path.join(REPO, "localization", "poruke.csv")

EXTS = {'.bas', '.cls', '.frm'}

# --------------------------------------------------------------------------
# Word map (same as scan_captions.py)
# --------------------------------------------------------------------------
SUPPLEMENTARY = {
    'voca':       'voća',
    'voce':       'voće',
    'vocu':       'voću',
    'vocama':     'voćama',
    'voci':       'voći',
    'drzava':     'država',
    'drzave':     'države',
    'drzavu':     'državu',
    'drzavni':    'državni',
    'postanski':  'poštanski',
    'postanskog': 'poštanskog',
    'postanskom': 'poštanskom',
    'racun':      'račun',
    'racuna':     'računa',
    'racune':     'račune',
    'racunu':     'računu',
    'tekuci':     'tekući',
    'tekuceg':    'tekućeg',
    'tekucem':    'tekućem',
    'citanje':    'čitanje',
    'citanju':    'čitanju',
    'citanja':    'čitanja',
    'sifarnik':   'šifarnik',
    'sifarnika':  'šifarnika',
    'sifarnici':  'šifarnici',  # nominative plural, not locative
    'sortu':      'sortu',      # no change -- sorta = variety, no diacritics
    'vrstu':      'vrstu',      # no change
}

STRIP = str.maketrans({
    'š': 's', 'Š': 'S', 'ž': 'z', 'Ž': 'Z',
    'č': 'c', 'Č': 'C', 'ć': 'c', 'Ć': 'C',
    'đ': 'd', 'Đ': 'D',
})

def strip_sr(s):
    return s.translate(STRIP)

def build_word_map():
    wmap = {}
    conflict = set()
    with open(PORUKE, encoding='utf-8', newline='') as f:
        for row in csv.DictReader(f, delimiter=';'):
            for word in re.findall(r"[A-Za-zÀ-ɏ']+", row.get('Tekst', '')):
                if strip_sr(word) == word:
                    continue
                key = strip_sr(word).lower()
                val = word.lower()
                if key in wmap and wmap[key] != val:
                    conflict.add(key)
                else:
                    wmap[key] = val
    for k in conflict:
        wmap.pop(k, None)
    for k, v in SUPPLEMENTARY.items():
        if k not in wmap and v.lower() != k:
            wmap[k] = v.lower()
    return wmap

# ALL_CAPS_UNDERSCORE pattern: matches Poruka() keys, config keys, etc.
# These are internal identifiers, NOT user-visible strings.
ALL_CAPS_ID_RE = re.compile(r'^[A-Z][A-Z0-9_]+$')

def correct_text(text, wmap):
    """Return corrected text. Returns same text if no change.
    Skips ALL_CAPS_UNDERSCORE strings (message keys, config keys)."""
    # Protect ALL_CAPS_UNDERSCORE strings (e.g. Poruka key arguments)
    if ALL_CAPS_ID_RE.match(text):
        return text
    tokens = re.findall(r'[A-Za-z]+|[^A-Za-z]+', text)
    out = []
    changed = False
    for tok in tokens:
        if not re.match(r'[A-Za-z]', tok):
            out.append(tok)
            continue
        key = tok.lower()
        if key in wmap and wmap[key] != key:
            fix_lower = wmap[key]
            if tok[0].isupper() and not tok[1:].isupper():
                fix = fix_lower[0].upper() + fix_lower[1:]
            elif tok.isupper():
                fix = fix_lower.upper()
            else:
                fix = fix_lower
            out.append(fix)
            if fix != tok:
                changed = True
        else:
            out.append(tok)
    return ''.join(out) if changed else text

def vba_text_expr(text):
    if not text:
        return '""'
    parts = []
    chunk = []
    for c in text:
        if ord(c) < 128:
            chunk.append(c)
        else:
            if chunk:
                s = ''.join(chunk).replace('"', '""')
                parts.append(f'"{s}"')
                chunk = []
            parts.append(f'ChrW({ord(c)})')
    if chunk:
        s = ''.join(chunk).replace('"', '""')
        parts.append(f'"{s}"')
    return ' & '.join(parts) if parts else '""'

# --------------------------------------------------------------------------
# Dangerous line patterns to SKIP entirely
# --------------------------------------------------------------------------
SKIP_LINE_PATTERNS = [
    re.compile(r'RequireColumnIndex\s*\('),
    re.compile(r'GetColumnIndex\s*\('),
    re.compile(r'LookupValue\s*\('),
    re.compile(r'AppendRow\s*\('),
    re.compile(r'UpdateCell\s*\('),
    re.compile(r'"Kolicina":'),          # JSON property
    re.compile(r'"""Kolicina"""'),       # JSON property escaped
    re.compile(r'GS_KOLICINA\s*,\s*"'), # sync column name param
    re.compile(r'RequireOTKHeaderValue'),
    re.compile(r'ChrW\('),              # already fixed
]

# Files to completely skip
SKIP_FILES = {
    'modBusinessFlowProTests.bas',  # test data, column arrays
    'modSEFMapper.bas',             # JSON/XML field names
    'modMasterSync.bas',            # sync column name parameters
    'modConfig.bas',                # schema identifiers only (column/table names)
}

STRING_LITERAL_RE = re.compile(r'"([^"]*)"')

def process_file(fajl, wmap):
    if fajl in SKIP_FILES:
        return 0

    path = os.path.join(SRC, fajl)
    with open(path, encoding='latin-1', newline='') as f:
        original = f.read()

    lines = original.split('\n')
    new_lines = []
    n_changed = 0

    for lineno, line in enumerate(lines, 1):
        stripped = line.lstrip()

        # Skip comment lines
        if stripped.startswith("'"):
            new_lines.append(line)
            continue

        # Skip dangerous line patterns
        skip = False
        for pat in SKIP_LINE_PATTERNS:
            if pat.search(line):
                skip = True
                break
        if skip:
            new_lines.append(line)
            continue

        # Find and replace string literals in this line
        def replace_match(m):
            nonlocal n_changed
            orig_text = m.group(1)
            corrected = correct_text(orig_text, wmap)
            if corrected == orig_text:
                return m.group(0)  # no change
            expr = vba_text_expr(corrected)
            n_changed += 1
            return expr

        new_line = STRING_LITERAL_RE.sub(replace_match, line)
        new_lines.append(new_line)

    content = '\n'.join(new_lines)
    if content == original:
        return 0

    with open(path, 'w', encoding='latin-1', newline='') as f:
        f.write(content)

    return n_changed

def main():
    print('Building word map...')
    wmap = build_word_map()
    print(f'  {len(wmap)} word corrections')

    grand_total = 0
    for fname in sorted(os.listdir(SRC)):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in EXTS:
            continue
        n = process_file(fname, wmap)
        if n > 0:
            print(f'  {n:3d}  {fname}')
            grand_total += n

    print(f'\nUkupno zamenjeno: {grand_total} string literala')
    print()
    print('NIJE PROMENJENO (zahteva manuelnu intervenciju):')
    print('  frmOtkup.frx   -- staticke labele "Vrsta voca" / "Sorta voca"')
    print('                    Fiksirati u Excel Form Designer (Caption property)')
    print('  modSEFMapper   -- "Kolicina" u JSON izlazu (API field name)')
    print('  modMasterSync  -- column name parametri za sync')
    print('  modConfig.bas  -- schema identifiers (column/table names)')

if __name__ == '__main__':
    main()
