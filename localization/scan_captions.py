#!/usr/bin/env python3
"""
Scan VBA source for .Caption = "string" assignments.
Builds a word-correction map from poruke.csv (where diacritics are already correct),
applies word-level suggestions, outputs caption_review.csv for human review.

Workflow:
  1. python scan_captions.py           -> caption_review.csv
  2. Open in Excel, edit SuggestedText, delete/skip OK rows
  3. python apply_caption_review.py    -> applies to VBA source files
"""
import re
import os
import csv
from collections import Counter

REPO   = "/home/user/otkupapp-pwa"
SRC    = os.path.join(REPO, "src-vba")
PORUKE = os.path.join(REPO, "localization", "poruke.csv")
OUT    = os.path.join(REPO, "localization", "caption_review.csv")

EXTS = {'.bas', '.cls', '.frm'}

# Supplementary word corrections not covered by poruke.csv
# (words that appear in form labels but not in user messages)
SUPPLEMENTARY = {
    'voca':       'voća',      # voća (ć)
    'voce':       'voće',
    'vocu':       'voću',
    'vocama':     'voćama',
    'drzava':     'država',    # država (ž)
    'drzave':     'države',
    'drzavu':     'državu',
    'drzavni':    'državni',
    'postanski':  'poštanski', # poštanski (š)
    'postanskog': 'poštanskog',
    'postanskom': 'poštanskom',
    'racun':      'račun',     # račun (č)
    'racuna':     'računa',
    'racune':     'račune',
    'racunu':     'računu',
    'otkupisce':  'otkupište', # otkupište (š)
    'opstina':    'opština',   # opština (š)
    'opstine':    'opštine',
    'opstinu':    'opštinu',
    'tekuci':     'tekući',    # tekući (ć)
    'tekuceg':    'tekućeg',
    'tekucem':    'tekućem',
    'citanje':    'čitanje',   # čitanje (č)
    'citanju':    'čitanju',
    'citanja':    'čitanja',
}

# Serbian diacritic strip table (for building ASCII keys)
STRIP = str.maketrans({
    'š': 's', 'Š': 'S',
    'ž': 'z', 'Ž': 'Z',
    'č': 'c', 'Č': 'C',
    'ć': 'c', 'Ć': 'C',
    'đ': 'd', 'Đ': 'D',
})

def strip_sr(s):
    return s.translate(STRIP)


def build_word_map():
    """
    Read all Tekst values from poruke.csv.
    For each word containing a Serbian diacritic, map its ASCII form -> unicode form.
    Ambiguous entries (same ASCII maps to 2+ different unicode words) are dropped.
    """
    wmap = {}       # ascii_lower -> unicode_lower
    conflict = set()

    with open(PORUKE, encoding='utf-8', newline='') as f:
        for row in csv.DictReader(f, delimiter=';'):
            text = row.get('Tekst', '')
            for word in re.findall(r"[A-Za-zÀ-ɏ']+", text):
                if strip_sr(word) == word:
                    continue  # no diacritics, skip
                key = strip_sr(word).lower()
                val = word.lower()
                if key in wmap and wmap[key] != val:
                    conflict.add(key)
                else:
                    wmap[key] = val

    for k in conflict:
        wmap.pop(k, None)

    # Merge supplementary (only for keys not already covered)
    for k, v in SUPPLEMENTARY.items():
        if k not in wmap and v != k:
            wmap[k] = v

    return wmap


def correct_text(text, wmap):
    """
    Apply word-map to text. Returns (corrected_text, status).
    Status: 'SUGGESTED' if something changed, 'OK' otherwise.
    """
    # Split into alternating (word, non-word) tokens
    tokens = re.findall(r'[A-Za-z]+|[^A-Za-z]+', text)
    out = []
    changed = False

    for tok in tokens:
        if not re.match(r'[A-Za-z]', tok):
            out.append(tok)
            continue

        key = tok.lower()
        if key in wmap:
            fix_lower = wmap[key]
            # Restore capitalization
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

    corrected = ''.join(out)
    return corrected, ('SUGGESTED' if changed else 'OK')


CAPTION_RE = re.compile(r'\.caption\s*=\s*"([^"]*)"', re.IGNORECASE)


def scan_files(wmap):
    results = []
    seen = set()  # (fajl, orig_text) -- deduplicate same string in same file

    for fname in sorted(os.listdir(SRC)):
        if os.path.splitext(fname)[1].lower() not in EXTS:
            continue
        path = os.path.join(SRC, fname)
        with open(path, encoding='latin-1', newline='') as f:
            lines = f.read().split('\n')

        for lineno, line in enumerate(lines, 1):
            # Skip comment lines
            stripped = line.lstrip()
            if stripped.startswith("'"):
                continue
            for m in CAPTION_RE.finditer(line):
                text = m.group(1)
                if not text.strip():
                    continue
                key = (fname, text)
                if key in seen:
                    continue
                seen.add(key)
                corrected, status = correct_text(text, wmap)
                results.append({
                    'Fajl':          fname,
                    'OrigLiteral':   f'"{text}"',
                    'SuggestedText': corrected,
                    'Status':        status,
                    'Line':          lineno,
                })

    # Sort: SUGGESTED first (most actionable), then OK
    results.sort(key=lambda r: (0 if r['Status'] == 'SUGGESTED' else 1,
                                r['Fajl'],
                                int(r['Line'])))
    return results


def main():
    print('Building word map from poruke.csv...')
    wmap = build_word_map()
    print(f'  {len(wmap)} unique word corrections')

    print('Scanning VBA source files...')
    rows = scan_files(wmap)

    with open(OUT, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f,
                                fieldnames=['Fajl', 'OrigLiteral', 'SuggestedText', 'Status', 'Line'],
                                delimiter=';',
                                lineterminator='\r\n')
        writer.writeheader()
        writer.writerows(rows)

    counts = Counter(r['Status'] for r in rows)
    print(f'\nOutput -> {OUT}')
    print(f"  SUGGESTED : {counts['SUGGESTED']:4d}  (word-map match, auto-correctable)")
    print(f"  OK        : {counts['OK']:4d}  (no change needed or unknown words)")
    print()
    print('Sledeci koraci:')
    print('  1. Otvori caption_review.csv u Excelu')
    print('  2. Pregledaj/ispravi SuggestedText kolonu')
    print('  3. Redove kojima je SuggestedText == OrigText (bez navodnika) mozes brisati')
    print('  4. python apply_caption_review.py')


if __name__ == '__main__':
    main()
