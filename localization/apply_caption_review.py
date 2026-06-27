#!/usr/bin/env python3
"""
Apply reviewed caption_review.csv to VBA source files.

Reads caption_review.csv (after human review/edit), converts SuggestedText to
VBA ChrW() expressions, and replaces the original string literals in-place.

Only processes rows where Status != 'OK' AND SuggestedText differs from the
original text extracted from OrigLiteral.

Files are read/written as latin-1 (byte-preserving for ANSI VBA sources).
"""
import re
import os
import csv
from collections import defaultdict

REPO       = "/home/user/otkupapp-pwa"
REVIEW_CSV = os.path.join(REPO, "localization", "caption_review.csv")
SRC        = os.path.join(REPO, "src-vba")


# --------------------------------------------------------------------------
# VBA expression builder (100% ASCII output)
# --------------------------------------------------------------------------
def vba_text_expr(text):
    """Convert unicode string to a VBA expression using ChrW() for non-ASCII."""
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
# Load reviewed CSV
# --------------------------------------------------------------------------
def load_review():
    """Returns dict: fajl -> list of (orig_literal, vba_zamena)."""
    by_file = defaultdict(list)
    seen = set()

    with open(REVIEW_CSV, encoding='utf-8', newline='') as f:
        for row in csv.DictReader(f, delimiter=';'):
            status      = row.get('Status', '').strip()
            fajl        = row.get('Fajl', '').strip()
            orig_lit    = row.get('OrigLiteral', '').strip()
            suggested   = row.get('SuggestedText', '').strip()

            if status == 'OK':
                continue

            # Extract original text (strip outer quotes)
            if orig_lit.startswith('"') and orig_lit.endswith('"'):
                orig_text = orig_lit[1:-1]
            else:
                orig_text = orig_lit

            if orig_text == suggested:
                continue  # reviewer left it unchanged

            key = (fajl, orig_lit)
            if key in seen:
                continue
            seen.add(key)

            zamena = vba_text_expr(suggested)
            by_file[fajl].append((orig_lit, zamena))

    return by_file


# --------------------------------------------------------------------------
# Apply to files
# --------------------------------------------------------------------------
def apply_file(fajl, pairs):
    path = os.path.join(SRC, fajl)
    if not os.path.exists(path):
        print(f'  SKIP (not found): {fajl}')
        return 0

    with open(path, encoding='latin-1', newline='') as f:
        content = f.read()

    original = content
    n = 0

    for orig, zamena in pairs:
        count = content.count(orig)
        if count == 0:
            print(f'  WARN not found: {orig[:60]}')
            continue
        content = content.replace(orig, zamena)
        n += count
        print(f'    {count}x: {orig[:50]} -> {zamena[:60]}')

    if content == original:
        return 0

    with open(path, 'w', encoding='latin-1', newline='') as f:
        f.write(content)

    return n


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    print(f'Reading {REVIEW_CSV}...')
    by_file = load_review()

    if not by_file:
        print('  Nema promena za primeniti (svi redovi OK ili SuggestedText = original).')
        return

    total = 0
    for fajl in sorted(by_file):
        pairs = by_file[fajl]
        print(f'\n{fajl} ({len(pairs)} zamena):')
        n = apply_file(fajl, pairs)
        total += n

    print(f'\nUkupno zamenjeno: {total}')
    print()
    print('Sledeci koraci:')
    print('  git diff src-vba/  -- proveri izmene')
    print('  file src-vba/*.frm src-vba/*.bas -- mora ostati ASCII (ChrW je ASCII)')
    print('  git add + commit + push')
    print('  ImportAllVBA -> Debug -> Compile -> test u Excelu')


if __name__ == '__main__':
    main()
