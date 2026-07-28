#!/usr/bin/env bash
# Materijal za prvi kontakt: .html je izvor istine, .pdf je deliverable.
#
#   tools/materijal.sh build   -> renderuje docs/Sales/AgriX_Materijal_za_prvi_kontakt.pdf
#   tools/materijal.sh check   -> proverava da tvrdnje o ceni i popustu prate vazece odluke
#
# Zavisnost: Chromium/Chrome preko tools/render-pdf.sh.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$ROOT/docs/Sales/AgriX_Materijal_za_prvi_kontakt.html"
OUT="$ROOT/docs/Sales/AgriX_Materijal_za_prvi_kontakt.pdf"

case "${1:-}" in
  build)
    "$ROOT/tools/render-pdf.sh" "$SRC" "$OUT"
    ;;

  check)
    python3 - "$ROOT" <<'PY'
import sys, os, re
root = sys.argv[1]
html = open(os.path.join(root, 'docs/Sales/AgriX_Materijal_za_prvi_kontakt.html'), encoding='utf-8').read()
log  = open(os.path.join(root, 'docs/Master Plan/09_QA_DECISION_LOG.md'), encoding='utf-8').read()
greske = []

# 1) Nijedna formulacija ne sme obecavati popust. Odluka 418 ih je ukinula,
#    a C3/111/112 obrisala. Ovo je greska koja je vec jednom prosla.
zabranjeno = [
    r'[Pp]opust postoji',
    r'popust\s+(?:za|samo za)\s+viš?e pravnih lica',
    r'[Mm]ogu[ćc] (?:je )?popust',
    r'individualni popust\w*\s+(?:se odobrava|postoji|je mogu)',
]
for uzorak in zabranjeno:
    for m in re.finditer(uzorak, html):
        kontekst = re.sub(r'\s+', ' ', html[max(0, m.start()-90):m.end()+90])
        greske.append(f'obecava popust (odluka 418 ih je ukinula): …{kontekst}…')

# 2) Svaki ID odluke naveden u dokumentu mora postojati u logu i ne sme biti obrisan.
obrisane = set()
for m in re.finditer(r'^(?:(\d{1,3})|([A-Z]+\d*))\. \*\*.*?`Deleted`', log, re.M):
    obrisane.add(m.group(1) or m.group(2))

pomenute = set()
for blok in re.findall(r'<span class="odl">\(([^)]+)\)</span>', html):
    for oid in re.split(r'[,\s]+', blok):
        if oid.strip():
            pomenute.add(oid.strip())

for oid in sorted(pomenute):
    if oid in obrisane:
        greske.append(f'poziva se na obrisanu odluku {oid}')
    elif not re.search(r'^' + re.escape(oid) + r'\. \*\*', log, re.M):
        greske.append(f'odluka {oid} ne postoji u decision logu')

# 3) Dokument mora ostati oznacen kao interni.
if 'ne šalje se klijentu' not in html:
    greske.append('nedostaje oznaka da je dokument interni')

if greske:
    print('GRESKE u materijalu za prvi kontakt:')
    for g in greske:
        print('  -', g)
    sys.exit(1)
print(f'OK: nema obecanja popusta; {len(pomenute)} navedenih odluka postoji i nijedna nije obrisana')
PY
    ;;

  *)
    sed -n '2,6p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
    exit 1
    ;;
esac
