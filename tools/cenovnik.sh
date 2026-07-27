#!/usr/bin/env bash
# Cenovnik 2027: .html je izvor istine, .pdf je deliverable.
#
#   tools/cenovnik.sh build   -> renderuje docs/Sales/AgriX_Cenovnik_2027.pdf iz .html
#   tools/cenovnik.sh check   -> poredi iznose u .html sa ostala tri mesta sa cenama
#
# Zavisnost: Chromium/Chrome (headless print-to-pdf) i Python + openpyxl za `check`.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$ROOT/docs/Sales/AgriX_Cenovnik_2027.html"
OUT="$ROOT/docs/Sales/AgriX_Cenovnik_2027.pdf"

nadji_chrome() {
  for c in "$CHROME_BIN" /opt/pw-browsers/chromium-*/chrome-linux/chrome \
           "$(command -v chromium || true)" "$(command -v chromium-browser || true)" \
           "$(command -v google-chrome || true)" \
           "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"; do
    [ -n "${c:-}" ] && [ -x "$c" ] && { echo "$c"; return 0; }
  done
  return 1
}

case "${1:-}" in
  build)
    CHROME_BIN="${CHROME_BIN:-}"
    chrome="$(nadji_chrome)" || {
      echo "GRESKA: Chromium/Chrome nije pronadjen." >&2
      echo "  Postavi CHROME_BIN=/putanja/do/chrome ili instaliraj Chromium." >&2
      exit 1
    }
    tmp="$(mktemp -d)"; trap 'rm -rf "$tmp"' EXIT
    "$chrome" --headless --disable-gpu --no-sandbox \
      --no-pdf-header-footer --print-to-pdf-no-header \
      --print-to-pdf="$OUT" "file://$SRC" >"$tmp/log" 2>&1 || { cat "$tmp/log" >&2; exit 1; }
    echo "OK: $OUT"
    ;;

  check)
    python3 - "$ROOT" <<'PY'
import sys, re, os
root = sys.argv[1]
html = open(os.path.join(root, 'docs/Sales/AgriX_Cenovnik_2027.html'), encoding='utf-8').read()

def iz_htmla(labela):
    m = re.search(labela, html, re.S)
    return m.group(1).replace('.', '') if m else None

ocek = {
    'Desktop':        iz_htmla(r'AgriX Desktop</h3>\s*<div class="price">([\d.]+)\s*€'),
    'Desktop all-in': iz_htmla(r'SEF, Banka i Hladnjača</span>\s*<b>([\d.]+)\s*€'),
    'Mobile':         iz_htmla(r'AgriX Mobile</h3>\s*<div class="price">([\d.]+)\s*€'),
    'Mobile all-in':  iz_htmla(r'SEF, Banka, Hladnjača i Dispatch</span>\s*<b>([\d.]+)\s*€'),
}
if None in ocek.values():
    print('GRESKA: ne mogu da procitam sve cene iz .html izvora:', ocek); sys.exit(1)
ocek = {k: int(v) for k, v in ocek.items()}

greske = []
try:
    import openpyxl
except ImportError:
    print('UPOZORENJE: openpyxl nije instaliran, preskacem provere .xlsx fajlova')
    openpyxl = None

if openpyxl:
    cz = openpyxl.load_workbook(os.path.join(root, 'docs/Sales/AgriX_Sablon_ponude.xlsx'))['Cenovnik']
    sab = {cz[f'A{r}'].value: cz[f'B{r}'].value for r in range(5, 9)}
    for k, v in ocek.items():
        if sab.get('AgriX ' + k) != v:
            greske.append(f'sablon ponude: {k} = {sab.get("AgriX " + k)}, ocekivano {v}')

    pr = openpyxl.load_workbook(os.path.join(root, 'docs/Finance/AgriX_Finansijski_model.xlsx'))['Pretpostavke']
    fin = {pr[f'A{r}'].value: pr[f'B{r}'].value for r in range(6, 10)}
    for naziv, k in [('Desktop Core', 'Desktop'), ('Desktop all-in', 'Desktop all-in'),
                     ('Mobile', 'Mobile'), ('Mobile all-in', 'Mobile all-in')]:
        if fin.get(naziv) != ocek[k]:
            greske.append(f'finansijski model: {naziv} = {fin.get(naziv)}, ocekivano {ocek[k]}')

ug = open(os.path.join(root, 'docs/Legal/AgriX_Ugovor_o_licenciranju.md'), encoding='utf-8').read()
nadjeno = [int(x.replace('.', '')) for x in re.findall(r'AgriX (?:Desktop|Mobile)(?: all-in)?\s*\|\s*([\d.]+)', ug)]
if nadjeno != [ocek['Desktop'], ocek['Desktop all-in'], ocek['Mobile'], ocek['Mobile all-in']]:
    greske.append(f'ugovor Prilog 1: {nadjeno}, ocekivano {list(ocek.values())}')

if greske:
    print('RAZLIKA u cenama:'); [print('  -', g) for g in greske]; sys.exit(1)
print('OK: cene se poklapaju na sva cetiri mesta ->', ocek)
PY
    ;;

  *)
    sed -n '2,7p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
    exit 1
    ;;
esac
