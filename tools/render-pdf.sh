#!/usr/bin/env bash
# Zajednicki deterministicki HTML -> PDF renderer za prodajne dokumente.
#
#   tools/render-pdf.sh <ulaz.html> <izlaz.pdf>
#
# Chromium u PDF upisuje /CreationDate i /ModDate, pa dva uzastopna build-a
# istog izvora daju razlicite bajtove i lazan git diff. Zato se ta dva polja
# posle rendera normalizuju na fiksnu vrednost. Zamena je iste duzine, pa
# xref tabela ostaje ispravna.
#
# Zavisnost: Chromium/Chrome. CHROME_BIN nadjacava automatsko pronalazenje.
set -euo pipefail

SRC="${1:?upotreba: render-pdf.sh <ulaz.html> <izlaz.pdf>}"
OUT="${2:?upotreba: render-pdf.sh <ulaz.html> <izlaz.pdf>}"

# Fiksan datum u PDF metapodacima. Nije datum dokumenta — dokument nosi
# svoju verziju u tekstu. Ovo postoji samo da build bude ponovljiv.
PDF_DATE="${PDF_DATE:-D:20270101000000+00'00'}"

nadji_chrome() {
  for c in "${CHROME_BIN:-}" /opt/pw-browsers/chromium-*/chrome-linux/chrome \
           "$(command -v chromium || true)" "$(command -v chromium-browser || true)" \
           "$(command -v google-chrome || true)" \
           "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"; do
    [ -n "${c:-}" ] && [ -x "$c" ] && { echo "$c"; return 0; }
  done
  return 1
}

chrome="$(nadji_chrome)" || {
  echo "GRESKA: Chromium/Chrome nije pronadjen." >&2
  echo "  Postavi CHROME_BIN=/putanja/do/chrome ili instaliraj Chromium." >&2
  exit 1
}

[ -f "$SRC" ] || { echo "GRESKA: nema izvornog fajla $SRC" >&2; exit 1; }

tmp="$(mktemp -d)"; trap 'rm -rf "$tmp"' EXIT
"$chrome" --headless --disable-gpu --no-sandbox \
  --no-pdf-header-footer --print-to-pdf-no-header \
  --print-to-pdf="$tmp/out.pdf" "file://$(cd "$(dirname "$SRC")" && pwd)/$(basename "$SRC")" \
  >"$tmp/log" 2>&1 || { cat "$tmp/log" >&2; exit 1; }

# Normalizacija vremenskih peciata -> ponovljiv build
python3 - "$tmp/out.pdf" "$OUT" "$PDF_DATE" <<'PY'
import re, sys
ulaz, izlaz, datum = sys.argv[1], sys.argv[2], sys.argv[3].encode()
d = open(ulaz, 'rb').read()
for polje in (b'/CreationDate', b'/ModDate'):
    def zameni(m, polje=polje):
        staro = m.group(1)
        novo = datum
        if len(novo) != len(staro):
            # duzina mora ostati ista da xref offseti ostanu tacni
            novo = novo[:len(staro)].ljust(len(staro), b' ')
        return polje + b' (' + novo + b')'
    d = re.sub(re.escape(polje) + rb'\s*\(([^)]*)\)', zameni, d)
open(izlaz, 'wb').write(d)
PY

echo "OK: $OUT"
