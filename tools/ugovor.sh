#!/usr/bin/env bash
# Ugovor o licenciranju: .md je izvor istine, .docx je deliverable.
#
#   tools/ugovor.sh build   -> generise docs/Legal/build/AgriX_Ugovor_o_licenciranju.docx iz .md
#   tools/ugovor.sh check   -> poredi tekst kuriranog .docx-a sa .md i prijavi razlike
#
# Zavisnost: pandoc (https://pandoc.org). Provera: pandoc --version
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
MD="$ROOT/docs/Legal/AgriX_Ugovor_o_licenciranju.md"
DOCX="$ROOT/docs/Legal/AgriX_Ugovor_o_licenciranju.docx"
OUTDIR="$ROOT/docs/Legal/build"
OUT="$OUTDIR/AgriX_Ugovor_o_licenciranju.docx"

# Mapa clanova na odluke je interni aparat i ne ide u ugovor.
MAP_HEADING='## Prilog uz izvor: mapa'

need_pandoc() {
  command -v pandoc >/dev/null 2>&1 || {
    echo "GRESKA: pandoc nije instaliran." >&2
    echo "  Ubuntu/Debian: sudo apt-get install pandoc" >&2
    echo "  macOS:         brew install pandoc" >&2
    echo "  Windows:       winget install --id JohnMacFarlane.Pandoc" >&2
    exit 1
  }
}

# Telo ugovora: bez HTML zaglavlja, bez metapodataka i bez mape odluka.
# Pocinje posle prvog samostalnog '---', zavrsava pred mapom odluka.
contract_body() {
  awk -v stop="$MAP_HEADING" '
    index($0, stop) == 1 { exit }
    !started { if ($0 == "---") started = 1; next }
    { print }
  ' "$MD"
}

# Zajednicka normalizacija za obe strane poredjenja.
# Uklanja markdown ukrase, granice tabela, prazne linije i SVE verzalne linije
# (naslovi clanova i priloga) - oni se legitimno razlikuju u .md i .docx formatu.
normalize() {
  sed -e 's/\\//g' \
      -e 's/\*\*//g' -e 's/\*//g' \
      -e 's/^#\+ *//' -e 's/^> *//' \
      -e 's/^|//' -e 's/|$//' \
      -e 's/_\{3,\}/__________/g' \
      -e 's/^\(Član [0-9]\+\) *—.*$/\1/' \
      -e 's/^\(PRILOG [0-9]\+\) *—.*$/\1/' \
    | tr -s ' |' ' ' \
    | sed -e 's/^ *//' -e 's/ *$//' \
    | grep -v '^-*$' \
    | grep -v '^$' \
    | grep -vP '^[^\p{Ll}]{4,}$'
}

case "${1:-}" in
  build)
    need_pandoc
    mkdir -p "$OUTDIR"
    ref=()
    [ -f "$DOCX" ] && ref=(--reference-doc="$DOCX")
    contract_body | pandoc -f gfm -t docx "${ref[@]}" -o "$OUT"
    echo "OK: $OUT"
    echo "Napomena: kurirani $DOCX NIJE prepisan. Uporedi pa zameni rucno ako je build ispravan."
    ;;

  check)
    need_pandoc
    [ -f "$DOCX" ] || { echo "GRESKA: nema $DOCX"; exit 1; }
    tmp="$(mktemp -d)"
    trap 'rm -rf "$tmp"' EXIT
    pandoc -f docx -t gfm --wrap=none "$DOCX" | normalize > "$tmp/docx.txt"
    contract_body | normalize > "$tmp/md.txt"
    if diff -u "$tmp/md.txt" "$tmp/docx.txt" > "$tmp/diff.txt"; then
      echo "OK: tekst .docx-a i .md se poklapaju."
    else
      echo "RAZLIKA u tekstu ugovora, .md (-) protiv .docx (+):"
      echo
      cat "$tmp/diff.txt"
      echo
      echo "Ako je .docx u pravu (npr. pravnik je menjao) - prenesi izmene u .md."
      echo "Ako je .md u pravu - pokreni 'tools/ugovor.sh build' i zameni .docx."
      exit 1
    fi
    ;;

  *)
    sed -n '2,8p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
    exit 1
    ;;
esac
