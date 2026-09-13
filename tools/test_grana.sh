#!/usr/bin/env bash
# Sastavi lokalnu TEST granu od izabranih PR-ova i proveri da li je bezbedno
# uvesti je u zatecenu svesku.
#
#   bash tools/test_grana.sh 308 311
#   bash tools/test_grana.sh --sveska "C:/Users/Dusan/Desktop/AGRIX DEV/AgriX.xlsm" 308 311
#   bash tools/test_grana.sh --sta-je-na-disku
#
# ZASTO POSTOJI
# ImportAllVBA cita src-vba/ ZATECENE grane i o tome ne pita nista. Uvoz koda
# cija se sema ne poklapa sa sveskom salje POZICIONE upise u pogresne kolone --
# a to je greska u podacima, ne pad upisa. Ovaj alat pre uvoza kaze sta je na
# disku i da li je bezbedno.
#
# Ne dira nijednu PR granu: pravi zasebnu `test/sveska`, koja se sme brisati.
#
# STA NE RADI
# Ne pokrece Excel i ne uvozi nista. Uvoz i `Debug > Compile VBAProject` ostaju
# rucni koraci -- alat samo kaze smes/ne smes.

set -u

# Root se izvodi iz gita, ne hardkoduje: skripta zivi u tools/ i mora da radi iz
# bilo kog radnog direktorijuma i na bilo kom klonu.
REPO="$(git rev-parse --show-toplevel 2>/dev/null)"
if [ -z "$REPO" ]; then
  echo "Nisam u git repou." >&2
  exit 2
fi
cd "$REPO" || exit 2

# gh nije na PATH-u na svakoj masini (v. docs). Redosled: $GH iz okruzenja ->
# PATH -> poznata Windows putanja.
GH="${GH:-}"
if [ -z "$GH" ]; then
  if command -v gh >/dev/null 2>&1; then
    GH="$(command -v gh)"
  else
    GH="/c/Program Files/GitHub CLI/gh.exe"
  fi
fi

GRANA="test/sveska"
# Manifest ide u .git/ -- NIKAD u radno stablo. Untracked fajl u repou bi oborio
# sopstvenu proveru ciste kopije nekoliko redova nize.
SADRZAJ="$(git rev-parse --git-dir)/test_grana_sadrzaj.txt"
SVESKA=""
PROVERA_SAMO=0
PROVI=()

pomoc() {
  sed -n '2,21p' "$0" | sed 's/^# \{0,1\}//'
}

while [ $# -gt 0 ]; do
  case "$1" in
    --sveska)          SVESKA="${2:-}"; shift 2 ;;
    --grana)           GRANA="${2:-}"; shift 2 ;;
    --sta-je-na-disku) PROVERA_SAMO=1; shift ;;
    -h|--help)         pomoc; exit 0 ;;
    -*)                echo "Nepoznat prekidac: $1" >&2; exit 2 ;;
    *)                 PROVI+=("$1"); shift ;;
  esac
done

# --- sta je trenutno na disku -------------------------------------------------
if [ "$PROVERA_SAMO" = "1" ]; then
  echo "grana:  $(git rev-parse --abbrev-ref HEAD)"
  echo "commit: $(git log --oneline -1)"
  if [ -n "$(git status --porcelain)" ]; then
    echo "PAZNJA: radna kopija NIJE cista -- na disku je nesto sto nije u commitu."
    echo "        Ako je medju izmenama src-vba, ImportAllVBA ce uvesti BAS TO:"
    git status --short | head -10
  else
    echo "radna kopija: cista"
  fi
  [ -f "$SADRZAJ" ] && { echo "sadrzaj test grane:"; cat "$SADRZAJ"; }
  exit 0
fi

if [ ${#PROVI[@]} -eq 0 ]; then
  echo "Reci koje PR-ove: bash tools/test_grana.sh 308 311" >&2
  exit 2
fi

if [ ! -x "$GH" ] && ! command -v "$GH" >/dev/null 2>&1; then
  echo "gh CLI nije nadjen: $GH" >&2
  echo "Postavi putanju: GH=\"/c/Program Files/GitHub CLI/gh.exe\" bash tools/test_grana.sh ..." >&2
  exit 2
fi

# --- preduslovi ---------------------------------------------------------------
if [ -n "$(git status --porcelain)" ]; then
  echo "Radna kopija nije cista -- commituj ili odbaci izmene pre sastavljanja:" >&2
  git status --short >&2
  exit 2
fi

echo "== fetch"
git fetch -q origin || exit 2

# --- PR broj -> head grana ----------------------------------------------------
GRANE=()
for pr in "${PROVI[@]}"; do
  b="$("$GH" pr view "$pr" --json headRefName -q .headRefName 2>/dev/null)"
  if [ -z "$b" ]; then
    echo "PR #$pr: ne mogu da nadjem granu" >&2
    exit 2
  fi
  t="$("$GH" pr view "$pr" --json title -q .title 2>/dev/null)"
  echo "  #$pr  $b"
  echo "        $t"
  GRANE+=("$b")
done

# --- sastavljanje -------------------------------------------------------------
echo
echo "== sastavljam $GRANA od origin/main"
git checkout -q -B "$GRANA" origin/main || exit 2

: > "$SADRZAJ"
for i in "${!GRANE[@]}"; do
  b="${GRANE[$i]}"
  echo "== merge $b"
  if ! git merge -q --no-edit "origin/$b"; then
    sukobi="$(git diff --name-only --diff-filter=U)"
    # Jedini ocekivan sukob je GENERISANA mapa vlasnistva -- ona se ne resava
    # rucno nego regeneracijom iz koda.
    if [ "$sukobi" = "docs/DOMEN/WHO_WRITES.md" ]; then
      echo "   sukob samo u generisanom WHO_WRITES.md -- regenerisem"
      git checkout --theirs docs/DOMEN/WHO_WRITES.md 2>/dev/null
      git add docs/DOMEN/WHO_WRITES.md
      git -c core.editor=true merge --continue >/dev/null 2>&1 || git commit -q --no-edit
    else
      echo "   SUKOB koji se ne resava sam:" >&2
      echo "$sukobi" >&2
      echo "   git merge --abort pa javi -- ova kombinacija PR-ova trazi ruke." >&2
      exit 2
    fi
  fi
  echo "#${PROVI[$i]}  $b" >> "$SADRZAJ"
done

python tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md >/dev/null
if [ -n "$(git status --porcelain docs/DOMEN/WHO_WRITES.md)" ]; then
  git add docs/DOMEN/WHO_WRITES.md
  git commit -q -m "test: regenerisan WHO_WRITES za test granu"
fi

# --- kapije -------------------------------------------------------------------
echo
echo "== vba_check"
python tools/vba_check.py | tail -2
vc=${PIPESTATUS[0]}

echo
echo "== modSchema u koraku sa kanonom"
python tools/gen_schema_module.py --check | tail -1
gs=${PIPESTATUS[0]}

sd=0
if [ -n "$SVESKA" ]; then
  lock="$(dirname "$SVESKA")/~\$$(basename "$SVESKA")"
  if [ -f "$lock" ]; then
    echo
    echo "== schema_diff: PRESKOCEN -- sveska je otvorena u Excelu"
    echo "   zatvori je pa pokreni: python tools/schema_diff.py \"$SVESKA\""
    sd=3
  else
    echo
    echo "== schema_diff nad sveskom"
    python tools/schema_diff.py "$SVESKA" | tail -4
    sd=${PIPESTATUS[0]}
  fi
fi

# --- ishod --------------------------------------------------------------------
echo
echo "============================================================"
echo "na disku je sada: $(git rev-parse --abbrev-ref HEAD)"
sed 's/^/   /' "$SADRZAJ"
echo "============================================================"

if [ "$vc" != "0" ]; then
  echo "vba_check PAO -- NE uvozi." >&2
  exit 2
fi
if [ "$gs" != "0" ]; then
  echo "modSchema.bas nije u koraku sa schema/schema.json -- NE uvozi." >&2
  echo "Regenerisi: python tools/gen_schema_module.py" >&2
  exit 2
fi
if [ "$sd" = "2" ]; then
  echo "schema_diff prijavljuje odstupanje -- NE uvozi dok se ne resi." >&2
  exit 2
fi

# Fixture je gitignored, pa ga prelazak grane NE menja -- a potpis pokriva i
# ugovor o formatu iz kanona (PR #318), koji se sa granom menja. Zato posle
# sastavljanja test grane run_vba po pravilu trazi regeneraciju; to nije kvar.
echo
echo "NB: ako posle ovoga pustas run_vba, fixture je verovatno ustajao:"
echo "  python tools/make_fixture.py --donor tests/fixtures/otkup_test.xlsm \\"
echo "      --out tests/fixtures/otkup_test_new.xlsm --force"
echo "  mv -f tests/fixtures/otkup_test_new.xlsm tests/fixtures/otkup_test.xlsm"
echo "  mv -f tests/fixtures/otkup_test_new.sig  tests/fixtures/otkup_test.sig"

if [ "$sd" = "3" ]; then
  echo
  echo "Sema NIJE proverena (sveska otvorena). Proveri pa tek onda uvezi."
  exit 0
fi

echo
echo "Bezbedno: otvori svesku -> Alt+F8 -> ImportAllVBA -> Alt+F11 -> Debug -> Compile."
