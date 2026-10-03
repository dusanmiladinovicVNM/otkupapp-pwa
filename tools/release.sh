#!/usr/bin/env bash
# tools/release.sh — release rutina (git deo): pull -> bump APP_VERSION -> KAPIJA -> commit -> tag -> push -> stamp.
# Pokreni iz repo klona NA BUILD MASINI (gde Excel modVbaTools.FOLDER pokazuje na isti src-vba):
#   bash tools/release.sh 2.2.2
# Posle ovoga ostaju Excel koraci + restore — skripta ih ispiše. Vidi docs/RELEASE_PROCEDURE.md.
#
# ZASTO KAPIJA STOJI PRED TAGOM
# -----------------------------
# Do 03.10.2026 je redosled bio: bump -> commit -> push main -> tag -> push tag,
# a Excel (ImportAllVBA + Debug > Compile) je bio uputstvo ISPOD toga. Tag je
# dakle nastajao i bio objavljen PRE nego sto je bilo sta izmereno nad tim
# izvorom. To nije teorija: `vba-v2.40.0` release notes izricito kaze da
# `run_vba.py` nije bio pokrenut, a tag je ipak postojao i bio push-ovan.
#
# Sada tag ne moze da nastane dok marker ne dokaze da je BAS OVAJ izvor
# (ukljucujuci bump APP_VERSION-a) prosao suite i rucni compile.
#
# SKRIPTA NE POKRECE EXCEL, i to je namerno. Pun prolaz traje 20-60 minuta, a
# compile je rucna kapija operatera. Posao skripte je da PROVERI MARKER
# (`tests/last_green.json`) protiv otiska izvora koji se isporucuje — to je
# jeftino, pa kapija moze da stoji u skripti i da se ne preskace.
#
# NA PADU KAPIJE BUMP OSTAJE U RADNOM DRVETU, nekomitovan. Operater tada testira
# tacno onaj izvor koji se isporucuje; da skripta vracala bump, testiralo bi se
# nesto drugo. Ponovno pokretanje je idempotentno.
#
# Statika se NE ponavlja cela: `release.sh` krece od `main`-a koji je CI vec
# izmerio, a bump dira jednu string konstantu. Vrti se `vba_check.py`, jer bas
# on pokriva ono sto bump moze da pokvari (ASCII, kraj reda, deklaracije) i
# jeftin je. Pun spisak statickih kapija drzi `.github/workflows/static.yml`.
set -euo pipefail

waive=""
reason=""
args=()
while [ $# -gt 0 ]; do
  case "$1" in
    --waive)  waive="${2:-}"; shift 2 ;;
    --reason) reason="${2:-}"; shift 2 ;;
    *)        args+=("$1"); shift ;;
  esac
done
set -- ${args[@]+"${args[@]}"}

ver="${1:-}"
if [ -z "$ver" ]; then
  echo "Upotreba: bash tools/release.sh <verzija>   (npr. 2.2.2)" >&2
  echo "          bash tools/release.sh 2.2.2 --waive ponasanje --reason \"<zasto>\"" >&2
  exit 1
fi
ver="${ver#v}"                         # dozvoli i "v2.2.2"
tag="vba-v${ver}"
root="$(git rev-parse --show-toplevel)"
cfg="$root/src-vba/modConfig.bas"
mbi="src-vba/modBuildInfo.bas"

# WAIVER TRAZI RAZLOG, i to je jedini nacin da kapija bude preskocena. Bez
# razloga se odbija: "preskoci kapiju" bez zapisanog zasto je tacno ono stanje
# zbog kog kapija i postoji. Razlog ulazi u ANOTIRAN TAG, pa ostaje trajno
# vidljiv u `git show <tag>` — waived release je zauvek obelezen kao waived.
if [ -n "$waive" ] && [ -z "$reason" ]; then
  echo "--waive $waive bez --reason: odbijeno." >&2
  echo "Waiver bez zapisanog razloga je tiho preskakanje kapije." >&2
  exit 1
fi
if [ -n "$waive" ] && [ "$waive" != "ponasanje" ] && [ "$waive" != "compile" ]; then
  echo "--waive prima 'ponasanje' ili 'compile', dobijeno: $waive" >&2
  exit 1
fi

PY="$(command -v python || command -v python3 || true)"
if [ -z "$PY" ]; then
  echo "Nema python-a u PATH-u — kapija se ne moze izvrsiti, pa se release prekida." >&2
  exit 1
fi
if [ ! -f "$root/tools/vba_gate.py" ]; then
  echo "Nema tools/vba_gate.py — bez njega nema cime da se dokaze izvor." >&2
  exit 1
fi

echo "== 1/6 main + pull =="
git -C "$root" checkout main
git -C "$root" pull --ff-only origin main

# Radni direktorijum mora biti cist. Dva izuzetka, oba sa imenom:
#   - modBuildInfo.bas — njega prepisuje stamp;
#   - modConfig.bas, ali SAMO ako je razlika bas bump APP_VERSION-a koji je ova
#     skripta ostavila kad je kapija pala. Bez tog izuzetka drugi pokusaj ne bi
#     mogao ni da pocne, pa bi operater "ociscio" bas izvor koji je testirao.
prljavo="$(git -C "$root" status --porcelain -- . ":(exclude)$mbi" || true)"
if [ -n "$prljavo" ]; then
  samo_cfg="$(printf '%s\n' "$prljavo" | awk '{print $2}' | sort -u)"
  van_bumpa="$(git -C "$root" diff -U0 -- "$cfg" \
               | grep -E '^[+-][^+-]' \
               | grep -vE 'APP_VERSION As String = ' | wc -l || true)"
  if [ "$samo_cfg" != "src-vba/modConfig.bas" ] || [ "$van_bumpa" -ne 0 ]; then
    echo "Radni direktorijum nije cist (osim $mbi i bump-a APP_VERSION-a)." >&2
    echo "Commit/oclisti pa ponovi." >&2
    git -C "$root" status --short >&2
    exit 1
  fi
  echo "   (zatecen nekomitovan bump APP_VERSION-a — nastavljam nad njim)"
fi

echo "== 2/6 bump APP_VERSION -> $ver =="
if ! grep -q 'APP_VERSION As String' "$cfg"; then
  echo "Ne nalazim APP_VERSION u $cfg" >&2; exit 1
fi
tmp="$(mktemp)"
sed -E 's/APP_VERSION As String = "[^"]*"/APP_VERSION As String = "'"$ver"'"/' "$cfg" > "$tmp" && mv "$tmp" "$cfg"

echo "== 3/6 KAPIJA nad izvorom koji se isporucuje =="
otisak="$("$PY" "$root/tools/vba_gate.py" --hash)"
echo "   izvor: ${otisak:0:16}"

kapija_pala=0
nalaz=""

echo "   -- vba_check (ASCII, kraj reda, deklaracije, duplikati)"
if ! "$PY" "$root/tools/vba_check.py" >/dev/null; then
  "$PY" "$root/tools/vba_check.py" >&2 || true
  nalaz="${nalaz}statika: vba_check ima nalaze"$'\n'
  kapija_pala=1
fi

trazi_compile=(--require-compile)
if [ "$waive" = "compile" ]; then trazi_compile=(); fi

echo "   -- marker zelenog: da li je BAS ovaj izvor prosao suite"
if [ "$waive" = "ponasanje" ]; then
  echo "      PRESKOCENO (--waive ponasanje)"
elif ! "$PY" "$root/tools/vba_gate.py" --require-green "${trazi_compile[@]+${trazi_compile[@]}}"; then
  "$PY" "$root/tools/vba_gate.py" --status >&2 || true
  nalaz="${nalaz}ponasanje/compile: izvor nije dokazan"$'\n'
  kapija_pala=1
elif [ "$waive" = "compile" ]; then
  echo "      suite dokazane; COMPILE PRESKOCEN (--waive compile)"
fi

if [ "$kapija_pala" -ne 0 ]; then
  echo >&2
  echo "------------------------------------------------------------" >&2
  echo "KAPIJA PALA — nema commita, nema taga, nista nije push-ovano." >&2
  printf '%s' "$nalaz" >&2
  echo >&2
  echo "Bump APP_VERSION-a OSTAJE u radnom drvetu, nekomitovan:" >&2
  echo "  to je tacno onaj izvor koji se isporucuje, pa ga testiraj takvog." >&2
  echo >&2
  echo "Na build masini (Windows + Excel), u ovom redu:" >&2
  echo "  python tools/run_vba.py                    # pun prolaz, pise marker" >&2
  echo "  # Alt+F11 -> Debug -> Compile VBAProject   (mora bez greske)" >&2
  echo "  python tools/vba_gate.py --mark-compile    # potvrdi compile nad OVIM izvorom" >&2
  echo "  python tools/vba_gate.py --status          # pogledaj sta jos fali" >&2
  echo "  bash tools/release.sh $ver                 # pa ponovo ovde" >&2
  echo >&2
  echo "Ako je kapija objektivno neizvodljiva, mora da ostane zapisano:" >&2
  echo "  bash tools/release.sh $ver --waive ponasanje --reason \"<zasto>\"" >&2
  echo "------------------------------------------------------------" >&2
  exit 2
fi
echo "   KAPIJA PROSLA"

echo "== 4/6 commit =="
if [ -n "$(git -C "$root" status --porcelain -- "$cfg")" ]; then
  git -C "$root" commit -m "release: vba v${ver}" -- "$cfg"
else
  echo "   APP_VERSION je vec $ver — nema commita."
fi

echo "== 5/6 anotiran tag $tag + push =="
# ANOTIRAN, ne lagan: verdikt pod kojim je tag nastao ide U TAG. Recenica u PR-u
# ili u release notes-u se gubi i ne moze se proveriti kasnije; `git show <tag>`
# moze. Ako je release waived, to je ovde — zauvek.
if git -C "$root" rev-parse -q --verify "refs/tags/$tag" >/dev/null; then
  echo "   Tag $tag vec postoji lokalno — preskacem kreiranje."
else
  {
    echo "AgriX VBA v${ver}"
    echo
    echo "Kapija pred tagom (tools/release.sh):"
    echo "  izvor      ${otisak}"
    if [ -n "$waive" ]; then
      echo "  WAIVED     $waive"
      echo "  razlog     $reason"
    fi
    echo
    "$PY" "$root/tools/vba_gate.py" --status 2>/dev/null \
      | sed 's/^/  /' || echo "  (status markera nedostupan)"
  } | git -C "$root" tag -a "$tag" -F -
fi
git -C "$root" push origin main
git -C "$root" push origin "$tag"

echo "== 6/6 stamp build otisak =="
bash "$root/tools/stamp-build.sh"

echo
echo "------------------------------------------------------------"
if [ -n "$waive" ]; then
  echo "!! Ovaj release je WAIVED: $waive"
  echo "!! Razlog je u anotiranom tagu: git show $tag"
  echo
fi
echo "Sad u Excelu (master .xlsm) — RUCNO:"
echo "  1) Alt+F8 -> ImportAllVBA"
echo "  2) Debug -> Compile VBAProject   (mora bez greske)"
echo "  3) Alt+F8 -> AssertBlankBuild     (mora 'BLANKO OK')"
echo "  3b) Alt+F8 -> PublishReleaseToDrive  (objavi kod + version.json u AgriX_Release; treba REL_FOLDER_ID)"
echo "  4) (opciono) VBE: Tools -> Digital Signature   (ako potpisuješ)"
echo "  5) Snimi / Save As builds (Ctrl+S)"
echo
echo "Zatim vrati placeholder i isporuci:"
echo "  git checkout -- $mbi"
echo "  -> pošalji .xlsm klijentima"
echo
echo "Ne zaboravi: dopuni docs/RELEASE_NOTES.md (par rečenica o ovom izdanju)."
echo
echo "Provera: GAS rebuildMonitoringFleet() / Fleet tab (ili auto-trigger na sat)."
echo "------------------------------------------------------------"
