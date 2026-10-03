#!/usr/bin/env bash
# tools/dokaz_release.sh — dvosmerni dokaz za kapiju u tools/release.sh.
#
#   bash tools/dokaz_release.sh
#
# Izlazni kod: 0 = kapija grize, 1 = ima nalaza.
#
# ZASTO POSTOJI
# -------------
# `release.sh` je od 03.10.2026 KAPIJA: tag ne moze da nastane dok marker ne
# dokaze da je bas taj src-vba prosao suite i rucni compile. Kapija koja nikad
# nije pokazana crvena ne dokazuje da ista zaustavlja — a ovu je nemoguce
# izmeriti usput, jer radi `checkout main`, `commit`, `tag` i `push`.
#
# Zato se gradi IZOLOVAN repo sa lokalnim bare "origin"-om: nema mreze, nema
# pravog push-a, i nijedna komanda ne dira repo iz koga je pozvana.
#
# Repo se gradi OD NULE (kopiranjem), ne kloniranjem: klon bi uzeo `main` —
# dakle skript koji se NE meri — a web/CI klon je uz to shallow, pa bi push u
# bare origin bio odbijen ("shallow update not allowed").
#
# Marker se za pozitivan slucaj PISE RUCNO, preko `vba_gate` API-ja, nad pravim
# otiscima tog repoa. Excel se ne pokrece; meri se kapija, ne suite.
#
# NE MERI: da li suite zaista prolaze (to je `run_vba.py`), ni da li compile
# prolazi (to je rucna kapija operatera). Meri da tag NE MOZE da nastane bez
# dokaza, i da moze sa njim.
set -uo pipefail

IZVOR="$(git rev-parse --show-toplevel)"
BAZA="$(mktemp -d)"
BARE="$BAZA/origin.git"
KLON="$BAZA/klon"
pali=0

ok()  { printf "DOKAZANO    %s\n" "$1"; }
pao() { printf "NE GRIZE    %s\n" "$1"; pali=$((pali+1)); }

# Repo se gradi OD NULE, ne klonira: izvorni repo je shallow klon, pa push u
# bare origin odbija ("shallow update not allowed"). Uz to, klon bi uzeo `main`
# -- dakle STARI skript -- a meri se onaj sa ove grane.
pripremi() {
  rm -rf "$BARE" "$KLON"
  git init -q --bare "$BARE"
  mkdir -p "$KLON"
  for d in src-vba tools tests schema; do
    [ -e "$IZVOR/$d" ] && cp -a "$IZVOR/$d" "$KLON/"
  done
  for f in .gitattributes .gitignore; do
    [ -e "$IZVOR/$f" ] && cp -a "$IZVOR/$f" "$KLON/"
  done
  rm -f "$KLON/tests/last_green.json"
  # stamp-build ne sme da pokusa pravi posao u testu
  printf '#!/usr/bin/env bash\necho "(stamp preskocen u testu)"\n' \
    > "$KLON/tools/stamp-build.sh"
  # vba_check se ZAMENJUJE na istoj putanji, iz dva razloga:
  #   1. pravi traje ~21s, a kapija se u ovom dokazu vrti sedam puta;
  #   2. tako se meri bas WIRE-UP -- da `release.sh` zove `tools/vba_check.py`
  #      i postuje njegov izlazni kod. Da je stub negde drugde, merilo bi se
  #      samo da checker radi, a ne i da ga kapija pita.
  statika 0
  git -C "$KLON" init -q
  git -C "$KLON" config user.email t@t; git -C "$KLON" config user.name t
  git -C "$KLON" checkout -q -B main
  git -C "$KLON" add -A >/dev/null 2>&1
  git -C "$KLON" commit -q -m "test: baza" >/dev/null
  git -C "$KLON" remote add origin "$BARE"
  git -C "$KLON" push -q -u origin main
}

# Postavi izlazni kod stub-ovanog vba_check-a u test repou.
#
# Izmena se KOMITUJE -- inace radni direktorijum ostane prljav, pa kapija stane
# na proveri cistoce i nikad ne stigne do statike. Prva verzija ovog dokaza je
# bila zelena bas zbog toga: tvrdila je "pala statika obara release", a merila
# je "prljav direktorijum obara release". Ista klasa greske kao placebo test.
statika() {
  printf '#!/usr/bin/env python3\nimport sys\nprint("stub vba_check: rc=%s")\nsys.exit(%s)\n' \
    "$1" "$1" > "$KLON/tools/vba_check.py"
  if [ -d "$KLON/.git" ]; then
    git -C "$KLON" commit -q -m "test: statika rc=$1" -- tools/vba_check.py \
      >/dev/null 2>&1 || true
    git -C "$KLON" push -q origin main >/dev/null 2>&1 || true
  fi
}

pokreni() { ( cd "$KLON" && bash tools/release.sh "$@" ) >"$BAZA/out" 2>&1; echo $?; }

ima_tag()  { git -C "$KLON" rev-parse -q --verify "refs/tags/$1" >/dev/null; }
ima_tag_na_origin() { git -C "$BARE" rev-parse -q --verify "refs/tags/$1" >/dev/null; }

# --- 1. BEZ MARKERA: kapija mora da zaustavi tag -------------------------
pripremi
rc="$(pokreni 9.9.9)"
if [ "$rc" = "2" ] && ! ima_tag vba-v9.9.9 && ! ima_tag_na_origin vba-v9.9.9; then
  ok "bez markera: rc=2, tag NIJE nastao ni lokalno ni na origin-u"
else
  pao "bez markera: rc=$rc, tag=$(ima_tag vba-v9.9.9 && echo DA || echo ne)"
  tail -6 "$BAZA/out"
fi
if grep -q "nema markera\|nije dokazan" "$BAZA/out"; then
  ok "nalaz imenuje da izvor nije dokazan"
else
  pao "nalaz ne imenuje razlog"; tail -8 "$BAZA/out"
fi
if [ -n "$(git -C "$KLON" status --porcelain -- src-vba/modConfig.bas)" ] \
   && grep -q '9.9.9' "$KLON/src-vba/modConfig.bas"; then
  ok "bump OSTAJE u radnom drvetu, nekomitovan"
else
  pao "bump je vracen ili komitovan"
fi
if [ -z "$(git -C "$KLON" log origin/main..main --oneline 2>/dev/null)" ]; then
  ok "nema commita na main-u posle pale kapije"
else
  pao "pala kapija je ostavila commit"
fi

# --- 2. PONOVNI POKUSAJ nad zatecenim bump-om je idempotentan -----------
rc="$(pokreni 9.9.9)"
if [ "$rc" = "2" ] && grep -q "zatecen nekomitovan bump" "$BAZA/out"; then
  ok "drugi pokusaj prepoznaje zatecen bump i nastavlja nad njim"
else
  pao "drugi pokusaj ne prepoznaje zatecen bump (rc=$rc)"; tail -8 "$BAZA/out"
fi

# --- 3. PRLJAV RADNI DIREKTORIJUM van bump-a se odbija ------------------
echo "' smece" >> "$KLON/src-vba/modTest.bas"
rc="$(pokreni 9.9.9)"
if [ "$rc" = "1" ] && grep -q "nije cist" "$BAZA/out"; then
  ok "prljav radni direktorijum van bump-a se odbija"
else
  pao "prljav radni direktorijum je prosao (rc=$rc)"; tail -6 "$BAZA/out"
fi
git -C "$KLON" checkout -q -- src-vba/modTest.bas

# --- 3a. STATIKA koja pada obara kapiju ---------------------------------
# Wire-up: kapija mora da PITA tools/vba_check.py i da postuje njegov rc.
statika 2
rc="$(pokreni 9.9.9)"
if [ "$rc" = "2" ] && grep -q "statika" "$BAZA/out" && ! ima_tag vba-v9.9.9; then
  ok "staticka kapija koja pada obara release (i nema taga)"
else
  pao "pala statika je prosla (rc=$rc)"; tail -8 "$BAZA/out"
fi
statika 0

# --- 4. WAIVER BEZ RAZLOGA se odbija ------------------------------------
rc="$(pokreni 9.9.9 --waive ponasanje)"
if [ "$rc" = "1" ] && grep -q "bez --reason" "$BAZA/out"; then
  ok "waiver bez razloga se odbija"
else
  pao "waiver bez razloga je prosao (rc=$rc)"; tail -6 "$BAZA/out"
fi
rc="$(pokreni 9.9.9 --waive izmisljeno --reason x)"
if [ "$rc" = "1" ]; then ok "nepoznata kapija u --waive se odbija"
else pao "nepoznata kapija u --waive je prosla (rc=$rc)"; fi

# --- 5. WAIVER SA RAZLOGOM pusti tag, i razlog je U TAGU ----------------
pripremi
rc="$(pokreni 9.9.9 --waive ponasanje --reason "nema Excela na CI masini")"
if [ "$rc" = "0" ] && ima_tag vba-v9.9.9; then
  ok "waiver sa razlogom pusti tag"
else
  pao "waiver sa razlogom nije pustio tag (rc=$rc)"; tail -10 "$BAZA/out"
fi
telo="$(git -C "$KLON" tag -l --format='%(contents)' vba-v9.9.9 2>/dev/null)"
if printf '%s' "$telo" | grep -q "WAIVED" \
   && printf '%s' "$telo" | grep -q "nema Excela na CI masini"; then
  ok "anotiran tag nosi WAIVED i razlog"
else
  pao "tag ne nosi waiver"; printf '%s\n' "$telo" | head -8
fi
if printf '%s' "$telo" | grep -qE "izvor      [0-9a-f]{64}"; then
  ok "anotiran tag nosi pun otisak izvora"
else
  pao "tag ne nosi otisak izvora"; printf '%s\n' "$telo" | head -8
fi
if ima_tag_na_origin vba-v9.9.9; then ok "tag je push-ovan na origin"
else pao "tag nije push-ovan"; fi

# --- 6. DOKAZAN IZVOR (lazan marker nad pravim otiskom) pusti tag -------
pripremi
PY="$(command -v python3 || command -v python)"
( cd "$KLON" && sed -E 's/APP_VERSION As String = "[^"]*"/APP_VERSION As String = "9.9.9"/' \
    src-vba/modConfig.bas > "$BAZA/_cfg" && mv "$BAZA/_cfg" src-vba/modConfig.bas )
"$PY" - "$KLON" <<'PY'
import importlib.util, json, os, sys, time
klon = sys.argv[1]
s = importlib.util.spec_from_file_location("g", os.path.join(klon, "tools", "vba_gate.py"))
g = importlib.util.module_from_spec(s); s.loader.exec_module(g)
src = os.path.join(klon, "src-vba")
delovi = g.ugovor_delovi(src, klon)
ug = g.otisak_ugovora(delovi)
sv = g.kontekst_sveske(g.podrazumevana_sveska(klon))
kada = time.strftime("%Y-%m-%dT%H:%M:%S")
podaci = {"verzija": g.MARKER_VERZIJA, "suites": {}, "compile": {
    delovi["izvor"]: {"izvor": delovi["izvor"], "kada": kada, "git": "test"}}}
for ime in g.potrebne_suite():
    podaci["suites"][ime] = {"status": "OK", "ukupno": 1, "palo": 0,
                             "izvor": delovi["izvor"], "ugovor": ug,
                             "ugovor_delovi": delovi, "sveska": sv,
                             "kada": kada, "git": "test", "platforma": "test"}
g.upisi_marker(podaci, os.path.join(klon, "tests", "last_green.json"))
print("lazan marker: %d suita + compile nad %s" % (len(podaci["suites"]), delovi["izvor"][:12]))
PY
git -C "$KLON" checkout -q -- src-vba/modConfig.bas
rc="$(pokreni 9.9.9)"
if [ "$rc" = "0" ] && ima_tag vba-v9.9.9 && ima_tag_na_origin vba-v9.9.9; then
  ok "dokazan izvor: kapija prosla, tag nastao i push-ovan"
else
  pao "dokazan izvor nije pustio tag (rc=$rc)"; tail -14 "$BAZA/out"
fi
telo="$(git -C "$KLON" tag -l --format='%(contents)' vba-v9.9.9 2>/dev/null)"
if printf '%s' "$telo" | grep -q "compile: potvrdjen" \
   && ! printf '%s' "$telo" | grep -q "WAIVED"; then
  ok "tag nosi verdikt markera i NIJE obelezen kao waived"
else
  pao "tag ne nosi verdikt"; printf '%s\n' "$telo" | head -12
fi
if [ -n "$(git -C "$KLON" log --oneline -1 --format=%s main | grep 'release: vba v9.9.9')" ]; then
  ok "commit 'release: vba v9.9.9' postoji"
else
  pao "nema release commita"
fi

# --- 7. IZMENA IZVORA POSLE DOKAZA obara kapiju -------------------------
git -C "$KLON" tag -d vba-v9.9.9 >/dev/null 2>&1
git -C "$BARE" tag -d vba-v9.9.9 >/dev/null 2>&1
echo "' izmena posle dokaza" >> "$KLON/src-vba/modTest.bas"
git -C "$KLON" commit -q -am "test: izmena posle dokaza" >/dev/null
git -C "$KLON" push -q origin main
rc="$(pokreni 9.9.10)"
if [ "$rc" = "2" ] && grep -qE "DRUGI izvor|nije dokazan" "$BAZA/out"; then
  ok "izmena izvora posle dokaza obara kapiju"
else
  pao "izmenjen izvor je prosao kapiju (rc=$rc)"; tail -10 "$BAZA/out"
fi

echo
echo "pali: $pali"
rm -rf "$BAZA"
exit $((pali > 0 ? 1 : 0))
