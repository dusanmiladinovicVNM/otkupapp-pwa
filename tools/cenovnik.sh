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

case "${1:-}" in
  build)
    "$ROOT/tools/render-pdf.sh" "$SRC" "$OUT"
    ;;

  check)
    python3 - "$ROOT" <<'PY'
import sys, re, os
root = sys.argv[1]
html = open(os.path.join(root, 'docs/Sales/AgriX_Cenovnik_2027.html'), encoding='utf-8').read()
greske = []

# --- 1) svi iznosi iz .html izvora, iz atributa a ne iz vidljivog teksta ---
def procitaj(atr):
    out = {}
    for m in re.finditer(r'data-%s="([^"]+)"\s+data-eur="(\d+)"' % atr, html):
        out[m.group(1)] = int(m.group(2))
    return out

cene = procitaj('cena')          # osnovne cene
izvedeno = procitaj('izvedeno')  # zbirovi u primerima

paketi = {'Desktop', 'Desktop all-in', 'Mobile', 'Mobile all-in'}
if not paketi <= set(cene):
    print('GRESKA: .html ne nosi data-cena za sve pakete:', sorted(paketi - set(cene)))
    sys.exit(1)

# vidljivi tekst mora da prati atribut
for naziv, eur in cene.items():
    m = re.search(r'data-cena="%s"\s+data-eur="%d"[^>]*>[^<]*?([\d.]+)' % (re.escape(naziv), eur), html)
    if not m or int(m.group(1).replace('.', '')) != eur:
        greske.append(f'{naziv}: atribut {eur} ne prati prikazani tekst')

# --- 2) struktura cena iz odluka 414 i 415 ---
if cene['Mobile'] - cene['Desktop'] != 1000:
    greske.append(f"odluka 414: Mobile dodatak je {cene['Mobile']-cene['Desktop']}, ocekivano 1000")
if cene['Mobile all-in'] - cene['Desktop all-in'] != 1000:
    greske.append(f"odluka 414: Mobile dodatak na all-in nivou je {cene['Mobile all-in']-cene['Desktop all-in']}, ocekivano 1000")
for baza, allin in (('Desktop', 'Desktop all-in'), ('Mobile', 'Mobile all-in')):
    if cene[allin] - cene[baza] != 700:
        greske.append(f"odluka 415: all-in doplata za {baza} je {cene[allin]-cene[baza]}, ocekivano 700")

# --- 3) mrtva tacka: nijedan podskup modula ne sme da kosta koliko all-in doplata ---
from itertools import combinations
for naziv, moduli in (('Desktop all-in', ['Modul SEF', 'Modul Banka', 'Modul Hladnjača / Proizvodnja']),
                      ('Mobile all-in',  ['Modul SEF', 'Modul Banka', 'Modul Hladnjača / Proizvodnja', 'Modul Dispatch'])):
    v = [cene[m] for m in moduli if m in cene]
    zbirovi = {sum(c) for r in range(1, len(v)+1) for c in combinations(v, r)}
    if 700 in zbirovi:
        greske.append(f'{naziv}: podskup modula se sabira na 700 — a la carte kosta isto kao all-in')
    if sum(v) <= 700:
        greske.append(f'odluka 415: doplata 700 nije niza od zbira modula ({sum(v)}) za {naziv}')

# --- 4) izvedeni zbirovi u primerima moraju da se slazu sa komponentama ---
def izracunaj(izraz):
    ukupno = 0
    for clan in izraz.split('+'):
        clan = clan.strip()
        faktor = 1.0
        m = re.match(r'^(\d+)\*(.+)$', clan)
        if m:
            faktor, clan = int(m.group(1)), m.group(2).strip()
        elif clan.endswith('/2'):
            faktor, clan = 0.5, clan[:-2].strip()
        if clan not in cene:
            return None, clan
        ukupno += faktor * cene[clan]
    return ukupno, None

for izraz, prikazano in izvedeno.items():
    vrednost, nepoznato = izracunaj(izraz)
    if nepoznato:
        greske.append(f'izvedeni zbir "{izraz}": nepoznata stavka "{nepoznato}"')
    elif abs(vrednost - prikazano) > 0.001:
        greske.append(f'izvedeni zbir "{izraz}": prikazano {prikazano}, izracunato {vrednost:.0f}')

# --- 5) ostala tri mesta sa cenama ---
try:
    import openpyxl
except ImportError:
    print('UPOZORENJE: openpyxl nije instaliran, preskacem .xlsx provere')
    openpyxl = None

if openpyxl:
    cz = openpyxl.load_workbook(os.path.join(root, 'docs/Sales/AgriX_Sablon_ponude.xlsx'))['Cenovnik']
    sab = {}
    for r in range(5, cz.max_row + 1):
        naziv, cena = cz[f'A{r}'].value, cz[f'B{r}'].value
        if naziv and isinstance(cena, (int, float)):
            sab[naziv.replace(' [u pripremi]', '')] = cena
    # nazivi u .html-u prate nazive u listu Cenovnik, uz mali broj mapiranja
    mapa = {'Desktop': 'AgriX Desktop', 'Desktop all-in': 'AgriX Desktop all-in',
            'Mobile': 'AgriX Mobile', 'Mobile all-in': 'AgriX Mobile all-in',
            'Gazdinstvo Basic — maloprodaja': None, 'Gazdinstvo Pro — maloprodaja': None}
    for naziv, eur in cene.items():
        kljuc = mapa.get(naziv, naziv)
        if kljuc is None:
            continue  # maloprodajne cene Gazdinstva ne postoje u sablonu ponude
        if kljuc not in sab:
            greske.append(f'sablon ponude: nema stavke "{kljuc}"')
        elif sab[kljuc] != eur:
            greske.append(f'sablon ponude: {kljuc} = {sab[kljuc]}, u cenovniku {eur}')

    pr = openpyxl.load_workbook(os.path.join(root, 'docs/Finance/AgriX_Finansijski_model.xlsx'))['Pretpostavke']
    fin = {pr[f'A{r}'].value: pr[f'B{r}'].value for r in range(6, 21)}
    for naziv, kljuc in (('Desktop', 'Desktop Core'), ('Desktop all-in', 'Desktop all-in'),
                         ('Mobile', 'Mobile'), ('Mobile all-in', 'Mobile all-in'),
                         ('Stanica preko pet', 'Stanica preko pet')):
        if fin.get(kljuc) != cene[naziv]:
            greske.append(f'finansijski model: {kljuc} = {fin.get(kljuc)}, u cenovniku {cene[naziv]}')

ug = open(os.path.join(root, 'docs/Legal/AgriX_Ugovor_o_licenciranju.md'), encoding='utf-8').read()
nadjeno = [int(x.replace('.', '')) for x in re.findall(r'AgriX (?:Desktop|Mobile)(?: all-in)?\s*\|\s*([\d.]+)', ug)]
ocekivano = [cene['Desktop'], cene['Desktop all-in'], cene['Mobile'], cene['Mobile all-in']]
if nadjeno != ocekivano:
    greske.append(f'ugovor Prilog 1: {nadjeno}, ocekivano {ocekivano}')

if greske:
    print('RAZLIKA u cenama:')
    for g in greske:
        print('  -', g)
    sys.exit(1)
print(f'OK: {len(cene)} cena i {len(izvedeno)} izvedenih zbirova; struktura 414/415 i sva cetiri mesta se slazu')
PY
    ;;

  *)
    sed -n '2,7p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
    exit 1
    ;;
esac
