---
paths:
  - "tools/release.sh"
  - "tools/release.ps1"
  - "tools/stamp-build.sh"
  - "tools/stamp-build.ps1"
  - "docs/RELEASE_NOTES.md"
  - "docs/RELEASE_PROCEDURE.md"
---

# Git, PR i release

> Preseljeno iz `CLAUDE.md` §6. Ovo je procedura na KRAJU rada — ne mora da
> zauzima kontekst dok se piše kod.

## Git

- Razvoj na zadatoj feature grani; commit poruke jasne i opisne.
- **Ne praviti PR bez eksplicitnog zahteva.**
- Pre merge-a u `main` proveriti konflikte (`git merge-tree`) i preklapanja
  fajlova.

## Integracija ažuriranog `main`-a u feature granu = UVEK „Opcija 3"

1. `git fetch origin main`
2. proveri preklapanja fajlova + `git merge-tree`
3. **rebase lokalno** na `origin/main`
4. **pokaži rezultat** (log, diff vs `main`, statičke provere — uključujući
   `python3 tools/vba_check.py`)
5. `git push --force-with-lease` **tek po eksplicitnom odobrenju**

Nikad force-push pre pokazivanja.

## Posle kreiranja PR-a ka `main`

Podseti korisnika na release/verzionisanje:
`tools/release.sh <verzija>` → Excel `ImportAllVBA` → `Compile` → snimi → ship →
`Fleet` provera, da se novi `AgriX_OtkupApp.xlsm` pravilno verzioniše.
Vidi `docs/RELEASE_PROCEDURE.md` i dopuni `docs/RELEASE_NOTES.md`.

## Na kraju SVAKE izmene koda (posle commit/push)

UVEK daj git bash komandu za preuzimanje feature grane radi testa kroz
`ImportAllVBA`. Lokalni klon je `~/Documents/GitHub/otkupapp-pwa`
(= `ImportAllVBA` folder):

```bash
cd ~/Documents/GitHub/otkupapp-pwa
git fetch origin <grana>
git checkout <grana>
git pull --ff-only origin <grana>
```

Zatim u Excelu: `Alt+F8 → ImportAllVBA → Debug → Compile → snimi → test`.

## I uz to — kratka test-checklista u chatu

Numerisani, konkretni koraci šta operater proba u Excelu (klik po klik +
očekivani rezultat), fokusiran na ono što je u toj izmeni dodato/promenjeno.
Kratko i praktično.
