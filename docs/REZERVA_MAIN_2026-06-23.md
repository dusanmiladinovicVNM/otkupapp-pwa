# Rezervna kopija `main` grane — 2026-06-23

Ova grana (`backup/main-2026-06-23`) je **rezervna kopija (snapshot) `main` grane**,
napravljena PRE nego što `radna-verzija` pregazi `main`. Cilj: sačuvati dotadašnji
`main` kao rezervu (rollback tačku), da se ništa ne izgubi.

## Tačno kada je urađeno

| | |
|---|---|
| **Backup napravljen** | **2026-06-23 17:39:12 CEST (+0200)** · `15:39:12 UTC` |
| **Snimljeni `main` (HEAD)** | `3d4c4ae3d51e18349612e84b141d31c6b6189666` |
| **Poruka tog commita** | „Add functions for screen dimensions and user form styling" |
| **Datum tog commita** | 2026-06-22 23:26:03 +0200 |

> Sam sadržaj `main`-a u trenutku backup-a je commit **`3d4c4ae`** — on je
> roditelj (`HEAD~1`) ove grane. Ovaj `.md` fajl je jedina razlika u odnosu na
> tadašnji `main`, dodat samo da dokumentuje snapshot.

## Stanje u trenutku backup-a (`main` vs `radna-verzija`)

- `main` → `3d4c4ae` (2026-06-22 23:26 +0200)
- `radna-verzija` → `d4b1c4d` „aktuelna radna verzija" (2026-06-23 12:37 +0200)
- Zajednički predak (merge-base): `1e6dd1c`
- Divergencija: `main` **+1** commit / `radna-verzija` **+16** commita; razlika **172 fajla** (5800+ / 1591−)

### ⚠️ Commit koji postoji samo na `main`-u

`3d4c4ae` „Add functions for screen dimensions and user form styling" je jedini
commit koji `main` ima a `radna-verzija` **nema**. Kada `radna-verzija` pregazi
`main`, taj rad nestaje iz `main` linije — ostaje sačuvan samo ovde, u rezervi.
Ako treba da uđe i u novu verziju, prebaciti ga (cherry-pick) u `radna-verzija`
pre gaženja.

## Kako vratiti `main` na ovu rezervu (rollback)

```bash
# pregled
git fetch origin
git log --oneline backup/main-2026-06-23

# vraćanje main-a na tačno stanje od 2026-06-23 (sadržaj = 3d4c4ae)
git checkout main
git reset --hard 3d4c4ae3d51e18349612e84b141d31c6b6189666
git push --force-with-lease origin main
```
