---
paths:
  - "src-vba/modSelfUpdate.bas"
  - "src-vba/modRelease.bas"
  - "src-vba/modDrive.bas"
  - "src-vba/modGoogleAuth.bas"
  - "src-vba/mod*Sync.bas"
  - "src-vba/modMain.bas"
  - "gas/**"
  - "src/**"
---

# Sync / PWA i self-update

> Preseljeno iz `CLAUDE.md` §3.

## Sync / PWA

`modStammdatenSync`, `modMasterSync`, `gas/`.

Google/PWA kredencijali žive u **`tblSEFConfig`** (`GetConfigValue`) — ne u
`tblLocalConfig`, ne u legacy `tblConfig` (vidi `.claude/rules/podaci-i-config.md`).
Auth setup: `modGoogleAuth.RunGoogleAuthSetup`.

## Self-update (kod)

- klijent: `modSelfUpdate` (`CheckForUpdateOnOpen` / `RunSelfUpdate`, dvofazni)
- build: `modRelease.PublishReleaseToDrive`
- Drive REST: `modDrive`
- **Pre bilo kakve izmene pročitaj `docs/SELF_UPDATE.md` — tamo su zamke.**

### Zamka #19 — `modSelfUpdate` je frozen (`SKIP_MODULES`)

Updatable moduli (`modMain`…) **NE smeju early-bind-ovati NOV `modSelfUpdate`
simbol**. Star klijent posle self-update-a = nov `modMain` + star `modSelfUpdate`
→ `Compile error: Sub or Function not defined` obori `StartApp`.

Nov cross-poziv sakrij iza **postojećeg stabilnog simbola** ili ga zovi
late-bound (`Application.Run`).

### Zamka #11 — nove `WithEvents` deklaracije u formama

Lome code-merge te forme pri self-update-u. Vidi
`.claude/rules/forme-i-kontrole.md`.
