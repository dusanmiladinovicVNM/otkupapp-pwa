# Release notes — AgriX / OtkupApp (VBA)

Za svaku release verziju (`vba-vX.Y.Z`) — par rečenica šta je urađeno.
Dopunjava se pri svakom `tools/release.sh` (korak B-11 u `RELEASE_PROCEDURE.md`).

> Razlika od `ARCHITECTURE_CHANGELOG.md`: tamo je detaljna arhitektonska istorija
> (interno `v6.xx`); ovde su **release-tagovi** (`vba-vX.Y.Z`) sa kratkim opisom za
> korisnika — „šta je novo u ovom .xlsm-u".

**Format:**
```
## vba-vX.Y.Z — YYYY-MM-DD
- promena 1
- promena 2
```

---

## vba-v2.2.x — sledeći release (u pripremi)
Tačan broj/datum se postavlja pri `tools/release.sh`.

- **Verzionisanje koda:** novi `modBuildInfo` (`BUILD_SHA` / `BUILD_VERSION` / `BUILD_DATE`), stamp pri buildu (`tools/stamp-build`).
- **Auto verzija:** `BUILD_VERSION` iz `git describe` (na tagu čisto, između tagova se sam diže).
- **Telemetrija builda:** `modMonitoring` i `modLicense` šalju `buildSha`/`buildVersion`/`buildDate`.
- **Fleet pregled „ko ima koju verziju":** GAS `Events`/`Fleet` + `rebuildMonitoringFleet` (auto na sat preko `installMonitoringTriggers`).
- **Release rutina:** `tools/release.sh|ps1` (jedna komanda) + procedura `docs/RELEASE_PROCEDURE.md`.
- **Min-version gate:** `modUpdateGate` na startu pita GAS (`checkVersion`) za minimalnu dozvoljenu verziju; zastarela verzija dobija upozorenje (ili blok uz `VERSION_ENFORCE=YES`). Opt-in + fail-open. Podešava se preko GAS Script Properties (`VERSION_MIN`/`VERSION_LATEST`/`VERSION_ENFORCE`/`VERSION_MESSAGE`), bez redeploy-a.
- **Blanko build guard:** `modBuildGuard.AssertBlankBuild` (ručno, build-only) proveri da fajl pre distribucije nema podataka — sprečava da blanko `AgriX_x.x.x` ode klijentima sa tuđim podacima.
- **Distribucija:** konvencija `builds\AgriX_x.x.x.xlsm` (blanko, isti za sve) dokumentovana u proceduri.
