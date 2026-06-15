# AgriX / OtkupApp — Desktop licenciranje (offline + online kill-switch)

> **Status:** Implementirano, ali NIJE testirano na realnom Windows/Excel okruženju.
> Pre produkcije obavezno proći "Verifikacija na realnoj mašini" na dnu.
> **VBA je otvoren** — ovo je deterrent protiv kopiranja/deljenja, ne kriptografski zid.

## 1. Šta radi

Na `Workbook_Open → StartApp` zove se `modLicense.LicenseGateOrQuit`:

1. **Enforcement flag** — ako `LICENSE_ENFORCE` u `tblSEFConfig` nije uključen, gate je **no-op** (uvek dozvoli). Tako dodavanje modula ne može da zaključa aplikaciju dok ti to svesno ne uključiš.
2. **Offline provera** — čita `license.lic`, proverava da je otisak mašine isti, da nije isteklo, i da je **digitalni potpis** validan (javni `license_pub.cer`). Radi bez mreže.
3. **Online kill-switch** (opciono) — ako je `LICENSE_CHECK_URL` podešen, pita GAS endpoint da li je licenca opozvana; odgovor je **potpisan** pa se ne može falsifikovati. Ako je mreža nedostupna → **grace prozor** (podrazumevano 30 dana) da kratak prekid ne zaključa korisnika.

## 2. Config ključevi (`tblSEFConfig`)

| Ključ | Vrednost | Podrazumevano |
|---|---|---|
| `LICENSE_ENFORCE` | `YES` da uključiš proveru | (prazno = isključeno) |
| `LICENSE_CHECK_URL` | GAS web app `/exec` URL (opciono, samo HTTPS) | (prazno = čisto offline) |
| `LICENSE_OFFLINE_GRACE_DAYS` | broj dana grace-a kad je online podešen ali nedostupan | `30` |
| `LICENSE_LAST_ONLINE_OK_AT` | (postavlja app sama) | — |

## 3. Fajlovi po mašini

Traže se u `...\Secrets\` pored radne sveske, pa pored same sveske:

- `license_pub.cer` — tvoj javni sertifikat (isti za sve stanice).
- `license.lic` — po mašini, vezan za otisak.

## 4. Postupak puštanja

**Jednom (kod tebe):**
```powershell
.\install\license\New-LicenseKeypair.ps1
# -> license_priv.pem (TAJNO), license_pub.cer (javno)
```
Kopiraj `license_pub.cer` u `...\Secrets` na svakoj stanici.

**Po stanici:**
1. Operater u Excelu pokrene makro `ShowMyFingerprint` i pošalje ti otisak.
2. Ti izdaš licencu:
   ```powershell
   .\install\license\Issue-License.ps1 -Fingerprint "A1B2..." `
       -Customer "Stanica Novi Sad" -ExpiresAt "2027-06-15"
   ```
3. Pošalji `license.lic`, operater ga ubaci u `...\Secrets`.
4. Postavi `LICENSE_ENFORCE = YES` u `tblSEFConfig`.

**Opciono — online kill-switch:**
1. Deploy `install/license/gas-checkLicense.gs` kao Web App.
2. U Script properties stavi `LICENSE_PRIV_PEM` (isti keypair) i `REVOKED_LIST`.
3. U `tblSEFConfig` postavi `LICENSE_CHECK_URL`.
Opozivanje ukradene mašine = dodaš njen otisak u `REVOKED_LIST`.

## 5. Failure-mode dizajn (svesne odluke)

- **Neispravna/nedostajuća licenca → fail-closed** (poruka + zatvaranje sveske), prikazuje otisak mašine.
- **Greška u samoj proveri (bag) → fail-open + glasan log** (`LicenseGateOrQuit` EH). Namerno, da bag u licenciranju ne zaključa legitimne korisnike dok se ne verifikuje na realnoj mašini. Posle verifikacije možeš da prebaciš na fail-closed.
- **Mreža nedostupna, online podešen → grace** (`LICENSE_OFFLINE_GRACE_DAYS`).

## 6. Zavisnosti

- Windows PowerShell 5.1 + .NET 4.6+ (stock Win10/11). Klijent NE traži OpenSSL.
- Issuer (ti): OpenSSL.
- GAS potpis: `Utilities.computeRsaSha256Signature` (RSA-SHA256, PKCS1) — kompatibilno sa klijentskom verifikacijom.

## 7. Ograničenja (budi iskren prema sebi)

- Ko otvori VBE može da zakomentariše `LicenseGateOrQuit` — zato: zaključaj VBA projekat, isporuči `.xlsb`, i razmisli o proveri na više mesta.
- Otisak se menja na reinstalaciji Windowsa / zameni diska C: → tad se izdaje nova licenca.
- Pravi nivo zaštite = logovati aktivacije (kad ima mreže) da **vidiš** neovlašćene kopije i rešiš ih komercijalno/pravno.

## 8. Verifikacija na realnoj mašini (uraditi pre produkcije)

- [ ] `?GetMachineFingerprint()` u Immediate vraća stabilan otisak (isti pri 2 pokretanja).
- [ ] Bez `LICENSE_ENFORCE` → app radi normalno (no-op).
- [ ] Sa validnim `license.lic` + `LICENSE_ENFORCE=YES` → app se otvara.
- [ ] Pokvari jedan karakter u `sig=` → app blokira ("potpis nije validan").
- [ ] `license.lic` sa tuđim fingerprint-om → blokira ("druga mašina").
- [ ] `expiresAt` u prošlosti → blokira ("istekla").
- [ ] Online: dodaj otisak u `REVOKED_LIST` → app blokira ("opozvana").
- [ ] Online podešen, internet isključen → app radi unutar grace prozora.
- [ ] Proveri da PowerShell prozor ne "bljesne" pri startu (Run sa hidden=0 + wait).
