# Ko pise koju tabelu

> **Generisan fajl -- ne menjaj rukom.**
> `python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md`

Izvedeno iz dva mehanicka signala u `src-vba/`:

- **mutate** -- `AppendRow` / `UpdateCell` / `RequireUpdateCell`:
  modul stvarno MENJA redove. Samo ovo je vlasnistvo (ugovor A11).
- **tx** -- `clsTransaction.AddTableSnapshot TBL_X`: operacija
  snapshotuje tabelu da bi `RollbackTx` umeo da je vrati. To je
  UCESCE u transakciji, ne vlasnistvo -- koordinator sme da
  snapshotuje tudju tabelu i zove API njenog vlasnika.

Test moduli su odvojeni: pisu uz rollback i nisu vlasnici podataka.

**Cemu sluzi:** kad isto polje pise vise mesta po razlicitim pravilima,
to je klasa buga koju test hvata tek posle nastanka. Pre nego sto
promenis pravilo upisa, ovde vidis ko jos pise istu tabelu.

| Tabela | Mutatora | Moduli koji MENJAJU redove |
|---|---|---|
| `tblOtkup` | 9 | `modAutoHladnjaca`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`, `modOtkupBlok`, `modSetup`, `modSledljivost`, `modStornoFlow` |
| `tblFakturaStavke` | 3 | `modDokumenta`, `modStorno`, `modUtovar` |
| `tblFakture` | 3 | `modFaktura`, `modSEFPersistance`, `modStorno` |
| `tblKorisnici` | 3 | `modAuth`, `modMaticniKorisnici`, `modSetup` |
| `tblNovac` | 3 | `modBankaMapiranje`, `modNovac`, `modStorno` |
| `tblPrijemnica` | 3 | `modDokumenta`, `modFaktura`, `modStorno` |
| `tblBankaImport` | 2 | `modBankaMapiranje`, `modStorno` |
| `tblPaleta` | 2 | `modPaletniList`, `modStorno` |
| `tblParcele` | 2 | `modGeoParcele`, `modMasterSync` |
| `tblUtovar` | 2 | `modStorno`, `modUtovar` |
| `tblAmbalaza` | 1 | `modStornoRecovery` |
| `tblArtikli` | 1 | `modAgrohemija` |
| `tblOtpremnica` | 1 | `modDokumenta` |
| `tblPaletaStavka` | 1 | `modPaletniList` |
| `tblPrevoznici` | 1 | `modUtovar` |
| `tblSEFConfig` | 1 | `modConfig` |
| `tblSEFSubmission` | 1 | `modSEFPersistance` |
| `tblStornoVeze` | 1 | `modStornoContext` |
| `tblUtovarStavke` | 1 | `modUtovar` |
| `tblZbirna` | 1 | `modDokumentInvariant` |
| `tblKooperanti` | 0 | _(samo testovi)_ |
| `tblKulture` | 0 | _(samo testovi)_ |
| `tblKupci` | 0 | _(samo testovi)_ |
| `tblKutije` | 0 | _(samo testovi)_ |
| `tblMagacin` | 0 | _(samo testovi)_ |
| `tblPartnerMap` | 0 | _(samo testovi)_ |
| `tblPrerada` | 0 | _(samo testovi)_ |
| `tblPreradaStavka` | 0 | _(samo testovi)_ |
| `tblSEFEventLog` | 0 | _(samo testovi)_ |
| `tblStanice` | 0 | _(samo testovi)_ |
| `tblStornoZurnal` | 0 | _(samo testovi)_ |
| `tblTipAmbalaze` | 0 | _(samo testovi)_ |

## Ucesnici transakcije (snapshot, NE vlasnistvo)

- `tblOtkup`: `modAutoHladnjaca`, `modBankaMapiranje`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`, `modOtkupBlok`, `modSledljivost`, `modStorno`, `modStornoFlow`, `modStornoRecovery`
- `tblFakturaStavke`: `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblFakture`: `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblKorisnici`: `modMaticniKorisnici`
- `tblNovac`: `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modOtkup`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblPrijemnica`: `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow`
- `tblBankaImport`: `modBankaImport`, `modBankaMapiranje`, `modStorno`
- `tblPaleta`: `modDokumenta`, `modPaletniList`, `modStorno`
- `tblParcele`: `modGeoParcele`, `modMasterSync`
- `tblUtovar`: `modStorno`, `modUtovar`
- `tblAmbalaza`: `modDokumenta`, `modMasterSync`, `modOtkup`, `modStorno`, `modStornoFlow`, `modStornoRecovery`
- `tblOtpremnica`: `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow`
- `tblPaletaStavka`: `modDokumenta`, `modPaletniList`, `modStorno`
- `tblSEFSubmission`: `modSEFService`, `modSEFStatusSync`, `modSEFValidator`
- `tblStornoVeze`: `modStornoContext`
- `tblUtovarStavke`: `modStorno`, `modUtovar`
- `tblZbirna`: `modDokumentInvariant`, `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow`
- `tblKooperanti`: `modKooperant`
- `tblMagacin`: `modAgroUnos`, `modAgrohemija`
- `tblPartnerMap`: `modBankaMapiranje`
- `tblPrerada`: `modPaletniList`, `modStorno`
- `tblPreradaStavka`: `modPaletniList`, `modStorno`
- `tblSEFEventLog`: `modSEFService`, `modSEFStatusSync`, `modSEFValidator`
- `tblStornoZurnal`: `modStorno`, `modStornoFlow`

## Test moduli po tabeli

- `tblOtkup`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblFakturaStavke`: `modTest`, `modTestStorno`
- `tblFakture`: `modSEFTests`, `modTest`, `modTestBanka`, `modTestStorno`
- `tblNovac`: `modNovacTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblPrijemnica`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblBankaImport`: `modTestBanka`, `modTestStorno`
- `tblPaleta`: `modBusinessFlowProTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblParcele`: `modAgrohemijaTests`
- `tblAmbalaza`: `modBusinessFlowProTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblArtikli`: `modAgrohemijaTests`
- `tblOtpremnica`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPaletaStavka`: `modBusinessFlowProTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblSEFConfig`: `modTestStorno`
- `tblSEFSubmission`: `modSEFTests`
- `tblStornoVeze`: `modBusinessFlowProTests`, `modTest`, `modTestStorno`, `modTestStornoCentar`
- `tblZbirna`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblKooperanti`: `modAgrohemijaTests`, `modTestBanka`
- `tblKulture`: `modTestPalete`
- `tblKupci`: `modTestBanka`
- `tblKutije`: `modTest`
- `tblMagacin`: `modAgrohemijaTests`, `modTest`
- `tblPartnerMap`: `modTestBanka`
- `tblSEFEventLog`: `modSEFTests`
- `tblStanice`: `modTestBanka`
- `tblStornoZurnal`: `modTestStornoCentar`
- `tblTipAmbalaze`: `modTestPalete`

## Sta ovo NE pokriva

- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad
  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --
  ako ga nadjes, to je nalaz, ne rupa u mapi.
- Granularnost je tabela, ne kolona.

