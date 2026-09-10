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
| `tblFakturaStavke` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modUtovar` |
| `tblFakture` | 4 | `modFaktura`, `modSEFPersistance`, `modStorno`, `modUtovar` |
| `tblBankaImport` | 3 | `modBankaImport`, `modBankaMapiranje`, `modStorno` |
| `tblKorisnici` | 3 | `modAuth`, `modMaticniKorisnici`, `modSetup` |
| `tblNovac` | 3 | `modBankaMapiranje`, `modNovac`, `modStorno` |
| `tblPrijemnica` | 3 | `modDokumenta`, `modFaktura`, `modStorno` |
| `tblZbirna` | 3 | `modDokumentInvariant`, `modDokumenta`, `modMasterSync` |
| `tblAmbalaza` | 2 | `modAmbalaza`, `modStornoRecovery` |
| `tblPaleta` | 2 | `modPaletniList`, `modStorno` |
| `tblParcele` | 2 | `modGeoParcele`, `modMasterSync` |
| `tblUtovar` | 2 | `modStorno`, `modUtovar` |
| `tblArtikli` | 1 | `modAgrohemija` |
| `tblCenovnik` | 1 | `modCenovnik` |
| `tblKooperanti` | 1 | `modKooperant` |
| `tblMagacin` | 1 | `modAgrohemija` |
| `tblOtkupStavke` | 1 | `modOtkup` |
| `tblOtpremnica` | 1 | `modDokumenta` |
| `tblOtpremnicaIzvori` | 1 | `modDokumenta` |
| `tblOtpremnicaStavke` | 1 | `modDokumenta` |
| `tblPaletaStavka` | 1 | `modPaletniList` |
| `tblPartnerMap` | 1 | `modNovac` |
| `tblPrevoznici` | 1 | `modUtovar` |
| `tblSEFConfig` | 1 | `modConfig` |
| `tblSEFEventLog` | 1 | `modSEFPersistance` |
| `tblSEFSubmission` | 1 | `modSEFPersistance` |
| `tblStornoVeze` | 1 | `modStornoContext` |
| `tblStornoZurnal` | 1 | `modStornoZurnal` |
| `tblUtovarStavke` | 1 | `modUtovar` |
| `tblVozaci` | 1 | `modMalina` |
| `tblZbirnaIzvori` | 1 | `modDokumenta` |
| `tblZbirnaStavke` | 1 | `modDokumenta` |
| `tblKulture` | 0 | _(samo testovi)_ |
| `tblKupci` | 0 | _(samo testovi)_ |
| `tblKutije` | 0 | _(samo testovi)_ |
| `tblPrerada` | 0 | _(samo testovi)_ |
| `tblPreradaStavka` | 0 | _(samo testovi)_ |
| `tblStanice` | 0 | _(samo testovi)_ |
| `tblTipAmbalaze` | 0 | _(samo testovi)_ |

## Ucesnici transakcije (snapshot, NE vlasnistvo)

- `tblOtkup`: `modAutoHladnjaca`, `modBankaMapiranje`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`, `modOtkupBlok`, `modSledljivost`, `modStorno`, `modStornoFlow`, `modStornoRecovery`
- `tblFakturaStavke`: `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblFakture`: `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblBankaImport`: `modBankaImport`, `modBankaMapiranje`, `modStorno`
- `tblKorisnici`: `modMaticniKorisnici`
- `tblNovac`: `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modOtkup`, `modStorno`, `modStornoFlow`, `modUtovar`
- `tblPrijemnica`: `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow`
- `tblZbirna`: `modDokumentInvariant`, `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow`
- `tblAmbalaza`: `modDokumenta`, `modMasterSync`, `modOtkup`, `modStorno`, `modStornoFlow`, `modStornoRecovery`
- `tblPaleta`: `modDokumenta`, `modPaletniList`, `modStorno`
- `tblParcele`: `modGeoParcele`, `modMasterSync`
- `tblUtovar`: `modStorno`, `modUtovar`
- `tblKooperanti`: `modKooperant`
- `tblMagacin`: `modAgroUnos`, `modAgrohemija`
- `tblOtkupStavke`: `modOtkup`
- `tblOtpremnica`: `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow`
- `tblOtpremnicaIzvori`: `modDokumenta`
- `tblOtpremnicaStavke`: `modDokumenta`
- `tblPaletaStavka`: `modDokumenta`, `modPaletniList`, `modStorno`
- `tblPartnerMap`: `modBankaMapiranje`
- `tblSEFEventLog`: `modSEFService`, `modSEFStatusSync`, `modSEFValidator`
- `tblSEFSubmission`: `modSEFService`, `modSEFStatusSync`, `modSEFValidator`
- `tblStornoVeze`: `modStornoContext`
- `tblStornoZurnal`: `modStorno`, `modStornoFlow`
- `tblUtovarStavke`: `modStorno`, `modUtovar`
- `tblZbirnaIzvori`: `modDokumenta`
- `tblZbirnaStavke`: `modDokumenta`
- `tblPrerada`: `modPaletniList`, `modStorno`
- `tblPreradaStavka`: `modPaletniList`, `modStorno`

## Test moduli po tabeli

- `tblOtkup`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoldenTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblFakturaStavke`: `modGoldenTests`, `modTest`, `modTestStorno`
- `tblFakture`: `modGoldenTests`, `modSEFTests`, `modTest`, `modTestBanka`, `modTestStorno`
- `tblBankaImport`: `modTestBanka`, `modTestStorno`
- `tblNovac`: `modGoldenTests`, `modNovacTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblPrijemnica`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoldenTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblZbirna`: `modBusinessFlowProTests`, `modGoldenTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblAmbalaza`: `modBusinessFlowProTests`, `modGoldenTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPaleta`: `modBusinessFlowProTests`, `modGoldenTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblParcele`: `modAgrohemijaTests`
- `tblArtikli`: `modAgrohemijaTests`
- `tblKooperanti`: `modAgrohemijaTests`, `modGoldenTests`, `modTestBanka`
- `tblMagacin`: `modAgrohemijaTests`, `modTest`
- `tblOtpremnica`: `modBusinessFlowProTests`, `modGoldenTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPaletaStavka`: `modBusinessFlowProTests`, `modGoldenTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblPartnerMap`: `modTestBanka`
- `tblSEFConfig`: `modTestStorno`
- `tblSEFEventLog`: `modSEFTests`
- `tblSEFSubmission`: `modSEFTests`
- `tblStornoVeze`: `modBusinessFlowProTests`, `modTest`, `modTestStorno`, `modTestStornoCentar`
- `tblStornoZurnal`: `modTestStornoCentar`
- `tblVozaci`: `modGoldenTests`
- `tblZbirnaIzvori`: `modBusinessFlowProTests`
- `tblKulture`: `modTestPalete`
- `tblKupci`: `modGoldenTests`, `modTestBanka`
- `tblKutije`: `modTest`
- `tblStanice`: `modGoldenTests`, `modTestBanka`
- `tblTipAmbalaze`: `modTestPalete`

## Sta ovo NE pokriva

- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad
  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --
  ako ga nadjes, to je nalaz, ne rupa u mapi.
- Granularnost je tabela, ne kolona.

