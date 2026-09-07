# Ko pise koju tabelu

> **Generisan fajl -- ne menjaj rukom.**
> `python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md`

Izvedeno iz dva mehanicka signala u `src-vba/`:

- **tx** -- `clsTransaction.AddTableSnapshot TBL_X`: operacija sama
  deklarise koje tabele menja, da bi `RollbackTx` umeo da ih vrati.
- **direct** -- `AppendRow` / `UpdateCell` kroz `modDataAccess`.

Test moduli su odvojeni: pisu uz rollback i nisu vlasnici podataka.

**Cemu sluzi:** kad isto polje pise vise mesta po razlicitim pravilima,
to je klasa buga koju test hvata tek posle nastanka. Pre nego sto
promenis pravilo upisa, ovde vidis ko jos pise istu tabelu.

| Tabela | Pisaca | Produkcioni moduli |
|---|---|---|
| `tblOtkup` | 12 | `modAutoHladnjaca`, `modBankaMapiranje`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`, `modOtkupBlok`, `modSetup`, `modSledljivost`, `modStorno`, `modStornoFlow`, `modStornoRecovery` |
| `tblFakture` | 10 | `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator`, `modStorno`, `modStornoFlow`, `modUtovar` |
| `tblNovac` | 8 | `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modOtkup`, `modStorno`, `modStornoFlow`, `modUtovar` |
| `tblAmbalaza` | 7 | `modAmbalaza`, `modDokumenta`, `modMasterSync`, `modOtkup`, `modStorno`, `modStornoFlow`, `modStornoRecovery` |
| `tblFakturaStavke` | 5 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow`, `modUtovar` |
| `tblZbirna` | 5 | `modDokumentInvariant`, `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblOtpremnica` | 4 | `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblPrijemnica` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow` |
| `tblSEFEventLog` | 4 | `modSEFPersistance`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblSEFSubmission` | 4 | `modSEFPersistance`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblBankaImport` | 3 | `modBankaImport`, `modBankaMapiranje`, `modStorno` |
| `tblKorisnici` | 3 | `modAuth`, `modMaticniKorisnici`, `modSetup` |
| `tblPaleta` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblPaletaStavka` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblStornoZurnal` | 3 | `modStorno`, `modStornoFlow`, `modStornoZurnal` |
| `tblMagacin` | 2 | `modAgroUnos`, `modAgrohemija` |
| `tblParcele` | 2 | `modGeoParcele`, `modMasterSync` |
| `tblPartnerMap` | 2 | `modBankaMapiranje`, `modNovac` |
| `tblPrerada` | 2 | `modPaletniList`, `modStorno` |
| `tblPreradaStavka` | 2 | `modPaletniList`, `modStorno` |
| `tblUtovar` | 2 | `modStorno`, `modUtovar` |
| `tblUtovarStavke` | 2 | `modStorno`, `modUtovar` |
| `tblArtikli` | 1 | `modAgrohemija` |
| `tblCenovnik` | 1 | `modCenovnik` |
| `tblKooperanti` | 1 | `modKooperant` |
| `tblPrevoznici` | 1 | `modUtovar` |
| `tblSEFConfig` | 1 | `modConfig` |
| `tblStornoVeze` | 1 | `modStornoContext` |
| `tblVozaci` | 1 | `modMalina` |
| `tblKulture` | 0 | _(samo testovi)_ |
| `tblKupci` | 0 | _(samo testovi)_ |
| `tblKutije` | 0 | _(samo testovi)_ |
| `tblStanice` | 0 | _(samo testovi)_ |
| `tblTipAmbalaze` | 0 | _(samo testovi)_ |

## Test moduli po tabeli

- `tblOtkup`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblFakture`: `modSEFTests`, `modTestBanka`, `modTestStorno`
- `tblNovac`: `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblAmbalaza`: `modBusinessFlowProTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblFakturaStavke`: `modTest`, `modTestStorno`
- `tblZbirna`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblOtpremnica`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPrijemnica`: `modFakturaTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblSEFEventLog`: `modSEFTests`
- `tblSEFSubmission`: `modSEFTests`
- `tblBankaImport`: `modTestBanka`, `modTestStorno`
- `tblPaleta`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblPaletaStavka`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblStornoZurnal`: `modTestStornoCentar`
- `tblMagacin`: `modAgrohemijaTests`, `modTest`
- `tblParcele`: `modAgrohemijaTests`
- `tblPartnerMap`: `modTestBanka`
- `tblArtikli`: `modAgrohemijaTests`
- `tblKooperanti`: `modAgrohemijaTests`, `modTestBanka`
- `tblSEFConfig`: `modTestStorno`
- `tblStornoVeze`: `modTest`, `modTestStorno`, `modTestStornoCentar`
- `tblKulture`: `modTestPalete`
- `tblKupci`: `modTestBanka`
- `tblKutije`: `modTest`
- `tblStanice`: `modTestBanka`
- `tblTipAmbalaze`: `modTestPalete`

## Sta ovo NE pokriva

- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad
  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --
  ako ga nadjes, to je nalaz, ne rupa u mapi.
- Granularnost je tabela, ne kolona.

