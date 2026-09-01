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
| `tblFakture` | 9 | `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator`, `modStorno`, `modStornoFlow` |
| `tblNovac` | 7 | `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modOtkup`, `modStorno`, `modStornoFlow` |
| `tblAmbalaza` | 6 | `modDokumenta`, `modMasterSync`, `modOtkup`, `modStorno`, `modStornoFlow`, `modStornoRecovery` |
| `tblZbirna` | 5 | `modDokumentInvariant`, `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblFakturaStavke` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow` |
| `tblOtpremnica` | 4 | `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblPrijemnica` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow` |
| `tblBankaImport` | 3 | `modBankaImport`, `modBankaMapiranje`, `modStorno` |
| `tblMagacin` | 3 | `frmAgrohemija`, `modAgroUnos`, `modAgrohemija` |
| `tblPaleta` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblPaletaStavka` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblPrerada` | 3 | `modFaktura`, `modPaletniList`, `modStorno` |
| `tblSEFEventLog` | 3 | `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblSEFSubmission` | 3 | `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblKorisnici` | 2 | `modAuth`, `modSetup` |
| `tblParcele` | 2 | `modGeoParcele`, `modMasterSync` |
| `tblPreradaStavka` | 2 | `modPaletniList`, `modStorno` |
| `tblStornoZurnal` | 2 | `modStorno`, `modStornoFlow` |
| `tblArtikli` | 1 | `modAgrohemija` |
| `tblKooperanti` | 1 | `modKooperant` |
| `tblPartnerMap` | 1 | `modBankaMapiranje` |
| `tblStornoVeze` | 1 | `modStornoContext` |
| `tblKulture` | 0 | _(samo testovi)_ |
| `tblKupci` | 0 | _(samo testovi)_ |
| `tblSEFConfig` | 0 | _(samo testovi)_ |
| `tblStanice` | 0 | _(samo testovi)_ |
| `tblTipAmbalaze` | 0 | _(samo testovi)_ |

## Test moduli po tabeli

- `tblOtkup`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblFakture`: `modSEFTests`, `modTestBanka`, `modTestStorno`
- `tblNovac`: `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblAmbalaza`: `modBusinessFlowProTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblZbirna`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblFakturaStavke`: `modTestStorno`
- `tblOtpremnica`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPrijemnica`: `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblBankaImport`: `modTestBanka`, `modTestStorno`
- `tblMagacin`: `modAgrohemijaTests`, `modTest`
- `tblPaleta`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblPaletaStavka`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblSEFEventLog`: `modSEFTests`
- `tblSEFSubmission`: `modSEFTests`
- `tblParcele`: `modAgrohemijaTests`
- `tblStornoZurnal`: `modTestStornoCentar`
- `tblArtikli`: `modAgrohemijaTests`
- `tblKooperanti`: `modAgrohemijaTests`, `modTestBanka`
- `tblPartnerMap`: `modTestBanka`
- `tblStornoVeze`: `modTest`, `modTestStorno`, `modTestStornoCentar`
- `tblKulture`: `modTestPalete`
- `tblKupci`: `modTestBanka`
- `tblSEFConfig`: `modTestStorno`
- `tblStanice`: `modTestBanka`
- `tblTipAmbalaze`: `modTestPalete`

## Sta ovo NE pokriva

- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad
  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --
  ako ga nadjes, to je nalaz, ne rupa u mapi.
- Granularnost je tabela, ne kolona.

