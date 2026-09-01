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
| `tblAmbalaza` | 7 | `modAmbalaza`, `modDokumenta`, `modMasterSync`, `modOtkup`, `modStorno`, `modStornoFlow`, `modStornoRecovery` |
| `tblNovac` | 7 | `modBankaMapiranje`, `modDokumenta`, `modFaktura`, `modNovac`, `modOtkup`, `modStorno`, `modStornoFlow` |
| `tblZbirna` | 5 | `modDokumentInvariant`, `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblFakturaStavke` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow` |
| `tblOtpremnica` | 4 | `modDokumenta`, `modMasterSync`, `modStorno`, `modStornoFlow` |
| `tblPrijemnica` | 4 | `modDokumenta`, `modFaktura`, `modStorno`, `modStornoFlow` |
| `tblSEFEventLog` | 4 | `modSEFPersistance`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblSEFSubmission` | 4 | `modSEFPersistance`, `modSEFService`, `modSEFStatusSync`, `modSEFValidator` |
| `tblBankaImport` | 3 | `modBankaImport`, `modBankaMapiranje`, `modStorno` |
| `tblMagacin` | 3 | `frmAgrohemija`, `modAgroUnos`, `modAgrohemija` |
| `tblPaleta` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblPaletaStavka` | 3 | `modDokumenta`, `modPaletniList`, `modStorno` |
| `tblStornoZurnal` | 3 | `modStorno`, `modStornoFlow`, `modStornoZurnal` |
| `tblKorisnici` | 2 | `modAuth`, `modSetup` |
| `tblParcele` | 2 | `modGeoParcele`, `modMasterSync` |
| `tblPartnerMap` | 2 | `modBankaMapiranje`, `modNovac` |
| `tblPrerada` | 2 | `modPaletniList`, `modStorno` |
| `tblPreradaStavka` | 2 | `modPaletniList`, `modStorno` |
| `tblArtikli` | 1 | `modAgrohemija` |
| `tblCenovnik` | 1 | `modCenovnik` |
| `tblKooperanti` | 1 | `modKooperant` |
| `tblSEFConfig` | 1 | `modConfig` |
| `tblStornoVeze` | 1 | `modStornoContext` |
| `tblVozaci` | 1 | `modMalina` |
| `tblKulture` | 0 | _(samo testovi)_ |
| `tblKupci` | 0 | _(samo testovi)_ |
| `tblStanice` | 0 | _(samo testovi)_ |
| `tblTipAmbalaze` | 0 | _(samo testovi)_ |

## Test moduli po tabeli

- `tblOtkup`: `modBusinessFlowProTests`, `modFakturaTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblFakture`: `modSEFTests`, `modTestBanka`, `modTestStorno`
- `tblAmbalaza`: `modBusinessFlowProTests`, `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblNovac`: `modTestBanka`, `modTestStorno`, `modTestStornoCentar`
- `tblZbirna`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblFakturaStavke`: `modTestStorno`
- `tblOtpremnica`: `modBusinessFlowProTests`, `modIzvestajTests`, `modTestStorno`, `modTestStornoCentar`
- `tblPrijemnica`: `modFakturaTests`, `modIzvestajTests`, `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblSEFEventLog`: `modSEFTests`
- `tblSEFSubmission`: `modSEFTests`
- `tblBankaImport`: `modTestBanka`, `modTestStorno`
- `tblMagacin`: `modAgrohemijaTests`, `modTest`
- `tblPaleta`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblPaletaStavka`: `modTestPalete`, `modTestStorno`, `modTestStornoCentar`
- `tblStornoZurnal`: `modTestStornoCentar`
- `tblParcele`: `modAgrohemijaTests`
- `tblPartnerMap`: `modTestBanka`
- `tblArtikli`: `modAgrohemijaTests`
- `tblKooperanti`: `modAgrohemijaTests`, `modTestBanka`
- `tblSEFConfig`: `modTestStorno`
- `tblStornoVeze`: `modTest`, `modTestStorno`, `modTestStornoCentar`
- `tblKulture`: `modTestPalete`
- `tblKupci`: `modTestBanka`
- `tblStanice`: `modTestBanka`
- `tblTipAmbalaze`: `modTestPalete`

## Sta ovo NE pokriva

- Upis mimo `AddTableSnapshot` i `modDataAccess` (direktan rad nad
  `ListObject`-om). Takav upis je van transakcije i van sloja podataka --
  ako ga nadjes, to je nalaz, ne rupa u mapi.
- Granularnost je tabela, ne kolona.

