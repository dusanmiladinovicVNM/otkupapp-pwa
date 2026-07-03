Attribute VB_Name = "modConfig"
'Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
' modConfig - Zentrale Konfiguration
' Alle Konstanten, Tabellennamen, Spaltenindizes
' KEINE Hardcoded Zellreferenzen im restlichen Code!
' ============================================================

' --- App Info ---
Public Const APP_NAME As String = "OtkupApp"
Public Const APP_VERSION As String = "2.11.0"

' --- Self-update / backup (Drive folder ID-jevi; popuni jednokratno) ---
' Vidi docs/RELEASE_PROCEDURE.md (self-update kanal) i modDrive/modRelease.
Public Const REL_FOLDER_ID As String = "1zL7ronXQUsOY56p7rULsqrM1u1U8sxod"      ' AgriX_Release: kod + version.json
Public Const BACKUP_FOLDER_ID As String = "199is7nQW3d4wfGX974AFTjpS4wo8itpl"   ' AgriX_Backup: xlsx data backup

' --- PDF dokument folderi (generisani PDF-ovi -> svoj podfolder pored sveske) ---
Public Const PDF_DIR_OTKUPNI As String = "Otkupni listovi"
Public Const PDF_DIR_PRIJEMNICE As String = "Prijemnice"
Public Const PDF_DIR_OTPREMNICE As String = "Otpremnice"
Public Const PDF_DIR_REVERS As String = "Revers ambalaze"
Public Const PDF_DIR_KARTICE As String = "Kartice kooperanata"
Public Const PDF_DIR_PALETNI As String = "Paletni listovi"
Public Const PDF_DIR_PRERADA As String = "Preradni listovi"
Public Const PDF_DIR_SPECIFIKACIJE As String = "Specifikacije"
Public Const PDF_DIR_IZVESTAJI As String = "Izvestaji"
' CSV nalozi za prenos (banka export) - isti root kao PDF folderi (EnsureDocFolder)
Public Const CSV_DIR_BANKA_NALOZI As String = "Nalozi za banku"

' --- Tabellennamen (ListObjects) ---
Public Const TBL_KOOPERANTI As String = "tblKooperanti"
Public Const TBL_STANICE As String = "tblStanice"
Public Const TBL_VOZACI As String = "tblVozaci"
Public Const TBL_KUPCI As String = "tblKupci"
Public Const TBL_KULTURE As String = "tblKulture"
Public Const TBL_OTKUP As String = "tblOtkup"
Public Const TBL_OTPREMNICA As String = "tblOtpremnica"
Public Const TBL_ZBIRNA As String = "tblZbirna"
Public Const TBL_PRIJEMNICA As String = "tblPrijemnica"
Public Const TBL_FAKTURE As String = "tblFakture"
Public Const TBL_FAKTURA_STAVKE As String = "tblFakturaStavke"
Public Const TBL_NOVAC As String = "tblNovac"
Public Const TBL_AMBALAZA As String = "tblAmbalaza"
Public Const TBL_CONFIG As String = "tblConfig"
Public Const TBL_PARTNER_MAP As String = "tblPartnerMap"
Public Const TBL_PARCELE As String = "tblParcele"
Public Const TBL_ARTIKLI As String = "tblArtikli"
Public Const TBL_MAGACIN As String = "tblMagacin"
Public Const TBL_BANKA_IMPORT As String = "tblBankaImport"
Public Const TBL_SEF_SUBMISSION As String = "tblSEFSubmission"
Public Const TBL_SEF_EVENT_LOG As String = "tblSEFEventLog"
Public Const TBL_SEF_CONFIG As String = "tblSEFConfig"
Public Const TBL_KORISNICI As String = "tblKorisnici"

' --- Phase 2 Tabellen ---
Public Const TBL_PROIZVODJACI As String = "tblProizvodjaci"
Public Const TBL_HLADNJACA As String = "tblHladnjaca"
Public Const TBL_LAGER As String = "tblLager"
Public Const TBL_PRERADA As String = "tblPrerada"

' --- Paletni list (Phase 2 / hladnjaca) ---
Public Const TBL_PALETA As String = "tblPaleta"
Public Const TBL_PALETA_STAVKA As String = "tblPaletaStavka"
Public Const TBL_PRERADA_STAVKA As String = "tblPreradaStavka"
Public Const TBL_TIP_AMBALAZE As String = "tblTipAmbalaze"
Public Const TBL_TIP_PALETE As String = "tblTipPalete"
Public Const TBL_KUTIJE As String = "tblKutije"
Public Const TBL_KESE As String = "tblKese"
Public Const TBL_VRSTA_GP As String = "tblVrstaGotovihProizvoda"
Public Const TBL_KVALITET As String = "tblKvalitet"
Public Const TBL_SEF As String = "tblSEF"
Public Const TBL_SLEDLJIVOST As String = "tblSledljivost"
Public Const TBL_METEO As String = "tblMeteo"

' --- Report Tabellen ---
Public Const TBL_RPT_SALDO_OM As String = "tblRptSaldoOM"
Public Const TBL_RPT_SALDO_KUPCI As String = "tblRptSaldoKupci"
Public Const TBL_RPT_MARZA As String = "tblRptMarza"
Public Const TBL_RPT_VOZACI As String = "tblRptVozaci"
Public Const TBL_RPT_ZBIRNI As String = "tblRptZbirni"

' --- Poruke (resource table) ---
Public Const TBL_PORUKE   As String = "tblPoruke"
Public Const COL_POR_KLJUC As String = "Kljuc"
Public Const COL_POR_TEKST  As String = "Tekst"
Public Const PORUKE_SHEET  As String = "sPoruke"

' --- Sheet-Codenamen (robust gegen Umbenennung) ---
Public Const SHT_CONFIG As String = "sConfig"
Public Const SHT_KOOPERANTI As String = "sKooperanti"
Public Const SHT_STANICE As String = "sStanice"
Public Const SHT_OTKUP As String = "sOtkup"
Public Const SHT_KUPCI As String = "sKupci"
Public Const SHT_FAKTURE As String = "sFakture"
Public Const SHT_NOVAC As String = "sNovac"

' --- Spaltennamen tblKooperanti ---
Public Const COL_KOOP_BPG As String = "BPGBroj"
Public Const COL_KOOP_STANICA As String = "StanicaID"
Public Const COL_KOOP_TEKUCI_RACUN As String = "TekuciRacun"
Public Const COL_KOOP_ID As String = "KooperantID"

' --- Spaltennamen tblKooperanti ---
Public Const COL_KUP_TEKUCI_RACUN As String = "TekuciRacun"
Public Const COL_KUP_ID As String = "KupacID"
Public Const COL_KUP_NAZIV As String = "Naziv"
' --- Spaltennamen tblParcele ---
Public Const COL_PAR_ID As String = "ParcelaID"
Public Const COL_PAR_KOOP As String = "KooperantID"
Public Const COL_PAR_KAT_BROJ As String = "KatBroj"
Public Const COL_PAR_KAT_OPSTINA As String = "KatOpstina"
Public Const COL_PAR_KULTURA As String = "Kultura"
Public Const COL_PAR_POVRSINA As String = "PovrsinaHa"
Public Const COL_PAR_GGAP As String = "GGAPStatus"
    ' --- Spaltennamen tblParcele - Geo ---
Public Const COL_PAR_AKTIVNA As String = "Aktivna"
Public Const COL_PAR_GEO_STATUS As String = "GeoStatus"
Public Const COL_PAR_GEO_SOURCE As String = "GeoSource"
Public Const COL_PAR_N As String = "N_Coord"
Public Const COL_PAR_E As String = "E_Coord"
Public Const COL_PAR_LAT As String = "Lat"
Public Const COL_PAR_LNG As String = "Lng"
Public Const COL_PAR_POLYGON As String = "PolygonGeoJSON"
Public Const COL_PAR_METEO As String = "MeteoEnabled"
Public Const COL_PAR_RIZIK As String = "RizikStatus"
Public Const COL_PAR_DATUM_GEO As String = "DatumGeoUnosa"
Public Const COL_PAR_DATUM_AZUR As String = "DatumAzuriranja"
Public Const COL_PAR_NAPOMENA As String = "Napomena"

' --- Spaltennamen tblArtikli ---
Public Const COL_ART_ID As String = "ArtikalID"
Public Const COL_ART_NAZIV As String = "Naziv"
Public Const COL_ART_TIP As String = "Tip"
Public Const COL_ART_JM As String = "JedinicaMere"
Public Const COL_ART_CENA As String = "CenaPoJedinici"
Public Const COL_ART_DOZA As String = "DozaPoHa"
Public Const COL_ART_KULTURA As String = "Kultura"
Public Const COL_ART_PAKOVANJE As String = "Pakovanje"

' Rezervisani virtuelni artikal za knjizenje POCETNOG DUGA kooperanta
' (migracija). Operater ga NE bira -- upisuje ga BookPocetniDug kao magacin
' IZLAZ stavku (Vrednost = iznos duga). Skriven iz combo lista i iz stanja
' magacina (vidi LoadArtikli/LoadArtikliUlaz i GetMagacinStanje).
Public Const ART_POCETNI_DUG As String = "ART-POC-DUG"

' --- Spaltennamen tblMagacin ---
Public Const COL_MAG_ID As String = "MagacinID"
Public Const COL_MAG_DATUM As String = "Datum"
Public Const COL_MAG_ARTIKAL As String = "ArtikalID"
Public Const COL_MAG_TIP As String = "Tip"
Public Const COL_MAG_KOLICINA As String = "Kolicina"
Public Const COL_MAG_KOOP As String = "KooperantID"
Public Const COL_MAG_PARCELA As String = "ParcelaID"
Public Const COL_MAG_BR_DOK As String = "BrojDokumenta"
Public Const COL_MAG_CENA As String = "CenaPoJedinici"
Public Const COL_MAG_VREDNOST As String = "Vrednost"
Public Const COL_MAG_NAPOMENA As String = "Napomena"
Public Const COL_MAG_STORNIRANO As String = "Stornirano"
Public Const COL_MAG_DOBAVLJAC As String = "DobavljacID"

Public Const MAG_ULAZ As String = "Ulaz"
Public Const MAG_IZLAZ As String = "Izlaz"

' --- Spaltennamen tblOtkup ---
Public Const COL_OTK_ID As String = "OtkupID"
Public Const COL_OTK_DATUM As String = "Datum"
Public Const COL_OTK_KOOPERANT As String = "KooperantID"
Public Const COL_OTK_STANICA As String = "StanicaID"
Public Const COL_OTK_KULTURA As String = "KulturaID"
Public Const COL_OTK_VRSTA As String = "VrstaVoca"
Public Const COL_OTK_SORTA As String = "SortaVoca"
Public Const COL_OTK_KOLICINA As String = "Kolicina"
Public Const COL_OTK_CENA As String = "Cena"
Public Const COL_OTK_TIP_AMB As String = "TipAmbalaze"
Public Const COL_OTK_KOL_AMB As String = "KolAmbalaze"
Public Const COL_OTK_KOL_AMB_IZDATA As String = "KolAmbIzdata"   ' OM izdao prazne kooperantu uz otkup (OM-Izlaz-Koop)
Public Const COL_OTK_VREME_UNOSA As String = "VremeUnosa"        ' timestamp snimanja otkupa (Now() pri upisu)
Public Const COL_OTK_BRUTO As String = "BrutoKg"                 ' bruto tezina uneta na otkupu (kad je OTKUP_BRUTO_UNOS); prazno = unet neto
Public Const COL_OTK_VOZAC As String = "VozacID"
Public Const COL_OTK_BR_DOK As String = "BrojDokumenta"
Public Const COL_OTK_NOVAC As String = "Novac"
Public Const COL_OTK_PRIMALAC As String = "PrimalacNovca"
Public Const COL_OTK_KLASA As String = "Klasa"
Public Const COL_OTK_STORNIRANO As String = "Stornirano"
Public Const COL_OTK_BROJ_ZBIRNE As String = "BrojZbirne"
Public Const COL_OTK_ISPLACENO As String = "Isplaceno"
Public Const COL_OTK_DATUM_ISPLATE As String = "DatumIsplate"
Public Const COL_OTK_OTPREMNICA_ID As String = "OtpremnicaID"
Public Const COL_OTK_PARCELA As String = "ParcelaID"

' --- Spaltennamen tblOtpremnica (NEU) ---
Public Const COL_OTP_ID As String = "OtpremnicaID"
Public Const COL_OTP_DATUM As String = "Datum"
Public Const COL_OTP_STANICA As String = "StanicaID"
Public Const COL_OTP_VOZAC As String = "VozacID"
Public Const COL_OTP_BROJ As String = "BrojOtpremnice"
Public Const COL_OTP_BROJ_ZBIRNE As String = "BrojZbirne"
Public Const COL_OTP_VRSTA As String = "VrstaVoca"
Public Const COL_OTP_SORTA As String = "SortaVoca"
Public Const COL_OTP_KOLICINA As String = "Kolicina"
Public Const COL_OTP_CENA As String = "Cena"
Public Const COL_OTP_TIP_AMB As String = "TipAmbalaze"
Public Const COL_OTP_KOL_AMB As String = "KolAmbalaze"
Public Const COL_OTP_KLASA As String = "Klasa"
Public Const COL_OTP_BRUTO As String = "BrutoKg"                 ' bruto tezina (kad je OTKUP_BRUTO_UNOS); prazno = neto

Public Const DOK_TIP_OM_ULAZ As String = "OMUlaz"

' --- Spaltennamen tblZbirna (NEU) ---
Public Const COL_ZBR_ID As String = "ZbirnaID"
Public Const COL_ZBR_DATUM As String = "Datum"
Public Const COL_ZBR_VOZAC As String = "VozacID"
Public Const COL_ZBR_BROJ As String = "BrojZbirne"
Public Const COL_ZBR_KUPAC As String = "KupacID"
Public Const COL_ZBR_HLADNJACA As String = "Hladnjaca"
Public Const COL_ZBR_POGON As String = "Pogon"
Public Const COL_ZBR_KOLICINA As String = "UkupnoKolicina"
Public Const COL_ZBR_TIP_AMB As String = "TipAmbalaze"
Public Const COL_ZBR_KOL_AMB As String = "UkupnoAmbalaze"
Public Const COL_ZBR_VRSTA As String = "VrstaVoca"
Public Const COL_ZBR_SORTA As String = "SortaVoca"
Public Const COL_ZBR_KLASA As String = "Klasa"

' --- Spaltennamen tblPrijemnica (NEU) ---
Public Const COL_PRJ_ID As String = "PrijemnicaID"
Public Const COL_PRJ_DATUM As String = "Datum"
Public Const COL_PRJ_KUPAC As String = "KupacID"
Public Const COL_PRJ_BROJ As String = "BrojPrijemnice"
Public Const COL_PRJ_BROJ_ZBIRNE As String = "BrojZbirne"
Public Const COL_PRJ_KOLICINA As String = "Kolicina"
Public Const COL_PRJ_CENA As String = "Cena"
Public Const COL_PRJ_TIP_AMB As String = "TipAmbalaze"
Public Const COL_PRJ_KOL_AMB As String = "KolAmbalaze"
Public Const COL_PRJ_KOL_AMB_VRACENA As String = "KolAmbVracena"
Public Const COL_PRJ_BRUTO As String = "BrutoKg"                 ' bruto tezina (preneto iz otkupa kad je OTKUP_BRUTO_UNOS); prazno = neto
Public Const COL_PRJ_VOZAC As String = "VozacID"
Public Const COL_PRJ_VRSTA As String = "VrstaVoca"
Public Const COL_PRJ_SORTA As String = "SortaVoca"
Public Const COL_PRJ_KLASA As String = "Klasa"
Public Const COL_PRJ_FAKTURISANO As String = "Fakturisano"
Public Const COL_PRJ_FAKTURA_ID As String = "FakturaID"

' ============================================================
' Paletni list (Phase 2) -- kolone novih tabela
' ============================================================
' tblKulture extension
Public Const COL_KUL_GAJBICA_PALETA As String = "GajbicaPoPaleti"

' tblPaleta (header)
Public Const COL_PAL_ID As String = "PaletaID"
Public Const COL_PAL_BROJ As String = "BrojPalete"
Public Const COL_PAL_GODINA As String = "Godina"
Public Const COL_PAL_DATUM As String = "Datum"
Public Const COL_PAL_VRSTA As String = "VrstaVoca"
Public Const COL_PAL_TIP_PALETE As String = "TipPalete"
Public Const COL_PAL_BR_GAJBICA As String = "BrojGajbica"
Public Const COL_PAL_NETO As String = "NetoKg"
Public Const COL_PAL_AMBALAZA As String = "AmbalazaKg"
Public Const COL_PAL_PALETA_KG As String = "PaletaKg"
Public Const COL_PAL_BRUTO As String = "BrutoKg"
Public Const COL_PAL_STATUS As String = "Status"
Public Const COL_PAL_PRERADJENO As String = "Preradjeno"
Public Const COL_PAL_SORTA As String = "SortaVoca"
Public Const COL_PAL_KLASA As String = "Klasa"
Public Const COL_PAL_TIP_AMBALAZE As String = "TipAmbalaze"
Public Const COL_PAL_KAPACITET As String = "KapacitetGajbica"
Public Const COL_PAL_CREATED As String = "CreatedAt"

' tblPaletaStavka (veza paleta -> prijemnica/zbirna)
Public Const COL_PALS_ID As String = "StavkaID"
Public Const COL_PALS_PALETA_ID As String = "PaletaID"
Public Const COL_PALS_BROJ_PRIJ As String = "BrojPrijemnice"
Public Const COL_PALS_BROJ_ZBIRNE As String = "BrojZbirne"
Public Const COL_PALS_BR_GAJBICA As String = "BrojGajbica"
Public Const COL_PALS_NETO As String = "NetoKg"
Public Const COL_PALS_AMBALAZA As String = "AmbalazaKg"
Public Const COL_PALS_PRIJEMNICA_ID As String = "PrijemnicaID"
Public Const COL_PALS_KLASA As String = "Klasa"
Public Const COL_PALS_VRSTA As String = "VrstaVoca"
Public Const COL_PALS_SORTA As String = "SortaVoca"
Public Const COL_PALS_CREATED As String = "CreatedAt"

' tblTipAmbalaze / tblTipPalete (sifarnici tezina)
Public Const COL_TAMB_TIP As String = "TipAmbalaze"
Public Const COL_TAMB_TEZINA As String = "TezinaGajbiceKg"
Public Const COL_TPAL_TIP As String = "TipPalete"
Public Const COL_TPAL_TEZINA As String = "TezinaKg"

' tblKutije / tblKese (sifarnici tezina) + tblVrstaGotovihProizvoda
Public Const COL_KUT_TIP As String = "TipKutije"
Public Const COL_KUT_TEZINA As String = "TezinaKg"
Public Const COL_KES_TIP As String = "TipKese"
Public Const COL_KES_TEZINA As String = "TezinaKg"
Public Const COL_VGP_TIP As String = "TipGotovogProizvoda"

' tblCenovnik (cene po proizvodu -- append-only istorija)
' Kljuc: VrstaVoca + SortaVoca + Klasa. Poslednji red (najnoviji Datum)
' je vazeci za otkup i dokumenta. Stari redovi ostaju radi kretanja cena.
Public Const TBL_CENOVNIK As String = "tblCenovnik"
Public Const COL_CEN_ID As String = "CenaID"
Public Const COL_CEN_DATUM As String = "Datum"
Public Const COL_CEN_VRSTA As String = "VrstaVoca"
Public Const COL_CEN_SORTA As String = "SortaVoca"
Public Const COL_CEN_KLASA As String = "Klasa"
Public Const COL_CEN_CENA As String = "Cena"
Public Const COL_CEN_CREATED As String = "CreatedAt"

' tblPrerada / tblPreradaStavka
Public Const COL_PRE_ID As String = "PreradaID"
Public Const COL_PRE_DATUM As String = "Datum"
Public Const COL_PRE_NETO As String = "NetoKolicina"
Public Const COL_PRE_KUTIJE As String = "BrojKutija"
Public Const COL_PRE_KESE As String = "BrojKesa"
Public Const COL_PRES_ID As String = "StavkaID"
Public Const COL_PRES_PRERADA_ID As String = "PreradaID"
Public Const COL_PRES_BROJ_PALETE As String = "BrojPalete"
Public Const COL_PRE_BROJ As String = "BrojPrerade"
Public Const COL_PRE_GODINA As String = "Godina"
Public Const COL_PRES_PALETA_ID As String = "PaletaID"
Public Const COL_PRES_NETO As String = "NetoKg"
Public Const COL_PRE_NETO_ULAZ As String = "NetoUlazKg"
Public Const COL_PRE_NETO_IZLAZ As String = "NetoIzlazKg"
Public Const COL_PRE_NAPOMENA As String = "Napomena"
Public Const COL_PRE_CREATED As String = "CreatedAt"
Public Const COL_PRES_CREATED As String = "CreatedAt"

' tblPrerada (Paletni list gotovih proizvoda) - tezine i izbori
Public Const COL_PRE_TEZINA_PALETE As String = "TezinaPaleteKg"
Public Const COL_PRE_BRUTO As String = "BrutoKg"
Public Const COL_PRE_AMBALAZA As String = "AmbalazaKg"
Public Const COL_PRE_TIP_KUTIJE As String = "TipKutije"
Public Const COL_PRE_TIP_KESE As String = "TipKese"
Public Const COL_PRE_TIP_GP As String = "TipGotovogProizvoda"

' Paleta status
Public Const PAL_STATUS_OTVORENA As String = "Otvorena"
Public Const PAL_STATUS_ZATVORENA As String = "Zatvorena"

' Paleta -- config kljuc (default tip palete) + fallback kapacitet
Public Const CFG_DEFAULT_TIP_PALETE As String = "DEFAULT_TIP_PALETE"
Public Const PALETA_DEFAULT_KAPACITET As Long = 240
Public Const CFG_PALETA_PRINT_MODE As String = "PALETA_PRINT_MODE"
' Paletiranje (izrada paletnih listova iz prijemnice). Default ON (kao do sada).
' NO -> PaletizePrijemnica se preskace; nema paleta ni paletnih listova.
Public Const CFG_PALETIRANJE As String = "PALETIRANJE"
Public Const CFG_OTKUP_PRINT_MODE As String = "OTKUP_PRINT_MODE"
' Stampa prijemnice -- auto izlaz na "Unos prijemnice". Default OFF (kao do sada;
' prijemnica se ne stampa). PDF | PRINT | PREVIEW | OFF (kao PALETA_PRINT_MODE).
Public Const CFG_PRIJEMNICA_PRINT_MODE As String = "PRIJEMNICA_PRINT_MODE"
' Izdavanje prazne ambalaze kooperantu (revers) - auto izlaz na "Unesi" u OM-Ulaz
' frame-u kad je toggle "Izdavanje kooperantu" aktivan. Default PDF (otvori PDF,
' kao OTKUP_PRINT_MODE). PDF | PRINT | PREVIEW | OFF (kao PALETA_PRINT_MODE).
Public Const CFG_OM_IZDAVANJE_PRINT_MODE As String = "OM_IZDAVANJE_PRINT_MODE"
' Rucni obrasci (faktura/kartica/sledljivost) - konfigurabilan izlaz (Faza 2B).
' PDF | PRINT | PREVIEW | OFF; prazno -> default tog dokumenta (DocResolveMode).
Public Const CFG_FAKTURA_PRINT_MODE As String = "FAKTURA_PRINT_MODE"
Public Const CFG_KARTICA_PRINT_MODE As String = "KARTICA_PRINT_MODE"
Public Const CFG_SLEDLJIVOST_PRINT_MODE As String = "SLEDLJIVOST_PRINT_MODE"
Public Const CFG_KARTICA_AMB_PRINT_MODE As String = "KARTICA_AMB_PRINT_MODE"
Public Const CFG_SPECIFIKACIJA_PRINT_MODE As String = "SPECIFIKACIJA_PRINT_MODE"
' Specifikacija isplata (banka nalozi, frmBankaExportPregled). Default PDF.
Public Const CFG_ISPLATA_SPEC_PRINT_MODE As String = "ISPLATA_SPEC_PRINT_MODE"
' Otpremnica / grupni otkupni list / paletni list got. proizvoda (prerada).
' PDF | PRINT | PREVIEW | OFF. Grupni: prazno -> prati CFG_OTKUP_PRINT_MODE.
Public Const CFG_OTPREMNICA_PRINT_MODE As String = "OTPREMNICA_PRINT_MODE"
Public Const CFG_GRUPNI_OTKUP_PRINT_MODE As String = "GRUPNI_OTKUP_PRINT_MODE"
Public Const CFG_PRERADA_PRINT_MODE As String = "PRERADA_PRINT_MODE"
Public Const CFG_PDV_NADOKNADA_STOPA As String = "PDV_NADOKNADA_STOPA"
Public Const PDV_NADOKNADA_DEFAULT As Double = 8

' Otkupni list - obavezni elementi otkupnog bloka (klauzula + rok isplate)
Public Const CFG_OTKUP_KLAUZULA As String = "OTKUP_KLAUZULA"
Public Const CFG_OTKUP_ROK As String = "OTKUP_ROK_ISPLATE"
Public Const OTKUP_ROK_DEFAULT As String = "Po dogovoru"

' Banka - nalozi za prenos (CSV export iz frmBankaExportPregled za uvoz u
' e-banking). Platilac = SELLER_NAME / SELLER_ACCOUNT (grupa "Prodavac (firma)").
' Sifra placanja: NBS sifarnik, default 221 (bezgotovinski promet robe i usluga).
' Svrha: osnovni tekst, po nalogu se dopisuje broj otkupnog bloka; poziv na broj
' odobrenja = broj bloka (jaki kljuc za auto-map pri uvozu izvoda, vidi
' docs/production-runbook-banka-import-setup.md).
Public Const CFG_BANKA_NALOG_SIFRA As String = "BANKA_NALOG_SIFRA_PLACANJA"
Public Const BANKA_NALOG_SIFRA_DEFAULT As String = "221"
Public Const CFG_BANKA_NALOG_SVRHA As String = "BANKA_NALOG_SVRHA"
Public Const BANKA_NALOG_SVRHA_DEFAULT As String = "Otkup poljoprivrednih proizvoda"
' Racuni firme za isplate. Firma moze imati vise racuna (razne banke); operater
' bira racun u frmBankaExportPregled ("Sa racuna"). Izvor istine su tri zasebna
' polja u Podesavanjima (grupa "Banka / nalozi"): CFG_BANKA_NALOG_RACUN_1/2/3.
' BankaNalogRacuniCSV ih spaja u ";"-listu (preskacuci prazna).
Public Const CFG_BANKA_NALOG_RACUN_1 As String = "BANKA_NALOG_RACUN_1"
Public Const CFG_BANKA_NALOG_RACUN_2 As String = "BANKA_NALOG_RACUN_2"
Public Const CFG_BANKA_NALOG_RACUN_3 As String = "BANKA_NALOG_RACUN_3"
' Legacy: stari jedinstveni ";"-spisak racuna (pre razdvajanja na 3 polja).
' Zadrzan zbog kompatibilnosti -- fallback u BankaNalogRacuniCSV + jednokratna
' migracija (modPodesavanja.MigrateBankaRacuniLegacy). Ne prikazuje se u editoru.
Public Const CFG_BANKA_NALOG_RACUNI As String = "BANKA_NALOG_RACUNI"

' --- Imena sablon sheet-ova (template worksheets za stampu) ---
' Jedinstveni izvor istine za nazive sheet-ova-sablona (umesto rasutih
' hardkodovanih stringova: modPrint/modPaletniList/modFaktura/modIzvestaj/
' modOtkupBlok/frmSledljivost).
Public Const WS_OTKUP_SABLON As String = "OtkupSablon"
Public Const WS_GRUPNI_OTKUP_SABLON As String = "GrupniOtkupSablon"
Public Const WS_OTPREMNICA_SABLON As String = "OtpremnicaSablon"
Public Const WS_PRIJEMNICA_SABLON As String = "PrijemnicaSablon"
Public Const WS_IZDAMB_SABLON As String = "IzdAmbSablon"
Public Const WS_PALETA_SABLON As String = "PaletaSablon"
Public Const WS_PRERADA_SABLON As String = "PreradaSablon"
Public Const WS_FAKTURA_SABLON As String = "FakturaSablon"
Public Const WS_KARTICA_SABLON As String = "KarticaSablon"
Public Const WS_KARTICA_AMB_SABLON As String = "KarticaAmbalazeSablon"
Public Const WS_SLEDLJIVOST_SABLON As String = "SledljivostSablon"
Public Const WS_SPECIFIKACIJA_SABLON As String = "SpecifikacijaSablon"
Public Const WS_ISPLATA_SPEC_SABLON As String = "IsplataSpecSablon"

' --- Dokument-Tipovi ---
Public Const COL_STORNIRANO As String = "Stornirano"
Public Const COL_OSIROCENO_OD As String = "OsirocenoOd"
Public Const DOK_TIP_OTKUP As String = "Otkup"
Public Const DOK_TIP_OTPREMNICA As String = "Otpremnica"
Public Const DOK_TIP_PRIJEMNICA As String = "Prijemnica"
Public Const DOK_TIP_IZLAZ_KUPCI As String = "Kupci-Otpremnica"
Public Const DOK_TIP_OM_IZLAZ_KOOP As String = "OM-Izlaz-Koop"  ' OM izdaje (praznu) ambalazu kooperantu
Public Const DOK_TIP_OM_ULAZ_KOOP As String = "OM-Ulaz-Koop"    ' kooperant vraca (praznu) ambalazu na OM (povrat)
Public Const DOK_TIP_OM_IZLAZ_FIRMA As String = "OM-Izlaz-Firma" ' OM vraca (praznu) ambalazu firmi (centrala)
Public Const DOK_TIP_OM_ULAZ_FIRMA As String = "OM-Ulaz-Firma"   ' firma (centrala) salje (praznu) ambalazu na OM

' --- Spaltennamen tblFakture ---
Public Const COL_FAK_ID As String = "FakturaID"
Public Const COL_FAK_BROJ As String = "BrojFakture"
Public Const COL_FAK_DATUM As String = "Datum"
Public Const COL_FAK_KUPAC As String = "KupacID"
Public Const COL_FAK_IZNOS As String = "Iznos"
Public Const COL_FAK_STATUS As String = "Status"
Public Const COL_FAK_DATUM_PLACANJA As String = "DatumPlacanja"

' --- Spaltennamen tblAmbalaza ---
Public Const COL_AMB_ID As String = "AmbID"
Public Const COL_AMB_DATUM As String = "Datum"
Public Const COL_AMB_TIP As String = "TipAmbalaze"
Public Const COL_AMB_KOLICINA As String = "Kolicina"
Public Const COL_AMB_SMER As String = "Smer"
Public Const COL_AMB_ENTITET As String = "EntitetID"
Public Const COL_AMB_ENTITET_TIP As String = "EntitetTip"
Public Const COL_AMB_VOZAC As String = "VozacID"
Public Const COL_AMB_DOK_ID As String = "DokumentID"
Public Const COL_AMB_DOK_TIP As String = "DokumentTip"

' --- tblNovac ---
Public Const COL_NOV_ID As String = "NovacID"
Public Const COL_NOV_BROJ_DOK As String = "BrojDokumenta"
Public Const COL_NOV_DATUM As String = "Datum"
Public Const COL_NOV_PARTNER As String = "Partner"
Public Const COL_NOV_PARTNER_ID As String = "PartnerID"
Public Const COL_NOV_ENTITET_TIP As String = "EntitetTip"
Public Const COL_NOV_OM_ID As String = "OMID"
Public Const COL_NOV_KOOP_ID As String = "KooperantID"
Public Const COL_NOV_FAKTURA_ID As String = "FakturaID"
Public Const COL_NOV_VRSTA As String = "VrstaVoca"
Public Const COL_NOV_TIP As String = "Tip"
Public Const COL_NOV_UPLATA As String = "Uplata"
Public Const COL_NOV_ISPLATA As String = "Isplata"
Public Const COL_NOV_NAPOMENA As String = "Napomena"
Public Const COL_NOV_OTKUP_ID As String = "OtkupID"


' --- Novac Tipovi ---
Public Const NOV_BANKA_UPLATA As String = "BankaUplata"
Public Const NOV_BANKA_ISPLATA As String = "BankaIsplata"
Public Const NOV_KUPCI_UPLATA As String = "KupciUplata"
Public Const NOV_KUPCI_AVANS As String = "KupciAvans"
Public Const NOV_KES_FIRMA_OTKUPAC As String = "KesFirmaOtkupac"
Public Const NOV_KES_OTKUPAC_KOOP As String = "KesOtkupacKoop"
Public Const NOV_VIRMAN_FIRMA_KOOP As String = "VirmanFirmaKoop"
Public Const NOV_VIRMAN_AVANS_KOOP As String = "VirmanAvansKoop"

' --- Banka Import ---
Public Const PREFIX_BANKA_IMPORT As String = "BIM-"
Public Const COL_BIM_ID As String = "BankaImportID"
Public Const COL_BIM_BROJ_DOKUMENTA As String = "BrojDokumenta"
Public Const COL_BIM_DATUM_IZVODA As String = "DatumIzvoda"
Public Const COL_BIM_BROJ_RACUNA As String = "BrojRacuna"
Public Const COL_BIM_DATUM_TRANSAKCIJE As String = "DatumTransakcije"
Public Const COL_BIM_PARTNER As String = "Partner"
Public Const COL_BIM_PARTNER_KONTO As String = "PartnerKonto"
Public Const COL_BIM_OPIS As String = "Opis"
Public Const COL_BIM_UPLATA As String = "Uplata"
Public Const COL_BIM_ISPLATA As String = "Isplata"
Public Const COL_BIM_VALUTA As String = "Valuta"
Public Const COL_BIM_POZIV_NA_BROJ As String = "PozivNaBroj"
Public Const COL_BIM_SVRHA_PLACANJA As String = "SvrhaPlacanja"
Public Const COL_BIM_BANKA_REFERENZ As String = "BankaReferenz"
Public Const COL_BIM_IZVOR_FAJL As String = "IzvorFajl"
Public Const COL_BIM_IMPORT_VREME As String = "ImportVreme"
Public Const COL_BIM_OBRADJENO As String = "Obradjeno"
Public Const COL_BIM_STORNIRANO As String = "Stornirano"
Public Const COL_BIM_POCETNO_STANJE As String = "PocetnoStanje"
Public Const COL_BIM_ZAVRSNO_STANJE As String = "ZavrsnoStanje"
Public Const COL_BIM_UKUPAN_DUGUJE As String = "UkupanDuguje"
Public Const COL_BIM_UKUPAN_POTRAZUJE As String = "UkupanPotrazuje"

Public Const CONFIG_KEY_PDFTOTEXT_EXE_PATH As String = "PDFTOTEXT_EXE_PATH"
Public Const APP_PDFTOTEXT_RELATIVE_EXE_PATH As String = "Tools\poppler\Library\bin\pdftotext.exe"
Public Const APP_BANKA_INBOX As String = "G:\My Drive\Bank_Izvodi\Inbox"
Public Const APP_BANKA_PROCESSED As String = "G:\My Drive\Bank_Izvodi\Processed"
Public Const APP_BANKA_ERROR As String = "G:\My Drive\Bank_Izvodi\Error"

' --- tblPartnerMap ---
Public Const COL_PM_BANKA_NAME As String = "BankaName"
Public Const COL_PM_PARTNER_ID As String = "PartnerID"
Public Const COL_PM_ENTITET_TIP As String = "EntitetTip"
Public Const COL_PM_OM_ID As String = "OMID"

' --- tblFakturaStavke Spalten ---
Public Const COL_FS_ID As String = "StavkaID"
Public Const COL_FS_FAKTURA_ID As String = "FakturaID"
Public Const COL_FS_PRIJEMNICA_ID As String = "PrijemnicaID"
Public Const COL_FS_KOLICINA As String = "Kolicina"
Public Const COL_FS_CENA As String = "Cena"
Public Const COL_FS_KLASA As String = "Klasa"
Public Const COL_FS_BROJ_PRIJEMNICE As String = "BrojPrijemnice"

' --- Typen Ambalaze ---
Public Const AMB_12_1 As String = "12/1"
Public Const AMB_6_1 As String = "6/1"

' --- Klasa ---
Public Const KLASA_I As String = "I"
Public Const KLASA_II As String = "II"

' --- Status ---
Public Const STATUS_AKTIVAN As String = "Aktivan"
Public Const STATUS_NEAKTIVAN As String = "Neaktivan"
Public Const STATUS_FAKTURISANO As String = "Fakturisano"
Public Const STATUS_NEFAKTURISANO As String = "Nefakturisano"
Public Const STATUS_PLACENO As String = "Placeno"
Public Const STATUS_NEPLACENO As String = "Neplaceno"
Public Const STATUS_ISPLACENO As String = "Da"

' Cloud/PWA sync toggle. Default: TRUE (cloud enabled).
' Desktop-only deploy postavlja CLOUD_SYNC_ENABLED = "NO" u Config sheet-u.
' Vrednost u Config sheet-u: "YES" | "NO" | "TRUE" | "FALSE" | "1" | "0".
' Empty / missing key tretira se kao "YES" (backward compatible).
Public Const CFG_KEY_CLOUD_SYNC_ENABLED As String = "CLOUD_SYNC_ENABLED"

' Malina mod (1 stanica = 1 vozilo): otpremnica == zbirna, auto-zbirna iz otpremnice.
' Default: FALSE (visnja klijenti rade nepromenjeno).
' Vrednost u Config sheet-u: "YES" | "TRUE" | "1" | "ON" | "ENABLED" za ukljuceno.
Public Const CFG_KEY_MALINA_MODE As String = "MALINA_MODE"
' Default kupac (Hladnjaca) za auto-zbirnu u malina modu. KupacID iz tblKupci.
Public Const CFG_MALINA_DEFAULT_KUPAC As String = "MALINA_DEFAULT_KUPAC"

' --- Dorade: podesavanja (UI: Maticni podaci -> Podesavanja) ---
' Default proizvod koji se auto-postavi pri otvaranju frmOtkup/frmDokumenta.
Public Const CFG_DEFAULT_VRSTA As String = "DEFAULT_VRSTA_VOCA"
Public Const CFG_DEFAULT_SORTA As String = "DEFAULT_SORTA_VOCA"
' Filter kooperanata po otkupnom mestu u frmOtkup (default ON / prazno = ON).
Public Const CFG_KOOP_FILTER_BY_OM As String = "KOOP_FILTER_BY_OM"
' Auto otpremnica+zbirna+prijemnica kad je otkupno mesto hladnjaca (default OFF).
Public Const CFG_AUTO_PRIJEMNICA_HLADNJACA As String = "AUTO_PRIJEMNICA_HLADNJACA"
' Automatsko generisanje brojeva dokumenata (forma prefill OTK/OTP/ZBR). Default
' ON (kao do sada). NO -> forma ne predlaze broj; operater unosi svoj.
Public Const CFG_AUTO_BROJ_DOK As String = "AUTO_BROJ_DOKUMENTA"
' Kupac unosi BRUTO tezinu (sa ambalazom); sistem oduzima taru -> cuva neto (default OFF).
Public Const CFG_OTKUP_BRUTO_UNOS As String = "OTKUP_BRUTO_UNOS"
' Auto-kreiranje kooperanta iz unetog imena u frmOtkup (default ON / prazno = ON).
Public Const CFG_KOOP_AUTO_CREATE As String = "KOOP_AUTO_CREATE"
' Pracenje parcela: dozvoli unos parcele u frmOtkup (default ON). OFF -> polje skip.
Public Const CFG_PRACENJE_PARCELA As String = "PRACENJE_PARCELA"
' Kompletna validacija unosa (obavezna polja pre snimanja) u frmOtkup, frmDokumenta
' i frmPalete. Default ON (kao od v2.8.1). OFF -> minimalna validacija kao pre te izmene.
Public Const CFG_VALIDACIJA_UNOSA As String = "VALIDACIJA_UNOSA"
' Kes isplate proizvodjacima postoje (default ON). OFF -> skip Novac/Primalac u
' frmOtkup i "Br. otk. blk." u Ulaz OM (frmDokumenta).
Public Const CFG_KES_ISPLATE As String = "KES_ISPLATE"

' --- tblKulture: podrazumevani tip ambalaze (auto-puni u otkupu/dokumentima) ---
Public Const COL_KUL_TIP_AMBALAZE As String = "TipAmbalaze"
' --- tblStanice: flag hladnjaca (auto-lanac; kupac = MALINA_DEFAULT_KUPAC) ---
Public Const COL_STA_JE_HLADNJACA As String = "JeHladnjaca"

' =========================
' Auth / korisnici (Faza 1) - opt-in prijava + prava po oblasti
' =========================
' Opt-in flag (mirror CLOUD_SYNC_ENABLED / LICENSE_ENABLED). Dok nije YES -> bez prijave.
Public Const CFG_KEY_AUTH_ENABLED As String = "AUTH_ENABLED"
' Faza 3 - PIN hashing (opt-in). Default NO -> plaintext PIN (kompatibilno sa Faza 1/2).
Public Const CFG_KEY_PIN_HASH_ENABLED As String = "PIN_HASH_ENABLED"

' Uloge
Public Const ULOGA_ADMIN As String = "Admin"
Public Const ULOGA_KORISNIK As String = "Korisnik"

' Kolone tblKorisnici (upis po imenu - drift-safe)
Public Const COL_KOR_ID As String = "KorisnikID"
Public Const COL_KOR_USERNAME As String = "Username"
Public Const COL_KOR_IME As String = "ImePrezime"
Public Const COL_KOR_PIN As String = "PIN"
Public Const COL_KOR_ULOGA As String = "Uloga"
Public Const COL_KOR_AKTIVAN As String = "Aktivan"
Public Const COL_KOR_STANICA As String = "StanicaID"
Public Const COL_KOR_CREATED As String = "CreatedAt"

' Oblasti = nazivi kolona prava u tblKorisnici (i vrednosti koje vraca OblastZaFormu).
' Vrednost u celiji: "DA" = dozvoljeno, sve ostalo = zabranjeno. Admin uvek bypass.
Public Const OBL_OTKUP As String = "Otkup"
Public Const OBL_DOKUMENTA As String = "Dokumenta"
Public Const OBL_AGROHEMIJA As String = "Agrohemija"
Public Const OBL_IZVESTAJI As String = "Izvestaji"
Public Const OBL_FAKTURISANJE As String = "Fakturisanje"
Public Const OBL_BANKA As String = "Banka"
Public Const OBL_MARZA As String = "Marza"
Public Const OBL_SLEDLJIVOST As String = "Sledljivost"
Public Const OBL_MATICNI As String = "MaticniPodaci"
Public Const OBL_PALETE As String = "Palete"
Public Const OBL_OTVORI_EXCEL As String = "OtvoriExcel"
Public Const OBL_SYNC_PWA As String = "SyncPWA"

' --- Audit (timestamp + userstamp) - zajednicke kolone za sve glavne tabele ---
' Upisuju se centralno u modDataAccess.AppendRow (insert) i UpdateCell (izmena).
' Kolone se kreiraju jednokratno: Alt+F8 -> EnsureAuditColumns.
Public Const COL_AUDIT_CREATED_AT As String = "CreatedAt"
Public Const COL_AUDIT_CREATED_BY As String = "CreatedBy"
Public Const COL_AUDIT_MODIFIED_AT As String = "ModifiedAt"
Public Const COL_AUDIT_MODIFIED_BY As String = "ModifiedBy"

' =========================
' Workflow states
' =========================
Public Const WF_LOCAL_DRAFT As String = "LOCAL_DRAFT"
Public Const WF_LOCAL_FINALIZED As String = "LOCAL_FINALIZED"
Public Const WF_SEF_READY As String = "SEF_READY"
Public Const WF_SEF_SENDING As String = "SEF_SENDING"
Public Const WF_SEF_SENT As String = "SEF_SENT"
Public Const WF_SEF_ACCEPTED As String = "SEF_ACCEPTED"
Public Const WF_SEF_REJECTED As String = "SEF_REJECTED"
Public Const WF_SEF_STORNO As String = "SEF_STORNO"
Public Const WF_SEF_SYNC_ERROR As String = "SEF_SYNC_ERROR"
Public Const WF_SEF_TECH_FAILED As String = "SEF_TECH_FAILED"
Public Const WF_SEF_UNKNOWN As String = "SEF_UNKNOWN"

' =========================
' Submission statuses
' =========================
Public Const SEF_SUB_CREATED As String = "CREATED"
Public Const SEF_SUB_SENT As String = "SENT"
Public Const SEF_SUB_ACCEPTED As String = "ACCEPTED"
Public Const SEF_SUB_REJECTED As String = "REJECTED"
Public Const SEF_SUB_FAILED As String = "FAILED"
Public Const SEF_SUB_UNKNOWN As String = "UNKNOWN"

' =========================
' Event types
' =========================
Public Const SEF_EVT_VALIDATION_STARTED As String = "VALIDATION_STARTED"
Public Const SEF_EVT_VALIDATION_OK As String = "VALIDATION_OK"
Public Const SEF_EVT_VALIDATION_FAILED As String = "VALIDATION_FAILED"
Public Const SEF_EVT_PAYLOAD_BUILT As String = "PAYLOAD_BUILT"
Public Const SEF_EVT_STATE_CHANGED As String = "STATE_CHANGED"
Public Const SEF_EVT_HTTP_SENT As String = "HTTP_SENT"
Public Const SEF_EVT_HTTP_RESPONSE As String = "HTTP_RESPONSE"
Public Const SEF_EVT_SYNC_OK As String = "SYNC_OK"
Public Const SEF_EVT_SYNC_FAILED As String = "SYNC_FAILED"

' =========================
' Error codes
' =========================
Public Const ERR_SEF_VALIDATION As Long = vbObjectError + 3100
Public Const ERR_SEF_STATE As Long = vbObjectError + 3101
Public Const ERR_SEF_DUPLICATE As Long = vbObjectError + 3102
Public Const ERR_SEF_CONFIG As Long = vbObjectError + 3103
Public Const ERR_SEF_HTTP As Long = vbObjectError + 3104
Public Const ERR_SEF_RESPONSE_PARSE As Long = vbObjectError + 3105
Public Const ERR_SEF_REJECTED As Long = vbObjectError + 3106

' Status values suggested:
Public Const BIM_STATUS_NOVO As String = "Novo"
Public Const BIM_STATUS_AUTO As String = "AutoMapirano"
Public Const BIM_STATUS_RUCNO As String = "RucnoMapirano"
Public Const BIM_STATUS_PRESKOCENO As String = "Preskoceno"
Public Const BIM_STATUS_GRESKA As String = "Greska"

' MapTip values suggested:
Public Const BIM_MAPTIP_FAKTURA As String = "FakturaUplata"
Public Const BIM_MAPTIP_KOOPERANT As String = "KooperantIsplata"
Public Const BIM_MAPTIP_NEP As String = "Nepoznato"
Public Const BIM_MAPTIP_PROVIZIJA As String = "Provizija"

' --- Izvestaj tipovi ---
Public Const IZV_POJEDINACNI As String = "Pojedinacni"
Public Const IZV_ZBIRNI As String = "Zbirni"


' --- SEF HTTPS ---
Public Const HTTP_TIMEOUT_RESOLVE_MS As Long = 10000
Public Const HTTP_TIMEOUT_CONNECT_MS As Long = 10000
Public Const HTTP_TIMEOUT_SEND_MS As Long = 30000
Public Const HTTP_TIMEOUT_RECEIVE_MS As Long = 30000

Public Function GetConfigValue(ByVal configKey As String) As String

    Dim v As Variant

    v = LookupValue("tblSEFConfig", "ConfigKey", configKey, "ConfigValue")

    If IsEmpty(v) Then
        GetConfigValue = ""
    Else
        GetConfigValue = Trim$(CStr(v))
    End If

End Function

Public Sub SetConfigValue(ByVal configKey As String, ByVal ConfigValue As String)
    Const SOURCE As String = "SetConfigValue"

    Dim data As Variant
    Dim colKey As Long
    Dim colValue As Long
    Dim i As Long
    Dim found As Boolean
    Dim lo As ListObject
    Dim rowData() As Variant
    Dim newRowIndex As Long

    On Error GoTo EH

    If Len(Trim$(configKey)) = 0 Then
        Err.Raise vbObjectError + 7500, SOURCE, "ConfigKey je prazan."
    End If

    Set lo = GetTable(TBL_SEF_CONFIG)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 7501, SOURCE, "Tabela ne postoji: " & TBL_SEF_CONFIG
    End If

    colKey = RequireColumnIndex(TBL_SEF_CONFIG, "ConfigKey", SOURCE)
    colValue = RequireColumnIndex(TBL_SEF_CONFIG, "ConfigValue", SOURCE)

    data = GetTableData(TBL_SEF_CONFIG)

    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Trim$(CStr(data(i, colKey))) = Trim$(configKey) Then
                RequireUpdateCell TBL_SEF_CONFIG, i, "ConfigValue", ConfigValue, SOURCE
                found = True
                Exit For
            End If
        Next i
    End If

    If Not found Then
        ReDim rowData(1 To lo.ListColumns.count)

        rowData(colKey) = Trim$(configKey)
        rowData(colValue) = ConfigValue

        newRowIndex = AppendRow(TBL_SEF_CONFIG, rowData)
        If newRowIndex = 0 Then
            Err.Raise vbObjectError + 7502, SOURCE, _
                      "AppendRow nije uspeo za config key: " & configKey
        End If
    End If

    Exit Sub

EH:
    LogErr SOURCE
    Err.Raise Err.Number, SOURCE, Err.description
End Sub


' ============================================================
' Cloud sync toggle.
'
' Pozivaju ga svi moduli koji zele da pisu/citaju Google sheet-ove
' (modStanicaLock, modBrojevi.MaxSeqFromGoogleSheet). Kada vraca False,
' VBA radi 100% lokalno -- nikakav HTTP saobracaj prema Google.
'
' Backward-compatible default: ako Config nema ovaj kljuc ili je prazno,
' tretira se kao True (postojeci klijenti rade kao i do sada).
' ============================================================
Public Function IsCloudSyncEnabled() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_CLOUD_SYNC_ENABLED)))
    
    Select Case v
        Case ""           ' nema kljuca -- default TRUE
            IsCloudSyncEnabled = True
        Case "NO", "FALSE", "0", "OFF", "DISABLED"
            IsCloudSyncEnabled = False
        Case Else         ' "YES", "TRUE", "1", "ON", "ENABLED" ili bilo sta drugo
            IsCloudSyncEnabled = True
    End Select
End Function


' ============================================================
' Malina mod toggle.
'
' Kada vraca True: 1 stanica = 1 vozilo; sistem auto-pravi zbirnu iz
' otpremnice (otpremnica == zbirna). Cita config kljuc MALINA_MODE.
'
' Default FALSE (suprotno od IsCloudSyncEnabled): bez kljuca ili prazno
' tretira se kao visnja ponasanje (nepromenjeno za postojece klijente).
' ============================================================
Public Function IsMalinaMode() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_MALINA_MODE)))

    Select Case v
        Case "YES", "TRUE", "1", "ON", "ENABLED"
            IsMalinaMode = True
        Case Else         ' "", "NO", "FALSE", "0" ili bilo sta drugo -> FALSE
            IsMalinaMode = False
    End Select
End Function

' Genericki bool config citac. Prazno/nepoznato -> defaultOn.
Public Function ConfigFlag(ByVal key As String, ByVal defaultOn As Boolean) As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(key)))
    Select Case v
        Case "YES", "TRUE", "1", "ON", "ENABLED", "DA"
            ConfigFlag = True
        Case "NO", "FALSE", "0", "OFF", "DISABLED", "NE"
            ConfigFlag = False
        Case Else
            ConfigFlag = defaultOn
    End Select
End Function

' Filter kooperanata po otkupnom mestu u frmOtkup (default ON).
Public Function KoopFilterByOM() As Boolean
    KoopFilterByOM = ConfigFlag(CFG_KOOP_FILTER_BY_OM, True)
End Function

' Auto otpremnica+zbirna+prijemnica kad je otkupno mesto hladnjaca (default OFF).
Public Function IsAutoPrijemnicaHladnjaca() As Boolean
    IsAutoPrijemnicaHladnjaca = ConfigFlag(CFG_AUTO_PRIJEMNICA_HLADNJACA, False)
End Function

' Automatsko generisanje brojeva dokumenata (forma prefill OTK/OTP/ZBR). Default ON.
' NO -> SuggestNextBroj vraca "" pa forma ostavlja polje prazno (operater unosi svoj).
Public Function IsAutoBrojDokumenta() As Boolean
    IsAutoBrojDokumenta = ConfigFlag(CFG_AUTO_BROJ_DOK, True)
End Function

' Paletiranje (izrada paletnih listova iz prijemnice). Default ON.
' NO -> PaletizePrijemnica se preskace (kao kad nema gajbica); prijemnica se i dalje snima.
Public Function IsPaletiranjeEnabled() As Boolean
    IsPaletiranjeEnabled = ConfigFlag(CFG_PALETIRANJE, True)
End Function

' Kupac unosi BRUTO tezinu (sa ambalazom); sistem oduzima taru ambalaze i cuva
' neto u Kolicina, a bruto u BrutoKg (frmOtkup). Default OFF -> postojece ponasanje.
Public Function OtkupBrutoUnos() As Boolean
    OtkupBrutoUnos = ConfigFlag(CFG_OTKUP_BRUTO_UNOS, False)
End Function

' Auto-kreiranje kooperanta kad se u frmOtkup unese ime koje nije u bazi.
' Default ON (kao do sada). NO -> nepoznato ime se ne kreira (operater bira
' postojeceg ili rucno doda u Maticnim podacima).
Public Function KoopAutoCreate() As Boolean
    KoopAutoCreate = ConfigFlag(CFG_KOOP_AUTO_CREATE, True)
End Function

' Pracenje parcela u frmOtkup (cmbParcela). Default ON. OFF -> polje se skip.
Public Function IsPracenjeParcela() As Boolean
    IsPracenjeParcela = ConfigFlag(CFG_PRACENJE_PARCELA, True)
End Function

' Kompletna validacija unosa (obavezna polja pre snimanja) u frmOtkup/frmDokumenta/
' frmPalete. Default ON (kao od v2.8.1 -- "Blokator obaveznih polja"). OFF vraca
' minimalnu validaciju kao pre te izmene (npr. broj gajbi u frmOtkup tada ostaje
' obavezan iskljucivo u bruto rezimu).
Public Function IsValidacijaUnosa() As Boolean
    IsValidacijaUnosa = ConfigFlag(CFG_VALIDACIJA_UNOSA, True)
End Function

' Postoje kes isplate proizvodjacima. Default ON. OFF -> skip Novac/Primalac
' (frmOtkup) i "Br. otk. blk." (Ulaz OM, frmDokumenta).
Public Function IsKesIsplate() As Boolean
    IsKesIsplate = ConfigFlag(CFG_KES_ISPLATE, True)
End Function

