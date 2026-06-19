Option Explicit

' ============================================================
' modConfig – Zentrale Konfiguration
' Alle Konstanten, Tabellennamen, Spaltenindizes
' KEINE Hardcoded Zellreferenzen im restlichen Code!
' ============================================================

' --- App Info ---
Public Const APP_NAME As String = "OtkupApp"
Public Const APP_VERSION As String = "2.2.1"

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
Public Const COL_PRJ_VOZAC As String = "VozacID"
Public Const COL_PRJ_VRSTA As String = "VrstaVoca"
Public Const COL_PRJ_SORTA As String = "SortaVoca"
Public Const COL_PRJ_KLASA As String = "Klasa"
Public Const COL_PRJ_FAKTURISANO As String = "Fakturisano"
Public Const COL_PRJ_FAKTURA_ID As String = "FakturaID"

' ============================================================
' Paletni list (Phase 2) — kolone novih tabela
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

' tblCenovnik (cene po proizvodu — append-only istorija)
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

' Paleta status
Public Const PAL_STATUS_OTVORENA As String = "Otvorena"
Public Const PAL_STATUS_ZATVORENA As String = "Zatvorena"

' Paleta — config kljuc (default tip palete) + fallback kapacitet
Public Const CFG_DEFAULT_TIP_PALETE As String = "DEFAULT_TIP_PALETE"
Public Const PALETA_DEFAULT_KAPACITET As Long = 240
Public Const CFG_PALETA_PRINT_MODE As String = "PALETA_PRINT_MODE"
Public Const CFG_OTKUP_PRINT_MODE As String = "OTKUP_PRINT_MODE"
Public Const CFG_PDV_NADOKNADA_STOPA As String = "PDV_NADOKNADA_STOPA"
Public Const PDV_NADOKNADA_DEFAULT As Double = 8

' Otkupni list - obavezni elementi otkupnog bloka (klauzula + rok isplate)
Public Const CFG_OTKUP_KLAUZULA As String = "OTKUP_KLAUZULA"
Public Const CFG_OTKUP_ROK As String = "OTKUP_ROK_ISPLATE"
Public Const OTKUP_ROK_DEFAULT As String = "Po dogovoru"

' --- Dokument-Tipovi ---
Public Const COL_STORNIRANO As String = "Stornirano"
Public Const COL_OSIROCENO_OD As String = "OsirocenoOd"
Public Const DOK_TIP_OTKUP As String = "Otkup"
Public Const DOK_TIP_OTPREMNICA As String = "Otpremnica"
Public Const DOK_TIP_PRIJEMNICA As String = "Prijemnica"
Public Const DOK_TIP_IZLAZ_KUPCI As String = "Kupci-Otpremnica"

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
Public Const APP_BANKA_INBOX As String = "G:\My Drive\Bank_Izvodi"
Public Const APP_BANKA_PROCESSED As String = "G:\My Drive\Bank_Izvodi\Verarbeitet"
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

' --- Typen Ambalaže ---
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

' --- tblKulture: podrazumevani tip ambalaze (auto-puni u otkupu/dokumentima) ---
Public Const COL_KUL_TIP_AMBALAZE As String = "TipAmbalaze"
' --- tblStanice: flag hladnjaca (auto-lanac; kupac = MALINA_DEFAULT_KUPAC) ---
Public Const COL_STA_JE_HLADNJACA As String = "JeHladnjaca"

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

' --- Izveštaj tipovi ---
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
' Pozivaju ga svi moduli koji žele da pišu/citaju Google sheet-ove
' (modStanicaLock, modBrojevi.MaxSeqFromGoogleSheet). Kada vraca False,
' VBA radi 100% lokalno — nikakav HTTP saobracaj prema Google.
'
' Backward-compatible default: ako Config nema ovaj kljuc ili je prazno,
' tretira se kao True (postojeci klijenti rade kao i do sada).
' ============================================================
Public Function IsCloudSyncEnabled() As Boolean
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_CLOUD_SYNC_ENABLED)))
    
    Select Case v
        Case ""           ' nema kljuca — default TRUE
            IsCloudSyncEnabled = True
        Case "NO", "FALSE", "0", "OFF", "DISABLED"
            IsCloudSyncEnabled = False
        Case Else         ' "YES", "TRUE", "1", "ON", "ENABLED" ili bilo šta drugo
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
