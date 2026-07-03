Attribute VB_Name = "modPoruke"
Option Explicit

Private mCache As Object

Public Function Poruka(ByVal kljuc As String) As String
    On Error GoTo EH
    If mCache Is Nothing Then BuildCache
    If mCache.Exists(kljuc) Then
        Poruka = mCache(kljuc)
    Else
        Debug.Print "modPoruke: missing key [" & kljuc & "]"
        Poruka = "[" & kljuc & "]"
    End If
    Exit Function
EH:
    Poruka = "[" & kljuc & "]"
End Function

Public Sub InvalidateCache()
    Set mCache = Nothing
End Sub

Private Sub BuildCache()
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 0
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = TBL_PORUKE Then GoTo Found
        Next lo
    Next ws
    Exit Sub
Found:
    Dim r As ListRow
    Dim k As String
    For Each r In lo.ListRows
        k = CStr(r.Range(1).Value)
        If k <> "" Then mCache(k) = CStr(r.Range(2).Value)
    Next r
End Sub

Public Sub UpsertPoruke(lo As ListObject)
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = 0
    Dim r As ListRow
    Dim k As String
    For Each r In lo.ListRows
        k = CStr(r.Range(1).Value)
        If k <> "" Then existing(k) = r.Index
    Next r
    UpsertRow lo, existing, "APP_MSG_GRESKA_PRI_POKRETANJU", "Gre" & ChrW(353) & "ka pri pokretanju aplikacije. Pogledajte log."
    UpsertRow lo, existing, "AGRO_LBL_MAGACIN_IZDAVANJE_ROBE", "Magacin " & ChrW(8212) & " izdavanje robe i prijem od dobavlja" & ChrW(269) & "a"
    UpsertRow lo, existing, "AGRO_LBL_IZLAZ_IZDAVANJE_ROBE", "Izlaz " & ChrW(8212) & " Izdavanje robe kooperantu"
    UpsertRow lo, existing, "AGRO_LBL_ULAZ_PRIJEM_ROBE", "Ulaz " & ChrW(8212) & " Prijem robe od dobavlja" & ChrW(269) & "a"
    UpsertRow lo, existing, "AGRO_LBL_ZAVRSI_IZDAVANJE", "Zavr" & ChrW(353) & "i izdavanje"
    UpsertRow lo, existing, "AGRO_LBL_ZAVRSI_PRIJEM", "Zavr" & ChrW(353) & "i prijem"
    UpsertRow lo, existing, "AGRO_LBL_POPUNI_PAKOVANJE_FIELD", "Popuni 'Pakovanje' field u tblArtikli pre kori" & ChrW(353) & "cenja."
    UpsertRow lo, existing, "AGRO_MSG_POKUSAVATE_DODATI", "Poku" & ChrW(353) & "avate dodati:"
    UpsertRow lo, existing, "AGRO_MSG_IZDAVANJE_ZAVRSENO", "Izdavanje zavr" & ChrW(353) & "eno:"
    UpsertRow lo, existing, "AGRO_MSG_GRESKA_PRI_CUVANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju izdavanja, promene vra" & ChrW(263) & "ene:"
    UpsertRow lo, existing, "AGRO_MSG_PRIJEM_ZAVRSEN", "Prijem zavr" & ChrW(353) & "en:"
    UpsertRow lo, existing, "AGRO_MSG_GRESKA_PRI_CUVANJU_2", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju prijema, promene vra" & ChrW(263) & "ene:"
    UpsertRow lo, existing, "BANKA_LBL_GENERISI_CSV_COMMIT", "Generi" & ChrW(353) & "i CSV naloge"
    UpsertRow lo, existing, "BANKA_LBL_UVOZ_TRANSAKCIJA_BANKARSKIH", "Uvoz transakcija iz bankarskih izvoda " & ChrW(8212) & " mapiranje na partnere"
    UpsertRow lo, existing, "DOK_LBL_OPERATIVNI_UNOS_DOKUMENATA", "Operativni unos dokumenata, ambala" & ChrW(382) & "e i novca"
    UpsertRow lo, existing, "DOK_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju dokumenata:"
    UpsertRow lo, existing, "DOK_MSG_UNOSI_SAMO_KLASA", "Unosi se samo II klasa: obri" & ChrW(353) & "ite koli" & ChrW(269) & "inu ambala" & ChrW(382) & "e za I klasu."
    UpsertRow lo, existing, "DOK_MSG_UNESITE_ISPRAVNU_KOLICINU", "Unesite ispravnu koli" & ChrW(269) & "inu ambala" & ChrW(382) & "e!"
    UpsertRow lo, existing, "DOK_MSG_UNESITE_ISPRAVNU_KOLICINU_2", "Unesite ispravnu koli" & ChrW(269) & "inu ambala" & ChrW(382) & "e za II klasu!"
    UpsertRow lo, existing, "DOK_MSG_TIP_AMBALAZE", "Tip ambala" & ChrW(382) & "e '"
    UpsertRow lo, existing, "DOK_MSG_NEMA_UNETU_TEZINU", "' nema unetu te" & ChrW(382) & "inu gajbice"
    UpsertRow lo, existing, "DOK_MSG_MATICNI_PODACI_TIP", "(Mati" & ChrW(269) & "ni podaci ? Tip ambala" & ChrW(382) & "e)."
    UpsertRow lo, existing, "DOK_MSG_BRUTO_MOZE_PRETVORITI", "Bruto se ne mo" & ChrW(382) & "e pretvoriti u neto."
    UpsertRow lo, existing, "DOK_MSG_TEZINA_AMBALAZE", "Te" & ChrW(382) & "ina ambala" & ChrW(382) & "e ("
    UpsertRow lo, existing, "DOK_MSG_JEDNAKA_BRUTO_TEZINI", "jednaka bruto te" & ChrW(382) & "ini ("
    UpsertRow lo, existing, "DOK_MSG_PROVERITE_BROJ_KOMADA", "Proverite broj komada ili tip ambala" & ChrW(382) & "e."
    UpsertRow lo, existing, "DOK_MSG_BRUTO_KLASA_MOZE", "Bruto (II klasa) se ne mo" & ChrW(382) & "e pretvoriti u neto."
    UpsertRow lo, existing, "DOK_MSG_TEZINA_AMBALAZE_KLASE", "Te" & ChrW(382) & "ina ambala" & ChrW(382) & "e II klase ("
    UpsertRow lo, existing, "DOK_MSG_GRESKA_PRI_CUVANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju otpremnice. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "DOK_ERR_GRESKA", "Gre" & ChrW(353) & "ka:"
    UpsertRow lo, existing, "DOK_MSG_UNESITE_AMBALAZU_ILI", "Unesite ambala" & ChrW(382) & "u ili novac za OM ulaz."
    UpsertRow lo, existing, "DOK_MSG_IZABERITE_TIP_AMBALAZE", "Izaberite tip ambala" & ChrW(382) & "e!"
    UpsertRow lo, existing, "DOK_MSG_NEDOVOLJNO_AVANSA_RASPOLOZIVO", "Nedovoljno OM avansa! Raspolo" & ChrW(382) & "ivo:"
    UpsertRow lo, existing, "DOK_MSG_GRESKA_PRI_CUVANJU_2", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju OM ulaza. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "DOK_MSG_VALIDACIJA_NIJE_PROSLA", "Validacija nije pro" & ChrW(353) & "la. Proverite kg i ambala" & ChrW(382) & "u (razlika mora biti 0)."
    UpsertRow lo, existing, "DOK_MSG_GRESKA_PRI_CUVANJU_3", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju zbirne. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "DOK_LBL_NEISPRAVNA_KOLICINA_AMBALAZE", "Neispravna koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e"
    UpsertRow lo, existing, "DOK_MSG_UNESITE_ISPRAVNU_KOLICINU_3", "Unesite ispravnu koli" & ChrW(269) & "inu vra" & ChrW(263) & "ene ambala" & ChrW(382) & "e!"
    UpsertRow lo, existing, "DOK_MSG_GRESKA_PRI_CUVANJU_4", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju prijemnice. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "DOK_MSG_UNESITE_AMBALAZU_ILI_2", "Unesite ambala" & ChrW(382) & "u ili uplatu kupca."
    UpsertRow lo, existing, "DOK_MSG_GRESKA_PRI_CUVANJU_5", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju izlaza kupca. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "DOK_ERR_GRESKA_PRI_PRIKAZU", "Gre" & ChrW(353) & "ka pri prikazu storniranih:"
    UpsertRow lo, existing, "FAK_LBL_STAMPAJ", ChrW(352) & "tampaj"
    UpsertRow lo, existing, "FAK_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju fakturisanja:"
    UpsertRow lo, existing, "FAK_ERR_GRESKA_PRI_UCITAVANJU", "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju prijemnica:"
    UpsertRow lo, existing, "FAK_MSG_ISTA_PRIJEMNICA_IZABRANA", "Ista prijemnica je izabrana vi" & ChrW(353) & "e puta:"
    UpsertRow lo, existing, "FAK_MSG_IZABRANA_PRIJEMNICA_STORNIRANA", "Izabrana prijemnica je stornirana i ne mo" & ChrW(382) & "e se fakturisati."
    UpsertRow lo, existing, "FAK_MSG_IZABRANA_PRIJEMNICA_VEC", "Izabrana prijemnica je ve" & ChrW(263) & " fakturisana i ne mo" & ChrW(382) & "e biti uklju" & ChrW(269) & "ena u novu fakturu."
    UpsertRow lo, existing, "FAK_MSG_GRESKA_PRI_KREIRANJU", "Gre" & ChrW(353) & "ka pri kreiranju fakture. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "FAK_ERR_GRESKA_PRI_IZRADI", "Gre" & ChrW(353) & "ka pri izradi fakture:"
    UpsertRow lo, existing, "FAK_MSG_NIJE_PRONADEN_FAKTURE", "Nije prona" & ChrW(273) & "en ID fakture za " & ChrW(353) & "tampu."
    UpsertRow lo, existing, "FAK_ERR_GRESKA_PRI_STAMPANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(353) & "tampanju:"
    UpsertRow lo, existing, "RPT_LBL_IZVESTAJI", "Izve" & ChrW(353) & "taji"
    UpsertRow lo, existing, "RPT_LBL_PREGLED_SALDO_OTKUPLJENE", "Pregled saldo, otkupljene robe, ambala" & ChrW(382) & "e i prose" & ChrW(269) & "ne cene"
    UpsertRow lo, existing, "RPT_LBL_PRIKAZI", "Prika" & ChrW(382) & "i"
    UpsertRow lo, existing, "RPT_LBL_STAMPAJ_KARTICU", ChrW(352) & "tampaj karticu"
    UpsertRow lo, existing, "RPT_LBL_TIP_IZVESTAJA", "TIP IZVE" & ChrW(352) & "TAJA"
    UpsertRow lo, existing, "RPT_LBL_AGRO_ZADUZENJE", "Agro zadu" & ChrW(382) & "enje"
    UpsertRow lo, existing, "RPT_LBL_AMBALAZA", "Ambala" & ChrW(382) & "a"
    UpsertRow lo, existing, "RPT_LBL_KES_OTKUPAC", "Ke" & ChrW(353) & " otkupac"
    UpsertRow lo, existing, "RPT_LBL_ZADUZENJE", "Zadu" & ChrW(382) & "enje"
    UpsertRow lo, existing, "RPT_LBL_RAZDUZENJE", "Razdu" & ChrW(382) & "enje"
    UpsertRow lo, existing, "MATICNI_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju menija " & ChrW(353) & "ifarnika:"
    UpsertRow lo, existing, "MATICNI_ERR_GRESKA_PRI_OTVARANJU_2", "Gre" & ChrW(353) & "ka pri otvaranju " & ChrW(353) & "ifarnika:"
    UpsertRow lo, existing, "OTKUP_MSG_POKUSAJ_PONOVO", ". Poku" & ChrW(353) & "aj ponovo."
    UpsertRow lo, existing, "OTKUP_MSG_UNESITE_ISPRAVNU_KOLICINU", "Unesite ispravnu koli" & ChrW(269) & "inu izdate ambala" & ChrW(382) & "e!"
    UpsertRow lo, existing, "OTKUP_MSG_IZABERITE_TIP_AMBALAZE", "Izaberite tip ambala" & ChrW(382) & "e (za II klasu)!"
    UpsertRow lo, existing, "OTKUP_MSG_IZABERITE_TIP_AMBALAZE_2", "Izaberite tip ambala" & ChrW(382) & "e (za izdatu ambala" & ChrW(382) & "u)!"
    UpsertRow lo, existing, "OTKUP_MSG_BRUTO_REZIM_UNESITE", "Bruto re" & ChrW(382) & "im: unesite broj gajbi za I klasu"
    UpsertRow lo, existing, "OTKUP_MSG_BRUTO_REZIM_UNESITE_2", "Bruto re" & ChrW(382) & "im: unesite broj gajbi za II klasu"
    UpsertRow lo, existing, "OTKUP_MSG_ZELITE_IPAK_NASTAVITE", ChrW(381) & "elite li ipak da nastavite?"
    UpsertRow lo, existing, "OTKUP_MSG_GRESKA_PRI_CUVANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju otkupa. Promene su vra" & ChrW(263) & "ene."
    UpsertRow lo, existing, "OTKUP_ERR_GRESKA_PRI_UNOSU", "Gre" & ChrW(353) & "ka pri unosu otkupa:"
    UpsertRow lo, existing, "OTKUP_ERR_GRESKA_PRI_POKRETANJU", "Gre" & ChrW(353) & "ka pri pokretanju glavnog ekrana:"
    UpsertRow lo, existing, "OTKUP_LBL_GRESKA_PRI_PROVERI", "Gre" & ChrW(353) & "ka pri proveri dokumenata. Pogledajte log."
    UpsertRow lo, existing, "OTKUP_LBL_IZVESTAJ", "Izve" & ChrW(353) & "taj"
    UpsertRow lo, existing, "OTKUP_LBL_MARZA", "Mar" & ChrW(382) & "a"
    UpsertRow lo, existing, "OTKUP_LBL_IZVESTAJ_SLEDLJIVOSTI", "Izve" & ChrW(353) & "taj o sledljivosti"
    UpsertRow lo, existing, "OTKUP_LBL_GRESKA_PRI_UVOZU", "Gre" & ChrW(353) & "ka pri uvozu banke. Otvaram mapiranje za postojece stavke."
    UpsertRow lo, existing, "OTKUP_MSG_GRESKA_PRI_UVOZU", "Gre" & ChrW(353) & "ka pri uvozu bankovnih izvoda:"
    UpsertRow lo, existing, "OTKUP_LBL_SEKCIJA_IZVESTAJ_SLEDLJIVOSTI", "Sekcija: Izve" & ChrW(353) & "taj o sledljivosti"
    UpsertRow lo, existing, "OTKUP_LBL_GRESKA_PRI_SNIMANJU", "Gre" & ChrW(353) & "ka pri snimanju. Pogledajte log."
    UpsertRow lo, existing, "OTKUP_ERR_GRESKA_PRI_SNIMANJU", "Gre" & ChrW(353) & "ka pri snimanju:"
    UpsertRow lo, existing, "OTKUP_ERR_FAJL_NIJE_USPESNO", "Fajl nije uspe" & ChrW(353) & "no sa" & ChrW(269) & "uvan ili aplikacija nije zatvorena:"
    UpsertRow lo, existing, "OTKUP_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju mati" & ChrW(269) & "nih podataka:"
    UpsertRow lo, existing, "OTKUP_LBL_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju sekcije:"
    UpsertRow lo, existing, "OTKUP_ERR_GRESKA_PRI_OTVARANJU_2", "Gre" & ChrW(353) & "ka pri otvaranju sekcije '"
    UpsertRow lo, existing, "OTKUP_LBL_DANASNJI_OTKUP", "Dana" & ChrW(353) & "nji otkup"
    UpsertRow lo, existing, "OTKUP_LBL_OTPREMNICA", "otpremnica  " & ChrW(8226)
    UpsertRow lo, existing, "PAL_LBL_PREGLED_PALETA_STAMPA", "Pregled paleta, " & ChrW(353) & "tampa i prerada"
    UpsertRow lo, existing, "SEF_LBL_POSALJI_SEF", "Po" & ChrW(353) & "alji na SEF"
    UpsertRow lo, existing, "SEF_LBL_OSVEZI_STATUS", "Osve" & ChrW(382) & "i status"
    UpsertRow lo, existing, "SEF_LBL_OTKAZI_SLANJE_SEF", "Otka" & ChrW(382) & "i slanje na SEF"
    UpsertRow lo, existing, "SEF_LBL_OSVEZI_SVE_PENDING", "Osve" & ChrW(382) & "i sve Pending"
    UpsertRow lo, existing, "SEF_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju SEF forme:"
    UpsertRow lo, existing, "SEF_MSG_SEF_STATUS_OSVEZEN", "SEF status osve" & ChrW(382) & "en."
    UpsertRow lo, existing, "SEF_MSG_CANCEL_USPESNO_POSLAT", "Cancel uspe" & ChrW(353) & "no poslat."
    UpsertRow lo, existing, "SEF_MSG_STORNO_USPESNO_POSLAT", "Storno uspe" & ChrW(353) & "no poslat."
    UpsertRow lo, existing, "SEF_MSG_RECOVERY_ZAVRSEN", "Recovery zavr" & ChrW(353) & "en."
    UpsertRow lo, existing, "SEF_MSG_PENDING_FAKTURE_OSVEZENE", "Pending fakture osve" & ChrW(382) & "ene."
    UpsertRow lo, existing, "SEF_MSG_SEF_SENDING_RECOVERY", "SEF_SENDING recovery zavr" & ChrW(353) & "en."
    UpsertRow lo, existing, "SLED_LBL_POVEZI", "Pove" & ChrW(382) & "i"
    UpsertRow lo, existing, "STM_LBL_OBRISI_GEO", "Obri" & ChrW(353) & "i geo"
    UpsertRow lo, existing, "STM_LBL_CENE_PROIZVODU_SVAKA", "Cene po proizvodu " & ChrW(8212) & " svaka promena je novi red (vazi poslednji)"
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_PROMENI", "Gre" & ChrW(353) & "ka pri promeni statusa:"
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_UCITAVANJU", "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju liste:"
    UpsertRow lo, existing, "STM_MSG_UNESITE_KATASTARSKU_OPSTINU", "Unesite katastarsku op" & ChrW(353) & "tinu!"
    UpsertRow lo, existing, "STM_MSG_UNESITE_VALIDNU_POVRSINU", "Unesite validnu povr" & ChrW(353) & "inu!"
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_DODAVANJU", "Gre" & ChrW(353) & "ka pri dodavanju!"
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_DODAVANJU_2", "Gre" & ChrW(353) & "ka pri dodavanju cene!"
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_DODAVANJU", "Gre" & ChrW(353) & "ka pri dodavanju:"
    UpsertRow lo, existing, "STM_MSG_NOVU_CENU_KORISTITE", "Za novu cenu koristite 'Dodaj' " & ChrW(8212) & " stari redovi ostaju radi istorije."
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_IZMENI", "Gre" & ChrW(353) & "ka pri izmeni:"
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_UCITAVANJU_2", "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju stanica:"
    UpsertRow lo, existing, "AGRO_MSG_GRESKA_PRI_UNOSU", "Gre" & ChrW(353) & "ka pri unosu magacina, promene vra" & ChrW(263) & "ene:"
    UpsertRow lo, existing, "BANKA_LBL_PDF_AUSWAHLEN", "Izaberi PDF"
    UpsertRow lo, existing, "GAUTH_MSG_TOKEN_AUSTAUSCH_FEHLGESCHLAGEN", "Razmena tokena nije uspela. Proveri Client ID/Secret."
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_OTK_NIJE", "Uvoz OTK nije potvr" & ChrW(273) & "en zbog fatal sync gre" & ChrW(353) & "ke."
    UpsertRow lo, existing, "SYNC_ERR_GRESKE", "Gre" & ChrW(353) & "ke:"
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_OTK_ZAVRSEN", "Uvoz OTK zavr" & ChrW(353) & "en sa gre" & ChrW(353) & "kama."
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_OTK_ZAVRSEN_2", "Uvoz OTK zavr" & ChrW(353) & "en!"
    UpsertRow lo, existing, "SYNC_MSG_GRESKA_PRI_UVOZU", "Gre" & ChrW(353) & "ka pri uvozu OTK:"
    UpsertRow lo, existing, "SYNC_MSG_PWA_UVOZ_NIJE", "PWA uvoz nije potvr" & ChrW(273) & "en. Promene su vra" & ChrW(263) & "ene zbog fatal sync gre" & ChrW(353) & "ke. Proveri log."
    UpsertRow lo, existing, "SYNC_MSG_PWA_UVOZ_ZAVRSEN", "PWA uvoz zavr" & ChrW(353) & "en i potvr" & ChrW(273) & "en."
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_ZBIRNIH_NIJE", "Uvoz zbirnih nije potvr" & ChrW(273) & "en zbog fatal sync gre" & ChrW(353) & "ke."
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_ZBIRNIH_ZAVRSEN", "Uvoz zbirnih zavr" & ChrW(353) & "en sa gre" & ChrW(353) & "kama."
    UpsertRow lo, existing, "SYNC_ERR_UVOZ_ZBIRNIH_ZAVRSEN_2", "Uvoz zbirnih zavr" & ChrW(353) & "en!"
    UpsertRow lo, existing, "SYNC_MSG_GRESKA_PRI_UVOZU_2", "Gre" & ChrW(353) & "ka pri uvozu zbirnih:"
    UpsertRow lo, existing, "SYNC_MSG_UVOZ_ZBIRNIH_ZAVRSEN", "Uvoz zbirnih zavr" & ChrW(353) & "en sa gre" & ChrW(353) & "kama ili fatal sync problemom."
    UpsertRow lo, existing, "SYNC_MSG_USPESNI_REDOVI_KOJI", "Uspe" & ChrW(353) & "ni redovi koji su potvr" & ChrW(273) & "eni kroz row-level TX ostaju upisani."
    UpsertRow lo, existing, "SYNC_MSG_NEUSPESNI_REDOVI_OZNACENI", "Neuspe" & ChrW(353) & "ni redovi su ozna" & ChrW(269) & "eni kao gre" & ChrW(353) & "ka gde je writeback uspeo."
    UpsertRow lo, existing, "SYNC_MSG_UVOZ_ZBIRNIH_ZAVRSEN_2", "Uvoz zbirnih zavr" & ChrW(353) & "en i potvr" & ChrW(273) & "en."
    UpsertRow lo, existing, "OTKUP_LBL_OTKUPNI_BLOKOVI", "Otkupni blokovi  " & ChrW(187)
    UpsertRow lo, existing, "OTKUP_LBL_SAKRIJ_BLOKOVE", ChrW(171) & "  Sakrij blokove"
    UpsertRow lo, existing, "OTKUP_LBL_UKUPNO", "Ukupno kg: " & ChrW(8212)
    UpsertRow lo, existing, "OTKUP_LBL_BLOKOVIMA", "U blokovima: " & ChrW(8212)
    UpsertRow lo, existing, "OTKUP_LBL_OSTATAK", "Ostatak: " & ChrW(8212)
    UpsertRow lo, existing, "OTKUP_LBL_UKUPNO_AMB", "Ukupno amb: " & ChrW(8212)
    UpsertRow lo, existing, "OTKUP_LBL_BLOKOVIMA_AMB", "U blokovima amb: " & ChrW(8212)
    UpsertRow lo, existing, "OTKUP_LBL_OSTATAK_AMB", "Ostatak amb: " & ChrW(8212)
    UpsertRow lo, existing, "CFG_LBL_PODESAVANJA_TBLSEFCONFIG", "Pode" & ChrW(353) & "avanja (tblSEFConfig)"
    UpsertRow lo, existing, "CFG_LBL_STR", ChrW(8212)
    UpsertRow lo, existing, "CFG_LBL_STR_2", ChrW(8212)
    UpsertRow lo, existing, "CFG_ERR_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju pode" & ChrW(353) & "avanja:"
    UpsertRow lo, existing, "CFG_MSG_PRESKOCENO_GRESKA", "Preskoceno (gre" & ChrW(353) & "ka):"
    UpsertRow lo, existing, "CFG_ERR_GRESKA_PRI_CUVANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju:"
    UpsertRow lo, existing, "CFG_MSG_IZLAZ_NUZDI_ALT", "(Izlaz u nu" & ChrW(382) & "di: Alt+F8 -> ShowConfigSheet.)"
    UpsertRow lo, existing, "SETUP_MSG_SETUP_USPESNO_ZAVRSEN", "Setup je uspe" & ChrW(353) & "no zavr" & ChrW(353) & "en."
    UpsertRow lo, existing, "SETUP_MSG_PODESAVANJA_MATICNI_PODACI", "Pode" & ChrW(353) & "avanja: Mati" & ChrW(269) & "ni podaci -> Pode" & ChrW(353) & "avanja (tblSEFConfig je skriven)."
    UpsertRow lo, existing, "SETUP_MSG_SETUP_ZAVRSEN_ALI", "Setup je zavr" & ChrW(353) & "en, ali postoje stavke za proveru:"
    UpsertRow lo, existing, "SETUP_ERR_GRESKA_TOKOM_SETUP", "Gre" & ChrW(353) & "ka tokom setup-a:"
    UpsertRow lo, existing, "SETUP_MSG_HEALTH_CHECK_PROSAO", "Health-check je pro" & ChrW(353) & "ao. Racunar je pode" & ChrW(353) & "en."
    UpsertRow lo, existing, "SETUP_MSG_HEALTH_CHECK_NASAO", "Health-check je na" & ChrW(353) & "ao stavke za proveru:"
    UpsertRow lo, existing, "SETUP_ERR_GRESKA_TOKOM_HEALTH", "Gre" & ChrW(353) & "ka tokom health-check-a:"
    UpsertRow lo, existing, "SETUP_MSG_APLIKACIJA_RADI_LOKALNO", "Aplikacija radi lokalno " & ChrW(8212) & " Google/PWA se ne koristi niti proverava."
    UpsertRow lo, existing, "SETUP_MSG_BANKARSKI_FOLDERI_PODESENI", "Bankarski folderi su pode" & ChrW(353) & "eni."
    UpsertRow lo, existing, "SETUP_ERR_GRESKA_PRI_PODESAVANJU", "Gre" & ChrW(353) & "ka pri pode" & ChrW(353) & "avanju bankarskih foldera:"
    UpsertRow lo, existing, "STM_ERR_GRESKA_PRI_EXPORTU", "Gre" & ChrW(353) & "ka pri exportu parcela:"
    UpsertRow lo, existing, "DOK_ERR_NEMA_AMBALAZE_NOVCA", "Nema ambala" & ChrW(382) & "e ni novca za " & ChrW(269) & "uvanje."
    UpsertRow lo, existing, "SEF_MSG_SENDING_FAKTURA_TRENUTNO", "- SENDING: Faktura se trenutno " & ChrW(353) & "alje."
    UpsertRow lo, existing, "SEF_MSG_SENT_FAKTURA_USPESNO", "- SENT: Faktura uspe" & ChrW(353) & "no primljena na SEF."
    UpsertRow lo, existing, "SEF_MSG_REJECTED_GRESKA_PROVERI", "- REJECTED: Gre" & ChrW(353) & "ka! Proveri 'Poslednja gre" & ChrW(353) & "ka'."
    UpsertRow lo, existing, "SEF_MSG_IZABERI_FAKTURU_LISTE", "Izaberi fakturu iz liste -> Klikni 'Po" & ChrW(353) & "alji na SEF'."
    UpsertRow lo, existing, "SEF_MSG_TEHNICKA_PODRSKA", "3. TEHNICKA PODR" & ChrW(352) & "KA:"
    UpsertRow lo, existing, "SEF_MSG_SVE_PROBLEME_KOJI", "Za sve probleme koji se ne re" & ChrW(353) & "avaju sa 'Osve" & ChrW(382) & "i status',"
    UpsertRow lo, existing, "SEF_MSG_KONTAKTIRAJ_ADMINISTRATORA_POSALJI", "kontaktiraj administratora i po" & ChrW(353) & "alji SEF Event Log (donja tabela)."
    UpsertRow lo, existing, "STM_MSG_PARCELA_NEMA_KATASTARSKI", "Parcela nema katastarski broj ili katastarsku op" & ChrW(353) & "tinu."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_OTVARANJU", "Gre" & ChrW(353) & "ka pri otvaranju GeoSrbije. Pogledaj log."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_CUVANJU", "Gre" & ChrW(353) & "ka pri " & ChrW(269) & "uvanju geo podataka. Pogledaj log."
    UpsertRow lo, existing, "STM_MSG_KLIKNI_JOS_JEDNOM", "Klikni jo" & ChrW(353) & " jednom za brisanje geo podataka."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_BRISANJU", "Gre" & ChrW(353) & "ka pri brisanju geo podataka. Pogledaj log."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_UCITAVANJU", "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju koordinata. Pogledaj log."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_OTVARANJU_2", "Gre" & ChrW(353) & "ka pri otvaranju Google Maps. Pogledaj log."
    UpsertRow lo, existing, "STM_MSG_GRESKA_PRI_OTVARANJU_3", "Gre" & ChrW(353) & "ka pri otvaranju polygon editora. Pogledaj log."
    UpsertRow lo, existing, "FAK_ERR_STORNIRANA_FAKTURA_MOZE", "Stornirana faktura se ne mo" & ChrW(382) & "e " & ChrW(353) & "tampati kao aktivna faktura:"
    UpsertRow lo, existing, "LIC_MSG_SERVER_NEDOSTUPAN_OFFLINE", "Server nedostupan " & ChrW(8212) & " offline grace (masina je vezana)."
    UpsertRow lo, existing, "LIC_MSG_PROPUSTAM_VEZANU_MASINU", "' " & ChrW(8212) & " propustam vezanu masinu."
    UpsertRow lo, existing, "MATICNI_MSG_PODESAVANJA", "Pode" & ChrW(353) & "avanja"
    UpsertRow lo, existing, "MONITO_ERR_NAMERNO_TESTIRANA_VBA", "Namerno testirana VBA gre" & ChrW(353) & "ka bez sensitive podataka"
    UpsertRow lo, existing, "CFG_MSG_STAMPA_OTKUPNOG_LISTA", ChrW(352) & "tampa otkupnog lista"
    UpsertRow lo, existing, "CFG_MSG_STAMPA_PALETNOG_LISTA", ChrW(352) & "tampa paletnog lista"
    UpsertRow lo, existing, "CFG_MSG_STAMPA_PRIJEMNICE_AUTO", ChrW(352) & "tampa prijemnice (auto na unos)"
    UpsertRow lo, existing, "CFG_MSG_KUPAC_UNOSI_BRUTO", "Kupac unosi BRUTO te" & ChrW(382) & "inu (oduzmi ambala" & ChrW(382) & "u)"
    UpsertRow lo, existing, "CFG_MSG_MALINA_REZIM", "Malina re" & ChrW(382) & "im"
    UpsertRow lo, existing, "CFG_MSG_PODESAVANJA_ZATVORENA", "Pode" & ChrW(353) & "avanja zatvorena."
    UpsertRow lo, existing, "CFG_MSG_PRIKAZI_CONFIG_TABELU", "Prika" & ChrW(382) & "i config tabelu"
    UpsertRow lo, existing, "SETUP_MSG_OVAJ_RACUNAR_PROSAO", "Da li je ovaj racunar pro" & ChrW(353) & "ao SetupNewPC"
    UpsertRow lo, existing, "SETUP_MSG_DATUM_VREME_ZAVRSETKA", "Datum i vreme zavr" & ChrW(353) & "etka setup-a"
    UpsertRow lo, existing, "SETUP_MSG_FOLDER_BANKARSKE_IZVODE", "Folder za bankarske izvode sa gre" & ChrW(353) & "kom"
    UpsertRow lo, existing, "SETUP_MSG_SEF_BASE_URL", "- SEF_BASE_URL nije pode" & ChrW(353) & "en."
    UpsertRow lo, existing, "SETUP_MSG_SEF_API_KEY", "- SEF_API_KEY nije pode" & ChrW(353) & "en."
    UpsertRow lo, existing, "SETUP_MSG_SEF_ENV_NIJE", "- SEF_ENV nije pode" & ChrW(353) & "en."
    UpsertRow lo, existing, "SETUP_MSG_NIJE_PODESEN", "nije pode" & ChrW(353) & "en:"
    UpsertRow lo, existing, "SETUP_MSG_FIRSTRUN_PONUDA", "Ovaj ra" & ChrW(269) & "unar jo" & ChrW(353) & " nije pode" & ChrW(353) & "en. " & ChrW(381) & "elite li da pokrenete pode" & ChrW(353) & "avanje (SetupNewPC) sada?"
    ' --- Self-update (modSelfUpdate); dijakritika preko ChrW ---
    UpsertRow lo, existing, "SU_AZURIRATI_SADA", "A" & ChrW(382) & "urirati sada? (preporu" & ChrW(269) & "eno)"
    UpsertRow lo, existing, "SU_BACKUP_NIJE", "Backup pre a" & ChrW(382) & "uriranja nije uspeo."
    UpsertRow lo, existing, "SU_PREUZIMANJE_OTKAZANO", "Preuzimanje nije uspelo (0 fajlova). A" & ChrW(382) & "uriranje otkazano."
    UpsertRow lo, existing, "SU_POKUSAJTE", "Poku" & ChrW(353) & "ajte ponovo ili preuzmite novu verziju ru" & ChrW(269) & "no."
    UpsertRow lo, existing, "SU_ZAVRSENO_FAJLOVA", "A" & ChrW(382) & "uriranje zavr" & ChrW(353) & "eno. Preuzeto fajlova: "
    UpsertRow lo, existing, "SU_ZAVRSENO_PREUZETO", "A" & ChrW(382) & "uriranje zavr" & ChrW(353) & "eno. Preuzeto: "
    UpsertRow lo, existing, "SU_GRESKA_AZURIRANJE", "Gre" & ChrW(353) & "ka pri a" & ChrW(382) & "uriranju: "
    UpsertRow lo, existing, "SU_GRESKA_2FAZA", "Gre" & ChrW(353) & "ka u 2. fazi a" & ChrW(382) & "uriranja: "
    UpsertRow lo, existing, "SU_TRUST", "Nema programskog pristupa VBA projektu." & vbCrLf & vbCrLf & "Uklju" & ChrW(269) & "i: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & "Macro Settings > 'Trust access to the VBA project object model'."
    ' --- Auth / korisnici / PIN hash / audit (modAuth, modSetup, frmStammdaten, modMaticniLookups) ---
    UpsertRow lo, existing, "AUTH_MSG_SAMO_ADMIN_PRIJAVA", "Samo administrator mo" & ChrW(382) & "e da menja prijavu."
    UpsertRow lo, existing, "AUTH_MSG_SAMO_ADMIN_ISKLJUCI", "Samo administrator mo" & ChrW(382) & "e da isklju" & ChrW(269) & "i prijavu."
    UpsertRow lo, existing, "AUTH_MSG_SAMO_ADMIN_KREIRA", "Samo administrator mo" & ChrW(382) & "e da kreira korisnike."
    UpsertRow lo, existing, "AUTH_MSG_SAMO_ADMIN_PINHASH", "Samo administrator mo" & ChrW(382) & "e da menja PIN hash."
    UpsertRow lo, existing, "AUTH_MSG_SAMO_ADMIN_KORISNICI", "Samo administrator mo" & ChrW(382) & "e da upravlja korisnicima."
    UpsertRow lo, existing, "AUTH_MSG_NEMA_ADMINA", "Ne mogu da uklju" & ChrW(269) & "im prijavu: ne postoji nijedan aktivan Admin." & vbCrLf & "Prvo pokreni: Alt+F8 -> KreirajPrvogAdmina."
    UpsertRow lo, existing, "AUTH_ERR_NE_MOGU_UKLJUCITI", "Ne mogu da uklju" & ChrW(269) & "im prijavu: "
    UpsertRow lo, existing, "AUTH_MSG_PRIJAVA_UKLJUCENA", "Prijava korisnika je UKLJU" & ChrW(268) & "ENA." & vbCrLf & "Pri slede" & ChrW(263) & "em otvaranju aplikacije tra" & ChrW(382) & "i se prijava."
    UpsertRow lo, existing, "AUTH_MSG_PRIJAVA_ISKLJUCENA", "Prijava korisnika je ISKLJU" & ChrW(268) & "ENA (app radi bez prijave)."
    UpsertRow lo, existing, "AUTH_MSG_VEC_POSTOJI", "' ve" & ChrW(263) & " postoji."
    UpsertRow lo, existing, "AUTH_MSG_ADMIN_KREIRAN", "' je kreiran." & vbCrLf & vbCrLf & "Slede" & ChrW(263) & "e: uklju" & ChrW(269) & "i prijavu sa Alt+F8 -> EnableAuth."
    UpsertRow lo, existing, "AUTH_ERR_SHA_NE_RADI", "SHA-256 ne radi u ovom okru" & ChrW(382) & "enju (pokreni Alt+F8 -> TestPinHash)." & vbCrLf & "PIN hash NIJE uklju" & ChrW(269) & "en."
    UpsertRow lo, existing, "AUTH_MSG_PINHASH_UKLJUCEN", "PIN hash je UKLJU" & ChrW(268) & "EN." & vbCrLf & "Postoje" & ChrW(263) & "i plaintext PIN-ovi migriraju na hash pri prvoj prijavi;" & vbCrLf & "novi/izmenjeni PIN-ovi se odmah " & ChrW(269) & "uvaju kao hash."
    UpsertRow lo, existing, "AUTH_MSG_PINHASH_ISKLJUCEN", "PIN hash je ISKLJU" & ChrW(268) & "EN (novi PIN-ovi se " & ChrW(269) & "uvaju kao plaintext)."
    UpsertRow lo, existing, "AUTH_MSG_PRIJAVA_NEUSPESNA", "Prijava neuspe" & ChrW(353) & "na. Aplikacija se zatvara."
    UpsertRow lo, existing, "AUTH_LBL_PRIJAVA_GRESKA", "Pogre" & ChrW(353) & "no korisni" & ChrW(269) & "ko ime ili PIN."
    UpsertRow lo, existing, "AUTH_MSG_PINHASH_RADI", "PIN hash (SHA-256) RADI ispravno u ovom okru" & ChrW(382) & "enju." & vbCrLf & "Mo" & ChrW(382) & "e" & ChrW(353) & " uklju" & ChrW(269) & "iti: Alt+F8 -> EnablePinHash."
    UpsertRow lo, existing, "AUTH_MSG_PINHASH_NE_RADI", "PIN hash NE radi u ovom okru" & ChrW(382) & "enju (SHA-256 nedostupan)." & vbCrLf & "NE uklju" & ChrW(269) & "uj PIN hash " & ChrW(8212) & " ostani na plaintext PIN-u." & vbCrLf & "Dobijeno: '"
    UpsertRow lo, existing, "AUD_MSG_KOLONE_SUFIX", " tabela:" & vbCrLf & "CreatedAt / CreatedBy / ModifiedAt / ModifiedBy." & vbCrLf & vbCrLf & "Od sada se svaki unos i izmena bele" & ChrW(382) & "e (ko i kada)."
    UpsertRow lo, existing, "GEN_ERR_GRESKA", "Gre" & ChrW(353) & "ka: "
    UpsertRow lo, existing, "KOR_LBL_KORISNICKO_IME", "Korisni" & ChrW(269) & "ko ime"
    UpsertRow lo, existing, "KOR_MSG_UNESITE_IME", "Unesite korisni" & ChrW(269) & "ko ime!"
    UpsertRow lo, existing, "KOR_MSG_VEC_POSTOJI", "' ve" & ChrW(263) & " postoji!"
    InvalidateCache
End Sub

Private Sub UpsertRow(lo As ListObject, existing As Object, _
                      ByVal kljuc As String, ByVal tekst As String)
    If existing.Exists(kljuc) Then
        If CStr(lo.ListRows(existing(kljuc)).Range(2).Value) <> tekst Then
            lo.ListRows(existing(kljuc)).Range(2).Value = tekst
        End If
    Else
        Dim r As ListRow
        Set r = lo.ListRows.Add
        r.Range(1).Value = kljuc
        r.Range(2).Value = tekst
        existing(kljuc) = r.Index
    End If
End Sub
