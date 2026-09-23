Attribute VB_Name = "modAutoHladnjaca"
Option Explicit

' ============================================================
' modAutoHladnjaca -- oznake hladnjace (stanica, kupac) i relink stanje.
'
' S3b-1: AUTO-LANAC JE OBRISAN. AutoChainHladnjaca je od Otkup cutover-a bio
' pauziran (modOtkupUnos ga ne zove), a pravio je lanac PO KLASI kroz stari
' pisac otpremnice (SaveOtpremnica_TX) i vezu Otkup.OtpremnicaID -- oba su
' obrisana. Sa njim su otisli i backfill prijemnica hladnjace (F-090, u mapi
' "nije sposobnost") i test seam ArmHladnjacaTestFail.
'
' S3d: LANAC SE VRACA U KORACIMA, nad kanonom. Prvi korak je otpremnica
' (AutoLanacHladnjaca ispod). Zbirna (S4) i prijemnica (S6) dodaju se u ISTU
' funkciju kad ti dokumenti predju na kanon -- ne u nov orkestrator.
'
' Stari lanac se NIJE prevodio: delio je dokument po klasi kroz SaveOtpremnica_TX
' i vezu Otkup.OtpremnicaID (oba obrisana u S3b-1).
' ============================================================

' Ispravka autohladnjace: posle storna otkupa-hladnjace ceo lanac
' (otpremnica+zbirna+prijemnica) je oboren, a palete su OSIROCENE. Kad operater
' izabere "Uneti ispravku", ekran zapamti broj te (stornirane) prijemnice ovde.
' Relink aparat je NEDOSTIZAN dok auto-lanca nema (modOtkupUnos).
Private mPendingRelinkOldPrij As String

Public Sub SetHladnjacaRelinkPending(ByVal oldPrijBroj As String)
    mPendingRelinkOldPrij = Trim$(oldPrijBroj)
End Sub

Public Function GetHladnjacaRelinkPending() As String
    GetHladnjacaRelinkPending = mPendingRelinkOldPrij
End Function

' Da li je stanica oznacena kao hladnjaca (tblStanice.JeHladnjaca = "Da").
Public Function IsHladnjacaStanica(ByVal stanicaID As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, COL_STA_JE_HLADNJACA), ""))
    IsHladnjacaStanica = (StrComp(v, "Da", vbTextCompare) = 0)
End Function

' Da li je KUPAC oznacen kao hladnjaca-kupac (interni cold-store tok). Isti signal
' kao frmDokumenta.RefreshBrojPrijSuggestion: kupac == CFG_MALINA_DEFAULT_KUPAC.
' Eksterni kupci -> False (za njih je zbirna poslednji interni dokument, a prijemnica
' eksterna -> storno framework ne kaskadira nizvodni tok). Prazan config / prazan
' kupac -> False.
Public Function IsHladnjacaKupac(ByVal kupacID As String) As Boolean
    On Error Resume Next
    kupacID = Trim$(kupacID)
    If Len(kupacID) = 0 Then Exit Function
    Dim h As String
    h = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(h) = 0 Then Exit Function
    IsHladnjacaKupac = (StrComp(kupacID, h, vbTextCompare) = 0)
End Function


' Da li je hladnjacki auto-lanac UKLJUCEN. Jedan autoritet, jedno mesto.
'
' Stanica (JeHladnjaca) kaze KOJI blok ide u lanac; ovaj prekidac kaze DA LI
' lanac uopste radi. Do sada su postojala dva gospodara: prekidac u Podesavanjima
' koji nijedan red koda nije citao, i nov put koji ga je ignorisao.
'
' Default je OFF i tako ostaje do S6: lanac sme da se upali tek kad ume da
' zavrsi ceo OTK -> OTP -> ZBR -> PRJ. Polovican lanac je gori od nikakvog --
' zbirna i prijemnica se danas ne mogu doraditi ni rucno (F3/F4 su pauzirani),
' pa bi operater ostao sa izdatom otpremnicom i bez ijednog puta napred.
Public Function LanacUkljucen() As Boolean
    LanacUkljucen = IsAutoPrijemnicaHladnjaca()
End Function

' Da li blokovi OVE STANICE idu u auto-lanac. Ista odluka kao LanacVaziZaBlok,
' samo pre nego sto blok postoji -- ekran je treba PRE svih pravila vezanih za
' aktivan rucni nacrt (potvrda prekoracenja, vezivanje), jer hladnjacki blok tim
' pravilima uopste ne podleze.
'
' FAIL-CLOSED, isto kao LanacVaziZaBlok: "ne znam" nije "nije hladnjaca" nego
' razlog. Pozivalac koji dobije razlog ne sme da primeni pravila rucnog toka.
Public Function LanacVaziZaStanicu(ByVal stanicaID As String, _
                                   Optional ByRef outGreska As String) As Boolean
    Dim errDesc As String

    outGreska = ""
    On Error GoTo EH

    If Not LanacUkljucen() Then Exit Function

    stanicaID = Trim$(stanicaID)
    If Len(stanicaID) = 0 Then Exit Function

    LanacVaziZaStanicu = HladnjacaStrogo(stanicaID)
    Exit Function

EH:
    errDesc = Err.description
    LogErr "modAutoHladnjaca.LanacVaziZaStanicu"
    outGreska = Poruka("OTKUI_ERR_LANAC_PUT") & " " & errDesc
    LanacVaziZaStanicu = False
End Function

' Da li OVAJ blok ide u auto-lanac umesto na radni sto. Ekran ovim grana PRE
' rucnog vezivanja; ceo sud je ovde, da ga sledeci pozivalac ne prepisuje.
'
' FAIL-CLOSED. Ovo nije prikaz nego RAZVODNICA izmedju dva toka, a JeHladnjaca je
' tvrda poslovna granica: hladnjacki blok MORA u lanac. Zato postoje tri ishoda,
' ne dva:
'
'   sigurno hladnjaca       -> True
'   sigurno nije hladnjaca  -> False
'   ne moze da se utvrdi    -> False + RAZLOG (outGreska)
'
' Pozivalac koji dobije razlog ne sme da nastavi NI JEDNIM putem: gurnuti blok u
' rucni nacrt "jer provera nije uspela" znaci da hladnjacki blok tiho zavrsi kao
' clan tudjeg dokumenta -- tacno ono sto je ovaj PR isao da spreci. Blok je vec
' upisan svojom transakcijom i ostaje nevezan, vidljiv u listi "Bez otpremnice".
'
' Isto pravilo drzi modStornoDok.StornoTraziIzborModa: neizvesnost vodi ka VISE
' pitanja, ne ka manje.
Public Function LanacVaziZaBlok(ByVal otkupID As String, _
                                Optional ByRef outGreska As String) As Boolean
    Dim stanicaID As String, redovi As Collection, errDesc As String

    outGreska = ""
    On Error GoTo EH

    If Not LanacUkljucen() Then Exit Function

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    ' Blok mora da postoji tacno jednom -- inace se o njemu ne zna nista, pa ni
    ' kojim putem ide.
    Set redovi = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If redovi Is Nothing Then
        Err.Raise vbObjectError + 8460, "LanacVaziZaBlok", _
                  "Citanje otkupa nije uspelo: " & otkupID
    End If
    If redovi.count <> 1 Then
        Err.Raise vbObjectError + 8461, "LanacVaziZaBlok", _
                  "Otkup ne postoji tacno jednom (" & redovi.count & "): " & otkupID
    End If

    If StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO), "")), _
               "Da", vbTextCompare) = 0 Then Exit Function

    stanicaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA), ""))
    If Len(stanicaID) = 0 Then
        Err.Raise vbObjectError + 8462, "LanacVaziZaBlok", _
                  "Otkup nema otkupno mesto: " & otkupID
    End If

    ' IsHladnjacaStanica je fail-open (prikaz), pa se ovde NE koristi: stanica
    ' koja ne postoji davala bi "nije hladnjaca" i blok bi tiho otisao u rucni tok.
    LanacVaziZaBlok = HladnjacaStrogo(stanicaID)
    Exit Function

EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr "modAutoHladnjaca.LanacVaziZaBlok"
    outGreska = Poruka("OTKUI_ERR_LANAC_PUT") & " " & errDesc
    LanacVaziZaBlok = False
End Function

' Je li stanica hladnjaca -- STROGO. Stanica koja ne postoji ili postoji dvaput
' nije "nije hladnjaca" nego nepoznato stanje, i to se dize kao greska.
Private Function HladnjacaStrogo(ByVal stanicaID As String) As Boolean
    Dim redovi As Collection

    Set redovi = FindRows(TBL_STANICE, "StanicaID", stanicaID)
    If redovi Is Nothing Then
        Err.Raise vbObjectError + 8463, "HladnjacaStrogo", _
                  "Citanje stanica nije uspelo: " & stanicaID
    End If
    If redovi.count <> 1 Then
        Err.Raise vbObjectError + 8464, "HladnjacaStrogo", _
                  "Otkupno mesto ne postoji tacno jednom (" & redovi.count & "): " & stanicaID
    End If

    If GetColumnIndex(TBL_STANICE, COL_STA_JE_HLADNJACA) = 0 Then
        Err.Raise vbObjectError + 8465, "HladnjacaStrogo", _
                  "Kolona " & COL_STA_JE_HLADNJACA & " ne postoji u " & TBL_STANICE
    End If

    HladnjacaStrogo = (StrComp(Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, _
                                                    COL_STA_JE_HLADNJACA), "")), _
                               "Da", vbTextCompare) = 0)
End Function

' ============================================================
' AUTO-LANAC HLADNJACE -- korak otpremnice (A-014, S3d).
'
' Roba na hladnjackoj stanici se MERI PRI PRIJEMU u hladnjacu, pa otpremnica
' nosi tacno ono sto i otkupni list: kilaza i ambalaza su 1:1, bez odstupanja.
' Zato otpremnica nema sta da ceka -- nastaje i IZDAJE se odmah, iz tog jednog
' bloka, u jednoj transakciji (CreateOtpremnicaIzIzvora_TX izvodi ocekivanje iz
' izvora, pa je "povezano = ocekivano" zadovoljeno samim nastankom).
'
' Vozac je OGLEDALO stanice (VozacID = StanicaID), i to nije podatak o coveku
' nego PROXY za samu hladnjacu: robu do hladnjace kooperant dovozi SAM, pa taj
' prevoz nema firminog vozaca. Zato je ogledalo na hladnjackoj stanici OBAVEZNO
' bez obzira na rezim -- malina rezim je zaseban razlog, i tice se ogledala na
' OSTALIM stanicama.
'
' Lanac ga trazi kroz modMalina.VozacOgledaloZaStanicu -- JEDNO telo koje dele
' hladnjacki lanac i malina auto-otpremnica (S5-1). Pravilo je "odlucuje
' ponovljena provera para, ne povratna vrednost Ensure-a"; ako ogledala ni tada
' nema, lanac staje i kaze na kojoj stanici.
'
' AUTOMATIKA JE OBAVEZNA, NE PONUDA: hladnjacki blok ne ide na radni sto nego u
' SVOJ lanac. Zato ekran grana PRE rucnog vezivanja (LanacVaziZaBlok), a ne posle
' njega -- aktivan rucni nacrt bi inace progutao blok i lanac se ne bi ni pokrenuo,
' a blok bi zavrsio kao jedan od clanova tudjeg dokumenta sa sasvim drugom kilazom.
'
' Provera clanstva ispod je zato ZASTITA OD DUPLIRANJA (ponovljen poziv, retry),
' a ne nacin da rucni tok pobedi automatiku.
'
' Otkup je vec upisan svojom transakcijom. Ako lanac padne, blok OSTAJE -- ne
' brise se i ne stornira. Operater dobija razlog i blok vidi na radnom stolu
' (lista "Bez otpremnice"). Tiho preskakanje bi znacilo da misli da je lanac
' odradjen.
'
' OPORAVAK NIJE RUCNO VEZIVANJE: taj blok ne sme da zavrsi u obicnoj otpremnici
' (kapija je u modScrDokumenti.VeziZaAktivnu). Put je "otkloni uzrok pa PONOVI
' auto-lanac" -- radnja "Ponovi auto-lanac" u listi "Bez otpremnice".
'
' Vraca OtpremnicaID ("" = nije pravljena); `izvestaj` je ono sto operater cita.
' ============================================================
Public Function AutoLanacHladnjaca(ByVal otkupID As String, _
                                   ByRef izvestaj As String) As String
    Dim stanicaID As String, kulturaID As String, tipAmb As String, broj As String
    Dim datum As Date, greska As String, errDesc As String
    Dim h As Object, izvori As Collection

    izvestaj = ""
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function
    If Not LanacUkljucen() Then Exit Function

    If StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO), "")), _
               "Da", vbTextCompare) = 0 Then Exit Function

    stanicaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA), ""))
    ' STROGO, isti primitiv kao router: slabija kopija istog pravila (fail-open
    ' IsHladnjacaStanica) bila bi kandidat za drift -- jedan bi rekao "ne znam",
    ' drugi "nije hladnjaca".
    If Not HladnjacaStrogo(stanicaID) Then Exit Function

    ' TVRDA KAPIJA ZA S6 -- OVAJ RED SE MORA PREISPITATI PRE PALJENJA LANCA.
    '
    ' Dok je lanac samo OTK -> OTP, "otpremnica vec postoji" znaci "posao je
    ' gotov" i izlazak je tacan. Cim lanac dobije ZBR i PRJ, isti red postaje
    ' zamka: posle ishoda "OTP uspeo, ZBR pao" ponovljen poziv bi izasao ODMAH i
    ' lanac se nikad ne bi dovrsio.
    '
    ' Pre S6 se bira JEDNO:
    '   A) ceo lanac je jedna atomska transakcija (nema polovicnog stanja), ili
    '   B) lanac je nastavljiv: OTP postoji -> proveri/nastavi ZBR, ZBR postoji
    '      -> proveri/nastavi PRJ.
    ' Izbor je izlazni uslov S6 (plan 14.20), ne stvar ukusa.
    If Len(modDokumenta.OtpremnicaZaOtkup(otkupID)) > 0 Then Exit Function

    Dim vozacID As String
    vozacID = modMalina.VozacOgledaloZaStanicu(stanicaID)
    If Len(vozacID) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_MIRROR") & " " & stanicaID
        Exit Function
    End If

    datum = CDate(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM))
    kulturaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_KULTURA), ""))
    tipAmb = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_TIP_AMB), ""))

    ' Broj ide iz niza STANICE, kao i svaka druga otpremnica; mirror prefiks je
    ' pravilo broja ZBIRNE (modBrojevi.ApplyMirrorPrefix) i ovde se ne primenjuje.
    broj = modBrojevi.GenerateBrojOtpremnice(stanicaID, datum)
    If Len(broj) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_BROJ") & " " & stanicaID
        Exit Function
    End If

    Set h = CreateObject("Scripting.Dictionary")
    h("Datum") = datum
    h("StanicaID") = stanicaID
    h("VozacID") = vozacID
    h("KulturaID") = kulturaID
    h("TipAmbalaze") = tipAmb
    h("BrojOtpremnice") = broj

    Set izvori = New Collection
    izvori.Add otkupID

    AutoLanacHladnjaca = modDokumenta.CreateOtpremnicaIzIzvora_TX(h, izvori, greska)
    If Len(AutoLanacHladnjaca) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & greska
        Exit Function
    End If

    ' Zbirna (S4) i prijemnica (S6) se dodaju OVDE, i svaka izvodi stavke iz svog
    ' kanonskog roditelja -- ne prima prepisane brojeve. 1:1 je tada posledica
    ' modela, a ne tri kopirane vrednosti koje mogu da se raziju.
    izvestaj = Poruka("OTKUI_MSG_LANAC_OTP") & " " & broj
    Exit Function

EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr "modAutoHladnjaca.AutoLanacHladnjaca"
    izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & errDesc
    AutoLanacHladnjaca = ""
End Function
