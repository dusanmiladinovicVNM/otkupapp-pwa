Attribute VB_Name = "modUiScreens"
'=====================================================================
' modUiScreens - REGISTAR EKRANA i ugovor prema njima (faza S3a).
'
' Jedno mesto koje zna koji ekrani postoje, kako se zovu, u koju grupu
' sidebara idu i koja im je oblast za dozvole. Sidebar se gradi odavde,
' ne iz nabrojanih nizova u BuildNav.
'
' KLJUCNO - LJUSKA NE SME DA ZNA NIJEDAN EKRAN PO IMENU. Svi pozivi ka
' ekranskom modulu idu kroz Application.Run, dakle KASNO VEZANO. Da je
' vezivanje rano, klijent kome jedan ekranski modul nedostaje ne bi se
' kompajlirao i pao bi ceo StartApp - ista klasa greske kao zamka #19 u
' CLAUDE.md. Ovako ekran koji nedostaje samo bude prikazan kao neaktivan.
'
' UGOVOR EKRANA (sve je opciono osim Scr_Meta):
'   Scr_Meta()                 -> opis ekrana; sluzi i kao provera da modul
'                                 postoji i da je ekran spreman
'   Scr_Build(z)               -> izgradi kontrole u svoju zonu (jednom)
'   Scr_Layout(z, w, h)        -> rasporedi; vraca zauzetu visinu
'   Scr_Rows(filter, q)        -> Array(kolone, redovi, n, zbirKg, zbirVal)
'                                 za DELJENU mrezu ljuske
'   Scr_Liste()                -> prekidac lista ekrana; niz redova
'                                 "KLJUC|natpis|naslov mreze|sirina"
'   Scr_Lista()                -> kljuc aktivne liste
'   Scr_Radnje()               -> radnje nad redom za AKTIVNU listu; redovi
'                                 "kljuc:natpis:sirina:stil:trebaRed" spojeni "|"
'   Scr_Event(tag, ev)         -> obradi klik; True ako je obradio
'   Scr_Dozvoljen()            -> dodatna brana ekrana (npr. administracija);
'                                 ekran koji je nema je dozvoljen
'   Scr_Save()                 -> upisi; "" ako je proslo, inace greska
'   Scr_ResetCache()           -> zaboravi izvedene mape (posle upisa)
'
' Od S4b SVI ekrani daju svoje redove kroz Scr_Rows - i ekran dokumenata,
' koji je do tada jedini punio mrezu po svom. Mreza je time neutralna.
' Ostali ekrani su u registru sa modulom koji jos ne postoji - sidebar ih
' prikazuje prigusene.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UISCR_BUILD As String = "v6-ui-187"

' Redosled polja u redu registra
Public Const SCR_KLJUC   As Long = 0
Public Const SCR_MODUL   As Long = 1
Public Const SCR_NASLOV  As Long = 2
Public Const SCR_IKONICA As Long = 3
Public Const SCR_GRUPA   As Long = 4
Public Const SCR_OBLAST  As Long = 5
' SEKCIJA ljuske -- koji SKUP stavki sidebara nosi ovaj ekran. Uvedena zato sto
' sidebar NEMA SKROL: na MIN_H (620) ostaje 492 pt za stavke, zauzeto je 405, pa
' bi maticne stavke (5 stavki + 2 grupe = 199 pt) ispale ispod profila -- tiho.
' Dva skupa se nikad ne crtaju zajedno; zlatno dugme u zaglavlju ih menja.
' V. docs/UI_MIGRACIJA_KATALOG.md, 24.1 i 24.2.
Public Const SCR_SEKCIJA As Long = 6

' Redosled polja u redu GRUPE sidebara
Public Const SCRG_KLJUC   As Long = 0
Public Const SCRG_NASLOV  As Long = 1
Public Const SCRG_SEKCIJA As Long = 2

Public Const SEK_RAD     As String = "RAD"
Public Const SEK_MATICNI As String = "MATICNI"

' Poslednja greska iz ekrana. Omotaci guse greske (ekran koji padne ne sme da
' obori aplikaciju), ali gusenje BEZ TRAGA znaci da dugme "ne radi" i niko ne
' zna zasto - tacno to se desilo sa "Po datumu". Ljuska ovo cita i prikazuje.
Public ScrLastErr As String

' Kes odgovora "da li modul postoji" - Application.Run na nepostojeci modul
' baca gresku, a to je skupo raditi pri svakom crtanju sidebara.
Private mHas As Object
' Kes odgovora "ima li ovaj ekran svoj Scr_Dozvoljen". Cuva se cinjenica o
' POSTOJANJU funkcije, ne njen odgovor -- v. ScrSopstvenaBrana.
Private mBrana As Object

'------------------------------------------------------------ REGISTAR
' kljuc | modul | naslov (kljuc kataloga) | MDL2 kod | grupa | oblast | sekcija
Public Function ScrRows() As Variant
    Dim c As Collection: Set c = New Collection
    c.Add "DOKUMENTI|modScrDokumenti|OTKUI_NAV_UNOS|" & IC_OTKUP & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    c.Add "PALETE|modScrPalete|OTKUI_NAV_PALETE|" & IC_PALETE & _
          "|OPERACIJE|" & OBL_PALETE & "|" & SEK_RAD
    ' Storno je do v6-ui-141 bio rezim F8 unosnog ekrana. Zaseban ekran zato sto
    ' NIJE unos: forma i "Sacuvaj" mu ne pripadaju (Scr_Save je za njega padao u
    ' Case Else), a pregled posledica pre odluke trazi svoju zonu -- cetiri moda,
    ' lanac, palete i blokovi ne staju u MsgBox. Oblast je OBL_DOKUMENTA: ko sme
    ' da unese dokument, sme i da ga stornira.
    c.Add "STORNO|modScrStorno|OTKUI_NAV_STORNO|" & IC_STORNO & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    ' Oporavak stoji uz Dokumenta i po oblasti prava: sve sto radi je
    ' prevezivanje i vracanje DOKUMENATA, pa ko sme da ih unosi sme i da ih
    ' popravi. Zaseban ekran, a ne jos jedan rezim, jer ovo nisu dokumenti nego
    ' POSAO koji ceka - i ne bira se po tipu nego po problemu.
    c.Add "OPORAVAK|modScrOporavak|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    c.Add "AGRO|modScrAgro|OTKUI_NAV_AGRO|" & IC_AGRO & _
          "|OPERACIJE|" & OBL_AGROHEMIJA & "|" & SEK_RAD
    c.Add "FAKTURE|modScrFakture|OTKUI_NAV_FAKT|" & IC_FAKT & _
          "|FINANSIJE|" & OBL_FAKTURISANJE & "|" & SEK_RAD
    c.Add "BANKA_UVOZ|modScrBankaUvoz|OTKUI_NAV_BANKA_UVOZ|" & IC_UVOZ & _
          "|FINANSIJE|" & OBL_BANKA & "|" & SEK_RAD
    c.Add "BANKA_NALOZI|modScrBankaNalozi|OTKUI_NAV_BANKA_NALOZI|" & IC_NALOZI & _
          "|FINANSIJE|" & OBL_BANKA & "|" & SEK_RAD
    c.Add "MARZA|modScrMarza|OTKUI_NAV_MARZA|" & IC_MARZA & _
          "|FINANSIJE|" & OBL_MARZA & "|" & SEK_RAD
    c.Add "IZVESTAJI|modScrIzvestaji|OTKUI_NAV_IZVESTAJI|" & IC_IZVEST & _
          "|ANALITIKA|" & OBL_IZVESTAJI & "|" & SEK_RAD
    c.Add "SLEDLJIVOST|modScrSledljivost|OTKUI_NAV_SLEDLJIVOST|" & IC_SLEDLJ & _
          "|ANALITIKA|" & OBL_SLEDLJIVOST & "|" & SEK_RAD

    ' ---- SEKCIJA MATICNI ------------------------------------------------
    ' Ono sto danas stoji iza zlatnog dugmeta: frmMaticniPodaci (popup meni,
    ' 16 sekcija u 4 grupe) -> frmStammdaten. Granica ekrana je LEGACY GRUPA,
    ' jer MAX_SEG (11) ne prima svih 13 sekcija sa podacima u jedan ekran, a
    ' grupisanje je u modMaticniLookups uvedeno svesno i operater ga zna.
    ' Plan i obrazlozenje: docs/UI_MIGRACIJA_KATALOG.md, 24.3.
    '
    ' Tri ekrana sa podacima jos nemaju modul -- sidebar ih zato crta prigusene,
    ' isto kao MARZA i SLEDLJIVOST. To NIJE propust nego zatecen nacin da se vidi
    ' sta dolazi (v. komentar na vrhu modula).
    c.Add "MAT_PARTNERI|modScrMatPartneri|OTKUI_NAV_MAT_PARTNERI|" & IC_MAT_PARTNERI & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_ROBA|modScrMatRoba|OTKUI_NAV_MAT_ROBA|" & IC_MAT_ROBA & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_PAKOVANJE|modScrMatPakovanje|OTKUI_NAV_MAT_PAKOVANJE|" & IC_MAT_PAKOVANJE & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    ' Korisnici i Sistem nose JOS JEDNU branu preko oblasti -- administraciju.
    ' Ona ne moze u SCR_OBLAST (to je naziv kolone prava u tblKorisnici), pa
    ' ekran sam odgovara kroz neobavezan Scr_Dozvoljen (v. ScrDozvoljen nize).
    c.Add "MAT_KORISNICI|modScrMatKorisnici|OTKUI_NAV_MAT_KORISNICI|" & IC_MAT_KORISNICI & _
          "|SISTEM|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_SISTEM|modScrMatSistem|OTKUI_NAV_MAT_SISTEM|" & IC_MAT_SISTEM & _
          "|SISTEM|" & OBL_MATICNI & "|" & SEK_MATICNI

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    ScrRows = a
End Function

' Grupe sidebara, redom. Naslov grupe je kljuc kataloga, trece polje je SEKCIJA.
' Redosled ovde je i redosled crtanja: prvo sve grupe radne sekcije, pa maticne.
Public Function ScrGroups() As Variant
    ScrGroups = Array("OPERACIJE|OTKUI_NAVG_OPERACIJE|" & SEK_RAD, _
                      "FINANSIJE|OTKUI_NAVG_FINANSIJE|" & SEK_RAD, _
                      "ANALITIKA|OTKUI_NAVG_ANALITIKA|" & SEK_RAD, _
                      "SIFARNICI|OTKUI_NAVG_SIFARNICI|" & SEK_MATICNI, _
                      "SISTEM|OTKUI_NAVG_SISTEM|" & SEK_MATICNI)
End Function

' Sekcija reda registra ili reda grupe. Prazno polje znaci RAD -- red koji je
' ostao bez sedmog polja (star export, rucna izmena) ne sme da nestane iz
' sidebara nego se ponasa kao pre uvodjenja sekcija.
Public Function SekcijaIli(ByVal s As String) As String
    SekcijaIli = SEK_RAD
    If Len(Trim$(s)) > 0 Then SekcijaIli = Trim$(s)
End Function

Public Function ScrSekcija(ByVal row As String) As String
    ScrSekcija = SekcijaIli(ScrField(row, SCR_SEKCIJA))
End Function

Public Function ScrGrupaSekcija(ByVal grpRow As String) As String
    ScrGrupaSekcija = SekcijaIli(ScrField(grpRow, SCRG_SEKCIJA))
End Function

' Prvi ekran sekcije na koji korisnik SME i koji postoji. Prazno = nijedan
' (npr. maticni ekrani pre nego sto im modul stigne). Ljuska ovim ne saznaje
' nijedan kljuc unapred -- pita registar i dobija onaj koji je dostupan.
Public Function ScrPrviUSekciji(ByVal sekcija As String) As String
    Dim r As Variant, kljuc As String
    For Each r In ScrRows()
        If ScrSekcija(CStr(r)) = sekcija Then
            kljuc = ScrField(CStr(r), SCR_KLJUC)
            If ScrAktivan(kljuc) Then
                ScrPrviUSekciji = kljuc
                Exit Function
            End If
        End If
    Next r
End Function

Public Function ScrField(ByVal row As String, ByVal idx As Long) As String
    Dim p() As String
    p = Split(row, "|")
    If idx > UBound(p) Then Exit Function
    ScrField = p(idx)
End Function

Public Function ScrRowByKey(ByVal kljuc As String) As String
    Dim r As Variant
    For Each r In ScrRows()
        If ScrField(CStr(r), SCR_KLJUC) = kljuc Then
            ScrRowByKey = CStr(r)
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- STANJE
' Postoji li modul ekrana i odgovara li na ugovor. Provera je namerno
' pokusaj poziva, ne pretraga po projektu: VBComponents trazi programski
' pristup VBA projektu, koji na korisnickoj masini cesto nije ukljucen.
Public Function ScrPostoji(ByVal kljuc As String) As Boolean
    Dim modul As String, dummy As Variant
    If mHas Is Nothing Then Set mHas = CreateObject("Scripting.Dictionary")
    If mHas.Exists(kljuc) Then
        ScrPostoji = mHas(kljuc)
        Exit Function
    End If
    modul = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(modul) = 0 Then Exit Function
    On Error Resume Next
    dummy = Application.Run(modul & ".Scr_Meta")
    ScrPostoji = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    mHas(kljuc) = ScrPostoji
End Function

' Ekran je dostupan ako postoji I ako korisnik ima pravo na njegovu oblast I
' ako sam ekran ne kaze da nije.
'
' TRECI uslov je uveden zbog maticnih sekcija Korisnici / Podesavanja / Admin:
' one traze ADMINISTRACIJU, a to nije oblast iz tblKorisnici pa ne moze da stane
' u SCR_OBLAST. Legacy je istu branu drzao u modMaticniLookups.MaticniMenu_OnClick;
' ovde je pita EKRAN, kroz neobavezan Scr_Dozvoljen -- ljuska i dalje ne zna
' nijedan ekran po imenu. Ekran koji ga ne implementira se ponasa kao pre.
'
' Brana je i dalje samo UI brana; tvrde su u modAdmin/modPodesavanja ulaznim
' tackama i tamo ostaju.
Public Function ScrDozvoljen(ByVal kljuc As String) As Boolean
    Dim obl As String
    On Error Resume Next
    obl = ScrField(ScrRowByKey(kljuc), SCR_OBLAST)
    If Len(obl) = 0 Then
        ScrDozvoljen = True
    Else
        ScrDozvoljen = modAuth.KorisnikImaPravo(obl)
    End If
    If ScrDozvoljen Then ScrDozvoljen = ScrSopstvenaBrana(kljuc)
End Function

' Odgovor samog ekrana na pitanje "smem li da te otvorim". Ekran koji nema
' Scr_Dozvoljen (a to su svi osim maticnog Sistema) dobija True -- greska poziva
' znaci "nema takvu funkciju", ne "zabranjeno".
'
' Kesira se SAMO cinjenica "ima li ekran tu funkciju", nikad njen ODGOVOR:
' prava se menjaju zamenom operatera, pa kesiran odgovor bi ostavio otvoren
' ekran kome novi operater nema pristup. Isti obrazac i isti razlog kao mHas
' kod ScrPostoji -- bez kesa bi svaki crtez sidebara bacio i uhvatio po jednu
' gresku za SVAKI ekran koji tu funkciju nema.
Private Function ScrSopstvenaBrana(ByVal kljuc As String) As Boolean
    Dim m As String, v As Variant
    ScrSopstvenaBrana = True
    If mBrana Is Nothing Then Set mBrana = CreateObject("Scripting.Dictionary")
    If mBrana.Exists(kljuc) Then
        If Not mBrana(kljuc) Then Exit Function
    End If
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    v = Application.Run(m & ".Scr_Dozvoljen")
    mBrana(kljuc) = (Err.Number = 0)
    If Err.Number = 0 Then ScrSopstvenaBrana = CBool(v)
    Err.Clear
End Function

Public Function ScrAktivan(ByVal kljuc As String) As Boolean
    ScrAktivan = ScrPostoji(kljuc)
    If ScrAktivan Then ScrAktivan = ScrDozvoljen(kljuc)
End Function

' Kes se prazni pri gradnji ekrana - posle self-update-a moduli koji nisu
' postojali mogu da postoje.
Public Sub ScrResetCache()
    Set mHas = Nothing
    Set mBrana = Nothing
End Sub

'--------------------------------------------------------------- POZIV
' Omotaci ugovora. Svaki gura gresku u "nije uspelo" umesto da je propusti
' u ljusku: ekran koji padne ne sme da obori aplikaciju.
Public Function ScrMeta(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrMeta = CStr(Application.Run(m & ".Scr_Meta"))
End Function

Public Function ScrBuild(ByVal kljuc As String, ByVal z As Object) As Boolean
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Application.Run m & ".Scr_Build", z
    ScrBuild = (Err.Number = 0)
End Function

Public Function ScrLayout(ByVal kljuc As String, ByVal z As Object, _
                          ByVal w As Single, ByVal h As Single) As Single
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrLayout = CSng(Application.Run(m & ".Scr_Layout", z, w, h))
End Function

Public Function ScrGrid(ByVal kljuc As String) As Variant
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrGrid = Application.Run(m & ".Scr_Grid")
End Function

Public Function ScrEvent(ByVal kljuc As String, ByVal tag As String, _
                         ByVal ev As String) As Boolean
    Dim m As String
    ScrLastErr = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrEvent = CBool(Application.Run(m & ".Scr_Event", tag, ev))
    If Err.Number <> 0 Then
        ScrLastErr = m & ".Scr_Event " & tag & " -> " & Err.Number & " " & Err.description
        Err.Clear
    End If
End Function

' Redovi za deljenu mrezu. Vraca Array(kolone, redovi, n, zbirKg, zbirVal)
' ili Empty ako ekran nema listu.
' NE zove se ScrRows - to ime vec nosi spisak redova REGISTRA, a dve funkcije
' istog imena u istom modulu su "Ambiguous name".
'
' GRESKA SE BELEZI, ne samo gusi. "On Error Resume Next" ovde mora da ostane --
' ekran koji padne ne sme da obori aplikaciju -- ali gusenje BEZ TRAGA je vec
' jednom skupo naplaceno ("Po datumu", v. ScrLastErr gore). Ponovilo se na cipu
' "Svi": Scr_Rows pukne, ovde se pretvori u Empty, LoadGridFromScreen na ne-niz
' radi Exit Sub -- pa mreza OSTANE na prethodnoj listi, sa prethodnim naslovom.
' Operater vidi dugme koje "ne radi", bez ijedne poruke i bez traga u logu.
Public Function ScrGridData(ByVal kljuc As String, ByVal filter As String, _
                            ByVal q As String) As Variant
    Dim m As String
    ScrLastErr = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrGridData = Application.Run(m & ".Scr_Rows", filter, q)
    If Err.Number <> 0 Then
        ScrLastErr = m & ".Scr_Rows -> " & Err.Number & " " & Err.description
        Err.Clear
    End If
End Function

' Cipovi AKTIVNE liste, oblika 'kljuc:KATALOG:sirina|...'. Opciono -- ekran
' koji ih nema vraca prazno i mreza se ne suzava nicim osim pretragom.
' Kljuc se vraca ljusci samo kao mFilter i putuje nazad u Scr_Rows; ljuska ga
' ne tumaci.
Public Function ScrCipovi(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    ScrCipovi = CStr(Application.Run(m & ".Scr_Cipovi"))
    Err.Clear
End Function

' Prekidac lista ekrana. Prazno = ekran ima samo jednu listu, pa prekidaca
' nema. Ljuska ne zna nijedan kljuc unapred - ni "OTPREMNICE" ni "PRERADE".
Public Function ScrListe(ByVal kljuc As String) As Variant
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrListe = Application.Run(m & ".Scr_Liste")
End Function

Public Function ScrLista(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrLista = CStr(Application.Run(m & ".Scr_Lista"))
End Function

' Radnje nad izabranim redom za trenutno aktivnu listu ekrana.
' Koliko stavki na ovom ekranu CEKA operatera. Opciono: ekran koji nema sta da
' broji je ne implementira i dobija nulu.
'
' Ljuska ovim ne saznaje NISTA o ekranu -- ne zna sta se broji ni zasto, samo
' dobija broj koji ce nacrtati uz stavku menija. Bez ovoga bi morala da zove
' GetNedovrseno po imenu, a to je tacno ono sto ceo ugovor izbegava.
Public Function ScrBrojac(ByVal kljuc As String) As Long
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrBrojac = CLng(Application.Run(m & ".Scr_Brojac"))
    If Err.Number <> 0 Then
        ScrBrojac = 0
        Err.Clear
    End If
End Function

' Broji li podnozje KOMADE umesto dinara na ovom ekranu.
'
' Ljuska je do sada pitala `ModeBrojiKomade(ActiveMode)` -- a ActiveMode je rezim
' UNOSA DOKUMENATA i na ugovornom ekranu ostaje onakav kakav ga je Dokumenta
' ostavila. Ko je bio na F7 (Reversi) pa presao na Uvoz izvoda, u podnozju je
' video "Ukupno 8.950 kom" umesto "Vrednost 8.950,00 RSD": novac izbrojan kao
' komadi, bez decimala. Ista klasa kao traka `zOtp` koja je ostajala upaljena.
'
' Ekran koji ovo ne implementira dobija False -- dinari. Ljuska ovim ne saznaje
' NISTA o ekranu, isto kao kod ScrBrojac.
Public Function ScrBrojiKomade(ByVal kljuc As String) As Boolean
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrBrojiKomade = CBool(Application.Run(m & ".Scr_BrojiKomade"))
    If Err.Number <> 0 Then
        ScrBrojiKomade = False
        Err.Clear
    End If
End Function

Public Function ScrRadnje(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrRadnje = CStr(Application.Run(m & ".Scr_Radnje"))
End Function

' Upis dokumenta. Ljuska predaje RECNIK vrednosti pod logickim imenima; ekran
' zna sta su i sta se sa njima radi. Vraca "" kad je proslo, inace poruku za
' operatera (jedan razmak = operater je odustao, ne prikazuje se nista).
' Ekran u isti recnik upisuje "fokus", "rezultat" i "poruke".
Public Function ScrSave(ByVal kljuc As String, ByVal polja As Object) As String
    Dim m As String
    ScrLastErr = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrSave = CStr(Application.Run(m & ".Scr_Save", polja))
    If Err.Number <> 0 Then
        ScrLastErr = m & ".Scr_Save -> " & Err.Number & " " & Err.description
        Err.Clear
    End If
End Function
