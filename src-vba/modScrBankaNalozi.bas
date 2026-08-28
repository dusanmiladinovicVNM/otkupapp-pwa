Attribute VB_Name = "modScrBankaNalozi"
'=====================================================================
' modScrBankaNalozi - ekran "Platni nalozi" (v6-ui-185). Faza E.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19). Red u registru
' (modUiScreens.ScrRows) je postojao od S3a -- stavka menija se do sada crtala
' prigusena jer modula nije bilo. Registar se NE dira.
'
' ODAKLE DOLAZI: frmBankaExportPregled prikazuje otvorene otkup blokove,
' operater cekira podskup, bira racun firme i generise CSV naloga za prenos
' (uvoz u e-banking) i PDF specifikaciju isplata. Mreza ljuske bira JEDAN red,
' pa je multiselect postao KORPA ("U NALOZIMA") -- isti obrazac kao prijemnice
' na ekranu Fakturisanje.
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED i PRIKAZ, plus JEDNO pravilo
' unosa -- BnPostaviIznos (cent-domen, > 0, nikad preko otvorenog, jednako
' otvorenom brise zadato). To je preslikana kopija legacy
' txtIsplatiti_Exit pravila, po obrascu iz kataloga par. 5/Faza B (dve
' kopije zive namerno dok legacy ne ode); klamp, selekcija i kapije novca
' su u domenu. Sve ostalo -- nijedna kapija i nijedan upis -- nije ovde:
'   - otvoreni blokovi (+identitet)  -> modBankaExportPregled.BuildBlokIsplataList
'                                       (fail-closed: dupli/prazan OtkupID,
'                                       dupli KooperantID OBARAJU citanje)
'   - redovi mreze                   -> modBankaExportPregled.GetBlokIsplataForGrid
'   - cetiri brojke zone             -> modBankaExportPregled.NalogeKpi
'   - izbor blokova za izvoz         -> modBankaExportPregled.OdaberiBlokoveZaNaloge
'   - CSV nalozi (finalna kapija)    -> modBankaExportPregled.GenerisiNalogeCSV
'                                       (ValidateNalogSaldo NAD SVEZIM saldom)
'   - PDF specifikacija              -> modBankaExportPregled.PrintIsplataSpecifikacija
'   - racuni firme                   -> modBankaExportPregled.BankaNalogRacuniCSV
'                                       + BankaNazivZaRacun
'   - vezivanje avansa               -> modNovac.ApplyAvansToOtkup_TX
'
' JEDNA LISTA. Predlozena druga lista "RACUNI" je izbacena: sav njen sadrzaj
' (racuni firme + naziv banke) nosi combo "Sa racuna" u zoni, pa bi lista bila
' pregled bez ijednog posla nad redom. Cip "iznad praga" je izbacen jer prag
' ne postoji ni u legacy formi ni u configu -- uvodjenje bi bilo novo poslovno
' pravilo, a ne prelazak ekrana.
'
' IZNOS PO NALOGU: podrazumevano SVEZ OTVOREN IZNOS; radnjom "Iznos..."
' operater za blok zadaje MANJI (delimicna isplata -- smoke 28.08.2026 je
' pokazao da je to stvaran tok, ne izuzetak). Zadati iznosi su legacy
' "Isplatiti" override, sa ISTIM aparatom odrzavanja: Dictionary
' OtkupID -> iznos, pri svakom citanju liste usklajden sa svezim otvorenim
' kroz modBankaExportPregled.ClampOverridesToOpenDict (nestao/zatvoren blok
' brise, veci spusta, manji ostaje -- uz poruku, nikad tiho). Korpa i dalje
' NE NOSI iznos: clanstvo i iznos su dva recnika, oba klampovana svezim
' stanjem, pa zastareli snimak ne moze da narucuje uplatu. Unos ide kroz
' InputBox (presedan: SEF komentar na Fakturisanju) -- polje u zoni vezano
' za "izabran red" bi zastarevalo na svaki sort/stranu/filter mreze.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRBN_BUILD As String = "v6-ui-185"

' Visina zone: red polja (combo racuna) + hint + red dugmadi, kao Fakturisanje.
Private Const BN_ZONA_H   As Single = 148

Private Const BN_Y_CAP    As Single = 6
Private Const BN_Y_KPI_V  As Single = 18
Private Const BN_Y_LBL    As Single = 48
Private Const BN_Y_HINT   As Single = 98
Private Const BN_Y_BTN    As Single = 116
Private Const BN_BTN_H    As Single = 24
Private Const BN_KPI_W    As Single = 140
Private Const BN_FLD_W    As Single = 260

' Desna traka zone nosi korpu "U NALOZIMA" -- isti raspored kao Fakturisanje.
Private Const BN_KORPA_W  As Single = 300
Private Const BN_KORPA_N  As Long = 4
Private Const BN_POLJA_MIN As Single = 460

' Kljuc jedine liste.
Private Const BN_NALOZI As String = "NALOZI"

' SKRIVENE KOLONE, prioritet 4 -- LayoutGrid crta do 3, pa vrednost postoji u
' modelu a celija se nikad ne pravi. Identitet ide U RED, ne pored njega:
' mreza redove sortira i deli na strane, pa bi svaka mapa "prikaz -> ID" koju
' ekran drzi sa strane zastarela na prvi klik po zaglavlju.
'
' Identitet je OtkupID -- broj bloka NIJE identitet (jedinstven je samo po
' stanici: fixture ima isti broj na dva otkupna mesta), a ziro racun dele svi
' blokovi istog kooperanta. Dvosmislen OtkupID ovde NE stize do mreze: citac
' (BuildBlokIsplataList) na dupli/prazan OtkupID medju otvorenima OBARA celo
' citanje (AUD-026), sto je strozije od "prazan identitet po redu".
Private Const BN_KOL_ID As Long = 11
' Sta radnje moraju da znaju a iz prikaza se ne vidi jednoznacno (prazna
' celija racuna lici na kolonu koja se nije nacrtala; avans se ne vidi nigde):
Private Const BN_KOL_KOOP As Long = 12
Private Const BN_KOL_TR As Long = 13
Private Const BN_KOL_AVANS As Long = 14
' Otvoren iznos je VIDLJIVA kolona 8; GridCell vraca vrednost modela (Double),
' pa se odatle i cita -- ne izvodi se iz formatiranog prikaza.
Private Const BN_KOL_OTVORENO As Long = 8

Private mLista As String

' KORPA "U NALOZIMA": blokovi koje je operater pokazao za izvoz. Prolazno
' stanje ekrana, NE podatak u tabeli. Svaka stavka je recnik:
' otkupID / broj / otvoreno -- broj i otvoreno sluze SAMO prikazu u traci
' (osvezavaju se pri svakom citanju liste); iznos naloga se NIKAD ne cita
' odavde nego svez pri izvozu.
Private mKorpa As Collection

' ZADATI IZNOSI po bloku (OtkupID -> iznos), za delimicnu isplatu. Ista
' struktura kao legacy m_OverrideAmounts; odrzava je ISTI klamp
' (ClampOverridesToOpenDict) pri svakom citanju liste. Iznos postoji samo za
' blok koji je u korpi -- izbacivanje iz korpe ga brise.
Private mIznosi As Object

Private mFill As Boolean           ' punjenje comboa okida Change
Private mRacunPunjen As Boolean

' Kes cetiri brojke zone (modBankaExportPregled.NalogeKpi = pun prolaz kroz
' tabele, a OsveziZonu se zove pri svakom citanju mreze). Cisti ga
' Scr_ResetCache. NEUSPEH CITANJA NIJE NULA -- v. Kpi / BnKpiPosleGreske.
Private mKpi As Variant
Private mKpiOK As Boolean

' KES SNIMKA LISTE za prikaz i filtriranje -- legacy m_FullBlokovi obrazac
' (frmBankaExportPregled: "LAGANI re-filter nad vec ucitanom listom, bez
' citanja tabela"). Bez njega SVAKI otkucaj u pretrazi placa pun
' BuildBlokIsplataList (na 1.500+ otkupa vise sekundi po slovu), pa je na
' pravoj svesci pretraga delovala mrtvo iako je model bio tacan -- izmereno
' kroz Diag_BnRedovi (smoke 3: q stize, 38 redova vraceno i drzano, a
' operater vidi zamrznut ekran). Invalidira ga Scr_ResetCache, koji ljuska
' zove posle SVAKOG upisa (RefreshFromData) -- IZVOZ ovaj kes NE koristi:
' BlokoviZaIzvoz i finalna kapija citaju svez saldo, kao i do sada.
Private mSnimak As Variant
Private mSnimakOK As Boolean
' Broj stvarnih citanja tabela -- bez njega se kes ne moze izmeriti (isti
' razlog kao mCiljPunjenja na Uvozu izvoda: broji se POZIV, ne uspeh).
Private mSnimakPunjenja As Long

' Izbor koji je postavio TEST. Zone u testu nema (forma se ne prikazuje), pa
' se combo ne moze procitati. Vazi SAMO u test rezimu.
Private mRacunTest As String

' Poslednji poziv Scr_Rows -- SAMO za Diag_BnRedovi (smoke 3: "filter ne
' radi"). Presudjuje da li upit pretrage uopste stize ekranu, sta ekran
' vraca, i pod kojim cipom -- bez ovoga se gubitak upita PRE ekrana i kvar
' POSLE ekrana (prikaz) ne razlikuju.
Private mDiagFilter As String
Private mDiagQ As String
Private mDiagN As Long

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=BANKA_NALOZI|naslov=OTKUI_NAV_BANKA_NALOZI|sub=OTKUI_SCRBN_SUB" & _
               "|lista=OTKUI_SCRBN_LISTA|oblik=zona+mreza|upis=zona"
End Function

Public Function Scr_Liste() As Variant
    Scr_Liste = Array(BN_NALOZI & "|OTKUI_SEG_BN_NALOZI|OTKUI_GRID_TITLE_BN_NALOZI|92")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = BN_NALOZI
    Scr_Lista = mLista
End Function

' Prvi cip je svuda "sve" -- ljuska na njega pada kad zatecen filter ne pripada
' listi (RefreshChipsForScreen). Zato prvi mora da bude NAJSIRI: povratak na
' uzi cip bi tiho sakrio redove.
Public Function Scr_Cipovi() As String
    Scr_Cipovi = BnCipoviZaListu(Scr_Lista())
End Function

' Cipovi PO KLJUCU LISTE -- da se ugovor moze izmeriti bez stanja ekrana.
Public Function BnCipoviZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case BN_NALOZI
            BnCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "imarac:OTKUI_CIPN_IMARAC:88|" & _
                              "bezrac:OTKUI_CIPN_BEZRAC:88|" & _
                              "avans:OTKUI_CIPN_AVANS:76"
    End Select
End Function

' PRAVILO CIPA. Kljuc je EKRANOV -- ljuska ga je samo vratila onakvog kakvog
' ga je dobila iz Scr_Cipovi. Nepoznat i prazan kljuc PUSTAJU sve.
'
' "Bez racuna" postoji zato sto takav blok NE MOZE u CSV (nema primaoca) --
' operater mora da vidi kome pre isplate treba upisati tekuci racun u maticne
' podatke. "Avans" pokazuje blokove ciji kooperant ima nerasporedjen avans:
' njih pre naloga treba vezati (radnja "Primeni avans"), inace se plati i ono
' sto je avansom vec pokriveno.
Public Function BnCipNalog(ByVal filter As String, ByVal imaTR As Boolean, _
                           ByVal avans As Double) As Boolean
    Select Case filter
        Case "imarac": BnCipNalog = imaTR
        Case "bezrac": BnCipNalog = Not imaTR
        Case "avans":  BnCipNalog = (avans > 0)
        Case Else:     BnCipNalog = True
    End Select
End Function

' Radnje nad izabranim redom. Tri od MAX_ACT (5). "Primeni avans" je jedino
' KNJIZENJE sa ovog ekrana (tblNovac, kroz postojecu transakciju) -- CSV i
' specifikacija ne pisu tabele, pa zive kao dugmad zone, ne kao radnje.
Public Function Scr_Radnje() As String
    Scr_Radnje = BnRadnjeZaListu(Scr_Lista())
End Function

' Tacno PET radnji -- granica bazena (MAX_ACT); sesta bi se tiho odsekla.
' Peta ("Svi sa racunom", trebaRed=0) je IZRICITI nacin da se izvezu svi:
' prazan izbor to vise ne znaci -- v. BlokoviZaIzvoz.
Public Function BnRadnjeZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case BN_NALOZI
            BnRadnjeZaListu = "bnadd:OTKUI_BTN_BN_UNALOG:96:primary:1|" & _
                              "bniznos:OTKUI_BTN_BN_IZNOS:80:soft:1|" & _
                              "bndel:OTKUI_BTN_BN_IZNALOG:80:ghost:1|" & _
                              "bnavans:OTKUI_BTN_BN_AVANS:112:soft:1|" & _
                              "bnsve:OTKUI_BTN_BN_SVE:116:ghost:0"
    End Select
End Function

' Znacka uz stavku menija: koliko otvorenih blokova ceka isplatu. To je
' PODATAK U TABELI (tblOtkup + tblNovac), ne prolazno stanje -- svaka promena
' nastaje upisom (uvoz izvoda, avans, storno), a ljuska posle upisa ionako
' zove RefreshFromData, pa je brojac time pokriven i privatan kanal ka
' OsveziNavBrojace NE treba. Korpa "U NALOZIMA" se NE broji u znacki: ona je
' izbor za izvoz, a ne posao koji ceka -- i nista se ne gubi ako se ekran
' napusti sa punom korpom (izvoz je eksplicitan, nema tihog knjizenja).
Public Function Scr_Brojac() As Long
    Dim k As Variant
    k = Kpi()
    Scr_Brojac = CLng(k(0))
End Function

Public Sub Scr_ResetCache()
    ' mKpi se NE brise, samo proglasava zastarelim -- v. Kpi.
    mKpiOK = False
    ' Snimak liste zastareva na svaki upis -- sledece citanje ide u tabele.
    mSnimakOK = False
    ' Racuni firme u configu su se mogli promeniti (Podesavanja idu kroz upis,
    ' a upis kroz RefreshFromData -> ResetCache). Combo se puni ponovo, uz
    ' cuvanje izbora -- v. PuniRacunCombo.
    mRacunPunjen = False
End Sub

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a: modLogError.LogError pocinje sa
    ' "On Error Resume Next", a svaka On Error naredba brise Err.
    errDesc = Err.description
    LogErr "modScrBankaNalozi.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

'=====================================================================
' KLIKOVI
'=====================================================================
Private Function ObradiKlik(ByVal tag As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        ObradiKlik = True
        Exit Function
    End If

    ' Izbor reda ne menja podatke ni prolazno stanje.
    If Left$(tag, 4) = "row:" Then Exit Function

    ' Promena u polju zone stize kao "chg:<tag kontrole>" na SVAKI otkucaj.
    ' Lista NE zavisi od polja (combo bira racun PLATIOCA, ne skup redova),
    ' pa se modOtkupUI.RefreshFromData ovde namerno NE zove (katalog par. 9.2);
    ' osvezava se samo hint, koji pokazuje banku izabranog racuna.
    If Left$(tag, 4) = "chg:" Then
        If Mid$(tag, 5) = "scrBnRacunT" Then OsveziHintSam
        Exit Function
    End If

    ' Dvoklik PREBACUJE red u korpu i iz nje -- povratna radnja nad prolaznim
    ' stanjem, isti obrazac kao Fakturisanje. Knjizenje (avans) i izvoz se NE
    ' pokrecu dvoklikom.
    If Left$(tag, 4) = "dbl:" Then
        ObradiKlik = PrebaciRed(CLng(val(Mid$(tag, 5))))
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(Mid$(tag, 5))
        Exit Function
    End If

    Select Case tag
        Case "scrBnCsv":    ObradiKlik = GenerisiNaloge()
        Case "scrBnSpec":   ObradiKlik = StampajSpecifikaciju()
        Case "scrBnOcisti": ObradiKlik = IsprazniKorpu()
    End Select
End Function

Private Function RadnjaNadRedom(ByVal spec As String) As Boolean
    Dim p() As String, red As Long, kljuc As String
    p = Split(spec, ":")
    If UBound(p) < 1 Then Exit Function
    kljuc = p(0)
    red = CLng(val(p(1)))

    ' Batch radnja ne trazi izabran red (peto polje u BnRadnjeZaListu je 0).
    If kljuc = "bnsve" Then
        RadnjaNadRedom = DodajSveSaRacunom()
        Exit Function
    End If

    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    Select Case kljuc
        Case "bnadd":   RadnjaNadRedom = DodajRedUKorpu(red)
        Case "bniznos": RadnjaNadRedom = ZadajIznos(red)
        Case "bndel":   RadnjaNadRedom = UkloniRedIzKorpe(red)
        Case "bnavans": RadnjaNadRedom = PrimeniAvans(red)
    End Select
End Function

' Identitet iza prikazanog reda. Prazno = red bez identiteta i radnja ODBIJA
' da bira. Do mreze takav red redovno ni ne stigne -- BuildBlokIsplataList
' fail-close-uje na dupli/prazan OtkupID -- pa je ovo poslednja linija, ne
' prva: GridCell van opsega ili pokvaren model ne sme da se pretvori u radnju
' nad pogresnim blokom.
Private Function IdReda(ByVal red As Long, ByVal kol As Long) As String
    Dim iD As String
    If red < 1 Then Exit Function
    iD = Trim$(CStr(modOtkupUI.GridCell(red, kol)))
    If Len(iD) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_DVOSMISLEN"), True
        Exit Function
    End If
    IdReda = iD
End Function

' Ima li blok tekuci racun -- cita se ono sto red NOSI (skrivena kolona), ne
' prazna celija racuna: prazna celija izgleda isto i kad se kolona nije
' nacrtala (v. katalog par. 9.6).
Private Function RedImaTR(ByVal red As Long) As Boolean
    RedImaTR = (Trim$(CStr(modOtkupUI.GridCell(red, BN_KOL_TR))) = "1")
End Function

Private Function RedD(ByVal red As Long, ByVal kol As Long) As Double
    Dim v As Variant
    On Error Resume Next
    v = modOtkupUI.GridCell(red, kol)
    If IsNumeric(v) Then RedD = CDbl(v)
    Err.Clear
End Function

Private Function RedOznaka(ByVal red As Long) As String
    On Error Resume Next
    RedOznaka = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    Err.Clear
End Function

'=====================================================================
' KORPA "U NALOZIMA"
'
' Collection recnika, kao korpa na Fakturisanju. Stavka nosi SAMO identitet
' (OtkupID) plus broj i otvoreno za prikaz u traci -- iznos naloga se NIKAD
' ne cita iz korpe nego svez pri izvozu (OdaberiBlokoveZaNaloge), pa snimak
' ne moze da narucuje uplatu.
'=====================================================================
Private Function Korpa() As Collection
    If mKorpa Is Nothing Then Set mKorpa = New Collection
    Set Korpa = mKorpa
End Function

Private Function UKorpi(ByVal otkupID As String) As Long
    Dim i As Long
    If mKorpa Is Nothing Then Exit Function
    If Len(otkupID) = 0 Then Exit Function
    For i = 1 To mKorpa.count
        If CStr(mKorpa(i)("otkupID")) = otkupID Then
            UKorpi = i
            Exit Function
        End If
    Next i
End Function

' Dodavanje ide preko identiteta; broj i otvoreno stizu iz istog reda mreze i
' sluze SAMO prikazu. Vraca poruku greske ili "" kad je dodato.
Public Function BnDodaj(ByVal otkupID As String, ByVal broj As String, _
                        ByVal otvoreno As Double, ByVal imaTR As Boolean) As String
    Dim red As Object
    If Len(Trim$(otkupID)) = 0 Then
        BnDodaj = Poruka("OTKUI_ERR_BN_DVOSMISLEN")
        Exit Function
    End If
    ' Blok bez tekuceg racuna ne moze u CSV (nema primaoca) -- ne sme ni u
    ' izbor, inace se u potvrdi broji nalog koji nikad ne nastane. Isto
    ' pravilo koje legacy drzi na check-u reda (HandleListSelectionChange).
    If Not imaTR Then
        BnDodaj = Poruka("OTKUI_ERR_BN_BEZ_RACUNA")
        Exit Function
    End If
    If UKorpi(otkupID) > 0 Then
        BnDodaj = Poruka("OTKUI_ERR_BN_VEC_U_NALOZIMA")
        Exit Function
    End If
    Set red = CreateObject("Scripting.Dictionary")
    red("otkupID") = Trim$(otkupID)
    red("broj") = broj
    red("otvoreno") = otvoreno
    Korpa().Add red
End Function

' Uklanjanje po IDENTITETU, ne po prikazu. Vraca True kad je nesto izbaceno.
' Sa clanstvom ide i zadati iznos: iznos bez stavke ne znaci nista.
Public Function BnUkloni(ByVal otkupID As String) As Boolean
    Dim i As Long
    i = UKorpi(otkupID)
    If i = 0 Then Exit Function
    Korpa().Remove i
    If Iznosi().Exists(otkupID) Then Iznosi().Remove otkupID
    BnUkloni = True
End Function

Public Function BnUKorpi(ByVal otkupID As String) As Boolean
    BnUKorpi = (UKorpi(otkupID) > 0)
End Function

Public Function BnKorpaBroj() As Long
    If mKorpa Is Nothing Then Exit Function
    BnKorpaBroj = mKorpa.count
End Function

' Zbir onoga sto bi se STVARNO izvezlo za stavke u korpi: zadati iznos gde
' postoji, inace otvoreno (snimak koji uskladjivanje drzi svezim).
Public Function BnKorpaZbir() As Double
    Dim i As Long, s As Double
    If mKorpa Is Nothing Then Exit Function
    For i = 1 To mKorpa.count
        s = s + BnIznosZa(CStr(mKorpa(i)("otkupID")), CDbl(mKorpa(i)("otvoreno")))
    Next i
    BnKorpaZbir = s
End Function

'--------------------------------------------------- ZADATI IZNOSI PO BLOKU
Private Function Iznosi() As Object
    If mIznosi Is Nothing Then Set mIznosi = CreateObject("Scripting.Dictionary")
    Set Iznosi = mIznosi
End Function

' Iznos koji bi blok poneo u nalog: zadat (delimicna isplata) ili otvoreno.
Public Function BnIznosZa(ByVal otkupID As String, ByVal otvoreno As Double) As Double
    If Iznosi().Exists(otkupID) Then
        BnIznosZa = CDbl(Iznosi()(otkupID))
    Else
        BnIznosZa = otvoreno
    End If
End Function

' Postavljanje zadatog iznosa -- ISTA pravila kao legacy txtIsplatiti_Exit:
' sve u cent-domenu (ZaokruziNovac PRE svake provere), iznos > 0, nikad preko
' otvorenog; jednak otvorenom = brise zadato (puna isplata je podrazumevana).
' Vraca poruku greske ili "" kad je prihvaceno.
Public Function BnPostaviIznos(ByVal otkupID As String, ByVal iznos As Double, _
                               ByVal otvoreno As Double) As String
    Dim iznosC As Double, otvorenoC As Double
    If Len(Trim$(otkupID)) = 0 Then
        BnPostaviIznos = Poruka("OTKUI_ERR_BN_DVOSMISLEN")
        Exit Function
    End If

    iznosC = ZaokruziNovac(iznos)
    otvorenoC = ZaokruziNovac(otvoreno)

    If iznosC <= 0 Then
        BnPostaviIznos = Poruka("OTKUI_ERR_BN_IZNOS_NULA")
        Exit Function
    End If
    If iznosC > otvorenoC Then
        BnPostaviIznos = Poruka("OTKUI_ERR_BN_IZNOS_PREKO") & " " & _
                         Format$(otvorenoC, "#,##0.00")
        Exit Function
    End If

    If iznosC = otvorenoC Then
        If Iznosi().Exists(otkupID) Then Iznosi().Remove otkupID
    Else
        Iznosi()(otkupID) = iznosC
    End If
End Function

' Uskladjivanje zadatih iznosa sa SVEZOM mapom otvorenih -- legacy
' PruneStaleOverrides pravilo, ISTI racun (ClampOverridesToOpenDict): nestao
' ili zatvoren blok gubi zadato, vece od otvorenog se spusta, manje ostaje.
' Vraca broj promena, da se prijave (tiho spustanje bi operater promasio).
Public Function BnUskladiIznose(ByVal ziviOtvoreno As Object) As Long
    BnUskladiIznose = modBankaExportPregled.ClampOverridesToOpenDict(Iznosi(), ziviOtvoreno)
End Function

' Identiteti korpe kao Dictionary -- oblik koji OdaberiBlokoveZaNaloge prima.
' Prazna korpa vraca Nothing = "svi blokovi", isto sto legacy radi kad nema
' selekcije (CollectIsplataBlokovi).
Public Function BnKorpaIDs() As Object
    Dim d As Object, i As Long
    If BnKorpaBroj() = 0 Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To mKorpa.count
        d(CStr(mKorpa(i)("otkupID"))) = True
    Next i
    Set BnKorpaIDs = d
End Function

' USKLADJIVANJE SA SVEZOM LISTOM, pri svakom citanju mreze. Stavka ciji blok
' vise nije otvoren (isplacen, storniran, avans ga zatvorio) IZLAZI iz korpe
' -- uz brojku, nikad tiho (isti razlog kao ClampOverridesToOpen u legacy-ju:
' tiho spustanje bi operater lako promasio). Zivim stavkama se osvezava
' snimak za prikaz, pa traka nikad ne pokazuje bajat iznos.
'
' Mrtva stavka u korpi ne bi mogla nista da narucuje ni bez ovoga (izvoz je
' presek korpe sa SVEZOM listom, a ValidateNalogSaldo je finalna kapija) --
' ovo cuva PRIKAZ: broj i zbir u traci moraju da govore istinu.
Public Function BnUskladiKorpu(ByVal ziviOtvoreno As Object) As Long
    Dim i As Long
    If mKorpa Is Nothing Then Exit Function
    If ziviOtvoreno Is Nothing Then Exit Function
    For i = mKorpa.count To 1 Step -1
        If ziviOtvoreno.Exists(CStr(mKorpa(i)("otkupID"))) Then
            mKorpa(i)("otvoreno") = CDbl(ziviOtvoreno(CStr(mKorpa(i)("otkupID"))))
        Else
            mKorpa.Remove i
            BnUskladiKorpu = BnUskladiKorpu + 1
        End If
    Next i
End Function

Private Function DodajRedUKorpu(ByVal red As Long) As Boolean
    Dim iD As String, greska As String
    iD = IdReda(red, BN_KOL_ID)
    If Len(iD) = 0 Then Exit Function

    greska = BnDodaj(iD, RedOznaka(red), RedD(red, BN_KOL_OTVORENO), RedImaTR(red))
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_DODATO"), False
    DodajRedUKorpu = True
End Function

Private Function UkloniRedIzKorpe(ByVal red As Long) As Boolean
    Dim iD As String
    iD = IdReda(red, BN_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    If Not BnUkloni(iD) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NIJE_U_NALOZIMA"), True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_UKLONJENO"), False
    UkloniRedIzKorpe = True
End Function

Private Function PrebaciRed(ByVal red As Long) As Boolean
    Dim iD As String
    iD = IdReda(red, BN_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    If BnUKorpi(iD) Then
        PrebaciRed = UkloniRedIzKorpe(red)
    Else
        PrebaciRed = DodajRedUKorpu(red)
    End If
End Function

Private Function IsprazniKorpu() As Boolean
    If BnKorpaBroj() = 0 Then Exit Function
    Set mKorpa = New Collection
    Set mIznosi = Nothing
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_KORPA_PRAZNA"), False
    IsprazniKorpu = True
End Function

' Radnja "Iznos...": delimicna isplata izabranog bloka. InputBox (presedan:
' SEF komentar), predlog = tekuci iznos za taj blok; prazan unos / otkaz ne
' menja nista. Blok koji jos nije u nalozima se dodaje -- operater koji mu
' zadaje iznos ocigledno bira bas njega.
Private Function ZadajIznos(ByVal red As Long) As Boolean
    Dim iD As String, unos As String, greska As String
    Dim iznos As Double, otvoreno As Double

    iD = IdReda(red, BN_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    If Not RedImaTR(red) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_BEZ_RACUNA"), True
        Exit Function
    End If
    otvoreno = RedD(red, BN_KOL_OTVORENO)

    unos = Trim$(InputBox(Poruka("OTKUI_ASK_BN_IZNOS") & " " & RedOznaka(red) & vbCrLf & vbCrLf & _
                          Poruka("OTKUI_LBL_BN_AVANS_OTVORENO") & " " & _
                          Format$(otvoreno, "#,##0.00") & " RSD", APP_NAME, _
                          Format$(BnIznosZa(iD, otvoreno), "0.00")))
    If Len(unos) = 0 Then Exit Function

    If Not TryParseDouble(unos, iznos) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_IZNOS_NEISPRAVAN"), True
        Exit Function
    End If

    greska = BnPostaviIznos(iD, iznos, otvoreno)
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If

    ' U naloge, ako vec nije -- bez ovoga bi zadat iznos stajao a blok ne bi
    ' isao u fajl, sto izgleda kao da unos ne radi.
    If Not BnUKorpi(iD) Then
        greska = BnDodaj(iD, RedOznaka(red), otvoreno, True)
        If Len(greska) > 0 Then
            modOtkupUI.ShowToast greska, True
            Exit Function
        End If
    End If

    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_IZNOS") & " " & _
                         Format$(BnIznosZa(iD, otvoreno), "#,##0.00") & " RSD", False
    ZadajIznos = True
End Function

'=====================================================================
' PRIMENI AVANS -- jedino knjizenje sa ovog ekrana (tblNovac).
'
' Motor je postojeci modNovac.ApplyAvansToOtkup_TX: transakcija, guard na
' dupli OtkupID u core-u. ApplyAvansToOtkup_TX vraca True i kad NISTA nije
' vezano (avans u medjuvremenu potrosen, blok zatvoren) -- stvarno proknjizen
' iznos se cita iz ByRef parametra i tek on je dokaz da se nesto desilo
' (AUD-026 c / RF-02); uspeh se ne prijavljuje na no-op.
'=====================================================================
Private Function PrimeniAvans(ByVal red As Long) As Boolean
    Dim iD As String, koopID As String
    Dim avans As Double, otvoreno As Double, vezuje As Double
    Dim primenjeno As Double

    iD = IdReda(red, BN_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    koopID = Trim$(CStr(modOtkupUI.GridCell(red, BN_KOL_KOOP)))
    If Len(koopID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_DVOSMISLEN"), True
        Exit Function
    End If

    avans = RedD(red, BN_KOL_AVANS)
    otvoreno = RedD(red, BN_KOL_OTVORENO)
    If avans <= 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NEMA_AVANSA"), True
        Exit Function
    End If
    If otvoreno <= 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NEMA_OTVORENOG"), True
        Exit Function
    End If

    vezuje = avans
    If vezuje > otvoreno Then vezuje = otvoreno

    If MsgBox(Poruka("OTKUI_ASK_BN_AVANS") & " " & RedOznaka(red) & "?" & vbCrLf & vbCrLf & _
              Poruka("OTKUI_LBL_BN_AVANS_OTVORENO") & " " & Format$(otvoreno, "#,##0.00") & vbCrLf & _
              Poruka("OTKUI_LBL_BN_AVANS_DOSTUPNO") & " " & Format$(avans, "#,##0.00") & vbCrLf & _
              Poruka("OTKUI_LBL_BN_AVANS_VEZUJE") & " " & Format$(vezuje, "#,##0.00"), _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    If Not ApplyAvansToOtkup_TX(koopID, iD, primenjeno) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_AVANS"), True
        Exit Function
    End If

    Scr_ResetCache
    If primenjeno > 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_AVANS_OK") & " " & _
                             Format$(primenjeno, "#,##0.00") & " RSD", False
    Else
        ' TX je prosla, ali nista nije proknjizeno -- ne prijavljuj uspeh.
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_AVANS_NOOP"), True
    End If
    PrimeniAvans = True
End Function

'=====================================================================
' IZVOZ: CSV NALOGA + PDF SPECIFIKACIJA (dugmad zone)
'
' Izbor blokova je jedno pravilo za oba izlaza, i zivi u domenu
' (OdaberiBlokoveZaNaloge): korpa ako u njoj ima stavki, inace SVI otvoreni;
' bez tekuceg racuna se preskace i broji; iznos = SVEZ otvoren, normalizovan
' u cent-domen PRE praga "> 0" (AUD-026). Ekran ne ponavlja nijedan deo toga.
'=====================================================================
Private Function BlokoviZaIzvoz(ByRef outBezTR As Long, ByRef outIzbaceno As Long) As Collection
    ' PRAZAN IZBOR NE IZVOZI NISTA. Prvi ugovor ("prazno = svi otvoreni",
    ' preuzet od legacy 'nema selekcije = svi') je oborila recenzija PR-a:
    ' CSV ne knjizi isplatu, pa su blokovi otvoreni i POSLE fajla -- a izbor
    ' se posle uspesnog izvoza prazni. Drugi klik bi tako tiho napravio
    ' naloge za SVE otvorene, ukljucujuci pun iznos bloka ciji je ZADATI deo
    ' upravo izvezen. "Svi" zato postoji samo kao izricita radnja
    ' (bnsve, "Svi sa racunom"), nikad kao znacenje praznog.
    outBezTR = 0
    outIzbaceno = 0
    If BnKorpaBroj() = 0 Then
        Set BlokoviZaIzvoz = New Collection
        Exit Function
    End If

    Dim sveze As Collection
    Set sveze = modBankaExportPregled.BuildBlokIsplataList()
    Set BlokoviZaIzvoz = modBankaExportPregled.OdaberiBlokoveZaNaloge( _
                             sveze, BnKorpaIDs(), outBezTR, outIzbaceno, Iznosi())
End Function

' Izricito "svi": puni izbor SVIM otvorenim blokovima sa racunom (svez
' snimak, ista selekcija koju bi izvoz uzeo). Radnja operatera -- prazan
' izbor to NE znaci sam od sebe.
Private Function DodajSveSaRacunom() As Boolean
    Dim sveze As Collection, blokovi As Collection
    Dim bezTR As Long, nepoznato As Long, dodato As Long
    Dim blk As clsBlokIsplata, v As Variant

    Set sveze = modBankaExportPregled.BuildBlokIsplataList()
    Set blokovi = modBankaExportPregled.OdaberiBlokoveZaNaloge( _
                      sveze, Nothing, bezTR, nepoznato)
    For Each v In blokovi
        Set blk = v
        If UKorpi(blk.otkupID) = 0 Then
            If Len(BnDodaj(blk.otkupID, blk.brojDokumenta, blk.OtvorenIznos, True)) = 0 Then
                dodato = dodato + 1
            End If
        End If
    Next v

    If dodato = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_SVE_NISTA"), True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_SVE_DODATO") & " " & CStr(dodato), False
    DodajSveSaRacunom = True
End Function

Private Function GenerisiNaloge() As Boolean
    Dim racun As String, racunInfo As String
    Dim blokovi As Collection
    Dim bezTR As Long, izbaceno As Long
    Dim ukupno As Double, csvPath As String, odbijeno As String
    Dim pitanje As String

    racun = IzabraniRacun()
    If Len(racun) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NEMA_RACUNA"), True
        Exit Function
    End If

    ' Prazan izbor i "izabrani vise ne mogu u nalog" su dve razlicite poruke
    ' -- gate je u BlokoviZaIzvoz, ovde se samo bira tekst.
    If BnKorpaBroj() = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_IZBOR_PRAZAN"), True
        Exit Function
    End If

    Set blokovi = BlokoviZaIzvoz(bezTR, izbaceno)
    If blokovi.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NEMA_BLOKOVA"), True
        Exit Function
    End If

    ukupno = ZbirIsplatiti(blokovi)
    racunInfo = racun
    If Len(BankaNazivZaRacun(racun)) > 0 Then
        racunInfo = racunInfo & " (" & BankaNazivZaRacun(racun) & ")"
    End If

    ' JEDNO dugme sa jasnim ishodom: potvrda kaze N, ukupan iznos, racun i
    ' datum valute PRE upisa fajla; nista se ne desava tiho.
    pitanje = Poruka("OTKUI_ASK_BN_CSV") & " " & CStr(blokovi.count) & vbCrLf & vbCrLf & _
              Poruka("OTKUI_LBL_BN_UKUPNO") & " " & Format$(ukupno, "#,##0.00") & " RSD" & vbCrLf & _
              Poruka("OTKUI_LBL_BN_SA_RACUNA") & " " & racunInfo & vbCrLf & _
              Poruka("OTKUI_LBL_BN_VALUTA") & " " & Format$(Date, "d.m.yyyy")
    If bezTR > 0 Then
        pitanje = pitanje & vbCrLf & Poruka("OTKUI_LBL_BN_PRESKOCENO_TR") & " " & CStr(bezTR)
    End If
    If izbaceno > 0 Then
        pitanje = pitanje & vbCrLf & Poruka("OTKUI_LBL_BN_PRESKOCENO_MRTVO") & " " & CStr(izbaceno)
    End If

    If MsgBox(pitanje, vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    csvPath = modBankaExportPregled.GenerisiNalogeCSV(blokovi, racun, odbijeno)

    ' Finalna kapija (ValidateNalogSaldo nad SVEZIM saldom) je odbila naloge:
    ' fajl NIJE napisan, razlog ide operateru, prikaz se osvezava da vidi
    ' trenutno stanje.
    If LenB(odbijeno) > 0 Then
        MsgBox odbijeno, vbExclamation, APP_NAME
        Scr_ResetCache
        GenerisiNaloge = True
        Exit Function
    End If

    If LenB(csvPath) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_CSV"), True
        Exit Function
    End If

    ' Fajl je napisan -- izbor je POTROSEN i prazni se (sa zadatim iznosima).
    ' CSV ne knjizi isplatu, pa blokovi ostaju otvoreni; upravo zato prazan
    ' izbor NE izvozi nista (v. BlokoviZaIzvoz) -- drugi klik dobija poruku,
    ' ne tihu reprizu svih otvorenih.
    Set mKorpa = New Collection
    Set mIznosi = Nothing
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_CSV") & " " & CStr(blokovi.count) & _
                         "  " & ChrW(183) & "  " & Format$(ukupno, "#,##0.00") & " RSD", False
    MsgBox Poruka("OTKUI_MSG_BN_CSV") & " " & CStr(blokovi.count) & vbCrLf & vbCrLf & csvPath, _
           vbInformation, APP_NAME

    ' Otvori folder sa oznacenim fajlom -- operater ga odatle uvozi u e-banking.
    On Error Resume Next
    Shell "explorer.exe /select,""" & csvPath & """", vbNormalFocus
    On Error GoTo 0

    GenerisiNaloge = True
End Function

' PDF specifikacija istih blokova (i istih iznosa) koje bi uzeo i CSV.
' Render i rezim (PDF/PRINT/PREVIEW/OFF) su u modPrint / config -- ekran samo
' zove postojecu rutinu. Ne moze se verifikovati automatski: ide na smoke
' checklistu operatera.
Private Function StampajSpecifikaciju() As Boolean
    Dim blokovi As Collection
    Dim bezTR As Long, izbaceno As Long
    Dim mode As String

    If BnKorpaBroj() = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_IZBOR_PRAZAN"), True
        Exit Function
    End If

    Set blokovi = BlokoviZaIzvoz(bezTR, izbaceno)
    If blokovi.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_NEMA_BLOKOVA"), True
        Exit Function
    End If

    ' OFF nije "nista se nije desilo" nego iskljucen izlaz -- kaze se, jer bi
    ' klik bez ijedne poruke izgledao kao dugme koje ne radi.
    mode = DocResolveMode(GetConfigValue(CFG_ISPLATA_SPEC_PRINT_MODE), "PDF")
    If mode = "OFF" Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BN_SPEC_OFF"), True
        Exit Function
    End If

    modBankaExportPregled.PrintIsplataSpecifikacija blokovi, IzabraniRacun()
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_SPEC") & " " & CStr(blokovi.count) & _
                         "  " & ChrW(183) & "  " & Format$(ZbirIsplatiti(blokovi), "#,##0.00") & _
                         " RSD", False
End Function

Private Function ZbirIsplatiti(ByVal blokovi As Collection) As Double
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        ZbirIsplatiti = ZbirIsplatiti + blk.IsplatitiIznos
    Next v
End Function

'=====================================================================
' REDOVI MREZE
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Dim rez As Variant
    ' Zona se puni odavde, kao na ostalim ugovornim ekranima: gradi se jednom,
    ' a podaci za nju postoje tek kad se lista cita.
    OsveziZonu
    rez = RedoviNalozi(filter, q)
    Scr_Rows = rez

    ' Trag za Diag_BnRedovi -- ne menja nista.
    On Error Resume Next
    mDiagFilter = filter
    mDiagQ = q
    mDiagN = CLng(rez(2))
    Err.Clear
End Function

' Opis kolona PO KLJUCU LISTE -- da se pravilo "identitet je u redu i ne crta
' se" moze tvrditi bez stanja ekrana.
Public Function BnKoloneZaListu(ByVal kljuc As String) As Variant
    BnKoloneZaListu = NaloziKolone()
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

' Prva kolona se uvek crta kao BROJ dokumenta (StyleGridCell, isBroj) -- tu
' stoji broj otkupnog bloka, koji je i poziv na broj u nalogu.
'
' Poslednje cetiri nose ono sto radnje moraju da znaju a iz prikaza se ne
' vidi jednoznacno; prioritet 4, pa se ne crtaju (v. BN_KOL_*).
Private Function NaloziKolone() As Variant
    ' ISPLATITI stoji uz OTVORENO: podrazumevano su isti broj, a razlikuju se
    ' tacno tamo gde je operater zadao delimicnu isplatu.
    NaloziKolone = Array( _
        "OTKUI_HD_BROJ||txt|96|1", _
        "OTKUI_HD_OZN||txt|32|1", _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_OM||txt|86|3", _
        "OTKUI_HDN_UKUPNO||rsd|96|3", _
        "OTKUI_HDN_ISPLACENO||rsd|96|3", _
        "OTKUI_HDN_OTVORENO||rsd|104|1", _
        "OTKUI_HDN_ISPLATITI||rsd|104|1", _
        "OTKUI_HDB_RACUN||txt|132|2", _
        "OTKUI_HDN_OTKID||txt|1|4", _
        "OTKUI_HDN_KOOPID||txt|1|4", _
        "OTKUI_HDN_IMATR||txt|1|4", _
        "OTKUI_HDN_AVANS||txt|1|4")
End Function

' Citac (GetBlokIsplataForGrid) vraca 1-bazirano:
'   1 OtkupID | 2 BrojDokumenta | 3 Datum | 4 KooperantNaziv | 5 KooperantID
'   6 StanicaID | 7 Ukupan | 8 Isplaceno | 9 Otvoren | 10 TekuciRacun
'   11 ImaTR | 12 AvansSaldo
Private Function RedoviNalozi(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String, iD As String, imaTR As Boolean, avans As Double
    Dim zbirOtv As Double, uskladjeno As Long, usklIznosa As Long
    Dim zivi As Object
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    src = Snimak()
    If Not IsArray(src) Then
        ' Nema otvorenih blokova -- ni korpa ni zadati iznosi nemaju nad cim
        ' da stoje.
        Set zivi = CreateObject("Scripting.Dictionary")
        uskladjeno = BnUskladiKorpu(zivi)
        If uskladjeno > 0 Then PrijaviUskladjivanje uskladjeno
        usklIznosa = BnUskladiIznose(zivi)
        If usklIznosa > 0 Then PrijaviUskladjivanjeIznosa usklIznosa
        RedoviNalozi = PrazanRezultat(NaloziKolone())
        Exit Function
    End If

    ' Korpa i zadati iznosi se uskladjuju sa SVEZIM skupom otvorenih, PRE cipa
    ' i pretrage: to je izbor za izvoz, a cip je pregled -- filter ne sme da
    ' izbacuje stavke iz izbora. Iznose drzi ISTI klamp kao legacy override
    ' (ClampOverridesToOpenDict): nestao blok gubi zadato, vece se spusta.
    ' U izboru ostaju samo blokovi koji JOS mogu u nalog: otvoreni SA
    ' racunom. Blok kome je racun u medjuvremenu obrisan izlazi (izvoz bi ga
    ' ionako preskocio, ali traka, zbir i potvrda ne smeju da ga pokazuju
    ' kao spreman) -- recenzija PR-a, tacka 4.
    Set zivi = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(src, 1)
        If CBool(src(i, 11)) Then zivi(Trim$(CStr(src(i, 1)))) = CDbl(src(i, 9))
    Next i
    uskladjeno = BnUskladiKorpu(zivi)
    If uskladjeno > 0 Then PrijaviUskladjivanje uskladjeno
    usklIznosa = BnUskladiIznose(zivi)
    If usklIznosa > 0 Then PrijaviUskladjivanjeIznosa usklIznosa

    ' Upit se normalizuje JEDNOM, haystack po redu -- v. TekstZaPretragu:
    ' imena nose kvake, operater ih (DE/EN tastatura) ne kuca.
    Dim qN As String
    qN = modUiData.TekstZaPretragu(q)

    ReDim outA(1 To UBound(src, 1), 1 To 14)
    For i = 1 To UBound(src, 1)
        iD = Trim$(CStr(src(i, 1)))
        imaTR = CBool(src(i, 11))
        avans = CDbl(src(i, 12))
        If Not BnCipNalog(filter, imaTR, avans) Then GoTo Sledeci
        hay = modUiData.TekstZaPretragu(CStr(src(i, 2)) & "|" & CStr(src(i, 4)) & "|" & _
              CStr(src(i, 6)) & "|" & CStr(src(i, 10)) & "|" & iD)
        If Len(qN) > 0 Then
            If InStr(1, hay, qN, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i, 2))
        ' Kvacica se racuna iz KORPE, ne iz tabele -- korpa je prolazno stanje.
        outA(n, 2) = IIf(UKorpi(iD) > 0, ChrW(10003), "")
        ' Datum ide kao serijski broj -- pravilo je ljuskino (modUiData.CellDate),
        ' ekran ga ne ponavlja.
        outA(n, 3) = modUiData.CellDate(src, i, 3)
        outA(n, 4) = CStr(src(i, 4))
        outA(n, 5) = CStr(src(i, 6))
        outA(n, 6) = CDbl(src(i, 7))
        outA(n, 7) = CDbl(src(i, 8))
        outA(n, 8) = CDbl(src(i, 9))
        ' Sta bi blok poneo u nalog: zadati iznos ili otvoreno. Klamp iznad
        ' garantuje da zadato nikad nije preko svezeg otvorenog.
        outA(n, 9) = BnIznosZa(iD, CDbl(src(i, 9)))
        outA(n, 10) = CStr(src(i, 10))
        outA(n, 11) = iD
        outA(n, 12) = CStr(src(i, 5))
        outA(n, 13) = IIf(imaTR, "1", "")
        outA(n, 14) = avans
        zbirOtv = zbirOtv + CDbl(src(i, 9))
Sledeci:
    Next i

    ' Zbir kolicine nema smisla nad blokovima isplate; vrednost je zbir
    ' OTVORENOG prikazanih redova -- izbrojan pod istim filterima kao redovi.
    RedoviNalozi = Array(NaloziKolone(), outA, n, 0#, zbirOtv, Array(0, 0, 0))
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.description
    Err.Raise errNum, "modScrBankaNalozi.RedoviNalozi", errDesc
End Function

' Snimak liste: iz tabela SAMO kad je zastareo (posle upisa), inace iz kesa.
' Pretraga i cipovi time postaju re-filter nad snimkom -- trenutni, kao u
' legacy formi. Greska citanja se NE kesira (isti obrazac kao Kpi).
Private Function Snimak() As Variant
    If Not mSnimakOK Then
        mSnimakPunjenja = mSnimakPunjenja + 1
        mSnimak = modBankaExportPregled.GetBlokIsplataForGrid()
        mSnimakOK = True
    End If
    Snimak = mSnimak
End Function

' Izbacivanje iz korpe se PRIJAVLJUJE -- tiho smanjenje izbora bi operater
' lako promasio (isti razlog kao poruka o usklajdenim override-ima u legacy
' LoadBlokovi). Bez forme je ShowToast no-op, pa test ovo ne vidi kao pad.
Private Sub PrijaviUskladjivanje(ByVal n As Long)
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_KORPA_USKLADJENA") & " " & CStr(n), True
End Sub

Private Sub PrijaviUskladjivanjeIznosa(ByVal n As Long)
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BN_IZNOSI_USKLADJENI") & " " & CStr(n), True
End Sub

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod reda polja. MORA da bude LABELA, ne Frame: Frame je
    ' prozorska kontrola i crta se IZNAD bezprozorskih bez obzira na z-order.
    ' Napravljena PRVA, labela ostaje ispod svega.
    modUiKit.NewLbl z, "bnBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "bnCap", UCase$(Poruka("OTKUI_SCRBN_CAP")), PAD, BN_Y_CAP, _
                    260, 11, TS_MICRO, True, C_MUTED, -1

    ' Cetiri brojke desno -- iste koje legacy forma drzi u KPI traci
    ' (Otvoreno / Selektovano -> U nalozima / Bez TR / Avans pool).
    For i = 0 To 3
        modUiKit.NewLbl z, "bnKL" & i, "", 0, BN_Y_CAP, BN_KPI_W, 11, _
                        TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "bnKV" & i, ChrW(8212), 0, BN_Y_KPI_V, BN_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' TRAKA KORPE "U NALOZIMA": naslov, poslednje stavke i zbir. Sadrzaj puni
    ' OsveziKorpuPanel, mesto daje RasporediPolja.
    modUiKit.NewLbl z, "bnKorpaCap", "", 0, BN_Y_LBL, BN_KORPA_W, 11, _
                    TS_MICRO, True, C_MUTED, -1
    For i = 0 To BN_KORPA_N - 1
        modUiKit.NewLbl z, "bnKorpaR" & i, "", 0, BN_Y_LBL + 16 + i * 13, _
                        BN_KORPA_W, 12, TS_META, False, C_FOREST, -1
    Next i
    modUiKit.NewLbl z, "bnKorpaZ", "", 0, BN_Y_LBL + 18 + BN_KORPA_N * 13, _
                    BN_KORPA_W, 13, TS_META, True, C_GREEN, -1

    ' POLJE "SA RACUNA". Pravi ga ljuska (NewFieldG); prefiks "scr" je
    ' OBAVEZAN, a kombo MORA biti polje (okvir nm + kontrola nmT) -- panel za
    ' izbor (modOtkupUI.FindCombo) trazi bas taj oblik.
    modOtkupUI.NewFieldG z, "scrBnRacun", Poruka("OTKUI_FLD_BN_RACUN"), "cmb", "", _
                         1, False, False, "BN"

    modUiKit.NewLbl z, "bnHint", "", PAD, BN_Y_HINT, 400, 12, TS_META, False, C_MUTED, -1

    ' Dva izlaza + ciscenje izbora. CSV je primaran -- to je posao ekrana.
    modUiKit.BtnV z, "scrBnCsv", Poruka("OTKUI_BTN_BN_CSV"), PAD, BN_Y_BTN, _
                  164, BN_BTN_H, "primary"
    modUiKit.BtnV z, "scrBnSpec", Poruka("OTKUI_BTN_BN_SPEC"), PAD + 172, BN_Y_BTN, _
                  150, BN_BTN_H, "soft"
    modUiKit.BtnV z, "scrBnOcisti", Poruka("OTKUI_BTN_BN_OCISTI"), PAD + 330, BN_Y_BTN, _
                  110, BN_BTN_H, "ghost"

    modUiKit.NewLbl z, "bnLnB", "", 0, BN_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = BN_ZONA_H
End Function

Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, kx As Single, kxK As Single
    Dim wPolja As Single, korpaVidi As Boolean, capDesno As Single
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    z.Controls("bnBg").Left = PAD - 10
    z.Controls("bnBg").top = BN_Y_LBL - 8
    z.Controls("bnBg").width = w - 2 * (PAD - 10)
    z.Controls("bnBg").Height = BN_Y_BTN - BN_Y_LBL + 2

    ' Desna traka (korpa) uzima svoje, polje i dugmad dele OSTATAK. Na uskom
    ' ekranu traka nestaje -- bolje bez trake nego sa dugmadima koja se ne vide.
    wPolja = w - BN_KORPA_W - PAD
    korpaVidi = (wPolja >= BN_POLJA_MIN)
    If Not korpaVidi Then wPolja = w
    kxK = w - BN_KORPA_W

    z.Controls("bnKorpaCap").Left = kxK
    z.Controls("bnKorpaCap").Visible = korpaVidi
    z.Controls("bnKorpaZ").Left = kxK
    z.Controls("bnKorpaZ").Visible = korpaVidi
    For i = 0 To BN_KORPA_N - 1
        z.Controls("bnKorpaR" & i).Left = kxK
        z.Controls("bnKorpaR" & i).Visible = korpaVidi
    Next i

    ' Brojke idu uz desnu ivicu; sakriva se ona koja bi nalegla na naslov zone.
    capDesno = PAD + 200
    For i = 0 To 3
        kx = w - PAD - (4 - i) * BN_KPI_W
        z.Controls("bnKL" & i).Left = kx
        z.Controls("bnKV" & i).Left = kx
        z.Controls("bnKL" & i).Visible = (kx > capDesno)
        z.Controls("bnKV" & i).Visible = (kx > capDesno)
    Next i

    PoljeX z, "scrBnRacun", PAD, BN_FLD_W, BN_Y_LBL

    ' Objasnjenje se zaustavlja pred trakom -- Label ne prelama, samo istece.
    z.Controls("bnHint").width = wPolja - PAD * 2

    modUiKit.MoveBtn z, "scrBnCsv", PAD, BN_Y_BTN
    modUiKit.MoveBtn z, "scrBnSpec", PAD + 172, BN_Y_BTN
    modUiKit.MoveBtn z, "scrBnOcisti", PAD + 330, BN_Y_BTN

    z.Controls("bnLnB").width = w
End Sub

Private Sub PoljeX(ByVal z As Object, ByVal nm As String, ByVal X As Single, _
                   ByVal w As Single, ByVal yLbl As Single)
    On Error Resume Next
    z.Controls(nm).Left = X
    z.Controls(nm).top = yLbl
    z.Controls(nm).width = w
    modOtkupUI.LayoutFieldInner z.Controls(nm)
End Sub

Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone("BANKA_NALOZI")
End Function

Private Function Kontrola(ByVal nm As String) As Object
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    Set Kontrola = z.Controls(nm).Controls(nm & "T")
End Function

Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PuniRacunCombo
    RasporediPolja z, z.width
    OsveziKorpuPanel z
    OsveziHint z
    OsveziBrojke z
End Sub

Private Sub OsveziHintSam()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    OsveziHint z
End Sub

' Racuni firme sa kojih idu nalozi: BankaNalogRacuniCSV (RACUN_1..4, fallback
' legacy ";"-spisak pa SELLER_ACCOUNT), prikaz "racun (Banka)". Cist racun ide
' u DRUGU kolonu -- prikaz je za coveka i sme da se menja, podatak ne (isto
' pravilo kao identitet u redu). Default izbor = SELLER_ACCOUNT ako je medju
' racunima, kao u legacy PopulateRacunCombo.
'
' Puni se ponovo posle svakog ResetCache (racuni u Podesavanjima su se mogli
' promeniti), ali CUVA izbor operatera.
Private Sub PuniRacunCombo()
    Dim c As Object, raw As String, parts() As String
    Dim i As Long, r As String, disp As String, bn As String
    Dim cur As String, def As String, idx As Long
    On Error GoTo EH

    If mRacunPunjen Then Exit Sub
    Set c = Kontrola("scrBnRacun")
    If c Is Nothing Then Exit Sub

    mFill = True
    cur = GetComboID(c)

    c.Clear
    c.ColumnCount = 2
    c.ColumnWidths = "180 pt;0 pt"
    c.BoundColumn = 1
    c.TextColumn = 1

    raw = modBankaExportPregled.BankaNalogRacuniCSV()
    parts = Split(raw, ";")
    For i = LBound(parts) To UBound(parts)
        r = Replace(Trim$(parts(i)), " ", "")
        If LenB(r) > 0 Then
            disp = r
            bn = modBankaExportPregled.BankaNazivZaRacun(r)
            If LenB(bn) > 0 Then disp = disp & "  (" & bn & ")"
            c.AddItem disp
            c.List(c.ListCount - 1, 1) = r
        End If
    Next i

    If c.ListCount > 0 Then
        idx = 0
        def = Replace(Trim$(DocConfigOr("SELLER_ACCOUNT", "")), " ", "")
        If LenB(cur) > 0 Then def = cur   ' izbor operatera preziviva refill
        If LenB(def) > 0 Then
            For i = 0 To c.ListCount - 1
                If CStr(c.List(i, 1)) = def Then idx = i: Exit For
            Next i
        End If
        c.ListIndex = idx
    End If

    mRacunPunjen = True
    mFill = False
    Exit Sub
EH:
    mFill = False
    ' Prazan combo bez traga je bio glavni razlog zasto je izgledalo da "nista
    ' nije povezano" -- isto kao u modOtkupUI.FillCombos.
    Debug.Print "modScrBankaNalozi.PuniRacunCombo PAO: " & Err.Number & " " & Err.description
End Sub

'------------------------------------------------------- PANEL KORPE
Private Sub OsveziKorpuPanel(ByVal z As Object)
    Dim i As Long, n As Long
    On Error Resume Next
    n = BnKorpaBroj()
    z.Controls("bnKorpaCap").caption = UCase$(Poruka("OTKUI_LBL_BN_KORPA_CAP"))
    For i = 0 To BN_KORPA_N - 1
        z.Controls("bnKorpaR" & i).caption = TrakaRed(i)
    Next i
    If n = 0 Then
        z.Controls("bnKorpaZ").caption = Poruka("OTKUI_LBL_BN_KORPA_SVI")
    Else
        z.Controls("bnKorpaZ").caption = n & " " & _
            Poruka("OTKUI_LBL_AG_KORPA_STAVKI") & "   " & ChrW(183) & "   " & _
            Format$(BnKorpaZbir(), "#,##0") & " RSD"
    End If
End Sub

' Tekst reda trake. NAJNOVIJE PRVO: operater upravo nesto doda, pa mu je
' potvrda ono sto trazi. PRELIV SE PRIJAVLJUJE: lista koja se tiho odseca
' izgleda kao cela. Racun je odvojen od crtanja, pa se meri bez forme.
Public Function TrakaRed(ByVal i As Long) As String
    Dim n As Long, sakriveno As Long
    If mKorpa Is Nothing Then Exit Function
    n = mKorpa.count
    If n = 0 Then Exit Function
    If i < 0 Or i > BN_KORPA_N - 1 Then Exit Function

    ' Sve staje: samo obrni redosled.
    If n <= BN_KORPA_N Then
        If i > n - 1 Then Exit Function
        TrakaRed = KorpaRedPrikaz(n - i)
        Exit Function
    End If

    ' Ne staje: poslednji red je prelivni.
    If i < BN_KORPA_N - 1 Then
        TrakaRed = KorpaRedPrikaz(n - i)
        Exit Function
    End If
    sakriveno = n - (BN_KORPA_N - 1)
    TrakaRed = ChrW(8230) & " " & Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & sakriveno
End Function

' Red trake nosi iznos koji bi se STVARNO izvezao (zadati ili otvoreno) --
' isti racun kao zbir ispod njega. Smoke 28.08: red je pokazivao otvoreno
' (21.798) dok je zbir pokazivao zadatih 10.000 -- dva broja jedan ispod
' drugog koja se ne slazu.
Private Function KorpaRedPrikaz(ByVal i As Long) As String
    Dim red As Object
    On Error Resume Next
    If mKorpa Is Nothing Then Exit Function
    If i < 1 Or i > mKorpa.count Then Exit Function
    Set red = mKorpa(i)
    KorpaRedPrikaz = CStr(red("broj")) & "   " & ChrW(183) & "   " & _
                     Format$(BnIznosZa(CStr(red("otkupID")), CDbl(red("otvoreno"))), "#,##0")
End Function

'---------------------------------------------------------- BROJKE I HINT
Private Sub OsveziHint(ByVal z As Object)
    Dim bn As String
    On Error Resume Next
    bn = modBankaExportPregled.BankaNazivZaRacun(IzabraniRacun())
    If Len(IzabraniRacun()) = 0 Then
        z.Controls("bnHint").caption = Poruka("OTKUI_LBL_BN_HINT_RACUN")
    ElseIf Len(bn) > 0 Then
        z.Controls("bnHint").caption = Poruka("OTKUI_LBL_BN_HINT") & "  " & _
                                       ChrW(183) & "  " & bn
    Else
        z.Controls("bnHint").caption = Poruka("OTKUI_LBL_BN_HINT")
    End If
End Sub

Private Sub OsveziBrojke(ByVal z As Object)
    Dim k As Variant, nepoznato As Boolean, crta As String
    On Error Resume Next
    k = Kpi()
    nepoznato = BnKpiNepoznat(k)
    crta = ChrW(8212)

    z.Controls("bnKL0").caption = UCase$(Poruka("OTKUI_KPI_BN_OTVORENO"))
    z.Controls("bnKL1").caption = UCase$(Poruka("OTKUI_KPI_BN_UNALOZIMA"))
    z.Controls("bnKL2").caption = UCase$(Poruka("OTKUI_KPI_BN_BEZTR"))
    z.Controls("bnKL3").caption = UCase$(Poruka("OTKUI_KPI_BN_AVANS"))

    ' Nula i "ne znam" nisu ista brojka -- v. BnKpiPosleGreske.
    If nepoznato Then
        z.Controls("bnKV0").caption = crta
        z.Controls("bnKV2").caption = crta
        z.Controls("bnKV3").caption = crta
    Else
        z.Controls("bnKV0").caption = Format$(CDbl(k(2)), "#,##0")
        z.Controls("bnKV2").caption = CStr(CLng(k(1)))
        z.Controls("bnKV3").caption = Format$(CDbl(k(3)), "#,##0")
    End If

    ' "U nalozima" je korpa -- prolazno stanje ekrana, ne KPI iz tabela.
    If BnKorpaBroj() = 0 Then
        z.Controls("bnKV1").caption = crta
    Else
        z.Controls("bnKV1").caption = CStr(BnKorpaBroj()) & " / " & _
                                      Format$(BnKorpaZbir(), "#,##0")
    End If
End Sub

' Cetiri brojke iz JEDNOG prolaza (modBankaExportPregled.NalogeKpi), kesirane
' do sledeceg upisa.
'
' NEUSPEH CITANJA NIJE NULA: znacka odgovara na "ima li blokova koji cekaju
' isplatu", pa bi nula posle greske znacila "nema posla" umesto "ne znam" --
' isti fail-open je vec placen u Stornu i na Uvozu izvoda. Greska se loguje,
' kes se ne proglasava vazecim, vraca se poslednja poznata vrednost; kad nje
' jos nema, brojka je NEPOZNATA (-1, ljuska je crta kao "!").
Private Function Kpi() As Variant
    Dim errDesc As String
    On Error GoTo EH
    If mKpiOK Then
        Kpi = mKpi
        Exit Function
    End If
    mKpi = modBankaExportPregled.NalogeKpi()
    mKpiOK = True
    Kpi = mKpi
    Exit Function
EH:
    ' Brojac ne sme da obori ljusku: OsveziNavBrojace pita SVAKI ekran.
    errDesc = Err.description
    LogErr "modScrBankaNalozi.Kpi"
    Kpi = BnKpiPosleGreske(mKpi)
    Err.Clear
End Function

' Sta brojac vraca kad citanje pukne: POSLEDNJU POZNATU vrednost, ne nule.
' Odvojeno od Kpi da bi se pravilo moglo izmeriti bez lomljenja seme.
Public Function BnKpiPosleGreske(ByVal poslednja As Variant) As Variant
    If IsArray(poslednja) Then
        BnKpiPosleGreske = poslednja
    Else
        BnKpiPosleGreske = BnKpiNepoznato()
    End If
End Function

' Brojke koje znace "ne znam". Znak nepoznatog nosi prva brojka (broj blokova);
' negativan broj brojac ne moze legitimno da bude, pa je to slobodan kanal
' kroz ugovor Scr_Brojac() As Long (ljuska ga crta kao "!").
Public Function BnKpiNepoznato() As Variant
    BnKpiNepoznato = Array(-1, -1, 0#, 0#)
End Function

Public Function BnKpiNepoznat(ByVal k As Variant) As Boolean
    If Not IsArray(k) Then
        BnKpiNepoznat = True
    Else
        BnKpiNepoznat = (CLng(k(0)) < 0)
    End If
End Function

'------------------------------------------------------- IZBOR U ZONI
' Cist racun platioca iz comboa (druga kolona), "" dok nista nije izabrano.
' NERAZRESEN UNOS NIJE RACUN: GetComboID daje stabilnu vrednost samo dok je
' stavka stvarno izabrana (ListIndex >= 0) -- ukucan deo racuna vraca "".
Private Function IzabraniRacun() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mRacunTest) > 0 Then
            IzabraniRacun = mRacunTest
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrBnRacun")
    If c Is Nothing Then Exit Function
    IzabraniRacun = GetComboID(c)
    Err.Clear
End Function

'=====================================================================
' DIJAGNOSTIKA
'
' Alt+F8 -> Diag_BnRedovi, pa Ctrl+G (Immediate). Ne menja nista.
'
' Isti razlog kao Diag_BuRedovi (katalog par. 9.10): celija mreze koja se
' vidi PRAZNA ili sa tudjim sadrzajem se ne razresava ni citanjem koda ni
' suite-om -- RenderGrid radi pod "On Error Resume Next", pa upis koji pukne
' ne ostavlja trag. Ispisuje se ono sto ekran PREDAJE mrezi i ono sto mreza
' od toga DRZI. Otvoren povod: smoke 28.08.2026, red bez racuna sa naizgled
' praznom celijom ISPLATITI (kolona 9).
'=====================================================================
Public Sub Diag_BnRedovi()
    Dim d As Variant, redovi As Variant, kolone As Variant, i As Long, n As Long
    Dim k As Long
    On Error Resume Next

    Debug.Print "--- Diag_BnRedovi (" & SCRBN_BUILD & ") ---"

    ' PRE naseg poziva: sta je LJUSKA poslednje trazila i sta je dobila.
    ' (Nas poziv ispod ce pregaziti trag -- zato se cita prvo.)
    Debug.Print "  POSLEDNJI POZIV: filter=[" & mDiagFilter & "] q=[" & mDiagQ & _
                "] vraceno redova=" & CStr(mDiagN)

    d = Scr_Rows("sve", "")
    If Not IsArray(d) Then
        Debug.Print "  Scr_Rows NIJE vratio niz"
        Exit Sub
    End If

    kolone = d(0)
    redovi = d(1)
    n = CLng(d(2))
    Debug.Print "  kolona=" & CStr(UBound(kolone) + 1) & "  redova=" & CStr(n)

    For i = 0 To UBound(kolone)
        Debug.Print "  spec " & CStr(i + 1) & ": " & CStr(kolone(i))
    Next i

    If IsArray(redovi) Then
        For i = 1 To 3
            If i > n Then Exit For
            For k = 1 To UBound(kolone) + 1
                Debug.Print "  EKRAN red " & CStr(i) & " kol" & CStr(k) & ": tip=" & _
                            TypeName(redovi(i, k)) & " vred=[" & CStr(redovi(i, k)) & "]"
            Next k
        Next i
    End If

    For i = 1 To 3
        For k = 1 To UBound(kolone) + 1
            Debug.Print "  MREZA red " & CStr(i) & " kol" & CStr(k) & ": tip=" & _
                        TypeName(modOtkupUI.GridCell(i, k)) & _
                        " vred=[" & CStr(modOtkupUI.GridCell(i, k)) & "]"
        Next k
    Next i

    Err.Clear
End Sub

'=====================================================================
' TEST SEAM
' Zona se u testu ne crta (forma se ne prikazuje), pa se stanje ekrana ne
' moze procitati iz kontrola. Ista kapija kao Scr_*Test na ostalim ekranima:
' seam koji MENJA stanje ekrana van test-rezima ne radi nista.
'=====================================================================
Public Sub Scr_BnRacunTestSet(ByVal racun As String)
    If Not IsTestMode() Then Exit Sub
    mRacunTest = racun
End Sub

' Dodavanje/uklanjanje bez mreze -- ide kroz ISTE rutine kao radnja nad redom.
Public Function Scr_BnKorpaTestDodaj(ByVal otkupID As String, ByVal broj As String, _
                                     ByVal otvoreno As Double, ByVal imaTR As Boolean) As String
    If Not IsTestMode() Then Exit Function
    Scr_BnKorpaTestDodaj = BnDodaj(otkupID, broj, otvoreno, imaTR)
End Function

Public Function Scr_BnTrakaRedTest(ByVal i As Long) As String
    If Not IsTestMode() Then Exit Function
    Scr_BnTrakaRedTest = TrakaRed(i)
End Function

' EKRANSKA putanja izvoza, ista koju zovu CSV i specifikacija -- ukljucujuci
' gate praznog izbora i prosledjivanje zadatih iznosa (", Iznosi()"). Bez
' ovoga bi domenska polovina bila dokazana, a jedan uklonjen argument u
' BlokoviZaIzvoz bi UI-ju pokazivao 250 a u fajl pustao 600 -- recenzija
' PR-a, tacka 3.
Public Function Scr_BnBlokoviZaIzvozTest(ByRef outBezTR As Long, _
                                         ByRef outIzbaceno As Long) As Collection
    If Not IsTestMode() Then Exit Function
    Set Scr_BnBlokoviZaIzvozTest = BlokoviZaIzvoz(outBezTR, outIzbaceno)
End Function

' Direktan upis zadatog iznosa BEZ validacije -- da se izmeri da citanje
' liste zaostali iznos STVARNO klampuje (validan unos preko otvorenog ne
' postoji, pa se stanje "zadato > otvoreno" bez ovoga ne moze ni napraviti).
Public Sub Scr_BnIznosTestSet(ByVal otkupID As String, ByVal iznos As Double)
    If Not IsTestMode() Then Exit Sub
    Iznosi()(otkupID) = iznos
End Sub

Public Function Scr_BnIznosZaTest(ByVal otkupID As String, ByVal otvoreno As Double) As Double
    If Not IsTestMode() Then Exit Function
    Scr_BnIznosZaTest = BnIznosZa(otkupID, otvoreno)
End Function

Public Function Scr_BnIznosPostojiTest(ByVal otkupID As String) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_BnIznosPostojiTest = Iznosi().Exists(otkupID)
End Function

' Koliko je puta snimak STVARNO citan iz tabela -- jedini nacin da se izmeri
' da pretraga ne placa pun prolaz po otkucaju (smoke 3).
Public Function Scr_BnSnimakPunjenjaTest() As Long
    If Not IsTestMode() Then Exit Function
    Scr_BnSnimakPunjenjaTest = mSnimakPunjenja
End Function

Public Sub Scr_BnTestReset()
    If Not IsTestMode() Then Exit Sub
    mLista = BN_NALOZI
    Set mKorpa = New Collection
    Set mIznosi = Nothing
    mRacunTest = ""
    mRacunPunjen = False
    mSnimakPunjenja = 0
    Scr_ResetCache
End Sub
