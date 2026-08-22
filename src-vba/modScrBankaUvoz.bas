Attribute VB_Name = "modScrBankaUvoz"
'=====================================================================
' modScrBankaUvoz - ekran "Uvoz izvoda" (v6-ui-177). Faza E, stavka 17.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19). Red u registru
' (modUiScreens.ScrRows) je postojao od S3a -- stavka menija se do sada crtala
' prigusena jer modula nije bilo. Registar se NE dira.
'
' ODAKLE DOLAZI: frmBankaImport puni ListBox otvorenim stavkama izvoda, pa se
' nad izabranom stavkom radi auto ili rucno mapiranje. Mreza ljuske radi isto,
' samo sto ono sto je forma imala u statusnoj liniji (saldo poslednjeg izvoda)
' ovde postaje SVOJA LISTA nad svim izvodima.
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED i PRIKAZ. Nijedno poslovno pravilo,
' nijedna kapija i nijedan upis nisu ovde:
'   - auto mapiranje reda / svih  -> modBankaMapiranje.AutoMap*_TX
'   - jaki kljucevi               -> modBankaMapiranje.AutoMapStrongKeysBankaImport_TX
'   - rucno mapiranje             -> modBankaMapiranje.MapBankaImportAs*_TX
'   - preskakanje                 -> modBankaMapiranje.SkipBankaImportRow_TX
'   - predlog podele po bloku     -> modBankaMapiranje.PlanBlokRaspodela
'   - smer stavke                 -> modBankaMapiranje.ClassifyBimSmer
'   - jak kljuc po redu           -> modBankaMapiranje.BimJakiKljucInfo
'   - redovi mreze                -> modBankaMapiranje.GetBankaImportForGrid
'                                    modBankaImport.GetBankaIzvodiForGrid
'   - integritet izvoda           -> modBankaImport.BimSaldoStatus
'
' DVE LISTE u deljenoj mrezi (prekidac iznad nje):
'   STAVKE   red za mapiranje; pet radnji nad redom
'   IZVODI   agregat po BimIzvodKljuc (broj + racun + datum): pocetno, uplate,
'            isplate,
'            zavrsno i da li se slaze. Legacy je isto to imao u JEDNOJ labeli i
'            samo za NAJNOVIJI izvod (UpdateIzvodSummaryLabel).
'
' ZASTO IZVODI JESU LISTA a "obradjeno"/"preskoceno" nisu: izvod je DRUGA
' forma podatka (agregat po izvodu, druge kolone, drugi identitet), a status
' stavke je isti citac sa filterom -- to su cipovi. Zasebna lista po statusu
' bila bi druga kopija istog citaca koja moze da se razidje.
'
' STA NIJE PRENETO: UVOZ (povlacenje PDF-ova, parsiranje, staging). Razlog nije
' duzina posla nego ishod: ImportBankaInbox_TX je Sub koji NE VRACA nista --
' SaveBankaImportRowsCore prebroji i upisane i duplikate, ali oba zavrse u
' Debug.Print. Dugme koje ne moze da kaze "uvezeno N, duplikata M" bilo bi tiho
' knjizenje, a uz to pomera i fajlove po disku (Inbox -> Processed/Error). Uz
' to, ni frmBankaImport nema uvozno dugme -- uvoz je oduvek zasebna komanda, pa
' ovaj ekran time NIJE uzi od legacy-ja.
'
' POLJA SU LJUSKINA, NE EKRANOVA. Sklop "natpis + shell + kontrola" pravi
' modOtkupUI.NewFieldG, raspored unutar polja modOtkupUI.LayoutFieldInner.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRBU_BUILD As String = "v6-ui-177"

' Visina zone: JEDAN red -- tri combo-a i objasnjenje uz njih. Dugmadi u zoni
' nema: sve radnje ovog ekrana su radnje NAD REDOM i zive u redu radnji mreze.
'
' Objasnjenje je u ISTOJ liniji sa poljima, desno od njih. Ispod njih je stajalo
' preko trake koju ljuska crta odmah ispod zone.
Private Const BU_ZONA_H   As Single = 104

Private Const BU_Y_CAP    As Single = 6
Private Const BU_Y_KPI_V  As Single = 18
Private Const BU_Y_LBL    As Single = 48
' Poravnato sa unosnim kutijama polja (NewFieldG ih stavlja na +16, visine 28).
Private Const BU_Y_HINT   As Single = 72
Private Const BU_KPI_W    As Single = 116
Private Const BU_FLD_W    As Single = 190
Private Const BU_FLD_GAP  As Single = 10
' Ispod ove sirine objasnjenje nema gde -- bolje bez njega nego preko brojki.
Private Const BU_HINT_MIN As Single = 100

' Kljucevi lista
Private Const BU_STAVKE As String = "STAVKE"
Private Const BU_IZVODI As String = "IZVODI"

' SKRIVENE KOLONE, prioritet 4 -- LayoutGrid crta do 3, pa vrednost postoji u
' modelu a celija se nikad ne pravi.
'
' Identitet ide U RED, ne pored njega: mreza redove sortira i deli na strane, pa
' bi svaka mapa "prikaz -> ID" koju ekran drzi sa strane zastarela na prvi klik
' po zaglavlju. Prazan identitet znaci DVOSMISLEN (isti BankaImportID postoji
' dvaput) i radnja tada ODBIJA da bira.
Private Const BU_STV_KOL_ID As Long = 10
' Otvorenost se NE cita iz prikazanog statusa: nov red ima PRAZAN status, pa se
' iz prikaza ne razlikuje od reda kome status nije upisan. Red je NOSI.
Private Const BU_STV_KOL_OTV As Long = 11
' Smer se NE izvodi iz toga koja je kolona iznosa popunjena: red sa I uplatom I
' isplatom izgleda kao uplata, a writer ga odbija (RequireBimSmer). Red ga NOSI.
Private Const BU_STV_KOL_SMER As Long = 12
Private Const BU_IZV_KOL_ID As Long = 10

Private mLista As String            ' BU_STAVKE | BU_IZVODI

Private mCombosPunjeni As Boolean
Private mTipPunjen As String        ' za koji je tip partner combo napunjen
Private mCiljPunjen As String       ' tip + "|" + partnerID za koji je cilj punjen
Private mFill As Boolean            ' punjenje comboa okida Change - v. mPopMute u ljusci

' Da li je lista faktura za izabranog kupca stvarno UCITANA. Prazan combo posle
' PADA ucitavanja izgleda isto kao "kupac nema otvorenih faktura", a prazan
' izbor znaci "knjizi kao AVANS". Zato se pad pamti i knjizenje kupca se
' blokira dok se ne osvezi. Isto pravilo koje frmBankaImport nosi u
' m_FaktureLoadOk; sam citac je izdvojen u modBankaMapiranje.
' UCITANOST LISTE CILJA -- vazi za OBE rucne rute, ne samo za fakture.
' Prazna lista nosi poslovno znacenje na obe: prazan izbor fakture je AVANS,
' prazan izbor bloka je "uzmi poziv na broj". Pad punjenja zato ne sme da se
' pretvori ni u jedno od to dvoje. v. CiljUcitan.
Private mCiljOK As Boolean
Private mCiljErr As String
' Koliko je puta punjenje liste POZVANO. Postoji zato sto se bez forme ne moze
' videti da li ga je kapija stvarno zvala: bez kontrole PuniCiljCombo izlazi
' odmah, pa uklonjen poziv ne menja nijedan drugi merljiv ishod -- a bas taj
' nedostajuci poziv je bio kvar (kapija je sudila po zastavici tudjeg izbora).
Private mCiljPunjenja As Long

' Kes cetiri brojke zone. Svaka je pun prolaz kroz tabelu, a OsveziZonu se zove
' pri svakom citanju mreze. Cisti ga Scr_ResetCache, koju ljuska zove posle
' svakog upisa (RefreshFromData).
Private mKpi As Variant
Private mKpiOK As Boolean

' Izbori koje je postavio TEST. Zone u testu nema (forma se ne prikazuje), pa
' se combo ne moze procitati. Vaze SAMO u test rezimu.
Private mTipTest As String
Private mPartnerTest As String
Private mCiljTest As String
Private mStanicaTest As String

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=BANKA_UVOZ|naslov=OTKUI_NAV_BANKA_UVOZ|sub=OTKUI_SCRBU_SUB" & _
               "|lista=OTKUI_SCRBU_LISTA|oblik=zona+mreza|upis=radnja"
End Function

Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        BU_STAVKE & "|OTKUI_SEG_BU_STAVKE|OTKUI_GRID_TITLE_BU_STAVKE|64", _
        BU_IZVODI & "|OTKUI_SEG_BU_IZVODI|OTKUI_GRID_TITLE_BU_IZVODI|60")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = BU_STAVKE
    Scr_Lista = mLista
End Function

' Scr_NaslovDopuna NAMERNO NE POSTOJI. Naslov mreze je labela fiksne sirine
' (grdTitle, 180pt), pa se dopuna odsecala usred reci ("-- 29 z"), a broj koji
' je nosila vec stoji u brojci OTVORENO iznad mreze i u cipu "za obradu".
' Odsecen tekst je gori od nikakvog.

' Prvi cip je svuda "sve" -- ljuska na njega pada kad zatecen filter ne pripada
' listi na koju se upravo preslo (RefreshChipsForScreen). Zato prvi mora da
' bude NAJSIRI: povratak na uzi cip bi tiho sakrio redove.
Public Function Scr_Cipovi() As String
    Scr_Cipovi = BuCipoviZaListu(Scr_Lista())
End Function

' Cipovi PO KLJUCU LISTE, odvojeno od Scr_Cipovi -- da se ugovor svake liste
' moze izmeriti bez prebacivanja stanja ekrana.
Public Function BuCipoviZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case BU_STAVKE
            BuCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "zaobradu:OTKUI_CIPB_ZAOBRADU:80|" & _
                              "jaki:OTKUI_CIPB_JAKI:92|" & _
                              "rucno:OTKUI_CIPB_RUCNO:72|" & _
                              "obradjeno:OTKUI_CIPB_OBRADJENO:84|" & _
                              "preskoceno:OTKUI_CIPB_PRESKOCENO:88"
        Case BU_IZVODI
            BuCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "otvoreni:OTKUI_CIPB_OTVORENI:96|" & _
                              "razlika:OTKUI_CIPB_RAZLIKA:88"
    End Select
End Function

' PRAVILO CIPA STAVKE. Kljuc je EKRANOV -- ljuska ga je samo vratila onakvog
' kakvog ga je dobila iz Scr_Cipovi. Nepoznat i prazan kljuc PUSTAJU sve: ekran
' koji dobije filter koji ne poznaje pokazuje punu listu, ne praznu.
'
' "Za obradu" i "za rucno" NISU isto: red sa statusom "Error" je auto vec
' pokusao i odbio, ali je i dalje OTVOREN (GetBankaImportOpen izbacuje samo
' "Da" i "Skip"). Bez oba cipa se ne vidi razlika izmedju "jos nije probano" i
' "probano pa vraceno operateru".
Public Function BuCipStavka(ByVal filter As String, ByVal obradjeno As String, _
                            ByVal jaki As Boolean) As Boolean
    Dim s As String
    s = Trim$(obradjeno)
    Select Case filter
        Case "zaobradu":   BuCipStavka = modBankaMapiranje.BimOtvoren(s)
        Case "jaki":       BuCipStavka = modBankaMapiranje.BimOtvoren(s) And jaki
        Case "rucno":      BuCipStavka = (s = BIM_OBR_ERROR)
        Case "obradjeno":  BuCipStavka = (s = BIM_OBR_DA)
        Case "preskoceno": BuCipStavka = (s = BIM_OBR_SKIP)
        Case Else:         BuCipStavka = True
    End Select
End Function

' PRAVILO CIPA IZVODA. "Ne slaze se" je BIM_SALDO_RAZLIKA i samo ona: legacy
' red bez saldo metapodataka (sva cetiri broja nula) NIJE neslaganje nego
' odsustvo podatka, i ne sme da se prikaze kao greska.
Public Function BuCipIzvod(ByVal filter As String, ByVal status As Long, _
                           ByVal otvorenih As Long) As Boolean
    Select Case filter
        Case "otvoreni": BuCipIzvod = (otvorenih > 0)
        Case "razlika":  BuCipIzvod = (status = BIM_SALDO_RAZLIKA)
        Case Else:       BuCipIzvod = True
    End Select
End Function

' Radnje nad izabranim redom. Peto polje je "trebaRed": jaki kljucevi i
' automatsko mapiranje SVIH rade bez izabranog reda.
Public Function Scr_Radnje() As String
    Scr_Radnje = BuRadnjeZaListu(Scr_Lista())
End Function

' Radnje PO KLJUCU LISTE -- isti razlog kao BuCipoviZaListu.
'
' Tacno PET na listi stavki, sto je granica bazena (MAX_ACT). Visak se tiho
' odseca (RefreshRowActions radi Exit For), pa bi operater dobio ekran kome
' fali dugme, bez ijedne poruke -- isti kvar je vec placen na listi paleta
' (v6-ui-162) i na SEF listi (v6-ui-176). Sesta radnja ovde ne sme da udje bez
' zasebne liste. "Osvezi" nije medju njima: to je posao ljuske.
'
' IZVODI NEMAJU RADNJE. To je pregled: nijedna operacija se ne radi nad
' izvodom kao celinom, a veza ka njegovim stavkama ide kroz pretragu ljuske
' (haystack liste stavki nosi i broj izvoda i broj racuna). Prazan spisak
' ljuska podnosi -- ActDefs vrati Empty, a raspored SAKRIJE zaostalu dugmad.
Public Function BuRadnjeZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case BU_STAVKE
            BuRadnjeZaListu = "bmauto:OTKUI_BTN_BU_AUTO:112:primary:1|" & _
                              "bmrucno:OTKUI_BTN_BU_RUCNO:112:soft:1|" & _
                              "bmskip:OTKUI_BTN_BU_SKIP:80:ghost:1|" & _
                              "bmjaki:OTKUI_BTN_BU_JAKI:104:soft:0|" & _
                              "bmsve:OTKUI_BTN_BU_SVE:116:ghost:0"
    End Select
End Function

' Znacka uz stavku menija. Red za mapiranje je PODATAK U TABELI, ne prolazno
' stanje ekrana: svaka radnja ga menja upisom, a ljuska posle upisa ionako zove
' RefreshFromData -- pa je brojac time vec pokriven i privatan kanal ka
' OsveziNavBrojace ovde NE treba (za razliku od korpe na Agrohemiji i
' Fakturisanju, koja u tabeli ne postoji).
'
' Broj se cita iz iste brojke koju vidi i cip "za obradu" -- ne iz zasebnog
' prolaza koji bi se s njim mogao raziciti.
Public Function Scr_Brojac() As Long
    Dim k As Variant
    k = Kpi()
    Scr_Brojac = CLng(k(0))
End Function

Public Sub Scr_ResetCache()
    ' mKpi se NE brise, samo proglasava zastarelim. Ako sledece citanje pukne,
    ' bolje je zadrzati poslednju poznatu brojku nego je zameniti nulom --
    ' v. Kpi i BuKpiPosleGreske.
    mKpiOK = False
    mCiljPunjen = ""
End Sub

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a: modLogError.LogError pocinje sa
    ' "On Error Resume Next", a svaka On Error naredba brise Err -- posle njega
    ' bi poruka operateru ostala bez uzroka.
    errDesc = Err.description
    LogErr "modScrBankaUvoz.Scr_Event"
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

    ' Izbor reda ne menja podatke. Menja PREDLOZEN tip mapiranja -- isto sto
    ' frmBankaImport.ShowSelectedRow radi na svaki klik po listi: uplata
    ' predlaze Kupca, isplata Kooperanta. Predlog, ne odluka: operater ga menja.
    If Left$(tag, 4) = "row:" Then
        PredloziTipZaRed CLng(val(Mid$(tag, 5)))
        Exit Function
    End If

    ' Promena u polju zone stize kao "chg:<tag kontrole>"; vrednost se cita iz
    ' same kontrole, ne iz taga.
    If Left$(tag, 4) = "chg:" Then
        ObradiKlik = ObradiPromenu(Mid$(tag, 5))
        Exit Function
    End If

    ' DVOKLIK NAMERNO NE RADI NISTA. Na Fakturisanju dvoklik prebacuje red u
    ' korpu i iz nje -- povratna radnja nad prolaznim stanjem. Ovde bi svaka
    ' radnja nad redom bila KNJIZENJE u tblNovac, a knjizenje se ne pokrece
    ' promasenim dvoklikom.
    If Left$(tag, 4) = "dbl:" Then Exit Function

    ' Radnja nad redom stize kao "act:<kljuc>:<red>".
    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(Mid$(tag, 5))
        Exit Function
    End If
End Function

' NERAZRESEN UNOS NIJE PROMENA. Ljuska Change salje ekranu na SVAKI otkucaj, a
' combo daje stabilnu vrednost samo dok je stavka stvarno izabrana
' (ListIndex >= 0). Uslov oblika "novo <> staro" je zato tacan vec na prvom
' slovu -- na Fakturisanju je bas takav uslov bacao celu korpu (review R2).
' Ovde nista ne pada, ali bi svaki znak povukao ponovno punjenje comboa nad
' tabelama.
'
' LISTA NE ZAVISI OD POLJA ZONE, pa se modOtkupUI.RefreshFromData ovde NE zove:
' polja biraju CILJ rucnog mapiranja, ne skup redova. (Fakturisanje mora, jer
' mu je lista prijemnica lista jednog kupca -- katalog par. 8.2.)
Private Function ObradiPromenu(ByVal tag As String) As Boolean
    If mFill Then Exit Function
    Select Case tag
        Case "scrBuTipT"
            If Len(IzabraniTip()) = 0 Then Exit Function
            OsveziZonu

        Case "scrBuPartnerT"
            If Len(IzabraniPartnerID()) = 0 Then Exit Function
            OsveziZonu

        Case "scrBuCiljT"
            OsveziObjasnjenjeSam
    End Select
End Function

Private Function RadnjaNadRedom(ByVal spec As String) As Boolean
    Dim p() As String, red As Long, kljuc As String
    p = Split(spec, ":")
    If UBound(p) < 1 Then Exit Function
    kljuc = p(0)
    red = CLng(val(p(1)))

    ' Batch radnje ne traze izabran red (peto polje u BuRadnjeZaListu je 0), pa
    ' se odvajaju PRE provere reda.
    Select Case kljuc
        Case "bmjaki"
            RadnjaNadRedom = JakiKljucevi()
            Exit Function
        Case "bmsve"
            RadnjaNadRedom = AutoSve()
            Exit Function
    End Select

    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    Select Case kljuc
        Case "bmauto":  RadnjaNadRedom = AutoRed(red)
        Case "bmrucno": RadnjaNadRedom = RucnoRed(red)
        Case "bmskip":  RadnjaNadRedom = PreskociRed(red)
    End Select
End Function

' Identitet iza prikazanog reda. PRAZNO znaci DVOSMISLENO -- isti BankaImportID
' postoji dvaput u tabeli -- i tada radnja ODBIJA da bira umesto da pogodi.
' Dvosmislenost prepoznaje CITAC (modFaktura.IdIliPrazno nad sirovom tabelom);
' ovde se samo ne pogadja. Bez toga bi radnja svakako pukla -- RequireSingleRow
' fail-close-uje na duplikat -- ali kao greska transakcije umesto kao poruka.
Private Function IdReda(ByVal red As Long, ByVal kol As Long) As String
    Dim iD As String
    If red < 1 Then Exit Function
    iD = Trim$(CStr(modOtkupUI.GridCell(red, kol)))
    If Len(iD) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_DVOSMISLEN"), True
        Exit Function
    End If
    IdReda = iD
End Function

' Sme li se nad redom jos raditi. Cita se ono sto red NOSI, ne ono sto se u
' njemu vidi -- v. BU_STV_KOL_OTV.
Private Function RedOtvoren(ByVal red As Long) As Boolean
    RedOtvoren = (Trim$(CStr(modOtkupUI.GridCell(red, BU_STV_KOL_OTV))) = "1")
End Function

Private Function RedSmer(ByVal red As Long) As String
    RedSmer = Trim$(CStr(modOtkupUI.GridCell(red, BU_STV_KOL_SMER)))
End Function

Private Function RedOznaka(ByVal red As Long) As String
    On Error Resume Next
    RedOznaka = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    Err.Clear
End Function

' Predlog tipa iz smera reda -- isto sto frmBankaImport.ShowSelectedRow radi.
' Ne dira izbor ako smer nije cist: nejasan red nema sta da predlozi.
Private Sub PredloziTipZaRed(ByVal red As Long)
    Dim smer As String, tip As String
    On Error Resume Next
    If red < 1 Then Exit Sub
    If Scr_Lista() <> BU_STAVKE Then Exit Sub
    smer = RedSmer(red)
    If smer = BIM_SMER_UPLATA Then
        tip = BIM_TIP_KUPAC
    ElseIf smer = BIM_SMER_ISPLATA Then
        tip = BIM_TIP_KOOPERANT
    Else
        Exit Sub
    End If
    If IzabraniTip() = tip Then Exit Sub
    PostaviTip tip
End Sub

Private Sub PostaviTip(ByVal tip As String)
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrBuTip")
    If c Is Nothing Then Exit Sub
    mFill = True
    c.value = tip
    mFill = False
    OsveziZonu
End Sub

'=====================================================================
' RADNJE NAD REDOM
'=====================================================================
Private Function AutoRed(ByVal red As Long) As Boolean
    Dim bimID As String, rezultat As String
    bimID = StavkaZaRad(red)
    If Len(bimID) = 0 Then Exit Function

    rezultat = modBankaMapiranje.AutoMapBankaImportRow_TX(bimID)

    ' Prazan rezultat NIJE greska nego ishod: auto nije nasao jednoznacan cilj i
    ' red je oznacen za rucno. Mreza se svejedno promenila (status), pa se vraca
    ' True da bi ljuska osvezila listu.
    Scr_ResetCache
    If Len(rezultat) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_AUTO"), True
    Else
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_AUTO"), False
    End If
    AutoRed = True
End Function

Private Function PreskociRed(ByVal red As Long) As Boolean
    Dim bimID As String
    bimID = StavkaZaRad(red)
    If Len(bimID) = 0 Then Exit Function

    If MsgBox(Poruka("OTKUI_ASK_BU_SKIP") & " " & RedOznaka(red) & "?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    If Not modBankaMapiranje.SkipBankaImportRow_TX(bimID) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_SKIP"), True
        Exit Function
    End If

    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_SKIP"), False
    PreskociRed = True
End Function

' Batch: knjizi samo stavke sa jednoznacnim jakim kljucem. Pad se PRIJAVLJUJE --
' ceo pass je tada rollback-ovan, a tiho "0 mapirano" je izgledalo kao "nema sta
' da se mapira" (AUD-014).
Private Function JakiKljucevi() As Boolean
    Dim n As Long, errDesc As String
    If MsgBox(Poruka("OTKUI_ASK_BU_JAKI"), vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    On Error GoTo EH
    n = modBankaMapiranje.AutoMapStrongKeysBankaImport_TX()
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_JAKI") & " " & CStr(n), False
    JakiKljucevi = True
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrBankaUvoz.JakiKljucevi"
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_BATCH") & " " & errDesc, True
    JakiKljucevi = True
End Function

Private Function AutoSve() As Boolean
    Dim n As Long, zaRucno As Long, tekst As String, errDesc As String
    If MsgBox(Poruka("OTKUI_ASK_BU_SVE"), vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    On Error GoTo EH
    n = modBankaMapiranje.AutoMapAllBankaImport_TX(zaRucno)
    Scr_ResetCache
    tekst = Poruka("OTKUI_MSG_BU_SVE") & " " & CStr(n)
    If zaRucno > 0 Then tekst = tekst & Poruka("OTKUI_MSG_BU_SVE_RUCNO") & " " & CStr(zaRucno)
    modOtkupUI.ShowToast tekst, False
    AutoSve = True
    Exit Function
EH:
    ' Batch koji padne je rollback-ovan i PROPAGIRA gresku. Bez ove grane bi
    ' operater posle greske dobio jos i "mapirano: 0", sto izgleda kao uredno
    ' zavrsen batch bez pogodaka.
    errDesc = Err.description
    LogErr "modScrBankaUvoz.AutoSve"
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_BATCH") & " " & errDesc, True
    AutoSve = True
End Function

' Identitet + otvorenost na jednom mestu: obe provere traze sve tri radnje nad
' redom, i obe citaju ono sto red NOSI.
Private Function StavkaZaRad(ByVal red As Long) As String
    Dim bimID As String
    If Scr_Lista() <> BU_STAVKE Then Exit Function
    bimID = IdReda(red, BU_STV_KOL_ID)
    If Len(bimID) = 0 Then Exit Function
    If Not RedOtvoren(red) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_ZATVOREN"), True
        Exit Function
    End If
    StavkaZaRad = bimID
End Function

'=====================================================================
' RUCNO MAPIRANJE
'
' Tri spregnuta polja zone: TIP (Kupac / Kooperant / OM), PARTNER i CILJ
' (faktura za kupca, blok za kooperanta; za OM cilja nema). Sva pravila su u
' modBankaMapiranje -- ovde je samo redosled pitanja i poruka.
'=====================================================================
Private Function RucnoRed(ByVal red As Long) As Boolean
    Dim bimID As String, tip As String, partnerID As String

    bimID = StavkaZaRad(red)
    If Len(bimID) = 0 Then Exit Function

    tip = IzabraniTip()
    If Len(tip) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_NEMA_TIPA"), True
        Exit Function
    End If

    partnerID = IzabraniPartnerID()
    If Len(partnerID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_NEMA_PARTNERA"), True
        Exit Function
    End If

    ' SMER-KAPIJA PRE KLIKA. Writer je ima (RequireBimSmer) i zadrzava je, ali
    ' bi je operater osetio tek kao gresku transakcije. Legacy je isto ovo
    ' pokazivao u preview-u pre klika (AUD-025).
    If Not modBankaMapiranje.BimSmerOdgovaraTipu(RedSmer(red), tip) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_SMER"), True
        Exit Function
    End If

    If MsgBox(Poruka("OTKUI_ASK_BU_RUCNO") & " " & RedOznaka(red) & "?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    Select Case tip
        Case BIM_TIP_KUPAC:     RucnoRed = RucnoKupac(bimID, partnerID)
        Case BIM_TIP_KOOPERANT: RucnoRed = RucnoKooperant(bimID, partnerID)
        Case BIM_TIP_OM:        RucnoRed = RucnoOM(bimID, partnerID)
    End Select

    If RucnoRed Then Scr_ResetCache
End Function

Private Function RucnoKupac(ByVal bimID As String, ByVal kupacID As String) As Boolean
    Dim fakturaID As String
    Dim greska As String

    ' AKO LISTA FAKTURA NIJE UCITANA, prazan izbor NE znaci "nema fakture" nego
    ' "ne znamo" -- a knjizenje avansa na osnovu takve liste je pogadjanje.
    ' Citac vraca zastavicu bas zbog ovoga (GetFaktureZaBimMapiranje).
    If Not CiljUcitan(greska) Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If

    fakturaID = IzabraniCiljID()

    ' Bez izabrane fakture uplata NIJE zatvaranje duga nego AVANS. To se nikad
    ' ne sme desiti tiho (FM-0024 #2).
    If Len(fakturaID) = 0 Then
        If MsgBox(Poruka("OTKUI_ASK_BU_AVANS"), vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function
    End If

    If Len(modBankaMapiranje.MapBankaImportAsKupac_TX(bimID, kupacID, fakturaID, True)) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_RUCNO"), True
        Exit Function
    End If

    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_RUCNO"), False
    RucnoKupac = True
End Function

' FAIL-CLOSED KAPIJA nad listom faktura. Prazna lista i PAD ucitavanja
' izgledaju isto, a znace suprotno: prazan izbor fakture knjizi AVANS. Odluka
' "smem li uopste da radim" je zato imenovana i odvojena od crtanja -- inace se
' ne moze izmeriti, a brisanje jednog If-a se ne bi videlo ni u jednom testu.
' SME LI RUCNO MAPIRANJE UOPSTE DA POCNE.
'
' Zajednicka za kupca i kooperanta, jer je opasnost ista: prazna lista cilja na
' obe rute znaci nesto konkretno -- prazan izbor fakture je AVANS, prazan izbor
' bloka je "uzmi poziv na broj" -- pa neuspelo punjenje ne sme da se pretvori u
' to. Kod kooperanta je posledica i grublja: ako iz poziva na broj ne ispadne
' nijedan kandidat, MapBankaImportAsKooperantBlockCore ceo iznos knjizi kao
' avans kooperanta i stavku oznaci obradjenom. Kvar tako postane USPESNO
' knjizenje drugog poslovnog ishoda.
Private Function BuSmeMapiranjeCilja() As Boolean
    BuSmeMapiranjeCilja = mCiljOK
End Function

' Puni listu cilja i kaze sme li se dalje. Obe rucne rute prolaze kroz OVO --
' dve kopije istog uslova bi se razisle, a prva je vec bila samo kod kupca.
Private Function CiljUcitan(ByRef outPoruka As String) As Boolean
    PuniCiljCombo
    CiljUcitan = BuSmeMapiranjeCilja()
    If Not CiljUcitan Then outPoruka = Poruka("OTKUI_ERR_BU_CILJ") & " " & mCiljErr
End Function

Private Function RucnoKooperant(ByVal bimID As String, ByVal kooperantID As String) As Boolean
    Dim blok As String, razlog As String, n As Long
    Dim potvrdjeno As Boolean
    Dim scope As String
    Dim stani As Boolean
    Dim greska As String

    ' AKO LISTA BLOKOVA NIJE UCITANA, prazan izbor NE znaci "operater nije birao
    ' blok". Fallback na poziv na broj bi tada bio pogadjanje, a scope bi ispao
    ' prazan -- pa bi raspodela zahvatila sva otkupna mesta sa tim brojem. Ako
    ' kandidata uopste nema, ceo iznos se knjizi kao avans kooperanta i stavka
    ' se oznacava obradjenom: kvar postaje uspesno knjizenje drugog ishoda.
    If Not CiljUcitan(greska) Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If

    ' Tek sad prazan izbor legitimno znaci "uzmi poziv na broj iz izvoda".
    blok = modBankaMapiranje.BimEfektivniBlok(bimID, IzabraniCiljID())

    ' Scope i odluka o zaustavljanju racunaju se na JEDNOM mestu -- v.
    ' ScopeIzbora. Izabran blok bez otkupnog mesta staje PRE ijednog citanja
    ' kandidata.
    scope = ScopeIzbora(stani)
    If stani Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_BLOK_BEZ_OM"), True
        Exit Function
    End If

    If modBankaMapiranje.BimBlokTraziPotvrdu(kooperantID, blok, razlog, scope) Then
        Select Case PitajZaPodelu(bimID, kooperantID, blok, scope)
            Case vbCancel
                Exit Function
            Case vbNo
                ' Bezbedan izlaz: dok je poreklo dvosmisleno, nista se ne vezuje
                ' za otkup. Ceo iznos ide kao avans kooperanta, a vezuje se
                ' kasnije dugmetom "Primeni avans na blok" u Banka izvestaju.
                If Len(modBankaMapiranje.MapBankaImportAsKooperant_TX( _
                        bimID, kooperantID, "", "", True)) = 0 Then
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_RUCNO"), True
                    Exit Function
                End If
                modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_AVANS_OK"), False
                RucnoKooperant = True
                Exit Function
            Case Else
                potvrdjeno = True
        End Select
    End If

    n = modBankaMapiranje.MapBankaImportAsKooperantBlockManual_TX( _
            bimID, kooperantID, blok, True, potvrdjeno, scope)

    If n <= 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_RUCNO"), True
        Exit Function
    End If

    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_RUCNO"), False
    RucnoKooperant = True
End Function

Private Function RucnoOM(ByVal bimID As String, ByVal omID As String) As Boolean
    If Len(modBankaMapiranje.MapBankaImportAsOM_TX(bimID, omID, "", True)) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_RUCNO"), True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_BU_RUCNO"), False
    RucnoOM = True
End Function

' Blok sa vise otvorenih stavki nego sto automatska raspodela sme da podeli:
' operater vidi TACNU podelu (isti planer po kome se knjizi) i bira izmedju tri
' ishoda. Ostaje MsgBox, kao u legacy-ju: tri ishoda nad izracunatim tekstom
' nisu forma nego pitanje.
Private Function PitajZaPodelu(ByVal bimID As String, ByVal kooperantID As String, _
                               ByVal blok As String, ByVal scope As String) As VbMsgBoxResult
    Dim kandidati As Variant, iznos As Double

    PitajZaPodelu = vbCancel

    ' ISTI scope kao kod knjizenja -- inace bi operater potvrdio jednu podelu, a
    ' knjizila bi se druga.
    kandidati = modBankaMapiranje.GetOtkupCandidatesForKooperantBlock( _
                    kooperantID, blok, True, scope)
    ' Kandidata nema, a granica je malopre bila prekoracena -- podaci su se
    ' promenili izmedju dva citanja. Tiho odustajanje bi izgledalo kao dugme
    ' koje ne radi, pa se prijavljuje.
    If IsEmpty(kandidati) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_BU_RUCNO"), True
        Exit Function
    End If

    iznos = CDbl(NzBIM(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_ISPLATA), 0#))

    PitajZaPodelu = MsgBox( _
        Poruka("OTKUI_ASK_BU_PODELA") & vbCrLf & vbCrLf & _
        blok & "  " & ChrW(183) & "  " & Format$(iznos, "#,##0.00") & vbCrLf & _
        TekstPodele(kandidati, iznos) & vbCrLf & _
        Poruka("OTKUI_ASK_BU_PODELA_IZBOR"), _
        vbQuestion + vbYesNoCancel, APP_NAME)
End Function

' Tekst predlozene podele. Racuna ga PlanBlokRaspodela -- ISTI planer po kome
' MapBankaImportAsKooperantBlockCore knjizi, pa prikaz i akcija ne mogu da se
' razidju. Visak preko otvorenih stavki ide u avans, kao i u knjizenju.
' Odvojen od MsgBox-a da bi se mogao izmeriti bez dijaloga.
Public Function TekstPodele(ByVal kandidati As Variant, ByVal iznos As Double) As String
    Dim plan As Variant, s As String, i As Long, podeljeno As Double

    plan = modBankaMapiranje.PlanBlokRaspodela(kandidati, iznos)

    If Not IsEmpty(plan) Then
        For i = 1 To UBound(plan, 1)
            s = s & " - " & CStr(plan(i, 1)) & ": " & _
                Format$(CDbl(plan(i, 2)), "#,##0.00") & " (" & CStr(plan(i, 3)) & ")" & vbCrLf
            podeljeno = podeljeno + CDbl(plan(i, 2))
        Next i
    End If

    If iznos - podeljeno > 0.009 Then
        s = s & " - " & Poruka("OTKUI_ASK_BU_PODELA_VISAK") & " " & _
            Format$(iznos - podeljeno, "#,##0.00") & " -> " & NOV_VIRMAN_AVANS_KOOP & vbCrLf
    End If

    TekstPodele = s
End Function

'=====================================================================
' REDOVI MREZE
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' Zona se puni odavde, kao i na ekranima Palete, Agrohemija i Fakturisanje:
    ' gradi se jednom, a podaci za nju postoje tek kad se lista cita.
    OsveziZonu
    If Scr_Lista() = BU_IZVODI Then
        Scr_Rows = RedoviIzvodi(filter, q)
        Exit Function
    End If
    Scr_Rows = RedoviStavke(filter, q)
End Function

' Opis kolona PO KLJUCU LISTE. Postoji da bi se pravilo "identitet je u redu i
' NE crta se" moglo tvrditi za SVAKU listu, bez prebacivanja stanja ekrana.
Public Function BuKoloneZaListu(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case BU_IZVODI: BuKoloneZaListu = IzvodiKolone()
        Case Else:      BuKoloneZaListu = StavkeKolone()
    End Select
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

'------------------------------------------------------- LISTA: STAVKE
Private Function StavkeKolone() As Variant
    ' Prva kolona se uvek crta kao BROJ dokumenta (StyleGridCell, isBroj) -- tu
    ' stoji BROJ IZVODA, jedini POSLOVNI broj koji stavka nosi.
    '
    ' BankaImportID se NE PRIKAZUJE. Legacy ga je imao u koloni "BIM", ali to je
    ' interna sifra: operater ne zna cemu sluzi i ne moze nista sa njom. Identitet
    ' i dalje ide U RED -- u skrivenu kolonu, gde mu je i mesto.
    '
    ' Poslednje tri nose ono sto radnja mora da zna a iz prikaza se ne vidi
    ' jednoznacno; prioritet 4, pa se ne crtaju.
    StavkeKolone = Array( _
        "OTKUI_HDB_IZVOD||txt|88|1", _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HDB_POZIV||txt|104|2", _
        "OTKUI_HDB_UPLATA||rsd|96|1", _
        "OTKUI_HDB_ISPLATA||rsd|96|1", _
        "OTKUI_HDB_STATUS||txt|72|1", _
        "OTKUI_HDB_PREDLOG||txt|160|2", _
        "OTKUI_HDB_RACUN||txt|132|3", _
        "OTKUI_HDB_BIMKEY||txt|1|4", _
        "OTKUI_HDB_OTVOREN||txt|1|4", _
        "OTKUI_HDB_SMER||txt|1|4")
End Function

' Citac vraca 1-bazirano:
'   1 BankaImportID (prazno = dvosmislen) | 2 BrojDokumenta | 3 BrojRacuna
'   4 DatumTransakcije | 5 Partner | 6 PozivNaBroj | 7 Uplata | 8 Isplata
'   9 Obradjeno | 10 Otvoren | 11 Smer | 12 JakKljuc | 13 CiljTip | 14 CiljOpis
Private Function RedoviStavke(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String, iD As String, obr As String
    Dim zbirU As Double, zbirI As Double
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    src = modBankaMapiranje.GetBankaImportForGrid()
    If Not IsArray(src) Then
        RedoviStavke = PrazanRezultat(StavkeKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 12)
    For i = 1 To UBound(src, 1)
        obr = CStr(src(i, 9))
        If Not BuCipStavka(filter, obr, CBool(src(i, 12))) Then GoTo Sledeci
        ' BankaImportID se ne PRIKAZUJE, ali ostaje u pretrazi: ko ga ima iz
        ' loga ili poruke o gresci mora moci da nadje red.
        hay = CStr(src(i, 15)) & "|" & CStr(src(i, 2)) & "|" & CStr(src(i, 3)) & "|" & _
              CStr(src(i, 5)) & "|" & CStr(src(i, 6)) & "|" & CStr(src(i, 14))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        iD = Trim$(CStr(src(i, 1)))
        outA(n, 1) = CStr(src(i, 2))
        ' DATUM IDE KAO SERIJSKI BROJ, ne kao Date. modUiData.CellDate odbija i
        ' broj koji CDate ne sme da primi (v. tamosnji komentar) -- ekran to
        ' pravilo NE ponavlja.
        outA(n, 2) = modUiData.CellDate(src, i, 4)
        outA(n, 3) = CStr(src(i, 5))
        outA(n, 4) = CStr(src(i, 6))
        outA(n, 5) = CDbl(src(i, 7))
        outA(n, 6) = CDbl(src(i, 8))
        outA(n, 7) = BuStatusTekst(obr)
        outA(n, 8) = BuPredlogTekst(CStr(src(i, 11)), CStr(src(i, 13)), CStr(src(i, 14)), _
                                    CBool(src(i, 10)))
        outA(n, 9) = CStr(src(i, 3))
        ' Identitet koji je citac PROVERIO (prazan = dvosmislen). Ne crta se.
        outA(n, 10) = iD
        outA(n, 11) = IIf(CBool(src(i, 10)), "1", "")
        outA(n, 12) = CStr(src(i, 11))
        zbirU = zbirU + CDbl(src(i, 7))
        zbirI = zbirI + CDbl(src(i, 8))
Sledeci:
    Next i

    ' Zbir kolicine nema smisla na izvodu, pa je nula.
    '
    ' Vrednost je PROMET prikazanih stavki (uplate + isplate), ne neto. Neto je
    ' na cipu "obradjeno" davao NEGATIVAN broj, koji nad izvodom ne znaci nista.
    ' Razdvojene brojke -- koliko uplata, koliko isplata -- stoje u traci iznad
    ' mreze; podnozje ljuske ima samo JEDAN slot (grdFoot.ftVal).
    RedoviStavke = Array(StavkeKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0))
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.description
    Err.Raise errNum, "modScrBankaUvoz.RedoviStavke", errDesc
End Function

' Status stavke kao tekst. Prazan status je NOV red -- prikazuje se kao crtica,
' jer prazna celija izgleda isto kao podatak koji nedostaje.
Public Function BuStatusTekst(ByVal obradjeno As String) As String
    Select Case Trim$(obradjeno)
        Case BIM_OBR_DA:    BuStatusTekst = Poruka("OTKUI_CIPB_OBRADJENO")
        Case BIM_OBR_SKIP:  BuStatusTekst = Poruka("OTKUI_CIPB_PRESKOCENO")
        Case BIM_OBR_ERROR: BuStatusTekst = Poruka("OTKUI_CIPB_RUCNO")
        Case Else:          BuStatusTekst = ChrW(8212)
    End Select
End Function

' PREDLOG PO REDU -- ono sto je legacy pokazivao samo za IZABRANU stavku, u
' viserednom preview panelu (BuildAutoPreviewText, oko 350 linija sa granama).
' Ovde je jedna celija po redu, pa se vidi za SVE redove odjednom.
'
' Racun ne radi ovaj tekst: cilj i njegovu oznaku dao je citac
' (modBankaMapiranje.BimJakiKljucInfo). Ovde je samo formulacija, i zato je
' Public -- da se moze izmeriti bez mreze.
Public Function BuPredlogTekst(ByVal smer As String, ByVal ciljTip As String, _
                               ByVal ciljOpis As String, ByVal otvoren As Boolean) As String
    ' Zatvorena stavka nema sta da predlozi.
    If Not otvoren Then Exit Function

    If smer = BIM_SMER_NEJASAN Then
        BuPredlogTekst = Poruka("OTKUI_LBL_BU_PRED_NEJASAN")
        Exit Function
    End If

    Select Case ciljTip
        Case BIM_CILJ_FAKTURA
            BuPredlogTekst = Poruka("OTKUI_LBL_BU_PRED_FAKTURA") & " " & ciljOpis
        Case BIM_CILJ_AVANS
            BuPredlogTekst = Poruka("OTKUI_LBL_BU_PRED_AVANS") & " " & ciljOpis
        Case BIM_CILJ_BLOK
            BuPredlogTekst = Poruka("OTKUI_LBL_BU_PRED_BLOK") & " " & ciljOpis
        Case Else
            BuPredlogTekst = Poruka("OTKUI_LBL_BU_PRED_NEMA")
    End Select
End Function

'------------------------------------------------------- LISTA: IZVODI
' DEVET VIDLJIVIH KOLONA, koliko ih ima i lista stavki. To nije kozmetika.
'
' Broj otvorenih i broj stavki stoje u JEDNOJ koloni -- "10 / 16", isti zapis
' koji traka iznad mreze vec koristi za "MAPIRANO 11 / 40".
'
' To je POCELO kao zaobilazak: ljuska pri promeni liste nije preracunavala
' sirine, pa je deseta kolona ove liste nasledjivala nulu od devete (skrivene)
' kolone liste STAVKE i ostajala prazna i kad joj je vrednost tacna. TAJ KVAR JE
' U MEDJUVREMENU POPRAVLJEN U LJUSCI (mGeomStara / OsveziGeometriju), pa
' zaobilazak vise nije potreban.
'
' Spojena kolona ipak OSTAJE, i to kao izbor a ne kao ostatak: dve susedne
' brojke bez konteksta citaju se gore od jedne sa kosom crtom.
Private Function IzvodiKolone() As Variant
    IzvodiKolone = Array( _
        "OTKUI_HDB_IZVOD||txt|96|1", _
        "OTKUI_HDB_RACUN||txt|0|1", _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HDB_POCETNO||rsd|104|2", _
        "OTKUI_HDB_UPLATA||rsd|104|1", _
        "OTKUI_HDB_ISPLATA||rsd|104|1", _
        "OTKUI_HDB_ZAVRSNO||rsd|104|1", _
        "OTKUI_HDB_SLAGANJE||txt|132|1", _
        "OTKUI_HDB_STAVKI||txt|118|2", _
        "OTKUI_HDB_IZVKEY||txt|1|4")
End Function

' Citac vraca 1-bazirano:
'   1 Kljuc | 2 BrojDokumenta | 3 BrojRacuna | 4 DatumIzvoda | 5 Pocetno
'   6 Potrazuje | 7 Duguje | 8 Zavrsno | 9 Razlika | 10 Status | 11 Stavki
'   12 Otvorenih
Private Function RedoviIzvodi(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String
    Dim zbirU As Double, zbirI As Double
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    src = modBankaImport.GetBankaIzvodiForGrid()
    If Not IsArray(src) Then
        RedoviIzvodi = PrazanRezultat(IzvodiKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 10)
    For i = 1 To UBound(src, 1)
        If Not BuCipIzvod(filter, CLng(src(i, 10)), CLng(src(i, 12))) Then GoTo Sledeci
        zbirU = zbirU + CDbl(src(i, 6))
        zbirI = zbirI + CDbl(src(i, 7))
        hay = CStr(src(i, 2)) & "|" & CStr(src(i, 3))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i, 2))
        outA(n, 2) = CStr(src(i, 3))
        ' v. RedoviStavke: datum se mrezi predaje kao serijski broj u opsegu.
        outA(n, 3) = modUiData.CellDate(src, i, 4)
        outA(n, 4) = CDbl(src(i, 5))
        outA(n, 5) = CDbl(src(i, 6))
        outA(n, 6) = CDbl(src(i, 7))
        outA(n, 7) = CDbl(src(i, 8))
        outA(n, 8) = BuSlaganjeTekst(CLng(src(i, 10)), CDbl(src(i, 9)))
        outA(n, 9) = BuStavkiTekst(CLng(src(i, 12)), CLng(src(i, 11)))
        outA(n, 10) = CStr(src(i, 1))
Sledeci:
    Next i

    ' Isto merilo kao na listi stavki: PROMET prikazanih izvoda (uplate +
    ' isplate). Nula bi u podnozju pisala "Vrednost 0,00 RSD", sto je tacno
    ' onoliko korisno koliko izgleda.
    RedoviIzvodi = Array(IzvodiKolone(), outA, n, 0#, zbirU + zbirI, Array(0, 0, 0))
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.description
    Err.Raise errNum, "modScrBankaUvoz.RedoviIzvodi", errDesc
End Function

' Koliko je od stavki izvoda jos otvoreno. Jedna kolona umesto dve -- v.
' IzvodiKolone. Odvojeno od mreze da bi se moglo izmeriti bez nje.
Public Function BuStavkiTekst(ByVal otvorenih As Long, ByVal stavki As Long) As String
    BuStavkiTekst = CStr(otvorenih) & " / " & CStr(stavki)
End Function

' Tekst kolone "Slaganje". Odluku je vec doneo modBankaImport.BimSaldoStatus --
' ovde je samo formulacija, i zato Public (meri se bez mreze).
Public Function BuSlaganjeTekst(ByVal status As Long, ByVal razlika As Double) As String
    Select Case status
        Case BIM_SALDO_OK
            BuSlaganjeTekst = Poruka("OTKUI_LBL_BU_SALDO_OK")
        Case BIM_SALDO_RAZLIKA
            BuSlaganjeTekst = Poruka("OTKUI_LBL_BU_SALDO_RAZLIKA") & " " & _
                              Format$(Abs(razlika), "#,##0.00")
        Case Else
            BuSlaganjeTekst = Poruka("OTKUI_LBL_BU_SALDO_NEMA")
    End Select
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod reda polja. Zona je krem, a polja su bela -- bez
    ' podloge se izmedju njih vidi pozadina zone. MORA da bude LABELA, ne Frame:
    ' Frame je prozorska kontrola i crta se IZNAD bezprozorskih bez obzira na
    ' z-order. Napravljena PRVA, labela ostaje ispod svega.
    modUiKit.NewLbl z, "buBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "buCap", UCase$(Poruka("OTKUI_SCRBU_CAP")), PAD, BU_Y_CAP, _
                    240, 11, TS_MICRO, True, C_MUTED, -1

    ' Cetiri brojke desno -- iste one koje legacy forma drzi u KPI traci
    ' (Otvoreno / Mapirano / Uplate / Isplate).
    For i = 0 To 3
        modUiKit.NewLbl z, "buKL" & i, "", 0, BU_Y_CAP, BU_KPI_W, 11, _
                        TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "buKV" & i, ChrW(8212), 0, BU_Y_KPI_V, BU_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' TRI SPREGNUTA POLJA. Prefiks "scr" je OBAVEZAN: bez njega promena teksta
    ' ide ljusci, koja o ovim poljima ne zna nista. Kombo MORA da bude polje
    ' (okvir + kontrola sa sufiksom "T"), a ne gola kontrola: panel za izbor
    ' (modOtkupUI.FindCombo) trazi bas taj oblik.
    modOtkupUI.NewFieldG z, "scrBuTip", Poruka("OTKUI_FLD_BU_TIP"), "cmb", "", _
                         1, False, False, "BU"
    modOtkupUI.NewFieldG z, "scrBuPartner", Poruka("OTKUI_FLD_BU_PARTNER"), "cmb", "", _
                         1, False, False, "BU"
    modOtkupUI.NewFieldG z, "scrBuCilj", Poruka("OTKUI_FLD_BU_CILJ"), "cmb", "", _
                         1, False, False, "BU"

    ' Mesto mu daje RasporediPolja -- stoji UZ polja, ne ispod njih.
    modUiKit.NewLbl z, "buHint", "", PAD, BU_Y_HINT, 200, 12, TS_META, False, C_MUTED, -1

    modUiKit.NewLbl z, "buLnB", "", 0, BU_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = BU_ZONA_H
End Function

Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, kx As Single, capDesno As Single
    Dim x3 As Single, hintX As Single, hintW As Single, kpiX As Single
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    ' Bela podloga pokriva CEO red polja, ukljucujuci i objasnjenje uz njih.
    z.Controls("buBg").Left = PAD - 10
    z.Controls("buBg").top = BU_Y_LBL - 8
    z.Controls("buBg").width = w - 2 * (PAD - 10)
    z.Controls("buBg").Height = BU_ZONA_H - BU_Y_LBL - 2

    PoljeX z, "scrBuTip", PAD, BU_FLD_W, BU_Y_LBL
    PoljeX z, "scrBuPartner", PAD + BU_FLD_W + BU_FLD_GAP, BU_FLD_W, BU_Y_LBL
    x3 = PAD + 2 * (BU_FLD_W + BU_FLD_GAP)
    PoljeX z, "scrBuCilj", x3, BU_FLD_W, BU_Y_LBL

    ' Polje cilja se GASI za OM: ni faktura ni blok se tada ne biraju, a polje
    ' koje ne radi nista poziva da se u njega nesto upise.
    z.Controls("scrBuCilj").Visible = (IzabraniTip() <> BIM_TIP_OM)

    ' Brojke idu uz desnu ivicu; sakriva se ona koja bi nalegla na naslov zone.
    capDesno = PAD + 180
    kpiX = w - PAD - 4 * BU_KPI_W
    For i = 0 To 3
        kx = w - PAD - (4 - i) * BU_KPI_W
        z.Controls("buKL" & i).Left = kx
        z.Controls("buKV" & i).Left = kx
        z.Controls("buKL" & i).Visible = (kx > capDesno)
        z.Controls("buKV" & i).Visible = (kx > capDesno)
    Next i

    ' OBJASNJENJE STOJI UZ POLJA, ne ispod njih: ispod je nalegalo na traku koju
    ' ljuska crta odmah po zavrsetku zone. Staje u prostor izmedju poslednjeg
    ' polja i brojki; kad tog prostora nema, sklanja se -- Label ne prelama, pa
    ' bi inace istekao preko brojki.
    ' Razmak je EKRANOV (BU_FLD_GAP). modOtkupUI.GAP je Private Const -- ovaj
    ' modul ga ne vidi, a VBA to javi tek kad se procedura prvi put izvrsi.
    hintX = x3 + BU_FLD_W + BU_FLD_GAP
    hintW = kpiX - BU_FLD_GAP - hintX
    z.Controls("buHint").Left = hintX
    z.Controls("buHint").top = BU_Y_HINT
    z.Controls("buHint").Visible = (hintW >= BU_HINT_MIN)
    If hintW >= BU_HINT_MIN Then z.Controls("buHint").width = hintW

    z.Controls("buLnB").width = w
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
    Set Zona = modOtkupUI.ScreenZone("BANKA_UVOZ")
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
    PuniTipCombo
    PuniPartnerCombo
    PuniCiljCombo
    RasporediPolja z, z.width
    OsveziObjasnjenje z
    OsveziBrojke z
End Sub

Private Sub OsveziObjasnjenjeSam()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    OsveziObjasnjenje z
End Sub

' Tip je zatvoren spisak od tri vrednosti i ne cita se ni iz jedne tabele.
Private Sub PuniTipCombo()
    Dim c As Object
    On Error Resume Next
    If mCombosPunjeni Then Exit Sub
    Set c = Kontrola("scrBuTip")
    If c Is Nothing Then Exit Sub

    mFill = True
    c.Clear
    c.ColumnCount = 1
    c.AddItem BIM_TIP_KUPAC
    c.AddItem BIM_TIP_KOOPERANT
    c.AddItem BIM_TIP_OM
    mCombosPunjeni = True
    mFill = False
End Sub

' Partner se puni PO TIPU, i puni se ponovo tek kad se tip promeni.
'
' Svuda se prikazuje i ID (ShowIDInComboDisplay): dva partnera istog naziva su
' u ovim sifarnicima obicna pojava, a izbor pogresnog salje novac pogresnom
' coveku. Isto sto frmBankaImport.LoadManualTargets radi (FM-0024 #7).
Private Sub PuniPartnerCombo()
    Dim c As Object, tip As String
    Dim mapa As Object, k As Variant
    On Error GoTo EH

    tip = IzabraniTip()
    Set c = Kontrola("scrBuPartner")
    If c Is Nothing Then Exit Sub
    If Len(tip) = 0 Then Exit Sub
    If mTipPunjen = tip Then Exit Sub

    mFill = True
    c.Clear
    c.ColumnCount = 2
    c.ColumnWidths = "180 pt;0 pt"
    c.BoundColumn = 1
    c.TextColumn = 1

    Select Case tip
        Case BIM_TIP_KUPAC
            Set mapa = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
        Case BIM_TIP_KOOPERANT
            Set mapa = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Case BIM_TIP_OM
            Set mapa = BuildLookupDict(TBL_STANICE, "StanicaID", "Naziv")
    End Select

    If Not mapa Is Nothing Then
        For Each k In mapa.keys
            c.AddItem Trim$(CStr(mapa(k)))
            c.List(c.ListCount - 1, 1) = CStr(k)
        Next k
    End If

    ShowIDInComboDisplay c

    mTipPunjen = tip
    mCiljPunjen = ""
    mFill = False
    Exit Sub
EH:
    mFill = False
    ' Prazan combo bez traga je bio glavni razlog zasto je izgledalo da "nista
    ' nije povezano" -- isto kao u modOtkupUI.FillCombos.
    Debug.Print "modScrBankaUvoz.PuniPartnerCombo PAO: " & Err.Number & " " & Err.description
End Sub

' Cilj se puni PO TIPU I PARTNERU: fakture za kupca, blokovi za kooperanta, za
' OM nista. Obe liste nose zastavicu ucitanosti -- v. mCiljOK.
Private Sub PuniCiljCombo()
    Dim c As Object, tip As String, partnerID As String, kljuc As String
    Dim src As Variant, i As Long
    On Error GoTo EH

    ' Broji se POZIV, ne uspeh -- v. mCiljPunjenja.
    mCiljPunjenja = mCiljPunjenja + 1

    tip = IzabraniTip()
    partnerID = IzabraniPartnerID()
    kljuc = tip & "|" & partnerID

    Set c = Kontrola("scrBuCilj")
    If c Is Nothing Then Exit Sub
    If mCiljPunjen = kljuc Then Exit Sub

    mFill = True
    c.Clear
    c.ColumnCount = 2
    c.ColumnWidths = "180 pt;0 pt"
    c.BoundColumn = 1
    c.TextColumn = 1

    ' Zastavica se resetuje na svako punjenje: "ucitano" vazi za TEKUCI izbor.
    mCiljOK = True
    mCiljErr = ""

    Select Case tip
        Case BIM_TIP_KUPAC
            PostaviNatpisCilja Poruka("OTKUI_FLD_BU_FAKTURA")
            src = modBankaMapiranje.GetFaktureZaBimMapiranje(partnerID, mCiljOK, mCiljErr)
            If IsArray(src) Then
                For i = 1 To UBound(src, 1)
                    c.AddItem CStr(src(i, 2)) & "  " & ChrW(183) & "  " & _
                              Format$(CDbl(src(i, 3)), "#,##0.00")
                    c.List(c.ListCount - 1, 1) = CStr(src(i, 1))
                Next i
            End If

        Case BIM_TIP_KOOPERANT
            PostaviNatpisCilja Poruka("OTKUI_FLD_BU_BLOK")
            ' TRI kolone: prikaz, broj bloka, OTKUPNO MESTO.
            '
            ' Broj otkupa je jedinstven PO STANICI, pa isti broj legitimno
            ' pripada dvama razlicitim blokovima. Kad bi combo nosio samo broj,
            ' posle izbora se ne bi znalo KOJI je -- a od toga zavisi na koji
            ' otkupni lanac ide novac. Scope zato ide u SVOJU kolonu; prikaz se
            ' NE parsira (v. IzabranaStanicaCilja).
            c.ColumnCount = 3
            c.ColumnWidths = "180 pt;0 pt;0 pt"
            src = modBankaMapiranje.GetBlokoviZaBimMapiranje(partnerID)
            If IsArray(src) Then
                For i = 1 To UBound(src, 1)
                    ' Blok bez upisanog otkupnog mesta OSTAJE u listi -- postoji
                    ' u podacima i precutati ga znacilo bi lagati o tome sta je
                    ' u tabeli. Ali se OZNACAVA, jer bi inace izgledao samo kao
                    ' "12" pored "12 . OM Naziv" i operater ne bi imao nacina da
                    ' zna zasto ga radnja odbija (v. BuScopeNedostaje).
                    If Len(Trim$(CStr(src(i, 2)))) = 0 Then
                        c.AddItem CStr(src(i, 1)) & "  " & ChrW(183) & "  " & _
                                  Poruka("OTKUI_LBL_BU_BLOK_BEZ_OM")
                    Else
                        c.AddItem CStr(src(i, 3))
                    End If
                    c.List(c.ListCount - 1, 1) = CStr(src(i, 1))
                    c.List(c.ListCount - 1, 2) = CStr(src(i, 2))
                Next i
            End If

        Case Else
            PostaviNatpisCilja Poruka("OTKUI_FLD_BU_CILJ")
    End Select

    mCiljPunjen = kljuc
    mFill = False
    Exit Sub
EH:
    mFill = False
    ' Pad punjenja NE sme da izgleda kao prazna lista. Na obe rute prazna lista
    ' znaci nesto konkretno -- avans, odnosno poziv na broj -- pa bi kvar postao
    ' drugi poslovni ishod. GetBlokoviZaBimMapiranje gresku DIZE (za razliku od
    ' faktura, koje je vracaju kroz zastavicu), pa oba puta zavrsavaju ovde.
    mCiljOK = False
    mCiljErr = "[" & CStr(Err.Number) & "] " & Err.description
    mCiljPunjen = ""
End Sub

Private Sub PostaviNatpisCilja(ByVal cap As String)
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    z.Controls("scrBuCilj").Controls("scrBuCiljL").caption = UCase$(cap)
End Sub

Private Sub OsveziObjasnjenje(ByVal z As Object)
    On Error Resume Next
    Select Case IzabraniTip()
        Case BIM_TIP_KUPAC:     z.Controls("buHint").caption = Poruka("OTKUI_LBL_BU_HINT_AVANS")
        Case BIM_TIP_KOOPERANT: z.Controls("buHint").caption = Poruka("OTKUI_LBL_BU_HINT_BLOK")
        Case BIM_TIP_OM:        z.Controls("buHint").caption = Poruka("OTKUI_LBL_BU_HINT_OM")
        Case Else:              z.Controls("buHint").caption = Poruka("OTKUI_LBL_BU_HINT")
    End Select
End Sub

Private Sub OsveziBrojke(ByVal z As Object)
    Dim k As Variant, nepoznato As Boolean, crta As String
    On Error Resume Next
    k = Kpi()
    nepoznato = BuKpiNepoznat(k)
    crta = ChrW(8212)

    z.Controls("buKL0").caption = UCase$(Poruka("OTKUI_KPI_BU_OTVORENO"))
    z.Controls("buKL1").caption = UCase$(Poruka("OTKUI_KPI_BU_MAPIRANO"))
    z.Controls("buKL2").caption = UCase$(Poruka("OTKUI_KPI_BU_UPLATE"))
    z.Controls("buKL3").caption = UCase$(Poruka("OTKUI_KPI_BU_ISPLATE"))

    ' Nula i "ne znam" nisu ista brojka. Kad citanje nije uspelo, sve cetiri
    ' plocice pokazuju crtu -- pola tacnih brojki uz dve nule bilo bi gore od
    ' iskrenog "nemam podatak".
    If nepoznato Then
        z.Controls("buKV0").caption = crta
        z.Controls("buKV1").caption = crta
        z.Controls("buKV2").caption = crta
        z.Controls("buKV3").caption = crta
        Exit Sub
    End If

    z.Controls("buKV0").caption = CStr(CLng(k(0)))
    z.Controls("buKV1").caption = CStr(CLng(k(1))) & " / " & CStr(CLng(k(2)))
    z.Controls("buKV2").caption = Format$(CDbl(k(3)), "#,##0")
    z.Controls("buKV3").caption = Format$(CDbl(k(4)), "#,##0")
End Sub

' Cetiri brojke iz JEDNOG prolaza kroz tabelu, kesirane do sledeceg upisa.
' Racun je u modBankaMapiranje.GetBankaImportKpi -- ekran ga ne ponavlja.
'
' NEUSPEH CITANJA NIJE NULA. Znacka uz stavku menija odgovara na pitanje "ima li
' finansijskih stavki koje cekaju coveka". Ako citanje pukne a mi vratimo nule,
' operater dobija "nema posla" umesto "ne znam" -- i to bas kad je nesto sa
' semom ili kesom poslo naopako. Isti fail-open je jednom vec placen u Stornu.
'
' Greska se zato LOGUJE, kes se NE proglasava vazecim (sledeci poziv pokusava
' ponovo), a vraca se POSLEDNJA POZNATA vrednost. Nula ide samo dok validne
' vrednosti jos nije ni bilo -- tada ni znacke nema, pa nema ni cega laznog.
Private Function Kpi() As Variant
    Dim errDesc As String
    On Error GoTo EH
    If mKpiOK Then
        Kpi = mKpi
        Exit Function
    End If
    mKpi = modBankaMapiranje.GetBankaImportKpi()
    mKpiOK = True
    Kpi = mKpi
    Exit Function
EH:
    ' Brojac ne sme da obori ljusku: OsveziNavBrojace pita SVAKI ekran, pa bi
    ' greska ovde ugasila i tudje znacke. Ali gusenje BEZ TRAGA je zaseban kvar.
    errDesc = Err.description
    LogErr "modScrBankaUvoz.Kpi"
    Kpi = BuKpiPosleGreske(mKpi)
    Err.Clear
End Function

' Sta brojac vraca kad citanje pukne: POSLEDNJU POZNATU vrednost, ne nule.
' Odvojeno od Kpi da bi se pravilo moglo izmeriti -- pad citanja se u testu ne
' moze izazvati bez lomljenja seme.
'
' A kad poslednje poznate vrednosti JOS NEMA (prvi pad u sesiji), brojka je
' NEPOZNATA -- i to se kaze. Ranije je tu stajala nula, uz obrazlozenje da "tada
' ni znacke nema, pa nema ni cega laznog". To je bilo naopako: u ovom UI-ju
' ODSUSTVO znacke jeste poruka, i glasi "nema sta da ceka". Prvi pad citanja bi
' tako i dalje bio fail-open, samo tise.
'
' Nepoznato se nosi kao NEGATIVAN broj: ugovor je Scr_Brojac() As Long i nema
' treci kanal, a negativan broj brojac ne moze legitimno da bude. Ljuska ga crta
' kao "!" (modOtkupUI.BrojacTekst), ekran kao crtu.
Public Function BuKpiPosleGreske(ByVal poslednja As Variant) As Variant
    If IsArray(poslednja) Then
        BuKpiPosleGreske = poslednja
    Else
        BuKpiPosleGreske = BuKpiNepoznato()
    End If
End Function

' Brojke koje znace "ne znam". Novac ostaje nula -- crta se ionako ne prikazuje
' kao iznos, a znak nepoznatog nosi prva brojka.
Public Function BuKpiNepoznato() As Variant
    BuKpiNepoznato = Array(-1, -1, -1, 0#, 0#)
End Function

' Da li skup brojki uopste nosi podatak.
Public Function BuKpiNepoznat(ByVal k As Variant) As Boolean
    If Not IsArray(k) Then
        BuKpiNepoznat = True
    Else
        BuKpiNepoznat = (CLng(k(0)) < 0)
    End If
End Function

'------------------------------------------------------- IZBORI U ZONI
' NERAZRESEN UNOS NIJE TIP. Ljuska Change salje na svaki otkucaj, pa "Koo" mora
' da vrati prazno -- ne "Kooperant".
Public Function BuTipIliPrazno(ByVal v As String) As String
    Select Case Trim$(v)
        Case BIM_TIP_KUPAC, BIM_TIP_KOOPERANT, BIM_TIP_OM
            BuTipIliPrazno = Trim$(v)
    End Select
End Function

Private Function IzabraniTip() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mTipTest) > 0 Then
            IzabraniTip = mTipTest
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrBuTip")
    If c Is Nothing Then Exit Function
    IzabraniTip = BuTipIliPrazno(CStr(c.value))
    Err.Clear
End Function

Private Function IzabraniPartnerID() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mPartnerTest) > 0 Then
            IzabraniPartnerID = mPartnerTest
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrBuPartner")
    If c Is Nothing Then Exit Function
    IzabraniPartnerID = GetComboID(c)
    Err.Clear
End Function

' Vrednost ciljnog polja kao ID (FakturaID) ili kao tekst (broj bloka).
' RUCNO IZABRAN BLOK MORA DA NOSI OTKUPNO MESTO.
'
' Tri stanja danas izgledaju isto -- prazan string -- a znace tri razlicite
' stvari:
'   1) scope nije ni trazen  (blok dolazi iz poziva na broj; legitimno bez
'      scope-a, isto kao automatsko mapiranje),
'   2) scope je trazen, ali red nema upisanu stanicu  (legacy/uvezen podatak),
'   3) scope je trazen, ali kolone nema  (schema drift -- resolver to sam
'      podigne kao gresku, v. GetOtkupCandidatesForKooperantBlock).
'
' Drugo stanje je opasno bas zato sto lici na prvo: operater je BIRAO blok, a
' writer bi dobio prazan scope i raspodelio novac preko svih otkupnih mesta sa
' tim brojem. Zato ovde stoji STOP, a ne tiho spustanje na nescope-ovan upis.
'
' Pravilo je izdvojeno da bi se moglo izmeriti bez forme. Njegovo VEZIVANJE u
' RucnoKooperant je jedan red i proveren je citanjem, ne testom.
Public Function BuScopeNedostaje(ByVal ciljID As String, ByVal stanica As String) As Boolean
    BuScopeNedostaje = (Len(Trim$(ciljID)) > 0 And Len(Trim$(stanica)) = 0)
End Function

' OTKUPNO MESTO izabranog bloka. Cita se iz TRECE kolone combo-a, ne iz prikaza:
' prikaz je za coveka i sme da se menja, a scope je podatak.
Private Function IzabranaStanicaCilja() As String
    Dim c As Object
    If IsTestMode() Then
        IzabranaStanicaCilja = mStanicaTest
        Exit Function
    End If
    On Error Resume Next
    Set c = Kontrola("scrBuCilj")
    If c Is Nothing Then Exit Function
    If c.ColumnCount < 3 Then Exit Function
    If c.ListIndex < 0 Then Exit Function
    IzabranaStanicaCilja = Trim$(CStr(c.List(c.ListIndex, 2)))
    Err.Clear
End Function

Private Function IzabraniCiljID() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mCiljTest) > 0 Then
            IzabraniCiljID = mCiljTest
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrBuCilj")
    If c Is Nothing Then Exit Function
    IzabraniCiljID = GetComboID(c)
    Err.Clear
End Function


'=====================================================================
' DIJAGNOSTIKA
'
' Alt+F8 -> Diag_BuRedovi, pa Ctrl+G (Immediate). Ne menja nista.
'
' Postoji zbog jedne klase kvara koju ni suite ni citanje koda ne razresavaju:
' celija mreze koja se vidi PRAZNA ili sa TUDJIM sadrzajem. modOtkupUI.RenderGrid
' radi pod "On Error Resume Next", pa upis koji pukne ne ostavlja trag -- natpis
' od ranijeg crtanja ostane, i to izgleda kao da ekran ispisuje tudje podatke.
' Ispisuje se ono sto ekran PREDAJE mrezi i ono sto mreza od toga DRZI.
'=====================================================================
Public Sub Diag_BuRedovi()
    Dim d As Variant, redovi As Variant, kolone As Variant, i As Long, n As Long
    On Error Resume Next

    Debug.Print "--- Diag_BuRedovi (" & SCRBU_BUILD & ") lista=" & Scr_Lista() & " ---"

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

    Dim k As Long
    If IsArray(redovi) Then
        For i = 1 To 2
            If i > n Then Exit For
            For k = 1 To UBound(kolone) + 1
                Debug.Print "  EKRAN red " & CStr(i) & " kol" & CStr(k) & ": tip=" & _
                            TypeName(redovi(i, k)) & " vred=[" & CStr(redovi(i, k)) & "]"
            Next k
        Next i
    End If

    For i = 1 To 2
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
' Zona se u testu ne crta (forma se ne prikazuje), pa se stanje ekrana ne moze
' procitati iz kontrola. Ista kapija kao Scr_*Test u modScrAgro i modScrFakture:
' seam koji MENJA stanje ekrana van test-rezima ne radi nista.
'=====================================================================
Public Sub Scr_BuListaTestSet(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mLista = kljuc
End Sub

Public Sub Scr_BuIzborTestSet(ByVal tip As String, ByVal partnerID As String, _
                              ByVal cilj As String, _
                              Optional ByVal stanicaCilja As String = "")
    If Not IsTestMode() Then Exit Sub
    mTipTest = BuTipIliPrazno(tip)
    mPartnerTest = partnerID
    mCiljTest = cilj
    mStanicaTest = stanicaCilja
End Sub

' SCOPE TRENUTNOG IZBORA, i odluka da li radnja sme da ide dalje.
'
' Kad je operater izabrao blok iz liste, zna se i sa kog je otkupnog mesta -- i
' to mora do writera, jer isti broj postoji na vise mesta. Kad blok dolazi iz
' poziva na broj, scope-a NEMA (poziv ga ne nosi) i ponasanje ostaje kao kod
' automatskog mapiranja.
'
' Racuna se na JEDNOM mestu, koje zovu i radnja i test. Dok je test ponavljao
' isti izraz, razilazenje to dvoje bi proslo neprimeceno -- sabotaza bi obarala
' kopiju u testu, a radnja bi i dalje slala prazan scope.
Private Function ScopeIzbora(ByRef stani As Boolean) As String
    ScopeIzbora = IzabranaStanicaCilja()
    If Len(IzabraniCiljID()) = 0 Then ScopeIzbora = ""
    stani = BuScopeNedostaje(IzabraniCiljID(), ScopeIzbora)
End Function

' Scope koji bi rucno mapiranje kooperanta poslalo writeru.
Public Function Scr_BuScopeBlokaTest() As String
    Dim stani As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_BuScopeBlokaTest = ScopeIzbora(stani)
End Function

' Da li bi rucno mapiranje kooperanta STALO nad trenutnim izborom.
Public Function Scr_BuStopBezOmTest() As Boolean
    Dim stani As Boolean
    Dim scope As String
    If Not IsTestMode() Then Exit Function
    scope = ScopeIzbora(stani)
    Scr_BuStopBezOmTest = stani
End Function

' Pad ucitavanja faktura se u testu ne moze izazvati bez lomljenja seme, a
' fail-closed grana je najskuplja stvar na ovom ekranu (avans umesto zatvaranja
' duga). Seam postavlja bas to stanje i vraca odluku koju RucnoKupac cita.
' Koliko je puta kapija zvala punjenje liste. Jedini nacin da se bez forme
' izmeri da poziv POSTOJI.
Public Function Scr_BuCiljPunjenoTest() As Long
    If Not IsTestMode() Then Exit Function
    Scr_BuCiljPunjenoTest = mCiljPunjenja
End Function

Public Function Scr_BuCiljStanjeTest(ByVal ucitane As Boolean) As Boolean
    If Not IsTestMode() Then Exit Function
    Dim greska As String
    mCiljOK = ucitane
    mCiljErr = "test"
    ' Ide kroz ISTU kapiju kroz koju idu i obe rucne rute. Dok je seam racunao
    ' svoj izraz, sabotaza je obarala kopiju u testu.
    Scr_BuCiljStanjeTest = CiljUcitan(greska)
End Function

Public Sub Scr_BuTestReset()
    If Not IsTestMode() Then Exit Sub
    mLista = BU_STAVKE
    mTipTest = ""
    mPartnerTest = ""
    mCiljTest = ""
    mStanicaTest = ""
    mCiljOK = True
    mCiljErr = ""
    mCiljPunjenja = 0
    Scr_ResetCache
End Sub
