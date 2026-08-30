Attribute VB_Name = "modScrIzvestaji"
'=====================================================================
' modScrIzvestaji - ekran "Izvestaji" (v6-ui-186). Faza E.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run (zamka #19).
' Red u registru (modUiScreens.ScrRows) je postojao od S3a -- stavka menija se
' do sada crtala prigusena jer modula nije bilo. Registar se NE dira.
'
' ODAKLE DOLAZI: frmIzvestaj (9 statickih tabova + 2 runtime taba) nad
' modIzvestaj Report* funkcijama. Poslovna logika je VEC izdvojena: svaki
' Report* prima (entitet, opseg) i vraca 2D niz sa dokumentovanim kolonama.
' Ekran je zato PRIKAZ NAD POSTOJECIM RACUNIMA: nijedan Report*, nijedna
' stampa i nijedno pravilo matrice se ovde ne pise ponovo niti menja.
'
' DESET LISTA deljene mreze (SALDO spaja legacy tabove 0 i 1: matrica ih
' nikad ne pokazuje istovremeno -- tab 0 samo za OM, tab 1 samo za Kupca --
' pa je za operatera to JEDAN slot ciji oblik prati entitet; jedanaesti slot
' bazena MAX_SEG ostaje slobodan). Dva runtime taba ("Otkupni listovi",
' "Pregled ambalaze") su ovde pune liste -- lista BLOKOVI na Dokumentima NIJE
' isto (blokovi JEDNE otpremnice, ne cele stanice u periodu).
'
' MATRICA JE IZVOR ISTINE: IzListaDostupna za 8 statickih lista pita
' modIzvestaj.IzvestajTabDostupan (FM-0029 #3), za 2 runtime liste preslikava
' legacy uslov iz UpdateReportMode. Matrica se NE siri i NE "popravlja":
' prazan tab za kooperanta u zbirnom rezimu je NEPOSTOJECI izvestaj (poslovna
' odluka), i tako se i prikazuje -- lista postoji uvek, kapija je na sadrzaju,
' a hint u zoni kaze ZASTO je prazna (nikad pun naslov nad trajno praznom
' listom, ali ni tiho nestajanje segmenta).
'
' KES SNIMKA (par. 22.9/N7 -- pun prolaz po otkucaju je placen kvar): sirov
' Report* povratak se kesira po KLJUCU KONTEKSTA (lista|tip|rezim|entitet|
' od|do). Pretraga i cipovi su re-filter nad snimkom (nula citanja tabela);
' promena konteksta legitimno cita ponovo. Invalidira ga Scr_ResetCache
' (ljuska ga zove posle svakog upisa). Broj stvarnih citanja meri
' mSnimakPunjenja (obrazac mCiljPunjenja / Diag_BnRedovi).
'
' SPECIJALNI REDOVI Report* povratka ("OM AVANS (nerasporedjen)",
' "AGROHEMIJA (nerasporedjena, van UKUPNO)", tri kontrolna reda isplate) NE
' idu u mrezu nego u BROJKE ZONE: to su podaci konteksta, ne redovi liste, a
' u tipiziranim kolonama mreze bi njihove prazne celije postale "0,00" --
' tacno FM-0028 #5 klasa lazi. Red UKUPNO (i POCETNO STANJE) zivi SAMO u
' nefiltriranom prikazu: pod cipom/pretragom bi tvrdio zbir koji ne odgovara
' vidljivim redovima; zbir prikazanih uvek daje podnozje mreze.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRIZ_BUILD As String = "v6-ui-186"

' Visina zone: red prekidaca + red polja + hint + red dugmadi.
Private Const IZ_ZONA_H   As Single = 148

Private Const IZ_Y_CAP    As Single = 6
Private Const IZ_Y_SEG    As Single = 20
Private Const IZ_Y_LBL    As Single = 48
Private Const IZ_Y_HINT   As Single = 98
Private Const IZ_Y_BTN    As Single = 116
Private Const IZ_BTN_H    As Single = 24
Private Const IZ_KPI_W    As Single = 150
Private Const IZ_SEG_H    As Single = 22

' Desna traka zone nosi DETALJ IZABRANOG REDA (drill-down bez detail panela
' mreze -- isti raspored kao traka korpe na Platnim nalozima): klik na red
' sa dokumentom pokazuje stavke tog dokumenta.
Private Const IZ_DET_W    As Single = 320
Private Const IZ_DET_N    As Long = 6
Private Const IZ_POLJA_MIN As Single = 470

' Kljucevi jedanaest lista.
Private Const IZ_SALDO As String = "SALDO"
Private Const IZ_ROBA As String = "ROBA"
Private Const IZ_AMB As String = "AMBALAZA"
Private Const IZ_ISPL As String = "ISPLATA"
Private Const IZ_ZBIR As String = "ZBIRNI"
Private Const IZ_CENA As String = "CENA"
Private Const IZ_MANJAK As String = "MANJAK"
Private Const IZ_KART As String = "KARTICA"
Private Const IZ_AMBK As String = "AMBKARTICA"
Private Const IZ_OTKL As String = "OTKLISTE"
Private Const IZ_RANG As String = "RANG"

' Labele specijalnih redova Report* povratka -- ISTI literali koje modIzvestaj
' upisuje (RF-06: "Labele su ASCII jer se po njima traze redovi"). Ako se tamo
' promene, test slaganja zone pukne -- to je i namera.
Private Const IZ_LBL_UKUPNO As String = "UKUPNO"
Private Const IZ_LBL_OM_AVANS As String = "OM AVANS (nerasporedjen)"
Private Const IZ_LBL_AGRO As String = "AGROHEMIJA (nerasporedjena, van UKUPNO)"
Private Const IZ_LBL_ISPL_PRIMLJENO As String = "OM Avans (primljeno)"
Private Const IZ_LBL_ISPL_PODELJENO As String = "OM Avans (podeljeno)"
Private Const IZ_LBL_ISPL_KOD As String = "Kod Otkupca"

' Granice "bez granice" opsega: Report* primaju Date, pa se prazno polje
' prevodi u pun opseg umesto u gresku.
Private Const IZ_DAT_MIN As Long = 2          ' 1.1.1900
Private Const IZ_DAT_MAX As Long = 2958465    ' 31.12.9999 (modUiData.DATUM_SERIJSKI_MAX)

'--------------------------------------------------------------- STANJE
Private mLista As String
Private mTip As String              ' "OM" / "Kupac" / "Vozac" / "Kooperant"
Private mZbirni As Boolean

Private mFill As Boolean            ' punjenje comboa okida Change
Private mComboTip As String         ' za koji tip je combo entiteta napunjen

' PODRAZUMEVAN ENTITET zivi u STANJU EKRANA, ne u tekstu comba. Ljuskin panel
' izbora filtrira stavke po TEKUCEM TEKSTU comba (PopIndex), pa bi legacy
' auto-izbor (ListIndex = 0) ostavio pun tekst u polju i panel bi zauvek
' nudio samo tu jednu stavku -- prvi smoke (28.08.2026) je to prijavio kao
' "dropdown ne radi". Combo zato ostaje PRAZAN dok operater ne izabere, a
' prikaz od prvog trenutka nosi prvi entitet tipa (legacy AutoRefresh
' ponasanje); hint ispod polja kaze KOJI je entitet stvarno prikazan.
Private mDefaultId As String

' KES SNIMAKA -- v. zaglavlje. MAPA kljuc konteksta -> sirov Report*
' povratak: setnja po listama (10 segmenata) ne placa pun prolaz ispocetka
' pri svakom povratku na vec vidjenu listu (prvi smoke: "sve je sporo").
' Kapa cuva memoriju; ResetCache prazni sve. mSnimakPunjenja broji POZIVE
' Report*-a, ne uspehe.
Private mSnimci As Object
Private mSnimakKljuc As String      ' kljuc POSLEDNJEG prikaza (kontekst stampe)
Private mSnimakPunjenja As Long
Private Const IZ_SNIMAK_KAPA As Long = 16
' Generacija podataka pod kojom je mapa punjena: upis sa DRUGOG ekrana ne
' prolazi kroz nas Scr_ResetCache, ali podize modUiData.DataGeneracija --
' kes starije generacije se odbacuje (recenzija PR #245, blocker 1).
Private mSnimakGen As Long

' Kontekst STVARNO ucitanih podataka -- naslov stampe i hint se grade iz njega
' (AUD-024 / FM-0029 #1: naslov opisuje ono sto je prikazano, ne trenutno
' stanje polja).
Private mCtxTip As String
Private mCtxZbirni As Boolean
Private mCtxId As String
Private mCtxEntNaziv As String
Private mCtxOd As Double            ' 0 = bez granice
Private mCtxDo As Double

' Brojke zone izdvojene iz Report* povratka (v. zaglavlje). -1 = nema podatka.
Private mZonaOmAvans As Variant
Private mZonaAgro As Variant
Private mZonaIsplPrimljeno As Variant
Private mZonaIsplPodeljeno As Variant
Private mZonaIsplKod As Variant
' Zavrsni saldo kartice (UKUPNO red Report* kartica: novac 7/amb 8, gajbe 6)
' -- red ne ide u mrezu, a brojka je STANJE konteksta pa zivi u zoni kao KPI
' (smoke krug 4: "kartice treba da imaju saldo i u pdf i u pregledu").
Private mZonaKartSaldo As Variant
Private mZonaKartSaldoAmb As Variant

' Hint kljuc poslednjeg citanja (postavlja RedoviZaListu, cita OsveziHint).
Private mHintKljuc As String

' Linije detalja izabranog reda (puni klik na red, prazni promena konteksta).
Private mDetalj As Variant

' Kontekst koji je postavio TEST. Zone u testu nema (forma se ne prikazuje),
' pa se combo i datumska polja ne mogu procitati. Vazi SAMO u test rezimu.
Private mTestId As String
Private mTestOd As Double
Private mTestDo As Double

' Poslednji poziv Scr_Rows -- SAMO za Diag_IzRedovi (N7 obrazac: bez ovoga se
' gubitak upita PRE ekrana i kvar POSLE ekrana ne razlikuju).
Private mDiagFilter As String
Private mDiagQ As String
Private mDiagN As Long

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=IZVESTAJI|naslov=OTKUI_NAV_IZVESTAJI|sub=OTKUI_SCRIZ_SUB" & _
               "|lista=OTKUI_SCRIZ_LISTA|oblik=zona+mreza|upis=zona"
End Function

' Tabovi lista su KONTEKSTNI (smoke krug 4): tab liste koja za izabrani tip
' entiteta ne postoji NI U JEDNOM rezimu je mrtvo dugme -- ne crta se (isti
' princip kao cipovi i radnje). Lista dostupna samo u drugom REZIMU istog
' tipa OSTAJE vidljiva: jedan klik na Pojedinacno/Zbirno je legitiman put,
' a hint objasnjava. Ljuska crta tabove iz ovog niza pri svakom rasporedu.
Public Function Scr_Liste() As Variant
    Scr_Liste = IzListeZaTip(TrenutniTip())
End Function

' Po tipu, ne po stanju ekrana -- da se ugovor moze izmeriti u testu.
Public Function IzListeZaTip(ByVal tip As String) As Variant
    Dim sve As Variant, res() As Variant, i As Long, n As Long, k As String
    sve = Array( _
        IZ_SALDO & "|OTKUI_SEG_IZ_SALDO|OTKUI_GRID_TITLE_IZ_SALDO|56", _
        IZ_ROBA & "|OTKUI_SEG_IZ_ROBA|OTKUI_GRID_TITLE_IZ_ROBA|52", _
        IZ_AMB & "|OTKUI_SEG_IZ_AMB|OTKUI_GRID_TITLE_IZ_AMB|72", _
        IZ_ISPL & "|OTKUI_SEG_IZ_ISPL|OTKUI_GRID_TITLE_IZ_ISPL|62", _
        IZ_ZBIR & "|OTKUI_SEG_IZ_ZBIR|OTKUI_GRID_TITLE_IZ_ZBIR|54", _
        IZ_CENA & "|OTKUI_SEG_IZ_CENA|OTKUI_GRID_TITLE_IZ_CENA|76", _
        IZ_MANJAK & "|OTKUI_SEG_IZ_MANJAK|OTKUI_GRID_TITLE_IZ_MANJAK|58", _
        IZ_KART & "|OTKUI_SEG_IZ_KART|OTKUI_GRID_TITLE_IZ_KART|58", _
        IZ_AMBK & "|OTKUI_SEG_IZ_AMBK|OTKUI_GRID_TITLE_IZ_AMBK|82", _
        IZ_RANG & "|OTKUI_SEG_IZ_RANG|OTKUI_GRID_TITLE_IZ_RANG|48", _
        IZ_OTKL & "|OTKUI_SEG_IZ_OTKL|OTKUI_GRID_TITLE_IZ_OTKL|82")
    ReDim res(0 To UBound(sve))
    n = 0
    For i = 0 To UBound(sve)
        k = Split(CStr(sve(i)), "|")(0)
        If IzListaZaTipPostoji(k, tip) Then
            res(n) = sve(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve res(0 To n - 1)
    IzListeZaTip = res
End Function

' Lista postoji za tip ako je matrica daje u BAR JEDNOM rezimu.
Public Function IzListaZaTipPostoji(ByVal kljuc As String, ByVal tip As String) As Boolean
    IzListaZaTipPostoji = IzListaDostupna(kljuc, tip, False) Or _
                          IzListaDostupna(kljuc, tip, True)
End Function

' Da li kombinacija (lista, rezim) trazi IZABRAN entitet. RANG nikad (nad
' svim kooperantima); zbirni rezim ne trazi -- OSIM ambalaze, ciji je
' zbirni oblik legacy agregat po tipu gajbe ZA izabranog entiteta.
Public Function IzTrebaEntitet(ByVal kljuc As String, ByVal zbirni As Boolean) As Boolean
    If kljuc = IZ_RANG Then Exit Function
    IzTrebaEntitet = (Not zbirni) Or (kljuc = IZ_AMB)
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = IZ_SALDO
    Scr_Lista = mLista
End Function

' Prvi cip je svuda "sve" -- ljuska na njega pada kad zatecen filter ne
' pripada listi (RefreshChipsForScreen). Liste bez prirodnog status-filtera
' NEMAJU cipove (prazno = ljuska ih krije): entitet, rezim i opseg su polja
' zone, a izmisljen cip bi bio novo poslovno pravilo. KARTICA ih nema ni
' zbog cega drugog: running saldo je kumulativ PUNOG skupa, pa bi filter po
' vrsti reda trajno prikazivao isecen saldo.
' Cipovi su KONTEKSTNI kao i radnje (smoke krug 3): na kombinaciji koja po
' matrici nema izvestaj lista je prazna sa objasnjenjem -- cip nad njom je
' filter necega cega nema i ne sme ni da se vidi.
Public Function Scr_Cipovi() As String
    Scr_Cipovi = IzCipoviZaKontekst(Scr_Lista(), TrenutniTip(), mZbirni)
End Function

Public Function IzCipoviZaKontekst(ByVal kljuc As String, ByVal tip As String, _
                                   ByVal zbirni As Boolean) As String
    If Not IzListaDostupna(kljuc, tip, zbirni) Then Exit Function
    IzCipoviZaKontekst = IzCipoviZaListu(kljuc)
End Function

' Cipovi PO KLJUCU LISTE -- da se ugovor moze izmeriti bez stanja ekrana.
Public Function IzCipoviZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case IZ_MANJAK
            IzCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "bezprij:OTKUI_CIPIZ_BEZPRIJ:88"
        Case IZ_AMB
            IzCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "ulaz:OTKUI_CIPIZ_ULAZ:56|" & _
                              "izlaz:OTKUI_CIPIZ_IZLAZ:56"
    End Select
End Function

' PRAVILO CIPA MANJKA: "bez prijema" propusta redove koji nose OZNAKU umesto
' brojke manjka (IZV_NEMA_PRIJEMA / IZV_VLASNIK_NEJASAN) -- ono sto se ne
' moze naplatiti. Nepoznat i prazan kljuc PUSTAJU sve.
Public Function IzCipManjak(ByVal filter As String, ByVal oznaka As String) As Boolean
    Select Case filter
        Case "bezprij": IzCipManjak = (Len(Trim$(oznaka)) > 0)
        Case Else:      IzCipManjak = True
    End Select
End Function

' PRAVILO CIPA AMBALAZE: red sa ulazom / red sa izlazom.
Public Function IzCipAmb(ByVal filter As String, ByVal ulaz As Double, _
                         ByVal izlaz As Double) As Boolean
    Select Case filter
        Case "ulaz":  IzCipAmb = (ulaz <> 0)
        Case "izlaz": IzCipAmb = (izlaz <> 0)
        Case Else:    IzCipAmb = True
    End Select
End Function

' Radnja nad redom postoji SAMO tamo gde red STVARNO ima dokument iza sebe:
' "Stampaj dokument" cita identitet iz skrivene kolone. Agregatne liste
' nemaju nijednu radnju (ljuska tada krije dugmad -- obrazac IZVODI,
' par. 9.2): agregatni red bez radnje nije greska, radnja koja pogadja jeste.
' Radnja je KONTEKSTNA, ne samo po listi (recenzija PR #245, nalaz 3): ROBA
' za kupca/vozaca je agregat po vrsti bez ref-kolone, a nedostupna
' kombinacija nema ni redove -- dugme tamo ne sme ni da se nudi.
Public Function Scr_Radnje() As String
    Scr_Radnje = IzRadnjeZaKontekst(Scr_Lista(), TrenutniTip(), mZbirni)
End Function

Public Function IzRadnjeZaKontekst(ByVal kljuc As String, ByVal tip As String, _
                                   ByVal zbirni As Boolean) As String
    If Not IzListaDostupna(kljuc, tip, zbirni) Then Exit Function
    ' ROBA nosi dokument-identitet za OM (OTP|) i kupca (PRJ|, lista
    ' prijemnica od kruga 5); vozacki oblik i ZBIRNI oblik (roba po
    ' kupcu, krug 11) su agregati bez dokumenta -- radnje nema.
    If kljuc = IZ_ROBA And (tip = "Vozac" Or zbirni) Then Exit Function
    IzRadnjeZaKontekst = IzRadnjeZaListu(kljuc)
End Function

Public Function IzRadnjeZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case IZ_OTKL, IZ_ROBA, IZ_AMB, IZ_KART
            IzRadnjeZaListu = "izprint:OTKUI_BTN_IZ_STAMPAJDOK:132:soft:1"
    End Select
End Function

' Ekran je read-only pregled: nista ovde ne ceka operatera, pa je brojac 0 i
' znacke nema -- kao lista IZVODI (par. 9.2). Ne izmislja se brojka da bi je
' bilo.
Public Function Scr_Brojac() As Long
    Scr_Brojac = 0
End Function

Public Sub Scr_ResetCache()
    ' Snimci zastarevaju na svaki upis -- sledece citanje ide u Report*.
    Set mSnimci = Nothing
    ' Sifarnici entiteta su se mogli promeniti (upis ide kroz
    ' RefreshFromData -> ResetCache) -- combo se puni ponovo, uz cuvanje
    ' izbora (v. PuniEntitetCombo).
    mComboTip = ""
End Sub

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a: LogError pocinje sa On Error Resume Next,
    ' a svaka On Error naredba brise Err.
    errDesc = Err.description
    LogErr "modScrIzvestaji.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

'=====================================================================
' MATRICA -> LISTE. Jedno javno pravilo, mereno testom nad SVIM
' kombinacijama: matrica (IzvestajTabDostupan) se NE prepisuje -- pita se.
'=====================================================================

' Indeks statickog taba mpReports za listu; -1 za runtime liste (one u legacy
' formi nemaju staticki indeks i ne prolaze kroz matricu). SALDO se razresava
' po TIPU -- legacy tabovi 0 i 1 (v. zaglavlje).
Public Function IzListaTab(ByVal kljuc As String, ByVal tip As String) As Long
    Select Case kljuc
        Case IZ_SALDO
            If tip = "Kupac" Then
                IzListaTab = IZV_TAB_SALDO_KUPCI
            Else
                IzListaTab = IZV_TAB_SALDO_OM
            End If
        Case IZ_ROBA:   IzListaTab = IZV_TAB_OTKUP_ROBA
        Case IZ_AMB:    IzListaTab = IZV_TAB_AMBALAZA
        Case IZ_ISPL:   IzListaTab = IZV_TAB_ISPLATA
        Case IZ_ZBIR:   IzListaTab = IZV_TAB_ZBIRNI
        Case IZ_CENA:   IzListaTab = IZV_TAB_PROSECNA_CENA
        Case IZ_MANJAK: IzListaTab = IZV_TAB_MANJAK
        Case IZ_KART:   IzListaTab = IZV_TAB_KARTICA
        Case Else:      IzListaTab = -1
    End Select
End Function

' Da li kombinacija (lista, tip, rezim) uopste ima izvestaj. Za 8 statickih
' lista odgovara MATRICA; za 2 runtime liste legacy uslov iz UpdateReportMode
' (Otkupni listovi samo OM-pojedinacno, Pregled ambalaze samo
' Kooperant-pojedinacno). Nedostupno NIJE prazan rezultat: lista postoji,
' hint kaze zasto je prazna.
Public Function IzListaDostupna(ByVal kljuc As String, ByVal tip As String, _
                                ByVal zbirni As Boolean) As Boolean
    Dim pg As Long
    Select Case kljuc
        Case IZ_OTKL
            IzListaDostupna = (tip = "OM" And Not zbirni)
        Case IZ_AMBK
            IzListaDostupna = (tip = "Kooperant" And Not zbirni)
        Case IZ_RANG
            ' Rang je nad SVIM kooperantima firme (legacy "Lista kooperanata"
            ' sa Unosa dokumenata, ovde uz period zone) -- izabrani entitet
            ' ga se ne tice, pa vazi u oba rezima tipa Kooperant.
            IzListaDostupna = (tip = "Kooperant")
        Case Else
            pg = IzListaTab(kljuc, tip)
            If pg < 0 Then Exit Function
            IzListaDostupna = IzvestajTabDostupan(tip, zbirni, pg)
    End Select
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

    ' Izbor reda puni DETALJ TRAKU u zoni (drill-down; smoke krug 3) -- ne
    ' menja podatke, pa se vraca False i ljuska nista ne osvezava; traka se
    ' crta direktno. Dvoklik NAMERNO ne radi nista (jedina radnja je stampa
    ' -- promasen dvoklik koji pokrene PDF je gori od nikakvog; par. 9.5).
    If Left$(tag, 4) = "row:" Then
        OsveziDetalj CLng(val(Mid$(tag, 5)))
        Exit Function
    End If
    If Left$(tag, 4) = "dbl:" Then Exit Function

    ' Promena u polju zone stize kao "chg:<kontrola>" na SVAKI otkucaj, a
    ' ljuska ne gleda povratnu vrednost (par. 8.2) -- SVE liste ovog ekrana
    ' zavise od polja zone, pa ekran sam trazi osvezavanje. Ali tek kad se
    ' kontekst STVARNO promeni: nerazresen unos nije promena.
    If Left$(tag, 4) = "chg:" Then
        Select Case Mid$(tag, 5)
            Case "scrIzEntT":              EntitetPromenjen
            Case "scrIzOdT", "scrIzDoT":   OpsegPromenjen
        End Select
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(Mid$(tag, 5))
        Exit Function
    End If

    Select Case tag
        Case "scrIzTipOM":   ObradiKlik = PostaviTip("OM")
        Case "scrIzTipKup":  ObradiKlik = PostaviTip("Kupac")
        Case "scrIzTipVoz":  ObradiKlik = PostaviTip("Vozac")
        Case "scrIzTipKoop": ObradiKlik = PostaviTip("Kooperant")
        Case "scrIzRezP":    ObradiKlik = PostaviRezim(False)
        Case "scrIzRezZ":    ObradiKlik = PostaviRezim(True)
        Case "scrIzPrint":   StampajIzvestaj
        Case "scrIzKartPdf": StampajKarticu
    End Select
End Function

' Povratna vrednost True = ljuska zove RefreshFromData, pa mreza dobija nov
' kontekst kroz redovan Scr_Rows (kljuc snimka se promenio -> novo citanje).
' Aktivna lista koje za novi tip NEMA (tab joj nestaje) prelazi na prvu
' postojecu -- highlight i naslov osvezava ljuska u istom RefreshFromData.
Private Function PostaviTip(ByVal tip As String) As Boolean
    If mTip = tip Then Exit Function
    mTip = tip
    If Not IzListaDostupna(Scr_Lista(), tip, mZbirni) Then
        mLista = PrvaListaZaKontekst(tip, mZbirni)
    End If
    PostaviTip = True
End Function

' Klik na Pojedinacno/Zbirno sa liste koje u novom rezimu nema prelazi na
' prvu dostupnu -- nikad prazan ekran sa hintom kao prvi utisak (krug 9,
' isto pravilo kao prelaz tipa iz S9).
Private Function PostaviRezim(ByVal zbirni As Boolean) As Boolean
    If mZbirni = zbirni And Len(mTip) > 0 Then Exit Function
    mZbirni = zbirni
    If Not IzListaDostupna(Scr_Lista(), TrenutniTip(), zbirni) Then
        mLista = PrvaListaZaKontekst(TrenutniTip(), zbirni)
    End If
    PostaviRezim = True
End Function

' Prva lista dostupna bas za (tip, rezim); fallback prva za tip pa SALDO.
Private Function PrvaListaZaKontekst(ByVal tip As String, _
                                     ByVal zbirni As Boolean) As String
    Dim liste As Variant, i As Long, k As String
    PrvaListaZaKontekst = IZ_SALDO
    liste = IzListeZaTip(tip)
    If Not IsArray(liste) Then Exit Function
    PrvaListaZaKontekst = Split(CStr(liste(0)), "|")(0)
    For i = LBound(liste) To UBound(liste)
        k = Split(CStr(liste(i)), "|")(0)
        If IzListaDostupna(k, tip, zbirni) Then
            PrvaListaZaKontekst = k
            Exit Function
        End If
    Next i
End Function

' Entitet iz comboa: PRAZAN ID je nerazresen unos, ne "drugi entitet"
' (par. 8.10/R2 -- prvo otkucano slovo ne sme da isprazni prikaz). Programsko
' punjenje comboa (mFill) nije unos operatera -- bez guarda bi refill usred
' Scr_Rows okinuo UGNEZDJEN RefreshFromData i dupli Report* prolaz (prvi
' smoke: "sve je sporo").
Private Sub EntitetPromenjen()
    Dim iD As String
    If mFill Then Exit Sub
    iD = SiroviEntitet()
    If Len(iD) = 0 Then Exit Sub
    If iD <> mCtxId Then modOtkupUI.RefreshFromData
End Sub

' Opseg: refresh SAMO kad se RAZRESENA granica promeni. Nepotpun datum tokom
' kucanja ("2", "21.") nije greska nego "jos nema granice" (specOdT obrazac,
' DatGranica pravilo iz modScrDokumenti) -- ne prazni listu i ne cita tabele.
' POTPUNO PRAZNO polje jeste nedvosmisleno "bez granice" (brisanje se zavrsava
' praznim), pa i ono osvezava.
Private Sub OpsegPromenjen()
    Dim odN As Double, doN As Double
    Dim odTxt As String, doTxt As String
    If mFill Then Exit Sub
    OpsegPolja odTxt, doTxt
    odN = IzDatGranica(odTxt)
    doN = IzDatGranica(doTxt)
    If odN = mCtxOd And doN = mCtxDo Then Exit Sub
    ' Nepotpun unos (tekst ima, granice nema) ne dira prikaz.
    If odN = 0 And Len(Trim$(odTxt)) > 0 And doN = mCtxDo Then Exit Sub
    If doN = 0 And Len(Trim$(doTxt)) > 0 And odN = mCtxOd Then Exit Sub
    modOtkupUI.RefreshFromData
End Sub

Private Function RadnjaNadRedom(ByVal spec As String) As Boolean
    Dim p() As String, red As Long
    p = Split(spec, ":")
    If UBound(p) < 1 Then Exit Function
    If p(0) <> "izprint" Then Exit Function
    red = CLng(val(p(1)))
    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    StampajDokumentReda red
    ' Stampa ne menja podatke -- mreza se ne osvezava.
End Function

'=====================================================================
' REDOVI MREZE
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Dim rez As Variant
    ' Combo mora biti napunjen PRE citanja konteksta (entitet se cita iz
    ' njega); ostatak zone (hint, brojke, dugmad) se osvezava POSLE citanja,
    ' jer se hint i brojke racunaju iz njega.
    PuniEntitetCombo
    rez = RedoviZaListu(filter, q)
    Scr_Rows = rez
    OsveziZonu

    ' Trag za Diag_IzRedovi -- ne menja nista.
    On Error Resume Next
    mDiagFilter = filter
    mDiagQ = q
    mDiagN = CLng(rez(2))
    Err.Clear
End Function

' Opis kolona PO KLJUCU LISTE I TIPU -- oblik kolona prati stvarni oblik
' Report* povratka (legacy tabovi ROBA i ZBIRNI menjaju kolone po tipu, SALDO
' po legacy tabu). Geometrija mreze prati opis pri svakom crtanju (par. 9.2).
Public Function IzKoloneZaListu(ByVal kljuc As String, ByVal tip As String, _
                                Optional ByVal zbirni As Boolean = False) As Variant
    ' Zbirne varijante (krug 9): SALDO/ISPLATA za OM = red po stanici;
    ' AMBALAZA = legacy agregat po tipu gajbe za izabranog entiteta.
    If zbirni Then
        Select Case kljuc
            Case IZ_SALDO
                If tip = "Kupac" Then
                    ' Kupac | Kg | Vrednost | Uplaceno | Saldo | Amb
                    IzKoloneZaListu = Array( _
                        "OTKUI_HDI_ENTITET||txt|150|1", _
                        "OTKUI_HD_KG||kg|84|1", _
                        "OTKUI_HD_VREDNOST||rsd|96|1", _
                        "OTKUI_HDI_NOVAC||rsd|92|1", _
                        "OTKUI_HDI_SALDO||rsd|96|1", _
                        "OTKUI_HDI_AMB||num|56|3", _
                        "OTKUI_HDI_REF||txt|1|4")
                    Exit Function
                End If
                ' Stanica | Kg | Vrednost | Isplaceno | Agro | Saldo | Amb
                IzKoloneZaListu = Array( _
                    "OTKUI_HD_OM||txt|140|1", _
                    "OTKUI_HD_KG||kg|82|1", _
                    "OTKUI_HD_VREDNOST||rsd|92|1", _
                    "OTKUI_HDI_ISPLACENO||rsd|92|1", _
                    "OTKUI_HDI_AGRO||rest|80|2", _
                    "OTKUI_HDI_SALDO||rsd|92|1", _
                    "OTKUI_HDI_AMB||num|56|3", _
                    "OTKUI_HDI_REF||txt|1|4")
                Exit Function
            Case IZ_ISPL
                ' Stanica | Kes | Virman firma | Virman avans | Ukupno
                IzKoloneZaListu = Array( _
                    "OTKUI_HD_OM||txt|150|1", _
                    "OTKUI_HDI_KES||rest|92|1", _
                    "OTKUI_HDI_VIRFIRMA||rest|92|1", _
                    "OTKUI_HDI_VIRAVANS||rest|92|1", _
                    "OTKUI_HDI_UKUPNO||rsd|96|1", _
                    "OTKUI_HDI_REF||txt|1|4")
                Exit Function
            Case IZ_ROBA
                ' Kupac | Kg | Vrednost (roba po kupcu, krug 11).
                IzKoloneZaListu = Array( _
                    "OTKUI_HDI_ENTITET||txt|0|1", _
                    "OTKUI_HD_KG||kg|92|1", _
                    "OTKUI_HD_VREDNOST||rsd|100|1", _
                    "OTKUI_HDI_REF||txt|1|4")
                Exit Function
            Case IZ_AMB
                ' Tip | Ulaz | Izlaz (legacy zbirni oblik).
                IzKoloneZaListu = Array( _
                    "OTKUI_HDA_TIP||txt|0|1", _
                    "OTKUI_HDA_ULAZ||txt|80|1", _
                    "OTKUI_HDA_IZLAZ||txt|80|1")
                Exit Function
        End Select
    End If
    Select Case kljuc
        Case IZ_SALDO
            If tip = "Kupac" Then
                ' Vrsta | Kolicina | Cena | Vrednost | Novac | Saldo | Ambalaza
                IzKoloneZaListu = Array( _
                    "OTKUI_HDI_VRSTA||txt|96|1", _
                    "OTKUI_HD_KG||kg|86|1", _
                    "OTKUI_HD_CENA||rest|76|2", _
                    "OTKUI_HD_VREDNOST||rsd|96|1", _
                    "OTKUI_HDI_NOVAC||rsd|96|1", _
                    "OTKUI_HDI_SALDO||rsd|96|1", _
                    "OTKUI_HDI_AMB||num|64|2")
            Else
                ' Kooperant | Kolicina | Vrednost | Isplaceno | Agro | Saldo | Amb
                IzKoloneZaListu = Array( _
                    "OTKUI_HDA_KOOPERANT||txt|130|1", _
                    "OTKUI_HD_KG||kg|82|1", _
                    "OTKUI_HD_VREDNOST||rsd|92|1", _
                    "OTKUI_HDI_ISPLACENO||rsd|92|1", _
                    "OTKUI_HDI_AGRO||rest|82|2", _
                    "OTKUI_HDI_SALDO||rsd|92|1", _
                    "OTKUI_HDI_AMB||num|56|3")
            End If
        Case IZ_ROBA
            If tip = "OM" Then
                ' Datum | BrOtp | Vrsta | Klasa | Vozac | Otp kg | Blokovi kg |
                ' Razlika | Prijemnica | Manjak kg | Manjak % | [OTP|id]
                ' Legacy je Manjak kg i % SPAJAO u jednu kolonu (ListBox limit
                ' 10) -- mreza ima MAX_COLS 14, pa su razdvojene. Prijemnica/
                ' manjak kolone su txt: prazno kad nema prijema je PORUKA, a
                ' u tipiziranoj koloni bi postalo "0,00" (FM-0028 #5).
                IzKoloneZaListu = Array( _
                    "OTKUI_HD_DATUM||date|64|1", _
                    "OTKUI_HDI_BROTP||txt|82|1", _
                    "OTKUI_HDI_VRSTA||txt|72|2", _
                    "OTKUI_HDI_KLASA||txt|40|2", _
                    "OTKUI_HDI_VOZAC||txt|100|3", _
                    "OTKUI_HDI_OTPKG||kg|72|1", _
                    "OTKUI_HDI_BLOKKG||kg|72|1", _
                    "OTKUI_HDI_RAZLIKA||kg|66|2", _
                    "OTKUI_HDI_PRIJKG||txt|76|1", _
                    "OTKUI_HDI_MANJKG||txt|70|1", _
                    "OTKUI_HDI_MANJPCT||txt|84|1", _
                    "OTKUI_HDI_REF||txt|1|4")
            ElseIf tip = "Kupac" Then
                ' LISTA PRIJEMNICA (smoke krug 4) -- ne agregat po vrsti:
                ' Datum | Broj | Zbirna | Vrsta | Kl. | Kg | Cena | Vrednost |
                ' [PRJ|id]
                IzKoloneZaListu = Array( _
                    "OTKUI_HD_DATUM||date|64|1", _
                    "OTKUI_HDI_BRDOK||txt|84|1", _
                    "OTKUI_HDI_BRZBIRNE||txt|84|2", _
                    "OTKUI_HDI_VRSTA||txt|72|2", _
                    "OTKUI_HDI_KLASA||txt|40|3", _
                    "OTKUI_HD_KG||kg|84|1", _
                    "OTKUI_HD_CENA||rest|70|2", _
                    "OTKUI_HD_VREDNOST||rsd|96|1", _
                    "OTKUI_HDI_REF||txt|1|4")
            Else
                ' Nr | Vrsta | Kolicina | Vrednost (agregat, bez identiteta)
                IzKoloneZaListu = Array( _
                    "OTKUI_HDI_NR||txt|40|2", _
                    "OTKUI_HDI_VRSTA||txt|110|1", _
                    "OTKUI_HD_KG||kg|92|1", _
                    "OTKUI_HD_VREDNOST||rsd|100|1")
            End If
        Case IZ_AMB
            ' Datum | Mesto | Tip | Dokument | Ulaz | Izlaz | [DokTip] | [DokID]
            ' Ulaz/Izlaz su txt (ekran formatira): nula se prikazuje PRAZNO,
            ' kao u legacy pregledu.
            IzKoloneZaListu = Array( _
                "OTKUI_HD_DATUM||date|64|1", _
                "OTKUI_HDI_MESTO||txt|110|1", _
                "OTKUI_HDA_TIP||txt|72|1", _
                "OTKUI_HDI_DOKUMENT||txt|96|1", _
                "OTKUI_HDA_ULAZ||txt|60|1", _
                "OTKUI_HDA_IZLAZ||txt|60|1", _
                "OTKUI_HDI_DOKTIP||txt|1|4", _
                "OTKUI_HDI_DOKID||txt|1|4")
        Case IZ_ISPL
            ' Kooperant | Kes otkupac | Virman firma | Virman avans | Ukupno
            ' Kanali su "rest": nula = prazno (isplata tim kanalom ne postoji).
            IzKoloneZaListu = Array( _
                "OTKUI_HDA_KOOPERANT||txt|140|1", _
                "OTKUI_HDI_KES||rest|92|1", _
                "OTKUI_HDI_VIRFIRMA||rest|92|1", _
                "OTKUI_HDI_VIRAVANS||rest|92|1", _
                "OTKUI_HDI_UKUPNO||rsd|96|1")
        Case IZ_ZBIR
            If tip = "Vozac" Then
                ' Vozac | Amb izlaz | Amb vracena | Manjak kg | Manjak %
                IzKoloneZaListu = Array( _
                    "OTKUI_HDI_VOZAC||txt|130|1", _
                    "OTKUI_HDI_AMBIZLAZ||num|76|1", _
                    "OTKUI_HDI_AMBVRAC||num|76|1", _
                    "OTKUI_HDI_MANJKG||kg|80|1", _
                    "OTKUI_HDI_MANJPCT||txt|70|1")
            Else
                ' Entitet | Vrsta | Kolicina | Vrednost | Prosek
                IzKoloneZaListu = Array( _
                    "OTKUI_HDI_ENTITET||txt|130|1", _
                    "OTKUI_HDI_VRSTA||txt|84|1", _
                    "OTKUI_HD_KG||kg|88|1", _
                    "OTKUI_HD_VREDNOST||rsd|98|1", _
                    "OTKUI_HDI_PROSEK||rest|76|2")
            End If
        Case IZ_CENA
            IzKoloneZaListu = Array( _
                "OTKUI_HDI_VRSTA||txt|110|1", _
                "OTKUI_HD_KG||kg|92|1", _
                "OTKUI_HD_VREDNOST||rsd|100|1", _
                "OTKUI_HDI_PROSCENA||rest|86|1")
        Case IZ_MANJAK
            ' BrZbirne | Zbirna kg | Prijemnica | Manjak kg | Manjak % | Prosek
            IzKoloneZaListu = Array( _
                "OTKUI_HDI_BRZBIRNE||txt|84|1", _
                "OTKUI_HDI_ZBIRKG||kg|80|1", _
                "OTKUI_HDI_PRIJKG||txt|80|1", _
                "OTKUI_HDI_MANJKG||txt|72|1", _
                "OTKUI_HDI_MANJPCT||txt|92|1", _
                "OTKUI_HDI_PROSEKGAJBE||rest|72|3")
        Case IZ_KART
            ' Datum | Broj dok. | Opis | Zaduzenje | Razduzenje | Saldo |
            ' Saldo amb. | [ref]. Zad/Razd su "rest": POCETNO STANJE i novcani
            ' redovi imaju praznu polovinu, a prazno je istina (ne "0,00").
            IzKoloneZaListu = Array( _
                "OTKUI_HD_DATUM||date|64|1", _
                "OTKUI_HDI_BRDOK||txt|76|1", _
                "OTKUI_HDI_OPIS||txt|170|1", _
                "OTKUI_HDI_ZAD||rest|88|1", _
                "OTKUI_HDI_RAZD||rest|88|1", _
                "OTKUI_HDI_SALDO||rsd|92|1", _
                "OTKUI_HDI_SALDOAMB||num|62|2", _
                "OTKUI_HDI_REF||txt|1|4")
        Case IZ_AMBK
            ' Datum | Broj dok. | Opis | Ulaz | Izlaz | Saldo (gajbe)
            IzKoloneZaListu = Array( _
                "OTKUI_HD_DATUM||date|64|1", _
                "OTKUI_HDI_BRDOK||txt|84|1", _
                "OTKUI_HDI_OPIS||txt|160|1", _
                "OTKUI_HDA_ULAZ||txt|58|1", _
                "OTKUI_HDA_IZLAZ||txt|58|1", _
                "OTKUI_HDI_SALDO||num|70|1")
        Case IZ_OTKL
            ' Datum | Broj dok. | Kooperant | Vrsta | Klasa | Kolicina |
            ' Vrednost | [OTK|id]
            IzKoloneZaListu = Array( _
                "OTKUI_HD_DATUM||date|64|1", _
                "OTKUI_HDI_BRDOK||txt|82|1", _
                "OTKUI_HDA_KOOPERANT||txt|130|1", _
                "OTKUI_HDI_VRSTA||txt|72|2", _
                "OTKUI_HDI_KLASA||txt|40|2", _
                "OTKUI_HD_KG||kg|84|1", _
                "OTKUI_HD_VREDNOST||rsd|96|1", _
                "OTKUI_HDI_REF||txt|1|4")
        Case IZ_RANG
            ' Rang | Kooperant | Otkupno mesto | Iznos | [KOP|id] -- isti
            ' kljucevi kao lista KOOPERANTI na Dokumentima (deljeni katalog).
            IzKoloneZaListu = Array( _
                "OTKUI_HDK_RANG||num|54|2", _
                "OTKUI_HDK_KOOPERANT||txt|0|1", _
                "OTKUI_HD_OM||txt|150|2", _
                "OTKUI_HDK_IZNOS||rsd|130|1", _
                "OTKUI_HDI_REF||txt|1|4")
    End Select
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

' Jedan poziv = jedan kontekst: dostupnost (matrica), identitet entiteta,
' opseg -> kljuc snimka -> sirovi podaci -> oblikovani redovi pod (filter, q).
Private Function RedoviZaListu(ByVal filter As String, ByVal q As String) As Variant
    Dim kljuc As String, tip As String, zbirni As Boolean, iD As String
    Dim odN As Double, doN As Double
    Dim kolone As Variant
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    kljuc = Scr_Lista()
    tip = TrenutniTip()
    zbirni = mZbirni

    ' Novi prikaz = nova selekcija; detalj prethodnog reda ne sme da ostane
    ' (isti razlog kao KarticaDetalji_Clear na mpReports_Change u legacy).
    OcistiDetalj

    kolone = IzKoloneZaListu(kljuc, tip, zbirni)

    ' 1) Kombinacija bez izvestaja (matrica) -- prazna lista, hint kaze zasto.
    If Not IzListaDostupna(kljuc, tip, zbirni) Then
        mHintKljuc = "OTKUI_IZ_HINT_NEDOSTUPNO"
        ResetZonskeBrojke
        RedoviZaListu = PrazanRezultat(kolone)
        Exit Function
    End If

    ' 2) Lista koja trazi entitet bez izabranog -- legacy guard ("Izaberite
    '    entitet"), samo kao prazna lista + hint umesto MsgBox-a. RANG je
    '    izuzet (nad SVIM kooperantima); AMBALAZA trazi entitet I U ZBIRNOM
    '    rezimu (legacy zbirni oblik = agregat po tipu ZA izabranog).
    iD = ""
    If IzTrebaEntitet(kljuc, zbirni) Then
        iD = IzabraniEntitet()
        If Len(iD) = 0 Then
            mHintKljuc = "OTKUI_IZ_HINT_IZABERI"
            ResetZonskeBrojke
            RedoviZaListu = PrazanRezultat(kolone)
            Exit Function
        End If
    End If

    OpsegGranice odN, doN

    ' 3) Snimak po kljucu konteksta; pretraga i cip NISU u kljucu.
    Dim k As String
    k = kljuc & "|" & tip & "|" & IIf(zbirni, "Z", "P") & "|" & iD & "|" & _
        CStr(odN) & "|" & CStr(doN)
    Dim src As Variant
    src = Snimak(k, kljuc, tip, zbirni, iD, odN, doN)

    mHintKljuc = ""
    RedoviZaListu = Oblikuj(kljuc, tip, zbirni, src, kolone, filter, q)
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.description
    Err.Raise errNum, "modScrIzvestaji.RedoviZaListu", errDesc
End Function

' Snimak konteksta: iz Report*-a SAMO kad kljuc jos nije u mapi (ili je kes
' ispraznjen posle upisa); inace iz mape. Pretraga i cipovi su re-filter nad
' snimkom (N7), a povratak na vec vidjenu listu je trenutan. Greska citanja
' se NE kesira (Err prekida pre upisa u mapu).
Private Function Snimak(ByVal k As String, ByVal kljuc As String, ByVal tip As String, _
                        ByVal zbirni As Boolean, ByVal iD As String, _
                        ByVal odN As Double, ByVal doN As Double) As Variant
    ' Upis sa drugog ekrana ne zove nas Scr_ResetCache -- generacija podataka
    ' je deljeni signal da je snimljeno stanje mozda staro (blocker 1).
    If mSnimakGen <> modUiData.DataGeneracija() Then
        Set mSnimci = Nothing
        mSnimakGen = modUiData.DataGeneracija()
    End If
    If mSnimci Is Nothing Then Set mSnimci = CreateObject("Scripting.Dictionary")

    If Not mSnimci.Exists(k) Then
        ' Kapa drzi memoriju: preko granice se krece ispocetka (najprostije
        ' ispravno; ResetCache ionako prazni sve posle svakog upisa).
        If mSnimci.count >= IZ_SNIMAK_KAPA Then mSnimci.RemoveAll
        mSnimakPunjenja = mSnimakPunjenja + 1
        mSnimci(k) = PuniSnimak(kljuc, tip, zbirni, iD, odN, doN)
    End If

    ' Kontekst STVARNO prikazanih podataka -- za naslov stampe i hint.
    mSnimakKljuc = k
    mCtxTip = tip
    mCtxZbirni = zbirni
    mCtxId = iD
    mCtxOd = odN
    mCtxDo = doN
    ' Lista koja i u zbirnom rezimu trazi entitet (zbirna AMBALAZA) nosi
    ' IME tog entiteta -- "Svi" bi lagao da je prikaz preko svih (smoke
    ' krug 9: podaci prve stanice pod naslovom "OM: Svi").
    mCtxEntNaziv = EntitetNaziv(tip, iD, zbirni And Not IzTrebaEntitet(kljuc, zbirni))

    Snimak = mSnimci(k)
End Function

' Jedan Report* poziv po listi -- ekran ne cita tabele sam. Prazna granica
' postaje pun opseg (Report* primaju Date).
Private Function PuniSnimak(ByVal kljuc As String, ByVal tip As String, _
                            ByVal zbirni As Boolean, ByVal iD As String, _
                            ByVal odN As Double, ByVal doN As Double) As Variant
    Dim dOd As Date, dDo As Date
    dOd = CDate(IIf(odN > 0, odN, IZ_DAT_MIN))
    dDo = CDate(IIf(doN > 0, doN, IZ_DAT_MAX))

    Select Case kljuc
        Case IZ_SALDO
            If zbirni Then
                If tip = "Kupac" Then
                    PuniSnimak = ReportSaldoKupciZbirni(dOd, dDo)
                Else
                    PuniSnimak = ReportSaldoOMZbirni(dOd, dDo)
                End If
            ElseIf tip = "Kupac" Then
                PuniSnimak = ReportSaldoKupci(iD, dOd, dDo)
            Else
                PuniSnimak = ReportSaldoOM(iD, dOd, dDo)
            End If
        Case IZ_ROBA
            ' Kupac gleda dokumenta (prijemnice), ne agregat po vrsti --
            ' agregat vec daje tab Zbirni (smoke krug 4). Zbirno (krug 11):
            ' roba PO KUPCU -- UKUPNO red kupcevog agregata.
            If zbirni And tip = "Kupac" Then
                PuniSnimak = ReportRobaKupciZbirni(dOd, dDo)
            ElseIf tip = "Kupac" Then
                PuniSnimak = ReportPrijemniceKupca(iD, dOd, dDo)
            Else
                PuniSnimak = ReportOtkupRoba(tip, iD, dOd, dDo)
            End If
        Case IZ_AMB:    PuniSnimak = ReportAmbalaza(tip, iD, dOd, dDo, zbirni)
        Case IZ_ISPL
            If zbirni Then
                PuniSnimak = ReportIsplataZbirniOM(dOd, dDo)
            Else
                PuniSnimak = ReportIsplata(tip, iD, dOd, dDo)
            End If
        Case IZ_ZBIR:   PuniSnimak = ReportZbirni(tip, dOd, dDo)
        Case IZ_CENA:   PuniSnimak = ReportProsecnaCena(tip, iD, dOd, dDo)
        Case IZ_MANJAK: PuniSnimak = ReportManjak(tip, iD, dOd, dDo)
        Case IZ_KART:   PuniSnimak = ReportKarticaKooperanta(iD, dOd, dDo)
        Case IZ_AMBK:   PuniSnimak = ReportKarticaAmbalaze(iD, dOd, dDo)
        Case IZ_OTKL:   PuniSnimak = ReportOtkupListe(iD, dOd, dDo)
        Case IZ_RANG
            ' Isti racun kao "Kooperanti po iznosu otkupa" na Unosu
            ' dokumenata (modOtkupBlok.KoopRangRows) -- ovde sa periodom
            ' zone umesto fiksne tekuce godine. Kontrolne sume racunu ne
            ' trebaju za prikaz.
            ' Granice idu kao PUN opseg (dOd/dDo), nikad 0: 0/0 bi u racunu
            ' znacilo legacy "tekuca godina", a prazna polja zone svuda na
            ' ekranu znace "sve".
            Dim rKg As Double, rVal As Double, eKg As Double, eVal As Double
            PuniSnimak = modOtkupBlok.KoopRangRows(rKg, rVal, eKg, eVal, _
                                                   CDbl(dOd), CDbl(dDo))
    End Select
End Function

'=====================================================================
' OBLIKOVANJE: sirov Report* povratak -> redovi mreze pod (filter, q).
'
' Pravila deljena svim listama:
'  - specijalni redovi (OM AVANS / AGRO / kontrolni redovi isplate) idu u
'    brojke zone, ne u mrezu;
'  - UKUPNO i POCETNO STANJE zive samo u NEFILTRIRANOM prikazu;
'  - haystack pretrage ide kroz modUiData.TekstZaPretragu (kvake u podacima,
'    DE tastatura kod operatera -- N3);
'  - identitet i sve sto radnja mora da zna a prikaz ne kaze jednoznacno ide
'    u red, prio 4 (GridCell ga cita, celija se nikad ne crta).
'=====================================================================
Private Function Oblikuj(ByVal kljuc As String, ByVal tip As String, _
                         ByVal zbirni As Boolean, _
                         ByVal src As Variant, ByVal kolone As Variant, _
                         ByVal filter As String, ByVal q As String) As Variant
    Dim nSrc As Long, nK As Long, i As Long, n As Long
    Dim outA() As Variant
    Dim qN As String, hay As String
    Dim filtrira As Boolean
    Dim sumKg As Double, sumVal As Double
    Dim vrsta As Long   ' 0 obican, 1 UKUPNO, 2 POCETNO, 3 za zonu

    ResetZonskeBrojke

    nK = UBound(kolone) + 1
    If IsEmpty(src) Or Not IsArray(src) Then
        Oblikuj = PrazanRezultat(kolone)
        Exit Function
    End If
    nSrc = UBound(src, 1)
    If nSrc = 0 Then
        Oblikuj = PrazanRezultat(kolone)
        Exit Function
    End If

    qN = modUiData.TekstZaPretragu(q)
    filtrira = (Len(qN) > 0) Or (Len(filter) > 0 And filter <> "sve")

    ReDim outA(1 To nSrc, 1 To nK)
    For i = 1 To nSrc
        vrsta = VrstaReda(kljuc, tip, zbirni, src, i)
        If vrsta = 3 Then GoTo Sledeci             ' izdvojen u zonu (VrstaReda)
        ' UKUPNO red NIKAD ne ide u mrezu: mreza sortira po koloni, pa bi
        ' legacy poslednji red PLUTAO usred liste (prvi smoke, lista Isplata).
        ' Zbir prikazanih daje podnozje; stampa dobija svoj izracunat UKUPNO.
        If vrsta = 1 Then GoTo Sledeci
        ' POCETNO STANJE je red konteksta -- pod filterom bi lagao.
        If filtrira And vrsta = 2 Then GoTo Sledeci

        If Not CipPropusta(kljuc, filter, src, i) Then GoTo Sledeci

        If Len(qN) > 0 Then
            hay = modUiData.TekstZaPretragu(HaystackReda(kljuc, tip, zbirni, src, i))
            If InStr(1, hay, qN, vbTextCompare) = 0 Then GoTo Sledeci
        End If

        n = n + 1
        UpisiRed kljuc, tip, zbirni, src, i, outA, n, (vrsta = 1), sumKg, sumVal
Sledeci:
    Next i

    Oblikuj = Array(kolone, outA, n, sumKg, sumVal, Array(0, 0, 0))
End Function

' Klasifikacija reda izvora. Vrsta 3 USPUT puni brojke zone (OM avans, agro,
' kontrolni redovi isplate) -- jedno mesto, da se izdvajanje i prikaz ne
' mogu razici.
Private Function VrstaReda(ByVal kljuc As String, ByVal tip As String, _
                           ByVal zbirni As Boolean, _
                           ByRef src As Variant, ByVal i As Long) As Long
    Dim lbl As String
    ' Zbirni oblici SALDO/ISPL nose UKUPNO u koloni 2 (kolona 1 je ID
    ' stanice); AMB zbirni deli granu pojedinacnog (UKUPNO u koloni 1).
    If zbirni Then
        Select Case kljuc
            Case IZ_SALDO, IZ_ISPL, IZ_ROBA
                If CStr(src(i, 2)) = IZ_LBL_UKUPNO Then VrstaReda = 1
                Exit Function
        End Select
    End If
    Select Case kljuc
        Case IZ_SALDO
            If tip = "Kupac" Then
                If CStr(src(i, 1)) = IZ_LBL_UKUPNO Then VrstaReda = 1
            Else
                lbl = CStr(src(i, 1))
                If lbl = IZ_LBL_UKUPNO Then
                    VrstaReda = 1
                ElseIf lbl = IZ_LBL_OM_AVANS Then
                    mZonaOmAvans = NzD(src(i, 4))
                    VrstaReda = 3
                ElseIf lbl = IZ_LBL_AGRO Then
                    mZonaAgro = NzD(src(i, 5))
                    VrstaReda = 3
                End If
            End If
        Case IZ_ROBA
            If CStr(src(i, 2)) = IZ_LBL_UKUPNO Then VrstaReda = 1
        Case IZ_AMB
            If CStr(src(i, 1)) = IZ_LBL_UKUPNO Then VrstaReda = 1
        Case IZ_ISPL
            lbl = CStr(src(i, 1))
            Select Case lbl
                Case IZ_LBL_UKUPNO:          VrstaReda = 1
                Case IZ_LBL_ISPL_PRIMLJENO:  mZonaIsplPrimljeno = NzD(src(i, 5)): VrstaReda = 3
                Case IZ_LBL_ISPL_PODELJENO:  mZonaIsplPodeljeno = NzD(src(i, 5)): VrstaReda = 3
                Case IZ_LBL_ISPL_KOD:        mZonaIsplKod = NzD(src(i, 5)): VrstaReda = 3
            End Select
        Case IZ_ZBIR
            If tip = "Vozac" Then
                If CStr(src(i, 1)) = IZ_LBL_UKUPNO Then VrstaReda = 1
            Else
                If CStr(src(i, 2)) = IZ_LBL_UKUPNO Then VrstaReda = 1
            End If
        Case IZ_MANJAK
            If CStr(src(i, 1)) = IZ_LBL_UKUPNO Then VrstaReda = 1
        Case IZ_KART
            lbl = CStr(src(i, 4))
            If lbl = IZ_LBL_UKUPNO Then
                ' UKUPNO red kartice nosi ZAVRSNI saldo (kol. 7) i zavrsni
                ' amb saldo (kol. 8) -- red ne ide u mrezu, brojke idu u zonu.
                mZonaKartSaldo = NzD(src(i, 7))
                mZonaKartSaldoAmb = NzD(src(i, 8))
                VrstaReda = 1
            End If
            If lbl = IZV_POCETNO_STANJE Then VrstaReda = 2
        Case IZ_AMBK
            lbl = CStr(src(i, 3))
            If lbl = IZ_LBL_UKUPNO Then
                mZonaKartSaldo = NzD(src(i, 6))     ' zavrsni saldo gajbi
                VrstaReda = 1
            End If
            If lbl = IZV_POCETNO_STANJE Then VrstaReda = 2
    End Select
End Function

Private Function CipPropusta(ByVal kljuc As String, ByVal filter As String, _
                             ByRef src As Variant, ByVal i As Long) As Boolean
    Select Case kljuc
        Case IZ_MANJAK
            ' Oznaka ("nema prijema" / "nejasan vlasnik") stize kao TEKST u
            ' koloni 5; broj znaci da prijem postoji.
            CipPropusta = IzCipManjak(filter, IIf(IsNumeric(src(i, 5)), "", _
                                      Trim$(CStr(NzS(src(i, 5))))))
        Case IZ_AMB
            CipPropusta = IzCipAmb(filter, NzD(src(i, 5)), NzD(src(i, 6)))
        Case Else
            CipPropusta = True
    End Select
End Function

' Tekst po kom se red trazi -- vidljive tekstualne kolone (brojevi dokumenata,
' imena, vrsta, opis).
Private Function HaystackReda(ByVal kljuc As String, ByVal tip As String, _
                              ByVal zbirni As Boolean, _
                              ByRef src As Variant, ByVal i As Long) As String
    If zbirni Then
        Select Case kljuc
            Case IZ_SALDO, IZ_ISPL, IZ_ROBA
                ' naziv entiteta (kolona 2; kolona 1 je ID)
                HaystackReda = NzS(src(i, 2))
                Exit Function
        End Select
    End If
    Select Case kljuc
        Case IZ_SALDO, IZ_ISPL, IZ_CENA
            HaystackReda = NzS(src(i, 1))
        Case IZ_ROBA
            If tip = "OM" Then
                HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3)) & "|" & _
                               NzS(src(i, 4)) & "|" & NzS(src(i, 5))
            ElseIf tip = "Kupac" Then
                ' broj | zbirna | vrsta prijemnice
                HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3)) & "|" & _
                               NzS(src(i, 4))
            Else
                HaystackReda = NzS(src(i, 2))
            End If
        Case IZ_AMB
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3)) & "|" & NzS(src(i, 4))
        Case IZ_ZBIR
            HaystackReda = NzS(src(i, 1)) & "|" & NzS(src(i, 2))
        Case IZ_MANJAK
            HaystackReda = NzS(src(i, 1))
        Case IZ_KART
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3)) & "|" & NzS(src(i, 4))
        Case IZ_AMBK
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3))
        Case IZ_RANG
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3))
        Case IZ_OTKL
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 3)) & "|" & _
                           NzS(src(i, 4)) & "|" & NzS(src(i, 5))
    End Select
End Function

' Upis jednog reda izvora u red mreze + zbirovi podnozja (POD ISTIM filterima
' kao redovi -- par. 13; UKUPNO red se NE broji u podnozje).
Private Sub UpisiRed(ByVal kljuc As String, ByVal tip As String, _
                     ByVal zbirni As Boolean, _
                     ByRef src As Variant, ByVal i As Long, _
                     ByRef outA() As Variant, ByVal n As Long, _
                     ByVal jeUkupno As Boolean, _
                     ByRef sumKg As Double, ByRef sumVal As Double)
    Dim ref As String, p() As String
    ' Zbirni oblici (krug 9): red po stanici -- naziv + brojevi + OM|
    ' identitet iz kolone 1 snimka; AMB zbirni = Tip | Ulaz | Izlaz.
    If zbirni And (kljuc = IZ_SALDO Or kljuc = IZ_ISPL Or kljuc = IZ_ROBA) Then
        Dim zc As Long, nSrcK As Long
        nSrcK = UBound(src, 2)
        outA(n, 1) = NzS(src(i, 2))
        For zc = 3 To nSrcK
            outA(n, zc - 1) = NzD(src(i, zc))
        Next zc
        outA(n, nSrcK) = IIf(Len(NzS(src(i, 1))) > 0, _
                             IIf(tip = "Kupac", "KUP|", "OM|") & NzS(src(i, 1)), "")
        If Not jeUkupno Then
            If kljuc = IZ_SALDO Then
                sumKg = sumKg + NzD(src(i, 3))
                ' saldo je pretposlednja brojcana kolona (OM: 7, kupci: 6)
                sumVal = sumVal + NzD(src(i, nSrcK - 1))
            ElseIf kljuc = IZ_ROBA Then
                sumKg = sumKg + NzD(src(i, 3))
                sumVal = sumVal + NzD(src(i, 4))
            Else
                sumVal = sumVal + NzD(src(i, 6))
            End If
        End If
        Exit Sub
    End If
    If zbirni And kljuc = IZ_AMB Then
        outA(n, 1) = NzS(src(i, 1))
        outA(n, 2) = GajbeIliPrazno(src(i, 5))
        outA(n, 3) = GajbeIliPrazno(src(i, 6))
        Exit Sub
    End If
    Select Case kljuc
        Case IZ_SALDO
            outA(n, 1) = NzS(src(i, 1))
            If tip = "Kupac" Then
                outA(n, 2) = NzD(src(i, 2))
                outA(n, 3) = NzD(src(i, 3))
                outA(n, 4) = NzD(src(i, 4))
                outA(n, 5) = NzD(src(i, 5))
                outA(n, 6) = NzD(src(i, 6))
                outA(n, 7) = NzD(src(i, 7))
            Else
                outA(n, 2) = NzD(src(i, 2))
                outA(n, 3) = NzD(src(i, 3))
                outA(n, 4) = NzD(src(i, 4))
                outA(n, 5) = NzD(src(i, 5))
                outA(n, 6) = NzD(src(i, 6))
                outA(n, 7) = NzD(src(i, 7))
            End If
            If Not jeUkupno Then
                sumKg = sumKg + NzD(src(i, 2))
                sumVal = sumVal + NzD(src(i, 6))
            End If
        Case IZ_ROBA
            If tip = "OM" Then
                outA(n, 1) = IzDatCell(src(i, 1))
                outA(n, 2) = NzS(src(i, 2))
                outA(n, 3) = NzS(src(i, 3))
                outA(n, 4) = NzS(src(i, 4))
                outA(n, 5) = NzS(src(i, 5))
                outA(n, 6) = NzD(src(i, 6))
                outA(n, 7) = NzD(src(i, 7))
                outA(n, 8) = NzD(src(i, 8))
                ' Prazno kad nema prijema JE poruka (RF-06) -- ne "0,00".
                outA(n, 9) = FmtIliPrazno(src(i, 9))
                outA(n, 10) = FmtIliPrazno(src(i, 10))
                If IsNumeric(src(i, 11)) And Not IsEmpty(src(i, 11)) Then
                    outA(n, 11) = Format$(CDbl(src(i, 11)), "0.00") & "%"
                Else
                    outA(n, 11) = NzS(src(i, 11))   ' "nema prijema" / "nejasan vlasnik"
                End If
                outA(n, 12) = NzS(src(i, 12))       ' "OTP|<id>" ili prazno
                If Not jeUkupno Then sumKg = sumKg + NzD(src(i, 6))
            ElseIf tip = "Kupac" Then
                ' Prijemnice kupca (ReportPrijemniceKupca fiksne kolone).
                outA(n, 1) = IzDatCell(src(i, 1))
                outA(n, 2) = NzS(src(i, 2))
                outA(n, 3) = NzS(src(i, 3))
                outA(n, 4) = NzS(src(i, 4))
                outA(n, 5) = NzS(src(i, 5))
                outA(n, 6) = NzD(src(i, 6))
                outA(n, 7) = NzD(src(i, 7))
                outA(n, 8) = NzD(src(i, 8))
                outA(n, 9) = IIf(Len(NzS(src(i, 9))) > 0, "PRJ|" & NzS(src(i, 9)), "")
                If Not jeUkupno Then
                    sumKg = sumKg + NzD(src(i, 6))
                    sumVal = sumVal + NzD(src(i, 8))
                End If
            Else
                outA(n, 1) = NzS(src(i, 1))
                outA(n, 2) = NzS(src(i, 2))
                outA(n, 3) = NzD(src(i, 3))
                outA(n, 4) = NzD(src(i, 4))
                If Not jeUkupno Then
                    sumKg = sumKg + NzD(src(i, 3))
                    sumVal = sumVal + NzD(src(i, 4))
                End If
            End If
        Case IZ_AMB
            ' UKUPNO red: legacy drzi "UKUPNO" u koloni DATUMA -- ovde ide u
            ' kolonu Mesto (datumska kolona ne sme tekst; kvar celije se broji).
            If jeUkupno Then
                outA(n, 1) = 0#
                outA(n, 2) = IZ_LBL_UKUPNO
            Else
                outA(n, 1) = IzDatCell(src(i, 1))
                outA(n, 2) = NzS(src(i, 2))
            End If
            outA(n, 3) = NzS(src(i, 3))
            outA(n, 4) = NzS(src(i, 4))
            outA(n, 5) = GajbeIliPrazno(src(i, 5))
            outA(n, 6) = GajbeIliPrazno(src(i, 6))
            ' "AMB|<DokTip>|<DokID>" -> dve prenosne kolone (ruta stampe trazi
            ' oba; tip ambalaze je vidljiva kolona 3 istog reda).
            outA(n, 7) = ""
            outA(n, 8) = ""
            ref = NzS(src(i, 7))
            If Left$(ref, 4) = "AMB|" Then
                p = Split(ref, "|")
                If UBound(p) >= 2 Then
                    outA(n, 7) = p(1)
                    outA(n, 8) = p(2)
                End If
            End If
        Case IZ_ISPL
            outA(n, 1) = NzS(src(i, 1))
            outA(n, 2) = NzD(src(i, 2))
            outA(n, 3) = NzD(src(i, 3))
            outA(n, 4) = NzD(src(i, 4))
            outA(n, 5) = NzD(src(i, 5))
            If Not jeUkupno Then sumVal = sumVal + NzD(src(i, 5))
        Case IZ_ZBIR
            outA(n, 1) = NzS(src(i, 1))
            If tip = "Vozac" Then
                outA(n, 2) = NzD(src(i, 2))
                outA(n, 3) = NzD(src(i, 3))
                outA(n, 4) = NzD(src(i, 4))
                If IsNumeric(src(i, 5)) Then
                    outA(n, 5) = Format$(CDbl(src(i, 5)), "0.00") & "%"
                Else
                    outA(n, 5) = NzS(src(i, 5))
                End If
            Else
                outA(n, 2) = NzS(src(i, 2))
                outA(n, 3) = NzD(src(i, 3))
                outA(n, 4) = NzD(src(i, 4))
                outA(n, 5) = NzD(src(i, 5))
                If Not jeUkupno Then
                    sumKg = sumKg + NzD(src(i, 3))
                    sumVal = sumVal + NzD(src(i, 4))
                End If
            End If
        Case IZ_CENA
            outA(n, 1) = NzS(src(i, 1))
            outA(n, 2) = NzD(src(i, 2))
            outA(n, 3) = NzD(src(i, 3))
            outA(n, 4) = NzD(src(i, 4))
            sumKg = sumKg + NzD(src(i, 2))
            sumVal = sumVal + NzD(src(i, 3))
        Case IZ_MANJAK
            outA(n, 1) = NzS(src(i, 1))
            outA(n, 2) = NzD(src(i, 2))
            outA(n, 3) = FmtIliPrazno(src(i, 3))
            outA(n, 4) = FmtIliPrazno(src(i, 4))
            If IsNumeric(src(i, 5)) And Not IsEmpty(src(i, 5)) Then
                outA(n, 5) = Format$(CDbl(src(i, 5)), "0.00") & "%"
            Else
                outA(n, 5) = NzS(src(i, 5))
            End If
            outA(n, 6) = NzD(src(i, 6))
            If Not jeUkupno Then sumKg = sumKg + NzD(src(i, 2))
        Case IZ_KART
            outA(n, 1) = IzDatCell(src(i, 1))
            outA(n, 2) = NzS(src(i, 2))
            ' Opis + parcela u istoj koloni, kao legacy prikaz.
            outA(n, 3) = NzS(src(i, 4))
            If Len(NzS(src(i, 3))) > 0 Then
                outA(n, 3) = outA(n, 3) & " / " & NzS(src(i, 3))
            End If
            outA(n, 4) = NzD(src(i, 5))
            outA(n, 5) = NzD(src(i, 6))
            outA(n, 6) = NzD(src(i, 7))
            outA(n, 7) = NzD(src(i, 8))
            outA(n, 8) = NzS(src(i, 9))            ' "OTK|<id>" / "NOV" / "MAG" / "AMB"
            ' Podnozje: neto promet prikazanih redova (zaduzenja - razduzenja)
            ' -- "Vrednost 0,00" uz punu karticu je izgledalo kao kvar.
            sumVal = sumVal + NzD(src(i, 5)) - NzD(src(i, 6))
        Case IZ_AMBK
            outA(n, 1) = IzDatCell(src(i, 1))
            outA(n, 2) = NzS(src(i, 2))
            outA(n, 3) = NzS(src(i, 3))
            outA(n, 4) = GajbeIliPrazno(src(i, 4))
            outA(n, 5) = GajbeIliPrazno(src(i, 5))
            outA(n, 6) = NzD(src(i, 6))
        Case IZ_OTKL
            outA(n, 1) = IzDatCell(src(i, 1))
            outA(n, 2) = NzS(src(i, 2))
            outA(n, 3) = NzS(src(i, 3))
            outA(n, 4) = NzS(src(i, 4))
            outA(n, 5) = NzS(src(i, 5))
            outA(n, 6) = NzD(src(i, 6))
            outA(n, 7) = NzD(src(i, 7))
            outA(n, 8) = NzS(src(i, 8))            ' "OTK|<id>"
            sumKg = sumKg + NzD(src(i, 6))
            sumVal = sumVal + NzD(src(i, 7))
        Case IZ_RANG
            ' Rang je mesto na CELOJ listi (i = indeks u sortiranom snimku),
            ' ne redni broj posle pretrage -- isti razlog kao na Dokumentima.
            outA(n, 1) = i
            outA(n, 2) = NzS(src(i, 2))
            outA(n, 3) = NzS(src(i, 3))
            outA(n, 4) = NzD(src(i, 4))
            outA(n, 5) = "KOP|" & NzS(src(i, 1))
            sumVal = sumVal + NzD(src(i, 4))
    End Select
End Sub

'--------------------------------------------------------------- POMOCNI
Private Function NzD(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then NzD = CDbl(v)
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    NzS = Trim$(CStr(v))
End Function

' Datum -> serijski broj za kolonu tipa "date" (0 = uredno prazno).
Private Function IzDatCell(ByVal v As Variant) As Double
    Dim d As Double
    If IsDate(v) Then
        d = Int(CDbl(CDate(v)))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        d = Int(CDbl(v))
    End If
    If d < 1 Or d > IZ_DAT_MAX Then Exit Function
    IzDatCell = d
End Function

' Kolone kod kojih je PRAZNO poruka ("nema prijema"): broj se formatira,
' sve ostalo ostaje prazno -- nikad "0,00" umesto oznake (FM-0028 #5).
Private Function FmtIliPrazno(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) Then FmtIliPrazno = FmtKolicina(CDbl(v))
End Function

' Gajbe: ceo broj, nula i prazno se prikazuju PRAZNO (legacy pregled).
Private Function GajbeIliPrazno(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) Then
        If CDbl(v) <> 0 Then GajbeIliPrazno = Format$(CDbl(v), "#,##0")
    End If
End Function

Private Sub ResetZonskeBrojke()
    mZonaOmAvans = Empty
    mZonaAgro = Empty
    mZonaIsplPrimljeno = Empty
    mZonaIsplPodeljeno = Empty
    mZonaIsplKod = Empty
    mZonaKartSaldo = Empty
    mZonaKartSaldoAmb = Empty
End Sub

'=====================================================================
' KONTEKST: tip, rezim, entitet, opseg.
'=====================================================================
Private Function TrenutniTip() As String
    If Len(mTip) = 0 Then mTip = "OM"
    TrenutniTip = mTip
End Function

' Da li POSLEDNJI prikazani kontekst jos ima snimak (za guard stampe i hint).
Private Function SnimakPostoji() As Boolean
    If mSnimci Is Nothing Then Exit Function
    If Len(mSnimakKljuc) = 0 Then Exit Function
    SnimakPostoji = mSnimci.Exists(mSnimakKljuc)
End Function

' Cist ID entiteta iz comboa (druga kolona). NERAZRESEN UNOS NIJE ENTITET:
' GetComboID daje stabilnu vrednost samo dok je stavka stvarno izabrana.
Private Function SiroviEntitet() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mTestId) > 0 Then
            SiroviEntitet = mTestId
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrIzEnt")
    If c Is Nothing Then Exit Function
    SiroviEntitet = GetComboID(c)
    Err.Clear
End Function

' Entitet za CITANJE: izbor operatera, a dok ga nema -- podrazumevani (prvi
' entitet tipa; v. mDefaultId). Tako prikaz od prvog trenutka nosi podatke
' (legacy AutoRefresh), a combo ostaje prazan i panel izbora zdrav.
Private Function IzabraniEntitet() As String
    IzabraniEntitet = SiroviEntitet()
    If Len(IzabraniEntitet) = 0 Then IzabraniEntitet = mDefaultId
End Function

' Tekst iz datumskih polja zone; u testu polja nema.
Private Sub OpsegPolja(ByRef odTxt As String, ByRef doTxt As String)
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrIzOd")
    If Not c Is Nothing Then odTxt = Trim$(CStr(c.text))
    Set c = Kontrola("scrIzDo")
    If Not c Is Nothing Then doTxt = Trim$(CStr(c.text))
    Err.Clear
End Sub

Private Sub OpsegGranice(ByRef odN As Double, ByRef doN As Double)
    Dim odTxt As String, doTxt As String
    If IsTestMode() Then
        If mTestOd > 0 Or mTestDo > 0 Then
            odN = mTestOd
            doN = mTestDo
            Exit Sub
        End If
    End If
    OpsegPolja odTxt, doTxt
    odN = IzDatGranica(odTxt)
    doN = IzDatGranica(doTxt)
End Sub

' Datum kao GRANICA opsega; 0 = nema granice. Prazan ili nepotpun unos nije
' greska -- dok operater kuca "21." nema smisla praznjenje liste. ISTO
' pravilo kao DatGranica u modScrDokumenti (specOdT presedan) -- ne izmislja
' se novo parsiranje.
Public Function IzDatGranica(ByVal s As String) As Double
    Dim d As Date
    On Error Resume Next
    If Len(Trim$(s)) = 0 Then Exit Function
    If TryParseDateValue(s, d) Then IzDatGranica = Int(CDbl(d))
End Function

' Naziv entiteta za naslov stampe -- iz sifarnika, po ID-u koji je STVARNO
' ucitan (ne iz comboa, koji se u medjuvremenu mogao promeniti).
Private Function EntitetNaziv(ByVal tip As String, ByVal iD As String, _
                              ByVal zbirni As Boolean) As String
    On Error Resume Next
    If zbirni Then
        EntitetNaziv = Poruka("OTKUI_IZ_SVI")
        Exit Function
    End If
    Select Case tip
        Case "OM"
            EntitetNaziv = NzS(LookupValue(TBL_STANICE, "StanicaID", iD, "Naziv"))
        Case "Kupac"
            EntitetNaziv = NzS(LookupValue(TBL_KUPCI, "KupacID", iD, "Naziv"))
        Case "Vozac"
            EntitetNaziv = Trim$(NzS(LookupValue(TBL_VOZACI, "VozacID", iD, "Ime")) & _
                          " " & NzS(LookupValue(TBL_VOZACI, "VozacID", iD, "Prezime")))
        Case "Kooperant"
            EntitetNaziv = Trim$(NzS(LookupValue(TBL_KOOPERANTI, "KooperantID", iD, "Ime")) & _
                          " " & NzS(LookupValue(TBL_KOOPERANTI, "KooperantID", iD, "Prezime")))
    End Select
    If Len(EntitetNaziv) = 0 Then EntitetNaziv = iD
    EntitetNaziv = EntitetNaziv & " (" & iD & ")"
    Err.Clear
End Function

Private Function TipLabela(ByVal tip As String) As String
    Select Case tip
        Case "OM":        TipLabela = Poruka("OTKUI_SEGIZ_TIP_OM")
        Case "Kupac":     TipLabela = Poruka("OTKUI_SEGIZ_TIP_KUP")
        Case "Vozac":     TipLabela = Poruka("OTKUI_SEGIZ_TIP_VOZ")
        Case "Kooperant": TipLabela = Poruka("OTKUI_SEGIZ_TIP_KOOP")
    End Select
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod reda polja -- LABELA, ne Frame (Frame je prozorska
    ' kontrola i crta se iznad bezprozorskih). Napravljena PRVA, ostaje ispod.
    modUiKit.NewLbl z, "izBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "izCap", UCase$(Poruka("OTKUI_SCRIZ_CAP")), PAD, IZ_Y_CAP, _
                    200, 11, TS_MICRO, True, C_MUTED, -1

    ' PREKIDACI: entitet-tip (4) i rezim (2). Vrsta "seg" (NewSegBtn), ne
    ' "btn" -- par. 7.7: clsFlatBtn.IsSelected priznaje izabrano stanje samo
    ' za nav/chip/seg; kao "btn" bi hover-out vracao belu.
    modUiKit.NewSegBtn z, "scrIzTipOM", Poruka("OTKUI_SEGIZ_TIP_OM"), _
                       PAD, IZ_Y_SEG, 46, IZ_SEG_H, True
    modUiKit.NewSegBtn z, "scrIzTipKup", Poruka("OTKUI_SEGIZ_TIP_KUP"), _
                       PAD + 50, IZ_Y_SEG, 56, IZ_SEG_H, False
    modUiKit.NewSegBtn z, "scrIzTipVoz", Poruka("OTKUI_SEGIZ_TIP_VOZ"), _
                       PAD + 110, IZ_Y_SEG, 56, IZ_SEG_H, False
    modUiKit.NewSegBtn z, "scrIzTipKoop", Poruka("OTKUI_SEGIZ_TIP_KOOP"), _
                       PAD + 170, IZ_Y_SEG, 76, IZ_SEG_H, False

    modUiKit.NewSegBtn z, "scrIzRezP", Poruka("OTKUI_SEGIZ_POJ"), _
                       PAD + 270, IZ_Y_SEG, 84, IZ_SEG_H, True
    modUiKit.NewSegBtn z, "scrIzRezZ", Poruka("OTKUI_SEGIZ_ZBI"), _
                       PAD + 358, IZ_Y_SEG, 60, IZ_SEG_H, False

    ' Dve brojke desno: OM avans (nerasporedjen) i agro (nerasporedjena) --
    ' specijalni redovi ReportSaldoOM, izdvojeni iz mreze (v. zaglavlje).
    ' Kontrolne brojke isplate dele iste dve kutije (v. OsveziBrojke).
    For i = 0 To 1
        modUiKit.NewLbl z, "izKL" & i, "", 0, IZ_Y_CAP, IZ_KPI_W, 11, _
                        TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "izKV" & i, ChrW(8212), 0, IZ_Y_SEG - 2, IZ_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' DETALJ TRAKA (drill-down): klik na red sa dokumentom pokazuje njegove
    ' stavke desno -- isti raspored kao traka korpe na Platnim nalozima.
    modUiKit.NewLbl z, "izDetCap", "", 0, IZ_Y_LBL, IZ_DET_W, 11, _
                    TS_MICRO, True, C_MUTED, -1
    For i = 0 To IZ_DET_N - 1
        modUiKit.NewLbl z, "izDetR" & i, "", 0, IZ_Y_LBL + 14 + i * 12, _
                        IZ_DET_W, 12, TS_META, False, C_FOREST, -1
    Next i

    ' POLJA. Pravi ih ljuska (NewFieldG); prefiks "scr" je OBAVEZAN, a kombo
    ' MORA biti polje (okvir nm + kontrola nmT) -- FindCombo trazi taj oblik.
    modOtkupUI.NewFieldG z, "scrIzEnt", Poruka("OTKUI_FLD_IZ_ENT"), "cmb", "", _
                         1, False, False, "IZ"
    modOtkupUI.NewFieldG z, "scrIzOd", Poruka("OTKUI_FLD_IZ_OD"), "txt", "", _
                         1, False, False, "IZ"
    modOtkupUI.NewFieldG z, "scrIzDo", Poruka("OTKUI_FLD_IZ_DO"), "txt", "", _
                         1, False, False, "IZ"

    ' Legacy default opsega: 1.1. tekuce godine -- danas.
    On Error Resume Next
    z.Controls("scrIzOd").Controls("scrIzOdT").text = "1.1." & Year(Date)
    z.Controls("scrIzDo").Controls("scrIzDoT").text = Format$(Date, "d.m.yyyy")
    On Error GoTo 0

    modUiKit.NewLbl z, "izHint", "", PAD, IZ_Y_HINT, 420, 12, TS_META, False, C_MUTED, -1

    ' Stampa aktivnog izvestaja (tabelarni PDF) + kartica po sablonu (samo
    ' na listama kartica -- vidljivost daje raspored).
    modUiKit.BtnV z, "scrIzPrint", Poruka("OTKUI_BTN_IZ_PRINT"), PAD, IZ_Y_BTN, _
                  156, IZ_BTN_H, "primary"
    modUiKit.BtnV z, "scrIzKartPdf", Poruka("OTKUI_BTN_IZ_KARTPDF"), PAD + 164, IZ_Y_BTN, _
                  170, IZ_BTN_H, "soft"

    modUiKit.NewLbl z, "izLnB", "", 0, IZ_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = IZ_ZONA_H
End Function

' Raspored + boje prekidaca. Boja stoji uz RASPORED (par. 7.7): koji je tip/
' rezim izabran je JEDNA odluka, pa boja i raspored ne mogu da se razidju.
Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, kx As Single
    Dim naKartici As Boolean
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    z.Controls("izBg").Left = PAD - 10
    z.Controls("izBg").top = IZ_Y_LBL - 8
    z.Controls("izBg").width = w - 2 * (PAD - 10)
    z.Controls("izBg").Height = IZ_Y_BTN - IZ_Y_LBL + 2

    OsveziPrekidace z

    ' Brojke uz desnu ivicu.
    For i = 0 To 1
        kx = w - PAD - (2 - i) * IZ_KPI_W
        z.Controls("izKL" & i).Left = kx
        z.Controls("izKV" & i).Left = kx
    Next i

    PoljeX z, "scrIzEnt", PAD, 250, IZ_Y_LBL
    PoljeX z, "scrIzOd", PAD + 258, 86, IZ_Y_LBL
    PoljeX z, "scrIzDo", PAD + 352, 86, IZ_Y_LBL

    ' Entitet se bira kad ga lista trazi: pojedinacni rezim (osim ranga) i
    ' zbirna AMBALAZA (legacy agregat po tipu ZA izabranog) -- krug 9.
    z.Controls("scrIzEnt").Visible = IzTrebaEntitet(Scr_Lista(), mZbirni)

    ' Detalj traka uzima desno; polja i hint dele OSTATAK. Na uskom ekranu
    ' traka nestaje -- isti kompromis kao korpa na Platnim nalozima.
    Dim wPolja As Single, detVidi As Boolean, dx As Single
    wPolja = w - IZ_DET_W - PAD
    detVidi = (wPolja >= IZ_POLJA_MIN)
    If Not detVidi Then wPolja = w
    dx = w - IZ_DET_W
    z.Controls("izDetCap").Left = dx
    z.Controls("izDetCap").Visible = detVidi
    For i = 0 To IZ_DET_N - 1
        z.Controls("izDetR" & i).Left = dx
        z.Controls("izDetR" & i).Visible = detVidi
    Next i

    z.Controls("izHint").width = wPolja - 2 * PAD

    ' Kartice imaju JEDNU stampu -- legacy sablon sa rekapitulacijom robe,
    ' BPG-om i potpisima je "ono sto je ispravno i potrebno" (odluka posle
    ' smoke kruga 5). Generic tabelarni PDF se na karticama NE nudi: dve
    ' konkurentske kartice su zbunjivost, ne izbor.
    naKartici = (Scr_Lista() = IZ_KART Or Scr_Lista() = IZ_AMBK)
    modUiKit.MoveBtn z, "scrIzPrint", PAD, IZ_Y_BTN
    modUiKit.MoveBtn z, "scrIzKartPdf", PAD, IZ_Y_BTN
    modUiKit.BoxShow z, "scrIzPrint", Not naKartici
    modUiKit.BoxShow z, "scrIzKartPdf", naKartici

    z.Controls("izLnB").width = w
End Sub

' Boje prekidaca tipa i rezima (BoxState + RebaseSink -- par. 7.7: render
' koji promeni boju javlja novu osnovu, inace je hover-out vrati).
Private Sub OsveziPrekidace(ByVal z As Object)
    Dim tip As String
    On Error Resume Next
    tip = TrenutniTip()
    SegBoja z, "scrIzTipOM", (tip = "OM")
    SegBoja z, "scrIzTipKup", (tip = "Kupac")
    SegBoja z, "scrIzTipVoz", (tip = "Vozac")
    SegBoja z, "scrIzTipKoop", (tip = "Kooperant")
    SegBoja z, "scrIzRezP", Not mZbirni
    SegBoja z, "scrIzRezZ", mZbirni
End Sub

Private Sub SegBoja(ByVal z As Object, ByVal nm As String, ByVal sel As Boolean)
    On Error Resume Next
    modUiKit.BoxState z, nm, IIf(sel, C_FOREST, C_WHITE), _
                      IIf(sel, C_CREAM, C_FOREST), sel
    modOtkupUI.RebaseSink nm
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
    Set Zona = modOtkupUI.ScreenZone("IZVESTAJI")
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
    RasporediPolja z, z.width
    OsveziHint z
    OsveziBrojke z
End Sub

' Hint: nedostupna kombinacija kaze ZASTO je lista prazna (FM-0029 merilo --
' nikad pun naslov nad trajno praznom listom); pojedinacni rezim bez izbora
' kaze sta fali; inace opis prikazanog konteksta.
Private Sub OsveziHint(ByVal z As Object)
    Dim s As String
    On Error Resume Next
    Select Case mHintKljuc
        Case "OTKUI_IZ_HINT_NEDOSTUPNO"
            s = Poruka("OTKUI_IZ_HINT_NEDOSTUPNO") & "  " & ChrW(183) & "  " & _
                TipLabela(TrenutniTip()) & ", " & _
                Poruka(IIf(mZbirni, "OTKUI_SEGIZ_ZBI", "OTKUI_SEGIZ_POJ"))
            ' Kartice postoje samo za kooperanta -- hint kaze i KUDA, ne samo
            ' zasto (prvi smoke: "nema gde da se izabere kooperant").
            If Scr_Lista() = IZ_KART Or Scr_Lista() = IZ_AMBK Then
                s = s & "  " & ChrW(183) & "  " & Poruka("OTKUI_IZ_HINT_KART_KOOP")
            End If
        Case "OTKUI_IZ_HINT_IZABERI"
            s = Poruka("OTKUI_IZ_HINT_IZABERI")
        Case Else
            If SnimakPostoji() Then
                s = TipLabela(mCtxTip) & ": " & mCtxEntNaziv & "  " & _
                    ChrW(183) & "  " & OpsegLabela()
            Else
                s = Poruka("OTKUI_IZ_HINT")
            End If
    End Select
    z.Controls("izHint").caption = s
End Sub

' Period STVARNO prikazanih podataka (nikad iz polja -- ona opisuju sledece
' citanje).
Private Function OpsegLabela() As String
    Dim s1 As String, s2 As String
    s1 = IIf(mCtxOd > 0, Format$(CDate(mCtxOd), "d.m.yyyy"), Poruka("OTKUI_IZ_BEZ_GRANICE"))
    s2 = IIf(mCtxDo > 0, Format$(CDate(mCtxDo), "d.m.yyyy"), Poruka("OTKUI_IZ_BEZ_GRANICE"))
    OpsegLabela = s1 & " - " & s2
End Function

' Dve brojke zone -- pune se iz SNIMKA aktivne liste (SALDO za OM: avans +
' agro; ISPLATA: primljeno/kod otkupca), inace crta. Nula i "nema podatka"
' nisu ista brojka.
Private Sub OsveziBrojke(ByVal z As Object)
    Dim crta As String
    On Error Resume Next
    crta = ChrW(8212)

    If Scr_Lista() = IZ_ISPL And Not IsEmpty(mZonaIsplPrimljeno) Then
        z.Controls("izKL0").caption = UCase$(Poruka("OTKUI_KPI_IZ_OMAVANS_PRIM"))
        z.Controls("izKV0").caption = Format$(NzD(mZonaIsplPrimljeno), "#,##0")
        z.Controls("izKL1").caption = UCase$(Poruka("OTKUI_KPI_IZ_KOD_OTKUPCA"))
        z.Controls("izKV1").caption = Format$(NzD(mZonaIsplKod), "#,##0")
        Exit Sub
    End If

    ' Kartice: zavrsni saldo perioda kao KPI (kolona salda u mrezi je running
    ' po redu; JEDNA brojka "gde smo na kraju" zivi ovde i u stampi).
    If Scr_Lista() = IZ_KART Or Scr_Lista() = IZ_AMBK Then
        z.Controls("izKL0").caption = UCase$(Poruka("OTKUI_KPI_IZ_SALDO"))
        If IsEmpty(mZonaKartSaldo) Then
            z.Controls("izKV0").caption = crta
        Else
            z.Controls("izKV0").caption = Format$(NzD(mZonaKartSaldo), "#,##0")
        End If
        z.Controls("izKL1").caption = UCase$(Poruka("OTKUI_KPI_IZ_SALDOAMB"))
        If Scr_Lista() = IZ_AMBK Or IsEmpty(mZonaKartSaldoAmb) Then
            ' AMBK saldo JESTE gajbe -- druga kutija bi ponovila prvu.
            z.Controls("izKV1").caption = crta
            If Scr_Lista() = IZ_AMBK Then z.Controls("izKL1").caption = ""
        Else
            z.Controls("izKV1").caption = Format$(NzD(mZonaKartSaldoAmb), "#,##0")
        End If
        Exit Sub
    End If

    z.Controls("izKL0").caption = UCase$(Poruka("OTKUI_KPI_IZ_OMAVANS"))
    z.Controls("izKL1").caption = UCase$(Poruka("OTKUI_KPI_IZ_AGRO"))
    If IsEmpty(mZonaOmAvans) Then
        z.Controls("izKV0").caption = crta
    Else
        z.Controls("izKV0").caption = Format$(NzD(mZonaOmAvans), "#,##0")
    End If
    If IsEmpty(mZonaAgro) Then
        z.Controls("izKV1").caption = crta
    Else
        z.Controls("izKV1").caption = Format$(NzD(mZonaAgro), "#,##0")
    End If
End Sub

'------------------------------------------------------- DETALJ TRAKA
' Klik na red -> stavke dokumenta u desnoj traci zone (drill-down; smoke
' krug 3 je trazio "onaj detalj o kom je bilo reci"). Pravi padajuci redovi
' mreze ostaju odlozen posao ljuske (par. 5/Faza C) -- traka je ono sto
' ekran moze BEZ dopune ugovora, i pokriva legacy "Detalji otkupa" sustinu:
' sve stavke izabranog otkupnog lista.
Private Sub OsveziDetalj(ByVal red As Long)
    Dim ref As String, kljuc As String
    Dim linije As Variant, naslov As String

    kljuc = Scr_Lista()
    mDetalj = Empty
    naslov = ""

    Select Case kljuc
        Case IZ_OTKL
            ref = NzS(modOtkupUI.GridCell(red, 8))
            If Left$(ref, 4) = "OTK|" Then
                naslov = Poruka("OTKUI_IZ_DET_OTKUP") & " " & NzS(modOtkupUI.GridCell(red, 2))
                linije = IzDetaljOtkupLista(Mid$(ref, 5))
            End If
        Case IZ_KART
            ref = NzS(modOtkupUI.GridCell(red, 8))
            If Left$(ref, 4) = "OTK|" Then
                naslov = Poruka("OTKUI_IZ_DET_OTKUP") & " " & NzS(modOtkupUI.GridCell(red, 2))
                linije = IzDetaljOtkupLista(Mid$(ref, 5))
            ElseIf Len(ref) > 0 Then
                ' NOV / MAG / AMB red: sustina je vec u opisu reda.
                naslov = Poruka("OTKUI_IZ_DET_RED")
                linije = Array(NzS(modOtkupUI.GridCell(red, 3)), _
                               Poruka("OTKUI_IZ_DET_IZNOS") & " " & _
                               DetIznosKartice(red))
            End If
        Case IZ_ROBA
            If mCtxTip = "Kupac" Then
                ref = NzS(modOtkupUI.GridCell(red, 9))
                If Left$(ref, 4) = "PRJ|" Then
                    naslov = Poruka("OTKUI_IZ_DET_PRIJEMNICA") & " " & NzS(modOtkupUI.GridCell(red, 2))
                    linije = IzDetaljPrijemnice(Mid$(ref, 5))
                End If
            Else
                ref = NzS(modOtkupUI.GridCell(red, 12))
                If Left$(ref, 4) = "OTP|" Then
                    naslov = Poruka("OTKUI_IZ_DET_OTPREMNICA") & " " & NzS(modOtkupUI.GridCell(red, 2))
                    linije = IzDetaljOtpremnice(Mid$(ref, 5))
                End If
            End If
        Case IZ_AMB
            If Len(NzS(modOtkupUI.GridCell(red, 8))) > 0 Then
                naslov = Poruka("OTKUI_IZ_DET_DOKUMENT") & " " & NzS(modOtkupUI.GridCell(red, 4))
                linije = Array( _
                    NzS(modOtkupUI.GridCell(red, 7)), _
                    Poruka("OTKUI_HDA_TIP") & ": " & NzS(modOtkupUI.GridCell(red, 3)), _
                    Poruka("OTKUI_HDA_ULAZ") & " " & NzS(modOtkupUI.GridCell(red, 5)) & _
                    "   " & Poruka("OTKUI_HDA_IZLAZ") & " " & NzS(modOtkupUI.GridCell(red, 6)))
            End If
    End Select

    If IsArray(linije) Then
        mDetalj = linije
        DetaljTraka naslov
    End If
End Sub

' Iznos novcanog/agro reda kartice: razduzenje ili zaduzenje, sta postoji.
Private Function DetIznosKartice(ByVal red As Long) As String
    Dim z As Double, r As Double
    On Error Resume Next
    z = CDbl(modOtkupUI.GridCell(red, 4))
    r = CDbl(modOtkupUI.GridCell(red, 5))
    If r <> 0 Then
        DetIznosKartice = Format$(r, "#,##0.00")
    Else
        DetIznosKartice = Format$(z, "#,##0.00")
    End If
    Err.Clear
End Function

' STAVKE OTKUPNOG LISTA -- legacy "Detalji otkupa" sustina, kao cist racun
' (testabilan bez forme): sve nestornirane linije ISTOG dokumenta (broj +
' stanica, kao ReprintOtkupniListByOtkupID koji stampa ceo BrDok), po
' linija "Vrsta Klasa  kg x cena = vrednost". Detalj nosi SAMO ono sto red
' liste NE pokazuje (smoke krug 4): kooperant je vec kolona reda pa se ne
' ponavlja, a UKUPNO dokumenta ide samo kad linija ima VISE (red pokazuje
' jednu) -- jednolinijski dokument bi njime dublirao sopstveni red.
Public Function IzDetaljOtkupLista(ByVal otkupID As String) As Variant
    Dim d As Variant, i As Long
    Dim cId As Long, cBr As Long, cSt As Long
    Dim cVr As Long, cKl As Long, cKol As Long, cCe As Long, cStorno As Long
    Dim cVoz As Long, cZb As Long
    Dim brDok As String, stanica As String, vozId As String, brZb As String
    Dim linije As Collection, kg As Double, cena As Double
    Dim totKg As Double, totVr As Double
    On Error GoTo EH

    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then Exit Function
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    cVr = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    cKl = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cCe = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    cStorno = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    cZb = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)

    ' Dokument izabrane linije (broj je scoped po stanici). Vozac i zbirna su
    ' dokumentski (sve linije ih dele) -- citaju se sa izabrane.
    For i = 1 To UBound(d, 1)
        If Trim$(CStr(d(i, cId))) = Trim$(otkupID) Then
            brDok = NzS(d(i, cBr))
            stanica = NzS(d(i, cSt))
            vozId = NzS(d(i, cVoz))
            brZb = NzS(d(i, cZb))
            Exit For
        End If
    Next i
    If Len(brDok) = 0 Then Exit Function

    Set linije = New Collection
    For i = 1 To UBound(d, 1)
        If NzS(d(i, cBr)) = brDok And NzS(d(i, cSt)) = stanica Then
            If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
                kg = NzD(d(i, cKol))
                cena = NzD(d(i, cCe))
                linije.Add NzS(d(i, cVr)) & " " & NzS(d(i, cKl)) & "  " & _
                           FmtKolicina(kg) & " x " & Format$(cena, "#,##0.00") & _
                           " = " & Format$(kg * cena, "#,##0.00")
                totKg = totKg + kg
                totVr = totVr + kg * cena
            End If
        End If
    Next i
    If linije.count = 0 Then Exit Function
    If linije.count > 1 Then
        linije.Add "UKUPNO  " & FmtKolicina(totKg) & " kg  " & _
                   ChrW(183) & "  " & Format$(totVr, "#,##0.00")
    End If

    ' Sledljivost dokumenta (smoke krug 5): vozac i zbirna sa lista, pa
    ' prijemnice te zbirne SA KUPCEM -- red kartice/liste nista od ovoga ne
    ' pokazuje. Prazan deo se preskace (nema = nema, ne izmislja se).
    Dim ctx As String
    If Len(vozId) > 0 Then
        ctx = Poruka("OTKUI_IZ_DET_VOZAC") & " " & EntitetNaziv("Vozac", vozId, False)
    End If
    If Len(brZb) > 0 Then
        If Len(ctx) > 0 Then ctx = ctx & "  " & ChrW(183) & "  "
        ctx = ctx & Poruka("OTKUI_IZ_DET_ZBIRNA") & " " & brZb
    End If
    If Len(ctx) > 0 Then linije.Add ctx
    If Len(brZb) > 0 Then DodajPrijemniceZbirne linije, brZb, vozId

    Dim res() As String, n As Long
    ReDim res(0 To linije.count - 1)
    For n = 1 To linije.count
        res(n - 1) = linije(n)
    Next n
    IzDetaljOtkupLista = res
    Exit Function
EH:
    ' Detalj je pregled -- pad citanja ostavlja praznu traku, ne obara klik.
End Function

' Osnovni podaci otpremnice + zbir njenih blokova (za red liste ROBA).
' Detalj otpremnice nosi SAMO ono sto kolone reda ROBA (OM) ne pokazuju
' (smoke krug 4): vozac (trazen izricito), broj otkupnih listova, ZBIRNA i
' PRIJEMNICE te zbirne (broj + kg). Otpremljene kg i kg blokova su vec
' kolone reda -- ne ponavljaju se. Prijemnica se vezuje kao u
' ReportOtkupRobaOM prvom koraku: isti BrojZbirne, pa suzenje po vozacu ako
' broj dele dve zbirne; detalj je pregled pa sme fail-open (pokazuje sve
' kandidate), racun manjka i dalje radi samo fail-closed put.
Public Function IzDetaljOtpremnice(ByVal otpID As String) As Variant
    Dim d As Variant, i As Long
    Dim cId As Long, cBr As Long, cVoz As Long, cZb As Long
    Dim cOtkOtp As Long, cStorno As Long
    Dim brOtp As String, vozac As String, vozId As String, brZb As String
    Dim blokova As Long
    Dim linije As Collection
    On Error GoTo EH

    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    cVoz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC)
    cZb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    For i = 1 To UBound(d, 1)
        If Trim$(CStr(d(i, cId))) = Trim$(otpID) Then
            brOtp = NzS(d(i, cBr))
            vozId = NzS(d(i, cVoz))
            vozac = EntitetNaziv("Vozac", vozId, False)
            brZb = NzS(d(i, cZb))
            Exit For
        End If
    Next i
    If Len(brOtp) = 0 Then Exit Function

    d = GetTableData(TBL_OTKUP)
    If IsArray(d) Then
        cOtkOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
        cStorno = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
        For i = 1 To UBound(d, 1)
            If Trim$(CStr(d(i, cOtkOtp))) = Trim$(otpID) Then
                If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
                    blokova = blokova + 1
                End If
            End If
        Next i
    End If

    Set linije = New Collection
    linije.Add Poruka("OTKUI_IZ_DET_VOZAC") & " " & vozac
    linije.Add Poruka("OTKUI_IZ_DET_BLOKOVA") & " " & CStr(blokova)
    If Len(brZb) > 0 Then
        linije.Add Poruka("OTKUI_IZ_DET_ZBIRNA") & " " & brZb
        DodajPrijemniceZbirne linije, brZb, vozId
    End If

    Dim res() As String, n As Long
    ReDim res(0 To linije.count - 1)
    For n = 1 To linije.count
        res(n - 1) = linije(n)
    Next n
    IzDetaljOtpremnice = res
    Exit Function
EH:
End Function

' Prijemnice date zbirne u detalj: prvo po (BrojZbirne, VozacID); ako se
' nijedna ne poklopi po vozacu, svi kandidati po broju zbirne (fail-open).
' Linija nosi i KUPCA prijemnice (smoke krug 5: "firma koja je izdala
' prijemnicu") -- to je karika sledljivosti koju nijedan red liste ne kaze.
Private Sub DodajPrijemniceZbirne(ByVal linije As Collection, _
                                  ByVal brZb As String, ByVal vozId As String)
    Dim d As Variant, i As Long, krug As Long, nasao As Boolean
    Dim cZb As Long, cVoz As Long, cBr As Long, cKol As Long, cStorno As Long
    Dim cKup As Long, s As String
    On Error Resume Next
    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Sub
    cZb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    cStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    cKup = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    For krug = 1 To 2
        For i = 1 To UBound(d, 1)
            If NzS(d(i, cZb)) = brZb Then
                If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
                    If krug = 2 Or NzS(d(i, cVoz)) = vozId Then
                        s = Poruka("OTKUI_IZ_DET_PRIJEMNICA") & " " & _
                            NzS(d(i, cBr)) & "  " & ChrW(183) & "  " & _
                            FmtKolicina(NzD(d(i, cKol))) & " kg"
                        If Len(NzS(d(i, cKup))) > 0 Then
                            s = s & "  " & ChrW(183) & "  " & _
                                EntitetNaziv("Kupac", NzS(d(i, cKup)), False)
                        End If
                        linije.Add s
                        nasao = True
                    End If
                End If
            End If
        Next i
        If nasao Then Exit For
    Next krug
    Err.Clear
End Sub

' Detalj reda prijemnice (ROBA za kupca) -- SAMO sto kolone reda ne kazu:
' vozac, sorta, ambalaza (tip x kolicina, vraceno) i status fakturisanja.
Public Function IzDetaljPrijemnice(ByVal prjID As String) As Variant
    Dim d As Variant, i As Long
    Dim cId As Long, cVoz As Long, cSor As Long, cTipA As Long
    Dim cKolA As Long, cVrac As Long, cFakt As Long, cFid As Long
    Dim linije As Collection
    On Error GoTo EH

    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Function
    cId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    cVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    cSor = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cTipA = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    cKolA = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    cVrac = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA)
    cFakt = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO)
    cFid = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)

    For i = 1 To UBound(d, 1)
        If Trim$(CStr(d(i, cId))) = Trim$(prjID) Then
            Set linije = New Collection
            linije.Add Poruka("OTKUI_IZ_DET_VOZAC") & " " & _
                       EntitetNaziv("Vozac", NzS(d(i, cVoz)), False)
            If Len(NzS(d(i, cSor))) > 0 Then
                linije.Add Poruka("OTKUI_IZ_DET_SORTA") & " " & NzS(d(i, cSor))
            End If
            If Len(NzS(d(i, cTipA))) > 0 Then
                linije.Add Poruka("OTKUI_IZ_DET_AMB") & " " & NzS(d(i, cTipA)) & _
                           " x " & Format$(NzD(d(i, cKolA)), "#,##0") & _
                           "  (" & Poruka("OTKUI_IZ_DET_AMB_VRAC") & " " & _
                           Format$(NzD(d(i, cVrac)), "#,##0") & ")"
            End If
            If CStr(d(i, cFakt)) = "Da" Then
                linije.Add Poruka("OTKUI_IZ_DET_FAKTURA") & " " & NzS(d(i, cFid))
            Else
                linije.Add Poruka("OTKUI_IZ_DET_FAKTURA") & " " & _
                           Poruka("OTKUI_IZ_DET_NEFAKT")
            End If

            Dim res() As String, n As Long
            ReDim res(0 To linije.count - 1)
            For n = 1 To linije.count
                res(n - 1) = linije(n)
            Next n
            IzDetaljPrijemnice = res
            Exit Function
        End If
    Next i
    Exit Function
EH:
    ' Detalj je pregled -- pad citanja ostavlja praznu traku, ne obara klik.
End Function

' Crtanje trake: naslov + linije + preliv ("... jos N") -- preliv se
' PRIJAVLJUJE, lista koja se tiho odseca izgleda kao cela (par. 7.8).
Private Sub DetaljTraka(ByVal naslov As String)
    Dim z As Object, i As Long, n As Long
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    n = 0
    If IsArray(mDetalj) Then n = UBound(mDetalj) - LBound(mDetalj) + 1
    z.Controls("izDetCap").caption = UCase$(naslov)
    For i = 0 To IZ_DET_N - 1
        If i < n Then
            If i = IZ_DET_N - 1 And n > IZ_DET_N Then
                z.Controls("izDetR" & i).caption = ChrW(8230) & " " & _
                    Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & CStr(n - IZ_DET_N + 1)
            Else
                z.Controls("izDetR" & i).caption = CStr(mDetalj(LBound(mDetalj) + i))
            End If
        Else
            z.Controls("izDetR" & i).caption = ""
        End If
    Next i
End Sub

Private Sub OcistiDetalj()
    mDetalj = Empty
    DetaljTraka ""
End Sub

' Combo entiteta se puni PO TIPU (obrazac PuniPartnerCombo, par. 9): dve
' kolone, cist ID u drugoj, prikaz sa ID-jem (dva ista imena su stvarnost --
' fixture ima dva istoimena kooperanta). Kooperanti: samo AKTIVNI, kao
' legacy LoadEntiteti. Puni se ponovo posle ResetCache; cuva izbor.
Private Sub PuniEntitetCombo()
    Dim c As Object, tip As String
    Dim mapa As Object, k As Variant
    Dim cur As String, i As Long, idx As Long
    On Error GoTo EH

    tip = TrenutniTip()
    Set c = Kontrola("scrIzEnt")
    If c Is Nothing Then Exit Sub
    If mComboTip = tip Then Exit Sub

    mFill = True
    cur = GetComboID(c)

    c.Clear
    c.ColumnCount = 2
    c.ColumnWidths = "180 pt;0 pt"
    c.BoundColumn = 1
    c.TextColumn = 1

    Select Case tip
        Case "OM"
            Set mapa = BuildLookupDict(TBL_STANICE, "StanicaID", "Naziv")
        Case "Kupac"
            Set mapa = BuildLookupDict(TBL_KUPCI, "KupacID", "Naziv")
        Case "Vozac"
            Set mapa = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
        Case "Kooperant"
            Set mapa = AktivniKooperanti()
    End Select

    mDefaultId = ""
    If Not mapa Is Nothing Then
        For Each k In mapa.keys
            c.AddItem Trim$(CStr(mapa(k)))
            c.List(c.ListCount - 1, 1) = CStr(k)
            If Len(mDefaultId) = 0 Then mDefaultId = CStr(k)
        Next k
    End If

    ShowIDInComboDisplay c

    ' BEZ auto-izbora: pun tekst u combu bi FILTRIRAO ljuskin panel na tu
    ' jednu stavku (PopIndex ide po tekstu) -- prvi smoke je to video kao
    ' "dropdown ne radi". Podrazumevani entitet zato zivi u mDefaultId
    ' (IzabraniEntitet ga vraca dok izbora nema), a combo drzi tekst SAMO
    ' kad ga je operater sam postavio: njegov izbor preziviva refill.
    If c.ListCount > 0 And Len(cur) > 0 Then
        idx = -1
        For i = 0 To c.ListCount - 1
            If CStr(c.List(i, 1)) = cur Then idx = i: Exit For
        Next i
        If idx >= 0 Then c.ListIndex = idx
    End If

    mComboTip = tip
    mFill = False
    Exit Sub
EH:
    mFill = False
    Debug.Print "modScrIzvestaji.PuniEntitetCombo PAO: " & Err.Number & " " & Err.description
End Sub

' Samo aktivni kooperanti -- legacy LoadEntiteti pravilo (BuildLookupDict ne
' filtrira aktivnost).
Private Function AktivniKooperanti() As Object
    Dim d As Object, src As Variant, i As Long
    Dim cId As Long, cIme As Long, cPr As Long, cAkt As Long
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo Gotovo

    src = GetTableData(TBL_KOOPERANTI)
    If Not IsArray(src) Then GoTo Gotovo
    cId = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    cAkt = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
    If cId = 0 Then GoTo Gotovo

    For i = 1 To UBound(src, 1)
        If cAkt = 0 Or CStr(src(i, cAkt)) = STATUS_AKTIVAN Then
            Dim iD As String
            iD = NzS(src(i, cId))
            If Len(iD) > 0 And Not d.Exists(iD) Then
                d(iD) = Trim$(NzS(src(i, cIme)) & " " & NzS(src(i, cPr)))
            End If
        End If
    Next i
Gotovo:
    Set AktivniKooperanti = d
End Function

'=====================================================================
' STAMPE. Ne verifikuju se automatski -- smoke checklista.
'=====================================================================

' KOJE se kolone SABIRAJU u stampanom UKUPNO redu -- po listi, 1-based
' indeksi VIDLJIVIH kolona. Tip kolone (kg/rsd/rest) opisuje PRIKAZ, ne
' aditivnost: prosecna cena, prosek gajbi i RUNNING SALDO su numericki a
' zbir im ne znaci nista (recenzija PR #245, blocker 2 -- generic suma je
' na kartici davala "UKUPNO SALDO" kao zbir medjustanja). Politika prati
' legacy UKUPNO redove: sabira se promet, nikad prosek i nikad stanje.
Public Function IzSabirljive(ByVal kljuc As String, ByVal tip As String, _
                             Optional ByVal zbirni As Boolean = False) As Variant
    ' Zbirni oblici (krug 9): sve su promet/stanje-po-entitetu kolone --
    ' aditivne preko stanica (i saldo: razlika je aditivna), agro "rest" isto.
    If zbirni Then
        Select Case kljuc
            Case IZ_SALDO
                If tip = "Kupac" Then
                    IzSabirljive = Array(2, 3, 4, 5, 6)
                Else
                    IzSabirljive = Array(2, 3, 4, 5, 6, 7)
                End If
                Exit Function
            Case IZ_ROBA
                IzSabirljive = Array(2, 3)
                Exit Function
            Case IZ_ISPL
                IzSabirljive = Array(2, 3, 4, 5)
                Exit Function
            Case IZ_AMB
                IzSabirljive = Array(2, 3)
                Exit Function
        End Select
    End If
    Select Case kljuc
        Case IZ_SALDO
            If tip = "Kupac" Then
                IzSabirljive = Array(2, 4, 5, 6, 7)    ' bez 3 = cena
            Else
                IzSabirljive = Array(2, 3, 4, 5, 6, 7)
            End If
        Case IZ_ROBA
            If tip = "OM" Then
                IzSabirljive = Array(6, 7, 8, 9, 10)   ' bez 11 = manjak %
            ElseIf tip = "Kupac" Then
                IzSabirljive = Array(6, 8)             ' kg i vrednost; bez 7 = cena
            Else
                IzSabirljive = Array(3, 4)
            End If
        Case IZ_AMB:    IzSabirljive = Array(5, 6)     ' ulaz/izlaz (txt, ali promet)
        Case IZ_ISPL:   IzSabirljive = Array(2, 3, 4, 5)
        Case IZ_ZBIR
            If tip = "Vozac" Then
                IzSabirljive = Array(2, 3, 4)          ' bez 5 = manjak %
            Else
                IzSabirljive = Array(3, 4)             ' bez 5 = prosek
            End If
        Case IZ_CENA:   IzSabirljive = Array(2, 3)     ' bez 4 = prosecna cena
        Case IZ_MANJAK: IzSabirljive = Array(2, 3, 4)  ' bez % i proseka gajbi
        Case IZ_KART:   IzSabirljive = Array(4, 5)     ' promet; NIKAD saldo (6, 7)
        Case IZ_AMBK:   IzSabirljive = Array(4, 5)     ' ulaz/izlaz; NIKAD saldo (6)
        Case IZ_OTKL:   IzSabirljive = Array(6, 7)
        Case IZ_RANG:   IzSabirljive = Array(4)        ' iznos; rang broj nikad
        Case Else:      IzSabirljive = Array()
    End Select
End Function

' Zaglavlja stampe po listi -- isti opis kolona kao mreza (vidljive kolone),
' natpisi iz kataloga. Javno da bi ga stampa i test delili.
Public Function IzHeaderiZaListu(ByVal kljuc As String, ByVal tip As String, _
                                 Optional ByVal zbirni As Boolean = False) As Variant
    Dim kolone As Variant, i As Long, n As Long
    Dim res() As String
    kolone = IzKoloneZaListu(kljuc, tip, zbirni)
    ReDim res(0 To UBound(kolone))
    For i = 0 To UBound(kolone)
        If val(Split(CStr(kolone(i)), "|")(4)) < 4 Then
            res(n) = Poruka(Split(CStr(kolone(i)), "|")(0))
            n = n + 1
        End If
    Next i
    ReDim Preserve res(0 To n - 1)
    IzHeaderiZaListu = res
End Function

' "Stampaj izvestaj": tabelarni PDF AKTIVNE liste -- tacno ono sto operater
' vidi (cip + pretraga; PrintSpecDat presedan), sa naslovom iz KONTEKSTA
' SNIMKA (AUD-024: naslov opisuje prikazano, ne trenutno stanje polja).
' Filtriran izvod na papiru KAZE da je filtriran.
Private Sub StampajIzvestaj()
    Dim rez As Variant, redovi As Variant, n As Long
    Dim kolone As Variant, headers As Variant
    Dim dataS() As String
    Dim i As Long, j As Long, vidljivih As Long
    Dim naslov As String

    If Not SnimakPostoji() Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZ_PRVO_PRIKAZI"), True
        Exit Sub
    End If

    ' Isto oblikovanje kroz koje je prosla mreza -- poslednji (filter, q).
    rez = RedoviZaListu(mDiagFilter, mDiagQ)
    n = CLng(rez(2))
    If n = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_PRAZNO"), True
        Exit Sub
    End If
    kolone = rez(0)
    redovi = rez(1)
    headers = IzHeaderiZaListu(Scr_Lista(), mCtxTip, mCtxZbirni)
    vidljivih = UBound(headers) + 1

    ' +1 red: UKUPNO se racuna NAD STAMPANIM redovima (mreza ga ne nosi --
    ' sort bi ga pomerao; a zbir filtriranog izvoda mora da odgovara bas
    ' onome sto je na papiru). Sabiraju se SAMO kolone koje politika liste
    ' proglasi sabirljivim (IzSabirljive) -- tip kolone opisuje prikaz, ne
    ' aditivnost, pa bi generic suma sabirala prosecne cene i running salda.
    ReDim dataS(1 To n + 1, 1 To vidljivih)
    Dim kind As String, tot() As Double, imaTot() As Boolean
    Dim sabir As Variant, s As Long
    ReDim tot(1 To vidljivih)
    ReDim imaTot(1 To vidljivih)
    sabir = IzSabirljive(Scr_Lista(), mCtxTip, mCtxZbirni)
    If IsArray(sabir) Then
        For s = LBound(sabir) To UBound(sabir)
            If CLng(sabir(s)) >= 1 And CLng(sabir(s)) <= vidljivih Then _
                imaTot(CLng(sabir(s))) = True
        Next s
    End If
    For i = 1 To n
        For j = 1 To vidljivih
            dataS(i, j) = CelijaZaStampu(CStr(kolone(j - 1)), redovi(i, j))
            ' Sabirljive txt kolone (gajbe, prijemnica kg) nose formatiran
            ' string istog sistemskog oblika -- NzD ih cita; prazno je 0.
            If imaTot(j) Then tot(j) = tot(j) + NzD(redovi(i, j))
        Next j
    Next i
    dataS(n + 1, 1) = "UKUPNO"
    For j = 2 To vidljivih
        If imaTot(j) Then
            kind = Split(CStr(kolone(j - 1)), "|")(2)
            If kind = "kg" Then
                dataS(n + 1, j) = FmtKolicina(tot(j))
            ElseIf kind = "num" Then
                dataS(n + 1, j) = Format$(tot(j), "#,##0")
            ElseIf kind = "txt" Then
                dataS(n + 1, j) = Format$(tot(j), "#,##0.##")
            Else
                dataS(n + 1, j) = Format$(tot(j), "#,##0.00")
            End If
        End If
    Next j

    ' Kontekst (podnaslov) nosi entitet, period i AKTIVAN filter -- papir bez
    ' te napomene izgleda kao ceo izvestaj. Kartice nose i ZAVRSNI saldo:
    ' UKUPNO red stampe sabira samo promet (saldo nije aditivan), a "gde smo
    ' na kraju" je bas ono zbog cega se kartica stampa (smoke krug 4).
    naslov = TipLabela(mCtxTip) & ": " & mCtxEntNaziv & " (" & OpsegLabela() & ")"
    If (Scr_Lista() = IZ_KART Or Scr_Lista() = IZ_AMBK) And _
       Not IsEmpty(mZonaKartSaldo) Then
        naslov = naslov & "  " & ChrW(183) & "  " & _
                 Poruka("OTKUI_KPI_IZ_SALDO") & ": " & _
                 Format$(NzD(mZonaKartSaldo), "#,##0.00")
        If Scr_Lista() = IZ_KART And Not IsEmpty(mZonaKartSaldoAmb) Then
            naslov = naslov & "  " & ChrW(183) & "  " & _
                     Poruka("OTKUI_KPI_IZ_SALDOAMB") & ": " & _
                     Format$(NzD(mZonaKartSaldoAmb), "#,##0")
        End If
    End If
    If Len(mDiagQ) > 0 Then
        naslov = naslov & "  " & ChrW(183) & "  " & Poruka("OTKUI_IZ_PRETRAGA") & _
                 " " & mDiagQ
    End If
    If Len(mDiagFilter) > 0 And mDiagFilter <> "sve" Then
        naslov = naslov & "  " & ChrW(183) & "  " & Poruka("OTKUI_IZ_FILTER") & _
                 " " & mDiagFilter
    End If

    ' House stil kao i ostali dokumenti (zaglavlje firme + naslov + tabela;
    ' smoke krug 3: "svi dokumenti uskladjeni"). Kolone brojeva desno.
    Dim desno() As Boolean
    ReDim desno(0 To vidljivih - 1)
    For j = 1 To vidljivih
        Select Case Split(CStr(kolone(j - 1)), "|")(2)
            Case "kg", "rsd", "num", "rest", "date"
                desno(j - 1) = True
            Case "txt"
                desno(j - 1) = imaTot(j)   ' sabirljive txt kolone su brojevi
        End Select
    Next j

    PrintIzvestajHouse dataS, n + 1, vidljivih, _
                       UCase$(Poruka(NaslovKljucListe(Scr_Lista()))), _
                       naslov, headers, desno
End Sub

' Vrednost celije za stampu -- isto formatiranje kao mreza (CelijaTekst
' pravila), ali u string matrici koju PrintIzvestaj upisuje u sheet.
Private Function CelijaZaStampu(ByVal spec As String, ByVal v As Variant) As String
    Dim kind As String
    kind = Split(spec, "|")(2)
    Select Case kind
        Case "date"
            If IsNumeric(v) Then
                If CDbl(v) >= 1 Then CelijaZaStampu = Format$(CDate(CDbl(v)), "d.m.yyyy")
            End If
        Case "kg":  CelijaZaStampu = FmtKolicina(NzD(v))
        Case "num": CelijaZaStampu = Format$(NzD(v), "#,##0")
        Case "rsd": CelijaZaStampu = Format$(NzD(v), "#,##0.00")
        Case "rest"
            If NzD(v) <> 0 Then CelijaZaStampu = Format$(NzD(v), "#,##0.00")
        Case Else
            CelijaZaStampu = NzS(v)
    End Select
End Function

Private Function NaslovKljucListe(ByVal kljuc As String) As String
    NaslovKljucListe = "OTKUI_GRID_TITLE_IZ_" & IzSufiks(kljuc)
End Function

Private Function IzSufiks(ByVal kljuc As String) As String
    Select Case kljuc
        Case IZ_SALDO:  IzSufiks = "SALDO"
        Case IZ_ROBA:   IzSufiks = "ROBA"
        Case IZ_AMB:    IzSufiks = "AMB"
        Case IZ_ISPL:   IzSufiks = "ISPL"
        Case IZ_ZBIR:   IzSufiks = "ZBIR"
        Case IZ_CENA:   IzSufiks = "CENA"
        Case IZ_MANJAK: IzSufiks = "MANJAK"
        Case IZ_KART:   IzSufiks = "KART"
        Case IZ_AMBK:   IzSufiks = "AMBK"
        Case IZ_OTKL:   IzSufiks = "OTKL"
    End Select
End Function

' "Stampaj karticu (PDF)": po sablonu, za kooperanta i period IZ SNIMKA --
' tab-aware kao legacy btnStampajKarticu (finansijska kartica ili kartica
' ambalaze). Postojece rutine, nepromenjene.
Private Sub StampajKarticu()
    Dim dOd As Date, dDo As Date
    If Not SnimakPostoji() Or mCtxTip <> "Kooperant" Or Len(mCtxId) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZ_PRVO_PRIKAZI"), True
        Exit Sub
    End If
    If Scr_Lista() <> IZ_KART And Scr_Lista() <> IZ_AMBK Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZ_PRVO_PRIKAZI"), True
        Exit Sub
    End If
    dOd = CDate(IIf(mCtxOd > 0, mCtxOd, IZ_DAT_MIN))
    dDo = CDate(IIf(mCtxDo > 0, mCtxDo, IZ_DAT_MAX))
    If Scr_Lista() = IZ_AMBK Then
        PrintKarticaAmbalazePDF mCtxId, dOd, dDo
    Else
        PrintKarticaPDF mCtxId, dOd, dDo
    End If
End Sub

' "Stampaj dokument" nad izabranim redom: ruta po listi i po tipu dokumenta.
' Red bez dokumenta ODBIJA porukom -- agregatni red bez radnje nije greska,
' radnja koja pogadja jeste.
Private Sub StampajDokumentReda(ByVal red As Long)
    Dim ref As String, dokTip As String, dokID As String, tipAmb As String

    Select Case Scr_Lista()
        Case IZ_OTKL
            ref = NzS(modOtkupUI.GridCell(red, 8))
            If Left$(ref, 4) = "OTK|" Then
                ReprintOtkupniListByOtkupID Mid$(ref, 5)
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
            End If
        Case IZ_KART
            ref = NzS(modOtkupUI.GridCell(red, 8))
            If Left$(ref, 4) = "OTK|" Then
                ReprintOtkupniListByOtkupID Mid$(ref, 5)
            ElseIf Len(ref) > 0 Then
                ' NOV / MAG / AMB red kartice nema dokument za stampu iz ovog
                ' pregleda -- isto kao legacy Case Else.
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_STAMPA_NEDOSTUPNA"), True
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
            End If
        Case IZ_ROBA
            If mCtxTip = "Kupac" Then
                ref = NzS(modOtkupUI.GridCell(red, 9))
                If Left$(ref, 4) = "PRJ|" Then
                    PrintPrijemnica Mid$(ref, 5)
                Else
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
                End If
            Else
                ref = NzS(modOtkupUI.GridCell(red, 12))
                If Left$(ref, 4) = "OTP|" Then
                    OutputOtpremnicaPDF Mid$(ref, 5)
                Else
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
                End If
            End If
        Case IZ_AMB
            dokTip = NzS(modOtkupUI.GridCell(red, 7))
            dokID = NzS(modOtkupUI.GridCell(red, 8))
            tipAmb = NzS(modOtkupUI.GridCell(red, 3))
            If Len(dokID) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
                Exit Sub
            End If
            Select Case dokTip
                Case DOK_TIP_PRIJEMNICA
                    PrintPrijemnica dokID
                Case DOK_TIP_OTKUP
                    ReprintOtkupniListByOtkupID dokID
                Case DOK_TIP_OTPREMNICA
                    OutputOtpremnicaPDF dokID
                Case DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, _
                     DOK_TIP_OM_IZLAZ_FIRMA, DOK_TIP_OM_ULAZ_FIRMA
                    ' Revers: rekonstrukcija iz dve noge ledgera -- racun
                    ' izdvojen u modIzvestaj.StampajReversAmbalaze (AUD-012:
                    ' tip ambalaze IZABRANOG reda je deo kljuca).
                    StampajReversAmbalaze dokID, dokTip, tipAmb
                Case Else
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_STAMPA_NEDOSTUPNA"), True
            End Select
        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
    End Select
End Sub

'=====================================================================
' DIJAGNOSTIKA
'
' Alt+F8 -> Diag_IzRedovi, pa Ctrl+G. Ne menja nista. Isti razlog kao
' Diag_BnRedovi (N7): "filter ne radi" se ne razresava citanjem -- ispisuje
' se sta je ljuska poslednje trazila, sta ekran vraca, koliko je snimak
' stvarno punjen i pod kojim kljucem.
'=====================================================================
Public Sub Diag_IzRedovi()
    Dim d As Variant, kolone As Variant, redovi As Variant, i As Long, k As Long, n As Long
    On Error Resume Next

    Debug.Print "--- Diag_IzRedovi (" & SCRIZ_BUILD & ") ---"
    Debug.Print "  POSLEDNJI POZIV: filter=[" & mDiagFilter & "] q=[" & mDiagQ & _
                "] vraceno redova=" & CStr(mDiagN)
    Debug.Print "  SNIMAK: kljuc=[" & mSnimakKljuc & "] ok=" & SnimakPostoji() & _
                " punjenja=" & CStr(mSnimakPunjenja)

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
' Zona se u testu ne crta (forma se ne prikazuje), pa se kontekst ne moze
' procitati iz kontrola. Ista kapija kao Scr_*Test na ostalim ekranima:
' seam koji MENJA stanje van test-rezima ne radi nista.
'=====================================================================
Public Sub Scr_IzTestSet(ByVal lista As String, ByVal tip As String, _
                         ByVal zbirni As Boolean, ByVal entitetID As String, _
                         ByVal odSerijski As Double, ByVal doSerijski As Double)
    If Not IsTestMode() Then Exit Sub
    If Len(lista) > 0 Then mLista = lista
    mTip = tip
    mZbirni = zbirni
    mTestId = entitetID
    mTestOd = odSerijski
    mTestDo = doSerijski
End Sub

Public Function Scr_IzSnimakPunjenjaTest() As Long
    If Not IsTestMode() Then Exit Function
    Scr_IzSnimakPunjenjaTest = mSnimakPunjenja
End Function

Public Function Scr_IzSnimakKljucTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_IzSnimakKljucTest = mSnimakKljuc
End Function

' Brojke zone poslednjeg oblikovanja -- za slaganje "specijalni redovi su
' izdvojeni, ne izgubljeni" (OM avans, agro, kontrolni redovi isplate).
' Empty = red nije postojao u povratku.
Public Function Scr_IzZonaBrojkaTest(ByVal koja As String) As Variant
    If Not IsTestMode() Then Exit Function
    Select Case koja
        Case "omavans":      Scr_IzZonaBrojkaTest = mZonaOmAvans
        Case "agro":         Scr_IzZonaBrojkaTest = mZonaAgro
        Case "primljeno":    Scr_IzZonaBrojkaTest = mZonaIsplPrimljeno
        Case "podeljeno":    Scr_IzZonaBrojkaTest = mZonaIsplPodeljeno
        Case "kod":          Scr_IzZonaBrojkaTest = mZonaIsplKod
        Case "kartsaldo":    Scr_IzZonaBrojkaTest = mZonaKartSaldo
        Case "kartsaldoamb": Scr_IzZonaBrojkaTest = mZonaKartSaldoAmb
    End Select
End Function

Public Function Scr_IzHintKljucTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_IzHintKljucTest = mHintKljuc
End Function

' Naziv entiteta STVARNO prikazanog konteksta (za hint/stampu) -- test
' meri da zbirna ambalaza ne lazira "Svi" (smoke krug 9).
Public Function Scr_IzCtxNazivTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_IzCtxNazivTest = mCtxEntNaziv
End Function

Public Sub Scr_IzTestReset()
    If Not IsTestMode() Then Exit Sub
    mLista = IZ_SALDO
    mTip = "OM"
    mZbirni = False
    mTestId = ""
    mTestOd = 0
    mTestDo = 0
    mSnimakPunjenja = 0
    mComboTip = ""
    ResetZonskeBrojke
    mHintKljuc = ""
    Scr_ResetCache
    mSnimakKljuc = ""
    mCtxTip = ""
    mCtxId = ""
End Sub
