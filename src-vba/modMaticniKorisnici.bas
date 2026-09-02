Attribute VB_Name = "modMaticniKorisnici"
'=====================================================================
' modMaticniKorisnici - korisnici i prava pristupa (korak M4).
'
' Poslednja sekcija menija "Maticni podaci" koja je do sada ostajala u formi.
' Ostala je namerno: tblKorisnici nosi PIN i matricu prava po oblasti, pa je
' cekala svoj ekran.
'
' ODAKLE DOLAZI: frmStammdaten grane Case "Korisnici" u btnDodaj_Click i
' btnIzmeni_Click, plus SetupKorisnici / BuildKorisniciOblasti / OblComboVal /
' OblastValueFromRow / OblastiCsvFromRow / NormalizeUloga.
'
' RECNIK KOLONE JE "DA" / "NE", i to nije stvar ukusa nego zatecena cinjenica
' na CETIRI mesta: modSetup upisuje "DA", modAuth deaktiviranim smatra samo
' "NE" (modAuth.bas:87), a prava se i pisu i citaju kao "DA"/"NE"
' (OblComboVal, OblastValueFromRow). Zato ovaj modul pise iskljucivo DA/NE.
'
' NALAZ KOJI SE OVIM ZATVARA: genericko dugme "Deaktiviraj" je u tu kolonu
' upisivalo "Aktivan"/"Neaktivan" (STATUS_AKTIVAN / STATUS_NEAKTIVAN). Kako
' modAuth prepoznaje samo "NE", deaktivacija korisnika tim dugmetom NIJE
' sprecavala prijavu -- u listi je pisalo "Neaktivan", a login je prolazio.
' V. UI_MIGRACIJA_KATALOG 26.15 i 26.18.
'
' STA MODUL NE RADI: ne dira modAuth. Citac ostaje kakav jeste -- zapisi koji
' vec nose "Neaktivan" i dalje se citaju kao aktivni, tacno kao do sada. Menja
' se samo sta se UPISUJE; nijedan korisnik ovom izmenom ne gubi pristup.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATKOR_BUILD As String = "v6-ui-193"

Public Const KOR_DA As String = "DA"
Public Const KOR_NE As String = "NE"

' Prefiks pod kojim pozivalac moze da preda vrednosti prava u recniku polja:
' "obl:Otkup" -> "DA". Forma ih salje (ima 12 combo-a), ekran ne salje nista i
' prava mu ostaju netaknuta -- njima upravlja svoja lista.
Public Const KOR_OBL_PREFIKS As String = "obl:"

' Kolona liste prava koja nosi NAZIV OBLASTI. Jedan broj, deljen izmedju opisa
' kolona i radnje -- da se indeks ne moze razici.
Public Const KOR_COL_OBLAST As Long = 4

Private Const SRC As String = "modMaticniKorisnici"

' Brojke iz POSLEDNJEG citanja -- isti obrazac koji modMaticniIzvor drzi za
' ostale sekcije. Racunaju se u istom prolazu kao redovi, pa zona i mreza ne
' mogu da se razidju.
Private mUkupno As Long
Private mAktivnih As Long
Private mNeaktivnih As Long

'-------------------------------------------------------------- POLJA
' Polja editora. Prava NISU ovde: dvanaest combo-a ne staje u bazen od sest, a
' i ne pripadaju istoj radnji -- menjaju se po jedno, iz svoje liste.
Public Function KorPolja() As Variant
    KorPolja = Array( _
        "korime|OTKUI_HDK_USERNAME|txt|1|" & COL_KOR_USERNAME & "|", _
        "ime|OTKUI_HDK_IME|txt|0|" & COL_KOR_IME & "|", _
        "pin|OTKUI_HDK_PIN|txt|0|" & COL_KOR_PIN & "|", _
        "uloga|OTKUI_HDK_ULOGA|cmb|1|" & COL_KOR_ULOGA & "|@uloge", _
        "aktivan|OTKUI_HDK_AKTIVAN|cmb|1|" & COL_KOR_AKTIVAN & "|@dane_kor", _
        "stanica|OTKUI_HDM_STANICA|cmb|0|" & COL_KOR_STANICA & "|@stanice")
End Function

Public Function KorComboStavke(ByVal izvor As String) As Variant
    Select Case izvor
        Case "@uloge":    KorComboStavke = Array(ULOGA_ADMIN, ULOGA_KORISNIK)
        Case "@dane_kor": KorComboStavke = Array(KOR_DA, KOR_NE)
    End Select
End Function

'------------------------------------------------------------- REDOVI
Public Function KorKolone() As Variant
    KorKolone = Array( _
        "OTKUI_HDM_ID|" & COL_KOR_ID & "|txt|84|1", _
        "OTKUI_HDK_USERNAME|" & COL_KOR_USERNAME & "|txt|130|1", _
        "OTKUI_HDK_IME|" & COL_KOR_IME & "|part|0|1", _
        "OTKUI_HDK_ULOGA|" & COL_KOR_ULOGA & "|txt|84|1", _
        "OTKUI_HDK_AKTIVAN|" & COL_KOR_AKTIVAN & "|txt|72|1", _
        "OTKUI_HDK_PRAVA|@prava|txt|0|2")
End Function

Public Function KorRedovi(ByVal filter As String, ByVal q As String) As Variant
    Dim data As Variant, outA() As Variant, i As Long, j As Long, n As Long
    Dim cID As Long, cU As Long, cIme As Long, cUl As Long, cAk As Long
    Dim hay As String, akt As Boolean, prava As String
    On Error GoTo EH

    mUkupno = 0: mAktivnih = 0: mNeaktivnih = 0
    data = GetTableData(TBL_KORISNICI)
    If IsEmpty(data) Then
        KorRedovi = Array(KorKolone(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    cID = GetColumnIndex(TBL_KORISNICI, COL_KOR_ID)
    cU = GetColumnIndex(TBL_KORISNICI, COL_KOR_USERNAME)
    cIme = GetColumnIndex(TBL_KORISNICI, COL_KOR_IME)
    cUl = GetColumnIndex(TBL_KORISNICI, COL_KOR_ULOGA)
    cAk = GetColumnIndex(TBL_KORISNICI, COL_KOR_AKTIVAN)
    If cID = 0 Then
        KorRedovi = Array(KorKolone(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To UBound(data, 1), 1 To 6)
    For i = 1 To UBound(data, 1)
        If Len(Trim$(NzToText(data(i, cID)))) = 0 Then GoTo Sledeci

        ' Aktivan je SVE sto nije "NE" -- isti kriterijum koji modAuth primenjuje
        ' pri prijavi. Zapis sa zatecenim "Neaktivan" se time i ovde prikazuje
        ' onako kako se stvarno ponasa: kao aktivan.
        akt = True
        If cAk > 0 Then akt = (UCase$(Trim$(NzToText(data(i, cAk)))) <> KOR_NE)
        ' Brojke se pune PRE cipa: "ukupno" broji sve naloge bez obzira na
        ' izabrani cip, inace bi brojka menjala znacenje sa cipom.
        mUkupno = mUkupno + 1
        If akt Then mAktivnih = mAktivnih + 1 Else mNeaktivnih = mNeaktivnih + 1
        If filter = modMaticniIzvor.MAT_CIP_AKT And Not akt Then GoTo Sledeci
        If filter = modMaticniIzvor.MAT_CIP_NEAKT And akt Then GoTo Sledeci

        prava = PravaOpis(data, i, JeAdmin(NzToText(data(i, cUl))))
        n = n + 1
        outA(n, 1) = NzToText(data(i, cID))
        If cU > 0 Then outA(n, 2) = NzToText(data(i, cU))
        If cIme > 0 Then outA(n, 3) = NzToText(data(i, cIme))
        If cUl > 0 Then outA(n, 4) = NzToText(data(i, cUl))
        If cAk > 0 Then outA(n, 5) = NzToText(data(i, cAk))
        outA(n, 6) = prava

        If Len(q) > 0 Then
            hay = ""
            For j = 1 To 6
                hay = hay & "|" & NzToText(outA(n, j))
            Next j
            If InStr(1, hay, q, vbTextCompare) = 0 Then n = n - 1
        End If
Sledeci:
    Next i

    If n = 0 Then
        KorRedovi = Array(KorKolone(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    KorRedovi = Array(KorKolone(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, SRC & ".KorRedovi", Err.description
End Function

'-------------------------------------------------------------- PRAVA
' Cetvrta kolona je NEVIDLJIVA i nosi naziv oblasti (naziv kolone u tabeli).
' Radnja bira po njoj, ne po rednom broju -- prikaz je lokalizovan naziv.
Public Function KorPravaKolone() As Variant
    KorPravaKolone = Array( _
        "OTKUI_HDK_OBLAST|@oblast|part|0|1", _
        "OTKUI_HDK_PRAVO|@pravo|txt|92|1", _
        "OTKUI_HDK_ODAKLE|@odakle|txt|150|2", _
        "OTKUI_HDK_OBLAST|@kljuc|txt|0|4")
End Function

' Prava IZABRANOG korisnika: jedan red po oblasti. Bez izabranog korisnika
' lista je prazna -- prava bez korisnika nisu podatak.
Public Function KorPravaRedovi(ByVal korID As String, ByVal q As String) As Variant
    Dim obl As Variant, outA() As Variant, n As Long, red As Long
    Dim data As Variant, adm As Boolean, v As String, hay As String
    On Error GoTo EH

    mUkupno = 0: mAktivnih = 0: mNeaktivnih = 0
    red = RedPoID(korID)
    If red = 0 Then
        KorPravaRedovi = Array(KorPravaKolone(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    data = GetTableData(TBL_KORISNICI)
    adm = JeAdmin(PoljeReda(data, red, COL_KOR_ULOGA))

    ReDim outA(1 To UBound(modAuth.OblastiList()) + 1, 1 To KOR_COL_OBLAST)
    For Each obl In modAuth.OblastiList()
        v = VrednostOblasti(data, red, CStr(obl))
        ' Admin ima sve -- to nije zapis nego pravilo (OblComboVal), pa se i
        ' prikazuje kao pravilo, a ne kao vrednost koja se moze menjati.
        If adm Then v = KOR_DA
        ' Za listu prava brojke znace: koliko oblasti ukupno, koliko sa pravom,
        ' koliko bez. Broji se pre pretrage, iz istog razloga kao gore.
        mUkupno = mUkupno + 1
        If v = KOR_DA Then mAktivnih = mAktivnih + 1 Else mNeaktivnih = mNeaktivnih + 1
        hay = KorOblastNaziv(CStr(obl)) & "|" & v
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeca
        End If
        n = n + 1
        outA(n, 1) = KorOblastNaziv(CStr(obl))
        outA(n, 2) = IIf(v = KOR_DA, Poruka("OTKUI_KOR_IMA"), Poruka("OTKUI_KOR_NEMA"))
        outA(n, 3) = IIf(adm, Poruka("OTKUI_KOR_JER_ADMIN"), Poruka("OTKUI_KOR_POJEDINACNO"))
        outA(n, KOR_COL_OBLAST) = CStr(obl)
Sledeca:
    Next obl

    If n = 0 Then
        KorPravaRedovi = Array(KorPravaKolone(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    KorPravaRedovi = Array(KorPravaKolone(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, SRC & ".KorPravaRedovi", Err.description
End Function

' Lokalizovan naziv oblasti. Kljuc kataloga se izvodi iz naziva kolone, pa
' dodavanje oblasti u modAuth.OblastiList trazi samo jedan red u modPoruke.
Public Function KorOblastNaziv(ByVal oblast As String) As String
    KorOblastNaziv = Poruka("OTKUI_OBL_" & UCase$(oblast))
    If Len(KorOblastNaziv) = 0 Then KorOblastNaziv = oblast
End Function

'-------------------------------------------------------------- RADNJE
' Nov korisnik. Vraca "" kad je proslo.
Public Function KorDodaj(ByVal polja As Object, ByRef noviID As String) As String
    Dim greska As String, tx As clsTransaction, red As Long, adm As Boolean
    On Error GoTo EH
    noviID = ""
    greska = Proveri(polja, "")
    If Len(greska) > 0 Then
        KorDodaj = greska
        Exit Function
    End If
    ' PIN je obavezan SAMO pri unosu -- pri izmeni prazno znaci "ostaje isti".
    If Len(Vred(polja, "pin")) = 0 Then
        polja(modMaticniUnos.MAT_FOKUS) = "pin"
        KorDodaj = Poruka("MATK_ERR_PIN")
        Exit Function
    End If

    noviID = GetNextID(TBL_KORISNICI, COL_KOR_ID, "KOR-")
    adm = JeAdmin(Vred(polja, "uloga"))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_KORISNICI

    red = DodajPrazanRed()
    If red = 0 Then Err.Raise vbObjectError + 9700, SRC, "AppendRow nije uspeo."

    RequireUpdateCell TBL_KORISNICI, red, COL_KOR_ID, noviID, SRC
    UpisiPolja red, polja, True
    RequireUpdateCell TBL_KORISNICI, red, COL_KOR_CREATED, _
                      Format$(Now, "yyyy-mm-dd hh:nn:ss"), SRC
    UpisiPrava red, polja, adm, True, False

    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    KorDodaj = Odustani(tx, "KorDodaj")
End Function

' Red, ne ID: pozivalac je modMaticniUnos, koji redom barata za sve sekcije.
' Sopstveni ID se cita IZ REDA -- inace bi provera jedinstvenosti korisnickog
' imena pala na sopstvenom zapisu pri svakoj izmeni.
Public Function KorIzmeni(ByVal red As Long, ByVal polja As Object) As String
    Dim greska As String, tx As clsTransaction, adm As Boolean, korID As String
    Dim bioAdmin As Boolean
    On Error GoTo EH
    If red < 1 Then
        KorIzmeni = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If
    korID = PoljeReda(GetTableData(TBL_KORISNICI), red, COL_KOR_ID)
    ' Uloga PRE izmene -- treba za pravilo o spustanju sa admina (v. UpisiPrava).
    bioAdmin = JeAdmin(PoljeReda(GetTableData(TBL_KORISNICI), red, COL_KOR_ULOGA))
    greska = Proveri(polja, korID)
    If Len(greska) > 0 Then
        KorIzmeni = greska
        Exit Function
    End If
    adm = JeAdmin(Vred(polja, "uloga"))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_KORISNICI

    UpisiPolja red, polja, False
    UpisiPrava red, polja, adm, False, bioAdmin

    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    KorIzmeni = Odustani(tx, "KorIzmeni")
End Function

' Obrce aktivnost -- i to iskljucivo u recniku DA/NE koji modAuth cita.
Public Function KorPromeniStatus(ByVal red As Long, ByRef novi As String) As String
    Dim tx As clsTransaction, cur As String
    On Error GoTo EH
    novi = ""
    If red < 1 Then
        KorPromeniStatus = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If
    cur = UCase$(PoljeReda(GetTableData(TBL_KORISNICI), red, COL_KOR_AKTIVAN))
    novi = IIf(cur = KOR_NE, KOR_DA, KOR_NE)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_KORISNICI
    RequireUpdateCell TBL_KORISNICI, red, COL_KOR_AKTIVAN, novi, SRC
    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    KorPromeniStatus = Odustani(tx, "KorPromeniStatus")
End Function

' Obrce jedno pravo. Adminu se odbija: njegova prava nisu zapis nego posledica
' uloge, pa bi promena bila prividna -- sledeci upis bi je vratio na DA.
Public Function KorPromeniPravo(ByVal korID As String, ByVal oblast As String, _
                                ByRef novo As String) As String
    Dim red As Long, tx As clsTransaction, data As Variant
    On Error GoTo EH
    novo = ""
    If Len(Trim$(oblast)) = 0 Then
        KorPromeniPravo = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If
    red = RedPoID(korID)
    If red = 0 Then
        KorPromeniPravo = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If
    data = GetTableData(TBL_KORISNICI)
    If JeAdmin(PoljeReda(data, red, COL_KOR_ULOGA)) Then
        KorPromeniPravo = Poruka("MATK_ERR_ADMIN_SVE")
        Exit Function
    End If
    If GetColumnIndex(TBL_KORISNICI, oblast) = 0 Then
        KorPromeniPravo = Poruka("MATK_ERR_NEMA_OBLASTI") & " " & oblast
        Exit Function
    End If

    novo = IIf(VrednostOblasti(data, red, oblast) = KOR_DA, KOR_NE, KOR_DA)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_KORISNICI
    RequireUpdateCell TBL_KORISNICI, red, oblast, novo, SRC
    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    KorPromeniPravo = Odustani(tx, "KorPromeniPravo")
End Function

'------------------------------------------------------------ CITANJE
Public Function KorVrednostiReda(ByVal red As Long) As Object
    Dim d As Object, data As Variant, r As Variant, kol As String
    Set d = CreateObject("Scripting.Dictionary")
    Set KorVrednostiReda = d
    On Error GoTo EH
    If red < 1 Then Exit Function
    data = GetTableData(TBL_KORISNICI)
    For Each r In KorPolja()
        kol = modMaticniIzvor.PoljeF(CStr(r), 4)
        ' PIN se NE vraca u editor: hes se ne prikazuje, a prazno polje pri
        ' izmeni ionako znaci "ostaje isti".
        If modMaticniIzvor.PoljeF(CStr(r), 0) = "pin" Then
            d("pin") = ""
        ElseIf kol = COL_KOR_STANICA Then
            d("stanica") = NazivStanice(PoljeReda(data, red, kol))
        Else
            d(modMaticniIzvor.PoljeF(CStr(r), 0)) = PoljeReda(data, red, kol)
        End If
    Next r
    Exit Function
EH:
    LogErr SRC & ".KorVrednostiReda"
End Function

Public Function KorRedPoID(ByVal korID As String) As Long
    KorRedPoID = RedPoID(korID)
End Function

' Citljivo ime korisnika za natpise. Vraca korisnicko ime, a ID samo kad ga
' nema -- napomena "Prava: KOR-004" ne kaze operateru cija su.
Public Function KorNaziv(ByVal korID As String) As String
    Dim red As Long
    KorNaziv = Trim$(korID)
    red = RedPoID(korID)
    If red = 0 Then Exit Function
    If Len(PoljeReda(GetTableData(TBL_KORISNICI), red, COL_KOR_USERNAME)) > 0 Then _
        KorNaziv = PoljeReda(GetTableData(TBL_KORISNICI), red, COL_KOR_USERNAME)
End Function

'-------------------------------------------------------------- BROJKE
' Brojke iz poslednjeg KorRedovi / KorPravaRedovi. Cita ih modMaticniIzvor i
' prosledjuje zoni -- jedno mesto koje ih racuna, jedno koje ih prikazuje.
Public Function KorUkupno() As Long
    KorUkupno = mUkupno
End Function

Public Function KorAktivnih() As Long
    KorAktivnih = mAktivnih
End Function

Public Function KorNeaktivnih() As Long
    KorNeaktivnih = mNeaktivnih
End Function

Public Function JeAdmin(ByVal uloga As String) As Boolean
    JeAdmin = (StrComp(Trim$(uloga), ULOGA_ADMIN, vbTextCompare) = 0)
End Function

'--------------------------------------------------------- UNUTRASNJE
' Korisnicko ime je obavezno i JEDINSTVENO. Pri izmeni se sopstveni red
' preskace -- inace bi svaka izmena pala na "vec postoji".
Private Function Proveri(ByVal polja As Object, ByVal sopstveniID As String) As String
    Dim u As String, nadjen As Variant
    u = Vred(polja, "korime")
    If Len(u) = 0 Then
        polja(modMaticniUnos.MAT_FOKUS) = "korime"
        Proveri = Poruka("MATK_ERR_USERNAME")
        Exit Function
    End If
    On Error Resume Next
    nadjen = LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_ID)
    Err.Clear
    On Error GoTo 0
    If Not IsEmpty(nadjen) Then
        If Len(Trim$(CStr(nadjen))) > 0 Then
            If StrComp(Trim$(CStr(nadjen)), Trim$(sopstveniID), vbTextCompare) <> 0 Then
                polja(modMaticniUnos.MAT_FOKUS) = "korime"
                Proveri = Poruka("MATK_ERR_POSTOJI") & " " & u
            End If
        End If
    End If
End Function

Private Sub UpisiPolja(ByVal red As Long, ByVal polja As Object, ByVal jeUnos As Boolean)
    Dim r As Variant, k As String, kol As String, v As String
    For Each r In KorPolja()
        k = modMaticniIzvor.PoljeF(CStr(r), 0)
        kol = modMaticniIzvor.PoljeF(CStr(r), 4)
        v = Vred(polja, k)
        Select Case k
            Case "pin"
                ' Prazan PIN pri IZMENI znaci "ostaje isti" -- zateceno pravilo
                ' forme, i jedini nacin da se korisnik menja bez otkucavanja PIN-a.
                If Len(v) > 0 Then _
                    RequireUpdateCell TBL_KORISNICI, red, kol, modAuth.PreparePin(v), SRC
            Case "uloga"
                RequireUpdateCell TBL_KORISNICI, red, kol, Uloga(v), SRC
            Case "aktivan"
                RequireUpdateCell TBL_KORISNICI, red, kol, DaNe(v, jeUnos), SRC
            Case "stanica"
                RequireUpdateCell TBL_KORISNICI, red, kol, StanicaID(v), SRC
            Case Else
                RequireUpdateCell TBL_KORISNICI, red, kol, v, SRC
        End Select
    Next r
End Sub

' Prava. Admin dobija DA na svemu -- to je pravilo, ne izbor. Za obicnog
' korisnika se pisu SAMO one oblasti koje je pozivalac stvarno poslao
' (prefiks "obl:"); ekran ne salje nista, pa mu prava ostaju netaknuta i menjaju
' se iz svoje liste. Pri UNOSU se, kad nista nije poslato, upisuje "NE" na sve --
' isto sto KorisniciSetDefaults radi u formi.
'
' SPUSTANJE SA ADMINA je treci slucaj i namerno se ponasa kao unos. Adminov red
' nosi "DA" na svim oblastima zato sto je to POSLEDICA uloge, a ne izbor koji je
' neko napravio. Kad uloga padne na Korisnik, prenos tih "DA" bi dao pun pristup
' koji niko nije dodelio -- a modAuth ga cita bukvalno. Zato se, kad prava nisu
' izricito poslata, sve gasi i dodeljuje ponovo iz liste prava.
Private Sub UpisiPrava(ByVal red As Long, ByVal polja As Object, ByVal adm As Boolean, _
                       ByVal jeUnos As Boolean, ByVal bioAdmin As Boolean)
    Dim obl As Variant, k As String, imaPoslatih As Boolean

    Dim gasi As Boolean

    For Each obl In modAuth.OblastiList()
        If polja.Exists(KOR_OBL_PREFIKS & CStr(obl)) Then imaPoslatih = True
    Next obl
    gasi = GasiSvaPrava(adm, imaPoslatih, jeUnos, bioAdmin)

    For Each obl In modAuth.OblastiList()
        If GetColumnIndex(TBL_KORISNICI, CStr(obl)) > 0 Then
            k = KOR_OBL_PREFIKS & CStr(obl)
            If adm Then
                RequireUpdateCell TBL_KORISNICI, red, CStr(obl), KOR_DA, SRC
            ElseIf polja.Exists(k) Then
                RequireUpdateCell TBL_KORISNICI, red, CStr(obl), _
                                  DaNe(CStr(polja(k)), False), SRC
            ElseIf gasi Then
                RequireUpdateCell TBL_KORISNICI, red, CStr(obl), KOR_NE, SRC
            End If
        End If
    Next obl
End Sub

' Da li se sva prava gase pa dodeljuju ponovo. Izdvojeno u funkciju zato sto je
' PRAVILO, a ne uslov: nov korisnik krece bez prava, a onaj kome je uloga pala sa
' admina krece bez prava iz istog razloga -- njegovih dvanaest "DA" nije dodelio
' niko, nego uloga koje vise nema. Prava koja su izricito poslata (forma salje
' svih dvanaest combo-a) ne diraju se; admin ih dobija po pravilu.
Private Function GasiSvaPrava(ByVal adm As Boolean, ByVal imaPoslatih As Boolean, _
                              ByVal jeUnos As Boolean, ByVal bioAdmin As Boolean) As Boolean
    If adm Or imaPoslatih Then Exit Function
    GasiSvaPrava = (jeUnos Or bioAdmin)
End Function

' "DA" samo za tacno "DA"; sve ostalo je "NE". Pri UNOSU prazno znaci DA --
' nov korisnik je podrazumevano aktivan, kao u KorisniciSetDefaults.
Private Function DaNe(ByVal v As String, ByVal praznoJeDa As Boolean) As String
    If Len(Trim$(v)) = 0 Then
        DaNe = IIf(praznoJeDa, KOR_DA, KOR_NE)
    ElseIf StrComp(Trim$(v), KOR_DA, vbTextCompare) = 0 Then
        DaNe = KOR_DA
    Else
        DaNe = KOR_NE
    End If
End Function

Private Function Uloga(ByVal v As String) As String
    Uloga = IIf(JeAdmin(v), ULOGA_ADMIN, ULOGA_KORISNIK)
End Function

' Combo stanice nosi "Naziv (ST-xxx)" ili goli ID; oba oblika daju ID.
Private Function StanicaID(ByVal v As String) As String
    Dim id As String
    StanicaID = Trim$(v)
    If Len(StanicaID) = 0 Then Exit Function
    id = ExtractIDFromDisplay(StanicaID)
    If Len(id) = 0 Or InStr(1, id, "ST-", vbTextCompare) = 0 Then _
        id = Trim$(CStr(LookupValue(TBL_STANICE, "Naziv", StanicaID, "StanicaID")))
    If Len(id) > 0 Then StanicaID = id
End Function

Private Function NazivStanice(ByVal id As String) As String
    Dim nm As String
    NazivStanice = Trim$(id)
    If Len(NazivStanice) = 0 Then Exit Function
    nm = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", NazivStanice, "Naziv")))
    If Len(nm) > 0 Then NazivStanice = nm & " (" & id & ")"
End Function

Private Function VrednostOblasti(ByVal data As Variant, ByVal red As Long, _
                                 ByVal oblast As String) As String
    If StrComp(PoljeReda(data, red, oblast), KOR_DA, vbTextCompare) = 0 Then
        VrednostOblasti = KOR_DA
    Else
        VrednostOblasti = KOR_NE
    End If
End Function

' Kratak opis prava za listu korisnika. Admin dobija jedan pojam umesto spiska
' od dvanaest -- isto sto legacy pise kao "SVE (admin)".
Private Function PravaOpis(ByRef data As Variant, ByVal red As Long, _
                           ByVal adm As Boolean) As String
    Dim obl As Variant, res As String
    If adm Then
        PravaOpis = Poruka("OTKUI_KOR_SVE_ADMIN")
        Exit Function
    End If
    For Each obl In modAuth.OblastiList()
        If VrednostOblasti(data, red, CStr(obl)) = KOR_DA Then
            If Len(res) > 0 Then res = res & ", "
            res = res & KorOblastNaziv(CStr(obl))
        End If
    Next obl
    PravaOpis = res
End Function

Private Function PoljeReda(ByRef data As Variant, ByVal red As Long, _
                           ByVal kolona As String) As String
    Dim c As Long
    On Error Resume Next
    If IsEmpty(data) Then Exit Function
    c = GetColumnIndex(TBL_KORISNICI, kolona)
    If c > 0 Then PoljeReda = Trim$(NzToText(data(red, c)))
    Err.Clear
End Function

Private Function RedPoID(ByVal korID As String) As Long
    Dim data As Variant, c As Long, i As Long
    On Error GoTo EH
    korID = Trim$(korID)
    If Len(korID) = 0 Then Exit Function
    c = GetColumnIndex(TBL_KORISNICI, COL_KOR_ID)
    If c = 0 Then Exit Function
    data = GetTableData(TBL_KORISNICI)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, c))), korID, vbTextCompare) = 0 Then
            RedPoID = i
            Exit Function
        End If
    Next i
    Exit Function
EH:
    LogErr SRC & ".RedPoID"
End Function

Private Function DodajPrazanRed() As Long
    Dim prazno() As Variant
    ReDim prazno(1 To GetTable(TBL_KORISNICI).ListColumns.count)
    DodajPrazanRed = AppendRow(TBL_KORISNICI, prazno)
End Function

Private Function Vred(ByVal polja As Object, ByVal k As String) As String
    On Error Resume Next
    If polja Is Nothing Then Exit Function
    If polja.Exists(k) Then Vred = Trim$(CStr(polja(k)))
    Err.Clear
End Function

Private Function Odustani(ByRef tx As clsTransaction, ByVal gde As String) As String
    Dim errDesc As String
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    LogError SRC & "." & gde, errDesc
    Err.Clear
    On Error GoTo 0
    Odustani = Poruka("MATU_ERR_UPIS") & " " & errDesc
End Function

'------------------------------------------------------------ TEST SEAM
Public Function KorProveriTest(ByVal polja As Object, ByVal sopstveniID As String) As String
    KorProveriTest = Proveri(polja, sopstveniID)
End Function

Public Function KorDaNeTest(ByVal v As String, ByVal praznoJeDa As Boolean) As String
    KorDaNeTest = DaNe(v, praznoJeDa)
End Function

Public Function KorGasiSvaPravaTest(ByVal adm As Boolean, ByVal imaPoslatih As Boolean, _
                                    ByVal jeUnos As Boolean, ByVal bioAdmin As Boolean) As Boolean
    KorGasiSvaPravaTest = GasiSvaPrava(adm, imaPoslatih, jeUnos, bioAdmin)
End Function
