Attribute VB_Name = "modSledljivost"
Option Explicit

' ============================================================
' modSledljivost -- JEZGRO GRAFA SLEDLJIVOSTI (S9-1, plan 14.64 i 14.65,
' docs/DOMEN/SLEDLJIVOST.md). SAMO CITANJE: ne pise nijednu tabelu.
'
' Cvor je dokument ili stavka sa svojim ID-om. Ivica postoji samo tamo gde je u
' kanonu zapisan ID drugog cvora (SLED-IDENT-01). Broj dokumenta, datum, klasa i
' kolone bez pisca (Otkup.BrojZbirne, Otkup.OtpremnicaID, Otpremnica.BrojZbirne,
' GeneracijaID) su osobine ili ostaci, nikad ivica -- jezgro ih ne cita.
'
' IVICE OVOG REZA (E7 i E8 dodaje S9-7, pa GP lanac do tada stoji za sebe):
'   E1 otkup -> stavke otkupa        modOtkup.StavkeOtkupaRedovi
'   E2 otkup <-> otpremnica          modDokumenta.AktivnoClanstvoOtpremnica
'   E3 otpremnica <-> zbirna         modDokumenta.AktivnoClanstvoZbirni
'   E4 zbirna <-> prijemnica         modAutoHladnjaca.AktivnePrijemniceZbirnih
'                                    (unapred) i Prijemnica.ZbirnaRoditeljID
'                                    (unazad) -- ista kolona, isto pravilo
'   E5 prijemnica -> stavke          modDokumenta.StavkePrijemniceRedovi
'   E6 stavka prijemnice -> faktura  modFaktura.FakturaLinijePoStavci
'   E9 prerada <-> utovar (po kg)    modUtovar.UtovarenoPoPreradi
'   E10 utovar -> GP faktura         marker Utovar.FakturaID, potvrdjen kroz
'                                    modUtovar.AktivnihFstZaUtovar (koje fakture
'                                    tvrde utovar) i modSEFMapper.GpFakturaRazlog
'
' VEZE DAJU KANONSKI STROGI CITACI, jedan prolaz po tabeli (N13). Jezgro nema
' svoju kopiju pravila clanstva: sta je "aktivno clanstvo" i sta je kvar zna
' samo citac koji to pravilo drzi i za pisce.
'
' FAIL-CLOSED (SLED-INV-01): ivica koja ne razresava se NE POGADJA. Kad strog
' citac odbije tabelu, jezgro ne pada i ne trazi drugi put do veze -- cvor na
' toj ivici dobija razlog u koloni kvara i graf od njega ne nastavlja. Ostale
' ivice rade. Granica koju jezgro nasledjuje (KI-081): strogi citaci su
' globalni, pa kvar JEDNOG dokumenta gasi tu ivicu za svaki trag dok se ne
' popravi; poruka citaca imenuje dokument.
'
' SAMO AKTIVNE VEZE (SLED-ODL-03): stornirani dokument se ne crta. Istorija
' storna i ispravki je poseban prikaz iz tblStornoVeze, ne deo ovog grafa.
'
' NADSKUP (SLED-ODL-01): alokacije nema. Kad zbirna ima vise aktivnih
' prijemnica, sve sto je sa druge strane te ivice nosi oznaku nadskupa
' ("moguce poreklo").
'
' KLASA (SLED-KLASA-01): odseca se SAMO u hladnjackom lancu, gde je 1:1 po
' klasi dokazano piscem. Svuda drugde se nista ne odseca: cvorovi iste klase
' idu prvi i nose oznaku, ostali dolaze posle njih sa oznakom druge klase.
' ============================================================

' --- vrste cvorova ---
Public Const SLG_OTK As String = "OTK"      ' otkup (blok)
Public Const SLG_OKS As String = "OKS"      ' stavka otkupa; ID = OtkupID & "#" & RedniBroj
Public Const SLG_OTP As String = "OTP"      ' otpremnica
Public Const SLG_ZBR As String = "ZBR"      ' zbirna
Public Const SLG_PRJ As String = "PRJ"      ' prijemnica
Public Const SLG_PRS As String = "PRS"      ' stavka prijemnice (PrijemnicaStavkaID)
Public Const SLG_FAK As String = "FAK"      ' faktura, dosegnuta kroz svoju liniju (E6) ili utovar (E10)
Public Const SLG_PRE As String = "PRE"      ' prerada
Public Const SLG_UTV As String = "UTV"      ' utovar

' --- kolone rezultata (2D niz, 1..n, 1..SLG_KOLONA) ---
Public Const SLG_C_NIVO As Long = 1         ' rastojanje od polaznog cvora (0 = polazni)
Public Const SLG_C_VRSTA As Long = 2
Public Const SLG_C_ID As Long = 3
Public Const SLG_C_LABELA As Long = 4       ' broj dokumenta; za stavku klasa
Public Const SLG_C_DATUM As Long = 5        ' SVOJ datum cvora; stavka ga nema (Empty)
Public Const SLG_C_STATUS As Long = 6       ' SVOJ status cvora; stavka ga nema ("")
Public Const SLG_C_KLASA As Long = 7        ' samo stavke
Public Const SLG_C_KG As Long = 8           ' samo stavke
Public Const SLG_C_OZNAKA As Long = 9       ' SLG_ISTA_KLASA / SLG_DRUGA_KLASA / ""
Public Const SLG_C_NADSKUP As Long = 10     ' True = moguce poreklo, ne tacno
Public Const SLG_C_KVAR As Long = 11        ' neprazno = cvor je oznacen, graf ne ide dalje
Public Const SLG_C_KOOPERANT As Long = 12   ' samo otkup
Public Const SLG_C_PARCELA As Long = 13     ' samo otkup
Public Const SLG_C_PREKO As Long = 14       ' ID cvora preko koga se doslo
Public Const SLG_KOLONA As Long = 14

Public Const SLG_ISTA_KLASA As String = "ISTA_KLASA"
Public Const SLG_DRUGA_KLASA As String = "DRUGA_KLASA"
Public Const SLG_STATUS_AKTIVAN As String = "AKTIVAN"
Public Const SLG_STATUS_STORNIRANO As String = "STORNIRANO"

' --- dokumenti sa zaglavljem: mesto u nizovima stanja ---
Private Const SL_OTK As Long = 1
Private Const SL_OTP As Long = 2
Private Const SL_ZBR As Long = 3
Private Const SL_PRJ As Long = 4
Private Const SL_FAK As Long = 5
Private Const SL_PRE As Long = 6
Private Const SL_UTV As Long = 7

' --- stanje jednog hoda; puni ga Ucitaj, prazni Oslobodi. Nije kes: svaki
' javni poziv cita tabele ispocetka, pa odgovor ne moze da zastari. ---
Private m_Data(1 To 7) As Variant           ' zaglavlja (GetTableData)
Private m_Idx(1 To 7) As Object             ' UCase(ID) -> red (prvi)
Private m_Dupli(1 To 7) As Object           ' UCase(ID) koji postoji vise puta
Private m_cID(1 To 7) As Long
Private m_cBroj(1 To 7) As Long
Private m_cDatum(1 To 7) As Long
Private m_cSto(1 To 7) As Long
Private m_cStatus(1 To 7) As Long
Private m_cOtkKoop As Long
Private m_cOtkParcela As Long
Private m_cPrjRod As Long

Private m_E1 As Variant                     ' stavke otkupa (StavkeOtkupaRedovi)
Private m_E1PoOtk As Object                 ' UCase(OtkupID) -> Collection(red u m_E1)
Private m_E1Kvar As String
Private m_E2 As Object                      ' UCase(OtkupID) -> OtpremnicaID
Private m_E2Obrnuto As Object               ' UCase(OtpremnicaID) -> Collection(OtkupID)
Private m_E2Kvar As String
Private m_E3 As Object                      ' UCase(OtpremnicaID) -> ZbirnaID
Private m_E3Obrnuto As Object               ' UCase(ZbirnaID) -> Collection(OtpremnicaID)
Private m_E3Kvar As String
Private m_E4 As Object                      ' UCase(ZbirnaID) -> Collection(PrijemnicaID)
Private m_E4KvarPoZbirni As Object          ' UCase(ZbirnaID) -> razlog
Private m_E4Kvar As String
Private m_E5 As Variant                     ' stavke prijemnice (StavkePrijemniceRedovi)
Private m_E5PoPrj As Object                 ' UCase(PrijemnicaID) -> Collection(red u m_E5)
Private m_E5PoStavci As Object              ' UCase(PrijemnicaStavkaID) -> red u m_E5
Private m_E5Kvar As String
Private m_E6 As Object                      ' UCase(PrijemnicaStavkaID) -> Array(FakturaID, ...)
Private m_E6PoFakturi As Object             ' UCase(FakturaID) -> Collection(kljuc stavke)
Private m_E6Kvar As String
Private m_E9Zbir As Object                  ' PreradaID -> ukupno aktivno utovareno (kg)
Private m_E9PoPreradi As Object             ' UCase(PreradaID) -> Collection(Array(kljuc utovara, kg))
Private m_E9PoUtovaru As Object             ' UCase(UtovarID) -> Collection(Array(kljuc prerade, kg))
Private m_E9KvarPoPreradi As Object         ' UCase(PreradaID) -> razlog
Private m_E9Kvar As String
Private m_E10PoFakturi As Object            ' UCase(Utovar.FakturaID) -> Collection(kljuc utovara)
Private m_cUtFaktura As Long

Private m_Redovi As Collection              ' redovi odgovora, redom hoda
Private m_Vidjen As Object                  ' cvorovi vec u odgovoru
Private m_Ctx As Object                     ' UCase(klasa) polaznog cvora; prazno = bez oznaka

' ============================================================
' JAVNI ULAZI
' ============================================================

' PRECI: odakle je cvor dosao (ka otkupu, kooperantu i parceli).
' Vraca 2D niz (kolone SLG_C_*), po nivoima, ili Empty kad polaznog cvora nema.
Public Function SledPreci(ByVal vrsta As String, ByVal id As String) As Variant
    SledPreci = Hodaj(vrsta, id, True)
End Function

' POTOMCI: kuda je cvor otisao (ka prijemnici i fakturi).
Public Function SledPotomci(ByVal vrsta As String, ByVal id As String) As Variant
    SledPotomci = Hodaj(vrsta, id, False)
End Function

' ============================================================
' HOD
' ============================================================

Private Function Hodaj(ByVal vrsta As String, ByVal id As String, _
                       ByVal unazad As Boolean) As Variant
    Const SRC As String = "modSledljivost.Hodaj"
    Dim kljuc As String, errNum As Long, errDesc As String, errSrc As String

    Hodaj = Empty
    vrsta = UCase$(Trim$(vrsta))
    kljuc = UCase$(Trim$(id))
    If Len(kljuc) = 0 Then Exit Function
    If Not VrstaPoznata(vrsta) Then
        Err.Raise vbObjectError + 9101, SRC, "Nepoznata vrsta cvora: " & vrsta
    End If

    On Error GoTo EH
    Ucitaj
    Set m_Redovi = New Collection
    Set m_Vidjen = CreateObject("Scripting.Dictionary")
    Set m_Ctx = KlaseStarta(vrsta, kljuc)

    If PolazniPostoji(vrsta, kljuc) Then
        If unazad Then
            Select Case vrsta
                Case SLG_FAK: UzFakturu kljuc, 0
                Case SLG_PRS: UzStavkuPrijemnice kljuc, 0, ""
                Case SLG_PRJ: UzPrijemnicu kljuc, 0, ""
                Case SLG_ZBR: UzZbirnu kljuc, 0, ""
                Case SLG_OTP: UzOtpremnicu kljuc, 0, "", False, False
                Case SLG_OTK: UzOtkup kljuc, 0, "", False, False
                Case SLG_OKS: StavkaOtkupa kljuc, 0
                Case SLG_UTV: UzUtovar kljuc, 0, ""
                Case SLG_PRE: DodajDok 0, SL_PRE, kljuc, False, KvarPrerade(kljuc), ""
            End Select
        Else
            Select Case vrsta
                Case SLG_OKS: NizStavkuOtkupa kljuc
                Case SLG_OTK: NizOtkup kljuc, 0, ""
                Case SLG_OTP: NizOtpremnicu kljuc, 0, ""
                Case SLG_ZBR: NizZbirnu kljuc, 0, ""
                Case SLG_PRJ: NizPrijemnicu kljuc, 0, "", False, False
                Case SLG_PRS: NizStavkuPrijemnicePoKljucu kljuc
                Case SLG_FAK: DodajDok 0, SL_FAK, kljuc, False, "", ""
                Case SLG_PRE: NizPreradu kljuc
                Case SLG_UTV: NizUtovar kljuc, 0, "", Empty
            End Select
        End If
        Hodaj = Rezultat()
    End If

    Oslobodi
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.Source
    Oslobodi
    Err.Raise errNum, errSrc, errDesc
End Function

Private Function VrstaPoznata(ByVal vrsta As String) As Boolean
    Select Case vrsta
        Case SLG_OTK, SLG_OKS, SLG_OTP, SLG_ZBR, SLG_PRJ, SLG_PRS, SLG_FAK, SLG_PRE, SLG_UTV
            VrstaPoznata = True
    End Select
End Function

' Polazni cvor koga nema daje prazan odgovor. Stavka cija se tabela ne moze
' procitati ipak prolazi: odgovor tada nosi jedan red sa razlogom.
Private Function PolazniPostoji(ByVal vrsta As String, ByVal kljuc As String) As Boolean
    Select Case vrsta
        Case SLG_PRS
            If Len(m_E5Kvar) > 0 Then
                PolazniPostoji = True
            Else
                PolazniPostoji = m_E5PoStavci.Exists(kljuc)
            End If
        Case SLG_OKS
            If Len(m_E1Kvar) > 0 Then
                PolazniPostoji = True
            Else
                PolazniPostoji = (RedStavkeOtkupa(kljuc) > 0)
            End If
        Case Else
            PolazniPostoji = (RedCvora(SlotVrste(vrsta), kljuc) > 0)
    End Select
End Function

' ------------------------------------------------------------
' UNAZAD: faktura -> stavka prijemnice -> prijemnica -> zbirna ->
'         otpremnica -> otkup -> stavke otkupa
' ------------------------------------------------------------

Private Sub UzFakturu(ByVal kljuc As String, ByVal nivo As Long)
    Dim lista As Collection, v As Variant, fakID As String, fak As String, kvar As String
    If Not DodajDok(nivo, SL_FAK, kljuc, False, SpojiKvar(m_E6Kvar, m_E9Kvar), "") Then Exit Sub
    fakID = IdCvora(SL_FAK, kljuc)

    ' Faktura prijemnice: njene linije po stavci (E6).
    If Len(m_E6Kvar) = 0 Then
        If m_E6PoFakturi.Exists(kljuc) Then
            Set lista = m_E6PoFakturi(kljuc)
            For Each v In lista
                UzStavkuPrijemnice CStr(v), nivo + 1, fakID
            Next v
        End If
    End If

    ' GP faktura: utovar ciji marker pokazuje na nju, potvrdjen pravilom E10.
    If Len(m_E9Kvar) > 0 Then Exit Sub
    If Not m_E10PoFakturi.Exists(kljuc) Then Exit Sub
    Set lista = m_E10PoFakturi(kljuc)
    For Each v In lista
        fak = FakturaUtovara(CStr(v), kvar)
        If Len(kvar) > 0 Then
            DodajDok nivo + 1, SL_UTV, CStr(v), False, kvar, fakID
        ElseIf StrComp(fak, fakID, vbTextCompare) = 0 Then
            UzUtovar CStr(v), nivo + 1, fakID
        End If
    Next v
End Sub

' E9 unazad: prerade utovara, svaka sa kilogramima koje je taj utovar poneo.
Private Sub UzUtovar(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim lista As Collection, v As Variant, utvID As String

    If Len(m_E9Kvar) > 0 Then
        DodajDok nivo, SL_UTV, kljuc, False, m_E9Kvar, preko
        Exit Sub
    End If
    If Not DodajDok(nivo, SL_UTV, kljuc, False, "", preko) Then Exit Sub
    If Not m_E9PoUtovaru.Exists(kljuc) Then Exit Sub
    utvID = IdCvora(SL_UTV, kljuc)
    Set lista = m_E9PoUtovaru(kljuc)
    For Each v In lista
        DodajDok nivo + 1, SL_PRE, CStr(v(0)), False, KvarPrerade(CStr(v(0))), utvID, "", kljuc, v(1)
    Next v
End Sub

Private Sub UzStavkuPrijemnice(ByVal kljuc As String, ByVal nivo As Long, _
                               ByVal preko As String)
    Dim r As Long, kl As String
    If m_Vidjen.Exists(SLG_PRS & "|" & kljuc) Then Exit Sub
    m_Vidjen.Add SLG_PRS & "|" & kljuc, True

    If Len(m_E5Kvar) > 0 Then
        Dodaj nivo, SLG_PRS, kljuc, "", Empty, "", "", Empty, "", False, m_E5Kvar, "", "", preko
        Exit Sub
    End If
    If Not m_E5PoStavci.Exists(kljuc) Then
        Dodaj nivo, SLG_PRS, kljuc, "", Empty, "", "", Empty, "", False, _
              "Stavka prijemnice " & kljuc & " ne postoji.", "", "", preko
        Exit Sub
    End If

    r = CLng(m_E5PoStavci(kljuc))
    kl = CStr(m_E5(r, 3))
    Dodaj nivo, SLG_PRS, CStr(m_E5(r, 7)), kl, Empty, "", kl, m_E5(r, 4), _
          OznakaKlase(kl), False, "", "", "", preko
    UzPrijemnicu UCase$(Trim$(CStr(m_E5(r, 1)))), nivo + 1, CStr(m_E5(r, 7))
End Sub

Private Sub UzPrijemnicu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim r As Long, zbr As String, kvar As String

    r = RedCvora(SL_PRJ, kljuc)
    If r > 0 Then
        If CvorAktivanJednom(SL_PRJ, kljuc) Then
            zbr = UCase$(ZbirnaPrijemnice(r))
            If Len(zbr) = 0 Then
                kvar = "Prijemnica " & IdCvora(SL_PRJ, kljuc) & _
                       " nema roditeljsku zbirnu (ZbirnaRoditeljID)."
            ElseIf Not CvorAktivanJednom(SL_ZBR, zbr) Then
                kvar = "Prijemnica " & IdCvora(SL_PRJ, kljuc) & _
                       " pokazuje na zbirnu koja nije aktivna tacno jednom: " & zbr & "."
            End If
        End If
    End If

    If Not DodajDok(nivo, SL_PRJ, kljuc, False, kvar, preko) Then Exit Sub
    If Len(kvar) > 0 Then Exit Sub
    UzZbirnu zbr, nivo + 1, IdCvora(SL_PRJ, kljuc)
End Sub

' E4 unazad: roditeljska zbirna prijemnice, ISKLJUCIVO po ZbirnaRoditeljID.
Private Function ZbirnaPrijemnice(ByVal r As Long) As String
    ZbirnaPrijemnice = Trim$(NzToText(m_Data(SL_PRJ)(r, m_cPrjRod)))
End Function

Private Sub UzZbirnu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim lista As Collection, v As Variant
    Dim nadskup As Boolean, hlad As Boolean, razlog As String

    If Len(m_E3Kvar) > 0 Then
        DodajDok nivo, SL_ZBR, kljuc, False, m_E3Kvar, preko
        Exit Sub
    End If

    nadskup = ZbirnaJeNadskup(kljuc)
    hlad = ZbirnaJeHladnjacka(kljuc, razlog)
    If Not DodajDok(nivo, SL_ZBR, kljuc, False, razlog, preko) Then Exit Sub
    If Not m_E3Obrnuto.Exists(kljuc) Then Exit Sub        ' zbirna bez clanova
    Set lista = m_E3Obrnuto(kljuc)
    For Each v In lista
        UzOtpremnicu UCase$(CStr(v)), nivo + 1, IdCvora(SL_ZBR, kljuc), nadskup, hlad
    Next v
End Sub

Private Sub UzOtpremnicu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String, _
                         ByVal nadskup As Boolean, ByVal hlad As Boolean)
    Dim lista As Collection, v As Variant

    If Len(m_E2Kvar) > 0 Then
        DodajDok nivo, SL_OTP, kljuc, nadskup, m_E2Kvar, preko
        Exit Sub
    End If
    If Not DodajDok(nivo, SL_OTP, kljuc, nadskup, "", preko) Then Exit Sub
    If Not m_E2Obrnuto.Exists(kljuc) Then Exit Sub
    Set lista = m_E2Obrnuto(kljuc)
    For Each v In lista
        UzOtkup UCase$(CStr(v)), nivo + 1, IdCvora(SL_OTP, kljuc), nadskup, hlad
    Next v
End Sub

Private Sub UzOtkup(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String, _
                    ByVal nadskup As Boolean, ByVal hlad As Boolean)
    Dim lista As Collection, ostaju As Collection, v As Variant
    Dim r As Long, kl As String, imaIstu As Boolean, oznaka As String, otkID As String

    If Len(m_E1Kvar) > 0 Then
        ' Stavke nisu citljive: klasa se ne zna, pa se otkup ne odseca.
        DodajDok nivo, SL_OTK, kljuc, nadskup, m_E1Kvar, preko
        Exit Sub
    End If

    Set ostaju = New Collection
    If m_E1PoOtk.Exists(kljuc) Then
        Set lista = m_E1PoOtk(kljuc)
        For Each v In lista
            kl = CStr(m_E1(CLng(v), 3))
            If Not KlasaSeOdseca(hlad, kl) Then
                ostaju.Add CLng(v)
                If ImaKontekst() Then
                    If m_Ctx.Exists(UCase$(kl)) Then imaIstu = True
                End If
            End If
        Next v
        ' Hladnjacki lanac: otkup kome nije ostala nijedna stavka nije izvor.
        If lista.count > 0 And ostaju.count = 0 Then Exit Sub
    End If

    If ImaKontekst() Then
        If imaIstu Then oznaka = SLG_ISTA_KLASA Else oznaka = SLG_DRUGA_KLASA
    End If
    If Not DodajDok(nivo, SL_OTK, kljuc, nadskup, "", preko, oznaka) Then Exit Sub

    otkID = IdCvora(SL_OTK, kljuc)
    For Each v In ostaju
        r = CLng(v)
        kl = CStr(m_E1(r, 3))
        Dodaj nivo + 1, SLG_OKS, otkID & "#" & CStr(m_E1(r, 2)), kl, Empty, "", kl, _
              m_E1(r, 4), OznakaKlase(kl), nadskup, "", "", "", otkID
    Next v
End Sub

' Polazna stavka otkupa: iznad nje nema nicega (kooperant i parcela su osobine
' njenog otkupa), pa je odgovor ona sama.
Private Sub StavkaOtkupa(ByVal kljuc As String, ByVal nivo As Long)
    Dim r As Long, kl As String
    If Len(m_E1Kvar) > 0 Then
        Dodaj nivo, SLG_OKS, kljuc, "", Empty, "", "", Empty, "", False, m_E1Kvar, "", "", ""
        Exit Sub
    End If
    r = RedStavkeOtkupa(kljuc)
    If r = 0 Then Exit Sub
    kl = CStr(m_E1(r, 3))
    Dodaj nivo, SLG_OKS, CStr(m_E1(r, 1)) & "#" & CStr(m_E1(r, 2)), kl, Empty, "", kl, _
          m_E1(r, 4), OznakaKlase(kl), False, "", "", "", ""
End Sub

' ------------------------------------------------------------
' UNAPRED: stavka otkupa -> otkup -> otpremnica -> zbirna -> prijemnica ->
'          stavka prijemnice -> faktura
' ------------------------------------------------------------

Private Sub NizStavkuOtkupa(ByVal kljuc As String)
    Dim poz As Long
    StavkaOtkupa kljuc, 0
    If Len(m_E1Kvar) > 0 Then Exit Sub
    poz = InStr(1, kljuc, "#")
    If poz < 2 Then Exit Sub
    NizOtkup Left$(kljuc, poz - 1), 1, kljuc
End Sub

Private Sub NizOtkup(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim otp As String

    If Len(m_E2Kvar) > 0 Then
        DodajDok nivo, SL_OTK, kljuc, False, m_E2Kvar, preko
        Exit Sub
    End If
    If Not DodajDok(nivo, SL_OTK, kljuc, False, "", preko) Then Exit Sub
    otp = OtpremnicaOtkupa(kljuc)
    If Len(otp) = 0 Then Exit Sub                         ' otkup bez otpremnice: kraj, nije kvar
    NizOtpremnicu UCase$(otp), nivo + 1, IdCvora(SL_OTK, kljuc)
End Sub

' E2 unapred: aktivna otpremnica otkupa, ISKLJUCIVO iz kanonskog clanstva.
Private Function OtpremnicaOtkupa(ByVal otkKljuc As String) As String
    If m_E2.Exists(otkKljuc) Then OtpremnicaOtkupa = Trim$(CStr(m_E2(otkKljuc)))
End Function

Private Sub NizOtpremnicu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim zbr As String

    If Len(m_E3Kvar) > 0 Then
        DodajDok nivo, SL_OTP, kljuc, False, m_E3Kvar, preko
        Exit Sub
    End If
    If Not DodajDok(nivo, SL_OTP, kljuc, False, "", preko) Then Exit Sub
    zbr = ZbirnaOtpremnice(kljuc)
    If Len(zbr) = 0 Then Exit Sub                         ' otpremnica bez zbirne: kraj, nije kvar
    NizZbirnu UCase$(zbr), nivo + 1, IdCvora(SL_OTP, kljuc)
End Sub

' E3 unapred: aktivna zbirna otpremnice, ISKLJUCIVO iz kanonskog clanstva.
Private Function ZbirnaOtpremnice(ByVal otpKljuc As String) As String
    If m_E3.Exists(otpKljuc) Then ZbirnaOtpremnice = Trim$(CStr(m_E3(otpKljuc)))
End Function

Private Sub NizZbirnu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String)
    Dim lista As Collection, v As Variant
    Dim nadskup As Boolean, hlad As Boolean, razlog As String, kvar As String

    If Len(m_E4Kvar) > 0 Then kvar = m_E4Kvar
    If Len(kvar) = 0 Then
        If m_E4KvarPoZbirni.Exists(kljuc) Then            ' skup prijemnica zbirne se ne zna: oznaci, ne vezuj
            kvar = CStr(m_E4KvarPoZbirni(kljuc))
        End If
    End If
    If Len(kvar) > 0 Then
        DodajDok nivo, SL_ZBR, kljuc, False, kvar, preko
        Exit Sub
    End If

    hlad = ZbirnaJeHladnjacka(kljuc, razlog)
    If Not DodajDok(nivo, SL_ZBR, kljuc, False, razlog, preko) Then Exit Sub
    If Not m_E4.Exists(kljuc) Then Exit Sub               ' zbirna bez prijemnice: kraj, nije kvar
    Set lista = m_E4(kljuc)
    nadskup = (lista.count > 1)
    For Each v In lista
        NizPrijemnicu UCase$(Trim$(CStr(v))), nivo + 1, IdCvora(SL_ZBR, kljuc), nadskup, hlad
    Next v
End Sub

Private Sub NizPrijemnicu(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String, _
                          ByVal nadskup As Boolean, ByVal hlad As Boolean)
    Dim lista As Collection, v As Variant, r As Long

    If Len(m_E5Kvar) > 0 Then
        DodajDok nivo, SL_PRJ, kljuc, nadskup, m_E5Kvar, preko
        Exit Sub
    End If
    If Not DodajDok(nivo, SL_PRJ, kljuc, nadskup, "", preko) Then Exit Sub
    If Not m_E5PoPrj.Exists(kljuc) Then Exit Sub
    Set lista = m_E5PoPrj(kljuc)
    For Each v In lista
        r = CLng(v)
        If Not KlasaSeOdseca(hlad, CStr(m_E5(r, 3))) Then
            NizStavkuPrijemnice r, nivo + 1, IdCvora(SL_PRJ, kljuc), nadskup
        End If
    Next v
End Sub

Private Sub NizStavkuPrijemnicePoKljucu(ByVal kljuc As String)
    If Len(m_E5Kvar) > 0 Then
        Dodaj 0, SLG_PRS, kljuc, "", Empty, "", "", Empty, "", False, m_E5Kvar, "", "", ""
        Exit Sub
    End If
    NizStavkuPrijemnice CLng(m_E5PoStavci(kljuc)), 0, "", False
End Sub

Private Sub NizStavkuPrijemnice(ByVal r As Long, ByVal nivo As Long, ByVal preko As String, _
                                ByVal nadskup As Boolean)
    Dim kljuc As String, kl As String, stavkaID As String, linija As Variant

    stavkaID = CStr(m_E5(r, 7))
    kljuc = UCase$(Trim$(stavkaID))
    If m_Vidjen.Exists(SLG_PRS & "|" & kljuc) Then Exit Sub
    m_Vidjen.Add SLG_PRS & "|" & kljuc, True

    kl = CStr(m_E5(r, 3))
    Dodaj nivo, SLG_PRS, stavkaID, kl, Empty, "", kl, m_E5(r, 4), OznakaKlase(kl), _
          nadskup, m_E6Kvar, "", "", preko
    If Len(m_E6Kvar) > 0 Then Exit Sub
    If Not m_E6.Exists(kljuc) Then Exit Sub               ' nefakturisana stavka: kraj, nije kvar
    linija = m_E6(kljuc)
    DodajDok nivo + 1, SL_FAK, UCase$(Trim$(CStr(linija(0)))), nadskup, "", stavkaID, "", kljuc
End Sub

' ------------------------------------------------------------
' GP LANAC UNAPRED: prerada -> utovar (po kg) -> GP faktura
' ------------------------------------------------------------

Private Sub NizPreradu(ByVal kljuc As String)
    Dim lista As Collection, v As Variant, kvar As String, ukupno As Variant, preID As String

    kvar = m_E9Kvar
    If Len(kvar) = 0 Then kvar = KvarPrerade(kljuc)
    preID = IdCvora(SL_PRE, kljuc)
    ukupno = Empty
    If Not m_E9Zbir Is Nothing Then
        If m_E9Zbir.Exists(preID) Then ukupno = m_E9Zbir(preID)
    End If
    ' Prerada nosi UKUPNO utovareno; svaki utovar ispod nje svoje kilograme.
    If Not DodajDok(0, SL_PRE, kljuc, False, kvar, "", "", "", ukupno) Then Exit Sub
    If Len(kvar) > 0 Then Exit Sub
    If Not m_E9PoPreradi.Exists(kljuc) Then Exit Sub      ' prerada bez utovara: kraj, nije kvar
    Set lista = m_E9PoPreradi(kljuc)
    For Each v In lista
        NizUtovar CStr(v(0)), 1, preID, v(1)
    Next v
End Sub

Private Sub NizUtovar(ByVal kljuc As String, ByVal nivo As Long, ByVal preko As String, _
                      ByVal kg As Variant)
    Dim fak As String, kvar As String

    fak = FakturaUtovara(kljuc, kvar)
    If Not DodajDok(nivo, SL_UTV, kljuc, False, kvar, preko, "", "", kg) Then Exit Sub
    If Len(fak) = 0 Then Exit Sub                         ' nefakturisan ili oznacen utovar
    DodajDok nivo + 1, SL_FAK, UCase$(fak), False, "", IdCvora(SL_UTV, kljuc)
End Sub

' E10: GP faktura utovara, ili "" kad je nema. Vezu potvrdjuju ZAJEDNO marker
' na utovaru, aktivne linije fakture koje utovar tvrde (kanonski brojac, isti
' koji drzi pisac i storno) i GP pravilo fakture (isto koje drzi SEF kapija).
' Kad se ne slazu, veze nema, a outKvar nosi razlog (SLED-INV-01).
Private Function FakturaUtovara(ByVal utvKljuc As String, ByRef outKvar As String) As String
    Dim r As Long, utvID As String, marker As String, n As Long
    Dim fakture As Object, razlog As String

    outKvar = ""
    r = RedCvora(SL_UTV, utvKljuc)
    If r = 0 Then Exit Function
    utvID = IdCvora(SL_UTV, utvKljuc)
    marker = Trim$(NzToText(m_Data(SL_UTV)(r, m_cUtFaktura)))

    On Error GoTo EH
    n = modUtovar.AktivnihFstZaUtovar(utvID, fakture)
    If n = 0 And Len(marker) = 0 Then Exit Function       ' nefakturisan utovar: kraj, nije kvar
    If n = 0 Then
        outKvar = "Utovar " & utvID & " nosi FakturaID " & marker & _
                  ", a nijedna aktivna linija fakture ga ne tvrdi."
        Exit Function
    End If
    If Len(marker) = 0 Then
        outKvar = "Aktivne linije fakture tvrde utovar " & utvID & ", a utovar ne nosi FakturaID."
        Exit Function
    End If
    If fakture.count <> 1 Or Not fakture.Exists(marker) Then
        outKvar = "Utovar " & utvID & " nosi FakturaID " & marker & _
                  ", a tvrde ga linije faktura: " & Join(fakture.Keys, ", ") & "."
        Exit Function
    End If
    If Not CvorAktivanJednom(SL_FAK, UCase$(marker)) Then
        outKvar = "Faktura " & marker & " utovara " & utvID & " nije aktivna tacno jednom."
        Exit Function
    End If
    razlog = modSEFMapper.GpFakturaRazlog(marker)
    If Len(razlog) > 0 Then
        outKvar = razlog
        Exit Function
    End If
    FakturaUtovara = marker
    Exit Function
EH:
    outKvar = "GP faktura utovara " & utvID & " nije citljiva (E10): " & Err.description
    Err.Clear
End Function

Private Function KvarPrerade(ByVal kljuc As String) As String
    If m_E9KvarPoPreradi Is Nothing Then Exit Function
    If m_E9KvarPoPreradi.Exists(kljuc) Then KvarPrerade = CStr(m_E9KvarPoPreradi(kljuc))
End Function

' ============================================================
' PRAVILA
' ============================================================

' SLED-KLASA-01: da li se cvor date klase ODSECA iz odgovora.
Private Function KlasaSeOdseca(ByVal hladnjacki As Boolean, ByVal klasa As String) As Boolean
    If Not ImaKontekst() Then Exit Function
    If Not hladnjacki Then Exit Function                  ' van hladnjackog lanca se nista ne odseca
    KlasaSeOdseca = Not m_Ctx.Exists(UCase$(Trim$(klasa)))
End Function

Private Function OznakaKlase(ByVal klasa As String) As String
    If Not ImaKontekst() Then Exit Function
    If m_Ctx.Exists(UCase$(Trim$(klasa))) Then
        OznakaKlase = SLG_ISTA_KLASA
    Else
        OznakaKlase = SLG_DRUGA_KLASA
    End If
End Function

Private Function ImaKontekst() As Boolean
    If m_Ctx Is Nothing Then Exit Function
    ImaKontekst = (m_Ctx.count > 0)
End Function

' Klase polaznog cvora: stavka nosi svoju, dokument klase svojih stavki. Cvor
' bez stavki (otpremnica, zbirna) nema kontekst, pa odgovor nema oznake klase.
Private Function KlaseStarta(ByVal vrsta As String, ByVal kljuc As String) As Object
    Dim ctx As Object, lista As Collection, v As Variant, r As Long
    Set ctx = CreateObject("Scripting.Dictionary")
    Set KlaseStarta = ctx

    Select Case vrsta
        Case SLG_PRS
            If Len(m_E5Kvar) = 0 Then
                If m_E5PoStavci.Exists(kljuc) Then DodajKlasu ctx, m_E5(CLng(m_E5PoStavci(kljuc)), 3)
            End If
        Case SLG_PRJ
            If Len(m_E5Kvar) = 0 Then
                If m_E5PoPrj.Exists(kljuc) Then
                    Set lista = m_E5PoPrj(kljuc)
                    For Each v In lista
                        DodajKlasu ctx, m_E5(CLng(v), 3)
                    Next v
                End If
            End If
        Case SLG_FAK
            If Len(m_E6Kvar) = 0 And Len(m_E5Kvar) = 0 Then
                If m_E6PoFakturi.Exists(kljuc) Then
                    Set lista = m_E6PoFakturi(kljuc)
                    For Each v In lista
                        If m_E5PoStavci.Exists(CStr(v)) Then
                            DodajKlasu ctx, m_E5(CLng(m_E5PoStavci(CStr(v))), 3)
                        End If
                    Next v
                End If
            End If
        Case SLG_OKS
            If Len(m_E1Kvar) = 0 Then
                r = RedStavkeOtkupa(kljuc)
                If r > 0 Then DodajKlasu ctx, m_E1(r, 3)
            End If
        Case SLG_OTK
            If Len(m_E1Kvar) = 0 Then
                If m_E1PoOtk.Exists(kljuc) Then
                    Set lista = m_E1PoOtk(kljuc)
                    For Each v In lista
                        DodajKlasu ctx, m_E1(CLng(v), 3)
                    Next v
                End If
            End If
    End Select
End Function

Private Sub DodajKlasu(ByVal ctx As Object, ByVal klasa As Variant)
    Dim k As String
    k = UCase$(Trim$(CStr(klasa)))
    If Len(k) = 0 Then Exit Sub
    If Not ctx.Exists(k) Then ctx.Add k, True
End Sub

' SLED-ODL-01: zbirna je nadskup kad ima vise aktivnih prijemnica, i kad se
' skup njenih prijemnica ne zna -- "ne znam" se ne prikazuje kao tacno.
Private Function ZbirnaJeNadskup(ByVal kljuc As String) As Boolean
    Dim lista As Collection
    If Len(m_E4Kvar) > 0 Then
        ZbirnaJeNadskup = True
        Exit Function
    End If
    If m_E4KvarPoZbirni.Exists(kljuc) Then
        ZbirnaJeNadskup = True
        Exit Function
    End If
    If Not m_E4.Exists(kljuc) Then Exit Function
    Set lista = m_E4(kljuc)
    ZbirnaJeNadskup = (lista.count > 1)
End Function

' Hladnjacki lanac je 1:1 (HLD-LANAC-01): zbirna lanca ima TACNO JEDNU
' otpremnicu, i ona je hladnjacka po istom primitivu koji koristi lanac.
' "Ne znam" nije "jeste": razlog se vraca, a klasa se tada ne odseca -- visak
' je za opoziv robe bezbedan, izostavljen izvor nije.
Private Function ZbirnaJeHladnjacka(ByVal kljuc As String, ByRef outRazlog As String) As Boolean
    Dim lista As Collection, otp As String
    outRazlog = ""
    If Len(m_E3Kvar) > 0 Then Exit Function
    If Not m_E3Obrnuto.Exists(kljuc) Then Exit Function
    Set lista = m_E3Obrnuto(kljuc)
    If lista.count <> 1 Then Exit Function
    otp = IdCvora(SL_OTP, UCase$(CStr(lista(1))))

    On Error GoTo EH
    ZbirnaJeHladnjacka = modAutoHladnjaca.OtpremnicaJeHladnjacka(otp)
    Exit Function
EH:
    outRazlog = "Nije utvrdjeno da li je lanac hladnjacki: " & Err.description
    ZbirnaJeHladnjacka = False
    Err.Clear
End Function

' ============================================================
' CVOROVI
' ============================================================

Private Function SlotVrste(ByVal vrsta As String) As Long
    Select Case vrsta
        Case SLG_OTK: SlotVrste = SL_OTK
        Case SLG_OTP: SlotVrste = SL_OTP
        Case SLG_ZBR: SlotVrste = SL_ZBR
        Case SLG_PRJ: SlotVrste = SL_PRJ
        Case SLG_FAK: SlotVrste = SL_FAK
        Case SLG_PRE: SlotVrste = SL_PRE
        Case SLG_UTV: SlotVrste = SL_UTV
    End Select
End Function

Private Function VrstaSlota(ByVal slot As Long) As String
    Select Case slot
        Case SL_OTK: VrstaSlota = SLG_OTK
        Case SL_OTP: VrstaSlota = SLG_OTP
        Case SL_ZBR: VrstaSlota = SLG_ZBR
        Case SL_PRJ: VrstaSlota = SLG_PRJ
        Case SL_FAK: VrstaSlota = SLG_FAK
        Case SL_PRE: VrstaSlota = SLG_PRE
        Case SL_UTV: VrstaSlota = SLG_UTV
    End Select
End Function

Private Function NazivSlota(ByVal slot As Long) As String
    Select Case slot
        Case SL_OTK: NazivSlota = "Otkup"
        Case SL_OTP: NazivSlota = "Otpremnica"
        Case SL_ZBR: NazivSlota = "Zbirna"
        Case SL_PRJ: NazivSlota = "Prijemnica"
        Case SL_FAK: NazivSlota = "Faktura"
        Case SL_PRE: NazivSlota = "Prerada"
        Case SL_UTV: NazivSlota = "Utovar"
    End Select
End Function

Private Function RedCvora(ByVal slot As Long, ByVal kljuc As String) As Long
    Dim idx As Object
    If slot = 0 Then Exit Function
    Set idx = m_Idx(slot)
    If idx Is Nothing Then Exit Function
    If idx.Exists(kljuc) Then RedCvora = CLng(idx(kljuc))
End Function

Private Function RedStorniran(ByVal slot As Long, ByVal r As Long) As Boolean
    If m_cSto(slot) = 0 Then Exit Function
    RedStorniran = (StrComp(Trim$(NzToText(m_Data(slot)(r, m_cSto(slot)))), "Da", vbTextCompare) = 0)
End Function

' Identitet cvora (SLED-INV-01): postoji TACNO JEDNOM i nije storniran.
Private Function CvorAktivanJednom(ByVal slot As Long, ByVal kljuc As String) As Boolean
    Dim r As Long, dupli As Object
    r = RedCvora(slot, kljuc)
    If r = 0 Then Exit Function
    Set dupli = m_Dupli(slot)
    If dupli.Exists(kljuc) Then Exit Function
    CvorAktivanJednom = Not RedStorniran(slot, r)
End Function

' ID onako kako je zapisan u zaglavlju; kljuc (velika slova) kad cvora nema.
Private Function IdCvora(ByVal slot As Long, ByVal kljuc As String) As String
    Dim r As Long
    r = RedCvora(slot, kljuc)
    If r = 0 Then
        IdCvora = kljuc
    Else
        IdCvora = Trim$(NzToText(m_Data(slot)(r, m_cID(slot))))
    End If
End Function

' Red stavke otkupa po kljucu "OTKUPID#RB"; 0 = nema je.
Private Function RedStavkeOtkupa(ByVal kljuc As String) As Long
    Dim poz As Long, otk As String, rb As String, lista As Collection, v As Variant
    poz = InStr(1, kljuc, "#")
    If poz < 2 Then Exit Function
    otk = Left$(kljuc, poz - 1)
    rb = Mid$(kljuc, poz + 1)
    If Not m_E1PoOtk.Exists(otk) Then Exit Function
    Set lista = m_E1PoOtk(otk)
    For Each v In lista
        If StrComp(Trim$(CStr(m_E1(CLng(v), 2))), rb, vbTextCompare) = 0 Then
            RedStavkeOtkupa = CLng(v)
            Exit Function
        End If
    Next v
End Function

' Dodaje cvor dokumenta u odgovor. Vraca True kad cvor postoji TACNO JEDNOM i
' aktivan je, tj. kad se od njega sme hodati dalje. Cvor koji to ne ispunjava,
' i cvor kome je pozivalac dao razlog, upisuje se sa razlogom u koloni kvara.
' zaVidjenje: dodatak kljucu "vec u odgovoru" (faktura se upisuje po liniji).
' kg: kilogrami koje cvor nosi u ovom odgovoru (GP lanac); izostavljeno = nema ih.
Private Function DodajDok(ByVal nivo As Long, ByVal slot As Long, ByVal kljuc As String, _
                          ByVal nadskup As Boolean, ByVal kvar As String, _
                          ByVal preko As String, Optional ByVal oznaka As String = "", _
                          Optional ByVal zaVidjenje As String = "", _
                          Optional ByVal kg As Variant) As Boolean
    Dim r As Long, vr As String, status As String, dupli As Object
    Dim koop As String, parcela As String, oznakaVidjenja As String, kgCvora As Variant

    kgCvora = Empty
    If Not IsMissing(kg) Then kgCvora = kg
    vr = VrstaSlota(slot)
    oznakaVidjenja = vr & "|" & kljuc & "|" & zaVidjenje
    If m_Vidjen.Exists(oznakaVidjenja) Then Exit Function
    m_Vidjen.Add oznakaVidjenja, True

    r = RedCvora(slot, kljuc)
    If r = 0 Then
        Dodaj nivo, vr, kljuc, "", Empty, "", "", Empty, oznaka, nadskup, _
              SpojiKvar(kvar, NazivSlota(slot) & " " & kljuc & " ne postoji."), "", "", preko
        Exit Function
    End If

    If RedStorniran(slot, r) Then
        status = SLG_STATUS_STORNIRANO
    ElseIf m_cStatus(slot) > 0 Then
        status = Trim$(NzToText(m_Data(slot)(r, m_cStatus(slot))))
    End If
    If Len(status) = 0 Then status = SLG_STATUS_AKTIVAN

    Set dupli = m_Dupli(slot)
    If dupli.Exists(kljuc) Then
        kvar = SpojiKvar(kvar, NazivSlota(slot) & " " & kljuc & " postoji vise puta.")
    End If

    If slot = SL_OTK Then
        koop = Trim$(NzToText(m_Data(slot)(r, m_cOtkKoop)))
        If m_cOtkParcela > 0 Then parcela = Trim$(NzToText(m_Data(slot)(r, m_cOtkParcela)))
    End If

    Dodaj nivo, vr, Trim$(NzToText(m_Data(slot)(r, m_cID(slot)))), _
          Trim$(NzToText(m_Data(slot)(r, m_cBroj(slot)))), m_Data(slot)(r, m_cDatum(slot)), _
          status, "", kgCvora, oznaka, nadskup, kvar, koop, parcela, preko

    DodajDok = (status <> SLG_STATUS_STORNIRANO And Not dupli.Exists(kljuc))
End Function

Private Function SpojiKvar(ByVal prvi As String, ByVal drugi As String) As String
    If Len(prvi) = 0 Then
        SpojiKvar = drugi
    ElseIf Len(drugi) = 0 Then
        SpojiKvar = prvi
    Else
        SpojiKvar = prvi & " " & drugi
    End If
End Function

Private Sub Dodaj(ByVal nivo As Long, ByVal vrsta As String, ByVal id As String, _
                  ByVal labela As String, ByVal datum As Variant, ByVal status As String, _
                  ByVal klasa As String, ByVal kg As Variant, ByVal oznaka As String, _
                  ByVal nadskup As Boolean, ByVal kvar As String, ByVal koop As String, _
                  ByVal parcela As String, ByVal preko As String)
    m_Redovi.Add Array(nivo, vrsta, id, labela, datum, status, klasa, kg, oznaka, _
                       nadskup, kvar, koop, parcela, preko)
End Sub

' Odgovor po nivoima; unutar nivoa ista klasa ide prva (SLED-KLASA-01), a
' redosled hoda se inace cuva (stabilno umetanje).
Private Function Rezultat() As Variant
    Dim n As Long, i As Long, j As Long, c As Long
    Dim red() As Variant, tmp As Variant, rez() As Variant

    Rezultat = Empty
    n = m_Redovi.count
    If n = 0 Then Exit Function

    ReDim red(1 To n)
    For i = 1 To n
        red(i) = m_Redovi(i)
    Next i

    For i = 2 To n
        tmp = red(i)
        j = i - 1
        Do While j >= 1
            If Not RedIdeIspred(tmp, red(j)) Then Exit Do
            red(j + 1) = red(j)
            j = j - 1
        Loop
        red(j + 1) = tmp
    Next i

    ReDim rez(1 To n, 1 To SLG_KOLONA)
    For i = 1 To n
        For c = 1 To SLG_KOLONA
            rez(i, c) = red(i)(c - 1)
        Next c
    Next i
    Rezultat = rez
End Function

Private Function RedIdeIspred(ByVal a As Variant, ByVal b As Variant) As Boolean
    If CLng(a(0)) <> CLng(b(0)) Then
        RedIdeIspred = (CLng(a(0)) < CLng(b(0)))
    Else
        RedIdeIspred = (RangOznake(CStr(a(8))) < RangOznake(CStr(b(8))))
    End If
End Function

Private Function RangOznake(ByVal oznaka As String) As Long
    Select Case oznaka
        Case SLG_ISTA_KLASA: RangOznake = 0
        Case SLG_DRUGA_KLASA: RangOznake = 2
        Case Else: RangOznake = 1
    End Select
End Function

' ============================================================
' UCITAVANJE -- jedan prolaz po tabeli (N13)
' ============================================================

Private Sub Ucitaj()
    Const SRC As String = "modSledljivost.Ucitaj"

    UcitajZaglavlja SL_OTK, TBL_OTKUP, COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_DATUM, COL_TRACE_IZDATO_STATUS
    UcitajZaglavlja SL_OTP, TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ, COL_OTP_DATUM, COL_TRACE_IZDATO_STATUS
    UcitajZaglavlja SL_ZBR, TBL_ZBIRNA, COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_DATUM, COL_TRACE_IZDATO_STATUS
    UcitajZaglavlja SL_PRJ, TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, ""
    UcitajZaglavlja SL_FAK, TBL_FAKTURE, COL_FAK_ID, COL_FAK_BROJ, COL_FAK_DATUM, COL_FAK_STATUS

    m_cOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    m_cOtkParcela = GetColumnIndex(TBL_OTKUP, COL_OTK_PARCELA)
    m_cPrjRod = RequireColumnIndex(TBL_PRIJEMNICA, COL_DETE_ZBIRNA_ROD, SRC)

    UcitajE1
    UcitajE2
    UcitajE3
    UcitajE4
    UcitajE5
    UcitajE6
    UcitajGp
End Sub

Private Sub UcitajZaglavlja(ByVal slot As Long, ByVal tbl As String, ByVal colID As String, _
                            ByVal colBroj As String, ByVal colDatum As String, _
                            ByVal colStatus As String)
    Const SRC As String = "modSledljivost.UcitajZaglavlja"
    Dim d As Variant, i As Long, k As String, idx As Object, dupli As Object

    Set idx = CreateObject("Scripting.Dictionary")
    Set dupli = CreateObject("Scripting.Dictionary")
    Set m_Idx(slot) = idx
    Set m_Dupli(slot) = dupli
    m_Data(slot) = Empty

    m_cID(slot) = RequireColumnIndex(tbl, colID, SRC)
    m_cBroj(slot) = RequireColumnIndex(tbl, colBroj, SRC)
    m_cDatum(slot) = RequireColumnIndex(tbl, colDatum, SRC)
    m_cSto(slot) = GetColumnIndex(tbl, COL_STORNIRANO)
    m_cStatus(slot) = 0
    If Len(colStatus) > 0 Then m_cStatus(slot) = GetColumnIndex(tbl, colStatus)

    d = GetTableData(tbl)
    If Not IsArray(d) Then Exit Sub
    m_Data(slot) = d
    For i = 1 To UBound(d, 1)
        k = UCase$(Trim$(NzToText(d(i, m_cID(slot)))))
        If Len(k) > 0 Then
            If idx.Exists(k) Then
                If Not dupli.Exists(k) Then dupli.Add k, True
            Else
                idx.Add k, i
            End If
        End If
    Next i
End Sub

Private Sub DodajUListu(ByVal d As Object, ByVal kljuc As String, ByVal vrednost As Variant)
    Dim lista As Collection
    If Len(kljuc) = 0 Then Exit Sub
    If d.Exists(kljuc) Then
        Set lista = d(kljuc)
    Else
        Set lista = New Collection
        d.Add kljuc, lista
    End If
    lista.Add vrednost
End Sub

Private Sub UcitajE1()
    Dim i As Long
    m_E1 = Empty
    m_E1Kvar = ""
    Set m_E1PoOtk = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    m_E1 = modOtkup.StavkeOtkupaRedovi()
    On Error GoTo 0
    If Not IsArray(m_E1) Then Exit Sub
    For i = 1 To UBound(m_E1, 1)
        DodajUListu m_E1PoOtk, UCase$(Trim$(CStr(m_E1(i, 1)))), i
    Next i
    Exit Sub
EH:
    m_E1Kvar = "Stavke otkupa nisu citljive (E1): " & Err.description
    m_E1 = Empty
    Err.Clear
End Sub

Private Sub UcitajE2()
    Dim k As Variant
    Set m_E2 = Nothing
    m_E2Kvar = ""
    Set m_E2Obrnuto = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    Set m_E2 = modDokumenta.AktivnoClanstvoOtpremnica()
    On Error GoTo 0
    For Each k In m_E2.Keys
        DodajUListu m_E2Obrnuto, UCase$(Trim$(CStr(m_E2(k)))), CStr(k)
    Next k
    Exit Sub
EH:
    m_E2Kvar = "Clanstvo otkupa u otpremnici nije citljivo (E2): " & Err.description
    Set m_E2 = Nothing
    Err.Clear
End Sub

Private Sub UcitajE3()
    Dim k As Variant
    Set m_E3 = Nothing
    m_E3Kvar = ""
    Set m_E3Obrnuto = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    Set m_E3 = modDokumenta.AktivnoClanstvoZbirni()
    On Error GoTo 0
    For Each k In m_E3.Keys
        DodajUListu m_E3Obrnuto, UCase$(Trim$(CStr(m_E3(k)))), CStr(k)
    Next k
    Exit Sub
EH:
    m_E3Kvar = "Clanstvo otpremnice u zbirnoj nije citljivo (E3): " & Err.description
    Set m_E3 = Nothing
    Err.Clear
End Sub

Private Sub UcitajE4()
    Set m_E4 = Nothing
    m_E4Kvar = ""
    Set m_E4KvarPoZbirni = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    Set m_E4 = modAutoHladnjaca.AktivnePrijemniceZbirnih(m_E4KvarPoZbirni)
    Exit Sub
EH:
    m_E4Kvar = "Prijemnice zbirnih nisu citljive (E4): " & Err.description
    Set m_E4 = CreateObject("Scripting.Dictionary")
    Set m_E4KvarPoZbirni = CreateObject("Scripting.Dictionary")
    Err.Clear
End Sub

Private Sub UcitajE5()
    Dim i As Long
    m_E5 = Empty
    m_E5Kvar = ""
    Set m_E5PoPrj = CreateObject("Scripting.Dictionary")
    Set m_E5PoStavci = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    m_E5 = modDokumenta.StavkePrijemniceRedovi()
    On Error GoTo 0
    If Not IsArray(m_E5) Then Exit Sub
    For i = 1 To UBound(m_E5, 1)
        DodajUListu m_E5PoPrj, UCase$(Trim$(CStr(m_E5(i, 1)))), i
        m_E5PoStavci(UCase$(Trim$(CStr(m_E5(i, 7))))) = i
    Next i
    Exit Sub
EH:
    m_E5Kvar = "Stavke prijemnica nisu citljive (E5): " & Err.description
    m_E5 = Empty
    Err.Clear
End Sub

Private Sub UcitajE6()
    Dim k As Variant, linija As Variant
    Set m_E6 = Nothing
    m_E6Kvar = ""
    Set m_E6PoFakturi = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    Set m_E6 = modFaktura.FakturaLinijePoStavci()
    On Error GoTo 0
    For Each k In m_E6.Keys
        linija = m_E6(k)
        DodajUListu m_E6PoFakturi, UCase$(Trim$(CStr(linija(0)))), CStr(k)
    Next k
    Exit Sub
EH:
    m_E6Kvar = "Linije faktura nisu citljive (E6): " & Err.description
    Set m_E6 = Nothing
    Err.Clear
End Sub

' GP lanac: zaglavlja prerade i utovara, ivica E9 (prerada <-> utovar, po kg) i
' marker ivice E10 (Utovar.FakturaID). Kvar ovde gasi samo GP deo grafa.
Private Sub UcitajGp()
    Const SRC As String = "modSledljivost.UcitajGp"
    Dim poUtovaru As Object, kvarCitaca As Object, poPreradi As Object, unutra As Object
    Dim k As Variant, v As Variant, lista As Collection, d As Variant
    Dim i As Long, marker As String, pre As String, utv As String

    m_E9Kvar = ""
    m_cUtFaktura = 0
    Set m_E9Zbir = Nothing
    Set m_E9PoPreradi = CreateObject("Scripting.Dictionary")
    Set m_E9PoUtovaru = CreateObject("Scripting.Dictionary")
    Set m_E9KvarPoPreradi = CreateObject("Scripting.Dictionary")
    Set m_E10PoFakturi = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    UcitajZaglavlja SL_PRE, TBL_PRERADA, COL_PRE_ID, COL_PRE_BROJ, COL_PRE_DATUM, ""
    UcitajZaglavlja SL_UTV, TBL_UTOVAR, COL_UT_ID, COL_UT_BROJ, COL_UT_DATUM, ""
    m_cUtFaktura = RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)

    d = m_Data(SL_UTV)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            marker = UCase$(Trim$(NzToText(d(i, m_cUtFaktura))))
            If Len(marker) > 0 Then
                If Not RedStorniran(SL_UTV, i) Then
                    DodajUListu m_E10PoFakturi, marker, UCase$(Trim$(NzToText(d(i, m_cID(SL_UTV)))))
                End If
            End If
        Next i
    End If

    Set m_E9Zbir = modUtovar.UtovarenoPoPreradi(poUtovaru, kvarCitaca)

    ' Ista prerada moze biti na istom utovaru u vise stavki: kg se sabira po paru.
    Set poPreradi = CreateObject("Scripting.Dictionary")
    For Each k In poUtovaru.Keys
        pre = UCase$(Trim$(CStr(k)))
        If poPreradi.Exists(pre) Then
            Set unutra = poPreradi(pre)
        Else
            Set unutra = CreateObject("Scripting.Dictionary")
            poPreradi.Add pre, unutra
        End If
        Set lista = poUtovaru(k)
        For Each v In lista
            utv = UCase$(Trim$(CStr(v(0))))
            If unutra.Exists(utv) Then
                unutra(utv) = CDbl(unutra(utv)) + CDbl(v(1))
            Else
                unutra.Add utv, CDbl(v(1))
            End If
        Next v
    Next k
    For Each k In poPreradi.Keys
        Set unutra = poPreradi(k)
        For Each v In unutra.Keys
            DodajUListu m_E9PoPreradi, CStr(k), Array(CStr(v), CDbl(unutra(v)))
            DodajUListu m_E9PoUtovaru, CStr(v), Array(CStr(k), CDbl(unutra(v)))
        Next v
    Next k
    For Each k In kvarCitaca.Keys
        m_E9KvarPoPreradi(UCase$(Trim$(CStr(k)))) = CStr(kvarCitaca(k))
    Next k
    Exit Sub
EH:
    m_E9Kvar = "GP lanac nije citljiv (E9): " & Err.description
    Err.Clear
End Sub

Private Sub Oslobodi()
    Dim i As Long
    For i = 1 To 7
        m_Data(i) = Empty
        Set m_Idx(i) = Nothing
        Set m_Dupli(i) = Nothing
    Next i
    m_E1 = Empty
    m_E5 = Empty
    Set m_E1PoOtk = Nothing
    Set m_E2 = Nothing
    Set m_E2Obrnuto = Nothing
    Set m_E3 = Nothing
    Set m_E3Obrnuto = Nothing
    Set m_E4 = Nothing
    Set m_E4KvarPoZbirni = Nothing
    Set m_E5PoPrj = Nothing
    Set m_E5PoStavci = Nothing
    Set m_E6 = Nothing
    Set m_E6PoFakturi = Nothing
    Set m_E9Zbir = Nothing
    Set m_E9PoPreradi = Nothing
    Set m_E9PoUtovaru = Nothing
    Set m_E9KvarPoPreradi = Nothing
    Set m_E10PoFakturi = Nothing
    Set m_Redovi = Nothing
    Set m_Vidjen = Nothing
    Set m_Ctx = Nothing
End Sub
