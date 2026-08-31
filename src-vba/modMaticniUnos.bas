Attribute VB_Name = "modMaticniUnos"
'=====================================================================
' modMaticniUnos - JEDAN pisac maticnih podataka (korak M2).
'
' Provere i upis su izdvojeni iz frmStammdaten.btnDodaj_Click (544 linije),
' btnIzmeni_Click (463) i OnSoftDeleteClick. Forma je prevezana NA OVAJ MODUL:
' od sada oba UI-ja rade isti upis, kroz isti kod.
'
' ZASTO OVDE NIJE PONOVLJENA ODLUKA IZ FAZE B (dve kopije namerno). Za dokumente
' su postojala dva puta -- operativan legacy i neproveren nov -- pa je kopija
' bila osiguranje. Kod maticnih podataka postoji SAMO JEDAN put (forma), tabele
' su povrsina sinhronizacije (modStammdatenSync, modMasterSync), a WHO_WRITES
' pokazuje 0-2 modula po tabeli. Druga kopija bi ovde bila prva prilika za
' razlaz, ne osiguranje. V. docs/UI_MIGRACIJA_KATALOG.md 24.5.
'
' STA MODUL NE RADI:
'   - ne crta i ne cita kontrole. Ulaz je RECNIK "kljuc polja -> vrednost", isti
'     obrazac koji ljuska koristi za Scr_Save. Poruku o gresci prikazuje pozivalac;
'     modul u recnik upise "fokus" = kljuc polja koje je odbijeno.
'   - ne racuna sam: cene idu kroz modCenovnik.AddCena, ID kroz GetNextID, upis
'     kroz AppendRow/RequireUpdateCell u clsTransaction, MALINA ogledalo kroz
'     modMalina -- sve postojece rutine, iste koje je forma zvala.
'   - NE dira Korisnike. tblKorisnici nosi matricu prava i PreparePin, i ide u
'     M4 zajedno sa svojim ekranom; do tada forma zadrzava svoju granu za njih.
'
' JEDNA NAMERNA IZMENA PONASANJA: unos ide PO IMENU KOLONE, ne pozicijski.
' Legacy je za deset sekcija zvao AppendRow sa nizom vrednosti, pa je tacnost
' zavisila od redosleda kolona u tabeli -- a sema se razlikuje po instalaciji
' (bas zato LoadList za Kupce vec ima toleranciju, a btnIzmeni za Stanice
' alias-probe). Isti obrazac koji legacy vec koristi za Korisnike i Kulture:
' prazan red pa RequireUpdateCell po imenu, sve u transakciji. Upisuju se ISTE
' celije; menja se samo kako se adresiraju.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATUNOS_BUILD As String = "v6-ui-193"

' Kljuc pod kojim modul javlja koje polje je odbijeno. Pozivalac ga koristi da
' vrati fokus tamo gde je greska -- forma je to radila sa SetFocus.
Public Const MAT_FOKUS As String = "fokus"

Private Const SRC As String = "modMaticniUnos"

'--------------------------------------------------------------- UNOS
' Nov zapis. Vraca "" kad je proslo, inace poruku za operatera.
' noviID dobija ID novog zapisa (ili unetu vrednost, gde je PK sam naziv).
Public Function MatDodaj(ByVal kljuc As String, ByVal polja As Object, _
                         ByRef noviID As String) As String
    Dim tbl As String, greska As String, red As Long, tx As clsTransaction
    Dim pk As String, prefiks As String, statKol As String

    On Error GoTo EH
    noviID = ""
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(tbl) = 0 Then
        MatDodaj = Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & kljuc
        Exit Function
    End If

    ' Korisnici imaju SVOG pisca. PIN se hesira, uloga i aktivnost se pisu u
    ' recniku "DA"/"NE" koji cita modAuth, a prava su kolone istog reda -- nista
    ' od toga opsti upis ne zna. V. modMaticniKorisnici i UI_MIGRACIJA_KATALOG
    ' 24.18.
    If kljuc = "KORISNICI" Then
        MatDodaj = modMaticniKorisnici.KorDodaj(polja, noviID)
        Exit Function
    End If

    ' Cenovnik je append-only i ne ide kroz opsti upis: nova cena je nov vazeci
    ' red, a racuna ga modCenovnik. Ista rutina koju je forma zvala.
    If kljuc = "CENOVNIK" Then
        MatDodaj = DodajCenu(polja, noviID)
        Exit Function
    End If

    greska = Proveri(kljuc, polja)
    If Len(greska) > 0 Then
        MatDodaj = greska
        Exit Function
    End If

    pk = modMaticniIzvor.MatPK(kljuc)
    prefiks = modMaticniIzvor.MatPrefiksID(kljuc)
    If Len(prefiks) > 0 Then
        noviID = GetNextID(tbl, pk, prefiks)
    Else
        ' Sekcija bez surogat kljuca: PK je sama uneta vrednost, pa duplikat
        ' mora da se odbije PRE upisa -- inace bi dva tipa ambalaze istog imena
        ' nizvodno znacila dve razlicite tezine za istu gajbicu.
        noviID = PrvaVrednost(kljuc, polja)
        If PostojiPK(tbl, pk, noviID) Then
            polja(MAT_FOKUS) = modMaticniIzvor.PoljeF(CStr(modMaticniIzvor.MatPolja(kljuc)(0)), 0)
            MatDodaj = Poruka("MATU_ERR_VEC_POSTOJI") & " " & noviID
            Exit Function
        End If
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot tbl

    red = DodajPrazanRed(tbl)
    If red = 0 Then Err.Raise vbObjectError + 9600, SRC, "AppendRow nije uspeo."

    RequireUpdateCell tbl, red, pk, noviID, SRC
    UpisiPolja kljuc, tbl, red, polja, True

    ' Status na unosu: v. modMaticniIzvor.MatStatusNaUnosu (Parcele su "Da").
    statKol = modMaticniIzvor.MatStatusKolona(kljuc)
    If Len(statKol) > 0 Then _
        RequireUpdateCell tbl, red, statKol, modMaticniIzvor.MatStatusNaUnosu(kljuc), SRC

    tx.CommitTx
    Set tx = Nothing

    ' MALINA: nova stanica dobija par-vozaca sa istim ID-em (izvestaji/ambalaza).
    ' Idempotentno i self-gated u modMalina; NE sme da obori unos stanice --
    ' zato je van transakcije i pod sopstvenim On Error, tacno kao u legacy formi.
    If kljuc = "STANICE" Then
        If IsMalinaMode() Then
            On Error Resume Next
            EnsureVozacMirrorForStanica noviID, Vred(polja, "naziv"), Vred(polja, "mesto"), ""
            Err.Clear
            On Error GoTo EH
        End If
    End If
    Exit Function
EH:
    MatDodaj = OdustaniUzGresku(tx, "MatDodaj", kljuc)
End Function

'------------------------------------------------------------- IZMENA
' Izmena postojeceg reda. red je redni broj U TABELI (ne u prikazu) -- pozivalac
' ga dobija iz MatRedPoID, dakle po IDENTITETU, ne po poziciji u listi.
Public Function MatIzmeni(ByVal kljuc As String, ByVal red As Long, _
                          ByVal polja As Object) As String
    Dim tbl As String, greska As String, tx As clsTransaction

    On Error GoTo EH
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(tbl) = 0 Then
        MatIzmeni = Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & kljuc
        Exit Function
    End If
    If kljuc = "KORISNICI" Then
        MatIzmeni = modMaticniKorisnici.KorIzmeni(red, polja)
        Exit Function
    End If
    If kljuc = "CENOVNIK" Then
        ' Istorija cena se ne menja -- nova cena je nov red. Isto sto legacy
        ' forma kaze, i zato je "Izmeni" tamo sakriveno za Cenovnik.
        MatIzmeni = Poruka("MATU_ERR_CENOVNIK_APPEND")
        Exit Function
    End If
    If red < 1 Then
        MatIzmeni = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If

    greska = Proveri(kljuc, polja)
    If Len(greska) > 0 Then
        MatIzmeni = greska
        Exit Function
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot tbl

    UpisiPolja kljuc, tbl, red, polja, False

    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    MatIzmeni = OdustaniUzGresku(tx, "MatIzmeni", kljuc)
End Function

'-------------------------------------------------------- SOFT-DELETE
' Obrce status reda. noviStatus dobija vrednost koja je upisana.
' Sekcija bez kolone statusa se ODBIJA umesto da tiho ne uradi nista.
Public Function MatPromeniStatus(ByVal kljuc As String, ByVal red As Long, _
                                 ByRef noviStatus As String) As String
    Dim tbl As String, statKol As String, data As Variant, c As Long
    Dim tx As clsTransaction, cur As String

    On Error GoTo EH
    noviStatus = ""
    ' Kolona Aktivan u tblKorisnici NIJE obicna kolona statusa: modAuth
    ' neaktivnim smatra samo "NE", pa bi opsti upis ovde napisao "Neaktivan" i
    ' korisnik bi se i dalje prijavljivao. Zato ide kroz svog pisca.
    If kljuc = "KORISNICI" Then
        MatPromeniStatus = modMaticniKorisnici.KorPromeniStatus(red, noviStatus)
        Exit Function
    End If
    tbl = modMaticniIzvor.MatTabela(kljuc)
    statKol = modMaticniIzvor.MatStatusKolona(kljuc)
    If Len(tbl) = 0 Or Len(statKol) = 0 Then
        MatPromeniStatus = Poruka("MATU_ERR_NEMA_STATUSA")
        Exit Function
    End If
    If red < 1 Then
        MatPromeniStatus = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If

    data = GetTableData(tbl)
    If IsEmpty(data) Then
        MatPromeniStatus = Poruka("MATU_ERR_NEMA_REDA")
        Exit Function
    End If
    c = GetColumnIndex(tbl, statKol)
    cur = Trim$(NzToText(data(red, c)))
    If StrComp(cur, STATUS_NEAKTIVAN, vbTextCompare) = 0 Then
        noviStatus = STATUS_AKTIVAN
    Else
        noviStatus = STATUS_NEAKTIVAN
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot tbl
    RequireUpdateCell tbl, red, statKol, noviStatus, SRC
    tx.CommitTx
    Set tx = Nothing
    Exit Function
EH:
    MatPromeniStatus = OdustaniUzGresku(tx, "MatPromeniStatus", kljuc)
End Function

'------------------------------------------------------------ POMOCNE
' Red u tabeli po PK. Ovo je jedino mesto na kom se red bira -- i bira se po
' IDENTITETU. Legacy je birao po poziciji u listboxu (m_RowMap); u mrezi koja se
' sortira i pretrazuje pozicija ne znaci nista, a istoimeni zapisi su u
' sifarnicima obicna pojava.
Public Function MatRedPoID(ByVal kljuc As String, ByVal id As String) As Long
    Dim tbl As String, data As Variant, c As Long, i As Long
    On Error GoTo EH
    id = Trim$(id)
    If Len(id) = 0 Then Exit Function
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(tbl) = 0 Then Exit Function
    c = GetColumnIndex(tbl, modMaticniIzvor.MatPK(kljuc))
    If c = 0 Then Exit Function
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, c))), id, vbTextCompare) = 0 Then
            MatRedPoID = i
            Exit Function
        End If
    Next i
    Exit Function
EH:
    LogErr SRC & ".MatRedPoID"
End Function

' Vrednosti postojeceg reda pod kljucevima polja -- za punjenje editora.
' Combo polja dobijaju PRIKAZ (naziv, ne ID), isto sto legacy puni u combo.
Public Function MatVrednostiReda(ByVal kljuc As String, ByVal red As Long) As Object
    Dim d As Object, tbl As String, data As Variant, a As Variant, r As Variant
    Dim kol As String, c As Long, v As String
    Set d = CreateObject("Scripting.Dictionary")
    Set MatVrednostiReda = d
    On Error GoTo EH
    ' PIN se ne vraca u editor (hes), a stanica se vraca kao naziv -- oba
    ' pravila zna modMaticniKorisnici.
    If kljuc = "KORISNICI" Then
        Set MatVrednostiReda = modMaticniKorisnici.KorVrednostiReda(red)
        Exit Function
    End If
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(tbl) = 0 Or red < 1 Then Exit Function
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        kol = modMaticniIzvor.MatKolonaPolja(kljuc, CStr(r))
        v = ""
        If Len(kol) > 0 Then
            c = GetColumnIndex(tbl, kol)
            If c > 0 Then v = Trim$(NzToText(data(red, c)))
        End If
        ' Strani kljuc se u editoru vidi kao prikaz, ne kao ID -- inace bi
        ' operater u polju "Stanica" video ST-003.
        If Len(v) > 0 Then v = PrikazZaCombo(CStr(r), v)
        d(modMaticniIzvor.PoljeF(CStr(r), 0)) = v
    Next r
    Exit Function
EH:
    LogErr SRC & ".MatVrednostiReda"
End Function

'---------------------------------------------------------- UNUTRASNJE
' Provera SVIH polja sekcije. Vraca "" kad je sve u redu; inace poruku i upisuje
' MAT_FOKUS u recnik. Redosled je redosled polja -- prva greska zaustavlja, isto
' kao u formi.
Private Function Proveri(ByVal kljuc As String, ByVal polja As Object) As String
    Dim a As Variant, r As Variant, spec As String
    Dim v As String, d As Double, upoz As Double, blok As Double

    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then
        Proveri = Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & kljuc
        Exit Function
    End If

    For Each r In a
        spec = CStr(r)
        v = Vred(polja, modMaticniIzvor.PoljeF(spec, 0))

        If modMaticniIzvor.PoljeF(spec, 3) = "1" And Len(v) = 0 Then
            Proveri = Odbij(polja, spec, Poruka("MATU_ERR_OBAVEZNO"))
            Exit Function
        End If

        If modMaticniIzvor.PoljeF(spec, 2) = "num" And Len(v) > 0 Then
            If Not TryParseDouble(v, d) Then
                Proveri = Odbij(polja, spec, Poruka("MATU_ERR_BROJ"))
                Exit Function
            End If
            ' Povrsina i cena moraju biti STROGO pozitivne -- nula parcela i
            ' nula dinara nisu podatak nego prazan unos. Ostali brojevi smeju
            ' biti nula (tezina prazne kese, doza koja se ne prati).
            If d < 0 Then
                Proveri = Odbij(polja, spec, Poruka("MATU_ERR_NEGATIVNO"))
                Exit Function
            End If
            If d = 0 And TraziPozitivan(kljuc, modMaticniIzvor.PoljeF(spec, 0)) Then
                Proveri = Odbij(polja, spec, Poruka("MATU_ERR_POZITIVNO"))
                Exit Function
            End If
        End If
    Next r

    ' Ukrstena provera kultura: prag blokade ne sme biti ispod praga upozorenja.
    ' Jedino pravilo koje gleda DVA polja odjednom, pa stoji van petlje.
    If kljuc = "KULTURE" Then
        If TryParseDouble(Vred(polja, "pragupoz"), upoz) Then
            If TryParseDouble(Vred(polja, "pragblok"), blok) Then
                If upoz > 0 And blok > 0 And blok < upoz Then
                    Proveri = Odbij(polja, modMaticniIzvor.MatPolje(kljuc, "pragblok"), _
                                    Poruka("MATU_ERR_PRAG_BLOK"))
                    Exit Function
                End If
            End If
        End If
    End If
End Function

' Polja kod kojih nula NIJE podatak. Preslikano iz legacy provera (povrsina > 0,
' cena cenovnika > 0); sve ostalo je tamo dozvoljavalo nulu.
Private Function TraziPozitivan(ByVal kljuc As String, ByVal poljeKljuc As String) As Boolean
    If kljuc = "PARCELE" And poljeKljuc = "povrsina" Then TraziPozitivan = True
    If kljuc = "CENOVNIK" And poljeKljuc = "cena" Then TraziPozitivan = True
End Function

Private Function Odbij(ByVal polja As Object, ByVal spec As String, _
                       ByVal razlog As String) As String
    On Error Resume Next
    polja(MAT_FOKUS) = modMaticniIzvor.PoljeF(spec, 0)
    Odbij = Poruka(modMaticniIzvor.PoljeF(spec, 1)) & ": " & razlog
End Function

' Upis svih polja sekcije u dati red, PO IMENU KOLONE.
' Polje cija kolona ne postoji u semi se PRESKACE -- drift ne sme da obori ceo
' upis, isto pravilo koje LoadList vec primenjuje na citanju.
Private Sub UpisiPolja(ByVal kljuc As String, ByVal tbl As String, ByVal red As Long, _
                       ByVal polja As Object, ByVal jeUnos As Boolean)
    Dim a As Variant, r As Variant, spec As String, kol As String
    Dim v As String, d As Double
    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then Exit Sub
    For Each r In a
        spec = CStr(r)
        kol = modMaticniIzvor.MatKolonaPolja(kljuc, spec)
        If Len(kol) > 0 Then
            v = Vred(polja, modMaticniIzvor.PoljeF(spec, 0))
            Select Case modMaticniIzvor.PoljeF(spec, 2)
                Case "num"
                    ' Prazan broj: pri UNOSU celija ostaje prazna, pri IZMENI se
                    ' upisuje nula. To je zateceno ponasanje forme (btnDodaj za
                    ' Kulture preskace praznu gajbicu, btnIzmeni upisuje 0) i
                    ' ovde se namerno CUVA -- poravnanje bi bilo neizmerena
                    ' izmena na jedinom operativnom putu upisa.
                    If Len(v) = 0 Then
                        If Not jeUnos Then RequireUpdateCell tbl, red, kol, 0, SRC
                    ElseIf TryParseDouble(v, d) Then
                        RequireUpdateCell tbl, red, kol, d, SRC
                    End If
                Case "cmb"
                    ' Combo koji nosi STRANI KLJUC pise ID, ne prikaz.
                    RequireUpdateCell tbl, red, kol, IdZaCombo(spec, v), SRC
                Case Else
                    RequireUpdateCell tbl, red, kol, v, SRC
            End Select
        End If
    Next r
End Sub

' Prikaz iz combo-a -> vrednost koja ide u tabelu.
' Stanica i kooperant se biraju po nazivu, a cuvaju kao ID; ostali combo-i su
' obicne liste vrednosti i pisu se kako su izabrani.
Private Function IdZaCombo(ByVal spec As String, ByVal prikaz As String) As String
    Dim izvor As String, id As String
    izvor = modMaticniIzvor.PoljeF(spec, 5)
    IdZaCombo = Trim$(prikaz)
    If Len(IdZaCombo) = 0 Then Exit Function

    Select Case izvor
        Case "@stanice"
            ' Prikaz je "Naziv (ST-xxx)" ili goli naziv -- oba oblika postoje u
            ' zatecenom kodu, pa se prvo pokusava ID iz zagrade, pa pretraga po
            ' nazivu. Isti redosled koji legacy forma ima.
            id = ExtractIDFromDisplay(IdZaCombo)
            If Len(id) = 0 Or InStr(1, id, "ST-", vbTextCompare) = 0 Then _
                id = Trim$(CStr(LookupValue(TBL_STANICE, "Naziv", IdZaCombo, "StanicaID")))
            If Len(id) > 0 Then IdZaCombo = id
        Case "@kooperanti"
            id = ExtractIDFromDisplay(IdZaCombo)
            If Len(id) > 0 Then IdZaCombo = id
    End Select
End Function

' Obrnut smer: vrednost iz tabele -> ono sto operater vidi u combo-u.
Private Function PrikazZaCombo(ByVal spec As String, ByVal v As String) As String
    Dim ime As String, prez As String
    PrikazZaCombo = v
    Select Case modMaticniIzvor.PoljeF(spec, 5)
        Case "@stanice"
            ime = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", v, "Naziv")))
            If Len(ime) > 0 Then PrikazZaCombo = ime & " (" & v & ")"
        Case "@kooperanti"
            ime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", v, "Ime")))
            prez = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", v, "Prezime")))
            If Len(ime & prez) > 0 Then _
                PrikazZaCombo = Trim$(ime & " " & prez) & " (" & v & ")"
    End Select
End Function

' Nova cena kroz modCenovnik -- racun se ne prepisuje ovde.
Private Function DodajCenu(ByVal polja As Object, ByRef noviID As String) As String
    Dim cena As Double, dat As Date, res As String, greska As String
    greska = Proveri("CENOVNIK", polja)
    If Len(greska) > 0 Then
        DodajCenu = greska
        Exit Function
    End If
    If Not TryParseDouble(Vred(polja, "cena"), cena) Then
        DodajCenu = Poruka("MATU_ERR_BROJ")
        Exit Function
    End If
    ' Prazan ili neispravan datum znaci DANAS -- isto sto legacy forma radi.
    If Not TryParseDateValue(Vred(polja, "datum"), dat) Then dat = Date
    res = AddCena(dat, Vred(polja, "vrsta"), Vred(polja, "sorta"), _
                  Vred(polja, "klasa"), cena)
    If Len(res) = 0 Then
        DodajCenu = Poruka("MATU_ERR_CENA_NIJE_UPISANA")
        Exit Function
    End If
    noviID = res
End Function

' Prva vrednost sekcije bez surogat kljuca (tip ambalaze, palete, kutije, kese,
' gotovog proizvoda) -- ona JESTE PK.
Private Function PrvaVrednost(ByVal kljuc As String, ByVal polja As Object) As String
    Dim a As Variant
    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then Exit Function
    PrvaVrednost = Vred(polja, modMaticniIzvor.PoljeF(CStr(a(LBound(a))), 0))
End Function

Private Function PostojiPK(ByVal tbl As String, ByVal pk As String, _
                           ByVal v As String) As Boolean
    On Error Resume Next
    PostojiPK = (Len(Trim$(NzToText(LookupValue(tbl, pk, v, pk)))) > 0)
    Err.Clear
End Function

' Prazan red pune sirine tabele. Pozicijski AppendRow je namerno izbegnut -- v.
' zaglavlje modula.
Private Function DodajPrazanRed(ByVal tbl As String) As Long
    Dim prazno() As Variant
    ReDim prazno(1 To GetTable(tbl).ListColumns.count)
    DodajPrazanRed = AppendRow(tbl, prazno)
End Function

Private Function Vred(ByVal polja As Object, ByVal k As String) As String
    On Error Resume Next
    If polja Is Nothing Then Exit Function
    If polja.Exists(k) Then Vred = Trim$(CStr(polja(k)))
    Err.Clear
End Function

' Jedan izlaz za sve greske: rollback, trag u logu, poruka operateru. Err se
' cita PRE rollback-a -- RollbackTx ide kroz On Error Resume Next i obrisao bi ga.
Private Function OdustaniUzGresku(ByRef tx As clsTransaction, ByVal gde As String, _
                                  ByVal kljuc As String) As String
    Dim errDesc As String
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    LogError SRC & "." & gde, kljuc & ": " & errDesc
    Err.Clear
    On Error GoTo 0
    OdustaniUzGresku = Poruka("MATU_ERR_UPIS") & " " & errDesc
End Function

'------------------------------------------------------------ TEST SEAM
' Provera bez upisa. Sabotaza koja obori pravilo pada po imenu, a test ne mora
' da pise u tabele da bi to izmerio.
Public Function MatProveriTest(ByVal kljuc As String, ByVal polja As Object) As String
    MatProveriTest = Proveri(kljuc, polja)
End Function

' Kolona u koju bi dato polje bilo upisano (posle razresenja alias-a).
Public Function MatKolonaZaPoljeTest(ByVal kljuc As String, ByVal poljeKljuc As String) As String
    MatKolonaZaPoljeTest = modMaticniIzvor.MatKolonaPolja(kljuc, _
                                modMaticniIzvor.MatPolje(kljuc, poljeKljuc))
End Function
