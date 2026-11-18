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
' razlaz, ne osiguranje. V. docs/UI_MIGRACIJA_KATALOG.md 26.5.
'
' STA MODUL NE RADI:
'   - ne crta i ne cita kontrole. Ulaz je RECNIK "kljuc polja -> vrednost", isti
'     obrazac koji ljuska koristi za Scr_Save. Poruku o gresci prikazuje pozivalac;
'     modul u recnik upise "fokus" = kljuc polja koje je odbijeno.
'   - ne racuna sam: cene idu kroz modCenovnik (provere CenaUnosProblem /
'     CenaIspravkaProblem, pisci AddCena / IspraviCenu / StornirajCenu -- v.
'     odeljak CENOVNIK ispod), ID kroz GetNextID, upis
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

Public Const MATUNOS_BUILD As String = "v6-ui-206"

' Kljuc pod kojim modul javlja koje polje je odbijeno. Pozivalac ga koristi da
' vrati fokus tamo gde je greska -- forma je to radila sa SetFocus.
Public Const MAT_FOKUS As String = "fokus"

Private Const SRC As String = "modMaticniUnos"

' Zatvorena kapija upisa -- SAMO za test. Otvaranje ide iskljucivo kroz False,
' koji vraca normalno ponasanje; ne postoji vrednost koja kapiju zaobilazi.
' Postoji zato sto je u headless runu MozeAdministraciju anti-lockout (bez
' AUTH-a su svi admini), pa bi tvrdnja "pisac postuje kapiju" inace merila dva
' puta prolaz i prolazila i kad kapije nema.
Private mBranaZatvorenaTest As Boolean

' Zatvara kapiju upisa za jednu tvrdnju (v. mBranaZatvorenaTest).
Public Sub MatBranaZatvoriTest(ByVal zatvori As Boolean)
    If Not IsTestMode() Then Exit Sub
    mBranaZatvorenaTest = zatvori
End Sub

'--------------------------------------------------------------- UNOS
' Nov zapis. Vraca "" kad je proslo, inace poruku za operatera.
' noviID dobija ID novog zapisa (ili unetu vrednost, gde je PK sam naziv).
' TVRDA KAPIJA UPISA. Stoji u piscu, ne samo u ekranu, jer se do ovih procedura
' stize i iz legacy forme i iz svakog buduceg pozivaoca -- a UI brana pada
' zajedno sa svojim ekranom. Isti obrazac koji vec vazi za ApplyAvansToOtkup i
' IsplataBlokProblem (.claude/rules/testovi.md par.5).
'
' Korisnici traze JOS administraciju: prava pristupa su jedina lista na kojoj se
' pristup moze dodeliti samom sebi.
'
' Vraca "" kad je dozvoljeno, inace poruku za operatera.
Public Function MatBranaUpisa(ByVal kljuc As String) As String
    On Error Resume Next
    ' Test seam: sme SAMO da ZATVORI kapiju, nikad da je otvori. U headless runu
    ' je MozeAdministraciju anti-lockout (svi su admini), pa bi tvrdnja "pisac
    ' postuje kapiju" inace merila dva puta prolaz.
    If mBranaZatvorenaTest And IsTestMode() Then
        MatBranaUpisa = Poruka("MATU_ERR_BEZ_PRAVA")
        Exit Function
    End If
    If Not modAuth.KorisnikImaPravo(OBL_MATICNI) Then
        MatBranaUpisa = Poruka("MATU_ERR_BEZ_PRAVA")
        Exit Function
    End If
    If kljuc = "KORISNICI" Or kljuc = "PRAVA" Then
        If Not modAuth.MozeAdministraciju() Then _
            MatBranaUpisa = Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA")
    End If
    Err.Clear
End Function

Public Function MatDodaj(ByVal kljuc As String, ByVal polja As Object, _
                         ByRef noviID As String) As String
    Dim tbl As String, greska As String, red As Long, tx As clsTransaction
    Dim pk As String, prefiks As String, statKol As String

    On Error GoTo EH
    noviID = ""
    MatDodaj = MatBranaUpisa(kljuc)
    If Len(MatDodaj) > 0 Then Exit Function
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(tbl) = 0 Then
        MatDodaj = Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & kljuc
        Exit Function
    End If

    ' Korisnici imaju SVOG pisca. PIN se hesira, uloga i aktivnost se pisu u
    ' recniku "DA"/"NE" koji cita modAuth, a prava su kolone istog reda -- nista
    ' od toga opsti upis ne zna. V. modMaticniKorisnici i UI_MIGRACIJA_KATALOG
    ' 26.18.
    If kljuc = "KORISNICI" Then
        MatDodaj = modMaticniKorisnici.KorDodaj(polja, noviID)
        Exit Function
    End If

    ' Cenovnik ne ide kroz opsti upis: red se posle unosa ne menja (CEN-02),
    ' a provere i upis drzi modCenovnik (DodajCenu -> AddCena).
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

    ' Ugovor o formatu + redosled kolona PRE upisa. Tabela se razresava u RUNTIME-u
    ' (DodajPrazanRed radi AppendRow(tbl, ...) sa promenljivim imenom), pa je ovaj
    ' upis nevidljiv i za tools/who_writes.py i za svaku staticku pretragu --
    ' kapija na javnom ulazu je jedina odbrana koja ga pokriva.
    modSchema.SchemaReadyOrFail SRC, tbl

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
    MatIzmeni = MatBranaUpisa(kljuc)
    If Len(MatIzmeni) > 0 Then Exit Function
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
        ' Red cenovnika se ne menja na mestu (CEN-02). Pogresan red se
        ' ISPRAVLJA (MatIspraviCenu: storno starog + nov red u jednoj
        ' transakciji, CEN-13) ili STORNIRA (MatStornirajCenu) -- nikad kroz
        ' izmenu celija, pa ovaj put i dalje odbija.
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

    ' PRIRODAN PK SE NE PREIMENUJE. Kod ambalaze, paleta, kutija, kesa i vrste
    ' gotovog proizvoda naziv JESTE kljuc, i na njega se nizvodno pozivaju
    ' kulture, prijemnice i palete po VREDNOSTI, ne po surogatu. Tiha izmena bi
    ' ostavila te reference da pokazuju na ime kog vise nema.
    '
    ' Preimenovanje bi trazilo transakcionu izmenu SVIH referenci; dok toga
    ' nema, polje se odbija sa uputstvom -- nov zapis pa deaktivacija starog.
    greska = PkNepromenjen(kljuc, red, polja)
    If Len(greska) > 0 Then
        MatIzmeni = greska
        Exit Function
    End If

    ' Ugovor o formatu + redosled kolona PRE upisa. Tabela se razresava u RUNTIME-u
    ' (DodajPrazanRed radi AppendRow(tbl, ...) sa promenljivim imenom), pa je ovaj
    ' upis nevidljiv i za tools/who_writes.py i za svaku staticku pretragu --
    ' kapija na javnom ulazu je jedina odbrana koja ga pokriva.
    modSchema.SchemaReadyOrFail SRC, tbl

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
    MatPromeniStatus = MatBranaUpisa(kljuc)
    If Len(MatPromeniStatus) > 0 Then Exit Function
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

    ' Ugovor o formatu + redosled kolona PRE upisa. Tabela se razresava u RUNTIME-u
    ' (DodajPrazanRed radi AppendRow(tbl, ...) sa promenljivim imenom), pa je ovaj
    ' upis nevidljiv i za tools/who_writes.py i za svaku staticku pretragu --
    ' kapija na javnom ulazu je jedina odbrana koja ga pokriva.
    modSchema.SchemaReadyOrFail SRC, tbl

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

'----------------------------------------------------------- CENOVNIK
' Cenovnik ima svoj tok (docs/DOMEN/CENOVNIK.md): red se posle unosa ne menja
' (CEN-02), pogresan se STORNIRA (CEN-11, CEN-12), a ISPRAVKA je storno starog
' i nov red u jednoj transakciji (CEN-13). Pravila i upis drzi modCenovnik;
' ovde je samo prevod recnika polja editora u argumente, i kljuca problema u
' polje za fokus. Pravo je postojece MaticniPodaci kroz MatBranaUpisa -- bez
' novog prava (odluka vlasnika).
'
' Recnik nosi kljuceve polja editora: vrsta, sorta, klasa, stanica, datum,
' cena. Stanica stize u PRIKAZ obliku iz combo-a ("Naziv (ST-xxx)") ili
' prazna (= globalna cena).

' Provera unosa NOVE cene. "" ili poruka operateru, uz polja(MAT_FOKUS) = polje
' koje ne valja. Ista provera koju DodajCenu pusta pre AddCena, a AddCena jos
' jednom pre upisa -- jedna implementacija (CenaUnosProblem), vise pozivalaca.
' Greska citanja je fail-closed: poruka, ne "prolazi".
Public Function MatCenaProblem(ByVal polja As Object) As String
    Dim dat As Date, vrsta As String, sorta As String, klasa As String
    Dim cena As Double, stanicaID As String, tx As clsTransaction
    On Error GoTo EH
    MatCenaProblem = CenaUnosOdbija(polja, dat, vrsta, sorta, klasa, cena, stanicaID)
    Exit Function
EH:
    MatCenaProblem = OdustaniUzGresku(tx, "MatCenaProblem", "CENOVNIK")
End Function

' Pregled pre upisa GLOBALNE cene (ADR 0019, odluka vlasnika): spisak stanica
' cija ranija POSEBNA cena od tog dana prestaje, ili "". Posebna cena (stanica
' izabrana) ne prekida nista, pa je odgovor "". izuzmiID je red koji ispravka
' stornira u istoj transakciji. Zove se POSLE provere; recnik koji ne prolazi
' proveru nema sta da pokaze ("").
'
' Greska citanja se NAMERNO ne guta: "" bi znacilo "nista ne prestaje", pa bi
' se globalna cena upisala bez pitanja. Ekran je pod branom Scr_Event (KI-016),
' pa pad zaustavlja potez i nista se ne upisuje.
Public Function MatCenaPregled(ByVal polja As Object, _
                               Optional ByVal izuzmiID As String = "") As String
    Dim dat As Date, vrsta As String, sorta As String, klasa As String
    Dim cena As Double, stanicaID As String
    If Len(CenaIzPolja(polja, dat, vrsta, sorta, klasa, cena, stanicaID, _
                       Len(izuzmiID) > 0)) > 0 Then Exit Function
    If Len(stanicaID) > 0 Then Exit Function
    MatCenaPregled = modCenovnik.PrekinutePosebneCene(dat, vrsta, sorta, klasa, izuzmiID)
End Function

' Provera ISPRAVKE reda staraID (CEN-13): "" ili poruka, uz fokus kao kod
' unosa. Ispravka bez ijedne promene se odbija PRE storna, a ispravka koja
' zadrzava stari Datum nije pod granicom datuma (CEN-09) -- oba pravila drzi
' modCenovnik.CenaIspravkaProblem.
Public Function MatIspravkaCeneProblem(ByVal staraID As String, ByVal polja As Object) As String
    Dim dat As Date, vrsta As String, sorta As String, klasa As String
    Dim cena As Double, stanicaID As String, k As String, tx As clsTransaction
    On Error GoTo EH
    MatIspravkaCeneProblem = CenaIzPolja(polja, dat, vrsta, sorta, klasa, cena, stanicaID, True)
    If Len(MatIspravkaCeneProblem) > 0 Then Exit Function
    k = modCenovnik.CenaIspravkaProblem(staraID, dat, vrsta, sorta, klasa, cena, stanicaID)
    If Len(k) > 0 Then MatIspravkaCeneProblem = CenaOdbij(polja, k, sorta)
    Exit Function
EH:
    MatIspravkaCeneProblem = OdustaniUzGresku(tx, "MatIspravkaCeneProblem", "CENOVNIK")
End Function

' Ispravka reda staraID (CEN-13) u modCenovnik.IspraviCenu: storno starog i nov
' red u JEDNOJ transakciji, veza u oba smera. "" = uspeh i noviID = ID novog
' reda; inace poruka i noviID prazan. Pisac sam ponovo pusta sve provere
' (razlog, red, kljuc dana, granica), pa poziv mimo ekrana ne zaobilazi nijednu.
Public Function MatIspraviCenu(ByVal staraID As String, ByVal polja As Object, _
                               ByVal razlog As String, ByRef noviID As String) As String
    Dim dat As Date, vrsta As String, sorta As String, klasa As String
    Dim cena As Double, stanicaID As String, tx As clsTransaction
    On Error GoTo EH
    noviID = ""
    MatIspraviCenu = MatBranaUpisa("CENOVNIK")
    If Len(MatIspraviCenu) > 0 Then Exit Function
    MatIspraviCenu = CenaIzPolja(polja, dat, vrsta, sorta, klasa, cena, stanicaID, True)
    If Len(MatIspraviCenu) > 0 Then Exit Function
    MatIspraviCenu = modCenovnik.IspraviCenu(staraID, dat, vrsta, sorta, klasa, cena, _
                                             stanicaID, razlog, noviID)
    ' "" bez novog ID-a nije uspeh -- ugovor pisca kaze da ga uspeh uvek nosi.
    If Len(MatIspraviCenu) = 0 And Len(noviID) = 0 Then _
        MatIspraviCenu = Poruka("CEN_ERR_ISPRAVKA_NIJE_UPISANA")
    Exit Function
EH:
    noviID = ""
    MatIspraviCenu = OdustaniUzGresku(tx, "MatIspraviCenu", "CENOVNIK")
End Function

' Storno reda cenaID (CEN-11, CEN-12): razlog obavezan, storno jednokratan i
' konacan -- pravila drzi modCenovnik.StornirajCenu. "" = uspeh, inace poruka.
Public Function MatStornirajCenu(ByVal cenaID As String, ByVal razlog As String) As String
    Dim tx As clsTransaction
    On Error GoTo EH
    MatStornirajCenu = MatBranaUpisa("CENOVNIK")
    If Len(MatStornirajCenu) > 0 Then Exit Function
    MatStornirajCenu = modCenovnik.StornirajCenu(cenaID, razlog)
    Exit Function
EH:
    MatStornirajCenu = OdustaniUzGresku(tx, "MatStornirajCenu", "CENOVNIK")
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
            If c > 0 Then
                ' Datum ide u editor kao dd.mm.yyyy -- oblik koji TryParseDateValue
                ' cita deterministicki (DMY), nezavisno od regionalnih podesavanja.
                ' CStr(Date) je lokalni kratki datum: "11/18/2026" sa m/d/y masine
                ' se ne cita nazad, pa bi ispravka cene tiho dobila danasnji datum.
                If modMaticniIzvor.PoljeF(CStr(r), 2) = "date" Then
                    v = DatumZaEditor(data(red, c))
                Else
                    v = Trim$(NzToText(data(red, c)))
                End If
            End If
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

        ' COMBO MORA DA BIRA IZ SVOJE LISTE. Kontrole su fmStyleDropDownCombo
        ' (slobodan tekst), pa je do sada prolazio izmisljen StanicaID ili
        ' proizvoljna vrednost enum-a: upis bi uspeo, a strani kljuc ne bi
        ' pokazivao ni na sta. Prazno i dalje prolazi -- obaveznost je zasebna
        ' provera iznad.
        If modMaticniIzvor.PoljeF(spec, 2) = "cmb" And Len(v) > 0 Then
            If Not ComboVrednostPostoji(kljuc, spec, v, polja) Then
                Proveri = Odbij(polja, spec, Poruka("MATU_ERR_VAN_LISTE"))
                Exit Function
            End If
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

' Nova cena kroz modCenovnik -- racun se ne prepisuje ovde. Provera je ista
' kao u MatCenaProblem (CenaUnosOdbija), pa operater dobija tacan razlog uz
' polje; AddCena je pusta jos jednom pre upisa (drugi zid, ista funkcija).
Private Function DodajCenu(ByVal polja As Object, ByRef noviID As String) As String
    Dim dat As Date, vrsta As String, sorta As String, klasa As String
    Dim cena As Double, stanicaID As String, res As String
    DodajCenu = CenaUnosOdbija(polja, dat, vrsta, sorta, klasa, cena, stanicaID)
    If Len(DodajCenu) > 0 Then Exit Function
    res = modCenovnik.AddCena(dat, vrsta, sorta, klasa, cena, stanicaID)
    If Len(res) = 0 Then
        DodajCenu = Poruka("MATU_ERR_CENA_NIJE_UPISANA")
        Exit Function
    End If
    noviID = res
End Function

' Recnik polja editora -> argumenti modCenovnik. "" kad se recnik moze
' procitati, inace poruka uz polja(MAT_FOKUS):
'   1) Proveri -- obaveznost i combo iz svog spiska (i stanica, u PRIKAZ obliku);
'   2) cena je broj (MATU_ERR_BROJ, fokus na cenu);
'   3) prazan ili neispravan datum znaci DANAS pri UNOSU -- isto sto legacy
'      forma radi (PLT-52). ISPRAVKA (datumObavezan) ga odbija sa CEN_ERR_DATUM:
'      AgriX trazi vaziOd za ispravku, a tiho "danas" bi stornirao stari red i
'      preselio cenu na danas, bez povratka za red stariji od granice (CEN-09);
'   4) stanica: prikaz "Naziv (ST-xxx)" -> ID, isti prevod koji opsti upis radi
'      za combo stranog kljuca (IdZaCombo); prazno ostaje prazno (= globalna).
Private Function CenaIzPolja(ByVal polja As Object, ByRef dat As Date, _
                             ByRef vrsta As String, ByRef sorta As String, _
                             ByRef klasa As String, ByRef cena As Double, _
                             ByRef stanicaID As String, _
                             Optional ByVal datumObavezan As Boolean = False) As String
    CenaIzPolja = Proveri("CENOVNIK", polja)
    If Len(CenaIzPolja) > 0 Then Exit Function
    If Not TryParseDouble(Vred(polja, "cena"), cena) Then
        CenaIzPolja = Odbij(polja, modMaticniIzvor.MatPolje("CENOVNIK", "cena"), _
                            Poruka("MATU_ERR_BROJ"))
        Exit Function
    End If
    vrsta = Vred(polja, "vrsta")
    sorta = Vred(polja, "sorta")
    If Not TryParseDateValue(Vred(polja, "datum"), dat) Then
        If datumObavezan Then
            CenaIzPolja = CenaOdbij(polja, "CEN_ERR_DATUM", sorta)
            Exit Function
        End If
        dat = Date
    End If
    klasa = Vred(polja, "klasa")
    stanicaID = IdZaCombo(modMaticniIzvor.MatPolje("CENOVNIK", "stanica"), _
                          Vred(polja, "stanica"))
End Function

' Unos nove cene: recnik, pa pravilo unosa (modCenovnik.CenaUnosProblem).
' Jedno telo za MatCenaProblem (editor) i DodajCenu (pisac).
Private Function CenaUnosOdbija(ByVal polja As Object, ByRef dat As Date, _
                                ByRef vrsta As String, ByRef sorta As String, _
                                ByRef klasa As String, ByRef cena As Double, _
                                ByRef stanicaID As String) As String
    Dim k As String
    CenaUnosOdbija = CenaIzPolja(polja, dat, vrsta, sorta, klasa, cena, stanicaID)
    If Len(CenaUnosOdbija) > 0 Then Exit Function
    k = modCenovnik.CenaUnosProblem(dat, vrsta, sorta, klasa, cena, stanicaID)
    If Len(k) > 0 Then CenaUnosOdbija = CenaOdbij(polja, k, sorta)
End Function

' Kljuc problema iz modCenovnik -> poruka operateru i polje za fokus (ugovor
' cenovnika, sekcija 2): VRSTA_OBAVEZNA -> vrsta; KULTURA_NEPOZNATA -> sorta
' kad je sorta data, inace vrsta; KLASA -> klasa; CENA_* -> cena;
' STANICA_NEPOZNATA -> stanica; DATUM, DATUM_GRANICA, DUPLIKAT -> datum. Kljuc
' bez svog polja (red ne postoji, vec storniran, ispravka bez promene) ostaje
' bez fokusa. Poruka je sam tekst kljuca: CEN_ERR_* su pune recenice koje
' imenuju polje, isti tekst koji vracaju i StornirajCenu / IspraviCenu.
Private Function CenaOdbij(ByVal polja As Object, ByVal kljucProblema As String, _
                           ByVal sorta As String) As String
    Dim poljeKljuc As String
    On Error Resume Next
    Select Case kljucProblema
        Case "CEN_ERR_VRSTA_OBAVEZNA"
            poljeKljuc = "vrsta"
        Case "CEN_ERR_KULTURA_NEPOZNATA"
            If Len(Trim$(sorta)) > 0 Then
                poljeKljuc = "sorta"
            Else
                poljeKljuc = "vrsta"
            End If
        Case "CEN_ERR_KLASA"
            poljeKljuc = "klasa"
        Case "CEN_ERR_CENA_POZITIVNA", "CEN_ERR_CENA_OPSEG", "CEN_ERR_CENA_DECIMALE"
            poljeKljuc = "cena"
        Case "CEN_ERR_STANICA_NEPOZNATA"
            poljeKljuc = "stanica"
        Case "CEN_ERR_DATUM", "CEN_ERR_DATUM_GRANICA", "CEN_ERR_DUPLIKAT"
            poljeKljuc = "datum"
    End Select
    If Len(poljeKljuc) > 0 Then polja(MAT_FOKUS) = poljeKljuc
    CenaOdbij = Poruka(kljucProblema)
End Function

' Celija datuma -> tekst za editor, dd.mm.yyyy (v. MatVrednostiReda). Vrednost
' koja nije citljiv datum ide kako jeste: editor ne sme da izmisli datum.
Private Function DatumZaEditor(ByVal x As Variant) As String
    Dim d As Date
    On Error GoTo Sirovo
    Select Case VarType(x)
        Case vbDate
            DatumZaEditor = Format$(x, "dd.mm.yyyy")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            If CDbl(x) >= 1 Then DatumZaEditor = Format$(CDate(CDbl(x)), "dd.mm.yyyy")
        Case Else
            If TryParseDateValue(NzToText(x), d) Then
                DatumZaEditor = Format$(d, "dd.mm.yyyy")
            Else
                DatumZaEditor = Trim$(NzToText(x))
            End If
    End Select
    Exit Function
Sirovo:
    DatumZaEditor = Trim$(NzToText(x))
End Function

' Prva vrednost sekcije bez surogat kljuca (tip ambalaze, palete, kutije, kese,
' gotovog proizvoda) -- ona JESTE PK.
' Da li dati tekst postoji u spisku tog combo-a. Kontekst zavisnog izvora se
' cita iz DRUGOG polja istog recnika (sorte zavise od vrste) -- isti spisak
' zavisnosti koji koristi i editor, pa se provera i ponuda ne mogu razici.
'
' Nepoznat izvor NE odbija: spisak koji se ne moze izgraditi nije dokaz da je
' vrednost pogresna, a lazno odbijanje bi zaustavilo upis koji je ispravan.
' "" kad je PK netaknut ili sekcija ima surogat kljuc; inace poruka.
Private Function PkNepromenjen(ByVal kljuc As String, ByVal red As Long, _
                               ByVal polja As Object) As String
    Dim pk As String, tbl As String, spec As String, poljeKljuc As String
    Dim staro As String, novo As String, c As Long, data As Variant
    On Error GoTo EH
    ' Sekcija sa prefiksom ima surogat kljuc (KOOP-001) koji operater i ne
    ' unosi -- tamo nema sta da se stiti.
    If Len(modMaticniIzvor.MatPrefiksID(kljuc)) > 0 Then Exit Function
    pk = modMaticniIzvor.MatPK(kljuc)
    tbl = modMaticniIzvor.MatTabela(kljuc)
    If Len(pk) = 0 Or Len(tbl) = 0 Then Exit Function

    spec = PoljeZaKolonu(kljuc, pk)
    If Len(spec) = 0 Then Exit Function
    poljeKljuc = modMaticniIzvor.PoljeF(spec, 0)
    novo = Vred(polja, poljeKljuc)
    If Len(novo) = 0 Then Exit Function

    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    c = GetColumnIndex(tbl, pk)
    If c = 0 Then Exit Function
    staro = Trim$(NzToText(data(red, c)))
    If StrComp(staro, novo, vbTextCompare) <> 0 Then
        polja(MAT_FOKUS) = poljeKljuc
        PkNepromenjen = Poruka("MATU_ERR_PK_ZAKLJUCAN")
    End If
    Exit Function
EH:
    LogErr SRC & ".PkNepromenjen"
End Function

' Opis polja koje pise u datu kolonu, ili "" ako takvog nema.
Private Function PoljeZaKolonu(ByVal kljuc As String, ByVal kolona As String) As String
    Dim a As Variant, r As Variant
    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        If StrComp(modMaticniIzvor.MatKolonaPolja(kljuc, CStr(r)), kolona, vbTextCompare) = 0 Then
            PoljeZaKolonu = CStr(r)
            Exit Function
        End If
    Next r
End Function

Private Function ComboVrednostPostoji(ByVal kljuc As String, ByVal spec As String, _
                                      ByVal v As String, ByVal polja As Object) As Boolean
    Dim izvor As String, zavisi As String, kontekst As String
    Dim stavke As Variant, i As Long
    On Error GoTo EH
    ComboVrednostPostoji = True
    izvor = modMaticniIzvor.PoljeF(spec, 5)
    If Len(izvor) = 0 Then Exit Function

    zavisi = modMaticniIzvor.MatComboZavisi(izvor)
    If Len(zavisi) > 0 Then kontekst = Vred(polja, zavisi)
    stavke = modMaticniIzvor.MatComboStavke(izvor, kontekst)
    ' PUKLO citanje nije prazan spisak. MatComboStavke i jedno i drugo vraca kao
    ' prazan niz (combo se ne sme srusiti), pa se razlika cita odvojeno --
    ' inace bi kvar pri citanju tblStanice bio "provera je prosla".
    If Len(modMaticniIzvor.MatComboGreska()) > 0 Then
        LogWarn "modMaticniUnos.ComboVrednostPostoji", _
                "spisak se ne cita (" & modMaticniIzvor.MatComboGreska() & _
                ") -- vrednost odbijena"
        ComboVrednostPostoji = False
        Exit Function
    End If
    ' PRAZAN SPISAK ODBIJA. Do v6-ui-205 je prolazio, uz obrazlozenje "prvi
    ' zapis u sekciji nema sta da ponudi" -- a to obrazlozenje je bilo pogresno:
    ' ovamo se dolazi SAMO za NEPRAZNU vrednost (v. Proveri, grana cmb). Prazno
    ' polje je vec proslo granom obaveznosti iznad i nikad ne stigne ovde.
    '
    ' Zato prazan spisak znaci "nema cega da bude" -- a ne "ne mogu da proverim".
    ' GetTableData za nepostojecu ili praznu tabelu vraca Empty BEZ greske, pa
    ' se taj slucaj ne vidi kroz MatComboGreska: bez ovoga bi nepostojeca
    ' tblStanice pustala proizvoljan tekst pravo u StanicaID.
    If Not IsArray(stavke) Then
        ComboVrednostPostoji = False
        Exit Function
    End If
    If UBound(stavke) < LBound(stavke) Then
        ComboVrednostPostoji = False
        Exit Function
    End If

    For i = LBound(stavke) To UBound(stavke)
        If StrComp(Trim$(CStr(stavke(i))), Trim$(v), vbTextCompare) = 0 Then Exit Function
    Next i
    ComboVrednostPostoji = False
    Exit Function
EH:
    ' Citanje spiska je PUKLO. Do v6-ui-203 se to citalo kao "prolazi" -- greska
    ' u proveri je postajala odobrenje, pa je proizvoljan tekst mogao da udje u
    ' kolonu stranog kljuca. Sada je fail-closed, kao i brana ekrana: neuspela
    ' provera nije prosla provera.
    '
    ' Od v6-ui-205 ni prazan spisak vise ne prolazi (v. telo funkcije): ovamo se
    ' dolazi samo za NEPRAZNU vrednost, pa prazan spisak znaci "nema cega da
    ' bude". Ovaj put ostaje za greske koje se dignu U SAMOJ proveri.
    LogWarn "modMaticniUnos.ComboVrednostPostoji", _
            "spisak za '" & izvor & "' se ne cita (" & Err.Number & " " & _
            Err.description & ") -- vrednost odbijena"
    ComboVrednostPostoji = False
End Function

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
Public Function MatPoljeZaKolonuTest(ByVal kljuc As String, ByVal kolona As String) As String
    MatPoljeZaKolonuTest = PoljeZaKolonu(kljuc, kolona)
End Function

Public Function MatKolonaZaPoljeTest(ByVal kljuc As String, ByVal poljeKljuc As String) As String
    MatKolonaZaPoljeTest = modMaticniIzvor.MatKolonaPolja(kljuc, _
                                modMaticniIzvor.MatPolje(kljuc, poljeKljuc))
End Function
