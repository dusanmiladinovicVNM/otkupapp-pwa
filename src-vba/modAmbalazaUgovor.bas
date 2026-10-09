Attribute VB_Name = "modAmbalazaUgovor"

Option Explicit

' UGOVOR AMBALAZNE KNJIGE (AMB-10a). Pun model: docs/DOMEN/AMBALAZA.md.
'
' Knjiga je APPEND-ONLY i dogadjaj je PRENOS: jedan red imenuje OBE strane
' (OdNalog -> NaNalog), kolicina je uvek pozitivna, smera kao podatka nema.
' Ovaj modul nosi ono sto uz taj model mora da bude izgovoreno na JEDNOM mestu:
'
'   - zatvorene liste naloga i vrsta kretanja;
'   - razresavanje naloga (AMB-INV-03) -- jedina kapija, fail-closed;
'   - provera prenosa (AMB-INV-01, -02, -03 i granica opticaja);
'   - doprinos obavezi (osnova AMB-INV-09), storno-svestan.
'
' NISTA OVDE NE PISE U TABELE i nijedan produkcioni put ga jos ne zove. To je
' namerno: 10a je ugovor, cutover je 10b. Modul zato sme da se testira i obara
' sabotazom pre nego sto ijedan pisac zavisi od njega.
'
' JEDNA IMPLEMENTACIJA, DVA POZIVAOCA (.claude/rules/testovi.md 5): *Problem
' vraca tekst i koristi ga ekran, Require* dize gresku i koristi je pisac. Druga
' kopija pravila bi se razisla prvom izmenom.

' ============================================================
' NALOZI -- ko moze da drzi gajbe
' ============================================================
'
' Stvarni drzaoci imaju svoj master zapis. Firma i SpoljniSvet su SISTEMSKI: po
' jedan jedini nalog, bez ID-a.
'
' SpoljniSvet NIJE partner nego GRANICA opticaja: ambalaza nastaje u knjizi samo
' kao SpoljniSvet -> neko, a izlazi iz opticaja samo kao neko -> SpoljniSvet.
Public Const AMB_NALOG_KOOPERANT As String = "Kooperant"
Public Const AMB_NALOG_STANICA As String = "Stanica"
Public Const AMB_NALOG_KUPAC As String = "Kupac"
Public Const AMB_NALOG_VOZAC As String = "Vozac"
Public Const AMB_NALOG_FIRMA As String = "Firma"
Public Const AMB_NALOG_SPOLJNI As String = "SpoljniSvet"

' ============================================================
' VRSTA KRETANJA -- zatvoren enum, osam vrednosti
' ============================================================
'
' Odgovara na ZASTO je prenos nastao. Ko kome daje vec nose Od/Na, a koji je
' dokument povod nosi DokumentTIP -- zato ovde nema ni OTKUP_* ni STANICA_KOOPERANT.
'
' Vrsta NIJE izvediva iz para naloga, i to je razlog sto postoji kao podatak:
' Stanica -> Kooperant je fizicki isti potez i kad firma ZADUZUJE partnera
' (IZDATA_PRAZNA) i kad mu VRACA njegove gajbe (VRACANJE_TUDJE_AMBALAZE), a te
' dve stvari imaju suprotno dejstvo na dug.
Public Const AMB_VK_UZ_ROBU As String = "AMBALAZA_UZ_ROBU"
Public Const AMB_VK_IZDATA_PRAZNA As String = "IZDATA_PRAZNA"
Public Const AMB_VK_POVRAT_PRAZNE As String = "POVRAT_PRAZNE"
Public Const AMB_VK_PRENOS_INTERNO As String = "PRENOS_INTERNO"
Public Const AMB_VK_ULAZ_TUDJE As String = "ULAZ_TUDJE_AMBALAZE"
Public Const AMB_VK_VRACANJE_TUDJE As String = "VRACANJE_TUDJE_AMBALAZE"
Public Const AMB_VK_NABAVKA As String = "NABAVKA"
Public Const AMB_VK_OTPIS As String = "OTPIS"

' Vrsta AMBALAZNOG DOKUMENTA (AMB-10-DOK) -- zatvoren enum, tri vrednosti.
' Ne mesati sa VrstaKretanja iznad: ovo je vrsta DOKUMENTA, ono je vrsta
' KRETANJA. Vezu medju njima drzi AmbDokDozvoljavaKretanje.
Public Const AMB_DOK_REVERS As String = "REVERS"

' SMEROVI REVERSA -- do 10b-2 goli literali na 12 mesta.
'
' Konstante postoje da bi mapa ispod mogla da bude ZATVORENA: nepoznat smer
' nema red, pa pisac fail-closed pada. Zateceni Case Else je radio isto, ali je
' pravilo zivelo u pisecu; sada zivi u ugovoru, uz par naloga i vrstu.
Public Const REV_SMER_IZDAVANJE As String = "IZDAVANJE"
Public Const REV_SMER_PRIJEM As String = "PRIJEM"
Public Const REV_SMER_IZDATO_OM As String = "IZDATO_OM"
Public Const REV_SMER_PRIJEM_OD_OM As String = "PRIJEM_OD_OM"
' Dokument koji je izdao PARTNER, a ne mi. Danas: revers kupca -- dokaz da je
' vozac preuzeo prazne gajbe (odluka operatera 03.10.2026, AMB-10-ODL-10).
' Broj je NJEGOV, pa je vlasnik numerickog niza partner, ne sopstveni nalog.
Public Const AMB_DOK_REVERS_PARTNERA As String = "REVERS_PARTNERA"
Public Const AMB_DOK_NABAVKA As String = "NABAVKA"
Public Const AMB_DOK_OTPIS As String = "OTPIS"

' ============================================================
' KLASE NALOGA -- azbuka matrice ispod
' ============================================================
'
'   GRANICA    SpoljniSvet i nista drugo
'   SOPSTVENI  Stanica, Firma, Vozac -- unutar firme
'   PARTNER    Kooperant, Kupac -- druga strana posla
'   REALAN     bilo koji nosilac salda (sve osim granice)
Public Const AMB_KLASA_GRANICA As String = "GRANICA"
Public Const AMB_KLASA_SOPSTVENI As String = "SOPSTVENI"
Public Const AMB_KLASA_PARTNER As String = "PARTNER"
Public Const AMB_KLASA_REALAN As String = "REALAN"

' ============================================================
' ZATVORENE LISTE -- jedan izvor, da se citalac i test ne razidju
' ============================================================

' NAZIV VRSTE AMBALAZNOG DOKUMENTA (zaglavlje), za coveka.
'
' Zivi u UGOVORU, a ne u ekranu, jer je spisak vrsta domenski: enum je ovde
' (AMB_DOK_*), pa i njegov naziv. Do 10c-2 je stajao u modScrDokumenti i zvao ga
' je i modStornoDok; kad mu je trebao i IZVESTAJ, zavisnost bi isla report ->
' ekran, sto je obrnuto. Tri spiska naziva za isti enum bi se razisla prvom
' izmenom.
'
' Fail-open je namerni: nepoznata vrsta se vraca kao sopstveni tekst, da nov enum
' ne bi proizveo prazan natpis dok mu se ne doda poruka.
Public Function AmbVrstaDokNaziv(ByVal v As String) As String
    AmbVrstaDokNaziv = Trim$(v)
    Select Case Trim$(v)
        Case AMB_DOK_REVERS:          AmbVrstaDokNaziv = Poruka("OTKUI_AMBD_REVERS")
        Case AMB_DOK_REVERS_PARTNERA: AmbVrstaDokNaziv = Poruka("OTKUI_AMBD_REVERS_PART")
        Case AMB_DOK_NABAVKA:         AmbVrstaDokNaziv = Poruka("OTKUI_AMBD_NABAVKA")
        Case AMB_DOK_OTPIS:           AmbVrstaDokNaziv = Poruka("OTKUI_AMBD_OTPIS")
    End Select
End Function

Public Function AmbNaloziSvi() As Variant
    AmbNaloziSvi = Array(AMB_NALOG_KOOPERANT, AMB_NALOG_STANICA, AMB_NALOG_KUPAC, _
                         AMB_NALOG_VOZAC, AMB_NALOG_FIRMA, AMB_NALOG_SPOLJNI)
End Function

Public Function AmbVrsteSve() As Variant
    AmbVrsteSve = Array(AMB_VK_UZ_ROBU, AMB_VK_IZDATA_PRAZNA, AMB_VK_POVRAT_PRAZNE, _
                        AMB_VK_PRENOS_INTERNO, AMB_VK_ULAZ_TUDJE, AMB_VK_VRACANJE_TUDJE, _
                        AMB_VK_NABAVKA, AMB_VK_OTPIS)
End Function

Private Function UNizu(ByVal niz As Variant, ByVal vrednost As String) As Boolean
    Dim i As Long
    For i = LBound(niz) To UBound(niz)
        If StrComp(CStr(niz(i)), Trim$(vrednost), vbTextCompare) = 0 Then
            UNizu = True
            Exit Function
        End If
    Next i
End Function

Public Function AmbNalogTipPoznat(ByVal tip As String) As Boolean
    AmbNalogTipPoznat = UNizu(AmbNaloziSvi(), tip)
End Function

Public Function AmbVrstaPoznata(ByVal vrsta As String) As Boolean
    AmbVrstaPoznata = UNizu(AmbVrsteSve(), vrsta)
End Function

' KANONSKI ZAPIS -- knjiga cuva vrednost iz zatvorene liste, ne onu koju je
' pozivalac otkucao. Provere su vbTextCompare pa bi i "stanica" prosla, ali bi u
' koloni ostala druga pisana forma iste stvari -- a nad njom svaki GROUP BY,
' Dictionary kljuc i izvestaj racunaju dva naloga. Prazno znaci "nije u listi".
Private Function KanonIzNiza(ByVal niz As Variant, ByVal vrednost As String) As String
    Dim i As Long
    For i = LBound(niz) To UBound(niz)
        If StrComp(CStr(niz(i)), Trim$(vrednost), vbTextCompare) = 0 Then
            KanonIzNiza = CStr(niz(i))
            Exit Function
        End If
    Next i
End Function

Public Function AmbVrstaKanon(ByVal vrsta As String) As String
    AmbVrstaKanon = KanonIzNiza(AmbVrsteSve(), vrsta)
End Function

Public Function AmbNalogTipKanon(ByVal tip As String) As String
    AmbNalogTipKanon = KanonIzNiza(AmbNaloziSvi(), tip)
End Function

' SOPSTVENI nalozi -- unutar firme, ne partneri. Prazne gajbe sa stanice najcesce
' idu VOZACU pa tek onda drugoj stanici, pa je vozac ovde "nas" iako je transporter
' (AMB-10-ODL-7).
Public Function AmbNalogSopstveni(ByVal tip As String) As Boolean
    Select Case Trim$(tip)
        Case AMB_NALOG_STANICA, AMB_NALOG_FIRMA, AMB_NALOG_VOZAC
            AmbNalogSopstveni = True
    End Select
End Function

' PARTNER -- druga strana posla, i jedini nalog kome firma moze da DUGUJE.
' Obaveza(partner, tip) nema smisla nad stanicom ili vozacem.
Public Function AmbNalogPartner(ByVal tip As String) As Boolean
    Select Case Trim$(tip)
        Case AMB_NALOG_KOOPERANT, AMB_NALOG_KUPAC
            AmbNalogPartner = True
    End Select
End Function

' Sistemski nalog: jedan jedini, bez ID-a i bez master zapisa.
Public Function AmbNalogSistemski(ByVal tip As String) As Boolean
    AmbNalogSistemski = (StrComp(Trim$(tip), AMB_NALOG_FIRMA, vbTextCompare) = 0) Or _
                        (StrComp(Trim$(tip), AMB_NALOG_SPOLJNI, vbTextCompare) = 0)
End Function

' Master tabela u kojoj ID mora da postoji. "" za sistemske naloge.
Public Function AmbNalogTabela(ByVal tip As String) As String
    Select Case Trim$(tip)
        Case AMB_NALOG_KOOPERANT: AmbNalogTabela = TBL_KOOPERANTI
        Case AMB_NALOG_STANICA:   AmbNalogTabela = TBL_STANICE
        Case AMB_NALOG_KUPAC:     AmbNalogTabela = TBL_KUPCI
        Case AMB_NALOG_VOZAC:     AmbNalogTabela = TBL_VOZACI
    End Select
End Function

Private Function AmbNalogKljuc(ByVal tip As String) As String
    Select Case Trim$(tip)
        Case AMB_NALOG_KOOPERANT: AmbNalogKljuc = "KooperantID"
        Case AMB_NALOG_STANICA:   AmbNalogKljuc = "StanicaID"
        Case AMB_NALOG_KUPAC:     AmbNalogKljuc = "KupacID"
        Case AMB_NALOG_VOZAC:     AmbNalogKljuc = "VozacID"
    End Select
End Function

' ============================================================
' AMB-INV-03 -- razresavanje naloga, JEDINA kapija
' ============================================================
'
' Par Tip+ID je polimorfan, pa je "Tip=Vozac, ID=KUP-17" sintaksno ispravan a
' semanticki nemoguc. Zasebna tabela naloga bi to ukinula strukturno, ali bi
' uvela peti registar koji mora da prati cetiri master tabele -- klasa koja je
' ovaj repo vec ujela (MRTAV_UNOS u vba_hard_census). Zato JEDNA kapija, i
' nijedan pisac ne proverava tipove sam (AMB-10-ODL-1).
'
' Vraca "" kad je nalog valjan, inace IMENOVAN razlog.
' STRUKTURA NALOGA -- sve sto se o nalogu moze tvrditi BEZ citanja tabela.
'
' Postoji odvojeno zato sto je citaocu knjige potrebno bas ovo: red zapisan u
' tabeli mora da nosi ispravan OBLIK naloga, a postojanje maticnog reda je kapija
' UPISA. Da citalac proverava i postojanje, obrisan maticni red bi retroaktivno
' oborio svako citanje knjige, a svaki saldo bi postao kvadratan nad tabelom.
Public Function AmbNalogStrukturaProblem(ByVal tip As String, ByVal id As String) As String
    Dim t As String, k As String
    t = Trim$(tip)
    k = Trim$(id)

    If Len(t) = 0 Then
        AmbNalogStrukturaProblem = "Nalog nema tip."
        Exit Function
    End If

    If Not AmbNalogTipPoznat(t) Then
        AmbNalogStrukturaProblem = "Nepoznat tip naloga: '" & t & "'."
        Exit Function
    End If

    If AmbNalogSistemski(t) Then
        ' Jedan jedini nalog -- ID bi bio drugi identitet iste stvari.
        If Len(k) > 0 Then
            AmbNalogStrukturaProblem = "Nalog '" & t & "' je sistemski i nema ID (dobio: '" & k & "')."
        End If
        Exit Function
    End If

    If Len(k) = 0 Then
        AmbNalogStrukturaProblem = "Nalog '" & t & "' trazi ID."
    End If
End Function

Public Function AmbNalogProblem(ByVal tip As String, ByVal id As String) As String
    Dim t As String, k As String, tbl As String
    t = Trim$(tip)
    k = Trim$(id)

    AmbNalogProblem = AmbNalogStrukturaProblem(t, k)
    If Len(AmbNalogProblem) > 0 Then Exit Function

    ' Sistemski nalog nema maticni red koji bi se razresavao.
    If AmbNalogSistemski(t) Then Exit Function

    tbl = AmbNalogTabela(t)
    If Len(tbl) = 0 Then
        ' Tip je poznat a tabela nije mapirana -- greska u ovom modulu, ne u podatku.
        AmbNalogProblem = "Tip naloga '" & t & "' nema mapiranu tabelu."
        Exit Function
    End If

    ' AMB-INV-03 trazi JEDNOZNACNO razresenje, ne "postoji bar jedan".
    ' Dva master reda sa istim ID-em nisu "jos bolje" nego kvar: pisac ne zna
    ' kome pripisuje gajbe. Dvosmislenost je fail-closed, kao kod zbirne
    ' (ZbirnaIdentResolve / ZBR_RES_UNIQUE).
    Dim n As Long
    n = modDataAccess.FindRows(tbl, AmbNalogKljuc(t), k).count

    If n = 0 Then
        AmbNalogProblem = "Nalog '" & t & "' sa ID '" & k & "' ne postoji u " & tbl & "."
    ElseIf n > 1 Then
        AmbNalogProblem = "Nalog '" & t & "' sa ID '" & k & "' nije jednoznacan: " & _
                          CStr(n) & " reda u " & tbl & "."
    End If
End Function

' ============================================================
' AMBALAZNI DOKUMENT -- vrsta i zaglavlje (AMB-10-DOK)
' ============================================================
'
' tblAmbalazaDokument nosi dogadjaje koji nemaju svoj poslovni dokument. Njegov
' AmbDokID ide u tblAmbalaza.DokumentID, cime ReversID prestaje da bude drugi,
' paralelan identitet -- ne brise se nego POSTAJE ovo.
Public Function AmbDokVrsteSve() As Variant
    AmbDokVrsteSve = Array(AMB_DOK_REVERS, AMB_DOK_REVERS_PARTNERA, _
                           AMB_DOK_NABAVKA, AMB_DOK_OTPIS)
End Function

Public Function AmbDokVrstaPoznata(ByVal vrsta As String) As Boolean
    AmbDokVrstaPoznata = UNizu(AmbDokVrsteSve(), vrsta)
End Function

Public Function AmbDokVrstaKanon(ByVal vrsta As String) As String
    AmbDokVrstaKanon = KanonIzNiza(AmbDokVrsteSve(), vrsta)
End Function

' KOJE KRETANJE SME NA KOM DOKUMENTU.
'
' Bez ovoga bi 10b morao da pretpostavi, a pretpostavka bi prosla tiho: NABAVKA
' okacena na revers izgledala bi kao uredan zapis.
'
'   REVERS   -> IZDATA_PRAZNA, POVRAT_PRAZNE, PRENOS_INTERNO, VRACANJE_TUDJE
'   NABAVKA  -> NABAVKA
'   OTPIS    -> OTPIS
'
' AMBALAZA_UZ_ROBU nikad nije ovde -- ona putuje sa robom, pa joj je izvorni
' dokument otkup, otpremnica ili prijemnica.
'
' ULAZ_TUDJE_AMBALAZE sme uz SVAKI AMBALAZNI dokument: ono nije vrsta posla nego
' POKRICE DEFICITA (AMB-INV-07), pa nastaje svuda gde bi realan nalog pao ispod
' nule -- i na reversu, ne samo uz robu.
'
' OPSEG OVE FUNKCIJE: ona odgovara SAMO za tblAmbalazaDokument. Da li ULAZ_TUDJE
' sme uz otkup ili prijemnicu je drugo pitanje, jer robni tipovi dokumenata nisu
' ovde zatvoren skup -- pola AmbDok validator, pola genericki source-document
' validator bila bi funkcija koja ni jedno ne tvrdi do kraja.
Public Function AmbDokDozvoljavaKretanje(ByVal dokVrsta As String, _
                                         ByVal vrstaKretanja As String) As Boolean
    ' FAIL-CLOSED NA NEPOZNATU VRSTU (review #399, P2).
    '
    ' Dozvola za pokrice deficita je ranije stajala IZNAD ove provere, pa je
    ' ("NEPOSTOJECI_DOKUMENT", ULAZ_TUDJE) vracalo True i zaobilazilo zatvoren enum.
    ' Kapija koja odgovori pre nego sto proveri preduslov nije kapija.
    If Not AmbDokVrstaPoznata(dokVrsta) Then Exit Function

    If StrComp(Trim$(vrstaKretanja), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
        AmbDokDozvoljavaKretanje = True      ' pokrice deficita ide uz svaki AmbDok
        Exit Function
    End If

    Select Case Trim$(dokVrsta)
        Case AMB_DOK_REVERS
            Select Case Trim$(vrstaKretanja)
                Case AMB_VK_IZDATA_PRAZNA, AMB_VK_POVRAT_PRAZNE, _
                     AMB_VK_PRENOS_INTERNO, AMB_VK_VRACANJE_TUDJE
                    AmbDokDozvoljavaKretanje = True
            End Select
        Case AMB_DOK_REVERS_PARTNERA
            ' Partnerov dokument dokazuje da je ambalaza STIGLA OD NJEGA: kupac
            ' vraca prazne, vozac ih preuzima. IZDATA_PRAZNA ovde NE sme -- kad
            ' mi izdajemo partneru, dokument je NAS, pa je to AMB_DOK_REVERS.
            ' PRENOS_INTERNO takodje ne: interno kretanje ne moze imati partnerov
            ' papir kao povod.
            AmbDokDozvoljavaKretanje = (StrComp(Trim$(vrstaKretanja), AMB_VK_POVRAT_PRAZNE, vbTextCompare) = 0)
        Case AMB_DOK_NABAVKA
            AmbDokDozvoljavaKretanje = (StrComp(Trim$(vrstaKretanja), AMB_VK_NABAVKA, vbTextCompare) = 0)
        Case AMB_DOK_OTPIS
            AmbDokDozvoljavaKretanje = (StrComp(Trim$(vrstaKretanja), AMB_VK_OTPIS, vbTextCompare) = 0)
    End Select
End Function

' KO SME DA POSEDUJE NUMERICKI NIZ, PO VRSTI DOKUMENTA.
'
' Nije dovoljno da nalog postoji: AmbNalogProblem dokazuje postojanje, ne pravo na
' seriju brojeva. Bez ovoga prolazi "REVERS, vlasnik = Kooperant" -- partner koji
' poseduje NASU seriju -- pa cak i "vlasnik = SpoljniSvet", granica koja uopste
' nije drzalac.
'
' Pravilo je bilo jedno za sve vrste, u jednoj recenici:
'
'   BROJ JE NAS, PROTIVPARTNER JE NJIHOV.
'
' 03.10.2026 je dobilo opseg: vazi za dokument KOJI PISEMO MI -- revers
' kooperantu, nabavku, otpis. Revers KUPCA je njegov papir i nosi njegov broj
' (AMB-10-ODL-10), pa je dobio svoju vrstu REVERS_PARTNERA. Stari tekst je ovde
' izricito nabrajao "i revers kupca" kao nas -- to je bilo netacno.
' Vlasnik numerickog niza, po vrsti dokumenta.
'
' NAS dokument nosi NAS broj: partner ne izdaje nasu seriju, pa je vlasnik
' sopstveni nalog (Stanica, Firma, Vozac).
'
' PARTNEROV dokument nosi NJEGOV broj, i to je 03.10.2026 ispravljeno kao
' pravilo, ne izuzetak: dokument od kupca je dokaz da je vozac preuzeo ambalazu --
' kupcev papir, kupcev broj. Isto vazi za prijemnicu i bankovne izvode; tamo kod
' to vec radi (modOtkupUI.PredlogPrijemnice: "Ostali kupci nose svoj eksterni,
' nezavisni broj -- polje se tada NE dira"). Izmisljati uz njih i nas broj znaci
' dva broja za jedan papir, pa ni jedan nije onaj po kome operater trazi.
'
' Funkcija je po VRSTI, ne po smeru, i to je namerno: AmbDokMatricaNepotpuna
' obara svaku vrstu bez odgovora, pa nova putanja ne moze da se provuce tiho.
' Da je po smeru, zaglavlje bi se upisivalo pre nego sto se smer zna.
Public Function AmbDokBrojOwnerKlasa(ByVal vrsta As String) As String
    Select Case Trim$(vrsta)
        Case AMB_DOK_REVERS, AMB_DOK_NABAVKA, AMB_DOK_OTPIS
            AmbDokBrojOwnerKlasa = AMB_KLASA_SOPSTVENI
        Case AMB_DOK_REVERS_PARTNERA
            AmbDokBrojOwnerKlasa = AMB_KLASA_PARTNER
    End Select
End Function

' KAPIJA POTPUNOSTI nad vrstama dokumenta -- isti oblik kao AmbMatricaNepotpuna.
' Vrsta bez odgovora o vlasniku broja prosla bi kroz proveru zaglavlja neprimetno.
Public Function AmbDokMatricaNepotpuna() As String
    Dim sve As Variant, i As Long, fale As String
    sve = AmbDokVrsteSve()
    For i = LBound(sve) To UBound(sve)
        If Len(AmbDokBrojOwnerKlasa(CStr(sve(i)))) = 0 Then fale = fale & " " & CStr(sve(i))
    Next i
    AmbDokMatricaNepotpuna = Trim$(fale)
End Function

' ZAGLAVLJE DOKUMENTA -- provera pre upisa. Vraca "" ili imenovan razlog.
'
' VLASNIK NUMERICKOG NIZA JE OBAVEZAN. Broj bez opsega u kom je jedinstven nije
' identitet nego niz znakova: dva dokumenta mogu nositi isti broj a da nijedna
' provera ne primeti. Ranije je tu stajala samo StanicaID, i to opciona -- sto je
' revers kupca ostavljalo bez ikakvog opsega (review #399, P2).
'
' KOJI nalog je vlasnik po vrsti dokumenta odlucuje 10b, zajedno sa numeracijom.
' Ovde se tvrdi samo da vlasnik POSTOJI i da se razresava ISTOM kapijom kao svaki
' drugi nalog -- dakle "BrojOwnerTip=Vozac, BrojOwnerID=KUP-17" pada pre upisa.
Public Function AmbDokProblem(ByVal vrsta As String, ByVal broj As String, _
                              ByVal datum As Date, _
                              ByVal brojOwnerTip As String, _
                              ByVal brojOwnerID As String) As String
    If Not AmbDokVrstaPoznata(vrsta) Then
        AmbDokProblem = "Nepoznata vrsta ambalaznog dokumenta: '" & Trim$(vrsta) & "'."
        Exit Function
    End If

    If Len(Trim$(broj)) = 0 Then
        AmbDokProblem = "Ambalazni dokument nema broj."
        Exit Function
    End If

    If datum = 0 Then
        AmbDokProblem = "Ambalazni dokument nema datum."
        Exit Function
    End If

    If Len(Trim$(brojOwnerTip)) = 0 Then
        AmbDokProblem = "Ambalazni dokument nema vlasnika numerickog niza -- " & _
                        "broj bez opsega nije jedinstven."
        Exit Function
    End If

    Dim p As String
    p = AmbNalogProblem(brojOwnerTip, brojOwnerID)
    If Len(p) > 0 Then
        AmbDokProblem = "Vlasnik broja: " & p
        Exit Function
    End If

    ' Postojanje naloga NIJE pravo na seriju brojeva (review #399).
    Dim klasa As String
    klasa = AmbDokBrojOwnerKlasa(vrsta)
    If Len(klasa) = 0 Then
        AmbDokProblem = "Vrsta '" & Trim$(vrsta) & "' nema definisanog vlasnika broja."
        Exit Function
    End If

    ' Poruka je NEUTRALNA od 03.10.2026: ranije je zavrsavala sa "broj je nas,
    ' protivpartner je njihov", a to za REVERS_PARTNERA vise nije tacno -- pa je
    ' validacija padala ispravno a objasnjavala suprotno od pravila.
    If Not AmbNalogUKlasi(klasa, brojOwnerTip) Then
        AmbDokProblem = "'" & Trim$(vrsta) & "' trazi " & klasa & " kao vlasnika broja, " & _
                        "a dobio je " & Trim$(brojOwnerTip) & "."
    End If
End Function

Public Sub RequireAmbDok(ByVal vrsta As String, ByVal broj As String, _
                         ByVal datum As Date, ByVal brojOwnerTip As String, _
                         ByVal brojOwnerID As String, ByVal sourceName As String)
    Dim p As String
    p = AmbDokProblem(vrsta, broj, datum, brojOwnerTip, brojOwnerID)
    If Len(p) > 0 Then Err.Raise vbObjectError + 4451, sourceName, p
End Sub

' ============================================================
' MATRICA: koja klasa naloga sme na kojoj strani, po vrsti kretanja
' ============================================================
'
' CEO ugovor o stranama stoji OVDE, u jednoj tabeli. Ranije je bio razbacan po
' sest If blokova, i svaki je tvrdio SAMO jedan smer implikacije -- pa je sest
' puta zaredom nadjeno da komplement prolazi (granica u jednom smeru,
' PRENOS_INTERNO bez klasa, obaveza bez partnera, izdavanje i povrat bez ikakvih
' klasa). Matrica taj oblik greske cini nemogucim: strana koja nije navedena ne
' postoji, a vrsta bez reda pada na kapiji potpunosti (AmbMatricaNepotpuna).
'
'   vrsta                      | Od         | Na
'   ---------------------------|------------|------------
'   AMBALAZA_UZ_ROBU           | REALAN     | REALAN
'   IZDATA_PRAZNA              | SOPSTVENI  | PARTNER
'   POVRAT_PRAZNE              | PARTNER    | SOPSTVENI
'   PRENOS_INTERNO             | SOPSTVENI  | SOPSTVENI
'   ULAZ_TUDJE_AMBALAZE        | GRANICA    | PARTNER
'   VRACANJE_TUDJE_AMBALAZE    | SOPSTVENI  | PARTNER
'   NABAVKA                    | GRANICA    | SOPSTVENI
'   OTPIS                      | REALAN     | GRANICA
'
' NABAVKA ide na SOPSTVENI, ne na partnera: nove gajbe ulaze u firmu, pa se
' partneru IZDAJU (IZDATA_PRAZNA). Direktno `SpoljniSvet -> Kooperant NABAVKA`
' preskocilo bi cin izdavanja, a s njim i zaduzenje partnera.
'
' OTPIS sme sa BILO KOG realnog naloga: gajbica se moze polomiti i kod
' kooperanta, ne samo na stanici. Da li partner tada duguje naknadu je poslovno
' pitanje, ne knjigovodstveno.
Public Function AmbKlaseVrste(ByVal vrsta As String) As Variant
    Select Case Trim$(vrsta)
        Case AMB_VK_UZ_ROBU
            AmbKlaseVrste = Array(AMB_KLASA_REALAN, AMB_KLASA_REALAN)
        Case AMB_VK_IZDATA_PRAZNA
            AmbKlaseVrste = Array(AMB_KLASA_SOPSTVENI, AMB_KLASA_PARTNER)
        Case AMB_VK_POVRAT_PRAZNE
            AmbKlaseVrste = Array(AMB_KLASA_PARTNER, AMB_KLASA_SOPSTVENI)
        Case AMB_VK_PRENOS_INTERNO
            AmbKlaseVrste = Array(AMB_KLASA_SOPSTVENI, AMB_KLASA_SOPSTVENI)
        Case AMB_VK_ULAZ_TUDJE
            AmbKlaseVrste = Array(AMB_KLASA_GRANICA, AMB_KLASA_PARTNER)
        Case AMB_VK_VRACANJE_TUDJE
            AmbKlaseVrste = Array(AMB_KLASA_SOPSTVENI, AMB_KLASA_PARTNER)
        Case AMB_VK_NABAVKA
            AmbKlaseVrste = Array(AMB_KLASA_GRANICA, AMB_KLASA_SOPSTVENI)
        Case AMB_VK_OTPIS
            AmbKlaseVrste = Array(AMB_KLASA_REALAN, AMB_KLASA_GRANICA)
        Case Else
            AmbKlaseVrste = Array()
    End Select
End Function

' IZVORNA TABELA PO TIPU DOKUMENTA -- zatvorena mapa, fail-closed.
'
' AMB-INV-08 trazi da knjiga i izvorni dokument dele rollback. Vezivanje
' dokumenta za transakciju (BindSourceDocument) dokazuje IDENTITET transakcije,
' ali ne i da je tabela tog dokumenta u njenom snapshotu -- pa bi ovo prolazilo:
'
'   BindSourceDocument(Otkup, OTK-123)
'   snapshot tblAmbalaza    DA
'   snapshot tblOtkup       NE     -> rollback nije zajednicki
'
' Mapa stoji u domenu, ne u clsTransaction: transakcija je genericki primitiv i
' ne sme da zna tipove poslovnih dokumenata. Nepoznat tip vraca prazno, a pisac
' to tretira kao odbijanje -- nov tip dokumenta ne moze da se provuce.
' IZVORNI DOKUMENTI -- JEDAN POPIS, OBA SMERA IZ NJEGA.
'
' AMB-INV-08 trazi TABELU po tipu; undo garda (modStornoZurnal) trazi TIP po
' tabeli. Dve Select Case mape bi se razisle prvim sledecim presecenim
' dokumentom, pa oba citaoca citaju OVAJ popis. Fail-closed ostaje: sto nije u
' popisu, nema ni tabelu ni tip.
Public Function AmbIzvorniParovi() As Variant
    AmbIzvorniParovi = Array( _
        Array(DOK_TIP_AMBALAZA_DOKUMENT, TBL_AMBALAZA_DOKUMENT), _
        Array(DOK_TIP_OTKUP, TBL_OTKUP), _
        Array(DOK_TIP_OTPREMNICA, TBL_OTPREMNICA), _
        Array(DOK_TIP_PRIJEMNICA, TBL_PRIJEMNICA))
End Function

Public Function AmbIzvornaTabela(ByVal dokTip As String) As String
    Dim p As Variant, i As Long
    p = AmbIzvorniParovi()
    For i = LBound(p) To UBound(p)
        If StrComp(Trim$(dokTip), CStr(p(i)(0)), vbTextCompare) = 0 Then
            AmbIzvornaTabela = CStr(p(i)(1))
            Exit Function
        End If
    Next i
End Function

' Obrnut smer: tip izvornog dokumenta za tabelu u kojoj zaglavlje zivi. Vraca
' "" za svaku tabelu koja nije izvorna -- pozivalac tada nema sta da pita.
Public Function AmbDokTipZaIzvornuTabelu(ByVal tabela As String) As String
    Dim p As Variant, i As Long
    p = AmbIzvorniParovi()
    For i = LBound(p) To UBound(p)
        If StrComp(Trim$(tabela), CStr(p(i)(1)), vbTextCompare) = 0 Then
            AmbDokTipZaIzvornuTabelu = CStr(p(i)(0))
            Exit Function
        End If
    Next i
End Function

' ZAGLAVLJE I KRETANJE SE PROVERAVAJU ZAJEDNO (AMB-10-ODL-9, -10).
'
' Klasa vlasnika broja i dozvoljeno kretanje su bile dve nezavisne provere, pa
' nijedna nije videla drugu. Tri zaobilaznice koje su time prolazile:
'
'   Kupac -> FIRMA sa REVERS_PARTNERA   Firma je SOPSTVENI, pa klase prolaze --
'                                       a ODL-9 kaze da firma NE ulazi u lanac
'   BrojOwner K1, kretanje K2 -> Vozac  broj jednog partnera, dogadjaj drugog
'   obican REVERS: Kupac -> Vozac       partnerov broj potpuno zaobidjen
'
' OBRNUTA KAPIJA JE DEO PRAVILA, ne dodatak: bez nje se isto kretanje moze
' knjiziti na nasu vrstu i dobiti nas broj, pa ODL-10 ne vazi ni za jedan
' dokument -- samo za one koje pozivalac izvoli da nazove REVERS_PARTNERA.
'
' Generisani redovi NE prolaze ovde: pisac ih zove iz vec provernog zahteva, a
' pokrice deficita ima par GRANICA -> PARTNER koji ni jedno pravilo ne pogadja.
Public Function AmbDokKretanjeProblem(ByVal vrsta As String, _
                                      ByVal brojOwnerTip As String, _
                                      ByVal brojOwnerID As String, _
                                      ByVal odTip As String, ByVal odID As String, _
                                      ByVal naTip As String, _
                                      ByVal vrstaKretanja As String) As String
    Dim jePartnerov As Boolean, povratOdKupca As Boolean
    Dim ko As String
    jePartnerov = (StrComp(Trim$(vrsta), AMB_DOK_REVERS_PARTNERA, vbTextCompare) = 0)
    ' OBRNUTA KAPIJA GLEDA SVAKI POVRAT OD KUPCA, ne samo kupac -> vozac.
    ' Prva verzija je trazila ceo par, pa su obican REVERS nad kupac -> FIRMA i
    ' kupac -> STANICA prolazili: par nije bio kupac -> vozac, pa se kapija nije
    ' ni palila. ODL-9 kaze da lanac ide kupac -> vozac -> stanica i da firma u
    ' njega NE ulazi, pa je uslov sam POVRAT OD KUPCA.
    povratOdKupca = (StrComp(Trim$(odTip), AMB_NALOG_KUPAC, vbTextCompare) = 0) And _
                    (StrComp(Trim$(vrstaKretanja), AMB_VK_POVRAT_PRAZNE, vbTextCompare) = 0)

    If jePartnerov Then
        If StrComp(Trim$(odTip), AMB_NALOG_KUPAC, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-9: partnerov revers polazi od " & _
                "kupca, a dobio je " & Trim$(odTip) & "."
            Exit Function
        End If
        If StrComp(Trim$(naTip), AMB_NALOG_VOZAC, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-9: lanac je kupac -> vozac -> " & _
                "stanica; partnerov revers ne ide na " & Trim$(naTip) & "."
            Exit Function
        End If
        If StrComp(Trim$(vrstaKretanja), AMB_VK_POVRAT_PRAZNE, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-10: partnerov revers nosi " & _
                "povrat praznih, a dobio je " & Trim$(vrstaKretanja) & "."
            Exit Function
        End If
        If StrComp(Trim$(brojOwnerTip), AMB_NALOG_KUPAC, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-10: broj partnerovog reversa " & _
                "pripada kupcu, a vlasnik niza je " & Trim$(brojOwnerTip) & "."
            Exit Function
        End If
        If StrComp(Trim$(brojOwnerID), Trim$(odID), vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-10: broj nosi kupac '" & _
                Trim$(brojOwnerID) & "' a ambalazu predaje '" & Trim$(odID) & _
                "' -- dokument bi imao tudj broj."
            Exit Function
        End If
        Exit Function
    End If

    If povratOdKupca Then
        If StrComp(Trim$(naTip), AMB_NALOG_VOZAC, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-9: povrat praznih od kupca ide " & _
                "na vozaca (lanac kupac -> vozac -> stanica), a ide na " & _
                Trim$(naTip) & "."
            Exit Function
        End If
        ' AMB-10-ODL-22: pravilo je VLASNIK BROJA, ne vrsta dokumenta.
        '
        ' Do 05.10.2026 je ovde stajalo "vrsta mora biti REVERS_PARTNERA".
        ' Operater je izmerio premisu: PRIJEMNICA je i sama partnerov dokument
        ' (eksterna je, osim kad nasa hladnjaca izdaje), a zamena pune ambalaze
        ' praznom se knjizi POD BROJEM PRIJEMNICE -- bez dodatnog broja. ODL-10
        ' tu nije zaobidjen nego ISPUNJEN: povrat nosi kupcev broj.
        '
        ' Zato se meri ono sto ODL-10 i kaze -- "nosi njegov broj" -- a ne vrsta
        ' dokumenta kao njena zamena. Obican REVERS nad kupac -> vozac i dalje
        ' pada, jer je njegov broj NAS; robni dokument koji vlasnika broja ne
        ' objavi isto pada, jer prazan vlasnik nije kupac (fail-closed).
        ko = Trim$(brojOwnerTip)
        If Len(ko) = 0 Then ko = "(nije zadat)"
        If StrComp(Trim$(brojOwnerTip), AMB_NALOG_KUPAC, vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-10: povrat praznih od kupca " & _
                "nosi kupcev broj, a vlasnik broja je " & ko & " (dokument '" & _
                Trim$(vrsta) & "')."
            Exit Function
        End If
        If StrComp(Trim$(brojOwnerID), Trim$(odID), vbTextCompare) <> 0 Then
            AmbDokKretanjeProblem = "AMB-10-ODL-10: broj nosi kupac '" & _
                Trim$(brojOwnerID) & "' a prazne vraca '" & Trim$(odID) & _
                "' -- dokument bi imao tudj broj."
        End If
    End If
End Function
' AMB-10-ODL-22: ROBNI dokument UME da bude partnerov.
'
' Prijemnica je eksterni dokument -- izdaje ju hladnjaca, a mi je primamo --
' pa je njen broj KUPCEV broj.
'
' KAD NASA HLADNJACA IZDAJE PRIJEMNICU, VLASNIK BROJA JE NAMERNO ISTI IZRAZ
' (odluka operatera 05.10.2026, posle P2 iz review-a). Review je tacno rekao da
' KupacID odgovara na "ko je kupac u poslu" a BrojOwner na "cijem nizu pripada
' broj" -- i da ih AMB-10 svuda drugde razdvaja. Ovde se NE razdvajaju, i to je
' IZMERENO a ne pretpostavljeno: modBrojevi.GenerateBrojPrijemnice scope-uje niz
' bas po (KupacID, dan) --
'
'   MaxSeqFromTable(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_DATUM,
'                   COL_PRJ_KUPAC, kupacID, datum)
'
' -- uz svoj komentar da auto-numeracija vazi SAMO za hladnjaca-kupca, a ostali
' kupci nose svoj eksterni broj. Dakle vlasnik broja JE kupac, u oba rezima, i
' poklapa se sa AMB-10-ODL-20 (BrojOwnerTip, BrojOwnerID, dan). Jedno pravilo,
' bez grananja po izdavaocu.
'
' Ovo bi se MORALO ponovo izmeriti ako broj prijemnice ikada pocne da se
' generise iz NASEG niza nezavisnog od kupca -- tada role i vlasnistvo prestaju
' da se poklapaju i mapa treba granu.
'
' Mapa je ZATVORENA i za ostale tipove vraca prazno. To nije rupa nego
' fail-closed: obrnuta kapija ODL-10 odbija povrat od kupca bez vlasnika
' broja, pa dokument koji ga ne objavi ne moze da knjizi taj povrat.
'
' Oblik reda: (dokTip, tabela, kolona ID-a, nalogTip vlasnika, kolona vlasnika)
Public Function AmbRobniVlasniciBroja() As Variant
    AmbRobniVlasniciBroja = Array( _
        Array(DOK_TIP_PRIJEMNICA, TBL_PRIJEMNICA, COL_PRJ_ID, _
              AMB_NALOG_KUPAC, COL_PRJ_KUPAC))
End Function

' SMER REVERSA -> PAR NALOGA I VRSTA KRETANJA (AMB-10-ODL-7, -ODL-8).
'
' Stari pisac je isti posao radio SA SEST NOGU u cetiri smera, i vozaca nosio
' kao ZIG (kolona VozacID) -- pa se njegov saldo dobijao inverzijom smera, sto
' je fail-open. Nov red imenuje obe strane, pa je po smeru dovoljan JEDAN red.
'
' Vrste nisu izvedene iz para nego PROCITANE iz 6.7:
'   stanica zaduzuje kooperanta praznim   -> IZDATA_PRAZNA
'   kooperant ih vraca                    -> POVRAT_PRAZNE
'   vozac <-> stanica, oba SOPSTVENA      -> PRENOS_INTERNO (ODL-7)
'
' Mapa je ZATVORENA: nepoznat smer nema red i pisac pada fail-closed. Oblik:
' (smer, odTip, naTip, vrstaKretanja).
Public Function AmbReversSmerovi() As Variant
    AmbReversSmerovi = Array( _
        Array(REV_SMER_IZDAVANJE, AMB_NALOG_STANICA, AMB_NALOG_KOOPERANT, _
              AMB_VK_IZDATA_PRAZNA), _
        Array(REV_SMER_PRIJEM, AMB_NALOG_KOOPERANT, AMB_NALOG_STANICA, _
              AMB_VK_POVRAT_PRAZNE), _
        Array(REV_SMER_IZDATO_OM, AMB_NALOG_VOZAC, AMB_NALOG_STANICA, _
              AMB_VK_PRENOS_INTERNO), _
        Array(REV_SMER_PRIJEM_OD_OM, AMB_NALOG_STANICA, AMB_NALOG_VOZAC, _
              AMB_VK_PRENOS_INTERNO))
End Function

' Pripada li nalog trazenoj klasi.
Public Function AmbNalogUKlasi(ByVal klasa As String, ByVal tip As String) As Boolean
    Select Case Trim$(klasa)
        Case AMB_KLASA_GRANICA
            AmbNalogUKlasi = (StrComp(Trim$(tip), AMB_NALOG_SPOLJNI, vbTextCompare) = 0)
        Case AMB_KLASA_SOPSTVENI
            AmbNalogUKlasi = AmbNalogSopstveni(tip)
        Case AMB_KLASA_PARTNER
            AmbNalogUKlasi = AmbNalogPartner(tip)
        Case AMB_KLASA_REALAN
            AmbNalogUKlasi = AmbNalogTipPoznat(tip) And _
                             (StrComp(Trim$(tip), AMB_NALOG_SPOLJNI, vbTextCompare) <> 0)
    End Select
End Function

' KAPIJA POTPUNOSTI: nijedna vrsta ne sme da ostane bez reda u matrici.
'
' Bez ovoga bi deveta vrsta dodata u enum tiho prolazila kroz sve provere strana.
' Vraca "" kad je matrica potpuna, inace imena vrsta koje fale.
Public Function AmbMatricaNepotpuna() As String
    Dim sve As Variant, i As Long, klase As Variant, fale As String
    sve = AmbVrsteSve()
    For i = LBound(sve) To UBound(sve)
        klase = AmbKlaseVrste(CStr(sve(i)))
        If UBound(klase) < 1 Then fale = fale & " " & CStr(sve(i))
    Next i
    AmbMatricaNepotpuna = Trim$(fale)
End Function

' ============================================================
' JEDAN ZAHTEV, VISE REDOVA -- sta pisac sme da proizvede
' ============================================================
'
' Pisac ne upisuje uvek tacno ono sto je trazeno, i to su dve odluke modela:
'
'   deficit (AMB-INV-07)  -> uz trazeni red ide POKRICE, na DRUGOM paru naloga
'   obaveza (AMB-INV-09)  -> trazeni red se DELI na dva, na ISTOM paru naloga
'
' Druga lomi naivnu idempotenciju. Ponovljen isti zahtev (vracanje 20 uz obavezu
' 12) nalazi red od 12, pa bi poredjenje "trazena kolicina == kolicina reda"
' prijavilo HARD CONFLICT nad potpuno ispravnim ponavljanjem -- a ponovno
' racunanje podele nije izlaz, jer je obaveza posle prvog upisa DRUGA. Zato se
' ponavljanje meri ZBIROM preko para naloga, a ovde stoji koje vrste jedan zahtev
' uopste sme da proizvede na tom paru.
'
' ULAZ_TUDJE_AMBALAZE NIJE ZAHTEV. Ona je posledica -- pokrice deficita -- i
' generise je pisac. Da je i zahtev, pokrice jednog zahteva i eksplicitan zahtev
' nad istim dokumentom delili bi i par i vrstu, pa bi jedan tiho progutao drugi
' kao "idempotentno ponavljanje".
Public Function AmbVrstaJeZahtev(ByVal vrsta As String) As Boolean
    If Not AmbVrstaPoznata(vrsta) Then Exit Function
    AmbVrstaJeZahtev = (StrComp(Trim$(vrsta), AMB_VK_ULAZ_TUDJE, vbTextCompare) <> 0)
End Function

' Vrste koje jedan zahtev sme da proizvede NA ISTOM PARU naloga. Prazno za vrstu
' koja nije zahtev.
Public Function AmbVrsteZahteva(ByVal vrsta As String) As Variant
    AmbVrsteZahteva = Array()
    If Not AmbVrstaJeZahtev(vrsta) Then Exit Function

    Select Case AmbVrstaKanon(vrsta)
        Case AMB_VK_VRACANJE_TUDJE
            ' AMB-INV-09: vracanje preko obaveze se cepa, a ostatak je NOVO
            ' zaduzenje partnera -- ne vracanje. Dva dogadjaja, dve vrste.
            AmbVrsteZahteva = Array(AMB_VK_VRACANJE_TUDJE, AMB_VK_IZDATA_PRAZNA)
        Case Else
            AmbVrsteZahteva = Array(AmbVrstaKanon(vrsta))
    End Select
End Function

' KAPIJA POTPUNOSTI nad proizvodnjom -- isti oblik kao AmbMatricaNepotpuna.
'
' Tri uslova, jer su tri nacina da ovo tiho pukne:
'   1. vrsta koja JE zahtev a ne proizvodi nista -- pisac bi upisao red koji
'      nijedno ponavljanje ne bi prepoznalo;
'   2. zahtev koji ne proizvodi SAMOG SEBE -- podela bez trazenog reda;
'   3. proizvedena vrsta sa DRUGIM parom klasa -- podela ne sme da promeni ko sme
'      da stoji na kojoj strani, inace zaobilazi matricu kroz sopstveni ostatak.
Public Function AmbProizvodnjaNepotpuna() As String
    Dim sve As Variant, prod As Variant, klaseZ As Variant, klaseP As Variant
    Dim i As Long, j As Long, v As String, fale As String

    sve = AmbVrsteSve()
    For i = LBound(sve) To UBound(sve)
        v = CStr(sve(i))
        prod = AmbVrsteZahteva(v)

        If Not AmbVrstaJeZahtev(v) Then
            If UBound(prod) >= LBound(prod) Then fale = fale & " " & v & ":posledica-proizvodi"
        ElseIf UBound(prod) < LBound(prod) Then
            fale = fale & " " & v & ":bez-proizvodnje"
        Else
            If Not UNizu(prod, v) Then fale = fale & " " & v & ":bez-sebe"

            klaseZ = AmbKlaseVrste(v)
            For j = LBound(prod) To UBound(prod)
                klaseP = AmbKlaseVrste(CStr(prod(j)))
                If UBound(klaseZ) < 1 Or UBound(klaseP) < 1 Then
                    fale = fale & " " & v & ">" & CStr(prod(j)) & ":bez-klasa"
                ElseIf StrComp(CStr(klaseZ(0)), CStr(klaseP(0)), vbTextCompare) <> 0 Or _
                       StrComp(CStr(klaseZ(1)), CStr(klaseP(1)), vbTextCompare) <> 0 Then
                    fale = fale & " " & v & ">" & CStr(prod(j)) & ":drugi-par-klasa"
                End If
            Next j
        End If
    Next i

    AmbProizvodnjaNepotpuna = Trim$(fale)
End Function

' ============================================================
' CIJI DEFICIT SE SME POKRITI -- AMB-10-ODL-8
' ============================================================
'
' Pokrice deficita je ULAZ_TUDJE_AMBALAZE, a matrica za nju kaze GRANICA ->
' PARTNER. Iz toga SLEDI, bez nove odluke, da se deficit SOPSTVENOG naloga ne
' pokriva: "nase gajbe iz vazduha" nije dogadjaj nego skriven manjak. Stanica
' koja nema gajbe ne sme da ih izda; put je NABAVKA, sa svojim dokumentom i svojom
' cenom.
'
' Odgovor se CITA iz matrice, a ne pise kao "PARTNER": kad bi se klasa odredista
' ULAZ_TUDJE ikad promenila, ovo ide za njom samo. Isti razlog zbog kog matrica
' postoji.
Public Function AmbPokriceKlasa() As String
    Dim klase As Variant
    klase = AmbKlaseVrste(AMB_VK_ULAZ_TUDJE)
    If UBound(klase) < 1 Then Exit Function
    AmbPokriceKlasa = CStr(klase(1))
End Function

Public Function AmbPokriceProblem(ByVal tip As String, ByVal id As String) As String
    Dim p As String, klasa As String

    p = AmbNalogProblem(tip, id)
    If Len(p) > 0 Then
        AmbPokriceProblem = p
        Exit Function
    End If

    klasa = AmbPokriceKlasa()
    If Len(klasa) = 0 Then
        AmbPokriceProblem = "Pokrice deficita nema definisanu klasu odredista."
        Exit Function
    End If

    If Not AmbNalogUKlasi(klasa, tip) Then
        AmbPokriceProblem = "Deficit naloga " & Trim$(tip) & " se NE pokriva ulazom tudje " & _
                            "ambalaze (pokrice trazi " & klasa & "): nase gajbe ne nastaju " & _
                            "iz vazduha, za njih ide NABAVKA."
    End If
End Function

' ============================================================
' PROVERA PRENOSA -- AMB-INV-01, -02, -03 + granica
' ============================================================
'
' Vraca "" kad je prenos valjan, inace IMENOVAN razlog. Ne pise nista i ne zna
' za transakciju: 10b je zove pre upisa, ekran je zove za poruku uz polje.
' STRUKTURA PRENOSA -- ceo ugovor osim postojanja naloga.
'
' Citalac knjige meri bas ovo nad ZAPISANIM redom (modAmbalaza.KnjigaRedProblem),
' a pisac isto plus postojanje. Jedna implementacija, dva pozivaoca -- druga kopija
' matrice bi se razisla prvom izmenom.
Public Function AmbPrenosStrukturaProblem(ByVal odTip As String, ByVal odID As String, _
                                          ByVal naTip As String, ByVal naID As String, _
                                          ByVal kolicina As Double, ByVal tipAmb As String, _
                                          ByVal vrsta As String) As String
    Dim p As String

    ' AMB-INV-01: kolicina je uvek pozitivna -- smer nosi par naloga, ne znak.
    If kolicina <= 0 Then
        AmbPrenosStrukturaProblem = "Kolicina mora biti veca od nule (dobio: " & CStr(kolicina) & ")."
        Exit Function
    End If

    If Len(Trim$(tipAmb)) = 0 Then
        AmbPrenosStrukturaProblem = "Tip ambalaze je obavezan -- stanje se vodi PO TIPU."
        Exit Function
    End If

    If Not AmbVrstaPoznata(vrsta) Then
        AmbPrenosStrukturaProblem = "Nepoznata vrsta kretanja: '" & Trim$(vrsta) & "'."
        Exit Function
    End If

    p = AmbNalogStrukturaProblem(odTip, odID)
    If Len(p) > 0 Then
        AmbPrenosStrukturaProblem = "Nalog OD: " & p
        Exit Function
    End If

    p = AmbNalogStrukturaProblem(naTip, naID)
    If Len(p) > 0 Then
        AmbPrenosStrukturaProblem = "Nalog NA: " & p
        Exit Function
    End If

    ' AMB-INV-02: prenos na samog sebe nije dogadjaj nego greska unosa.
    If StrComp(Trim$(odTip), Trim$(naTip), vbTextCompare) = 0 And _
       StrComp(Trim$(odID), Trim$(naID), vbTextCompare) = 0 Then
        AmbPrenosStrukturaProblem = "Prenos na isti nalog: " & Trim$(odTip) & " '" & Trim$(odID) & "'."
        Exit Function
    End If

    ' KLASE STRANA -- iz matrice, oba smera odjednom.
    '
    ' Ovde je ranije stajalo sest If blokova, svaki sa svojim smerom implikacije.
    ' Matrica ih zamenjuje: sta nije navedeno, ne prolazi.
    Dim klase As Variant
    klase = AmbKlaseVrste(vrsta)
    If UBound(klase) < 1 Then
        ' Vrsta je u enumu a nema red u matrici -- kvar ugovora, ne podatka.
        AmbPrenosStrukturaProblem = "Vrsta '" & Trim$(vrsta) & "' nema definisane klase strana."
        Exit Function
    End If

    If Not AmbNalogUKlasi(CStr(klase(0)), odTip) Then
        AmbPrenosStrukturaProblem = "'" & Trim$(vrsta) & "' trazi " & CStr(klase(0)) & _
                           " kao IZVOR, a dobio je " & Trim$(odTip) & "."
        Exit Function
    End If

    If Not AmbNalogUKlasi(CStr(klase(1)), naTip) Then
        AmbPrenosStrukturaProblem = "'" & Trim$(vrsta) & "' trazi " & CStr(klase(1)) & _
                           " kao ODREDISTE, a dobio je " & Trim$(naTip) & "."
    End If
End Function

' PUN ugovor prenosa: struktura + postojanje oba naloga. Ovo zove PISAC.
Public Function AmbPrenosProblem(ByVal odTip As String, ByVal odID As String, _
                                 ByVal naTip As String, ByVal naID As String, _
                                 ByVal kolicina As Double, ByVal tipAmb As String, _
                                 ByVal vrsta As String) As String
    AmbPrenosProblem = AmbPrenosStrukturaProblem(odTip, odID, naTip, naID, _
                                                 kolicina, tipAmb, vrsta)
    If Len(AmbPrenosProblem) > 0 Then Exit Function

    Dim p As String
    p = AmbNalogProblem(odTip, odID)
    If Len(p) > 0 Then
        AmbPrenosProblem = "Nalog OD: " & p
        Exit Function
    End If

    p = AmbNalogProblem(naTip, naID)
    If Len(p) > 0 Then AmbPrenosProblem = "Nalog NA: " & p
End Function

' Ista provera, za pisca: greska umesto poruke.
Public Sub RequireAmbPrenos(ByVal odTip As String, ByVal odID As String, _
                            ByVal naTip As String, ByVal naID As String, _
                            ByVal kolicina As Double, ByVal tipAmb As String, _
                            ByVal vrsta As String, ByVal sourceName As String)
    Dim p As String
    p = AmbPrenosProblem(odTip, odID, naTip, naID, kolicina, tipAmb, vrsta)
    If Len(p) > 0 Then
        Err.Raise vbObjectError + 4450, sourceName, p
    End If
End Sub

' ============================================================
' DOPRINOS OBAVEZI -- osnova AMB-INV-09
' ============================================================
'
' Obaveza firme prema partneru se IZVODI iz knjige, bez ijedne mutabilne kolone:
'
'   Obaveza(partner, tip) = SUM AmbDoprinosObavezi(svi dogadjaji partnera i tipa)
'
' Prosta razlika dve sume NIJE dovoljna nad append-only knjigom: storno ulaza
' tudje ambalaze upisuje kontra-stav, fizicko stanje se anulira, a obaveza bi
' ostala da visi. Zato storno daje MINUS doprinos originala -- jedno pravilo gasi
' i stanje i obavezu.
'
' vrstaOriginala je prazna za obican dogadjaj, a za kontra-stav nosi vrstu reda
' koji se ponistava.
Public Function AmbDoprinosObavezi(ByVal vrsta As String, ByVal kolicina As Double, _
                                   ByVal vrstaOriginala As String) As Double
    If Len(Trim$(vrstaOriginala)) > 0 Then
        AmbDoprinosObavezi = -AmbDoprinosObavezi(vrstaOriginala, kolicina, "")
        Exit Function
    End If

    Select Case Trim$(vrsta)
        Case AMB_VK_ULAZ_TUDJE:      AmbDoprinosObavezi = kolicina
        Case AMB_VK_VRACANJE_TUDJE:  AmbDoprinosObavezi = -kolicina
        Case Else:                   AmbDoprinosObavezi = 0
    End Select
End Function
