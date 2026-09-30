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
Public Function AmbNalogProblem(ByVal tip As String, ByVal id As String) As String
    Dim t As String, k As String, tbl As String
    t = Trim$(tip)
    k = Trim$(id)

    If Len(t) = 0 Then
        AmbNalogProblem = "Nalog nema tip."
        Exit Function
    End If

    If Not AmbNalogTipPoznat(t) Then
        AmbNalogProblem = "Nepoznat tip naloga: '" & t & "'."
        Exit Function
    End If

    If AmbNalogSistemski(t) Then
        ' Jedan jedini nalog -- ID bi bio drugi identitet iste stvari.
        If Len(k) > 0 Then
            AmbNalogProblem = "Nalog '" & t & "' je sistemski i nema ID (dobio: '" & k & "')."
        End If
        Exit Function
    End If

    If Len(k) = 0 Then
        AmbNalogProblem = "Nalog '" & t & "' trazi ID."
        Exit Function
    End If

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
    AmbDokVrsteSve = Array(AMB_DOK_REVERS, AMB_DOK_NABAVKA, AMB_DOK_OTPIS)
End Function

Public Function AmbDokVrstaPoznata(ByVal vrsta As String) As Boolean
    AmbDokVrstaPoznata = UNizu(AmbDokVrsteSve(), vrsta)
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
' ULAZ_TUDJE_AMBALAZE je izuzetak i sme UZ SVAKI dokument: ono nije vrsta posla
' nego POKRICE DEFICITA (AMB-INV-07), pa nastaje svuda gde bi realan nalog pao
' ispod nule -- i uz otkup i uz revers.
Public Function AmbDokDozvoljavaKretanje(ByVal dokVrsta As String, _
                                         ByVal vrstaKretanja As String) As Boolean
    If StrComp(Trim$(vrstaKretanja), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
        AmbDokDozvoljavaKretanje = True      ' pokrice deficita ide svuda
        Exit Function
    End If

    Select Case Trim$(dokVrsta)
        Case AMB_DOK_REVERS
            Select Case Trim$(vrstaKretanja)
                Case AMB_VK_IZDATA_PRAZNA, AMB_VK_POVRAT_PRAZNE, _
                     AMB_VK_PRENOS_INTERNO, AMB_VK_VRACANJE_TUDJE
                    AmbDokDozvoljavaKretanje = True
            End Select
        Case AMB_DOK_NABAVKA
            AmbDokDozvoljavaKretanje = (StrComp(Trim$(vrstaKretanja), AMB_VK_NABAVKA, vbTextCompare) = 0)
        Case AMB_DOK_OTPIS
            AmbDokDozvoljavaKretanje = (StrComp(Trim$(vrstaKretanja), AMB_VK_OTPIS, vbTextCompare) = 0)
    End Select
End Function

' ZAGLAVLJE DOKUMENTA -- provera pre upisa. Vraca "" ili imenovan razlog.
'
' Stanica nije strana dogadjaja nego OPSEG JEDINSTVENOSTI BROJA (modBrojevi trazi
' slobodan broj po stanici i danu), pa se proverava kad je data. Da li je obavezna
' po vrsti dokumenta odlucuje 10b, zajedno sa numeracijom.
Public Function AmbDokProblem(ByVal vrsta As String, ByVal broj As String, _
                              ByVal datum As Date, ByVal stanicaID As String) As String
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

    If Len(Trim$(stanicaID)) > 0 Then
        Dim p As String
        p = AmbNalogProblem(AMB_NALOG_STANICA, stanicaID)
        If Len(p) > 0 Then AmbDokProblem = "Stanica dokumenta: " & p
    End If
End Function

Public Sub RequireAmbDok(ByVal vrsta As String, ByVal broj As String, _
                         ByVal datum As Date, ByVal stanicaID As String, _
                         ByVal sourceName As String)
    Dim p As String
    p = AmbDokProblem(vrsta, broj, datum, stanicaID)
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
' PROVERA PRENOSA -- AMB-INV-01, -02, -03 + granica
' ============================================================
'
' Vraca "" kad je prenos valjan, inace IMENOVAN razlog. Ne pise nista i ne zna
' za transakciju: 10b je zove pre upisa, ekran je zove za poruku uz polje.
Public Function AmbPrenosProblem(ByVal odTip As String, ByVal odID As String, _
                                 ByVal naTip As String, ByVal naID As String, _
                                 ByVal kolicina As Double, ByVal tipAmb As String, _
                                 ByVal vrsta As String) As String
    Dim p As String

    ' AMB-INV-01: kolicina je uvek pozitivna -- smer nosi par naloga, ne znak.
    If kolicina <= 0 Then
        AmbPrenosProblem = "Kolicina mora biti veca od nule (dobio: " & CStr(kolicina) & ")."
        Exit Function
    End If

    If Len(Trim$(tipAmb)) = 0 Then
        AmbPrenosProblem = "Tip ambalaze je obavezan -- stanje se vodi PO TIPU."
        Exit Function
    End If

    If Not AmbVrstaPoznata(vrsta) Then
        AmbPrenosProblem = "Nepoznata vrsta kretanja: '" & Trim$(vrsta) & "'."
        Exit Function
    End If

    p = AmbNalogProblem(odTip, odID)
    If Len(p) > 0 Then
        AmbPrenosProblem = "Nalog OD: " & p
        Exit Function
    End If

    p = AmbNalogProblem(naTip, naID)
    If Len(p) > 0 Then
        AmbPrenosProblem = "Nalog NA: " & p
        Exit Function
    End If

    ' AMB-INV-02: prenos na samog sebe nije dogadjaj nego greska unosa.
    If StrComp(Trim$(odTip), Trim$(naTip), vbTextCompare) = 0 And _
       StrComp(Trim$(odID), Trim$(naID), vbTextCompare) = 0 Then
        AmbPrenosProblem = "Prenos na isti nalog: " & Trim$(odTip) & " '" & Trim$(odID) & "'."
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
        AmbPrenosProblem = "Vrsta '" & Trim$(vrsta) & "' nema definisane klase strana."
        Exit Function
    End If

    If Not AmbNalogUKlasi(CStr(klase(0)), odTip) Then
        AmbPrenosProblem = "'" & Trim$(vrsta) & "' trazi " & CStr(klase(0)) & _
                           " kao IZVOR, a dobio je " & Trim$(odTip) & "."
        Exit Function
    End If

    If Not AmbNalogUKlasi(CStr(klase(1)), naTip) Then
        AmbPrenosProblem = "'" & Trim$(vrsta) & "' trazi " & CStr(klase(1)) & _
                           " kao ODREDISTE, a dobio je " & Trim$(naTip) & "."
    End If
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
