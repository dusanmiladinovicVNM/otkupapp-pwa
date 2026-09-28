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
' GRANICA OPTICAJA -- gde SpoljniSvet sme da stoji
' ============================================================
'
' Ambalaza nastaje i nestaje iz opticaja SAMO kroz tri vrste kretanja, i uvek na
' tacno odredjenoj strani. Bez ovoga bi granica mogla da se pojavi bilo gde i
' "gubitak" bi izgledao kao obican prenos.
Public Function AmbGranicaSmeKaoOd(ByVal vrsta As String) As Boolean
    Select Case Trim$(vrsta)
        Case AMB_VK_ULAZ_TUDJE, AMB_VK_NABAVKA: AmbGranicaSmeKaoOd = True
    End Select
End Function

Public Function AmbGranicaSmeKaoNa(ByVal vrsta As String) As Boolean
    AmbGranicaSmeKaoNa = (StrComp(Trim$(vrsta), AMB_VK_OTPIS, vbTextCompare) = 0)
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

    ' GRANICA OPTICAJA -- ekvivalencija, ne jednosmerna implikacija.
    '
    ' Prva verzija je proveravala samo "ako je SpoljniSvet tu, vrsta mora biti X".
    ' Komplement je prolazio: `Stanica -> Kooperant, NABAVKA` je bio validan, iako
    ' ugovor kaze da tim dogadjajem ambalaza ULAZI u opticaj. Zato se svaki uslov
    ' izgovara kao "vazi tacno tada i nikad inace".
    Dim odJeGranica As Boolean, naJeGranica As Boolean
    odJeGranica = (StrComp(Trim$(odTip), AMB_NALOG_SPOLJNI, vbTextCompare) = 0)
    naJeGranica = (StrComp(Trim$(naTip), AMB_NALOG_SPOLJNI, vbTextCompare) = 0)

    If AmbGranicaSmeKaoOd(vrsta) And Not odJeGranica Then
        AmbPrenosProblem = "'" & Trim$(vrsta) & "' mora imati " & AMB_NALOG_SPOLJNI & _
                           " kao izvor -- tom vrstom ambalaza ULAZI u opticaj."
        Exit Function
    End If

    If odJeGranica And Not AmbGranicaSmeKaoOd(vrsta) Then
        AmbPrenosProblem = AMB_NALOG_SPOLJNI & " kao izvor nije dozvoljen za '" & _
                           Trim$(vrsta) & "' -- ambalaza ulazi u opticaj samo kao " & _
                           AMB_VK_ULAZ_TUDJE & " ili " & AMB_VK_NABAVKA & "."
        Exit Function
    End If

    If AmbGranicaSmeKaoNa(vrsta) And Not naJeGranica Then
        AmbPrenosProblem = "'" & Trim$(vrsta) & "' mora imati " & AMB_NALOG_SPOLJNI & _
                           " kao odrediste -- tom vrstom ambalaza IZLAZI iz opticaja."
        Exit Function
    End If

    If naJeGranica And Not AmbGranicaSmeKaoNa(vrsta) Then
        AmbPrenosProblem = AMB_NALOG_SPOLJNI & " kao odrediste nije dozvoljen za '" & _
                           Trim$(vrsta) & "' -- iz opticaja se izlazi samo kao " & _
                           AMB_VK_OTPIS & "."
        Exit Function
    End If

    ' PRENOS_INTERNO je po definiciji izmedju SOPSTVENIH naloga (AMB-10-ODL-7).
    ' Bez ovoga bi `Kooperant -> Kupac, PRENOS_INTERNO` prosao kroz centralnu
    ' kapiju sa potpuno pogresnim poslovnim znacenjem.
    If StrComp(Trim$(vrsta), AMB_VK_PRENOS_INTERNO, vbTextCompare) = 0 Then
        If Not (AmbNalogSopstveni(odTip) And AmbNalogSopstveni(naTip)) Then
            AmbPrenosProblem = AMB_VK_PRENOS_INTERNO & " ide samo izmedju sopstvenih " & _
                               "naloga (Stanica, Firma, Vozac) -- dobio: " & _
                               Trim$(odTip) & " -> " & Trim$(naTip) & "."
            Exit Function
        End If
    End If

    ' VRSTE KOJE DIRAJU OBAVEZU MORAJU IMATI PARTNERA NA PARTNERSKOJ STRANI.
    '
    ' Nije trazeno u review-u nego je ista klasa: Obaveza(partner, tip) se racuna
    ' PO PARTNERU, pa `SpoljniSvet -> Stanica, ULAZ_TUDJE_AMBALAZE` nema kome da
    ' pripise dug. Stanica i vozac nisu partneri -- firma sebi ne duguje.
    If StrComp(Trim$(vrsta), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
        If Not AmbNalogPartner(naTip) Then
            AmbPrenosProblem = AMB_VK_ULAZ_TUDJE & " trazi PARTNERA kao odrediste " & _
                               "(Kooperant ili Kupac) -- inace obaveza nema kome da se " & _
                               "pripise. Dobio: " & Trim$(naTip) & "."
            Exit Function
        End If
    End If

    If StrComp(Trim$(vrsta), AMB_VK_VRACANJE_TUDJE, vbTextCompare) = 0 Then
        If Not AmbNalogPartner(naTip) Then
            AmbPrenosProblem = AMB_VK_VRACANJE_TUDJE & " trazi PARTNERA kao odrediste " & _
                               "-- firma vraca gajbe onome od koga ih je uzela. Dobio: " & _
                               Trim$(naTip) & "."
            Exit Function
        End If
        If Not AmbNalogSopstveni(odTip) Then
            AmbPrenosProblem = AMB_VK_VRACANJE_TUDJE & " ide sa SOPSTVENOG naloga -- " & _
                               "dobio: " & Trim$(odTip) & "."
            Exit Function
        End If
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
