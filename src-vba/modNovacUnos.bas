Attribute VB_Name = "modNovacUnos"
'=====================================================================
' modNovacUnos - UNOS NOVCA I AMBALAZE (isplate F5, uplate kupaca F6,
' reversi F7), bez ijedne kontrole.
'
' Treci modul istog oblika: modOtkupUnos (F1), modDokUnos (F2-F4),
' modNovacUnos (F5-F7). Razlog je isti - poslovni posao ne sme da zivi
' u formi, jer ga onda drugi ekran ne moze pozvati bez prepisivanja.
'
'   IsplataValidiraj / IsplataUpisi   F5, SaveOMUlaz_TX (samo novac)
'   UplataValidiraj  / UplataUpisi    F6, SaveKupciIzlaz_TX (samo novac)
'   ReversValidiraj  / ReversUpisi    F7, SaveOMUlaz_TX (samo ambalaza)
'
' Sve tri Validiraj rutine vracaju "" kad je proslo, inace poruku za
' operatera, i pune LOGICKO ime polja na koje treba vratiti fokus.
' Sve tri Upisi rutine vracaju broj dokumenta ("" = nije upisano).
'
' ODAKLE DOLAZE PRAVILA: frmDokumenta.btnUnosOMUlaz_Click (F5 i F7 su
' dve polovine tog jednog handlera) i frmDokumenta.btnUnosIzlaz_Click
' (F6). Redosled provera je prepisan, jer je operater navikao koje ga
' polje prvo zaustavi.
'
' TRI RAZLIKE U ODNOSU NA LEGACY - namerne, ne propust:
'
'   1) VOZAC SE NE TRAZI ZA CIST NOVAC (F5, F6). Legacy ga trazi uz
'      VALIDACIJA_UNOSA, ali u tom handleru isti dokument sme da nosi i
'      ambalazu, a vozac postoji zbog nje. U novom UI-ju gotovinski
'      rezimi ambalazu NEMAJU (modOtkupUI.ApplyFormFields), pa se
'      vozac ne upisuje nigde: ni SaveNovac ni tblNovac nemaju to polje.
'      Provera bi zaustavljala operatera na podatku koji se odbacuje.
'      U F7, gde ambalaza postoji, vozac se trazi - i to strozije nego
'      uz VALIDACIJA_UNOSA (vidi ReversValidiraj).
'
'   2) F5, PARTNER JE OTKUPNO MESTO -> TO OTKUPNO MESTO JE ENTITET
'      NOVCA. Polje se u tom rezimu zove "Primalac", pa izabrano
'      otkupno mesto JESTE onaj ko prima novac. Legacy tu mogucnost
'      nije imao: primalac je bio samo kooperant, a otkupno mesto se
'      podrazumevalo iz konteksta forme. Kad je partner kooperant,
'      entitet ostaje otkupno mesto iz konteksta - tacno kao legacy.
'
'   3) F7 NE PRIMA KUPCA KAO PARTNERA. Cetiri smera reversa
'      (SaveOMUlaz_TX) idu iskljucivo kooperant <-> OM <-> firma;
'      ambalaza kupca u legacy ide kroz prijemnicu (povrat) i kroz
'      kupci-izlaz, ne kroz revers. Izbor kupca za smer "izdato/prijem
'      kooperantu" se zato odbija porukom, umesto da se tiho proknjizi
'      na pogresan entitet.
'
' VAZNO: legacy frmDokumenta OSTAJE netaknut i potpuno operativan. Ovaj
' modul je drugi pozivalac istih poslovnih rutina, ne zamena za formu.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const NOVUNOS_BUILD As String = "v6-ui-118"

' Smer reversa: redni broj segmenta u polju SMER REVERSA (modOtkupUI,
' isti raspored kao frmDokumenta.SetupOMIzdavanjeToggle) -> vrednost
' koju SaveOMUlaz_TX ocekuje. Nula znaci "operater nije izabrao".
Public Const SMER_REV_IZD_KOOP As Long = 1
Public Const SMER_REV_PRI_KOOP As Long = 2
Public Const SMER_REV_IZD_OM As Long = 3
Public Const SMER_REV_PRI_OM As Long = 4

'--------------------------------------------------------------- ULAZ
' Prazan recnik sa svim kljucevima - da pozivalac ne mora da pamti spisak.
Public Function NoviIsplataUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("stanicaID") = ""
    p("stanicaTekst") = ""
    p("partnerID") = ""
    p("partnerTip") = ""
    p("partnerTekst") = ""
    p("vrsta") = ""
    p("brDok") = ""
    p("novac") = 0#
    p("otkupID") = ""
    ' Vidljivi tekst polja "otvoreni blok" ide uz ID bas zbog NerazresenIzbor:
    ' bez njega se ukucan a nerazresen blok ne razlikuje od praznog polja.
    p("blokTekst") = ""
    p("otkupOstatak") = 0#
    p("izAvansa") = False
    p("tipNovca") = ""
    Set NoviIsplataUnos = p
End Function

Public Function NoviUplataUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("partnerID") = ""
    p("partnerTekst") = ""
    p("vrsta") = ""
    p("brDok") = ""
    p("novac") = 0#
    p("fakturaID") = ""
    p("fakturaTekst") = ""
    p("fakturaOstatak") = 0#
    p("tipNovca") = ""
    p("napomena") = ""
    Set NoviUplataUnos = p
End Function

Public Function NoviReversUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("stanicaID") = ""
    p("stanicaTekst") = ""
    p("partnerID") = ""
    p("partnerTip") = ""
    p("partnerTekst") = ""
    p("vozacID") = ""
    p("vrsta") = ""
    p("brDok") = ""
    p("tipAmb") = ""
    p("kolAmb") = 0&
    p("smerRev") = 0&
    Set NoviReversUnos = p
End Function

Private Function S(ByVal p As Object, ByVal k As String) As String
    On Error Resume Next
    If p.Exists(k) Then S = Trim$(CStr(p(k)))
End Function

Private Function D(ByVal p As Object, ByVal k As String) As Double
    On Error Resume Next
    If p.Exists(k) Then
        If IsNumeric(p(k)) Then D = CDbl(p(k))
    End If
End Function

Private Function L(ByVal p As Object, ByVal k As String) As Long
    On Error Resume Next
    If p.Exists(k) Then
        If IsNumeric(p(k)) Then L = CLng(p(k))
    End If
End Function

Private Function B(ByVal p As Object, ByVal k As String) As Boolean
    On Error Resume Next
    If p.Exists(k) Then B = CBool(p(k))
End Function

' Broj dokumenta se u gotovinskim rezimima unosi slobodno (Z3a), pa sme
' da ostane prazan kad je VALIDACIJA_UNOSA iskljucena. Prazan povratak
' iz *Upisi je rezervisan za NEUSPEH, pa upis bez broja vraca ovu
' oznaku - inace bi uspeo upis izgledao kao pad.
Private Function BrojIliOznaka(ByVal brDok As String) As String
    If Len(Trim$(brDok)) > 0 Then BrojIliOznaka = Trim$(brDok) Else BrojIliOznaka = Poruka("NOVUNOS_BEZ_BROJA")
End Function

' UKUCAN A NERAZRESEN IZBOR = GRESKA, NE "NIJE IZABRANO".
'
' Combo-i ljuske dopustaju kucanje, a ID i tip stizu iz SKRIVENIH kolona koje
' postoje samo kad je stavka stvarno izabrana (ListIndex >= 0). Operater zato
' moze da vidi ime u polju, a da ljuska posalje prazan ID - i tada bi svaka od
' ovih grana tiho promenila znacenje dokumenta:
'
'   partner  -> isplata prestaje da bude isplata kooperantu i postaje
'               kes firma->otkupac, vezan za otkupno mesto umesto za coveka
'   blok     -> razduzenje otkupnog bloka postaje avans kooperantu
'   faktura  -> uplata po fakturi postaje avans kupca, faktura ostaje otvorena
'
' Sve tri se knjize kao ISPRAVAN dokument, samo pogresan. Zato prazan ID uz
' NEPRAZAN tekst nije stanje nego greska: dokument staje dok operater ne izabere
' stavku iz liste. Prazno polje i dalje znaci "nije izabrano" i prolazi.
Private Function NerazresenIzbor(ByVal tekst As String, ByVal iD As String) As Boolean
    NerazresenIzbor = (Len(Trim$(tekst)) > 0 And Len(Trim$(iD)) = 0)
End Function

' Isti par provera kao legacy: broj dokumenta je zajednicki namespace za
' ambalazu i za novac, pa se duplikat trazi u OBE tabele.
Private Function DuplBroj(ByVal brDok As String) As String
    If Len(brDok) = 0 Then Exit Function
    DuplBroj = CheckDuplicate(TBL_AMBALAZA, COL_AMB_DOK_ID, brDok, COL_AMB_DATUM)
    If Len(DuplBroj) > 0 Then Exit Function
    DuplBroj = CheckDuplicate(TBL_NOVAC, COL_NOV_BROJ_DOK, brDok, COL_NOV_DATUM)
End Function

'=====================================================================
' F5 ISPLATE - novac izlazi
'
' Kome se isplacuje odredjuje TIP NOVCA, a tip novca je jedini razlog
' zbog kog ovaj rezim uopste ima polja "otvoreni blok" i "isplata iz":
' isti iznos knjizen pod pogresnim tipom ne vidi se u saldu otkupnog
' mesta niti razduzuje otkupni blok.
'=====================================================================
Public Function IsplataValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim strogo As Boolean, novac As Double, dup As String
    Dim partTip As String, partID As String
    Dim omSaldo As Double, blokErr As String
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    If Len(S(p, "stanicaID")) = 0 Then
        fokus = "stanicaID": IsplataValidiraj = Poruka("OTKUNOS_ERR_OM"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": IsplataValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If

    ' Ukucano ime primaoca koje nije izabrano iz liste ne sme da prodje kao
    ' "nema primaoca" - vidi NerazresenIzbor.
    If NerazresenIzbor(S(p, "partnerTekst"), S(p, "partnerID")) Then
        fokus = "partnerID": IsplataValidiraj = Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"): Exit Function
    End If
    If NerazresenIzbor(S(p, "blokTekst"), S(p, "otkupID")) Then
        fokus = "otkupID": IsplataValidiraj = Poruka("NOVUNOS_ERR_BLOK_NEIZABRAN"): Exit Function
    End If

    ' Duplikat broja se proverava PRE iznosa - isti redosled kao legacy.
    dup = DuplBroj(S(p, "brDok"))
    If Len(dup) > 0 Then
        fokus = "brDok": IsplataValidiraj = dup: Exit Function
    End If

    novac = D(p, "novac")
    If novac <= 0 Then
        fokus = "novac": IsplataValidiraj = Poruka("NOVUNOS_ERR_NOVAC"): Exit Function
    End If

    partTip = UCase$(S(p, "partnerTip"))
    partID = S(p, "partnerID")

    If partTip = "KOOP" And Len(partID) > 0 Then
        ' --- isplata kooperantu ---
        If Len(S(p, "otkupID")) > 0 Then
            ' VLASNISTVO I TRENUTNI OSTATAK BLOKA. Provera se NE radi nad
            ' vrednoscu koju je ekran poslao: taj broj je snimak iz trenutka
            ' kad je lista punjena, a izmedju punjenja i potvrde stanje se moze
            ' promeniti (drugi unos, uvoz izvoda, drugi pozivalac writer-a).
            ' modNovac.IsplataBlokProblem cita stanje SADA, i istu kapiju
            ' nezavisno podize i SaveOMUlaz_TX - ovde je samo da operater dobije
            ' poruku uz polje umesto izuzetka.
            blokErr = IsplataBlokProblem(S(p, "otkupID"), partID, S(p, "stanicaID"), novac)
            If Len(blokErr) > 0 Then
                fokus = "novac": IsplataValidiraj = blokErr: Exit Function
            End If
            ' "Isplata iz OM avansa" ima smisla samo dok su kes isplate
            ' ukljucene - legacy taj prekidac tada i ne pusta da se pritisne.
            If B(p, "izAvansa") And IsKesIsplate() Then
                omSaldo = GetOMAvansSaldo(S(p, "stanicaID"))
                If novac > omSaldo Then
                    fokus = "novac"
                    IsplataValidiraj = Poruka("DOK_MSG_NEDOVOLJNO_AVANSA_RASPOLOZIVO") & " " & _
                                       Format$(omSaldo, "#,##0.00") & " RSD"
                    Exit Function
                End If
                p("tipNovca") = NOV_KES_OTKUPAC_KOOP
            Else
                p("tipNovca") = NOV_VIRMAN_FIRMA_KOOP
            End If
        Else
            ' Bez bloka isplata nije razduzenje nego avans kooperantu.
            p("tipNovca") = NOV_VIRMAN_AVANS_KOOP
        End If
    Else
        ' --- isplata otkupnom mestu (ili bez izabranog primaoca) ---
        ' Kooperant otpada, pa otpada i veza sa otkupnim blokom: red bez
        ' kooperanta sa OtkupID-em bio bi razduzenje nicijeg bloka.
        p("partnerID") = ""
        p("otkupID") = ""
        p("tipNovca") = NOV_KES_FIRMA_OTKUPAC
        ' Izabrano otkupno mesto JESTE primalac, pa je ono i entitet novca
        ' (razlika 2 iz zaglavlja). Bez partnera ostaje kontekst.
        If partTip = "OM" And Len(partID) > 0 Then
            p("stanicaID") = partID
            If Len(S(p, "partnerTekst")) > 0 Then p("stanicaTekst") = S(p, "partnerTekst")
        End If
    End If

    If strogo And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": IsplataValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If
    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error
    ' Resume Next, koji cisti Err.
    errDesc = Err.description
    LogErr "modNovacUnos.IsplataValidiraj"
    IsplataValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

Public Function IsplataUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    If Not SaveOMUlaz_TX( _
        datum:=CDate(p("datum")), _
        brojDok:=S(p, "brDok"), _
        stanicaNaziv:=S(p, "stanicaTekst"), _
        stanicaID:=S(p, "stanicaID"), _
        vozacID:="", _
        tipAmb:="", _
        kolAmb:=0, _
        vrstaVoca:=S(p, "vrsta"), _
        novac:=D(p, "novac"), _
        kooperantID:=S(p, "partnerID"), _
        primalacDisplay:=S(p, "partnerTekst"), _
        otkupID:=S(p, "otkupID"), _
        tipNovca:=S(p, "tipNovca"), _
        koopSmer:="") Then Exit Function

    IsplataUpisi = BrojIliOznaka(S(p, "brDok"))
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modNovacUnos.IsplataUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'=====================================================================
' F6 UPLATE KUPACA - novac ulazi
'
' Izabrana faktura je jedina razlika izmedju uplate i avansa: bez nje
' novac nema sta da zatvori, pa red ostaje avans kupca.
'=====================================================================
Public Function UplataValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim strogo As Boolean, novac As Double, dup As String
    Dim fakErr As String
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    ' Kupac je obavezan UVEK (SaveKupciIzlaz_TX ga i sam trazi), za razliku
    ' od vrste koja pada pod VALIDACIJA_UNOSA.
    ' Ukucano ime kupca ide PRE opste provere "kupac je obavezan": operater vidi
    ' ime u polju, pa mu poruka "izaberi kupca" ne bi rekla sta nije u redu.
    If NerazresenIzbor(S(p, "partnerTekst"), S(p, "partnerID")) Then
        fokus = "partnerID": UplataValidiraj = Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"): Exit Function
    End If
    If Len(S(p, "partnerID")) = 0 Then
        fokus = "partnerID": UplataValidiraj = Poruka("DOKUNOS_ERR_KUPAC"): Exit Function
    End If
    If NerazresenIzbor(S(p, "fakturaTekst"), S(p, "fakturaID")) Then
        fokus = "fakturaID": UplataValidiraj = Poruka("NOVUNOS_ERR_FAKTURA_NEIZABRANA"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": UplataValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If

    dup = DuplBroj(S(p, "brDok"))
    If Len(dup) > 0 Then
        fokus = "brDok": UplataValidiraj = dup: Exit Function
    End If

    novac = D(p, "novac")
    If novac <= 0 Then
        fokus = "novac": UplataValidiraj = Poruka("NOVUNOS_ERR_NOVAC"): Exit Function
    End If

    If Len(S(p, "fakturaID")) > 0 Then
        ' VLASNISTVO I TRENUTNO PREOSTALO. Isti razlog kao kod bloka: vrednost
        ' koju je ekran poslao je snimak iz trenutka punjenja liste. Kapija se
        ' cita SADA i istu nezavisno podize SaveKupciIzlaz_TX.
        fakErr = UplataFakturaProblem(S(p, "fakturaID"), S(p, "partnerID"), novac)
        If Len(fakErr) > 0 Then
            fokus = "novac": UplataValidiraj = fakErr: Exit Function
        End If
        p("tipNovca") = NOV_KUPCI_UPLATA
        p("napomena") = Poruka("NOVUNOS_NAP_FAKTURA") & " " & S(p, "fakturaTekst")
    Else
        p("tipNovca") = NOV_KUPCI_AVANS
        p("napomena") = Poruka("NOVUNOS_NAP_AVANS_KUP")
    End If

    If strogo And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": UplataValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modNovacUnos.UplataValidiraj"
    UplataValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

Public Function UplataUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    If Not SaveKupciIzlaz_TX( _
        datum:=CDate(p("datum")), _
        brojDok:=S(p, "brDok"), _
        kupacNaziv:=S(p, "partnerTekst"), _
        kupacID:=S(p, "partnerID"), _
        vozacID:="", _
        tipAmb:="", _
        kolAmb:=0, _
        vrstaVoca:=S(p, "vrsta"), _
        novac:=D(p, "novac"), _
        fakturaID:=S(p, "fakturaID"), _
        napomena:=S(p, "napomena"), _
        tipNovca:=S(p, "tipNovca")) Then Exit Function

    UplataUpisi = BrojIliOznaka(S(p, "brDok"))
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modNovacUnos.UplataUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'=====================================================================
' F7 REVERSI - kretanje prazne ambalaze
'
' Smer je OBAVEZAN i nije stvar prikaza: bez njega je SaveOMUlaz_TX
' ranije tiho knjizio legacy "OM prima od vozaca", pa je prazan smer
' davao pogresan red bez ijedne poruke. Danas core podize gresku, a
' ovde se zaustavlja pre nje, da operater dobije polje a ne izuzetak.
'=====================================================================
Public Function ReversValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim strogo As Boolean, kolAmb As Long, smer As Long, dup As String
    Dim partTip As String
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    If Len(S(p, "stanicaID")) = 0 Then
        fokus = "stanicaID": ReversValidiraj = Poruka("OTKUNOS_ERR_OM"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": ReversValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If

    dup = DuplBroj(S(p, "brDok"))
    If Len(dup) > 0 Then
        fokus = "brDok": ReversValidiraj = dup: Exit Function
    End If

    kolAmb = L(p, "kolAmb")
    If kolAmb <= 0 Then
        fokus = "kolAmb": ReversValidiraj = Poruka("NOVUNOS_ERR_KOL_AMB"): Exit Function
    End If
    If Len(S(p, "tipAmb")) = 0 Then
        fokus = "tipAmb": ReversValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
    End If

    smer = L(p, "smerRev")
    If smer < SMER_REV_IZD_KOOP Or smer > SMER_REV_PRI_OM Then
        fokus = "smerRev": ReversValidiraj = Poruka("NOVUNOS_ERR_SMER"): Exit Function
    End If

    partTip = UCase$(S(p, "partnerTip"))
    If smer = SMER_REV_IZD_KOOP Or smer = SMER_REV_PRI_KOOP Then
        ' Ukucano ime bez izbora iz liste dobija svoju poruku: opsta ("izaberi
        ' kooperanta") ne bi objasnila zasto polje sa vidljivim imenom ne valja.
        If NerazresenIzbor(S(p, "partnerTekst"), S(p, "partnerID")) Then
            fokus = "partnerID": ReversValidiraj = Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"): Exit Function
        End If
        ' Dvojni upis (kooperant + razduzenje/zaduzenje OM) trazi kooperanta.
        ' Kupac ovde ne postoji ni u legacy (razlika 3 iz zaglavlja).
        If Len(S(p, "partnerID")) = 0 Or partTip <> "KOOP" Then
            fokus = "partnerID": ReversValidiraj = Poruka("NOVUNOS_ERR_SMER_KOOP"): Exit Function
        End If
    Else
        ' Firma <-> OM uvek ide preko vozaca, pa je vozac obavezan i kad je
        ' VALIDACIJA_UNOSA iskljucena - SaveOMUlaz_TX bi inace podigao gresku.
        If Len(S(p, "vozacID")) = 0 Then
            fokus = "vozacID": ReversValidiraj = Poruka("NOVUNOS_ERR_VOZAC_OM"): Exit Function
        End If
    End If

    ' Auto-broj reversa (postuje AUTO_BROJ_DOKUMENTA kroz SuggestNextBroj):
    ' ako broj nije unet ni predlozen, generise se sada, po modelu
    ' x/ddmmyy[-N] iz revers tokova nad tblAmbalaza. Isto radi legacy, i to
    ' tek posle izbora smera - zato je ovde, a ne pre njega.
    If Len(S(p, "brDok")) = 0 Then
        p("brDok") = SuggestNextBroj(KIND_REV, S(p, "stanicaID"), CDate(p("datum")))
        dup = DuplBroj(S(p, "brDok"))
        If Len(dup) > 0 Then
            fokus = "brDok": ReversValidiraj = dup: Exit Function
        End If
    End If

    If strogo And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": ReversValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modNovacUnos.ReversValidiraj"
    ReversValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Redni broj segmenta -> vrednost koju SaveOMUlaz_TX poznaje. Nepoznat
' smer vraca prazno, pa core guard puca umesto da knjizi nasumice.
Public Function SmerRevKljuc(ByVal smer As Long) As String
    Select Case smer
        Case SMER_REV_IZD_KOOP: SmerRevKljuc = "IZDAVANJE"
        Case SMER_REV_PRI_KOOP: SmerRevKljuc = "PRIJEM"
        Case SMER_REV_IZD_OM:   SmerRevKljuc = "IZDATO_OM"
        Case SMER_REV_PRI_OM:   SmerRevKljuc = "PRIJEM_OD_OM"
    End Select
End Function

Public Function ReversUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim smer As Long, brDok As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""
    smer = L(p, "smerRev")
    brDok = S(p, "brDok")

    If Not SaveOMUlaz_TX( _
        datum:=CDate(p("datum")), _
        brojDok:=brDok, _
        stanicaNaziv:=S(p, "stanicaTekst"), _
        stanicaID:=S(p, "stanicaID"), _
        vozacID:=S(p, "vozacID"), _
        tipAmb:=S(p, "tipAmb"), _
        kolAmb:=L(p, "kolAmb"), _
        vrstaVoca:=S(p, "vrsta"), _
        novac:=0, _
        kooperantID:=S(p, "partnerID"), _
        primalacDisplay:=S(p, "partnerTekst"), _
        otkupID:="", _
        tipNovca:="", _
        koopSmer:=SmerRevKljuc(smer)) Then Exit Function

    ' ISPRAVKA reversa (druga faza): ako je revers-ispravka na cekanju,
    ' upravo snimljeni revers je njena zamena. No-op inace.
    modDokUnos.ZavrsiIspravkuAko FLOW_DOC_REVERS, brDok, poruke

    ' PDF revers - best-effort, isto kao u legacy: pad stampe ne sme da
    ' obori potvrdu upisa (rutina loguje sama).
    StampajRevers p, smer

    ReversUpisi = BrojIliOznaka(brDok)
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modNovacUnos.ReversUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Dva oblika reversa, kao u legacy: kooperantski nosi ime kooperanta i
' ID, firma<->OM nosi ime vozaca (bez ID-ja) i drugu protivstranu.
Private Sub StampajRevers(ByVal p As Object, ByVal smer As Long)
    Dim vozacNaziv As String
    On Error GoTo EH
    If smer = SMER_REV_IZD_KOOP Or smer = SMER_REV_PRI_KOOP Then
        OutputIzdavanjeAmbalaze CDate(p("datum")), S(p, "brDok"), _
                                S(p, "stanicaTekst"), S(p, "stanicaID"), _
                                S(p, "partnerTekst"), S(p, "partnerID"), _
                                S(p, "tipAmb"), L(p, "kolAmb"), S(p, "vrsta"), _
                                (smer = SMER_REV_PRI_KOOP)
    Else
        vozacNaziv = Trim$(NzToText(LookupValue(TBL_VOZACI, "VozacID", S(p, "vozacID"), "Ime")) & " " & _
                           NzToText(LookupValue(TBL_VOZACI, "VozacID", S(p, "vozacID"), "Prezime")))
        OutputIzdavanjeAmbalaze CDate(p("datum")), S(p, "brDok"), _
                                S(p, "stanicaTekst"), S(p, "stanicaID"), _
                                vozacNaziv, "", _
                                S(p, "tipAmb"), L(p, "kolAmb"), S(p, "vrsta"), _
                                (smer = SMER_REV_IZD_OM), "FIRMA"
    End If
    Exit Sub
EH:
    LogErr "modNovacUnos.StampajRevers"
End Sub
