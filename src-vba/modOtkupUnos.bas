Attribute VB_Name = "modOtkupUnos"
'=====================================================================
' modOtkupUnos - UNOS OTKUPNOG LISTA, bez ijedne kontrole.
'
' Zasto postoji: ceo posao unosa otkupa (provere, bruto->neto, upis, stampa,
' auto-lanac hladnjace, prevezivanje paleta pri ispravci) do sada je ziveo u
' frmOtkup.btnUnos_Click - tri stotine linija poslovne logike zakljucane u
' jednoj formi. Novi UI ne moze da je pozove, a prepisivanje bi napravilo dve
' kopije koje se razilaze.
'
' Ovde je taj posao izdvojen tako da ga zovu OBE forme:
'
'   OtkupValidiraj(p, fokus)   provere + bruto->neto; vraca poruku o gresci
'                              ("" = proslo) i LOGICKO ime polja na koje treba
'                              vratiti fokus
'   OtkupUpisi(p, poruke)      SaveOtkupMulti_TX + stampa + auto-lanac
'                              hladnjace; vraca OtkupID (prazno = nije upisano)
'
' Ulaz je RECNIK (Scripting.Dictionary) sa vrednostima polja, da ga moze
' napuniti bilo koja forma. Kljucevi su LOGICKI, ne imena kontrola:
'
'   datum, stanicaID, kooperantID, vrsta, sorta, tipAmb, vozacID, brDok,
'   brojZbirne, parcelaID, primalac
'   kolicinaI, cenaI, kolAmb, kolAmbIzdata
'   dveKlase, kolicinaII, cenaII, kolAmbII
'   novac
'
' OtkupValidiraj UPISUJE nazad u recnik: kolicinaI/kolicinaII postaju NETO, a
' brutoKgI/brutoKgII zamrznuti uneti bruto (kad je OTKUP_BRUTO_UNOS ukljucen).
'
' Potvrde koje trazi operater (prekoracenje otpremnice, prosek po gajbici,
' neslaganje kulture parcele) ostaju MsgBox u ovom modulu - identicne su u obe
' forme, pa nema razloga da ih svaka postavlja po svom.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const OTKUNOS_BUILD As String = "v6-ui-114"

'--------------------------------------------------------------- ULAZ
' Prazan recnik sa svim kljucevima - da pozivalac ne mora da pamti spisak.
Public Function NoviOtkupUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("stanicaID") = ""
    p("kooperantID") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("tipAmb") = ""
    p("vozacID") = ""
    p("brDok") = ""
    p("brojZbirne") = ""
    p("parcelaID") = ""
    p("primalac") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("kolAmbIzdata") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("novac") = 0#
    p("brutoKgI") = 0#
    p("brutoKgII") = 0#
    Set NoviOtkupUnos = p
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

'---------------------------------------------------------- PROVERE
' Vraca "" kad je sve u redu; inace poruku za operatera. U fokus se upisuje
' LOGICKO ime polja koje treba istaci (vidi spisak kljuceva u zaglavlju).
'
' Redosled provera je isti kao u frmOtkup.btnUnos_Click - to nije stil nego
' ponasanje: operater je navikao koje ga polje prvo zaustavi.
Public Function OtkupValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim kolI As Double, cenI As Double, kolII As Double, cenII As Double
    Dim kolAmb As Long, kolAmbII As Long, kolAmbIzd As Long
    Dim imaKlasaI As Boolean, dveKl As Boolean
    Dim tara As Double, taraII As Double
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""

    If Len(S(p, "stanicaID")) = 0 Then
        fokus = "stanicaID": OtkupValidiraj = Poruka("OTKUNOS_ERR_OM"): Exit Function
    End If
    If Len(S(p, "kooperantID")) = 0 Then
        fokus = "kooperantID": OtkupValidiraj = Poruka("OTKUNOS_ERR_KOOP"): Exit Function
    End If
    If Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": OtkupValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If
    If IsValidacijaUnosa() And Len(S(p, "sorta")) = 0 Then
        fokus = "sorta": OtkupValidiraj = Poruka("OTKUNOS_ERR_SORTA"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    cenI = D(p, "cenaI")
    kolII = D(p, "kolicinaII")
    cenII = D(p, "cenaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    kolAmbIzd = L(p, "kolAmbIzdata")
    dveKl = B(p, "dveKlase")
    imaKlasaI = (kolI > 0)

    ' Klasa I je opciona SAMO kad je ukljucena Klasa II i I ostavljena prazna
    ' (unosi se samo II klasa). Tada ambalaza I MORA biti prazna.
    If imaKlasaI Then
        If cenI <= 0 Then
            fokus = "cenaI": OtkupValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
        End If
    Else
        If Not dveKl Then
            fokus = "kolicinaI": OtkupValidiraj = Poruka("OTKUI_ERR_KOLICINA"): Exit Function
        End If
        If kolAmb > 0 Then
            fokus = "kolAmb": OtkupValidiraj = Poruka("DOK_MSG_UNOSI_SAMO_KLASA"): Exit Function
        End If
    End If

    If dveKl Then
        If kolII <= 0 Then
            fokus = "kolicinaII": OtkupValidiraj = Poruka("OTKUNOS_ERR_KOLICINA_II"): Exit Function
        End If
        If cenII <= 0 Then
            fokus = "cenaII": OtkupValidiraj = Poruka("OTKUNOS_ERR_CENA_II"): Exit Function
        End If
    End If

    If (kolAmb > 0 Or kolAmbII > 0 Or kolAmbIzd > 0) And Len(S(p, "tipAmb")) = 0 Then
        fokus = "tipAmb": OtkupValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
    End If

    ' Broj gajbi je OBAVEZAN za svaku unetu klasu kad je validacija ukljucena.
    ' Kad je iskljucena, obavezan je samo u bruto rezimu - bez njega se bruto ne
    ' pretvara u neto, pa bi se tezina gajbi platila kao voce.
    If IsValidacijaUnosa() Then
        If kolI > 0 And kolAmb <= 0 Then
            fokus = "kolAmb": OtkupValidiraj = Poruka("OTKUNOS_ERR_GAJBE_I"): Exit Function
        End If
        If dveKl And kolII > 0 And kolAmbII <= 0 Then
            fokus = "kolAmbII": OtkupValidiraj = Poruka("OTKUNOS_ERR_GAJBE_II"): Exit Function
        End If
    Else
        If OtkupBrutoUnos() And kolI > 0 And kolAmb <= 0 Then
            fokus = "kolAmb": OtkupValidiraj = Poruka("OTKUP_MSG_BRUTO_REZIM_UNESITE"): Exit Function
        End If
        If dveKl And OtkupBrutoUnos() And kolII > 0 And kolAmbII <= 0 Then
            fokus = "kolAmbII": OtkupValidiraj = Poruka("OTKUP_MSG_BRUTO_REZIM_UNESITE_2"): Exit Function
        End If
    End If

    ' --- BRUTO -> NETO. Operater unosi bruto (voce + ambalaza); u Kolicinu ide
    ' neto, a uneti bruto se zamrzava u BrutoKg. Tara se vezuje za klasu ciji su
    ' to gajbici. ---
    If OtkupBrutoUnos() And kolAmb > 0 Then
        tara = kolAmb * GetTezinaGajbice(S(p, "tipAmb"))
        If tara <= 0 Then
            fokus = "tipAmb"
            OtkupValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                             Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If tara >= kolI Then
            fokus = "kolicinaI"
            OtkupValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _
                             " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgI") = kolI
        kolI = kolI - tara
        p("kolicinaI") = kolI
    End If

    If dveKl And OtkupBrutoUnos() And kolAmbII > 0 Then
        taraII = kolAmbII * GetTezinaGajbice(S(p, "tipAmb"))
        If taraII <= 0 Then
            fokus = "tipAmb"
            OtkupValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                             Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If taraII >= kolII Then
            fokus = "kolicinaII"
            OtkupValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE_KLASE") & Format$(taraII, "#,##0.00") & _
                             " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgII") = kolII
        kolII = kolII - taraII
        p("kolicinaII") = kolII
    End If

    If IsValidacijaUnosa() And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": OtkupValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If

    ' Dupli broj: ISTA kapija koju drzi pisac, ne druga implementacija.
    '
    ' Zatecena provera je isla kroz CheckDuplicate(broj, datum) -- BEZ stanice --
    ' pa je bila uza od pisca: dokument koji CreateOtkup_TX smatra legalnim (isti
    ' broj, druga stanica, isti dan) ekran bi odbio pre nego sto pisac dobije
    ' priliku. Pravilo i opseg zive u modOtkup; ovde je samo rana povratna
    ' informacija operateru.
    If Len(S(p, "brDok")) > 0 And Len(S(p, "stanicaID")) > 0 Then
        Dim zauzeo As String
        zauzeo = modOtkup.BrojDokumentaZauzet(S(p, "stanicaID"), _
                                             CDate(p("datum")), S(p, "brDok"))
        If Len(zauzeo) > 0 Then
            fokus = "brDok"
            OtkupValidiraj = Poruka("OTKUNOS_ERR_BROJ_ZAUZET") & " " & zauzeo
            Exit Function
        End If
    End If

    ' Kultura parcele vs izabrana vrsta - pitanje, ne greska.
    If Len(S(p, "parcelaID")) > 0 Then
        Dim parK As String
        parK = NzToText(LookupValue(TBL_PARCELE, COL_PAR_ID, S(p, "parcelaID"), COL_PAR_KULTURA))
        If Len(parK) > 0 And Len(S(p, "vrsta")) > 0 Then
            If StrComp(parK, S(p, "vrsta"), vbTextCompare) <> 0 Then
                If MsgBox(Poruka("OTKUNOS_ASK_PARCELA_1") & " (" & parK & ") " & _
                          Poruka("OTKUNOS_ASK_PARCELA_2") & " (" & S(p, "vrsta") & ")." & vbCrLf & _
                          Poruka("OTKUP_MSG_ZELITE_IPAK_NASTAVITE"), _
                          vbExclamation + vbYesNo, APP_NAME) = vbNo Then
                    fokus = "parcelaID": OtkupValidiraj = " ": Exit Function
                End If
            End If
        End If
    End If

    ' Prosek neto kg po gajbici (pragovi iz tblKulture) - sam pita operatera.
    If Not OtkupProsekGajbiceOK(S(p, "vrsta"), kolI, kolAmb, kolII, kolAmbII) Then
        fokus = "kolicinaI": OtkupValidiraj = " ": Exit Function
    End If
    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error Resume
    ' Next, koji cisti Err - operater bi inace dobio poruku bez objasnjenja.
    errDesc = Err.description
    LogErr "modOtkupUnos.OtkupValidiraj"
    OtkupValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' (Vrsta, Sorta) -> KulturaID. Razresavanje je posao ADAPTERA, ne pisca (S4.1f),
' ali pravilo zivi na JEDNOM mestu -- isti razresivac koristi i PWA ingest.
Private Function RazresiKulturu(ByVal vrsta As String, ByVal sorta As String, _
                                ByRef outGreska As String) As String
    Dim detalj As String
    RazresiKulturu = modOtkup.RazresiKulturuIzVrsteSorte(vrsta, sorta, detalj)

    If Len(RazresiKulturu) = 0 Then
        outGreska = Poruka("OTKUNOS_ERR_KULTURA") & " " & vrsta & " / " & sorta
    End If
End Function

Private Function OtkStavkaDTO(ByVal klasa As String, ByVal kol As Double, _
                              ByVal cena As Double, ByVal amb As Double, _
                              ByVal bruto As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "Cena", cena
    s.Add "KolAmbalaze", amb
    If bruto > 0 Then s.Add "BrutoKg", bruto
    Set OtkStavkaDTO = s
End Function

'------------------------------------------------------------- UPIS
' Upisuje otkup i radi sve sto ide uz njega. Vraca OtkupID (ili spojene ID-eve
' obe klase); prazno znaci da upis nije uspeo. U "poruke" se skupljaju
' napomene koje pozivalac prikazuje posle uspeha.
Public Function OtkupUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, hlPending As String, hlNewPrij As String, hlWarn As String
    Dim doHlRelink As Boolean, hlRelWarn As String, hlGajbDiff As Boolean
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    ' ISPRAVKA HLADNJACKOG LANCA JE FAIL-CLOSED DOK LANAC STOJI.
    '
    ' Pending relink znaci: operater je stornirao stari hladnjacki dokument i
    ' izabrao "Uneti ispravku". Ispravka nije samo nov otkup -- ona trazi i nov
    ' nizvodni lanac, da bi se palete stare prijemnice imale gde prevezati.
    '
    ' Lanac je pauziran do PR7, pa se taj drugi deo ne moze izvrsiti. Zato se
    ' STAJE PRE PISCA: bez ovoga bi nastao nov otkup, pending bi bio POTROSEN, a
    ' operater bi dobio samo "nema prijemnice" -- pola ispravke, i to nepovratno.
    If Len(GetHladnjacaRelinkPending()) > 0 Then
        poruke = poruke & Poruka("OTKUNOS_MSG_ISPRAVKA_PAUZIRANA") & " " & _
                 GetHladnjacaRelinkPending() & vbCrLf
        Exit Function
    End If

    ' KANONSKI PISAC. Sta vise NE ide u upis, i zasto:
    '
    '   vozacID     vozac pripada otpremnici, ne otkupu (S4.1c)
    '   brojZbirne  broj nikad nije bio veza nego labela (A2)
    '   novac       kes ne ulazi kroz otkupni list (S4.1b). Mereno: ovaj ekran
    '               p("novac") NIKAD ne postavlja -- polje pripada modNovacUnos,
    '               pa je grana u starom piscu bila mrtva na izvoru.
    '   primalac    isto
    Dim kulturaID As String
    kulturaID = RazresiKulturu(S(p, "vrsta"), S(p, "sorta"), errDesc)
    If Len(kulturaID) = 0 Then
        poruke = poruke & errDesc
        Exit Function
    End If

    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", CDate(p("datum"))
    h.Add "KooperantID", S(p, "kooperantID")
    h.Add "StanicaID", S(p, "stanicaID")
    h.Add "KulturaID", kulturaID
    h.Add "VrstaVoca", S(p, "vrsta")
    h.Add "SortaVoca", S(p, "sorta")
    h.Add "TipAmbalaze", S(p, "tipAmb")
    h.Add "BrojDokumenta", S(p, "brDok")
    h.Add "ParcelaID", S(p, "parcelaID")
    h.Add "KolAmbIzdata", L(p, "kolAmbIzdata")

    Dim stavke As Collection
    Set stavke = New Collection
    If D(p, "kolicinaI") > 0 Then
        stavke.Add OtkStavkaDTO(KLASA_I, D(p, "kolicinaI"), D(p, "cenaI"), _
                                L(p, "kolAmb"), D(p, "brutoKgI"))
    End If
    If B(p, "dveKlase") And D(p, "kolicinaII") > 0 Then
        stavke.Add OtkStavkaDTO(KLASA_II, D(p, "kolicinaII"), D(p, "cenaII"), _
                                L(p, "kolAmbII"), D(p, "brutoKgII"))
    End If

    Dim greska As String
    res = CreateOtkup_TX(h, stavke, greska)

    If Len(res) = 0 Then
        poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & greska
        Exit Function
    End If

    ' Stampa otkupnog lista - best-effort, greska ne sme da obori potvrdu upisa.
    On Error Resume Next
    OutputOtkupniList res
    Err.Clear
    On Error GoTo EH

    ' AUTO-LANAC JE PAUZIRAN (Otkup cutover, korak 2).
    '
    ' Lanac deli dokument PO KLASI: iz "ID1 + ID2" vadi idI i idII i svaku klasu
    ' vodi kroz svoju otpremnicu, zbirnu i prijemnicu (modAutoHladnjaca:182).
    ' Nov pisac daje JEDAN OtkupID, a jedan otkup red drzi JEDAN OtpremnicaID --
    ' pa je veza strukturno gubitna dok otpremnica ne predje na header+stavke.
    '
    ' Odluka operatera: lanac se gasi do PR7, umesto da se upisuje polovicna veza.
    ' Kod lanca OSTAJE netaknut -- pauzira se poziv, i to glasno.
    If IsHladnjacaStanica(S(p, "stanicaID")) Then
        poruke = poruke & Poruka("OTKUNOS_MSG_LANAC_PAUZIRAN") & vbCrLf
    End If
    SetPaletizeSkip False        ' toggle se vraca i kad lanac nije ni pokrenut

    ' Relink aparat je NEDOSTIZAN dok je lanac pauziran: pending se hvata iznad,
    ' pre pisca, pa dovde nikad ne stigne sa vrednoscu. Ostaje netaknut da ga PR7
    ' vrati u pogon zajedno sa lancem.
    If doHlRelink Then
        SetHladnjacaRelinkPending ""         ' potrosi (idempotentno)
        If Len(hlNewPrij) = 0 Then
            poruke = poruke & Poruka("OTKUNOS_MSG_NEMA_PRIJEMNICE") & vbCrLf
        ElseIf ReassignPaleteToPrijemnica_TX(hlPending, hlNewPrij, hlRelWarn, True, hlGajbDiff) Then
            poruke = poruke & Poruka("OTKUNOS_MSG_PALETE_PREVEZANE") & " " & _
                     hlPending & " " & ChrW(8594) & " " & hlNewPrij & vbCrLf
            If Len(hlRelWarn) > 0 Then poruke = poruke & hlRelWarn & vbCrLf
            If hlGajbDiff Then poruke = poruke & PaletaAdjustPrompt(hlNewPrij) & vbCrLf
        Else
            LogRelinkFailure hlPending, hlNewPrij, hlRelWarn
            poruke = poruke & Poruka("OTKUNOS_MSG_PALETE_NISU") & " " & hlRelWarn & vbCrLf
        End If
    End If

    OtkupUpisi = res
    Exit Function
EH:
    ' isto pravilo: opis pre logovanja (LogErr / Poruka cuvaju svoj Err)
    errDesc = Err.description
    SetPaletizeSkip False        ' toggle ne sme da ostane ukljucen ni na gresci
    LogErr "modOtkupUnos.OtkupUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function
