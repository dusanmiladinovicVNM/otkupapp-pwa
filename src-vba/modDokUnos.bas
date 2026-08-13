Attribute VB_Name = "modDokUnos"
'=====================================================================
' modDokUnos - UNOS DOKUMENATA (otpremnica, zbirna, prijemnica), bez
' ijedne kontrole.
'
' Isti razlog i isti oblik kao modOtkupUnos, samo za rezime iz
' frmDokumenta. Poslovni posao (provere, bruto->neto, upis, ono sto ide
' posle upisa) ne sme da zivi u formi, jer ga onda drugi ekran ne moze
' pozvati bez prepisivanja.
'
'   OtpremnicaValidiraj(p, fokus)  provere + bruto->neto; vraca poruku o
'                                  gresci ("" = proslo) i LOGICKO ime
'                                  polja na koje treba vratiti fokus
'   OtpremnicaUpisi(p, poruke)     SaveOtpremnicaMulti_TX + auto-zbirna
'                                  (MALINA) + zavrsetak ispravke; vraca
'                                  BrojOtpremnice (prazno = nije upisano)
'
' Ulaz je RECNIK sa LOGICKIM imenima polja (NoviOtpremnicaUnos):
'
'   datum, stanicaID, vozacID, brDok, brojZbirne, vrsta, sorta, tipAmb
'   kolicinaI, cenaI, kolAmb
'   dveKlase, kolicinaII, cenaII, kolAmbII
'
' OtpremnicaValidiraj UPISUJE nazad: kolicinaI/kolicinaII postaju NETO, a
' brutoKgI/brutoKgII zamrznuti uneti bruto (kad je OTKUP_BRUTO_UNOS ON).
'
' RAZLIKE U ODNOSU NA OTKUPNI LIST (nisu greske - tako je u legacy):
'   - vozac je OBAVEZAN (otkupni list ga ne trazi)
'   - vrsta i sorta su obavezne SAMO uz VALIDACIJA_UNOSA
'   - cena I je obavezna samo uz VALIDACIJA_UNOSA; inace sme i 0
'   - gajbe i tip ambalaze su obavezni samo uz VALIDACIJA_UNOSA
'   - nema parcele, nema izdate ambalaze, nema proseka po gajbici
'
' VAZNO: legacy frmDokumenta OSTAJE netaknut i potpuno operativan. Ovaj
' modul je drugi pozivalac istih poslovnih rutina, ne zamena za formu.
' Dok oba sistema ne budu potpuna, obe kopije postoje namerno.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const DOKUNOS_BUILD As String = "v6-ui-115"

'--------------------------------------------------------------- ULAZ
Public Function NoviOtpremnicaUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("stanicaID") = ""
    p("vozacID") = ""
    p("brDok") = ""
    p("brojZbirne") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("tipAmb") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("brutoKgI") = 0#
    p("brutoKgII") = 0#
    Set NoviOtpremnicaUnos = p
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
' Vraca "" kad je sve u redu; inace poruku za operatera. Redosled provera
' je isti kao u frmDokumenta.btnUnosOtp_Click - to nije stil nego
' ponasanje: operater je navikao koje ga polje prvo zaustavi.
Public Function OtpremnicaValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim kolI As Double, cenI As Double, kolII As Double, cenII As Double
    Dim kolAmb As Long, kolAmbII As Long
    Dim imaKlasaI As Boolean, dveKl As Boolean, strogo As Boolean
    Dim tara As Double, taraII As Double
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    If Len(S(p, "stanicaID")) = 0 Then
        fokus = "stanicaID": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_OM"): Exit Function
    End If
    ' Otpremnica bez vozaca ne postoji - roba nekim putem ide sa otkupnog mesta.
    If Len(S(p, "vozacID")) = 0 Then
        fokus = "vozacID": OtpremnicaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If
    If strogo And Len(S(p, "sorta")) = 0 Then
        fokus = "sorta": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_SORTA"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    cenI = D(p, "cenaI")
    kolII = D(p, "kolicinaII")
    cenII = D(p, "cenaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    dveKl = B(p, "dveKlase")
    imaKlasaI = (kolI > 0)

    ' Klasa I je opciona SAMO uz ukljucenu Klasu II (unosi se samo II klasa).
    ' Tada ambalaza I mora ostati prazna.
    If Not imaKlasaI Then
        If Not dveKl Then
            fokus = "kolicinaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_KOLICINA"): Exit Function
        End If
        If kolAmb > 0 Then
            fokus = "kolAmb": OtpremnicaValidiraj = Poruka("DOK_MSG_UNOSI_SAMO_KLASA"): Exit Function
        End If
    End If

    ' Cena I: obavezna samo uz strogu validaciju. Van nje otpremnica sme da ode
    ' bez cene (cena stize sa cenovnikom ili kasnije), ali ne sme biti negativna.
    If strogo And imaKlasaI Then
        If cenI <= 0 Then
            fokus = "cenaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
        End If
    ElseIf cenI < 0 Then
        fokus = "cenaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
    End If

    If dveKl Then
        If kolII <= 0 Then
            fokus = "kolicinaII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_KOLICINA_II"): Exit Function
        End If
        If cenII <= 0 Then
            fokus = "cenaII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_CENA_II"): Exit Function
        End If
    End If

    ' Gajbe i tip ambalaze su obavezni samo uz strogu validaciju - drugacije nego
    ' kod otkupnog lista, gde ih bruto rezim trazi i bez nje.
    If strogo Then
        If imaKlasaI And kolAmb <= 0 Then
            fokus = "kolAmb": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_I"): Exit Function
        End If
        If dveKl And kolAmbII <= 0 Then
            fokus = "kolAmbII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_II"): Exit Function
        End If
        If (kolAmb > 0 Or kolAmbII > 0) And Len(S(p, "tipAmb")) = 0 Then
            fokus = "tipAmb": OtpremnicaValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
        End If
    End If

    ' --- BRUTO -> NETO. Otpremnica se cuva u NETO, isto kao otkupni list, da
    ' panel blokova poredi neto sa neto. Uneti bruto se zamrzava u BrutoKg. ---
    If OtkupBrutoUnos() And kolAmb > 0 Then
        tara = kolAmb * GetTezinaGajbice(S(p, "tipAmb"))
        If tara <= 0 Then
            fokus = "tipAmb"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If tara >= kolI Then
            fokus = "kolicinaI"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _
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
            OtpremnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If taraII >= kolII Then
            fokus = "kolicinaII"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE_KLASE") & Format$(taraII, "#,##0.00") & _
                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgII") = kolII
        kolII = kolII - taraII
        p("kolicinaII") = kolII
    End If

    If strogo And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": OtpremnicaValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If

    If Len(S(p, "brDok")) > 0 Then
        Dim dup As String
        dup = CheckDuplicate(TBL_OTPREMNICA, COL_OTP_BROJ, S(p, "brDok"), COL_OTP_DATUM)
        If Len(dup) > 0 Then
            fokus = "brDok": OtpremnicaValidiraj = dup: Exit Function
        End If
    End If
    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error Resume
    ' Next, koji cisti Err - operater bi inace dobio poruku bez objasnjenja.
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaValidiraj"
    OtpremnicaValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'------------------------------------------------------------- UPIS
' Upisuje otpremnicu i radi sve sto ide uz nju. Vraca broj otpremnice;
' prazno znaci da upis nije uspeo. U "poruke" se skupljaju napomene koje
' pozivalac prikazuje posle uspeha.
Public Function OtpremnicaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, brZbrSave As String, createdZbr As Long, autoZbrErr As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    ' MALINA: otpremnica se snima sa PRAZNIM BrojZbirne da je auto-zbirna pokupi
    ' (broj u formi je samo predlog; auto-zbirna dodeli "S" + broj otpremnice).
    ' Van malina moda ide vrednost iz polja.
    If IsMalinaMode() Then brZbrSave = "" Else brZbrSave = S(p, "brojZbirne")

    res = SaveOtpremnicaMulti_TX( _
        datum:=CDate(p("datum")), _
        stanicaID:=S(p, "stanicaID"), _
        vozacID:=S(p, "vozacID"), _
        brojOtp:=S(p, "brDok"), _
        brojZbirne:=brZbrSave, _
        vrsta:=S(p, "vrsta"), _
        sorta:=S(p, "sorta"), _
        kolicinaI:=D(p, "kolicinaI"), _
        cenaI:=D(p, "cenaI"), _
        tipAmb:=S(p, "tipAmb"), _
        kolAmb:=L(p, "kolAmb"), _
        hasKlasaII:=B(p, "dveKlase"), _
        kolicinaII:=D(p, "kolicinaII"), _
        cenaII:=D(p, "cenaII"), _
        brutoKgI:=D(p, "brutoKgI"), _
        kolAmbII:=L(p, "kolAmbII"), _
        brutoKgII:=D(p, "brutoKgII"))

    If Len(res) = 0 Then Exit Function

    ' MALINA: otpremnica JESTE zbirna -> auto-zbirna iz upravo snimljene.
    ' Upravo snimljena je otvorena (brZbrSave=""), pa mora nastati bar jedna;
    ' tih pad se prijavljuje, da operater vidi da zbirne nema.
    If IsMalinaMode() Then
        On Error Resume Next
        createdZbr = AutoCreateZbirnaFromOtpremnice_TX()
        If Err.Number <> 0 Then
            autoZbrErr = Err.description
            LogErr "modDokUnos.OtpremnicaUpisi.AutoZbirna"
            Err.Clear
        End If
        On Error GoTo EH
        If Len(autoZbrErr) > 0 Or createdZbr < 1 Then
            poruke = poruke & Poruka("DOKUNOS_MSG_ZBIRNA_NIJE")
            If Len(autoZbrErr) > 0 Then poruke = poruke & " " & autoZbrErr
            poruke = poruke & vbCrLf
        End If
    End If

    ' ISPRAVKA_ODMAH: ako na cekanju stoji ispravka otpremnice, upravo snimljena
    ' je njena zamena -> prevezi blokove i rekalkulisi zbirnu. No-op inace.
    ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA, res, poruke

    OtpremnicaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'--------------------------------------------------------- ISPRAVKA
' Zavrsetak ispravke posle snimanja zamenskog dokumenta.
'
' Radi SAMO nad PERSISTENTNOM ispravkom na cekanju (tblStornoVeza), ne nad
' stanjem sesije: storno panel jos zivi iskljucivo u frmDokumenta (Faza D),
' pa ovaj modul nema odakle da zna sta je operater upravo stornirao u toj
' formi. Persistentan zapis prezivljava zatvaranje forme i Excela, pa je i
' dovoljan da se veza ne izgubi.
'
' SAFE-STOP kao u legacy: dve ili vise otvorenih ispravki istog tipa = ne
' biraj naslepo, nego pusti operatera kroz "Osiroceni dokumenti".
Private Sub ZavrsiIspravkuAko(ByVal docType As String, ByVal newBroj As String, _
                              ByRef poruke As String)
    Dim cnt As Long, cid As String, oldBroj As String, res As Object
    On Error GoTo EH
    newBroj = Trim$(newBroj)
    If Len(newBroj) = 0 Then Exit Sub

    cnt = modStornoContext.CountPendingCorrectionsByDocType(docType, SV_MODE_ISPRAVKA)
    If cnt = 0 Then Exit Sub                    ' nema ispravke -> obican unos
    If cnt > 1 Then
        poruke = poruke & Poruka("DOKUNOS_MSG_VISE_ISPRAVKI") & vbCrLf
        Exit Sub
    End If

    cid = modStornoContext.FindLatestPending(docType, SV_MODE_ISPRAVKA)
    If Len(cid) = 0 Then Exit Sub

    ' Potvrda je obavezna: operater je mozda napustio ispravku pa uneo DRUGI
    ' dokument - automatsko vezivanje bi tada spojilo pogresna dva.
    oldBroj = modStornoContext.GetCorrectionField(cid, COL_SV_OLD_BROJ)
    If MsgBox(Poruka("DOKUNOS_ASK_ISPRAVKA_1") & " '" & oldBroj & "'." & vbCrLf & vbCrLf & _
              Poruka("DOKUNOS_ASK_ISPRAVKA_2") & " '" & newBroj & "'?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    Select Case docType
        Case FLOW_DOC_OTPREMNICA: Set res = CompleteOtpremnicaIspravka(cid, newBroj)
        Case Else: Exit Sub
    End Select

    If res Is Nothing Then Exit Sub
    If CBool(res("success")) Then
        poruke = poruke & Poruka("DOKUNOS_MSG_ISPRAVKA_OK") & " " & CStr(res("message")) & vbCrLf
    Else
        poruke = poruke & Poruka("DOKUNOS_MSG_ISPRAVKA_NIJE") & " " & CStr(res("message")) & vbCrLf
    End If
    Exit Sub
EH:
    LogErr "modDokUnos.ZavrsiIspravkuAko"
End Sub
