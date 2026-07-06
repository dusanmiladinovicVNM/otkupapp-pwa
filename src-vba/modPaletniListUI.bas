Attribute VB_Name = "modPaletniListUI"
'Attribute VB_Name = "modPaletniListUI"
Option Explicit

' ============================================================
' modPaletniListUI -- operater-facing ulazi (InputBox/MsgBox) za paletni
' list i preradu. Drzi UI/business granicu: business sloj (modPaletniList)
' ne koristi MsgBox za kontrolu toka; ovde su Alt+F8 stubovi dok ne stigne
' frmPalete (PR #44). Resolucija broj->PaletaID i poziv SavePrerada_TX su
' tanki; sva poslovna pravila/transakcija su u modPaletniList.
' ============================================================

' Alt+F8: broj palete (tekuca godina) -> PDF paletnog lista.
Public Sub ExportPaletniListPDF_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj palete (godina " & Year(Date) & "):", "Paletni list -> PDF")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub

    Dim broj As Long: broj = CLng(val(ans))
    If broj <= 0 Then Exit Sub

    Dim palID As String: palID = FindPaletaIDByBroj(broj, Year(Date))
    If palID = "" Then
        MsgBox "Nije nadjena paleta br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ExportPaletniListPDF palID, True
    Exit Sub
EH:
    LogErr "modPaletniListUI.ExportPaletniListPDF_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka pri izvozu paletnog lista: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8: broj prerade (tekuca godina) -> PDF preradnog lista.
Public Sub ExportPreradaPDF_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj prerade (godina " & Year(Date) & "):", "Preradni list -> PDF")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub

    Dim broj As Long: broj = CLng(val(ans))
    If broj <= 0 Then Exit Sub

    Dim preID As String: preID = FindPreradaIDByBroj(broj, Year(Date))
    If preID = "" Then
        MsgBox "Nije nadjena prerada br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ExportPreradaPDF preID, True
    Exit Sub
EH:
    LogErr "modPaletniListUI.ExportPreradaPDF_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka pri izvozu preradnog lista: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8: unos paleta (brojevi, zarezom) + kutije/kese/neto/napomena ->
' SavePrerada_TX + PDF. Privremeni stub dok ne stigne frmPalete (PR #44).
Public Sub SavePrerada_Prompt()
    On Error GoTo EH

    Dim sp As String
    sp = InputBox("Brojevi paleta za preradu (zarezom, npr. 1,2,5):", "Prerada")
    If Trim$(sp) = "" Then Exit Sub

    Dim yr As Long: yr = Year(Date)
    Dim ids As Collection: Set ids = New Collection
    Dim parts() As String: parts = Split(sp, ",")
    Dim i As Long, pbr As Long, pid As String
    For i = LBound(parts) To UBound(parts)
        pbr = CLng(val(Trim$(parts(i))))
        If pbr > 0 Then
            pid = FindPaletaIDByBroj(pbr, yr)
            If pid = "" Then
                MsgBox "Paleta " & pbr & "/" & yr & " ne postoji.", vbExclamation, APP_NAME
                Exit Sub
            End If
            ids.Add pid
        End If
    Next i

    Dim sk As String: sk = InputBox("Broj kutija:", "Prerada", "0")
    If StrPtr(sk) = 0 Then Exit Sub
    Dim se As String: se = InputBox("Broj kesa:", "Prerada", "0")
    If StrPtr(se) = 0 Then Exit Sub
    Dim sn As String: sn = InputBox("Neto izlaz (kg):", "Prerada", "0")
    If StrPtr(sn) = 0 Then Exit Sub
    Dim snap As String: snap = InputBox("Napomena (opciono):", "Prerada", "")

    Dim preID As String
    preID = SavePrerada_TX(ids, CLng(val(sk)), CLng(val(se)), _
                           CDbl(val(Replace(sn, ",", "."))), snap)
    If preID <> "" Then OutputPreradaList preID
    Exit Sub
EH:
    LogErr "modPaletniListUI.SavePrerada_Prompt"
    MsgBox "Prerada nije sacuvana: " & Err.description, vbExclamation, APP_NAME
End Sub

' Alt+F8 / dugme: izlaz za sve nepotpune palete + poruka operateru.
Public Sub PrintNepotpunePalete_Prompt()
    On Error GoTo EH
    Dim n As Long: n = PrintNepotpunePalete()
    If n = 0 Then
        MsgBox "Nema otvorenih (nepotpunih) paleta.", vbInformation, APP_NAME
    Else
        MsgBox n & " nepotpunih paleta poslato na izlaz (po PALETA_PRINT_MODE).", _
               vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modPaletniListUI.PrintNepotpunePalete_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi nepotpunih paleta: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8 / dugme: rucno zatvori otvorenu paletu po broju (tekuca godina).
Public Sub ClosePaleta_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj palete za rucno zatvaranje (godina " & Year(Date) & "):", _
                   "Zatvori paletu")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    Dim broj As Long: broj = CLng(val(ans))
    If broj <= 0 Then Exit Sub

    Dim palID As String: palID = FindPaletaIDByBroj(broj, Year(Date))
    If palID = "" Then
        MsgBox "Nije nadjena paleta br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ClosePaletaManual_TX palID
    MsgBox "Paleta " & broj & "/" & Year(Date) & " je zatvorena.", vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "modPaletniListUI.ClosePaleta_Prompt"
    MsgBox "Paleta nije zatvorena: " & Err.description, vbExclamation, APP_NAME
End Sub

' Alt+F8: storno palete po broju (tekuca godina), uz potvrdu.
Public Sub StornoPaleta_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj palete za STORNO (godina " & Year(Date) & "):", "Storno palete")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    Dim broj As Long: broj = CLng(val(ans))
    If broj <= 0 Then Exit Sub

    Dim palID As String: palID = FindPaletaIDByBroj(broj, Year(Date))
    If palID = "" Then
        MsgBox "Nije nadjena paleta br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    If MsgBox("Stornirati paletu " & broj & "/" & Year(Date) & "?", _
              vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    If StornoPaleta_TX(palID) Then
        MsgBox "Paleta " & broj & "/" & Year(Date) & " je stornirana.", vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo (vidi log).", vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modPaletniListUI.StornoPaleta_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka pri stornu: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8: storno prerade po broju (tekuca godina), uz potvrdu.
Public Sub StornoPrerada_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj prerade za STORNO (godina " & Year(Date) & "):", "Storno prerade")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    Dim broj As Long: broj = CLng(val(ans))
    If broj <= 0 Then Exit Sub

    Dim preID As String: preID = FindPreradaIDByBroj(broj, Year(Date))
    If preID = "" Then
        MsgBox "Nije nadjena prerada br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    If MsgBox("Stornirati preradu " & broj & "/" & Year(Date) & "? Palete se vracaju u lager.", _
              vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    If StornoPrerada_TX(preID) Then
        MsgBox "Prerada " & broj & "/" & Year(Date) & " je stornirana.", vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo (vidi log).", vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modPaletniListUI.StornoPrerada_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka pri stornu: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' KOREKCIJA KOLICINA (Adjust) -- zajednicki UI tok za frmDokumenta/frmOtkup/rucno.
' Poziva AdjustPaletaGajbiceZaPrijemnicu_TX; ako visak gajbica ne staje na paletu,
' operater bira: PRELIJ na sledecu paletu ili PREKO kapaciteta na istoj.
' Vraca tekst-rezime za prikaz (ili "" ako nije bilo nista / operater odustao).
' ============================================================
Public Function PaletaAdjustPrompt(ByVal brojPrij As String) As String
    On Error GoTo EH
    Dim info As String
    Dim res As Long
    res = AdjustPaletaGajbiceZaPrijemnicu_TX(brojPrij, "", info)

    If res = ADJ_NEEDS_CHOICE Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Korekcija gajbica za prijemnicu " & brojPrij & ":" & vbCrLf & info & vbCrLf & vbCrLf & _
                     "DA  = preliti visak na sledecu (otvorenu/novu) paletu" & vbCrLf & _
                     "NE  = slozi SVE na istu paletu, PREKO kapaciteta" & vbCrLf & _
                     "OTKAZI = ne diraj palete sada", _
                     vbYesNoCancel + vbQuestion, APP_NAME)
        If ans = vbYes Then
            res = AdjustPaletaGajbiceZaPrijemnicu_TX(brojPrij, "PRELIJ", info)
        ElseIf ans = vbNo Then
            res = AdjustPaletaGajbiceZaPrijemnicu_TX(brojPrij, "PREKO", info)
        Else
            PaletaAdjustPrompt = "Korekcija kolicina na paletama PRESKOCENA (uradi kasnije: " & _
                                 "Alt+F8 -> PaletaAdjust_Prompt)."
            Exit Function
        End If
    End If

    If res = ADJ_BLOCKED Then
        PaletaAdjustPrompt = "Korekcija kolicina NIJE uradjena: " & info
    ElseIf res > 0 Then
        PaletaAdjustPrompt = "Kolicine na paletama korigovane. " & info
    Else
        PaletaAdjustPrompt = info      ' 0 = nista za korekciju
    End If
    Exit Function
EH:
    LogErr "modPaletniListUI.PaletaAdjustPrompt"
    PaletaAdjustPrompt = "Greska pri korekciji kolicina: " & Err.description
End Function

' Alt+F8: rucna korekcija kolicina na paletama za AKTIVNU prijemnicu (escape kad
' je auto-korekcija ranije preskocena, ili Integritet A3/A4 prijavi odstupanje).
Public Sub PaletaAdjust_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj prijemnice cije kolicine treba uskladiti sa paletama:", _
                   "Korekcija kolicina na paletama")
    If Trim$(ans) = "" Then Exit Sub
    Dim msg As String: msg = PaletaAdjustPrompt(Trim$(ans))
    If Len(msg) > 0 Then MsgBox msg, vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "modPaletniListUI.PaletaAdjust_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' Sifarnik helperi za frmPalete (tip/tezina kutija i kesa, vrste gotovog
' proizvoda). Reuse GetLookupList/LookupValue; samo aktivni u izboru.
' ============================================================

' Distinct aktivni tipovi kutija (dropdown na frmPalete).
Public Function GetKutijeOptions() As Variant
    GetKutijeOptions = GetLookupList(TBL_KUTIJE, COL_KUT_TIP, , , True)
End Function

' Distinct aktivni tipovi kesa.
Public Function GetKeseOptions() As Variant
    GetKeseOptions = GetLookupList(TBL_KESE, COL_KES_TIP, , , True)
End Function

' Distinct aktivni tipovi gotovog proizvoda.
Public Function GetVrstaGPOptions() As Variant
    GetVrstaGPOptions = GetLookupList(TBL_VRSTA_GP, COL_VGP_TIP, , , True)
End Function

' Tezina (kg) prazne kutije za dati tip; 0 ako nije nadjeno.
Public Function GetTezinaKutije(ByVal tip As String) As Double
    On Error Resume Next
    If Trim$(tip) = "" Then Exit Function
    Dim v As Variant
    v = LookupValue(TBL_KUTIJE, COL_KUT_TIP, Trim$(tip), COL_KUT_TEZINA)
    If IsNumeric(v) Then GetTezinaKutije = CDbl(v)
End Function

' Tezina (kg) prazne kese za dati tip; 0 ako nije nadjeno.
Public Function GetTezinaKese(ByVal tip As String) As Double
    On Error Resume Next
    If Trim$(tip) = "" Then Exit Function
    Dim v As Variant
    v = LookupValue(TBL_KESE, COL_KES_TIP, Trim$(tip), COL_KES_TEZINA)
    If IsNumeric(v) Then GetTezinaKese = CDbl(v)
End Function

