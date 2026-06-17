Attribute VB_Name = "modPaletniListUI"
Option Explicit

' ============================================================
' modPaletniListUI — operater-facing ulazi (InputBox/MsgBox) za paletni
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

    Dim broj As Long: broj = CLng(Val(ans))
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
    MsgBox "Greska pri izvozu paletnog lista: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8: broj prerade (tekuca godina) -> PDF preradnog lista.
Public Sub ExportPreradaPDF_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj prerade (godina " & Year(Date) & "):", "Preradni list -> PDF")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub

    Dim broj As Long: broj = CLng(Val(ans))
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
    MsgBox "Greska pri izvozu preradnog lista: " & Err.description, vbCritical, APP_NAME
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
        pbr = CLng(Val(Trim$(parts(i))))
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
    preID = SavePrerada_TX(ids, CLng(Val(sk)), CLng(Val(se)), _
                           CDbl(Val(Replace(sn, ",", "."))), snap)
    If preID <> "" Then ExportPreradaPDF preID, True
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
    MsgBox "Greska pri stampi nepotpunih paleta: " & Err.description, vbCritical, APP_NAME
End Sub

' Alt+F8 / dugme: rucno zatvori otvorenu paletu po broju (tekuca godina).
Public Sub ClosePaleta_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj palete za rucno zatvaranje (godina " & Year(Date) & "):", _
                   "Zatvori paletu")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    Dim broj As Long: broj = CLng(Val(ans))
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
