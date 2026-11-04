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

' TEST SEAM (KI-086): tekst poslednjeg pitanja korekcije kolicina ("" = nije
' postavljeno). U test rezimu PaletaAdjustPrompt ne otvara MsgBox (u nevidljivom
' Excelu visi), nego BELEZI pitanje i bira "Otkazi" -- samo zabranjuje, ne menja.
' Gasi ga modTest.ResetSeamova (PaletaAdjustPitanjeTestReset).
Private mAdjustPitanjeTest As String

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

' Cetiri Alt+F8 makroa koja su menjala palete i preradu po BROJU i tekucoj godini
' (ClosePaleta_Prompt, StornoPaleta_Prompt, SavePrerada_Prompt, StornoPrerada_Prompt)
' obrisana su u S8 rezu 1b (PAL-ODL-04): birala su prvi pogodak broja, a dupli broj
' palete u godini moze da postoji (integritet C5). Ekran Palete pokriva iste radnje
' po ID-u (C-052, C-053, C-055, C-058).

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

' ============================================================
' KOREKCIJA KOLICINA (Adjust) -- zajednicki UI tok za frmDokumenta/frmOtkup/rucno.
' Poziva AdjustPaletaGajbiceZaPrijemnicu_TX; ako visak gajbica ne staje na paletu,
' operater bira: PRELIJ na sledecu paletu ili PREKO kapaciteta na istoj.
' Vraca tekst-rezime za prikaz (ili "" ako nije bilo nista / operater odustao).
'
' Dokument je IDENTITET (S8 rez 1a, N1): pozivalac salje PrijemnicaID koji vec
' drzi. Broj je samo labela u pitanju operateru -- izbor po broju je korigovao i
' palete drugog kupca istog broja.
' ============================================================
' outKorekcijaCeka (KI-086): True kad je korekcija bila potrebna, a NIJE uradjena
' (ADJ_BLOCKED, "Otkazi" na pitanju, greska) -- pozivalac to vidi kao podatak, ne
' samo kroz tekst. False kad je uradjena ili nije bila potrebna.
Public Function PaletaAdjustPrompt(ByVal prijemnicaID As String, ByVal brojPrij As String, _
                                   Optional ByRef outKorekcijaCeka As Boolean) As String
    outKorekcijaCeka = False
    On Error GoTo EH
    Dim info As String
    Dim res As Long
    res = AdjustPaletaGajbiceZaPrijemnicu_TX(prijemnicaID, "", info)

    If res = ADJ_NEEDS_CHOICE Then
        Dim ans As VbMsgBoxResult, pitanje As String
        pitanje = "Korekcija gajbica za prijemnicu " & brojPrij & ":" & vbCrLf & info & vbCrLf & vbCrLf & _
                  "DA  = preliti visak na sledecu (otvorenu/novu) paletu" & vbCrLf & _
                  "NE  = slozi SVE na istu paletu, PREKO kapaciteta" & vbCrLf & _
                  "OTKAZI = ne diraj palete sada"
        If IsTestMode() Then
            mAdjustPitanjeTest = pitanje
            ans = vbCancel
        Else
            ans = MsgBox(pitanje, vbYesNoCancel + vbQuestion, APP_NAME)
        End If
        If ans = vbYes Then
            res = AdjustPaletaGajbiceZaPrijemnicu_TX(prijemnicaID, "PRELIJ", info)
        ElseIf ans = vbNo Then
            res = AdjustPaletaGajbiceZaPrijemnicu_TX(prijemnicaID, "PREKO", info)
        Else
            outKorekcijaCeka = True
            PaletaAdjustPrompt = "Korekcija kolicina na paletama PRESKOCENA (uradi kasnije: " & _
                                 "Alt+F8 -> PaletaAdjust_Prompt)."
            Exit Function
        End If
    End If

    If res = ADJ_BLOCKED Then
        outKorekcijaCeka = True
        PaletaAdjustPrompt = "Korekcija kolicina NIJE uradjena: " & info
    ElseIf res > 0 Then
        PaletaAdjustPrompt = "Kolicine na paletama korigovane. " & info
    Else
        PaletaAdjustPrompt = info      ' 0 = nista za korekciju
    End If
    Exit Function
EH:
    LogErr "modPaletniListUI.PaletaAdjustPrompt"
    outKorekcijaCeka = True
    PaletaAdjustPrompt = "Greska pri korekciji kolicina: " & Err.description
End Function

' TEST SEAM (KI-086): tekst poslednjeg pitanja korekcije. Tvrdo gejtovan.
Public Function PaletaAdjustPitanjeTest() As String
    If Not IsTestMode() Then Exit Function
    PaletaAdjustPitanjeTest = mAdjustPitanjeTest
End Function

' TEST SEAM (KI-086): brise zabelezeno pitanje -- test pre poteza, i ResetSeamova.
Public Sub PaletaAdjustPitanjeTestReset()
    mAdjustPitanjeTest = ""
End Sub

' Alt+F8: rucna korekcija kolicina na paletama za AKTIVNU prijemnicu (escape kad
' je auto-korekcija ranije preskocena, ili Integritet A3/A4 prijavi odstupanje).
' Operater kuca broj; korekcija ide po dokumentu u koji se broj razresi
' (PrijemnicaZaKorekciju, PAL-ODL-04: makro ostaje, ali strog).
Public Sub PaletaAdjust_Prompt()
    On Error GoTo EH
    Dim ans As String, razlog As String, prjID As String
    ans = InputBox("Broj prijemnice cije kolicine treba uskladiti sa paletama:", _
                   "Korekcija kolicina na paletama")
    If Trim$(ans) = "" Then Exit Sub
    prjID = PrijemnicaZaKorekciju(ans, razlog)
    If Len(prjID) = 0 Then
        MsgBox "Korekcija kolicina NIJE uradjena: " & razlog, vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim msg As String: msg = PaletaAdjustPrompt(prjID, Trim$(ans))
    If Len(msg) > 0 Then MsgBox msg, vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "modPaletniListUI.PaletaAdjust_Prompt"
    MsgBox "Gre" & ChrW(353) & "ka: " & Err.description, vbCritical, APP_NAME
End Sub

' Broj prijemnice -> PrijemnicaID TACNO JEDNOG aktivnog dokumenta; prazno uz
' razlog kad ga nema ili kad ih je vise.
'
' Broj se racuna po kupcu (modStorno:1848), pa ga dva kupca legitimno nose
' istovremeno. Izbor "prvog pogotka" bi tada korigovao palete tudjeg
' dokumenta; razlog zato imenuje sve kupce, da operater zna zasto je odbijen i
' koji dokument da trazi (Oporavak ili ispravka, gde je dokument poznat).
Public Function PrijemnicaZaKorekciju(ByVal broj As String, ByRef outRazlog As String) As String
    Const SRC As String = "modPaletniListUI.PrijemnicaZaKorekciju"
    outRazlog = ""
    broj = Trim$(broj)
    If Len(broj) = 0 Then
        outRazlog = "Broj prijemnice nije unet."
        Exit Function
    End If

    Dim d As Variant: d = GetTableData(TBL_PRIJEMNICA)
    Dim ids As Collection: Set ids = New Collection
    Dim kupci As String
    If Not IsEmpty(d) Then
        Dim cId As Long, cBr As Long, cKu As Long, cSt As Long, r As Long
        cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
        cBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
        cKu = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, SRC)
        cSt = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, SRC)
        For r = 1 To UBound(d, 1)
            If BrojJednak(d(r, cBr), broj) _
               And UCase$(Trim$(NzToText(d(r, cSt)))) <> "DA" Then
                ids.Add Trim$(NzToText(d(r, cId)))
                If Len(kupci) > 0 Then kupci = kupci & ", "
                kupci = kupci & KupacOznaka(Trim$(NzToText(d(r, cKu))))
            End If
        Next r
    End If

    Select Case ids.count
        Case 0
            outRazlog = "Nema aktivne prijemnice broja " & broj & "."
        Case 1
            PrijemnicaZaKorekciju = CStr(ids(1))
        Case Else
            outRazlog = "Broj " & broj & " nosi " & ids.count & " aktivnih prijemnica " & _
                        "(kupci: " & kupci & "). Korekcija po broju bi mogla da " & _
                        "zahvati tudje palete -- pokreni je iz Oporavka ili kroz ispravku, " & _
                        "gde je dokument poznat."
    End Select
End Function

' "Naziv (ID)" kupca za poruku; sam ID kad naziva nema. Javno: zovu ga i potvrde
' radnji Oporavka (modScrOporavak.OpisDokumentaZaPotvrdu).
Public Function KupacOznaka(ByVal kupacID As String) As String
    Dim nm As String
    On Error Resume Next
    nm = Trim$(NzToText(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, COL_KUP_NAZIV)))
    On Error GoTo 0
    If Len(nm) > 0 Then
        KupacOznaka = nm & " (" & kupacID & ")"
    Else
        KupacOznaka = kupacID
    End If
End Function

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

