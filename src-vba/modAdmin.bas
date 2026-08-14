Attribute VB_Name = "modAdmin"
Option Explicit

' ============================================================
' modAdmin -- Admin panel (runtime) u frmStammdaten, sekcija "Admin"
' iz menija Maticni podaci (grupa "Sistem", ispod "Podesavanja").
'
' Tok (isti kao Podesavanja):
'   modMaticniLookups.MaticniSekcijeGrupisano -> frmMaticniPodaci.OpenSekcija ->
'   frmStammdaten (Tag = "Admin") -> UserForm_Activate -> BuildAdminPanel.
'
' Kontrole se grade u RUNTIME-u (Controls.Add) -- frmStammdaten.frx se NE dira,
' isti obrazac kao modPodesavanja/clsConfigBtn i modMaticniLookups/clsLookupMenuBtn.
' Klik svakog dugmeta hvata clsAdminBtn (WithEvents) -> AdminPanel_OnClick.
'
' Panel NE implementira novu logiku -- svako dugme zove POSTOJECU javnu ulaznu
' tacku (reuse). Mapa:
'   Proveri azuriranje  -> modUpdateGate.ReleaseManifestVersion + OnTime RunSelfUpdate
'   Ensure (setup+seme) -> modSetup.SetupNewPC + Ensure* seme (agregat)
'   Health check setup  -> modSetup.RunSetupHealthCheck
'   Production health   -> modProductionHealthCheck.RunProductionHealthCheck
'   Google autorizacija -> modGoogleAuth.RunGoogleAuthSetup
'   Objavi na Drive     -> modRelease.PublishReleaseToDrive  (build/dev)
'   VBA Import / Export -> modVbaTools.ImportAllVBA / ExportAllVBA  (dev)
'   Otvori VBA editor   -> modPregledListova.OtvoriVBA
'   Migracija           -> modMigracija.MigrirajPodatkeIzStarog  (ima svoju potvrdu)
'   Ocisti tabele       -> modPregledListova.OcistiTabele       (ima svoju potvrdu)
'
' ASCII-only modul (vidi CLAUDE.md, sekcija 4): dijakritika u captionima ide
' iskljucivo preko ChrW, kao u modMaticniLookups.
' ============================================================

Private mFrm As Object              ' frmStammdaten instanca (host)
Private mWrappers As Collection     ' clsAdminBtn (drzi WithEvents zivim)

' ============================================================
' PUBLIC -- izgradnja panela (poziva frmStammdaten.UserForm_Activate za Tag="Admin")
' ============================================================
Public Sub BuildAdminPanel(ByVal frm As Object)
    Const SRC As String = "modAdmin.BuildAdminPanel"
    On Error GoTo EH

    ' AUD-033: tvrda brana -- Admin panel gradi samo admin (ili dok je AUTH iskljucen;
    ' MozeAdministraciju je anti-lockout). Defense-in-depth uz meni gate (modMaticniLookups).
    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    Set mFrm = frm
    Set mWrappers = New Collection

    ' Sakri sve postojece (maticni-podaci) kontrole -- gradimo svoj panel preko.
    Dim ctl As MSForms.Control
    For Each ctl In frm.Controls
        On Error Resume Next
        ctl.Visible = False
        On Error GoTo EH
    Next ctl

    Dim w As Single
    w = frm.InsideWidth
    If w < 400 Then w = 960

    Const m As Single = 12

    ' Naslov
    Dim lblTitle As MSForms.label
    Set lblTitle = AddLabel("adm_title", m, 8, w - 2 * m, 20)
    lblTitle.caption = "Admin"
    StyleLabel lblTitle, TXT_LIGHT(), True
    lblTitle.Font.Size = FONT_SIZE_HEADER

    ' Povratak (gore desno -- vidljiv pre skrolovanja)
    Dim btnBack As MSForms.CommandButton
    Set btnBack = AddButton("btnAdmBack", w - m - 120, 32, 120, 24)
    StyleExitButton btnBack, "Povratak"
    WireButton btnBack, "back"

    ' Hint
    Dim lblHint As MSForms.label
    Set lblHint = AddLabel("adm_hint", m, 36, w - m - 140, 16)
    lblHint.caption = "Operativne i razvojne komande. Neke su destruktivne -- koristi oprezno."
    StyleLabel lblHint, TXT_MUTED(), False
    lblHint.Font.Size = FONT_SIZE_SMALL

    ' Grupe dugmadi (data-driven). Dodavanje komande = jedan red u AdminGroups.
    Dim groups As Variant
    groups = AdminGroups()

    Const COLS As Long = 2
    Const COLGAP As Single = 12
    Const BTNH As Single = 26
    Const ROWGAP As Single = 6
    Const HDRH As Single = 16
    Const HDRGAP As Single = 3
    Const GROUPGAP As Single = 12
    Dim btnW As Single
    btnW = (w - 2 * m - (COLS - 1) * COLGAP) / COLS

    Dim Y As Single: Y = 64
    Dim gi As Long, ii As Long
    Dim grp As Variant, items As Variant, it As Variant
    Dim hl As MSForms.label, b As MSForms.CommandButton
    Dim col As Long, cx As Single
    Dim cap As String, act As String

    For gi = LBound(groups) To UBound(groups)
        grp = groups(gi)

        ' Sekcijski header (labela)
        Set hl = AddLabel("admgrp_" & CStr(gi), m, Y, w - 2 * m, HDRH)
        hl.caption = UCase$(CStr(grp(0)))
        StyleLabel hl, TXT_MUTED(), True
        hl.Font.Size = FONT_SIZE_SMALL
        Y = Y + HDRH + HDRGAP

        items = grp(1)
        col = 0
        For ii = LBound(items) To UBound(items)
            it = items(ii)
            cap = CStr(it(0))
            act = CStr(it(1))

            cx = m + col * (btnW + COLGAP)
            Set b = AddButton("btnAdm_" & act, cx, Y, btnW, BTNH)
            If act = "checkupdate" Then
                StylePrimaryButton b, cap
            Else
                StyleMenuButton b, cap
            End If
            WireButton b, act

            If col = COLS - 1 Then
                col = 0
                Y = Y + BTNH + ROWGAP
            Else
                col = col + 1
            End If
        Next ii
        If col <> 0 Then Y = Y + BTNH + ROWGAP
        Y = Y + GROUPGAP
    Next gi

    On Error Resume Next
    mFrm.ScrollBars = fmScrollBarsVertical
    mFrm.ScrollHeight = Y + 16
    mFrm.KeepScrollBarsVisible = fmScrollBarsVertical
    On Error GoTo 0

    Exit Sub
EH:
    LogErr SRC
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbCritical, APP_NAME
End Sub

' Registar komandi: Array(GrupaNaziv, Array(Array(Caption, action), ...)).
' Redosled = redosled u panelu. Gradi se preko Collection da se izbegne VBA
' limit "Too many line continuations" (isti razlog kao modPodesavanja.CfgAdd).
Private Function AdminGroups() As Variant
    Dim g As Collection: Set g = New Collection
    Dim a As Collection

    Set a = New Collection
    a.Add Array("Proveri a" & ChrW(382) & "uriranje", "checkupdate")
    g.Add Array("A" & ChrW(382) & "uriranje", CollToArr(a))

    Set a = New Collection
    a.Add Array("Ensure (setup + " & ChrW(353) & "eme)", "ensure")
    a.Add Array("Health check (setup)", "healthsetup")
    a.Add Array("Production health check", "healthprod")
    a.Add Array("Integritet provere (tabele)", "integritet")
    g.Add Array("Setup i provere", CollToArr(a))

    Set a = New Collection
    a.Add Array("Google autorizacija", "googleauth")
    a.Add Array("Objavi release na Drive", "publish")
    g.Add Array("Google / Drive", CollToArr(a))

    Set a = New Collection
    a.Add Array("VBA Import", "vbaimport")
    a.Add Array("VBA Export", "vbaexport")
    a.Add Array("Otvori VBA editor", "vbaopen")
    g.Add Array("VBA (dev)", CollToArr(a))

    Set a = New Collection
    a.Add Array("Migracija iz starog fajla", "migracija")
    a.Add Array("O" & ChrW(269) & "isti tabele od podataka", "ocisti")
    g.Add Array("Podaci (oprezno)", CollToArr(a))

    AdminGroups = CollToArr(g)
End Function

Private Function CollToArr(ByVal c As Collection) As Variant
    If c.count = 0 Then CollToArr = Array(): Exit Function
    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    CollToArr = a
End Function

' ============================================================
' PUBLIC -- click ruter (zove clsAdminBtn). Svaka grana = reuse postojece tacke.
' ============================================================
Public Sub AdminPanel_OnClick(ByVal action As String)
    On Error GoTo EH

    ' AUD-033: tvrda brana i na akcijama (ne samo na izgradnji panela).
    If Not modAuth.MozeAdministraciju() Then
        MsgBox Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    Select Case LCase$(action)
        Case "checkupdate":  AdminCheckUpdate
        Case "ensure":       AdminEnsureEverything
        Case "healthsetup":  RunSetupHealthCheck
        Case "healthprod":   RunProductionHealthCheck
        Case "integritet":   RunIntegritetProvere
        Case "googleauth":   RunGoogleAuthSetup
        Case "publish":      AdminPublishToDrive
        Case "vbaimport":    AdminVbaImport
        Case "vbaexport":    ExportAllVBA
        Case "vbaopen":      OtvoriVBA
        Case "migracija":    MigrirajPodatkeIzStarog
        Case "ocisti":       OcistiTabele
        Case "back":         CloseAdminPanel
    End Select
    Exit Sub
EH:
    LogErr "modAdmin.AdminPanel_OnClick(" & action & ")"
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbExclamation, APP_NAME
End Sub

' ============================================================
' PRIVATE -- akcije koje zahtevaju omotac (potvrda / OnTime / agregacija)
' ============================================================

' Rucna provera azuriranja: read-only modUpdateGate.ReleaseManifestVersion
' (NE dira modSelfUpdate/SKIP -> self-update-safe) + VersionCompare. Panel SAM
' vodi dijalog, pa daje JASAN feedback i kad nema novije (bez dvostruke poruke
' koju bi dao reuse CheckForUpdateOnOpen).
' ZAMKA: RunSelfUpdate NIKAD direktno -- preko OnTime (prazan stack), kao modMain.
Private Sub AdminCheckUpdate()
    On Error GoTo EH
    Dim remote As String
    remote = ReleaseManifestVersion()

    If Len(remote) > 0 And VersionCompare(APP_VERSION, remote) < 0 Then
        If MsgBox("Dostupno je a" & ChrW(382) & "uriranje: " & remote & vbCrLf & _
                  "Trenutna verzija: " & APP_VERSION & vbCrLf & vbCrLf & _
                  "A" & ChrW(382) & "urirati sada? (program " & ChrW(263) & "e se zatvoriti i ponovo otvoriti)", _
                  vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            CloseAdminPanel
            Application.OnTime Now, "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!RunSelfUpdate"
        End If
    Else
        MsgBox "Nema novih a" & ChrW(382) & "uriranja. Koristite najnoviju verziju (" & APP_VERSION & ")." & _
               IIf(Len(remote) = 0, vbCrLf & "(Napomena: kanal a" & ChrW(382) & "uriranja trenutno nije dostupan -- offline / nije objavljeno.)", ""), _
               vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modAdmin.AdminCheckUpdate"
End Sub

' Agregat "ensure everything": glavni setup + sve pojedinacne seme (reuse).
' Sve su idempotentne.
'
' Ranije je ovo zvalo MsgBox-varijante jezgara, pa je jedan klik davao CETIRI
' dijaloga -- a gore od toga, pad koraka je zavrsavao u sopstvenom MsgBox-u i
' procedura se normalno vracala, pa je agregat odmah zatim javljao "sve
' provereno". Sada se zovu TIHA jezgra (*Core), rezultat se skuplja i na kraju
' se javlja STVARNO stanje. Jedini tudji dijalog je SetupNewPC -- on je Setup*
' ulazna tacka i njegov izvestaj nosi dijagnostiku koju nema smisla gutati.
Private Sub AdminEnsureEverything()
    Dim problems As String
    Dim fails As Long

    On Error GoTo EH
    SetupNewPC                       ' prikazuje svoj izvestaj (setup health-check)

    On Error Resume Next

    ' SetupNewPC nema povratnu vrednost: na nalaze upise APP_SETUP_COMPLETED="NE",
    ' prikaze svoj warning i vrati se NORMALNO -- a i njegov EH zavrsi isto tako.
    ' Bez ove provere agregat je posle "setup ima probleme" mogao da javi
    ' "setup i sve seme su provereni". IsSetupHealthy cita bas taj ishod
    ' (APP_SETUP_COMPLETED + core folderi + obavezne tabele).
    Err.Clear
    If Not IsSetupHealthy() Then
        problems = problems & "- Setup (SetupNewPC): health-check nije prosao" & vbCrLf
    End If

    Err.Clear: EnsurePoruke
    problems = problems & EnsureStepLine("Poruke (tblPoruke)", 0)

    ' EnsureCenovnikSchema se NE zove posebno -- EnsurePaletniListSchema ga
    ' vec zove, pa je raniji direktan poziv bio cist duplikat.
    ' fails = 0 pre poziva: pod "Resume Next" dodela izostane ako poziv pukne, pa
    ' bi promenljiva zadrzala broj iz PRETHODNOG koraka i lazno ga prijavila.
    Err.Clear: fails = 0: fails = EnsurePaletniListSchema()
    problems = problems & EnsureStepLine("Paletni list + cenovnik", fails)

    Err.Clear: fails = 0: fails = EnsureDoradeSchema()
    problems = problems & EnsureStepLine("Dorade + runtime schema", fails)

    Err.Clear: EnsureKorisniciSchema
    problems = problems & EnsureStepLine("Korisnici", 0)

    Err.Clear: EnsureAuditColumns
    problems = problems & EnsureStepLine("Audit kolone", 0)

    On Error GoTo EH

    If Len(problems) = 0 Then
        MsgBox Poruka("SETUP_MSG_ENSURE_SVE_PROSLO"), vbInformation, APP_NAME
    Else
        MsgBox Poruka("SETUP_MSG_ENSURE_SA_PADOVIMA") & vbCrLf & vbCrLf & problems & _
               vbCrLf & "Detalji: list SETUP_LOG (Level = ERROR).", _
               vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modAdmin.AdminEnsureEverything"
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbExclamation, APP_NAME
End Sub

' Rezultat jednog Ensure koraka -> red izvestaja (prazno kad je proslo).
' Pokriva oba oblika signala koja jezgra danas koriste: Err (jezgra koja
' re-raise-uju) i broj palih koraka (jezgra koja nastavljaju posle pada).
' errNum/errDesc: konvencija projekta za EH promenljive.
Private Function EnsureStepLine(ByVal labelText As String, ByVal failCount As Long) As String
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description
    Err.Clear

    If errNum <> 0 Then
        EnsureStepLine = "- " & labelText & ": " & errDesc & " (" & errNum & ")" & vbCrLf
    ElseIf failCount > 0 Then
        EnsureStepLine = "- " & labelText & ": " & failCount & " korak(a) nije proslo" & vbCrLf
    End If
End Function

' Objava na Drive je outward-facing (ceo fleet) i build/dev komanda -> trazi
' sifru (RELEASE_PUBLISH_SIFRA, modConfig). Unos sifre je ujedno i potvrda
' (jedan dijalog); prazno/Cancel = tiho odustao, pogresna sifra = poruka + prekid.
Private Sub AdminPublishToDrive()
    On Error GoTo EH
    Dim s As String
    s = InputBox("Objaviti trenutni kod (src-vba) i version.json na Google Drive" & vbCrLf & _
                 "(AgriX_Release -- vidljivo celom fleetu)? Ovo je build/dev komanda." & vbCrLf & vbCrLf & _
                 "Unesite " & ChrW(353) & "ifru za objavu:", APP_NAME)
    If Len(s) = 0 Then Exit Sub
    If s <> RELEASE_PUBLISH_SIFRA Then
        MsgBox "Pogre" & ChrW(353) & "na " & ChrW(353) & "ifra. Objava je otkazana.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    PublishReleaseToDrive
    Exit Sub
EH:
    LogErr "modAdmin.AdminPublishToDrive"
End Sub

' VBA Import prepisuje postojeci kod -> potvrda (putanja je fiksna, dev masina).
Private Sub AdminVbaImport()
    On Error GoTo EH
    If MsgBox("Uvezi sve VBA module iz src-vba foldera?" & vbCrLf & _
              "Ovo PREPISUJE postojeci kod (putanja je fiksna -- dev masina).", _
              vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub
    ImportAllVBA
    Exit Sub
EH:
    LogErr "modAdmin.AdminVbaImport"
End Sub

' Povratak na dashboard + cleanup (isti obrazac kao modPodesavanja.CloseConfigEditor).
Private Sub CloseAdminPanel()
    On Error Resume Next
    frmOtkupAPP.ReturnToDashboard "Admin zatvoren."
    Unload mFrm
    Set mFrm = Nothing
    Set mWrappers = Nothing
End Sub

' ============================================================
' PRIVATE -- runtime control helperi (Controls.Add; .frx se ne dira)
' ============================================================
Private Sub WireButton(ByVal b As MSForms.CommandButton, ByVal act As String)
    Dim wrp As clsAdminBtn
    Set wrp = New clsAdminBtn
    wrp.action = act
    Set wrp.btn = b
    mWrappers.Add wrp
End Sub

Private Function AddLabel(ByVal nm As String, ByVal X As Single, ByVal Y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.label
    RemoveCtl nm
    Dim c As MSForms.label
    Set c = mFrm.Controls.Add("Forms.Label.1", nm, True)
    c.Left = X: c.top = Y: c.width = w: c.Height = h
    Set AddLabel = c
End Function

Private Function AddButton(ByVal nm As String, ByVal X As Single, ByVal Y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    RemoveCtl nm
    Dim c As MSForms.CommandButton
    Set c = mFrm.Controls.Add("Forms.CommandButton.1", nm, True)
    c.Left = X: c.top = Y: c.width = w: c.Height = h
    Set AddButton = c
End Function

Private Sub RemoveCtl(ByVal nm As String)
    On Error Resume Next
    mFrm.Controls.Remove nm
    On Error GoTo 0
End Sub
