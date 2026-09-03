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

Private mFrm As Object              ' okvir domacina (radna povrsina ljuske)
Private mWrappers As Collection     ' clsAdminBtn (drzi WithEvents zivim)

' Prekidac grupa -- isti obrazac koji Podesavanja vec koriste. Admin ima
' dvanaest komandi u pet grupa; do v6-ui-201 su sve stajale naslagane, sto je
' bio jedini ekran u aplikaciji koji tako izgleda.
Private mAktivnaGrupa As String
Private mSegBtns As Collection      ' key = naziv grupe -> segment dugme
Private mKomande As Collection      ' key = naziv grupe -> Collection dugmadi
Private mAdmW As Single

Private Const ADM_SEG_H As Single = 24
Private Const ADM_SEG_GAP As Single = 6
Private Const ADM_M As Single = PAD

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

    Const m As Single = PAD

    ' Naslov
    Dim lblTitle As MSForms.label
    Set lblTitle = AddLabel("adm_title", m, 14, w - 2 * m, 18)
    lblTitle.caption = Poruka("OTKUI_MS_ADMIN")
    modUiKit.PanelStilNaslov lblTitle

    ' Povratak (gore desno -- vidljiv pre skrolovanja)
    Dim btnBack As MSForms.CommandButton
    Set btnBack = AddButton("btnAdmBack", w - m - 132, 38, 132, 26)
    btnBack.caption = Poruka("OTKUI_BTN_PANEL_NAZAD")
    modUiKit.PanelStilDugme btnBack, "ghost"
    WireButton btnBack, "back"

    ' Hint
    Dim lblHint As MSForms.label
    Set lblHint = AddLabel("adm_hint", m, 40, w - m - 150, 15)
    lblHint.caption = "Operativne i razvojne komande. Neke su destruktivne -- koristi oprezno."
    modUiKit.PanelStilNapomena lblHint

    ' Grupe dugmadi (data-driven). Dodavanje komande = jedan red u AdminGroups.
    Dim groups As Variant
    groups = AdminGroups()

    ' Mere rasporeda stoje u RelayoutAdmin -- gradnja ih ne treba, jer kontrole
    ' pravi na nuli i tek ih raspored postavlja (isti obrazac kao Podesavanja).
    Const BTNH As Single = 28

    Dim gi As Long, ii As Long
    Dim grp As Variant, items As Variant, it As Variant
    Dim b As MSForms.CommandButton, seg As MSForms.CommandButton
    Dim cap As String, act As String, gname As String
    Dim lista As Collection

    mAdmW = w
    Set mSegBtns = New Collection
    Set mKomande = New Collection

    ' Gradnja: sve komande postoje, raspored ih pali i gasi. Isti obrazac kao
    ' polja u Podesavanjima -- crta se JEDNA grupa, ostale su sklonjene.
    For gi = LBound(groups) To UBound(groups)
        grp = groups(gi)
        gname = CStr(grp(0))

        Set seg = AddButton("admgrp_" & CStr(gi), ADM_M, 0, _
                            modUiKit.PanelSirinaSegmenta(gname), ADM_SEG_H)
        modUiKit.PanelStilSegment seg, gname, False
        ' Grupa se nosi U AKCIJI: clsAdminBtn nema polje za nosivost, a klasa se
        ' po pravilu ne siri (v. .claude/rules/forme-i-kontrole.md).
        WireButton seg, "grp:" & gname
        mSegBtns.Add seg, gname

        Set lista = New Collection
        items = grp(1)
        For ii = LBound(items) To UBound(items)
            it = items(ii)
            cap = CStr(it(0))
            act = CStr(it(1))
            Set b = AddButton("btnAdm_" & act, ADM_M, 0, 200, BTNH)
            b.caption = cap
            If act = "checkupdate" Then
                modUiKit.PanelStilDugme b, "primary"
            ElseIf act = "ocisti" Or act = "migracija" Then
                ' Jedine dve komande koje diraju podatke -- boja to kaze pre
                ' nego sto se klikne, ne tek u dijalogu potvrde.
                modUiKit.PanelStilDugme b, "danger"
            Else
                modUiKit.PanelStilDugme b, "ghost"
            End If
            WireButton b, act
            lista.Add b
        Next ii
        mKomande.Add lista, gname
        If gi = LBound(groups) Then mAktivnaGrupa = gname
    Next gi

    ' Skrol postavlja RelayoutAdmin -- on jedini zna dokle je raspored stigao.
    ' Do v6-ui-201 je isti blok stajao i ovde, sa promenljivom Y koju gradnja
    ' vise nema (raspored je preuzeo mere): "Variable not defined" pri compile-u.
    RelayoutAdmin

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

    ' Segment prekidaca nosi grupu u akciji ("grp:<naziv>") -- clsAdminBtn nema
    ' polje za nosivost, a klasa se ne siri.
    If Left$(action, 4) = "grp:" Then
        If mAktivnaGrupa <> Mid$(action, 5) Then
            mAktivnaGrupa = Mid$(action, 5)
            RelayoutAdmin
        End If
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
' Sve su idempotentne; svaka javlja svoj rezultat, pa finalni rezime.
Private Sub AdminEnsureEverything()
    On Error GoTo EH
    SetupNewPC
    EnsurePoruke
    EnsureCenovnikSchema
    EnsurePaletniListSchema
    EnsureDoradeSchema
    EnsureKorisniciSchema
    EnsureAuditColumns
    MsgBox "Ensure zavr" & ChrW(353) & "en (setup + sve " & ChrW(353) & "eme provereno).", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr "modAdmin.AdminEnsureEverything"
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbExclamation, APP_NAME
End Sub

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
' Oslobadja reference modula bez zatvaranja domacina -- zove ga modUiPanel pre
' nego sto isprazni okvir. Isti obrazac koji modPodesavanja vec ima
' (Podesavanja_Release): WithEvents omotaci moraju da odu PRE kontrola, inace
' drze kontrole kojih vise nema.
Public Sub Admin_Release()
    On Error Resume Next
    Set mFrm = Nothing
    Set mWrappers = Nothing
    Set mSegBtns = Nothing
    Set mKomande = Nothing
    mAktivnaGrupa = ""
End Sub

Private Sub CloseAdminPanel()
    On Error Resume Next
    ' Domacin moze biti FORMA (legacy put, frmStammdaten) ili OKVIR u radnoj
    ' povrsini nove ljuske (modUiPanel). Forma se gasi, okvir se vraca ekranu --
    ' pa se pita sta je domacin, umesto da se pretpostavi. Ova grana nestaje
    ' zajedno sa legacy formom.
    '
    ' Pita se po MODULU, ne po kljucu panela: modul zna svoje ime, a kljuc je
    ' strano ime koje se moze preimenovati u registru (i jeste, u v6-ui-201).
    If modUiPanel.PanelZatvoriAko("modAdmin") Then Exit Sub
    frmOtkupAPP.ReturnToDashboard "Admin zatvoren."
    Unload mFrm
    Set mFrm = Nothing
    Set mWrappers = Nothing
End Sub

' ============================================================
' PRIVATE -- runtime control helperi (Controls.Add; .frx se ne dira)
' ============================================================
' Raspored: prekidac grupa, pa komande SAMO aktivne grupe. Isti ritam kao
' Podesavanja -- segmenti se prelamaju, komande idu u dve kolone.
Private Sub RelayoutAdmin()
    Dim gname As Variant, grp As String, seg As MSForms.CommandButton
    Dim X As Single, Y As Single, sw As Single, cx As Single
    Dim lista As Collection, i As Long, col As Long, btnW As Single
    Const COLS As Long = 2
    Const COLGAP As Single = 14
    Const BTNH As Single = 28
    Const ROWGAP As Single = 8
    On Error GoTo EH
    If mSegBtns Is Nothing Then Exit Sub

    X = ADM_M
    Y = 76
    For Each gname In AdminGrupeImena()
        grp = CStr(gname)
        Set seg = mSegBtns(grp)
        sw = modUiKit.PanelSirinaSegmenta(grp)
        If X > ADM_M And X + sw > mAdmW - ADM_M Then
            X = ADM_M
            Y = Y + ADM_SEG_H + ADM_SEG_GAP
        End If
        seg.Left = X: seg.top = Y: seg.width = sw: seg.Height = ADM_SEG_H
        seg.Visible = True
        modUiKit.PanelStilSegment seg, grp, (grp = mAktivnaGrupa)
        X = X + sw + ADM_SEG_GAP
    Next gname
    Y = Y + ADM_SEG_H + 18

    btnW = (mAdmW - 2 * ADM_M - (COLS - 1) * COLGAP) / COLS
    For Each gname In AdminGrupeImena()
        grp = CStr(gname)
        Set lista = mKomande(grp)
        If grp <> mAktivnaGrupa Then
            For i = 1 To lista.count
                lista(i).Visible = False
            Next i
        Else
            col = 0
            For i = 1 To lista.count
                cx = ADM_M + col * (btnW + COLGAP)
                lista(i).Left = cx
                lista(i).top = Y
                lista(i).width = btnW
                lista(i).Height = BTNH
                lista(i).Visible = True
                If col = COLS - 1 Then
                    col = 0
                    Y = Y + BTNH + ROWGAP
                Else
                    col = col + 1
                End If
            Next i
            If col <> 0 Then Y = Y + BTNH + ROWGAP
        End If
    Next gname

    On Error Resume Next
    mFrm.ScrollBars = fmScrollBarsVertical
    mFrm.ScrollHeight = Y + 16
    mFrm.KeepScrollBarsVisible = fmScrollBarsVertical
    Exit Sub
EH:
    LogErr "modAdmin.RelayoutAdmin"
End Sub

' Imena grupa u redosledu registra -- raspored ih mora obici tim redom, a
' Collection sa kljucem to ne garantuje.
Private Function AdminGrupeImena() As Variant
    Dim g As Variant, a() As Variant, i As Long
    g = AdminGroups()
    ReDim a(LBound(g) To UBound(g))
    For i = LBound(g) To UBound(g)
        a(i) = CStr(g(i)(0))
    Next i
    AdminGrupeImena = a
End Function

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
    mFrm.Controls.Remove CStr(nm)
    On Error GoTo 0
End Sub
