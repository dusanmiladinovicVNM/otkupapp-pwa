Attribute VB_Name = "modPodesavanja"
Option Explicit

' ============================================================
' modPodesavanja — editor podesavanja (tblSEFConfig) u UI-u.
'
' Cilj: kad je tblSEFConfig sakriven (VeryHidden), operativna podesavanja se
' uredjuju kroz formu umesto rucnog editovanja celija. Otvara se kao nova
' sekcija "Podesavanja" u meniju Maticni podaci:
'   modMaticniLookups.MaticniSekcije -> frmMaticniPodaci.OpenSekcija ->
'   frmStammdaten (Tag = "Podesavanja") -> UserForm_Activate -> BuildConfigEditor.
'
' Kontrole se grade u RUNTIME-u (Controls.Add) — frmStammdaten.frx se NE dira,
' isti obrazac kao modMaticniLookups/clsLookupMenuBtn i modOtkupBlok/clsBlokUI.
' Klik Sacuvaj/Sakrij/Povratak hvata clsConfigBtn (WithEvents).
'
' BEZBEDNOST: prikazuju se SAMO operativna (slobodan-unos) polja. Interni kes /
' anti-tamper kljucevi (LICENSE_TOKEN, LICENSE_BOUND_PARTS, LICENSE_NEXT_CHECK,
' LICENSE_STATUS, LICENSE_HWM, TRIAL_HWM, GOOGLE_*_TOKEN, APP_SETUP_*...) se
' NAMERNO NE prikazuju — njihovo rucno menjanje je upravo bypass koji sakrivanje
' tabele zatvara.
'
' Izlaz u nuzdi (ako forma nije dostupna): Alt+F8 -> ShowConfigSheet.
' ============================================================

Private mFrm As Object              ' frmStammdaten instanca (host)
Private mInputs As Collection       ' input kontrole, key = ConfigKey
Private mWrappers As Collection     ' clsConfigBtn (drzi WithEvents zivim)
Private mBtnToggle As MSForms.CommandButton

' --- Registar polja: Array(Grupa, ConfigKey, Labela, Tip) ---
' Tip: "bool" (YES/NO), "list:A;B;C" (combo sa zadatim opcijama), "int",
'      "secret", "memo", "text".
' Dodavanje novog operativnog kljuca = jedan red ovde (data-driven, kao
' MaticniSekcije). NE dodavati interne kes kljuceve (vidi BEZBEDNOST gore).
Public Function ConfigEditorFields() As Variant
    ' Gradi se preko CfgAdd (Collection) da se izbegne VBA limit "Too many line
    ' continuations". Vraca 0-based Variant niz redova Array(Grupa,Key,Labela,Tip),
    ' pa potrosaci (BuildConfigEditor/SaveConfigEditor) ostaju nepromenjeni.
    Dim c As Collection: Set c = New Collection

    CfgAdd c, "Licenca", "LICENSE_ENABLED", "Licenciranje ukljuceno", "bool"
    CfgAdd c, "Licenca", "LICENSE_KEY", "Licencni kljuc", "text"
    CfgAdd c, "Licenca", "LICENSE_ENDPOINT", "Licencni endpoint (URL)", "text"

    CfgAdd c, "Probni period", "TRIAL_ENABLED", "Probni period ukljucen", "bool"
    CfgAdd c, "Probni period", "TRIAL_START", "Pocetak (yyyy-mm-dd)", "text"
    CfgAdd c, "Probni period", "TRIAL_DAYS", "Trajanje (dana)", "int"

    CfgAdd c, "Sinhronizacija", "CLOUD_SYNC_ENABLED", "Cloud sync ukljucen", "bool"
    CfgAdd c, "Sinhronizacija", "SHEETS_SYNC_ENABLED", "Google Sheets sync ukljucen", "bool"
    CfgAdd c, "Sinhronizacija", "SYNC_AUTO_INTERVAL_MIN", "Auto-sync interval (min, >=15)", "int"

    CfgAdd c, "Monitoring", "MONITORING_ENDPOINT", "Monitoring endpoint (URL)", "text"
    CfgAdd c, "Monitoring", "MONITORING_SECRET", "Monitoring secret", "secret"
    CfgAdd c, "Monitoring", "MONITORING_ENV", "Okruzenje (DEV/PROD)", "text"

    CfgAdd c, "Google", "GOOGLE_CLIENT_ID", "Google Client ID", "text"
    CfgAdd c, "Google", "GOOGLE_CLIENT_SECRET", "Google Client Secret", "secret"
    CfgAdd c, "Google", "GOOGLE_PWA_FOLDER_ID", "PWA Folder ID", "text"
    CfgAdd c, "Google", "GOOGLE_STAMMDATEN_SHEET_ID", "Stammdaten Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_KARTICE_SHEET_ID", "Kartice Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_MGMT_SHEET_ID", "Management Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_REPORTS_FOLDER_ID", "Reports Folder ID", "text"

    CfgAdd c, "SEF", "SEF_BASE_URL", "SEF Base URL", "text"
    CfgAdd c, "SEF", "SEF_API_KEY", "SEF API kljuc", "secret"
    CfgAdd c, "SEF", "SEF_ENV", "SEF okruzenje", "text"
    CfgAdd c, "SEF", "SEF_PAYMENT_DUE_DAYS", "Rok placanja (dana)", "int"
    CfgAdd c, "SEF", "SEF_PAYMENT_MEANS_CODE", "Sifra nacina placanja", "text"
    CfgAdd c, "SEF", "SEF_NOTE_DEFAULT", "Podrazumevana napomena", "text"
    CfgAdd c, "SEF", "SEF_FORCE_TODAY_ISSUE_DATE", "Forsiraj danasnji datum izdavanja", "bool"

    CfgAdd c, "Prodavac (firma)", "SELLER_NAME", "Naziv firme", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_PIB", "PIB", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_MATICNI_BROJ", "Maticni broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_STREET", "Ulica i broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_CITY", "Grad", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_POSTAL_CODE", "Postanski broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_COUNTRY_CODE", "Drzava (kod, npr. RS)", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_ACCOUNT", "Tekuci racun", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_EMAIL", "Email", "text"

    CfgAdd c, "Otkup / dokumenta", "OTKUP_KLAUZULA", "Klauzula (otkupni list)", "memo"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_ROK_ISPLATE", "Rok isplate (otkupni list)", "text"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_PRINT_MODE", "Štampa otkupnog lista", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, "Otkup / dokumenta", "PALETA_PRINT_MODE", "Štampa paletnog lista", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, "Otkup / dokumenta", "PDV_NADOKNADA_STOPA", "PDV nadoknada stopa (%)", "int"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_TIP_PALETE", "Podrazumevani tip palete", "text"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_BLOK_PANEL", "Panel za blokove (Otkup)", "bool"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_VRSTA_VOCA", "Podrazumevana vrsta voća (auto pri otvaranju)", "text"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_SORTA_VOCA", "Podrazumevana sorta voća (auto pri otvaranju)", "text"
    CfgAdd c, "Otkup / dokumenta", "KOOP_FILTER_BY_OM", "Filtriraj kooperante po otkupnom mestu", "bool"
    CfgAdd c, "Otkup / dokumenta", "AUTO_PRIJEMNICA_HLADNJACA", "Auto otpremnica+zbirna+prijemnica (OM=hladnjača)", "bool"

    CfgAdd c, "Malina režim", "MALINA_MODE", "Auto-zbirna iz otpremnice (1 stanica = 1 vozilo)", "bool"
    CfgAdd c, "Malina režim", "MALINA_DEFAULT_KUPAC", "Podrazumevani kupac (KupacID, auto-zbirna)", "text"

    CfgAdd c, "Alati / putanje", "PDFTOTEXT_EXE_PATH", "pdftotext.exe (banka import)", "text"

    CfgAdd c, "Management / Klijent", "MGMT_USER_1", "Management korisnik 1", "text"
    CfgAdd c, "Management / Klijent", "MGMT_USER_2", "Management korisnik 2", "text"
    CfgAdd c, "Management / Klijent", "MGMT_USER_3", "Management korisnik 3", "text"
    CfgAdd c, "Management / Klijent", "CLIENT_ID", "Client ID", "text"
    CfgAdd c, "Management / Klijent", "CLIENT_NAME", "Client naziv", "text"
    CfgAdd c, "Management / Klijent", "ENV", "Okruzenje (klijent, DEV/PROD)", "text"

    CfgAdd c, "Napredno / Test", "SEF_TEST_ALLOW_LIVE", "SEF test: dozvoli LIVE slanje", "bool"
    CfgAdd c, "Napredno / Test", "SEF_TEST_ALLOW_CANCEL_STORNO", "SEF test: dozvoli cancel/storno", "bool"
    CfgAdd c, "Napredno / Test", "SEF_DEBUG_LOG", "SEF debug log", "bool"

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    ConfigEditorFields = a
End Function

' Helper: dodaj jedan red u registar (izbegava line-continuation limit).
Private Sub CfgAdd(ByRef c As Collection, ByVal grp As String, ByVal key As String, _
                   ByVal lbl As String, ByVal typ As String)
    c.Add Array(grp, key, lbl, typ)
End Sub

' ============================================================
' PUBLIC — izgradnja editora (poziva frmStammdaten.UserForm_Activate za Tag)
' ============================================================
Public Sub BuildConfigEditor(ByVal frm As Object)
    Const SRC As String = "modPodesavanja.BuildConfigEditor"
    On Error GoTo EH

    Set mFrm = frm
    Set mInputs = New Collection
    Set mWrappers = New Collection
    Set mBtnToggle = Nothing

    ' Sakri sve postojece (maticni-podaci) kontrole — gradimo svoj panel preko.
    Dim ctl As MSForms.Control
    For Each ctl In frm.Controls
        On Error Resume Next
        ctl.Visible = False
        On Error GoTo EH
    Next ctl

    Dim w As Single
    w = frm.InsideWidth
    If w < 400 Then w = 960

    Const M As Single = 12
    Const LBLW As Single = 250
    Dim inLeft As Single: inLeft = M + LBLW + 10
    Dim inW As Single: inW = w - inLeft - M - 18      ' rezerva za scrollbar
    If inW < 120 Then inW = 120

    ' Naslov
    Dim lblTitle As MSForms.label
    Set lblTitle = AddLabel("cfg_title", M, 8, w - 2 * M, 20)
    lblTitle.caption = "Podešavanja (tblSEFConfig)"
    StyleLabel lblTitle, TXT_LIGHT(), True
    lblTitle.Font.Size = FONT_SIZE_HEADER

    ' Footer dugmad (na vrhu — vidljiva pre skrolovanja)
    Dim btnSave As MSForms.CommandButton
    Set btnSave = AddButton("btnCfgSave", M, 32, 120, 24)
    StylePrimaryButton btnSave, "Sačuvaj"
    WireButton btnSave, "save"

    Set mBtnToggle = AddButton("btnCfgToggle", M + 130, 32, 200, 24)
    StyleExitButton mBtnToggle, ToggleCaption()
    WireButton mBtnToggle, "toggle"

    Dim btnBack As MSForms.CommandButton
    Set btnBack = AddButton("btnCfgBack", w - M - 120, 32, 120, 24)
    StyleExitButton btnBack, "Povratak"
    WireButton btnBack, "back"

    Dim lblHint As MSForms.label
    Set lblHint = AddLabel("cfg_hint", M, 60, w - 2 * M, 16)
    lblHint.caption = "Interna polja (token, bound, status, HWM, OAuth token...) se namerno NE prikazuju."
    StyleLabel lblHint, TXT_MUTED(), False
    lblHint.Font.Size = FONT_SIZE_SMALL

    ' Polja
    Dim flds As Variant: flds = ConfigEditorFields()
    Dim y As Single: y = 86
    Dim curGroup As String: curGroup = ""
    Dim f As Variant, grp As String, key As String, cap As String, typ As String, typRaw As String
    Dim rowH As Single, cur As String
    Dim hdr As MSForms.label, lbl As MSForms.label
    Dim cmb As MSForms.ComboBox, tb As MSForms.TextBox
    Dim opts As Variant, oi As Long
    Dim i As Long

    For i = LBound(flds) To UBound(flds)
        f = flds(i)
        grp = CStr(f(0))
        key = CStr(f(1))
        cap = CStr(f(2))
        typRaw = CStr(f(3))
        typ = LCase$(typRaw)

        If grp <> curGroup Then
            curGroup = grp
            y = y + 8
            Set hdr = AddLabel("cfghdr_" & i, M, y, w - 2 * M, 18)
            hdr.caption = "— " & grp & " —"
            StyleLabel hdr, TXT_LIGHT(), True
            y = y + 22
        End If

        rowH = IIf(typ = "memo", 46, 18)

        Set lbl = AddLabel("cfglbl_" & key, M, y + 1, LBLW, 16)
        lbl.caption = cap
        StyleLabel lbl, TXT_MUTED(), False
        lbl.Font.Size = FONT_SIZE_SMALL

        cur = GetConfigValue(key)

        If typ = "bool" Or Left$(typ, 5) = "list:" Then
            Set cmb = AddCombo("cfg_" & key, inLeft, y, 160, 18)
            cmb.Style = fmStyleDropDownCombo
            If typ = "bool" Then
                opts = Array("YES", "NO")
            Else
                opts = Split(Mid$(typRaw, 6), ";")
            End If
            For oi = LBound(opts) To UBound(opts)
                cmb.AddItem Trim$(CStr(opts(oi)))
            Next oi
            cmb.value = cur
            StyleComboBox cmb
            mInputs.Add cmb, key
        Else
            Set tb = AddText("cfg_" & key, inLeft, y, inW, rowH)
            If typ = "memo" Then
                tb.MultiLine = True
                tb.WordWrap = True
                tb.ScrollBars = fmScrollBarsVertical
            End If
            tb.value = cur
            StyleTextBox tb
            mInputs.Add tb, key
        End If

        y = y + rowH + 8
    Next i

    ' Skrol forme (footer je na vrhu pa je dostupan na scroll=0)
    On Error Resume Next
    frm.ScrollBars = fmScrollBarsVertical
    frm.ScrollHeight = y + 16
    frm.KeepScrollBarsVisible = fmScrollBarsVertical
    On Error GoTo EH

    Exit Sub
EH:
    LogErr SRC
    MsgBox "Greška pri otvaranju podešavanja: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' PUBLIC — click ruter (zove clsConfigBtn)
' ============================================================
Public Sub ConfigEditor_OnClick(ByVal action As String)
    On Error GoTo EH
    Select Case LCase$(action)
        Case "save": SaveConfigEditor
        Case "toggle": ToggleConfigSheet
        Case "back": CloseConfigEditor
    End Select
    Exit Sub
EH:
    LogErr "modPodesavanja.ConfigEditor_OnClick"
End Sub

' ============================================================
' PRIVATE — save / back
' ============================================================
Private Sub SaveConfigEditor()
    Const SRC As String = "modPodesavanja.SaveConfigEditor"
    On Error GoTo EH
    If mInputs Is Nothing Then Exit Sub

    Dim flds As Variant: flds = ConfigEditorFields()
    Dim errs As String, n As Long
    Dim f As Variant, key As String, typ As String, v As String
    Dim i As Long

    For i = LBound(flds) To UBound(flds)
        f = flds(i)
        key = CStr(f(1))
        typ = LCase$(CStr(f(3)))

        v = ""
        On Error Resume Next
        v = Trim$(CStr(mInputs(key).value))
        On Error GoTo EH

        If typ = "int" And Len(v) > 0 And Not IsNumeric(v) Then
            errs = errs & " - " & key & " mora biti broj." & vbCrLf
        Else
            SetConfigValue key, v
            n = n + 1
        End If
    Next i

    If Len(errs) > 0 Then
        MsgBox "Sačuvano: " & n & " polja." & vbCrLf & vbCrLf & _
               "Preskočeno (greška):" & vbCrLf & errs, vbExclamation, APP_NAME
    Else
        MsgBox "Sačuvano: " & n & " polja.", vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr SRC
    MsgBox "Greška pri čuvanju: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub CloseConfigEditor()
    On Error Resume Next
    frmOtkupAPP.ReturnToDashboard "Podešavanja zatvorena."
    Unload mFrm
    Set mFrm = Nothing
    Set mInputs = Nothing
    Set mWrappers = Nothing
    Set mBtnToggle = Nothing
End Sub

' ============================================================
' PUBLIC — vidljivost tblSEFConfig sheet-a (toggle + Alt+F8 makroi)
' ============================================================
Public Sub ToggleConfigSheet()
    On Error GoTo EH
    If ConfigSheetIsHidden() Then
        ShowConfigSheet
        MsgBox "tblSEFConfig je sada VIDLJIV.", vbInformation, APP_NAME
    Else
        HideConfigSheet
        MsgBox "tblSEFConfig je sada SAKRIVEN (VeryHidden)." & vbCrLf & _
               "Uređuj ga isključivo preko ove forme." & vbCrLf & _
               "(Izlaz u nuždi: Alt+F8 -> ShowConfigSheet.)", vbInformation, APP_NAME
    End If
    If Not mBtnToggle Is Nothing Then mBtnToggle.caption = ToggleCaption()
    Exit Sub
EH:
    LogErr "modPodesavanja.ToggleConfigSheet"
End Sub

' Alt+F8 ulazne tacke (izlaz u nuzdi / setup).
Public Sub HideConfigSheet()
    On Error GoTo EH
    ConfigSheet().Visible = xlSheetVeryHidden
    Exit Sub
EH:
    LogErr "modPodesavanja.HideConfigSheet"
End Sub

Public Sub ShowConfigSheet()
    On Error GoTo EH
    ConfigSheet().Visible = xlSheetVisible
    Exit Sub
EH:
    LogErr "modPodesavanja.ShowConfigSheet"
End Sub

Private Function ToggleCaption() As String
    ToggleCaption = IIf(ConfigSheetIsHidden(), "Prikaži config tabelu", "Sakrij config tabelu")
End Function

Private Function ConfigSheetIsHidden() As Boolean
    On Error Resume Next
    ConfigSheetIsHidden = (ConfigSheet().Visible <> xlSheetVisible)
End Function

' Worksheet koji nosi tblSEFConfig (ListObject.Parent = Worksheet).
Private Function ConfigSheet() As Object
    Dim lo As ListObject
    Set lo = GetTable(TBL_SEF_CONFIG)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 7611, "modPodesavanja.ConfigSheet", _
                  "Tabela " & TBL_SEF_CONFIG & " ne postoji."
    End If
    Set ConfigSheet = lo.Parent
End Function

' ============================================================
' PRIVATE — runtime control helperi (Controls.Add; .frx se ne dira)
' ============================================================
Private Sub WireButton(ByVal b As MSForms.CommandButton, ByVal act As String)
    Dim wrp As clsConfigBtn
    Set wrp = New clsConfigBtn
    wrp.action = act
    Set wrp.btn = b
    mWrappers.Add wrp
End Sub

Private Function AddLabel(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.label
    RemoveCtl nm
    Dim c As MSForms.label
    Set c = mFrm.Controls.Add("Forms.Label.1", nm, True)
    c.Left = x: c.top = y: c.width = w: c.Height = h
    Set AddLabel = c
End Function

Private Function AddText(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                         ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    RemoveCtl nm
    Dim c As MSForms.TextBox
    Set c = mFrm.Controls.Add("Forms.TextBox.1", nm, True)
    c.Left = x: c.top = y: c.width = w: c.Height = h
    Set AddText = c
End Function

Private Function AddCombo(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.ComboBox
    RemoveCtl nm
    Dim c As MSForms.ComboBox
    Set c = mFrm.Controls.Add("Forms.ComboBox.1", nm, True)
    c.Left = x: c.top = y: c.width = w: c.Height = h
    Set AddCombo = c
End Function

Private Function AddButton(ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal h As Single) As MSForms.CommandButton
    RemoveCtl nm
    Dim c As MSForms.CommandButton
    Set c = mFrm.Controls.Add("Forms.CommandButton.1", nm, True)
    c.Left = x: c.top = y: c.width = w: c.Height = h
    Set AddButton = c
End Function

Private Sub RemoveCtl(ByVal nm As String)
    On Error Resume Next
    mFrm.Controls.Remove nm
    On Error GoTo 0
End Sub

' Postavi podrazumevanu vrstu/sortu (iz config-a) na prosledjene combo-e.
' Postavljanje cmbVrsta.Value okida _Change u formi (puni sortu + auto-cena/tip).
Public Sub ApplyDefaultProizvod(ByVal cmbVrsta As Object, ByVal cmbSorta As Object)
    On Error Resume Next
    Dim v As String, s As String
    v = Trim$(GetConfigValue(CFG_DEFAULT_VRSTA))
    s = Trim$(GetConfigValue(CFG_DEFAULT_SORTA))
    If Len(v) = 0 Then Exit Sub
    cmbVrsta.value = v
    If Len(s) > 0 Then cmbSorta.value = s
End Sub
