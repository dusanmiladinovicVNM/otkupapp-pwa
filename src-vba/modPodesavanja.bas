Attribute VB_Name = "modPodesavanja"
'Attribute VB_Name = "modPodesavanja"
Option Explicit

' ============================================================
' modPodesavanja -- editor podesavanja (tblSEFConfig) u UI-u.
'
' Cilj: kad je tblSEFConfig sakriven (VeryHidden), operativna podesavanja se
' uredjuju kroz formu umesto rucnog editovanja celija. Otvara se kao nova
' sekcija "Podesavanja" u meniju Maticni podaci:
'   modMaticniLookups.MaticniSekcije -> frmMaticniPodaci.OpenSekcija ->
'   frmStammdaten (Tag = "Podesavanja") -> UserForm_Activate -> BuildConfigEditor.
'
' Kontrole se grade u RUNTIME-u (Controls.Add) -- frmStammdaten.frx se NE dira,
' isti obrazac kao modMaticniLookups/clsLookupMenuBtn i modOtkupBlok/clsBlokUI.
' Klik Sacuvaj/Sakrij/Povratak hvata clsConfigBtn (WithEvents).
'
' BEZBEDNOST: prikazuju se SAMO operativna (slobodan-unos) polja. Interni kes /
' anti-tamper kljucevi (LICENSE_TOKEN, LICENSE_BOUND_PARTS, LICENSE_NEXT_CHECK,
' LICENSE_STATUS, LICENSE_HWM, TRIAL_HWM, GOOGLE_*_TOKEN, APP_SETUP_*...) se
' NAMERNO NE prikazuju -- njihovo rucno menjanje je upravo bypass koji sakrivanje
' tabele zatvara.
'
' Izlaz u nuzdi (ako forma nije dostupna): Alt+F8 -> ShowConfigSheet.
' ============================================================

Private mFrm As Object              ' frmStammdaten instanca (host)
Private mInputs As Collection       ' input kontrole, key = ConfigKey
Private mWrappers As Collection     ' clsConfigBtn (drzi WithEvents zivim)
Private mBtnToggle As MSForms.CommandButton
Private mGroupOrder As Collection       ' imena grupa po redosledu
Private mRowsByGroup As Collection      ' key=grupa -> Collection of Array(lbl, input, typ)
Private mGroupHeaders As Collection     ' key=grupa -> header CommandButton (Tag "1"=collapsed)
Private mFormW As Single
Private mMargin As Single
Private mColCount As Long
Private mCellW As Single
Private mCol0X As Single
Private mCol1X As Single
Private mTopContentY As Single

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

    ' Redosled grupa: operativno gore -> tehnika dole.
    ' (Prodavac, Otkup, Malina rezim, Management/Klijent [+licenca/probni],
    '  SEF; pa Monitoring, Sinhronizacija, Google, Alati, Napredno.)

    CfgAdd c, "Prodavac (firma)", "SELLER_NAME", "Naziv firme", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_PIB", "PIB", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_MATICNI_BROJ", "Mati" & ChrW(269) & "ni broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_STREET", "Ulica i broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_CITY", "Grad", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_POSTAL_CODE", "Po" & ChrW(353) & "tanski broj", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_COUNTRY_CODE", "Dr" & ChrW(382) & "ava (kod, npr. RS)", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_ACCOUNT", "Teku" & ChrW(263) & "i ra" & ChrW(269) & "un", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_EMAIL", "Email", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_OBJEKAT_MESTO", "Objekat - lokacija/mesto", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_OBJEKAT_BR_REGISTRA", "Objekat - broj registra", "text"
    CfgAdd c, "Prodavac (firma)", "SELLER_LOGO_PATH", "Putanja do logotipa (prazno = <workbook>\logo.png)", "text"

    CfgAdd c, "Otkup / dokumenta", "OTKUP_KLAUZULA", "Klauzula (otkupni list)", "memo"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_ROK_ISPLATE", "Rok isplate (otkupni list)", "text"
    CfgAdd c, "Otkup / dokumenta", "KOOP_FILTER_BY_OM", "Filtriraj kooperante po otkupnom mestu", "bool"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_VRSTA_VOCA", "Podrazumevana vrsta vo" & ChrW(263) & "a", "list:" & LookupCSV(TBL_KULTURE, "VrstaVoca", True)
    CfgAdd c, "Otkup / dokumenta", "KOOP_AUTO_CREATE", "Auto-kreiraj kooperanta iz unetog imena", "bool"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_SORTA_VOCA", "Podrazumevana sorta vo" & ChrW(263) & "a", "list:" & LookupCSV(TBL_KULTURE, "SortaVoca", True)
    CfgAdd c, "Otkup / dokumenta", "AUTO_BROJ_DOKUMENTA", "Automatsko generisanje brojeva dokumenata", "bool"
    CfgAdd c, "Otkup / dokumenta", "DEFAULT_TIP_PALETE", "Podrazumevani tip palete", "text"
    CfgAdd c, "Otkup / dokumenta", "KES_ISPLATE", "Postoje ke" & ChrW(353) & " isplate proizvodjacima", "bool"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_BRUTO_UNOS", Poruka("CFG_MSG_KUPAC_UNOSI_BRUTO"), "bool"
    CfgAdd c, "Otkup / dokumenta", "PALETIRANJE", "Paletiranje (izrada paletnih listova)", "bool"
    CfgAdd c, "Otkup / dokumenta", "PDV_NADOKNADA_STOPA", "PDV nadoknada stopa (%)", "int"
    CfgAdd c, "Otkup / dokumenta", "OTKUP_BLOK_PANEL", "Panel za blokove (Otkup)", "bool"
    CfgAdd c, "Otkup / dokumenta", "PRACENJE_PARCELA", "Pracenje parcela (unos parcele u otkupu)", "bool"
    CfgAdd c, "Otkup / dokumenta", "VALIDACIJA_UNOSA", "Kompletna validacija unosa (obavezna polja pre snimanja)", "bool"

    ' --- Stampa (centralni izlazni dispecer; DocResolveMode cita ove kljuceve) ---
    ' Po dokumentu: PDF | PRINT | PREVIEW | OFF (prazno -> default tog dokumenta).
    CfgAdd c, ChrW(352) & "tampa", "OTKUP_PRINT_MODE", Poruka("CFG_MSG_STAMPA_OTKUPNOG_LISTA"), "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "GRUPNI_OTKUP_PRINT_MODE", "Grupni otkupni list", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "PRIJEMNICA_PRINT_MODE", Poruka("CFG_MSG_STAMPA_PRIJEMNICE_AUTO"), "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "OTPREMNICA_PRINT_MODE", "Otpremnica", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "PALETA_PRINT_MODE", Poruka("CFG_MSG_STAMPA_PALETNOG_LISTA"), "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "PRERADA_PRINT_MODE", "Paletni list got. proizvoda", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "OM_IZDAVANJE_PRINT_MODE", "Revers ambala" & ChrW(382) & "e (izdavanje/povrat)", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "FAKTURA_PRINT_MODE", "Faktura", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "KARTICA_PRINT_MODE", "Kartica kooperanta", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "KARTICA_AMB_PRINT_MODE", "Kartica ambala" & ChrW(382) & "e", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "SLEDLJIVOST_PRINT_MODE", "Sledljivost", "list:PDF;PRINT;PREVIEW;OFF"
    CfgAdd c, ChrW(352) & "tampa", "SPECIFIKACIJA_PRINT_MODE", "Specifikacija", "list:PDF;PRINT;PREVIEW;OFF"

    CfgAdd c, Poruka("CFG_MSG_MALINA_REZIM"), "MALINA_MODE", "Auto-zbirna iz otpremnice (1 stanica = 1 vozilo)", "bool"
    CfgAdd c, Poruka("CFG_MSG_MALINA_REZIM"), "AUTO_PRIJEMNICA_HLADNJACA", "Auto otpremnica+zbirna+prijemnica (OM=hladnjaca)", "bool"
    CfgAdd c, Poruka("CFG_MSG_MALINA_REZIM"), "MALINA_DEFAULT_KUPAC", "Podrazumevani kupac/hladnjaca (KupacID)", "list:" & LookupCSV(TBL_KUPCI, COL_KUP_ID, False)

    CfgAdd c, "Management / Klijent", "MGMT_USER_1", "Management korisnik 1", "text"
    CfgAdd c, "Management / Klijent", "MGMT_USER_2", "Management korisnik 2", "text"
    CfgAdd c, "Management / Klijent", "MGMT_USER_3", "Management korisnik 3", "text"
    CfgAdd c, "Management / Klijent", "CLIENT_ID", "Client ID", "text"
    CfgAdd c, "Management / Klijent", "CLIENT_NAME", "Client naziv", "text"
    CfgAdd c, "Management / Klijent", "ENV", "Okruzenje (klijent, DEV/PROD)", "text"
    CfgAdd c, "Management / Klijent", "LICENSE_ENABLED", "Licenciranje ukljuceno", "bool"
    CfgAdd c, "Management / Klijent", "LICENSE_KEY", "Licencni kljuc", "text"
    CfgAdd c, "Management / Klijent", "LICENSE_ENDPOINT", "Licencni endpoint (URL)", "text"
    CfgAdd c, "Management / Klijent", "TRIAL_ENABLED", "Probni period ukljucen", "bool"
    CfgAdd c, "Management / Klijent", "TRIAL_START", "Pocetak (yyyy-mm-dd)", "text"
    CfgAdd c, "Management / Klijent", "TRIAL_DAYS", "Trajanje (dana)", "int"

    CfgAdd c, "SEF", "SEF_BASE_URL", "SEF Base URL", "text"
    CfgAdd c, "SEF", "SEF_API_KEY", "SEF API kljuc", "secret"
    CfgAdd c, "SEF", "SEF_ENV", "SEF okruzenje", "text"
    CfgAdd c, "SEF", "SEF_PAYMENT_DUE_DAYS", "Rok placanja (dana)", "int"
    CfgAdd c, "SEF", "SEF_PAYMENT_MEANS_CODE", "Sifra nacina placanja", "text"
    CfgAdd c, "SEF", "SEF_NOTE_DEFAULT", "Podrazumevana napomena", "text"
    CfgAdd c, "SEF", "SEF_FORCE_TODAY_ISSUE_DATE", "Forsiraj dana" & ChrW(353) & "nji datum izdavanja", "bool"

    CfgAdd c, "Monitoring", "MONITORING_ENDPOINT", "Monitoring endpoint (URL)", "text"
    CfgAdd c, "Monitoring", "MONITORING_SECRET", "Monitoring secret", "secret"
    CfgAdd c, "Monitoring", "MONITORING_ENV", "Okruzenje (DEV/PROD)", "text"

    CfgAdd c, "Sinhronizacija", "CLOUD_SYNC_ENABLED", "Cloud sync ukljucen", "bool"
    CfgAdd c, "Sinhronizacija", "SHEETS_SYNC_ENABLED", "Google Sheets sync ukljucen", "bool"
    CfgAdd c, "Sinhronizacija", "SYNC_AUTO_INTERVAL_MIN", "Auto-sync interval (min, >=15)", "int"

    CfgAdd c, "Google", "GOOGLE_CLIENT_ID", "Google Client ID", "text"
    CfgAdd c, "Google", "GOOGLE_CLIENT_SECRET", "Google Client Secret", "secret"
    CfgAdd c, "Google", "GOOGLE_PWA_FOLDER_ID", "PWA Folder ID", "text"
    CfgAdd c, "Google", "GOOGLE_STAMMDATEN_SHEET_ID", "Stammdaten Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_KARTICE_SHEET_ID", "Kartice Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_MGMT_SHEET_ID", "Management Sheet ID", "text"
    CfgAdd c, "Google", "GOOGLE_REPORTS_FOLDER_ID", "Reports Folder ID", "text"

    ' Banka / lokalno -- per-masina putanje i parametri (tblLocalConfig, store="local").
    ' Citaju se preko GetLocalConfigValue u modBankaImport / modBankaImportParserPdfToText,
    ' zato OVDE moraju ici u "local" (inace se upisu u tblSEFConfig a ne procitaju).
    CfgAdd c, "Banka / lokalno", "PDFTOTEXT_EXE_PATH", "pdftotext.exe (banka import)", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_DRIVE_SOURCE_PATH", "Drive izvor (sinhronizovan 00_Inbox\01_Bank folder)", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_DRIVE_DOWNLOADED_PATH", "Drive obradjeno (opciono; prazno = ne premesta)", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_INBOX_PATH", "Lokalni inbox (staging)", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_PROCESSED_PATH", "Lokalni processed", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_ERROR_PATH", "Lokalni error", "text", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_DRIVE_MAX_FILES", "Max fajlova po povlacenju (Drive)", "int", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_DRIVE_MIN_FILE_AGE_SECONDS", "Min. starost fajla (s) pre povlacenja", "int", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_AUTO_IMPORT_ON_START", "Auto-uvoz izvoda pri startu", "list:DA;NE", "local"
    CfgAdd c, "Banka / lokalno", "BANKA_ALLOWED_EXTENSIONS", "Dozvoljene ekstenzije (npr. pdf)", "text", "local"

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
                   ByVal lbl As String, ByVal typ As String, _
                   Optional ByVal store As String = "sef")
    ' store: "sef" -> tblSEFConfig (Get/SetConfigValue); "local" -> tblLocalConfig
    ' (Get/SetLocalConfigValue), za per-masina putanje (Poppler, banka lokalne putanje).
    c.Add Array(grp, key, lbl, typ, store)
End Sub

' Helper: distinct vrednosti kolone kao "A;B;C" za dinamicki "list:" tip.
' Reuse GetLookupList; prazno ako nema podataka.
Private Function LookupCSV(ByVal tbl As String, ByVal col As String, _
                           ByVal onlyActive As Boolean) As String
    On Error Resume Next
    Dim arr As Variant
    arr = GetLookupList(tbl, col, "", , onlyActive)
    If IsArray(arr) Then
        If (UBound(arr) - LBound(arr) + 1) > 0 Then LookupCSV = Join(arr, ";")
    End If
End Function

' ============================================================
' PUBLIC -- izgradnja editora (poziva frmStammdaten.UserForm_Activate za Tag)
' ============================================================
Public Sub BuildConfigEditor(ByVal frm As Object)
    Const SRC As String = "modPodesavanja.BuildConfigEditor"
    On Error GoTo EH

    Set mFrm = frm
    Set mInputs = New Collection
    Set mWrappers = New Collection
    Set mBtnToggle = Nothing

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
    Const LBLW As Single = 250
    Dim inLeft As Single: inLeft = m + LBLW + 10
    Dim inW As Single: inW = w - inLeft - m - 18      ' rezerva za scrollbar
    If inW < 120 Then inW = 120

    ' Naslov
    Dim lblTitle As MSForms.label
    Set lblTitle = AddLabel("cfg_title", m, 8, w - 2 * m, 20)
    lblTitle.caption = Poruka("CFG_LBL_PODESAVANJA_TBLSEFCONFIG")
    StyleLabel lblTitle, TXT_LIGHT(), True
    lblTitle.Font.Size = FONT_SIZE_HEADER

    ' Footer dugmad (na vrhu -- vidljiva pre skrolovanja)
    Dim btnSave As MSForms.CommandButton
    Set btnSave = AddButton("btnCfgSave", m, 32, 120, 24)
    StylePrimaryButton btnSave, "Sacuvaj"
    WireButton btnSave, "save"

    Set mBtnToggle = AddButton("btnCfgToggle", m + 130, 32, 200, 24)
    StyleExitButton mBtnToggle, ToggleCaption()
    WireButton mBtnToggle, "toggle"

    Dim btnBack As MSForms.CommandButton
    Set btnBack = AddButton("btnCfgBack", w - m - 120, 32, 120, 24)
    StyleExitButton btnBack, "Povratak"
    WireButton btnBack, "back"

    Dim lblHint As MSForms.label
    Set lblHint = AddLabel("cfg_hint", m, 60, w - 2 * m, 16)
    lblHint.caption = "Interna polja (token, bound, status, HWM, OAuth token...) se namerno NE prikazuju. Grupa 'Banka / lokalno' se cuva u tblLocalConfig (per-masina)."
    StyleLabel lblHint, TXT_MUTED(), False
    lblHint.Font.Size = FONT_SIZE_SMALL

    ' Alatna traka: rasiri / skupi sve grupe odjednom.
    Dim btnExpand As MSForms.CommandButton
    Set btnExpand = AddButton("btnCfgExpandAll", m, 80, 150, 22)
    StyleExitButton btnExpand, "[+] Rasiri sve"
    WireButton btnExpand, "expandall"

    Dim btnCollapse As MSForms.CommandButton
    Set btnCollapse = AddButton("btnCfgCollapseAll", m + 158, 80, 150, 22)
    StyleExitButton btnCollapse, "[-] Skupi sve"
    WireButton btnCollapse, "collapseall"

    ' Direktan izbor Poppler-a (pdftotext.exe) bez Alt+F8 -- otvara folder picker
    ' (modSetup.SetupPopplerInteractive) i osvezi polje PDFTOTEXT_EXE_PATH u formi.
    Dim btnPoppler As MSForms.CommandButton
    Set btnPoppler = AddButton("btnCfgPoppler", m + 316, 80, 210, 22)
    StyleExitButton btnPoppler, "Izaberi Poppler (pdftotext.exe)"
    WireButton btnPoppler, "poppler"

    ' Geometrija kolona: 2 kolone kad ima dovoljno sirine, inace 1.
    Const COLGAP As Single = 16
    mFormW = w
    mMargin = m
    mColCount = IIf(w >= 720, 2, 1)
    mCellW = (w - 2 * m - (mColCount - 1) * COLGAP) / mColCount
    mCol0X = m
    mCol1X = m + mCellW + COLGAP
    mTopContentY = 112

    ' Izgradnja grupa (data-driven iz ConfigEditorFields; kolona "Grupa").
    ' Svaka grupa = collapsible sekcija (header dugme) sa svojim poljima.
    Set mGroupOrder = New Collection
    Set mRowsByGroup = New Collection
    Set mGroupHeaders = New Collection

    Dim flds As Variant: flds = ConfigEditorFields()
    Dim i As Long, f As Variant
    Dim grp As String, key As String, cap As String, typRaw As String, typ As String
    Dim store As String
    Dim cur As String
    Dim lbl As MSForms.label, tb As MSForms.TextBox, cmb As MSForms.ComboBox
    Dim opts As Variant, oi As Long
    Dim inCtl As MSForms.Control
    Dim rowsColl As Collection
    Dim hdrBtn As MSForms.CommandButton
    Dim curGroup As String: curGroup = ""
    Dim grpIdx As Long: grpIdx = 0

    For i = LBound(flds) To UBound(flds)
        f = flds(i)
        grp = CStr(f(0)): key = CStr(f(1)): cap = CStr(f(2))
        typRaw = CStr(f(3)): typ = LCase$(typRaw)
        store = "sef": If UBound(f) >= 4 Then store = LCase$(CStr(f(4)))

        If grp <> curGroup Then
            curGroup = grp
            grpIdx = grpIdx + 1
            mGroupOrder.Add grp
            Set rowsColl = New Collection
            mRowsByGroup.Add rowsColl, grp
            Set hdrBtn = AddGroupHeaderBtn(grp, grpIdx)
            mGroupHeaders.Add hdrBtn, grp
        End If

        Set lbl = AddLabel("cfglbl_" & key, 0, 0, mCellW, 14)
        lbl.caption = cap
        StyleLabel lbl, TXT_MUTED(), False
        lbl.Font.Size = FONT_SIZE_SMALL

        If store = "local" Then
            cur = GetLocalConfigValue(key, "")
        Else
            cur = GetConfigValue(key)
        End If

        If typ = "bool" Or Left$(typ, 5) = "list:" Then
            Set cmb = AddCombo("cfg_" & key, 0, 0, mCellW, 18)
            cmb.style = fmStyleDropDownCombo
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
            Set inCtl = cmb
        Else
            Set tb = AddText("cfg_" & key, 0, 0, mCellW, IIf(typ = "memo", 46, 18))
            If typ = "memo" Then
                tb.MultiLine = True
                tb.WordWrap = True
                tb.ScrollBars = fmScrollBarsVertical
            End If
            tb.value = cur
            StyleTextBox tb
            mInputs.Add tb, key
            Set inCtl = tb
        End If

        mRowsByGroup(grp).Add Array(lbl, inCtl, typ)
    Next i

    ' Pocetno stanje: sve grupe skupljene (pregledan "sadrzaj"); siri po potrebi.
    Dim gname As Variant
    For Each gname In mGroupOrder
        mGroupHeaders(CStr(gname)).Tag = "1"
    Next gname

    RelayoutGroups

    Exit Sub
EH:
    LogErr SRC
    MsgBox Poruka("CFG_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' PUBLIC -- click ruter (zove clsConfigBtn)
' ============================================================
Public Sub ConfigEditor_OnClick(ByVal action As String, Optional ByVal groupKey As String = "")
    On Error GoTo EH
    Select Case LCase$(action)
        Case "save": SaveConfigEditor
        Case "toggle": ToggleConfigSheet
        Case "back": CloseConfigEditor
        Case "grp": ConfigEditor_ToggleGroup groupKey
        Case "expandall": ConfigEditor_SetAll False
        Case "collapseall": ConfigEditor_SetAll True
        Case "poppler": ConfigEditor_PickPoppler
    End Select
    Exit Sub
EH:
    LogErr "modPodesavanja.ConfigEditor_OnClick"
End Sub

' Toggle collapse/expand jedne grupe (header.Tag "1"=collapsed).
Public Sub ConfigEditor_ToggleGroup(ByVal grp As String)
    On Error GoTo EH
    If mGroupHeaders Is Nothing Then Exit Sub
    Dim hdr As MSForms.CommandButton
    On Error Resume Next
    Set hdr = mGroupHeaders(grp)
    On Error GoTo EH
    If hdr Is Nothing Then Exit Sub
    hdr.Tag = IIf(hdr.Tag = "1", "0", "1")
    RelayoutGroups
    Exit Sub
EH:
    LogErr "modPodesavanja.ConfigEditor_ToggleGroup"
End Sub

' Skupi (collapseAll=True) ili rasiri (False) sve grupe odjednom.
Public Sub ConfigEditor_SetAll(ByVal collapseAll As Boolean)
    On Error GoTo EH
    If mGroupOrder Is Nothing Then Exit Sub
    Dim g As Variant
    For Each g In mGroupOrder
        mGroupHeaders(CStr(g)).Tag = IIf(collapseAll, "1", "0")
    Next g
    RelayoutGroups
    Exit Sub
EH:
    LogErr "modPodesavanja.ConfigEditor_SetAll"
End Sub

' Dugme "Izaberi Poppler" -- direktan izbor pdftotext.exe iz Podesavanja (bez Alt+F8).
' Pokrece folder picker + upis (modSetup.SetupPopplerInteractive), pa osvezi prikaz
' polja PDFTOTEXT_EXE_PATH da odmah pokaze novu vrednost (prazno = auto pored xlsm-a).
Public Sub ConfigEditor_PickPoppler()
    On Error GoTo EH
    modSetup.SetupPopplerInteractive
    On Error Resume Next
    If Not mInputs Is Nothing Then
        mInputs("PDFTOTEXT_EXE_PATH").value = GetLocalConfigValue("PDFTOTEXT_EXE_PATH", "")
    End If
    Exit Sub
EH:
    LogErr "modPodesavanja.ConfigEditor_PickPoppler"
End Sub

' Preracunaj pozicije svih grupa/polja (collapse-aware, 2-kolonski grid).
' Memo polja zauzimaju ceo red; ostala se ravnaju u kolone radi gustine.
Private Sub RelayoutGroups()
    On Error GoTo EH
    If mGroupOrder Is Nothing Then Exit Sub

    Const HDRH As Single = 20
    Const HDRGAP As Single = 4
    Const LBLH As Single = 12
    Const LBLGAP As Single = 2
    Const INH As Single = 18
    Const ROWGAP As Single = 8
    Const MEMOH As Single = 46
    Const GROUPGAP As Single = 10

    Dim cellH As Single: cellH = LBLH + LBLGAP + INH + ROWGAP
    Dim memoCellH As Single: memoCellH = LBLH + LBLGAP + MEMOH + ROWGAP

    Dim Y As Single: Y = mTopContentY
    Dim gname As Variant, grp As String
    Dim hdr As MSForms.CommandButton
    Dim collapsed As Boolean
    Dim grpRows As Collection, r As Variant
    Dim lbl As MSForms.label, inCtl As MSForms.Control, typ As String
    Dim col As Long, cx As Single, idx As Long

    For Each gname In mGroupOrder
        grp = CStr(gname)
        Set hdr = mGroupHeaders(grp)
        hdr.Left = mMargin
        hdr.width = mFormW - 2 * mMargin
        hdr.top = Y
        hdr.Visible = True
        collapsed = (hdr.Tag = "1")
        hdr.caption = IIf(collapsed, "[+]  ", "[-]  ") & grp
        Y = Y + HDRH + HDRGAP

        Set grpRows = mRowsByGroup(grp)
        col = 0
        For idx = 1 To grpRows.count
            r = grpRows(idx)
            Set lbl = r(0)
            Set inCtl = r(1)
            typ = CStr(r(2))

            If collapsed Then
                lbl.Visible = False
                inCtl.Visible = False
            ElseIf typ = "memo" Then
                If col = 1 Then Y = Y + cellH: col = 0
                lbl.Left = mMargin: lbl.top = Y: lbl.width = mFormW - 2 * mMargin
                lbl.Visible = True
                inCtl.Left = mMargin: inCtl.top = Y + LBLH + LBLGAP
                inCtl.width = mFormW - 2 * mMargin: inCtl.Height = MEMOH
                inCtl.Visible = True
                Y = Y + memoCellH
                col = 0
            Else
                cx = IIf(col = 0, mCol0X, mCol1X)
                lbl.Left = cx: lbl.top = Y: lbl.width = mCellW
                lbl.Visible = True
                inCtl.Left = cx: inCtl.top = Y + LBLH + LBLGAP
                inCtl.width = mCellW
                inCtl.Visible = True
                If mColCount = 1 Then
                    Y = Y + cellH
                ElseIf col = 1 Then
                    Y = Y + cellH: col = 0
                Else
                    col = 1
                End If
            End If
        Next idx

        If (Not collapsed) And col = 1 Then Y = Y + cellH
        If Not collapsed Then Y = Y + GROUPGAP
    Next gname

    On Error Resume Next
    mFrm.ScrollBars = fmScrollBarsVertical
    mFrm.ScrollHeight = Y + 16
    mFrm.KeepScrollBarsVisible = fmScrollBarsVertical
    On Error GoTo 0
    Exit Sub
EH:
    LogErr "modPodesavanja.RelayoutGroups"
End Sub

' Header dugme jedne grupe (collapsible). Klik hvata clsConfigBtn (action="grp").
Private Function AddGroupHeaderBtn(ByVal grp As String, ByVal idx As Long) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = AddButton("cfggrp_" & CStr(idx), mMargin, 0, mFormW - 2 * mMargin, 20)
    StyleMenuButton b, grp
    WireGroupButton b, grp
    Set AddGroupHeaderBtn = b
End Function

Private Sub WireGroupButton(ByVal b As MSForms.CommandButton, ByVal grp As String)
    Dim wrp As clsConfigBtn
    Set wrp = New clsConfigBtn
    wrp.action = "grp"
    wrp.groupKey = grp
    Set wrp.btn = b
    mWrappers.Add wrp
End Sub

' ============================================================
' PRIVATE -- save / back
' ============================================================
Private Sub SaveConfigEditor()
    Const SRC As String = "modPodesavanja.SaveConfigEditor"
    On Error GoTo EH
    If mInputs Is Nothing Then Exit Sub

    Dim flds As Variant: flds = ConfigEditorFields()
    Dim errs As String, n As Long
    Dim f As Variant, key As String, typ As String, v As String
    Dim store As String
    Dim i As Long

    For i = LBound(flds) To UBound(flds)
        f = flds(i)
        key = CStr(f(1))
        typ = LCase$(CStr(f(3)))
        store = "sef": If UBound(f) >= 4 Then store = LCase$(CStr(f(4)))

        v = ""
        On Error Resume Next
        v = Trim$(CStr(mInputs(key).value))
        On Error GoTo EH

        If typ = "int" And Len(v) > 0 And Not IsNumeric(v) Then
            errs = errs & " - " & key & " mora biti broj." & vbCrLf
        Else
            If store = "local" Then
                SetLocalConfigValue key, v
            Else
                SetConfigValue key, v
            End If
            n = n + 1
        End If
    Next i

    If Len(errs) > 0 Then
        MsgBox "Sacuvano: " & n & " polja." & vbCrLf & vbCrLf & _
               Poruka("CFG_MSG_PRESKOCENO_GRESKA") & vbCrLf & errs, vbExclamation, APP_NAME
    Else
        MsgBox "Sacuvano: " & n & " polja.", vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr SRC
    MsgBox Poruka("CFG_ERR_GRESKA_PRI_CUVANJU") & Err.description, vbCritical, APP_NAME
End Sub

Private Sub CloseConfigEditor()
    On Error Resume Next
    frmOtkupAPP.ReturnToDashboard Poruka("CFG_MSG_PODESAVANJA_ZATVORENA")
    Unload mFrm
    Set mFrm = Nothing
    Set mInputs = Nothing
    Set mWrappers = Nothing
    Set mBtnToggle = Nothing
    Set mGroupOrder = Nothing
    Set mRowsByGroup = Nothing
    Set mGroupHeaders = Nothing
End Sub

' ============================================================
' PUBLIC -- vidljivost tblSEFConfig sheet-a (toggle + Alt+F8 makroi)
' ============================================================
Public Sub ToggleConfigSheet()
    On Error GoTo EH
    If ConfigSheetIsHidden() Then
        ShowConfigSheet
        MsgBox "tblSEFConfig je sada VIDLJIV.", vbInformation, APP_NAME
    Else
        HideConfigSheet
        MsgBox "tblSEFConfig je sada SAKRIVEN (VeryHidden)." & vbCrLf & _
               "Ureduj ga iskljucivo preko ove forme." & vbCrLf & _
               Poruka("CFG_MSG_IZLAZ_NUZDI_ALT"), vbInformation, APP_NAME
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
    ToggleCaption = IIf(ConfigSheetIsHidden(), Poruka("CFG_MSG_PRIKAZI_CONFIG_TABELU"), "Sakrij config tabelu")
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
    Set ConfigSheet = lo.parent
End Function

' ============================================================
' PRIVATE -- runtime control helperi (Controls.Add; .frx se ne dira)
' ============================================================
Private Sub WireButton(ByVal b As MSForms.CommandButton, ByVal act As String)
    Dim wrp As clsConfigBtn
    Set wrp = New clsConfigBtn
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

Private Function AddText(ByVal nm As String, ByVal X As Single, ByVal Y As Single, _
                         ByVal w As Single, ByVal h As Single) As MSForms.TextBox
    RemoveCtl nm
    Dim c As MSForms.TextBox
    Set c = mFrm.Controls.Add("Forms.TextBox.1", nm, True)
    c.Left = X: c.top = Y: c.width = w: c.Height = h
    Set AddText = c
End Function

Private Function AddCombo(ByVal nm As String, ByVal X As Single, ByVal Y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.ComboBox
    RemoveCtl nm
    Dim c As MSForms.ComboBox
    Set c = mFrm.Controls.Add("Forms.ComboBox.1", nm, True)
    c.Left = X: c.top = Y: c.width = w: c.Height = h
    Set AddCombo = c
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


' Otpusti module-level reference (clsConfigBtn WithEvents + kontrole) pre
' self-update importa (zivi event-sink obara CodeModule edit). Idempotentno.
Public Sub Podesavanja_Release()
    On Error Resume Next
    Set mFrm = Nothing
    Set mInputs = Nothing
    Set mWrappers = Nothing
    Set mBtnToggle = Nothing
    Set mGroupOrder = Nothing
    Set mRowsByGroup = Nothing
    Set mGroupHeaders = Nothing
End Sub
