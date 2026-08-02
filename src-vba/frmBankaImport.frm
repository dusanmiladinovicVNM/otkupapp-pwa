VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBankaImport 
   Caption         =   "UserForm1"
   ClientHeight    =   13590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmBankaImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBankaImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_Data As Variant
Private m_BimIDs() As String
Private mChromeRemoved As Boolean

' Self-update bezbedno hvatanje evenata RUNTIME kontrola: nove WithEvents
' deklaracije NE smeju u formu (lome code-merge te forme pri update-u --
' docs/SELF_UPDATE.md zamka #11), pa event sink zivi u clsUiSink instancama.
' Isti obrazac kao frmDokumenta (WireSink + UiSinkEvent na dnu modula).
Private m_uiSinks As Object          ' tag -> clsUiSink

' Filter stavki izvoda (runtime kontrole; .frx se ne dira). Do sada je lstBanka
' prikazivala ISKLJUCIVO otvorene (nemapirane) stavke, pa uvezeni izvod nije
' imao gde da se vidi u celini. Combo + opseg datuma to otvaraju:
' Otvorene / Obradjene / Preskocene / Sve.
'
' TIP JE NAMERNO "As Object", NE "As MSForms.*": modSelfUpdate.IsHardModuleBody
' tretira module-level " AS MSFORMS." kao TVRDO telo i rutira celu formu na
' reinstall (SELF_UPDATE.md zamka #20). frmBankaImport je do sada bila "meka"
' (self-updatable) forma i mora to da ostane -- zato late-bound reference.
' NE menjati u MSForms.* tipove.
Private m_cmbBimFilter As Object          ' MSForms.ComboBox (late-bound)
Private m_txtBimOd As Object              ' MSForms.TextBox
Private m_txtBimDo As Object              ' MSForms.TextBox
Private m_lblBimFilter As Object          ' MSForms.Label
Private m_lblBimOd As Object              ' MSForms.Label
Private m_lblBimDo As Object              ' MSForms.Label
Private m_filterBuilt As Boolean

' Za koliko se lstBanka spusta da bi filter traka stala iznad nje.
Private Const BIM_FILTER_SHIFT As Single = 36

Private Sub UserForm_Activate()
    On Error GoTo EH
    
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    ' Headers
    StyleFrameTitleLabel lblKopf, "Bank Import"
    StyleSubtitle lblSubtitle, Poruka("BANKA_LBL_UVOZ_TRANSAKCIJA_BANKARSKIH")
    
    ' Section headers (ako koristis frames)
    ' StyleSectionHeader fraDetail, "Detalji selektovane stavke"
    ' StyleSectionHeader fraPreview, "Pregled automatskog mapiranja"
    
    ' Action buttons
    StylePrimaryButton btnAutoJedan, "Automatski mapiraj red"
    StylePrimaryButton btnAutoSve, "Automatski mapiraj sve"
    StylePrimaryButton btnSacuvajRucno, "Rucno mapiraj red"
    StylePrimaryButton btnSkip, "Preskoci red"
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StyleExitButton btnPovratak, "Zatvori"     ' ili "Povratak"
    
    ' Status labels
    StyleLabel lblStatus, TXT_MUTED(), True
    StyleLabel lblIzvodSummary, TXT_MUTED(), True
    
    ' Detail labels (suptilno)
    StyleLabel lblBimID, TXT_MUTED(), False
    StyleLabel lblPartner, TXT_MUTED(), False
    StyleLabel lblPozivNaBroj, TXT_MUTED(), False
    StyleLabel lblOpis, TXT_MUTED(), False
    StyleLabel lblSvrha, TXT_MUTED(), False
    StyleLabel lblIznos, TXT_MUTED(), True       ' bold jer je money
    StyleLabel lblPreview, TXT_MUTED(), False
    
    ' Initialize cmbMapTip
    If cmbMapTip.ListCount = 0 Then
        cmbMapTip.AddItem "Kupac"
        cmbMapTip.AddItem "Kooperant"
        cmbMapTip.AddItem "OM"
    End If
    
    SetupList
    SetupBankaFilter            ' pomera lstBanka nadole -> mora PRE BuildListHeaders
    BuildListHeaders

    ' Auto-map sve sto se moze preko jakih kljuceva (poziv->otkup/faktura, tekuci racun)
    ' pre prikaza; dvosmislene ostaju otvorene za rucno. Ne obara formu ako padne.
    On Error Resume Next
    AutoMapStrongKeysBankaImport_TX
    On Error GoTo EH

    LoadBankaRows
    
    ' KPI strip (opciono -- vidi Izmena 2)
     LayoutTopKpis
     RefreshTopKpis
    
    lstBanka.SetFocus
    
    Exit Sub
    
EH:
    LogErr "frmBankaImport.UserForm_Activate"
    MsgBox "Gre" & ChrW(353) & "ka pri otvaranju forme: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub SetupList()
    With lstBanka
        .ColumnCount = 7
        .ColumnWidths = "70;70;140;80;70;70;60"
    End With
End Sub

' Runtime kolone-headeri iznad lstBanka (listbox se puni AddItem-om, pa ColumnHeads
' ne radi bez RowSource-a). Idempotentno: ukloni pa dodaj (Activate moze vise puta).
' Sirine odgovaraju SetupList .ColumnWidths "70;70;140;80;70;70;60".
Private Sub BuildListHeaders()
    On Error Resume Next

    Dim titles As Variant
    Dim widths As Variant
    Dim i As Long
    Dim x As Single
    Dim lbl As MSForms.label
    Dim nm As String

    titles = Array("BIM", "Datum", "Partner", "Poziv na broj", "Uplata", "Isplata", "Status")
    widths = Array(70, 70, 140, 80, 70, 70, 60)

    x = lstBanka.Left
    For i = LBound(titles) To UBound(titles)
        nm = "hdrBanka_" & CStr(i)
        Me.Controls.Remove nm

        Set lbl = Me.Controls.Add("Forms.Label.1", nm, True)
        lbl.Left = x
        lbl.Top = lstBanka.Top - 13
        lbl.Width = CSng(widths(i))
        lbl.Height = 12
        lbl.caption = CStr(titles(i))
        lbl.Font.Bold = True
        lbl.Font.Size = 8

        x = x + CSng(widths(i))
    Next i
End Sub

' ============================================================
' FILTER STAVKI IZVODA -- runtime traka iznad lstBanka (Controls.Add + WireSink;
' .frx se ne dira). Prostor se uzima od same liste (spusti se za BIM_FILTER_SHIFT),
' pa ne moze da preklopi nijednu zatecenu kontrolu. Idempotentno (Activate moze
' vise puta). Podaci: modBankaMapiranje.GetBankaImportRows.
' ============================================================
Private Sub SetupBankaFilter()
    On Error GoTo done
    If m_filterBuilt Then Exit Sub

    Dim oldTop As Single
    oldTop = lstBanka.Top

    ' Oslobodi prostor iznad liste (lista se skrati odozgo). Guard se podize ODMAH
    ' po pomeranju: ako izgradnja kontrola posle toga padne, lista se ne sme pomeriti
    ' jos jednom pri sledecem Activate (BimFilterMode tada vraca default "Otvorene").
    lstBanka.Top = oldTop + BIM_FILTER_SHIFT
    lstBanka.Height = lstBanka.Height - BIM_FILTER_SHIFT
    If lstBanka.Height < 60 Then lstBanka.Height = 60
    m_filterBuilt = True

    Const LBLH As Single = 12
    Const ROWH As Single = 20
    Const GAP As Single = 8
    Dim lblY As Single: lblY = oldTop - 13      ' red gde su ranije stajali headeri
    Dim ctlY As Single: ctlY = oldTop + 1
    Dim x As Single: x = lstBanka.Left

    Set m_lblBimFilter = NewBimLabel("lblBimFilterRT", "Prikaz", x, lblY, 120, LBLH)
    Set m_cmbBimFilter = Me.Controls.Add("Forms.ComboBox.1", "cmbBimFilterRT", True)
    With m_cmbBimFilter
        .Style = fmStyleDropDownList
        .AddItem BIM_F_OTVORENE
        .AddItem BIM_F_OBRADJENE
        .AddItem BIM_F_PRESKOCENE
        .AddItem BIM_F_SVE
        .Move x, ctlY, 120, ROWH
        .value = BIM_F_OTVORENE          ' default = dosadasnje ponasanje
    End With
    ' Sink TEK posle postavljanja default vrednosti -- inace bi _Change okinuo
    ' LoadBankaRows jos pre nego sto je forma zavrsila Activate.
    WireSink m_cmbBimFilter, "m_cmbBimFilter"
    StyleComboBox m_cmbBimFilter
    x = x + 120 + GAP

    Set m_lblBimOd = NewBimLabel("lblBimOdRT", "Od", x, lblY, 80, LBLH)
    Set m_txtBimOd = Me.Controls.Add("Forms.TextBox.1", "txtBimOdRT", True)
    m_txtBimOd.Move x, ctlY, 80, ROWH
    StyleTextBox m_txtBimOd
    x = x + 80 + GAP

    Set m_lblBimDo = NewBimLabel("lblBimDoRT", "Do", x, lblY, 80, LBLH)
    Set m_txtBimDo = Me.Controls.Add("Forms.TextBox.1", "txtBimDoRT", True)
    m_txtBimDo.Move x, ctlY, 80, ROWH
    StyleTextBox m_txtBimDo

    Exit Sub
done:
    LogErr "frmBankaImport.SetupBankaFilter"
End Sub

Private Function NewBimLabel(ByVal ctlName As String, ByVal cap As String, _
                             ByVal l As Single, ByVal t As Single, _
                             ByVal w As Single, ByVal h As Single) As Object
    Dim lbl As Object
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_MUTED()
        .Font.name = APP_FONT: .Font.Size = 8
        .caption = cap
        .Move l, t, w, h
    End With
    Set NewBimLabel = lbl
End Function

' Izabrani filter (default = Otvorene, dok traka jos nije izgradjena).
Private Function BimFilterMode() As String
    On Error Resume Next
    BimFilterMode = BIM_F_OTVORENE
    If m_cmbBimFilter Is Nothing Then Exit Function
    Dim s As String: s = Trim$(CStr(m_cmbBimFilter.value))
    If Len(s) > 0 Then BimFilterMode = s
End Function

' Tekst polja -> Date; prazno / neparsivo = 0 (bez te granice).
Private Function BimDatum(ByVal t As Object) As Date
    On Error Resume Next
    If t Is Nothing Then Exit Function
    Dim s As String: s = Trim$(CStr(t.value))
    If Len(s) = 0 Then Exit Function
    Dim d As Date
    If TryParseDateValue(s, d) Then BimDatum = d
End Function

Private Sub m_cmbBimFilter_Change()
    LoadBankaRows
End Sub

Private Sub LoadBankaRows()
    Dim i As Long
    Dim colID As Long, colDatum As Long, colPartner As Long
    Dim colPoziv As Long, colUplata As Long, colIsplata As Long, colObr As Long

    Dim fMode As String
    fMode = BimFilterMode()

    lstBanka.Clear
    Erase m_BimIDs

    m_Data = GetBankaImportRows(fMode, BimDatum(m_txtBimOd), BimDatum(m_txtBimDo))
    If IsEmpty(m_Data) Then
        lblStatus.caption = "Nema stavki za izbor: " & fMode & "."
        UpdateIzvodSummaryLabel
        Exit Sub
    End If
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    colDatum = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE)
    colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
    colPoziv = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ)
    colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)
    
    ReDim m_BimIDs(0 To UBound(m_Data, 1) - 1)
    
    For i = 1 To UBound(m_Data, 1)
        lstBanka.AddItem CStr(m_Data(i, colID))
        lstBanka.List(lstBanka.ListCount - 1, 1) = Format$(m_Data(i, colDatum), "d.m.yyyy")
        lstBanka.List(lstBanka.ListCount - 1, 2) = CStr(m_Data(i, colPartner))
        lstBanka.List(lstBanka.ListCount - 1, 3) = CStr(m_Data(i, colPoziv))
        lstBanka.List(lstBanka.ListCount - 1, 4) = Format$(CDbl(nz(m_Data(i, colUplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 5) = Format$(CDbl(nz(m_Data(i, colIsplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 6) = CStr(nz(m_Data(i, colObr), ""))
        
        m_BimIDs(i - 1) = CStr(m_Data(i, colID))
    Next i
    
    lblStatus.caption = lstBanka.ListCount & " stavki  (" & fMode & _
                        BimPeriodSuffix() & ")"

    UpdateIzvodSummaryLabel
    RefreshTopKpis

End Sub

' ", period 1.1.2026 - 30.6.2026" ili "" kad opseg nije zadat.
Private Function BimPeriodSuffix() As String
    Dim dOd As Date, dDo As Date
    dOd = BimDatum(m_txtBimOd)
    dDo = BimDatum(m_txtBimDo)
    If dOd = 0 And dDo = 0 Then Exit Function

    Dim a As String, b As String
    If dOd > 0 Then a = Format$(dOd, "d.m.yyyy") Else a = "pocetak"
    If dDo > 0 Then b = Format$(dDo, "d.m.yyyy") Else b = "danas"
    BimPeriodSuffix = ", period " & a & " - " & b
End Function

Private Sub lstBanka_Click()
    If lstBanka.ListIndex < 0 Then Exit Sub
    ShowSelectedRow
    UpdateAutoPreview
End Sub

Private Sub ShowSelectedRow()
    Dim bimID As String
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    lblBimID.caption = bimID
    lblPartner.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_PARTNER))
    lblPozivNaBroj.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_POZIV_NA_BROJ))
    lblOpis.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_OPIS))
    lblSvrha.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_SVRHA_PLACANJA))
    
    Dim uplata As Double, isplata As Double
    uplata = CDbl(nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_UPLATA), "0"))
    isplata = CDbl(nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_ISPLATA), "0"))
    
    lblIznos.caption = "Uplata: " & Format$(uplata, "#,##0.00") & _
                       " | Isplata: " & Format$(isplata, "#,##0.00")
    
    If uplata > 0 Then
        cmbMapTip.value = "Kupac"
    ElseIf isplata > 0 Then
        cmbMapTip.value = "Kooperant"
    End If
    
    LoadManualTargets
End Sub

Private Sub cmbMapTip_Change()
    LoadManualTargets
    UpdateAutoPreview
End Sub

Private Sub LoadManualTargets()
    cmbPartner.Clear
    cmbFaktura.Clear
    
    Select Case cmbMapTip.value
        Case "Kupac"
            FillCmb cmbPartner, GetLookupList(TBL_KUPCI, "Naziv")
            
        Case "Kooperant"
            Dim data As Variant
            Dim i As Long
            Dim colID As Long, colIme As Long, colPrezime As Long
            
            data = GetTableData(TBL_KOOPERANTI)
            If IsEmpty(data) Then Exit Sub
            
            colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
            colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
            colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
            
            For i = 1 To UBound(data, 1)
                cmbPartner.AddItem CStr(data(i, colID)) & " - " & _
                                   CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime))
            Next i
            
        Case "OM"
            FillCmb cmbPartner, GetLookupList(TBL_STANICE, "Naziv")
    End Select
End Sub
Private Sub cmbPartner_Change()
    If cmbMapTip.value = "Kooperant" Then
        LoadOtkupBlokoviForSelectedKooperant
    End If
    UpdateAutoPreview
End Sub
Private Sub cmbOtkupBlok_Change()
    UpdateAutoPreview
End Sub
Private Sub LoadOtkupBlokoviForSelectedKooperant()
    Dim kooperantID As String
    Dim data As Variant
    Dim colKoop As Long, colBrDok As Long
    Dim dict As Object
    Dim i As Long
    
    cmbOtkupBlok.Clear
    
    If cmbPartner.value = "" Then Exit Sub
    
    kooperantID = ExtractIDFromDisplay(cmbPartner.value)
    
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID Then
            If Trim$(CStr(data(i, colBrDok))) <> "" Then
                If Not dict.Exists(CStr(data(i, colBrDok))) Then
                    dict.Add CStr(data(i, colBrDok)), True
                End If
            End If
        End If
    Next i
    
    Dim k As Variant
    For Each k In dict.keys
        cmbOtkupBlok.AddItem CStr(k)
    Next k
End Sub

Private Sub btnAutoJedan_Click()
    Dim bimID As String
    Dim result As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    result = AutoMapBankaImportRow_TX(bimID)
    
    If result <> "" Then
        MsgBox "Automatski mapirano.", vbInformation, APP_NAME
    End If
    
    LoadBankaRows
End Sub

Private Sub btnAutoSve_Click()
    Dim n As Long
    n = AutoMapAllBankaImport_TX()
    MsgBox "Automatski mapirano: " & n, vbInformation, APP_NAME
    LoadBankaRows
End Sub

Private Sub btnSacuvajRucno_Click()
    Dim bimID As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    Select Case cmbMapTip.value
        Case "Kupac"
            Dim kupacID As String
            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.value, "KupacID"))
            Call MapBankaImportAsKupac_TX(bimID, kupacID, "", True)
            
    Case "Kooperant"
        Dim kooperantID As String
        Dim brojBloka As String
        Dim n As Long
    
        kooperantID = ExtractIDFromDisplay(cmbPartner.value)
        brojBloka = Trim$(cmbOtkupBlok.value)
    
        If brojBloka <> "" Then
            n = MapBankaImportAsKooperantBlockManual_TX(bimID, kooperantID, brojBloka, True)
        Else
            n = MapBankaImportAsKooperantBlock_TX(bimID, kooperantID, True)
        End If
            
        Case "OM"
            Dim omID As String
            omID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbPartner.value, "StanicaID"))
            Call MapBankaImportAsOM_TX(bimID, omID, "", True)
    End Select
    
    LoadBankaRows
End Sub

Private Sub btnSkip_Click()
    Dim bimID As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    If SkipBankaImportRow_TX(bimID) Then
        LoadBankaRows
    End If
End Sub

Private Sub btnOsvezi_Click()
    LoadBankaRows
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmBankaImport.btnPovratak_Click"
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    ReleaseUiSinks                    ' raskini krug forma<->clsUiSink pre gasenja

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
End Sub

'PREVIEW
Private Sub UpdateAutoPreview()
    Dim bimID As String
    
    lblPreview.caption = ""
    
    If lstBanka.ListIndex < 0 Then Exit Sub
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    lblPreview.caption = BuildAutoPreviewText(bimID)
End Sub

Private Function BuildAutoPreviewText(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim uplata As Double
    Dim isplata As Double
    Dim mapped As Variant
    Dim s As String
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        BuildAutoPreviewText = "Preview nije dostupan."
        Exit Function
    End If
    
    partnerName = CStr(bim(1, 3))
    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))
    
    s = "BIM ID: " & bankaImportID & vbCrLf
    s = s & "Partner: " & partnerName & vbCrLf
    
    If Trim$(CStr(bim(1, 10))) <> "" Then
        s = s & "Poziv na broj: " & CStr(bim(1, 10)) & vbCrLf
    End If
    
    If uplata > 0 And isplata = 0 Then
        s = s & BuildIncomingPreview(bankaImportID, partnerName)
    ElseIf isplata > 0 And uplata = 0 Then
        s = s & BuildOutgoingPreview(bankaImportID, partnerName)
    Else
        s = s & "Status: Nije cist smer uplata/isplata"
    End If
    
    BuildAutoPreviewText = s
End Function

Private Function BuildIncomingPreview(ByVal bankaImportID As String, ByVal partnerName As String) As String
    Dim mapped As Variant
    Dim kupacID As String
    Dim fakturaID As String
    Dim s As String
    Dim konto As String
    Dim poziv As String
    Dim kupacByKonto As Variant
    Dim fakByPoziv As Variant
    Dim idDoc As String
    Dim idKonto As String
    Dim resolvedKupac As String
    Dim matchVia As String

    ' PRIORITET (kao AutoMapIncomingKupac): tekuci racun (kupac) + poziv na broj (faktura).
    konto = CStr(NzBIM(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_PARTNER_KONTO), ""))
    poziv = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))

    kupacByKonto = TryResolveKupacByKonto(konto)
    fakByPoziv = TryResolveFakturaByPoziv(poziv)

    fakturaID = ""
    idDoc = ""
    If Not IsEmpty(fakByPoziv) Then
        fakturaID = CStr(fakByPoziv(0))
        idDoc = CStr(fakByPoziv(1))
    End If

    If IsEmpty(kupacByKonto) Then
        idKonto = ""
    Else
        idKonto = CStr(kupacByKonto(0))
    End If

    If idDoc <> "" And idKonto <> "" And UCase$(idDoc) <> UCase$(idKonto) Then
        BuildIncomingPreview = "Smer: Uplata" & vbCrLf & _
            "Auto match: KONFLIKT (poziv->" & idDoc & " / racun->" & idKonto & ") -> rucno"
        Exit Function
    End If

    resolvedKupac = idDoc
    matchVia = "poziv na broj"
    If resolvedKupac = "" Then
        resolvedKupac = idKonto
        matchVia = "tekuci racun"
    End If

    If resolvedKupac <> "" Then
        kupacID = resolvedKupac
        If fakturaID = "" Then fakturaID = TryResolveFakturaForKupac(bankaImportID, kupacID)

        s = "Smer: Uplata" & vbCrLf
        s = s & "Auto match: Kupac (" & matchVia & ")" & vbCrLf
        s = s & "KupacID: " & kupacID & vbCrLf
        s = s & "Kupac: " & CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")) & vbCrLf

        If fakturaID <> "" Then
            s = s & "FakturaID: " & fakturaID & vbCrLf
            s = s & "Broj fakture: " & CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ)) & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
        Else
            s = s & "Faktura: nije jednoznacno nadjena" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
        End If

        BuildIncomingPreview = s
        Exit Function
    End If

    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        If CStr(mapped(1)) = "Kupac" Then
            kupacID = CStr(mapped(0))
            fakturaID = TryResolveFakturaForKupac(bankaImportID, kupacID)
            
            s = "Smer: Uplata" & vbCrLf
            s = s & "Auto match: Kupac" & vbCrLf
            s = s & "KupacID: " & kupacID & vbCrLf
            s = s & "Kupac: " & CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")) & vbCrLf
            
            If fakturaID <> "" Then
                s = s & "FakturaID: " & fakturaID & vbCrLf
                s = s & "Broj fakture: " & CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ)) & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
            Else
                s = s & "Faktura: nije jednoznacno nadjena" & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
            End If
            
            BuildIncomingPreview = s
            Exit Function
        End If
    End If
    
    mapped = TryResolveKupacBIM(partnerName)
    If Not IsEmpty(mapped) Then
        kupacID = CStr(mapped(0))
        fakturaID = TryResolveFakturaForKupac(bankaImportID, kupacID)
        
        s = "Smer: Uplata" & vbCrLf
        s = s & "Auto match: Kupac (heuristika)" & vbCrLf
        s = s & "KupacID: " & kupacID & vbCrLf
        s = s & "Kupac: " & CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")) & vbCrLf
        
        If fakturaID <> "" Then
            s = s & "FakturaID: " & fakturaID & vbCrLf
            s = s & "Broj fakture: " & CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ)) & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
        Else
            s = s & "Faktura: nije jednoznacno nadjena" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
        End If
        
        BuildIncomingPreview = s
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        s = "Smer: Uplata" & vbCrLf
        s = s & "Auto match: OM" & vbCrLf
        s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
        s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
        s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
        BuildIncomingPreview = s
        Exit Function
    End If
    
    BuildIncomingPreview = "Smer: Uplata" & vbCrLf & "Auto match: Nije pronadjen"
End Function

Private Function BuildOutgoingPreview(ByVal bankaImportID As String, ByVal partnerName As String) As String
    Dim mapped As Variant
    Dim kooperantID As String
    Dim kandidati As Variant
    Dim s As String
    Dim blockNo As String
    Dim i As Long
    Dim konto As String
    Dim poziv As String
    Dim koopByPoziv As String
    Dim koopByKonto As Variant
    Dim idDoc As String
    Dim idKonto As String
    Dim resolvedKoop As String
    Dim matchVia As String

    ' PRIORITET (kao AutoMapOutgoingKooperantOrOM): poziv na broj (otkup) + tekuci racun.
    konto = CStr(NzBIM(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_PARTNER_KONTO), ""))
    poziv = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))

    koopByPoziv = TryResolveKooperantByOtkupPoziv(poziv)
    koopByKonto = TryResolveKooperantByKonto(konto)

    idDoc = koopByPoziv
    If IsEmpty(koopByKonto) Then
        idKonto = ""
    Else
        idKonto = CStr(koopByKonto(0))
    End If

    If idDoc <> "" And idKonto <> "" And UCase$(idDoc) <> UCase$(idKonto) Then
        BuildOutgoingPreview = "Smer: Isplata" & vbCrLf & _
            "Auto match: KONFLIKT (poziv->" & idDoc & " / racun->" & idKonto & ") -> rucno"
        Exit Function
    End If

    resolvedKoop = idDoc
    matchVia = "poziv na broj"
    If resolvedKoop = "" Then
        resolvedKoop = idKonto
        matchVia = "tekuci racun"
    End If

    If resolvedKoop <> "" Then
        kooperantID = resolvedKoop
        blockNo = poziv
        kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)

        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (" & matchVia & ")" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf

        If Trim$(blockNo) <> "" Then
            s = s & "Blok: " & blockNo & vbCrLf
        End If

        If IsEmpty(kandidati) Then
            s = s & "Otkup kandidati: nema otvorenih stavki" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
        Else
            s = s & "Otkup kandidati:" & vbCrLf
            For i = 1 To UBound(kandidati, 1)
                s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
                    Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
                    CStr(kandidati(i, 3)) & vbCrLf
            Next i
            s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP
        End If

        BuildOutgoingPreview = s
        Exit Function
    End If

    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        Select Case CStr(mapped(1))
            Case "Kooperant"
                kooperantID = CStr(mapped(0))
                blockNo = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))
                kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)
                
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: Kooperant" & vbCrLf
                s = s & "KooperantID: " & kooperantID & vbCrLf
                s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf
                
                If Trim$(blockNo) <> "" Then
                    s = s & "Blok: " & blockNo & vbCrLf
                End If
                
                If IsEmpty(kandidati) Then
                    s = s & "Otkup kandidati: nema otvorenih stavki" & vbCrLf
                    s = s & "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
                Else
                    s = s & "Otkup kandidati:" & vbCrLf
                    For i = 1 To UBound(kandidati, 1)
                        s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
                            Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
                            CStr(kandidati(i, 3)) & vbCrLf
                    Next i
                    s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP
                End If
                
                BuildOutgoingPreview = s
                Exit Function
                
            Case "OM"
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: OM" & vbCrLf
                s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
                s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
                BuildOutgoingPreview = s
                Exit Function
        End Select
    End If
    
    mapped = TryResolveKooperantBIM(partnerName)
    If Not IsEmpty(mapped) Then
        kooperantID = CStr(mapped(0))
        blockNo = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))
        kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)
        
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (heuristika)" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf
        
        If Trim$(blockNo) <> "" Then
            s = s & "Blok: " & blockNo & vbCrLf
        End If
        
        If IsEmpty(kandidati) Then
            s = s & "Otkup kandidati: nema otvorenih stavki" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
        Else
            s = s & "Otkup kandidati:" & vbCrLf
            For i = 1 To UBound(kandidati, 1)
                s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
                    Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
                    CStr(kandidati(i, 3)) & vbCrLf
            Next i
            s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP
        End If
        
        BuildOutgoingPreview = s
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: OM" & vbCrLf
        s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
        s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
        s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
        BuildOutgoingPreview = s
        Exit Function
    End If
    
    BuildOutgoingPreview = "Smer: Isplata" & vbCrLf & "Auto match: Nije pronadjen"
End Function

'======================================================================
' v6.18+: Statement Summary Display
'
' Prikazuje saldo info iz najnovijeg aktivnog (non-stornirano) izvoda
' u tblBankaImport. Citaje BrojIzvoda + DatumIzvoda i 4 saldo polja,
' pa prikazuje math integrity status (Level 1) read-side kao vizualnu
' potvrdu operateru.
'
' Pozvana iz LoadBankaRows() na kraju, sto pokriva sve refresh tacke
' (Activate, btnOsvezi_Click, i posle svake Auto*/Sacuvaj/Skip akcije).
'======================================================================
Private Sub UpdateIzvodSummaryLabel()
    On Error GoTo EH
    
    Dim data As Variant
    Dim colBrojIzvoda As Long
    Dim colDatumIzvoda As Long
    Dim colPocetno As Long
    Dim colZavrsno As Long
    Dim colDuguje As Long
    Dim colPotrazuje As Long
    
    Dim i As Long
    Dim maxRow As Long
    Dim maxDatum As Date
    
    data = GetTableData(TBL_BANKA_IMPORT)
    
    If IsEmpty(data) Then
        Me.lblIzvodSummary.caption = "Nema importovanih izvoda."
        Exit Sub
    End If
    
    If UBound(data, 1) < 1 Then
        Me.lblIzvodSummary.caption = "Nema importovanih izvoda."
        Exit Sub
    End If
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    
    If IsEmpty(data) Then
        Me.lblIzvodSummary.caption = "Nema aktivnih izvoda."
        Exit Sub
    End If
    
    colBrojIzvoda = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    colDatumIzvoda = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_IZVODA)
    colPocetno = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POCETNO_STANJE)
    colZavrsno = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ZAVRSNO_STANJE)
    colDuguje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_DUGUJE)
    colPotrazuje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_POTRAZUJE)
    
    maxRow = -1
    maxDatum = #1/1/1900#
    
    For i = 1 To UBound(data, 1)
        Dim rowDatumStr As String
        rowDatumStr = Trim$(CStr(data(i, colDatumIzvoda)))
        
        If LenB(rowDatumStr) > 0 Then
            Dim rowDatum As Date
            On Error Resume Next
            rowDatum = CDate(rowDatumStr)
            If Err.Number = 0 Then
                On Error GoTo EH
                If rowDatum >= maxDatum Then
                    maxDatum = rowDatum
                    maxRow = i
                End If
            Else
                Err.Clear
                On Error GoTo EH
            End If
        End If
    Next i
    
    If maxRow < 0 Then
        Me.lblIzvodSummary.caption = "Nema aktivnih izvoda sa validnim datumom."
        Exit Sub
    End If
    
    Dim brIzv As String
    Dim datumIzv As String
    Dim pocetno As Double
    Dim zavrsno As Double
    Dim duguje As Double
    Dim potraz As Double
    
    brIzv = CStr(data(maxRow, colBrojIzvoda))
    datumIzv = CStr(data(maxRow, colDatumIzvoda))
    pocetno = CDbl(nz(data(maxRow, colPocetno), "0"))
    zavrsno = CDbl(nz(data(maxRow, colZavrsno), "0"))
    duguje = CDbl(nz(data(maxRow, colDuguje), "0"))
    potraz = CDbl(nz(data(maxRow, colPotrazuje), "0"))
    
    ' Level 1 read-side math check
    Dim expected As Double
    Dim diff As Double
    expected = pocetno + potraz - duguje
    diff = Abs(expected - zavrsno)
    
    Dim statusTxt As String
    Dim statusColor As Long
    
    If pocetno = 0 And zavrsno = 0 And duguje = 0 And potraz = 0 Then
        ' Legacy row (pre-v6.18 import, bez saldo metapodataka)
        statusTxt = "(legacy - bez saldo metapodataka)"
        statusColor = RGB(128, 128, 128)
    ElseIf diff <= 0.01 Then
        statusTxt = "OK"
        statusColor = RGB(0, 128, 0)
    Else
        statusTxt = "DIFF " & Format$(diff, "#,##0.00")
        statusColor = RGB(192, 0, 0)
    End If
    
    Me.lblIzvodSummary.caption = _
        "Izvod " & brIzv & " (" & datumIzv & ")  |  " & _
        "Pocetno: " & Format$(pocetno, "#,##0.00") & "  |  " & _
        "Uplate: " & Format$(potraz, "#,##0.00") & "  |  " & _
        "Isplate: " & Format$(duguje, "#,##0.00") & "  |  " & _
        "Zavrsno: " & Format$(zavrsno, "#,##0.00") & "  |  " & _
        statusTxt
    Me.lblIzvodSummary.ForeColor = statusColor
    
    Exit Sub

EH:
    On Error Resume Next
    Me.lblIzvodSummary.caption = "(gre" & ChrW(353) & "ka pri " & ChrW(269) & "itanju saldo info-a)"
    Me.lblIzvodSummary.ForeColor = RGB(128, 128, 128)
    
    LogErr "frmBankaImport.UpdateIzvodSummaryLabel"
End Sub

Private Sub LayoutTopKpis()
    On Error GoTo EH
    
    LayoutTopKpiInternals fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent
    LayoutTopKpiInternals fraKpiAutoMatch, lblKpiAutoTitle, lblKpiAutoValue, lblKpiAutoAccent
    LayoutTopKpiInternals fraKpiUplate, lblKpiUplTitle, lblKpiUplValue, lblKpiUplAccent
    LayoutTopKpiInternals fraKpiIsplate, lblKpiIspTitle, lblKpiIspValue, lblKpiIspAccent
    
    Exit Sub
EH:
    LogErr "frmBankaImport.LayoutTopKpis"
End Sub

Private Sub RefreshTopKpis()
    On Error GoTo EH
    
    Dim totalCount As Long
    Dim totalUplata As Double
    Dim totalIsplata As Double
    Dim mappedCount As Long
    Dim totalStaged As Long

    If IsArray(m_Data) Then
        totalCount = UBound(m_Data, 1)

        Dim colUplata As Long, colIsplata As Long
        colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
        colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)

        Dim i As Long
        For i = 1 To UBound(m_Data, 1)
            totalUplata = totalUplata + CDbl(nz(m_Data(i, colUplata), "0"))
            totalIsplata = totalIsplata + CDbl(nz(m_Data(i, colIsplata), "0"))
        Next i
    End If

    ' Stvarno stanje mapiranja iz cele tblBankaImport (ne samo otvorene):
    ' Mapirano = Obradjeno "Da"; Ukupno = sve nestornirane staging stavke.
    ComputeBankaMapState mappedCount, totalStaged
    
    ' Card 1: Otvoreno
    StyleTopKpi fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent, "neutral"
    lblKpiOtvTitle.caption = "Otvoreno"
    lblKpiOtvValue.caption = totalCount & " stavki"
    
    ' Card 2: Mapirano (stvarno stanje: Obradjeno "Da" / Ukupno staged)
    Dim autoKind As String
    If mappedCount > 0 Then autoKind = "ok" Else autoKind = "neutral"
    StyleTopKpi fraKpiAutoMatch, lblKpiAutoTitle, lblKpiAutoValue, lblKpiAutoAccent, autoKind
    lblKpiAutoTitle.caption = "Mapirano"
    lblKpiAutoValue.caption = mappedCount & " / " & totalStaged
    
    ' Card 3: Uplate ukupno
    StyleTopKpi fraKpiUplate, lblKpiUplTitle, lblKpiUplValue, lblKpiUplAccent, "neutral"
    lblKpiUplTitle.caption = "Uplate"
    lblKpiUplValue.caption = Format$(totalUplata, "#,##0") & " RSD"
    
    ' Card 4: Isplate ukupno
    StyleTopKpi fraKpiIsplate, lblKpiIspTitle, lblKpiIspValue, lblKpiIspAccent, "neutral"
    lblKpiIspTitle.caption = "Isplate"
    lblKpiIspValue.caption = Format$(totalIsplata, "#,##0") & " RSD"
    
    Exit Sub
EH:
    LogErr "frmBankaImport.RefreshTopKpis"
End Sub

' Realno stanje mapiranja iz cele tblBankaImport (ne samo otvorenih redova u m_Data).
Private Sub ComputeBankaMapState(ByRef mappedCount As Long, ByRef totalStaged As Long)
    On Error GoTo EH

    Dim data As Variant
    Dim colObr As Long
    Dim i As Long
    Dim st As String

    mappedCount = 0
    totalStaged = 0

    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Sub

    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)

    For i = 1 To UBound(data, 1)
        totalStaged = totalStaged + 1
        st = UCase$(Trim$(CStr(nz(data(i, colObr), ""))))
        If st = "DA" Then mappedCount = mappedCount + 1
    Next i

    Exit Sub
EH:
    LogErr "frmBankaImport.ComputeBankaMapState"
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnAutoJedan, "Automatski mapiraj red"
    StylePrimaryButton btnAutoSve, "Automatski mapiraj sve"
    StylePrimaryButton btnSacuvajRucno, "Rucno mapiraj red"
    StylePrimaryButton btnSkip, "Preskoci red"
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StyleExitButton btnPovratak, "Zatvori"
End Sub

Private Sub btnAutoJedan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnAutoJedan
End Sub

Private Sub btnAutoSve_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnAutoSve
End Sub

Private Sub btnSacuvajRucno_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnSacuvajRucno
End Sub

Private Sub btnSkip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnSkip
End Sub

Private Sub btnOsvezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnOsvezi
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

' ------------------------------------------------------------
' UI sink (clsUiSink) - eventi runtime kontrola bez WithEvents u formi
' (self-update bezbedno; docs/SELF_UPDATE.md zamka #11). Isti obrazac kao
' frmDokumenta: Bind -> UiSinkEvent dispatcher -> postojeci handleri.
' ------------------------------------------------------------

' Vrati True ako je sink stvarno vezan. Fail-visible (log) - tiho neuspesno
' vezivanje bi dalo vidljivu kontrolu koja ne reaguje.
Private Function WireSink(ByVal ctl As Object, ByVal tagName As String) As Boolean
    On Error GoTo Fail
    If ctl Is Nothing Then Err.Raise 91, , "kontrola je Nothing"
    If m_uiSinks Is Nothing Then Set m_uiSinks = CreateObject("Scripting.Dictionary")
    Dim s As clsUiSink
    Set s = New clsUiSink
    s.Bind Me, ctl, tagName
    Set m_uiSinks(tagName) = s
    WireSink = True
    Exit Function
Fail:
    LogErr "frmBankaImport.WireSink(" & tagName & ")", Err.description
End Function

' Otpusti sve clsUiSink omotace (raskini krug forma<->sink i reference kontrola).
' Idempotentno; pozvati iz QueryClose i Terminate.
Private Sub ReleaseUiSinks()
    On Error Resume Next
    Dim k As Variant
    If Not m_uiSinks Is Nothing Then
        For Each k In m_uiSinks.Keys
            m_uiSinks(k).ReleaseSink
        Next k
        m_uiSinks.RemoveAll
    End If
    Set m_uiSinks = Nothing
End Sub

' Dispatcher za clsUiSink (Public po nuznosti - klasa dobacuje event formi;
' ne zvati direktno).
Public Sub UiSinkEvent(ByVal tagName As String, ByVal ev As String, ByVal arg As Object)
    Select Case tagName & "." & ev
        Case "m_cmbBimFilter.Change":   m_cmbBimFilter_Change
    End Select
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    ReleaseUiSinks                    ' raskini krug forma<->clsUiSink
End Sub
