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

' Runtime kontrola (.frx se ne dira) + njen event sink. WithEvents ZIVI U clsUiSink,
' nikad u formi (self-update code-merge; docs/SELF_UPDATE.md zamka #11).
Private mBtnStrongMap As MSForms.CommandButton
Private m_uiSinks As Object                  ' tag -> clsUiSink
Private Const STRONG_MAP_CAPTION As String = "Mapiraj jake kljuceve"

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
    BuildListHeaders
    EnsureRuntimeControls

    ' AUD-014: otvaranje forme NE knjizi novac. Ranije je Activate pod
    ' "On Error Resume Next" pozivao AutoMapStrongKeysBankaImport_TX -> samo
    ' otvaranje pregleda je pravilo redove u tblNovac, bez potvrde, bez prikaza
    ' rezultata, a greska (i rollback celog pass-a) se gutala. Sada se pri
    ' otvaranju SAMO prebroji sta bi jaki kljucevi mapirali (read-only), a
    ' knjizenje ide na klik dugmeta "Mapiraj jake kljuceve (N)".
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

Private Sub LoadBankaRows()
    Dim i As Long
    Dim colID As Long, colDatum As Long, colPartner As Long
    Dim colPoziv As Long, colUplata As Long, colIsplata As Long, colObr As Long
    
    lstBanka.Clear
    Erase m_BimIDs
    
    m_Data = GetBankaImportOpen()
    If IsEmpty(m_Data) Then
        lblStatus.caption = "Nema otvorenih stavki."
        UpdateIzvodSummaryLabel
        UpdateStrongMapHint
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
    
    lblStatus.caption = lstBanka.ListCount & " otvorenih stavki"

    UpdateIzvodSummaryLabel
    UpdateStrongMapHint
    RefreshTopKpis

End Sub

'======================================================================
' Runtime kontrole + jaki kljucevi iza dugmeta (AUD-014)
'
' .frx se NE dira: dugme se pravi u runtime-u (Controls.Add), a event ide kroz
' clsUiSink (nikad "Private WithEvents" u formi -- zamka #11).
'======================================================================
Private Sub EnsureRuntimeControls()
    On Error GoTo EH

    If Not mBtnStrongMap Is Nothing Then Exit Sub

    Dim host As Object
    Set host = btnAutoSve.Parent

    ' Idempotentno (kao BuildListHeaders): ako je kontrola ostala iz ranijeg
    ' prikaza a referenca je otpustena, prvo je ukloni pa dodaj ponovo.
    On Error Resume Next
    host.Controls.Remove "btnStrongMap"
    On Error GoTo EH

    Set mBtnStrongMap = host.Controls.Add("Forms.CommandButton.1", "btnStrongMap", True)
    WireSink mBtnStrongMap, "mBtnStrongMap"

    With mBtnStrongMap
        .width = btnAutoSve.width
        .Height = btnAutoSve.Height
        .top = btnAutoSve.top
        .Left = ButtonRowRightEdge(host) + 8

        ' Ako nema mesta desno od postojeceg reda dugmadi, spusti ga ispod "Auto sve".
        If .Left + .width > host.InsideWidth Then
            .Left = btnAutoSve.Left
            .top = btnAutoSve.top + btnAutoSve.Height + 6
        End If

        .ControlTipText = "Knjizi samo stavke sa jednoznacnim jakim kljucem " & _
                          "(poziv na broj -> otkup/faktura, tekuci racun)"
    End With

    StylePrimaryButton mBtnStrongMap, STRONG_MAP_CAPTION
    Exit Sub

EH:
    LogErr "frmBankaImport.EnsureRuntimeControls"
End Sub

' Desna ivica postojeceg reda dugmadi u istom kontejneru.
Private Function ButtonRowRightEdge(ByVal host As Object) As Single
    Dim ctls As Variant
    Dim v As Variant
    Dim edge As Single

    edge = btnAutoSve.Left + btnAutoSve.width
    ctls = Array(btnAutoJedan, btnAutoSve, btnSacuvajRucno, btnSkip, btnOsvezi, btnPovratak)

    On Error Resume Next
    For Each v In ctls
        If v.Parent Is host Then
            If v.Left + v.width > edge Then edge = v.Left + v.width
        End If
    Next v
    On Error GoTo 0

    ButtonRowRightEdge = edge
End Function

' Koliko bi jaki kljucevi mapirali -- CISTO CITANJE, bez knjizenja.
Private Sub UpdateStrongMapHint()
    On Error GoTo EH

    Dim n As Long
    n = CountStrongKeyReadyBankaImport()

    lblStatus.caption = lblStatus.caption & "  |  jaki kljucevi: " & CStr(n) & " spremno"

    If Not mBtnStrongMap Is Nothing Then
        StylePrimaryButton mBtnStrongMap, STRONG_MAP_CAPTION & " (" & CStr(n) & ")"
        mBtnStrongMap.enabled = (n > 0)
    End If

    Exit Sub

EH:
    LogErr "frmBankaImport.UpdateStrongMapHint"
End Sub

Private Sub mBtnStrongMap_Click()
    On Error GoTo EH

    Dim n As Long
    Dim mapped As Long

    n = CountStrongKeyReadyBankaImport()

    If n = 0 Then
        MsgBox "Nema stavki koje jaki kljucevi (poziv na broj / tekuci racun) mogu " & _
               "jednoznacno da mapiraju.", vbInformation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Mapirati " & CStr(n) & " stavki po jakim kljucevima?" & vbCrLf & vbCrLf & _
              "Knjizi se u tblNovac. Dvosmislene stavke ostaju otvorene za rucno mapiranje.", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    mapped = AutoMapStrongKeysBankaImport_TX()

    LoadBankaRows

    MsgBox "Mapirano po jakim kljucevima: " & CStr(mapped) & _
           " (prepoznato: " & CStr(n) & ").", vbInformation, APP_NAME
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.mBtnStrongMap_Click"

    MsgBox "Mapiranje po jakim kljucevima nije proslo, promene su vracene: " & errDesc, _
           vbCritical, APP_NAME

    On Error Resume Next
    LoadBankaRows
End Sub

Private Sub mBtnStrongMap_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover mBtnStrongMap
End Sub

' ------------------------------------------------------------
' UI sink (clsUiSink) - eventi runtime kontrola bez WithEvents u formi
' ------------------------------------------------------------
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

Private Sub ReleaseUiSinks()
    On Error Resume Next
    Dim k As Variant
    If Not m_uiSinks Is Nothing Then
        For Each k In m_uiSinks.keys
            m_uiSinks(k).ReleaseSink
        Next k
        m_uiSinks.RemoveAll
    End If
    Set m_uiSinks = Nothing
    Set mBtnStrongMap = Nothing
End Sub

Public Sub UiSinkEvent(ByVal tagName As String, ByVal ev As String, ByVal arg As Object)
    Select Case tagName & "." & ev
        Case "mBtnStrongMap.Click":      mBtnStrongMap_Click
        Case "mBtnStrongMap.MouseMove":  mBtnStrongMap_MouseMove 0, 0, 0, 0
    End Select
End Sub

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
    ElseIf cmbMapTip.value = "Kupac" Then
        LoadFaktureForSelectedKupac
    End If
    UpdateAutoPreview
End Sub

Private Sub cmbFaktura_Change()
    UpdateAutoPreview
End Sub

' FM-0024 #2/#26: cmbFaktura je bila mrtva kontrola (samo Clear), pa je RUCNO
' mapiranje kupca UVEK slalo prazan FakturaID -> svaka rucna uplata je zavrsavala
' kao avans (NOV_KUPCI_AVANS), i onda kad je faktura postojala i bila vidljiva u
' preview-u. Lista nudi samo NEstornirane fakture tog kupca sa otvorenim saldom;
' prazan izbor = svesno avans (btnSacuvajRucno trazi potvrdu).
Private Sub LoadFaktureForSelectedKupac()
    On Error GoTo EH

    Dim kupacID As String
    Dim data As Variant
    Dim colFID As Long, colBroj As Long, colKup As Long, colIznos As Long
    Dim i As Long
    Dim otvoreno As Double

    cmbFaktura.Clear

    If Trim$(nz(cmbPartner.value, "")) = "" Then Exit Sub

    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.value, "KupacID"))
    If Trim$(kupacID) = "" Then Exit Sub

    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub

    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colIznos = GetColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS)

    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKup)) = kupacID Then
            otvoreno = CDbl(nz(data(i, colIznos), "0")) - _
                       GetUplataForFaktura(CStr(data(i, colFID)))

            If otvoreno > 0.009 Then
                cmbFaktura.AddItem CStr(data(i, colFID)) & " - " & _
                                   CStr(data(i, colBroj)) & " | otvoreno: " & _
                                   Format$(otvoreno, "#,##0.00")
            End If
        End If
    Next i

    Exit Sub

EH:
    LogErr "frmBankaImport.LoadFaktureForSelectedKupac"
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
    Dim manualRequired As Long
    Dim msg As String

    n = AutoMapAllBankaImport_TX(manualRequired)

    msg = "Automatski mapirano: " & n
    If manualRequired > 0 Then
        msg = msg & vbCrLf & "Za rucno mapiranje (nejednoznacno): " & manualRequired
    End If

    MsgBox msg, vbInformation, APP_NAME
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
            Dim fakturaID As String

            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.value, "KupacID"))
            If Trim$(kupacID) = "" Then
                MsgBox "Izaberite kupca.", vbExclamation, APP_NAME
                Exit Sub
            End If

            fakturaID = ExtractIDFromDisplay(Trim$(nz(cmbFaktura.value, "")))

            ' FM-0024 #2: bez izabrane fakture uplata NIJE zatvaranje duga nego
            ' avans. Ranije se to desavalo tiho (uvek prazan FakturaID).
            If fakturaID = "" Then
                If MsgBox("Nije izabrana faktura." & vbCrLf & vbCrLf & _
                          "Uplata se knjizi kao AVANS kupca (" & NOV_KUPCI_AVANS & "), " & _
                          "ne kao zatvaranje fakture." & vbCrLf & vbCrLf & _
                          "Nastaviti?", vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
            End If

            ReportManualResult MapBankaImportAsKupac_TX(bimID, kupacID, fakturaID, True) <> "", "Kupac"

        Case "Kooperant"
            Dim kooperantID As String
            Dim brojBloka As String
            Dim n As Long

            kooperantID = ExtractIDFromDisplay(cmbPartner.value)
            If Trim$(kooperantID) = "" Then
                MsgBox "Izaberite kooperanta.", vbExclamation, APP_NAME
                Exit Sub
            End If

            brojBloka = Trim$(cmbOtkupBlok.value)

            If brojBloka <> "" Then
                n = MapBankaImportAsKooperantBlockManual_TX(bimID, kooperantID, brojBloka, True)
            Else
                n = MapBankaImportAsKooperantBlock_TX(bimID, kooperantID, True)
            End If

            ReportManualResult (n > 0), "Kooperant"

        Case "OM"
            Dim omID As String
            omID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbPartner.value, "StanicaID"))
            If Trim$(omID) = "" Then
                MsgBox "Izaberite OM.", vbExclamation, APP_NAME
                Exit Sub
            End If

            ReportManualResult MapBankaImportAsOM_TX(bimID, omID, "", True) <> "", "OM"
    End Select

    LoadBankaRows
End Sub

' FM-0024 #11: rezultat rucnog mapiranja se vise ne ignorise (ranije "Call ...").
' Odbijeno mapiranje (npr. pogresan smer) je do sada izgledalo kao da se nista
' nije desilo -- lista se samo osvezi i stavka ostane otvorena bez objasnjenja.
Private Sub ReportManualResult(ByVal okFlag As Boolean, ByVal tipNaziv As String)
    If okFlag Then
        MsgBox "Rucno mapirano (" & tipNaziv & ").", vbInformation, APP_NAME
    Else
        MsgBox "Rucno mapiranje (" & tipNaziv & ") NIJE izvrseno." & vbCrLf & _
               "Proveri preview (sekcija RUCNO) i status stavke.", vbExclamation, APP_NAME
    End If
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

    ReleaseUiSinks
    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmBankaImport.btnPovratak_Click"
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    ' Raskini krug forma<->clsUiSink pre zatvaranja (self-update bezbedno).
    ReleaseUiSinks

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
    lblPreview.caption = BuildAutoPreviewText(bimID) & BuildManualPreviewText(bimID)
End Sub

' FM-0024 #3: izvor za PRIKAZ mora biti isti kao izvor za KOMANDU. Preview je
' racunao kandidate po pozivu na broj IZ IZVODA, dok btnSacuvajRucno knjizi po
' bloku izabranom u cmbOtkupBlok -- operater je gledao jedan blok, a knjizio drugi.
Private Function EffectiveBlockNo(ByVal bankaImportID As String) As String
    Dim manualBlok As String

    If Trim$(nz(cmbMapTip.value, "")) = "Kooperant" Then
        manualBlok = Trim$(nz(cmbOtkupBlok.value, ""))
    End If

    If manualBlok <> "" Then
        EffectiveBlockNo = manualBlok
    Else
        EffectiveBlockNo = Trim$(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, _
                                                  bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    End If
End Function

' Sta bi uradilo RUCNO mapiranje sa trenutnim izborom u combo-ima -- cita iz ISTIH
' kontrola iz kojih cita komanda (tip, partner, blok, faktura), pa preview i akcija
' ne mogu da se razidju. Ukljucuje i smer-kapiju: red pogresnog smera je odbijen
' (AUD-025) i to se vidi PRE klika, ne posle.
Private Function BuildManualPreviewText(ByVal bankaImportID As String) As String
    On Error GoTo EH

    Dim bim As Variant
    Dim s As String
    Dim tipMap As String
    Dim uplata As Double
    Dim isplata As Double

    tipMap = Trim$(nz(cmbMapTip.value, ""))
    If tipMap = "" Then Exit Function

    s = vbCrLf & "--- RUCNO (dugme 'Rucno mapiraj red') ---" & vbCrLf

    If Trim$(nz(cmbPartner.value, "")) = "" Then
        BuildManualPreviewText = s & "Nije izabran partner."
        Exit Function
    End If

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function

    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))

    s = s & "Tip: " & tipMap & vbCrLf

    Select Case tipMap
        Case "Kupac"
            If uplata <= 0 Then
                BuildManualPreviewText = s & "ODBIJENO: stavka je isplata, a tip Kupac trazi uplatu."
                Exit Function
            End If

            Dim kupacID As String
            Dim fakturaID As String

            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.value, "KupacID"))
            fakturaID = ExtractIDFromDisplay(Trim$(nz(cmbFaktura.value, "")))

            s = s & "KupacID: " & kupacID & vbCrLf

            If fakturaID <> "" Then
                s = s & "FakturaID: " & fakturaID & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
            Else
                s = s & "Faktura: nije izabrana -> knjizi se kao AVANS" & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
            End If

        Case "Kooperant"
            If isplata <= 0 Then
                BuildManualPreviewText = s & "ODBIJENO: stavka je uplata, a tip Kooperant trazi isplatu."
                Exit Function
            End If

            Dim kooperantID As String
            Dim blokNo As String

            kooperantID = ExtractIDFromDisplay(cmbPartner.value)
            blokNo = EffectiveBlockNo(bankaImportID)

            s = s & "KooperantID: " & kooperantID & vbCrLf
            If Trim$(nz(cmbOtkupBlok.value, "")) <> "" Then
                s = s & "Blok (rucni izbor): " & blokNo & vbCrLf
            ElseIf blokNo <> "" Then
                s = s & "Blok (poziv na broj): " & blokNo & vbCrLf
            End If

            s = s & CandidatesPreviewText(kooperantID, blokNo)

        Case "OM"
            Dim omID As String
            omID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbPartner.value, "StanicaID"))
            s = s & "OMID: " & omID & vbCrLf
            s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_OTKUPAC
    End Select

    BuildManualPreviewText = s
    Exit Function

EH:
    LogErr "frmBankaImport.BuildManualPreviewText"
    BuildManualPreviewText = vbCrLf & "--- RUCNO ---" & vbCrLf & "(preview nije dostupan)"
End Function

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
        s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_OTKUPAC
        BuildIncomingPreview = s
        Exit Function
    End If
    
    BuildIncomingPreview = "Smer: Uplata" & vbCrLf & "Auto match: Nije pronadjen"
End Function

' GetOtkupCandidatesForKooperantBlock od RF-09 moze da digne gresku (3+ otvorenih
' stavki u bloku = rucno mapiranje). Preview to mora da PRIKAZE, ne da obori.
Private Function SafeBlockCandidates(ByVal kooperantID As String, _
                                     ByVal blokNo As String, _
                                     ByRef errText As String) As Variant
    errText = ""

    On Error Resume Next
    SafeBlockCandidates = GetOtkupCandidatesForKooperantBlock(kooperantID, blokNo)
    If Err.Number <> 0 Then
        errText = Err.description
        SafeBlockCandidates = Empty
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Jedan tekst kandidata bloka za SVE preview grane (auto: jak kljuc / PartnerMap /
' heuristika, i rucno) - ranije je isti blok bio prepisan cetiri puta.
Private Function CandidatesPreviewText(ByVal kooperantID As String, _
                                       ByVal blokNo As String) As String
    Dim kandidati As Variant
    Dim kandErr As String
    Dim s As String
    Dim i As Long

    kandidati = SafeBlockCandidates(kooperantID, blokNo, kandErr)

    If kandErr <> "" Then
        CandidatesPreviewText = "Otkup kandidati: " & kandErr & vbCrLf & _
                                "Tip knjizenja: RUCNO (automatska raspodela odbijena)"
        Exit Function
    End If

    If IsEmpty(kandidati) Then
        CandidatesPreviewText = "Otkup kandidati: nema otvorenih stavki" & vbCrLf & _
                                "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
        Exit Function
    End If

    s = "Otkup kandidati:" & vbCrLf
    For i = 1 To UBound(kandidati, 1)
        s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
            Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
            CStr(kandidati(i, 3)) & vbCrLf
    Next i
    s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP

    CandidatesPreviewText = s
End Function

Private Function BuildOutgoingPreview(ByVal bankaImportID As String, ByVal partnerName As String) As String
    Dim mapped As Variant
    Dim kooperantID As String
    Dim s As String
    Dim blockNo As String
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
        blockNo = EffectiveBlockNo(bankaImportID)

        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (" & matchVia & ")" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf

        If Trim$(blockNo) <> "" Then
            s = s & "Blok: " & blockNo & vbCrLf
        End If

        s = s & CandidatesPreviewText(kooperantID, blockNo)

        BuildOutgoingPreview = s
        Exit Function
    End If

    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        Select Case CStr(mapped(1))
            Case "Kooperant"
                kooperantID = CStr(mapped(0))
                blockNo = EffectiveBlockNo(bankaImportID)

                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: Kooperant" & vbCrLf
                s = s & "KooperantID: " & kooperantID & vbCrLf
                s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf

                If Trim$(blockNo) <> "" Then
                    s = s & "Blok: " & blockNo & vbCrLf
                End If

                s = s & CandidatesPreviewText(kooperantID, blockNo)

                BuildOutgoingPreview = s
                Exit Function
                
            Case "OM"
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: OM" & vbCrLf
                s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
                s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
                s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_OTKUPAC
                BuildOutgoingPreview = s
                Exit Function
        End Select
    End If
    
    mapped = TryResolveKooperantBIM(partnerName)
    If Not IsEmpty(mapped) Then
        kooperantID = CStr(mapped(0))
        blockNo = EffectiveBlockNo(bankaImportID)

        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (heuristika)" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf

        If Trim$(blockNo) <> "" Then
            s = s & "Blok: " & blockNo & vbCrLf
        End If

        s = s & CandidatesPreviewText(kooperantID, blockNo)

        BuildOutgoingPreview = s
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: OM" & vbCrLf
        s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
        s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
        s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_OTKUPAC
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

    ' Runtime dugme zadrzava svoj caption (sadrzi broj spremnih stavki).
    If Not mBtnStrongMap Is Nothing Then
        StylePrimaryButton mBtnStrongMap, mBtnStrongMap.caption
    End If
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
