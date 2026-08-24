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

' Da li je lista faktura za izabranog kupca zaista ucitana. Prazan combo posle
' PADA ucitavanja izgleda isto kao "kupac nema otvorenih faktura", a prazan izbor
' znaci "knjizi kao avans" -- operater bi potvrdjivao odluku na osnovu netacne
' liste. Zato se pad pamti i knjizenje kupca se blokira dok se ne osvezi.
Private m_FaktureLoadOk As Boolean
Private m_FaktureLoadErr As String
' ISTO ZA BLOKOVE. Prazna lista blokova znaci "uzmi poziv na broj iz izvoda", a
' odatle prazan skup kandidata zavrsi kao AVANS kooperanta uz stavku oznacenu
' obradjenom. Pad ucitavanja zato ne sme da izgleda kao prazna lista -- ista
' klasa zbog koje m_FaktureLoadOk vec postoji.
Private m_BlokoviLoadOk As Boolean
Private m_BlokoviLoadErr As String

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
    ' Opis se hvata PRE LogErr-a: LogErr ide kroz On Error Resume Next i resetuje
    ' Err, pa bi poruka operateru ostala bez uzroka (AUD-054 obrazac).
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.UserForm_Activate"
    MsgBox "Gre" & ChrW(353) & "ka pri otvaranju forme: " & errDesc, vbCritical, APP_NAME
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
    ' Brojac vise ne guta gresku, pa je ovde i prikazujemo: tiho "0 spremno" bi
    ' schema/read problem prikazalo kao "nema sta da se mapira" (AUD-014).
    ' Opis se hvata PRE LogErr-a (BUG-1/AUD-054 idiom).
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.UpdateStrongMapHint"

    On Error Resume Next

    lblStatus.caption = lblStatus.caption & "  |  jaki kljucevi: GRESKA (" & errDesc & ")"

    If Not mBtnStrongMap Is Nothing Then
        StylePrimaryButton mBtnStrongMap, STRONG_MAP_CAPTION & " (?)"
        mBtnStrongMap.enabled = False
    End If
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
            ' FM-0024 #7: identitet je ID, ne naziv. `GetLookupList` spaja kupce
            ' istog naziva u JEDNU stavku, a `LookupValue` po nazivu vraca prvi
            ' pogodak -- uplata (ili avans) je mogla da zavrsi na pogresnom kupcu
            ' bez ijednog znaka operateru. Isti obrazac kao frmFakturisanje.
            FillComboDisplayID cmbPartner, TBL_KUPCI, COL_KUP_NAZIV, COL_KUP_ID
            ShowIDInComboDisplay cmbPartner    ' dva kupca istog naziva moraju da se razlikuju

        Case "Kooperant"
            Dim data As Variant
            Dim i As Long
            Dim colID As Long, colIme As Long, colPrezime As Long

            ' Kooperant lista je jednokolonska ("ID - Ime Prezime"); vrati
            ' ColumnCount posle Kupac/OM punjenja (koje je bound 2-kolonsko).
            cmbPartner.ColumnCount = 1

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
            ' Isto kao za kupce: stanice istog naziva se ne smeju stopiti u jednu
            ' stavku, a izbor mora da nosi StanicaID.
            FillComboDisplayID cmbPartner, TBL_STANICE, "Naziv", "StanicaID"
            ShowIDInComboDisplay cmbPartner    ' isto i za istoimene stanice
    End Select
End Sub

' Stabilan ID izabranog partnera -- iz bound kolone combo-a (Kupac/OM) ili iz
' "ID - Ime Prezime" prikaza (Kooperant). Jedno mesto za sve tri grane, da se
' preview i komanda ne bi razlikovali u tome KOGA su izabrali.
Private Function SelectedPartnerID() As String
    Select Case Trim$(nz(cmbMapTip.value, ""))
        Case "Kooperant"
            SelectedPartnerID = ExtractIDFromDisplay(nz(cmbPartner.value, ""))
        Case Else
            SelectedPartnerID = GetComboID(cmbPartner)
    End Select
End Function
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
    Dim colFID As Long, colBroj As Long, colKup As Long
    Dim i As Long
    Dim otvoreno As Double

    cmbFaktura.Clear
    m_FaktureLoadOk = True
    m_FaktureLoadErr = ""

    If Trim$(nz(cmbPartner.value, "")) = "" Then Exit Sub

    kupacID = SelectedPartnerID()
    If Trim$(kupacID) = "" Then Exit Sub

    ' Nedostajuca tabela NIJE "nema faktura": GetTableData vraca Empty za oba, a
    ' prazan izbor fakture znaci AVANS. Zastavica ispod bi ostala True.
    RequireTable TBL_FAKTURE, "frmBankaImport.LoadFaktureForSelectedKupac"

    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub

    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)

    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKup)) = kupacID Then
            ' Isti obracun otvorenog iznosa koji koristi i writer pri raspodeli
            ' uplate (GetOtvorenoNaFakturi) -- prikaz i knjizenje jedan izvor.
            otvoreno = GetOtvorenoNaFakturi(CStr(data(i, colFID)))

            If otvoreno > 0.009 Then
                cmbFaktura.AddItem CStr(data(i, colFID)) & " - " & _
                                   CStr(data(i, colBroj)) & " | otvoreno: " & _
                                   Format$(otvoreno, "#,##0.00")
            End If
        End If
    Next i

    Exit Sub

EH:
    ' Pad ucitavanja NE sme da izgleda kao "nema otvorenih faktura" (= avans).
    m_FaktureLoadOk = False
    m_FaktureLoadErr = Err.description

    LogErr "frmBankaImport.LoadFaktureForSelectedKupac"

    On Error Resume Next
    cmbFaktura.Clear
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
    m_BlokoviLoadOk = True
    m_BlokoviLoadErr = ""
    
    On Error GoTo EH
    
    If cmbPartner.value = "" Then Exit Sub
    
    kooperantID = SelectedPartnerID()
    
    ' Nedostajuca tabela NIJE "kooperant nema blokova": prazna lista vodi na
    ' poziv na broj, a odatle prazan skup kandidata zavrsi kao AVANS.
    RequireTable TBL_OTKUP, "frmBankaImport.LoadOtkupBlokoviForSelectedKooperant"
    
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

    Exit Sub

EH:
    ' Pad ucitavanja NE sme da izgleda kao "kooperant nema blokova". Prazna lista
    ' vodi na poziv na broj, a odatle prazan skup kandidata zavrsi kao AVANS uz
    ' stavku oznacenu obradjenom -- ista klasa zbog koje postoji m_FaktureLoadOk.
    m_BlokoviLoadOk = False
    m_BlokoviLoadErr = Err.description

    LogErr "frmBankaImport.LoadOtkupBlokoviForSelectedKooperant"

    On Error Resume Next
    cmbOtkupBlok.Clear
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
    On Error GoTo EH

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
    Exit Sub

EH:
    ' Batch koji padne se rollback-uje i sada PROPAGIRA gresku; bez ove grane bi
    ' korisnik posle critical poruke dobio jos i "Automatski mapirano: 0", sto
    ' izgleda kao uredno zavrsen batch bez pogodaka.
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.btnAutoSve_Click"

    MsgBox "Automatsko mapiranje NIJE izvrseno, promene su vracene: " & errDesc, _
           vbCritical, APP_NAME

    On Error Resume Next
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

            kupacID = SelectedPartnerID()
            If Trim$(kupacID) = "" Then
                MsgBox "Izaberite kupca.", vbExclamation, APP_NAME
                Exit Sub
            End If

            fakturaID = ExtractIDFromDisplay(Trim$(nz(cmbFaktura.value, "")))

            ' Ako lista faktura nije ucitana, prazan izbor NE znaci "nema fakture"
            ' nego "ne znamo" -- knjizenje avansa na osnovu takve liste je pogadjanje.
            If Not m_FaktureLoadOk Then
                MsgBox "Lista faktura za ovog kupca nije ucitana:" & vbCrLf & _
                       m_FaktureLoadErr & vbCrLf & vbCrLf & _
                       "Mapiranje kupca je zaustavljeno (prazna lista bi znacila avans). " & _
                       "Osvezi listu ili resi gresku pa pokusaj ponovo.", _
                       vbCritical, APP_NAME
                Exit Sub
            End If

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

            kooperantID = SelectedPartnerID()
            If Trim$(kooperantID) = "" Then
                MsgBox "Izaberite kooperanta.", vbExclamation, APP_NAME
                Exit Sub
            End If

            ' AKO LISTA BLOKOVA NIJE UCITANA, prazan combo NE znaci "operater
            ' nije birao blok". Fallback na poziv na broj bi tada bio pogadjanje,
            ' a ako iz njega ne ispadne nijedna otkupna stavka, ceo iznos se
            ' knjizi kao AVANS kooperanta i stavka se oznaci obradjenom.
            ' Isto pravilo koje ova forma vec ima za fakture (m_FaktureLoadOk).
            Dim rucnoPoruka As String
            If Not KooperantRucnoSme(rucnoPoruka) Then
                MsgBox rucnoPoruka, vbExclamation, APP_NAME
                Exit Sub
            End If

            ' ISTI blok koji pokazuje RUCNI preview: izbor iz liste ako postoji,
            ' inace poziv na broj iz izvoda. Ranije je prazan combo isao u zasebnu
            ' granu (`MapBankaImportAsKooperantBlock_TX`, bez potvrde podele), pa
            ' je blok sa 3+ otvorenih stavki tu zavrsavao generickom greskom --
            ' iako preview (koji vec racuna efektivni blok) nudi podelu. Kako je
            ' lista blokova samo napunjena a NIJE auto-selektovana, prazan combo
            ' je bio DEFAULT slucaj, tj. "3+ blok ima izlaz" je radilo samo ako
            ' operater rucno klikne blok.
            brojBloka = EffectiveManualBlockNo(bimID)

            ' Blok sa vise od MAX_BLOK_KANDIDATA otvorenih stavki automatski put
            ' ODBIJA (ne pogadja raspodelu). Rucni put ga zavrsava, ali tek kad
            ' operater vidi tacnu podelu i potvrdi je.
            Dim potvrdjenaPodela As Boolean

            If Not ConfirmManyCandidatesSplit(bimID, kooperantID, brojBloka, potvrdjenaPodela) Then Exit Sub

            ' Poslednji argument: blok je IZABRAN iz liste, ne izveden iz poziva
            ' na broj. Bez njega bi izabran a vec placen blok tiho postao avans
            ' kooperanta, a stavka bila oznacena obradjenom.
            ' Poslednji argument vraca da li je writer VEC prijavio gresku --
            ' inace operater za jednu ocekivanu validacionu situaciju dobija DVA
            ' dijaloga: konkretan iz writera i genericki odavde.
            Dim vecPrijavljeno As Boolean
            n = MapBankaImportAsKooperantBlockManual_TX(bimID, kooperantID, brojBloka, _
                                                       True, potvrdjenaPodela, "", _
                                                       ManualBlokIzabran(), vecPrijavljeno)

            If n > 0 Or Not vecPrijavljeno Then ReportManualResult (n > 0), "Kooperant"

        Case "OM"
            Dim omID As String
            omID = SelectedPartnerID()
            If Trim$(omID) = "" Then
                MsgBox "Izaberite OM.", vbExclamation, APP_NAME
                Exit Sub
            End If

            ReportManualResult MapBankaImportAsOM_TX(bimID, omID, "", True) <> "", "OM"
    End Select

    LoadBankaRows
End Sub

' Blok sa 3+ otvorenih stavki: prikazi TACNU podelu (isti planer koji knjizi) i
' trazi izricitu potvrdu. Vraca False = operater je odustao; `potvrdjeno` = True
' znaci da smemo da predjemo granicu automatske raspodele.
Private Function ConfirmManyCandidatesSplit(ByVal bankaImportID As String, _
                                            ByVal kooperantID As String, _
                                            ByVal brojBloka As String, _
                                            ByRef potvrdjeno As Boolean) As Boolean
    On Error GoTo EH

    Dim kandidati As Variant
    Dim manualRequiredText As String

    potvrdjeno = False
    ConfirmManyCandidatesSplit = True

    ' Bez prelaska granice: ako ovo prodje, blok je u granicama i nista se ne pita.
    ' Prazan tekst = blok je U GRANICAMA (nema sta da se potvrdjuje). Bilo koja
    ' DRUGA greska se iz SafeBlockCandidates propagira i hvata je EH ispod --
    ' ne tumaci se kao "previse kandidata".
    kandidati = SafeBlockCandidates(kooperantID, brojBloka, manualRequiredText)
    If manualRequiredText = "" Then Exit Function

    ' Preko granice -> ucitaj PUNU listu i predlozi podelu.
    kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, brojBloka, True)
    If IsEmpty(kandidati) Then Exit Function

    Dim isplata As Double
    isplata = CDbl(nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_ISPLATA), "0"))

    ' Tri ishoda, jer je podela ovde HEURISTIKA (veci otvoreni prvi), a komentar
    ' uz granicu kaze da 3+ stavki moze znaciti recikliran broj bloka ili dupliran
    ' unos -- greedy raspodela tada moze da plati pogresan otkup. Zato operater
    ' bira: DA = knjizi predlozenu podelu; NE = ceo iznos kao AVANS kooperanta
    ' (lineage ostaje otvoren, avans se kasnije precizno vezuje dugmetom "Primeni
    ' avans na blok" u Banka izvestaju); OTKAZI = ne diraj stavku.
    Dim odgovor As VbMsgBoxResult

    odgovor = MsgBox("Blok " & brojBloka & " ima " & CStr(UBound(kandidati, 1)) & _
              " otvorene otkupne stavke -- vise nego sto automatska raspodela sme da podeli," & _
              vbCrLf & "pa predlozena podela moze da pogodi POGRESAN otkup " & _
              "(recikliran broj bloka, dupliran unos)." & vbCrLf & vbCrLf & _
              "Predlog podele iznosa " & Format$(isplata, "#,##0.00") & ":" & vbCrLf & _
              SplitPreviewText(kandidati, isplata) & vbCrLf & _
              "DA = knjizi ovu podelu" & vbCrLf & _
              "NE = knjizi ceo iznos kao AVANS kooperanta (vezes ga kasnije rucno)" & vbCrLf & _
              "OTKAZI = ne diraj stavku", _
              vbQuestion + vbYesNoCancel, APP_NAME)

    If odgovor = vbCancel Then
        ConfirmManyCandidatesSplit = False
        Exit Function
    End If

    If odgovor = vbNo Then
        ' Bezbedan izlaz: nista se ne vezuje za otkup dok je poreklo dvosmisleno.
        ReportManualResult MapBankaImportAsKooperant_TX(bankaImportID, kooperantID, "", "", True) <> "", _
                           "Kooperant (avans)"
        LoadBankaRows
        ConfirmManyCandidatesSplit = False
        Exit Function
    End If

    potvrdjeno = True
    Exit Function

EH:
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.ConfirmManyCandidatesSplit"

    MsgBox "Ne mogu da pripremim podelu po bloku: " & errDesc, vbCritical, APP_NAME
    ConfirmManyCandidatesSplit = False
End Function

' Tekst predlozene podele -- racuna ga `PlanBlokRaspodela`, ISTI planer po kome
' `MapBankaImportAsKooperantBlockCore` knjizi (prikaz i akcija ne mogu da se
' razidju). Visak preko otvorenih stavki ide u avans, kao i u knjizenju.
Private Function SplitPreviewText(ByVal kandidati As Variant, ByVal iznos As Double) As String
    Dim plan As Variant
    Dim s As String
    Dim i As Long
    Dim podeljeno As Double

    plan = PlanBlokRaspodela(kandidati, iznos)

    If Not IsEmpty(plan) Then
        For i = 1 To UBound(plan, 1)
            s = s & " - " & CStr(plan(i, 1)) & ": " & _
                Format$(CDbl(plan(i, 2)), "#,##0.00") & " (" & CStr(plan(i, 3)) & ")" & vbCrLf
            podeljeno = podeljeno + CDbl(plan(i, 2))
        Next i
    End If

    If iznos - podeljeno > 0.009 Then
        s = s & " - visak " & Format$(iznos - podeljeno, "#,##0.00") & " -> " & _
            NOV_VIRMAN_AVANS_KOOP & vbCrLf
    End If

    SplitPreviewText = s
End Function

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
    On Error GoTo EH

    Dim bimID As String

    lblPreview.caption = ""

    If lstBanka.ListIndex < 0 Then Exit Sub

    bimID = m_BimIDs(lstBanka.ListIndex)
    lblPreview.caption = BuildAutoPreviewText(bimID) & BuildManualPreviewText(bimID)
    Exit Sub

EH:
    ' Preview sme da padne (npr. schema greska iz resolvera kandidata koja se
    ' NAMERNO propagira umesto da se tumaci kao "previse kandidata"). Prikazi
    ' razlog u panelu -- klik na listu ne sme da otvori VBA runtime dijalog.
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmBankaImport.UpdateAutoPreview"

    On Error Resume Next
    lblPreview.caption = "Preview nije dostupan: " & errDesc
End Sub

' FM-0024 #3: izvor za PRIKAZ mora biti isti kao izvor za KOMANDU -- ali AUTO i
' RUCNO imaju RAZLICITE komande, pa i razlicite izvore bloka:
'
'   AUTO   (btnAutoJedan/btnAutoSve -> AutoMapBankaImportRow_TX)
'          = iskljucivo poziv na broj iz izvoda; writer ne vidi kontrole forme.
'          -> `AutoBlockNoForBim` (modBankaMapiranje, ISTA funkcija koju zove writer)
'
'   RUCNO  (btnSacuvajRucno -> MapBankaImportAsKooperantBlockManual_TX)
'          = izbor u cmbOtkupBlok, a ako nije izabran onda poziv na broj.
'          -> ova funkcija
'
' Kad je ovo bila JEDNA funkcija za oba preview-a, rucni izbor bloka je menjao i
' AUTO prikaz: preview je pokazivao blok B, a "Automatski mapiraj red" knjizio A.
' SME LI RUCNO MAPIRANJE KOOPERANTA UOPSTE DA POCNE.
'
' Izdvojeno iz btnSacuvajRucno_Click da bi se moglo IZMERITI. Dok je uslov stajao
' u samom handleru, jedini nacin da se proveri bio je da neko rukom otvori formu:
' kapija je bila tacna, ali "provereno citanjem" -- a bas tu klasu gresaka
' (prazna lista tumacena kao izbor) su poslednja tri PR-a nalazila tri puta.
'
' Poruka se vraca pozivaocu umesto da se prikaze ovde, da funkcija ostane bez
' dijaloga i time pozivljiva iz testa.
Private Function KooperantRucnoSme(ByRef outPoruka As String) As Boolean
    outPoruka = ""

    If Not m_BlokoviLoadOk Then
        outPoruka = "Lista blokova NIJE u" & ChrW(269) & "itana (" & _
                    m_BlokoviLoadErr & ")." & vbCrLf & _
                    "Prazna lista NE zna" & ChrW(269) & "i poziv na broj."
        Exit Function
    End If

    KooperantRucnoSme = True
End Function

' ============================================================
' TEST SEAM-OVI
'
' Postoje samo za test i TVRDO su gejtovani -- van test rezima ne rade nista.
' Isti obrazac kao modScrDokumenti.Scr_OtpTestSet.
'
' Zasto uopste: pravila ove forme (ucitanost liste, "blok je izabran") odlucuju
' hoce li uplata postati avans. Do sada su bila proverljiva samo rukom, pa su
' ista greska i njena ispravka dva puta prosle kroz review umesto kroz suite.
' Forma se u testu NE prikazuje -- UserForm_Activate (koji cita tabele) se zato
' nikad ne izvrsava.
' ============================================================
Public Sub BiTestSetUcitanost(ByVal blokoviOk As Boolean, ByVal greska As String)
    If Not IsTestMode() Then Exit Sub
    m_BlokoviLoadOk = blokoviOk
    m_BlokoviLoadErr = greska
End Sub

Public Sub BiTestSetFaktureUcitanost(ByVal faktureOk As Boolean, ByVal greska As String)
    If Not IsTestMode() Then Exit Sub
    m_FaktureLoadOk = faktureOk
    m_FaktureLoadErr = greska
End Sub

' Odluka koju handler stvarno cita, ne njena kopija.
Public Function BiTestKooperantSme() As Boolean
    Dim poruka As String
    If Not IsTestMode() Then Exit Function
    BiTestKooperantSme = KooperantRucnoSme(poruka)
End Function

Public Function BiTestKooperantPoruka() As String
    Dim poruka As String
    If Not IsTestMode() Then Exit Function
    KooperantRucnoSme poruka
    BiTestKooperantPoruka = poruka
End Function

' Izbor u combo-ima se postavlja OVDE, a ne iz testa: kontrole su Private clanovi
' forme, a i sam upis mora da prodje kroz istu formu kroz koju prolazi operater.
Public Sub BiTestSetIzbor(ByVal tip As String, ByVal blok As String)
    If Not IsTestMode() Then Exit Sub
    On Error Resume Next
    cmbMapTip.value = tip
    cmbOtkupBlok.value = blok
End Sub

Public Function BiTestBlokIzabran() As Boolean
    If Not IsTestMode() Then Exit Function
    BiTestBlokIzabran = ManualBlokIzabran()
End Function

' Da li je blok IZABRAN iz liste, ili je izveden iz poziva na broj.
'
' Ista razlika koju vec pravi EffectiveManualBlockNo, samo imenovana: writer na
' osnovu nje odlucuje sme li prazan skup kandidata da postane avans. Izabran blok
' bez otvorenih stavki je protivrecnost -- operater je rekao KOJI dug placa.
Private Function ManualBlokIzabran() As Boolean
    If Trim$(nz(cmbMapTip.value, "")) <> "Kooperant" Then Exit Function
    ManualBlokIzabran = (Trim$(nz(cmbOtkupBlok.value, "")) <> "")
End Function

Private Function EffectiveManualBlockNo(ByVal bankaImportID As String) As String
    Dim manualBlok As String

    If Trim$(nz(cmbMapTip.value, "")) = "Kooperant" Then
        manualBlok = Trim$(nz(cmbOtkupBlok.value, ""))
    End If

    If manualBlok <> "" Then
        EffectiveManualBlockNo = manualBlok
    Else
        EffectiveManualBlockNo = AutoBlockNoForBim(bankaImportID)
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

    ' ISTI klasifikator koji koristi writer (RequireBimSmer preko ClassifyBimSmer),
    ' da preview ne bi prikazao kao validan red koji ce komanda odbiti -- npr. red
    ' sa OBA iznosa (nejasan smer) je ranije u preview-u prolazio kao ispravan.
    Dim smer As String
    smer = ClassifyBimSmer(uplata, isplata)

    s = s & "Tip: " & tipMap & vbCrLf

    If smer = BIM_SMER_NEJASAN Then
        BuildManualPreviewText = s & "ODBIJENO: stavka nema cist smer " & _
            "(uplata " & Format$(uplata, "#,##0.00") & " / isplata " & _
            Format$(isplata, "#,##0.00") & ") -- nijedan tip nije dozvoljen."
        Exit Function
    End If

    Select Case tipMap
        Case "Kupac"
            If smer <> BIM_SMER_UPLATA Then
                BuildManualPreviewText = s & "ODBIJENO: stavka je isplata, a tip Kupac trazi uplatu."
                Exit Function
            End If

            Dim kupacID As String
            Dim fakturaID As String

            kupacID = SelectedPartnerID()
            fakturaID = ExtractIDFromDisplay(Trim$(nz(cmbFaktura.value, "")))

            s = s & "KupacID: " & kupacID & vbCrLf

            If Not m_FaktureLoadOk Then
                BuildManualPreviewText = s & "ODBIJENO: lista faktura nije ucitana (" & _
                    m_FaktureLoadErr & "). Prazna lista NE znaci avans."
                Exit Function
            End If

            If fakturaID <> "" Then
                ' Ista raspodela koju ce writer da proknjizi: na fakturu ide
                ' najvise njen otvoren iznos, visak je avans kupca.
                Dim otvorenoFak As Double
                Dim naFakturu As Double

                otvorenoFak = GetOtvorenoNaFakturi(fakturaID)

                s = s & "FakturaID: " & fakturaID & vbCrLf
                s = s & "Otvoreno na fakturi: " & Format$(otvorenoFak, "#,##0.00") & vbCrLf

                If otvorenoFak <= 0.009 Then
                    BuildManualPreviewText = s & "ODBIJENO: faktura nema otvoren iznos " & _
                        "(u medjuvremenu je placena) -- osvezi listu."
                    Exit Function
                End If

                If uplata <= otvorenoFak Then
                    naFakturu = uplata
                Else
                    naFakturu = otvorenoFak
                End If

                s = s & "Na fakturu: " & Format$(naFakturu, "#,##0.00") & _
                    " (" & NOV_KUPCI_UPLATA & ")" & vbCrLf

                If uplata - naFakturu > 0.009 Then
                    s = s & "Visak -> avans: " & Format$(uplata - naFakturu, "#,##0.00") & _
                        " (" & NOV_KUPCI_AVANS & ")"
                End If
            Else
                s = s & "Faktura: nije izabrana -> knjizi se kao AVANS" & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
            End If

        Case "Kooperant"
            If smer <> BIM_SMER_ISPLATA Then
                BuildManualPreviewText = s & "ODBIJENO: stavka je uplata, a tip Kooperant trazi isplatu."
                Exit Function
            End If

            Dim kooperantID As String
            Dim blokNo As String
            Dim kandidati As Variant
            Dim manualRequiredText As String

            kooperantID = SelectedPartnerID()
            blokNo = EffectiveManualBlockNo(bankaImportID)

            s = s & "KooperantID: " & kooperantID & vbCrLf
            If Trim$(nz(cmbOtkupBlok.value, "")) <> "" Then
                s = s & "Blok (rucni izbor): " & blokNo & vbCrLf
            ElseIf blokNo <> "" Then
                s = s & "Blok (poziv na broj): " & blokNo & vbCrLf
            End If

            kandidati = SafeBlockCandidates(kooperantID, blokNo, manualRequiredText)

            If manualRequiredText = "" Then
                s = s & CandidatesPreviewText(kooperantID, blokNo)
            Else
                ' Preko granice automatske raspodele: rucni put ovo MOZE da zavrsi
                ' (uz potvrdu), pa preview mora da pokaze punu listu i predlozenu
                ' podelu -- isti planer koji ce knjiziti.
                kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blokNo, True)

                s = s & "Automatska raspodela ODBIJENA (blok ima " & _
                    CStr(UBound(kandidati, 1)) & " otvorene stavke)." & vbCrLf
                s = s & "Rucno knjizenje trazi potvrdu ove podele:" & vbCrLf
                s = s & SplitPreviewText(kandidati, isplata)
            End If

        Case "OM"
            Dim omID As String
            omID = SelectedPartnerID()
            s = s & "OMID: " & omID & vbCrLf
            s = s & "Smer: " & smer & vbCrLf
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

' Kandidati bloka za PRIKAZ, sa jasno ogranicenim izuzetkom.
'
' SAMO `ERR_BMAP_MANUAL_REQUIRED` (3+ otvorenih stavki) se pretvara u tekst i
' otvara "rucno uz potvrdu" tok -- to je jedina greska koja znaci "operater
' odlucuje". Svaka druga (schema, citanje, obracun salda) se PROPAGIRA: ranije su
' sve greske izgledale isto, pa je npr. schema problem vodio u tok koji nudi
' potvrdu podele, kao da je rec o previse kandidata.
Private Function SafeBlockCandidates(ByVal kooperantID As String, _
                                     ByVal blokNo As String, _
                                     ByRef manualRequiredText As String) As Variant
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    manualRequiredText = ""

    On Error Resume Next
    SafeBlockCandidates = GetOtkupCandidatesForKooperantBlock(kooperantID, blokNo)
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    If errNum <> 0 Then Err.Clear
    On Error GoTo 0

    If errNum = 0 Then Exit Function

    SafeBlockCandidates = Empty

    If errNum = ERR_BMAP_MANUAL_REQUIRED Then
        manualRequiredText = errDesc
        Exit Function
    End If

    Err.Raise errNum, errSrc, errDesc
End Function

' Jedan tekst kandidata bloka za SVE preview grane (auto: jak kljuc / PartnerMap /
' heuristika, i rucno) - ranije je isti blok bio prepisan cetiri puta.
Private Function CandidatesPreviewText(ByVal kooperantID As String, _
                                       ByVal blokNo As String) As String
    Dim kandidati As Variant
    Dim manualRequiredText As String
    Dim s As String
    Dim i As Long

    kandidati = SafeBlockCandidates(kooperantID, blokNo, manualRequiredText)

    If manualRequiredText <> "" Then
        CandidatesPreviewText = "Otkup kandidati: " & manualRequiredText & vbCrLf & _
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
        blockNo = AutoBlockNoForBim(bankaImportID)

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
                blockNo = AutoBlockNoForBim(bankaImportID)

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
        blockNo = AutoBlockNoForBim(bankaImportID)

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
    ' LogErr PRE "On Error Resume Next" -- ona resetuje Err, pa bi poziv ispod
    ' nje upisao nista.
    LogErr "frmBankaImport.UpdateIzvodSummaryLabel"
    On Error Resume Next
    Me.lblIzvodSummary.caption = "(gre" & ChrW(353) & "ka pri " & ChrW(269) & "itanju saldo info-a)"
    Me.lblIzvodSummary.ForeColor = RGB(128, 128, 128)
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
