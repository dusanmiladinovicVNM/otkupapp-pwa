VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBankaExportPregled 
   Caption         =   "UserForm1"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmBankaExportPregled.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBankaExportPregled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_Blokovi As Collection            ' prikazani blokovi (posle kooperant-filtera)
Private m_FullBlokovi As Collection        ' pre kooperant-filtera (prune + combo lista)
Private m_OverrideAmounts As Object        ' Dictionary OtkupID -> custom amount
Private m_SetupDone As Boolean
Private mChromeRemoved As Boolean

' Runtime kontrole (.frx se ne dira): filter po kooperantu + izbor racuna firme
Private mCmbKooperant As MSForms.ComboBox
Private m_uiSinks As Object          ' tag -> clsUiSink (self-update: WithEvents van formi)
Private mLblKooperant As MSForms.label
Private mCmbRacun As MSForms.ComboBox
Private mLblRacun As MSForms.label
Private mKoopComboFilling As Boolean       ' guard: refill ne sme da okine Change
Private mRacuni() As String                ' cisti racuni, paralelno sa mCmbRacun stavkama
Private mRacuniCount As Long

' Runtime dugmad: vezivanje virman avansa (NOV_VIRMAN_AVANS_KOOP) na blokove.
' Motor je ApplyAvansToOtkup_TX (postojeci, transakciono bezbedan); dugmad
' ga samo pozivaju za izabran blok / cekirane blokove.
Private mBtnAvansBlok As MSForms.CommandButton    ' fokusiran blok (detail panel)
Private mBtnAvansSel As MSForms.CommandButton     ' cekirani blokovi (akcije)
Private Const AVANS_BLOK_CAPTION As String = "Primeni avans na blok"
Private Const AVANS_SEL_CAPTION As String = "Primeni avans (sel.)"

Private Sub UserForm_Activate()
    On Error GoTo EH
    MouseWheel_Attach Me

    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    If m_SetupDone Then
        PopulateRacunCombo    ' racuni u configu su se mogli promeniti u medjuvremenu
        LoadBlokovi
        Exit Sub
    End If
    m_SetupDone = True

    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    EnsureRuntimeControls
    
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StylePrimaryButton btnExport, "PDF specifikacija"
    StylePrimaryButton btnPostaviFull, "Postavi na otvoreno"
    StylePrimaryButton btnGenerisiCSV, Poruka("BANKA_LBL_GENERISI_CSV_COMMIT")
    btnGenerisiCSV.enabled = True
    StyleExitButton btnPovratak, "Povratak"

    StyleLabel lblStatus, TXT_MUTED(), True
    StyleLabel lblSelectionSummary, TXT_MUTED(), True
    StyleSubtitle lblSubtitle, "Pregled otvorenih blokova za isplatu"
    
    StyleSectionHeader fraFilter, "Filteri"
    StyleSectionHeader fraDetail, "Detalji izabranog bloka"
    
    StyleListHeaderLabel lblColDatum
    StyleListHeaderLabel lblColKooperant
    StyleListHeaderLabel lblColStanica
    StyleListHeaderLabel lblColBrojDok
    StyleListHeaderLabel lblColUkupan
    StyleListHeaderLabel lblColIsplaceno
    StyleListHeaderLabel lblColOtvoren
    StyleListHeaderLabel lblColTR
    StyleListHeaderLabel lblColIsplatiti
    
    LayoutTopKpis
    RefreshTopKpis
    
    Set m_OverrideAmounts = CreateObject("Scripting.Dictionary")
    
    SetupList
    PopulateStanicaCombo
    LoadBlokovi
    ClearDetailPanel
    
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "frmBankaExportPregled.UserForm_Activate"
    MsgBox "Gre" & ChrW(353) & "ka pri otvaranju pregleda: " & errDesc, vbCritical, APP_NAME
End Sub

Private Sub SetupList()
    With lstBlokovi
        .ColumnCount = 9   ' jedna vise za "Isplatiti"
        ' Datum | Kooperant | Stanica | BrojDok | Ukupan | Isplaceno | Otvoreno | TR | Isplatiti
        .ColumnWidths = "60;140;50;60;75;75;75;30;75"
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption    ' native checkboxovi
    End With
End Sub

Private Sub PopulateStanicaCombo()
    cmbStanica.Clear
    cmbStanica.AddItem ""    ' empty = "Sve stanice"
    FillCmb cmbStanica, GetLookupList(TBL_STANICE, "Naziv")
End Sub

'======================================================================
' Runtime kontrole - .frx se ne dira (Controls.Add idiom kao frmPalete/
' frmAgrohemija): kooperant filter u filter redu + izbor racuna firme
' uz action dugmad.
'======================================================================
Private Sub EnsureRuntimeControls()
    EnsureKooperantFilter
    EnsureRacunCombo
    EnsureAvansButtons        ' posle EnsureRacunCombo (sel. dugme se sidri levo od "Sa racuna")
End Sub

'======================================================================
' EnsureAvansButtons - runtime dugmad za vezivanje virman avansa na blokove
' (.frx se ne dira). Po bloku (dole u detail panelu) + na cekirane blokove
' (u akcijskoj traci, levo od "Sa racuna" combo-a).
'======================================================================
Private Sub EnsureAvansButtons()
    On Error GoTo EH

    ' 1) Po bloku -- dno detalj frejma (ispod avans info-a)
    If mBtnAvansBlok Is Nothing Then
        Dim dp As Object
        Set dp = btnPostaviFull.Parent
        Set mBtnAvansBlok = dp.Controls.Add("Forms.CommandButton.1", "btnAvansBlok", True)
        WireSink mBtnAvansBlok, "mBtnAvansBlok"
        With mBtnAvansBlok
            .Left = btnPostaviFull.Left
            .width = dp.InsideWidth - btnPostaviFull.Left - 10
            If .width < 80 Then .width = btnPostaviFull.width
            .Height = btnPostaviFull.Height
            .top = dp.InsideHeight - .Height - 10
            .enabled = False
        End With
        StylePrimaryButton mBtnAvansBlok, AVANS_BLOK_CAPTION
    End If

    ' 2) Na cekirane -- akcijska traka, levo od "Sa racuna" combo-a
    If mBtnAvansSel Is Nothing Then
        Dim ah As Object
        Set ah = btnGenerisiCSV.Parent
        Set mBtnAvansSel = ah.Controls.Add("Forms.CommandButton.1", "btnAvansSel", True)
        WireSink mBtnAvansSel, "mBtnAvansSel"
        With mBtnAvansSel
            .width = 130
            .Height = btnGenerisiCSV.Height
            .top = btnGenerisiCSV.top
            If Not mLblRacun Is Nothing Then
                .Left = mLblRacun.Left - .width - 18
            Else
                .Left = btnGenerisiCSV.Left - .width - 18
            End If
            If .Left < 6 Then .Left = 6
        End With
        StylePrimaryButton mBtnAvansSel, AVANS_SEL_CAPTION
    End If
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.EnsureAvansButtons"
End Sub

' Desna ivica postojeceg filter reda (u istom kontejneru kao cmbStanica).
Private Function FilterRowRightEdge(ByVal host As Object) As Single
    Dim edge As Single
    edge = cmbStanica.Left + cmbStanica.width
    On Error Resume Next
    If txtDatumOd.Parent Is host Then
        If txtDatumOd.Left + txtDatumOd.width > edge Then edge = txtDatumOd.Left + txtDatumOd.width
    End If
    If txtDatumDo.Parent Is host Then
        If txtDatumDo.Left + txtDatumDo.width > edge Then edge = txtDatumDo.Left + txtDatumDo.width
    End If
    On Error GoTo 0
    FilterRowRightEdge = edge
End Function

' Combo "Kooperant" u filter redu: radi i na unos (autocomplete +
' substring filter) i kao padajuca lista. Prazno = svi kooperanti.
Private Sub EnsureKooperantFilter()
    On Error GoTo EH
    If Not mCmbKooperant Is Nothing Then Exit Sub

    Dim host As Object
    Set host = cmbStanica.Parent

    ' Osvezi dugme (iz .frx) je stajalo odmah posle datum polja -> nalegalo je
    ' na kooperant combo. Pomeri ga desno (ali odmaknuto od ivice) i visinski
    ' centriraj u frame; combo ostaje odmah posle datuma (gde je i bio).
    On Error Resume Next
    Dim ob As Object
    Set ob = btnOsvezi.Parent
    btnOsvezi.Left = ob.InsideWidth - btnOsvezi.width - 48      ' malo od desne ivice
    btnOsvezi.top = (ob.InsideHeight - btnOsvezi.Height) / 2    ' visinski centrirano u frame
    On Error GoTo EH

    Dim edge As Single
    edge = FilterRowRightEdge(host)

    Set mLblKooperant = host.Controls.Add("Forms.Label.1", "lblKooperantFilter", True)
    With mLblKooperant
        .caption = "Kooperant:"
        .Left = edge + 14
        .top = cmbStanica.top + 3
        .width = 56
        .Height = 12
    End With
    StyleLabel mLblKooperant, TXT_MUTED(), True

    Set mCmbKooperant = host.Controls.Add("Forms.ComboBox.1", "cmbKooperantFilter", True)
    WireSink mCmbKooperant, "mCmbKooperant"
    With mCmbKooperant
        .Left = mLblKooperant.Left + mLblKooperant.width + 4
        .top = cmbStanica.top
        .width = 160
        .Height = cmbStanica.Height
        .style = fmStyleDropDownCombo         ' radi i na unos i kao dropdown
        .MatchEntry = fmMatchEntryComplete    ' autocomplete pri kucanju
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        .ControlTipText = "Kucaj deo imena ili izaberi; prazno = svi kooperanti"
    End With
    StyleComboBox mCmbKooperant
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.EnsureKooperantFilter"
End Sub

' Combo "Sa racuna" uz action dugmad: racun firme sa koga idu nalozi
' (BankaNalogRacuniCSV: RACUN_1/2/3, fallback legacy pa SELLER_ACCOUNT). Samo izbor.
Private Sub EnsureRacunCombo()
    On Error GoTo EH
    If Not mCmbRacun Is Nothing Then Exit Sub

    Dim host As Object
    Set host = btnGenerisiCSV.Parent

    ' sidro = najlevlje od action dugmadi (Export / CSV)
    Dim anchorLeft As Single, anchorTop As Single, anchorH As Single
    anchorLeft = btnGenerisiCSV.Left
    anchorTop = btnGenerisiCSV.top
    anchorH = btnGenerisiCSV.Height
    On Error Resume Next
    If btnExport.Parent Is host Then
        If btnExport.Left < anchorLeft Then
            anchorLeft = btnExport.Left
            anchorTop = btnExport.top
            anchorH = btnExport.Height
        End If
    End If
    On Error GoTo EH

    Const CMB_W As Single = 185
    Const LBL_W As Single = 52

    Set mLblRacun = host.Controls.Add("Forms.Label.1", "lblRacunIsplate", True)
    Set mCmbRacun = host.Controls.Add("Forms.ComboBox.1", "cmbRacunIsplate", True)

    With mLblRacun
        .caption = "Sa ra" & ChrW(269) & "una:"
        .width = LBL_W
        .Height = 12
    End With

    With mCmbRacun
        .width = CMB_W
        .style = fmStyleDropDownList      ' samo izbor - racun mora biti tacan
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        .ControlTipText = "Ra" & ChrW(269) & "un firme sa koga idu nalozi (Pode" & _
                          ChrW(353) & "avanja -> Banka / nalozi)"
    End With

    If anchorLeft - CMB_W - LBL_W - 24 >= 6 Then
        ' ista linija, levo od dugmadi
        mCmbRacun.Left = anchorLeft - CMB_W - 12
        mCmbRacun.top = anchorTop + (anchorH - 18) / 2
        mLblRacun.Left = mCmbRacun.Left - LBL_W - 4
        mLblRacun.top = mCmbRacun.top + 3
    Else
        ' fallback: iznad dugmeta
        mCmbRacun.Left = anchorLeft
        mCmbRacun.top = anchorTop - 22
        mLblRacun.Left = anchorLeft - LBL_W - 4
        If mLblRacun.Left < 4 Then mLblRacun.Left = 4
        mLblRacun.top = mCmbRacun.top + 3
    End If

    StyleLabel mLblRacun, TXT_MUTED(), True
    StyleComboBox mCmbRacun

    PopulateRacunCombo
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.EnsureRacunCombo"
End Sub

' Napuni combo racuna iz BankaNalogRacuniCSV (tri polja RACUN_1/2/3; fallback na
' legacy ";"-spisak pa SELLER_ACCOUNT). Default izbor = SELLER_ACCOUNT ako je
' medju racunima.
Private Sub PopulateRacunCombo()
    On Error GoTo EH
    If mCmbRacun Is Nothing Then Exit Sub

    Dim raw As String
    raw = BankaNalogRacuniCSV()

    mRacuniCount = 0
    ReDim mRacuni(0 To 0)
    mCmbRacun.Clear

    Dim parts() As String
    parts = Split(raw, ";")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim r As String
        r = Replace(Trim$(parts(i)), " ", "")
        If LenB(r) > 0 Then
            ReDim Preserve mRacuni(0 To mRacuniCount)
            mRacuni(mRacuniCount) = r
            Dim disp As String
            disp = r
            Dim bn As String
            bn = BankaNazivZaRacun(r)
            If LenB(bn) > 0 Then disp = disp & "  (" & bn & ")"
            mCmbRacun.AddItem disp
            mRacuniCount = mRacuniCount + 1
        End If
    Next i

    If mRacuniCount > 0 Then
        mCmbRacun.ListIndex = 0
        Dim def As String
        def = Replace(Trim$(DocConfigOr("SELLER_ACCOUNT", "")), " ", "")
        If LenB(def) > 0 Then
            Dim k As Long
            For k = 0 To mRacuniCount - 1
                If mRacuni(k) = def Then mCmbRacun.ListIndex = k: Exit For
            Next k
        End If
    End If
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.PopulateRacunCombo"
End Sub

' Cist (normalizovan) racun trenutno izabran u combo-u; "" ako nema.
Private Function SelectedRacun() As String
    If mCmbRacun Is Nothing Then Exit Function
    If mCmbRacun.ListIndex >= 0 And mCmbRacun.ListIndex < mRacuniCount Then
        SelectedRacun = mRacuni(mCmbRacun.ListIndex)
    End If
End Function

' Distinct imena kooperanata iz PUNE liste u combo (sortirano); cuva
' trenutno ukucani tekst, guard da ne okine Change -> LoadBlokovi.
Private Sub RefillKooperantCombo()
    On Error GoTo EH
    If mCmbKooperant Is Nothing Then Exit Sub
    If m_FullBlokovi Is Nothing Then Exit Sub

    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1   ' TextCompare

    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In m_FullBlokovi
        Set blk = v
        If LenB(Trim$(blk.kooperantNaziv)) > 0 Then
            If Not names.Exists(blk.kooperantNaziv) Then names.Add blk.kooperantNaziv, True
        End If
    Next v

    Dim arr As Variant
    arr = names.keys

    ' insertion sort (mala lista) - abecedno, case-insensitive
    Dim a As Long, b As Long
    Dim t As String
    For a = LBound(arr) + 1 To UBound(arr)
        t = arr(a): b = a - 1
        Do While b >= LBound(arr)
            If StrComp(CStr(arr(b)), t, vbTextCompare) <= 0 Then Exit Do
            arr(b + 1) = arr(b): b = b - 1
        Loop
        arr(b + 1) = t
    Next a

    mKoopComboFilling = True
    Dim cur As String: cur = CStr(mCmbKooperant.value)
    mCmbKooperant.Clear
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        mCmbKooperant.AddItem CStr(arr(i))
    Next i
    mCmbKooperant.value = cur
    mKoopComboFilling = False
    Exit Sub
EH:
    mKoopComboFilling = False
    LogErr "frmBankaExportPregled.RefillKooperantCombo"
End Sub

' Substring filter po nazivu kooperanta (case-insensitive): radi i za
' kucani deo imena i za pun izbor iz liste. Prazno -> ista kolekcija.
Private Function FilterBlokoviPoKooperantu(ByVal src As Collection, _
                                           ByVal filter As String) As Collection
    If src Is Nothing Then
        Set FilterBlokoviPoKooperantu = New Collection
        Exit Function
    End If
    If LenB(Trim$(filter)) = 0 Then
        Set FilterBlokoviPoKooperantu = src
        Exit Function
    End If

    Dim result As New Collection
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In src
        Set blk = v
        If InStr(1, blk.kooperantNaziv, Trim$(filter), vbTextCompare) > 0 Then
            result.Add blk
        End If
    Next v
    Set FilterBlokoviPoKooperantu = result
End Function

Private Sub mCmbKooperant_Change()
    If mKoopComboFilling Then Exit Sub
    ApplyKooperantFilter          ' lagani re-filter, BEZ rebuild-a pune liste
End Sub

'======================================================================
' LoadBlokovi - PUN rebuild otvorenih blokova (cita tabele). Skupo, pa se
' zove SAMO kad se izvor menja: otvaranje forme, Osvezi, datum/stanica
' filter. NE zove se na kooperant filter (to je view-filter, ApplyKooperantFilter).
'======================================================================
Private Sub LoadBlokovi()
    On Error GoTo EH

    Dim datumOd As Date, datumDo As Date
    Dim stanicaID As String

    On Error Resume Next
    If Len(Trim$(txtDatumOd.value)) > 0 Then datumOd = CDate(txtDatumOd.value)
    If Len(Trim$(txtDatumDo.value)) > 0 Then datumDo = CDate(txtDatumDo.value)
    On Error GoTo EH

    If Len(Trim$(cmbStanica.value)) > 0 Then
        stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbStanica.value, "StanicaID"))
    End If

    Set m_FullBlokovi = BuildBlokIsplataList(datumOd, datumDo, stanicaID)

    ' Override cleanup + punjenje combo imena idu protiv PUNE liste -> SAMO
    ' pri rebuild-u (kooperant-filter ne sme da obrise "Isplatiti" unose)
    Dim uskladjeno As Long
    uskladjeno = PruneStaleOverrides()
    RefillKooperantCombo

    ApplyKooperantFilter

    ' Tiho spustanje iznosa bi operater lako promasio -- javi ga u statusu
    ' (ApplyKooperantFilter tek sto je prepisao lblStatus).
    If uskladjeno > 0 Then
        lblStatus.caption = lblStatus.caption & " | Uskla" & ChrW(273) & "eno 'Isplatiti': " & _
                            CStr(uskladjeno)
    End If
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.LoadBlokovi"
    lblStatus.caption = "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju."
End Sub

'======================================================================
' ApplyKooperantFilter - LAGANI re-filter nad vec ucitanom m_FullBlokovi
' (bez citanja tabela). Zove se na svaku promenu kooperant combo-a i na
' kraju LoadBlokovi. Substring nad nazivom radi i za kucani deo imena i za
' pun izbor iz padajuce liste; prazno = svi.
'======================================================================
Private Sub ApplyKooperantFilter()
    On Error GoTo EH

    Dim koopFilter As String
    If Not mCmbKooperant Is Nothing Then koopFilter = Trim$(CStr(mCmbKooperant.value))
    Set m_Blokovi = FilterBlokoviPoKooperantu(m_FullBlokovi, koopFilter)

    RenderListbox
    UpdateEmptyState
    lblStatus.caption = SummarizeBlokList(m_Blokovi)
    UpdateSelectionSummary
    RefreshTopKpis
    ClearDetailPanel
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.ApplyKooperantFilter"
    lblStatus.caption = "Gre" & ChrW(353) & "ka pri filtriranju."
End Sub

'======================================================================
' PruneStaleOverrides - AUD-026 / FM-0020 #1.
'
' Zaostali "Isplatiti" override-i se pri svakom rebuild-u usklade sa
' TRENUTNO otvorenim iznosima: blok kog vise nema (ili je zatvoren) gubi
' override, override veci od otvorenog se spusta na otvoreno. Ranije se
' brisao samo override za nestao blok, pa je posle delimicne isplate ili
' vezanog avansa u prikazu (i u CSV nalozima) mogao ostati iznos veci od
' otvorenog. Logika je u modBankaExportPregled.ClampOverridesToOpen, da
' bude testabilna van forme.
'
' Vraca broj promenjenih unosa (za status liniju).
'======================================================================
Private Function PruneStaleOverrides() As Long
    PruneStaleOverrides = ClampOverridesToOpen(m_OverrideAmounts, m_FullBlokovi)
End Function

Private Sub RenderListbox()
    If m_Blokovi Is Nothing Then
        lstBlokovi.Clear
        Exit Sub
    End If

    Dim n As Long: n = m_Blokovi.count
    If n = 0 Then
        lstBlokovi.Clear
        Exit Sub
    End If

    ' 2D niz + JEDAN .List upis (umesto AddItem + po-celija: ~9x manje COM
    ' poziva, bitno na 800+ redova -- isti .List=arr pattern kao izvestaji).
    Dim arr() As Variant
    ReDim arr(0 To n - 1, 0 To 8)

    Dim i As Long: i = 0
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In m_Blokovi
        Set blk = v
        arr(i, 0) = Format$(blk.datum, "d.m.yyyy")
        arr(i, 1) = blk.kooperantNaziv
        arr(i, 2) = blk.stanicaID
        arr(i, 3) = blk.brojDokumenta
        arr(i, 4) = Format$(blk.UkupanIznos, "#,##0.00")
        arr(i, 5) = Format$(blk.VecIsplaceno, "#,##0.00")
        arr(i, 6) = Format$(blk.OtvorenIznos, "#,##0.00")
        arr(i, 7) = IIf(blk.HasTekuciRacun, "OK", "--")
        arr(i, 8) = Format$(GetIsplatitiAmount(blk), "#,##0.00")
        i = i + 1
    Next v

    lstBlokovi.List = arr
End Sub

'======================================================================
' GetIsplatitiAmount
' Vraca current "Isplatiti" iznos za blok:
'   - override iz m_OverrideAmounts ako postoji
'   - inace OtvorenIznos
'======================================================================
Private Function GetIsplatitiAmount(ByVal blk As clsBlokIsplata) As Double
    If m_OverrideAmounts Is Nothing Then
        GetIsplatitiAmount = blk.OtvorenIznos
        Exit Function
    End If
    
    If m_OverrideAmounts.Exists(blk.otkupID) Then
        GetIsplatitiAmount = CDbl(m_OverrideAmounts(blk.otkupID))
    Else
        GetIsplatitiAmount = blk.OtvorenIznos
    End If
End Function

'======================================================================
' lstBlokovi_Click - pokazi detail panel za clicked row
'======================================================================
Private Sub lstBlokovi_Click()
    HandleListSelectionChange
End Sub

Private Sub lstBlokovi_Change()
    HandleListSelectionChange
End Sub

Private Sub HandleListSelectionChange()
    On Error GoTo EH
    
    ' Selection summary mora uvek da se update-uje
    UpdateSelectionSummary
    
    ' Detail panel samo ako ima fokusiran row (ListIndex >= 0)
    If lstBlokovi.ListIndex < 0 Then
        ClearDetailPanel
        Exit Sub
    End If
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    
    If blk Is Nothing Then
        ClearDetailPanel
        Exit Sub
    End If
    
    ' Bez TR redovi: skini check, prikazi gresku u detail panelu
    If Not blk.HasTekuciRacun Then
        lstBlokovi.Selected(lstBlokovi.ListIndex) = False
        lblDetailValidacija.caption = "Ovaj kooperant nema TekuciRacun. Ne mo" & ChrW(382) & "e biti u paketu."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
    Else
        lblDetailValidacija.Visible = False
    End If
    
    PopulateDetailPanel blk
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.HandleListSelectionChange"
End Sub

'======================================================================
' Empty state
'======================================================================

Private Sub UpdateEmptyState()
    If m_Blokovi Is Nothing Then
        lblEmptyState.caption = "Nema podataka."
        lblEmptyState.Visible = True
        lstBlokovi.Visible = False
        Exit Sub
    End If
    
    Dim koopF As String
    If Not mCmbKooperant Is Nothing Then koopF = Trim$(CStr(mCmbKooperant.value))

    If m_Blokovi.count = 0 Then
        If Len(Trim$(cmbStanica.value)) > 0 Or _
           Len(Trim$(txtDatumOd.value)) > 0 Or _
           Len(Trim$(txtDatumDo.value)) > 0 Or _
           Len(koopF) > 0 Then
            lblEmptyState.caption = "Nema rezultata za izabran filter." & vbCrLf & _
                                    "Probaj sira pravila ili klikni Osve" & ChrW(382) & "i."
        Else
            lblEmptyState.caption = "Sve otvorene stavke su zatvorene. ?" & vbCrLf & _
                                    "Nema blokova za isplatu."
        End If
        lblEmptyState.Visible = True
        lstBlokovi.Visible = False
    Else
        lblEmptyState.Visible = False
        lstBlokovi.Visible = True
    End If
End Sub

'======================================================================
' GetBlokByListIndex - mapiranje ListBox index -> clsBlokIsplata
'======================================================================
Private Function GetBlokByListIndex(ByVal idx As Long) As clsBlokIsplata
    If m_Blokovi Is Nothing Then Exit Function
    If idx < 0 Then Exit Function
    If idx >= m_Blokovi.count Then Exit Function
    
    ' Collection index = 1-based; ListBox index = 0-based
    Set GetBlokByListIndex = m_Blokovi(idx + 1)
End Function

'======================================================================
' PopulateDetailPanel - pokazi info izabranog bloka
'======================================================================
Private Sub PopulateDetailPanel(ByVal blk As clsBlokIsplata)
    lblDetailBlok.caption = blk.brojDokumenta & " -- " & blk.kooperantNaziv
    lblDetailOtvoreno.caption = "Otvoreno: " & Format$(blk.OtvorenIznos, "#,##0.00") & " RSD"
    
    Dim currentAmount As Double
    currentAmount = GetIsplatitiAmount(blk)
    txtIsplatiti.value = Format$(currentAmount, "0.00")
    
    lblDetailAvans.caption = "Kooperant avans: " & Format$(blk.KooperantAvansSaldo, "#,##0.00") & " RSD"
    If blk.KooperantAvansSaldo > 0 Then
        lblDetailAvans.ForeColor = CLR_WARNING()    ' info, ne error
    Else
        lblDetailAvans.ForeColor = TXT_MUTED()
    End If
    
    lblDetailTR.caption = "Tek. ra" & ChrW(269) & "un:" & IIf(LenB(blk.TekuciRacun) > 0, blk.TekuciRacun, "--nedostaje--")
    
    Dim imaAvans As Boolean
    imaAvans = (blk.KooperantAvansSaldo > 0 And blk.OtvorenIznos > 0)
    If imaAvans Then
        lblDetailAvansHint.caption = "Klikni 'Primeni avans na blok' da avans vezes na ovaj otkup"
        lblDetailAvansHint.Visible = True
    Else
        lblDetailAvansHint.Visible = False
    End If

    If Not mBtnAvansBlok Is Nothing Then mBtnAvansBlok.enabled = imaAvans

    EnableField txtIsplatiti
    btnPostaviFull.enabled = True
End Sub

Private Sub ClearDetailPanel()
    lblDetailBlok.caption = "(izaberi blok klikom na red)"
    lblDetailOtvoreno.caption = ""
    txtIsplatiti.value = ""
    lblDetailAvans.caption = ""
    lblDetailTR.caption = ""
    lblDetailValidacija.Visible = False
    DisableField txtIsplatiti
    btnPostaviFull.enabled = False
    lblDetailAvansHint.Visible = False
    If Not mBtnAvansBlok Is Nothing Then mBtnAvansBlok.enabled = False
End Sub

'======================================================================
' txtIsplatiti_Exit - apply custom amount sa validacijom
'======================================================================
Private Sub txtIsplatiti_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH
    RemoveFocusBorder txtIsplatiti
    
    If lstBlokovi.ListIndex < 0 Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    If blk Is Nothing Then Exit Sub
    
    Dim newAmount As Double
    If Not TryParseDouble(txtIsplatiti.value, newAmount) Then
        lblDetailValidacija.caption = "Neispravan iznos."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If
    
    ' AUD-026: sve provere idu u cent-domenu, bez tolerancije (isto pravilo
    ' kao klamp i finalna kapija -- vidi vrh modBankaExportPregled). Prag
    ' "+ 0.01" je propustao preplatu od punog centa, a bas taj cent banka
    ' isplati. Normalizacija ide PRE provere nule, da sub-cent unos ne bi
    ' prosao validaciju pa tiho ispao iz CSV-a (writer trazi iznos > 0).
    Dim iznosCent As Double, otvorenoCent As Double
    iznosCent = ZaokruziNovac(newAmount)
    otvorenoCent = ZaokruziNovac(blk.OtvorenIznos)

    If iznosCent <= 0 Then
        lblDetailValidacija.caption = "Iznos mora biti veci od 0."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If

    If iznosCent > otvorenoCent Then
        lblDetailValidacija.caption = "Iznos veci od otvorenog (" & _
                                       Format$(otvorenoCent, "#,##0.00") & ")."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If

    ' Validacija OK, zapamti override (normalizovan, da prikaz, kapija i CSV
    ' rade nad istom vrednoscu).
    If iznosCent = otvorenoCent Then
        ' Vraceno na otvoreno = ne treba override
        If m_OverrideAmounts.Exists(blk.otkupID) Then
            m_OverrideAmounts.Remove blk.otkupID
        End If
    Else
        m_OverrideAmounts(blk.otkupID) = iznosCent
    End If
    
    lblDetailValidacija.Visible = False
    
    ' Re-render samo te kolone u listbox-u
    lstBlokovi.List(lstBlokovi.ListIndex, 8) = Format$(iznosCent, "#,##0.00")
    
    UpdateSelectionSummary
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.txtIsplatiti_Exit"
End Sub

Private Sub txtIsplatiti_Enter()
    ApplyFocusBorder txtIsplatiti
End Sub

Private Sub btnPostaviFull_Click()
    On Error GoTo EH
    
    If lstBlokovi.ListIndex < 0 Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    If blk Is Nothing Then Exit Sub
    
    ' Skloni override, vrati na full
    If m_OverrideAmounts.Exists(blk.otkupID) Then
        m_OverrideAmounts.Remove blk.otkupID
    End If
    
    txtIsplatiti.value = Format$(blk.OtvorenIznos, "0.00")
    lstBlokovi.List(lstBlokovi.ListIndex, 8) = Format$(blk.OtvorenIznos, "#,##0.00")
    
    lblDetailValidacija.Visible = False
    UpdateSelectionSummary
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.btnPostaviFull_Click"
End Sub

'======================================================================
' UpdateSelectionSummary - bottom toolbar live count + sum
'======================================================================
Private Sub UpdateSelectionSummary()
    If m_Blokovi Is Nothing Then
        lblSelectionSummary.caption = ""
        Exit Sub
    End If
    
    Dim selectedCount As Long
    Dim selectedSum As Double
    Dim missingTR As Long
    
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then
            Dim blk As clsBlokIsplata
            Set blk = GetBlokByListIndex(i)
            If Not blk Is Nothing Then
                If Not blk.HasTekuciRacun Then
                    missingTR = missingTR + 1
                Else
                    selectedCount = selectedCount + 1
                    selectedSum = selectedSum + GetIsplatitiAmount(blk)
                End If
            End If
        End If
    Next i
    
    Dim msg As String
    If selectedCount = 0 And missingTR = 0 Then
        msg = "Nista nije selektovano"
        lblSelectionSummary.ForeColor = TXT_MUTED()
    Else
        msg = "Selektovano: " & selectedCount & " blokova | Suma: " & _
              Format$(selectedSum, "#,##0.00") & " RSD"
        If missingTR > 0 Then
            msg = msg & "  ? Preskoceno (bez TR): " & missingTR
        End If
        lblSelectionSummary.ForeColor = CLR_SUCCESS()
    End If
    
    lblSelectionSummary.caption = msg
End Sub
Private Sub btnOsvezi_Click()
    LoadBlokovi
End Sub

Private Sub cmbStanica_Change()
    LoadBlokovi
End Sub

Private Sub txtDatumOd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RemoveFocusBorder txtDatumOd
    LoadBlokovi
End Sub

Private Sub txtDatumDo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RemoveFocusBorder txtDatumDo
    LoadBlokovi
End Sub

Private Sub txtDatumOd_Enter():    ApplyFocusBorder txtDatumOd:    End Sub
Private Sub txtDatumDo_Enter():    ApplyFocusBorder txtDatumDo:    End Sub

'======================================================================
' btnExport - PDF specifikacija isplata (umesto ranijeg TSV clipboard-a).
' Isti izbor blokova i iznosa kao CSV nalozi: selektovani (ili svi),
' iznos po bloku = "Isplatiti" (operater unos ili otvoreno).
'======================================================================
Private Sub btnExport_Click()
    On Error GoTo EH

    If m_Blokovi Is Nothing Then
        MsgBox "Nema podataka za specifikaciju.", vbInformation, APP_NAME
        Exit Sub
    End If

    If m_Blokovi.count = 0 Then
        MsgBox "Nema podataka za specifikaciju.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim missingTR As Long
    Dim blokovi As Collection
    Set blokovi = CollectIsplataBlokovi(missingTR)

    If blokovi.count = 0 Then
        MsgBox "Nema blokova za specifikaciju: izabrani blokovi nemaju teku" & _
               ChrW(263) & "i ra" & ChrW(269) & "un.", vbExclamation, APP_NAME
        Exit Sub
    End If

    PrintIsplataSpecifikacija blokovi, SelectedRacun()

    lblStatus.caption = "Specifikacija: " & blokovi.count & " blokova | " & _
                        Format$(SumIsplatiti(blokovi), "#,##0.00") & " RSD"
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "frmBankaExportPregled.btnExport_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri izradi specifikacije: " & errDesc, vbCritical, APP_NAME
End Sub

Private Function CountSelected() As Long
    Dim n As Long
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then n = n + 1
    Next i
    CountSelected = n
End Function

'======================================================================
' CollectIsplataBlokovi - blokovi za CSV naloge / specifikaciju isplata:
'   - selektovani redovi ako selekcije ima, inace svi prikazani
'   - preskace blokove bez tekuceg racuna (broji ih u outMissingTR)
'   - IsplatitiIznos = operater unos (override) ili OtvorenIznos
'======================================================================
Private Function CollectIsplataBlokovi(ByRef outMissingTR As Long) As Collection
    Dim result As New Collection
    outMissingTR = 0

    Dim hasSelection As Boolean
    hasSelection = (CountSelected() > 0)

    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If hasSelection And Not lstBlokovi.Selected(i) Then GoTo NextRow

        Dim blk As clsBlokIsplata
        Set blk = GetBlokByListIndex(i)
        If blk Is Nothing Then GoTo NextRow

        If Not blk.HasTekuciRacun Then
            outMissingTR = outMissingTR + 1
            GoTo NextRow
        End If

        blk.IsplatitiIznos = GetIsplatitiAmount(blk)
        If blk.IsplatitiIznos > 0 Then result.Add blk
NextRow:
    Next i

    Set CollectIsplataBlokovi = result
End Function

' Suma IsplatitiIznos preko kolekcije (za potvrdu i status liniju).
Private Function SumIsplatiti(ByVal blokovi As Collection) As Double
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        SumIsplatiti = SumIsplatiti + blk.IsplatitiIznos
    Next v
End Function

Private Sub btnPovratak_Click()
    On Error GoTo EH
    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.btnPovratak_Click"
    Unload Me
End Sub

Private Sub UserForm_Deactivate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    ReleaseUiSinks
    MouseWheel_Detach
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    ReleaseUiSinks
    MouseWheel_Detach
    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' UI sink (clsUiSink) - eventi runtime kontrola bez WithEvents u formi
' (self-update bezbedno; docs/SELF_UPDATE.md zamke #11/#20)
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
    LogErr "frmBankaExportPregled.WireSink(" & tagName & ")", Err.description
End Function

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

Public Sub UiSinkEvent(ByVal tagName As String, ByVal ev As String, ByVal arg As Object)
    Select Case tagName & "." & ev
        Case "mCmbKooperant.Change":      mCmbKooperant_Change
        Case "mBtnAvansBlok.Click":       mBtnAvansBlok_Click
        Case "mBtnAvansBlok.MouseMove":   mBtnAvansBlok_MouseMove 0, 0, 0, 0
        Case "mBtnAvansSel.Click":        mBtnAvansSel_Click
        Case "mBtnAvansSel.MouseMove":    mBtnAvansSel_MouseMove 0, 0, 0, 0
    End Select
End Sub

' Mouse hover pattern
'======================================================================
' Vezivanje virman avansa na blokove (NOV_VIRMAN_AVANS_KOOP -> OtkupID).
' Motor: ApplyAvansToOtkup_TX (postojeci, transakcija). Po vezivanju se
' skida "Isplatiti" override (otvoreno bloka se menja) i radi pun reload
' (LoadBlokovi) jer je tblNovac promenjen.
'
' AUD-026 (c) / FM-0020 #2:
' ApplyAvansToOtkup_TX vraca True i kada nije bilo sta da se veze (nema
' slobodnog avansa, blok vise nije otvoren) -- transakcija je uspela, samo
' je bila no-op. Zato se stvarno proknjizen iznos cita iz ByRef parametra
' (RF-02 / AUD-010) i tek on je dokaz da se nesto desilo: pozivalac po
' njemu broji "primenjene" avanse, a override se skida SAMO ako se otvoreno
' bloka zaista promenilo.
'======================================================================
Private Function PrimeniAvansTX(ByVal blk As clsBlokIsplata, _
                                Optional ByRef appliedAmount As Double) As Boolean
    appliedAmount = 0
    PrimeniAvansTX = ApplyAvansToOtkup_TX(blk.kooperantID, blk.otkupID, appliedAmount)
    If PrimeniAvansTX And appliedAmount > 0 Then
        If Not m_OverrideAmounts Is Nothing Then
            If m_OverrideAmounts.Exists(blk.otkupID) Then m_OverrideAmounts.Remove blk.otkupID
        End If
    End If
End Function

' Po bloku: veze avans kooperanta na fokusiran blok (do njegovog otvorenog).
Private Sub mBtnAvansBlok_Click()
    On Error GoTo EH
    If lstBlokovi.ListIndex < 0 Then Exit Sub

    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    If blk Is Nothing Then Exit Sub

    If blk.KooperantAvansSaldo <= 0 Then
        MsgBox "Kooperant nema neraspore" & ChrW(273) & "en avans.", vbInformation, APP_NAME
        Exit Sub
    End If
    If blk.OtvorenIznos <= 0 Then
        MsgBox "Blok nema otvoren iznos.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim vezuje As Double
    vezuje = blk.KooperantAvansSaldo
    If vezuje > blk.OtvorenIznos Then vezuje = blk.OtvorenIznos

    Dim msg As String
    msg = "Vezati avans na blok " & blk.brojDokumenta & " (" & blk.kooperantNaziv & ")?" & vbCrLf & vbCrLf & _
          "Otvoreno bloka:  " & Format$(blk.OtvorenIznos, "#,##0.00") & " RSD" & vbCrLf & _
          "Dostupan avans:  " & Format$(blk.KooperantAvansSaldo, "#,##0.00") & " RSD" & vbCrLf & _
          "Veze se pribl.:  " & Format$(vezuje, "#,##0.00") & " RSD" & vbCrLf & vbCrLf & _
          "(upisuje se u tblNovac -- OtkupID na avans red)"
    If MsgBox(msg, vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    Dim brDok As String: brDok = blk.brojDokumenta
    Dim primenjeno As Double
    If PrimeniAvansTX(blk, primenjeno) Then
        LoadBlokovi
        If primenjeno > 0 Then
            lblStatus.caption = "Avans vezan na blok " & brDok & ": " & _
                                Format$(primenjeno, "#,##0.00") & " RSD."
        Else
            ' TX je prosla, ali nista nije proknjizeno (avans u medjuvremenu
            ' potrosen ili blok vise nije otvoren) -- ne prijavljuj uspeh.
            lblStatus.caption = "Blok " & brDok & ": nije bilo avansa za vezivanje."
            MsgBox "Ni" & ChrW(353) & "ta nije proknji" & ChrW(382) & "eno: kooperant nema " & _
                   "slobodan avans ili blok vi" & ChrW(353) & "e nije otvoren." & vbCrLf & _
                   "Pregled je osve" & ChrW(382) & "en.", vbInformation, APP_NAME
        End If
    Else
        MsgBox "Gre" & ChrW(353) & "ka pri vezivanju avansa. Pogledajte log.", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "frmBankaExportPregled.mBtnAvansBlok_Click"
    MsgBox "Gre" & ChrW(353) & "ka: " & errDesc, vbCritical, APP_NAME
End Sub

' Na cekirane: veze avans na svaki cekiran blok sa raspolozivim avansom.
Private Sub mBtnAvansSel_Click()
    On Error GoTo EH

    Dim sel As Collection: Set sel = New Collection
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then
            Dim b As clsBlokIsplata
            Set b = GetBlokByListIndex(i)
            If Not b Is Nothing Then
                If b.KooperantAvansSaldo > 0 And b.OtvorenIznos > 0 Then sel.Add b
            End If
        End If
    Next i

    If sel.count = 0 Then
        MsgBox "Nema selektovanih blokova sa raspolo" & ChrW(382) & "ivim avansom." & vbCrLf & _
               "(cekiraj blokove ciji kooperant ima avans)", vbInformation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Vezati avans na " & sel.count & " selektovanih blokova?" & vbCrLf & vbCrLf & _
              "(upisuje se u tblNovac za svaki blok)", _
              vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    ' AUD-026 (c): broji se STVARNO proknjizen iznos, ne puko True. Uspela
    ' transakcija bez ijednog dinara je "bez promene", ne "primenjen avans".
    Dim okCount As Long, noopCount As Long, failCount As Long
    Dim okSum As Double
    Dim primenjeno As Double
    Dim v As Variant
    For Each v In sel
        Dim blk As clsBlokIsplata
        Set blk = v
        If PrimeniAvansTX(blk, primenjeno) Then
            If primenjeno > 0 Then
                okCount = okCount + 1
                okSum = okSum + primenjeno
            Else
                noopCount = noopCount + 1
            End If
        Else
            failCount = failCount + 1
        End If
    Next v

    LoadBlokovi
    lblStatus.caption = "Avans primenjen: " & okCount & " blokova | " & _
                        Format$(okSum, "#,##0.00") & " RSD" & _
                        IIf(noopCount > 0, " | bez promene: " & noopCount, "") & _
                        IIf(failCount > 0, " | gre" & ChrW(353) & "ka: " & failCount, "")
    If failCount > 0 Then
        MsgBox "Primenjeno: " & okCount & " (" & Format$(okSum, "#,##0.00") & " RSD)." & vbCrLf & _
               "Bez promene: " & noopCount & ". Gre" & ChrW(353) & "ka: " & failCount & _
               ". Pogledajte log.", vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "frmBankaExportPregled.mBtnAvansSel_Click"
    MsgBox "Gre" & ChrW(353) & "ka: " & errDesc, vbCritical, APP_NAME
End Sub

Private Sub mBtnAvansBlok_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    If mBtnAvansBlok.enabled Then ButtonHover mBtnAvansBlok
End Sub

Private Sub mBtnAvansSel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover mBtnAvansSel
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StylePrimaryButton btnExport, "PDF specifikacija"
    StylePrimaryButton btnPostaviFull, "Postavi na otvoreno"
    StylePrimaryButton btnGenerisiCSV, Poruka("BANKA_LBL_GENERISI_CSV_COMMIT")
    StyleExitButton btnPovratak, "Povratak"
    If Not mBtnAvansBlok Is Nothing Then StylePrimaryButton mBtnAvansBlok, AVANS_BLOK_CAPTION
    If Not mBtnAvansSel Is Nothing Then StylePrimaryButton mBtnAvansSel, AVANS_SEL_CAPTION
End Sub

Private Sub btnOsvezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnOsvezi
End Sub

Private Sub btnExport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnExport
End Sub

Private Sub btnPostaviFull_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPostaviFull
End Sub

Private Sub btnGenerisiCSV_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnGenerisiCSV
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

'======================================================================
' btnGenerisiCSV - CSV naloga za prenos (uvoz u e-banking).
' Selektovani blokovi (ili svi ako selekcije nema), iznos po bloku =
' "Isplatiti" (operater unos ili otvoreno). Potvrda pre upisa fajla.
'======================================================================
Private Sub btnGenerisiCSV_Click()
    On Error GoTo EH

    If m_Blokovi Is Nothing Then
        MsgBox "Nema podataka za naloge.", vbInformation, APP_NAME
        Exit Sub
    End If
    If m_Blokovi.count = 0 Then
        MsgBox "Nema podataka za naloge.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Racun platioca: combo "Sa racuna" (BankaNalogRacuniCSV: RACUN_1/2/3 / SELLER_ACCOUNT)
    Dim racun As String
    racun = SelectedRacun()
    If LenB(racun) = 0 Then
        MsgBox "Izaberite ra" & ChrW(269) & "un sa koga ide isplata." & vbCrLf & _
               "Ra" & ChrW(269) & "uni firme se unose u Pode" & ChrW(353) & "avanja -> Banka / nalozi" & vbCrLf & _
               "(ili Prodavac (firma) -> Teku" & ChrW(263) & "i ra" & ChrW(269) & "un).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim missingTR As Long
    Dim blokovi As Collection
    Set blokovi = CollectIsplataBlokovi(missingTR)

    If blokovi.count = 0 Then
        MsgBox "Nema blokova za naloge: izabrani blokovi nemaju teku" & ChrW(263) & "i ra" & ChrW(269) & "un.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim total As Double
    total = SumIsplatiti(blokovi)

    Dim racunInfo As String
    racunInfo = racun
    If LenB(BankaNazivZaRacun(racun)) > 0 Then
        racunInfo = racunInfo & " (" & BankaNazivZaRacun(racun) & ")"
    End If

    Dim msg As String
    msg = "Generisati " & blokovi.count & " naloga za prenos?" & vbCrLf & vbCrLf & _
          "Ukupan iznos: " & Format$(total, "#,##0.00") & " RSD" & vbCrLf & _
          "Sa ra" & ChrW(269) & "una: " & racunInfo & vbCrLf & _
          "Datum valute: " & Format$(Date, "d.m.yyyy")
    If missingTR > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Presko" & ChrW(269) & "eno (bez TR): " & missingTR & " blokova"
    End If

    If MsgBox(msg, vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    Dim csvPath As String
    Dim odbijeno As String
    csvPath = GenerisiNalogeCSV(blokovi, racun, odbijeno)

    ' AUD-026: finalna saldo kapija je odbila naloge (neki blok trazi vise
    ' nego sto je otvoreno). Fajl NIJE napisan; osvezi prikaz da operater
    ' vidi trenutno stanje (clamp tada spusta zaostale override-e).
    If LenB(odbijeno) > 0 Then
        MsgBox odbijeno, vbExclamation, APP_NAME
        LoadBlokovi
        Exit Sub
    End If

    If LenB(csvPath) = 0 Then
        MsgBox "Gre" & ChrW(353) & "ka pri generisanju CSV fajla. Pogledajte log.", vbCritical, APP_NAME
        Exit Sub
    End If

    lblStatus.caption = "CSV: " & blokovi.count & " naloga | " & _
                        Format$(total, "#,##0.00") & " RSD"

    MsgBox "Generisano " & blokovi.count & " naloga." & vbCrLf & vbCrLf & csvPath, _
           vbInformation, APP_NAME

    ' Otvori folder sa oznacenim fajlom (operater ga odatle uvozi u e-banking)
    On Error Resume Next
    Shell "explorer.exe /select,""" & csvPath & """", vbNormalFocus
    On Error GoTo 0
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "frmBankaExportPregled.btnGenerisiCSV_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri generisanju naloga: " & errDesc, vbCritical, APP_NAME
End Sub

'======================================================================
' Top KPI Strip
'======================================================================
Private Sub LayoutTopKpis()
    On Error GoTo EH
    
    LayoutTopKpiInternals fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent
    LayoutTopKpiInternals fraKpiSelected, lblKpiSelTitle, lblKpiSelValue, lblKpiSelAccent
    LayoutTopKpiInternals fraKpiBezTR, lblKpiTRTitle, lblKpiTRValue, lblKpiTRAccent
    LayoutTopKpiInternals fraKpiAvansPool, lblKpiAvTitle, lblKpiAvValue, lblKpiAvAccent
    
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.LayoutTopKpis"
End Sub

Private Sub RefreshTopKpis()
    On Error GoTo EH
    
    ' --- 1) OTVORENO total ---
    Dim totalOpen As Double
    Dim koopCount As Long
    Dim missingTR As Long
    Dim avansPool As Double
    Dim selSum As Double
    Dim selCount As Long
    
    Dim koopSet As Object
    Set koopSet = CreateObject("Scripting.Dictionary")
    Dim koopAvansSet As Object
    Set koopAvansSet = CreateObject("Scripting.Dictionary")
    
    If Not m_Blokovi Is Nothing Then
        Dim v As Variant
        For Each v In m_Blokovi
            Dim blk As clsBlokIsplata
            Set blk = v
            totalOpen = totalOpen + blk.OtvorenIznos
            If Not koopSet.Exists(blk.kooperantID) Then koopSet.Add blk.kooperantID, True
            If Not blk.HasTekuciRacun Then missingTR = missingTR + 1
            
            ' Avans pool po kooperantu (ne duplicirati istog kooperanta)
            If Not koopAvansSet.Exists(blk.kooperantID) Then
                koopAvansSet.Add blk.kooperantID, True
                avansPool = avansPool + blk.KooperantAvansSaldo
            End If
        Next v
    End If
    
    ' Selected
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then
            Dim sblk As clsBlokIsplata
            Set sblk = GetBlokByListIndex(i)
            If Not sblk Is Nothing Then
                If sblk.HasTekuciRacun Then
                    selCount = selCount + 1
                    selSum = selSum + GetIsplatitiAmount(sblk)
                End If
            End If
        End If
    Next i
    
    ' Card 1: OTVORENO (neutral, info)
    StyleTopKpi fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent, "neutral"
    lblKpiOtvTitle.caption = "Otvoreno ukupno"
    lblKpiOtvValue.caption = Format$(totalOpen, "#,##0") & " RSD"
    
    ' Card 2: SELECTED (ok kad >0)
    Dim selKind As String
    If selCount > 0 Then selKind = "ok" Else selKind = "neutral"
    StyleTopKpi fraKpiSelected, lblKpiSelTitle, lblKpiSelValue, lblKpiSelAccent, selKind
    lblKpiSelTitle.caption = "Selektovano"
    lblKpiSelValue.caption = selCount & " bl / " & Format$(selSum, "#,##0") & " RSD"
    
    ' Card 3: BEZ TR (warn kad >0)
    Dim trKind As String
    If missingTR > 0 Then trKind = "warn" Else trKind = "ok"
    StyleTopKpi fraKpiBezTR, lblKpiTRTitle, lblKpiTRValue, lblKpiTRAccent, trKind
    lblKpiTRTitle.caption = "Bez TR"
    If missingTR = 0 Then
        lblKpiTRValue.caption = "0 (svi imaju)"
    Else
        lblKpiTRValue.caption = missingTR & " blokova"
    End If
    
    ' Card 4: AVANS POOL (ok kad >0, info)
    Dim avKind As String
    If avansPool > 0 Then avKind = "ok" Else avKind = "neutral"
    StyleTopKpi fraKpiAvansPool, lblKpiAvTitle, lblKpiAvValue, lblKpiAvAccent, avKind
    lblKpiAvTitle.caption = "Avans pool"
    lblKpiAvValue.caption = Format$(avansPool, "#,##0") & " RSD"
    
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.RefreshTopKpis"
End Sub
