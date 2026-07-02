VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtkup 
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "frmOtkup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ============================================================
' frmOtkup v2.1 - NUR Otkup (Kooperant ? Station)
' Rechte Seite (Isporuka) wurde entfernt.
' Otpremnica/Zbirna/Prijemnica sind jetzt in frmDokumenta.
' ============================================================
Private mChromeRemoved As Boolean

' Runtime polje "Izdata ambala" & ChrW(382) & "a" (OM izdaje prazne kooperantu uz otkup).
' CLAUDE.md: nove kontrole se ne dodaju u .frx -> Controls.Add u runtime-u.
Private m_txtAmbIzdata As MSForms.TextBox

' Runtime polje "Kolicina ambalaze (II)" -- Klasa II je zaseban tblOtkup red sa
' svojom KolAmbalaze; deli "Kolicina ambalaze" red sa Klasom I (kao Kolicina/Cena),
' vidljivo samo kad je chkDveKlase ukljucen (CLAUDE.md: ne dira .frx).
Private m_txtKolAmbalazeII As MSForms.TextBox
Private m_kolAmbFullWidth As Single

' Runtime labela "Gajbe do zatvaranja aktivne palete" (info; .frx se ne dira).
Private m_lblPaletaInfo As MSForms.label

Private Sub UserForm_Activate()
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    MouseWheel_Attach Me
End Sub

Private Sub UserForm_Deactivate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_Initialize()

    ApplyFormTheme Me, BG_MAIN
    ApplyThemeToControls Me

    ' bitne kontrole - eksplicitni stil
    StylePrimaryButton btnUnos, "Unos"
    StyleExitButton btnPovratak, "Povratak"
    StyleStornoButton btnStornoOtkup, "Storno"
    StyleLabel lblUkupnoKG, TXT_MUTED, True

    StyleComboBox cmbVrstaVoca
    StyleComboBox cmbSortaVoca
    StyleComboBox cmbOtkupnoMesto
    StyleComboBox cmbKooperant
    StyleComboBox cmbParcela
    StyleComboBox cmbVozac
    StyleComboBox cmbTipAmbalaze

    StyleTextBox txtDatum
    StyleTextBox txtKolicina
    StyleTextBox txtCena
    StyleTextBox txtKolAmbalaze
    StyleTextBox txtNovac
    StyleTextBox txtPrimalac
    StyleTextBox txtBrojDokumenta
    StyleTextBox txtBrojZbirne
    StyleTextBox txtKolicinaKLII
    StyleTextBox txtCenaKLII
    
    ' Datum defaults
    txtDatum.value = Format$(Date, "d.m.yyyy")
    
    ' ComboBoxen fuellen
    FillCmb cmbVrstaVoca, GetLookupList(TBL_KULTURE, "VrstaVoca", , , True)
    FillComboDisplayID cmbOtkupnoMesto, TBL_STANICE, "Naziv", "StanicaID"
    FillCmb cmbVozac, GetVozacDisplayList()
    FillCmb cmbTipAmbalaze, GetTipAmbalazeOptions()
    
    ' Numerische Felder auf 0 setzen
    txtKolicina.value = ""
    txtCena.value = ""
    txtKolAmbalaze.value = ""
    txtNovac.value = "0"
    
    ' Klasa II - initial disabled
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    chkDveKlase.value = False
    lblUkupnoKG.caption = ""

    ' Opcioni panel "Otkupni blokovi" (na dugme; ne dira postojeci unos)
    On Error Resume Next
    AttachOtkupBlokPanel Me
    On Error GoTo 0

    ' Podrazumevana vrsta/sorta (Podesavanja) -> okida auto-cenu i auto-tip ambalaze.
    On Error Resume Next
    ApplyDefaultProizvod cmbVrstaVoca, cmbSortaVoca
    On Error GoTo 0

    ' Runtime polje "Izdata ambala" & ChrW(382) & "a" (ne dira .frx).
    SetupAmbIzdataField

    ' Runtime polje "Kol. ambalaze (II)" za Klasu II (ne dira .frx; skriveno dok
    ' chkDveKlase nije ukljucen).
    SetupKolAmbalazeIIField

    ' Runtime labela: gajbe do zatvaranja aktivne palete (info; .frx se ne dira).
    SetupPaletaInfoField
    UpdatePaletaInfo

    ' Podesavanja: disable parcela / novac+primalac kad su toggle-i iskljuceni.
    ApplyOtkupTogglesState
End Sub

' Kreira runtime "Izdata ambala" & ChrW(382) & "a" u SOPSTVENOM redu ispod "Kolicina ambalaze":
' otvara prazan red tako sto Novac/Primalac (+ njihove labele) i dugmad spusti za
' jednu visinu reda, pa popuni vakantno mesto (textbox levo + labela desno, kao
' ostala polja). Sva geometrija se MERI u runtime-u (.frx se ne cita iz koda).
Private Sub SetupAmbIzdataField()
    Const ROW_GAP As Single = 6
    On Error GoTo done

    If Not m_txtAmbIzdata Is Nothing Then Exit Sub

    ' Referentna labela ("Kolicina ambalaze") za poravnanje nove labele.
    Dim refLbl As MSForms.Control: Set refLbl = RowLabelRightOf(txtKolAmbalaze)

    ' Vakantni red = trenutna pozicija "Novac"; pomak = razmak izmedju dva reda.
    Dim yIzdata As Single: yIzdata = txtNovac.top
    Dim dy As Single: dy = txtNovac.top - txtKolAmbalaze.top
    If dy < txtKolAmbalaze.Height + ROW_GAP Then dy = txtKolAmbalaze.Height + ROW_GAP

    ' Labele susednih polja nadji PRE pomeranja (anchor.Top se menja pomakom).
    Dim lblNovac As MSForms.Control: Set lblNovac = RowLabelRightOf(txtNovac)
    Dim lblPrimalac As MSForms.Control: Set lblPrimalac = RowLabelRightOf(txtPrimalac)

    ' Spusti donji blok za jedan red da se oslobodi mesto za "Izdata ambala" & ChrW(382) & "a".
    ShiftCtlDown txtNovac, dy
    ShiftCtlDown lblNovac, dy
    ShiftCtlDown txtPrimalac, dy
    ShiftCtlDown lblPrimalac, dy
    ShiftCtlDown btnUnos, dy
    ShiftCtlDown btnStornoOtkup, dy
    ShiftCtlDown btnPovratak, dy

    ' TextBox u levoj koloni (kao ostala polja).
    Set m_txtAmbIzdata = Me.Controls.Add("Forms.TextBox.1", "txtAmbIzdataRT", True)
    With m_txtAmbIzdata
        .Left = txtKolAmbalaze.Left
        .top = yIzdata
        .width = txtKolAmbalaze.width
        .Height = txtKolAmbalaze.Height
        .ControlTipText = "Ambala" & ChrW(382) & "a koju OM izdaje kooperantu (preuzima od OM)"
    End With
    StyleTextBox m_txtAmbIzdata

    ' TabOrder: odmah posle "Kolicina ambalaze" (prati vizuelnu poziciju u formi).
    On Error Resume Next
    m_txtAmbIzdata.TabIndex = txtKolAmbalaze.TabIndex + 1
    On Error GoTo done

    ' Labela desno od textbox-a, poravnata sa ostalim labelama.
    Dim lbl As MSForms.label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblAmbIzdataRT", True)
    With lbl
        .caption = "Izdata ambala" & ChrW(382) & "a"
        .top = yIzdata + 2
        .Height = 14
        If Not refLbl Is Nothing Then
            .Left = refLbl.Left
            .width = refLbl.width
            .Font.Size = refLbl.Font.Size
        Else
            .Left = txtKolAmbalaze.Left + txtKolAmbalaze.width + 6
            .width = 120
        End If
    End With
    Exit Sub
done:
    LogErr "frmOtkup.SetupAmbIzdataField"
    Set m_txtAmbIzdata = Nothing
End Sub

' Kreira runtime polje "Kol. ambalaze (II)" za Klasu II (zaseban tblOtkup red sa
' svojom KolAmbalaze). .frx se ne dira -> Controls.Add u runtime-u. Geometrija se
' MERI iz postojecih Klasa II polja; tacan raspored proveriti u Excelu (forme se
' renderuju samo tamo). Skriveno dok chkDveKlase nije ukljucen.
Private Sub SetupKolAmbalazeIIField()
    On Error GoTo done

    If Not m_txtKolAmbalazeII Is Nothing Then Exit Sub

    ' Zapamti punu sirinu "Kolicina ambalaze" (Klasa I) da je vratimo kad se II iskljuci.
    m_kolAmbFullWidth = txtKolAmbalaze.width

    ' Polje Klase II deli "Kolicina ambalaze" red sa Klasom I, isto kao sto
    ' txtKolicinaKLII deli red sa txtKolicina (desna polovina istog reda).
    Set m_txtKolAmbalazeII = Me.Controls.Add("Forms.TextBox.1", "txtKolAmbalazeIIRT", True)
    With m_txtKolAmbalazeII
        .Left = txtKolicinaKLII.Left
        .top = txtKolAmbalaze.top
        .width = txtKolicinaKLII.width
        .Height = txtKolAmbalaze.Height
        .ControlTipText = "Broj gajbi za Klasu II (zasebno od Klase I)"
        .Visible = False
    End With
    StyleTextBox m_txtKolAmbalazeII

    On Error Resume Next
    m_txtKolAmbalazeII.TabIndex = txtKolAmbalaze.TabIndex + 1
    On Error GoTo done
    Exit Sub
done:
    LogErr "frmOtkup.SetupKolAmbalazeIIField"
    Set m_txtKolAmbalazeII = Nothing
End Sub

' Prikazi/sakrij polje Klase II; deli "Kolicina ambalaze" red sa Klasom I:
' kad je ON -> Klasu I skupi na levu polovinu (kao txtKolicina) i otkrij Klasu II;
' kad je OFF -> vrati Klasu I na punu sirinu i sakrij/ocisti Klasu II.
Private Sub ShowKolAmbalazeII(ByVal bShow As Boolean)
    On Error Resume Next
    If bShow Then
        txtKolAmbalaze.width = txtKolicina.width
    Else
        If m_kolAmbFullWidth > 0 Then txtKolAmbalaze.width = m_kolAmbFullWidth
    End If
    If Not m_txtKolAmbalazeII Is Nothing Then
        m_txtKolAmbalazeII.Visible = bShow
        If Not bShow Then m_txtKolAmbalazeII.value = ""
    End If
End Sub

' Vraca Label u istom redu, najblizi DESNO od date kontrole (labela polja).
' Nothing ako nema razumnog kandidata (cuva od hvatanja udaljene desne tabele).
Private Function RowLabelRightOf(ByVal anchor As MSForms.Control) As MSForms.Control
    On Error Resume Next
    Dim c As MSForms.Control, best As MSForms.Control
    Dim bestDx As Single: bestDx = 1000000
    For Each c In Me.Controls
        If TypeOf c Is MSForms.label Then
            If Abs(c.top - anchor.top) <= 6 And c.Left >= anchor.Left Then
                If (c.Left - anchor.Left) < bestDx Then
                    bestDx = c.Left - anchor.Left
                    Set best = c
                End If
            End If
        End If
    Next c
    If bestDx <= anchor.width + 220 Then Set RowLabelRightOf = best
End Function

' Spusti kontrolu za dy (bezbedno i kad je kontrola Nothing).
Private Sub ShiftCtlDown(ByVal ctl As MSForms.Control, ByVal dy As Single)
    If ctl Is Nothing Then Exit Sub
    On Error Resume Next
    ctl.top = ctl.top + dy
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnUnos, "Unos"
    StyleExitButton btnPovratak, "Povratak"
    StyleStornoButton btnStornoOtkup, "Storno"
End Sub

Private Sub btnUnos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnos
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub btnStornootkup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnStornoOtkup
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

Private Sub chkDveKlase_Click()
    If chkDveKlase.value Then
        EnableField txtKolicinaKLII
        EnableField txtCenaKLII
        StyleTextBox txtKolicinaKLII
        StyleTextBox txtCenaKLII
        ShowKolAmbalazeII True
        AutoFillCenaOtkup
    Else
        DisableField txtKolicinaKLII
        DisableField txtCenaKLII
        ShowKolAmbalazeII False
        lblUkupnoKG.caption = ""
        UpdatePaletaInfo
    End If
End Sub

Private Sub txtKolicinaKlII_Change()
    UpdateUkupnoKg
End Sub

Private Sub txtKolicina_Change()
    UpdateUkupnoKg
End Sub

' Zivi prikaz neto pri bruto-unosu reaguje i na promenu broja/tipa ambalaze.
Private Sub txtKolAmbalaze_Change()
    UpdateUkupnoKg
End Sub

Private Sub cmbTipAmbalaze_Change()
    UpdateUkupnoKg
End Sub

Private Sub UpdateUkupnoKg()
    On Error GoTo EH

    ' Bruto unos: prikazi NETO posle oduzimanja tare (informativno, pre snimanja).
    If OtkupBrutoUnos() Then
        Dim kb As Double, ka As Long
        If Trim$(txtKolicina.value) <> "" Then TryParseDouble txtKolicina.value, kb
        If Trim$(txtKolAmbalaze.value) <> "" Then TryParseLong txtKolAmbalaze.value, ka
        If kb > 0 And ka > 0 And Trim$(cmbTipAmbalaze.value) <> "" Then
            Dim tw As Double: tw = ka * GetTezinaGajbice(cmbTipAmbalaze.value)
            If tw > 0 And tw < kb Then
                lblUkupnoKG.caption = "Neto: " & Format$(kb - tw, "#,##0.00") & _
                    " kg  (bruto " & Format$(kb, "#,##0.00") & " - amb " & _
                    Format$(tw, "#,##0.00") & ")"
                Exit Sub
            End If
        End If
    End If

    If Not chkDveKlase.value Then
        lblUkupnoKG.caption = ""
        Exit Sub
    End If

    Dim kl1 As Double
    Dim kl2 As Double

    If Trim$(txtKolicina.value) <> "" Then
        TryParseDouble txtKolicina.value, kl1
    End If

    If Trim$(txtKolicinaKLII.value) <> "" Then
        TryParseDouble txtKolicinaKLII.value, kl2
    End If

    lblUkupnoKG.caption = "Ukupno: " & Format$(kl1 + kl2, "#,##0.00") & " kg"
    Exit Sub

EH:
    LogErr "frmOtkup.UpdateUkupnoKg"
    lblUkupnoKG.caption = ""
End Sub

' ============================================================
' KASKADIERUNG - VrstaVoca ? SortaVoca
' ============================================================

Private Sub cmbVrstaVoca_Change()
    ' Wenn VrstaVoca gewaehlt wird, SortaVoca-Liste filtern
    cmbSortaVoca.Clear
    If cmbVrstaVoca.value <> "" Then
        FillCmb cmbSortaVoca, _
            GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbVrstaVoca.value, True)
    End If
    AutoFillCenaOtkup
End Sub

Private Sub cmbSortaVoca_Change()
    AutoFillCenaOtkup
End Sub

' Auto-popunjavanje cene (cenovnik) i tipa ambalaze (kultura) po proizvodu.
' Postavlja samo ako postoji vrednost; rucni unos ostaje moguc.
Private Sub AutoFillCenaOtkup()
    On Error Resume Next

    Dim vrsta As String, sorta As String
    vrsta = Trim$(cmbVrstaVoca.value)
    sorta = Trim$(cmbSortaVoca.value)
    If vrsta = "" Then Exit Sub

    Dim cI As Double
    cI = GetVazecaCena(vrsta, sorta, KLASA_I)
    If cI > 0 Then txtCena.value = Format$(cI, "0.######")

    If chkDveKlase.value Then
        Dim cII As Double
        cII = GetVazecaCena(vrsta, sorta, KLASA_II)
        If cII > 0 Then txtCenaKLII.value = Format$(cII, "0.######")
    End If

    ' #6 podrazumevani tip ambalaze iz kulture
    Dim ta As String
    ta = GetKulturaTipAmbalaze(vrsta, sorta)
    If Len(ta) > 0 Then cmbTipAmbalaze.value = ta

    UpdatePaletaInfo
End Sub

' Runtime labela "gajbe do zatvaranja aktivne palete": meri poziciju ispod
' dugmeta Povratak (forme se renderuju samo u Excelu; .frx se ne dira).
Private Sub SetupPaletaInfoField()
    On Error GoTo done
    If Not m_lblPaletaInfo Is Nothing Then Exit Sub

    Set m_lblPaletaInfo = Me.Controls.Add("Forms.Label.1", "lblPaletaInfoRT", True)
    With m_lblPaletaInfo
        .Left = txtKolicina.Left
        .top = btnPovratak.top + btnPovratak.Height + 6
        .width = Me.InsideWidth - txtKolicina.Left - 12
        If .width < 240 Then .width = 240
        .Height = 28
        .WordWrap = True
        .caption = ""
    End With
    On Error Resume Next
    StyleLabel m_lblPaletaInfo, TXT_MUTED, False
    On Error GoTo done
    Exit Sub
done:
    LogErr "frmOtkup.SetupPaletaInfoField"
    Set m_lblPaletaInfo = Nothing
End Sub

' Osvezi info koliko gajbi treba da se zatvori aktivna paleta za trenutno
' izabranu vrstu/sortu. Klasa I uvek; Klasa II kad je chkDveKlase ON.
Private Sub UpdatePaletaInfo()
    On Error GoTo EH
    If m_lblPaletaInfo Is Nothing Then Exit Sub

    Dim vrsta As String, sorta As String
    vrsta = Trim$(cmbVrstaVoca.value)
    sorta = Trim$(cmbSortaVoca.value)
    If Len(vrsta) = 0 Then
        m_lblPaletaInfo.caption = ""
        Exit Sub
    End If

    Dim info As String
    info = GajbeDoZatvaranjaPaleteInfo(vrsta, sorta, KLASA_I)
    If chkDveKlase.value Then
        Dim info2 As String
        info2 = GajbeDoZatvaranjaPaleteInfo(vrsta, sorta, KLASA_II)
        If Len(info2) > 0 Then
            If Len(info) > 0 Then
                info = info & vbCrLf & "II: " & info2
            Else
                info = "II: " & info2
            End If
        End If
    End If

    m_lblPaletaInfo.caption = info
    Exit Sub
EH:
    LogErr "frmOtkup.UpdatePaletaInfo"
End Sub

' Podesavanja: stanje polja prema toggle-ima (parcele / kes isplate). Polja
' ostaju VIDLJIVA, ali su disabled kad je toggle OFF (bez unosa; tab ih
' preskace). Postavlja se u oba smera pa re-otvaranje forme prati config.
Private Sub ApplyOtkupTogglesState()
    On Error Resume Next
    If IsPracenjeParcela() Then EnableCombo cmbParcela Else DisableCombo cmbParcela
    If IsKesIsplate() Then
        EnableField txtNovac
        EnableField txtPrimalac
    Else
        DisableField txtNovac
        DisableField txtPrimalac
    End If
End Sub

Private Sub cmbOtkupnoMesto_Change()
    On Error GoTo EH

    cmbKooperant.Clear
    cmbParcela.Clear

    If cmbOtkupnoMesto.value = "" Then
        ' Operater je obrisao izbor -- pusti aktivnu stanicu (sa bulk push)
        If Len(GetActiveStanica()) > 0 Then
            ShowLockStatus "Sinhronizujem prethodnu stanicu..."
            ReleaseStanicaLock GetActiveStanica()
            HideLockStatus
        End If
        txtBrojDokumenta.value = ""
        Exit Sub
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)
    If stanicaID = "" Then Exit Sub

    ' Promena stanice VAN konteksta selektovane otpremnice (panel): datum i broj
    ' zbirne su nasledjeni sa te otpremnice -> vrati datum na danas i ocisti
    ' zaostalu zbirnu, da svez unos ne nosi stari datum. mPrefilling gard: prefill
    ' sam postavlja stanicu i NE sme da se resetuje (vidi modOtkupBlok).
    If Not OtkupBlok_IsPrefilling() Then
        Dim otpStanica As String: otpStanica = OtkupBlok_ActiveStanica()
        If Len(otpStanica) > 0 And otpStanica <> stanicaID Then
            txtDatum.value = Format$(Date, "d.m.yyyy")
            txtBrojZbirne.value = ""
            ResetProizvodNaDefault
            OtkupBlok_ClearActiveOtp
        End If
    End If

    ' Parse datum (vec treba da je popunjen u txtDatum)
    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        datumDok = Date   ' fallback na danas; korisnik moze promeniti
    End If

    ' Lock acquire (interno: bulk push + release prethodne ako postoji + acquire nove)
    Dim isChanging As Boolean
    isChanging = (Len(GetActiveStanica()) > 0 And GetActiveStanica() <> stanicaID)

    If isChanging Then
        ShowLockStatus "Sinhronizujem prethodnu stanicu i preuzimam novu..."
    Else
        ShowLockStatus "Preuzimam stanicu..."
    End If

    Dim acquired As Boolean
    acquired = AcquireStanicaLock(stanicaID, datumDok)

    HideLockStatus

    If Not acquired Then
        MsgBox "Nije moguce preuzeti stanicu " & stanicaID & Poruka("OTKUP_MSG_POKUSAJ_PONOVO"), _
               vbExclamation, APP_NAME
        cmbOtkupnoMesto.value = ""
        Exit Sub
    End If

    ' #4 toggle: ON -> kooperanti po OM; OFF -> svi kooperanti ("" = svi)
    If KoopFilterByOM() Then
        FillComboKooperantiByStanica cmbKooperant, stanicaID
    Else
        FillComboKooperantiByStanica cmbKooperant, ""
    End If
    RefreshBrojDokumentaSuggestion

    ' MALINA: vozac == par-vozac OM (VozacID == StanicaID) -> auto-izbor, da
    ' otkupac ne mora rucno da bira vozaca. Popunjen vozac na otkupu omogucava i
    ' auto-povezivanje u sledljivosti. Ako par-vozac ne postoji, ostaje prazno.
    If IsMalinaMode() Then SelectComboByDisplayID cmbVozac, stanicaID
    Exit Sub

EH:
    LogErr "frmOtkup.cmbOtkupnoMesto_Change"
    HideLockStatus
    cmbKooperant.Clear
    cmbParcela.Clear
End Sub

Private Sub cmbKooperant_Change()
    On Error GoTo EH

    ' Panel "Otkupni blokovi": osvezi inline "ukupan iznos otk. listova" za izabranog
    ' kooperanta (no-op ako panel nije otvoren). Pre early-exit-a za parcele.
    On Error Resume Next
    OtkupBlok_RefreshKoopTotal
    On Error GoTo EH

    ' Pracenje parcela OFF (Podesavanja) -> parcela polje se preskace.
    If Not IsPracenjeParcela() Then Exit Sub

    cmbParcela.Clear

    If cmbKooperant.ListIndex < 0 Then Exit Sub

    Dim kooperantID As String
    kooperantID = GetComboID(cmbKooperant)

    If kooperantID = "" Then Exit Sub

    Dim parData As Variant
    parData = GetTableData(TBL_PARCELE)

    If IsEmpty(parData) Then Exit Sub

    Dim colKoop As Long
    Dim colID As Long
    Dim colKat As Long
    Dim colKultura As Long
    Dim colPovrsina As Long

    colID = RequireColumnIndex(TBL_PARCELE, COL_PAR_ID, _
                               "frmOtkup.cmbKooperant_Change")
    colKoop = RequireColumnIndex(TBL_PARCELE, COL_PAR_KOOP, _
                                 "frmOtkup.cmbKooperant_Change")
    colKat = RequireColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ, _
                                "frmOtkup.cmbKooperant_Change")
    colKultura = RequireColumnIndex(TBL_PARCELE, COL_PAR_KULTURA, _
                                    "frmOtkup.cmbKooperant_Change")
    colPovrsina = RequireColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA, _
                                     "frmOtkup.cmbKooperant_Change")

    Dim i As Long
    Dim povrsina As Double

    For i = 1 To UBound(parData, 1)
        If CStr(parData(i, colKoop)) = kooperantID Then

            povrsina = 0
            If IsNumeric(parData(i, colPovrsina)) Then povrsina = CDbl(parData(i, colPovrsina))

            cmbParcela.AddItem CStr(parData(i, colKat)) & " | " & _
                               CStr(parData(i, colKultura)) & " | " & _
                               Format$(povrsina, "#,##0.00") & " ha (" & _
                               CStr(parData(i, colID)) & ")"
        End If
    Next i

    Exit Sub

EH:
    LogErr "frmOtkup.cmbKooperant_Change"
    cmbParcela.Clear
End Sub

Private Sub cmbParcela_Change()
    On Error GoTo EH

    If cmbParcela.ListIndex < 0 Then Exit Sub

    Dim parcelaID As String
    parcelaID = ExtractParcelaID(cmbParcela.value)

    If parcelaID = "" Then Exit Sub

    Dim kultura As String
    kultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))

    If kultura = "" Then Exit Sub

    Dim vrsta As String
    vrsta = CStr(LookupValue(TBL_KULTURE, "SortaVoca", kultura, "VrstaVoca"))

    If vrsta <> "" Then
        cmbVrstaVoca.value = vrsta

        On Error Resume Next
        cmbSortaVoca.value = kultura
        On Error GoTo EH
    End If

    Exit Sub

EH:
    LogErr "frmOtkup.cmbParcela_Change"
End Sub

' Helper: ParcelaID aus Display-String extrahieren
' Format: "KatBroj | Kultura | 2.50 ha (PAR001)"
Private Function ExtractParcelaID(ByVal display As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(display, "(")
    p2 = InStrRev(display, ")")
    If p1 > 0 And p2 > p1 Then
        ExtractParcelaID = Mid$(display, p1 + 1, p2 - p1 - 1)
    End If
End Function

Private Sub txtDatum_AfterUpdate()
    On Error GoTo EH

    ' Ako nema aktivne stanice, samo refresh predlog (suggestion zavisi od datuma)
    If Len(GetActiveStanica()) = 0 Then
        RefreshBrojDokumentaSuggestion
        Exit Sub
    End If

    ' Parse novi datum
    Dim newDatum As Date
    If Not TryParseDateValue(txtDatum.value, newDatum) Then
        ' Los format -- operator vidi u polju, ne menjamo lock state
        Exit Sub
    End If

    ' Ako je isti datum, nista
    If GetActiveDatum() = newDatum Then
        RefreshBrojDokumentaSuggestion
        Exit Sub
    End If

    ' Drugaciji datum -- re-acquire (bulk push staro + acquire novo)
    ShowLockStatus "Sinhronizujem prethodni datum..."
    Dim acquired As Boolean
    acquired = AcquireStanicaLock(GetActiveStanica(), newDatum)
    HideLockStatus

    If acquired Then
        RefreshBrojDokumentaSuggestion
    End If
    Exit Sub

EH:
    LogErr "frmOtkup.txtDatum_AfterUpdate"
    HideLockStatus
End Sub

Private Sub RefreshBrojDokumentaSuggestion(Optional ByVal checkRemote As Boolean = True)
    On Error GoTo EH

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)
    If Len(stanicaID) = 0 Then Exit Sub

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then Exit Sub

    Dim suggested As String
    suggested = SuggestNextBroj(KIND_OTK, stanicaID, datumDok, checkRemote)

    If Len(suggested) > 0 Then
        txtBrojDokumenta.value = suggested
    End If
    Exit Sub

EH:
    LogErr "frmOtkup.RefreshBrojDokumentaSuggestion"
End Sub

' Vrati levu formu na "danas" kontekst kad se napusti otpremnica iz panela
' (Sakrij blokove): datum -> danas, ocisti zaostali broj zbirne, pa preracunaj
' predlog broja otkupnog lista za tekuci kontekst. Public: zove modOtkupBlok.
Public Sub ResetDatumKontekst()
    On Error Resume Next
    txtDatum.value = Format$(Date, "d.m.yyyy")
    txtBrojZbirne.value = ""
    ResetProizvodNaDefault
    RefreshBrojDokumentaSuggestion False
End Sub

' Vrati vrstu/sortu voca na podrazumevani proizvod (kao pri otvaranju forme) kad
' se napusti otpremnica iz panela -- da svez unos ne nosi vrstu/sortu otpremnice.
' Cisti pa primeni default (CFG_DEFAULT_VRSTA/SORTA); bez podesenog default-a
' ostaju prazni. ApplyDefaultProizvod okida auto-cenu/tip ambalaze (cmbVrsta_Change).
Private Sub ResetProizvodNaDefault()
    On Error Resume Next
    cmbVrstaVoca.value = ""
    cmbSortaVoca.value = ""
    ApplyDefaultProizvod cmbVrstaVoca, cmbSortaVoca
End Sub

' ============================================================
' OTKUP
' ============================================================

Private Sub btnUnos_Click()
    On Error GoTo EH

    ButtonActive btnUnos

    If cmbOtkupnoMesto.value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        cmbOtkupnoMesto.SetFocus
        Exit Sub
    End If

    If Trim$(cmbKooperant.value) = "" Then
        MsgBox "Izaberite ili unesite kooperanta!", vbExclamation, APP_NAME
        cmbKooperant.SetFocus
        Exit Sub
    End If

    If cmbVrstaVoca.value = "" Then
        MsgBox "Izaberite vrstu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
        cmbVrstaVoca.SetFocus
        Exit Sub
    End If

    If IsValidacijaUnosa() And cmbSortaVoca.value = "" Then
        MsgBox "Izaberite sortu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
        cmbSortaVoca.SetFocus
        Exit Sub
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    ' Klasa I je opciona SAMO kad je ukljucena Klasa II i Klasa I ostavljena prazna
    ' (unosi se samo II klasa). Tada Kolicina I i Ambalaza I MORAJU biti prazne.
    Dim kolicinaI As Double
    Dim cenaI As Double
    Dim hasKlasaI As Boolean: hasKlasaI = (Trim$(txtKolicina.value) <> "")

    If hasKlasaI Then
        If Not TryParseDouble(txtKolicina.value, kolicinaI) Or kolicinaI <= 0 Then
            MsgBox "Unesite ispravnu koli" & ChrW(269) & "inu!", vbExclamation, APP_NAME
            txtKolicina.SetFocus
            Exit Sub
        End If
        If Not TryParseDouble(txtCena.value, cenaI) Or cenaI <= 0 Then
            MsgBox "Unesite ispravnu cenu!", vbExclamation, APP_NAME
            txtCena.SetFocus
            Exit Sub
        End If
    Else
        If Not chkDveKlase.value Then
            MsgBox "Unesite ispravnu koli" & ChrW(269) & "inu!", vbExclamation, APP_NAME
            txtKolicina.SetFocus
            Exit Sub
        End If
        If Trim$(txtKolAmbalaze.value) <> "" Then
            MsgBox Poruka("DOK_MSG_UNOSI_SAMO_KLASA"), _
                   vbExclamation, APP_NAME
            txtKolAmbalaze.SetFocus
            Exit Sub
        End If
    End If

    Dim kolicinaII As Double
    Dim cenaII As Double

    If chkDveKlase.value Then
        If Not TryParseDouble(txtKolicinaKLII.value, kolicinaII) Or kolicinaII <= 0 Then
            MsgBox "Unesite koli" & ChrW(269) & "inu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKLII.SetFocus
            Exit Sub
        End If

        If Not TryParseDouble(txtCenaKLII.value, cenaII) Or cenaII <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKLII.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If Trim$(txtKolAmbalaze.value) <> "" Then
        If Not TryParseLong(txtKolAmbalaze.value, kolAmb) Then
            MsgBox Poruka("DOK_MSG_UNESITE_ISPRAVNU_KOLICINU"), vbExclamation, APP_NAME
            txtKolAmbalaze.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmbII As Long
    If Not m_txtKolAmbalazeII Is Nothing Then
        If chkDveKlase.value And Trim$(m_txtKolAmbalazeII.value) <> "" Then
            If Not TryParseLong(m_txtKolAmbalazeII.value, kolAmbII) Then
                MsgBox Poruka("DOK_MSG_UNESITE_ISPRAVNU_KOLICINU_2"), vbExclamation, APP_NAME
                m_txtKolAmbalazeII.SetFocus
                Exit Sub
            End If
        End If
    End If

    Dim kolAmbIzdata As Long
    If Not m_txtAmbIzdata Is Nothing Then
        If Trim$(m_txtAmbIzdata.value) <> "" Then
            If Not TryParseLong(m_txtAmbIzdata.value, kolAmbIzdata) Then
                MsgBox Poruka("OTKUP_MSG_UNESITE_ISPRAVNU_KOLICINU"), vbExclamation, APP_NAME
                m_txtAmbIzdata.SetFocus
                Exit Sub
            End If
        End If
    End If

    If kolAmb > 0 And cmbTipAmbalaze.value = "" Then
        MsgBox Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"), vbExclamation, APP_NAME
        cmbTipAmbalaze.SetFocus
        Exit Sub
    End If

    If kolAmbII > 0 And cmbTipAmbalaze.value = "" Then
        MsgBox Poruka("OTKUP_MSG_IZABERITE_TIP_AMBALAZE"), vbExclamation, APP_NAME
        cmbTipAmbalaze.SetFocus
        Exit Sub
    End If

    If kolAmbIzdata > 0 And cmbTipAmbalaze.value = "" Then
        MsgBox Poruka("OTKUP_MSG_IZABERITE_TIP_AMBALAZE_2"), vbExclamation, APP_NAME
        cmbTipAmbalaze.SetFocus
        Exit Sub
    End If

    ' Broj gajbi (ambalaza) je OBAVEZAN za svaku unetu klasu. U bruto rezimu je
    ' dodatno kriticno (inace se bruto ne pretvara u neto -> tezina gajbi bi se
    ' platila kao voce).
    If IsValidacijaUnosa() Then
        If kolicinaI > 0 And kolAmb <= 0 Then
            MsgBox "Unesite broj gajbi za I klasu!", vbExclamation, APP_NAME
            txtKolAmbalaze.SetFocus
            Exit Sub
        End If
        If chkDveKlase.value And kolicinaII > 0 And kolAmbII <= 0 Then
            MsgBox "Unesite broj gajbi za II klasu!", vbExclamation, APP_NAME
            If Not m_txtKolAmbalazeII Is Nothing Then m_txtKolAmbalazeII.SetFocus
            Exit Sub
        End If
    Else
        ' Minimalna validacija (kao pre utegnute izmene): broj gajbi obavezan SAMO
        ' u bruto rezimu (bez toga se bruto ne pretvara u neto).
        If OtkupBrutoUnos() And kolicinaI > 0 And kolAmb <= 0 Then
            MsgBox Poruka("OTKUP_MSG_BRUTO_REZIM_UNESITE") & _
                   "(bez toga se bruto ne pretvara u neto).", vbExclamation, APP_NAME
            txtKolAmbalaze.SetFocus
            Exit Sub
        End If
        If chkDveKlase.value And OtkupBrutoUnos() And kolicinaII > 0 And kolAmbII <= 0 Then
            MsgBox Poruka("OTKUP_MSG_BRUTO_REZIM_UNESITE_2") & _
                   "(bez toga se bruto ne pretvara u neto).", vbExclamation, APP_NAME
            If Not m_txtKolAmbalazeII Is Nothing Then m_txtKolAmbalazeII.SetFocus
            Exit Sub
        End If
    End If

    ' --- BRUTO unos (toggle OTKUP_BRUTO_UNOS): kupac unosi bruto (voce + ambalaza).
    ' Oduzmi taru (kolAmb * tezina gajbice) -> u Kolicina ide NETO, bruto se zamrzava
    ' u BrutoKg. Tara se vezuje za Klasu I (kolAmb se i inace odnosi na Klasu I). ---
    Dim brutoKgI As Double
    If OtkupBrutoUnos() And kolAmb > 0 Then
        Dim taraKg As Double
        taraKg = kolAmb * GetTezinaGajbice(cmbTipAmbalaze.value)
        If taraKg <= 0 Then
            MsgBox Poruka("DOK_MSG_TIP_AMBALAZE") & cmbTipAmbalaze.value & Poruka("DOK_MSG_NEMA_UNETU_TEZINU") & _
                   Poruka("DOK_MSG_MATICNI_PODACI_TIP") & vbCrLf & _
                   Poruka("DOK_MSG_BRUTO_MOZE_PRETVORITI"), vbExclamation, APP_NAME
            cmbTipAmbalaze.SetFocus
            Exit Sub
        End If
        If taraKg >= kolicinaI Then
            MsgBox Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(taraKg, "#,##0.00") & " kg) je veca ili " & _
                   Poruka("DOK_MSG_JEDNAKA_BRUTO_TEZINI") & Format$(kolicinaI, "#,##0.00") & " kg)." & vbCrLf & _
                   Poruka("DOK_MSG_PROVERITE_BROJ_KOMADA"), vbExclamation, APP_NAME
            txtKolicina.SetFocus
            Exit Sub
        End If
        brutoKgI = kolicinaI             ' zamrzni uneti bruto
        kolicinaI = kolicinaI - taraKg   ' u Kolicina ide neto
    End If

    ' BRUTO unos za Klasu II (zasebne gajbe -> zasebna tara). Isto kao Klasa I.
    Dim brutoKgII As Double
    If chkDveKlase.value And OtkupBrutoUnos() And kolAmbII > 0 Then
        Dim taraKgII As Double
        taraKgII = kolAmbII * GetTezinaGajbice(cmbTipAmbalaze.value)
        If taraKgII <= 0 Then
            MsgBox Poruka("DOK_MSG_TIP_AMBALAZE") & cmbTipAmbalaze.value & Poruka("DOK_MSG_NEMA_UNETU_TEZINU") & _
                   Poruka("DOK_MSG_MATICNI_PODACI_TIP") & vbCrLf & _
                   Poruka("DOK_MSG_BRUTO_KLASA_MOZE"), vbExclamation, APP_NAME
            cmbTipAmbalaze.SetFocus
            Exit Sub
        End If
        If taraKgII >= kolicinaII Then
            MsgBox Poruka("DOK_MSG_TEZINA_AMBALAZE_KLASE") & Format$(taraKgII, "#,##0.00") & " kg) je veca ili " & _
                   Poruka("DOK_MSG_JEDNAKA_BRUTO_TEZINI") & Format$(kolicinaII, "#,##0.00") & " kg)." & vbCrLf & _
                   Poruka("DOK_MSG_PROVERITE_BROJ_KOMADA"), vbExclamation, APP_NAME
            If Not m_txtKolAmbalazeII Is Nothing Then m_txtKolAmbalazeII.SetFocus
            Exit Sub
        End If
        brutoKgII = kolicinaII             ' zamrzni uneti bruto (II)
        kolicinaII = kolicinaII - taraKgII ' u Kolicina (II) ide neto
    End If

    Dim novac As Double
    If Trim$(txtNovac.value) <> "" Then
        If Not TryParseDouble(txtNovac.value, novac) Or novac < 0 Then
            MsgBox "Unesite ispravan iznos novca!", vbExclamation, APP_NAME
            txtNovac.SetFocus
            Exit Sub
        End If
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        MsgBox "Nije prona" & ChrW(273) & "en ID otkupnog mesta!", vbExclamation, APP_NAME
        cmbOtkupnoMesto.SetFocus
        Exit Sub
    End If

    Dim kooperantID As String
    kooperantID = ResolveKooperantByName(cmbKooperant, stanicaID)

    If kooperantID = "" Then
        MsgBox "Nije prona" & ChrW(273) & "en ID kooperanta!", vbExclamation, APP_NAME
        cmbKooperant.SetFocus
        Exit Sub
    End If

    Dim vozacID As String
    If cmbVozac.value <> "" Then
        vozacID = ExtractIDFromDisplay(cmbVozac.value)
    End If

    If IsValidacijaUnosa() And Trim$(txtBrojDokumenta.value) = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        txtBrojDokumenta.SetFocus
        Exit Sub
    End If

    If Trim$(txtBrojDokumenta.value) <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_OTKUP, COL_OTK_BR_DOK, _
                                Trim$(txtBrojDokumenta.value), COL_OTK_DATUM)

        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim parcelaID As String

    If cmbParcela.ListIndex >= 0 Then
        parcelaID = ExtractParcelaID(cmbParcela.value)

        If parcelaID <> "" Then
            Dim parKultura As String
            parKultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))

            If parKultura <> "" Then
                Dim selectedKultura As String
                selectedKultura = Trim$(cmbVrstaVoca.value)

                If selectedKultura <> "" Then
                    If StrComp(parKultura, selectedKultura, vbTextCompare) <> 0 Then
                        Dim ans As VbMsgBoxResult
                        ans = MsgBox("Kultura parcele (" & parKultura & _
                                     ") ne odgovara izabranoj sorti/kulturi (" & _
                                     selectedKultura & ")!" & vbCrLf & vbCrLf & _
                                     Poruka("OTKUP_MSG_ZELITE_IPAK_NASTAVITE"), _
                                     vbExclamation + vbYesNo, APP_NAME)
                        If ans = vbNo Then Exit Sub
                    End If
                End If
            End If
        End If
    End If

    ' Panel "Otkupni blokovi": upozorenje na prekoracenje preostale kolicine otpremnice
    If Not OtkupBlok_ConfirmUnos() Then Exit Sub

    Dim result As String

    result = SaveOtkupMulti_TX( _
        datum:=datumDok, _
        kooperantID:=kooperantID, _
        stanicaID:=stanicaID, _
        vrstaVoca:=cmbVrstaVoca.value, _
        sortaVoca:=cmbSortaVoca.value, _
        kolicinaI:=kolicinaI, _
        cenaI:=cenaI, _
        tipAmb:=cmbTipAmbalaze.value, _
        kolAmb:=kolAmb, _
        vozacID:=vozacID, _
        brDok:=Trim$(txtBrojDokumenta.value), _
        novac:=novac, _
        primalac:=txtPrimalac.value, _
        parcelaID:=parcelaID, _
        brojZbirne:=Trim$(txtBrojZbirne.value), _
        hasKlasaII:=chkDveKlase.value, _
        kolicinaII:=kolicinaII, _
        cenaII:=cenaII, _
        kolAmbIzdata:=kolAmbIzdata, _
        brutoKgI:=brutoKgI, _
        kolAmbII:=kolAmbII, _
        brutoKgII:=brutoKgII)

    If result = "" Then
        MsgBox Poruka("OTKUP_MSG_GRESKA_PRI_CUVANJU"), vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Otkup sa" & ChrW(269) & "uvan: " & result, vbInformation, APP_NAME

    ' Otkupni list: PDF (po CFG_OTKUP_PRINT_MODE; podrazumevano PDF + otvori) za
    ' upravo sacuvani blok. Best-effort: greska ne sme da obori potvrdu snimanja.
    On Error Resume Next
    OutputOtkupniList result
    On Error GoTo 0

    ' #3 Hladnjaca auto-lanac: ako je OM hladnjaca i toggle ON ->
    ' auto otpremnica+zbirna+prijemnica iz otkupa. Best-effort (pre ClearOtkupFields,
    ' dok combo-i jos drze vrednosti).
    On Error Resume Next
    Dim hlWarn As String
    hlWarn = AutoChainHladnjaca(datumDok, stanicaID, cmbVrstaVoca.value, cmbSortaVoca.value, _
                       vozacID, cmbTipAmbalaze.value, kolAmb, kolicinaI, cenaI, _
                       chkDveKlase.value, kolicinaII, cenaII, Trim$(txtBrojDokumenta.value), _
                       result, brutoKgI, kolAmbII, brutoKgII)
    On Error GoTo 0
    If Len(hlWarn) > 0 Then MsgBox hlWarn, vbExclamation, APP_NAME

    ClearOtkupFields

    ' Panel "Otkupni blokovi": vezi upravo sacuvani red za izabranu otpremnicu
    On Error Resume Next
    OtkupBlok_AfterUnos result
    On Error GoTo 0

    Exit Sub

EH:
    LogErr "frmOtkup.btnUnos"
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & Err.description, vbCritical, APP_NAME
End Sub

Private Sub ClearOtkupFields()
    txtBrojDokumenta.value = ""
    txtKolicina.value = ""
    txtCena.value = ""
    txtKolAmbalaze.value = ""
    If Not m_txtAmbIzdata Is Nothing Then m_txtAmbIzdata.value = ""
    txtNovac.value = "0"
    txtPrimalac.value = ""
    cmbKooperant.value = ""
    cmbParcela.Clear

    chkDveKlase.value = False
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    ShowKolAmbalazeII False
    lblUkupnoKG.caption = ""

    ' Auto-cena/tip se gube posle prvog unosa (txtCena ocisceno, a vrsta/sorta
    ' combo nije menjan pa _Change ne okida). Vrati ih za jos izabran proizvod.
    ' (AutoFillCenaOtkup osvezava i info o aktivnoj paleti.)
    On Error Resume Next
    AutoFillCenaOtkup
    On Error GoTo 0

    ' Kooperant je ociscen -> fokus na njega (sledeci unos = novi kooperant).
    cmbKooperant.SetFocus

    ' Lokalni predlog (bez Google) -- just-saved red je vec u tblOtkup-u
    RefreshBrojDokumentaSuggestion False
End Sub

Private Sub btnStornoOtkup_Click()
    ButtonActive btnStornoOtkup
End Sub

' ============================================================
' NAVIGATION
' ============================================================

Private Sub btnPovratak_Click()
    On Error GoTo EH
    
    If Len(GetActiveStanica()) > 0 Then
        ShowLockStatus "Sinhronizujem unos pre zatvaranja..."
        ReleaseStanicaLock GetActiveStanica()
        HideLockStatus
    End If

    ButtonActive btnPovratak

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmOtkup.btnPovratak_Click"
    HideLockStatus
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    MouseWheel_Detach

    ' Release lock pre zatvaranja (vbFormControlMenu = X klik, ostalo = Code/Excel close)
    If Len(GetActiveStanica()) > 0 Then
        ShowLockStatus "Sinhronizujem unos pre zatvaranja..."
        ReleaseStanicaLock GetActiveStanica()
        HideLockStatus
    End If

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    ' Oslobodi dinamicki panel (WithEvents wrappers + runtime kontrole) PRE unload-a.
    ' Bez ovoga Excel pri rusenju forme cisti ~35 event-sink objekata -> sporo
    ' zatvaranje. OtkupBlok_Release je idempotentan; AttachOtkupBlokPanel ga
    ' ponovo izgradi pri sledecem otvaranju sekcije.
    OtkupBlok_Release

    On Error GoTo 0
End Sub

Private Sub ShowLockStatus(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
    Application.Cursor = xlWait
    DoEvents
End Sub

Private Sub HideLockStatus()
    On Error Resume Next
    Application.StatusBar = False
    Application.Cursor = xlDefault
End Sub

