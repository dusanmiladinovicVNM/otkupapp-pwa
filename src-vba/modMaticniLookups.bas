Attribute VB_Name = "modMaticniLookups"
'Attribute VB_Name = "modMaticniLookups"
Option Explicit

' ============================================================
' modMaticniLookups - jedinstveni (data-driven) meni "Maticni podaci"
'
' Ceo meni frmMaticniPodaci se gradi iz JEDNE registracije sekcija
' (MaticniSekcijeGrupisano). Za svaku sekciju se dinamicki kreira dugme
' (Controls.Add) pa se njegov klik hvata preko clsLookupMenuBtn
' (WithEvents). Tako:
'   - frmMaticniPodaci.frx se NE dira,
'   - sve sekcije (postojece + nove) idu kroz isti mehanizam,
'   - dodavanje nove sekcije = jedan red u MaticniSekcije + Case u
'     frmStammdaten.
'
' Postojeca staticna dugmad na formi se sakrivaju (ostaju u .frx kao
' fallback ako dinamicka izgradnja ne uspe).
'
' Otvaranje sekcije ide kroz frmMaticniPodaci.OpenSekcija (koji vec
' ispravno upravlja m_IsOpeningChild flagom, da se meni ne zatvori).
' ============================================================

Private mWrappers As Collection   ' clsLookupMenuBtn instance (drzi WithEvents zivim)
Private mBtns As Collection       ' MSForms.CommandButton kontrole (za reset/hover)

Private Const STATIC_BTNS As String = _
    "btnKooperanti;btnStanice;btnKupci;btnVozaci;btnArtikli;btnParcele"

' Registracija svih sekcija: Array(Naziv u meniju, Tag za frmStammdaten).
' Redosled ovde = redosled u meniju.
Public Function MaticniSekcije() As Variant
    ' Ravna lista (zadrzana radi kompatibilnosti) izvedena iz grupisane
    ' registracije MaticniSekcijeGrupisano - jedan izvor istine.
    Dim groups As Variant
    groups = MaticniSekcijeGrupisano()

    Dim out As Collection
    Set out = New Collection

    Dim gi As Long, ii As Long
    Dim grp As Variant, items As Variant
    For gi = LBound(groups) To UBound(groups)
        grp = groups(gi)
        items = grp(1)
        For ii = LBound(items) To UBound(items)
            out.Add items(ii)
        Next ii
    Next gi

    Dim a() As Variant, k As Long
    ReDim a(0 To out.count - 1)
    For k = 1 To out.count
        a(k - 1) = out(k)
    Next k
    MaticniSekcije = a
End Function

' Grupisana registracija sekcija menija "Maticni podaci".
' Vraca Array(Array(GrupaNaziv, Array(Array(Caption, Tag), ...)), ...).
' Redosled grupa i stavki = redosled u meniju. Pakovanje (ambalaza, palete,
' kutije, kese) je svoja grupa - vizuelno podredjena osnovnim sifarnicima,
' a ne ravnopravno sa Kooperanti/Kupci/Stanice.
' Dodavanje sekcije = jedan red u odgovarajucoj grupi + Case u frmStammdaten.
Public Function MaticniSekcijeGrupisano() As Variant
    MaticniSekcijeGrupisano = Array( _
        Array(ChrW(352) & "ifarnici", Array( _
            Array("Kooperanti", "Kooperanti"), _
            Array("Stanice", "Stanice"), _
            Array("Kupci", "Kupci"), _
            Array("Vozaci", "Vozaci"), _
            Array("Parcele", "Parcele"))), _
        Array("Proizvodi i cene", Array( _
            Array("Artikli", "Artikli"), _
            Array("Kulture", "Kulture"), _
            Array("Cenovnik", "Cenovnik"), _
            Array("Vrsta got. proizvoda", "VrstaGP"))), _
        Array("Ambala" & ChrW(382) & "a i pakovanje", Array( _
            Array("Ambala" & ChrW(382) & "a", "TipAmbalaze"), _
            Array("Palete", "TipPalete"), _
            Array("Kutije", "Kutije"), _
            Array("Kese", "Kese"))), _
        Array("Sistem", Array( _
            Array(Poruka("MATICNI_MSG_PODESAVANJA"), "Pode" & ChrW(353) & "avanja"))))
End Function

' Gradi ceo meni na prosledjenoj formi (frmMaticniPodaci).
' Poziva se iz frmMaticniPodaci.UserForm_Initialize.
Public Sub AttachMaticniMenu(ByVal frm As Object)
    On Error GoTo EH

    Set mWrappers = New Collection
    Set mBtns = New Collection

    ' Geometrija se cita sa postojeceg dugmeta (robusno, bez magic broja).
    Dim tmpl As MSForms.CommandButton
    Set tmpl = frm.Controls("btnKooperanti")

    Dim exitBtn As MSForms.CommandButton
    Set exitBtn = frm.Controls("btnExit")

    Dim X As Single, w As Single, top0 As Single
    X = tmpl.Left
    w = tmpl.width
    top0 = tmpl.top

    Dim btnH As Single
    btnH = tmpl.Height
    If btnH < 16 Then btnH = 16

    Const ITEMGAP As Single = 3
    Const HDRH As Single = 15
    Const HDRGAP As Single = 2
    Const GROUPGAP As Single = 9

    Dim groups As Variant
    groups = MaticniSekcijeGrupisano()

    Dim Y As Single
    Y = top0

    Dim gi As Long, ii As Long, hdrIdx As Long
    Dim grp As Variant, items As Variant, it As Variant
    Dim grpName As String, cap As String, tg As String
    Dim hnm As String, nm As String
    Dim hl As MSForms.label
    Dim c As MSForms.CommandButton
    Dim wrp As clsLookupMenuBtn
    hdrIdx = 0

    For gi = LBound(groups) To UBound(groups)
        grp = groups(gi)
        grpName = CStr(grp(0))
        items = grp(1)

        If gi > LBound(groups) Then Y = Y + GROUPGAP

        ' Sekcijski header - labela (nije klikabilna); vizuelno grupise sekcije.
        hdrIdx = hdrIdx + 1
        hnm = "lblMDgrp_" & CStr(hdrIdx)
        On Error Resume Next
        frm.Controls.Remove hnm
        On Error GoTo EH

        ' Duzi naslovi se prelamaju u 2 reda; daj im visu kutiju (sa WordWrap)
        ' da se ceo tekst vidi. Kratki naslovi ostaju 1 red.
        Dim hdrThisH As Single
        hdrThisH = IIf(Len(grpName) > 12, HDRH + 13, HDRH)
        Set hl = frm.Controls.Add("Forms.Label.1", hnm, True)
        hl.Left = X
        hl.width = w
        hl.top = Y
        hl.Height = hdrThisH
        hl.WordWrap = True
        hl.caption = UCase$(grpName)
        hl.BackStyle = fmBackStyleTransparent
        hl.TextAlign = fmTextAlignLeft
        hl.Font.name = APP_FONT_BOLD
        hl.Font.Size = FONT_SIZE_SMALL
        hl.Font.Bold = True
        hl.ForeColor = TXT_MUTED()
        Y = Y + hdrThisH + HDRGAP

        For ii = LBound(items) To UBound(items)
            it = items(ii)
            cap = CStr(it(0))
            tg = CStr(it(1))

            nm = "btnMD_" & tg
            On Error Resume Next
            frm.Controls.Remove nm
            On Error GoTo EH

            Set c = frm.Controls.Add("Forms.CommandButton.1", nm, True)
            c.Left = X
            c.width = w
            c.top = Y
            c.Height = btnH
            StyleMenuButton c, cap

            Set wrp = New clsLookupMenuBtn
            wrp.sekcijaTag = tg
            wrp.sekcijaCaption = cap
            Set wrp.btn = c

            mWrappers.Add wrp
            mBtns.Add c

            Y = Y + btnH + ITEMGAP
        Next ii
    Next gi

    ' Exit ispod sadrzaja; rastegni formu da sve sekcije ostanu citljive.
    Const BOTTOMPAD As Single = 10
    exitBtn.top = Y + BOTTOMPAD
    Dim desiredInside As Single
    desiredInside = exitBtn.top + exitBtn.Height + BOTTOMPAD
    If frm.InsideHeight < desiredInside Then
        frm.Height = frm.Height + (desiredInside - frm.InsideHeight)
    End If

    ' Uspeh - sakrij staticna dugmad (dinamicka su preko njih).
    HideStaticButtons frm
    Exit Sub

EH:
    LogErr "modMaticniLookups.AttachMaticniMenu"
    ' Fallback: ostavi staticna dugmad vidljiva (postojeca funkcionalnost).
End Sub

Private Sub HideStaticButtons(ByVal frm As Object)
    Dim names() As String
    names = Split(STATIC_BTNS, ";")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        On Error Resume Next
        frm.Controls(names(i)).Visible = False
        On Error GoTo 0
    Next i
End Sub

' Reset stila svih dinamickih dugmadi (za hover efekat).
Public Sub MaticniMenu_ResetAll()
    On Error Resume Next
    If mBtns Is Nothing Then Exit Sub
    Dim c As MSForms.CommandButton
    For Each c In mBtns
        StyleMenuButton c
    Next c
End Sub

Public Sub MaticniMenu_OnHover(ByVal b As Object)
    On Error Resume Next
    MaticniMenu_ResetAll
    ButtonHover b
End Sub

' Klik na sekciju -- otvori frmStammdaten preko forme (ona drzi flag
' m_IsOpeningChild, pa se meni ne zatvori usput).
Public Sub MaticniMenu_OnClick(ByVal sekTag As String, ByVal sekCaption As String)
    On Error GoTo EH

    ButtonActiveByTag sekTag
    frmMaticniPodaci.OpenSekcija sekTag, sekCaption
    Exit Sub

EH:
    LogErr "modMaticniLookups.MaticniMenu_OnClick"
End Sub

Private Sub ButtonActiveByTag(ByVal sekTag As String)
    On Error Resume Next
    If mBtns Is Nothing Then Exit Sub
    Dim nm As String
    nm = "btnMD_" & sekTag
    Dim c As MSForms.CommandButton
    For Each c In mBtns
        If StrComp(c.name, nm, vbTextCompare) = 0 Then
            ButtonActive c
            Exit For
        End If
    Next c
End Sub


' Otpusti module-level reference (clsLookupMenuBtn WithEvents + dugmad) pre
' self-update importa (zivi event-sink obara CodeModule edit). Idempotentno.
Public Sub MaticniMenu_Release()
    On Error Resume Next
    Set mWrappers = Nothing
    Set mBtns = Nothing
End Sub
