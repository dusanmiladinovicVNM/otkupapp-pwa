Attribute VB_Name = "modMaticniLookups"
'Attribute VB_Name = "modMaticniLookups"
Option Explicit

' ============================================================
' modMaticniLookups – jedinstveni (data-driven) meni "Maticni podaci"
'
' Ceo meni frmMaticniPodaci se gradi iz JEDNE registracije sekcija
' (MaticniSekcije). Za svaku sekciju se dinamicki kreira dugme
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
    MaticniSekcije = Array( _
        Array("Kooperanti", "Kooperanti"), _
        Array("Stanice", "Stanice"), _
        Array("Kupci", "Kupci"), _
        Array("Vozaci", "Vozaci"), _
        Array("Artikli", "Artikli"), _
        Array("Parcele", "Parcele"), _
        Array("Kulture", "Kulture"), _
        Array("Ambalaza", "TipAmbalaze"), _
        Array("Palete", "TipPalete"), _
        Array("Cenovnik", "Cenovnik"), _
        Array("Kutije", "Kutije"), _
        Array("Kese", "Kese"), _
        Array("Vrsta got. proizvoda", "VrstaGP"), _
        Array(Poruka("MATICNI_MSG_PODESAVANJA"), "Podesavanja"))
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

    Dim X As Single, w As Single, top0 As Single, bandBottom As Single
    X = tmpl.Left
    w = tmpl.width
    top0 = tmpl.top
    bandBottom = exitBtn.top      ' Exit ostaje na svom mestu; dugmad popunjavaju iznad

    Dim secs As Variant
    secs = MaticniSekcije()

    Dim n As Long
    n = UBound(secs) - LBound(secs) + 1
    If n <= 0 Then Exit Sub

    ' Spakuj n dugmadi u isti vertikalni opseg koji su zauzimala staticna
    ' dugmad (top0 .. Exit.Top) — bez resize-a forme.
    Const SPACING As Single = 3

    ' Citljiv pitch = visina sablonskog dugmeta + razmak. Ako n dugmadi ne
    ' staju u postojeci opseg (top0 .. Exit), povecaj formu i spusti Exit
    ' nanize, umesto da se dugmad gnjece -> sve sekcije ostaju citljive.
    Dim pitch As Single
    pitch = tmpl.Height + SPACING

    If top0 + n * pitch > bandBottom Then
        Dim grow As Single
        grow = (top0 + n * pitch) - bandBottom
        exitBtn.top = exitBtn.top + grow
        frm.Height = frm.Height + grow
        bandBottom = exitBtn.top
    Else
        pitch = (bandBottom - top0) / n
    End If

    Dim btnH As Single
    btnH = pitch - SPACING
    If btnH < 14 Then btnH = 14

    Dim i As Long
    For i = 0 To n - 1
        Dim sec As Variant
        sec = secs(LBound(secs) + i)

        Dim nm As String
        nm = "btnMD_" & CStr(sec(1))

        On Error Resume Next
        frm.Controls.Remove nm      ' u slucaju re-init
        On Error GoTo EH

        Dim c As MSForms.CommandButton
        Set c = frm.Controls.Add("Forms.CommandButton.1", nm, True)
        c.Left = X
        c.width = w
        c.top = top0 + i * pitch
        c.Height = btnH
        StyleMenuButton c, CStr(sec(0))

        Dim wrp As clsLookupMenuBtn
        Set wrp = New clsLookupMenuBtn
        wrp.sekcijaTag = CStr(sec(1))
        wrp.sekcijaCaption = CStr(sec(0))
        Set wrp.btn = c

        mWrappers.Add wrp
        mBtns.Add c
    Next i

    ' Uspeh — sakrij staticna dugmad (dinamicka su preko njih).
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

' Klik na sekciju — otvori frmStammdaten preko forme (ona drzi flag
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

