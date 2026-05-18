Attribute VB_Name = "modTheme"
Option Explicit

' ============================================================
' modTheme – OtkupAPP shared theme
' Dark UI theme + shared styling helpers for all UserForms
' ============================================================

' =========================
' FONTOVI
' =========================
Public Const APP_FONT As String = "Segoe UI"
Public Const APP_FONT_BOLD As String = "Segoe UI Semibold"
Public Const FONT_SIZE_NORMAL As Single = 10
Public Const FONT_SIZE_SMALL As Single = 9
Public Const FONT_SIZE_HEADER As Single = 11
Public Const FONT_SIZE_TITLE As Single = 13

' =========================
' PALETA
' =========================

Public Function BG_MAIN() As Long: BG_MAIN = RGB(18, 20, 18): End Function
Public Function BG_TOP() As Long: BG_TOP = RGB(24, 30, 24): End Function
Public Function BG_PANEL() As Long: BG_PANEL = RGB(28, 36, 30): End Function

Public Function BTN_BG() As Long: BTN_BG = RGB(46, 74, 48): End Function
Public Function BTN_HOVER() As Long: BTN_HOVER = RGB(66, 104, 68): End Function
Public Function BTN_ACTIVE() As Long: BTN_ACTIVE = RGB(212, 180, 76): End Function

Public Function TXT_LIGHT() As Long: TXT_LIGHT = RGB(244, 242, 232): End Function
Public Function TXT_MUTED() As Long: TXT_MUTED = RGB(182, 188, 172): End Function
Public Function TXT_ALERT() As Long: TXT_ALERT = RGB(230, 95, 95): End Function
Public Function BORDER_SOFT() As Long: BORDER_SOFT = RGB(88, 108, 86): End Function

Public Function INPUT_BG() As Long: INPUT_BG = RGB(36, 52, 38): End Function
Public Function INPUT_DISABLED_BG() As Long: INPUT_DISABLED_BG = RGB(28, 34, 29): End Function
Public Function INPUT_BORDER() As Long: INPUT_BORDER = RGB(96, 122, 92): End Function

Public Function CLR_SUCCESS() As Long: CLR_SUCCESS = RGB(126, 204, 96): End Function
Public Function CLR_WARNING() As Long: CLR_WARNING = RGB(240, 204, 92): End Function
Public Function CLR_ERROR() As Long: CLR_ERROR = RGB(230, 95, 95): End Function


' ============================================================
' GLAVNA THEME PROCEDURA
' ============================================================
Public Sub ApplyTheme(ByVal frm As Object, Optional ByVal formBackColor As Long = -1)
    On Error Resume Next

    If formBackColor = -1 Then
        frm.BackColor = BG_MAIN
    Else
        frm.BackColor = formBackColor
    End If

    SetFont frm, APP_FONT, FONT_SIZE_NORMAL
    StyleControls frm

    On Error GoTo 0
End Sub

' alias ako si vec koristio ovo ime
Public Sub ApplyFormTheme(ByVal frm As Object, Optional ByVal formBackColor As Long = -1)
    ApplyTheme frm, formBackColor
End Sub

' ============================================================
' REKURZIVNO STILIZOVANJE
' ============================================================
Private Sub StyleControls(ByVal parent As Object)
    Dim c As Object

    For Each c In parent.Controls
        On Error Resume Next

        If Not IsDirectChild(c, parent) Then GoTo NextCtrl

        SetFont c, APP_FONT, FONT_SIZE_NORMAL

        Select Case TypeName(c)

            Case "TextBox"
                StyleTextBox c

            Case "ComboBox"
                StyleComboBox c

            Case "ListBox"
                StyleListBox c

            Case "CommandButton"
                StyleButton c

            Case "Label"
                StyleLabel c

            Case "Frame"
                StyleFrame c
                StyleControls c

            Case "CheckBox"
                StyleCheckBox c

            Case "MultiPage"
                StyleMultiPage c

            Case "Page"
                c.BackColor = BG_MAIN
                StyleControls c

        End Select

NextCtrl:
        On Error GoTo 0
    Next c
End Sub

' ako negde rucno zoveš ovo ime
Public Sub ApplyThemeToControls(ByVal frm As Object)
    StyleControls frm
End Sub

' ============================================================
' CONTROL STYLING
' ============================================================
Public Sub StyleTextBox(ByVal c As MSForms.TextBox)
    With c
        .BackColor = INPUT_BG
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = INPUT_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleComboBox(ByVal c As MSForms.ComboBox)
    With c
        .BackColor = INPUT_BG
        .BorderColor = INPUT_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleListBox(ByVal c As MSForms.ListBox)
    With c
        .BackColor = BG_PANEL
        .BorderColor = INPUT_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
    End With
End Sub

Public Sub StyleLabel(ByVal c As MSForms.label, _
                      Optional ByVal forceColor As Long = -1, _
                      Optional ByVal isBold As Boolean = False)

    With c
        .BackStyle = fmBackStyleTransparent

        If forceColor = -1 Then
            .ForeColor = TXT_LIGHT
        Else
            .ForeColor = forceColor
        End If

        If isBold Then
            .Font.name = APP_FONT_BOLD
        Else
            .Font.name = APP_FONT
        End If

        .Font.Size = FONT_SIZE_NORMAL
    End With

    Dim nm As String
    nm = LCase$(c.name)

    If nm Like "*title*" Or nm Like "*naslov*" Then
        c.Font.name = APP_FONT_BOLD
        c.Font.Size = FONT_SIZE_TITLE
    ElseIf nm Like "*field*" Or nm Like "lbl*" Then
        c.Font.Size = FONT_SIZE_SMALL
    End If
End Sub

Public Sub StyleFrame(ByVal c As MSForms.Frame)
    With c
        .BackColor = BG_PANEL
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_SOFT
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_SMALL
    End With
End Sub

Public Sub StyleCheckBox(ByVal c As MSForms.CheckBox)
    With c
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleMultiPage(ByVal c As MSForms.MultiPage)
    Dim i As Long

    c.BackColor = BG_MAIN
    c.Font.name = APP_FONT
    c.Font.Size = FONT_SIZE_SMALL

    For i = 0 To c.Pages.count - 1
        c.Pages(i).BackColor = BG_MAIN
        StyleControls c.Pages(i)
    Next i
End Sub

' ============================================================
' BUTTON STYLING
' ============================================================
Public Sub StyleButton(ByVal c As MSForms.CommandButton)
    c.SpecialEffect = fmSpecialEffectFlat
    c.Font.name = APP_FONT
    c.Font.Size = FONT_SIZE_NORMAL
    c.TakeFocusOnClick = False

    Dim nm As String
    Dim cap As String

    nm = LCase$(c.name)
    cap = LCase$(c.caption)

    If nm Like "*unos*" Or nm Like "*save*" Or nm Like "*sacuvaj*" _
       Or nm Like "*izradi*" Or nm Like "*prikazi*" _
       Or cap Like "*unos*" Or cap Like "*sacuvaj*" _
       Or cap Like "*izradi*" Or cap Like "*prikazi*" Then

        SetButtonPrimary c

    ElseIf nm Like "*obrisi*" Or nm Like "*delete*" _
       Or cap Like "*obrisi*" Or cap Like "*delete*" Then

        SetButtonDanger c

    ElseIf nm Like "*povrat*" Or nm Like "*nazad*" Or nm Like "*zatvori*" _
       Or cap Like "*povrat*" Or cap Like "*nazad*" Or cap Like "*zatvori*" _
       Or cap Like "*izadji*" Then

        SetButtonNav c

    Else
        SetButtonSecondary c
    End If
End Sub

Public Sub StyleMenuButton(ByVal btn As MSForms.CommandButton, Optional ByVal captionText As String = "")
    With btn
        If Len(captionText) > 0 Then .caption = captionText
        .BackColor = BTN_BG
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_NORMAL
        .TakeFocusOnClick = False
        On Error Resume Next
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        On Error GoTo 0
    End With
End Sub

Public Sub StylePrimaryButton(ByVal btn As MSForms.CommandButton, Optional ByVal captionText As String = "")
    With btn
        If Len(captionText) > 0 Then .caption = captionText
        .BackColor = BTN_ACTIVE
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_NORMAL
        .TakeFocusOnClick = False
        On Error Resume Next
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        On Error GoTo 0
    End With
End Sub

Public Sub StyleExitButton(ByVal btn As MSForms.CommandButton, Optional ByVal captionText As String = "")
    With btn
        If Len(captionText) > 0 Then .caption = captionText
        .BackColor = BG_TOP
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_NORMAL
        .TakeFocusOnClick = False
        On Error Resume Next
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        On Error GoTo 0
    End With
End Sub

Public Sub StyleStornoButton(ByVal btn As MSForms.CommandButton, Optional ByVal captionText As String = "")
    With btn
        If Len(captionText) > 0 Then .caption = captionText
        .BackColor = BG_TOP
        .ForeColor = TXT_LIGHT
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_NORMAL
        .TakeFocusOnClick = False
        On Error Resume Next
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        On Error GoTo 0
    End With
End Sub

Public Sub ButtonHover(ByVal btn As MSForms.CommandButton)
    btn.BackColor = BTN_HOVER
    btn.ForeColor = TXT_LIGHT
End Sub

Public Sub ButtonActive(ByVal btn As MSForms.CommandButton)
    btn.BackColor = BTN_ACTIVE
    btn.ForeColor = TXT_LIGHT
End Sub

Public Sub SetButtonPrimary(ByVal btn As Object)
    btn.BackColor = BTN_ACTIVE
    btn.ForeColor = TXT_LIGHT
    btn.Font.Bold = True
End Sub

Public Sub SetButtonSecondary(ByVal btn As Object)
    btn.BackColor = BTN_BG
    btn.ForeColor = TXT_LIGHT
    btn.Font.Bold = False
End Sub

Public Sub SetButtonDanger(ByVal btn As Object)
    btn.BackColor = TXT_ALERT
    btn.ForeColor = TXT_LIGHT
    btn.Font.Bold = True
End Sub

Public Sub SetButtonNav(ByVal btn As Object)
    btn.BackColor = BG_TOP
    btn.ForeColor = TXT_MUTED
    btn.Font.Bold = False
End Sub

Public Sub ResetButtonGroup(ParamArray buttons() As Variant)
    Dim i As Long
    For i = LBound(buttons) To UBound(buttons)
        If Not buttons(i) Is Nothing Then
            StyleMenuButton buttons(i)
        End If
    Next i
End Sub

Public Sub ResetButtonGroupWithExit(ByVal exitBtn As MSForms.CommandButton, ParamArray buttons() As Variant)
    Dim i As Long

    For i = LBound(buttons) To UBound(buttons)
        If Not buttons(i) Is Nothing Then
            StyleMenuButton buttons(i)
        End If
    Next i

    If Not exitBtn Is Nothing Then
        StyleExitButton exitBtn
    End If
End Sub

' ============================================================
' FIELD HELPERS
' ============================================================
Public Sub EnableField(ByVal txt As MSForms.TextBox)
    txt.enabled = True
    txt.BackColor = INPUT_BG
    txt.ForeColor = TXT_LIGHT
    txt.BorderColor = INPUT_BORDER
End Sub

Public Sub DisableField(ByVal txt As MSForms.TextBox)
    txt.enabled = False
    txt.value = ""
    txt.BackColor = INPUT_DISABLED_BG
    txt.ForeColor = TXT_MUTED
    txt.BorderColor = INPUT_DISABLED_BG
End Sub

Public Sub EnableCombo(ByVal cmb As MSForms.ComboBox)
    cmb.enabled = True
    cmb.BackColor = INPUT_BG
    cmb.ForeColor = TXT_LIGHT
End Sub

Public Sub DisableCombo(ByVal cmb As MSForms.ComboBox)
    cmb.enabled = False
    cmb.value = ""
    cmb.BackColor = INPUT_DISABLED_BG
    cmb.ForeColor = TXT_MUTED
End Sub

' ============================================================
' LABEL STATUS HELPERS
' ============================================================
Public Sub SetLabelSuccess(ByVal lbl As MSForms.label)
    lbl.ForeColor = CLR_SUCCESS
End Sub

Public Sub SetLabelWarning(ByVal lbl As MSForms.label)
    lbl.ForeColor = CLR_WARNING
End Sub

Public Sub SetLabelError(ByVal lbl As MSForms.label)
    lbl.ForeColor = CLR_ERROR
End Sub

Public Sub SetLabelMuted(ByVal lbl As MSForms.label)
    lbl.ForeColor = TXT_MUTED
End Sub

' ============================================================
' SECTION HELPERS
' ============================================================
Public Sub SetSectionHeader(ByVal frm As MSForms.Frame, ByVal title As String)
    frm.caption = UCase$(title)
    frm.Font.Bold = True
    frm.Font.Size = FONT_SIZE_SMALL
    frm.ForeColor = TXT_LIGHT
End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================
Private Sub SetFont(ByVal c As Object, ByVal fontName As String, ByVal fontSize As Single)
    On Error Resume Next
    c.Font.name = fontName
    c.Font.Size = fontSize
    On Error GoTo 0
End Sub

Private Function IsDirectChild(ByVal c As Object, ByVal parent As Object) As Boolean
    On Error Resume Next

    If TypeName(parent) = "UserForm" Then
        IsDirectChild = (TypeName(c.parent) = TypeName(parent))
        If Err.Number <> 0 Then IsDirectChild = True

    ElseIf TypeName(parent) = "Frame" Then
        Dim p As Object
        Set p = c.parent
        IsDirectChild = (p.name = parent.name)
        If Err.Number <> 0 Then IsDirectChild = True

    ElseIf TypeName(parent) = "Page" Then
        IsDirectChild = True

    Else
        IsDirectChild = True
    End If

    On Error GoTo 0
End Function

' ============================================================
' v6.11 UI ENHANCEMENTS — focus border, section header,
' status dot, accent bar, badge, KPI card, step indicator
' ============================================================

' --- FOCUS BORDER (gold accent on active input) ---

Public Sub ApplyFocusBorder(ByVal c As Object)
    On Error Resume Next
    c.BorderColor = BTN_ACTIVE()
    c.SpecialEffect = fmSpecialEffectFlat
    On Error GoTo 0
End Sub

Public Sub RemoveFocusBorder(ByVal c As Object)
    On Error Resume Next
    c.BorderColor = INPUT_BORDER()
    c.SpecialEffect = fmSpecialEffectFlat
    On Error GoTo 0
End Sub

' --- SECTION HEADER UPPERCASE (used in frmDokumenta frames) ---

Public Sub StyleSectionHeader(ByVal frm As MSForms.Frame, ByVal title As String)
    With frm
        .caption = " " & UCase$(title)
        .BackColor = BG_PANEL()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_SOFT()
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = BTN_ACTIVE()      ' gold caption
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_SMALL
    End With
End Sub

' --- STATUS DOT (small colored circle via Label) ---

Public Sub StyleStatusDot(ByVal lbl As MSForms.label, _
                          Optional ByVal okState As Boolean = True)
    With lbl
        .caption = ChrW(&H25CF)         ' filled circle
        .BackStyle = fmBackStyleTransparent
        .Font.name = "Segoe UI Symbol"
        .Font.Size = 12
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        If okState Then
            .ForeColor = CLR_SUCCESS()
        Else
            .ForeColor = CLR_ERROR()
        End If
    End With
End Sub

' --- SIDEBAR ACCENT BAR (vertical gold strip on active nav item) ---

Public Sub StyleAccentBar(ByVal lbl As MSForms.label)
    With lbl
        .caption = ""
        .BackStyle = fmBackStyleOpaque
        .BackColor = BTN_ACTIVE()
        .BorderStyle = fmBorderStyleNone
    End With
End Sub

' --- BADGE (red counter pill, e.g. "3" on Banka nav) ---

Public Sub StyleBadge(ByVal lbl As MSForms.label, ByVal countValue As Long)
    With lbl
        If countValue <= 0 Then
            .Visible = False
            Exit Sub
        End If
        .Visible = True
        .caption = CStr(countValue)
        .BackStyle = fmBackStyleOpaque
        .BackColor = TXT_ALERT()
        .ForeColor = TXT_LIGHT()
        .TextAlign = fmTextAlignCenter
        .Font.name = APP_FONT_BOLD
        .Font.Size = 8
        .BorderStyle = fmBorderStyleNone
    End With
End Sub

' --- KPI CARD (small stat block, used at sidebar bottom and top bar) ---

Public Sub StyleKpiCardFrame(ByVal frm As MSForms.Frame)
    With frm
        .caption = ""
        .BackColor = BG_TOP()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_SOFT()
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Public Sub StyleKpiLabel(ByVal lbl As MSForms.label, ByVal kind As String)
    ' kind: "title" | "value" | "delta_up" | "delta_down" | "muted"
    With lbl
        .BackStyle = fmBackStyleTransparent
        Select Case LCase$(kind)
            Case "title"
                .ForeColor = TXT_MUTED()
                .Font.name = APP_FONT
                .Font.Size = FONT_SIZE_SMALL
                .Font.Bold = False
            Case "value"
                .ForeColor = TXT_LIGHT()
                .Font.name = APP_FONT_BOLD
                .Font.Size = FONT_SIZE_TITLE
                .Font.Bold = True
            Case "delta_up"
                .ForeColor = CLR_SUCCESS()
                .Font.name = APP_FONT
                .Font.Size = FONT_SIZE_SMALL
            Case "delta_down"
                .ForeColor = CLR_ERROR()
                .Font.name = APP_FONT
                .Font.Size = FONT_SIZE_SMALL
            Case Else
                .ForeColor = TXT_MUTED()
                .Font.name = APP_FONT
                .Font.Size = FONT_SIZE_SMALL
        End Select
    End With
End Sub

' --- INLINE PREVIEW BOX (used for "UKUPNO XXX RSD") ---

Public Sub StylePreviewBox(ByVal lbl As MSForms.label, _
                            Optional ByVal kind As String = "info")
    ' kind: "info" | "warn" | "danger" | "ok"
    With lbl
        .BackStyle = fmBackStyleOpaque
        .BackColor = BG_MAIN()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_SOFT()
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_SMALL
        .TextAlign = fmTextAlignCenter
        Select Case LCase$(kind)
            Case "ok"
                .ForeColor = CLR_SUCCESS()
            Case "warn"
                .ForeColor = CLR_WARNING()
            Case "danger"
                .ForeColor = CLR_ERROR()
            Case Else
                .ForeColor = TXT_LIGHT()
        End Select
    End With
End Sub

' --- F-KEY ACCELERATOR HINT (small gold "[F2]" label) ---

Public Sub StyleFkeyHint(ByVal lbl As MSForms.label, ByVal fKey As String)
    With lbl
        .caption = "[" & fKey & "]"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = BTN_ACTIVE()
        .Font.name = APP_FONT
        .Font.Size = 8
        .Font.Bold = False
        .TextAlign = fmTextAlignRight
    End With
End Sub

' ============================================================
' v6.11 UI: TOP KPI CARD HELPERS
' ============================================================

Public Sub StyleTopKpi(ByVal frm As MSForms.Frame, _
                       ByVal lblTitle As MSForms.label, _
                       ByVal lblValue As MSForms.label, _
                       ByVal lblAccent As MSForms.label, _
                       ByVal kind As String)
    ' kind: "ok" | "warn" | "danger" | "neutral"

    With frm
        .caption = ""
        .BackColor = BG_TOP()
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_SOFT()
        .SpecialEffect = fmSpecialEffectFlat
    End With

    With lblTitle
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_MUTED()
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
        .Font.Bold = False
    End With

    With lblValue
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_LIGHT()
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_TITLE
        .Font.Bold = True
    End With

    ' Accent strip (uska vertikalna ili horizontalna linija desno)
    With lblAccent
        .caption = ""
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
    End With

    Select Case LCase$(kind)
        Case "ok":      lblAccent.BackColor = CLR_SUCCESS()
        Case "warn":    lblAccent.BackColor = CLR_WARNING()
        Case "danger":  lblAccent.BackColor = CLR_ERROR()
        Case Else:      lblAccent.BackColor = BTN_ACTIVE()
    End Select
End Sub

Public Sub LayoutTopKpiInternals(ByVal frm As MSForms.Frame, _
                                  ByVal lblTitle As MSForms.label, _
                                  ByVal lblValue As MSForms.label, _
                                  ByVal lblAccent As MSForms.label)
    ' Pozicionira label-e unutar Frame-a po istom obrascu kao na renderu.
    Dim padL As Single, padT As Single

    padL = 12
    padT = 8

    With lblTitle
        .Left = padL
        .Top = padT
        .width = frm.InsideWidth - padL - 16
        .Height = 14
    End With

    With lblValue
        .Left = padL
        .Top = lblTitle.Top + lblTitle.Height + 4
        .width = frm.InsideWidth - padL - 16
        .Height = 22
    End With

    ' Tanka vertikalna linija u desnom uglu Frame-a
    With lblAccent
        .width = 3
        .Height = frm.InsideHeight - 16
        .Left = frm.InsideWidth - 11
        .Top = 8
    End With
End Sub

' --- Section accent line (tanka horizontalna traka iznad Frame-a) ---

Public Sub StyleSectionAccent(ByVal lbl As MSForms.label, _
                               ByVal frmBelow As MSForms.Frame, _
                               Optional ByVal kind As String = "primary")
    On Error Resume Next

    With lbl
        .caption = ""
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .Visible = True
        .Height = 3
    End With

    ' Detektuj da li je label child of Frame ili child of UserForm
    Dim isChildOfFrame As Boolean
    isChildOfFrame = False

    Dim p As Object
    Set p = lbl.parent

    If TypeName(p) = "Frame" Then
        If p.name = frmBelow.name Then
            isChildOfFrame = True
        End If
    End If

    If isChildOfFrame Then
        ' label zivi unutar Frame-a — pozicija u Frame koordinatama
        lbl.Left = 0
        lbl.Top = 0
        lbl.width = frmBelow.InsideWidth
    Else
        ' label zivi na formi — pozicija u form koordinatama
        lbl.Left = frmBelow.Left
        lbl.Top = frmBelow.Top
        lbl.width = frmBelow.width
    End If

    lbl.ZOrder 0   ' iznad Frame border-a

    Select Case LCase$(kind)
        Case "ok":       lbl.BackColor = CLR_SUCCESS()
        Case "warn":     lbl.BackColor = CLR_WARNING()
        Case "danger":   lbl.BackColor = CLR_ERROR()
        Case "info":     lbl.BackColor = RGB(80, 130, 200)
        Case Else:       lbl.BackColor = BTN_ACTIVE()
    End Select

    On Error GoTo 0
End Sub

' --- Subtitle ispod naslova ---

Public Sub StyleSubtitle(ByVal lbl As MSForms.label, ByVal text As String)
    With lbl
        .caption = text
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_MUTED()
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
        .Font.Bold = False
    End With
End Sub

Public Sub StyleFrameTitleLabel(ByVal lbl As MSForms.label, ByVal title As String)
    With lbl
        .caption = UCase$(title)
        .BackStyle = fmBackStyleTransparent
        .ForeColor = BTN_ACTIVE()        ' gold
        .Font.name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_HEADER
        .Font.Bold = True
        .TextAlign = fmTextAlignLeft
    End With
End Sub

'======================================================================
' StyleListHeaderLabel
'
' Header label iznad ListBox kolone. Vizuelno suptilan ali jasan —
' bold + muted boja, manji font od subtitle-a.
'
' Koristi se za "fake header row" iznad VBA ListBox-a (MSForms ListBox
' nema native column header).
'======================================================================
Public Sub StyleListHeaderLabel(ByVal lbl As MSForms.label)
    On Error Resume Next
    
    With lbl
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = BTN_ACTIVE()
        .BackStyle = fmBackStyleOpaque        ' bilo Transparent
        .BackColor = BG_PANEL()                ' dodaj ovo
        .TextAlign = fmTextAlignLeft
    End With
    
    On Error GoTo 0
End Sub
