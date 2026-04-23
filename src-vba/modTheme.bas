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
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleComboBox(ByVal c As MSForms.ComboBox)
    With c
        .BackColor = INPUT_BG
        .BorderColor = INPUT_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleListBox(ByVal c As MSForms.ListBox)
    With c
        .BackColor = BG_PANEL
        .BorderColor = INPUT_BORDER
        .SpecialEffect = fmSpecialEffectFlat
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
    End With
End Sub

Public Sub StyleLabel(ByVal c As MSForms.Label, _
                      Optional ByVal forceColor As Long = -1, _
                      Optional ByVal isBold As Boolean = False)

    With c
        .BackStyle = fmBackStyleTransparent

        If forceColor = -1 Then
            .foreColor = TXT_LIGHT
        Else
            .foreColor = forceColor
        End If

        If isBold Then
            .Font.Name = APP_FONT_BOLD
        Else
            .Font.Name = APP_FONT
        End If

        .Font.Size = FONT_SIZE_NORMAL
    End With

    Dim nm As String
    nm = LCase$(c.Name)

    If nm Like "*title*" Or nm Like "*naslov*" Then
        c.Font.Name = APP_FONT_BOLD
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
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT_BOLD
        .Font.Size = FONT_SIZE_SMALL
    End With
End Sub

Public Sub StyleCheckBox(ByVal c As MSForms.CheckBox)
    With c
        .BackStyle = fmBackStyleTransparent
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT
        .Font.Size = FONT_SIZE_NORMAL
    End With
End Sub

Public Sub StyleMultiPage(ByVal c As MSForms.MultiPage)
    Dim i As Long

    c.BackColor = BG_MAIN
    c.Font.Name = APP_FONT
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
    c.Font.Name = APP_FONT
    c.Font.Size = FONT_SIZE_NORMAL
    c.TakeFocusOnClick = False

    Dim nm As String
    Dim cap As String

    nm = LCase$(c.Name)
    cap = LCase$(c.Caption)

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
        If Len(captionText) > 0 Then .Caption = captionText
        .BackColor = BTN_BG
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT_BOLD
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
        If Len(captionText) > 0 Then .Caption = captionText
        .BackColor = BTN_ACTIVE
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT_BOLD
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
        If Len(captionText) > 0 Then .Caption = captionText
        .BackColor = BG_TOP
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT_BOLD
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
        If Len(captionText) > 0 Then .Caption = captionText
        .BackColor = BG_TOP
        .foreColor = TXT_LIGHT
        .Font.Name = APP_FONT_BOLD
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
    btn.foreColor = TXT_LIGHT
End Sub

Public Sub ButtonActive(ByVal btn As MSForms.CommandButton)
    btn.BackColor = BTN_ACTIVE
    btn.foreColor = TXT_LIGHT
End Sub

Public Sub SetButtonPrimary(ByVal btn As Object)
    btn.BackColor = BTN_ACTIVE
    btn.foreColor = TXT_LIGHT
    btn.Font.Bold = True
End Sub

Public Sub SetButtonSecondary(ByVal btn As Object)
    btn.BackColor = BTN_BG
    btn.foreColor = TXT_LIGHT
    btn.Font.Bold = False
End Sub

Public Sub SetButtonDanger(ByVal btn As Object)
    btn.BackColor = TXT_ALERT
    btn.foreColor = TXT_LIGHT
    btn.Font.Bold = True
End Sub

Public Sub SetButtonNav(ByVal btn As Object)
    btn.BackColor = BG_TOP
    btn.foreColor = TXT_MUTED
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
    txt.Enabled = True
    txt.BackColor = INPUT_BG
    txt.foreColor = TXT_LIGHT
    txt.BorderColor = INPUT_BORDER
End Sub

Public Sub DisableField(ByVal txt As MSForms.TextBox)
    txt.Enabled = False
    txt.Value = ""
    txt.BackColor = INPUT_DISABLED_BG
    txt.foreColor = TXT_MUTED
    txt.BorderColor = INPUT_DISABLED_BG
End Sub

Public Sub EnableCombo(ByVal cmb As MSForms.ComboBox)
    cmb.Enabled = True
    cmb.BackColor = INPUT_BG
    cmb.foreColor = TXT_LIGHT
End Sub

Public Sub DisableCombo(ByVal cmb As MSForms.ComboBox)
    cmb.Enabled = False
    cmb.Value = ""
    cmb.BackColor = INPUT_DISABLED_BG
    cmb.foreColor = TXT_MUTED
End Sub

' ============================================================
' LABEL STATUS HELPERS
' ============================================================
Public Sub SetLabelSuccess(ByVal lbl As MSForms.Label)
    lbl.foreColor = CLR_SUCCESS
End Sub

Public Sub SetLabelWarning(ByVal lbl As MSForms.Label)
    lbl.foreColor = CLR_WARNING
End Sub

Public Sub SetLabelError(ByVal lbl As MSForms.Label)
    lbl.foreColor = CLR_ERROR
End Sub

Public Sub SetLabelMuted(ByVal lbl As MSForms.Label)
    lbl.foreColor = TXT_MUTED
End Sub

' ============================================================
' SECTION HELPERS
' ============================================================
Public Sub SetSectionHeader(ByVal frm As MSForms.Frame, ByVal title As String)
    frm.Caption = UCase$(title)
    frm.Font.Bold = True
    frm.Font.Size = FONT_SIZE_SMALL
    frm.foreColor = TXT_LIGHT
End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================
Private Sub SetFont(ByVal c As Object, ByVal fontName As String, ByVal fontSize As Single)
    On Error Resume Next
    c.Font.Name = fontName
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
        IsDirectChild = (p.Name = parent.Name)
        If Err.Number <> 0 Then IsDirectChild = True

    ElseIf TypeName(parent) = "Page" Then
        IsDirectChild = True

    Else
        IsDirectChild = True
    End If

    On Error GoTo 0
End Function

