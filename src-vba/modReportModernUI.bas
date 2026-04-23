Attribute VB_Name = "modReportModernUI"
' ============================================================
' UI UPGRADE PACK FOR frmIzvestaj
' Purpose:
'   Keep ALL existing reporting logic intact and only modernize the UI.
'
' What this gives you:
'   - dark modern app styling
'   - better visual hierarchy
'   - styled toggle buttons
'   - styled listboxes
'   - cleaner tabs / pages
'   - safer UI helpers that avoid unsupported properties
'
' How to use:
'   1) Keep your existing frmIzvestaj business/report code.
'   2) Add the helper module below as: modReportUITheme
'   3) Add the form procedures below into frmIzvestaj.
'   4) In UserForm_Activate, keep your existing logic, but call:
'        ApplyReportTheme Me
'        InitReportUI
'      near the top.
'   5) Replace only UI-related procedures with the upgraded versions below:
'        - UpdateEntitetToggleUI
'        - UpdateTipToggleUI
'        - SetupListBoxes
'        - UpdateReportMode   (optional safer visual refresh version)
'
' IMPORTANT:
'   This pack does NOT change your report generation logic.
'   All Generate...Report procedures stay as they are.
' ============================================================


' ============================================================
' STANDARD MODULE: modReportUITheme
' ============================================================
Option Explicit

Public Const UI_BG As Long = &H1A2233
Public Const UI_PANEL As Long = &H223049
Public Const UI_PANEL_ALT As Long = &H2A3955
Public Const UI_ACCENT As Long = &HC06A1F
Public Const UI_ACCENT_SOFT As Long = &H8A4E1A
Public Const UI_TEXT As Long = &HF4F7FB
Public Const UI_MUTED As Long = &HC8D3E2
Public Const UI_OK As Long = &H3E8E41
Public Const UI_DANGER As Long = &H3F3FBF
Public Const UI_INPUT As Long = &H202B40

Public Sub ApplyReportTheme(frm As Object)
    On Error Resume Next

    frm.BackColor = UI_BG
    frm.Caption = "Izvestaji"

    StyleLabelSafe frm, "lblTitle", UI_TEXT, 18, True
    StyleLabelSafe frm, "lblKarticaKoop", UI_TEXT, 11, True
    StyleLabelSafe frm, "lblKarticaPeriod", UI_MUTED, 10, False

    StyleFrameSafe frm, "fraFilters"
    StyleFrameSafe frm, "fraTop"
    StyleFrameSafe frm, "fraTabs"
    StyleFrameSafe frm, "fraActions"

    StyleInputSafe frm, "cmbEntitet"
    StyleInputSafe frm, "cmbVrstaRobe"
    StyleInputSafe frm, "txtDatumOd"
    StyleInputSafe frm, "txtDatumDo"

    StyleButtonSafe frm, "btnUnos", True
    StyleButtonSafe frm, "btnStampaj", False
    StyleButtonSafe frm, "btnPovratak", False
    StyleButtonSafe frm, "btnStampajKarticu", False

    StyleToggleSafe frm, "tglOM", False, False
    StyleToggleSafe frm, "tglKupci", False, False
    StyleToggleSafe frm, "tglVozaci", False, False
    StyleToggleSafe frm, "tglKooperanti", False, False
    StyleToggleSafe frm, "tglPojedinacni", False, False
    StyleToggleSafe frm, "tglZbirni", False, False

    StyleMultiPageSafe frm, "mpReports"

    On Error GoTo 0
End Sub

Public Sub InitReportUICommon(frm As Object)
    On Error Resume Next

    frm.BackColor = UI_BG

    If HasControl(frm, "lblTitle") Then
        With frm.Controls("lblTitle")
            .Caption = "Izvestaji"
            .BackStyle = fmBackStyleTransparent
        End With
    End If

    On Error GoTo 0
End Sub

Public Sub RefreshEntityToggleTheme(frm As Object)
    StyleToggleSafe frm, "tglOM", CBool(frm.Controls("tglOM").Value), False
    StyleToggleSafe frm, "tglKupci", CBool(frm.Controls("tglKupci").Value), False
    StyleToggleSafe frm, "tglVozaci", CBool(frm.Controls("tglVozaci").Value), False
    StyleToggleSafe frm, "tglKooperanti", CBool(frm.Controls("tglKooperanti").Value), False
End Sub

Public Sub RefreshModeToggleTheme(frm As Object)
    StyleToggleSafe frm, "tglPojedinacni", CBool(frm.Controls("tglPojedinacni").Value), True
    StyleToggleSafe frm, "tglZbirni", CBool(frm.Controls("tglZbirni").Value), False
End Sub

Public Sub StyleListBoxSafe(frm As Object, ByVal ctlName As String, ByVal colCount As Long, ByVal colWidths As String)
    On Error Resume Next
    With frm.Controls(ctlName)
        .ColumnCount = colCount
        .ColumnWidths = colWidths
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = UI_PANEL
        .foreColor = UI_TEXT
        .BorderStyle = fmBorderStyleSingle
        .IntegralHeight = False
        .MultiSelect = fmMultiSelectSingle
    End With
    On Error GoTo 0
End Sub

Public Sub StyleLabelSafe(frm As Object, ByVal ctlName As String, ByVal foreColor As Long, ByVal fontSize As Single, ByVal isBold As Boolean)
    On Error Resume Next
    With frm.Controls(ctlName)
        .Font.Name = "Segoe UI"
        .Font.Size = fontSize
        .Font.Bold = isBold
        .foreColor = foreColor
        .BackStyle = fmBackStyleTransparent
    End With
    On Error GoTo 0
End Sub

Public Sub StyleFrameSafe(frm As Object, ByVal ctlName As String)
    On Error Resume Next
    With frm.Controls(ctlName)
        .BackColor = UI_BG
        .Caption = Replace(.Caption, "-", "")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .foreColor = UI_MUTED
    End With
    On Error GoTo 0
End Sub

Public Sub StyleInputSafe(frm As Object, ByVal ctlName As String)
    On Error Resume Next
    With frm.Controls(ctlName)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = UI_INPUT
        .foreColor = UI_TEXT
        .BorderStyle = fmBorderStyleSingle
    End With
    On Error GoTo 0
End Sub

Public Sub StyleButtonSafe(frm As Object, ByVal ctlName As String, ByVal isPrimary As Boolean)
    On Error Resume Next
    With frm.Controls(ctlName)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10.5
        .foreColor = UI_TEXT
        .BackColor = IIf(isPrimary, UI_ACCENT, UI_PANEL_ALT)
        .TakeFocusOnClick = False
    End With
    On Error GoTo 0
End Sub

Public Sub StyleToggleSafe(frm As Object, ByVal ctlName As String, ByVal isActive As Boolean, ByVal isPrimaryGroup As Boolean)
    On Error Resume Next
    With frm.Controls(ctlName)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10.5
        .Font.Bold = isActive
        .foreColor = IIf(isActive, UI_TEXT, UI_MUTED)

        If isActive Then
            If isPrimaryGroup Then
                .BackColor = UI_ACCENT
            Else
                .BackColor = UI_ACCENT_SOFT
            End If
        Else
            .BackColor = UI_PANEL
        End If
    End With
    On Error GoTo 0
End Sub

Public Sub StyleMultiPageSafe(frm As Object, ByVal ctlName As String)
    On Error Resume Next
    With frm.Controls(ctlName)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10
        .BackColor = UI_BG
    End With
    On Error GoTo 0
End Sub

Public Function HasControl(frm As Object, ByVal ctlName As String) As Boolean
    On Error Resume Next
    Dim tmp As Object
    Set tmp = frm.Controls(ctlName)
    HasControl = (Err.Number = 0)
    Set tmp = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Public Sub SetPageVisibleSafe(mp As MSForms.MultiPage, ByVal pageIndex As Long, ByVal makeVisible As Boolean)
    On Error Resume Next
    mp.Pages(pageIndex).Visible = makeVisible
    On Error GoTo 0
End Sub
