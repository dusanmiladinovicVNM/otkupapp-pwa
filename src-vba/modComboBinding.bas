Attribute VB_Name = "modComboBinding"

Option Explicit

' ============================================================
' modComboBinding
' Shared ComboBox binding helpers
'
' Rule:
'   Column 0 = display text shown to operator
'   Column 1 = hidden stable ID used by code
'
' No business logic here.
' No MsgBox here.
' No form-specific assumptions here.
' ============================================================

Public Sub FillComboDisplayID(ByVal cmb As MSForms.ComboBox, _
                              ByVal tableName As String, _
                              ByVal displayCol As String, _
                              ByVal idCol As String, _
                              Optional ByVal displayWidth As String = "180 pt")
    On Error GoTo EH

    cmb.Clear
    cmb.ColumnCount = 2
    cmb.ColumnWidths = displayWidth & ";0 pt"
    cmb.BoundColumn = 1
    cmb.TextColumn = 1

    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then Exit Sub

    Dim colDisplay As Long
    Dim colID As Long

    colDisplay = RequireColumnIndex(tableName, displayCol, "modComboBinding.FillComboDisplayID")
    colID = RequireColumnIndex(tableName, idCol, "modComboBinding.FillComboDisplayID")

    Dim i As Long
    Dim displayText As String
    Dim idValue As String

    For i = 1 To UBound(data, 1)
        idValue = Trim$(NzToText(data(i, colID)))
        displayText = Trim$(NzToText(data(i, colDisplay)))

        If idValue <> "" Then
            cmb.AddItem displayText
            cmb.List(cmb.ListCount - 1, 1) = idValue
        End If
    Next i

    Exit Sub

EH:
    LogErr "modComboBinding.FillComboDisplayID"
End Sub

Public Function GetComboID(ByVal cmb As MSForms.ComboBox) As String
    On Error GoTo Fallback

    If cmb.ListIndex >= 0 Then
        If cmb.ColumnCount >= 2 Then
            GetComboID = Trim$(CStr(cmb.List(cmb.ListIndex, 1)))
            Exit Function
        End If
    End If

Fallback:
    ' Fallback for legacy combos like "Name (ID)".
    GetComboID = ExtractIDFromDisplaySafe(Trim$(cmb.value))
End Function

Public Function GetComboDisplay(ByVal cmb As MSForms.ComboBox) As String
    On Error GoTo Fallback

    If cmb.ListIndex >= 0 Then
        GetComboDisplay = Trim$(CStr(cmb.List(cmb.ListIndex, 0)))
        Exit Function
    End If

Fallback:
    GetComboDisplay = Trim$(cmb.value)
End Function

Public Function SetComboByID(ByVal cmb As MSForms.ComboBox, ByVal idValue As String) As Boolean
    On Error GoTo EH

    Dim wanted As String
    wanted = Trim$(idValue)

    If wanted = "" Then
        cmb.ListIndex = -1
        cmb.value = ""
        SetComboByID = True
        Exit Function
    End If

    Dim i As Long

    If cmb.ColumnCount >= 2 Then
        For i = 0 To cmb.ListCount - 1
            If Trim$(CStr(cmb.List(i, 1))) = wanted Then
                cmb.ListIndex = i
                SetComboByID = True
                Exit Function
            End If
        Next i
    End If

    SetComboByID = False
    Exit Function

EH:
    LogErr "modComboBinding.SetComboByID"
    SetComboByID = False
End Function

Public Function ExtractIDFromDisplaySafe(ByVal displayText As String) As String
    On Error GoTo EH

    Dim p1 As Long
    Dim p2 As Long

    p1 = InStrRev(displayText, "(")
    p2 = InStrRev(displayText, ")")

    If p1 > 0 And p2 > p1 Then
        ExtractIDFromDisplaySafe = Trim$(Mid$(displayText, p1 + 1, p2 - p1 - 1))
    Else
        ExtractIDFromDisplaySafe = ""
    End If

    Exit Function

EH:
    ExtractIDFromDisplaySafe = ""
End Function

