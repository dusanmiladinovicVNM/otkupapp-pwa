VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSnapshots As Object ' Scripting.Dictionary
Private mActive As Boolean

Private mCalc As XlCalculation
Private mScreen As Boolean
Private mEvents As Boolean
Private mDisplayStatusBar As Boolean
Private mStatusBarText As Variant


Public Sub BeginTx()
    If mActive Then Err.Raise 5, , "Transaction already active."

    Set mSnapshots = CreateObject("Scripting.Dictionary")
    mActive = True

    ' Performance / stability
    mCalc = Application.Calculation
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    mDisplayStatusBar = Application.DisplayStatusBar
    mStatusBarText = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' Snimi snapshot tabele (ListObject) po imenu
Public Sub AddTableSnapshot(ByVal tableName As String)
    If Not mActive Then Err.Raise 5, , "No active transaction."
    
    Dim lo As ListObject
    Set lo = modDataAccess.GetTable(tableName)
    
    Dim snap(0 To 1) As Variant  ' (0)=data, (1)=colCount
    If lo.DataBodyRange Is Nothing Then
        snap(0) = Empty
    Else
        snap(0) = lo.DataBodyRange.Value2
    End If
    snap(1) = lo.ListColumns.count
    
    If Not mSnapshots.Exists(tableName) Then
        mSnapshots.Add tableName, snap
    End If
End Sub

Public Sub CommitTx()
    If Not mActive Then Exit Sub
    Cleanup
    mActive = False
End Sub

Public Sub RollbackTx()
    If Not mActive Then Exit Sub

    Dim key As Variant
    For Each key In mSnapshots.keys
        RestoreTable CStr(key), mSnapshots(key)
    Next key

    Cleanup
    mActive = False
End Sub

Private Sub Cleanup()
    On Error Resume Next
    Application.ScreenUpdating = mScreen
    Application.EnableEvents = mEvents
    Application.Calculation = mCalc
    Application.DisplayStatusBar = mDisplayStatusBar
    Application.StatusBar = mStatusBarText
    On Error GoTo 0
    
    Set mSnapshots = Nothing
    mActive = False
End Sub

Private Sub RestoreTable(ByVal tableName As String, ByVal snap As Variant)
    Dim lo As ListObject
    Set lo = modDataAccess.GetTable(tableName)
    
    Dim snapData As Variant: snapData = snap(0)
    Dim snapCols As Long: snapCols = CLng(snap(1))
    
    ' Sanity check
    If lo.ListColumns.count <> snapCols Then
        Err.Raise 5, , "Column count mismatch: " & tableName
    End If
    
    If IsEmpty(snapData) Then
        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.Delete
        End If
        Exit Sub
    End If
    
    Dim neededRows As Long
    neededRows = UBound(snapData, 1)
    
    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.rows.count
    End If
    
    ' Rows anpassen
    Do While currentRows < neededRows
        lo.ListRows.Add
        currentRows = currentRows + 1
    Loop
    Do While currentRows > neededRows
        lo.ListRows(currentRows).Delete
        currentRows = currentRows - 1
    Loop
    
    lo.DataBodyRange.Value2 = snapData
End Sub
