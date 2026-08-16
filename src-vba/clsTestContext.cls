VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsTestContext -- snimi stanje Excela na ulasku, vrati ga na izlasku.
'
' Zamenjuje obrazac koji je do sada svaki test modul nosio rucno: wasQuiet,
' quietSet, pa RestoreJournalQuiet u dva mesta (normalan izlaz i EH), pa jos
' Application.Calculation negde treci put. Svaka od tih kopija je jedna prilika
' da se stanje ne vrati -- a nevraceno stanje ne obara suite koja ga je ostavila,
' nego SLEDECU.
'
' VBA nema Try/Finally. Najblize sto postoji je Class_Terminate: kad objekat
' izadje iz opsega -- i na urednom Exit Sub, i posle Err.Raise, i kad suite
' pukne na nacin koji EH ne pokrije -- VBA ga zove. Zato je restore ovde, a ne u
' proceduri koju neko mora da se seti da pozove.
'
' Upotreba:
'     Dim ctx As clsTestContext
'     Set ctx = New clsTestContext     ' snapshot se uzima odmah
'     ctx.Quiet                        ' gasi journal trag i UI cekanja
'     ...
'     ' nema sta da se cisti rucno; Set ctx = Nothing ili kraj procedure vraca sve
'
' Provera da je stanje stvarno vraceno: ctx.Drift() posle Restore mora biti "".
' ============================================================

Private mTaken As Boolean
Private mRestored As Boolean
Private mRestoreError As String

Private mCalculation As Long
Private mEnableEvents As Boolean
Private mScreenUpdating As Boolean
Private mDisplayAlerts As Boolean
Private mStatusBar As Variant
Private mCursor As Long

Private mTestMode As Boolean
Private mQuiet As Boolean
Private mQuietTouched As Boolean

Private Sub Class_Initialize()
    Snapshot
End Sub

' Ovo je "finally" ovog frameworka. Ne oslanja se na to da je suite napisala
' cleanup granu -- VBA zove Terminate i kad je nije napisala.
Private Sub Class_Terminate()
    On Error Resume Next
    Restore
    On Error GoTo 0
End Sub

Public Sub Snapshot()
    On Error Resume Next
    mCalculation = Application.Calculation
    mEnableEvents = Application.EnableEvents
    mScreenUpdating = Application.ScreenUpdating
    mDisplayAlerts = Application.DisplayAlerts
    mStatusBar = Application.StatusBar
    mCursor = Application.Cursor
    mTestMode = modTestMode.IsTestMode()
    mQuiet = modJournaling.IsTestModeQuiet()
    On Error GoTo 0

    mTaken = True
    mRestored = False
    mQuietTouched = False
End Sub

' Gasi sve sto ceka operatera ili ostavlja trag: MsgBox/SetFocus gard
' (modTestMode) i CSV crash-recovery journal (modJournaling). Journal je bitan
' zato sto ga tx.RollbackTx NE povlaci -- test redovi bi ostali u Journal
' folderu i sledeci start bi javio lazno upozorenje o gubitku podataka.
Public Sub Quiet()
    On Error Resume Next
    modTestMode.SetTestMode True
    modJournaling.SetTestModeQuiet True
    On Error GoTo 0
    mQuietTouched = True
End Sub

Public Sub Restore()
    If Not mTaken Then Exit Sub
    If mRestored Then Exit Sub
    mRestored = True

    ' On Error Resume Next mora da stoji: restore se zove i iz Class_Terminate,
    ' gde podignuta greska nema ko da uhvati. Ali greska se ZAPAMTI i izlazi kroz
    ' Drift() -- inace bi neuspeo restore nestao bez traga, a to je tacno stanje
    ' koje ova klasa postoji da spreci.
    On Error Resume Next
    Err.Clear

    ' Journal se vraca na ZATECENO stanje, ne bezuslovno na False -- suite ume da
    ' se pokrene iz konteksta koji je journal vec bio stisao.
    If mQuietTouched Then
        modJournaling.SetTestModeQuiet mQuiet
        modJournaling.StopAutoSaveTimer      ' posle rollback-a nema sta da se snima
    End If
    modTestMode.SetTestMode mTestMode

    Application.Calculation = mCalculation
    Application.EnableEvents = mEnableEvents
    Application.ScreenUpdating = mScreenUpdating
    Application.DisplayAlerts = mDisplayAlerts
    Application.StatusBar = mStatusBar
    Application.Cursor = mCursor

    If Err.Number <> 0 Then
        mRestoreError = "restore pao: " & CStr(Err.Number) & " " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Prazan string = stanje je isto kao na ulasku. Sve ostalo je opis razlike.
' Zove se posle Restore da bi se DOKAZALO da je vraceno, umesto da se
' pretpostavi. Isto sto CLEANUP kapija radi nad tabelama, samo za stanje aplikacije.
Public Function Drift() As String
    Dim s As String
    If Len(mRestoreError) > 0 Then s = mRestoreError & "; "
    On Error Resume Next
    If Application.Calculation <> mCalculation Then _
        s = s & "Calculation=" & Application.Calculation & " (bilo " & mCalculation & "); "
    If Application.EnableEvents <> mEnableEvents Then _
        s = s & "EnableEvents=" & Application.EnableEvents & "; "
    If Application.ScreenUpdating <> mScreenUpdating Then _
        s = s & "ScreenUpdating=" & Application.ScreenUpdating & "; "
    If Application.DisplayAlerts <> mDisplayAlerts Then _
        s = s & "DisplayAlerts=" & Application.DisplayAlerts & "; "
    If Application.Cursor <> mCursor Then _
        s = s & "Cursor=" & Application.Cursor & "; "
    If modTestMode.IsTestMode() <> mTestMode Then _
        s = s & "TestMode=" & modTestMode.IsTestMode() & "; "
    If modJournaling.IsTestModeQuiet() <> mQuiet Then _
        s = s & "JournalQuiet=" & modJournaling.IsTestModeQuiet() & "; "
    On Error GoTo 0
    Drift = s
End Function
