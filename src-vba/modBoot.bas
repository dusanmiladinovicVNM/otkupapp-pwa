Attribute VB_Name = "modBoot"
Option Explicit

' ============================================================
' modBoot - pokretanje / gasenje aplikacije nezavisno od toga
' da li kod radi kao .xlam (ENGINE) ili kao jedan .xlsm (COMBINED).
'
' COMBINED (.xlsm): ThisWorkbook.Workbook_Open poziva BootFromData(ThisWorkbook).
' ENGINE   (.xlam): add-in oslukuje Application evente (clsAppEvents) i kad se
'                   otvori klijentski .xlsx (sa marker tabelom), poziva BootFromData.
'
' Sva startup logika (monitoring, lock cleanup, StartApp) je ovde - centralno,
' u kodu - pa se azurira iskljucivo update-om .xlam-a, bez diranja podataka.
' ============================================================

' Application-event hook (samo u ENGINE rezimu).
Private gAppEvents As clsAppEvents

' Instalira oslukivanje otvaranja/zatvaranja workbook-ova (ENGINE rezim).
' Poziva se iz ThisWorkbook.Workbook_Open add-in-a kad se Excel pokrene.
Public Sub InstallAppEvents()
    On Error Resume Next
    Set gAppEvents = New clsAppEvents
    Set gAppEvents.App = Application

    ' Ako je klijentski .xlsx vec otvoren pre nego sto se add-in ucitao,
    ' startuj ga odmah (redosled ucitavanja ume da varira).
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If modContext.IsClientDataBook(wb) Then
            BootFromData wb
            Exit Sub
        End If
    Next wb
End Sub

' Glavna ulazna tacka: vezi klijentski workbook i pokreni aplikaciju.
' Idempotentno: ponovni poziv za vec startovan workbook ne radi nista.
Public Sub BootFromData(ByVal dataWb As Workbook)
    If modContext.IsBoundDataBook(dataWb) Then Exit Sub

    ' Vezi PODATKE pre bilo kakvog pristupa tabelama/config-u.
    modContext.BindDataBook dataWb

    On Error Resume Next
    Monitor_AppOpen "Operator"
    On Error GoTo EH

    StartApp

    ' Stale lock cleanup (prethodna sesija crash-ovala -> lock ostao u SyncControl).
    ' Fail-soft: ako Google nije dostupan, samo se loguje i ide dalje.
    On Error Resume Next
    modStanicaLock.CleanupOrphanedLocks
    On Error GoTo EH

    On Error Resume Next
    Monitor_Event _
        eventType:="VBA_STARTUP_SUCCESS", _
        severity:="INFO", _
        message:="Workbook startup completed", _
        userId:="Operator", _
        moduleName:="modBoot", _
        procedureName:="BootFromData", _
        entityType:="App", _
        entityID:="Startup", _
        correlationId:="VBA-STARTUP"
    On Error GoTo 0

    Exit Sub

EH:
    On Error Resume Next

    Monitor_Error _
        moduleName:="modBoot", _
        procedureName:="BootFromData", _
        entityType:="App", _
        entityID:="Startup", _
        correlationId:="VBA-STARTUP", _
        errorNumber:=Err.Number, _
        errorDescription:=Err.description, _
        errorSource:=Err.SOURCE

    LogErr "modBoot.BootFromData"

    Application.Visible = True
    MsgBox "Greska pri pokretanju aplikacije. Pogledajte log.", vbCritical
End Sub

' Gasenje: release aktivnog stanica lock-a + ShutdownApp + odvezi podatke.
Public Sub ShutdownFromData()
    On Error GoTo EH

    ' Final release aktivne stanica lock-a sa bulk push.
    ' Fail-soft: ako bulk push padne (offline), lock postaje stale i sledeca
    ' sesija ga cisti kroz CleanupOrphanedLocks.
    On Error Resume Next
    If Len(modStanicaLock.GetActiveStanica()) > 0 Then
        Application.StatusBar = "Sinhronizujem unos pre zatvaranja..."
        Application.Cursor = xlWait
        DoEvents

        modStanicaLock.ReleaseStanicaLock modStanicaLock.GetActiveStanica()

        Application.StatusBar = False
        Application.Cursor = xlDefault
    End If
    On Error GoTo EH

    modMain.ShutdownApp
    modContext.UnbindDataBook

    Exit Sub

EH:
    LogErr "modBoot.ShutdownFromData"
End Sub
