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

' Imena tabela koje rollback NIJE uspeo da vrati. Prazno = rollback potpun.
' Nosi se kao stanje jer RollbackTx ne sme da digne gresku -- vidi komentar
' iznad RollbackTx.
Private mNevracene As String


Public Sub BeginTx()
    If mActive Then Err.Raise 5, , "Transaction already active."

    ' Nepotpun rollback NIJE nova normalna transakcija. Dok ova sesija zna da
    ' jedna snapshot-ovana tabela nije vracena, nov upis bi pisao PREKO
    ' nekonzistentnog stanja -- i prvi sledeci commit bi ga zakazao za Save.
    ' Jedini izlaz je zatvaranje bez snimanja i ponovno otvaranje: na disku
    ' stoji stanje pre te transakcije, jer je Save zatvoren (modTxState).
    If Not modTxState.UpisDozvoljen() Then
        Err.Raise vbObjectError + 2700, "clsTransaction.BeginTx", _
                  "UPIS ZATVOREN: prethodni rollback nije bio potpun (" & _
                  modTxState.NevraceneTabeleSesije() & _
                  "). Zatvorite aplikaciju BEZ snimanja i otvorite je ponovo."
    End If

    Set mSnapshots = CreateObject("Scripting.Dictionary")
    mActive = True
    mNevracene = ""

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
    
    ' Capture snapshot summary BEFORE Cleanup destroys mSnapshots reference,
    ' so AutoSave log gets a meaningful source name.
    Dim sourceName As String
    sourceName = BuildAutoSaveSourceName()
    
    CleanUp
    mActive = False
    
    ' AR-002a: zakazi background save (OnTime) umesto sinhronog Save-a.
    ' Stvarni save i dalje izvrsava AutoSaveAfterCommit. Failure NIKAD
    ' ne propaga jer je TX vec commit-ovan u memoriji.
    Call MarkDirtyAndSchedule(sourceName)

    ' Storno cockpit kes je posle izmene zastareo -> oznaci + prezakazi idle warm.
    ' Fail-soft: nikad ne diramo commit-ovanu TX (InvalidateStornoWarm guta greske).
    On Error Resume Next
    InvalidateStornoWarm
    On Error GoTo 0
End Sub

' Rollback je i sam operacija koja moze da padne: RestoreTable dize gresku na
' neuskladjen broj kolona, a GetTable / ListRows.Add / DataBodyRange.Value2 na
' zasticen list ili nestalu tabelu. Dok je petlja stajala na prvom raise-u,
' jedan pad je nosio tri posledice, sve tihe:
'   1. ostale snapshot-ovane tabele su ostajale NEVRACENE (parcijalan rollback);
'   2. CleanUp se nije izvrsio, pa je Application.EnableEvents ostajao False do
'      kraja sesije -- svaki kasniji Workbook event tiho prestane da radi;
'   3. mActive je ostajao True, pa je sledeci BeginTx dizao
'      "Transaction already active."
'
' Primitiv NE DIZE gresku i to je namerno, ne propust. Od 276 poziva .RollbackTx
' u izvoru 182 ne stoje pod On Error Resume Next, a vecina je unutar aktivnog EH
' bloka (obrazac "If Not tx Is Nothing Then tx.RollbackTx"). Nov raise bi tamo
' ZAMENIO originalnu poslovnu gresku, a u testovima pretvorio cleanup u pad.
' Zato cinjenica ide dva puta: kao citljivo stanje (RollbackNepotpun /
' NevraceneTabele) i van trake (lokalni log + Monitor_Critical).
Public Sub RollbackTx()
    If Not mActive Then Exit Sub

    Dim key As Variant
    Dim prvaGreska As String

    For Each key In mSnapshots.keys
        On Error Resume Next
        RestoreTable CStr(key), mSnapshots(key)
        If Err.Number <> 0 Then
            If Len(prvaGreska) = 0 Then
                prvaGreska = CStr(Err.Number) & ": " & Err.description
            End If
            If Len(mNevracene) > 0 Then mNevracene = mNevracene & ", "
            mNevracene = mNevracene & CStr(key)
            Err.Clear
        End If
        On Error GoTo 0
    Next key

    ' CleanUp UVEK: Application state se ne sme ostaviti suspendovan ni kad
    ' restore padne. CleanUp i sam guta greske (On Error Resume Next u njemu).
    CleanUp
    mActive = False

    ' Granica bezbednosti je GLOBALNA, ne na ovom objektu: tx nestane kad
    ' pozivalac izadje iz procedure, a nekonzistentno stanje ostaje. modTxState
    ' zatvara upis i snimanje do reload-a; property-ji ispod ostaju za
    ' pozivaoca koji zeli detalj odmah, ali oni nisu brana.
    If Len(mNevracene) > 0 Then
        modTxState.OznaciNepotpunRollback mNevracene, prvaGreska
        PrijaviNepotpunRollback prvaGreska
    End If
End Sub

' Je li tabela u ovoj transakciji -- citac za AMB-INV-08.
'
' Pisac knjige ne moze da zna da li ga je pozivalac snapshotovao: tx je lokalan
' objekat i nema globalnog "tekuceg tx". Zato ga pisac TRAZI kao argument i pita
' ovo -- invarijanta time prestaje da bude tvrdnja o obliku koda.
'
' FAIL-CLOSED: neaktivna transakcija (posle Commit/Rollback, ili pre BeginTx)
' vraca False. Upis posle commit-a nije pokriven rollback-om, pa "nema snapshot"
' jeste tacan odgovor, ne tehnikalija.
Public Function ImaSnapshot(ByVal tableName As String) As Boolean
    If Not mActive Then Exit Function
    If mSnapshots Is Nothing Then Exit Function
    ImaSnapshot = mSnapshots.Exists(tableName)
End Function

' Nepotpun rollback je cinjenica koju pozivalac mora moci da PROCITA -- signal
' koji ne stigne do citaoca ne postoji. Vazi od zadnjeg RollbackTx-a nad ovom
' instancom; BeginTx ga resetuje.
Public Property Get RollbackNepotpun() As Boolean
    RollbackNepotpun = (Len(mNevracene) > 0)
End Property

' Imena tabela koje nisu vracene, zapetom razdvojena. Prazno = rollback potpun.
Public Property Get NevraceneTabele() As String
    NevraceneTabele = mNevracene
End Property

' Dva kanala, jer ni jedan nije uvek dostupan: Monitor_Critical je no-op kad
' monitoring nije podesen (Monitoring_IsEnabled), a lokalni log radi uvek.
' Oba sama gutaju greske, pa ovaj poziv ne moze da obori rollback.
Private Sub PrijaviNepotpunRollback(ByVal prvaGreska As String)
    Dim msg As String
    msg = "ROLLBACK NEPOTPUN -- nevracene tabele: " & mNevracene
    If Len(prvaGreska) > 0 Then msg = msg & " | prva greska: " & prvaGreska

    On Error Resume Next
    LogError "clsTransaction.RollbackTx", msg
    Monitor_Critical "clsTransaction", "RollbackTx", msg
    On Error GoTo 0
End Sub

Private Sub CleanUp()
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

Private Function BuildAutoSaveSourceName() As String
    ' Builds a source identifier for the AutoSave log line that lists
    ' the tables snapshotted in this transaction. Useful for post-mortem
    ' investigation when something goes wrong.
    '
    ' Example output:
    '   "clsTransaction[tblOtkup,tblNovac,tblAmbalaza]"
    '
    ' Failure to build a name must not block save - returns generic fallback.
    
    On Error GoTo Fallback
    
    If mSnapshots Is Nothing Then
        BuildAutoSaveSourceName = "clsTransaction[no-snapshots]"
        Exit Function
    End If
    
    If mSnapshots.count = 0 Then
        BuildAutoSaveSourceName = "clsTransaction[empty]"
        Exit Function
    End If
    
    Dim names As String
    Dim k As Variant
    Dim count As Long
    count = 0
    
    For Each k In mSnapshots.keys
        If Len(names) > 0 Then names = names & ","
        names = names & CStr(k)
        count = count + 1
        ' Cap at 10 table names to keep log line readable.
        If count >= 10 Then
            names = names & ",..."
            Exit For
        End If
    Next k
    
    BuildAutoSaveSourceName = "clsTransaction[" & names & "]"
    Exit Function
    
Fallback:
    BuildAutoSaveSourceName = "clsTransaction[unknown]"
End Function
