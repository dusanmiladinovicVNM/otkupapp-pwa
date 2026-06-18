Attribute VB_Name = "modContext"
Option Explicit

' ============================================================
' modContext - Workbook-Kontext (KOD vs. PODACI)
'
' Posle podele aplikacije na:
'   - OtkupApp.xlam  (zajednicki KOD za sve klijente)
'   - <klijent>.xlsx (PODACI + config, poseban po klijentu)
'
' vazi pravilo:
'   CodeBook = workbook u kome je ovaj kod (sam .xlam)         -> ugradjeni asseti
'   DataBook = klijentski workbook sa tabelama i config-om     -> SVI podaci/config
'
' SVE sto cita/pise podatke ili config MORA ici preko DataBook (ne ThisWorkbook).
'
' Dva rezima rada (isti izvorni kod podrzava oba):
'   1) COMBINED  - jedan .xlsm (kod + podaci zajedno, kao do sada).
'                  DataBook i CodeBook pokazuju na isti workbook.
'   2) ENGINE    - kod je izgradjen kao add-in (.xlam), podaci su u .xlsx.
'                  DataBook pokazuje na otvoreni klijentski .xlsx.
'
' Detekcija rezima: ThisWorkbook.IsAddin (True samo kad je kod .xlam).
' ============================================================

' Naziv add-in fajla pod kojim se .xlam snima (vidi build/Build-Engine.ps1).
Public Const ENGINE_FILE_NAME As String = "OtkupApp.xlam"

' Marker tabela po kojoj prepoznajemo klijentski .xlsx medju otvorenim fajlovima.
Public Const DATA_MARKER_TABLE As String = "tblConfig"

' Vezani klijentski workbook (postavlja modBoot.BootFromData u ENGINE rezimu).
Private gDataBook As Workbook

' --- CodeBook: uvek workbook u kome zivi ovaj kod ---
Public Property Get CodeBook() As Workbook
    Set CodeBook = ThisWorkbook
End Property

' --- DataBook: klijentski .xlsx (u COMBINED rezimu = ThisWorkbook) ---
Public Property Get DataBook() As Workbook
    ' COMBINED rezim: kod i podaci su u istom fajlu.
    If Not ThisWorkbook.IsAddin Then
        Set DataBook = ThisWorkbook
        Exit Property
    End If

    ' ENGINE rezim: vrati vezani klijentski workbook (fallback: pronadji ga).
    If gDataBook Is Nothing Then ResolveDataBook
    If gDataBook Is Nothing Then
        Err.Raise vbObjectError + 8000, "modContext.DataBook", _
            "Klijentski podaci (.xlsx) nisu otvoreni ili nisu prepoznati. " & _
            "Otvorite klijentski fajl koji sadrzi tabelu '" & DATA_MARKER_TABLE & "'."
    End If
    Set DataBook = gDataBook
End Property

' Vezuje klijentski workbook (poziva modBoot pre svakog pristupa podacima).
Public Sub BindDataBook(ByVal wb As Workbook)
    Set gDataBook = wb
End Sub

Public Sub UnbindDataBook()
    Set gDataBook = Nothing
End Sub

' True kad kod radi kao add-in (.xlam).
Public Function IsEngineMode() As Boolean
    IsEngineMode = ThisWorkbook.IsAddin
End Function

' True ako je wb validan klijentski data-fajl (ima marker tabelu, nije add-in).
Public Function IsClientDataBook(ByVal wb As Workbook) As Boolean
    If wb Is Nothing Then Exit Function
    If wb Is ThisWorkbook Then Exit Function
    If wb.IsAddin Then Exit Function
    IsClientDataBook = WorkbookHasTable(wb, DATA_MARKER_TABLE)
End Function

' True ako je wb trenutno vezani DataBook.
Public Function IsBoundDataBook(ByVal wb As Workbook) As Boolean
    If gDataBook Is Nothing Then Exit Function
    IsBoundDataBook = (gDataBook Is wb)
End Function

' Fallback: nadji prvi otvoreni workbook koji ima marker config tabelu.
Private Sub ResolveDataBook()
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If IsClientDataBook(wb) Then
            Set gDataBook = wb
            Exit Sub
        End If
    Next wb
End Sub

Private Function WorkbookHasTable(ByVal wb As Workbook, ByVal tblName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    On Error GoTo done
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tblName, vbTextCompare) = 0 Then
                WorkbookHasTable = True
                Exit Function
            End If
        Next lo
    Next ws
done:
End Function
