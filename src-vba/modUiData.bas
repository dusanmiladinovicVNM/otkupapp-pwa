Attribute VB_Name = "modUiData"
'=====================================================================
' modUiData - PRISTUP TABELAMA za novi UI (faza S4b).
'
' Kes pune tabele i cetiri citaca celije. Nista vise: ni jedno ime ekrana,
' ni jedno ime kolone, nijedno poslovno pravilo.
'
' Postoji zato sto od S4b i ljuska i ekranski moduli citaju iste tabele:
' ljuska za KPI i combo-e, ekran za svoje redove. Da je ostalo u ljusci,
' ekran bi morao da zove njeno privatno telo; da je otislo u ekran, ljuska
' bi zavisila od ekrana. Zajednicki sloj resava oba.
'
' Kes se prazni pri gradnji ekrana i posle upisa (modOtkupUI.RefreshFromData).
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UIDATA_BUILD As String = "v6-ui-107"

Private mCache As Object

Public Sub ResetCache()
    Set mCache = Nothing
End Sub

' KES NE SME DA MEMOISE NEUSPEH.
'
' Zatecen incident: operater je gledao PRAZNE liste za svaki tip dokumenta, bez
' ijedne greske, iako je tblOtkup pun. Dijagnostika je pokazala tacno to:
'
'     IsArray(CachedTable("tblOtkup"))  -> False
'     IsArray(GetTableData("tblOtkup")) -> True
'
' Tabela je pri PRVOM citanju bila prazna (podaci stizu posle -- sync, uvoz,
' legacy forma), pa je Empty ostao u kesu do kraja sesije. ResetCache se zove
' samo pri gradnji ekrana i posle upisa KROZ NOVI UI, a nijedan od tih puteva
' nije prosao.
'
' Zato: kesira se samo USPEH. Neuspeh je "ne znam jos", ne "nema nista" -- pa se
' sledeci poziv ponovo pita. Cena je jedan promasen sken po tabeli koja je
' zaista prazna; korist je da podaci koji stignu kasnije budu vidljivi.
Public Function CachedTable(ByVal tblName As String) As Variant
    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")
    If mCache.Exists(tblName) Then
        CachedTable = mCache(tblName)
        Exit Function
    End If
    Dim src As Variant
    src = GetTableData(tblName)
    If IsArray(src) Then mCache(tblName) = src
    CachedTable = src
End Function

' TEST SEAM: da li je kljuc u kesu. Postoji zato sto se "neuspeh se ne kesira"
' ne moze izmeriti kroz vracenu vrednost -- ona je Empty u oba slucaja.
Public Function KesImaKljuc(ByVal tblName As String) As Boolean
    If mCache Is Nothing Then Exit Function
    KesImaKljuc = mCache.Exists(tblName)
End Function

' Je li tabela uopste CITLJIVA (postoji kao ListObject). GetTableData vraca Empty
' i za praznu i za nepostojecu tabelu, a to su dva razlicita ishoda: prazna je
' normalno stanje, nepostojeca je greska koja mora da se vidi.
Public Function TabelaCitljiva(ByVal tblName As String) As Boolean
    On Error Resume Next
    TabelaCitljiva = Not (GetTable(tblName) Is Nothing)
End Function

Public Function ColIdx(ByVal tblName As String, ByVal colName As String) As Long
    If Len(colName) = 0 Then Exit Function
    On Error Resume Next
    ColIdx = GetColumnIndex(tblName, colName)
End Function

Public Function CellS(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsEmpty(v) Then Exit Function
    CellS = Trim$(CStr(v))
End Function

Public Function CellD(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then CellD = CDbl(v)
End Function

Public Function CellDate(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then
        CellDate = Int(CDbl(v))
    ElseIf IsDate(v) Then
        CellDate = Int(CDbl(CDate(v)))
    End If
End Function
