VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsTestResult -- rezultat JEDNE suite, u obliku koji ume da procita i masina.
'
' Postoji zato sto je do sada svaki test modul nosio sopstveni mPass/mFail i
' ispisivao ih u Immediate prozor. Immediate prozor ne moze da se procita preko
' COM-a, pa runner nije imao nacina da sazna da je Banka suite sa 189 provera
' tiho spala na 120 -- videla bi se samo boja: zeleno.
'
' Jedan red teksta po suite-u (modTestRunner ih upisuje u suite_results.txt
' pored sveske), a tools/run_vba.py ih poredi sa `min_asserts` iz
' tests/suite_manifest.json.
'
' NotRun je namerno zasebno polje, a ne "0 provera": suite koja se NIJE
' pokrenula i suite koja je prosla bez ijedne provere izgledaju isto u brojevima,
' a znace suprotno.
' ============================================================

Public SuiteId As String
Public PassCount As Long
Public FailCount As Long
Public NotRun As Boolean
Public Message As String
Public Seconds As Double

' Format mora da prati read_suite_counts() u tools/run_vba.py:
'   SUITE|<id>|<pass>|<fail>|<notrun 0/1>|<sekunde>|<poruka>
' Uspravna crta je razdvajac, pa se iz poruke uklanja (poruke dolaze iz
' Err.Description i imena testova, gde crta nema semantiku).
Public Function ToLine() As String
    ToLine = "SUITE|" & SuiteId & "|" & CStr(PassCount) & "|" & CStr(FailCount) & _
             "|" & IIf(NotRun, "1", "0") & "|" & Format$(Seconds, "0.0") & _
             "|" & Replace$(Replace$(Message, "|", "/"), vbLf, " ")
End Function
