VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlokIsplata1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsBlokIsplata
' Class wrapper za jedan otvoreni otkup blok u pregledu isplata.
' Class-based jer VBA UDT-i ne mogu da budu Variant payload
' (Collection, Variant array, Dictionary).
' Pattern dosledan sa clsTransaction, clsFilterParam.
' ============================================================

Public otkupID As String
Public datum As Date
Public kooperantID As String
Public kooperantNaziv As String
Public stanicaID As String
Public TekuciRacun As String
Public brojDokumenta As String
Public UkupanIznos As Double
Public VecIsplaceno As Double
Public OtvorenIznos As Double
Public HasTekuciRacun As Boolean
Public KooperantAvansSaldo As Double      ' info display: nedodeljeni avans
