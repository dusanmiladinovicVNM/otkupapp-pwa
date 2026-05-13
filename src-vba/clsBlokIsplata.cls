Option Explicit

' ============================================================
' clsBlokIsplata
' Class wrapper za jedan otvoreni otkup blok u pregledu isplata.
' Class-based jer VBA UDT-i ne mogu da budu Variant payload
' (Collection, Variant array, Dictionary).
' Pattern dosledan sa clsTransaction, clsFilterParam.
' ============================================================

Public otkupID As String
Public Datum As Date
Public kooperantID As String
Public kooperantNaziv As String
Public stanicaID As String
Public TekuciRacun As String
Public BrojDokumenta As String
Public UkupanIznos As Double
Public VecIsplaceno As Double
Public OtvorenIznos As Double
Public HasTekuciRacun As Boolean
