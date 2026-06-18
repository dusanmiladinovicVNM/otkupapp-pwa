VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNalogIzlaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsNalogIzlaz — jedna IZLAZNA linija radnog naloga (data holder).
' Stil kao clsBlokIsplata / clsFilterParam: gola javna polja, bez logike.
' Koristi modRadniNalog.SaveRadniNalog_TX (kolekcija ovih objekata = izlazi).
'
' TipIzlaza:
'   RN_IZLAZ_PALETA    -> nova paleta u lager (Vrsta/Sorta/Frakcija/TipAmbalaze/
'                         Kapacitet/BrojGajbica/NetoKg)
'   RN_IZLAZ_PAKOVANJE -> finalno pakovanje (Kutije/Kese/NetoKg)
'   RN_IZLAZ_OTPAD     -> merena otpadna frakcija (NetoKg)
' ============================================================

Public TipIzlaza As String     ' RN_IZLAZ_PALETA | RN_IZLAZ_PAKOVANJE | RN_IZLAZ_OTPAD
Public Frakcija As String      ' Klasa I / II / izvoz / otpad (oznaka frakcije)
Public VrstaVoca As String
Public SortaVoca As String
Public TipAmbalaze As String
Public Kapacitet As Long
Public BrojGajbica As Long
Public Kutije As Long
Public Kese As Long
Public NetoKg As Double
