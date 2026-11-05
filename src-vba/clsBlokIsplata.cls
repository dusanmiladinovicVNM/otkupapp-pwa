VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlokIsplata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' clsBlokIsplata
' Class wrapper za jedan otvoreni otkup blok u pregledu isplata.
'
' Od F4 doplate isti oblik nosi i otvoreno KNJIZNO ODOBRENJE: otkupID je tada
' KLJUC OBAVEZE (DoplataID), a doplataID <> "" je razlikovac. Mape naloga
' (korpa, zadati iznosi, saldo kapija) rade nad kljucem i ne razlikuju ih;
' radnja koja vazi samo za otkup (primena avansa) gleda doplataID.
' Od DOP-ODL-27 i SPECIFIKACIJU doplate: otkupID je tada broj specifikacije,
' specifikacija <> "" je razlikovac, a njeni dokumenti se ne nude posebno.
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
Public IsplatitiIznos As Double           ' finalni iznos za nalog/specifikaciju (operater unos ili OtvorenIznos)
Public doplataID As String                ' KO (F4 doplate): DoplataID; "" = otkupni blok
Public specifikacija As String            ' specifikacija doplate (DOP-ODL-27): broj; "" = blok ili KO
