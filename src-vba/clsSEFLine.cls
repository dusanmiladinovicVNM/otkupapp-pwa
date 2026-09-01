VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSEFLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public prijemnicaID As String
Public brojPrijemnice As String
' GP grana (R2, revizija #248): stavka fakture gotove robe nosi
' preradu umesto prijemnice -- tacno jedan od parova je popunjen.
Public preradaID As String
Public brojPrerade As String
Public naziv As String
Public kolicina As Double
Public cena As Double
Public klasa As String
Public neto As Double
Public pdv As Double
Public iznos As Double
