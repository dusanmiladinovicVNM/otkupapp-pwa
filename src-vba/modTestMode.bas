Attribute VB_Name = "modTestMode"
Option Explicit

' ============================================================
' modTestMode
' Globalni test-rezim. Dok je TRUE, kod preskace sve sto ceka operatera:
' MsgBox / InputBox / FileDialog, ali i SetFocus na formi koja nije prikazana.
' U nevidljivom Excelu svaki takav poziv je TRAJNO visenje -- COM poziv ne
' puca, samo stoji dok ga watchdog ne raskloni ili dok se proces ne ubije.
'
' Isti oblik kao modJournaling.SetTestModeQuiet / IsTestModeQuiet (koji gasi
' journal trag za vreme smoke suite-ova). Zaseban modul iz dva razloga:
'   1) ovaj flag nije o journalingu nego o UI-ju -- ne mesa se u tudji modul;
'   2) MORA da postoji i u produkcijskom build-u, jer ga referencira frmOtkup,
'      a modTest.bas se u produkciju ne uvozi (bez ovoga bi frmOtkup kod
'      klijenta prestao da se kompajlira: "Sub or Function not defined").
'
' Postavlja se ISKLJUCIVO iz test modula; produkcioni tok ga nikad ne dira.
' ============================================================

Private m_TestMode As Boolean

Public Sub SetTestMode(ByVal onOff As Boolean)
    m_TestMode = onOff
End Sub

Public Function IsTestMode() As Boolean
    IsTestMode = m_TestMode
End Function
