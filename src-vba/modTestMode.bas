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

' Seam: podmetnuta greska U AKTIVNOM obradjivacu (v. ObradjivacKvarAkoTest).
Private mObradjivacKvar As String
Private mObradjivacKvarPogodjen As Boolean

Public Sub SetTestMode(ByVal onOff As Boolean)
    m_TestMode = onOff
End Sub

Public Function IsTestMode() As Boolean
    IsTestMode = m_TestMode
End Function

' ============================================================
' Seam: greska U AKTIVNOM obradjivacu (vba_check GARD_U_OBRADJIVACU)
'
' Dok je obradjivac aktivan (posle greske, pre Resume / Exit / On Error GoTo -1),
' procedura NE MOZE da obradi novu gresku -- ni kroz On Error Resume Next: nova
' greska izlazi iz procedure, a originalna se gubi. Ispravka je On Error GoTo -1
' pre garda. Kod pozvane procedure koja i sama guta greske (LogError, Monitor_*)
' ta ispravka ne menja nista vidljivo, pa je bez ove kuke test ne bi mogao da
' razlikuje od njenog odsustva.
'
' Kuka stoji u obradjivacu clana, ODMAH POSLE garda. Nema svoj On Error, pa njena
' greska stize u kontekst pozivaoca: sa GoTo -1 je proguta njegov Resume Next,
' bez njega izlazi iz clana -- i test pada po imenu. Seam samo kvari (testovi.md
' par.4): postavljac odbija rad van test rezima, a kuka ponovo pita IsTestMode().
' Van test rezima i za drugi clan kuka ne radi nista i ne dira Err. Gasi je
' modTest.ResetSeamova.
'
' ObradjivacKvarPogodjen kaze da je kuka STVARNO podigla gresku. Test ga tvrdi
' pre svega ostalog: kuka koja ne stoji u obradjivacu (ili se zove pod drugim
' imenom) ostavlja test zelen i sa GoTo -1 i bez njega -- placebo.
' ============================================================
Public Sub ObradjivacKvarTestSet(ByVal clan As String)
    If Not IsTestMode() Then Exit Sub
    mObradjivacKvar = clan
    mObradjivacKvarPogodjen = False
End Sub

Public Function ObradjivacKvarPogodjen() As Boolean
    ObradjivacKvarPogodjen = mObradjivacKvarPogodjen
End Function

Public Sub ObradjivacKvarAkoTest(ByVal clan As String)
    If Len(mObradjivacKvar) = 0 Then Exit Sub
    If mObradjivacKvar <> clan Then Exit Sub
    If Not IsTestMode() Then Exit Sub
    mObradjivacKvarPogodjen = True
    Err.Raise 76, "modTestMode.ObradjivacKvarAkoTest[" & clan & "]", _
              "podmetnuta greska u aktivnom obradjivacu"
End Sub
