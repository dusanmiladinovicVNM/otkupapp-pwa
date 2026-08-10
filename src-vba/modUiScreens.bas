Attribute VB_Name = "modUiScreens"
'=====================================================================
' modUiScreens - REGISTAR EKRANA i ugovor prema njima (faza S3a).
'
' Jedno mesto koje zna koji ekrani postoje, kako se zovu, u koju grupu
' sidebara idu i koja im je oblast za dozvole. Sidebar se gradi odavde,
' ne iz nabrojanih nizova u BuildNav.
'
' KLJUCNO - LJUSKA NE SME DA ZNA NIJEDAN EKRAN PO IMENU. Svi pozivi ka
' ekranskom modulu idu kroz Application.Run, dakle KASNO VEZANO. Da je
' vezivanje rano, klijent kome jedan ekranski modul nedostaje ne bi se
' kompajlirao i pao bi ceo StartApp - ista klasa greske kao zamka #19 u
' CLAUDE.md. Ovako ekran koji nedostaje samo bude prikazan kao neaktivan.
'
' UGOVOR EKRANA (sve je opciono osim Scr_Meta):
'   Scr_Meta()                 -> opis ekrana; sluzi i kao provera da modul
'                                 postoji i da je ekran spreman
'   Scr_Build(z)               -> izgradi kontrole u svoju zonu (jednom)
'   Scr_Layout(z, w, h)        -> rasporedi; vraca zauzetu visinu
'   Scr_Grid()                 -> deskriptor mreze (tabela, kolone, cipovi)
'   Scr_Event(tag, ev)         -> obradi klik; True ako je obradio
'   Scr_Save()                 -> upisi; "" ako je proslo, inace greska
'
' Ugovor danas ispunjavaju: modScrDokumenti (Scr_Meta; ljuska ga jos crta
' po starom) i modScrPalete (Scr_Meta/Scr_Build/Scr_Layout/Scr_Event - prvi
' ekran koji ljuska crta ISKLJUCIVO kroz ugovor). Ostali su u registru sa
' modulom koji jos ne postoji - sidebar ih prikazuje prigusene.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UISCR_BUILD As String = "v6-ui-74"

' Redosled polja u redu registra
Public Const SCR_KLJUC   As Long = 0
Public Const SCR_MODUL   As Long = 1
Public Const SCR_NASLOV  As Long = 2
Public Const SCR_IKONICA As Long = 3
Public Const SCR_GRUPA   As Long = 4
Public Const SCR_OBLAST  As Long = 5

' Kes odgovora "da li modul postoji" - Application.Run na nepostojeci modul
' baca gresku, a to je skupo raditi pri svakom crtanju sidebara.
Private mHas As Object

'------------------------------------------------------------ REGISTAR
' kljuc | modul | naslov (kljuc kataloga) | MDL2 kod | grupa | oblast
Public Function ScrRows() As Variant
    Dim c As Collection: Set c = New Collection
    c.Add "DOKUMENTI|modScrDokumenti|OTKUI_NAV_UNOS|" & IC_OTKUP & _
          "|OPERACIJE|" & OBL_DOKUMENTA
    c.Add "PALETE|modScrPalete|OTKUI_NAV_PALETE|" & IC_PALETE & _
          "|OPERACIJE|" & OBL_PALETE
    c.Add "AGRO|modScrAgro|OTKUI_NAV_AGRO|" & IC_AGRO & _
          "|OPERACIJE|" & OBL_AGROHEMIJA
    c.Add "FAKTURE|modScrFakture|OTKUI_NAV_FAKT|" & IC_FAKT & _
          "|FINANSIJE|" & OBL_FAKTURISANJE
    c.Add "BANKA_UVOZ|modScrBankaUvoz|OTKUI_NAV_BANKA_UVOZ|" & IC_UVOZ & _
          "|FINANSIJE|" & OBL_BANKA
    c.Add "BANKA_NALOZI|modScrBankaNalozi|OTKUI_NAV_BANKA_NALOZI|" & IC_NALOZI & _
          "|FINANSIJE|" & OBL_BANKA
    c.Add "MARZA|modScrMarza|OTKUI_NAV_MARZA|" & IC_MARZA & _
          "|FINANSIJE|" & OBL_MARZA
    c.Add "IZVESTAJI|modScrIzvestaji|OTKUI_NAV_IZVESTAJI|" & IC_IZVEST & _
          "|ANALITIKA|" & OBL_IZVESTAJI
    c.Add "SLEDLJIVOST|modScrSledljivost|OTKUI_NAV_SLEDLJIVOST|" & IC_SLEDLJ & _
          "|ANALITIKA|" & OBL_SLEDLJIVOST

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    ScrRows = a
End Function

' Grupe sidebara, redom. Naslov grupe je kljuc kataloga.
Public Function ScrGroups() As Variant
    ScrGroups = Array("OPERACIJE|OTKUI_NAVG_OPERACIJE", _
                      "FINANSIJE|OTKUI_NAVG_FINANSIJE", _
                      "ANALITIKA|OTKUI_NAVG_ANALITIKA")
End Function

Public Function ScrField(ByVal row As String, ByVal idx As Long) As String
    Dim p() As String
    p = Split(row, "|")
    If idx > UBound(p) Then Exit Function
    ScrField = p(idx)
End Function

Public Function ScrRowByKey(ByVal kljuc As String) As String
    Dim r As Variant
    For Each r In ScrRows()
        If ScrField(CStr(r), SCR_KLJUC) = kljuc Then
            ScrRowByKey = CStr(r)
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- STANJE
' Postoji li modul ekrana i odgovara li na ugovor. Provera je namerno
' pokusaj poziva, ne pretraga po projektu: VBComponents trazi programski
' pristup VBA projektu, koji na korisnickoj masini cesto nije ukljucen.
Public Function ScrPostoji(ByVal kljuc As String) As Boolean
    Dim modul As String, dummy As Variant
    If mHas Is Nothing Then Set mHas = CreateObject("Scripting.Dictionary")
    If mHas.Exists(kljuc) Then
        ScrPostoji = mHas(kljuc)
        Exit Function
    End If
    modul = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(modul) = 0 Then Exit Function
    On Error Resume Next
    dummy = Application.Run(modul & ".Scr_Meta")
    ScrPostoji = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    mHas(kljuc) = ScrPostoji
End Function

' Ekran je dostupan ako postoji I ako korisnik ima pravo na njegovu oblast.
Public Function ScrDozvoljen(ByVal kljuc As String) As Boolean
    Dim obl As String
    On Error Resume Next
    obl = ScrField(ScrRowByKey(kljuc), SCR_OBLAST)
    If Len(obl) = 0 Then
        ScrDozvoljen = True
    Else
        ScrDozvoljen = modAuth.KorisnikImaPravo(obl)
    End If
End Function

Public Function ScrAktivan(ByVal kljuc As String) As Boolean
    ScrAktivan = ScrPostoji(kljuc)
    If ScrAktivan Then ScrAktivan = ScrDozvoljen(kljuc)
End Function

' Kes se prazni pri gradnji ekrana - posle self-update-a moduli koji nisu
' postojali mogu da postoje.
Public Sub ScrResetCache()
    Set mHas = Nothing
End Sub

'--------------------------------------------------------------- POZIV
' Omotaci ugovora. Svaki gura gresku u "nije uspelo" umesto da je propusti
' u ljusku: ekran koji padne ne sme da obori aplikaciju.
Public Function ScrMeta(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrMeta = CStr(Application.Run(m & ".Scr_Meta"))
End Function

Public Function ScrBuild(ByVal kljuc As String, ByVal z As Object) As Boolean
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Application.Run m & ".Scr_Build", z
    ScrBuild = (Err.Number = 0)
End Function

Public Function ScrLayout(ByVal kljuc As String, ByVal z As Object, _
                          ByVal w As Single, ByVal h As Single) As Single
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrLayout = CSng(Application.Run(m & ".Scr_Layout", z, w, h))
End Function

Public Function ScrGrid(ByVal kljuc As String) As Variant
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrGrid = Application.Run(m & ".Scr_Grid")
End Function

Public Function ScrEvent(ByVal kljuc As String, ByVal tag As String, _
                         ByVal ev As String) As Boolean
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrEvent = CBool(Application.Run(m & ".Scr_Event", tag, ev))
End Function

Public Function ScrSave(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrSave = CStr(Application.Run(m & ".Scr_Save"))
End Function
