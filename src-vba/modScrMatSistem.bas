Attribute VB_Name = "modScrMatSistem"
'=====================================================================
' modScrMatSistem - ekran "Podesavanja i alati" (sekcija MATICNI, korak M0).
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run (zamka #19).
'
' ODAKLE DOLAZI: dve stavke grupe "Sistem" iz legacy menija Maticni podaci
' (modMaticniLookups.MaticniSekcijeGrupisano):
'   "Podesavanja" -> frmStammdaten (Tag="Podesavanja") -> modPodesavanja.BuildConfigEditor
'   "Admin"       -> frmStammdaten (Tag="Admin")       -> modAdmin.BuildAdminPanel
'
' STA OVAJ EKRAN JESTE: pokretac, ne editor. Oba panela OSTAJU gde jesu i ne
' menjaju se -- v. docs/UI_MIGRACIJA_KATALOG.md, 24.9. Podesavanja rade, imaju
' grupisanje i nose bezbednosno pravilo (interni i anti-tamper kljucevi se
' namerno ne prikazuju); Admin je spisak od deset radnji, a ne lista. Prelazak
' bilo kog od njih bi bio redizajn onoga sto radi.
'
' ZASTO ONDA EKRAN, A NE SAMO DUGME: ugovorni ekran uvek dobija mrezu
' (LayoutScreenZone bezuslovno slaze zTitle + zGrid), pa ekran bez liste ne
' postoji. Lista alatki je jedina lista koja tu ima smisla, i uz to daje
' maticnoj sekciji BAR JEDAN ekran od prvog dana -- bez njega bi zlatno dugme
' vodilo u sidebar u kom je sve prigusen (v. PostaviSekciju).
'
' IDENTITET REDA: cetvrta kolona je NEVIDLJIVA (prioritet 4) i nosi Tag za
' frmStammdaten. Radnja bira po njoj, ne po rednom broju -- isti obrazac koji
' modScrStorno koristi za GeneracijaID. Ovde je opasnost mala (dva reda), ali
' pravilo koje vazi samo dok je lista kratka nije pravilo.
'
' BRANA: administracija. Ne moze u SCR_OBLAST (to je naziv kolone prava u
' tblKorisnici), pa ide kroz neobavezan Scr_Dozvoljen -- ljuska i dalje ne zna
' nijedan ekran po imenu. Ovo je UI brana, ista koju legacy meni ima u
' MaticniMenu_OnClick; tvrde brane su i dalje u modAdmin.BuildAdminPanel i
' modPodesavanja.BuildConfigEditor i ne diraju se.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRMS_BUILD As String = "v6-ui-196"

' Izvor za log. Otvaranje panela je jedino sto ovaj ekran radi, pa mu svaki
' ishod -- i odbijanje -- ide u log pod istim imenom.
Private Const MOD_TRAG As String = "modScrMatSistem.OtvoriAlatku"

' Kolona koja nosi KLJUC PANELA. JEDAN broj, deljen izmedju opisa kolona,
' punjenja redova i radnje -- da se indeks ne moze razici.
'
' Do M6 je nosila legacy Tag za frmStammdaten. Sada nosi kljuc iz
' modUiPanel.PanelRedovi: ekran vise ne zna nista o legacy formi, a panel se
' otvara u radnoj povrsini nove ljuske (UI_MIGRACIJA_KATALOG 24.21).
Public Const MS_COL_TAG As Long = 4

Private Const MS_ZONA_H As Single = KPI_H

' Test seam za branu. Sme SAMO da je zatvori, nikad da je otvori -- brana koju
' test moze da otkljuca nije brana. Postoji zato sto se u headless runu
' administracija ne moze iskljuciti (MozeAdministraciju je anti-lockout: bez
' AUTH-a svi su admini), pa bi tvrdnja "ljuska postuje branu ekrana" inace
' merila dva puta True i prolazila i kad brane nema.
Private mBranaZatvorenaTest As Boolean

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=MAT_SISTEM|naslov=OTKUI_NAV_MAT_SISTEM|sub=OTKUI_SCRMS_SUB" & _
               "|lista=OTKUI_SCRMS_LISTA|oblik=lista|upis=ne"
End Function

' Dodatna brana ekrana. Vraca se ljusci kroz modUiScreens.ScrDozvoljen, pa
' stavka menija bude prigusena, a ne skrivena: operater koji nije admin treba
' da vidi DA ovo postoji i da mu je zabranjeno, ne da mu nestane.
Public Function Scr_Dozvoljen() As Boolean
    If mBranaZatvorenaTest Then Exit Function
    Scr_Dozvoljen = modAuth.MozeAdministraciju()
End Function

' Jedna lista - prekidaca nema.
Public Function Scr_Liste() As Variant
End Function

Public Function Scr_Lista() As String
    Scr_Lista = "ALATKE"
End Function

Public Function Scr_Cipovi() As String
End Function

Public Function Scr_Radnje() As String
    Scr_Radnje = "otvori:OTKUI_BTN_MS_OTVORI:96:soft:1"
End Function

Public Sub Scr_ResetCache()
    ' Spisak alatki je nepromenljiv - nema izvedenih mapa. Metod postoji zbog
    ' ugovora: ljuska ga zove posle svake promene podataka.
End Sub

'--------------------------------------------------------------- ZONA
Public Sub Scr_Build(ByVal z As Object)
    modUiKit.NewLbl z, "msCap", UCase$(Poruka("OTKUI_SCRMS_LISTA")), PAD, 8, 260, 11, _
                    TS_MICRO, True, C_MUTED, -1
    ' Zasto upozorenje stoji ovde, a ne u dijalogu pri otvaranju: dijalog se
    ' zatvara i zaboravlja, a linija u zoni ne moze da se zaboravi (isti razlog
    ' zbog kog Oporavak vise ne javlja siroticice modalnim dijalogom).
    modUiKit.NewLbl z, "msHint", Poruka("OTKUI_SCRMS_SUB"), PAD, 24, 620, 16, _
                    TS_BODY, False, C_FOREST, -1
    modUiKit.NewLbl z, "msLnB", "", 0, MS_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    On Error Resume Next
    z.Controls("msHint").width = w - 2 * PAD
    z.Controls("msLnB").width = w
    Scr_Layout = MS_ZONA_H
End Function

'-------------------------------------------------------------- REDOVI
' Cetvrta kolona je NEVIDLJIVA (prioritet 4) i nosi Tag za frmStammdaten.
Private Function MsGridCols() As Variant
    MsGridCols = Array( _
        "OTKUI_HDMS_ALATKA||txt|180|1", _
        "OTKUI_HDMS_OPIS||part|0|1", _
        "OTKUI_HDMS_GDE||txt|140|2", _
        "OTKUI_HDMS_ALATKA||txt|0|4")
End Function

' Spisak alatki. Treci clan je KLJUC iz modUiPanel.PanelRedovi -- ne natpis i
' ne legacy Tag, nego identitet po kom registar nalazi graditelja.
Private Function Alatke() As Variant
    Alatke = Array( _
        Array(Poruka("OTKUI_MS_PODESAVANJA"), Poruka("OTKUI_MS_PODESAVANJA_OPIS"), _
              "PODESAVANJA"), _
        Array(Poruka("OTKUI_MS_ADMIN"), Poruka("OTKUI_MS_ADMIN_OPIS"), "ADMIN"))
End Function

' Tagovi alatki koje OVAJ ekran otvara. Javno zato sto je to druga polovina
' tvrdnje "novi UI dostize sve sto stari meni dostize": trinaest sifarnika i
' Korisnici idu kroz modMaticniIzvor.MatKljucIzLegacyTag, a Podesavanja i Admin
' kroz ovaj spisak. Test 160 spaja obe polovine.
Public Function MsAlatkaTagovi() As Variant
    Dim src As Variant, a() As Variant, i As Long
    src = Alatke()
    ReDim a(0 To UBound(src))
    For i = LBound(src) To UBound(src)
        a(i) = CStr(src(i)(2))
    Next i
    MsAlatkaTagovi = a
End Function

Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, outA() As Variant, i As Long, n As Long, hay As String
    On Error GoTo EH
    src = Alatke()
    ReDim outA(1 To UBound(src) + 1, 1 To MS_COL_TAG)
    For i = LBound(src) To UBound(src)
        hay = CStr(src(i)(0)) & "|" & CStr(src(i)(1))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i)(0))
        outA(n, 2) = CStr(src(i)(1))
        outA(n, 3) = Poruka("OTKUI_MS_PROZOR")
        outA(n, MS_COL_TAG) = CStr(src(i)(2))
Sledeci:
    Next i
    If n = 0 Then
        Scr_Rows = Array(MsGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    Scr_Rows = Array(MsGridCols(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrMatSistem.Scr_Rows", Err.description
End Function

'-------------------------------------------------------------- RADNJE
' Ugovor je isti kao modScrPalete.Scr_Event: na USPESNOM izlazu Err.Number mora
' biti 0, pa se Err cisti u OBA smera.
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    If Left$(tag, 4) = "act:" Then Scr_Event = MsAkcija(tag)
    Err.Clear
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrMatSistem.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

' Vraca False uvek: otvaranje panela ne menja redove ove liste, pa mreza nema
' sta da procita ponovo.
Private Function MsAkcija(ByVal tag As String) As Boolean
    Dim p() As String, red As Long, sekTag As String
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then
        LogWarn MOD_TRAG, "Radnja bez rednog broja u tagu: '" & tag & "'."
        Exit Function
    End If
    If p(0) <> "otvori" Then
        LogWarn MOD_TRAG, "Nepoznata radnja '" & p(0) & "'."
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
        Exit Function
    End If
    red = CLng(val(p(1)))
    sekTag = Trim$(CStr(modOtkupUI.GridCell(red, MS_COL_TAG)))
    If Len(sekTag) = 0 Then
        LogWarn MOD_TRAG, "Red " & red & " nema Tag u koloni " & MS_COL_TAG & _
                ". Mreza drzi " & modOtkupUI.GridBrojRedova() & " redova."
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    OtvoriPanel sekTag
End Function

' JEZGRO radnje: otvara BAS taj panel i nista drugo. Odvojeno od reda mreze
' zato sto je to tvrdnja koju M0 nosi -- da se otvara izabrana alatka, a ne
' prva ili susedna. Kroz UI se u headless runu ne moze izmeriti.
'
' Od M6 panel se otvara U RADNOJ POVRSINI nove ljuske, ne kao zaseban prozor:
' sidebar i zaglavlje ostaju, povratak je jedan klik. Otvaranjem upravlja
' modUiPanel -- ovaj ekran zna samo KOJI kljuc je izabran, ne i ko ga gradi.
Public Function OtvoriPanel(ByVal kljucPanela As String) As Boolean
    Dim odgovor As String
    kljucPanela = UCase$(Trim$(kljucPanela))
    If Len(kljucPanela) = 0 Then Exit Function
    LogInfo MOD_TRAG, "Otvaranje panela '" & kljucPanela & "'."

    odgovor = modUiPanel.PanelOtvori(kljucPanela)
    If Len(odgovor) > 0 Then
        modOtkupUI.ShowToast odgovor, True
        Exit Function
    End If
    OtvoriPanel = True
End Function

'------------------------------------------------------------ TEST SEAM
' Spisak alatki i njihovi Tag-ovi, za tvrdnju da red nosi TACAN Tag. Klik kroz
' formu se u harnessu ne moze odigrati, a OtvoriPanel prikazuje formu -- pa se
' meri ono sto prethodi: da li red i kolona identiteta nose ono sto treba.
' Zatvara branu za jednu tvrdnju. Otvaranje ide iskljucivo kroz False, koji
' vraca normalno ponasanje -- ne postoji vrednost koja branu zaobilazi.
Public Sub Scr_MsBranaZatvoriTest(ByVal zatvori As Boolean)
    mBranaZatvorenaTest = zatvori
End Sub

Public Function Scr_MsTagZaRedTest(ByVal red As Long) As String
    Dim src As Variant
    src = Alatke()
    If red < 1 Or red > UBound(src) + 1 Then Exit Function
    Scr_MsTagZaRedTest = CStr(src(red - 1)(2))
End Function
