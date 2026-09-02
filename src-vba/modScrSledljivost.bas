Attribute VB_Name = "modScrSledljivost"
'=====================================================================
' modScrSledljivost - ekran "Sledljivost" (v6-ui-187). Faza E.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run (zamka #19).
' Red u registru (modUiScreens.ScrRows) je postojao od S3a -- stavka menija se
' do sada crtala prigusena jer modula nije bilo. Registar se NE dira.
'
' MERILO: LANAC KOJI SE NE IZMISLJA. Ekran odgovara na dva pitanja -- "od ovog
' otkupnog lista, gde je roba zavrsila?" (napred) i "od ove fakture/prijemnice,
' od kojih kooperanata i parcela je roba dosla?" (nazad). Oba pitanja odgovara
' ISTO zrno: jedan red = jedan otkupni list sa razresenim karikama kao
' kolonama; smer daje pretraga ljuske (haystack nosi SVE brojeve lanca), ne
' poseban prekidac. Zato ekran nema entitet combo ni tip/rezim -- nema ni S1
' klase zamki (PopIndex trovanje panela).
'
' ODAKLE DOLAZI: frmSledljivost (trag po zbirnoj + povezivanje) nad
' modSledljivost.TraceByZbirna. Racuni lanca su za novi UI izdvojeni u
' modIzvestaj (ReportSledljivostLanac / ReportSledljivostProblemi -- obrazac
' ReportPrijemniceKupca): otkup->otpremnica ide iskljucivo po OtpremnicaID,
' otpremnica/prijemnica->zbirna kroz ISTO vlasnicko pravilo kao
' ReportOtkupRobaOM (BuildManjakDict + PrijemZaZbirnu, fail-closed),
' prijemnica->faktura po denorm FakturaID koloni. Ekran je PRIKAZ nad ta dva
' racuna: nijedno pravilo razresenja ne zivi ovde.
'
' POVEZIVANJE (auto-link, rucno vezivanje otkupa za otpremnicu) NE ulazi u
' v1: to je upis sa izborom kandidata i trazi svoj UX; legacy frmSledljivost
' ostaje operativna (par. 5/Faza B -- dve kopije zive namerno). Lista
' NEPOTPUNI je PREGLED tog posla, ne alat.
'
' FAIL-CLOSED (merilo #2 zadatka): nepotpun ili visesmislen lanac se
' prikazuje kao takav (kolona OZNAKA + lista NEPOTPUNI); kg koji "nestaje"
' niz lanac je vidljiva razlika sa oznakom, nikad precutana. Tiho
' premoscenje veze ne postoji ni ovde ni u racunima.
'
' KES SNIMKA (par. 22.9/N7 + 23.10/R1 od prvog dana): JEDNO punjenje
' (lanac + problemi zajedno) po kljucu konteksta (od|do) puni SVE TRI liste;
' prelaz liste, cip i svaki otkucaj pretrage su re-filter nad snimkom -- nula
' citanja tabela. Invalidira ga Scr_ResetCache i generacija podataka
' (modUiData.DataGeneracija -- upis sa DRUGOG ekrana). mSnimakPunjenja broji
' punjenja; Diag_SlRedovi (Alt+F8) od prvog dana.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRSL_BUILD As String = "v6-ui-189"

' Visina zone: red polja + hint + red dugmadi (isti raspored kao Izvestaji).
Private Const SL_ZONA_H   As Single = 148

Private Const SL_Y_CAP    As Single = 6
Private Const SL_Y_LBL    As Single = 48
Private Const SL_Y_HINT   As Single = 98
Private Const SL_Y_BTN    As Single = 116
Private Const SL_BTN_H    As Single = 24
Private Const SL_KPI_W    As Single = 150

' Desna traka zone nosi DETALJ LANCA izabranog reda (karika po karika sa kg)
' -- isti raspored kao detalj traka Izvestaja.
Private Const SL_DET_W    As Single = 320
Private Const SL_DET_N    As Long = 6
Private Const SL_POLJA_MIN As Single = 470

' Kljucevi tri liste.
Private Const SL_LANAC As String = "LANAC"
Private Const SL_PARC As String = "PARCELE"
Private Const SL_NEP As String = "NEPOTPUNI"

' Granice "bez granice" opsega (isto kao Izvestaji).
Private Const SL_DAT_MIN As Long = 2          ' 1.1.1900
Private Const SL_DAT_MAX As Long = 2958465    ' 31.12.9999

Private Const SL_SNIMAK_KAPA As Long = 8

'--------------------------------------------------------------- STANJE
Private mLista As String

' KES SNIMKA -- v. zaglavlje. Mapa kljuc konteksta (od|do) -> Array(lanac,
' problemi). Jedno punjenje = OBA read-modela, pa je prelaz na svaku listu
' bez ponovnog citanja.
Private mSnimci As Object
Private mSnimakKljuc As String      ' kljuc POSLEDNJEG prikaza (kontekst stampe)
Private mSnimakPunjenja As Long
Private mSnimakGen As Long

' Kontekst STVARNO ucitanih podataka (naslov stampe i hint -- AUD-024).
Private mCtxOd As Double            ' 0 = bez granice
Private mCtxDo As Double

' KPI zone: potpuni lanci / nepotpune karike. Empty = jos nije citano.
Private mKpiPotpun As Variant
Private mKpiProblemi As Variant

' Hint kljuc poslednjeg citanja (postavlja RedoviZaListu, cita OsveziHint).
Private mHintKljuc As String

' Izabran red: OtkupID za detalj traku i "Lanac (PDF)"; broj zbirne
' izabranog reda za sablon-PDF (smoke krug 2).
Private mIzabranOtkupID As String
Private mIzabranaZbirna As String
Private mDetalj As Variant

' Polje izbora dokumenta sledljivosti (smoke krug 3b): guard punjenja
' (upis u kontrolu okida Change) + kljuc poslednjeg punjenja comba
' (snimak-kljuc # punjenja # generacija -- da se ne puni iznova bez potrebe).
Private mDokFill As Boolean
Private mCmbDokKljuc As String

' Otkup cije kandidate trenutno nosi polje "Otpremnica za povezivanje"
' (krug 4 S9) -- odbrana od ustajalog para red/polje.
Private mPovOtkupID As String

' Kontekst koji je postavio TEST (zone u testu nema).
Private mTestOd As Double
Private mTestDo As Double

' Poslednji poziv Scr_Rows -- SAMO za Diag_SlRedovi (N7 obrazac).
Private mDiagFilter As String
Private mDiagQ As String
Private mDiagN As Long

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=SLEDLJIVOST|naslov=OTKUI_NAV_SLEDLJIVOST|sub=OTKUI_SCRSL_SUB" & _
               "|lista=OTKUI_SCRSL_LISTA|oblik=zona+mreza|upis=zona"
End Function

' Tri liste, sve tri UVEK dostupne: ekran nema tip/rezim, pa matrica
' nedostupnosti ne postoji -- prazna kombinacija objasnjava hint, ne
' nestajanje taba.
Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        SL_LANAC & "|OTKUI_SEG_SL_LANAC|OTKUI_GRID_TITLE_SL_LANAC|54", _
        SL_PARC & "|OTKUI_SEG_SL_PARC|OTKUI_GRID_TITLE_SL_PARC|76", _
        SL_NEP & "|OTKUI_SEG_SL_NEP|OTKUI_GRID_TITLE_SL_NEP|76")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = SL_LANAC
    Scr_Lista = mLista
End Function

Public Function Scr_Cipovi() As String
    Scr_Cipovi = SlCipoviZaListu(Scr_Lista())
End Function

' Cipovi PO KLJUCU LISTE -- prvi je svuda najsiri ("sve"). Filteri su
' prirodan podatak reda (oznaka lanca / parcela / klasa problema), ne novo
' poslovno pravilo.
Public Function SlCipoviZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case SL_LANAC
            SlCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "potpun:OTKUI_CIPSL_POTPUN:86|" & _
                              "nepotpun:OTKUI_CIPSL_NEPOTPUN:76"
        Case SL_PARC
            SlCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "bezpar:OTKUI_CIPSL_BEZPARC:82"
        Case SL_NEP
            SlCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                              "veze:OTKUI_CIPSL_VEZE:52|" & _
                              "prijem:OTKUI_CIPSL_PRIJEM:60|" & _
                              "fakture:OTKUI_CIPSL_FAKTURE:62|" & _
                              "kg:OTKUI_CIPSL_KG:72"
    End Select
End Function

' PRAVILO CIPA LANCA: "potpun" = prazna oznaka, "nepotpun" = bilo koja
' oznaka (ukljucujuci kg razliku -- lanac koji curi nije potpun). Nepoznat i
' prazan kljuc pustaju sve.
Public Function SlCipLanac(ByVal filter As String, ByVal oznaka As String) As Boolean
    Select Case filter
        Case "potpun":   SlCipLanac = (Len(Trim$(oznaka)) = 0)
        Case "nepotpun": SlCipLanac = (Len(Trim$(oznaka)) > 0)
        Case Else:       SlCipLanac = True
    End Select
End Function

Public Function SlCipParcele(ByVal filter As String, ByVal parcelaID As String) As Boolean
    Select Case filter
        Case "bezpar": SlCipParcele = (Len(Trim$(parcelaID)) = 0)
        Case Else:     SlCipParcele = True
    End Select
End Function

' PRAVILO CIPA PROBLEMA: grupe klasa. veze = karike veza (bez otpremnice /
' neusaglasena / bez zbirne); prijem = vlasnistvo i prijem (dvosmislen broj /
' bez prijema); fakture = nefakturisane; kg = kg razlike.
Public Function SlCipProblemi(ByVal filter As String, ByVal klasa As String) As Boolean
    Select Case filter
        Case "veze"
            SlCipProblemi = (klasa = SLEDP_BEZ_OTPREMNICE Or klasa = SLEDP_VEZA _
                             Or klasa = SLEDP_BEZ_ZBIRNE)
        Case "prijem"
            SlCipProblemi = (klasa = SLEDP_BROJ_DVOSMISLEN Or klasa = SLEDP_BEZ_PRIJEMA)
        Case "fakture"
            SlCipProblemi = (klasa = SLEDP_FAK_NEISPRAVNA)
        Case "kg"
            SlCipProblemi = (klasa = SLEDP_KG_RAZLIKA)
        Case Else
            SlCipProblemi = True
    End Select
End Function

' Radnja "Stampaj dokument" postoji na sve tri liste, ali red bez dokumenta
' ODBIJA porukom (zbirna nema svoju stampu -- legacy Case Else obrazac).
' NEPOTPUNI nosi i "Povezi..." (smoke krug 2: povezivanje je iz novog UI-ja
' bilo nedostupno) -- radi SAMO nad redom klase OTKUP-BEZ-OTPREMNICE,
' ostali odbijaju porukom; upis ide kroz ReassignOtkupToOtpremnica_TX.
Public Function Scr_Radnje() As String
    Scr_Radnje = SlRadnjeZaListu(Scr_Lista())
End Function

Public Function SlRadnjeZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case SL_NEP
            SlRadnjeZaListu = "sledprint:OTKUI_BTN_IZ_STAMPAJDOK:132:soft:1|" & _
                              "sledpovezi:OTKUI_BTN_SL_POVEZI:86:soft:1"
        Case SL_LANAC, SL_PARC
            SlRadnjeZaListu = "sledprint:OTKUI_BTN_IZ_STAMPAJDOK:132:soft:1"
    End Select
End Function

' Ekran je read-only pregled: nista ovde ne CEKA operatera (popravke rade
' legacy frmSledljivost i ekran Oporavak), pa je brojac 0 kao na Izvestajima
' -- ne izmislja se brojka da bi je bilo. Broj nepotpunih karika je KPI zone.
Public Function Scr_Brojac() As Long
    Scr_Brojac = 0
End Function

Public Sub Scr_ResetCache()
    Set mSnimci = Nothing
    mKpiPotpun = Empty
    mKpiProblemi = Empty
    mCmbDokKljuc = ""
End Sub

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a: LogError pocinje sa On Error Resume Next,
    ' a svaka On Error naredba brise Err.
    errDesc = Err.description
    LogErr "modScrSledljivost.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

'=====================================================================
' KLIKOVI
'=====================================================================
Private Function ObradiKlik(ByVal tag As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        ObradiKlik = True
        Exit Function
    End If

    ' Izbor reda puni DETALJ TRAKU (pun lanac karika po karika) -- ne menja
    ' podatke, ljuska nista ne osvezava. Dvoklik NAMERNO ne radi nista
    ' (jedina radnja je stampa -- promasen dvoklik koji pokrene PDF je gori
    ' od nikakvog; par. 9.5).
    If Left$(tag, 4) = "row:" Then
        OsveziDetalj CLng(val(Mid$(tag, 5)))
        Exit Function
    End If
    If Left$(tag, 4) = "dbl:" Then Exit Function

    ' Promena datumskih polja: refresh SAMO kad se RAZRESENA granica promeni
    ' (DatGranica pravilo -- nepotpun unos "21." nije promena).
    If Left$(tag, 4) = "chg:" Then
        Select Case Mid$(tag, 5)
            Case "scrSlOdT", "scrSlDoT": OpsegPromenjen
            ' Kucanje u polja izbora NE obradjuje ekran: suzavanje radi
            ' ljuskin panel (PopFromTyping/PopIndex -- podniz nad
            ' prikazom). Ekranski filter je bio DUPLO suzavanje i 2N COM
            ' poziva po slovu (krug 6 S13).
        End Select
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(Mid$(tag, 5))
        Exit Function
    End If

    Select Case tag
        Case "scrSlPrint": StampajIzvestaj
        Case "scrSlLanac": StampajLanacIzabranog
        Case "scrSlSab":   StampajSledljivostReda
        Case "scrSlAuto":  ObradiKlik = AutoPovezi()
    End Select
End Function

Private Sub OpsegPromenjen()
    Dim odN As Double, doN As Double
    Dim odTxt As String, doTxt As String
    OpsegPolja odTxt, doTxt
    odN = SlDatGranica(odTxt)
    doN = SlDatGranica(doTxt)
    If odN = mCtxOd And doN = mCtxDo Then Exit Sub
    ' Nepotpun unos (tekst ima, granice nema) ne dira prikaz.
    If odN = 0 And Len(Trim$(odTxt)) > 0 And doN = mCtxDo Then Exit Sub
    If doN = 0 And Len(Trim$(doTxt)) > 0 And odN = mCtxOd Then Exit Sub
    modOtkupUI.RefreshFromData
End Sub

Private Function RadnjaNadRedom(ByVal spec As String) As Boolean
    Dim p() As String, red As Long
    p = Split(spec, ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    Select Case p(0)
        Case "sledprint"
            StampajDokumentReda red
            ' Stampa ne menja podatke -- mreza se ne osvezava.
        Case "sledpovezi"
            ' Povratna vrednost True = ljuska zove RefreshFromData, pa upis
            ' odmah osvezi liste (i red nestane iz NEPOTPUNI).
            RadnjaNadRedom = PoveziRed(red)
    End Select
End Function

' RUCNO povezivanje (smoke krug 2; UI iz kruga 4 S9 -- bez InputBox-a).
' Radi SAMO nad redom klase OTKUP-BEZ-OTPREMNICE (klasa-kod iz prenosne
' kolone 9); svaki drugi red ODBIJA porukom. Kandidate NUDI polje
' "Otpremnica za povezivanje" (puni ga izbor reda, legacy pravilo ista
' stanica + isti datum); izbor se RAZRESAVA iz liste -- delimican tekst
' odbija. Upis kroz ReassignOtkupToOtpremnica_TX -- kapije cilja
' (postoji, nije storniran) ostaju u writeru.
Private Function PoveziRed(ByVal red As Long) As Boolean
    Dim klasaKod As String, dokTip As String, otkupID As String
    Dim izbor As String, otpID As String, prikaz As String
    Dim CB As Object

    klasaKod = NzS(modOtkupUI.GridCell(red, 9))
    dokTip = NzS(modOtkupUI.GridCell(red, 7))
    otkupID = NzS(modOtkupUI.GridCell(red, 8))
    If klasaKod <> SLEDP_BEZ_OTPREMNICE Or dokTip <> DOK_TIP_OTKUP _
       Or Len(otkupID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_SAMO_NEPOVEZANI"), True
        Exit Function
    End If

    ' Polje mora da nosi kandidate BAS OVOG otkupa (odbrana od ustajalog
    ' para red/polje -- radnja moze stici i bez novog izbora reda).
    If mPovOtkupID <> otkupID Then NapuniPovKandidate otkupID
    Set CB = Kontrola("scrSlPov")
    If CB Is Nothing Then Exit Function
    If CB.ListCount = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_NEMA_KANDIDATA"), True
        Exit Function
    End If

    izbor = IzabraniIzComba("scrSlPov")
    If Len(izbor) = 0 Then
        ' Uz poruku se ODMAH otvara i ponuda (krug 5) -- ali LJUSKIN
        ' panel kroz front door (UiEvent/Drop), ne nativna lista:
        ' CB.DropDown je crtao penzionisanu nativnu listu PREKO panela
        ' (krug 6 S12). UiEvent sam izlazi kad forme nema (testovi).
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_POV_NEIZABRAN"), True
        modOtkupUI.UiEvent "scrSlPov", "Drop", Empty
        Exit Function
    End If
    otpID = Split(izbor, "|")(0)
    prikaz = Mid$(izbor, InStr(1, izbor, "|") + 1)

    If ReassignOtkupToOtpremnica_TX(otkupID, otpID) Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_POVEZI_OK") & " " & prikaz, False
        NapuniPovKandidate ""            ' red nestaje iz liste -- polje se prazni
        PoveziRed = True
    Else
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_POVEZI_NEUSPEH"), True
    End If
End Function

' Kandidati za povezivanje u polju zone (krug 4 S9). Prazan otkupID
' prazni polje. Stavka se NE bira automatski (S1 klasa zamki) -- samo se
' lista puni; skrivena kolona nosi "OtpremnicaID|broj".
Private Sub NapuniPovKandidate(ByVal otkupID As String)
    Dim CB As Object, k As Variant, i As Long
    mPovOtkupID = otkupID
    Set CB = Kontrola("scrSlPov")
    If CB Is Nothing Then Exit Sub
    mDokFill = True
    On Error Resume Next
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = CStr(Int(CB.width) - 8) & " pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    CB.text = ""
    If Len(otkupID) > 0 Then
        k = modSledljivost.GetOtpremnicaKandidatiZaOtkup(otkupID)
        If IsArray(k) Then
            ' Isti ritam separatora kao polje dokumenta (krug 5) i isti
            ' JEDAN upis liste (krug 6 S13).
            Dim arr() As Variant, n As Long
            n = UBound(k, 1)
            ReDim arr(0 To n - 1, 0 To 1)
            For i = 1 To n
                arr(i - 1, 0) = NzS(k(i, 2)) & _
                    IIf(Len(NzS(k(i, 3))) > 0, _
                        "  " & ChrW(183) & " " & NzS(k(i, 3)), "") & _
                    "  " & ChrW(183) & " " & FmtKolicina(NzD(k(i, 4))) & _
                    " kg" & "  " & ChrW(183) & " " & NzS(k(i, 5))
                arr(i - 1, 1) = NzS(k(i, 1)) & "|" & NzS(k(i, 2))
            Next i
            CB.List = arr
        End If
    End If
    On Error GoTo 0
    mDokFill = False
End Sub

' AUTOMATSKO povezivanje -- legacy btnAutoLink, isti TX. Krug 8 R7:
' pravilo je GLOBALNO (svi periodi, ne prikazani opseg -- poruka to i
' kaze), a greska/rollback se razlikuje od legitimne nule kroz ByRef.
' True kad je nesto povezano -> ljuska osvezava liste.
Private Function AutoPovezi() As Boolean
    Dim n As Long, greska As Boolean
    n = modSledljivost.AutoLinkOtkupOtpremnica_TX(greska)
    If greska Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_AUTO_GRESKA"), True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_POVEZANO") & " " & CStr(n), False
    AutoPovezi = (n > 0)
End Function

' "Sledljivost (PDF)" (smoke krug 3b): 1) RAZRESEN izbor u polju
' "Dokument sledljivosti" ide prvi; 2) kucan a nerazresen tekst ODBIJA
' porukom -- ne pogadja se (pravilo nerazresenog izbora, testovi.md par. 5);
' 3) prazno polje pada na izabrani red -- jedna meta ide odmah, vise njih
' upucuje na polje. InputBox biranja NEMA (nalaz operatera 31.08).
' OFF rezim se prijavljuje.
Private Sub StampajSledljivostReda()
    Dim izbor As String, CB As Object
    Dim mete As Variant

    If DocResolveMode(GetConfigValue(CFG_SLEDLJIVOST_PRINT_MODE), "PDF") = "OFF" Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_PRINT_OFF"), True
        Exit Sub
    End If

    izbor = IzabraniDok()
    If Len(izbor) > 0 Then
        StampajMetu Split(izbor, "|")(0), Mid$(izbor, InStr(1, izbor, "|") + 1)
        Exit Sub
    End If
    Set CB = Kontrola("scrSlDok")
    If Not CB Is Nothing Then
        If Len(Trim$(CStr(CB.text))) > 0 Then
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_DOK_NEIZABRAN"), True
            Exit Sub
        End If
    End If

    ' Prazno polje: kontekst je izabrani red.
    If Len(mIzabranaZbirna) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_NEMA_ZBIRNE"), True
        Exit Sub
    End If
    mete = ReportSledljivostMete(mIzabranaZbirna)
    If Not IsArray(mete) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_NEMA_ZBIRNE"), True
        Exit Sub
    End If
    If UBound(mete, 1) = 1 Then
        StampajMetu NzS(mete(1, 1)), NzS(mete(1, 2))
    Else
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_VISE_META"), False
    End If
End Sub

' Prikazno ime tipa mete -- kroz katalog, kao SlProblemNaziv.
Private Function SlMetaNaziv(ByVal tip As String) As String
    Select Case tip
        Case SLEDM_ZBIRNA:  SlMetaNaziv = Poruka("OTKUI_SL_META_ZBIRNA")
        Case SLEDM_PALETA:  SlMetaNaziv = Poruka("OTKUI_SL_META_PALETA")
        Case SLEDM_PRERADA: SlMetaNaziv = Poruka("OTKUI_SL_META_PRERADA")
        Case SLEDM_NEJASNA: SlMetaNaziv = Poruka("OTKUI_SL_META_NEJASNA")
        Case Else:          SlMetaNaziv = tip
    End Select
End Function

' Izlaz jednog dokumenta sledljivosti. Zbirna ide kroz postojecu rutu
' (postuje ceo SLEDLJIVOST_PRINT_MODE); paletni i preradni list izlaze
' kao PDF -- dugme je "(PDF)", a njihova fizicka stampa ostaje na
' ekranu Palete.
Private Sub StampajMetu(ByVal tip As String, ByVal iD As String)
    Dim ishod As String
    Select Case tip
        Case SLEDM_ZBIRNA
            ishod = StampajSledljivostZbirne(iD)
            Select Case ishod
                Case "OFF":  modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_PRINT_OFF"), True
                Case "NEMA": modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_PRAZNO"), True
                Case "DVOSMISLEN": modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_DVOSMISLENA"), True
            End Select
        Case SLEDM_NEJASNA
            ' Krug 8 R3: broj dele razliciti vlasnici -- nema stampe.
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_DVOSMISLENA"), True
        Case SLEDM_PALETA
            If Len(ExportPaletniListPDF(iD, True)) = 0 Then _
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_STAMPA_NEUSPEH"), True
        Case SLEDM_PRERADA
            If Len(ExportPreradaPDF(iD, True)) = 0 Then _
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_STAMPA_NEUSPEH"), True
    End Select
End Sub

'=====================================================================
' POLJE IZBORA DOKUMENTA SLEDLJIVOSTI (smoke krug 3b)
'=====================================================================

' Ponuda polja: (1..N, 1..2) prikaz | "tip|id" iz snimka -- SVI
' dokumenti perioda. Suzavanje PRI KUCANJU radi LJUSKIN panel
' (PopFromTyping/PopIndex, podniz nad prikazom) -- ekran ponudu NE
' filtrira sam (krug 6 S13: duplo suzavanje, 2N COM poziva po slovu).
' Prikaz zato NOSI sve po cemu se trazi: tip, broj, datum, opis.
' Cist sklop nad snimkom, testabilan bez kontrole; Empty van snimka.
Public Function SlDokPonuda() As Variant
    Dim snap As Variant, dok As Variant, i As Long, n As Long
    Dim res() As Variant

    If Not SnimakPostoji() Then Exit Function
    snap = mSnimci(mSnimakKljuc)
    If Not IsArray(snap) Then Exit Function
    If UBound(snap) < 2 Then Exit Function
    dok = snap(2)
    If Not IsArray(dok) Then Exit Function

    n = UBound(dok, 1)
    ReDim res(1 To n, 1 To 2)
    For i = 1 To n
        res(i, 1) = SlMetaNaziv(NzS(dok(i, 1))) & " " & NzS(dok(i, 3)) & _
                    IIf(IsEmpty(dok(i, 4)), "", _
                        "  " & ChrW(183) & " " & Format$(CDate(dok(i, 4)), "d.M.")) & _
                    IIf(Len(NzS(dok(i, 5))) > 0, _
                        "  " & ChrW(183) & " " & NzS(dok(i, 5)), "")
        res(i, 2) = NzS(dok(i, 1)) & "|" & NzS(dok(i, 2))
    Next i
    SlDokPonuda = res
End Function

' Punjenje polja iz snimka -- guard po (kljuc snimka # punjenja #
' generacija), pa se ne puni iznova na svaki layout prolaz.
Private Sub PuniDokCombo()
    Dim CB As Object, k As String
    If Not SnimakPostoji() Then Exit Sub
    Set CB = Kontrola("scrSlDok")
    If CB Is Nothing Then Exit Sub
    k = mSnimakKljuc & "#" & CStr(mSnimakPunjenja) & "#" & CStr(mSnimakGen)
    If k = mCmbDokKljuc Then Exit Sub
    mCmbDokKljuc = k
    NapuniDokStavke CB, SlDokPonuda()
End Sub

' Prepis stavki u kontrolu (2 kolone: prikaz + skriveni "tip|id" --
' obrazac scrAgKoop). Tekst operatera se CUVA: Clear ume da ga dira, a
' upis ide pod mDokFill da povratni Change ne pokrene punjenje iznova.
' Sirina kolone prati KONTROLU (krug 4 S7: tvrda "244 pt" je bila sira
' od liste). Lista se dodeljuje JEDNIM upisom (.List = matrica) -- 2N
' AddItem poziva nad 1000+ dokumenata je bilo vidljivo sporo prvo
' otvaranje ekrana (krug 6 S13).
Private Sub NapuniDokStavke(ByVal CB As Object, ByVal stavke As Variant)
    Dim i As Long, n As Long, t As String
    Dim arr() As Variant
    mDokFill = True
    On Error Resume Next
    t = CStr(CB.text)
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = CStr(Int(CB.width) - 8) & " pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    If IsArray(stavke) Then
        n = UBound(stavke, 1)
        ReDim arr(0 To n - 1, 0 To 1)
        For i = 1 To n
            arr(i - 1, 0) = CStr(stavke(i, 1))
            arr(i - 1, 1) = CStr(stavke(i, 2))
        Next i
        CB.List = arr
    End If
    If CStr(CB.text) <> t Then CB.text = t
    On Error GoTo 0
    mDokFill = False
End Sub

' Razresen izbor polja dokumenta: "tip|id" ili "" (nerazreseno).
Private Function IzabraniDok() As String
    IzabraniDok = IzabraniIzComba("scrSlDok")
End Function

' Razresen izbor comba: skrivena kolona stavke ili "" (nerazreseno).
' Izbor iz liste resava; PUN kucan tekst jednak prikazu stavke takodje
' (kucan do kraja). Delimican tekst NIJE izbor -- ne pogadja se
' (pravilo nerazresenog izbora, testovi.md par. 5).
Private Function IzabraniIzComba(ByVal nm As String) As String
    Dim CB As Object, i As Long, t As String
    Set CB = Kontrola(nm)
    If CB Is Nothing Then Exit Function
    On Error Resume Next
    If CB.ListIndex >= 0 Then
        IzabraniIzComba = CStr(CB.List(CB.ListIndex, 1))
        Exit Function
    End If
    t = Trim$(CStr(CB.text))
    If Len(t) = 0 Then Exit Function
    For i = 0 To CB.ListCount - 1
        If CStr(CB.List(i, 0)) = t Then
            IzabraniIzComba = CStr(CB.List(i, 1))
            Exit Function
        End If
    Next i
End Function

'=====================================================================
' REDOVI MREZE
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Dim rez As Variant
    rez = RedoviZaListu(filter, q)
    Scr_Rows = rez
    OsveziZonu

    ' Trag za Diag_SlRedovi -- ne menja nista.
    On Error Resume Next
    mDiagFilter = filter
    mDiagQ = q
    mDiagN = CLng(rez(2))
    Err.Clear
End Function

' Opis kolona PO KLJUCU LISTE. Format ljuske: "KLJUC|IZVOR|VRSTA|SIRINA|PRIO"
' (IZVOR prazan -- redove daje ekran). Identitet je poslednja kolona
' prioriteta 4 (mreza crta do 3), citan kroz GridCell.
Public Function SlKoloneZaListu(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case SL_LANAC
            ' Datum | Br. dok | Kooperant | Kg | Otpremnica | Zbirna |
            ' Prijem | Pal. sveze | Pal. gotovog | Faktura | Kupac |
            ' Oznaka | Stanje | [OTK|id] -- GP grana. Smoke GP-1: kolone
            ' prate TOK ROBE (posle prijemnice roba lezi na paletama, pa
            ' se tek onda prodaje), stanje je zakljucak i ide poslednje.
            ' Karike su txt: prazno = karika ne postoji, razlog nosi
            ' OZNAKA (FM-0028 #5 -- nikad "0,00" umesto poruke).
            SlKoloneZaListu = Array( _
                "OTKUI_HD_DATUM||date|60|1", _
                "OTKUI_HDI_BRDOK||txt|66|1", _
                "OTKUI_HDA_KOOPERANT||txt|104|1", _
                "OTKUI_HD_KG||kg|66|1", _
                "OTKUI_HDI_BROTP||txt|68|3", _
                "OTKUI_HDI_BRZBIRNE||txt|74|2", _
                "OTKUI_HDS_PRIJEM||txt|74|1", _
                "OTKUI_HDS_PALETE||txt|96|2", _
                "OTKUI_HDS_PRERADAGP||txt|118|2", _
                "OTKUI_HDS_FAKTURA||txt|68|1", _
                "OTKUI_HDS_KUPAC||txt|88|3", _
                "OTKUI_HDS_OZNAKA||txt|90|1", _
                "OTKUI_HDS_STANJE||txt|78|2", _
                "OTKUI_HDI_REF||txt|1|4")
        Case SL_PARC
            ' Kooperant | BPG | Kat. broj | Kultura | Ha | GGAP | Kg | Datum |
            ' Br. dok | Zbirna | Oznaka | [OTK|id] -- legacy TraceByZbirna
            ' kolone, za sertifikacioni odgovor "od kojih parcela".
            SlKoloneZaListu = Array( _
                "OTKUI_HDA_KOOPERANT||txt|112|1", _
                "OTKUI_HDS_BPG||txt|84|2", _
                "OTKUI_HDS_KATBROJ||txt|68|1", _
                "OTKUI_HDS_KULTURA||txt|76|2", _
                "OTKUI_HDS_POVRSINA||rest|52|3", _
                "OTKUI_HDS_GGAP||txt|60|2", _
                "OTKUI_HD_KG||kg|70|1", _
                "OTKUI_HD_DATUM||date|60|1", _
                "OTKUI_HDI_BRDOK||txt|70|1", _
                "OTKUI_HDI_BRZBIRNE||txt|76|3", _
                "OTKUI_HDS_OZNAKA||txt|76|1", _
                "OTKUI_HDI_REF||txt|1|4")
        Case SL_NEP
            ' Problem | Datum | Broj | Nosilac | Kg | Detalj | [DokTip] |
            ' [DokID] | [KlasaKod]. Kg je txt (FmtIliPrazno): dvosmislen broj
            ' nema kg, a u tipiziranoj koloni bi Empty postao "0,00".
            ' KlasaKod (SLEDP_*) nosi radnja "Povezi..." -- prikazno ime
            ' klase ide kroz Poruka() pa nije stabilan kljuc.
            SlKoloneZaListu = Array( _
                "OTKUI_HDS_PROBLEM||txt|128|1", _
                "OTKUI_HD_DATUM||date|60|1", _
                "OTKUI_HDI_BRDOK||txt|76|1", _
                "OTKUI_HDS_NOSILAC||txt|108|1", _
                "OTKUI_HD_KG||txt|64|2", _
                "OTKUI_HDS_DETALJ||txt|190|1", _
                "OTKUI_HDI_DOKTIP||txt|1|4", _
                "OTKUI_HDI_DOKID||txt|1|4", _
                "OTKUI_HDS_KLASAKOD||txt|1|4")
    End Select
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

' Jedan poziv = jedan kontekst: opseg -> kljuc snimka -> sirovi podaci ->
' oblikovani redovi pod (filter, q).
Private Function RedoviZaListu(ByVal filter As String, ByVal q As String) As Variant
    Dim kljuc As String
    Dim odN As Double, doN As Double
    Dim kolone As Variant, snap As Variant
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    kljuc = Scr_Lista()
    OcistiDetalj
    kolone = SlKoloneZaListu(kljuc)

    OpsegGranice odN, doN

    Dim k As String
    k = CStr(odN) & "|" & CStr(doN)
    snap = Snimak(k, odN, doN)

    mHintKljuc = ""
    RedoviZaListu = Oblikuj(kljuc, snap, kolone, filter, q)
    Exit Function
EH:
    errNum = Err.Number
    errDesc = Err.description
    Err.Raise errNum, "modScrSledljivost.RedoviZaListu", errDesc
End Function

' Snimak konteksta: Array(lanac, problemi) -- OBA read-modela u JEDNOM
' punjenju, pa prelaz na bilo koju listu ne cita tabele ponovo. Iz Report*
' se ide SAMO kad kljuc nije u mapi ili je generacija podataka odmakla
' (par. 23.10/R1). Greska citanja se NE kesira (Err prekida pre upisa).
Private Function Snimak(ByVal k As String, ByVal odN As Double, _
                        ByVal doN As Double) As Variant
    If mSnimakGen <> modUiData.DataGeneracija() Then
        Set mSnimci = Nothing
        mSnimakGen = modUiData.DataGeneracija()
    End If
    If mSnimci Is Nothing Then Set mSnimci = CreateObject("Scripting.Dictionary")

    If Not mSnimci.Exists(k) Then
        If mSnimci.count >= SL_SNIMAK_KAPA Then mSnimci.RemoveAll
        mSnimakPunjenja = mSnimakPunjenja + 1
        Dim dOd As Date, dDo As Date
        dOd = CDate(IIf(odN > 0, odN, SL_DAT_MIN))
        dDo = CDate(IIf(doN > 0, doN, SL_DAT_MAX))
        ' TRI read-modela u istom punjenju: liste (lanac+problemi) i
        ' ponuda polja izbora dokumenta (krug 3b) dele kontekst.
        ' TableCache (krug 8 R6): sva tri prolaze kroz iste velike tabele
        ' -- jedno punjenje kesa umesto ponovljenih citanja listova.
        ' Greska citanja se i dalje NE kesira, ali kes MORA da se zatvori
        ' i na gresci (EHKes), inace bi ostao otvoren preko celog ekrana.
        Dim lanacV As Variant, problemiV As Variant, dokV As Variant
        BeginTableCache
        On Error GoTo EHKes
        lanacV = ReportSledljivostLanac(dOd, dDo)
        problemiV = ReportSledljivostProblemi(dOd, dDo)
        dokV = ReportSledljivostDokumenti(dOd, dDo)
        On Error GoTo 0
        EndTableCache
        mSnimci(k) = Array(lanacV, problemiV, dokV)
    End If

    mSnimakKljuc = k
    mCtxOd = odN
    mCtxDo = doN

    Dim rezultat As Variant
    rezultat = mSnimci(k)
    Snimak = rezultat

    ' KPI zone iz snimka -- nula i "jos nije citano" nisu ista brojka.
    Dim lanac As Variant, problemi As Variant, i As Long, potpunih As Long
    lanac = rezultat(0)
    problemi = rezultat(1)
    If IsArray(lanac) Then
        For i = 1 To UBound(lanac, 1)
            If Len(Trim$(CStr(NzS(lanac(i, 14))))) = 0 Then potpunih = potpunih + 1
        Next i
    End If
    mKpiPotpun = potpunih
    If IsArray(problemi) Then
        mKpiProblemi = UBound(problemi, 1)
    Else
        mKpiProblemi = 0
    End If
    Exit Function

EHKes:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    EndTableCache
    Err.Raise errNum, errSrc, errDesc
End Function

'=====================================================================
' OBLIKOVANJE: snimak -> redovi mreze pod (filter, q). Pravila deljena:
'  - haystack ide kroz modUiData.TekstZaPretragu (kvake u podacima, ASCII
'    upit operatera -- N3);
'  - identitet prio 4, GridCell ga cita, celija se ne crta;
'  - UKUPNO reda nema ni u izvoru (Report* ga ne vraca) -- podnozje daje
'    zbir prikazanih, stampa dodaje svoj izracunat UKUPNO.
'=====================================================================
Private Function Oblikuj(ByVal kljuc As String, ByVal snap As Variant, _
                         ByVal kolone As Variant, _
                         ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant
    Dim nSrc As Long, nK As Long, i As Long, n As Long
    Dim outA() As Variant
    Dim qN As String, hay As String
    Dim sumKg As Double, sumVal As Double

    If Not IsArray(snap) Then
        Oblikuj = PrazanRezultat(kolone)
        HintZaPrazno kljuc, snap
        Exit Function
    End If
    If kljuc = SL_NEP Then src = snap(1) Else src = snap(0)

    nK = UBound(kolone) + 1
    If IsEmpty(src) Or Not IsArray(src) Then
        Oblikuj = PrazanRezultat(kolone)
        HintZaPrazno kljuc, snap
        Exit Function
    End If
    nSrc = UBound(src, 1)

    qN = modUiData.TekstZaPretragu(q)

    ReDim outA(1 To nSrc, 1 To nK)
    For i = 1 To nSrc
        If Not CipPropusta(kljuc, filter, src, i) Then GoTo Sledeci
        If Len(qN) > 0 Then
            hay = modUiData.TekstZaPretragu(HaystackReda(kljuc, src, i))
            If InStr(1, hay, qN, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        UpisiRed kljuc, src, i, outA, n, sumKg, sumVal
Sledeci:
    Next i

    If n = 0 Then HintZaPrazno kljuc, snap
    Oblikuj = Array(kolone, outA, n, sumKg, sumVal, Array(0, 0, 0))
End Function

' Hint kad je lista prazna: kaze ZASTO i KUDA (merilo #3 zadatka) -- prazan
' period ("prosiri period") nije isto sto i "sve karike potpune" na listi
' NEPOTPUNI (to je dobro stanje i tako se i kaze).
Private Sub HintZaPrazno(ByVal kljuc As String, ByVal snap As Variant)
    Dim lanacPrazan As Boolean
    lanacPrazan = True
    If IsArray(snap) Then
        If IsArray(snap(0)) Then lanacPrazan = False
    End If
    If kljuc = SL_NEP And Not lanacPrazan Then
        mHintKljuc = "OTKUI_SL_HINT_SVIPOTPUNI"
    ElseIf lanacPrazan Then
        mHintKljuc = "OTKUI_SL_HINT_PRAZNO"
    End If
End Sub

Private Function CipPropusta(ByVal kljuc As String, ByVal filter As String, _
                             ByRef src As Variant, ByVal i As Long) As Boolean
    Select Case kljuc
        Case SL_LANAC
            CipPropusta = SlCipLanac(filter, NzS(src(i, 14)))
        Case SL_PARC
            CipPropusta = SlCipParcele(filter, NzS(src(i, 17)))
        Case SL_NEP
            CipPropusta = SlCipProblemi(filter, NzS(src(i, 1)))
        Case Else
            CipPropusta = True
    End Select
End Function

' Haystack nosi SVE brojeve lanca -- to je "smer nazad": upit sa brojem
' fakture/prijemnice/zbirne suzava listu na njihove otkupe, pa se kooperanti
' i parcele citaju iz redova.
Private Function HaystackReda(ByVal kljuc As String, ByRef src As Variant, _
                              ByVal i As Long) As String
    Select Case kljuc
        Case SL_LANAC
            ' Kolona 27 = SearchRefs (krug 8 R2): svi brojevi prijemnica i
            ' faktura reda -- "2 prij."/"2 fakt." prikaz ih guta, a smer
            ' NAZAD ih mora naci.
            HaystackReda = NzS(src(i, 2)) & "|" & NzS(src(i, 4)) & "|" & _
                           NzS(src(i, 8)) & "|" & NzS(src(i, 9)) & "|" & _
                           NzS(src(i, 10)) & "|" & NzS(src(i, 12)) & "|" & _
                           NzS(src(i, 13)) & "|" & NzS(src(i, 23)) & "|" & _
                           NzS(src(i, 14)) & "|" & NzS(src(i, 27)) & "|" & _
                           NzS(src(i, 30))
        Case SL_PARC
            ' Isti SearchRefs + kupac (krug 8 R2): "znam fakturu -> nadji
            ' kooperante i parcele" mora da radi i na sertifikacionoj
            ' projekciji.
            HaystackReda = NzS(src(i, 4)) & "|" & NzS(src(i, 22)) & "|" & _
                           NzS(src(i, 18)) & "|" & NzS(src(i, 19)) & "|" & _
                           NzS(src(i, 2)) & "|" & NzS(src(i, 9)) & "|" & _
                           NzS(src(i, 23)) & "|" & NzS(src(i, 13)) & "|" & _
                           NzS(src(i, 27))
        Case SL_NEP
            ' Kolona 9 = lanac-brojevi reda (npr. broj zbirne nefakturisane
            ' prijemnice) -- obecanje "pretraga nalazi svaki broj u lancu"
            ' vazi i na NEPOTPUNIMA (krug 4 S8).
            HaystackReda = NzS(src(i, 3)) & "|" & NzS(src(i, 4)) & "|" & _
                           NzS(src(i, 6)) & "|" & SlProblemNaziv(NzS(src(i, 1))) & _
                           "|" & NzS(src(i, 9))
    End Select
End Function

' Upis jednog reda snimka u red mreze + zbir podnozja (POD ISTIM filterima
' kao redovi -- par. 13).
Private Sub UpisiRed(ByVal kljuc As String, ByRef src As Variant, ByVal i As Long, _
                     ByRef outA() As Variant, ByVal n As Long, _
                     ByRef sumKg As Double, ByRef sumVal As Double)
    Select Case kljuc
        Case SL_LANAC
            outA(n, 1) = SlDatCell(src(i, 1))
            outA(n, 2) = NzS(src(i, 2))
            outA(n, 3) = NzS(src(i, 4))
            outA(n, 4) = NzD(src(i, 7))
            outA(n, 5) = NzS(src(i, 8))
            outA(n, 6) = NzS(src(i, 9))
            outA(n, 7) = NzS(src(i, 10))
            outA(n, 8) = NzS(src(i, 28))
            outA(n, 9) = NzS(src(i, 29))
            outA(n, 10) = NzS(src(i, 12))
            outA(n, 11) = NzS(src(i, 13))
            outA(n, 12) = NzS(src(i, 14))
            outA(n, 13) = NzS(src(i, 30))
            outA(n, 14) = SlRef(src(i, 15))
            sumKg = sumKg + NzD(src(i, 7))
        Case SL_PARC
            outA(n, 1) = NzS(src(i, 4))
            outA(n, 2) = NzS(src(i, 22))
            outA(n, 3) = NzS(src(i, 18))
            outA(n, 4) = NzS(src(i, 19))
            outA(n, 5) = NzD(src(i, 20))
            outA(n, 6) = NzS(src(i, 21))
            outA(n, 7) = NzD(src(i, 7))
            outA(n, 8) = SlDatCell(src(i, 1))
            outA(n, 9) = NzS(src(i, 2))
            outA(n, 10) = NzS(src(i, 9))
            If Len(NzS(src(i, 17))) = 0 Then
                outA(n, 11) = SLED_OZN_BEZ_PARCELE
            Else
                outA(n, 11) = ""
            End If
            outA(n, 12) = SlRef(src(i, 15))
            sumKg = sumKg + NzD(src(i, 7))
        Case SL_NEP
            outA(n, 1) = SlProblemNaziv(NzS(src(i, 1)))
            outA(n, 2) = SlDatCell(src(i, 2))
            outA(n, 3) = NzS(src(i, 3))
            outA(n, 4) = NzS(src(i, 4))
            outA(n, 5) = FmtIliPrazno(src(i, 5))
            outA(n, 6) = NzS(src(i, 6))
            outA(n, 7) = NzS(src(i, 7))
            outA(n, 8) = NzS(src(i, 8))
            outA(n, 9) = NzS(src(i, 1))
            sumKg = sumKg + NzD(src(i, 5))
    End Select
End Sub

Private Function SlRef(ByVal v As Variant) As String
    If Len(NzS(v)) > 0 Then SlRef = "OTK|" & NzS(v)
End Function

' Prikazni naziv klase problema -- kod iz Report* je ASCII konstanta,
' operater vidi tekst iz kataloga.
Public Function SlProblemNaziv(ByVal klasa As String) As String
    Select Case klasa
        Case SLEDP_BEZ_OTPREMNICE:  SlProblemNaziv = Poruka("OTKUI_SLP_BEZOTP")
        Case SLEDP_VEZA:            SlProblemNaziv = Poruka("OTKUI_SLP_VEZA")
        Case SLEDP_BEZ_ZBIRNE:      SlProblemNaziv = Poruka("OTKUI_SLP_BEZZBR")
        Case SLEDP_BROJ_DVOSMISLEN: SlProblemNaziv = Poruka("OTKUI_SLP_DVOSM")
        Case SLEDP_BEZ_PRIJEMA:     SlProblemNaziv = Poruka("OTKUI_SLP_BEZPRIJ")
        Case SLEDP_FAK_NEISPRAVNA:  SlProblemNaziv = Poruka("OTKUI_SLP_NEFAKT")
        Case SLEDP_KG_RAZLIKA:      SlProblemNaziv = Poruka("OTKUI_SLP_KG")
        Case Else:                  SlProblemNaziv = klasa
    End Select
End Function

'--------------------------------------------------------------- POMOCNI
Private Function NzD(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then NzD = CDbl(v)
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    On Error Resume Next
    NzS = Trim$(CStr(v))
End Function

' Datum -> serijski broj za kolonu tipa "date" (0 = uredno prazno).
Private Function SlDatCell(ByVal v As Variant) As Double
    Dim d As Double
    If IsDate(v) Then
        d = Int(CDbl(CDate(v)))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        d = Int(CDbl(v))
    End If
    If d < 1 Or d > SL_DAT_MAX Then Exit Function
    SlDatCell = d
End Function

' Kolone kod kojih je PRAZNO poruka: broj se formatira, Empty ostaje prazno
' -- nikad "0,00" umesto oznake (FM-0028 #5).
Private Function FmtIliPrazno(ByVal v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) Then FmtIliPrazno = FmtKolicina(CDbl(v))
End Function

' Datum kao GRANICA opsega; 0 = nema granice. ISTO pravilo kao DatGranica u
' modScrDokumenti / IzDatGranica u modScrIzvestaji (svaki ekran nosi svoju
' 3-linijsku kopiju nad deljenim TryParseDateValue) -- ne izmislja se novo
' parsiranje.
Public Function SlDatGranica(ByVal s As String) As Double
    Dim d As Date
    On Error Resume Next
    If Len(Trim$(s)) = 0 Then Exit Function
    If TryParseDateValue(s, d) Then SlDatGranica = Int(CDbl(d))
End Function

'=====================================================================
' KONTEKST: opseg datuma.
'=====================================================================
Private Function SnimakPostoji() As Boolean
    If mSnimci Is Nothing Then Exit Function
    If Len(mSnimakKljuc) = 0 Then Exit Function
    SnimakPostoji = mSnimci.Exists(mSnimakKljuc)
End Function

Private Sub OpsegPolja(ByRef odTxt As String, ByRef doTxt As String)
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrSlOd")
    If Not c Is Nothing Then odTxt = Trim$(CStr(c.text))
    Set c = Kontrola("scrSlDo")
    If Not c Is Nothing Then doTxt = Trim$(CStr(c.text))
    Err.Clear
End Sub

Private Sub OpsegGranice(ByRef odN As Double, ByRef doN As Double)
    Dim odTxt As String, doTxt As String
    If IsTestMode() Then
        If mTestOd > 0 Or mTestDo > 0 Then
            odN = mTestOd
            doN = mTestDo
            Exit Sub
        End If
    End If
    OpsegPolja odTxt, doTxt
    odN = SlDatGranica(odTxt)
    doN = SlDatGranica(doTxt)
End Sub

Private Function OpsegLabela() As String
    Dim s1 As String, s2 As String
    s1 = IIf(mCtxOd > 0, Format$(CDate(mCtxOd), "d.m.yyyy"), Poruka("OTKUI_IZ_BEZ_GRANICE"))
    s2 = IIf(mCtxDo > 0, Format$(CDate(mCtxDo), "d.m.yyyy"), Poruka("OTKUI_IZ_BEZ_GRANICE"))
    OpsegLabela = s1 & " - " & s2
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod reda polja -- LABELA, ne Frame (par. 7.7).
    modUiKit.NewLbl z, "slBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "slCap", UCase$(Poruka("OTKUI_SCRSL_CAP")), PAD, SL_Y_CAP, _
                    200, 11, TS_MICRO, True, C_MUTED, -1

    ' Dve KPI brojke desno: potpuni lanci / nepotpune karike (iz snimka).
    For i = 0 To 1
        modUiKit.NewLbl z, "slKL" & i, "", 0, SL_Y_CAP, SL_KPI_W, 11, _
                        TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "slKV" & i, ChrW(8212), 0, SL_Y_CAP + 14, SL_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' DETALJ TRAKA: pun lanac izabranog reda, karika po karika.
    modUiKit.NewLbl z, "slDetCap", "", 0, SL_Y_LBL, SL_DET_W, 11, _
                    TS_MICRO, True, C_MUTED, -1
    For i = 0 To SL_DET_N - 1
        modUiKit.NewLbl z, "slDetR" & i, "", 0, SL_Y_LBL + 14 + i * 12, _
                        SL_DET_W, 12, TS_META, False, C_FOREST, -1
    Next i

    ' POLJA. Pravi ih ljuska (NewFieldG); prefiks "scr" je OBAVEZAN.
    modOtkupUI.NewFieldG z, "scrSlOd", Poruka("OTKUI_FLD_IZ_OD"), "txt", "", _
                         1, False, False, "SL"
    modOtkupUI.NewFieldG z, "scrSlDo", Poruka("OTKUI_FLD_IZ_DO"), "txt", "", _
                         1, False, False, "SL"

    ' Legacy default opsega: 1.1. tekuce godine -- danas (kao Izvestaji).
    On Error Resume Next
    z.Controls("scrSlOd").Controls("scrSlOdT").text = "1.1." & Year(Date)
    z.Controls("scrSlDo").Controls("scrSlDoT").text = Format$(Date, "d.m.yyyy")
    On Error GoTo 0

    ' POLJE IZBORA DOKUMENTA SLEDLJIVOSTI (smoke krug 3b): dropdown sa
    ' filterom -- kucanje (broj, datum, tip, status...) suzava ponudu.
    ' MatchEntry none: MSForms prefix-autocomplete bi se tukao sa
    ' substring filterom i prepisivao kucani tekst.
    modOtkupUI.NewFieldG z, "scrSlDok", Poruka("OTKUI_FLD_SL_DOK"), "cmb", "", _
                         1, False, False, "SL"
    On Error Resume Next
    z.Controls("scrSlDok").Controls("scrSlDokT").MatchEntry = fmMatchEntryNone
    On Error GoTo 0

    ' POLJE KANDIDATA ZA POVEZIVANJE (krug 4 S9 -- "povezivanje treba
    ' lepse resiti"): puni ga IZBOR reda klase 'Otkup bez otpremnice' na
    ' NEPOTPUNIMA, radnja "Povezi..." cita razresen izbor. Bez InputBox-a.
    modOtkupUI.NewFieldG z, "scrSlPov", Poruka("OTKUI_FLD_SL_POVEZI"), "cmb", "", _
                         1, False, False, "SL"

    ' Fontovi oba comba EKSPLICITNO prate datumska polja (krug 5:
    ' "cudne razlike u fontu" -- ne oslanja se na default kontrole).
    On Error Resume Next
    With z.Controls("scrSlOd").Controls("scrSlOdT").Font
        z.Controls("scrSlDok").Controls("scrSlDokT").Font.name = .name
        z.Controls("scrSlDok").Controls("scrSlDokT").Font.Size = .Size
        z.Controls("scrSlPov").Controls("scrSlPovT").Font.name = .name
        z.Controls("scrSlPov").Controls("scrSlPovT").Font.Size = .Size
    End With
    On Error GoTo 0

    modUiKit.NewLbl z, "slHint", "", PAD, SL_Y_HINT, 420, 12, TS_META, False, C_MUTED, -1

    ' Stampa aktivne liste (house PDF) + lanac izabranog reda (house PDF sa
    ' kontekst-linijom: koren, opseg, kompletnost) + sledljivost zbirne po
    ' POSTOJECEM sablonu + auto-povezivanje (samo NEPOTPUNI -- vidljivost
    ' daje raspored). Smoke krug 2.
    modUiKit.BtnV z, "scrSlPrint", Poruka("OTKUI_BTN_IZ_PRINT"), PAD, SL_Y_BTN, _
                  156, SL_BTN_H, "primary"
    modUiKit.BtnV z, "scrSlLanac", Poruka("OTKUI_BTN_SL_LANACPDF"), PAD + 164, SL_Y_BTN, _
                  120, SL_BTN_H, "soft"
    modUiKit.BtnV z, "scrSlSab", Poruka("OTKUI_BTN_SL_SABLON"), PAD + 292, SL_Y_BTN, _
                  158, SL_BTN_H, "soft"
    modUiKit.BtnV z, "scrSlAuto", Poruka("OTKUI_BTN_SL_AUTO"), PAD + 458, SL_Y_BTN, _
                  132, SL_BTN_H, "soft"

    modUiKit.NewLbl z, "slLnB", "", 0, SL_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    ' Ponuda polja izbora ide iz ISTOG snimka kao liste (0 dodatnih
    ' citanja); guard u punjenju preskace kad se nista nije promenilo.
    PuniDokCombo
    Scr_Layout = SL_ZONA_H
End Function

Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, kx As Single
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    z.Controls("slBg").Left = PAD - 10
    z.Controls("slBg").top = SL_Y_LBL - 8
    z.Controls("slBg").width = w - 2 * (PAD - 10)
    ' Kartica obuhvata i CELU detalj traku (smoke S2: redovi trake su
    ' visili ispod bele podloge i vizuelno ulazili u sledeci blok) --
    ' dno ide do pred donju liniju zone, ne do reda dugmadi.
    z.Controls("slBg").Height = SL_ZONA_H - (SL_Y_LBL - 8) - 8

    ' Brojke uz desnu ivicu.
    For i = 0 To 1
        kx = w - PAD - (2 - i) * SL_KPI_W
        z.Controls("slKL" & i).Left = kx
        z.Controls("slKV" & i).Left = kx
    Next i

    PoljeX z, "scrSlOd", PAD, 86, SL_Y_LBL
    PoljeX z, "scrSlDo", PAD + 94, 86, SL_Y_LBL
    PoljeX z, "scrSlDok", PAD + 188, 264, SL_Y_LBL
    ' Kandidati povezivanja su posao liste NEPOTPUNI -- na ostalima se
    ' polje ne crta (isti obrazac kao dugme scrSlAuto).
    PoljeX z, "scrSlPov", PAD + 460, 240, SL_Y_LBL
    PoljeVidi z, "scrSlPov", (Scr_Lista() = SL_NEP)

    ' Detalj traka uzima desno; polja i hint dele ostatak. Na uskom ekranu
    ' traka nestaje umesto da se preklapa (isti kompromis kao Izvestaji).
    Dim wPolja As Single, detVidi As Boolean, dx As Single
    wPolja = w - SL_DET_W - PAD
    detVidi = (wPolja >= SL_POLJA_MIN)
    If Not detVidi Then wPolja = w
    dx = w - SL_DET_W
    z.Controls("slDetCap").Left = dx
    z.Controls("slDetCap").Visible = detVidi
    For i = 0 To SL_DET_N - 1
        z.Controls("slDetR" & i).Left = dx
        z.Controls("slDetR" & i).Visible = detVidi
    Next i

    z.Controls("slHint").width = wPolja - 2 * PAD

    modUiKit.MoveBtn z, "scrSlPrint", PAD, SL_Y_BTN
    modUiKit.MoveBtn z, "scrSlLanac", PAD + 164, SL_Y_BTN
    modUiKit.MoveBtn z, "scrSlSab", PAD + 292, SL_Y_BTN
    modUiKit.MoveBtn z, "scrSlAuto", PAD + 458, SL_Y_BTN
    ' Auto-povezivanje je posao liste NEPOTPUNI -- na ostalima je mrtvo
    ' dugme i ne crta se (kontekstna dugmad, obrazac scrIzKartPdf).
    modUiKit.BoxShow z, "scrSlAuto", (Scr_Lista() = SL_NEP)

    z.Controls("slLnB").width = w
End Sub

Private Sub PoljeX(ByVal z As Object, ByVal nm As String, ByVal X As Single, _
                   ByVal w As Single, ByVal yLbl As Single)
    On Error Resume Next
    z.Controls(nm).Left = X
    z.Controls(nm).top = yLbl
    z.Controls(nm).width = w
    modOtkupUI.LayoutFieldInner z.Controls(nm)
End Sub

' Vidljivost celog polja (okvir nosi sve unutrasnje kontrole) -- isti
' helper kao na Agrohemiji.
Private Sub PoljeVidi(ByVal z As Object, ByVal nm As String, ByVal vis As Boolean)
    On Error Resume Next
    z.Controls(nm).Visible = vis
End Sub

Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone("SLEDLJIVOST")
End Function

Private Function Kontrola(ByVal nm As String) As Object
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    Set Kontrola = z.Controls(nm).Controls(nm & "T")
End Function

Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    RasporediPolja z, z.width
    OsveziHint z
    OsveziBrojke z
End Sub

' Hint: prazna lista kaze ZASTO i KUDA; inace opis prikazanog konteksta
' (period STVARNO ucitanih podataka, nikad iz polja -- AUD-024).
Private Sub OsveziHint(ByVal z As Object)
    Dim s As String
    On Error Resume Next
    Select Case mHintKljuc
        Case "OTKUI_SL_HINT_PRAZNO"
            s = Poruka("OTKUI_SL_HINT_PRAZNO")
        Case "OTKUI_SL_HINT_SVIPOTPUNI"
            s = Poruka("OTKUI_SL_HINT_SVIPOTPUNI")
        Case Else
            If SnimakPostoji() Then
                s = Poruka("OTKUI_SL_PERIOD") & " " & OpsegLabela() & "  " & _
                    ChrW(183) & "  " & Poruka("OTKUI_SL_HINT")
            Else
                s = Poruka("OTKUI_SL_HINT")
            End If
    End Select
    z.Controls("slHint").caption = s
End Sub

Private Sub OsveziBrojke(ByVal z As Object)
    Dim crta As String
    On Error Resume Next
    crta = ChrW(8212)
    z.Controls("slKL0").caption = UCase$(Poruka("OTKUI_KPI_SL_POTPUN"))
    z.Controls("slKL1").caption = UCase$(Poruka("OTKUI_KPI_SL_PROBLEMI"))
    If IsEmpty(mKpiPotpun) Then
        z.Controls("slKV0").caption = crta
    Else
        z.Controls("slKV0").caption = Format$(NzD(mKpiPotpun), "#,##0")
    End If
    If IsEmpty(mKpiProblemi) Then
        z.Controls("slKV1").caption = crta
    Else
        z.Controls("slKV1").caption = Format$(NzD(mKpiProblemi), "#,##0")
    End If
End Sub

'------------------------------------------------------- DETALJ TRAKA
' Klik na red -> PUN LANAC tog reda, karika po karika sa kg po karici --
' tacno ono sto ravan red NE pokazuje (par. 23.12/S10: samo novi podaci).
' Redovi NEPOTPUNI liste sa otkup-zrnom (DokTip=Otkup) dobijaju isti lanac;
' ostale karike ostavljaju traku praznu (v1 -- zapisano u katalogu).
Private Sub OsveziDetalj(ByVal red As Long)
    Dim ref As String, kljuc As String
    Dim linije As Variant, naslov As String

    kljuc = Scr_Lista()
    mDetalj = Empty
    mIzabranOtkupID = ""
    mIzabranaZbirna = ""
    naslov = ""

    Select Case kljuc
        Case SL_LANAC
            ref = NzS(modOtkupUI.GridCell(red, 14))
            mIzabranaZbirna = NzS(modOtkupUI.GridCell(red, 6))
        Case SL_PARC
            ref = NzS(modOtkupUI.GridCell(red, 12))
            mIzabranaZbirna = NzS(modOtkupUI.GridCell(red, 10))
        Case SL_NEP
            If NzS(modOtkupUI.GridCell(red, 7)) = DOK_TIP_OTKUP Then
                ref = "OTK|" & NzS(modOtkupUI.GridCell(red, 8))
            ElseIf NzS(modOtkupUI.GridCell(red, 7)) = SLED_DOK_ZBIRNA Then
                ' Red zbirne: broj je vidljiva kolona 3 -- sablon-PDF radi
                ' i odavde.
                mIzabranaZbirna = NzS(modOtkupUI.GridCell(red, 3))
            End If
            ' Kandidati povezivanja prate izabrani red (krug 4 S9): pune
            ' se SAMO za klasu OTKUP-BEZ-OTPREMNICE, ostali prazne polje.
            If NzS(modOtkupUI.GridCell(red, 9)) = SLEDP_BEZ_OTPREMNICE _
               And NzS(modOtkupUI.GridCell(red, 7)) = DOK_TIP_OTKUP Then
                NapuniPovKandidate NzS(modOtkupUI.GridCell(red, 8))
            Else
                NapuniPovKandidate ""
            End If
    End Select

    If Left$(ref, 4) = "OTK|" Then
        mIzabranOtkupID = Mid$(ref, 5)
        naslov = Poruka("OTKUI_SL_DET_KARIKA_OTKUP") & " " & _
                 NzS(modOtkupUI.GridCell(red, IIf(kljuc = SL_NEP, 3, IIf(kljuc = SL_PARC, 9, 2))))
        linije = SlDetaljLanca(mIzabranOtkupID, kljuc)
    End If

    If IsArray(linije) Then
        mDetalj = linije
        DetaljTraka naslov
    Else
        DetaljTraka ""
    End If
End Sub

' Linije detalja iz SNIMKA (nula citanja tabela): karika po karika sa kg,
' vozac/stanica (nisu kolone reda), parcela (nije kolona LANAC reda) i
' oznaka. Javna radi testa; van snimka vraca Empty.
Public Function SlDetaljLanca(ByVal otkupID As String, ByVal kljuc As String) As Variant
    Dim snap As Variant, lanac As Variant, i As Long, r As Long
    Dim linije As Collection

    If Not SnimakPostoji() Then Exit Function
    snap = mSnimci(mSnimakKljuc)
    If Not IsArray(snap) Then Exit Function
    lanac = snap(0)
    If Not IsArray(lanac) Then Exit Function

    r = 0
    For i = 1 To UBound(lanac, 1)
        If NzS(lanac(i, 15)) = Trim$(otkupID) Then r = i: Exit For
    Next i
    If r = 0 Then Exit Function

    Set linije = New Collection
    linije.Add Poruka("OTKUI_SL_DET_STANICA") & " " & NzS(lanac(r, 23)) & _
               IIf(Len(NzS(lanac(r, 24))) > 0, "   " & _
               Poruka("OTKUI_IZ_DET_VOZAC") & " " & NzS(lanac(r, 24)), "")
    If kljuc <> SL_PARC Then
        If Len(NzS(lanac(r, 18))) > 0 Then
            linije.Add Poruka("OTKUI_SL_DET_PARCELA") & " " & NzS(lanac(r, 18)) & _
                       IIf(Len(NzS(lanac(r, 19))) > 0, " (" & NzS(lanac(r, 19)) & _
                       IIf(NzD(lanac(r, 20)) > 0, ", " & _
                       Format$(NzD(lanac(r, 20)), "0.##") & " ha", "") & ")", "")
        End If
    End If
    If Len(NzS(lanac(r, 8))) > 0 Then
        linije.Add Poruka("OTKUI_IZ_DET_OTPREMNICA") & " " & NzS(lanac(r, 8)) & _
                   IIf(NzD(lanac(r, 25)) > 0, "  " & ChrW(183) & "  " & _
                   FmtKolicina(NzD(lanac(r, 25))) & " kg", "")
    End If
    If Len(NzS(lanac(r, 9))) > 0 Then
        linije.Add Poruka("OTKUI_IZ_DET_ZBIRNA") & " " & NzS(lanac(r, 9)) & _
                   IIf(NzD(lanac(r, 26)) > 0, "  " & ChrW(183) & "  " & _
                   FmtKolicina(NzD(lanac(r, 26))) & " kg", "") & _
                   IIf(Len(NzS(lanac(r, 13))) > 0, "   " & _
                   Poruka("OTKUI_SL_DET_KUPAC") & " " & NzS(lanac(r, 13)), "")
    End If
    If Len(NzS(lanac(r, 10))) > 0 Then
        linije.Add Poruka("OTKUI_IZ_DET_PRIJEMNICA") & " " & NzS(lanac(r, 10)) & _
                   IIf(Not IsEmpty(lanac(r, 11)), "  " & ChrW(183) & "  " & _
                   FmtKolicina(NzD(lanac(r, 11))) & " kg", "")
    End If
    If Len(NzS(lanac(r, 12))) > 0 Then
        linije.Add Poruka("OTKUI_IZ_DET_FAKTURA") & " " & NzS(lanac(r, 12))
    ElseIf NzS(lanac(r, 14)) = SLED_OZN_FAK_NEISPRAVNA Then
        linije.Add Poruka("OTKUI_IZ_DET_FAKTURA") & " " & Poruka("OTKUI_SL_DET_FAKNEISPRAVNA")
    End If
    If Len(NzS(lanac(r, 14))) > 0 Then
        linije.Add UCase$(NzS(lanac(r, 14)))
    End If

    Dim res() As String, n As Long
    ReDim res(0 To linije.count - 1)
    For n = 1 To linije.count
        res(n - 1) = linije(n)
    Next n
    SlDetaljLanca = res
End Function

' Crtanje trake: naslov + linije + preliv ("... jos N") -- preliv se
' PRIJAVLJUJE (par. 7.8).
Private Sub DetaljTraka(ByVal naslov As String)
    Dim z As Object, i As Long, n As Long
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    n = 0
    If IsArray(mDetalj) Then n = UBound(mDetalj) - LBound(mDetalj) + 1
    z.Controls("slDetCap").caption = UCase$(naslov)
    For i = 0 To SL_DET_N - 1
        If i < n Then
            If i = SL_DET_N - 1 And n > SL_DET_N Then
                z.Controls("slDetR" & i).caption = ChrW(8230) & " " & _
                    Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & CStr(n - SL_DET_N + 1)
            Else
                z.Controls("slDetR" & i).caption = CStr(mDetalj(LBound(mDetalj) + i))
            End If
        Else
            z.Controls("slDetR" & i).caption = ""
        End If
    Next i
End Sub

Private Sub OcistiDetalj()
    mDetalj = Empty
    DetaljTraka ""
End Sub

'=====================================================================
' STAMPE. Ne verifikuju se automatski -- smoke checklista.
'=====================================================================

' KOJE se kolone SABIRAJU u stampanom UKUPNO redu (politika sabirljivosti,
' par. 23.10/R2): sabira se SAMO kg promet. Brojevi dokumenata, povrsina
' parcele (atribut, ne promet -- ponavlja se po redu iste parcele) i kg
' NEPOTPUNI liste (mesa zrna: otkup + otpremnica + zbirna) se NE sabiraju.
Public Function SlSabirljive(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case SL_LANAC: SlSabirljive = Array(4)
        Case SL_PARC:  SlSabirljive = Array(7)
        Case Else:     SlSabirljive = Array()
    End Select
End Function

' Zaglavlja stampe -- isti opis kolona kao mreza (vidljive kolone).
Public Function SlHeaderiZaListu(ByVal kljuc As String) As Variant
    Dim kolone As Variant, i As Long, n As Long
    Dim res() As String
    kolone = SlKoloneZaListu(kljuc)
    ReDim res(0 To UBound(kolone))
    For i = 0 To UBound(kolone)
        If val(Split(CStr(kolone(i)), "|")(4)) < 4 Then
            res(n) = Poruka(Split(CStr(kolone(i)), "|")(0))
            n = n + 1
        End If
    Next i
    ReDim Preserve res(0 To n - 1)
    SlHeaderiZaListu = res
End Function

' "Stampaj izvestaj": house PDF AKTIVNE liste -- tacno ono sto operater
' vidi (cip + pretraga), sa naslovom iz KONTEKSTA SNIMKA (AUD-024) i
' vidljivom napomenom kad je filter aktivan.
Private Sub StampajIzvestaj()
    Dim rez As Variant, redovi As Variant, n As Long
    Dim kolone As Variant, headers As Variant
    Dim dataS() As String
    Dim i As Long, j As Long, vidljivih As Long
    Dim naslov As String

    If Not SnimakPostoji() Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZ_PRVO_PRIKAZI"), True
        Exit Sub
    End If

    rez = RedoviZaListu(mDiagFilter, mDiagQ)
    n = CLng(rez(2))
    If n = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_PRAZNO"), True
        Exit Sub
    End If
    kolone = rez(0)
    redovi = rez(1)
    headers = SlHeaderiZaListu(Scr_Lista())
    vidljivih = UBound(headers) + 1

    ReDim dataS(1 To n + 1, 1 To vidljivih)
    Dim kind As String, tot() As Double, imaTot() As Boolean
    Dim sabir As Variant, s As Long
    ReDim tot(1 To vidljivih)
    ReDim imaTot(1 To vidljivih)
    sabir = SlSabirljive(Scr_Lista())
    If IsArray(sabir) Then
        For s = LBound(sabir) To UBound(sabir)
            If CLng(sabir(s)) >= 1 And CLng(sabir(s)) <= vidljivih Then _
                imaTot(CLng(sabir(s))) = True
        Next s
    End If
    For i = 1 To n
        For j = 1 To vidljivih
            dataS(i, j) = CelijaZaStampu(CStr(kolone(j - 1)), redovi(i, j))
            If imaTot(j) Then tot(j) = tot(j) + NzD(redovi(i, j))
        Next j
    Next i
    dataS(n + 1, 1) = "UKUPNO"
    For j = 2 To vidljivih
        If imaTot(j) Then
            kind = Split(CStr(kolone(j - 1)), "|")(2)
            If kind = "kg" Then
                dataS(n + 1, j) = FmtKolicina(tot(j))
            Else
                dataS(n + 1, j) = Format$(tot(j), "#,##0.##")
            End If
        End If
    Next j

    naslov = Poruka("OTKUI_SL_PERIOD") & " " & OpsegLabela()
    If Not IsEmpty(mKpiProblemi) Then
        naslov = naslov & "  " & ChrW(183) & "  " & _
                 Poruka("OTKUI_KPI_SL_PROBLEMI") & ": " & _
                 Format$(NzD(mKpiProblemi), "#,##0")
    End If
    If Len(mDiagQ) > 0 Then
        naslov = naslov & "  " & ChrW(183) & "  " & Poruka("OTKUI_IZ_PRETRAGA") & _
                 " " & mDiagQ
    End If
    If Len(mDiagFilter) > 0 And mDiagFilter <> "sve" Then
        naslov = naslov & "  " & ChrW(183) & "  " & Poruka("OTKUI_IZ_FILTER") & _
                 " " & mDiagFilter
    End If

    Dim desno() As Boolean
    ReDim desno(0 To vidljivih - 1)
    For j = 1 To vidljivih
        Select Case Split(CStr(kolone(j - 1)), "|")(2)
            Case "kg", "rsd", "num", "rest", "date"
                desno(j - 1) = True
            Case "txt"
                desno(j - 1) = imaTot(j)
        End Select
    Next j

    PrintIzvestajHouse dataS, n + 1, vidljivih, _
                       UCase$(Poruka(NaslovKljucListe(Scr_Lista()))), _
                       naslov, headers, desno
End Sub

Private Function CelijaZaStampu(ByVal spec As String, ByVal v As Variant) As String
    Dim kind As String
    kind = Split(spec, "|")(2)
    Select Case kind
        Case "date"
            If IsNumeric(v) Then
                If CDbl(v) >= 1 Then CelijaZaStampu = Format$(CDate(CDbl(v)), "d.m.yyyy")
            End If
        Case "kg":  CelijaZaStampu = FmtKolicina(NzD(v))
        Case "num": CelijaZaStampu = Format$(NzD(v), "#,##0")
        Case "rsd": CelijaZaStampu = Format$(NzD(v), "#,##0.00")
        Case "rest"
            If NzD(v) <> 0 Then CelijaZaStampu = Format$(NzD(v), "#,##0.##")
        Case Else
            CelijaZaStampu = NzS(v)
    End Select
End Function

Private Function NaslovKljucListe(ByVal kljuc As String) As String
    Select Case kljuc
        Case SL_LANAC: NaslovKljucListe = "OTKUI_GRID_TITLE_SL_LANAC"
        Case SL_PARC:  NaslovKljucListe = "OTKUI_GRID_TITLE_SL_PARC"
        Case SL_NEP:   NaslovKljucListe = "OTKUI_GRID_TITLE_SL_NEP"
    End Select
End Function

' "Lanac (PDF)": sledljivosni izvestaj IZABRANOG reda -- karike kao redovi,
' kontekst-linija nosi koren + opseg + kompletnost. Postuje
' SLEDLJIVOST_PRINT_MODE: OFF se PRIJAVLJUJE (klik bez poruke izgleda kao
' dugme koje ne radi -- par. 22.2/C).
Private Sub StampajLanacIzabranog()
    Dim mode As String
    If Len(mIzabranOtkupID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_IZABERI_RED"), True
        Exit Sub
    End If
    mode = DocResolveMode(GetConfigValue(CFG_SLEDLJIVOST_PRINT_MODE), "PDF")
    If mode = "OFF" Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_SL_PRINT_OFF"), True
        Exit Sub
    End If

    Dim paket As Variant
    paket = SlLanacZaPdf(mIzabranOtkupID)
    If Not IsArray(paket) Then
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZ_PRVO_PRIKAZI"), True
        Exit Sub
    End If

    ' Ozbiljan A4 dokument po ugledu na SledljivostSablon (krug 6 S14 --
    ' "detalj na A4 listu je smesno"): sopstveni list sa info blokom,
    ' tabelom karika sa nosiocima i potpis/pecat podnozjem. House
    ' kompozer lista (PrintIzvestajHouse) je za SIROKE liste i tamo
    ' MERGE naslov sece na sirinu uske tabele -- ne dira se.
    StampajSledljivostLanacDoc paket, mode
End Sub

' Detalj-kljucevi nose dvotacku (traka: "Zbirna: S1/..."); u PDF tabeli
' karika stoji sama u svojoj koloni, pa se dvotacka skida NA MESTU
' UPOTREBE -- tekst kljuceva deli i detalj Izvestaja i ne sme se menjati.
Private Function BezDvotacke(ByVal s As String) As String
    BezDvotacke = Trim$(s)
    If Right$(BezDvotacke, 1) = ":" Then _
        BezDvotacke = Trim$(Left$(BezDvotacke, Len(BezDvotacke) - 1))
End Function

' Redovi lanac-PDF-a iz SNIMKA -- cist sklop, testabilan bez stampe:
' Array(dataS 2D, brojRedova, kontekstLinija) ili Empty kad reda nema.
Public Function SlLanacZaPdf(ByVal otkupID As String) As Variant
    Dim snap As Variant, lanac As Variant, i As Long, r As Long

    If Not SnimakPostoji() Then Exit Function
    snap = mSnimci(mSnimakKljuc)
    If Not IsArray(snap) Then Exit Function
    lanac = snap(0)
    If Not IsArray(lanac) Then Exit Function
    r = 0
    For i = 1 To UBound(lanac, 1)
        If NzS(lanac(i, 15)) = Trim$(otkupID) Then r = i: Exit For
    Next i
    If r = 0 Then Exit Function

    ' Karike bez dvotacke (krug 5 S11) -- kljucevi su deljeni sa detalj
    ' trakama, skida se ovde. Od kruga 6 (S14) red nosi i NOSIOCA
    ' (kolona 3): kooperant / vozac / kupac po karici -- dokument A4
    ' bez nosilaca je bio "detalj na listu".
    Dim dataS(1 To 5, 1 To 5) As String
    dataS(1, 1) = BezDvotacke(Poruka("OTKUI_SL_DET_KARIKA_OTKUP"))
    dataS(1, 2) = NzS(lanac(r, 2))
    dataS(1, 3) = NzS(lanac(r, 4))
    dataS(1, 4) = FmtKolicina(NzD(lanac(r, 7)))
    dataS(1, 5) = ""
    dataS(2, 1) = BezDvotacke(Poruka("OTKUI_IZ_DET_OTPREMNICA"))
    dataS(2, 2) = NzS(lanac(r, 8))
    dataS(2, 3) = NzS(lanac(r, 24))
    If NzD(lanac(r, 25)) > 0 Then dataS(2, 4) = FmtKolicina(NzD(lanac(r, 25)))
    dataS(3, 1) = BezDvotacke(Poruka("OTKUI_IZ_DET_ZBIRNA"))
    dataS(3, 2) = NzS(lanac(r, 9))
    dataS(3, 3) = NzS(lanac(r, 13))
    If NzD(lanac(r, 26)) > 0 Then dataS(3, 4) = FmtKolicina(NzD(lanac(r, 26)))
    dataS(4, 1) = BezDvotacke(Poruka("OTKUI_IZ_DET_PRIJEMNICA"))
    dataS(4, 2) = NzS(lanac(r, 10))
    dataS(4, 3) = NzS(lanac(r, 13))
    If Not IsEmpty(lanac(r, 11)) Then dataS(4, 4) = FmtKolicina(NzD(lanac(r, 11)))
    dataS(5, 1) = BezDvotacke(Poruka("OTKUI_IZ_DET_FAKTURA"))
    dataS(5, 2) = NzS(lanac(r, 12))
    dataS(5, 3) = NzS(lanac(r, 13))

    ' Oznaka stoji uz kariku na kojoj lanac staje/curi.
    Dim ozn As String
    ozn = NzS(lanac(r, 14))
    Select Case ozn
        Case SLED_OZN_NEPOVEZAN, SLED_OZN_OTP_STORNIRANA, SLED_OZN_VEZA
            dataS(2, 5) = ozn
        Case SLED_OZN_BEZ_ZBIRNE, SLED_OZN_ZBIRNA_NEMA, IZV_VLASNIK_NEJASAN
            dataS(3, 5) = ozn
        Case IZV_NEMA_PRIJEMA
            dataS(4, 5) = ozn
        Case SLED_OZN_FAK_NEISPRAVNA
            dataS(5, 5) = ozn
        Case Else
            If Len(ozn) > 0 Then dataS(1, 5) = ozn
    End Select

    Dim ctx As String
    ctx = Poruka("OTKUI_SL_DET_KARIKA_OTKUP") & " " & NzS(lanac(r, 2)) & "  " & _
          ChrW(183) & "  " & NzS(lanac(r, 4)) & "  " & ChrW(183) & "  " & _
          Poruka("OTKUI_SL_PERIOD") & " " & OpsegLabela() & "  " & ChrW(183) & "  " & _
          IIf(Len(ozn) = 0, Poruka("OTKUI_SL_POTPUN"), ozn)

    ' Info blok dokumenta (krug 6 S14): koren + prevoz + period + oznaka.
    Dim info(0 To 7) As String
    info(0) = NzS(lanac(r, 2))
    info(1) = NzS(lanac(r, 4))
    info(2) = NzS(lanac(r, 23))
    If IsDate(lanac(r, 1)) Then info(3) = Format$(CDate(lanac(r, 1)), "dd.MM.yyyy")
    info(4) = NzS(lanac(r, 24))
    info(5) = NzS(lanac(r, 13))
    info(6) = OpsegLabela()
    info(7) = ozn

    SlLanacZaPdf = Array(dataS, 5, ctx, info)
End Function

' "Stampaj dokument" nad izabranim redom: LANAC/PARCELE stampaju otkupni
' list (zrno reda); NEPOTPUNI rutira po vrsti karike. Red bez dokumenta
' ODBIJA porukom; zbirna nema svoju stampu i odbija s razlogom.
Private Sub StampajDokumentReda(ByVal red As Long)
    Dim ref As String, dokTip As String, dokID As String

    Select Case Scr_Lista()
        Case SL_LANAC, SL_PARC
            ref = NzS(modOtkupUI.GridCell(red, IIf(Scr_Lista() = SL_LANAC, 14, 12)))
            If Left$(ref, 4) = "OTK|" Then
                ReprintOtkupniListByOtkupID Mid$(ref, 5)
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
            End If
        Case SL_NEP
            dokTip = NzS(modOtkupUI.GridCell(red, 7))
            dokID = NzS(modOtkupUI.GridCell(red, 8))
            If Len(dokID) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
                Exit Sub
            End If
            Select Case dokTip
                Case DOK_TIP_OTKUP
                    ReprintOtkupniListByOtkupID dokID
                Case DOK_TIP_OTPREMNICA
                    OutputOtpremnicaPDF dokID
                Case DOK_TIP_PRIJEMNICA
                    PrintPrijemnica dokID
                Case SLED_DOK_ZBIRNA
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_SL_ZBIRNA_STAMPA"), True
                Case SLED_DOK_PRERADA
                    ' GP grana: kontradiktorna prerada -- ista ruta kao
                    ' meta PRERADA (preradni list).
                    If Len(ExportPreradaPDF(dokID, True)) = 0 Then _
                        modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
                Case SLED_DOK_UTOVAR
                    ' Krug 5b: utovarna lista ima obrazac -- stampa se
                    ' direktno (dokument koji ide sa robom).
                    modUtovar.PrintUtovar dokID
                    modOtkupUI.ShowToast Poruka("OTKUI_MSG_UT_STAMPA"), False
                Case Else
                    modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
            End Select
        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_IZ_NEMA_DOK"), True
    End Select
End Sub

'=====================================================================
' DIJAGNOSTIKA. Alt+F8 -> Diag_SlRedovi, pa Ctrl+G (N7 obrazac: bez ovoga
' se gubitak upita PRE ekrana i kvar POSLE ekrana ne razlikuju).
'=====================================================================
' DIJAGNOSTIKA (krug 7): "8. red dropdown-a je uvek veceg fonta".
' Nijedan put u kodu ne dira Font pop-redova (BuildPopup ih gradi
' identicno, PopRender menja samo caption/boju/top/sirinu/vidljivost,
' hover samo boju) -- zato se STVARNO stanje svih redova ljuskinog
' panela ispisuje ovde: Immediate + povratni string (za COM sondu).
' Radi nad otvorenom formom; bez nje ucita svoju headless (SetTestMode
' obavezno pre toga u tom slucaju) i istovari je na kraju.
' Alt+F8 -> Diag_SlPopFont dok je aplikacija otvorena.
Public Function Diag_SlPopFont() As String
    Dim f As Object, z As Object, c As Object, i As Long
    Dim svoja As Boolean, s As String

    On Error Resume Next
    If VBA.UserForms.count > 0 Then
        Set f = VBA.UserForms(0)
    Else
        Set f = VBA.UserForms.Add("frmOtkupUI")
        Dim n As Long
        n = f.Controls.count            ' okida UserForm_Initialize (par. 7.9)
        svoja = True
    End If
    If f Is Nothing Then
        Diag_SlPopFont = "forme nema i ne moze da se ucita"
        Debug.Print Diag_SlPopFont
        Exit Function
    End If

    s = SlPopIzvestaj(f)
    If svoja Then Unload f

    Debug.Print s
    Diag_SlPopFont = s
End Function

' DIJAGNOSTIKA (GP-1): "prvi put laguje" -- lag zivi u punjenju snimka
' (tri read-modela; kasniji ulasci su kes hit). Ovo meri SVAKI model u
' ms na PRAVOJ svesci, pod istim TableCache uslovima kao ekran, i ne
' dira kes ekrana. Alt+F8 -> Diag_SlPerf, ispis u Immediate (Ctrl+G);
' brojke presudjuju gde je vreme, umesto nagadjanja po kodu.
Public Function Diag_SlPerf() As String
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    Dim dOd As Date, dDo As Date, s As String
    Dim lanacV As Variant, problemiV As Variant, dokV As Variant

    dOd = CDate(SL_DAT_MIN)
    dDo = CDate(SL_DAT_MAX)

    BeginTableCache
    t0 = Timer
    lanacV = ReportSledljivostLanac(dOd, dDo)
    t1 = Timer
    problemiV = ReportSledljivostProblemi(dOd, dDo)
    t2 = Timer
    dokV = ReportSledljivostDokumenti(dOd, dDo)
    t3 = Timer
    EndTableCache

    s = "=== Diag_SlPerf (ceo period, hladan prolaz) ===" & vbLf & _
        "lanac:     " & Format$((t1 - t0) * 1000, "0") & " ms" & _
        IIf(IsArray(lanacV), "  (" & UBound(lanacV, 1) & " redova)", "  (prazno)") & vbLf & _
        "problemi:  " & Format$((t2 - t1) * 1000, "0") & " ms" & _
        IIf(IsArray(problemiV), "  (" & UBound(problemiV, 1) & " redova)", "  (prazno)") & vbLf & _
        "dokumenti: " & Format$((t3 - t2) * 1000, "0") & " ms" & _
        IIf(IsArray(dokV), "  (" & UBound(dokV, 1) & " redova)", "  (prazno)") & vbLf & _
        "UKUPNO:    " & Format$((t3 - t0) * 1000, "0") & " ms"
    Debug.Print s
    Diag_SlPerf = s
End Function

' Zajednicki citac stanja pop-redova za obe dijagnostike.
Private Function SlPopIzvestaj(ByVal f As Object) As String
    Dim z As Object, c As Object, i As Long, s As String
    On Error Resume Next
    Set z = f.Controls("zPop")
    If z Is Nothing Then
        SlPopIzvestaj = "zPop ne postoji na " & f.name & vbCrLf
        Exit Function
    End If
    s = "=== zPop / " & f.name & " (vis=" & z.Visible & ", h=" & _
        Format$(z.Height, "0.0") & ") ===" & vbCrLf
    For i = 0 To 13
        Set c = Nothing
        Set c = z.Controls("pop" & i)
        If c Is Nothing Then
            s = s & "pop" & i & ": NEMA" & vbCrLf
        Else
            s = s & "pop" & i & ": size=" & c.Font.Size & _
                "  name=" & c.Font.name & _
                "  bold=" & c.Font.bold & _
                "  h=" & Format$(c.Height, "0.0") & _
                "  top=" & Format$(c.top, "0.0") & _
                "  vis=" & c.Visible & vbCrLf
        End If
    Next i
    SlPopIzvestaj = s
End Function

Public Sub Diag_SlRedovi()
    Dim d As Variant, kolone As Variant, redovi As Variant, i As Long, k As Long, n As Long
    On Error Resume Next

    Debug.Print "--- Diag_SlRedovi (" & SCRSL_BUILD & ") ---"
    Debug.Print "  POSLEDNJI POZIV: filter=[" & mDiagFilter & "] q=[" & mDiagQ & _
                "] vraceno redova=" & CStr(mDiagN)
    Debug.Print "  SNIMAK: kljuc=[" & mSnimakKljuc & "] ok=" & SnimakPostoji() & _
                " punjenja=" & CStr(mSnimakPunjenja)

    d = Scr_Rows("sve", "")
    If Not IsArray(d) Then
        Debug.Print "  Scr_Rows NIJE vratio niz"
        Exit Sub
    End If

    kolone = d(0)
    redovi = d(1)
    n = CLng(d(2))
    Debug.Print "  kolona=" & CStr(UBound(kolone) + 1) & "  redova=" & CStr(n)

    For i = 0 To UBound(kolone)
        Debug.Print "  spec " & CStr(i + 1) & ": " & CStr(kolone(i))
    Next i

    If IsArray(redovi) Then
        For i = 1 To 3
            If i > n Then Exit For
            For k = 1 To UBound(kolone) + 1
                Debug.Print "  EKRAN red " & CStr(i) & " kol" & CStr(k) & ": tip=" & _
                            TypeName(redovi(i, k)) & " vred=[" & CStr(redovi(i, k)) & "]"
            Next k
        Next i
    End If

    For i = 1 To 3
        For k = 1 To UBound(kolone) + 1
            Debug.Print "  MREZA red " & CStr(i) & " kol" & CStr(k) & ": tip=" & _
                        TypeName(modOtkupUI.GridCell(i, k)) & _
                        " vred=[" & CStr(modOtkupUI.GridCell(i, k)) & "]"
        Next k
    Next i

    Err.Clear
End Sub

'=====================================================================
' TEST SEAM. Zona se u testu ne crta, pa se opseg ne cita iz kontrola.
' Seam koji MENJA stanje van test-rezima ne radi nista.
'=====================================================================
Public Sub Scr_SlTestSet(ByVal lista As String, ByVal odSerijski As Double, _
                         ByVal doSerijski As Double)
    If Not IsTestMode() Then Exit Sub
    If Len(lista) > 0 Then mLista = lista
    mTestOd = odSerijski
    mTestDo = doSerijski
End Sub

Public Sub Scr_SlIzaberiTest(ByVal otkupID As String)
    If Not IsTestMode() Then Exit Sub
    mIzabranOtkupID = otkupID
End Sub

Public Function Scr_SlSnimakPunjenjaTest() As Long
    If Not IsTestMode() Then Exit Function
    Scr_SlSnimakPunjenjaTest = mSnimakPunjenja
End Function

Public Function Scr_SlSnimakKljucTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_SlSnimakKljucTest = mSnimakKljuc
End Function

Public Function Scr_SlHintKljucTest() As String
    If Not IsTestMode() Then Exit Function
    Scr_SlHintKljucTest = mHintKljuc
End Function

Public Function Scr_SlKpiTest(ByVal koja As String) As Variant
    If Not IsTestMode() Then Exit Function
    Select Case koja
        Case "potpun":   Scr_SlKpiTest = mKpiPotpun
        Case "problemi": Scr_SlKpiTest = mKpiProblemi
    End Select
End Function

Public Sub Scr_SlTestReset()
    If Not IsTestMode() Then Exit Sub
    mLista = SL_LANAC
    mTestOd = 0
    mTestDo = 0
    mSnimakPunjenja = 0
    mIzabranOtkupID = ""
    mIzabranaZbirna = ""
    mDetalj = Empty
    mHintKljuc = ""
    Scr_ResetCache
    mSnimakKljuc = ""
    mCmbDokKljuc = ""
    mPovOtkupID = ""
    mCtxOd = 0
    mCtxDo = 0
End Sub
