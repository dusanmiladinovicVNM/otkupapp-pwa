Attribute VB_Name = "modScrDokumenti"
'=====================================================================
' modScrDokumenti - OPIS ekrana "Unos dokumenata" (F1..F7), faza S2b.
'
' Sve sto zna KOJI dokument gde zivi: tabela po rezimu (ModeTable), imena
' kolona po rezimu (Col*), sastav mreze (GridCols/ColumnSpec), zastavice
' rezima (ModeHas*), sifrarnici stanja (StatusCode/PayCode/KanalCode) i
' ikonica rezima. Ovde NEMA crtanja - samo odgovori na pitanja koja mreza
' postavlja.
'
' Zasto ovoliko i ne vise: ostatak ekrana (BuildForm, SelectModeCore,
' ApplyFormFields...) deli modul-level stanje sa ljuskom, pa bi njegovo
' premestanje bilo prepravka, ne premestanje. Merenje pre reza: od 32
' procedure koje diraju stanje mreze, njih 20 mesa mrezno i ekransko
' stanje. Zato u S2b izlazi samo ono sto je stvarno bez stanja - 30
' procedura koje su ciste funkcije rezima. Ostatak dolazi u S3, kad
' ugovor ekrana (Scr_Meta/Scr_Build/Scr_Grid/Scr_Event) da stanju gde
' da zivi.
'
' OVAJ MODUL JE SABLON ZA SVAKI SLEDECI EKRAN: palete, agrohemija,
' fakture, banka i ostali pisu svoj isti ovakav opis, a mrezu i fabriku
' kontrola ne diraju.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRDOK_BUILD As String = "v6-ui-143"

' Gde je Scr_Rows stigao - ime koraka ulazi u poruku o gresci.
Private mStep As String
' Danasnji datum i pocetak meseca; postavlja ih Scr_Rows, cita MatchFilterFast.
Private mToday As Double
Private mMonthStart As Double
' Kes izvedenih mapa ovog ekrana (iznosi faktura, kooperant po otkupu).
' Ranije su delile mPartMap sa ljuskom; posle preseljenja to bi bio poziv u
' njeno privatno telo. Prazni ga Scr_ResetCache.
Private mMape As Object

'--------------------------------------------------- RADNI STO OTPREMNICE
' F1 nije obicna "forma + lista". Otpremnica je IZVOR robe, a otkupni listovi
' (blokovi) su njen sastav po kooperantima; ekran postoji da operater vidi
' koliko od otpremnice jos nije povezano, i da je izda kad je sve povezano
' (odluka 14.8 t. 4). Zato F1 ima cetiri liste, koje se biraju prekidacem.
'
' Od S3b-2 radni sto stoji na KANONU (stari, na vezi Otkup.OtpremnicaID,
' obrisan je u S1b-3): sastav je tblOtpremnicaIzvori, ocekivano su stavke
' otpremnice, a svaku promenu pise modDokumenta (DodajOtpremnicaIzvor_TX,
' UkloniOtpremnicaIzvor_TX, IzdajOtpremnicu_TX). Ekran ne pise nijednu tabelu.
'
' AKTIVNA otpremnica je uvek NACRT: samo nacrt prima i pusta izvore. Izdata se
' vidi u listi, ali se ne bira; posle izdavanja ekran izlazi iz njenog
' konteksta.
'
' Identitet ide REDOM, ne brojem: liste OTPREMNICE i BLOKOVI nose ID u
' nevidljivoj poslednjoj koloni (prioritet 4), pa ga sortiranje ne odvaja od
' reda. Broj otpremnice je jedinstven tek po (stanica, dan).
'
' Specifikacija blokova (A-018, A-019, A-021) stoji nad ISTIM kanonom: oznaci
' vise otpremnica ili izaberi jednu pa "Stampaj specifikaciju", a "Po datumu"
' stampa ceo prikazani opseg. Oznaka je OtpremnicaID (ljuska kljuca red po
' nevidljivoj koloni identiteta), pa se broj nigde ne razresava u dokument.
'
' Lista NEVEZANI (A-025) pokazuje blokove bez aktivne otpremnice: upisane bez
' izabrane, uklonjene iz nacrta i oslobodjene stornom otpremnice. Odatle se
' vezuju za aktivni nacrt.
Private mLista As String          ' "SVI" | "OTPREMNICE" | "BLOKOVI" | "KOOPERANTI"
Private mOtpID As String          ' aktivna otpremnica (OtpremnicaID) -- NACRT
Private mOtpBroj As String        ' njen broj - za traku i naslov liste

' IZMENA NACRTA u F2 (odluka 19.09.2026): klik na nacrt u listi F2 popuni formu,
' a snimanje tada MENJA taj nacrt (UpdateOtpremnicaDraft_TX) umesto da pravi
' nov. Prazno = snimanje pravi nov nacrt. Ljuska izmenu otkazuje pri praznjenju
' forme i promeni rezima (Scr_IzmenaOtkazi).
Private mIzmenaOtpID As String
' Otvorena ISPRAVKA BLOKA (B-040, S3d-2): sledece snimanje F1 ne pravi nov
' dokument nego ZAMENU ovog -- stari se stornira, nastaje nov sa novim brojem.
' Ishodi RutaPosleUpisa -- imena, ne goli literali na tri mesta.
Public Const RUTA_LANAC As String = "LANAC"
Public Const RUTA_NACRT As String = "NACRT"
Private mIspravkaOtkupID As String
Private mIspravkaBroj As String
Private mIzmenaBroj As String

' Prekidac lista: "KLJUC|natpis|naslov mreze|sirina". Van F1 nema prekidaca -
' ostali rezimi imaju jednu listu, pa se dugmad ne prikazuju.
Public Function Scr_Liste() As Variant
    If modeKey(ActiveMode) <> "OTKUP" Then Exit Function
    ' Peto polje "opseg" trazi od ljuske polja OD / DO iznad liste (S3b-2b):
    ' specifikacija po datumu stampa tacno ono sto je u listi, pa granicu
    ' postavlja isti par koji listu i filtrira.
    Scr_Liste = Array( _
        "SVI|OTKUI_SEG_LS_SVI|OTKUI_GRID_TITLE_OTKUP|96", _
        "OTPREMNICE|OTKUI_SEG_LS_OTP|OTKUI_GRID_TITLE_OTPREMNICA|96|opseg", _
        "BLOKOVI|OTKUI_SEG_LS_BLOK|OTKUI_GRID_TITLE_BLOKOVI|110", _
        "NEVEZANI|OTKUI_SEG_LS_NEVEZ|OTKUI_GRID_TITLE_NEVEZANI|104", _
        "KOOPERANTI|OTKUI_SEG_LS_KOOP|OTKUI_GRID_TITLE_KOOP|100")
End Function

' Dopuna naslova mreze: u listi blokova stoji broj aktivne otpremnice.
Public Function Scr_NaslovDopuna() As String
    If Scr_Lista() = "BLOKOVI" Then Scr_NaslovDopuna = mOtpBroj
End Function

Public Function Scr_Cipovi() As String
    Scr_Cipovi = CipoviZaListu(Scr_Lista())
End Function

' Cipovi po listi. Izdvojeno iz Scr_Cipovi da bi pravilo moglo da se izmeri:
' lista otpremnica postoji samo u rezimu OTKUP.
'
' Lista dokumenata ('SVI') NE prijavljuje cipove: njeni zavise od rezima
' (ima li zbirnu, ima li fakturu), pa ostaju ljuskini.
' "otvorene" = NACRTI -- otpremnice koje jos primaju blokove.
Public Function CipoviZaListu(ByVal lista As String) As String
    If lista = "OTPREMNICE" Then _
        CipoviZaListu = "sve:OTKUI_CHIP_SVE:40|otvorene:OTKUI_CHIP_OTVORENE:132"
End Function

' Radnje nad redom za AKTIVNU listu: kljuc : natpis : sirina : stil : trebaRed.
' Kljuc se vraca u Scr_Event kao "act:<kljuc>:<red>".
Public Function Scr_Radnje() As String
    If modeKey(ActiveMode) <> "OTKUP" Then Exit Function
    Select Case Scr_Lista()
        Case "SVI"
            ' Ispravka bloka (B-040): pisac postoji od S1, a od S3d-2 ima i
            ' ulaz sa ekrana. Kapije su u modOtkup -- dugme ih ne ponavlja.
            Scr_Radnje = "print:OTKUI_BTN_RED_PRINT:116:ghost:1|" & _
                         "storno:OTKUI_BTN_RED_STORNO:88:danger:1|" & _
                         "ispravblok:OTKUI_BTN_RED_ISPRAVI_BLOK:96:soft:1"
            ' Vezivanje postojeceg otkupa za aktivnu otpremnicu: oporavak kad
            ' vezivanje posle unosa nije uspelo, i blok upisan pre izbora.
            If Len(mOtpID) > 0 Then _
                Scr_Radnje = Scr_Radnje & "|vezi:OTKUI_BTN_RED_VEZI:132:soft:1"
        Case "BLOKOVI"
            ' Blok u sastavu se NE stornira (kapija u modStorno.StornoOtkup) --
            ' prvo izlazi iz nacrta. Izdavanje ne trazi red: radi nad aktivnom.
            Scr_Radnje = "print:OTKUI_BTN_RED_PRINT:116:ghost:1|" & _
                         "ukloni:OTKUI_BTN_RED_UKLONI:136:danger:1|" & _
                         "ispravblok:OTKUI_BTN_RED_ISPRAVI_BLOK:96:soft:1|" & _
                         "izdaj:OTKUI_BTN_RED_IZDAJ:120:soft:0"
        Case "OTPREMNICE"
            ' Specifikacija (A-018, A-021): "2" znaci da radnja radi nad
            ' OZNACENIM redovima, a bez oznaka nad izabranim. "Po datumu"
            ' (A-019) ne trazi red -- stampa celu prikazanu listu.
            ' Ispravka izdate (B-022) trazi red: radi nad TACNO jednim
            ' dokumentom, pa oznake ne uzima -- dve otpremnice se ne
            ' ispravljaju jednim potezom.
            Scr_Radnje = "mark:OTKUI_BTN_RED_MARK:104:ghost:0|" & _
                         "spec:OTKUI_BTN_RED_SPEC:152:ghost:2|" & _
                         "specdat:OTKUI_BTN_RED_SPECDAT:96:ghost:0|" & _
                         "ispravi:OTKUI_BTN_RED_ISPRAVI:88:danger:1"
        Case "NEVEZANI"
            ' Blok bez otpremnice se vezuje za AKTIVNI nacrt -- isti pisac kao
            ' posle unosa; bez aktivne otpremnice radnje nema.
            '
            ' Hladnjacki blok se NE vezuje rucno: njemu je oporavak ponavljanje
            ' auto-lanca. Obe radnje stoje u redu radnji, a svaka na svojoj
            ' granici odbija blok koji joj ne pripada -- red jos ne moze da nosi
            ' svoje radnje (ugovor ljuske, v. plan 15).
            Scr_Radnje = "print:OTKUI_BTN_RED_PRINT:116:ghost:1"
            If Len(mOtpID) > 0 Then _
                Scr_Radnje = Scr_Radnje & "|vezi:OTKUI_BTN_RED_VEZI:132:soft:1"
            If modAutoHladnjaca.LanacUkljucen() Then _
                Scr_Radnje = Scr_Radnje & "|ponovi:OTKUI_BTN_RED_PONOVI:140:soft:1"
    End Select
End Function

' Koju listu F1 trenutno pokazuje. Van F1 uvek "SVI".
Public Function Scr_Lista() As String
    If modeKey(ActiveMode) <> "OTKUP" Then
        Scr_Lista = "SVI"
    ElseIf Len(mLista) = 0 Then
        Scr_Lista = "SVI"
    Else
        Scr_Lista = mLista
    End If
End Function

' Stanica aktivne otpremnice. Ljuska time proverava da li je promena otkupnog
' mesta izasla iz konteksta otpremnice (legacy cmbOtkupnoMesto_Change).
Public Function Scr_OtpStanica() As String
    On Error Resume Next
    If Len(mOtpID) = 0 Then Exit Function
    Scr_OtpStanica = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                mOtpID, COL_OTP_STANICA)))
End Function

' Napusti otpremnicu: njen datum i roba vise ne vaze, pa se lista vraca na
' otpremnice. Zovu je ljuska (promena otkupnog mesta) i izdavanje.
Public Sub Scr_OtpOtkazi()
    mOtpID = ""
    mOtpBroj = ""
    mLista = "OTPREMNICE"
End Sub

' TEST SEAM za ugovor ClearForm-a (datum aktivne otpremnice ostaje). Van
' test-rezima ne radi NISTA; produkcija bira otpremnicu kroz AktivirajOtpremnicu.
Public Sub Scr_OtpTestSet(ByVal otpID As String, ByVal broj As String)
    If Not IsTestMode() Then Exit Sub
    mOtpID = otpID
    mOtpBroj = broj
End Sub

' TEST SEAM: radni sto u pocetno stanje (bez otpremnice, lista SVI).
' Scr_OtpOtkazi to ne radi -- ona je produkcioni izlazak iz konteksta i
' namerno vodi na listu otpremnica. Van test-rezima ne radi nista.
Public Sub Scr_OtpTestReset()
    If Not IsTestMode() Then Exit Sub
    mOtpID = ""
    mOtpBroj = ""
    mLista = ""
End Sub

' OtpremnicaID aktivne otpremnice ("" = nema).
Public Function Scr_OtpID() As String
    Scr_OtpID = mOtpID
End Function

' Izbor aktivne otpremnice. Vraca "" kad je izabrana, inace razlog za operatera.
'
' Samo NACRT: izdata i stornirana ne primaju izvore, pa bi ih izbor prevario --
' sledeci upis bi pao tek na vezivanju. Status se trazi TACNO: prazan ili
' nepoznat status nije nacrt (isto pravilo kao RequireOtpDraft u piscu).
Public Function AktivirajOtpremnicu(ByVal otpID As String) As String
    On Error GoTo EH
    otpID = Trim$(otpID)
    AktivirajOtpremnicu = NacrtRazlog(otpID)
    If Len(AktivirajOtpremnicu) > 0 Then Exit Function
    mOtpID = otpID
    mOtpBroj = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))
    Exit Function
EH:
    AktivirajOtpremnicu = Poruka("OTKUI_ERR_RADNJA") & " " & Err.description
End Function

' "" kad je otpremnica aktivan NACRT, inace razlog za operatera. Jedno pravilo
' za izbor na radnom stolu (F1) i za izmenu nacrta (F2).
Private Function NacrtRazlog(ByVal otpID As String) As String
    Dim st As String
    If Len(otpID) = 0 Then
        NacrtRazlog = Poruka("OTKUI_ERR_OTP_NEPOZNATA")
        Exit Function
    End If
    If Len(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_ID)))) = 0 _
       Or UCase$(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, _
                                             COL_STORNIRANO)))) = "DA" Then
        NacrtRazlog = Poruka("OTKUI_ERR_OTP_NEPOZNATA")
        Exit Function
    End If
    st = UCase$(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, _
                                           COL_TRACE_IZDATO_STATUS))))
    If st <> UCase$(IZDATO_DRAFT) Then NacrtRazlog = Poruka("OTKUI_ERR_OTP_IZDATA")
End Function

' Otvara izmenu nacrta u F2. "" = otvorena (sledece snimanje menja nacrt), a
' spec je forma nacrta za ApplyPrefill; inace razlog. Izdata se ne menja: pisac
' trazi DRAFT, a izdat dokument se ispravlja stornom (S3c).
'
' Forma se sastavlja PRE otvaranja (review #363, P2): kanonski citalac koji
' padne prekida otvaranje i ostavlja izmenu ZATVORENU -- nikad otvorena izmena
' nad delimicnom formom.
Public Function OtvoriIzmenuNacrta(ByVal otpID As String, Optional ByRef spec As String) As String
    spec = ""
    On Error GoTo EH
    otpID = Trim$(otpID)
    OtvoriIzmenuNacrta = NacrtRazlog(otpID)
    If Len(OtvoriIzmenuNacrta) > 0 Then Exit Function
    spec = PrefillNacrta(otpID)
    mIzmenaOtpID = otpID
    mIzmenaBroj = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))
    Exit Function
EH:
    Scr_IzmenaOtkazi
    spec = ""
    OtvoriIzmenuNacrta = Poruka("OTKUI_ERR_RADNJA") & " " & Err.description
End Function

' Otvara ISPRAVKU BLOKA u F1 (B-040). "" = otvorena (sledece snimanje pravi
' zamenu), a `spec` je forma starog bloka za ApplyPrefill; inace razlog.
'
' Kapije se NE prepisuju ovde: pita se modOtkup.IspravkaOtkupaRazlog, isti izvor
' pravila koji pisac dize kao gresku. Ekran ga zove PRE forme da operater ne
' kuca zamenu za dokument koji je ne moze primiti (storniran, vec zamenjen, ili
' izvor IZDATE otpremnice).
'
' Forma se sastavlja PRE otvaranja: citalac koji padne prekida otvaranje i
' ostavlja ispravku ZATVORENU -- nikad otvorena ispravka nad delimicnom formom
' (isti razlog kao izmena nacrta, review #363 P2).
'
' BROJ se ne prenosi: ispravka je nov dokument i dobija NOV broj (A9). To vec
' drzi PrefillIzStorniranog, koji broj namerno izostavlja.
Public Function OtvoriIspravkuBloka(ByVal otkupID As String, _
                                    Optional ByRef spec As String) As String
    Dim brDok As String
    spec = ""
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    OtvoriIspravkuBloka = modOtkup.IspravkaOtkupaRazlog(otkupID)
    If Len(OtvoriIspravkuBloka) > 0 Then Exit Function

    brDok = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_BR_DOK)))
    spec = modStornoDok.PrefillIzStorniranog(STIP_OTKUP, brDok, otkupID)
    If Len(spec) = 0 Then
        OtvoriIspravkuBloka = Poruka("OTKUI_ERR_ISPRAVI_BLOK") & " " & brDok
        Exit Function
    End If

    mIspravkaOtkupID = otkupID
    mIspravkaBroj = brDok
    Exit Function
EH:
    Scr_IspravkaOtkazi
    spec = ""
    OtvoriIspravkuBloka = Poruka("OTKUI_ERR_RADNJA") & " " & Err.description
End Function

' STA SE RADI SA BLOKOM POSLE UPISA -- jedna odluka, bez forme, pa se meri.
'
'   RUTA_LANAC  obavezan hladnjacki lanac
'   RUTA_NACRT  rucno vezivanje za aktivan nacrt
'   ""          nista
'
' `outGreska` != "" znaci da se put NE ZNA: tada se ne radi NISTA (fail-closed),
' jer bi i lanac i nacrt bili pogadjanje.
'
' Ugovor je DETERMINISTICAN i ne zavisi od toga sta je slucajno otvoreno na
' ekranu:
'
'   ispravka bloka koji je BIO u nacrtu -> ""      (pisac je clanstvo vec preneo)
'   ispravka SLOBODNOG hladnjackog bloka -> LANAC  (lanac je obavezan i za nju)
'   ispravka slobodnog obicnog bloka     -> ""     (naslednik ostaje slobodan)
'   nov unos                             -> LANAC ili NACRT, po pravilima
'
' Zasto ispravka slobodnog obicnog bloka NE ide u nacrt: naslednik bi upao u
' dokument koji je slucajno otvoren u tom trenutku, iako original nije bio ni u
' jednom. Ispravka menja dokument, ne njegovu pripadnost.
Public Function RutaPosleUpisa(ByVal otkupID As String, _
                               ByVal jeIspravka As Boolean, _
                               ByVal imaoRoditelja As Boolean, _
                               Optional ByRef outGreska As String) As String
    Dim ideULanac As Boolean

    outGreska = ""
    If jeIspravka And imaoRoditelja Then Exit Function

    ideULanac = modAutoHladnjaca.LanacVaziZaBlok(otkupID, outGreska)
    If Len(outGreska) > 0 Then Exit Function
    If ideULanac Then
        RutaPosleUpisa = RUTA_LANAC
        Exit Function
    End If

    If jeIspravka Then Exit Function
    If Len(mOtpID) > 0 Then RutaPosleUpisa = RUTA_NACRT
End Function

' Otkazuje ispravku bloka -- sledece snimanje opet pravi NOV dokument.
Public Sub Scr_IspravkaOtkazi()
    mIspravkaOtkupID = ""
    mIspravkaBroj = ""
End Sub

' OtkupID bloka koji je otvoren za ispravku ("" = nema).
Public Function Scr_IspravkaOtkupID() As String
    Scr_IspravkaOtkupID = mIspravkaOtkupID
End Function

' Otkazuje izmenu nacrta -- sledece snimanje pravi nov nacrt.
Public Sub Scr_IzmenaOtkazi()
    mIzmenaOtpID = ""
    mIzmenaBroj = ""
End Sub

' OtpremnicaID nacrta koji je otvoren za izmenu ("" = nema).
Public Function Scr_IzmenaOtpID() As String
    Scr_IzmenaOtpID = mIzmenaOtpID
End Function

' Opis aktivne otpremnice za traku iznad forme. Prazno = nema izabrane.
' Oblik (13 polja):
'   broj | stanica | datum | ocekKg | povKg | ostKg | ocekAmb | povAmb | ostAmb |
'   cenaI | semafor | ostatak po klasi | cenaII
' semafor: -1 = neka klasa je PREKORACENA (kg ili gajbe), 0 = sve klase su na
' nuli (spremna za izdavanje), 1 = u toku. Racuna se PO KLASI: izdavanje trazi
' povezano = ocekivano za svaku klasu, pa zbir preko klasa (+20 u I, -20 u II)
' ne sme da izgleda kao "spremna".
Public Function Scr_OtpInfo() As String
    Dim prog As Object, k As Variant, r As Object
    Dim ukKg As Double, blKg As Double, ukAmb As Double, blAmb As Double
    Dim prek As Boolean, sveNula As Boolean, poKlasi As String, sem As String
    Dim dat As String, cenaII As Double
    On Error GoTo EH
    If Len(mOtpID) = 0 Then Exit Function

    Set prog = modDokumenta.GetOtpremnicaProgress(mOtpID)
    sveNula = True
    For Each k In prog.Keys
        Set r = prog(k)
        ukKg = ukKg + CDbl(r("ocekivano"))
        blKg = blKg + CDbl(r("povezano"))
        ukAmb = ukAmb + CDbl(r("ocekivanoAmb"))
        blAmb = blAmb + CDbl(r("povezanoAmb"))
        If CDbl(r("preostalo")) < -0.0001 Or CDbl(r("preostaloAmb")) < -0.0001 Then prek = True
        If Abs(CDbl(r("preostalo"))) > 0.0001 Or Abs(CDbl(r("preostaloAmb"))) > 0.0001 Then sveNula = False
        If Len(poKlasi) > 0 Then poKlasi = poKlasi & " / "
        poKlasi = poKlasi & CStr(k) & " " & modStornoDok.KgTekst(CDbl(r("preostalo")))
    Next k
    If prog.count < 2 Then poKlasi = ""
    If prek Then
        sem = "-1"
    ElseIf sveNula Then
        sem = "0"
    Else
        sem = "1"
    End If

    dat = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mOtpID, COL_OTP_DATUM))
    If IsDate(dat) Then dat = Format$(CDate(dat), "dd.mm.yyyy.")
    cenaII = CenaKlase(mOtpID, KLASA_II)

    ' Brojevi IDU KROZ NumStr: Str$ uvek pise tacku, a ljuska cita sa Val(),
    ' koji staje na zarezu -- lokalni zarez bi 60,6 kg pretvorio u 60.
    Scr_OtpInfo = mOtpBroj & "|" & _
                  OtpKupacNaziv(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                     mOtpID, COL_OTP_STANICA))) & "|" & _
                  dat & "|" & _
                  NumStr(ukKg) & "|" & NumStr(blKg) & "|" & NumStr(ukKg - blKg) & "|" & _
                  NumStr(ukAmb) & "|" & NumStr(blAmb) & "|" & NumStr(ukAmb - blAmb) & "|" & _
                  NumStr(CenaKlase(mOtpID, KLASA_I)) & "|" & sem & "|" & poKlasi & "|" & _
                  IIf(cenaII > 0, NumStr(cenaII), "")
    Exit Function
EH:
    ' Traka ne sme da obori ekran, a ni da cuti: broj ostaje, a umesto brojeva
    ' stoji opis greske (ljuska crta samo broj i podnaslov kad polja fale).
    Scr_OtpInfo = mOtpBroj & "|" & Poruka("OTKUI_ERR_RADNJA") & " " & Err.description
End Function

' Cena klase za traku i prefill: ona po kojoj su vec pisani blokovi te klase
' (prosek izvora, vrednost / kg), a tek ako njih nema -- predlog sa stavke
' otpremnice (odluka 14.8 t. 2). Nula = nema ni jednog ni drugog.
Private Function CenaKlase(ByVal otpID As String, ByVal klasa As String) As Double
    Dim vred As Object, rec As Variant, poDok As Object, c As Collection
    Dim red As Variant, i As Long
    Set vred = modDokumenta.VrednostIzvoraPoOtpremnici()
    If vred.Exists(otpID & "|" & klasa) Then
        rec = vred(otpID & "|" & klasa)
        If CDbl(rec(0)) > 0 Then
            CenaKlase = CDbl(rec(1)) / CDbl(rec(0))
            Exit Function
        End If
    End If
    Set poDok = modDokumenta.StavkeOtpremnicePoDokumentu()
    If Not poDok.Exists(otpID) Then Exit Function
    Set c = poDok(otpID)
    For i = 1 To c.count
        red = c(i)
        If UCase$(Trim$(CStr(red(3)))) = klasa Then
            If IsNumeric(red(5)) Then CenaKlase = CDbl(red(5))
            Exit Function
        End If
    Next i
End Function

' Broj u niz koji ne zavisi od lokalnih podesavanja - Str$ uvek pise tacku
' kao decimalni znak i nikad ne grupise hiljade, tacno onako kako ga Val()
' na drugoj strani i ocekuje.
Private Function NumStr(ByVal v As Double) As String
    NumStr = Trim$(Str$(v))
End Function

Private Function OtpKupacNaziv(ByVal stanicaID As String) As String
    OtpKupacNaziv = stanicaID
    If Len(stanicaID) = 0 Then Exit Function
    On Error Resume Next
    OtpKupacNaziv = NzToText(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"))
    If Len(OtpKupacNaziv) = 0 Then OtpKupacNaziv = stanicaID
End Function

' Koliko blok prelazi ostatak aktivne otpremnice, po klasi: npr. "12,50 kg (I)",
' ili "" kad ne prelazi. Klasa koju otpremnica uopste ne ocekuje ima ostatak
' nula -- svaki kilogram u njoj je prekoracenje (izdavanje bi ga odbilo).
' Read-model koji padne ovde PROPAGIRA gresku -- pozivalac odlucuje.
Public Function PrekoracenjeOpis(ByVal kgI As Double, ByVal kgII As Double) As String
    Dim prog As Object, deoI As String, deoII As String
    If Len(mOtpID) = 0 Then Exit Function
    Set prog = modDokumenta.GetOtpremnicaProgress(mOtpID)
    deoI = PrekoracenjeKlase(prog, KLASA_I, kgI)
    deoII = PrekoracenjeKlase(prog, KLASA_II, kgII)
    PrekoracenjeOpis = deoI
    If Len(deoII) > 0 Then
        If Len(PrekoracenjeOpis) > 0 Then PrekoracenjeOpis = PrekoracenjeOpis & ", "
        PrekoracenjeOpis = PrekoracenjeOpis & deoII
    End If
End Function

Private Function PrekoracenjeKlase(ByVal prog As Object, ByVal klasa As String, _
                                   ByVal kg As Double) As String
    Dim ost As Double
    If kg <= 0 Then Exit Function
    If prog.Exists(klasa) Then ost = CDbl(prog(klasa)("preostalo"))
    If kg <= ost + 0.0001 Then Exit Function
    PrekoracenjeKlase = modStornoDok.KgTekst(kg - ost) & " kg (" & klasa & ")"
End Function

' Veze otkup za aktivnu otpremnicu. "" = vezan, inace razlog za operatera.
' HLADNJACKI BLOK SE NE VEZUJE RUCNO. Kapija stoji na GRANICI RADNJE, ne samo u
' tome koje se dugme crta: oporavak posle pale automatike je "ispravi uzrok pa
' PONOVI lanac", a ne "veze ga za prvu otvorenu otpremnicu". Bez ove kapije bi se
' obavezan lanac zaobilazio jednim klikom iz liste "Bez otpremnice".
'
' Neizvesnost takodje odbija: blok za koji se ne zna kojim putem ide ne sme da
' zavrsi u rucnom dokumentu.
Public Function VeziZaAktivnu(ByVal otkupID As String) As String
    Dim g As String, putGreska As String
    On Error GoTo EH
    If modAutoHladnjaca.LanacVaziZaBlok(Trim$(otkupID), putGreska) Then
        VeziZaAktivnu = Poruka("OTKUI_ERR_VEZA_HLADNJACA")
        Exit Function
    End If
    If Len(putGreska) > 0 Then
        VeziZaAktivnu = putGreska
        Exit Function
    End If
    If Len(mOtpID) = 0 Then
        VeziZaAktivnu = Poruka("OTKUI_ERR_NEMA_AKT_OTP")
        Exit Function
    End If
    If Not modDokumenta.DodajOtpremnicaIzvor_TX(mOtpID, Trim$(otkupID), g) Then
        VeziZaAktivnu = Poruka("OTKUI_ERR_VEZA_OTP") & " " & g
    End If
    Exit Function
EH:
    VeziZaAktivnu = Poruka("OTKUI_ERR_VEZA_OTP") & " " & Err.description
End Function

' Izvadi otkup iz aktivne otpremnice. "" = uklonjen, inace razlog.
Public Function UkloniIzAktivne(ByVal otkupID As String) As String
    Dim g As String
    On Error GoTo EH
    If Len(mOtpID) = 0 Then
        UkloniIzAktivne = Poruka("OTKUI_ERR_NEMA_AKT_OTP")
        Exit Function
    End If
    If Not modDokumenta.UkloniOtpremnicaIzvor_TX(mOtpID, Trim$(otkupID), g) Then
        UkloniIzAktivne = Poruka("OTKUI_ERR_UKLONI") & " " & g
    End If
    Exit Function
EH:
    UkloniIzAktivne = Poruka("OTKUI_ERR_UKLONI") & " " & Err.description
End Function

' Izdaje aktivnu otpremnicu. "" = izdata, inace razlog -- pisac imenuje klasu i
' brojeve kad povezano nije jednako ocekivanom. Posle izdavanja otpremnica vise
' ne prima izvore, pa ekran izlazi iz njenog konteksta.
Public Function IzdajAktivnu() As String
    Dim g As String
    On Error GoTo EH
    If Len(mOtpID) = 0 Then
        IzdajAktivnu = Poruka("OTKUI_ERR_NEMA_AKT_OTP")
        Exit Function
    End If
    If Not modDokumenta.IzdajOtpremnicu_TX(mOtpID, g) Then
        IzdajAktivnu = Poruka("OTKUI_ERR_IZDAJ") & " " & g
        Exit Function
    End If
    Scr_OtpOtkazi
    Exit Function
EH:
    IzdajAktivnu = Poruka("OTKUI_ERR_IZDAJ") & " " & Err.description
End Function

'--------------------------------------------------------- UGOVOR EKRANA
' Prva tacka ugovora iz modUiScreens. Sluzi dvostruko: opisuje ekran i
' javlja registru da modul POSTOJI - registar ga trazi bas ovim pozivom
' (Application.Run), jer rano vezivanje bi oborilo compile klijentu kome
' neki ekranski modul nedostaje.
'
' Ostatak ugovora (Scr_Build / Scr_Layout / Scr_Grid / Scr_Event /
' Scr_Save) dolazi u S3b, kad se stanje mreze i forme preseli ovamo iz
' ljuske. Do tada ljuska crta ovaj ekran po starom.
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=DOKUMENTI|naslov=OTKUI_NAV_UNOS|oblik=forma+mreza|rezima=7"
End Function

' Radnje ovog ekrana. Ljuska ne zna nijednu - prosledjuje tag i, ako je ekran
' vratio True, osvezi mrezu i traku.
'   lsSVI / lsKOOPERANTI - prekidac liste u F1
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    On Error Resume Next
    ' F2: klik na NACRT otvara njegovu izmenu (odluka 19.09.2026). Nista se ne
    ' upisuje -- forma se samo popuni, pa mreza ne mora da se cita ponovo.
    If modeKey(ActiveMode) = "OTPREMNICA" Then
        If Left$(tag, 4) = "row:" Then IzaberiNacrtZaIzmenu CLng(Mid$(tag, 5))
        Exit Function
    End If
    If modeKey(ActiveMode) <> "OTKUP" Then Exit Function

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        Scr_Event = True
        Exit Function
    End If

    ' Radnja nad izabranim redom: "act:<sta>:<red>". Ljuska zna samo koji je
    ' red izabran; sta se nad njim radi zna ovaj modul, a POSAO rade postojece
    ' rutine (modPrint / modStorno) - ovde se nista ne upisuje rucno.
    If Left$(tag, 4) = "act:" Then
        Scr_Event = RowAction(tag)
        Exit Function
    End If

    ' Klik na red u listi otpremnica BIRA aktivnu otpremnicu. ID dolazi iz
    ' nevidljive kolone reda, ne iz broja.
    If Left$(tag, 4) = "row:" And Scr_Lista() = "OTPREMNICE" Then
        Dim razlog As String
        razlog = AktivirajOtpremnicu(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), _
                                                               UBound(OtpGridCols()) + 1)))
        If Len(razlog) > 0 Then
            modOtkupUI.ShowToast razlog, True
            Exit Function
        End If
        ' Sve sto otpremnica zna o robi prepisuje se u formu; operateru ostaju
        ' kooperant i kolicine.
        modOtkupUI.ApplyPrefill PrefillSpec(mOtpID)
        ' izbor otpremnice vodi pravo na njene blokove - to je sledeci potez
        mLista = "BLOKOVI"
        Scr_Event = True
    End If
End Function

'---------------------------------------- HLADNJACA ISPRAVKA (Faza D/13)
' Autohladnjaca: jedan otkupni list na hladnjaci povlaci ceo lanac
' (otpremnica -> zbirna -> prijemnica -> palete). Storno tog otkupa obara
' ceo lanac, ali PALETE ostaju - one su fizicke, roba je stvarno na njima.
'
' Zato se posle storna pita sta sa njima. Kontekst se cita PRE storna:
' posle njega otkup vise nije aktivan pa se veza ka prijemnici gubi.
' Prazan palInfo znaci "nema sta da se pita" (nije hladnjaca, nema zbirne
' ili prijemnica nije paletizovana).
Private Sub HladnjacaLanac(ByVal otkupID As String, ByRef prijBroj As String, _
                           ByRef palInfo As String)
    Dim stanica As String, zbirna As String
    On Error Resume Next
    prijBroj = "": palInfo = ""
    stanica = NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA))
    zbirna = NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_BROJ_ZBIRNE))
    If Len(zbirna) = 0 Then Exit Sub
    If Not IsHladnjacaStanica(stanica) Then Exit Sub
    prijBroj = NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, zbirna, COL_PRJ_BROJ))
    If Len(prijBroj) = 0 Then Exit Sub
    palInfo = GetPaleteInfoForPrijemnicaBroj(prijBroj)
End Sub

' Tri ishoda, isti kao u legacy modOtkupBlok.OfferHladnjacaIspravka:
'   DA     ISPRAVKA - zapamti pending relink i prefiluj otkup iz storniranog;
'          po Unosu se lanac gradi bez sveze paletizacije, a stare palete se
'          prevezuju (posao vec radi modOtkupUnos.OtkupUpisi)
'   NE     DUPLI UNOS - roba nije primljena dvaput, pa se fantomske stavke
'          odmah skidaju; paleta koja ostane prazna se stornira
'   OTKAZI nista - palete ostaju osirocene i dalje broje robu
Private Sub PonudiHladnjacaIspravku(ByVal brDok As String, ByVal otkupID As String, _
                                    ByVal prijBroj As String, ByVal palInfo As String)
    Dim odg As VbMsgBoxResult, info As String, spec As String
    On Error GoTo EH
    odg = MsgBox(Poruka("OTKUI_MSG_HLAD_LANAC") & vbCrLf & _
                 Poruka("OTKUI_MSG_HLAD_PALETE") & " " & prijBroj & " (" & palInfo & ")" & vbCrLf & vbCrLf & _
                 Poruka("OTKUI_ASK_HLAD"), vbQuestion + vbYesNoCancel, APP_NAME)

    If odg = vbYes Then
        SetHladnjacaRelinkPending prijBroj
        ' Prefill bez correction context-a: ovaj tok ne stvara zapis u
        ' tblStornoVeza (veza se cuva kroz pending relink). Polazi se od
        ' OtkupID-a storniranog reda, ne od broja (S1e).
        spec = modStornoDok.PrefillIzStorniranog(STIP_OTKUP, brDok, otkupID)
        If Len(spec) > 0 Then modOtkupUI.ApplyPrefill spec
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_HLAD_ISPRAVKA"), False
    ElseIf odg = vbNo Then
        If DetachOsirocenePaletaStavke_TX(prijBroj, info) > 0 Then
            MsgBox info, vbInformation, APP_NAME
        Else
            MsgBox Poruka("OTKUI_MSG_HLAD_NISTA") & IIf(Len(info) > 0, " " & info, ""), _
                   vbExclamation, APP_NAME
        End If
    Else
        MsgBox Poruka("OTKUI_MSG_HLAD_OSIROCENE"), vbInformation, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modScrDokumenti.PonudiHladnjacaIspravku"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Sub

' Vraca True ako je radnja PROMENILA podatke (pa mreza mora ponovo da se cita).
' Stampa vraca False - ona nista ne menja.
'
' IDENTITET REDA JE OtkupID iz nevidljive kolone (S1e), ne broj: broj je
' jedinstven tek po (otkupno mesto, dan), pa stampa ili storno po broju hvata i
' tudji dokument (AUD-057). Broj sluzi samo za poruke operateru.
Private Function RowAction(ByVal tag As String) As Boolean
    Dim p() As String, red As Long, broj As String, otkupID As String
    On Error GoTo EH
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    ' Ko trazi red, trazi ga sam - odmah ispod. Prva kolona je BROJ dokumenta;
    ' GridCell na red 0 vraca prazno.
    broj = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    ' Identitet reda: svaka lista ima svoje kolone, pa i svoj indeks. Lista
    ' otpremnica ne nosi otkup nego OtpremnicaID -- cita ga specifikacija.
    Select Case Scr_Lista()
        Case "BLOKOVI"
            otkupID = Trim$(CStr(modOtkupUI.GridCell(red, UBound(BlokGridCols()) + 1)))
        Case "NEVEZANI"
            otkupID = Trim$(CStr(modOtkupUI.GridCell(red, UBound(NevGridCols()) + 1)))
        Case "OTPREMNICE"
            otkupID = ""
        Case Else
            otkupID = Trim$(CStr(modOtkupUI.GridCell(red, IdentKolonaIndeks("OTKUP"))))
    End Select
    Dim razlog As String
    Select Case p(0)
        ' "ispravi" NIJE ovde: to je ispravka OTPREMNICE, a lista otpremnica
        ' nema OtkupID u redu. Kljucevi su zato razdvojeni -- isti kljuc za dve
        ' radnje je jednom vec ugasio onu drugu (review #368).
        Case "print", "storno", "vezi", "ukloni", "ponovi", "ispravblok"
            If Len(otkupID) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
                Exit Function
            End If
    End Select

    Select Case p(0)
        Case "print"
            modPrint.OutputOtkupniList otkupID
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_STAMPA") & " " & broj, False

        Case "storno"
            ' Razlog zbog kog pisac odbija storno (izvor aktivne otpremnice)
            ' operater cuje PRE potvrde -- isti preflight koji koristi F8.
            razlog = modStornoDok.StornoRazlog(modStornoDok.STIP_OTKUP, broj, "", otkupID)
            If Len(razlog) > 0 Then
                modOtkupUI.ShowToast razlog, True
                Exit Function
            End If
            If MsgBox(Poruka("OTKUI_ASK_STORNO") & " " & broj & _
                      Poruka("OTKUI_ASK_STORNO2"), vbQuestion + vbYesNo, _
                      APP_NAME) = vbNo Then Exit Function
            ' Autohladnjaca: kontekst lanca se cita PRE storna - posle njega
            ' otkup vise nije aktivan, pa se veza ka prijemnici ne bi nasla.
            Dim hlPrij As String, hlPal As String
            HladnjacaLanac otkupID, hlPrij, hlPal
            If modStorno.StornoOtkup_TX(otkupID) Then
                Scr_ResetCache
                RowAction = True
                If Len(hlPal) > 0 Then
                    PonudiHladnjacaIspravku broj, otkupID, hlPrij, hlPal
                Else
                    modOtkupUI.ShowToast Poruka("OTKUI_MSG_STORNIRANO") & " " & broj, False
                End If
            Else
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_STORNO") & " " & broj, True
            End If

        Case "ponovi"
            ' Ponavljanje auto-lanca posle otklonjenog uzroka. Blok koji lancu ne
            ' pripada se odbija imenom -- ista kapija kao kod vezivanja, samo sa
            ' druge strane.
            Dim ponPoruka As String, ponPut As String
            If Not modAutoHladnjaca.LanacVaziZaBlok(otkupID, ponPut) Then
                modOtkupUI.ShowToast IIf(Len(ponPut) > 0, ponPut, _
                                         Poruka("OTKUI_ERR_PONOVI_NIJE_HLAD")), True
                Exit Function
            End If
            If Len(modAutoHladnjaca.AutoLanacHladnjaca(otkupID, ponPoruka)) = 0 Then
                modOtkupUI.ShowToast IIf(Len(ponPoruka) > 0, ponPoruka, _
                                         Poruka("OTKUI_ERR_LANAC")), True
                Exit Function
            End If
            Scr_ResetCache
            RowAction = True
            modOtkupUI.ShowToast ponPoruka, False

        Case "ispravblok"
            ' Ispravka bloka: forma se puni starim podacima, a SLEDECE snimanje
            ' pravi ZAMENU (nov broj, nov ID, stari storniran). Ako je blok u
            ' NACRTU otpremnice, nacrt u istom potezu prelazi na naslednika.
            Dim ispSpec As String, ispRazlog As String
            ispRazlog = OtvoriIspravkuBloka(otkupID, ispSpec)
            If Len(ispRazlog) > 0 Then
                modOtkupUI.ShowToast ispRazlog, True
                Exit Function
            End If
            modOtkupUI.ApplyPrefill ispSpec
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_ISPRAVKA_BLOKA") & " " & mIspravkaBroj, False

        Case "vezi"
            razlog = VeziZaAktivnu(otkupID)
            If Len(razlog) > 0 Then
                modOtkupUI.ShowToast razlog, True
                Exit Function
            End If
            Scr_ResetCache
            RowAction = True
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_VEZAN") & " " & mOtpBroj & ": " & broj, False

        Case "ukloni"
            razlog = UkloniIzAktivne(otkupID)
            If Len(razlog) > 0 Then
                modOtkupUI.ShowToast razlog, True
                Exit Function
            End If
            Scr_ResetCache
            RowAction = True
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_UKLONJEN") & " " & mOtpBroj & ": " & broj, False

        Case "izdaj"
            If Len(mOtpID) = 0 Then
                modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_AKT_OTP"), True
                Exit Function
            End If
            ' Izdavanje je konacno za sastav i knjizi izlaz gajbi -- pita se.
            If MsgBox(Poruka("OTKUI_ASK_IZDAJ") & " " & mOtpBroj & Poruka("OTKUI_ASK_IZDAJ2"), _
                      vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function
            Dim brIzd As String
            brIzd = mOtpBroj
            razlog = IzdajAktivnu()
            If Len(razlog) > 0 Then
                modOtkupUI.ShowToast razlog, True
                Exit Function
            End If
            Scr_ResetCache
            RowAction = True
            modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZDATA") & " " & brIzd, False

        Case "spec"
            ' Stampa nista ne menja -- mreza se ne cita ponovo (False).
            SpecZaIzbor red

        Case "specdat"
            SpecPoDatumu

        Case "ispravi"
            RowAction = IspraviIzdatu(red)

        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
    Exit Function
EH:
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Ispravka izdate otpremnice (B-022): jedan potez -- stara se stornira, a nova
' nastaje kao nacrt sa istim zaglavljem, ocekivanjem i blokovima. Posle toga je
' nova AKTIVNA otpremnica, pa operater odmah radi nad njom (doda ili ukloni
' blok, doradi ocekivanje u F2, izda).
'
' Identitet je OtpremnicaID iz nevidljive kolone reda, ne broj -- isto pravilo
' kao specifikacija i storno.
Private Function IspraviIzdatu(ByVal red As Long) As Boolean
    Dim oid As String, stariBroj As String, greska As String
    Dim novaID As String, razlog As String

    If red <= 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If
    oid = Trim$(CStr(modOtkupUI.GridCell(red, UBound(OtpGridCols()) + 1)))
    If Len(oid) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_OTP"), True
        Exit Function
    End If
    stariBroj = Trim$(CStr(modOtkupUI.GridCell(red, 1)))

    If MsgBox(Poruka("OTKUI_ASK_ISPRAVI") & " " & stariBroj & _
              Poruka("OTKUI_ASK_ISPRAVI2"), vbQuestion + vbYesNo, _
              APP_NAME) = vbNo Then Exit Function

    novaID = modDokumenta.IspravkaOtpremnice_TX(oid, greska)
    If Len(novaID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_ISPRAVI") & " " & greska, True
        Exit Function
    End If

    Scr_ResetCache
    IspraviIzdatu = True

    ' Nova je nacrt, pa sme da bude aktivna. Ako izbor ne uspe, ispravka je
    ' ipak gotova -- razlog se kaze, a podaci se ne diraju.
    razlog = AktivirajOtpremnicu(novaID)
    If Len(razlog) > 0 Then
        modOtkupUI.ShowToast razlog, True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_ISPRAVLJENA") & " " & mOtpBroj, False
End Function

'--------------------------------------- SPECIFIKACIJA BLOKOVA (A-018, A-019)
' Oznacene otpremnice, a kad oznaka nema -- ona na izabranom redu. Kljuc oznake
' je OtpremnicaID (ljuska kljuca red po koloni identiteta), pa broja u ovom
' putu nema: stari tok je oznake drzao po broju i posle ih razresavao nazad u
' dokument, a broj je jedinstven tek po (stanica, dan).
Private Sub SpecZaIzbor(ByVal red As Long)
    Dim ids As Collection, k As Variant, kljucevi As String, oid As String
    Set ids = New Collection
    kljucevi = modOtkupUI.MarkedKeys()
    If Len(kljucevi) > 0 Then
        For Each k In Split(kljucevi, "|")
            If Len(Trim$(CStr(k))) > 0 Then ids.Add Trim$(CStr(k))
        Next k
    ElseIf red > 0 Then
        oid = Trim$(CStr(modOtkupUI.GridCell(red, UBound(OtpGridCols()) + 1)))
        If Len(oid) > 0 Then ids.Add oid
    End If
    If ids.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_OTP"), True
        Exit Sub
    End If
    StampajSpecifikaciju ids
End Sub

' "Po datumu" stampa SVE STO JE PRIKAZANO: opseg, cip i pretraga zajedno, preko
' svih strana. ID-evi se citaju iz mreze ljuske -- to je tacno ono sto operater
' vidi, pa se lista i stampa ne mogu razici. Bez oba datuma nema "po datumu".
Private Sub SpecPoDatumu()
    Dim ids As Collection
    If DatGranica(modOtkupUI.GridDatOd()) = 0 Or DatGranica(modOtkupUI.GridDatDo()) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SPEC_OPSEG"), True
        Exit Sub
    End If
    Set ids = SpecOtpremniceIzPrikaza()
    If ids.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_FILT"), True
        Exit Sub
    End If
    StampajSpecifikaciju ids
End Sub

' OtpremnicaID-evi svih redova koje mreza drzi, redom prikaza. Javno zbog testa
' opsega: mreza je jedino mesto koje zna sta je posle filtera i strana ostalo.
Public Function SpecOtpremniceIzPrikaza() As Collection
    Dim c As Collection, r As Long, n As Long, oid As String
    Set c = New Collection
    Set SpecOtpremniceIzPrikaza = c
    If Scr_Lista() <> "OTPREMNICE" Then Exit Function
    n = modOtkupUI.GridBrojRedova()
    For r = 1 To n
        oid = Trim$(CStr(modOtkupUI.GridCell(r, UBound(OtpGridCols()) + 1)))
        If Len(oid) > 0 Then c.Add oid
    Next r
End Function

' Jedna ulazna tacka stampe: izdate idu na stampu, nacrti se preskacu i broje
' (nacrt se ne stampa -- review #362, P1). Provere i posao su u modPrint; ovde
' se samo razvrstava i javlja operateru.
Private Sub StampajSpecifikaciju(ByVal ids As Collection)
    Dim izdate As Collection, nacrta As Long, razlog As String
    Set izdate = IzdateZaSpecifikaciju(ids, nacrta)
    If izdate.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_SPEC_NACRT"), True
        Exit Sub
    End If
    razlog = modPrint.PrintSpecifikacijaBlokova(izdate)
    If Len(razlog) > 0 Then
        modOtkupUI.ShowToast razlog, True
        Exit Sub
    End If
    razlog = Poruka("OTKUI_MSG_SPEC") & " " & izdate.count
    If nacrta > 0 Then razlog = razlog & " " & Poruka("OTKUI_MSG_SPEC_NACRT") & " " & nacrta
    modOtkupUI.ShowToast razlog, False
End Sub

' Izdate otpremnice iz izbora, redom izbora i bez ponavljanja; nacrta = koliko
' ih nije izdato. Status se cita po ID-u, istim pravilom kao stampa otpremnice
' (IzdatoStatusJeIzdato). Javno zbog testa.
Public Function IzdateZaSpecifikaciju(ByVal ids As Collection, _
                                      ByRef nacrta As Long) As Collection
    Dim c As Collection, v As Variant, vidjen As Object, k As String
    Set c = New Collection
    Set IzdateZaSpecifikaciju = c
    Set vidjen = CreateObject("Scripting.Dictionary")
    nacrta = 0
    If ids Is Nothing Then Exit Function
    For Each v In ids
        k = UCase$(Trim$(CStr(v)))
        If Len(k) > 0 Then
            If Not vidjen.Exists(k) Then
                vidjen.Add k, True
                If modDokumenta.OtpremnicaJeIzdata(Trim$(CStr(v))) Then
                    c.Add Trim$(CStr(v))
                Else
                    nacrta = nacrta + 1
                End If
            End If
        End If
    Next v
End Function

' Datum kao GRANICA opsega; 0 = nema granice. Prazan ili nepotpun unos nije
' greska -- dok operater kuca "21." nema smisla praznjenje liste.
Private Function DatGranica(ByVal s As String) As Double
    Dim d As Date
    On Error Resume Next
    If Len(Trim$(s)) = 0 Then Exit Function
    If TryParseDateValue(s, d) Then DatGranica = Int(CDbl(d))
End Function

' Klik na red liste F2. ID dolazi iz nevidljive kolone reda, ne iz broja.
' Izdata otpremnica se samo imenuje (nije greska: lista F2 pokazuje i izdate).
Private Sub IzaberiNacrtZaIzmenu(ByVal red As Long)
    Dim otpID As String, razlog As String, spec As String
    otpID = Trim$(CStr(modOtkupUI.GridCell(red, IdentKolonaIndeks("OTPREMNICA"))))
    razlog = OtvoriIzmenuNacrta(otpID, spec)
    If Len(razlog) = 0 Then
        modOtkupUI.ApplyPrefill spec
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_IZMENA_NACRTA") & " " & mIzmenaBroj, False
    ElseIf razlog = Poruka("OTKUI_ERR_OTP_IZDATA") Then
        modOtkupUI.ShowToast razlog, False
    Else
        ' Izmena je zatvorena; forma PRETHODNE izmene ne sme da ostane, jer bi
        ' je sledece snimanje upisalo kao nov nacrt -- duplikat.
        modOtkupUI.ClearForm
        modOtkupUI.ShowToast razlog, True
    End If
End Sub

' Nacrt u formu F2: zaglavlje i ocekivanje po klasi (kolicina, gajbe, predlog
' cene). Dvoklasni unos se ukljucuje pre polja klase II.
' Greska kanonskog citaoca stavki se PROPAGIRA -- pozivalac (OtvoriIzmenuNacrta)
' tada ne otvara izmenu.
Private Function PrefillNacrta(ByVal otpID As String) As String
    Dim res As String, poDok As Object, c As Collection
    Dim red As Variant, i As Long, imaII As Boolean
    res = PrefillZaglavlja(otpID)
    res = Dodaj(res, "brdok", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))

    Set poDok = modDokumenta.StavkeOtpremnicePoDokumentu()
    If poDok.Exists(otpID) Then
        Set c = poDok(otpID)
        For i = 1 To c.count
            red = c(i)
            If UCase$(Trim$(CStr(red(3)))) = KLASA_II Then imaII = True
        Next i
        res = Dodaj(res, "dveklase", IIf(imaII, "2", "1"))
        For i = 1 To c.count
            red = c(i)
            Select Case UCase$(Trim$(CStr(red(3))))
                Case KLASA_I
                    res = Dodaj(res, "kol1", BrojTekst(red(4)))
                    res = Dodaj(res, "amb1", BrojTekst(red(6)))
                    res = Dodaj(res, "cena", BrojTekst(red(5)))
                Case KLASA_II
                    res = Dodaj(res, "kol2", BrojTekst(red(4)))
                    res = Dodaj(res, "amb2", BrojTekst(red(6)))
                    res = Dodaj(res, "cena2", BrojTekst(red(5)))
            End Select
        Next i
    End If
    PrefillNacrta = res
End Function

' Zaglavlje otpremnice kao prefill: datum, roba, otkupno mesto, vozac, tip
' ambalaze. JEDNO mesto koje oba prefilla citaju -- izbor na radnom stolu (F1)
' i izmena nacrta (F2).
Private Function PrefillZaglavlja(ByVal otpID As String) As String
    Dim vDat As Variant, res As String
    On Error Resume Next
    If Len(otpID) = 0 Then Exit Function
    vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    If IsDate(vDat) Then res = "datum=" & Format$(CDate(vDat), "dd.mm.yyyy")
    res = Dodaj(res, "vrsta", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VRSTA)))
    res = Dodaj(res, "sorta", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_SORTA)))
    res = Dodaj(res, "omid", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA)))
    res = Dodaj(res, "vozacid", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VOZAC)))
    res = Dodaj(res, "tipamb", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_TIP_AMB)))
    PrefillZaglavlja = res
End Function

' Broj iz celije u tekst polja forme (lokalni zapis); prazno ostaje prazno.
Private Function BrojTekst(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    BrojTekst = CStr(CDbl(v))
End Function

' Sta se sa otpremnice prepisuje u formu otkupnog lista. Cena je po klasi
' (CenaKlase); broj otkupnog lista se trazi od kanonskog generatora, pa kad je
' auto-broj iskljucen polje ostaje operateru.
Private Function PrefillSpec(ByVal otpID As String) As String
    Dim vDat As Variant, stanica As String, res As String, cena As Double
    On Error Resume Next
    If Len(otpID) = 0 Then Exit Function

    vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    stanica = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))
    res = PrefillZaglavlja(otpID)
    cena = CenaKlase(otpID, KLASA_I)
    If cena > 0 Then res = Dodaj(res, "cena", Format$(cena, "0.00"))
    cena = CenaKlase(otpID, KLASA_II)
    If cena > 0 Then res = Dodaj(res, "cena2", Format$(cena, "0.00"))
    If Len(stanica) > 0 And IsDate(vDat) Then
        res = Dodaj(res, "brdok", SuggestNextBroj(KIND_OTK, stanica, CDate(vDat), False))
    End If
    PrefillSpec = res
End Function

Private Function Dodaj(ByVal res As String, ByVal k As String, ByVal v As String) As String
    Dodaj = res
    If Len(v) = 0 Then Exit Function
    Dodaj = res & IIf(Len(res) > 0, "|", "") & k & "=" & v
End Function

' Pita operatera kad blok prelazi ostatak aktivne otpremnice u nekoj klasi.
' Kolicine su vec NETO (OtkupValidiraj je oduzeo taru).
'
' Kad read-model padne, pitanje se preskace: upis otkupa ne zavisi od njega, a
' isti kvar ce sledeci korak (vezivanje) prijaviti glasno i imenom.
Private Function PotvrdiPrekoracenje(ByVal kgI As Double, ByVal kgII As Double) As Boolean
    Dim opis As String
    PotvrdiPrekoracenje = True
    On Error GoTo EH
    opis = PrekoracenjeOpis(kgI, kgII)
    If Len(opis) = 0 Then Exit Function
    PotvrdiPrekoracenje = (MsgBox(Poruka("OTKUI_ASK_PREKORACENJE_1") & " " & opis & " " & _
                                  Poruka("OTKUI_ASK_PREKORACENJE_2"), _
                                  vbQuestion + vbYesNo, APP_NAME) = vbYes)
    Exit Function
EH:
    PotvrdiPrekoracenje = True
End Function

'---------------------------------------------- LISTA: KOOPERANTI (F1)
' Rang kooperanata po iznosu otkupnih listova u tekucoj godini - legacy "Lista
' kooperanata". Racun je u modOtkupBlok.KoopRangRows, isti koji puni i legacy
' panel; ovde se samo prevodi u redove mreze.
'
' Za razliku od legacy panela (listbox preko pola forme, bez sortiranja i
' pretrage), ovo je obicna lista ljuske - pa ima pretragu, sortiranje po bilo
' kojoj koloni i strane.
Private Function KoopGridCols() As Variant
    KoopGridCols = Array( _
        "OTKUI_HDK_RANG||num|54|1", _
        "OTKUI_HDK_KOOPERANT||part|0|1", _
        "OTKUI_HD_OM||txt|170|2", _
        "OTKUI_HDK_IZNOS||rsd|130|1")
End Function

Private Function RowsKooperanti(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim hay As String, sumVal As Double, iznos As Double
    Dim rawKg As Double, rawVal As Double, emptyKg As Double, emptyVal As Double
    On Error GoTo EH
    mStep = "kooperanti"

    src = modOtkupBlok.KoopRangRows(rawKg, rawVal, emptyKg, emptyVal)
    If Not IsArray(src) Then
        RowsKooperanti = Array(KoopGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 4)
    For r = 1 To UBound(src, 1)
        hay = CStr(src(r, 2)) & "|" & CStr(src(r, 3))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        iznos = CDbl(src(r, 4))
        n = n + 1
        ' rang je mesto na CELOJ listi, ne redni broj posle pretrage - inace bi
        ' pretraga "prepakovala" rang i broj bi lagao
        outA(n, 1) = r
        outA(n, 2) = CStr(src(r, 2))
        outA(n, 3) = CStr(src(r, 3))
        outA(n, 4) = iznos
        sumVal = sumVal + iznos
Sledeci:
    Next r

    mStep = "OK"
    RowsKooperanti = Array(KoopGridCols(), outA, n, 0#, sumVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsKooperanti[" & mStep & "]", Err.description
End Function

'--------------------------------------------------------------- UPIS
' Ljuska predaje vrednosti forme pod logickim imenima; ovde se odlucuje sta su
' i sta se sa njima radi. Vraca "" kad je proslo, poruku kad nije, jedan razmak
' kad je operater odustao. U isti recnik se upisuju "fokus" (polje na koje
' treba vratiti kursor), "rezultat" (broj/ID upisanog) i "poruke" (napomene).
'
' NIJEDAN racun nije ovde: provere, bruto->neto, upis, stampa i auto-lanac
' hladnjace rade modOtkupUnos i rutine koje on zove - iste one koje zove i
' legacy frmOtkup.
Public Function Scr_Save(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    On Error GoTo EH
    polja("fokus") = ""
    polja("rezultat") = ""
    polja("poruke") = ""

    Select Case CStr(polja("rezim"))
        Case "OTKUP"        ' nastavlja se ispod
        Case "OTPREMNICA"
            Scr_Save = SnimiOtpremnicu(polja)
            Exit Function
        Case "ZBIRNA"
            Scr_Save = SaveZbirna(polja)
            Exit Function
        Case "PRIJEMNICA"
            Scr_Save = SavePrijemnica(polja)
            Exit Function
        Case "AMB_ISPLATE"
            Scr_Save = SaveIsplata(polja)
            Exit Function
        Case "AMB_UPLATE"
            Scr_Save = SaveUplata(polja)
            Exit Function
        Case "REVERSI"
            Scr_Save = SaveRevers(polja)
            Exit Function
        Case Else
            Scr_Save = Poruka("OTKUI_TODO_NEVEZANO")
            Exit Function
    End Select

    Set p = modOtkupUnos.NoviOtkupUnos()
    p("datum") = polja("datum")
    p("stanicaID") = polja("stanicaID")
    ' Kooperant moze biti izabran iz liste (ima ID) ili otkucan rukom. U drugom
    ' slucaju se trazi po imenu, pa i kreira ako je auto-kreiranje ukljuceno -
    ' isto sto legacy radi kroz ResolveKooperantByName.
    Dim koopID As String, koopNov As Boolean
    koopID = CStr(polja("kooperantID"))
    If Len(koopID) = 0 Then
        koopID = ResolveKooperantByText(CStr(polja("partnerTekst")), _
                                        CStr(polja("stanicaID")), koopNov)
    End If
    p("kooperantID") = koopID
    p("vrsta") = polja("vrsta")
    p("sorta") = polja("sorta")
    p("tipAmb") = polja("tipAmb")
    p("vozacID") = polja("vozacID")
    p("brDok") = polja("brDok")
    p("brojZbirne") = polja("brojZbirne")
    p("parcelaID") = polja("parcelaID")
    p("kolicinaI") = polja("kolicinaI")
    p("cenaI") = polja("cenaI")
    p("kolAmb") = polja("kolAmb")
    p("kolAmbIzdata") = polja("kolAmbIzdata")
    p("dveKlase") = polja("dveKlase")
    p("kolicinaII") = polja("kolicinaII")
    p("cenaII") = polja("cenaII")
    p("kolAmbII") = polja("kolAmbII")
    ' Kes se uz otkupni list vise ne knjizi - ide kroz F5/F6.
    p("novac") = 0#
    p("primalac") = ""

    greska = modOtkupUnos.OtkupValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        Scr_Save = greska
        Exit Function
    End If

    ' ROUTING PRE SVIH PRAVILA RUCNOG TOKA.
    '
    ' Prekoracenje se meri nad AKTIVNIM RUCNIM NACRTOM, a hladnjacki blok tom
    ' nacrtu ne pripada -- ide u svoj lanac. Da se pitanje postavilo pre nego sto
    ' se to zna, operater bi na "Ne" izgubio ceo upis zbog ogranicenja dokumenta
    ' sa kojim blok nema veze.
    '
    ' Pita se SAMO kad se pouzdano zna da blok ide rucnim tokom: i "ide u lanac"
    ' i "ne moze da se utvrdi" preskacu pitanje (fail-closed).
    ' Otvorena ispravka: snimanje pravi ZAMENU ovog bloka, ne nov dokument.
    ' Zna se PRE pitanja o prekoracenju, jer ispravka tom pitanju ne podleze.
    Dim ispravkaID As String, imaoRoditelja As Boolean
    ispravkaID = mIspravkaOtkupID
    If Len(ispravkaID) > 0 Then _
        imaoRoditelja = (Len(modDokumenta.OtpremnicaZaOtkup(ispravkaID)) > 0)

    Dim putGreska As String, ideULanac As Boolean
    ideULanac = modAutoHladnjaca.LanacVaziZaStanicu(CStr(polja("stanicaID")), putGreska)

    ' PREKORACENJE MERI DODAVANJE, A ISPRAVKA NE DODAJE NEGO ZAMENJUJE.
    '
    ' Nacrt koji ocekuje 60 i vec ima izvor od 60 posle ispravke na 55 ima
    ' povezano 55, ne 115. Da se pitanje postavilo, operater bi potvrdjivao
    ' prekoracenje koje ne postoji -- a na "Ne" bi izgubio legitimnu ispravku.
    ' Zato ga preskace i ispravka, kao i hladnjacki blok i neizvestan put.
    If Len(ispravkaID) = 0 And Not ideULanac And Len(putGreska) = 0 Then
        If Not PotvrdiPrekoracenje(CDbl(p("kolicinaI")), CDbl(p("kolicinaII"))) Then
            Scr_Save = " "
            Exit Function
        End If
    End If

    res = modOtkupUnos.OtkupUpisi(p, poruke, ispravkaID)
    If Len(res) = 0 Then
        Scr_Save = Poruka("OTKUP_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    ' Ispravka je uspela -- sledece snimanje je opet nov dokument. Gasi se TEK
    ' posle uspeha: na gresku operater popravi polja i snimi ponovo.
    If Len(ispravkaID) > 0 Then Scr_IspravkaOtkazi

    ' Sta se radi sa blokom posle upisa -- odluka je u RutaPosleUpisa, ovde samo
    ' izvrsenje. Lanac ne obara upis: otkup je snimljen svojom transakcijom i
    ' ostaje i kad otpremnica ne uspe, a razlog ide operateru u istu poruku.
    Dim lanacPoruka As String, ruta As String, veza As String
    ruta = RutaPosleUpisa(res, Len(ispravkaID) > 0, imaoRoditelja, putGreska)
    If Len(putGreska) > 0 Then
        ' Ne zna se kojim putem blok ide -- ne gura se NI U LANAC NI U NACRT.
        ' Vezivanje "jer provera nije uspela" bi hladnjacki blok tiho smestilo u
        ' tudji dokument. Blok je upisan i ceka u listi "Bez otpremnice".
        poruke = Trim$(poruke & "  " & putGreska)
    ElseIf ruta = RUTA_LANAC Then
        modAutoHladnjaca.AutoLanacHladnjaca res, lanacPoruka
        If Len(lanacPoruka) > 0 Then poruke = Trim$(poruke & "  " & lanacPoruka)
    ElseIf ruta = RUTA_NACRT Then
        veza = VeziZaAktivnu(res)
        If Len(veza) > 0 Then poruke = Trim$(poruke & "  " & veza)
    End If

    ' Nov kooperant je kreiran tokom upisa - lista partnera mora da ga vidi
    ' odmah, bez zatvaranja ekrana.
    If koopNov Then modOtkupUI.RefreshPartnerLista

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
    Exit Function
EH:
    Scr_Save = Poruka("OTKUI_ERR_RADNJA") & " " & Err.description
End Function

' F2 OTPREMNICA. Ekran samo prevodi polja u recnik i vraca poruku - posao radi
' modDokUnos. Ovde nema nijedne provere: sve sto je provera zivi u modulu.
'
' Od S3a upis otvara NACRT, pa modDokUnos vraca OtpremnicaID. Operateru se u
' toast-u i dalje pokazuje BROJ -- identitet je za masinu, broj za coveka. ID
' ostaje u recniku pod svojim imenom, za radnju "Izdaj" (S3b).
Private Function SnimiOtpremnicu(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modDokUnos.NoviOtpremnicaUnos()
    p("datum") = polja("datum")
    p("stanicaID") = polja("stanicaID")
    p("vozacID") = polja("vozacID")
    p("brDok") = polja("brDok")
    p("brojZbirne") = polja("brojZbirne")
    p("vrsta") = polja("vrsta")
    p("sorta") = polja("sorta")
    p("tipAmb") = polja("tipAmb")
    p("kolicinaI") = polja("kolicinaI")
    p("cenaI") = polja("cenaI")
    p("kolAmb") = polja("kolAmb")
    p("dveKlase") = polja("dveKlase")
    p("kolicinaII") = polja("kolicinaII")
    p("cenaII") = polja("cenaII")
    p("kolAmbII") = polja("kolAmbII")
    ' Otvorena izmena: validacija broja izuzima sopstveni red nacrta.
    p("izmenaOtpID") = mIzmenaOtpID

    greska = modDokUnos.OtpremnicaValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SnimiOtpremnicu = greska
        Exit Function
    End If

    ' Otvorena izmena nacrta: snimanje MENJA taj nacrt, ne pravi nov.
    If Len(mIzmenaOtpID) > 0 Then
        If Not modDokUnos.OtpremnicaIzmeniNacrt(mIzmenaOtpID, p, poruke) Then
            SnimiOtpremnicu = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
            Exit Function
        End If
        Scr_ResetCache
        polja("otpremnicaID") = mIzmenaOtpID
        polja("rezultat") = CStr(polja("brDok"))
        polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
        Scr_IzmenaOtkazi
        Exit Function
    End If

    res = modDokUnos.OtpremnicaUpisi(p, poruke)
    If Len(res) = 0 Then
        SnimiOtpremnicu = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("otpremnicaID") = res
    polja("rezultat") = CStr(polja("brDok"))
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' F3 ZBIRNA. Isti obrazac kao SnimiOtpremnicu: ekran samo prevodi polja u recnik.
' Dve razlike koje dolaze iz same forme, ne iz odluke ovog modula:
'   - BROJ DOKUMENTA JE BROJ ZBIRNE (u F3 polje "broj zbirne" i ne postoji -
'     modOtkupUI.ModeVezujeZbirnu je False za taj rezim), pa ide kao "brDok";
'   - PARTNER je kupac. Ljuska ga skuplja pod kljucem "kooperantID" jer je to
'     ista kontrola (cbKupac) u svim rezimima; ovde dobija svoje ime.
Private Function SaveZbirna(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modDokUnos.NoviZbirnaUnos()
    p("datum") = polja("datum")
    p("vozacID") = polja("vozacID")
    p("kupacID") = polja("kooperantID")
    p("brDok") = polja("brDok")
    ' ODREDISTE (MIG-001): hladnjaca i pogon su kolone tblZbirna koje writer vec
    ' pise; do v6-ui-215 ih ekran nije slao, pa su isle prazne. Ekran ih SAMO
    ' prevodi - pravilo (odakle hladnjaca dolazi) zivi u ljusci, provera u
    ' modDokUnos, upis u modDokumenta.
    p("hladnjaca") = polja("hladnjaca")
    p("pogon") = polja("pogon")
    p("vrsta") = polja("vrsta")
    p("sorta") = polja("sorta")
    p("tipAmb") = polja("tipAmb")
    p("kolicinaI") = polja("kolicinaI")
    p("kolAmb") = polja("kolAmb")
    p("dveKlase") = polja("dveKlase")
    p("kolicinaII") = polja("kolicinaII")
    p("kolAmbII") = polja("kolAmbII")

    greska = modDokUnos.ZbirnaValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SaveZbirna = greska
        Exit Function
    End If

    res = modDokUnos.ZbirnaUpisi(p, poruke)
    If Len(res) = 0 Then
        SaveZbirna = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' F4 PRIJEMNICA. Kao gore; ovde broj dokumenta jeste broj prijemnice, a broj
' zbirne dolazi iz svog polja. "kolAmbIzdata" je u ovom rezimu VRACENA ambalaza
' (isto polje forme, drugo znacenje - modOtkupUI.ApplyFormFields).
Private Function SavePrijemnica(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modDokUnos.NoviPrijemnicaUnos()
    p("datum") = polja("datum")
    p("kupacID") = polja("kooperantID")
    p("vozacID") = polja("vozacID")
    p("brDok") = polja("brDok")
    p("brojZbirne") = polja("brojZbirne")
    p("vrsta") = polja("vrsta")
    p("sorta") = polja("sorta")
    p("tipAmb") = polja("tipAmb")
    p("kolicinaI") = polja("kolicinaI")
    p("cenaI") = polja("cenaI")
    p("kolAmb") = polja("kolAmb")
    p("kolAmbVracena") = polja("kolAmbIzdata")
    p("dveKlase") = polja("dveKlase")
    p("kolicinaII") = polja("kolicinaII")
    p("cenaII") = polja("cenaII")
    p("kolAmbII") = polja("kolAmbII")

    greska = modDokUnos.PrijemnicaValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SavePrijemnica = greska
        Exit Function
    End If

    res = modDokUnos.PrijemnicaUpisi(p, poruke)
    If Len(res) = 0 Then
        SavePrijemnica = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' F5 ISPLATE. Gotovinski rezimi nemaju ni robu ni ambalazu, pa je recnik kratak.
' Partner se u ljusci zove "kooperantID" (ista kontrola u svim rezimima); ovde
' dobija svoje ime i, uz njega, TIP partnera - od koga zavisi tip novca.
Private Function SaveIsplata(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modNovacUnos.NoviIsplataUnos()
    p("datum") = polja("datum")
    p("stanicaID") = polja("stanicaID")
    p("stanicaTekst") = polja("stanicaTekst")
    p("partnerID") = polja("kooperantID")
    p("partnerTip") = polja("partnerTip")
    p("partnerTekst") = polja("partnerTekst")
    p("vrsta") = polja("vrsta")
    p("brDok") = polja("brDok")
    p("novac") = polja("novac")
    p("otkupID") = polja("otkupID")
    p("blokTekst") = polja("blokTekst")
    p("otkupOstatak") = polja("otkupOstatak")
    p("izAvansa") = polja("izAvansa")

    greska = modNovacUnos.IsplataValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SaveIsplata = greska
        Exit Function
    End If

    res = modNovacUnos.IsplataUpisi(p, poruke)
    If Len(res) = 0 Then
        SaveIsplata = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' F6 UPLATE KUPACA. Partner je kupac, a izabrana faktura odlucuje da li je red
' uplata po fakturi ili avans - zato i njen preostali iznos ide u recnik.
Private Function SaveUplata(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modNovacUnos.NoviUplataUnos()
    p("datum") = polja("datum")
    p("partnerID") = polja("kooperantID")
    p("partnerTekst") = polja("partnerTekst")
    p("vrsta") = polja("vrsta")
    p("brDok") = polja("brDok")
    p("novac") = polja("novac")
    p("fakturaID") = polja("fakturaID")
    p("fakturaTekst") = polja("fakturaTekst")
    p("fakturaOstatak") = polja("fakturaOstatak")

    greska = modNovacUnos.UplataValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SaveUplata = greska
        Exit Function
    End If

    res = modNovacUnos.UplataUpisi(p, poruke)
    If Len(res) = 0 Then
        SaveUplata = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' F7 REVERSI. Jedini rezim u kome ambalaza ide bez robe; smer je redni broj
' izabranog segmenta, a modNovacUnos ga prevodi u ono sto core ocekuje.
' Broj reversa se predlaze u ljusci, a ako je ostao prazan generise ga
' ReversValidiraj - pa se posle upisa cita IZ RECNIKA, ne iz polja forme.
Private Function SaveRevers(ByVal polja As Object) As String
    Dim p As Object, fokus As String, greska As String, res As String, poruke As String
    Set p = modNovacUnos.NoviReversUnos()
    p("datum") = polja("datum")
    p("stanicaID") = polja("stanicaID")
    p("stanicaTekst") = polja("stanicaTekst")
    p("partnerID") = polja("kooperantID")
    p("partnerTip") = polja("partnerTip")
    p("partnerTekst") = polja("partnerTekst")
    p("vozacID") = polja("vozacID")
    p("vrsta") = polja("vrsta")
    p("brDok") = polja("brDok")
    p("tipAmb") = polja("tipAmb")
    p("kolAmb") = polja("kolAmb")
    p("smerRev") = polja("smerRev")

    greska = modNovacUnos.ReversValidiraj(p, fokus)
    If Len(greska) > 0 Then
        polja("fokus") = fokus
        SaveRevers = greska
        Exit Function
    End If

    res = modNovacUnos.ReversUpisi(p, poruke)
    If Len(res) = 0 Then
        SaveRevers = Poruka("DOK_MSG_GRESKA_PRI_CUVANJU") & " " & poruke
        Exit Function
    End If

    Scr_ResetCache
    polja("rezultat") = res
    polja("poruke") = Replace(Trim$(poruke), vbCrLf, "  ")
End Function

' Ikonica u markeru uz naslov - po DOKUMENTU, ne po modulu. Sve kodne tacke su
' vec proverene i koriste se drugde u ovom modulu; nijedna nije pogodjena "po
' opisu". Spisak svih glifova sa kodovima: Alt+F8 -> DumpMdl2Sheet.
Public Function ModeIco(ByVal mode As String) As Long
    Select Case mode
        Case "F1": ModeIco = IC_OTKUP       ' QuickNote  - otkupni list
        Case "F2": ModeIco = IC_BLOKOVI     ' Document   - otpremnica
        Case "F3": ModeIco = IC_NALOZI      ' CheckList  - zbirna je spisak
        Case "F4": ModeIco = IC_IZVEST      ' ReportDocument - prijemnica
        Case "F5": ModeIco = IC_ISPLATA     ' Upload     - novac izlazi
        Case "F6": ModeIco = IC_UPLATA      ' Download   - novac ulazi
        Case "F7": ModeIco = IC_REVERS      ' privremeno - ceka izbor
        Case Else: ModeIco = IC_OTKUP
    End Select
End Function

' dupli klik na red -> ucitaj dokument u polja iznad
' Rezimi koji NOSE vezu na zbirnu (imaju polje BROJ ZBIRNE).
Public Function ModeVezujeZbirnu(ByVal mode As String) As Boolean
    Select Case mode
        Case "F1", "F2", "F4": ModeVezujeZbirnu = True
    End Select
End Function

' Izvor KANONSKOG IDENTITETA po tipu dokumenta.
'
' Broj je labela. Za prijemnicu to nije teorija: GenerateBrojPrijemnice ima
' fiksan prefiks "1", broji sekvencu PO KUPCU i NEMA proveru jedinstvenosti, a
' auto-broj postoji samo za hladnjacu -- ostali kupci unose slobodno.
' Broj zbirne generator drzi jedinstvenim; tamo je identitet pojas za rucni
' unos. Za robna dokumenta identitet je GeneracijaID --
' Klasa I i II iz istog upisa dele vrednost, sto je tacno "jedan logicki
' dokument". Novac i faktura su jednoredni, pa im je identitet sopstveni PK.
'
' Revers je dve noge (Kooperant + Stanica) istog broja, a broj je jedinstven tek
' u nizu (stanica, dan) -- ni broj ni broj + smer ne razlucuju dokument. Identitet
' reda je AmbID kliknute noge; ReversID dokumenta se iz nje cita nizvodno
' (modStorno.ReversIDRazresi). Izvod nije ovde: ide uz BROJ RACUNA, kompozit
' koji vec razlucuje dokument.
Public Function IdKolonaTipa(ByVal tk As String) As String
    Select Case tk
        ' Otkup je JEDNO zaglavlje po dokumentu (S1e): identitet je OtkupID.
        ' GeneracijaID otkupni pisac ne upisuje, pa bi kolona bila prazna.
        Case "OTKUP":                                     IdKolonaTipa = COL_OTK_ID
        ' Isto za otpremnicu od S3a (review #362): nacrt je jedno zaglavlje sa
        ' OtpremnicaID-em, a GeneracijaID ne dobija. Sa generacijom bi skrivena
        ' kolona bila prazna, pa bi F8 dokument trazio PO BROJU -- a broj je
        ' jedinstven tek po (stanica, dan).
        Case "OTPREMNICA":                                IdKolonaTipa = COL_OTP_ID
        Case "ZBIRNA", "PRIJEMNICA":                       IdKolonaTipa = COL_GENERACIJA_ID
        Case "FAKTURA":                                     IdKolonaTipa = COL_FAK_ID
        Case "AMB_ISPLATE", "AMB_UPLATE":                   IdKolonaTipa = COL_NOV_ID
        Case "REVERSI":                                     IdKolonaTipa = COL_AMB_ID
        Case Else:                                          IdKolonaTipa = ""
    End Select
End Function

' Indeks nevidljive kolone identiteta u mrezi tipa: uvek POSLEDNJA koju
' GridCols(tk, True) doda. Racuna se iz istog niza koji mreza dobija, pa ne moze
' da se razidje sa njim. 0 = tip identitet nema.
Public Function IdentKolonaIndeks(ByVal tk As String) As Long
    If Len(IdKolonaTipa(tk)) = 0 Then Exit Function
    Dim cols As Variant: cols = GridCols(tk, True)
    If Not IsArray(cols) Then Exit Function
    IdentKolonaIndeks = UBound(cols) + 1
End Function

Public Function GridCols(ByVal mk As String, Optional ByVal saIdentitetom As Boolean = False) As Variant
    Dim c As Collection: Set c = New Collection
    ' Izvod nije red tabele nego GRUPA redova (jedan izvod = mnogo stavki),
    ' pa mu kolone daje sopstvena rutina, kao listi otpremnica u F1.
    If mk = "IZVOD" Then
        GridCols = IzvGridCols()
        Exit Function
    End If
    c.Add "OTKUI_HD_BROJ|" & ColBroj(mk) & "|txt|110|1"
    c.Add "OTKUI_HD_DATUM|" & ColDatum(mk) & "|date|58|1"
    c.Add "OTKUI_HD_PARTNER|" & ColPartner(mk) & "|part|0|1"

    Select Case mk
        Case "OTKUP", "OTPREMNICA", "ZBIRNA", "PRIJEMNICA"
            c.Add "OTKUI_HD_VRSTA|" & ColVrsta(mk) & "|txt|72|2"
            ' 104 pt = najduza realna sorta ("Willamette teren") na TS_BODY;
            ' visak uzima fleksibilna kolona PARTNER, ostale se ne pomeraju
            c.Add "OTKUI_HD_SORTA|" & ColSorta(mk) & "|txt|104|3"
            c.Add "OTKUI_HD_KLASA|" & ColKlasa(mk) & "|txt|46|2"
            c.Add "OTKUI_HD_KG|" & ColKolicina(mk) & "|kg|60|1"
            c.Add "OTKUI_HD_KOL_AMB|" & ColKolAmb(mk) & "|num|54|3"
            c.Add "OTKUI_HD_TIP_AMB|" & ColTipAmb(mk) & "|txt|78|3"
            ' tblZbirna nema Cenu - taj rezim ostaje bez kolone vrednosti
            If Len(ColCena(mk)) > 0 Then
                c.Add "OTKUI_HD_VREDNOST|" & ColCena(mk) & "|mult|92|1"
            End If
            ' Placanje se NE cita iz zastavice: tblOtkup.Isplaceno je samo "Da"
            ' ili prazno, pa ne razlikuje delimicno od nista. Pravo stanje se
            ' racuna iz tblNovac (modNovac.BuildIsplataDictByOtkup) i poredi sa
            ' vrednoscu dokumenta - odatle tri stanja i ostatak duga.
            If mk = "OTKUP" Or mk = "PRIJEMNICA" Then
                c.Add "OTKUI_HD_PLACENO||paypill|86|1"
                c.Add "OTKUI_HD_OSTATAK||rest|84|2"
            End If
        Case "AMB_ISPLATE"
            c.Add "OTKUI_HD_KANAL||kanal|82|1"
            c.Add "OTKUI_HD_VREDNOST|" & COL_NOV_ISPLATA & "|rsd|110|1"
        Case "AMB_UPLATE"
            c.Add "OTKUI_HD_KANAL||kanal|82|1"
            c.Add "OTKUI_HD_VREDNOST|" & COL_NOV_UPLATA & "|rsd|110|1"
        Case "REVERSI"
            ' OSNOV nosi najduzi tekst u mrezi ("Revers " & em-dash & " OM prijem",
            ' 18 znakova) - 112pt ga je seklo. 150pt prima i najduzu varijantu sa
            ' rezervom, a mesta ima: ovaj rezim ima samo 7 kolona.
            c.Add "OTKUI_HD_SMER|" & COL_AMB_SMER & "|txt|62|1"
            c.Add "OTKUI_HD_OSNOV||osnov|150|1"
            c.Add "OTKUI_HD_TIP_AMB|" & COL_AMB_TIP & "|txt|96|2"
            c.Add "OTKUI_HD_KOMADA|" & COL_AMB_KOLICINA & "|sum0|80|1"
        Case "FAKTURA"
            ' Faktura nema ni robu ni ambalazu - nosi iznos i svoj status.
            c.Add "OTKUI_HD_IZNOS|" & COL_FAK_IZNOS & "|rsd|110|1"
    End Select

    c.Add "OTKUI_HD_STATUS||pill|88|1"

    ' KANONSKI IDENTITET izabranog reda -- samo za Storno, i NEVIDLJIVO.
    '
    ' Zasto kolona a ne mapa sa strane: ljuska sortira redove POSLE Scr_Rows
    ' (SortedView), pa izlazni indeks nije indeks u mrezi -- mapa "red -> ID"
    ' bi posle prvog klika na zaglavlje pokazivala na pogresan dokument.
    ' SortedView kopira tacno mColN kolona, pa ono sto nije u ovom nizu ne
    ' prezivi sortiranje.
    '
    ' Prioritet 4 znaci NIKAD vidljiva: petlja vidljivosti ide "For pass = 3
    ' To 1", pa uslov (4 <= pass) nikad nije tacan. Podatak svejedno putuje,
    ' jer mColN broji sve deklarisane kolone. Sirina 0 je bez znacaja dok je
    ' kolona nevidljiva, ali stoji da flex-raspodela ne bi imala sta da uzme.
    '
    ' Kapija je ARGUMENT, ne ActiveMode. Do v6-ui-143 je ovde stajalo
    ' "If modOtkupUI.ActiveMode = "F8"", sto je radilo dok je storno bio rezim
    ' te iste ljuske. Ekran nema rezim: ostavljena, ta provera bi cutke bila
    ' False, kolona bi nestala, IdentIzReda bi vracao prazno i ceo lanac
    ' identiteta iz #198 bi pao na fail-closed po broju -- a nijedna suite to
    ' ne bi videla, jer testovi identiteta mere sloj ISPOD mreze.
    '
    ' Kapija i dalje postoji (a ne "uvek dodaj"): GridCols je zajednicki za
    ' rezim unosa i za Storno nad istim tipom (F4 i Storno/Prijemnica daju
    ' isti mk), pa bi bezuslovno dodavanje menjalo i liste unosnih rezima.
    If saIdentitetom Then
        If Len(IdKolonaTipa(mk)) > 0 Then
            c.Add "OTKUI_HD_IDENT|" & IdKolonaTipa(mk) & "|txt|0|4"
        End If
    End If

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    GridCols = a
End Function

Public Function ColBroj(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                   ColBroj = COL_OTK_BR_DOK
        Case "OTPREMNICA":              ColBroj = COL_OTP_BROJ
        Case "ZBIRNA":                  ColBroj = COL_ZBR_BROJ
        Case "PRIJEMNICA":              ColBroj = COL_PRJ_BROJ
        Case "AMB_ISPLATE", "AMB_UPLATE": ColBroj = COL_NOV_BROJ_DOK
        Case "REVERSI":                 ColBroj = COL_AMB_DOK_ID
        Case "FAKTURA":                 ColBroj = COL_FAK_BROJ
    End Select
End Function

Public Function ColDatum(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                   ColDatum = COL_OTK_DATUM
        Case "OTPREMNICA":              ColDatum = COL_OTP_DATUM
        Case "ZBIRNA":                  ColDatum = COL_ZBR_DATUM
        Case "PRIJEMNICA":              ColDatum = COL_PRJ_DATUM
        Case "AMB_ISPLATE", "AMB_UPLATE": ColDatum = COL_NOV_DATUM
        Case "REVERSI":                 ColDatum = COL_AMB_DATUM
        Case "FAKTURA":                 ColDatum = COL_FAK_DATUM
    End Select
End Function

Public Function ColPartner(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                   ColPartner = COL_OTK_KOOPERANT
        Case "OTPREMNICA":              ColPartner = COL_OTP_STANICA
        Case "ZBIRNA":                  ColPartner = COL_ZBR_KUPAC
        Case "PRIJEMNICA":              ColPartner = COL_PRJ_KUPAC
        Case "AMB_ISPLATE", "AMB_UPLATE": ColPartner = COL_NOV_PARTNER
        Case "REVERSI":                 ColPartner = COL_AMB_ENTITET
        Case "FAKTURA":                 ColPartner = COL_FAK_KUPAC
    End Select
End Function

Public Function ColVrsta(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColVrsta = COL_OTK_VRSTA
        Case "OTPREMNICA":           ColVrsta = COL_OTP_VRSTA
        Case "ZBIRNA":               ColVrsta = COL_ZBR_VRSTA
        Case "PRIJEMNICA":           ColVrsta = COL_PRJ_VRSTA
    End Select
End Function

Public Function ColSorta(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColSorta = COL_OTK_SORTA
        Case "OTPREMNICA":           ColSorta = COL_OTP_SORTA
        Case "ZBIRNA":               ColSorta = COL_ZBR_SORTA
        Case "PRIJEMNICA":           ColSorta = COL_PRJ_SORTA
    End Select
End Function

Public Function ColKlasa(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKlasa = COL_OKS_KLASA        ' stavka (ovStav)
        Case "OTPREMNICA":           ColKlasa = COL_OPS_KLASA        ' stavka (ovStav)
        Case "ZBIRNA":               ColKlasa = COL_ZBR_KLASA
        Case "PRIJEMNICA":           ColKlasa = COL_PRJ_KLASA
    End Select
End Function

Public Function ColKolicina(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKolicina = COL_OKS_KOLICINA  ' stavka (ovStav)
        Case "OTPREMNICA":           ColKolicina = COL_OPS_KOLICINA  ' stavka (ovStav)
        Case "ZBIRNA":               ColKolicina = COL_ZBR_KOLICINA
        Case "PRIJEMNICA":           ColKolicina = COL_PRJ_KOLICINA
    End Select
End Function

Public Function ColKolAmb(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKolAmb = COL_OKS_KOL_AMB     ' stavka (ovStav)
        Case "OTPREMNICA":           ColKolAmb = COL_OPS_KOL_AMB     ' stavka (ovStav)
        Case "ZBIRNA":               ColKolAmb = COL_ZBR_KOL_AMB
        Case "PRIJEMNICA":           ColKolAmb = COL_PRJ_KOL_AMB
    End Select
End Function

Public Function ColTipAmb(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColTipAmb = COL_OTK_TIP_AMB
        Case "OTPREMNICA":           ColTipAmb = COL_OTP_TIP_AMB
        Case "ZBIRNA":               ColTipAmb = COL_ZBR_TIP_AMB
        Case "PRIJEMNICA":           ColTipAmb = COL_PRJ_TIP_AMB
    End Select
End Function

' Prazno = rezim nema cenu, pa ni kolonu vrednosti.
'
' OTPREMNICA je nema (review #362, P1). Njena PredlogCena je predlog za prefill
' otkupa, izricito NE-finansijsko polje, pa Kolicina x PredlogCena nije vrednost
' dokumenta. Prava vrednost je vrednost izvornih otkupa, a nju nacrt -- koga
' ova mreza uglavnom prikazuje -- jos nema. Kolona bi zato bila ili izmisljena
' cifra, ili nula; oba lazu, pa je nema.
Public Function ColCena(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColCena = COL_OKS_CENA          ' stavka (ovStav)
        Case "PRIJEMNICA":           ColCena = COL_PRJ_CENA
    End Select
End Function

Public Function ColBrojZbirne(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColBrojZbirne = COL_OTK_BROJ_ZBIRNE
        Case "OTPREMNICA":           ColBrojZbirne = COL_OTP_BROJ_ZBIRNE
        Case "PRIJEMNICA":           ColBrojZbirne = COL_PRJ_BROJ_ZBIRNE
    End Select
End Function

' Polje opisa kolone: 0=kljuc naslova 1=izvorna kolona 2=vrsta 3=sirina 4=prio
Public Function ColF(ByVal spec As String, ByVal idx As Long) As String
    Dim p As Variant: p = Split(spec, "|")
    If idx > UBound(p) Then Exit Function
    ColF = CStr(p(idx))
End Function

Public Function StatusCode(ByVal isStorno As Boolean, ByVal bezZbirne As Boolean) As Long
    If isStorno Then
        StatusCode = 2
    ElseIf bezZbirne Then
        StatusCode = 0
    Else
        StatusCode = 1
    End If
End Function

' SEDAM dokumenata, F1..F7. Ranija sema (F2..F6+F8) je spajala dva razlicita
' dokumenta u F5: kartica je pisala "Ulaz OM" a mreza je citala tblOtkup.
' Sada je "Otkupni list" (tblOtkup) zaseban rezim F1, a F5/F6 su gotovinski
' promet iz tblNovac (isplate kooperantu / uplate od kupca) - isti smer kao
' frmDokumenta frame-ovi "Ulaz OM (Novac kooperantu)" i "Izlaz Kupci
' (Novac od kupca)".
' Tabela rezima. Ista tabela do koje se stize preko kljuca tipa -- rezim je
' samo drugo ime za tip koji se u njemu unosi, pa ovde nema drugog spiska.
Public Function ModeTable(ByVal mode As String) As String
    ModeTable = TabelaTipa(modeKey(mode))
End Function

' Tabela po kljucu tipa. Sedam tipova su rezimi F1..F7 i vec imaju svoju
' tabelu; fakture i izvodi je nemaju kroz ModeTable jer nemaju rezim.
Public Function TabelaTipa(ByVal tk As String) As String
    Select Case tk
        Case "OTKUP":       TabelaTipa = TBL_OTKUP
        Case "OTPREMNICA":  TabelaTipa = TBL_OTPREMNICA
        Case "ZBIRNA":      TabelaTipa = TBL_ZBIRNA
        Case "PRIJEMNICA":  TabelaTipa = TBL_PRIJEMNICA
        Case "AMB_ISPLATE", "AMB_UPLATE": TabelaTipa = TBL_NOVAC
        Case "REVERSI":     TabelaTipa = TBL_AMBALAZA
        Case "FAKTURA":     TabelaTipa = TBL_FAKTURE
        Case "IZVOD":       TabelaTipa = TBL_BANKA_IMPORT
        Case Else:          TabelaTipa = TBL_OTKUP
    End Select
End Function

Public Function modeKey(ByVal mode As String) As String
    Select Case mode
        Case "F1": modeKey = "OTKUP"
        Case "F2": modeKey = "OTPREMNICA"
        Case "F3": modeKey = "ZBIRNA"
        Case "F4": modeKey = "PRIJEMNICA"
        Case "F5": modeKey = "AMB_ISPLATE"
        Case "F6": modeKey = "AMB_UPLATE"
        Case "F7": modeKey = "REVERSI"
        Case Else: modeKey = "OTKUP"
    End Select
End Function

' Rezimi bez pojma "zbirne" - cipovi "Bez zbirne" / "Nefakturisane" se skrivaju.
' Faktura postoji SAMO nad prijemnicom (tblPrijemnica.FakturaID). Nad otkupnim
' listom pojam "fakturisano" nema smisla - otkup je nabavka, ne prodaja - pa se
' cip tamo i ne prikazuje. Ranije je "Nefakturisane" bio doslovan duplikat cipa
' "Bez zbirne": isti brojac i isti izraz u MatchFilterFast.
Public Function ColFakturaID(ByVal mk As String) As String
    If mk = "PRIJEMNICA" Then ColFakturaID = COL_PRJ_FAKTURA_ID
End Function

Public Function ModeHasFaktura(ByVal mode As String) As Boolean
    ModeHasFaktura = (Len(ColFakturaID(modeKey(mode))) > 0)
End Function

' Pojam "bez zbirne" postoji samo tamo gde dokument NOSI broj zbirne.
Public Function ModeHasZbirna(ByVal mode As String) As Boolean
    Select Case modeKey(mode)
        Case "OTKUP", "OTPREMNICA", "PRIJEMNICA": ModeHasZbirna = True
        Case Else:                                ModeHasZbirna = False
    End Select
End Function

Public Function ModeTextCol3(ByVal mode As String) As Boolean
    Select Case mode
        Case "F5", "F6", "F7": ModeTextCol3 = True
    End Select
End Function

' Broji li lista TOG TIPA komade umesto dinara.
'
' Pitanje pripada TIPU LISTE ("REVERSI"), ne rezimu unosa ("F7"): istu listu
' prikazuje i storno ekran, koji rezime uopste nema nego bira tip. Dok je
' poredjenje stajalo samo unutar ModeBrojiKomade, drugi pozivalac ga nije mogao
' upotrebiti a da ne prepise literal -- pa bi se dva mesta vremenom razisla.
'
' Zamka za onoga ko ovo dira: ModeBrojiKomade prima F-KLJUC. Pozvati ga sa
' tip-kljucem ("REVERSI") izgleda ispravno a tiho vraca False, jer modeKey
' nepoznat kljuc svodi na "OTKUP".
Public Function TipBrojiKomade(ByVal tk As String) As Boolean
    TipBrojiKomade = (tk = "REVERSI")
End Function

' Broji li podnozje komade umesto dinara u rezimu unosa dokumenata.
Public Function ModeBrojiKomade(ByVal mode As String) As Boolean
    ModeBrojiKomade = TipBrojiKomade(modeKey(mode))
End Function

' UGOVOR EKRANA. Ljuska vise ne cita ActiveMode sama: taj rezim pripada OVOM
' ekranu, pa na Uvozu izvoda ili Fakturisanju nema nikakvo znacenje. Ovde se
' odgovara iz sopstvenog stanja, a ekrani koji komade ne broje ovo ne
' implementiraju i dobijaju dinare.
Public Function Scr_BrojiKomade() As Boolean
    Scr_BrojiKomade = ModeBrojiKomade(ActiveMode)
End Function

' Svako kretanje ambalaze je DVOJNI upis - dva reda sa istim brojem i istim
' DokumentTip-om, jedna noga na kooperantu, druga na otkupnom mestu (vidi
' modOtkup.SaveOtkup i modDokumenta.SaveOMUlaz_TX). Prikazivati obe znaci
' duplirati svaki dokument i pokazivati otkupno mesto kao "partnera" tamo gde
' ono nije protivpartner nego samo knjigovodstvena protivstavka.
' Zato se po tipu dokumenta bira SAMO noga koja nosi znacenje:
'   revers/otkup ka kooperantu  -> noga kooperanta
'   revers firma <-> OM         -> noga otkupnog mesta (tada OM JESTE partner)
Public Function RevRowVisible(ByVal dokTip As String, ByVal entTip As String) As Boolean
    Select Case Trim$(dokTip)
        Case DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, DOK_TIP_OTKUP
            RevRowVisible = (Trim$(entTip) = "Kooperant")
        Case DOK_TIP_OM_IZLAZ_FIRMA, DOK_TIP_OM_ULAZ_FIRMA
            RevRowVisible = (Trim$(entTip) = "Stanica")
    End Select
End Function

' Gotovinski promet (tblNovac) nema kilograme, ali ima KANAL: novac je stigao
' na blagajnu (kes) ili preko izvoda / virmana (banka). Zato 4. kolona mreze
' u tim rezimima nosi kanal umesto kilograma - nista se ne gubi.
Public Function ModeHasKanal(ByVal mode As String) As Boolean
    Select Case mode
        Case "F5", "F6": ModeHasKanal = True
    End Select
End Function

' 1 = kes (blagajna), 2 = banka (izvod ili virman)
'
' "BIM:" u Napomeni je JEDINI trag veze novac -> bankovni izvod: tblNovac nema
' BankaImportID kolonu (modBankaMapiranje.BuildBIMNapomena, modConfig
' NOV_NAPOMENA_BIM_PREFIX). Zato se gleda PRVO, pre Tip-a: na strani kupaca
' kanal uopste nije razdvojen u Tip-u (isti KupciUplata nastaje i rucnim unosom
' u frmDokumenta i mapiranjem izvoda), pa je Napomena tamo jedini izvor.
' Redovi uvezeni iz izvoda PRE razdvajanja kanala nose KES tip - i njih
' "BIM:" ispravno svrstava u banku.
Public Function KanalCode(ByVal tip As String, ByVal napomena As String) As Long
    If Left$(LTrim$(napomena), Len(NOV_NAPOMENA_BIM_PREFIX)) = NOV_NAPOMENA_BIM_PREFIX Then
        KanalCode = 2
        Exit Function
    End If
    Select Case Trim$(tip)
        Case NOV_VIRMAN_FIRMA_OTKUPAC, NOV_VIRMAN_FIRMA_KOOP, NOV_VIRMAN_AVANS_KOOP, _
             NOV_BANKA_UPLATA, NOV_BANKA_ISPLATA
            KanalCode = 2
        Case Else
            KanalCode = 1
    End Select
End Function

' Kod reversa EntitetID pokazuje u RAZLICITU tabelu zavisno od EntitetTip
' (modAmbalaza koristi "Kooperant" / "Stanica" / "Kupac"), pa se partner ne moze
' razresiti jednim recnikom kao kod ostalih rezima.
Public Function RevPartner(ByVal entTip As String, ByVal entID As String, _
                            mKoop As Object, mStan As Object, mKup As Object) As String
    Dim d As Object
    RevPartner = entID
    Select Case Trim$(entTip)
        Case "Kooperant": Set d = mKoop
        Case "Stanica":   Set d = mStan
        Case "Kupac":     Set d = mKup
        Case Else:        Exit Function
    End Select
    If d Is Nothing Then Exit Function
    If d.Exists(entID) Then RevPartner = d(entID)
End Function

Public Function PayCode(ByVal duguje As Double, ByVal placeno As Double) As Long
    If duguje <= 0 Then
        PayCode = IIf(placeno > 0, PAY_PLACENO, PAY_NEPLAC)
    ElseIf placeno >= duguje - 0.005 Then      ' tolerancija na zaokruzenje para
        PayCode = PAY_PLACENO
    ElseIf placeno > 0 Then
        PayCode = PAY_DELIM
    Else
        PayCode = PAY_NEPLAC
    End If
End Function

Public Function KanalNaziv(ByVal code As Long) As String
    If code = 2 Then KanalNaziv = Poruka("OTKUI_KANAL_BANKA") Else KanalNaziv = Poruka("OTKUI_KANAL_KES")
End Function

' Ugovor: Array(kolone, redovi, n, zbirKg, zbirVal, brojaciCipova).
' Ovo je bivsi modOtkupUI.FillGrid. Do S4b je pisao PRAVO u stanje ljuske
' (mView, mViewN, mSumKg, mCnt*), pa je mreza mogla da sluzi samo njega.
' Sada vraca - i mreza je time postala neutralna. Sortiranje radi ljuska.
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' F1: rang kooperanata je svoja lista; ostalo je lista dokumenata rezima.
    Select Case Scr_Lista()
        Case "OTPREMNICE": Scr_Rows = RowsOtpremnice(filter, q): Exit Function
        Case "BLOKOVI":    Scr_Rows = RowsBlokovi(q): Exit Function
        Case "NEVEZANI":   Scr_Rows = RowsNevezani(q): Exit Function
        Case "KOOPERANTI": Scr_Rows = RowsKooperanti(q): Exit Function
    End Select
    ' Tip dolazi iz rezima -- ovaj ekran pokazuje dokument koji se u njemu
    ' unosi. OTKUP nosi nevidljiv OtkupID: radnje reda (stampa, storno) idu po
    ' njemu, ne po broju (S1e). Ostali tipovi ovde radnje reda nemaju.
    ' OTPREMNICA nosi nevidljiv OtpremnicaID: klik na nacrt otvara njegovu
    ' izmenu po ID-u (S3b-2).
    Dim mk As String: mk = modeKey(ActiveMode)
    Scr_Rows = RedoviZaTip(mk, filter, q, (mk = "OTKUP" Or mk = "OTPREMNICA"))
End Function

' Lista dokumenata JEDNOG TIPA. Javna i parametrizovana tipom, jer je ista
' lista potrebna dvama ekranima: unosni rezim pokazuje svoj tip, a Storno
' pokazuje tip koji je operater izabrao cipom.
'
' Do v6-ui-143 je ovo bilo Private i citalo ActiveMode na tri mesta (tabela,
' kljuc tipa, recnik partnera), pa je "koji tip" dolazilo iz rezima ljuske.
' Kad je storno postao svoj ekran, rezim vise ne kaze nista o tipu -- a kopija
' ovih 250 linija u ekranski modul bi bila drugo mesto na kome se odlucuje
' sta je kolona "partner". Zato tip ulazi kao argument.
'
' saIdentitetom dodaje NEVIDLJIVU kolonu kanonskog identiteta (v. GridCols).
' Trazi je samo Storno: unosni rezimi je ne prikazuju i ne koriste, a
' bezuslovno dodavanje bi menjalo sirinu mreze svima.
Public Function RedoviZaTip(ByVal tk As String, ByVal filter As String, ByVal q As String, _
                            Optional ByVal saIdentitetom As Boolean = False) As Variant
    Dim src As Variant, r As Long, n As Long, keep As Boolean, nRows As Long
    Dim outA() As Variant, c As Long, mk As String, tblName As String
    Dim cols As Variant, colN As Long
    Dim sumKg As Double, sumVal As Double
    Dim cOtk As Long, cBez As Long, cNef As Long
    Dim fltV As String, fltP As String
    Dim ix() As Long, kind() As String
    Dim iStorno As Long, iZbir As Long, iKg As Long, iDokTip As Long, iFakt As Long
    Dim iTip As Long, iNap As Long, iEntTip As Long, iBrojCol As Long
    Dim iKoopID As Long, iPartID As Long, iOtkID As Long
    Dim mKoop As Object, mStan As Object, mKup As Object
    Dim rev As Boolean, kanal As Boolean

    ' Izvod nije red tabele nego GRUPA redova (jedan izvod = mnogo stavki), pa
    ' se ne moze prikazati kroz listu dokumenata (koja je red = dokument).
    ' Dispecer stoji OVDE, a ne kod pozivaoca: tada ekran zna samo "daj mi
    ' redove ovog tipa" i ne mora da pamti koji je tip izuzetak.
    If tk = "IZVOD" Then
        RedoviZaTip = RowsIzvodi(filter, q)
        Exit Function
    End If

    On Error GoTo EH
    mStep = "start"
    tblName = TabelaTipa(tk)
    ' suzavanje iz panela "Filteri" i danasnji datum drzi ljuska - ekran ih
    ' cita, ne pamti
    fltV = modOtkupUI.FltVrsta()
    fltP = modOtkupUI.FltPart()
    mToday = Int(Now)
    mMonthStart = CDbl(DateSerial(Year(Now), Month(Now), 1))
    If Len(tblName) = 0 Then Exit Function

    src = modUiData.CachedTable(tblName)
    If Not IsArray(src) Then
        ' Prazna tabela i NECITLJIVA tabela su ovde izgledale isto: tiho
        ' "Exit Function" je crtalo praznu mrezu bez ijedne reci, pa operater
        ' nije imao nacin da razlikuje "nema dokumenata" od "ne umem da
        ' procitam tabelu". Prazna prolazi kao prazna; nepostojeca je greska.
        If Not modUiData.TabelaCitljiva(tblName) Then
            Err.Raise ERR_UI_BASE + 21, "modScrDokumenti.RowsDokumenti", _
                      "Tabela '" & tblName & "' nije nadjena u svesci."
        End If
        Exit Function
    End If

    mStep = "GridCols"
    ' Sve odluke ispod (koje kolone, koji recnik partnera, ima li placanja)
    ' pripadaju TIPU koji je pozivalac zatrazio. Do v6-ui-118 je ovde stajao
    ' goli modeKey pa je storno mogao da cita samo tblOtpremnica; do
    ' v6-ui-143 je stajao EffKey nad ActiveMode, pa je tip dolazio iz rezima.
    '
    ' Nekadasnja "druga brana" (filter tvrdo na "otkazane") je uklonjena
    ' NAMERNO: Storno sada stornira, pa mu je radna lista lista AKTIVNIH
    ' dokumenata. Pregled storniranih ostaje - kroz cip "Otkazane", isti
    ' onaj koji rade svi ostali rezimi.
    mk = tk
    cols = GridCols(mk, saIdentitetom)
    colN = UBound(cols) + 1

    ' indeksi izvornih kolona - JEDNOM po pozivu, ne po redu
    ReDim ix(0 To colN - 1)
    ReDim kind(0 To colN - 1)
    iKg = -1
    For c = 0 To colN - 1
        kind(c) = ColF(CStr(cols(c)), 2)
        ix(c) = ColIdx(tblName, ColF(CStr(cols(c)), 1))
        If kind(c) = "kg" Then iKg = c
    Next c

    ' OTKUP i OTPREMNICA: kolicina, gajbe, klase i vrednost dokumenta su na
    ' STAVKAMA, ne na zaglavlju -- CreateOtkup_TX i CreateOtpremnicaDraft_TX ih
    ' tamo ostavljaju prazne, pa je mreza za nov dokument pokazivala 0 kg i
    ' pilulu "placeno" posle prve delimicne isplate (REFAKTOR S14.7, kvar 2).
    ' Opis kolona ostaje isti; menja se samo izvor tih celija. Stavke se citaju
    ' JEDNOM po pozivu, kao i novac ispod.
    '
    ' Isti mehanizam za oba tipa, jer je i kvar isti: cim jedan tip dobije
    ' svoju kopiju petlje, sledeci cutover je trece mesto na kome se odlucuje
    ' odakle dolazi kilaza reda.
    Dim otkStav As Boolean, dStav As Object, ovStav() As String, iStavID As Long
    ReDim ovStav(0 To colN - 1)
    otkStav = (mk = "OTKUP" Or mk = "OTPREMNICA")
    If otkStav Then
        mStep = "stavke dokumenta"
        If mk = "OTKUP" Then
            Set dStav = modOtkup.ZbirStavkiPoOtkupu()
            iStavID = ColIdx(tblName, COL_OTK_ID)
        Else
            Set dStav = modDokumenta.ZbirStavkiPoOtpremnici()
            iStavID = ColIdx(tblName, COL_OTP_ID)
        End If
        ' Kolone se prepoznaju kroz ISTE Col* funkcije koje su ih i dodale u
        ' GridCols. Golim konstantama bi se nabrajala oba tipa, a tblOtkupStavke
        ' i tblOtpremnicaStavke dele imena kolona ("Klasa", "Kolicina",
        ' "KolAmbalaze") -- dve grane bi izgledale kao izbor, a bile isti string.
        '
        ' Kolona BEZ izvorne kolone (status, placanje) se preskace: Col* vraca ""
        ' za rezim koji to polje nema (OTPREMNICA nema cenu, review #362), a
        ' "Case """ bi se poklopio bas sa njima -- pilula statusa bi dobila
        ' vrednost dokumenta.
        Dim izvCol As String
        For c = 0 To colN - 1
            izvCol = ColF(CStr(cols(c)), 1)
            If Len(izvCol) > 0 Then
                Select Case izvCol
                    Case ColKolicina(mk): ovStav(c) = "kg"
                    Case ColCena(mk):     ovStav(c) = "vr"
                    Case ColKolAmb(mk):   ovStav(c) = "amb"
                    Case ColKlasa(mk):    ovStav(c) = "kl"
                End Select
            End If
        Next c
    End If

    mStep = "indeksi kolona"
    iStorno = ColIdx(tblName, COL_STORNIRANO)
    iZbir = ColIdx(tblName, ColBrojZbirne(mk))
    If Len(ColFakturaID(mk)) > 0 Then iFakt = ColIdx(tblName, ColFakturaID(mk))

    rev = (mk = "REVERSI")
    If rev Then
        iBrojCol = ColIdx(tblName, COL_AMB_DOK_ID)
        iDokTip = ColIdx(tblName, COL_AMB_DOK_TIP)
        iEntTip = ColIdx(tblName, COL_AMB_ENTITET_TIP)
    End If
    kanal = (mk = "AMB_ISPLATE" Or mk = "AMB_UPLATE")
    If kanal Then
        iTip = ColIdx(tblName, COL_NOV_TIP)
        iNap = ColIdx(tblName, COL_NOV_NAPOMENA)
        iEntTip = ColIdx(tblName, COL_NOV_ENTITET_TIP)
        iKoopID = ColIdx(tblName, COL_NOV_KOOP_ID)
        iPartID = ColIdx(tblName, COL_NOV_PARTNER_ID)
        iOtkID = ColIdx(tblName, COL_NOV_OTKUP_ID)
    End If

    mStep = "partner mape"
    If rev Or kanal Then
        Set mKoop = PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Set mStan = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")
        Set mKup = PartnerMap(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
    End If

    ' Placanje: gotove bulk rutine iz modNovac, jedan prolaz po tabeli novca.
    ' Za otkup je vezivanje direktno (tblNovac.OtkupID); za prijemnicu ide preko
    ' fakture, pa se stanje cita NA NIVOU FAKTURE - vidi PayCode.
    Dim pay As Boolean, dPay As Object, dFakIzn As Object
    Dim iPayID As Long
    mStep = "placanje"
    pay = (mk = "OTKUP" Or mk = "PRIJEMNICA")
    If pay Then
        If mk = "OTKUP" Then
            Set dPay = modNovac.BuildIsplataDictByOtkup()
            iPayID = ColIdx(tblName, COL_OTK_ID)
        Else
            Set dPay = modNovac.BuildUplataDictByFaktura()
            Set dFakIzn = FakturaIznosMap()
            iPayID = ColIdx(tblName, COL_PRJ_FAKTURA_ID)
        End If
        If dPay Is Nothing Then pay = False
    End If

    Dim pl As Variant, pmap As Object
    pl = PartnerLookupTip(tk)
    If Len(CStr(pl(0))) > 0 Then _
        Set pmap = PartnerMap(CStr(pl(0)), CStr(pl(1)), CStr(pl(2)), CStr(pl(3)))

    mStep = "petlja po redovima"
    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To colN)
    n = 0

    For r = 1 To nRows
        Dim vDatK As Double, vZbir As String, hay As String
        Dim isStorno As Boolean, bezZbirne As Boolean, bezFakture As Boolean
        Dim vKgRow As Double, cell As Variant

        ' reversi su podskup tblAmbalaza - ostalo iz te knjige ne ulazi
        If rev Then
            If Not RevRowVisible(CellS(src, r, iDokTip), CellS(src, r, iEntTip)) Then GoTo NextRow
        End If

        vDatK = 0
        vKgRow = 0
        hay = ""
        Dim zStav As Variant
        zStav = Empty
        If otkStav Then
            ' Nedostajuci kljuc NIJE nula (review #334, P1): red dokumenta bez
            ' stavki pada po imenu. Ranije je takav red imao duguje = 0, pa je
            ' pilula pokazivala "placeno" na dokumentu bez ijedne stavke.
            If mk = "OTKUP" Then
                zStav = modOtkup.ZbirStavkiZaOtkup(dStav, CellS(src, r, iStavID), _
                            "modScrDokumenti.RedoviZaTip")
            Else
                zStav = modDokumenta.ZbirStavkiZaOtpremnicu(dStav, CellS(src, r, iStavID), _
                            "modScrDokumenti.RedoviZaTip")
            End If
        End If
        If iKg >= 0 Then vKgRow = CellD(src, r, ix(iKg))

        Dim pCode As Long, pRest As Double, duguje As Double, placeno As Double
        Dim payKey As String
        pCode = 0: pRest = 0
        If pay Then
            duguje = 0: placeno = 0
            payKey = CellS(src, r, iPayID)
            If mk = "OTKUP" Then
                duguje = CDbl(zStav(1))
                If dPay.Exists(payKey) Then placeno = CDbl(dPay(payKey))
                pCode = PayCode(duguje, placeno)
            ElseIf Len(payKey) = 0 Then
                pCode = PAY_NEFAKT              ' prijemnica jos nije na fakturi
            Else
                If Not dFakIzn Is Nothing Then
                    If dFakIzn.Exists(payKey) Then duguje = CDbl(dFakIzn(payKey))
                End If
                If dPay.Exists(payKey) Then placeno = CDbl(dPay(payKey))
                pCode = PayCode(duguje, placeno)
            End If
            pRest = duguje - placeno
            If pRest < 0 Then pRest = 0
        End If

        For c = 0 To colN - 1
            Select Case kind(c)
                Case "txt"
                    cell = CellS(src, r, ix(c))
                    hay = hay & "|" & cell
                Case "part"
                    cell = CellS(src, r, ix(c))
                    If rev Then
                        cell = RevPartner(CellS(src, r, iEntTip), CStr(cell), mKoop, mStan, mKup)
                    ElseIf kanal Then
                        cell = NovacPartner(CellS(src, r, iEntTip), CellS(src, r, iKoopID), _
                                            CellS(src, r, iPartID), CellS(src, r, iOtkID), _
                                            CStr(cell), mKoop, mStan, mKup)
                    ElseIf Not pmap Is Nothing Then
                        If pmap.Exists(cell) Then cell = pmap(cell)
                    End If
                    hay = hay & "|" & cell
                Case "date"
                    vDatK = CellDate(src, r, ix(c))
                    cell = vDatK
                Case "kg", "num"
                    cell = CellD(src, r, ix(c))
                Case "sum0", "rsd"
                    cell = CellD(src, r, ix(c))
                    ' F5 i F6 dele tblNovac - red bez iznosa u SVOJOJ koloni
                    ' pripada drugom smeru i odbacuje se ovde, ne u filteru
                    If cell = 0 Then GoTo NextRow
                Case "mult"
                    cell = CellD(src, r, ix(c)) * vKgRow
                Case "paypill"
                    cell = pCode
                Case "rest"
                    ' ostatak ima smisla samo dok nije placeno do kraja
                    If pCode = PAY_DELIM Or pCode = PAY_NEPLAC Then
                        cell = pRest
                    Else
                        cell = 0
                    End If
                Case "osnov"
                    cell = OsnovNaziv(CellS(src, r, iDokTip), CellS(src, r, iBrojCol))
                    hay = hay & "|" & cell
                Case "kanal"
                    cell = KanalNaziv(KanalCode(CellS(src, r, iTip), CellS(src, r, iNap)))
                    hay = hay & "|" & cell
                Case Else
                    cell = ""
            End Select
            ' OTKUP: izvor ovih celija su stavke (ovStav iznad), ne zaglavlje.
            If otkStav Then
                Select Case ovStav(c)
                    Case "kg":  cell = CDbl(zStav(0))
                    Case "vr":  cell = CDbl(zStav(1))
                    Case "amb": cell = CDbl(zStav(2))
                    Case "kl"
                        cell = CStr(zStav(3))
                        hay = hay & "|" & cell
                End Select
            End If
            outA(n + 1, c + 1) = cell
        Next c

        vZbir = CellS(src, r, iZbir)
        isStorno = (iStorno > 0)
        If isStorno Then isStorno = (UCase$(CellS(src, r, iStorno)) = "DA")
        bezZbirne = (Len(vZbir) = 0)
        hay = hay & "|" & vZbir

        bezFakture = False
        If iFakt > 0 Then bezFakture = (Len(CellS(src, r, iFakt)) = 0)

        ' brojaci cipova - isti prolaz, bez zasebnog skena po cipu.
        ' "Bez zbirne" i "Nefakturisane" su RAZLICITI uslovi i broje se odvojeno.
        If isStorno Then
            cOtk = cOtk + 1
        Else
            If bezZbirne Then cBez = cBez + 1
            If bezFakture Then cNef = cNef + 1
        End If

        keep = MatchFilterFast(filter, vDatK, bezZbirne, isStorno, bezFakture)
        If keep And Len(q) > 0 Then keep = (InStr(1, hay, q, vbTextCompare) > 0)
        ' dodatni uslovi iz panela Filteri - isti "hay", bez drugog prolaza
        If keep And Len(fltV) > 0 Then keep = (InStr(1, hay, fltV, vbTextCompare) > 0)
        If keep And Len(fltP) > 0 Then keep = (InStr(1, hay, fltP, vbTextCompare) > 0)

        If keep Then
            n = n + 1
            For c = 0 To colN - 1
                Select Case kind(c)
                    Case "kg":                 sumKg = sumKg + CDbl(outA(n, c + 1))
                    Case "rsd", "mult", "sum0": sumVal = sumVal + CDbl(outA(n, c + 1))
                    Case "pill":               outA(n, c + 1) = StatusCode(isStorno, bezZbirne)
                End Select
            Next c
        End If
NextRow:
    Next r

    ' Sortiranje se vise NE radi ovde. Redovi idu ljusci u redosledu citanja,
    ' a ona ih rasporedjuje po koloni koju je korisnik izabrao - isto za svaki
    ' ekran. Dok je sortiranje bilo unutar ovog koda, mreza je umela da sortira
    ' samo dokumenta.
    mStep = "OK"
    RedoviZaTip = Array(cols, outA, n, sumKg, sumVal, Array(cOtk, cBez, cNef))
    Exit Function
EH:
    ' greska se NE guta - ReloadGrid je prijavljuje sa imenom koraka
    Err.Raise Err.Number, "modScrDokumenti.RedoviZaTip[" & mStep & "]", Err.description
End Function

Public Function MatchFilterFast(ByVal filter As String, ByVal vDatK As Double, _
                                 ByVal bezZbirne As Boolean, ByVal isStorno As Boolean, _
                                 ByVal bezFakture As Boolean) As Boolean
    Select Case filter
        Case "otkazane":  MatchFilterFast = isStorno
        Case "bezzbirne": MatchFilterFast = (Not isStorno) And bezZbirne
        Case "nefakt":    MatchFilterFast = (Not isStorno) And bezFakture
        Case "danas":     MatchFilterFast = (Not isStorno) And (vDatK = mToday)
        Case "nedelja":   MatchFilterFast = (Not isStorno) And (vDatK >= mToday - 6)
        Case "mesec":     MatchFilterFast = (Not isStorno) And (vDatK >= mMonthStart)
        Case Else:        MatchFilterFast = Not isStorno
    End Select
End Function

' Citljiv osnov reda. DOK_TIP_OTKUP je tu jer kooperant i pri predaji PUNIH
' gajbi ima izlaz ambalaze - to nije revers, ali jeste njegovo kretanje.
Public Function OsnovNaziv(ByVal dokTip As String, ByVal dokID As String) As String
    Dim izOtkupa As Boolean
    On Error Resume Next
    izOtkupa = OtkupKoopMap().Exists(Trim$(dokID))
    Select Case Trim$(dokTip)
        Case DOK_TIP_OM_IZLAZ_KOOP
            ' Prazne gajbe uz otkup se knjize ISTIM tipom kao pravi revers
            ' (modOtkup.bas:611), samo im je DokumentID = OtkupID. Bez ove
            ' razlike bi dva reda istog otkupa nosila razlicit osnov: jedan
            ' "Uz otkup", drugi "Revers" - a revers dokument ne postoji.
            OsnovNaziv = IIf(izOtkupa, Poruka("OTKUI_OSN_OTKUP_PRAZNE"), _
                                       Poruka("OTKUI_OSN_REV_IZDATO"))
        Case DOK_TIP_OM_ULAZ_KOOP:   OsnovNaziv = Poruka("OTKUI_OSN_REV_POVRAT")
        Case DOK_TIP_OM_IZLAZ_FIRMA: OsnovNaziv = Poruka("OTKUI_OSN_REV_OM_IZDATO")
        Case DOK_TIP_OM_ULAZ_FIRMA:  OsnovNaziv = Poruka("OTKUI_OSN_REV_OM_PRIJEM")
        Case DOK_TIP_OTKUP:          OsnovNaziv = Poruka("OTKUI_OSN_OTKUP_PUNE")
        Case Else:                   OsnovNaziv = dokTip
    End Select
End Function

' FakturaID -> Iznos. Prijemnica ne nosi svoj dug nego ga nasledjuje od fakture,
' pa se ostatak racuna NA NIVOU FAKTURE. Ako jedna faktura pokriva vise
' prijemnica, isti ostatak stoji u svakom njenom redu - to je tacno, ali se
' odnosi na fakturu, ne na pojedinacnu prijemnicu.
Public Function FakturaIznosMap() As Object
    Dim d As Object, src As Variant, iId As Long, iIzn As Long, r As Long, k As String
    If mMape Is Nothing Then Set mMape = CreateObject("Scripting.Dictionary")
    If mMape.Exists("#FAKIZN") Then
        Set FakturaIznosMap = mMape("#FAKIZN")
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    src = CachedTable(TBL_FAKTURE)
    If IsArray(src) Then
        iId = ColIdx(TBL_FAKTURE, COL_FAK_ID)
        iIzn = ColIdx(TBL_FAKTURE, COL_FAK_IZNOS)
        If iId > 0 And iIzn > 0 Then
            For r = 1 To UBound(src, 1)
                k = CellS(src, r, iId)
                If Len(k) > 0 Then d(k) = CellD(src, r, iIzn)
            Next r
        End If
    End If
    Set mMape("#FAKIZN") = d
    Set FakturaIznosMap = d
End Function

' Kolona Partner u tblNovac NIJE primalac novca. SaveNovac se za isplatu
' kooperantu poziva sa partner:=naziv OTKUPNOG MESTA, entitetTip:="OM",
' omID:=stanicaID, a stvarni primalac ide u KooperantID (modDokumenta:3791,
' modBankaMapiranje.MapBankaImportAsKooperant). Zato se ime vuce iz KooperantID
' kad postoji; tek ako ga nema, red se odnosi na sam OM ili na kupca.
' OtkupID -> KooperantID. Isti pristup koji koristi pregled otkupnih blokova
' (modBankaExportPregled: BuildLookupDict(TBL_OTKUP, COL_OTK_ID, COL_OTK_KOOPERANT)),
' jednom umesto LookupValue po redu. Sluzi dvema stvarima: da se primalac isplate
' nadje i kad red novca nema KooperantID, i da se prepozna da li je red ambalaze
' nastao iz otkupa (DokumentID je tada OtkupID) ili iz zasebnog reversa.
Public Function OtkupKoopMap() As Object
    On Error Resume Next
    If mMape Is Nothing Then Set mMape = CreateObject("Scripting.Dictionary")
    If mMape.Exists("#OTKKOOP") Then
        Set OtkupKoopMap = mMape("#OTKKOOP")
        Exit Function
    End If
    Dim d As Object
    Set d = BuildLookupDict(TBL_OTKUP, COL_OTK_ID, COL_OTK_KOOPERANT)
    If d Is Nothing Then Set d = CreateObject("Scripting.Dictionary")
    Set mMape("#OTKKOOP") = d
    Set OtkupKoopMap = d
End Function

Public Function NovacPartner(ByVal entTip As String, ByVal koopID As String, _
                              ByVal partID As String, ByVal otkID As String, _
                              ByVal partTekst As String, _
                              mKoop As Object, mStan As Object, mKup As Object) As String
    If Len(Trim$(koopID)) > 0 Then
        If Not mKoop Is Nothing Then
            If mKoop.Exists(koopID) Then NovacPartner = mKoop(koopID): Exit Function
        End If
        NovacPartner = koopID
        Exit Function
    End If
    ' red vezan za otkup, a bez KooperantID -> primaoca daje sam otkup
    If Len(Trim$(otkID)) > 0 Then
        Dim k As String
        k = ""
        If OtkupKoopMap().Exists(Trim$(otkID)) Then k = CStr(OtkupKoopMap()(Trim$(otkID)))
        If Len(k) > 0 Then
            If Not mKoop Is Nothing Then
                If mKoop.Exists(k) Then NovacPartner = mKoop(k): Exit Function
            End If
            NovacPartner = k
            Exit Function
        End If
    End If
    Dim d As Object
    Select Case Trim$(entTip)
        Case "Kupac":     Set d = mKup
        Case "OM":        Set d = mStan
        Case "Kooperant": Set d = mKoop
    End Select
    If Not d Is Nothing Then
        If d.Exists(partID) Then NovacPartner = d(partID): Exit Function
    End If
    NovacPartner = partTekst
End Function

Public Function PartnerLookup(ByVal mode As String) As Variant
    Select Case mode
        Case "F2":        PartnerLookup = Array(TBL_STANICE, "StanicaID", "Naziv", "")
        Case "F3", "F4":  PartnerLookup = Array(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
        Case "F1":        PartnerLookup = Array(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Case Else:        PartnerLookup = Array("", "", "", "")
                          ' F5/F6: tblNovac vec ima tekstualnu kolonu Partner
    End Select
End Function

Public Function PartnerLookupTip(ByVal tk As String) As Variant
    Select Case tk
        Case "OTKUP":                PartnerLookupTip = Array(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Case "OTPREMNICA":           PartnerLookupTip = Array(TBL_STANICE, "StanicaID", "Naziv", "")
        Case "ZBIRNA", "PRIJEMNICA", "FAKTURA": _
                                     PartnerLookupTip = Array(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
        Case Else:                   PartnerLookupTip = Array("", "", "", "")
    End Select
End Function

' Ljuska ovo zove kad se podaci promene (RefreshFromData) - ekran mora da
' zaboravi svoje izvedene mape, inace bi posle upisa racunao po starom.
Public Sub Scr_ResetCache()
    Set mMape = Nothing
End Sub

'------------------------------------------------- LISTA: OTPREMNICE (F1)
' Otpremnice kao IZVOR robe. Kljucna kolona je OSTATAK: ocekivano (stavke
' otpremnice) minus povezano (stavke njenih izvora). Oba zbira se racunaju
' JEDNIM prolazom za celu listu, kroz stroge kanonske citace.
Private Function OtpGridCols() As Variant
    OtpGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HDO_KUPAC||part|0|1", _
        "OTKUI_HD_VRSTA||txt|80|2", _
        "OTKUI_HD_SORTA||txt|100|2", _
        "OTKUI_HD_KG||kg|66|1", _
        "OTKUI_HDO_UBLOK||kg|76|1", _
        "OTKUI_HDO_OSTATAK||kg|76|1", _
        "OTKUI_HD_KOL_AMB||num|54|3", _
        "OTKUI_HD_STATUS||txt|64|1", _
        "OTKUI_HD_IDENT||txt|0|4")
End Function

' Filter "otvorene" = NACRTI (jos primaju blokove); svaki drugi znaci "sve".
' Brojac otvorenih ide preko svih otpremnica, i kad je cip ugasen.
Private Function RowsOtpremnice(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, stan As Object, zbir As Object, vred As Object
    Dim iID As Long, iBroj As Long, iDat As Long, iSt As Long, iVr As Long
    Dim iSo As Long, iStat As Long, iStorno As Long
    Dim otpID As String, ocek As Variant, ukKg As Double, blKg As Double
    Dim hay As String, kup As String, st As String, jeNacrt As Boolean
    Dim samoOtvorene As Boolean, cntOtvor As Long, sumOst As Double
    Dim dOd As Double, dDo As Double, vDat As Double
    samoOtvorene = (filter = "otvorene")
    ' Opseg datuma iznad liste (A-022; polja drzi ljuska). Nula znaci "nema
    ' granice", pa prazno polje i nepotpun datum tokom kucanja ne prazne listu.
    dOd = DatGranica(modOtkupUI.GridDatOd())
    dDo = DatGranica(modOtkupUI.GridDatDo())
    On Error GoTo EH
    mStep = "otpremnice"

    src = modUiData.CachedTable(TBL_OTPREMNICA)
    If Not IsArray(src) Then
        RowsOtpremnice = Array(OtpGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0, 0))
        Exit Function
    End If
    mStep = "stavke otpremnica"
    Set zbir = modDokumenta.ZbirStavkiPoOtpremnici()
    mStep = "izvori otpremnica"
    Set vred = modDokumenta.VrednostIzvoraPoOtpremnici()
    mStep = "otpremnice"
    Set stan = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")

    iID = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_ID)
    iBroj = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_BROJ)
    iDat = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_DATUM)
    iSt = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_STANICA)
    iVr = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_VRSTA)
    iSo = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_SORTA)
    iStat = modUiData.ColIdx(TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS)
    iStorno = modUiData.ColIdx(TBL_OTPREMNICA, COL_STORNIRANO)

    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To 11)
    For r = 1 To nRows
        If iStorno > 0 Then
            If UCase$(modUiData.CellS(src, r, iStorno)) = "DA" Then GoTo Sledeca
        End If
        ' Opseg sece i brojac cipa: "otvorenih 3" mora da znaci tri reda koje
        ' operater u tom opsegu i vidi.
        vDat = modUiData.CellDate(src, r, iDat)
        If dOd > 0 Then
            If vDat < dOd Then GoTo Sledeca
        End If
        If dDo > 0 Then
            If vDat > dDo Then GoTo Sledeca
        End If
        otpID = modUiData.CellS(src, r, iID)
        st = UCase$(modUiData.CellS(src, r, iStat))
        jeNacrt = (st = UCase$(IZDATO_DRAFT))
        If jeNacrt Then cntOtvor = cntOtvor + 1
        If samoOtvorene And Not jeNacrt Then GoTo Sledeca

        kup = modUiData.CellS(src, r, iSt)
        If Not stan Is Nothing Then
            If stan.Exists(kup) Then kup = CStr(stan(kup))
        End If
        hay = modUiData.CellS(src, r, iBroj) & "|" & modUiData.CellS(src, r, iVr) & _
              "|" & modUiData.CellS(src, r, iSo) & "|" & kup
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeca
        End If

        ocek = modDokumenta.ZbirStavkiZaOtpremnicu(zbir, otpID, "modScrDokumenti.RowsOtpremnice")
        ukKg = CDbl(ocek(0))
        blKg = PovezanoKg(vred, otpID)

        n = n + 1
        outA(n, 1) = modUiData.CellS(src, r, iBroj)
        outA(n, 2) = vDat
        outA(n, 3) = kup
        outA(n, 4) = modUiData.CellS(src, r, iVr)
        outA(n, 5) = modUiData.CellS(src, r, iSo)
        outA(n, 6) = ukKg
        outA(n, 7) = blKg
        outA(n, 8) = ukKg - blKg
        outA(n, 9) = CDbl(ocek(2))
        If jeNacrt Then
            outA(n, 10) = Poruka("OTKUI_OTP_ST_NACRT")
        ElseIf modDokumenta.IzdatoStatusJeIzdato(st) Then
            outA(n, 10) = Poruka("OTKUI_OTP_ST_IZDATA")
        Else
            outA(n, 10) = st
        End If
        outA(n, 11) = otpID
        sumOst = sumOst + (ukKg - blKg)
Sledeca:
    Next r

    mStep = "OK"
    ' cetvrti brojac je cip "otvorene"; prva tri pripadaju listi dokumenata
    RowsOtpremnice = Array(OtpGridCols(), outA, n, sumOst, 0#, Array(0, 0, 0, cntOtvor))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsOtpremnice[" & mStep & "]", Err.description
End Function

' Povezano kg otpremnice = kg stavki njenih izvora, po svim klasama.
Private Function PovezanoKg(ByVal vred As Object, ByVal otpID As String) As Double
    If vred.Exists(otpID & "|" & KLASA_I) Then PovezanoKg = CDbl(vred(otpID & "|" & KLASA_I)(0))
    If vred.Exists(otpID & "|" & KLASA_II) Then _
        PovezanoKg = PovezanoKg + CDbl(vred(otpID & "|" & KLASA_II)(0))
End Function

'---------------------------------------------------- LISTA: BLOKOVI (F1)
' Otkupni listovi u SASTAVU aktivne otpremnice (tblOtpremnicaIzvori). Bez
' izabrane otpremnice lista je prazna - to je tacno, ne greska.
Private Function BlokGridCols() As Variant
    BlokGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_KG||kg|66|1", _
        "OTKUI_HD_KOL_AMB||num|54|2", _
        "OTKUI_HD_CENA||num|70|2", _
        "OTKUI_HD_VREDNOST||mult|96|1", _
        "OTKUI_HD_IDENT||txt|0|4")
End Function

Private Function RowsBlokovi(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, koop As Object, clan As Object, izv As Collection
    Dim iBroj As Long, iDat As Long, iKoop As Long, iId As Long
    Dim zbir As Object, z As Variant, i As Long, oid As String
    Dim kg As Double, vr As Double, cena As Double, hay As String
    Dim sumKg As Double, sumVal As Double, ime As String
    On Error GoTo EH
    mStep = "blokovi"

    If Len(mOtpID) = 0 Then
        RowsBlokovi = Array(BlokGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ' Sastav kroz strog citac pisaca -- ekran ne cita tblOtpremnicaIzvori sam.
    mStep = "sastav"
    Set izv = modDokumenta.IzvoriOtpremnice(mOtpID)
    Set clan = CreateObject("Scripting.Dictionary")
    For i = 1 To izv.count
        clan(UCase$(CStr(izv(i)))) = True
    Next i
    If clan.count = 0 Then
        RowsBlokovi = Array(BlokGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    mStep = "blokovi"
    src = modUiData.CachedTable(TBL_OTKUP)
    If Not IsArray(src) Then Exit Function
    Set koop = PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")

    iBroj = modUiData.ColIdx(TBL_OTKUP, COL_OTK_BR_DOK)
    iDat = modUiData.ColIdx(TBL_OTKUP, COL_OTK_DATUM)
    iKoop = modUiData.ColIdx(TBL_OTKUP, COL_OTK_KOOPERANT)
    iId = modUiData.ColIdx(TBL_OTKUP, COL_OTK_ID)

    ' Kolicina, gajbe i vrednost bloka su na STAVKAMA. Dokument sa dve klase
    ' nema jednu cenu, pa je cena reda prosek (vrednost / kg).
    mStep = "stavke blokova"
    Set zbir = modOtkup.ZbirStavkiPoOtkupu()
    mStep = "blokovi"

    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To 8)
    For r = 1 To nRows
        oid = modUiData.CellS(src, r, iId)
        If Not clan.Exists(UCase$(oid)) Then GoTo Sledeci
        ime = modUiData.CellS(src, r, iKoop)
        If Not koop Is Nothing Then
            If koop.Exists(ime) Then ime = CStr(koop(ime))
        End If
        hay = modUiData.CellS(src, r, iBroj) & "|" & ime
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If

        z = modOtkup.ZbirStavkiZaOtkup(zbir, oid, "modScrDokumenti.RowsBlokovi")
        kg = CDbl(z(0))
        vr = CDbl(z(1))
        cena = 0
        If kg > 0 Then cena = vr / kg
        n = n + 1
        outA(n, 1) = modUiData.CellS(src, r, iBroj)
        outA(n, 2) = modUiData.CellDate(src, r, iDat)
        outA(n, 3) = ime
        outA(n, 4) = kg
        outA(n, 5) = CDbl(z(2))
        outA(n, 6) = cena
        outA(n, 7) = vr
        outA(n, 8) = oid
        sumKg = sumKg + kg
        sumVal = sumVal + vr
Sledeci:
    Next r

    mStep = "OK"
    RowsBlokovi = Array(BlokGridCols(), outA, n, sumKg, sumVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsBlokovi[" & mStep & "]", Err.description
End Function

'--------------------------------------------------- LISTA: NEVEZANI (F1)
' Blokovi BEZ AKTIVNE OTPREMNICE (A-025, odluka operatera 19.09.2026). Skup
' racuna kanon (modDokumenta.NevezaniOtkupi) -- ekran ne cita tblOtpremnicaIzvori
' sam, pa lista ne moze da ponudi za vezivanje otkup koji pisac smatra zauzetim.
' Kolona "bila u" nosi broj stornirane otpremnice iz koje je blok oslobodjen;
' prazna je za blok koji nikad nije bio vezan.
Private Function NevGridCols() As Variant
    NevGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_OM||txt|120|2", _
        "OTKUI_HD_KG||kg|66|1", _
        "OTKUI_HD_KOL_AMB||num|54|3", _
        "OTKUI_HD_VREDNOST||mult|96|1", _
        "OTKUI_HDN_BILA_U||txt|110|2", _
        "OTKUI_HD_IDENT||txt|0|4")
End Function

Private Function RowsNevezani(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, koop As Object, stan As Object, nev As Object
    Dim iBroj As Long, iDat As Long, iKoop As Long, iId As Long, iSt As Long
    Dim zbir As Object, z As Variant, oid As String, ime As String, om As String
    Dim bila As String, hay As String, sumKg As Double, sumVal As Double
    On Error GoTo EH
    mStep = "nevezani"

    Set nev = modDokumenta.NevezaniOtkupi()
    If nev.count = 0 Then
        RowsNevezani = Array(NevGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    src = modUiData.CachedTable(TBL_OTKUP)
    If Not IsArray(src) Then
        RowsNevezani = Array(NevGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    Set koop = PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Set stan = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")

    iBroj = modUiData.ColIdx(TBL_OTKUP, COL_OTK_BR_DOK)
    iDat = modUiData.ColIdx(TBL_OTKUP, COL_OTK_DATUM)
    iKoop = modUiData.ColIdx(TBL_OTKUP, COL_OTK_KOOPERANT)
    iSt = modUiData.ColIdx(TBL_OTKUP, COL_OTK_STANICA)
    iId = modUiData.ColIdx(TBL_OTKUP, COL_OTK_ID)

    ' Kilogrami, gajbe i vrednost bloka su na STAVKAMA, kao i u listi blokova.
    mStep = "stavke nevezanih"
    Set zbir = modOtkup.ZbirStavkiPoOtkupu()
    mStep = "nevezani"

    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To 9)
    For r = 1 To nRows
        oid = modUiData.CellS(src, r, iId)
        If Not nev.Exists(UCase$(oid)) Then GoTo Sledeci
        bila = CStr(nev(UCase$(oid)))
        ime = modUiData.CellS(src, r, iKoop)
        If Not koop Is Nothing Then
            If koop.Exists(ime) Then ime = CStr(koop(ime))
        End If
        om = modUiData.CellS(src, r, iSt)
        If Not stan Is Nothing Then
            If stan.Exists(om) Then om = CStr(stan(om))
        End If
        hay = modUiData.CellS(src, r, iBroj) & "|" & ime & "|" & om & "|" & bila
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If

        z = modOtkup.ZbirStavkiZaOtkup(zbir, oid, "modScrDokumenti.RowsNevezani")
        n = n + 1
        outA(n, 1) = modUiData.CellS(src, r, iBroj)
        outA(n, 2) = modUiData.CellDate(src, r, iDat)
        outA(n, 3) = ime
        outA(n, 4) = om
        outA(n, 5) = CDbl(z(0))
        outA(n, 6) = CDbl(z(2))
        outA(n, 7) = CDbl(z(1))
        outA(n, 8) = bila
        outA(n, 9) = oid
        sumKg = sumKg + CDbl(z(0))
        sumVal = sumVal + CDbl(z(1))
Sledeci:
    Next r

    mStep = "OK"
    RowsNevezani = Array(NevGridCols(), outA, n, sumKg, sumVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsNevezani[" & mStep & "]", Err.description
End Function

'--------------------------------------------------------- LISTA: IZVODI
' Bankovni izvod je jedini stornirljiv "dokument" koji NIJE red tabele: tblBankaImport
' cuva pojedinacne stavke izvoda, a stornira se ceo izvod. Zato ova lista
' grupise po (broj izvoda, broj racuna) - isti par koji StornoIzvod_TX trazi.
' Broj sam nije kljuc: dve banke mogu imati izvod istog broja.
Private Function IzvGridCols() As Variant
    IzvGridCols = Array( _
        "OTKUI_HD_BROJ||txt|100|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_RACUN||txt|0|1", _
        "OTKUI_HD_STAVKI||sum0|60|1", _
        "OTKUI_HD_IZNOS||rsd|110|1", _
        "OTKUI_HD_STATUS||pill|88|1")
End Function

' Filter je isti kao svuda: sve osim "otkazane" znaci AKTIVNI izvodi.
' Storniran izvod ovde postoji samo kao pregled - StornoIzvod_TX ga drugi
' put ne prima (GetIzvodStornoBlokade to i kaze).
Private Function RowsIzvodi(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, d As Object, kljuc As String, k As Variant
    Dim iBroj As Long, iDat As Long, iRac As Long, iUpl As Long
    Dim iIsp As Long, iSt As Long
    Dim samoOtkazane As Boolean, jeStorno As Boolean, hay As String
    Dim sumVal As Double, cntOtkaz As Long
    On Error GoTo EH
    mStep = "izvodi"
    samoOtkazane = (filter = "otkazane")

    src = modUiData.CachedTable(TBL_BANKA_IMPORT)
    If Not IsArray(src) Then
        RowsIzvodi = Array(IzvGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    iBroj = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    iDat = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_DATUM_IZVODA)
    iRac = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA)
    iUpl = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    iIsp = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    iSt = modUiData.ColIdx(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO)
    If iBroj = 0 Then
        RowsIzvodi = Array(IzvGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ' Jedan prolaz: kljuc "broj|racun" -> "datum|stavki|iznos|stornirano".
    ' Recnik cuva REDOSLED prvog pojavljivanja, pa se lista ne meni sama od
    ' sebe izmedju dva citanja (sortiranje posle radi ljuska).
    Set d = CreateObject("Scripting.Dictionary")
    nRows = UBound(src, 1)
    For r = 1 To nRows
        Dim br As String, rac As String
        br = modUiData.CellS(src, r, iBroj)
        If Len(br) > 0 Then
            rac = ""
            If iRac > 0 Then rac = modUiData.CellS(src, r, iRac)
            kljuc = br & "|" & rac
            jeStorno = False
            If iSt > 0 Then jeStorno = (UCase$(modUiData.CellS(src, r, iSt)) = "DA")
            Dim v As Variant
            If d.Exists(kljuc) Then
                v = d(kljuc)
            Else
                ' datum | stavki | iznos | stornirano
                v = Array(modUiData.CellDate(src, r, iDat), 0&, 0#, jeStorno)
            End If
            v(1) = CLng(v(1)) + 1
            v(2) = CDbl(v(2)) + Abs(modUiData.CellD(src, r, iUpl)) + Abs(modUiData.CellD(src, r, iIsp))
            ' Izvod je storniran samo ako su SVE njegove stavke stornirane -
            ' delimicno stornirani izvod ne postoji (StornoIzvod_TX ide u
            ' celosti), pa bi "bilo koja" davala lazan status.
            v(3) = CBool(v(3)) And jeStorno
            d(kljuc) = v
        End If
    Next r

    ReDim outA(1 To IIf(d.count = 0, 1, d.count), 1 To 6)
    For Each k In d.keys
        Dim p As Variant, rec As Variant
        p = Split(CStr(k), "|")
        rec = d(k)
        jeStorno = CBool(rec(3))
        If jeStorno Then cntOtkaz = cntOtkaz + 1
        If samoOtkazane <> jeStorno Then GoTo SledeciIzvod
        hay = CStr(p(0)) & "|" & CStr(p(1))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo SledeciIzvod
        End If
        n = n + 1
        outA(n, 1) = CStr(p(0))
        outA(n, 2) = rec(0)
        outA(n, 3) = CStr(p(1))
        outA(n, 4) = CLng(rec(1))
        outA(n, 5) = CDbl(rec(2))
        outA(n, 6) = StatusCode(jeStorno, False)
        sumVal = sumVal + CDbl(rec(2))
SledeciIzvod:
    Next k

    mStep = "OK"
    RowsIzvodi = Array(IzvGridCols(), outA, n, 0#, sumVal, Array(cntOtkaz, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsIzvodi[" & mStep & "]", Err.description
End Function
