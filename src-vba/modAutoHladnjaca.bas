Attribute VB_Name = "modAutoHladnjaca"
Option Explicit

' ============================================================
' modAutoHladnjaca -- hladnjacki lanac (HLD-LANAC-01) i oznake hladnjace.
'
' Blok upisan na stanici sa JeHladnjaca = "Da" sam pravi ceo lanac
' OTK -> OTP -> ZBR -> PRJ, 1:1 po klasi (S6 korak 7, odluke operatera
' 24.10.2026). Stanica je JEDINI autoritet: prekidac AUTO_PRIJEMNICA_HLADNJACA
' je obrisan (odluka 2), pa ne postoji konfiguracija koja bi hladnjacki blok
' pustila u rucni tok.
'
' Lanac NIJE jedna transakcija nego niz postojecih pisaca, svaki u svojoj
' (odluka 1, B): CreateOtpremnicaIzIzvora_TX, jezgro auto-zbirne
' (modMasterSync.AutoZbirnaZaOtpremnicu -- isto kao malina, odluka 3) i
' SavePrijemnicaMulti_TX sa roditeljem po ID-u (odluka 7). Zato je NASTAVLJIV:
' svaki korak je idempotentan po kanonskom roditelju, a ponovljen poziv krece
' od prvog koraka koji fali. Nedovrsen lanac (otpremnica bez zbirne ili zbirna
' bez prijemnice) vidi Oporavak (NedovrseniLanci, odluka 8); blok kome je pala
' vec otpremnica ceka u F1 "Bez otpremnice".
' ============================================================

' Da li je stanica oznacena kao hladnjaca (tblStanice.JeHladnjaca = "Da").
Public Function IsHladnjacaStanica(ByVal stanicaID As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, COL_STA_JE_HLADNJACA), ""))
    IsHladnjacaStanica = (StrComp(v, "Da", vbTextCompare) = 0)
End Function

' Da li je KUPAC oznacen kao hladnjaca-kupac (interni cold-store tok). Isti signal
' kao frmDokumenta.RefreshBrojPrijSuggestion: kupac == CFG_MALINA_DEFAULT_KUPAC.
' Eksterni kupci -> False (za njih je zbirna poslednji interni dokument, a prijemnica
' eksterna -> storno framework ne kaskadira nizvodni tok). Prazan config / prazan
' kupac -> False.
Public Function IsHladnjacaKupac(ByVal kupacID As String) As Boolean
    On Error Resume Next
    kupacID = Trim$(kupacID)
    If Len(kupacID) = 0 Then Exit Function
    Dim h As String
    h = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(h) = 0 Then Exit Function
    IsHladnjacaKupac = (StrComp(kupacID, h, vbTextCompare) = 0)
End Function

' Da li blokovi OVE STANICE idu u auto-lanac. Ista odluka kao LanacVaziZaBlok,
' samo pre nego sto blok postoji -- ekran je treba PRE svih pravila vezanih za
' aktivan rucni nacrt (potvrda prekoracenja, vezivanje), jer hladnjacki blok tim
' pravilima uopste ne podleze.
'
' FAIL-CLOSED, isto kao LanacVaziZaBlok: "ne znam" nije "nije hladnjaca" nego
' razlog. Pozivalac koji dobije razlog ne sme da primeni pravila rucnog toka.
Public Function LanacVaziZaStanicu(ByVal stanicaID As String, _
                                   Optional ByRef outGreska As String) As Boolean
    Dim errDesc As String

    outGreska = ""
    On Error GoTo EH

    stanicaID = Trim$(stanicaID)
    If Len(stanicaID) = 0 Then Exit Function

    LanacVaziZaStanicu = HladnjacaStrogo(stanicaID)
    Exit Function

EH:
    errDesc = Err.description
    LogErr "modAutoHladnjaca.LanacVaziZaStanicu"
    outGreska = Poruka("OTKUI_ERR_LANAC_PUT") & " " & errDesc
    LanacVaziZaStanicu = False
End Function

' Da li OVAJ blok ide u auto-lanac umesto na radni sto. Ekran ovim grana PRE
' rucnog vezivanja; ceo sud je ovde, da ga sledeci pozivalac ne prepisuje.
'
' FAIL-CLOSED. Ovo nije prikaz nego RAZVODNICA izmedju dva toka, a JeHladnjaca je
' tvrda poslovna granica: hladnjacki blok MORA u lanac. Zato postoje tri ishoda,
' ne dva:
'
'   sigurno hladnjaca       -> True
'   sigurno nije hladnjaca  -> False
'   ne moze da se utvrdi    -> False + RAZLOG (outGreska)
'
' Pozivalac koji dobije razlog ne sme da nastavi NI JEDNIM putem: gurnuti blok u
' rucni nacrt "jer provera nije uspela" znaci da hladnjacki blok tiho zavrsi kao
' clan tudjeg dokumenta -- tacno ono sto je ovaj PR isao da spreci. Blok je vec
' upisan svojom transakcijom i ostaje nevezan, vidljiv u listi "Bez otpremnice".
'
' Isto pravilo drzi modStornoDok.StornoTraziIzborModa: neizvesnost vodi ka VISE
' pitanja, ne ka manje.
Public Function LanacVaziZaBlok(ByVal otkupID As String, _
                                Optional ByRef outGreska As String) As Boolean
    Dim stanicaID As String, redovi As Collection, errDesc As String

    outGreska = ""
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    ' Blok mora da postoji tacno jednom -- inace se o njemu ne zna nista, pa ni
    ' kojim putem ide.
    Set redovi = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If redovi Is Nothing Then
        Err.Raise vbObjectError + 8460, "LanacVaziZaBlok", _
                  "Citanje otkupa nije uspelo: " & otkupID
    End If
    If redovi.count <> 1 Then
        Err.Raise vbObjectError + 8461, "LanacVaziZaBlok", _
                  "Otkup ne postoji tacno jednom (" & redovi.count & "): " & otkupID
    End If

    If StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO), "")), _
               "Da", vbTextCompare) = 0 Then Exit Function

    stanicaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA), ""))
    If Len(stanicaID) = 0 Then
        Err.Raise vbObjectError + 8462, "LanacVaziZaBlok", _
                  "Otkup nema otkupno mesto: " & otkupID
    End If

    ' IsHladnjacaStanica je fail-open (prikaz), pa se ovde NE koristi: stanica
    ' koja ne postoji davala bi "nije hladnjaca" i blok bi tiho otisao u rucni tok.
    LanacVaziZaBlok = HladnjacaStrogo(stanicaID)
    Exit Function

EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr "modAutoHladnjaca.LanacVaziZaBlok"
    outGreska = Poruka("OTKUI_ERR_LANAC_PUT") & " " & errDesc
    LanacVaziZaBlok = False
End Function

' Da li je OTPREMNICA hladnjacka: njena stanica je hladnjaca. STROGO, isti
' primitiv kao razvodnica bloka -- otpremnica koja ne postoji tacno jednom ili
' nema stanicu DIZE gresku, jer "ne znam" nije "nije hladnjacka".
'
' Zovu je dva mesta sa suprotnim posledicama (S6 korak 7): jezgro auto-zbirne
' (odluka 3) hladnjacku otpremnicu PUSTA i van maline, a rucno vezivanje u zbirnu
' (DodajZbirnaIzvor_TX, odluka 10) je ODBIJA. Pozivalac gresku vraca kao razlog.
Public Function OtpremnicaJeHladnjacka(ByVal otpremnicaID As String) As Boolean
    Dim redovi As Collection, stanicaID As String

    otpremnicaID = Trim$(otpremnicaID)
    Set redovi = FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID)
    If redovi Is Nothing Then
        Err.Raise vbObjectError + 8466, "OtpremnicaJeHladnjacka", _
                  "Citanje otpremnica nije uspelo: " & otpremnicaID
    End If
    If redovi.count <> 1 Then
        Err.Raise vbObjectError + 8467, "OtpremnicaJeHladnjacka", _
                  "Otpremnica ne postoji tacno jednom (" & redovi.count & "): " & otpremnicaID
    End If

    stanicaID = Trim$(nz(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID, COL_OTP_STANICA), ""))
    If Len(stanicaID) = 0 Then
        Err.Raise vbObjectError + 8468, "OtpremnicaJeHladnjacka", _
                  "Otpremnica nema otkupno mesto: " & otpremnicaID
    End If

    OtpremnicaJeHladnjacka = HladnjacaStrogo(stanicaID)
End Function

' Je li stanica hladnjaca -- STROGO. Stanica koja ne postoji ili postoji dvaput
' nije "nije hladnjaca" nego nepoznato stanje, i to se dize kao greska.
Private Function HladnjacaStrogo(ByVal stanicaID As String) As Boolean
    Dim redovi As Collection

    Set redovi = FindRows(TBL_STANICE, "StanicaID", stanicaID)
    If redovi Is Nothing Then
        Err.Raise vbObjectError + 8463, "HladnjacaStrogo", _
                  "Citanje stanica nije uspelo: " & stanicaID
    End If
    If redovi.count <> 1 Then
        Err.Raise vbObjectError + 8464, "HladnjacaStrogo", _
                  "Otkupno mesto ne postoji tacno jednom (" & redovi.count & "): " & stanicaID
    End If

    If GetColumnIndex(TBL_STANICE, COL_STA_JE_HLADNJACA) = 0 Then
        Err.Raise vbObjectError + 8465, "HladnjacaStrogo", _
                  "Kolona " & COL_STA_JE_HLADNJACA & " ne postoji u " & TBL_STANICE
    End If

    HladnjacaStrogo = (StrComp(Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, _
                                                    COL_STA_JE_HLADNJACA), "")), _
                               "Da", vbTextCompare) = 0)
End Function

' ============================================================
' AUTO-LANAC HLADNJACE -- OTK -> OTP -> ZBR -> PRJ (A-014, HLD-LANAC-01).
'
' Roba na hladnjackoj stanici se MERI PRI PRIJEMU u hladnjacu, pa je ceo lanac
' 1:1 sa otkupnim listom, po klasi, i nosi datum otkupa. Nijedan dokument nema
' sta da ceka: otpremnica i zbirna nastaju IZDATE, a prijemnica odmah posle.
'
' Vozac je OGLEDALO stanice (VozacID = StanicaID), i to nije podatak o coveku
' nego PROXY za samu hladnjacu: robu do hladnjace kooperant dovozi SAM, pa taj
' prevoz nema firminog vozaca. Lanac ga trazi kroz modMalina.VozacOgledaloZaStanicu
' -- JEDNO telo koje dele hladnjacki lanac i malina auto-otpremnica (S5-1).
' Kupac zbirne i prijemnice je hladnjaca (MALINA_DEFAULT_KUPAC).
'
' SVAKI NIVO IZVODI STAVKE IZ SVOG KANONSKOG RODITELJA I VEZUJE SE ZA NJEGA PO
' ID-u -- ne prima prepisane brojeve. 1:1 je tada posledica modela, a ne tri
' kopirane vrednosti koje mogu da se raziju. Izuzetak je cena: zbirna je nema, pa
' prijemnica nosi cenu IZVORNOG otkupa te klase (odluka 6), nadjenog po ID-u kroz
' lanac koji nosi tacno jedan blok.
'
' NASTAVLJIV (odluka 1, B). Stanje se cita po kanonskom roditelju: otpremnica po
' clanstvu bloka, zbirna po clanstvu otpremnice, prijemnica po ZbirnaRoditeljID.
' Korak koji postoji se preskace, pa ponovljen poziv nad kompletnim lancem ne
' radi nista, a nad nedovrsenim krece od prvog koraka koji fali. Preduslovi
' (kupac hladnjace, ogledalo, brojevi) se proveravaju PRE PRVOG UPISA -- da se to
' sazna tek na zbirnoj, blok bi ostao sa izdatom otpremnicom i bez puta napred.
'
' AUTOMATIKA JE OBAVEZNA, NE PONUDA: hladnjacki blok ne ide na radni sto nego u
' SVOJ lanac. Zato ekran grana PRE rucnog vezivanja (LanacVaziZaBlok), a rucni
' putevi ne primaju ni blok (VeziZaAktivnu) ni otpremnicu (DodajZbirnaIzvor_TX).
'
' Otkup je vec upisan svojom transakcijom. Ako lanac padne, blok OSTAJE -- ne
' brise se i ne stornira. Oporavak je "otkloni uzrok pa PONOVI": blok bez
' otpremnice u F1 ("Bez otpremnice"), nedovrsen lanac u Oporavku ("Nedovrseno").
'
' UGOVOR POVRATKA (plan 14.55 "Korak 7"):
'   PrijemnicaID         lanac je kompletiran OVIM pozivom; izvestaj nosi brojeve
'                        otpremnice, zbirne i prijemnice;
'   "" + prazan izvestaj nije bilo posla (nije hladnjaca, storniran blok, lanac
'                        vec kompletan);
'   "" + izvestaj        korak nije uspeo; izvestaj kaze koji i sta je vec
'                        napravljeno.
' ============================================================
Public Function AutoLanacHladnjaca(ByVal otkupID As String, _
                                   ByRef izvestaj As String) As String
    Dim stanicaID As String, kupacID As String, vozacID As String
    Dim otpID As String, zbrID As String, prjID As String
    Dim brOtp As String, brPrj As String, datum As Date
    Dim korak As String, greska As String, errDesc As String
    Dim h As Object, izvori As Collection

    izvestaj = ""
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    stanicaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA), ""))
    If StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO), "")), _
               "Da", vbTextCompare) = 0 Then Exit Function

    ' STROGO, isti primitiv kao router: slabija kopija istog pravila (fail-open
    ' IsHladnjacaStanica) bila bi kandidat za drift -- jedan bi rekao "ne znam",
    ' drugi "nije hladnjaca".
    If Not HladnjacaStrogo(stanicaID) Then Exit Function

    ' --- GDE JE LANAC STAO: po kanonskom roditelju, po ID-u ----------------
    otpID = modDokumenta.OtpremnicaZaOtkup(otkupID)
    If Len(otpID) > 0 Then zbrID = modDokumenta.AktivnaZbirnaZaOtpremnicu(otpID)
    If Len(zbrID) > 0 Then prjID = PrijemnicaZaZbirnu(zbrID)
    If Len(prjID) > 0 Then Exit Function

    ' --- PREDUSLOVI, PRE PRVOG UPISA ---------------------------------------
    korak = "preduslov"
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_KUPAC")
        Exit Function
    End If

    vozacID = modMalina.VozacOgledaloZaStanicu(stanicaID)
    If Len(vozacID) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_MIRROR") & " " & stanicaID
        Exit Function
    End If

    datum = CDate(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM))

    ' Broj otpremnice ide iz niza STANICE, kao i svaka druga otpremnica; mirror
    ' prefiks je pravilo broja ZBIRNE (modBrojevi.ApplyMirrorPrefix).
    If Len(otpID) = 0 Then
        brOtp = modBrojevi.GenerateBrojOtpremnice(stanicaID, datum)
        If Len(brOtp) = 0 Then
            izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_BROJ") & " " & stanicaID
            Exit Function
        End If
    End If
    ' Broj zbirne racuna jezgro (niz vozaca-ogledala, bez Google-a); ovde se samo
    ' proverava da ga ima -- prazan znaci ugasen auto-broj, a lanac nema
    ' operatera koji bi ga ukucao.
    If Len(zbrID) = 0 Then
        If Len(modBrojevi.SuggestNextBroj(KIND_ZBR, vozacID, datum, False)) = 0 Then
            izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_BROJ_ZBR") & " " & vozacID
            Exit Function
        End If
    End If
    brPrj = modBrojevi.GenerateBrojPrijemnice(kupacID, datum)
    If Len(brPrj) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_BROJ_PRJ") & " " & kupacID
        Exit Function
    End If

    ' --- OTPREMNICA: iz bloka, odmah IZDATA --------------------------------
    If Len(otpID) = 0 Then
        korak = "OTP"
        Set h = CreateObject("Scripting.Dictionary")
        h("Datum") = datum
        h("StanicaID") = stanicaID
        h("VozacID") = vozacID
        h("KulturaID") = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_KULTURA), ""))
        h("TipAmbalaze") = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_TIP_AMB), ""))
        h("BrojOtpremnice") = brOtp
        Set izvori = New Collection
        izvori.Add otkupID

        otpID = modDokumenta.CreateOtpremnicaIzIzvora_TX(h, izvori, greska)
        If Len(otpID) = 0 Then
            izvestaj = LanacStao(korak, greska, "", "")
            Exit Function
        End If
    End If

    ' --- ZBIRNA: isto jezgro kao malina (odluka 3) -------------------------
    If Len(zbrID) = 0 Then
        korak = "ZBR"
        zbrID = modMasterSync.AutoZbirnaZaOtpremnicu(otpID, greska)
        If Len(zbrID) = 0 Then
            ' "" bez razloga znaci da jezgro nije videlo posao -- za otpremnicu
            ' koju je lanac upravo napravio to je kvar, ne "nista".
            If Len(greska) = 0 Then greska = Poruka("OTKUI_ERR_LANAC_ZBR_TIHO")
            izvestaj = LanacStao(korak, greska, otpID, "")
            Exit Function
        End If
    End If

    ' --- PRIJEMNICA: stavke iz zbirne, cena iz otkupa, roditelj po ID-u -----
    korak = "PRJ"
    prjID = LanacPrijemnica(otkupID, zbrID, kupacID, brPrj, greska)
    If Len(prjID) = 0 Then
        izvestaj = LanacStao(korak, greska, otpID, zbrID)
        Exit Function
    End If

    izvestaj = Poruka("OTKUI_MSG_LANAC") & " " & LanacSpisak(otpID, zbrID, prjID)
    AutoLanacHladnjaca = prjID
    Exit Function

EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr "modAutoHladnjaca.AutoLanacHladnjaca", "korak=" & korak & " otkup=" & otkupID
    izvestaj = LanacStao(korak, errDesc, otpID, zbrID)
    AutoLanacHladnjaca = ""
End Function

' Prijemnica lanca. Zaglavlje (datum, vozac-ogledalo, broj, vrsta, sorta, tip) i
' kilaza/gajbe po klasi dolaze iz ZBIRNE -- kanonskog roditelja. Cena dolazi iz
' IZVORNOG OTKUPA te klase (odluka 6): zbirna cenu nema, a do otkupa se ide po
' ID-u, nivo po nivo, kroz lanac koji mora da nosi TACNO JEDAN blok -- inace
' "cena izvornog otkupa" ne bi bila jedna.
'
' Roditelj se predaje piscu PO ID-u (odluka 7): broj zbirne ume da bude
' dvosmislen (KR-001), a prijemnica vezana po broju tada ostaje bez roditelja.
Private Function LanacPrijemnica(ByVal otkupID As String, ByVal zbrID As String, _
                                 ByVal kupacID As String, ByVal brPrj As String, _
                                 ByRef outGreska As String) As String
    Const SRC As String = "LanacPrijemnica"
    Dim izvZbr As Collection, izvOtp As Collection, poKlasi As Object, cene As Object
    Dim kl As Variant, kgI As Double, ambI As Long, cenaI As Double
    Dim kgII As Double, ambII As Long, cenaII As Double, imaII As Boolean, errNum As Long

    outGreska = ""

    Set izvZbr = modDokumenta.IzvoriZbirne(zbrID)
    If izvZbr.count <> 1 Then
        Err.Raise vbObjectError + 8469, SRC, _
                  "Zbirna lanca " & zbrID & " ne nosi tacno jednu otpremnicu (" & izvZbr.count & ")."
    End If
    Set izvOtp = modDokumenta.IzvoriOtpremnice(CStr(izvZbr(1)))
    If izvOtp.count <> 1 Then
        Err.Raise vbObjectError + 8470, SRC, _
                  "Otpremnica lanca " & CStr(izvZbr(1)) & " ne nosi tacno jedan blok (" & izvOtp.count & ")."
    End If
    If StrComp(CStr(izvOtp(1)), otkupID, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 8470, SRC, _
                  "Otpremnica lanca " & CStr(izvZbr(1)) & " ne nosi blok " & otkupID & "."
    End If

    Set poKlasi = modDokumenta.ZbirnaPoKlasi(zbrID)
    Set cene = CeneOtkupaPoKlasi(otkupID)
    For Each kl In poKlasi.Keys
        If Not cene.Exists(CStr(kl)) Then
            Err.Raise vbObjectError + 8471, SRC, _
                      "Otkup " & otkupID & " nema stavku klase " & CStr(kl) & " koju nosi zbirna " & zbrID & "."
        End If
        Select Case CStr(kl)
            Case UCase$(KLASA_I)
                kgI = CDbl(poKlasi(kl)(0))
                ambI = CLng(poKlasi(kl)(1))
                cenaI = CDbl(cene(kl))
            Case UCase$(KLASA_II)
                imaII = True
                kgII = CDbl(poKlasi(kl)(0))
                ambII = CLng(poKlasi(kl)(1))
                cenaII = CDbl(cene(kl))
            Case Else
                Err.Raise vbObjectError + 8472, SRC, _
                          "Zbirna " & zbrID & " nosi nepoznatu klasu: " & CStr(kl) & "."
        End Select
    Next kl

    LanacPrijemnica = modDokumenta.SavePrijemnicaMulti_TX( _
        datum:=CDate(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_DATUM)), _
        kupacID:=kupacID, _
        vozacID:=Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_VOZAC), "")), _
        brojPrij:=brPrj, _
        brojZbirne:=Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_BROJ), "")), _
        vrstaVoca:=Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_VRSTA), "")), _
        sortaVoca:=Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_SORTA), "")), _
        kolicinaI:=kgI, cenaI:=cenaI, _
        tipAmb:=Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_TIP_AMB), "")), _
        kolAmb:=ambI, kolAmbVracena:=0, _
        hasKlasaII:=imaII, kolicinaII:=kgII, cenaII:=cenaII, kolAmbII:=ambII, _
        outErrNum:=errNum, outGreska:=outGreska, zbirnaRoditeljID:=zbrID)
    If Len(LanacPrijemnica) = 0 And Len(outGreska) = 0 Then
        outGreska = "pisac prijemnice nije vratio PrijemnicaID (greska " & CStr(errNum) & ")"
    End If
End Function

' Cena stavke otkupa po klasi: UCase(klasa) -> Cena. Strog citac stavki
' (StavkeOtkupaRedovi) -- pokvarena stavka pada, ne postaje cena nula.
Private Function CeneOtkupaPoKlasi(ByVal otkupID As String) As Object
    Dim rez As Object, s As Variant, i As Long
    Set rez = CreateObject("Scripting.Dictionary")
    Set CeneOtkupaPoKlasi = rez

    s = modOtkup.StavkeOtkupaRedovi()
    If Not IsArray(s) Then Exit Function
    For i = 1 To UBound(s, 1)
        If StrComp(CStr(s(i, 1)), otkupID, vbTextCompare) = 0 Then
            rez(UCase$(Trim$(CStr(s(i, 3))))) = CDbl(s(i, 5))
        End If
    Next i
End Function

' Aktivna prijemnica ciji je roditelj (po ID-u) data zbirna; "" = nema je.
Private Function PrijemnicaZaZbirnu(ByVal zbrID As String) As String
    Dim mapa As Object
    Set mapa = AktivnePrijemnicePoZbirni()
    If mapa.Exists(UCase$(Trim$(zbrID))) Then PrijemnicaZaZbirnu = CStr(mapa(UCase$(Trim$(zbrID))))
End Function

' UCase(ZbirnaRoditeljID) -> PrijemnicaID, samo AKTIVNE prijemnice.
'
' JEDNA definicija "zbirna ima prijemnicu" za nastavak lanca i za Oporavak.
' Strogo: nedostajuca kolona pada (RequireColumnIndex), jer bi citalac koji pad
' vrati kao "nema prijemnice" pustio lanac da napravi DRUGU (modStorno.
' LookupActiveID gresku guta i vraca "", pa ovde nije upotrebljen).
Private Function AktivnePrijemnicePoZbirni() As Object
    Const SRC As String = "AktivnePrijemnicePoZbirni"
    Dim mapa As Object, d As Variant, i As Long, rod As String
    Dim cId As Long, cRod As Long, cSto As Long

    Set mapa = CreateObject("Scripting.Dictionary")
    Set AktivnePrijemnicePoZbirni = mapa

    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Function
    cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    cRod = RequireColumnIndex(TBL_PRIJEMNICA, COL_DETE_ZBIRNA_ROD, SRC)
    cSto = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, SRC)
    For i = 1 To UBound(d, 1)
        rod = UCase$(Trim$(NzToText(d(i, cRod))))
        If Len(rod) > 0 Then
            If StrComp(Trim$(NzToText(d(i, cSto))), "Da", vbTextCompare) <> 0 Then
                If Not mapa.Exists(rod) Then mapa.Add rod, Trim$(NzToText(d(i, cId)))
            End If
        End If
    Next i
End Function

' Izvestaj o koraku koji nije uspeo: koji korak, zasto, i sta je VEC napravljeno
' -- operater mora da zna da otpremnica postoji i bez zbirne.
Private Function LanacStao(ByVal korak As String, ByVal razlog As String, _
                           ByVal otpID As String, ByVal zbrID As String) As String
    Dim s As String
    s = Poruka("OTKUI_ERR_LANAC")
    Select Case korak
        Case "OTP": s = s & " " & Poruka("OTKUI_ERR_LANAC_KORAK_OTP")
        Case "ZBR": s = s & " " & Poruka("OTKUI_ERR_LANAC_KORAK_ZBR")
        Case "PRJ": s = s & " " & Poruka("OTKUI_ERR_LANAC_KORAK_PRJ")
    End Select
    s = s & " " & razlog
    If Len(otpID) > 0 Then s = s & " " & Poruka("OTKUI_LANAC_VEC") & " " & LanacSpisak(otpID, zbrID, "")
    LanacStao = s
End Function

' "otpremnica <broj>, zbirna <broj>, prijemnica <broj>" -- BROJEVI, po kojima ih
' operater trazi u listama; ID je za masinu.
Private Function LanacSpisak(ByVal otpID As String, ByVal zbrID As String, _
                             ByVal prjID As String) As String
    Dim s As String
    On Error Resume Next
    If Len(otpID) > 0 Then s = Poruka("OTKUI_LANAC_OTP") & " " & _
        Trim$(nz(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ), ""))
    If Len(zbrID) > 0 Then s = s & ", " & Poruka("OTKUI_LANAC_ZBR") & " " & _
        Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_BROJ), ""))
    If Len(prjID) > 0 Then s = s & ", " & Poruka("OTKUI_LANAC_PRJ") & " " & _
        Trim$(nz(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prjID, COL_PRJ_BROJ), ""))
    LanacSpisak = s
End Function

' ============================================================
' SYNC KORAK: NEVEZANI HLADNJACKI BLOKOVI IZ PWA (S6 korak 7, odluka 9).
'
' Malina batch (modMasterSync.AutoCreateOtpremniceFromPWA_TX) hladnjacke blokove
' PRESKACE: grupisao bi ih po stanici i danu, a lanac drzi 1 blok = 1 lanac. Ovaj
' korak ih vodi kroz AutoLanacHladnjaca, BLOK PO BLOK, po obrascu batch-a:
'
'   povratna vrednost  = broj STVARNO kompletiranih lanaca
'   outGreske          = imenovani razlozi ("blok <broj>: <izvestaj>"), "; "
'
' Pad jednog bloka ne obara ostale: svaki korak lanca je svoja transakcija, pa
' ono sto je proslo JESTE upisano. Pad PROLAZA (citanje tabela) ide gore kao
' greska -- orkestrator ga razlikuje od delimicnog uspeha.
'
' Uzima samo blok koji NIKAD nije bio u otpremnici. Blok koji je storno
' otpremnice oslobodio je odluka operatera (F8 PONISTENJE) i ceka u Oporavku
' kao izgubljen blok -- sync koji bi mu sam napravio nov lanac ponistio bi tu
' odluku. Isto vazi za neizdat blok (filter malina batch-a).
'
' Sistemski pad unutar lanca jednog bloka (pisac, sema) izlazi kao razlog tog
' bloka, ne kao pad prolaza: AutoLanacHladnjaca ne nosi tu razliku (imenovano
' ogranicenje, plan 14.55 "Korak 7", tok koda deo 2).
' ============================================================
Public Function AutoLanacZaNevezaneBlokove(Optional ByRef outGreske As String) As Long
    Const SRC As String = "AutoLanacZaNevezaneBlokove"

    outGreske = ""
    On Error GoTo EH

    Dim nevezani As Object
    Set nevezani = modDokumenta.NevezaniOtkupi()
    If nevezani.count = 0 Then Exit Function

    Dim otk As Variant
    otk = GetTableData(TBL_OTKUP)
    If Not IsArray(otk) Then Exit Function

    Dim cId As Long, cSta As Long, cBr As Long, cIzd As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cIzd = RequireColumnIndex(TBL_OTKUP, COL_TRACE_IZDATO_STATUS, SRC)

    Dim hladne As Object, greske As Collection
    Set hladne = HladnjackeStanice()
    If hladne.count = 0 Then Exit Function
    Set greske = New Collection

    Dim i As Long, oid As String, prj As String
    Dim izv As String, n As Long
    For i = 1 To UBound(otk, 1)
        oid = Trim$(NzToText(otk(i, cId)))
        If Len(oid) > 0 Then
            If nevezani.Exists(UCase$(oid)) Then
                If Len(CStr(nevezani(UCase$(oid)))) = 0 Then
                    If modDokumenta.IzdatoStatusJeIzdato(otk(i, cIzd)) Then
                        If hladne.Exists(UCase$(Trim$(NzToText(otk(i, cSta))))) Then
                            izv = ""
                            prj = AutoLanacHladnjaca(oid, izv)
                            If Len(prj) > 0 Then
                                n = n + 1
                            ElseIf Len(izv) > 0 Then
                                greske.Add Poruka("OTKUI_LANAC_BLOK") & " " & _
                                           Trim$(NzToText(otk(i, cBr))) & ": " & izv
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    AutoLanacZaNevezaneBlokove = n
    outGreske = SpojiRazlogeLanca(greske)
    If n > 0 Then LogInfo SRC, "Hladnjacki lanac kompletiran za blokova: " & CStr(n)
    If Len(outGreske) > 0 Then LogWarn SRC, "Blokovi bez lanca: " & outGreske
    Exit Function

EH:
    ' Broj, opis i izvor PRE LogErr-a -- LogErr usput brise stanje greske.
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    LogErr SRC
    AutoLanacZaNevezaneBlokove = 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Private Function SpojiRazlogeLanca(ByVal c As Collection) As String
    Dim i As Long, s As String
    For i = 1 To c.count
        If Len(s) > 0 Then s = s & "; "
        s = s & CStr(c(i))
    Next i
    SpojiRazlogeLanca = s
End Function

' ============================================================
' NEDOVRSEN LANAC ZA OPORAVAK (S6 korak 7, odluka 8).
'
' Aktivan blok na hladnjackoj stanici koji JE u aktivnoj otpremnici, a otpremnica
' nema aktivnu zbirnu ili zbirna nema aktivnu prijemnicu. Takav blok ne vidi
' nijedna druga lista: "Bez otpremnice" cita nevezane blokove, a ovaj je vezan.
' Blok BEZ otpremnice nije ovde -- on je u F1 "Bez otpremnice", sa radnjom
' "Ponovi".
'
' Vraca Collection od Array(OtkupID, BrojDokumenta, OtpremnicaID, ZbirnaID).
' Clanstvo i prijemnice se citaju JEDNOM, kroz recnike -- lista se cita na svaku
' promenu podataka (brojac menija), pa citanje po bloku ne dolazi u obzir.
' Citaoci su strogi i greska izlazi gore: pozivalac je prikazuje kao red.
' ============================================================
Public Function NedovrseniLanci() As Collection
    Const SRC As String = "NedovrseniLanci"
    Dim res As Collection
    Set res = New Collection
    Set NedovrseniLanci = res

    Dim hladne As Object
    Set hladne = HladnjackeStanice()
    If hladne.count = 0 Then Exit Function

    Dim otk As Variant
    otk = GetTableData(TBL_OTKUP)
    If Not IsArray(otk) Then Exit Function

    Dim cId As Long, cSta As Long, cBr As Long, cSto As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cSto = RequireColumnIndex(TBL_OTKUP, COL_STORNIRANO, SRC)

    Dim clanOtp As Object, clanZbr As Object, prjPoZbr As Object
    Set clanOtp = modDokumenta.AktivnoClanstvoOtpremnica()
    Set clanZbr = modDokumenta.AktivnoClanstvoZbirni()
    Set prjPoZbr = AktivnePrijemnicePoZbirni()

    Dim i As Long, oid As String, otpID As String, zbrID As String
    For i = 1 To UBound(otk, 1)
        oid = Trim$(NzToText(otk(i, cId)))
        If Len(oid) > 0 Then
            If StrComp(Trim$(NzToText(otk(i, cSto))), "Da", vbTextCompare) <> 0 Then
                If hladne.Exists(UCase$(Trim$(NzToText(otk(i, cSta))))) Then
                    If clanOtp.Exists(UCase$(oid)) Then
                        otpID = CStr(clanOtp(UCase$(oid)))
                        zbrID = ""
                        If clanZbr.Exists(UCase$(otpID)) Then zbrID = CStr(clanZbr(UCase$(otpID)))
                        If Len(zbrID) = 0 Then
                            res.Add Array(oid, Trim$(NzToText(otk(i, cBr))), otpID, "")
                        ElseIf Not prjPoZbr.Exists(UCase$(zbrID)) Then
                            res.Add Array(oid, Trim$(NzToText(otk(i, cBr))), otpID, zbrID)
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' Broj otpremnice / zbirne za opis reda Oporavka.
Public Function LanacBrojDokumenta(ByVal tip As String, ByVal id As String) As String
    On Error Resume Next
    If Len(id) = 0 Then Exit Function
    Select Case tip
        Case "OTP": LanacBrojDokumenta = Trim$(nz(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, id, COL_OTP_BROJ), ""))
        Case "ZBR": LanacBrojDokumenta = Trim$(nz(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, id, COL_ZBR_BROJ), ""))
    End Select
End Function

' UCase(StanicaID) svih stanica oznacenih kao hladnjaca -- za citaoce koji pitaju
' za MNOGO blokova odjednom (sync korak, malina batch, Oporavak); HladnjacaStrogo
' po bloku bi za svaki prosao celu tabelu stanica. Kolona JeHladnjaca je obavezna
' (RequireColumnIndex): bez nje se ne zna nista.
'
' Stanica koja ne postoji NIJE u skupu: njen blok ide malina batch-u, a pisac
' otpremnice ga tamo odbija po imenu stanice (Test_OTP_AutoKvarGrupeNeObaraOstale).
' Stanica upisana dvaput sa "Da" JESTE u skupu, pa njen blok ide u lanac, gde ga
' HladnjacaStrogo odbija imenovanim razlogom -- ne zavrsava u malina otpremnici.
Public Function HladnjackeStanice() As Object
    Const SRC As String = "HladnjackeStanice"
    Dim rez As Object, d As Variant, i As Long, cId As Long, cH As Long
    Set rez = CreateObject("Scripting.Dictionary")
    Set HladnjackeStanice = rez

    d = GetTableData(TBL_STANICE)
    If Not IsArray(d) Then Exit Function
    cId = RequireColumnIndex(TBL_STANICE, "StanicaID", SRC)
    cH = RequireColumnIndex(TBL_STANICE, COL_STA_JE_HLADNJACA, SRC)
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cH))), "Da", vbTextCompare) = 0 Then
            rez(UCase$(Trim$(NzToText(d(i, cId))))) = True
        End If
    Next i
End Function
