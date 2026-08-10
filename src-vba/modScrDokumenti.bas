Attribute VB_Name = "modScrDokumenti"
'=====================================================================
' modScrDokumenti - OPIS ekrana "Unos dokumenata" (F1..F8), faza S2b.
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

Public Const SCRDOK_BUILD As String = "v6-ui-71"

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
    Scr_Meta = "kljuc=DOKUMENTI|naslov=OTKUI_NAV_UNOS|oblik=forma+mreza|rezima=8"
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
        Case "F8": ModeIco = IC_STORNO      ' Undo       - storno
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

Public Function GridCols(ByVal mk As String) As Variant
    Dim c As Collection: Set c = New Collection
    c.Add "OTKUI_HD_BROJ|" & ColBroj(mk) & "|txt|110|1"
    c.Add "OTKUI_HD_DATUM|" & ColDatum(mk) & "|date|58|1"
    c.Add "OTKUI_HD_PARTNER|" & ColPartner(mk) & "|part|0|1"

    Select Case mk
        Case "OTKUP", "OTPREMNICA", "ZBIRNA", "PRIJEMNICA", "STORNO"
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
    End Select

    c.Add "OTKUI_HD_STATUS||pill|88|1"

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
        Case "OTPREMNICA", "STORNO":    ColBroj = COL_OTP_BROJ
        Case "ZBIRNA":                  ColBroj = COL_ZBR_BROJ
        Case "PRIJEMNICA":              ColBroj = COL_PRJ_BROJ
        Case "AMB_ISPLATE", "AMB_UPLATE": ColBroj = COL_NOV_BROJ_DOK
        Case "REVERSI":                 ColBroj = COL_AMB_DOK_ID
    End Select
End Function

Public Function ColDatum(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                   ColDatum = COL_OTK_DATUM
        Case "OTPREMNICA", "STORNO":    ColDatum = COL_OTP_DATUM
        Case "ZBIRNA":                  ColDatum = COL_ZBR_DATUM
        Case "PRIJEMNICA":              ColDatum = COL_PRJ_DATUM
        Case "AMB_ISPLATE", "AMB_UPLATE": ColDatum = COL_NOV_DATUM
        Case "REVERSI":                 ColDatum = COL_AMB_DATUM
    End Select
End Function

Public Function ColPartner(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                   ColPartner = COL_OTK_KOOPERANT
        Case "OTPREMNICA", "STORNO":    ColPartner = COL_OTP_STANICA
        Case "ZBIRNA":                  ColPartner = COL_ZBR_KUPAC
        Case "PRIJEMNICA":              ColPartner = COL_PRJ_KUPAC
        Case "AMB_ISPLATE", "AMB_UPLATE": ColPartner = COL_NOV_PARTNER
        Case "REVERSI":                 ColPartner = COL_AMB_ENTITET
    End Select
End Function

Public Function ColVrsta(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColVrsta = COL_OTK_VRSTA
        Case "OTPREMNICA", "STORNO": ColVrsta = COL_OTP_VRSTA
        Case "ZBIRNA":               ColVrsta = COL_ZBR_VRSTA
        Case "PRIJEMNICA":           ColVrsta = COL_PRJ_VRSTA
    End Select
End Function

Public Function ColSorta(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColSorta = COL_OTK_SORTA
        Case "OTPREMNICA", "STORNO": ColSorta = COL_OTP_SORTA
        Case "ZBIRNA":               ColSorta = COL_ZBR_SORTA
        Case "PRIJEMNICA":           ColSorta = COL_PRJ_SORTA
    End Select
End Function

Public Function ColKlasa(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKlasa = COL_OTK_KLASA
        Case "OTPREMNICA", "STORNO": ColKlasa = COL_OTP_KLASA
        Case "ZBIRNA":               ColKlasa = COL_ZBR_KLASA
        Case "PRIJEMNICA":           ColKlasa = COL_PRJ_KLASA
    End Select
End Function

Public Function ColKolicina(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKolicina = COL_OTK_KOLICINA
        Case "OTPREMNICA", "STORNO": ColKolicina = COL_OTP_KOLICINA
        Case "ZBIRNA":               ColKolicina = COL_ZBR_KOLICINA
        Case "PRIJEMNICA":           ColKolicina = COL_PRJ_KOLICINA
    End Select
End Function

Public Function ColKolAmb(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColKolAmb = COL_OTK_KOL_AMB
        Case "OTPREMNICA", "STORNO": ColKolAmb = COL_OTP_KOL_AMB
        Case "ZBIRNA":               ColKolAmb = COL_ZBR_KOL_AMB
        Case "PRIJEMNICA":           ColKolAmb = COL_PRJ_KOL_AMB
    End Select
End Function

Public Function ColTipAmb(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColTipAmb = COL_OTK_TIP_AMB
        Case "OTPREMNICA", "STORNO": ColTipAmb = COL_OTP_TIP_AMB
        Case "ZBIRNA":               ColTipAmb = COL_ZBR_TIP_AMB
        Case "PRIJEMNICA":           ColTipAmb = COL_PRJ_TIP_AMB
    End Select
End Function

' Prazno = rezim nema cenu (tblZbirna), pa ni kolonu vrednosti.
Public Function ColCena(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColCena = COL_OTK_CENA
        Case "OTPREMNICA", "STORNO": ColCena = COL_OTP_CENA
        Case "PRIJEMNICA":           ColCena = COL_PRJ_CENA
    End Select
End Function

Public Function ColBrojZbirne(ByVal m As String) As String
    Select Case m
        Case "OTKUP":                ColBrojZbirne = COL_OTK_BROJ_ZBIRNE
        Case "OTPREMNICA", "STORNO": ColBrojZbirne = COL_OTP_BROJ_ZBIRNE
        Case "PRIJEMNICA":           ColBrojZbirne = COL_PRJ_BROJ_ZBIRNE
    End Select
End Function

' Polje opisa kolone: 0=kljuc naslova 1=izvorna kolona 2=vrsta 3=sirina 4=prio
Public Function ColF(ByVal spec As String, ByVal idx As Long) As String
    Dim p As Variant: p = Split(spec, "|")
    If idx > UBound(p) Then Exit Function
    ColF = CStr(p(idx))
End Function

' 0=broj 1=datum 2=partner 3=kolicina 4=cena 5=brojZbirne 6=direktna vrednost
Public Function ColumnSpec(ByVal mk As String) As Variant
    Select Case mk
        Case "OTKUP"
            ' partner na otkupu je KOOPERANT (ranije je stajao BrojOtpremnice)
            ColumnSpec = Array(COL_OTK_BR_DOK, COL_OTK_DATUM, COL_OTK_KOOPERANT, _
                               COL_OTK_KOLICINA, COL_OTK_CENA, COL_OTK_BROJ_ZBIRNE, "")
        Case "OTPREMNICA", "STORNO"
            ColumnSpec = Array(COL_OTP_BROJ, COL_OTP_DATUM, COL_OTP_STANICA, _
                               COL_OTP_KOLICINA, COL_OTP_CENA, COL_OTP_BROJ_ZBIRNE, "")
        Case "ZBIRNA"
            ' 5. slot (cena) - tblZbirna NEMA kolonu Cena, pa je VREDNOST 0.
            ' 6. slot prazan: ranije je pokazivao na samu sebe (COL_ZBR_BROJ),
            ' pa je svaka zbirna dobijala status "Poslato".
            ColumnSpec = Array(COL_ZBR_BROJ, COL_ZBR_DATUM, COL_ZBR_KUPAC, _
                               COL_ZBR_KOLICINA, "", "", "")
        Case "PRIJEMNICA"
            ColumnSpec = Array(COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_KUPAC, _
                               COL_PRJ_KOLICINA, COL_PRJ_CENA, COL_PRJ_BROJ_ZBIRNE, "")
        Case "AMB_ISPLATE"
            ' tblNovac nema kolicinu/cenu - vrednost je sam iznos isplate
            ColumnSpec = Array(COL_NOV_BROJ_DOK, COL_NOV_DATUM, COL_NOV_PARTNER, _
                               "", "", "", COL_NOV_ISPLATA)
        Case "AMB_UPLATE"
            ColumnSpec = Array(COL_NOV_BROJ_DOK, COL_NOV_DATUM, COL_NOV_PARTNER, _
                               "", "", "", COL_NOV_UPLATA)
        Case "REVERSI"
            ' broj reversa zivi u DokumentID ("x/ddmmyy" namespace, vidi
            ' modBrojevi.MaxSeqReversAmbalaza) - tblAmbalaza nema BrojDokumenta.
            ' 4. kolona = tip ambalaze (tekst), 5. = kolicina u komadima.
            ColumnSpec = Array(COL_AMB_DOK_ID, COL_AMB_DATUM, COL_AMB_ENTITET, _
                               COL_AMB_TIP, "", "", COL_AMB_KOLICINA)
        Case Else
            ColumnSpec = Array("", "", "", "", "", "", "")
    End Select
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
Public Function ModeTable(ByVal mode As String) As String
    Select Case mode
        Case "F1": ModeTable = TBL_OTKUP
        Case "F2": ModeTable = TBL_OTPREMNICA
        Case "F3": ModeTable = TBL_ZBIRNA
        Case "F4": ModeTable = TBL_PRIJEMNICA
        Case "F5": ModeTable = TBL_NOVAC
        Case "F6": ModeTable = TBL_NOVAC
        Case "F7": ModeTable = TBL_AMBALAZA
        Case "F8": ModeTable = TBL_OTPREMNICA
        Case Else: ModeTable = TBL_OTKUP
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
        Case "F8": modeKey = "STORNO"
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

Public Function ModeHasZbirna(ByVal mode As String) As Boolean
    Select Case mode
        Case "F1", "F2", "F4", "F8": ModeHasZbirna = True
        Case Else:                   ModeHasZbirna = False
    End Select
End Function

Public Function ModeTextCol3(ByVal mode As String) As Boolean
    Select Case mode
        Case "F5", "F6", "F7": ModeTextCol3 = True
    End Select
End Function

' Jedinica 5. kolone i podnozja: dinar za robu i novac, komad za reverse.
Public Function ModeValUnit(ByVal mode As String) As String
    If mode = "F7" Then ModeValUnit = Poruka("OTKUI_UNIT_KOM") Else ModeValUnit = Poruka("OTKUI_UNIT_RSD")
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
