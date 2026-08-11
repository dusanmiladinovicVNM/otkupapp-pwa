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

Public Const SCRDOK_BUILD As String = "v6-ui-83"

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
' (blokovi) su njena raspodela po kooperantima; ekran postoji da operater vidi
' koliko je od otpremnice jos neraspodeljeno. Zato F1 ima tri liste umesto
' jedne, i one se biraju prekidacem u zoni mreze.
'
' Stanje je OVDE, a ne u modOtkupBlok - taj drzi svoje mActiveOtpID vezano za
' frmOtkup, pa bi dva ziva ekrana delila istu promenljivu. Racun bilansa se
' NE duplira: zovu se njegove funkcije (SumKolByOtp i ostale, javne od F1).
Private mLista As String          ' "SVI" | "OTPREMNICE" | "BLOKOVI"
Private mOtpID As String          ' aktivna otpremnica (OtpremnicaID)
Private mOtpBroj As String        ' njen broj - za traku i naslov liste
' broj otpremnice -> OtpremnicaID; puni ga RowsOtpremnice u istom prolazu.
' Mreza prikazuje broj (to je ono sto operater vidi), a ekranu treba ID.
Private mOtpIds As Object

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

' Opis aktivne otpremnice za traku iznad forme. Prazno = nema izabrane.
' Oblik: broj | kupac | datum | ukupnoKg | uBlokKg | ostatakKg |
'        ukupnoAmb | uBlokAmb | ostatakAmb | cena
Public Function Scr_OtpInfo() As String
    Dim ukKg As Double, blKg As Double, ukAmb As Double, blAmb As Double
    Dim kupac As String, dat As String, cena As Double
    On Error Resume Next
    If Len(mOtpID) = 0 Then Exit Function

    ukKg = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mOtpID, COL_OTP_KOLICINA))
    ukAmb = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mOtpID, COL_OTP_KOL_AMB))
    kupac = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mOtpID, COL_OTP_STANICA))
    dat = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mOtpID, COL_OTP_DATUM))
    If IsDate(dat) Then dat = Format$(CDate(dat), "dd.mm.yyyy.")
    blKg = modOtkupBlok.SumKolByOtp(mOtpID)
    blAmb = modOtkupBlok.SumAmbByOtp(mOtpID)
    cena = modOtkupBlok.ExistingBlokCena(mOtpID)

    Scr_OtpInfo = mOtpBroj & "|" & OtpKupacNaziv(kupac) & "|" & dat & "|" & _
                  ukKg & "|" & blKg & "|" & (ukKg - blKg) & "|" & _
                  ukAmb & "|" & blAmb & "|" & (ukAmb - blAmb) & "|" & cena
End Function

' Broj iz Variant-a. modOtkupBlok ima svoj NumVal ali je Private; ovde je
' jeftinije imati tri linije nego otvarati jos jedan simbol u produkcionom
' modulu samo zbog konverzije.
Private Function NumVal(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumVal = CDbl(v)
End Function

Private Function OtpKupacNaziv(ByVal stanicaID As String) As String
    OtpKupacNaziv = stanicaID
    If Len(stanicaID) = 0 Then Exit Function
    On Error Resume Next
    OtpKupacNaziv = NzToText(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"))
    If Len(OtpKupacNaziv) = 0 Then OtpKupacNaziv = stanicaID
End Function

' Broj otpremnice - stoji u naslovu liste blokova.
Public Function Scr_OtpBroj() As String
    Scr_OtpBroj = mOtpBroj
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
    Scr_Meta = "kljuc=DOKUMENTI|naslov=OTKUI_NAV_UNOS|oblik=forma+mreza|rezima=8"
End Function

' Radnje ovog ekrana. Ljuska ne zna nijednu - prosledjuje tag i, ako je ekran
' vratio True, osvezi mrezu i traku.
'   lsSVI / lsOTPREMNICE / lsBLOKOVI - prekidac liste u F1
'   row:<n>                          - izabran red; u listi otpremnica to
'                                      BIRA aktivnu otpremnicu
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim broj As String
    On Error Resume Next
    If modeKey(ActiveMode) <> "OTKUP" Then Exit Function

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        Scr_Event = True
        Exit Function
    End If

    If Left$(tag, 4) = "row:" And Scr_Lista() = "OTPREMNICE" Then
        broj = CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 1))
        If Len(broj) = 0 Then Exit Function
        If mOtpIds Is Nothing Then Exit Function
        If Not mOtpIds.Exists(broj) Then Exit Function
        mOtpID = CStr(mOtpIds(broj))
        mOtpBroj = broj
        ' izbor otpremnice vodi pravo na njene blokove - to je sledeci potez
        mLista = "BLOKOVI"
        Scr_Event = True
    End If
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

' Ugovor: Array(kolone, redovi, n, zbirKg, zbirVal, brojaciCipova).
' Ovo je bivsi modOtkupUI.FillGrid. Do S4b je pisao PRAVO u stanje ljuske
' (mView, mViewN, mSumKg, mCnt*), pa je mreza mogla da sluzi samo njega.
' Sada vraca - i mreza je time postala neutralna. Sortiranje radi ljuska.
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' F1 ima tri liste. Dve su svoje - otpremnice kao izvor i blokovi aktivne
    ' otpremnice; treca je zatecena lista dokumenata, ista za svih osam rezima.
    Select Case Scr_Lista()
        Case "OTPREMNICE": Scr_Rows = RowsOtpremnice(q): Exit Function
        Case "BLOKOVI":    Scr_Rows = RowsBlokovi(q): Exit Function
    End Select
    Scr_Rows = RowsDokumenti(filter, q)
End Function

Private Function RowsDokumenti(ByVal filter As String, ByVal q As String) As Variant
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

    On Error GoTo EH
    mStep = "start"
    tblName = ModeTable(ActiveMode)
    ' suzavanje iz panela "Filteri" i danasnji datum drzi ljuska - ekran ih
    ' cita, ne pamti
    fltV = modOtkupUI.FltVrsta()
    fltP = modOtkupUI.FltPart()
    mToday = Int(Now)
    mMonthStart = CDbl(DateSerial(Year(Now), Month(Now), 1))
    If Len(tblName) = 0 Then Exit Function

    src = modUiData.CachedTable(tblName)
    If Not IsArray(src) Then Exit Function

    mStep = "GridCols"
    mk = modeKey(ActiveMode)
    ' Druga brana za F8: filter se ne sme izgubiti ni ako neko dodje ovde
    ' zaobilazeci cipove. MatchFilterFast svaki filter osim "otkazane"
    ' ogranicava na Not isStorno - bez ovoga ekran "Stornirani dokumenti"
    ' prikazuje upravo AKTIVNE dokumente.
    If mk = "STORNO" Then filter = "otkazane"
    cols = GridCols(mk)
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
    Dim iPayID As Long, iPayKol As Long, iPayCena As Long
    mStep = "placanje"
    pay = (mk = "OTKUP" Or mk = "PRIJEMNICA")
    If pay Then
        If mk = "OTKUP" Then
            Set dPay = modNovac.BuildIsplataDictByOtkup()
            iPayID = ColIdx(tblName, COL_OTK_ID)
            iPayKol = ColIdx(tblName, COL_OTK_KOLICINA)
            iPayCena = ColIdx(tblName, COL_OTK_CENA)
        Else
            Set dPay = modNovac.BuildUplataDictByFaktura()
            Set dFakIzn = FakturaIznosMap()
            iPayID = ColIdx(tblName, COL_PRJ_FAKTURA_ID)
        End If
        If dPay Is Nothing Then pay = False
    End If

    Dim pl As Variant, pmap As Object
    pl = PartnerLookup(ActiveMode)
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
        If iKg >= 0 Then vKgRow = CellD(src, r, ix(iKg))

        Dim pCode As Long, pRest As Double, duguje As Double, placeno As Double
        Dim payKey As String
        pCode = 0: pRest = 0
        If pay Then
            duguje = 0: placeno = 0
            payKey = CellS(src, r, iPayID)
            If mk = "OTKUP" Then
                duguje = CellD(src, r, iPayKol) * CellD(src, r, iPayCena)
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
    RowsDokumenti = Array(cols, outA, n, sumKg, sumVal, Array(cOtk, cBez, cNef))
    Exit Function
EH:
    ' greska se NE guta - ReloadGrid je prijavljuje sa imenom koraka
    Err.Raise Err.Number, "modScrDokumenti.Scr_Rows[" & mStep & "]", Err.description
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
        Case "F2", "F8":  PartnerLookup = Array(TBL_STANICE, "StanicaID", "Naziv", "")
        Case "F3", "F4":  PartnerLookup = Array(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
        Case "F1":        PartnerLookup = Array(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Case Else:        PartnerLookup = Array("", "", "", "")
                          ' F5/F6: tblNovac vec ima tekstualnu kolonu Partner
    End Select
End Function

' Ljuska ovo zove kad se podaci promene (RefreshFromData) - ekran mora da
' zaboravi svoje izvedene mape, inace bi posle upisa racunao po starom.
Public Sub Scr_ResetCache()
    Set mMape = Nothing
End Sub

'------------------------------------------------- LISTA: OTPREMNICE (F1)
' Otpremnice kao IZVOR robe. Kljucna kolona je OSTATAK - koliko od otpremnice
' jos nije raspodeljeno u blokove. Zbir po otpremnici se racuna JEDNIM
' prolazom kroz tblOtkup (modOtkupBlok.BuildNapisanoByOtp), ne pozivom po
' redu: 867 otpremnica puta 1625 otkupa bi bilo milion i po poredjenja.
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
        "OTKUI_HD_KOL_AMB||num|54|3")
End Function

Private Function RowsOtpremnice(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, d As Object, stan As Object
    Dim iID As Long, iBroj As Long, iDat As Long, iSt As Long, iVr As Long
    Dim iSo As Long, iKol As Long, iAmb As Long, iStorno As Long
    Dim otpID As String, ukKg As Double, blKg As Double, hay As String
    Dim sumOst As Double
    On Error GoTo EH
    mStep = "otpremnice"

    src = modUiData.CachedTable(TBL_OTPREMNICA)
    If Not IsArray(src) Then Exit Function
    Set d = modOtkupBlok.BuildNapisanoByOtp()
    Set mOtpIds = CreateObject("Scripting.Dictionary")
    Set stan = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")

    iID = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_ID)
    iBroj = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_BROJ)
    iDat = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_DATUM)
    iSt = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_STANICA)
    iVr = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_VRSTA)
    iSo = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_SORTA)
    iKol = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    iAmb = modUiData.ColIdx(TBL_OTPREMNICA, COL_OTP_KOL_AMB)
    iStorno = modUiData.ColIdx(TBL_OTPREMNICA, COL_STORNIRANO)

    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To 9)
    For r = 1 To nRows
        If iStorno > 0 Then
            If UCase$(modUiData.CellS(src, r, iStorno)) = "DA" Then GoTo Sledeca
        End If
        otpID = modUiData.CellS(src, r, iID)
        ukKg = modUiData.CellD(src, r, iKol)
        blKg = 0
        If d.Exists(otpID) Then blKg = CDbl(d(otpID))

        hay = modUiData.CellS(src, r, iBroj) & "|" & modUiData.CellS(src, r, iVr) & _
              "|" & modUiData.CellS(src, r, iSo)
        Dim kup As String
        kup = modUiData.CellS(src, r, iSt)
        If Not stan Is Nothing Then
            If stan.Exists(kup) Then kup = CStr(stan(kup))
        End If
        hay = hay & "|" & kup
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeca
        End If

        n = n + 1
        outA(n, 1) = modUiData.CellS(src, r, iBroj)
        mOtpIds(CStr(outA(n, 1))) = otpID
        outA(n, 2) = modUiData.CellDate(src, r, iDat)
        outA(n, 3) = kup
        outA(n, 4) = modUiData.CellS(src, r, iVr)
        outA(n, 5) = modUiData.CellS(src, r, iSo)
        outA(n, 6) = ukKg
        outA(n, 7) = blKg
        outA(n, 8) = ukKg - blKg
        outA(n, 9) = modUiData.CellD(src, r, iAmb)
        sumOst = sumOst + (ukKg - blKg)
Sledeca:
    Next r

    mStep = "OK"
    RowsOtpremnice = Array(OtpGridCols(), outA, n, sumOst, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsOtpremnice[" & mStep & "]", Err.description
End Function

'---------------------------------------------------- LISTA: BLOKOVI (F1)
' Otkupni listovi AKTIVNE otpremnice. Bez izabrane otpremnice lista je prazna
' - to je tacno, ne greska.
Private Function BlokGridCols() As Variant
    BlokGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_KG||kg|66|1", _
        "OTKUI_HD_KOL_AMB||num|54|2", _
        "OTKUI_HD_CENA||num|70|2", _
        "OTKUI_HD_VREDNOST||mult|96|1")
End Function

Private Function RowsBlokovi(ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, nRows As Long
    Dim outA() As Variant, koop As Object
    Dim iOtp As Long, iBroj As Long, iDat As Long, iKoop As Long
    Dim iKol As Long, iAmb As Long, iCena As Long, iStorno As Long
    Dim kg As Double, cena As Double, hay As String
    Dim sumKg As Double, sumVal As Double, ime As String
    On Error GoTo EH
    mStep = "blokovi"

    If Len(mOtpID) = 0 Then
        RowsBlokovi = Array(BlokGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    src = modUiData.CachedTable(TBL_OTKUP)
    If Not IsArray(src) Then Exit Function
    Set koop = PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")

    iOtp = modUiData.ColIdx(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    iBroj = modUiData.ColIdx(TBL_OTKUP, COL_OTK_BR_DOK)
    iDat = modUiData.ColIdx(TBL_OTKUP, COL_OTK_DATUM)
    iKoop = modUiData.ColIdx(TBL_OTKUP, COL_OTK_KOOPERANT)
    iKol = modUiData.ColIdx(TBL_OTKUP, COL_OTK_KOLICINA)
    iAmb = modUiData.ColIdx(TBL_OTKUP, COL_OTK_KOL_AMB)
    iCena = modUiData.ColIdx(TBL_OTKUP, COL_OTK_CENA)
    iStorno = modUiData.ColIdx(TBL_OTKUP, COL_STORNIRANO)

    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To 7)
    For r = 1 To nRows
        If modUiData.CellS(src, r, iOtp) <> mOtpID Then GoTo Sledeci
        If iStorno > 0 Then
            If UCase$(modUiData.CellS(src, r, iStorno)) = "DA" Then GoTo Sledeci
        End If
        ime = modUiData.CellS(src, r, iKoop)
        If Not koop Is Nothing Then
            If koop.Exists(ime) Then ime = CStr(koop(ime))
        End If
        hay = modUiData.CellS(src, r, iBroj) & "|" & ime
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If

        kg = modUiData.CellD(src, r, iKol)
        cena = modUiData.CellD(src, r, iCena)
        n = n + 1
        outA(n, 1) = modUiData.CellS(src, r, iBroj)
        outA(n, 2) = modUiData.CellDate(src, r, iDat)
        outA(n, 3) = ime
        outA(n, 4) = kg
        outA(n, 5) = modUiData.CellD(src, r, iAmb)
        outA(n, 6) = cena
        outA(n, 7) = kg * cena
        sumKg = sumKg + kg
        sumVal = sumVal + kg * cena
Sledeci:
    Next r

    mStep = "OK"
    RowsBlokovi = Array(BlokGridCols(), outA, n, sumKg, sumVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrDokumenti.RowsBlokovi[" & mStep & "]", Err.description
End Function
