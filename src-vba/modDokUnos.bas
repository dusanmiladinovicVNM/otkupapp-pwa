Attribute VB_Name = "modDokUnos"
'=====================================================================
' modDokUnos - UNOS DOKUMENATA (otpremnica, zbirna, prijemnica), bez
' ijedne kontrole.
'
' Isti razlog i isti oblik kao modOtkupUnos, samo za rezime iz
' frmDokumenta. Poslovni posao (provere, bruto->neto, upis, ono sto ide
' posle upisa) ne sme da zivi u formi, jer ga onda drugi ekran ne moze
' pozvati bez prepisivanja.
'
'   OtpremnicaValidiraj(p, fokus)  provere + bruto->neto; vraca poruku o
'                                  gresci ("" = proslo) i LOGICKO ime
'                                  polja na koje treba vratiti fokus
'   OtpremnicaUpisi(p, poruke)     CreateOtpremnicaDraft_TX + zavrsetak
'                                  ispravke; vraca OtpremnicaID (prazno =
'                                  nije upisano)
'   ZbirnaValidiraj(p, fokus)      provere polja i broja; zbirna NEMA
'                                  bruto->neto pa recnik ostaje netaknut
'   ZbirnaUpisi(p, poruke)         CreateZbirnaDraft_TX; vraca ZbirnaID
'   ZbirnaIzmeniNacrt(id, p, ...)  UpdateZbirnaDraft_TX nad postojecim nacrtom
'   PrijemnicaValidiraj / PrijemnicaUpisi  isto za F4 (SavePrijemnicaMulti_TX)
'
' Ulaz je RECNIK sa LOGICKIM imenima polja (NoviOtpremnicaUnos):
'
'   datum, stanicaID, vozacID, brDok, brojZbirne, vrsta, sorta, tipAmb
'   kolicinaI, cenaI, kolAmb
'   dveKlase, kolicinaII, cenaII, kolAmbII
'
' OtpremnicaValidiraj UPISUJE nazad: kolicinaI/kolicinaII postaju NETO, a
' brutoKgI/brutoKgII zamrznuti uneti bruto (kad je OTKUP_BRUTO_UNOS ON).
'
' Od S3a se taj uneti bruto NE upisuje nigde, i to je odluka a ne propust:
' BrutoKg na stavci otpremnice je bruto IZVORA, koji izdavanje sabira iz
' otkupnih blokova. Uneti bruto je ulaz u racun (bruto - tara = neto), pa bi u
' istoj koloni bile dve razlicite cinjenice pod istim imenom.
'
' RAZLIKE U ODNOSU NA OTKUPNI LIST (nisu greske - tako je u legacy):
'   - vozac je OBAVEZAN (otkupni list ga ne trazi)
'   - vrsta i sorta su obavezne SAMO uz VALIDACIJA_UNOSA
'   - cena I je obavezna samo uz VALIDACIJA_UNOSA; inace sme i 0
'   - gajbe i tip ambalaze su obavezni samo uz VALIDACIJA_UNOSA
'   - nema parcele, nema izdate ambalaze, nema proseka po gajbici
'
' RAZLIKE MEDJU REZIMIMA (sve tri iz legacy frmDokumenta, ne izmisljene):
'   - ZBIRNA nema cenu i NEMA bruto->neto: tblZbirna nema ni kolonu Cena ni
'     BrutoKg. Zbirna je zbir SVOJIH otpremnica, a one su vec u netu, pa bi
'     oduzimanje tare i drugi put spustilo kilograme. Zato ZbirnaValidiraj
'     recnik NE menja.
'   - ZBIRNA SE VISE NE POREDI SA IZVOROM PRI UNOSU (S4-2c/2b-1). Do ovog reza
'     je validator sabirao otpremnice po Otkup/Otpremnica.BrojZbirne i trazio da
'     se zbir poklopi. Ta veza je ukinuta (ZBR-KANON-01): clanstvo je zapis u
'     tblZbirnaIzvori, a pokrivenost meri IZDAVANJE (IzdajZbirnu_TX), nad tacno
'     onim izvorima koje je operater vezao. Unos zato prijavljuje NAJAVU.
'   - PRIJEMNICA ima i cenu i bruto->neto (tblPrijemnica ima BrutoKg), broj
'     zbirne joj je obavezan, a zbirna mora da postoji u sistemu.
'
' VAZNO: legacy frmDokumenta OSTAJE netaknut i potpuno operativan. Ovaj
' modul je drugi pozivalac istih poslovnih rutina, ne zamena za formu.
' Dok oba sistema ne budu potpuna, obe kopije postoje namerno.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const DOKUNOS_BUILD As String = "v6-ui-135"

'--------------------------------------------------------------- ULAZ
Public Function NoviOtpremnicaUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("stanicaID") = ""
    p("vozacID") = ""
    p("brDok") = ""
    p("brojZbirne") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("tipAmb") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("brutoKgI") = 0#
    p("brutoKgII") = 0#
    ' OtpremnicaID nacrta koji se MENJA (izmena u F2); prazno = nov nacrt.
    ' Provera broja tada izuzima taj red -- isto kao pisac (OtpIzmeniDraft).
    p("izmenaOtpID") = ""
    Set NoviOtpremnicaUnos = p
End Function

Private Function S(ByVal p As Object, ByVal k As String) As String
    On Error Resume Next
    If p.Exists(k) Then S = Trim$(CStr(p(k)))
End Function

Private Function D(ByVal p As Object, ByVal k As String) As Double
    On Error Resume Next
    If p.Exists(k) Then
        If IsNumeric(p(k)) Then D = CDbl(p(k))
    End If
End Function

Private Function L(ByVal p As Object, ByVal k As String) As Long
    On Error Resume Next
    If p.Exists(k) Then
        If IsNumeric(p(k)) Then L = CLng(p(k))
    End If
End Function

Private Function B(ByVal p As Object, ByVal k As String) As Boolean
    On Error Resume Next
    If p.Exists(k) Then B = CBool(p(k))
End Function

'---------------------------------------------------------- PROVERE
' Vraca "" kad je sve u redu; inace poruku za operatera. Redosled provera
' je isti kao u frmDokumenta.btnUnosOtp_Click - to nije stil nego
' ponasanje: operater je navikao koje ga polje prvo zaustavi.
Public Function OtpremnicaValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim kolI As Double, cenI As Double, kolII As Double, cenII As Double
    Dim kolAmb As Long, kolAmbII As Long
    Dim imaKlasaI As Boolean, dveKl As Boolean, strogo As Boolean
    Dim tara As Double, taraII As Double
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    If Len(S(p, "stanicaID")) = 0 Then
        fokus = "stanicaID": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_OM"): Exit Function
    End If
    ' Otpremnica bez vozaca ne postoji - roba nekim putem ide sa otkupnog mesta.
    If Len(S(p, "vozacID")) = 0 Then
        fokus = "vozacID": OtpremnicaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If
    If strogo And Len(S(p, "sorta")) = 0 Then
        fokus = "sorta": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_SORTA"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    cenI = D(p, "cenaI")
    kolII = D(p, "kolicinaII")
    cenII = D(p, "cenaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    dveKl = B(p, "dveKlase")
    imaKlasaI = (kolI > 0)

    ' Klasa I je opciona SAMO uz ukljucenu Klasu II (unosi se samo II klasa).
    ' Tada ambalaza I mora ostati prazna.
    If Not imaKlasaI Then
        If Not dveKl Then
            fokus = "kolicinaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_KOLICINA"): Exit Function
        End If
        If kolAmb > 0 Then
            fokus = "kolAmb": OtpremnicaValidiraj = Poruka("DOK_MSG_UNOSI_SAMO_KLASA"): Exit Function
        End If
    End If

    ' Cena I: obavezna samo uz strogu validaciju. Van nje otpremnica sme da ode
    ' bez cene (cena stize sa cenovnikom ili kasnije), ali ne sme biti negativna.
    If strogo And imaKlasaI Then
        If cenI <= 0 Then
            fokus = "cenaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
        End If
    ElseIf cenI < 0 Then
        fokus = "cenaI": OtpremnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
    End If

    If dveKl Then
        If kolII <= 0 Then
            fokus = "kolicinaII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_KOLICINA_II"): Exit Function
        End If
        If cenII <= 0 Then
            fokus = "cenaII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_CENA_II"): Exit Function
        End If
    End If

    ' Gajbe i tip ambalaze su obavezni samo uz strogu validaciju - drugacije nego
    ' kod otkupnog lista, gde ih bruto rezim trazi i bez nje.
    If strogo Then
        If imaKlasaI And kolAmb <= 0 Then
            fokus = "kolAmb": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_I"): Exit Function
        End If
        If dveKl And kolAmbII <= 0 Then
            fokus = "kolAmbII": OtpremnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_II"): Exit Function
        End If
        If (kolAmb > 0 Or kolAmbII > 0) And Len(S(p, "tipAmb")) = 0 Then
            fokus = "tipAmb": OtpremnicaValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
        End If
    End If

    ' --- BRUTO -> NETO. Otpremnica se cuva u NETO, isto kao otkupni list, da
    ' panel blokova poredi neto sa neto. Uneti bruto se zamrzava u BrutoKg. ---
    If OtkupBrutoUnos() And kolAmb > 0 Then
        tara = kolAmb * GetTezinaGajbice(S(p, "tipAmb"))
        If tara <= 0 Then
            fokus = "tipAmb"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If tara >= kolI Then
            fokus = "kolicinaI"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _
                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgI") = kolI
        kolI = kolI - tara
        p("kolicinaI") = kolI
    End If

    If dveKl And OtkupBrutoUnos() And kolAmbII > 0 Then
        taraII = kolAmbII * GetTezinaGajbice(S(p, "tipAmb"))
        If taraII <= 0 Then
            fokus = "tipAmb"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If taraII >= kolII Then
            fokus = "kolicinaII"
            OtpremnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE_KLASE") & Format$(taraII, "#,##0.00") & _
                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgII") = kolII
        kolII = kolII - taraII
        p("kolicinaII") = kolII
    End If

    If strogo And Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": OtpremnicaValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If

    ' Dupli broj: ISTA provera koju drzi pisac (modBrojevi.BrojZauzetUNizu), po
    ' nizu (stanica, dan) i sa storniranima. Zatecena je isla kroz CheckDuplicate
    ' -- cela tabela, sirovo poredjenje, bez storniranih -- pa je odbijala isti
    ' broj na drugoj stanici, a pustala broj stornirane otpremnice.
    '
    ' Izmena nacrta izuzima SAMO svoj red (po ID-u), kao i pisac: bez toga je
    ' nacrt odbijao sopstveni broj, a sa sirim izuzimanjem bi preuzeo tudj.
    If Len(S(p, "brDok")) > 0 Then
        Dim zauzeo As String
        zauzeo = modBrojevi.BrojZauzetUNizu(modBrojevi.KIND_OTP, S(p, "stanicaID"), _
                                            CDate(p("datum")), S(p, "brDok"), _
                                            S(p, "izmenaOtpID"))
        If Len(zauzeo) > 0 Then
            fokus = "brDok"
            OtpremnicaValidiraj = Poruka("DOKUNOS_ERR_BROJ_ZAUZET") & " " & zauzeo
            Exit Function
        End If
    End If
    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error Resume
    ' Next, koji cisti Err - operater bi inace dobio poruku bez objasnjenja.
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaValidiraj"
    OtpremnicaValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Ocekivana stavka otpremnice: sta operater PRIJAVLJUJE da ce otpremnica nositi.
' Spisak kljuceva drzi pisac (modDokumenta.OtpOcekKljucPoznat) -- ovde se samo
' popunjava. PredlogCena se salje samo kad postoji: prazno polje je "cena jos
' nije dogovorena", a nula bi u prefillu otkupa bila tvrdnja da je cena 0.
Private Function OtpStavkaDTO(ByVal klasa As String, ByVal kol As Double, _
                              ByVal amb As Double, ByVal predlogCena As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "KolAmbalaze", amb
    If predlogCena > 0 Then s.Add "PredlogCena", predlogCena
    Set OtpStavkaDTO = s
End Function

'------------------------------------------------------------- UPIS
' Otvara otpremnicu kao NACRT i radi sve sto ide uz nju. Vraca OtpremnicaID;
' prazno znaci da upis nije uspeo. U "poruke" se skupljaju napomene koje
' pozivalac prikazuje posle uspeha.
'
' ZASTO NACRT, A NE GOTOV DOKUMENT (S3a, odluka S14.8 t. 4): otpremnica je
' isporuka otkupnih blokova, pa dok se blokovi ne vezu ona jos ne postoji kao
' dokument -- postoji kao NAJAVA sta ce nositi. Izdavanje (IzdajOtpremnicu_TX)
' tek tada poredi najavljeno sa vezanim i knjizi ambalazu.
'
' Vraca ID, ne broj: broj je labela jedinstvena tek po (otkupno mesto, dan), a
' mutacija (izdavanje, izmena nacrta) ide po ID-u. Ekran operateru i dalje
' pokazuje broj -- to je prikaz, ne identitet.
Public Function OtpremnicaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, greska As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    Dim h As Object, ocek As Collection
    If Not OtpremnicaNacrtIzUnosa(p, h, ocek, poruke) Then Exit Function

    res = CreateOtpremnicaDraft_TX(h, ocek, greska)

    If Len(res) = 0 Then
        poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & greska
        Exit Function
    End If

    poruke = poruke & Poruka("DOKUNOS_MSG_OTP_NACRT") & vbCrLf

    ' MALINA AUTO-ZBIRNA JE PAUZIRANA (S3a), i to glasno.
    '
    ' AutoCreateZbirnaFromOtpremnice cita Kolicina / Klasa / KolAmbalaze sa
    ' ZAGLAVLJA otpremnice i ne gleda IzdatoStatus. Nad nacrtom bi napravila
    ' zbirnu sa 0 kg, i to od dokumenta koji jos nije isporuka. Zbirna prelazi na
    ' nov model u S4 -- do tada je poziv ugasen, umesto da se upise polovicna
    ' veza. Isti postupak kao sa hladnjackim lancem u S1.
    If IsMalinaMode() Then
        poruke = poruke & Poruka("DOKUNOS_MSG_ZBIRNA_PAUZIRANA") & vbCrLf
    End If

    ' ISPRAVKA OTPREMNICE JE PAUZIRANA (S3a), ne prevedena.
    '
    ' Ovde je do S3a stajalo ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA. Taj tok nije
    ' zatvaranje konteksta nego pisac STAROG modela (obrisan u S3c)
    ' preko GetBlokOtkupIDs zove ReassignOtkupToOtpremnica_TX, koji upisuje
    ' Otkup.OtpremnicaID i BrojZbirne, pa rekalkulise ili stornira zbirnu.
    '
    ' Pustiti ga nad upravo otvorenim NACRTOM znacilo bi dve stvari, obe lose:
    '   - nov model bi se vezivao starom vezom (Otkup.OtpremnicaID), a to je
    '     tacno most koji se po pravilu refaktora ne pravi;
    '   - dokument koji jos NEMA nijedan izvor i nije IZDATO bio bi proglasen
    '     zamenom izdate otpremnice, a correction kontekst zatvoren. Zamena sme
    '     da bude gotova tek posle: clanstvo -> ocekivano = povezano -> IZDATO.
    '
    ' Sposobnost se vraca u S3c, nad kanonskom vezom (tblOtpremnicaIzvori i
    ' IspravkaOd/ZamenjenSa po ID-u). Do tada operater mora da ZNA da kontekst
    ' stoji otvoren -- inace bi mislio da je ispravka zavrsena.
    If modStornoContext.CountPendingCorrectionsByDocType(FLOW_DOC_OTPREMNICA, _
                                                         SV_MODE_ISPRAVKA) > 0 Then
        poruke = poruke & Poruka("DOKUNOS_MSG_OTP_ISPRAVKA_PAUZIRANA") & vbCrLf
    End If

    OtpremnicaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Zaglavlje i ocekivanje nacrta iz unosa F2 -- JEDNO mesto za upis i izmenu
' nacrta, da dva puta ne bi razlicito citala ista polja. False = unos se ne
' moze prevesti; razlog je dopisan u "poruke".
Private Function OtpremnicaNacrtIzUnosa(ByVal p As Object, ByRef h As Object, _
                                        ByRef ocek As Collection, _
                                        ByRef poruke As String) As Boolean
    Dim kulturaID As String, detalj As String

    ' (Vrsta, Sorta) -> KulturaID. Razresavanje je posao ADAPTERA, ne pisca
    ' (S4.1f) -- isti razresivac koristi i otkupni list.
    kulturaID = modOtkup.RazresiKulturuIzVrsteSorte(S(p, "vrsta"), S(p, "sorta"), detalj)
    If Len(kulturaID) = 0 Then
        poruke = poruke & Poruka("OTKUNOS_ERR_KULTURA") & " " & _
                 S(p, "vrsta") & " / " & S(p, "sorta")
        Exit Function
    End If

    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", CDate(p("datum"))
    h.Add "StanicaID", S(p, "stanicaID")
    h.Add "VozacID", S(p, "vozacID")
    h.Add "KulturaID", kulturaID
    h.Add "BrojOtpremnice", S(p, "brDok")
    h.Add "TipAmbalaze", S(p, "tipAmb")

    ' BrojZbirne se NE salje: broj ne sme da bude veza nego labela (A2), a
    ' zbirna svoje otpremnice drzi kroz tblZbirnaIzvori. Do S3a je ovde stajao
    ' malina trik "snimi sa PRAZNIM brojem da ga auto-zbirna pokupi" -- v.
    ' OtpremnicaUpisi zasto je i sama auto-zbirna pauzirana.
    Set ocek = New Collection
    If D(p, "kolicinaI") > 0 Then
        ocek.Add OtpStavkaDTO(KLASA_I, D(p, "kolicinaI"), L(p, "kolAmb"), D(p, "cenaI"))
    End If
    If B(p, "dveKlase") And D(p, "kolicinaII") > 0 Then
        ocek.Add OtpStavkaDTO(KLASA_II, D(p, "kolicinaII"), L(p, "kolAmbII"), D(p, "cenaII"))
    End If
    OtpremnicaNacrtIzUnosa = True
End Function

' IZMENA NACRTA (S3b-2, odluka 19.09.2026). Kad povezano nije jednako
' ocekivanom -- operater je pogresio ocekivanje, ili je teret stvarno drugaciji
' -- nacrt se ispravlja ovde, a ne izjednacavanjem pri izdavanju: ocekivanje
' ostaje nezavisna kontrola. Clanstvo se ne dira (UpdateOtpremnicaDraft_TX), a
' izdata otpremnica se ne menja (pisac trazi DRAFT).
Public Function OtpremnicaIzmeniNacrt(ByVal otpremnicaID As String, ByVal p As Object, _
                                      ByRef poruke As String) As Boolean
    Dim h As Object, ocek As Collection, greska As String, errDesc As String
    On Error GoTo EH
    poruke = ""
    If Not OtpremnicaNacrtIzUnosa(p, h, ocek, poruke) Then Exit Function
    If Not UpdateOtpremnicaDraft_TX(otpremnicaID, h, ocek, greska) Then
        poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & greska
        Exit Function
    End If
    poruke = poruke & Poruka("DOKUNOS_MSG_OTP_NACRT_IZMENJEN") & vbCrLf
    OtpremnicaIzmeniNacrt = True
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaIzmeniNacrt"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'=====================================================================
' F3 ZBIRNA
'
' Zbirna nije "jos jedan robni dokument": ona je POKLOPAC nad otpremnicama
' jednog vozaca. Zato je vozac njen entitet niza (Z3a) i zato se njen zbir
' poredi sa izvorom pre upisa.
'=====================================================================

' Prazan recnik sa svim kljucevima - da pozivalac ne mora da pamti spisak.
' BROJ ZBIRNE JE brDok: u rezimu F3 broj dokumenta JESTE broj zbirne, pa
' zasebnog polja nema (modOtkupUI.ApplyFormFields / ModeVezujeZbirnu).
Public Function NoviZbirnaUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("vozacID") = ""
    p("kupacID") = ""
    p("brDok") = ""
    p("hladnjaca") = ""
    p("pogon") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("tipAmb") = ""
    p("kolicinaI") = 0#
    p("kolAmb") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("kolAmbII") = 0&
    Set NoviZbirnaUnos = p
End Function

' Redosled provera je isti kao u frmDokumenta.btnUnosZbr_Click - to nije stil
' nego ponasanje: operater je navikao koje ga polje prvo zaustavi.
' Kod razloga za RODITELJA -> korisnicki tekst. NEMA se prevodi u postojecu
' poruku "zbirna ne postoji": do nje se kroz PrijemnicaValidiraj ne stize (kapija
' je u Else grani, gde zbirna postoji), ali funkcija ne sme da vrati prazno i tako
' izgleda kao prolaz.
Private Function ZbirnaRoditeljPoruka(ByVal razlog As String) As String
    Select Case razlog
        Case ZBR_PARENT_DVOSMISLEN: ZbirnaRoditeljPoruka = Poruka("DOKUNOS_ERR_ZBR_P_DVOSMISLEN")
        Case ZBR_PARENT_TUDJ: ZbirnaRoditeljPoruka = Poruka("DOKUNOS_ERR_ZBR_P_TUDJ")
        Case ZBR_PARENT_ISTORIJA: ZbirnaRoditeljPoruka = Poruka("DOKUNOS_ERR_ZBR_P_ISTORIJA")
        Case ZBR_PARENT_NEMA: ZbirnaRoditeljPoruka = Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_1") & " " & _
                                                     Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_2")
        Case Else: ZbirnaRoditeljPoruka = Poruka("DOKUNOS_ERR_ZBR_INTEGRITET")
    End Select
End Function

' VALIDACIJA UNOSA ZBIRNE NAD KANONSKIM MODELOM (S4-2c/2b-1).
'
' Sta je NESTALO, i zasto:
'   - poredjenje sa izvorom po BrojZbirne -- clanstvo je zapis, ne labela
'     (ZBR-KANON-01), pa se pokrivenost meri pri IZDAVANJU, nad vezanim
'     otpremnicama, a ne pri unosu nad pogodjenim brojem;
'   - trazenje VRSTE i SORTE -- one su cinjenica ROBE koju donosi prvi izvor
'     (ZBR-KANON-04); zaglavlje nacrta ih ni ne prima, pa bi ih ekran trazio
'     samo da bi ih pisac odbio;
'   - GeneracijaID kapije (ZbirnaIdentResolve / ZbirnaNovUnosRazlog) -- identitet
'     je ZbirnaID od S4-2a, a taj okvir umire u S4-3.
'
' Sta je OSTALO: polja bez kojih dokument ne postoji, bar jedna klasa sa
' kilazom, nenegativna ambalaza, i BROJ.
'
' Broj se sudi ISTIM alatom kojim ga sudi pisac (.claude/rules/testovi.md SS5):
' modBrojevi.BrojOdgovaraKontekstu i BrojZauzetUNizu -- jedna implementacija,
' dva pozivaoca. Pisac ih zove kroz Require* i die glasno; unos ih zove direktno
' i vraca poruku uz polje. Kopija pravila ovde bi se razisla prvom izmenom.
'
' zbirnaID se prosledjuje pri IZMENI nacrta: nacrt sme da zadrzi svoj broj, pa
' se sopstveni red izuzima -- isto kao u ZbrIzmeniDraft.
Public Function ZbirnaValidiraj(ByVal p As Object, ByRef fokus As String, _
                                Optional ByVal zbirnaID As String = "") As String
    Dim kolI As Double, kolII As Double
    Dim kolAmb As Long, kolAmbII As Long
    Dim dveKl As Boolean
    Dim datum As Date
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""

    ' Vozac je entitet NIZA zbirne (Z3a): po njemu se broji i njegova je tura.
    If Len(S(p, "vozacID")) = 0 Then
        fokus = "vozacID": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function
    End If
    If Len(S(p, "kupacID")) = 0 Then
        fokus = "kupacID": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_KUPAC"): Exit Function
    End If
    If Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_BROJ_ZBIRNE"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    kolII = D(p, "kolicinaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    dveKl = B(p, "dveKlase")

    ' Bez ukljucene druge klase njena polja NE ulaze u najavu -- inace bi
    ' zaostala vrednost tiho postala stavka koju operater ne vidi.
    If Not dveKl Then
        kolII = 0
        kolAmbII = 0
    End If

    If kolI <= 0 And kolII <= 0 Then
        fokus = "kolicinaI"
        ZbirnaValidiraj = Poruka("DOKUNOS_ERR_ZBR_NEMA_KOLICINE")
        Exit Function
    End If

    If kolAmb < 0 Or kolAmbII < 0 Then
        fokus = "kolAmb"
        ZbirnaValidiraj = Poruka("DOK_LBL_NEISPRAVNA_KOLICINA_AMBALAZE")
        Exit Function
    End If

    datum = CDate(p("datum"))

    If modBrojevi.BrojKontekstOdbija( _
           modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, S(p, "vozacID"), _
                                            datum, S(p, "brDok"))) Then
        fokus = "brDok"
        ZbirnaValidiraj = Poruka("DOKUNOS_ERR_ZBR_TUDJ")
        Exit Function
    End If

    If Len(modBrojevi.BrojZauzetUNizu(modBrojevi.KIND_ZBR, S(p, "vozacID"), _
                                      datum, S(p, "brDok"), zbirnaID)) > 0 Then
        fokus = "brDok"
        ZbirnaValidiraj = Poruka("DOKUNOS_ERR_ZBR_BROJ_ZAUZET")
        Exit Function
    End If

    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error Resume
    ' Next, koji cisti Err - operater bi inace dobio poruku bez objasnjenja.
    errDesc = Err.description
    LogErr "modDokUnos.ZbirnaValidiraj"
    ZbirnaValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Zaglavlje i ocekivanje nacrta iz unosa F3 -- JEDNO mesto za upis i izmenu,
' da dva puta ne bi razlicito citala ista polja.
'
' VRSTA, SORTA I TIP AMBALAZE SE NE SALJU. HdrProveriKljuceve ih izricito ne
' prima: one su cinjenica robe koju donosi prvi izvor (ZBR-KANON-04). Poslati ih
' znacilo bi napraviti drugi izvor istine i dobiti pad na piscu.
Private Function ZbirnaNacrtIzUnosa(ByVal p As Object, ByRef h As Object, _
                                    ByRef ocek As Collection, _
                                    ByRef poruke As String) As Boolean
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", CDate(p("datum"))
    h.Add "VozacID", S(p, "vozacID")
    h.Add "BrojZbirne", S(p, "brDok")
    h.Add "KupacID", S(p, "kupacID")
    h.Add "Hladnjaca", S(p, "hladnjaca")
    h.Add "Pogon", S(p, "pogon")

    Set ocek = New Collection
    If D(p, "kolicinaI") > 0 Then
        ocek.Add ZbrStavkaDTO(KLASA_I, D(p, "kolicinaI"), L(p, "kolAmb"))
    End If
    If B(p, "dveKlase") And D(p, "kolicinaII") > 0 Then
        ocek.Add ZbrStavkaDTO(KLASA_II, D(p, "kolicinaII"), L(p, "kolAmbII"))
    End If

    If ocek.count = 0 Then
        poruke = poruke & Poruka("DOKUNOS_ERR_ZBR_NEMA_KOLICINE")
        Exit Function
    End If

    ZbirnaNacrtIzUnosa = True
End Function

Private Function ZbrStavkaDTO(ByVal klasa As String, ByVal kol As Double, _
                              ByVal amb As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "KolAmbalaze", amb
    Set ZbrStavkaDTO = s
End Function

' ZASTO NACRT, A NE GOTOV DOKUMENT (odluka operatera, 21.09.2026): zbirna je
' poklopac nad IZDATIM otpremnicama, pa dok se one ne vezu ona jos ne postoji
' kao dokument -- postoji kao NAJAVA sta ce nositi. Izdavanje (IzdajZbirnu_TX)
' tek tada poredi najavljeno sa vezanim.
'
' Vraca ID, ne broj: broj je labela jedinstvena tek po (vozac, dan), a mutacija
' ide po ID-u. Ekran operateru i dalje pokazuje broj -- to je prikaz, ne
' identitet.
Public Function ZbirnaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, greska As String, errDesc As String
    On Error GoTo EH
    poruke = ""

    Dim h As Object, ocek As Collection
    If Not ZbirnaNacrtIzUnosa(p, h, ocek, poruke) Then Exit Function

    res = CreateZbirnaDraft_TX(h, ocek, greska)

    If Len(res) = 0 Then
        poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & greska
        Exit Function
    End If

    poruke = poruke & Poruka("DOKUNOS_MSG_ZBR_NACRT") & vbCrLf

    ' ISPRAVKA ZBIRNE JE PAUZIRANA, ne prevedena -- isti stav kao kod otpremnice.
    ' Zavrsetak ispravke ide po BROJU i kroz okvir koji umire u S4-3; pustiti ga
    ' nad upravo otvorenim NACRTOM proglasilo bi zamenom dokument koji jos nema
    ' nijedan izvor. Operater mora da ZNA da kontekst stoji otvoren.
    If modStornoContext.CountPendingCorrectionsByDocType(FLOW_DOC_ZBIRNA, _
                                                         SV_MODE_ISPRAVKA) > 0 Then
        poruke = poruke & Poruka("DOKUNOS_MSG_ZBR_ISPRAVKA_PAUZIRANA") & vbCrLf
    End If

    ZbirnaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.ZbirnaUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' IZMENA NACRTA. Kad povezano nije jednako najavljenom -- operater je pogresio
' najavu, ili je teret stvarno drugaciji -- nacrt se ispravlja ovde, a ne
' izjednacavanjem pri izdavanju: najava ostaje nezavisna kontrola. Clanstvo se
' ne dira, a izdata zbirna se ne menja (pisac trazi DRAFT).
Public Function ZbirnaIzmeniNacrt(ByVal zbirnaID As String, ByVal p As Object, _
                                  ByRef poruke As String) As Boolean
    Dim h As Object, ocek As Collection, greska As String, errDesc As String
    On Error GoTo EH
    poruke = ""
    If Not ZbirnaNacrtIzUnosa(p, h, ocek, poruke) Then Exit Function
    If Not UpdateZbirnaDraft_TX(zbirnaID, h, ocek, greska) Then
        poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & greska
        Exit Function
    End If
    poruke = poruke & Poruka("DOKUNOS_MSG_ZBR_NACRT_IZMENJEN") & vbCrLf
    ZbirnaIzmeniNacrt = True
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.ZbirnaIzmeniNacrt"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Upisuje zbirnu. Vraca ZbirnaID (ili spojene ID-eve obe klase); prazno znaci
' da upis nije uspeo.
'
' HLADNJACA I POGON dolaze iz ekrana (v6-ui-215, MIG-001): F3 ima svoja dva
' polja, ljuska ih salje kroz p("hladnjaca") / p("pogon"), a ovde se samo
' prosledjuju writeru. Do tada su isla prazna, pa je svaka zbirna uneta kroz
' ljusku imala prazne kolone koje writer i modDokumentInvariant ipak nose.
' Prazna vrednost je i dalje legitimna (kupac bez hladnjace, kolone nema).
'=====================================================================
' F4 PRIJEMNICA
'
' Prijemnica je prijem robe kod kupca po JEDNOJ zbirnoj. Zato joj je broj
' zbirne obavezan, zbirna mora da postoji, i vazi pravilo 1 zbirna = 1
' prijemnica.
'=====================================================================

' Prazan recnik sa svim kljucevima. "kolAmbVracena" je prazna ambalaza koju
' kupac VRACA (u F1 isto polje znaci izdatu - modOtkupUI.ApplyFormFields).
Public Function NoviPrijemnicaUnos() As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = Date
    p("kupacID") = ""
    p("vozacID") = ""
    p("brDok") = ""
    p("brojZbirne") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("tipAmb") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("kolAmbVracena") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("brutoKgI") = 0#
    p("brutoKgII") = 0#
    ' Ispravka posle storna: puni ih PrijemnicaValidiraj kad prepozna da je
    ' ovaj unos ZAMENA za storniranu prijemnicu, a cita ih PrijemnicaUpisi.
    ' Prazan ispravkaID znaci obican unos.
    p("ispravkaID") = ""
    p("ispravkaStariBroj") = ""
    Set NoviPrijemnicaUnos = p
End Function

' Redosled provera je isti kao u frmDokumenta.btnUnosPrij_Click.
' UPISUJE NAZAD: kolicinaI/kolicinaII postaju NETO, brutoKgI/brutoKgII
' zamrznuti uneti bruto (kad je OTKUP_BRUTO_UNOS ON).
Public Function PrijemnicaValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim kolI As Double, cenI As Double, kolII As Double, cenII As Double
    Dim kolAmb As Long, kolAmbII As Long
    Dim imaKlasaI As Boolean, dveKl As Boolean, strogo As Boolean
    Dim tara As Double, taraII As Double
    Dim postojeca As String, errDesc As String
    On Error GoTo EH
    fokus = ""

    ' F4 JE PAUZIRAN DO S6 (Prijemnica cutover; review #362) -- ne do S4: S4 vraca
    ' zbirnu, a prijemnica dobija svoj header + stavke + izvore tek u S6. Razlog
    ' pauze je isti kao kod F3 (ZbirnaValidiraj): prijemnica
    ' trazi postojecu zbirnu, a zbirna se od S3a ne moze napraviti. Operater
    ' dobija razlog umesto poruke o zbirnoj koja "ne postoji".
    PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_PRIJEMNICA_PAUZIRANA")
    Exit Function

    strogo = IsValidacijaUnosa()

    If Len(S(p, "kupacID")) = 0 Then
        fokus = "kupacID": PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_KUPAC"): Exit Function
    End If
    If Len(S(p, "vozacID")) = 0 Then
        fokus = "vozacID": PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function
    End If
    ' Broj prijemnice se trazi uvek: hladnjaca ga dobija kao predlog, ostali
    ' kupci ga kucaju (Z3a), ali prazan ne prolazi ni u jednom slucaju.
    If Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": PrijemnicaValidiraj = Poruka("OTKUI_ERR_BROJ"): Exit Function
    End If
    If Len(S(p, "brojZbirne")) = 0 Then
        fokus = "brojZbirne": PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_BROJ_ZBIRNE"): Exit Function
    End If

    ' Zbirna mora da postoji u sistemu. Ponasanje po PRIJEMNICA_ZBIRNA_PROVERA:
    ' BLOK = prekid; UPOZORENJE = potvrda pa nastavak.
    If Not ZbirnaPostoji(S(p, "brojZbirne")) Then
        fokus = "brojZbirne"
        If PrijemnicaZbirnaBlokira() Then
            PrijemnicaValidiraj = Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_1") & " '" & _
                                  S(p, "brojZbirne") & "' " & Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_2")
            Exit Function
        End If
        If MsgBox(Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_1") & " '" & S(p, "brojZbirne") & "' " & _
                  Poruka("DOKUNOS_ERR_ZBIRNA_NEMA_2") & vbCrLf & vbCrLf & _
                  Poruka("DOKUNOS_ASK_ZBIRNA_NEMA"), _
                  vbQuestion + vbYesNo, APP_NAME) <> vbYes Then
            PrijemnicaValidiraj = " ": Exit Function
        End If
        fokus = ""            ' operater je potvrdio - polje vise nije sporno
    Else
        ' I2 (docs/DOMEN/ZBR_IDENTITET.md par.6): "postoji" NIJE isto sto i
        ' "jednoznacna". ZbirnaPostoji odgovara samo na prvo -- vraca True cim
        ' ijedan aktivan red nosi taj broj, pa bi se prijemnica vezala i kad pod
        ' brojem stoje DVA dokumenta, ili dokument DRUGOG vlasnika.
        '
        ' Veza se upisuje kao GOLA LABELA (COL_PRJ_BROJ_ZBIRNE), pa nizvodne
        ' operacije po broju mogu da zahvate tudje. Zato i UNIQUE danas pada ako
        ' je broj IKAD imao dva vlasnika -- isto pravilo koje drzi
        ' modStorno.RequireJedanVlasnikIkadPoBroju.
        '
        ' Ovo je TVRDA blokada, ne UPOZORENJE: PRIJEMNICA_ZBIRNA_PROVERA bira
        ' politiku za "zbirne nema" -- stanje koje operater moze da zna unapred.
        ' Dvosmislen ili tudj dokument nije to; tu potvrda ne pomaze jer ekran ne
        ' moze ni da ponudi koji je pravi.
        Dim zbrRod As ZbirnaIdent
        Dim rodRazlog As String
        zbrRod = ZbirnaIdentResolve(S(p, "brojZbirne"), S(p, "vozacID"), S(p, "kupacID"))
        rodRazlog = ZbirnaRoditeljRazlog(zbrRod)
        If Len(rodRazlog) > 0 Then
            fokus = "brojZbirne"
            PrijemnicaValidiraj = ZbirnaRoditeljPoruka(rodRazlog)
            Exit Function
        End If
    End If

    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If
    If strogo And Len(S(p, "sorta")) = 0 Then
        fokus = "sorta": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_SORTA"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    cenI = D(p, "cenaI")
    kolII = D(p, "kolicinaII")
    cenII = D(p, "cenaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    dveKl = B(p, "dveKlase")
    imaKlasaI = (kolI > 0)

    If Not imaKlasaI Then
        If Not dveKl Then
            fokus = "kolicinaI": PrijemnicaValidiraj = Poruka("OTKUI_ERR_KOLICINA"): Exit Function
        End If
        If kolAmb > 0 Then
            fokus = "kolAmb": PrijemnicaValidiraj = Poruka("DOK_MSG_UNOSI_SAMO_KLASA"): Exit Function
        End If
    End If

    ' Cena I: obavezna samo uz strogu validaciju; van nje sme i 0, ali ne negativna.
    If strogo And imaKlasaI Then
        If cenI <= 0 Then
            fokus = "cenaI": PrijemnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
        End If
    ElseIf cenI < 0 Then
        fokus = "cenaI": PrijemnicaValidiraj = Poruka("OTKUI_ERR_CENA"): Exit Function
    End If

    If strogo Then
        If imaKlasaI And kolAmb <= 0 Then
            fokus = "kolAmb": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_I"): Exit Function
        End If
        If dveKl And kolAmbII <= 0 Then
            fokus = "kolAmbII": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_GAJBE_II"): Exit Function
        End If
        If (kolAmb > 0 Or kolAmbII > 0) And Len(S(p, "tipAmb")) = 0 Then
            fokus = "tipAmb": PrijemnicaValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
        End If
    End If

    If dveKl Then
        If kolII <= 0 Then
            fokus = "kolicinaII": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_KOLICINA_II"): Exit Function
        End If
        If cenII <= 0 Then
            fokus = "cenaII": PrijemnicaValidiraj = Poruka("OTKUNOS_ERR_CENA_II"): Exit Function
        End If
    End If

    ' --- BRUTO -> NETO. Prijemnica se cuva u NETO, isto kao otkup i otpremnica,
    ' da manjak i izvestaji porede neto sa neto. Uneti bruto se zamrzava u
    ' BrutoKg. Tara se vezuje za klasu ciji su to gajbici. ---
    If OtkupBrutoUnos() And kolAmb > 0 Then
        tara = kolAmb * GetTezinaGajbice(S(p, "tipAmb"))
        If tara <= 0 Then
            fokus = "tipAmb"
            PrijemnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If tara >= kolI Then
            fokus = "kolicinaI"
            PrijemnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE") & Format$(tara, "#,##0.00") & _
                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgI") = kolI
        kolI = kolI - tara
        p("kolicinaI") = kolI
    End If

    If dveKl And OtkupBrutoUnos() And kolAmbII > 0 Then
        taraII = kolAmbII * GetTezinaGajbice(S(p, "tipAmb"))
        If taraII <= 0 Then
            fokus = "tipAmb"
            PrijemnicaValidiraj = Poruka("DOK_MSG_TIP_AMBALAZE") & S(p, "tipAmb") & _
                                  Poruka("DOK_MSG_NEMA_UNETU_TEZINU")
            Exit Function
        End If
        If taraII >= kolII Then
            fokus = "kolicinaII"
            PrijemnicaValidiraj = Poruka("DOK_MSG_TEZINA_AMBALAZE_KLASE") & Format$(taraII, "#,##0.00") & _
                                  " kg) " & Poruka("OTKUNOS_ERR_TARA_VECA")
            Exit Function
        End If
        p("brutoKgII") = kolII
        kolII = kolII - taraII
        p("kolicinaII") = kolII
    End If

    Dim dup As String
    dup = CheckDuplicate(TBL_PRIJEMNICA, COL_PRJ_BROJ, S(p, "brDok"), COL_PRJ_DATUM)
    If Len(dup) > 0 Then
        fokus = "brDok": PrijemnicaValidiraj = dup: Exit Function
    End If

    ' ISPRAVKA POSLE STORNA - mora PRE provere "1 zbirna = 1 prijemnica".
    ' Ako je ovaj unos zamena za storniranu prijemnicu, ta provera ne vazi:
    ' zbirna namerno dobija novu prijemnicu umesto oborene.
    Dim isprPoruka As String
    isprPoruka = PrepoznajIspravkuPrijemnice(p)
    If Len(isprPoruka) > 0 Then
        fokus = "brojZbirne": PrijemnicaValidiraj = isprPoruka: Exit Function
    End If

    ' 1 zbirna = 1 prijemnica. Ako zbirna vec ima AKTIVNU prijemnicu, ovo je
    ' verovatno dupli unos -> pitanje, ne greska. Uz ispravku se preskace
    ' (isti izuzetak ima i legacy btnUnosPrij_Click).
    postojeca = ""
    If Len(S(p, "ispravkaID")) = 0 Then _
        postojeca = LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                   S(p, "brojZbirne"), COL_PRJ_BROJ)
    If Len(postojeca) > 0 Then
        If MsgBox(Poruka("DOKUNOS_ASK_DUPLA_PRIJ_1") & " " & S(p, "brojZbirne") & " " & _
                  Poruka("DOKUNOS_ASK_DUPLA_PRIJ_2") & " " & postojeca & "." & vbCrLf & vbCrLf & _
                  Poruka("DOKUNOS_ASK_DUPLA_PRIJ_3"), _
                  vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then
            fokus = "brojZbirne": PrijemnicaValidiraj = " ": Exit Function
        End If
    End If
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.PrijemnicaValidiraj"
    PrijemnicaValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

'--------------------------------------------- ISPRAVKA PRIJEMNICE
' Da li je OVAJ unos zamena za storniranu prijemnicu.
'
' Novi UI nema stanje sesije izmedju storna i unosa (storno se pokrece u
' F8, unos u F4), pa se ispravka na cekanju trazi u tblStornoVeza - istoj
' persistentnoj evidenciji koju legacy koristi kao crash-safe putanju kad
' je forma bila zatvorena izmedju ta dva koraka.
'
' SAFE-STOP: dve ili vise ispravki na cekanju = ne biraj naslepo. Pogresno
' pogodjena veza bi palete jedne prijemnice prevezala na tudju robu.
'
' Vraca False SAMO kad je operater rekao "ne snimaj jos" (promenjena
' zbirna, odgovor OTKAZI). U svim ostalim slucajevima vraca True, a da li
' je ispravka prihvacena vidi se po p("ispravkaID").
' Ima li ispravke na cekanju za dati tip dokumenta. ODVOJENO od pitanja
' operateru, iz dva razloga: da se moze testirati bez dijaloga, i da
' odluka "sme li se uopste nastaviti" ne zavisi od toga da li je neko
' kliknuo Da.
'
' Vraca:
'    0  nema ispravke na cekanju -> obican unos
'    1  ima tacno jedna -> cid / stariBroj / parentBroj su popunjeni
'   -1  STOP: ne zna se koja, ili citanje evidencije nije uspelo; razlog
'       nosi poruku za operatera
'
' FAIL-CLOSED je ovde poslovno pravilo, ne stil. Ako se evidencija
' ispravki ne moze procitati, a ispravka mozda postoji, "nastavi kao
' obican unos" znaci: nova prijemnica dobija SVEZE palete, stare ostaju
' osirocene, a correction ostaje pending i ceka jos jednu prijemnicu.
' Neizvesnost mora da zaustavi upis, ne da ga propusti.
'
' SAFE-STOP nad dve ili vise: pogresno pogodjena veza prevezala bi palete
' jedne prijemnice na tudju robu.
Public Function NadjiIspravku(ByVal docType As String, ByRef cid As String, _
                              ByRef stariBroj As String, ByRef parentBroj As String, _
                              ByRef razlog As String) As Long
    Dim cnt As Long
    On Error GoTo EH
    cid = "": stariBroj = "": parentBroj = "": razlog = ""

    cnt = modStornoContext.CountPendingCorrectionsByDocType(docType, SV_MODE_ISPRAVKA)
    If cnt = 0 Then Exit Function
    If cnt > 1 Then
        razlog = Poruka("DOKUNOS_MSG_VISE_ISPRAVKI_PRIJ")
        NadjiIspravku = -1
        Exit Function
    End If

    cid = modStornoContext.FindLatestPending(docType, SV_MODE_ISPRAVKA)
    stariBroj = modStornoContext.GetCorrectionField(cid, COL_SV_OLD_BROJ)
    If Len(cid) = 0 Or Len(stariBroj) = 0 Then
        ' Brojac kaze da ispravka postoji, a ne moze da se procita koja -
        ' to je isti stepen neizvesnosti kao greska, pa isti ishod.
        razlog = Poruka("DOKUNOS_ERR_ISPRAVKA_NECITLJIVA")
        NadjiIspravku = -1
        Exit Function
    End If
    parentBroj = modStornoContext.GetCorrectionField(cid, COL_SV_PARENT_BROJ)
    NadjiIspravku = 1
    Exit Function
EH:
    LogErr "modDokUnos.NadjiIspravku"
    razlog = Poruka("DOKUNOS_ERR_ISPRAVKA_NECITLJIVA") & " " & Err.description
    NadjiIspravku = -1
End Function

' Pita operatera i, ako je odgovorio potvrdno, oznaci unos kao zamenu.
'
' Vraca "" kad upis sme dalje, jedan RAZMAK kad je operater sam odustao
' (poruka se ne prikazuje), inace poruku koja zaustavlja upis.
Private Function PrepoznajIspravkuPrijemnice(ByVal p As Object) As String
    Dim ishod As Long, cid As String, stariBroj As String
    Dim parentZbr As String, razlog As String
    On Error GoTo EH

    ishod = NadjiIspravku(FLOW_DOC_PRIJEMNICA, cid, stariBroj, parentZbr, razlog)
    If ishod = 0 Then Exit Function                 ' obican unos
    If ishod < 0 Then
        PrepoznajIspravkuPrijemnice = razlog        ' neizvesno -> upis staje
        Exit Function
    End If

    ' Potvrda je obavezna: operater je mozda napustio ispravku pa uneo DRUGU
    ' prijemnicu - tiho vezivanje bi joj dodelilo tudje palete.
    If MsgBox(Poruka("DOKUNOS_ASK_ISPRAVKA_PRIJ_1") & " '" & stariBroj & "'." & vbCrLf & vbCrLf & _
              Poruka("DOKUNOS_ASK_ISPRAVKA_PRIJ_2"), _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    ' Promena zbirne NE gasi ispravku sama od sebe - zbirna bira podrazumevanu
    ' vrednost, nije kapija. Ali se PITA: ako je operater u medjuvremenu
    ' promenio zbirnu, mozda ovo vise nije zamena za tu prijemnicu.
    If Len(parentZbr) > 0 And StrComp(S(p, "brojZbirne"), parentZbr, vbTextCompare) <> 0 Then
        Select Case MsgBox(Poruka("DOKUNOS_ASK_ISPRAVKA_ZBIRNA") & " (" & parentZbr & " " & _
                           ChrW(8594) & " " & S(p, "brojZbirne") & ")", _
                           vbQuestion + vbYesNoCancel, APP_NAME)
            Case vbYes:  ' ostaje ispravka
            Case vbNo:   Exit Function                        ' obican, nepovezan unos
            Case Else:   PrepoznajIspravkuPrijemnice = " "     ' ne snimaj jos
                         Exit Function
        End Select
    End If

    p("ispravkaID") = cid
    p("ispravkaStariBroj") = stariBroj
    Exit Function
EH:
    LogErr "modDokUnos.PrepoznajIspravkuPrijemnice"
    ' I ovde fail-closed: greska u prepoznavanju zaustavlja upis.
    PrepoznajIspravkuPrijemnice = Poruka("DOKUNOS_ERR_ISPRAVKA_NECITLJIVA") & " " & Err.description
End Function

' Upisuje prijemnicu. Vraca PrijemnicaID (ili spojene ID-eve obe klase);
' prazno znaci da upis nije uspeo.
'
' ISPRAVKA POSLE STORNA (Faza D/13): kad je ovaj unos zamena za storniranu
' prijemnicu, sveza paletizacija se PRESKACE (SetPaletizeSkip PRE upisa),
' pa se palete stare prevezuju na novu - ista roba, iste palete. Bez toga
' bi ista roba bila paletizovana dvaput: jednom pod starom prijemnicom
' (osirocene palete) i jednom pod novom.
Public Function PrijemnicaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, defHlad As String, palStatus As String, errDesc As String
    Dim ispravka As Boolean
    On Error GoTo EH
    poruke = ""

    ispravka = (Len(S(p, "ispravkaID")) > 0)
    If ispravka Then SetPaletizeSkip True

    res = SavePrijemnicaMulti_TX( _
        datum:=CDate(p("datum")), _
        kupacID:=S(p, "kupacID"), _
        vozacID:=S(p, "vozacID"), _
        brojPrij:=S(p, "brDok"), _
        brojZbirne:=S(p, "brojZbirne"), _
        vrstaVoca:=S(p, "vrsta"), _
        sortaVoca:=S(p, "sorta"), _
        kolicinaI:=D(p, "kolicinaI"), _
        cenaI:=D(p, "cenaI"), _
        tipAmb:=S(p, "tipAmb"), _
        kolAmb:=L(p, "kolAmb"), _
        kolAmbVracena:=L(p, "kolAmbVracena"), _
        hasKlasaII:=B(p, "dveKlase"), _
        kolicinaII:=D(p, "kolicinaII"), _
        cenaII:=D(p, "cenaII"), _
        kolAmbII:=L(p, "kolAmbII"), _
        brutoKgI:=D(p, "brutoKgI"), _
        brutoKgII:=D(p, "brutoKgII"))

    ' Toggle se vraca UVEK, i kad upis nije uspeo: ostavljen ukljucen bi
    ' sledecoj prijemnici tiho preskocio paletizaciju.
    SetPaletizeSkip False

    If Len(res) = 0 Then Exit Function

    ' Prevezivanje paleta ide ODMAH posle upisa, pre stampe: stampa je
    ' best-effort i sme da padne, a palete ne smeju da ostanu osirocene.
    If ispravka Then PreveziPaleteIspravke p, res, poruke

    ' Auto-stampa SAMO za default hladnjacu: eksterni kupci se ne stampaju sami.
    ' Best-effort - greska u izlazu ne sme da obori potvrdu upisa.
    defHlad = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(defHlad) > 0 And StrComp(S(p, "kupacID"), defHlad, vbTextCompare) = 0 Then
        On Error Resume Next
        OutputPrijemnica res
        ' Malina: grupni otkupni list (obrazac otkupnog lista, podaci s prijemnice).
        If IsMalinaMode() Then OutputGrupniOtkupniList res
        Err.Clear
        On Error GoTo EH
    End If

    ' Status palete (Klasa I = prvi token rezultata) - citanje vec iskomitovanih
    ' tabela, pa ide u napomene uz potvrdu upisa.
    On Error Resume Next
    palStatus = GetPaletaStatusForPrijemnica(Split(res, " + ")(0))
    Err.Clear
    On Error GoTo EH
    If Len(palStatus) > 0 Then poruke = poruke & palStatus & vbCrLf

    PrijemnicaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    SetPaletizeSkip False        ' toggle ne sme da ostane ukljucen ni na gresci
    LogErr "modDokUnos.PrijemnicaUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Nova prijemnica preuzima palete stare (bez ponovne paletizacije - ista
' roba). Razlika u broju gajbica se koriguje U MESTU, na istoj paleti
' (PaletaAdjustPrompt), pa se ne pravi nova.
'
' Sav posao rade postojece rutine: ReassignPaleteToPrijemnica_TX veze,
' modStornoContext zatvara ili oznacava kontekst. Ovde je samo redosled i
' ono sto se javi operateru.
'
' Neuspeh prevezivanja NE obara upis - prijemnica je vec proknjizena. Umesto
' toga se kontekst oznacava kao MANUAL, da posao ostane vidljiv u
' "Osirocenim dokumentima" umesto da se izgubi u poruci koja prodje.
Private Sub PreveziPaleteIspravke(ByVal p As Object, ByVal res As String, _
                                  ByRef poruke As String)
    Dim stariBroj As String, noviBroj As String, cid As String
    Dim relWarn As String, gajbDiff As Boolean, relOk As Boolean
    On Error GoTo EH
    stariBroj = S(p, "ispravkaStariBroj")
    cid = S(p, "ispravkaID")
    ' Ispravka se trosi ODMAH: ponovljen poziv nad istim recnikom ne sme da
    ' prevezuje drugi put.
    p("ispravkaID") = ""
    p("ispravkaStariBroj") = ""
    If Len(stariBroj) = 0 Then Exit Sub

    ' Broj nove prijemnice: ono sto je operater uneo; ako je polje bilo
    ' prazno (auto-broj), procita se sa upravo upisanog reda.
    noviBroj = Trim$(S(p, "brDok"))
    If Len(noviBroj) = 0 Then _
        noviBroj = Trim$(NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, _
                   Trim$(Split(res, " + ")(0)), COL_PRJ_BROJ)))

    ' IZVOR SE SALJE PO IDENTITETU. Correction context nosi PK stornirane
    ' prijemnice; iz njega se cita njena GENERACIJA, pa writer bira bas njene
    ' paletne stavke. Bez toga bi izbor isao po broju - a broj se racuna po
    ' kupcu, pa bi dokument drugog kupca istog broja bio zahvacen.
    '
    ' Kapija nad jednoznacnoscu broja vise nije ovde: ista provera sada stoji u
    ' writeru i vazi za svakog pozivaoca, ukljucujuci legacy formu.
    Dim stariGen As String
    stariGen = GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, _
                              modStornoDok.StorniraniDocID(cid))

    ' CILJ SE TAKODJE SALJE PO IDENTITETU. Nova prijemnica je upravo upisana, pa
    ' joj je PK poznat -- nema razloga da se cilj trazi po broju koji je labela.
    Dim noviGen As String
    noviGen = GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, Trim$(Split(res, " + ")(0)))

    relOk = ReassignPaleteToPrijemnica_TX(stariBroj, noviBroj, relWarn, True, gajbDiff, _
                                          stariGen, noviGen)
    If relOk Then
        poruke = poruke & Poruka("DOKUNOS_MSG_PALETE_PREVEZANE") & " " & _
                 stariBroj & " " & ChrW(8594) & " " & noviBroj & vbCrLf
        If Len(relWarn) > 0 Then poruke = poruke & relWarn & vbCrLf
        If gajbDiff Then poruke = poruke & PaletaAdjustPrompt(noviBroj) & vbCrLf
        If Len(cid) > 0 Then
            modStornoContext.CompleteCorrectionContext cid, "", noviBroj, _
                "Ispravka prijemnice: palete prevezane na " & noviBroj & "."
            StampIspravkaTrace TBL_PRIJEMNICA, COL_PRJ_BROJ, noviBroj, stariBroj, cid
        End If
    Else
        LogRelinkFailure stariBroj, noviBroj, relWarn
        poruke = poruke & Poruka("DOKUNOS_MSG_PALETE_NISU") & " " & relWarn & vbCrLf
        OznaciRucnuIspravku cid, "Auto-prevezivanje paleta nije uspelo: " & relWarn
    End If
    Exit Sub
EH:
    ' NEOCEKIVANA greska je opasnija od ocekivanog neuspeha: prijemnica je vec
    ' snimljena, paletizacija je bila PRESKOCENA, a correction bi bez ovoga
    ' ostao PENDING - pa bi sledeca prijemnica opet bila ponudjena kao zamena
    ' za isti stari dokument. Zato i ova grana zavrsava u MANUAL.
    LogErr "modDokUnos.PreveziPaleteIspravke"
    poruke = poruke & Poruka("DOKUNOS_MSG_PALETE_NISU") & " " & Err.description & vbCrLf
    OznaciRucnuIspravku cid, "Greska pri prevezivanju paleta: " & Err.description
End Sub

' Zatvori correction kao "trazi rucnu intervenciju". Posao tako ostaje vidljiv
' u "Nedovrsenom" umesto da nestane u poruci koja prodje, i - vazno - context
' vise nije PENDING, pa sledeci unos nije lazno ponudjen kao zamena.
Private Sub OznaciRucnuIspravku(ByVal cid As String, ByVal razlog As String)
    On Error Resume Next
    If Len(cid) = 0 Then Exit Sub
    modStornoContext.MarkCorrectionManual cid, _
        "Prevezi palete rucno (Oporavak -> Osirocene palete).", razlog
End Sub

'--------------------------------------------------------- ISPRAVKA
' Zavrsetak ispravke posle snimanja zamenskog dokumenta.
'
' Radi SAMO nad PERSISTENTNOM ispravkom na cekanju (tblStornoVeza), ne nad
' stanjem sesije. Razlog vazi i sad kad F8 ume da pokrene ispravku: storno
' se pokrece u jednom rezimu (F8), a zamenski dokument se unosi u drugom
' (F2/F3/F7), i izmedju to dvoje operater sme da zatvori Excel. Persistentan
' zapis to prezivljava, promenljiva u modulu ne bi.
'
' SAFE-STOP kao u legacy: dve ili vise otvorenih ispravki istog tipa = ne
' biraj naslepo, nego pusti operatera kroz "Osiroceni dokumenti".
'
' PRIJEMNICA NIJE U OVOM SELECT-u namerno: njena ispravka nije samo zatvaranje
' konteksta nego i prevezivanje paleta, koje mora da krene PRE upisa
' (SetPaletizeSkip). Zato ima svoj tok - PrepoznajIspravkuPrijemnice u
' PrijemnicaValidiraj i PreveziPaleteIspravke u PrijemnicaUpisi.
'
' PUBLIC je zbog modNovacUnos (revers, F7): pravilo "posle zamenskog dokumenta
' zavrsi ispravku" je isto za sva tri tipa, pa se zove odavde umesto da se
' prepise u treci modul.
'
' stanicaID / datum nosi samo revers: pisac vraca samo uspeh, pa
' CompleteReversIspravka ReversID zamene nalazi po (broj, stanica, dan) --
' zamena sme na drugu stanicu ili drugi dan. Ostali tipovi ih ne citaju.
Public Sub ZavrsiIspravkuAko(ByVal docType As String, ByVal newBroj As String, _
                             ByRef poruke As String, _
                             Optional ByVal stanicaID As String = "", _
                             Optional ByVal datum As Variant = Empty)
    Dim cnt As Long, cid As String, res As Object
    On Error GoTo EH
    newBroj = Trim$(newBroj)
    If Len(newBroj) = 0 Then Exit Sub

    cnt = modStornoContext.CountPendingCorrectionsByDocType(docType, SV_MODE_ISPRAVKA)
    If cnt = 0 Then Exit Sub                    ' nema ispravke -> obican unos
    If cnt > 1 Then
        poruke = poruke & Poruka("DOKUNOS_MSG_VISE_ISPRAVKI") & vbCrLf
        Exit Sub
    End If

    cid = modStornoContext.FindLatestPending(docType, SV_MODE_ISPRAVKA)
    If Len(cid) = 0 Then Exit Sub

    ' Potvrda je obavezna: operater je mozda napustio ispravku pa uneo DRUGI
    ' dokument - automatsko vezivanje bi tada spojilo pogresna dva.
    If MsgBox(ZavrsiIspravkuPitanje(docType, cid, newBroj, stanicaID, datum), _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    Select Case docType
        Case FLOW_DOC_ZBIRNA:     Set res = CompleteZbirnaIspravka(cid, newBroj)
        Case FLOW_DOC_REVERS:     Set res = CompleteReversIspravka(cid, newBroj, stanicaID, datum)
        Case Else: Exit Sub
    End Select

    If res Is Nothing Then Exit Sub
    If CBool(res("success")) Then
        poruke = poruke & Poruka("DOKUNOS_MSG_ISPRAVKA_OK") & " " & CStr(res("message")) & vbCrLf
    Else
        poruke = poruke & Poruka("DOKUNOS_MSG_ISPRAVKA_NIJE") & " " & CStr(res("message")) & vbCrLf
    End If
    Exit Sub
EH:
    LogErr "modDokUnos.ZavrsiIspravkuAko"
End Sub

' Pitanje pre vezivanja zamene za ispravku na cekanju. Ispravka se bira po TIPU
' (jedna otvorena ispravka tog tipa), pa je ovo pitanje jedina kapija protiv
' vezivanja pogresnog dokumenta. Za revers zato nosi i STANICU i DAN oba
' dokumenta: broj reversa je jedinstven tek u nizu (stanica, dan), pa bi
' "'45' -> '45'?" izgledalo isto i kad je snimljen tudj revers istog broja na
' drugoj stanici -- a jedno "Da" bi zatvorilo pogresnu ispravku.
' Stari revers se cita iz traga (OldDocID = ReversID; stanica i dan iz njegovih
' nogu Stanica), novi iz snimanja.
' Za ostale tipove tekst je nepromenjen. PUBLIC zbog testa: MsgBox se ne meri.
Public Function ZavrsiIspravkuPitanje(ByVal docType As String, ByVal cid As String, _
                                      ByVal newBroj As String, _
                                      Optional ByVal stanicaID As String = "", _
                                      Optional ByVal datum As Variant = Empty) As String
    Dim oldBroj As String, oldID As String, oldOpis As String, newOpis As String
    Dim oldSt As String, oldDan As Long
    oldBroj = modStornoContext.GetCorrectionField(cid, COL_SV_OLD_BROJ)
    If docType = FLOW_DOC_REVERS Then
        oldID = Trim$(modStornoContext.GetCorrectionField(cid, COL_SV_OLD_DOCID))
        If Len(oldID) > 0 Then
            If Len(modStorno.ReversStanicaDan(oldID, oldSt, oldDan)) = 0 Then _
                oldOpis = ReversOpis(oldSt, CDate(oldDan))
        End If
        newOpis = ReversOpis(stanicaID, datum)
    End If
    ZavrsiIspravkuPitanje = Poruka("DOKUNOS_ASK_ISPRAVKA_1") & " '" & oldBroj & "'" & oldOpis & "." & _
                            vbCrLf & vbCrLf & _
                            Poruka("DOKUNOS_ASK_ISPRAVKA_2") & " '" & Trim$(newBroj) & "'" & newOpis & "?"
End Function

' " (naziv / StanicaID, dd.mm.yyyy)" -- ID ostaje i uz naziv, jer dve stanice
' mogu imati slican naziv. "" kad stanica nije poznata.
' Public: isti opis nose potvrda storna reversa i "Vrati storno" (REV-IDENT-01
' Faza 2b -- isti KOOP broj, smer i dan legalno nose reversi dve stanice).
Public Function ReversOpis(ByVal stanicaID As String, ByVal datum As Variant) As String
    Dim naziv As String
    If Len(Trim$(stanicaID)) = 0 Then Exit Function
    On Error Resume Next
    naziv = Trim$(NzToText(LookupValue(TBL_STANICE, "StanicaID", Trim$(stanicaID), "Naziv")))
    On Error GoTo 0
    If Len(naziv) = 0 Then
        naziv = Trim$(stanicaID)
    Else
        naziv = naziv & " / " & Trim$(stanicaID)
    End If
    If IsDate(datum) Then
        ReversOpis = " (" & naziv & ", " & Format$(CDate(datum), "dd.mm.yyyy") & ")"
    Else
        ReversOpis = " (" & naziv & ")"
    End If
End Function
