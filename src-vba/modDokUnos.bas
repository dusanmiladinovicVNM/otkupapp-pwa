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
'   OtpremnicaUpisi(p, poruke)     SaveOtpremnicaMulti_TX + auto-zbirna
'                                  (MALINA) + zavrsetak ispravke; vraca
'                                  BrojOtpremnice (prazno = nije upisano)
'   ZbirnaValidiraj / ZbirnaUpisi          isto za F3 (SaveZbirnaMulti_TX)
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
'   - ZBIRNA se poredi sa izvorom: zbir kg i ambalaze mora da se poklopi sa
'     nestorniranim otpremnicama te zbirne (ValidateZbirnaPreUnosa). Ta provera
'     NIJE gejtovana VALIDACIJA_UNOSA - u legacy je hard-gate (UpdateValidacija).
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

Public Const DOKUNOS_BUILD As String = "v6-ui-133"

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

    If Len(S(p, "brDok")) > 0 Then
        Dim dup As String
        dup = CheckDuplicate(TBL_OTPREMNICA, COL_OTP_BROJ, S(p, "brDok"), COL_OTP_DATUM)
        If Len(dup) > 0 Then
            fokus = "brDok": OtpremnicaValidiraj = dup: Exit Function
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

'------------------------------------------------------------- UPIS
' Upisuje otpremnicu i radi sve sto ide uz nju. Vraca broj otpremnice;
' prazno znaci da upis nije uspeo. U "poruke" se skupljaju napomene koje
' pozivalac prikazuje posle uspeha.
Public Function OtpremnicaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, brZbrSave As String, createdZbr As Long, autoZbrErr As String
    Dim errDesc As String
    On Error GoTo EH
    poruke = ""

    ' MALINA: otpremnica se snima sa PRAZNIM BrojZbirne da je auto-zbirna pokupi
    ' (broj u formi je samo predlog; auto-zbirna dodeli "S" + broj otpremnice).
    ' Van malina moda ide vrednost iz polja.
    If IsMalinaMode() Then brZbrSave = "" Else brZbrSave = S(p, "brojZbirne")

    res = SaveOtpremnicaMulti_TX( _
        datum:=CDate(p("datum")), _
        stanicaID:=S(p, "stanicaID"), _
        vozacID:=S(p, "vozacID"), _
        brojOtp:=S(p, "brDok"), _
        brojZbirne:=brZbrSave, _
        vrsta:=S(p, "vrsta"), _
        sorta:=S(p, "sorta"), _
        kolicinaI:=D(p, "kolicinaI"), _
        cenaI:=D(p, "cenaI"), _
        tipAmb:=S(p, "tipAmb"), _
        kolAmb:=L(p, "kolAmb"), _
        hasKlasaII:=B(p, "dveKlase"), _
        kolicinaII:=D(p, "kolicinaII"), _
        cenaII:=D(p, "cenaII"), _
        brutoKgI:=D(p, "brutoKgI"), _
        kolAmbII:=L(p, "kolAmbII"), _
        brutoKgII:=D(p, "brutoKgII"))

    If Len(res) = 0 Then Exit Function

    ' MALINA: otpremnica JESTE zbirna -> auto-zbirna iz upravo snimljene.
    ' Upravo snimljena je otvorena (brZbrSave=""), pa mora nastati bar jedna;
    ' tih pad se prijavljuje, da operater vidi da zbirne nema.
    If IsMalinaMode() Then
        On Error Resume Next
        createdZbr = AutoCreateZbirnaFromOtpremnice_TX()
        If Err.Number <> 0 Then
            autoZbrErr = Err.description
            LogErr "modDokUnos.OtpremnicaUpisi.AutoZbirna"
            Err.Clear
        End If
        On Error GoTo EH
        If Len(autoZbrErr) > 0 Or createdZbr < 1 Then
            poruke = poruke & Poruka("DOKUNOS_MSG_ZBIRNA_NIJE")
            If Len(autoZbrErr) > 0 Then poruke = poruke & " " & autoZbrErr
            poruke = poruke & vbCrLf
        End If
    End If

    ' ISPRAVKA_ODMAH: ako na cekanju stoji ispravka otpremnice, upravo snimljena
    ' je njena zamena -> prevezi blokove i rekalkulisi zbirnu. No-op inace.
    ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA, res, poruke

    OtpremnicaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.OtpremnicaUpisi"
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
Public Function ZbirnaValidiraj(ByVal p As Object, ByRef fokus As String) As String
    Dim kolI As Double, kolII As Double
    Dim kolAmb As Long, kolAmbII As Long
    Dim imaKlasaI As Boolean, dveKl As Boolean, strogo As Boolean
    Dim errDesc As String
    On Error GoTo EH
    fokus = ""
    strogo = IsValidacijaUnosa()

    ' Vozac je entitet NIZA zbirne (Z3a): po njemu se broji i njegova je tura.
    If Len(S(p, "vozacID")) = 0 Then
        fokus = "vozacID": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_VOZAC"): Exit Function
    End If
    If Len(S(p, "kupacID")) = 0 Then
        fokus = "kupacID": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_KUPAC"): Exit Function
    End If
    ' Broj se trazi UVEK, i van stroge validacije (drugacije nego kod otpremnice):
    ' po njemu se nalaze otpremnice koje ova zbirna pokriva, pa bez njega nema ni
    ' provere ispod ni veze sa izvorom.
    If Len(S(p, "brDok")) = 0 Then
        fokus = "brDok": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_BROJ_ZBIRNE"): Exit Function
    End If
    If strogo And Len(S(p, "vrsta")) = 0 Then
        fokus = "vrsta": ZbirnaValidiraj = Poruka("OTKUNOS_ERR_VRSTA"): Exit Function
    End If
    If strogo And Len(S(p, "sorta")) = 0 Then
        fokus = "sorta": ZbirnaValidiraj = Poruka("OTKUNOS_ERR_SORTA"): Exit Function
    End If

    kolI = D(p, "kolicinaI")
    kolII = D(p, "kolicinaII")
    kolAmb = L(p, "kolAmb")
    kolAmbII = L(p, "kolAmbII")
    dveKl = B(p, "dveKlase")
    imaKlasaI = (kolI > 0)

    ' Klasa I je opciona SAMO uz ukljucenu Klasu II (unosi se samo II klasa).
    ' Tada ambalaza I mora ostati prazna.
    If Not imaKlasaI Then
        If Not dveKl Then
            fokus = "kolicinaI": ZbirnaValidiraj = Poruka("OTKUI_ERR_KOLICINA"): Exit Function
        End If
        If kolAmb > 0 Then
            fokus = "kolAmb": ZbirnaValidiraj = Poruka("DOK_MSG_UNOSI_SAMO_KLASA"): Exit Function
        End If
    End If

    ' Cene NEMA: tblZbirna nema kolonu Cena i SaveZbirnaMulti_TX je ne prima.
    If dveKl Then
        If kolII <= 0 Then
            fokus = "kolicinaII": ZbirnaValidiraj = Poruka("OTKUNOS_ERR_KOLICINA_II"): Exit Function
        End If
    End If

    If strogo And (kolAmb > 0 Or kolAmbII > 0) And Len(S(p, "tipAmb")) = 0 Then
        fokus = "tipAmb": ZbirnaValidiraj = Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"): Exit Function
    End If

    ' BRUTO->NETO OVDE NAMERNO NEMA. Zbirna je zbir svojih otpremnica, a one su
    ' vec u netu; oduzimanje tare i drugi put spustilo bi kilograme ispod izvora
    ' i oborilo bas provere ispod. tblZbirna zato nema ni kolonu BrutoKg.

    ' Hard-blokada: izvorne otpremnice imaju Klasu II a prekidac je iskljucen ->
    ' SaveZbirnaMulti_TX bi dobio hasKlasaII:=False i Kl.II bi se tiho izgubila.
    If Not dveKl Then
        If ZbirnaIzvorImaKlasuII(S(p, "brDok")) Then
            fokus = "kolicinaII": ZbirnaValidiraj = Poruka("DOKUNOS_ERR_IZVOR_KL2"): Exit Function
        End If
    End If

    ' Zbir mora da se poklopi sa izvorom. U legacy je to hard-gate koji NE zavisi
    ' od VALIDACIJA_UNOSA (btnUnosZbr_Click: "If Not UpdateValidacija()").
    If Not ZbirnaSeSlazeSaIzvorom(S(p, "brDok"), kolI, kolII, kolAmb + kolAmbII, dveKl) Then
        fokus = "kolicinaI"
        ZbirnaValidiraj = Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA")
        Exit Function
    End If

    Dim dup As String
    dup = CheckDuplicate(TBL_ZBIRNA, COL_ZBR_BROJ, S(p, "brDok"), COL_ZBR_DATUM)
    If Len(dup) > 0 Then
        fokus = "brDok": ZbirnaValidiraj = dup: Exit Function
    End If
    Exit Function
EH:
    ' Opis se cita PRE logovanja: LogErr (i Poruka) imaju svoj On Error Resume
    ' Next, koji cisti Err - operater bi inace dobio poruku bez objasnjenja.
    errDesc = Err.description
    LogErr "modDokUnos.ZbirnaValidiraj"
    ZbirnaValidiraj = Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

' Verdikt koji u legacy daje frmDokumenta.UpdateValidacija, bez ijednog natpisa:
' racun je isti (ValidateZbirnaPreUnosa), samo se ovde ne crta.
'   val(0-3)  Klasa I : suma otpremnica | uneto | razlika | poklapa se
'   val(4-7)  Klasa II: isto
'   val(8-10) ambalaza: suma otpremnica | uneto | razlika
' Ambalaza se poredi kao ZBIR OBE KLASE - ValidateZbirnaPreUnosa sabira sumaAmb
' preko obe klase, pa i ulaz mora biti zbir (legacy: inputAmb + inputAmbII).
Private Function ZbirnaSeSlazeSaIzvorom(ByVal brojZbirne As String, _
                                        ByVal kgI As Double, ByVal kgII As Double, _
                                        ByVal amb As Long, ByVal dveKl As Boolean) As Boolean
    Dim val As Variant, kgOK As Boolean
    On Error GoTo EH
    val = ValidateZbirnaPreUnosa(brojZbirne, kgI, kgII, amb)
    If Not IsArray(val) Then Exit Function
    If UBound(val) < 10 Then Exit Function

    If dveKl Then
        kgOK = CBool(val(3)) And CBool(val(7))
    Else
        kgOK = CBool(val(3))
    End If
    ' "zbrAmb > 0" je iz legacy: zbirna bez unete ambalaze se ne pusta, jer bi
    ' razlika ispala 0 i kad izvor ambalazu ima.
    ZbirnaSeSlazeSaIzvorom = kgOK And (CLng(val(10)) = 0) And (CLng(val(9)) > 0)
    Exit Function
EH:
    LogErr "modDokUnos.ZbirnaSeSlazeSaIzvorom"
    ZbirnaSeSlazeSaIzvorom = False
End Function

' Upisuje zbirnu. Vraca ZbirnaID (ili spojene ID-eve obe klase); prazno znaci
' da upis nije uspeo.
'
' HLADNJACA I POGON idu prazni: novi UI nema ta dva polja (odrediste otpremnice,
' katalog Z3b "cmbKupac_Change -> cmbHladnjaca / cmbPogon: NEMA"). Kljucevi u
' recniku postoje da ekran koji ih dobije nema sta da menja ovde.
Public Function ZbirnaUpisi(ByVal p As Object, ByRef poruke As String) As String
    Dim res As String, errDesc As String
    On Error GoTo EH
    poruke = ""

    res = SaveZbirnaMulti_TX( _
        datum:=CDate(p("datum")), _
        vozacID:=S(p, "vozacID"), _
        brojZbirne:=S(p, "brDok"), _
        kupacID:=S(p, "kupacID"), _
        hladnjaca:=S(p, "hladnjaca"), _
        pogon:=S(p, "pogon"), _
        vrstaVoca:=S(p, "vrsta"), _
        sortaVoca:=S(p, "sorta"), _
        ukupnoKolI:=D(p, "kolicinaI"), _
        tipAmb:=S(p, "tipAmb"), _
        ukupnoAmb:=L(p, "kolAmb"), _
        hasKlasaII:=B(p, "dveKlase"), _
        ukupnoKolII:=D(p, "kolicinaII"), _
        ukupnoAmbII:=L(p, "kolAmbII"))

    If Len(res) = 0 Then Exit Function

    ' ISPRAVKA_ODMAH: zavrsetak ide po BROJU zbirne, ne po vracenim ID-evima
    ' (legacy: TryAutoCompleteIspravka FLOW_DOC_ZBIRNA, txtBrojZbirne.value).
    ZavrsiIspravkuAko FLOW_DOC_ZBIRNA, S(p, "brDok"), poruke

    ZbirnaUpisi = res
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modDokUnos.ZbirnaUpisi"
    poruke = poruke & Poruka("OTKUP_ERR_GRESKA_PRI_UNOSU") & errDesc
End Function

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
Public Sub ZavrsiIspravkuAko(ByVal docType As String, ByVal newBroj As String, _
                             ByRef poruke As String)
    Dim cnt As Long, cid As String, oldBroj As String, res As Object
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
    oldBroj = modStornoContext.GetCorrectionField(cid, COL_SV_OLD_BROJ)
    If MsgBox(Poruka("DOKUNOS_ASK_ISPRAVKA_1") & " '" & oldBroj & "'." & vbCrLf & vbCrLf & _
              Poruka("DOKUNOS_ASK_ISPRAVKA_2") & " '" & newBroj & "'?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    Select Case docType
        Case FLOW_DOC_OTPREMNICA: Set res = CompleteOtpremnicaIspravka(cid, newBroj)
        Case FLOW_DOC_ZBIRNA:     Set res = CompleteZbirnaIspravka(cid, newBroj)
        Case FLOW_DOC_REVERS:     Set res = CompleteReversIspravka(cid, newBroj)
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
