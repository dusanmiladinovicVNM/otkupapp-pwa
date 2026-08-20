Attribute VB_Name = "modAgroUnos"
'=====================================================================
' modAgroUnos - UNOS AGROHEMIJE (magacin ulaz/izlaz), bez ijedne kontrole.
'
' Zasto postoji: ceo posao izdavanja i prijema do sada je ziveo u
' frmAgrohemija - korpa u Private tipu forme, provera stanja u
' ValidateKorpaIzlazStanje / BuildArtikalStanjeDict, a transakcija u
' btnZavrsiIzlaz_Click / btnZavrsiUlaz_Click. Novi UI to ne moze da pozove,
' a prepisivanje bi napravilo dve kopije koje se razilaze. Isti razlog i isti
' oblik kao modOtkupUnos (F1), modDokUnos (F2-F4) i modNovacUnos (F5-F7).
'
' Ovde je taj posao izdvojen tako da ga zove i novi ekran (modScrAgro):
'
'   NovaAgroKorpa()                  prazna korpa (Collection redova)
'   AgroArtikalInfo(id)              naziv/jm/cena/pakovanje/doza + greska
'   AgroPreporukaInfo(id, ha)        smart doza -> broj pakovanja
'   AgroStanjeMapa()                 artikalID -> stanje magacina
'   AgroKorpaKolicina(korpa, id)     koliko je tog artikla vec u korpi
'   AgroDodajIzlaz(...)              provere + red u korpu; "" = proslo
'   AgroDodajUlaz(...)               isto za prijem
'   AgroProveriKorpuIzlaz(korpa)     kapija stanja AGREGIRANO po artiklu
'   AgroUpisiIzlaz(korpa, ...)       jedna transakcija za celu korpu
'   AgroUpisiUlaz(korpa, ...)        isto za prijem
'   AgroZbirKorpe(korpa)             zbir vrednosti
'
' Red korpe je RECNIK (Scripting.Dictionary) sa logickim kljucevima:
'   artikalID, naziv, jm, cena, pakovanje, brojPak, kolicina, vrednost,
'   parcelaID, nula
' "kolicina" je uvek u JM artikla (kg/l) - to je ono sto ide u tblMagacin.
' "brojPak" je ono sto operater kuca kod IZLAZA; kod ULAZA je 0.
'
' STA OVDE NIJE: nijedno citanje tabele mimo postojecih rutina
' (GetMagacinStanje, LookupValue, SaveMagacinCore) i nijedna kontrola.
' Cena agrohemije je single-current po artiklu (tblArtikli.CenaPoJedinici) -
' modCenovnik se ovde NE zove, to je model za otkup voca.
'
' LEGACY SE NE DIRA. frmAgrohemija zadrzava svoju kopiju ove logike, isto kao
' frmOtkup i frmDokumenta - dok novi ekran ne prodje rad u pogonu. Pravilo se
' menja OVDE pa se rucno preslikava u formu, i to se zapise uz izmenu.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const AGROUNOS_BUILD As String = "v6-ui-171"

' Odustajanje operatera. Ista konvencija kao Scr_Save u ljusci: jedan razmak
' znaci "nista se nije desilo, ne prikazuj gresku".
Public Const AGRO_ODUSTAO As String = " "

'=====================================================================
' KORPA
'=====================================================================
Public Function NovaAgroKorpa() As Collection
    Set NovaAgroKorpa = New Collection
End Function

Public Function AgroZbirKorpe(ByVal korpa As Collection) As Double
    Dim i As Long
    If korpa Is Nothing Then Exit Function
    For i = 1 To korpa.count
        AgroZbirKorpe = AgroZbirKorpe + AD(korpa(i), "vrednost")
    Next i
End Function

' Koliko je tog artikla vec u korpi (u JM artikla). Izdvojeno iz
' frmAgrohemija.GetKorpaIzlazKolicinaZaArtikal.
Public Function AgroKorpaKolicina(ByVal korpa As Collection, _
                                  ByVal artikalID As String) As Double
    Dim i As Long
    If korpa Is Nothing Then Exit Function
    For i = 1 To korpa.count
        If AS_(korpa(i), "artikalID") = Trim$(artikalID) Then
            AgroKorpaKolicina = AgroKorpaKolicina + AD(korpa(i), "kolicina")
        End If
    Next i
End Function

'=====================================================================
' ARTIKAL I SMART DOZA
'=====================================================================
' Sve sto se o artiklu cita na jednom mestu, sa INVARIJANTOM nad pakovanjem.
' Legacy tu istu proveru ima tri puta (UpdatePreporuka, UpdateVrednost,
' btnDodajIzlaz), pa je i poruka bila tri puta prepisana.
'
' Vraca recnik: naziv, jm, cena, pakovanje, doza, greska.
' Neprazna "greska" znaci da se artiklom ne sme raditi.
Public Function AgroArtikalInfo(ByVal artikalID As String) As Object
    Dim r As Object, pakStr As String, cenaStr As String
    Set r = CreateObject("Scripting.Dictionary")
    r.CompareMode = vbTextCompare
    r("naziv") = ""
    r("jm") = ""
    r("cena") = 0#
    r("pakovanje") = 0#
    r("doza") = 0#
    r("greska") = ""

    If Len(Trim$(artikalID)) = 0 Then
        r("greska") = Poruka("AGROU_ERR_NEMA_ARTIKLA")
        Set AgroArtikalInfo = r
        Exit Function
    End If

    On Error GoTo EH
    r("naziv") = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_NAZIV))
    r("jm") = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_JM))

    cenaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_CENA))
    If IsNumeric(cenaStr) Then r("cena") = CDbl(cenaStr)

    Dim dozaStr As String
    dozaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_DOZA))
    If IsNumeric(dozaStr) Then r("doza") = CDbl(dozaStr)

    ' Invarijanta: svaki artikal mora imati popunjeno Pakovanje. Bez njega se
    ' broj pakovanja ne moze prevesti u kg, pa bi se u tblMagacin upisala
    ' kolicina koja nije ni kg ni komad.
    pakStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_PAKOVANJE))
    If Not IsNumeric(pakStr) Then
        r("greska") = Poruka("AGROU_ERR_PAKOVANJE")
    ElseIf CDbl(pakStr) <= 0 Then
        r("greska") = Poruka("AGROU_ERR_PAKOVANJE")
    Else
        r("pakovanje") = CDbl(pakStr)
    End If

    Set AgroArtikalInfo = r
    Exit Function
EH:
    r("greska") = Poruka("AGROU_ERR_ARTIKAL_CITANJE") & " " & Err.description
    Set AgroArtikalInfo = r
End Function

' Smart doza: doza po hektaru * ha -> zaokruzeno NAGORE na cela pakovanja.
' Racun (CalculatePreporuka) ostaje u modAgrohemija; ovde je samo prevod u
' pakovanja, isti kao u frmAgrohemija.UpdatePreporuka.
'
' Vraca recnik: dozaKg, pakovanje, brojPak, izdajKol, jm, greska.
Public Function AgroPreporukaInfo(ByVal artikalID As String, _
                                  ByVal ukupnoHa As Double) As Object
    Dim r As Object, info As Object, dozaKg As Double, pak As Double
    Set r = CreateObject("Scripting.Dictionary")
    r.CompareMode = vbTextCompare
    r("dozaKg") = 0#
    r("pakovanje") = 0#
    r("brojPak") = 0&
    r("izdajKol") = 0#
    r("jm") = ""
    r("greska") = ""

    Set info = AgroArtikalInfo(artikalID)
    r("jm") = info("jm")
    If Len(CStr(info("greska"))) > 0 Then
        r("greska") = info("greska")
        Set AgroPreporukaInfo = r
        Exit Function
    End If
    If ukupnoHa <= 0 Then
        Set AgroPreporukaInfo = r
        Exit Function
    End If

    pak = CDbl(info("pakovanje"))
    dozaKg = CalculatePreporuka(artikalID, ukupnoHa)
    r("dozaKg") = dozaKg
    r("pakovanje") = pak
    ' zaokruzenje NAGORE - pola pakovanja se ne izdaje
    r("brojPak") = CLng(-Int(-dozaKg / pak))
    r("izdajKol") = CDbl(r("brojPak")) * pak
    Set AgroPreporukaInfo = r
End Function

'=====================================================================
' STANJE MAGACINA
'=====================================================================
' artikalID -> stanje (kolona 7 iz GetMagacinStanje). Izdvojeno iz
' frmAgrohemija.BuildArtikalStanjeDict; racun ostaje u modAgrohemija.
Public Function AgroStanjeMapa() As Object
    Dim d As Object, stanje As Variant, i As Long, artID As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    stanje = GetMagacinStanje()
    If Not IsArray(stanje) Then
        Set AgroStanjeMapa = d
        Exit Function
    End If

    For i = 1 To UBound(stanje, 1)
        artID = Trim$(CStr(stanje(i, 1)))
        If Len(artID) > 0 Then
            If IsNumeric(stanje(i, 7)) Then
                d(artID) = CDbl(stanje(i, 7))
            Else
                d(artID) = 0#
            End If
        End If
    Next i
    Set AgroStanjeMapa = d
End Function

'=====================================================================
' DODAVANJE U KORPU
'=====================================================================
' IZLAZ. Kolicina se kuca kao BROJ PAKOVANJA i tek se ovde prevodi u kg/l -
' isto kao u legacy formi, gde je to bila "kljucna promena".
'
' Vraca "" kad je red dodat, inace poruku za operatera. "fokus" nosi logicko
' ime polja na koje treba vratiti kursor (kao OtkupValidiraj).
Public Function AgroDodajIzlaz(ByVal korpa As Collection, _
                               ByVal artikalID As String, _
                               ByVal brojPakUnos As Double, _
                               ByVal parcelaIDs As String, _
                               ByRef fokus As String) As String
    Dim info As Object, red As Object, brojPak As Long
    Dim kolicina As Double, dostupno As Double, uKorpi As Double
    Dim mapa As Object

    fokus = ""
    If korpa Is Nothing Then
        AgroDodajIzlaz = Poruka("AGROU_ERR_NEMA_KORPE")
        Exit Function
    End If

    Set info = AgroArtikalInfo(artikalID)
    If Len(CStr(info("greska"))) > 0 Then
        fokus = "artikal"
        AgroDodajIzlaz = CStr(info("greska"))
        Exit Function
    End If

    ' Broj pakovanja mora biti CEO broj veci od nule. Legacy je isti uslov
    ' delio na tri poruke; razlika za operatera je ista - broj nije ispravan.
    If brojPakUnos <= 0 Then
        fokus = "kolicina"
        AgroDodajIzlaz = Poruka("AGROU_ERR_PAKOVANJA_BROJ")
        Exit Function
    End If
    brojPak = CLng(Int(brojPakUnos))
    If brojPakUnos <> CDbl(brojPak) Then
        fokus = "kolicina"
        AgroDodajIzlaz = Poruka("AGROU_ERR_PAKOVANJA_CEO")
        Exit Function
    End If

    ' Parcela je obavezna samo dok je pracenje parcela ukljuceno - isti flag
    ' koji gasi i smart dozu (IsPracenjeParcela, kao u frmOtkup).
    If IsPracenjeParcela() Then
        If Len(Trim$(parcelaIDs)) = 0 Then
            fokus = "parcela"
            AgroDodajIzlaz = Poruka("AGROU_ERR_NEMA_PARCELE")
            Exit Function
        End If
    End If

    kolicina = CDbl(brojPak) * CDbl(info("pakovanje"))

    ' Kapija stanja gleda i ono sto je VEC u korpi - inace bi se ista roba
    ' mogla dodati dva puta i tek pri upisu ispasti da je nema.
    Set mapa = AgroStanjeMapa()
    If mapa.Exists(Trim$(artikalID)) Then dostupno = CDbl(mapa(Trim$(artikalID)))
    uKorpi = AgroKorpaKolicina(korpa, artikalID)
    If uKorpi + kolicina > dostupno Then
        fokus = "kolicina"
        AgroDodajIzlaz = Poruka("AGROU_ERR_NEDOVOLJNO") & " " & _
                         CStr(info("naziv")) & vbCrLf & _
                         Poruka("AGROU_LBL_NA_STANJU") & " " & _
                         AgroFmtKol(dostupno) & " " & CStr(info("jm")) & vbCrLf & _
                         Poruka("AGROU_LBL_U_KORPI") & " " & _
                         AgroFmtKol(uKorpi) & " " & CStr(info("jm")) & vbCrLf & _
                         Poruka("AGROU_LBL_DODAJE_SE") & " " & _
                         AgroFmtKol(kolicina) & " " & CStr(info("jm"))
        Exit Function
    End If

    Set red = NoviRed(artikalID, info)
    red("brojPak") = brojPak
    red("kolicina") = kolicina
    red("vrednost") = kolicina * CDbl(info("cena"))
    red("parcelaID") = Trim$(parcelaIDs)
    korpa.Add red
End Function

' ULAZ. Kolicina se kuca u JM artikla (kg/l), ne u pakovanjima - tako radi i
' legacy prijem. Cena se predlaze iz tblArtikli pa se sme ispraviti.
'
' Cena 0 je dozvoljena samo za dokumentovan besplatan / korektivni prijem i
' tek uz izricitu potvrdu. Potvrda stoji OVDE, a ne u ekranu: identicna je za
' svakog pozivaoca, isto kao potvrde u modOtkupUnos.
Public Function AgroDodajUlaz(ByVal korpa As Collection, _
                              ByVal artikalID As String, _
                              ByVal kolicina As Double, _
                              ByVal cena As Double, _
                              ByRef fokus As String) As String
    Dim info As Object, red As Object, nula As Boolean

    fokus = ""
    If korpa Is Nothing Then
        AgroDodajUlaz = Poruka("AGROU_ERR_NEMA_KORPE")
        Exit Function
    End If

    Set info = AgroArtikalInfo(artikalID)
    ' Prijem ne prevodi pakovanja u kg, pa mu invarijanta nad Pakovanjem nije
    ' kapija - trazi se samo da artikal postoji. Zato se ovde gleda naziv, a ne
    ' "greska": inace se roba bez popunjenog Pakovanja ne bi mogla ni primiti,
    ' a bas prijem je trenutak kad se sifarnik dopunjava.
    If Len(Trim$(artikalID)) = 0 Then
        fokus = "artikal"
        AgroDodajUlaz = Poruka("AGROU_ERR_NEMA_ARTIKLA")
        Exit Function
    End If

    If kolicina <= 0 Then
        fokus = "kolicina"
        AgroDodajUlaz = Poruka("AGROU_ERR_KOLICINA")
        Exit Function
    End If

    If cena < 0 Then
        fokus = "cena"
        AgroDodajUlaz = Poruka("AGROU_ERR_CENA")
        Exit Function
    ElseIf cena = 0 Then
        If MsgBox(Poruka("AGRO_MSG_POTVRDI_BESPLATAN_ULAZ"), _
                  vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
            fokus = "cena"
            AgroDodajUlaz = AGRO_ODUSTAO
            Exit Function
        End If
        nula = True
    End If

    Set red = NoviRed(artikalID, info)
    red("brojPak") = 0&
    red("kolicina") = kolicina
    red("cena") = cena
    red("vrednost") = kolicina * cena
    red("nula") = nula
    korpa.Add red
End Function

Private Function NoviRed(ByVal artikalID As String, ByVal info As Object) As Object
    Dim red As Object
    Set red = CreateObject("Scripting.Dictionary")
    red.CompareMode = vbTextCompare
    red("artikalID") = Trim$(artikalID)
    red("naziv") = CStr(info("naziv"))
    red("jm") = CStr(info("jm"))
    red("cena") = CDbl(info("cena"))
    red("pakovanje") = CDbl(info("pakovanje"))
    red("brojPak") = 0&
    red("kolicina") = 0#
    red("vrednost") = 0#
    red("parcelaID") = ""
    red("nula") = False
    Set NoviRed = red
End Function

'=====================================================================
' KAPIJA STANJA PRE UPISA
'=====================================================================
' Ista provera kao pri dodavanju, ali AGREGIRANO po artiklu i nad CELOM
' korpom. Postoji zato sto se stanje moglo promeniti izmedju dodavanja i
' upisa (drugi operater, sync), a delimican prolaz kroz petlju pa rollback je
' losija poruka od jedne recenice pre nego sto transakcija uopste pocne.
' Izdvojeno iz frmAgrohemija.ValidateKorpaIzlazStanje.
Public Function AgroProveriKorpuIzlaz(ByVal korpa As Collection) As String
    Dim mapa As Object, treba As Object, imena As Object, jmd As Object
    Dim i As Long, artID As String, k As Variant
    Dim dostupno As Double, potrebno As Double

    If korpa Is Nothing Then Exit Function
    If korpa.count = 0 Then Exit Function

    Set treba = CreateObject("Scripting.Dictionary")
    treba.CompareMode = vbTextCompare
    Set imena = CreateObject("Scripting.Dictionary")
    imena.CompareMode = vbTextCompare
    Set jmd = CreateObject("Scripting.Dictionary")
    jmd.CompareMode = vbTextCompare

    For i = 1 To korpa.count
        artID = AS_(korpa(i), "artikalID")
        If Len(artID) = 0 Then
            AgroProveriKorpuIzlaz = Poruka("AGROU_ERR_RED_BEZ_ARTIKLA")
            Exit Function
        End If
        If Not treba.Exists(artID) Then
            treba(artID) = 0#
            imena(artID) = AS_(korpa(i), "naziv")
            jmd(artID) = AS_(korpa(i), "jm")
        End If
        treba(artID) = CDbl(treba(artID)) + AD(korpa(i), "kolicina")
    Next i

    Set mapa = AgroStanjeMapa()
    For Each k In treba.keys
        dostupno = 0#
        If mapa.Exists(CStr(k)) Then dostupno = CDbl(mapa(CStr(k)))
        potrebno = CDbl(treba(CStr(k)))
        If dostupno < potrebno Then
            AgroProveriKorpuIzlaz = Poruka("AGROU_ERR_NEDOVOLJNO") & " " & _
                CStr(imena(CStr(k))) & vbCrLf & _
                Poruka("AGROU_LBL_NA_STANJU") & " " & _
                AgroFmtKol(dostupno) & " " & CStr(jmd(CStr(k))) & vbCrLf & _
                Poruka("AGROU_LBL_U_KORPI") & " " & _
                AgroFmtKol(potrebno) & " " & CStr(jmd(CStr(k)))
            Exit Function
        End If
    Next k
End Function

'=====================================================================
' UPIS - jedna transakcija za CELU korpu
'=====================================================================
' Zove se SaveMagacinCore, ne omotac SaveMagacin: omotac gutu typed gresku i
' vraca prazan string, pa operater vidi "nije uspelo" umesto tacnog razloga
' (cena / artikal / kooperant). Isti izbor kao u legacy formi.
'
' Vraca "" kad je proslo, inace poruku. Broj upisanih redova ide u "upisano".
Public Function AgroUpisiIzlaz(ByVal korpa As Collection, _
                               ByVal kooperantID As String, _
                               ByVal brojDok As String, _
                               ByVal datum As Date, _
                               ByRef upisano As Long) As String
    Const SRC As String = "modAgroUnos.AgroUpisiIzlaz"
    Dim tx As clsTransaction, txStarted As Boolean
    Dim i As Long, res As String, greska As String

    upisano = 0
    greska = ProveriZaglavlje(korpa, brojDok)
    If Len(greska) > 0 Then
        AgroUpisiIzlaz = greska
        Exit Function
    End If
    If Len(Trim$(kooperantID)) = 0 Then
        AgroUpisiIzlaz = Poruka("AGROU_ERR_NEMA_KOOPERANTA")
        Exit Function
    End If

    greska = AgroProveriKorpuIzlaz(korpa)
    If Len(greska) > 0 Then
        AgroUpisiIzlaz = greska
        Exit Function
    End If

    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN

    For i = 1 To korpa.count
        res = SaveMagacinCore( _
            datum, _
            AS_(korpa(i), "artikalID"), _
            MAG_IZLAZ, _
            AD(korpa(i), "kolicina"), _
            Trim$(kooperantID), _
            AS_(korpa(i), "parcelaID"), _
            Trim$(brojDok), _
            overrideCena:=AD(korpa(i), "cena"))
        If Len(Trim$(res)) = 0 Then
            Err.Raise vbObjectError + 4301, SRC, _
                      Poruka("AGROU_ERR_UPIS_IZLAZ") & " " & _
                      AS_(korpa(i), "artikalID")
        End If
        upisano = upisano + 1
    Next i

    tx.CommitTx
    txStarted = False
    Set tx = Nothing
    Exit Function
EH:
    AgroUpisiIzlaz = Poruka("AGRO_MSG_GRESKA_PRI_CUVANJU") & " " & Err.description
    upisano = 0
    LogErr SRC
    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
End Function

Public Function AgroUpisiUlaz(ByVal korpa As Collection, _
                              ByVal dobavljacID As String, _
                              ByVal brojDok As String, _
                              ByVal datum As Date, _
                              ByRef upisano As Long) As String
    Const SRC As String = "modAgroUnos.AgroUpisiUlaz"
    Dim tx As clsTransaction, txStarted As Boolean
    Dim i As Long, res As String, greska As String

    upisano = 0
    greska = ProveriZaglavlje(korpa, brojDok)
    If Len(greska) > 0 Then
        AgroUpisiUlaz = greska
        Exit Function
    End If

    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN

    For i = 1 To korpa.count
        ' allowZeroValue nosi dokumentovan besplatan / korektivni prijem; bez
        ' njega ULAZ sa cenom 0 pada isto kao IZLAZ.
        res = SaveMagacinCore( _
            datum, _
            AS_(korpa(i), "artikalID"), _
            MAG_ULAZ, _
            AD(korpa(i), "kolicina"), _
            "", _
            "", _
            Trim$(brojDok), _
            "", _
            Trim$(dobavljacID), _
            AD(korpa(i), "cena"), _
            allowZeroValue:=AB(korpa(i), "nula"))
        If Len(Trim$(res)) = 0 Then
            Err.Raise vbObjectError + 4311, SRC, _
                      Poruka("AGROU_ERR_UPIS_ULAZ") & " " & _
                      AS_(korpa(i), "artikalID")
        End If
        upisano = upisano + 1
    Next i

    tx.CommitTx
    txStarted = False
    Set tx = Nothing
    Exit Function
EH:
    AgroUpisiUlaz = Poruka("AGRO_MSG_GRESKA_PRI_CUVANJU_2") & " " & Err.description
    upisano = 0
    LogErr SRC
    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
End Function

Private Function ProveriZaglavlje(ByVal korpa As Collection, _
                                  ByVal brojDok As String) As String
    If korpa Is Nothing Then
        ProveriZaglavlje = Poruka("AGROU_ERR_KORPA_PRAZNA")
        Exit Function
    End If
    If korpa.count = 0 Then
        ProveriZaglavlje = Poruka("AGROU_ERR_KORPA_PRAZNA")
        Exit Function
    End If
    If Len(Trim$(brojDok)) = 0 Then
        ProveriZaglavlje = Poruka("AGROU_ERR_NEMA_BROJA")
    End If
End Function

'=====================================================================
' SITNO
'=====================================================================
' Cele kolicine bez decimala, ostale sa najvise dve - isto kao legacy
' FormatKol. Postoji zato sto "5 kg" i "5,00 kg" u istoj poruci izgledaju kao
' dva razlicita podatka.
Public Function AgroFmtKol(ByVal v As Double) As String
    If v = Int(v) Then
        AgroFmtKol = CStr(CLng(v))
    Else
        AgroFmtKol = Format$(v, "0.##")
    End If
End Function

' Citaci reda korpe. Ime AS_ ima donju crtu jer je "As" rezervisana rec, pa bi
' se ime funkcije "As" case-insensitive poklopilo sa njom.
Private Function AS_(ByVal red As Object, ByVal k As String) As String
    On Error Resume Next
    If red.Exists(k) Then AS_ = Trim$(CStr(red(k)))
End Function

Private Function AD(ByVal red As Object, ByVal k As String) As Double
    On Error Resume Next
    If red.Exists(k) Then
        If IsNumeric(red(k)) Then AD = CDbl(red(k))
    End If
End Function

Private Function AB(ByVal red As Object, ByVal k As String) As Boolean
    On Error Resume Next
    If red.Exists(k) Then AB = CBool(red(k))
End Function
