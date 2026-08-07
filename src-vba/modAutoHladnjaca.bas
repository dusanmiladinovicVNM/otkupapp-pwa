Attribute VB_Name = "modAutoHladnjaca"
'Attribute VB_Name = "modAutoHladnjaca"
Option Explicit

' ============================================================
' modAutoHladnjaca (#3)
'
' Kada je otkupno mesto (stanica) HLADNJACA (tblStanice.JeHladnjaca = "Da")
' i toggle AUTO_PRIJEMNICA_HLADNJACA je ukljucen, iz otkupa se automatski
' kreira ceo lanac: otpremnica + zbirna + prijemnica.
'
' Kupac (hladnjaca) = MALINA_DEFAULT_KUPAC (KupacID iz tblKupci).
' Cena = cena iz otkupa. Svi ostali podaci dolaze iz otkupa.
'
' Broj prijemnice (modBrojevi.GenerateBrojPrijemnice): "1/ddmmyy" za prvu tog
' dana, pa "1/ddmmyy-2", "-3" ... -n. Dvoklasna prijemnica (Kl I + Kl II) nosi
' ISTI broj -- jedna prijemnica = jedan broj (kao i kod rucnog unosa).
' Broj otpremnice = broj zbirne = broj otkupnog dokumenta (malina konvencija);
' ako otkup nema broj, generise se HL-ddmmyy-hhnnss.
'
' Posle kreiranja, izvorni tblOtkup red dobija nazad: OtpremnicaID + BrojZbirne +
' VozacID (LinkOtkupRedNaDokument), da otkup ostane povezan sa dokumentima.
'
' Reuse: SaveOtpremnica_TX / SaveZbirna_TX / SavePrijemnica_TX (modDokumenta),
'        GenerateBrojPrijemnice (modBrojevi), FindRows / RequireUpdateCell /
'        clsTransaction (vezivanje nazad u otkup), LookupValue / GetConfigValue,
'        IsAutoPrijemnicaHladnjaca (modConfig).
' ============================================================

' Ispravka autohladnjace: posle storna otkupa-hladnjace ceo lanac
' (otpremnica+zbirna+prijemnica) je oboren, a palete su OSIROCENE. Kad operater
' izabere "Uneti ispravku", panel Otkupni blokovi zapamti broj te (stornirane)
' prijemnice ovde. SLEDECI auto-lanac (ponovni Unos otkupa) prevezuje osirocene
' palete na novu prijemnicu -- ista roba, bez ponovne paletizacije.
' Cross-form stanje (modOtkupBlok -> frmOtkup); cisti se na Terminate forme.
Private mPendingRelinkOldPrij As String

' ============================================================
' TEST SEAM -- simulacija pada pojedinacnog koraka lanca.
' Isti obrazac kao modPaletniList.SetPaletizeSkip: Public toggle nad Private
' modul-stanjem. Koristi ga modBusinessFlowProTests; produkcija ga NE poziva.
'
' Leak-proof po konstrukciji: AutoChainHladnjaca kopira vrednost u lokalnu i
' ODMAH je brise, pa vazi samo za JEDAN poziv -- ne moze da procuri u operativni
' rad ni ako test padne pre ciscenja, ni ako se izadje ranim Exit Function.
' Kodovi: "OTP" | "ZBR" | "PRJ" | "LINK" | "" (bez simulacije).
Private mTestFailStep As String

Public Sub ArmHladnjacaTestFail(ByVal stepKod As String)
    mTestFailStep = UCase$(Trim$(stepKod))
End Sub

Public Sub SetHladnjacaRelinkPending(ByVal oldPrijBroj As String)
    mPendingRelinkOldPrij = Trim$(oldPrijBroj)
End Sub

Public Function GetHladnjacaRelinkPending() As String
    GetHladnjacaRelinkPending = mPendingRelinkOldPrij
End Function

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

' Auto-lanac za hladnjacu. Poziva se posle uspesnog SaveOtkupMulti_TX (frmOtkup).
' Best-effort: greska NE sme da obori potvrdu otkupa. Vraca "" kad je lanac
' kompletan; inace tekst upozorenja (frmOtkup ga prikaze operateru).
Public Function AutoChainHladnjaca(ByVal datum As Date, ByVal stanicaID As String, _
                              ByVal vrsta As String, ByVal sorta As String, _
                              ByVal vozacID As String, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, _
                              ByVal kolicinaI As Double, ByVal cenaI As Double, _
                              ByVal hasKlasaII As Boolean, _
                              ByVal kolicinaII As Double, ByVal cenaII As Double, _
                              ByVal brDok As String, _
                              ByVal otkupIDs As String, _
                              Optional ByVal brutoKgI As Double = 0, _
                              Optional ByVal kolAmbII As Long = 0, _
                              Optional ByVal brutoKgII As Double = 0, _
                              Optional ByRef outBrPrij As String) As String
    On Error GoTo EH
    outBrPrij = ""       ' broj tek generisane prijemnice (za relink osirocenih paleta)

    ' Test seam: procitaj armirani korak i ODMAH ga potrosi -> vazi samo za ovaj
    ' poziv, bez obzira kojim izlazom se izadje. U produkciji je uvek "".
    Dim failStep As String
    failStep = mTestFailStep
    mTestFailStep = ""

    ' Vraca "" kad je lanac kompletan; inace tekst upozorenja za frmOtkup.
    ' Prati se SVAKI korak lanca po klasi (ne samo prijemnica): otpremnica, zbirna,
    ' prijemnica i veza nazad u otkup red. Svaki neuspeh ulazi u upozorenje --
    ' lanac se NE sme prijaviti kao uspesan ako je bilo koji korak pao.
    Dim failOtp As String, failZbr As String, failPrij As String, failLink As String

    If Not IsAutoPrijemnicaHladnjaca() Then Exit Function
    If Not IsHladnjacaStanica(stanicaID) Then Exit Function

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        ' Toggle ukljucen ali kupac-hladnjaca nije podesen -> ne mozemo dalje.
        LogError "modAutoHladnjaca.AutoChainHladnjaca", _
                 "AUTO_PRIJEMNICA_HLADNJACA ukljucen ali MALINA_DEFAULT_KUPAC prazan."
        AutoChainHladnjaca = "Auto-lanac hladnjace nije pokrenut: kupac-hladnjaca " & _
            "(MALINA_DEFAULT_KUPAC) nije pode" & ChrW(353) & "en u Podesavanjima."
        Exit Function
    End If

    Dim hladnjaca As String
    hladnjaca = Trim$(nz(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca"), ""))

    ' Vozac je obavezan na otpremnici/zbirnoj/prijemnici. U malina/hladnjaca
    ' konvenciji par-vozac ima VozacID == StanicaID. Ako otkup nema vozaca,
    ' koristimo stanicu kao vozaca (i napravimo par-vozaca ako je malina mod).
    If Len(Trim$(vozacID)) = 0 Then vozacID = stanicaID

    ' AUD-046: mirror konvencija (vozac == stanica) vazi SAMO ako par stvarno
    ' postoji (red u tblStanice I red u tblVozaci). Provera se radi kad god je
    ' vozacID == stanicaID -- bez obzira da li ga je fallback iznad postavio ili
    ' ga je pozivalac PROSLEDIO. (Guard vezan samo za prazan vozacID se
    ' zaobilazio prostim prosledjivanjem stanicaID-a, i lanac je i dalje mogao
    ' dobiti FK bez reda u tblVozaci.)
    If StrComp(Trim$(vozacID), Trim$(stanicaID), vbTextCompare) = 0 Then
        If Not IsManagedStationMirror(stanicaID) Then
            ' Jedan pokusaj kreiranja para; Ensure re-raise-uje, a ovde smo
            ' best-effort (greska ne sme da obori potvrdu otkupa) -> lokalno se
            ' guta, a odluku donosi re-provera ispod.
            On Error Resume Next
            EnsureVozacMirrorForStanica stanicaID, _
                Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"), "")), _
                "(hladnjaca)", ""
            On Error GoTo EH
        End If

        If Not IsManagedStationMirror(stanicaID) Then
            LogError "modAutoHladnjaca.AutoChainHladnjaca", _
                     "Nema vozac-mirror para za StanicaID=" & stanicaID & _
                     " -- auto-lanac se ne pokrece (izbegnut FK bez pokrica u " & TBL_VOZACI & ")."
            AutoChainHladnjaca = "Auto-lanac hladnjace nije pokrenut: za stanicu " & stanicaID & _
                " ne postoji par-voza" & ChrW(269) & " u " & TBL_VOZACI & "."
            Exit Function
        End If
    End If

    ' Broj otpremnice = broj otkupa (ili generisan). Zbirna se razdvaja: mirror-
    ' stanica (vozac==stanica) nosi "S" prefiks (S1/ddmmyy), otpremnica/otkup
    ' zadrzavaju svoj broj (1/ddmmyy) da se ne sudaraju sa realnim vozacem.
    Dim brOtp As String
    brOtp = Trim$(brDok)
    If Len(brOtp) = 0 Then brOtp = "HL-" & Format$(datum, "ddmmyy") & "-" & Format$(Now, "hhnnss")
    Dim brZbr As String
    brZbr = ApplyMirrorPrefix(vozacID, brOtp)

    ' Klasa I je opciona (kolicinaI = 0 -> kroz lanac ide samo Klasa II).
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)

    ' OtkupID-jevi za vezivanje nazad u tblOtkup. Format iz SaveOtkupMulti_TX:
    ' "resultI", "resultI + resultII", ili (samo II klasa) "resultII".
    Dim idI As String, idII As String
    If Len(Trim$(otkupIDs)) > 0 Then
        Dim parts() As String
        parts = Split(otkupIDs, " + ")
        If hasKlasaI Then
            idI = Trim$(parts(LBound(parts)))
            If UBound(parts) > LBound(parts) Then idII = Trim$(parts(LBound(parts) + 1))
        Else
            idII = Trim$(parts(LBound(parts)))
        End If
    End If

    ' Broj prijemnice: jedan po dokumentu; Klasa I i II nose ISTI broj
    ' (GenerateBrojPrijemnice, modBrojevi). Generise se i kad se unosi samo Klasa II.
    Dim brPrij As String
    brPrij = GenerateBrojPrijemnice(kupacID, datum)
    ' outBrPrij se NE izlaze ovde: caller ga koristi za relink osirocenih paleta, pa
    ' sme da pokazuje samo na STVARNO kreiranu prijemnicu. Izlaze se tek posle
    ' uspesnog SavePrijemnica_TX (dovoljna je jedna klasa -- Klasa I i II dele broj).

    ' Lanac je po klasi FAIL-FAST nizvodno: cim jedan korak padne, ostatak te klase
    ' se PRESKACE. Razlog nije cistoca -- SavePrijemnica_TX pise i u tblPaleta/
    ' tblPaletaStavka/tblFakturaStavke, pa bi prijemnica u vec palom lancu
    ' paletizovala broj i time zablokirala ponovni unos ("broj prijemnice je vec
    ' paletizovan"). Uz to, jedini recovery alat (BackfillPrijemniceHladnjaca) je
    ' usidren na OTPREMNICU -- za "zbirna/prijemnica bez otpremnice" nema backfilla.
    ' Neiskorisceni brPrij ne pravi rupu u nizu: GenerateBrojPrijemnice je racunat
    ' (MaxSeqFromTable + 1), ne rezervise broj.
    ' Klase su i dalje NEZAVISNE (saga): pad Klase I ne obara Klasu II.

    ' Klasa I (svoja kolicina ambalaze = kolAmb). Preskace se ako se unosi samo II.
    If hasKlasaI Then
        Dim otpID As String
        If failStep <> "OTP" Then _
            otpID = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                      kolicinaI, cenaI, tipAmb, kolAmb, KLASA_I, brutoKgI)
        If Len(otpID) = 0 Then
            failOtp = AddKlasa(failOtp, "I")
            GoTo KlasaIDone
        End If
        Dim zbrIDI As String
        If failStep <> "ZBR" Then _
            zbrIDI = SaveZbirna_TX(datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                                   kolicinaI, tipAmb, kolAmb, KLASA_I)
        If Len(zbrIDI) = 0 Then
            failZbr = AddKlasa(failZbr, "I")
            GoTo KlasaIDone
        End If
        Dim prjI As String
        If failStep <> "PRJ" Then _
            prjI = SavePrijemnica_TX(datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                                     kolicinaI, cenaI, tipAmb, kolAmb, 0, KLASA_I, brutoKgI)
        If Len(prjI) = 0 Then
            failPrij = AddKlasa(failPrij, "I")
        Else
            outBrPrij = brPrij       ' prijemnica postoji -> relink paleta je bezbedan
        End If
        ' Veza nazad u otkup red: OtpremnicaID + BrojZbirne + VozacID.
        Dim linkOkI As Boolean
        If failStep <> "LINK" Then _
            linkOkI = LinkOtkupRedNaDokument(idI, otpID, brZbr, vozacID)
        If Not linkOkI Then failLink = AddKlasa(failLink, "I")
    End If
KlasaIDone:

    ' Klasa II: zasebne gajbe (kolAmbII) -> ceo lanac kao Klasa I; ISTI broj
    ' prijemnice (brPrij) kao Klasa I (jedna prijemnica = jedan broj).
    If hasKlasaII And kolicinaII > 0 Then
        Dim otpID2 As String
        If failStep <> "OTP" Then _
            otpID2 = SaveOtpremnica_TX(datum, stanicaID, vozacID, brOtp, brZbr, vrsta, sorta, _
                                       kolicinaII, cenaII, tipAmb, kolAmbII, KLASA_II, brutoKgII)
        If Len(otpID2) = 0 Then
            failOtp = AddKlasa(failOtp, "II")
            GoTo KlasaIIDone
        End If
        Dim zbrIDII As String
        If failStep <> "ZBR" Then _
            zbrIDII = SaveZbirna_TX(datum, vozacID, brZbr, kupacID, hladnjaca, "", vrsta, sorta, _
                                    kolicinaII, tipAmb, kolAmbII, KLASA_II)
        If Len(zbrIDII) = 0 Then
            failZbr = AddKlasa(failZbr, "II")
            GoTo KlasaIIDone
        End If
        Dim prjII As String
        If failStep <> "PRJ" Then _
            prjII = SavePrijemnica_TX(datum, kupacID, vozacID, brPrij, brZbr, vrsta, sorta, _
                                      kolicinaII, cenaII, tipAmb, kolAmbII, 0, KLASA_II, brutoKgII)
        If Len(prjII) = 0 Then
            failPrij = AddKlasa(failPrij, "II")
        Else
            outBrPrij = brPrij
        End If
        Dim linkOkII As Boolean
        If failStep <> "LINK" Then _
            linkOkII = LinkOtkupRedNaDokument(idII, otpID2, brZbr, vozacID)
        If Not linkOkII Then failLink = AddKlasa(failLink, "II")
    End If
KlasaIIDone:

    ' Vidljivo upozorenje za frmOtkup: nabraja SAMO korake koji su stvarno pali
    ' (otpremnica / zbirna / prijemnica / veza sa otkupom), po klasi. Najcesci uzrok
    ' pada prijemnice: broj prijemnice je vec paletizovan (zaostala stavka u
    ' tblPaletaStavka posle ciscenja tblPrijemnica). Tehnicki razlog je u logu
    ' (Monitor_Event DOKUMENT_SAVE_FAIL / Save*_TX).
    Dim stavke As String
    If Len(failOtp) > 0 Then stavke = stavke & vbCrLf & _
        "- OTPREMNICA nije kreirana (Klasa " & failOtp & ")."
    If Len(failZbr) > 0 Then stavke = stavke & vbCrLf & _
        "- ZBIRNA nije kreirana (Klasa " & failZbr & ")."
    If Len(failPrij) > 0 Then stavke = stavke & vbCrLf & _
        "- PRIJEMNICA nije kreirana (Klasa " & failPrij & ")."
    If Len(failLink) > 0 Then stavke = stavke & vbCrLf & _
        "- OTKUP red nije povezan sa dokumentom (Klasa " & failLink & ")."

    If Len(stavke) > 0 Then
        AutoChainHladnjaca = "Otkup je sa" & ChrW(269) & "uvan, ali AUTO-LANAC hladnjace je NEPOTPUN:" & _
            stavke & vbCrLf & _
            "Najcesci uzrok pada prijemnice: broj prijemnice je ve" & ChrW(263) & " paletizovan " & _
            "(zaostala stavka u tblPaletaStavka). Detalji su u logu."
    End If
    Exit Function

EH:
    Dim eDesc As String: eDesc = Err.description
    LogErr "modAutoHladnjaca.AutoChainHladnjaca"
    AutoChainHladnjaca = "Auto-lanac hladnjace prekinut gre" & ChrW(353) & "kom: " & eDesc & " (vidi log)."
End Function

' Upisi vezu nazad u tblOtkup red(ove) za dati OtkupID:
'   OtpremnicaID + BrojZbirne (tek kreirani -> upisuju se),
'   VozacID samo ako je prazan (da se ne pregazi operaterov izbor).
' Reuse postojecih primitiva: FindRows / RequireUpdateCell / clsTransaction.
'
' Vraca True SAMO kad je veza stvarno i potpuno upisana; False kad red nije nadjen,
' kad je upis pao (EH radi RollbackTx + LogErr, bez re-raise -- lanac je
' best-effort) ili kad ulaz ne moze dati potpunu vezu. Pozivalac
' (AutoChainHladnjaca) neuspeh ukljucuje u upozorenje, da se otkup bez
' OtpremnicaID/BrojZbirne ne prijavi kao uspesan lanac.
'
' Prazan OtkupID NIJE legitimno "nema sta da se veze": SaveOtkupMulti_TX radi
' Err.Raise ako bilo koja aktivna klasa vrati "" i sastavlja "ID1 + ID2", pa se
' dovde stize sa ID-jem za svaku aktivnu klasu. Prazno => prekrsen ugovor ili
' pogresno parsiran rezultat -> prijavi se kao neuspeh veze.
Private Function LinkOtkupRedNaDokument(ByVal otkupID As String, ByVal otpID As String, _
                                        ByVal brZbr As String, ByVal vozacID As String) As Boolean
    Const SRC As String = "modAutoHladnjaca.LinkOtkupRedNaDokument"
    Dim tx As clsTransaction
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function     ' prekrsen ugovor (vidi zaglavlje)
    ' Parcijalna veza nije uspeh: bez OtpremnicaID ili bez BrojZbirne otkup red
    ' ostaje polu-povezan. Sa fail-fast lancem se dovde ne stize sa praznim otpID
    ' (otpremnica je preduslov za nizvodne korake) -- guard je zastita ako se
    ' redosled koraka ikad promeni.
    If Len(otpID) = 0 Or Len(brZbr) = 0 Then Exit Function

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows Is Nothing Then Exit Function  ' red nije nadjen -> veza NIJE upisana
    If rows.count = 0 Then Exit Function

    Dim curVoz As String
    curVoz = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_VOZAC), ""))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim k As Long, r As Long
    For k = 1 To rows.count
        r = rows(k)
        If Len(otpID) > 0 Then RequireUpdateCell TBL_OTKUP, r, COL_OTK_OTPREMNICA_ID, otpID, SRC
        If Len(otpID) > 0 Then SetOtkupBrojOtpremnice r, otpID
        If Len(brZbr) > 0 Then RequireUpdateCell TBL_OTKUP, r, COL_OTK_BROJ_ZBIRNE, brZbr, SRC
        If Len(vozacID) > 0 And curVoz = "" Then _
            RequireUpdateCell TBL_OTKUP, r, COL_OTK_VOZAC, vozacID, SRC
    Next k

    tx.CommitTx
    Set tx = Nothing
    LinkOtkupRedNaDokument = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
End Function

' Nabrajanje klasa u poruci o neuspehu: "I", "II" ili "I i II".
Private Function AddKlasa(ByVal acc As String, ByVal klasa As String) As String
    If Len(acc) = 0 Then
        AddKlasa = klasa
    Else
        AddKlasa = acc & " i " & klasa
    End If
End Function

' ============================================================
' Jednokratni backfill: kreira prijemnice za vec postojece otpremnice
' hladnjaca-lanca kojima prijemnica NEDOSTAJE (npr. posle reseta tblPrijemnica
' kada je auto-lanac pravio otpremnicu+zbirnu, a prijemnica padala).
'
' Anchor = otpremnica (jedina nosi sve sto prijemnici treba: cena + bruto +
' kol. ambalaze po klasi). Idempotentno: preskace (BrojZbirne + Klasa) za koji
' vec postoji ne-stornirana prijemnica. Klasa I i II istog dokumenta (isti
' BrojZbirne) dele JEDAN broj prijemnice (kao i live auto-lanac).
'
' VAZNO (pre pokretanja): ako je tblPrijemnica praznjena, obrisi orphan redove u
' tblPaleta i tblPaletaStavka -> inace paletizacija u SavePrijemnica_TX puca
' ("vec paletizovana") i prijemnica se ne kreira. Macro tada NE rusi nista
' (svaka prijemnica je atomicna): samo prijavi koliko je palo.
'
' Pokretanje: Alt+F8 -> BackfillPrijemniceHladnjaca.
' Reuse: SavePrijemnica_TX, GenerateBrojPrijemnice, IsHladnjacaStanica,
'        GetTableData / RequireColumnIndex / GetColumnIndex.
' ============================================================
Public Sub BackfillPrijemniceHladnjaca()
    BackfillPrijemniceHladnjacaCore False
End Sub

' Jezgro backfill-a. TEST SEAM: silent:=True preskace sve MsgBox-ove i vraca
' brojace kroz outOk/outFail, pa je makro pozivljiv iz modBusinessFlowProTests
' bez interaktivnih prompta. Javni ulaz iznad je namerno BEZ argumenata --
' Sub sa parametrima se ne pojavljuje u Alt+F8 listi makroa.
'
' onlyBrojZbirne ogranicava obradu na JEDAN dokument. Bez toga backfill skenira
' SVE hladnjaca-otpremnice u svesci, pa bi pokretanje test suite-a nad realnim
' klijentskim fajlom tiho kreiralo prijemnice za prave dokumente (a posto testovi
' privremeno preusmere MALINA_DEFAULT_KUPAC na test-kupca, jos gore: pod pogresnim
' kupcem i mimo `have` provere). Testovi zato UVEK salju svoj BrojZbirne.
' Operativni poziv (wrapper) ga ostavlja prazan = ponasanje kao do sada.
Public Sub BackfillPrijemniceHladnjacaCore(ByVal silent As Boolean, _
                                           Optional ByVal onlyBrojZbirne As String = "", _
                                           Optional ByRef outOk As Long, _
                                           Optional ByRef outFail As Long)
    Const SRC As String = "modAutoHladnjaca.BackfillPrijemniceHladnjaca"
    On Error GoTo EH
    outOk = 0
    outFail = 0

    Dim kupacID As String
    kupacID = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(kupacID) = 0 Then
        If Not silent Then _
            MsgBox "MALINA_DEFAULT_KUPAC (kupac-hladnjaca) nije pode" & ChrW(353) & "en. " & _
                   "Podesi ga u Podesavanjima pa pokreni ponovo.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim otp As Variant: otp = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otp) Then
        If Not silent Then MsgBox "Nema otpremnica za backfill.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Kolone otpremnice (izvor podataka).
    Dim cDat As Long, cSta As Long, cVoz As Long, cZbr As Long, cVrs As Long
    Dim cSor As Long, cKol As Long, cCen As Long, cTip As Long, cAmb As Long
    Dim cKla As Long, cBru As Long, cStorno As Long
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cSta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, SRC)
    cVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cVrs = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    cSor = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cCen = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, SRC)
    cTip = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, SRC)
    cAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, SRC)
    cKla = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    cBru = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BRUTO)        ' opciono
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)    ' opciono

    ' Postojece (ne-stornirane) prijemnice -> set kljuceva (BrojZbirne|Klasa).
    ' Uz to, mapa BrojZbirne -> vec dodeljeni BrojPrijemnice: ako jedna klasa
    ' dokumenta vec ima prijemnicu, druga klasa mora dobiti ISTI broj (jedna
    ' prijemnica = jedan broj, kao i u zivom auto-lancu), a ne novi.
    '
    ' OBE mape su ogranicene na hladnjaca-kupca (kupacID). Broj prijemnice je
    ' PER-KUPAC (GenerateBrojPrijemnice -> MaxSeqFromTable filtriran po
    ' COL_PRJ_KUPAC), pa bi nasledjivanje broja iz prijemnice drugog kupca ubacilo
    ' broj iz tudjeg niza u hladnjaca niz. Isto vazi i za idempotentnost: prijemnica
    ' drugog kupca sa istim BrojZbirne ne sme da preskoci legitimnog kandidata.
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    Dim brByZbr As Object: Set brByZbr = CreateObject("Scripting.Dictionary")
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prj) Then
        Dim pZbr As Long, pKla As Long, pStorno As Long, pBroj As Long, pKup As Long
        pZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
        pKla = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
        pBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
        pKup = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, SRC)
        pStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
        Dim pr As Long
        For pr = 1 To UBound(prj, 1)
            If pStorno = 0 Or UCase$(Trim$(CStr(prj(pr, pStorno)))) <> "DA" Then
                If StrComp(Trim$(CStr(prj(pr, pKup))), kupacID, vbTextCompare) = 0 Then
                    have(KeyZbrKlasa(CStr(prj(pr, pZbr)), CStr(prj(pr, pKla)))) = True
                    Dim pzb As String: pzb = Trim$(CStr(prj(pr, pZbr)))
                    Dim pbr As String: pbr = Trim$(CStr(prj(pr, pBroj)))
                    If Len(pzb) > 0 And Len(pbr) > 0 Then
                        If Not brByZbr.Exists(pzb) Then brByZbr(pzb) = pbr
                    End If
                End If
            End If
        Next pr
    End If

    ' Kandidati: ne-stornirane hladnjaca-otpremnice bez (BrojZbirne|Klasa) prijemnice.
    Dim cand As Collection: Set cand = New Collection
    Dim r As Long
    For r = 1 To UBound(otp, 1)
        If (cStorno = 0) Or (UCase$(Trim$(CStr(otp(r, cStorno)))) <> "DA") Then
            Dim sid As String: sid = Trim$(CStr(otp(r, cSta)))
            Dim zbr As String: zbr = Trim$(CStr(otp(r, cZbr)))
            Dim kla As String: kla = Trim$(CStr(otp(r, cKla)))
            If Len(zbr) > 0 And IsHladnjacaStanica(sid) Then
                Dim uOpsegu As Boolean
                uOpsegu = (Len(onlyBrojZbirne) = 0)
                If Not uOpsegu Then _
                    uOpsegu = (StrComp(zbr, Trim$(onlyBrojZbirne), vbTextCompare) = 0)
                If uOpsegu Then
                    If Not have.Exists(KeyZbrKlasa(zbr, kla)) Then cand.Add r
                End If
            End If
        End If
    Next r

    If cand.count = 0 Then
        If Not silent Then _
            MsgBox "Nema hladnjaca-otpremnica bez prijemnice. Nista za backfill.", _
                   vbInformation, APP_NAME
        Exit Sub
    End If

    If Not silent Then
        If MsgBox("Pronadjeno " & cand.count & " otpremnica (hladnjaca) bez prijemnice." & _
                  vbCrLf & "Kreirati prijemnice za njih sada?" & vbCrLf & vbCrLf & _
                  "NAPOMENA: ako je tblPrijemnica ranije praznjena, prvo ocisti orphan " & _
                  "redove u tblPaleta i tblPaletaStavka (inace paletizacija puca).", _
                  vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
    End If

    ' Jedan broj prijemnice po dokumentu (BrojZbirne) -> Klasa I i II dele broj.
    ' brByZbr je vec pre-punjen brojevima postojecih prijemnica (gore), pa se broj
    ' preuzima i kad je sestrinska klasa backfill-ovana u ranijem prolazu.
    Dim ok As Long, fail As Long, i As Long
    For i = 1 To cand.count
        r = cand(i)
        If Not IsDate(otp(r, cDat)) Then
            fail = fail + 1
            GoTo ContinueLoop
        End If
        Dim dDat As Date: dDat = CDate(otp(r, cDat))
        Dim zb As String: zb = Trim$(CStr(otp(r, cZbr)))

        Dim brPrij As String
        If brByZbr.Exists(zb) Then
            brPrij = CStr(brByZbr(zb))
        Else
            brPrij = GenerateBrojPrijemnice(kupacID, dDat)
            brByZbr(zb) = brPrij
        End If

        Dim bru As Double: bru = 0
        If cBru > 0 Then bru = AsDbl(otp(r, cBru))

        Dim res As String
        res = SavePrijemnica_TX(dDat, kupacID, Trim$(CStr(otp(r, cVoz))), brPrij, zb, _
                  Trim$(CStr(otp(r, cVrs))), Trim$(CStr(otp(r, cSor))), _
                  AsDbl(otp(r, cKol)), AsDbl(otp(r, cCen)), Trim$(CStr(otp(r, cTip))), _
                  AsLng(otp(r, cAmb)), 0, ClassOrDefault(otp(r, cKla)), bru)
        If Len(res) > 0 Then ok = ok + 1 Else fail = fail + 1
ContinueLoop:
    Next i

    LogInfo SRC, "Backfill prijemnice: ok=" & ok & " fail=" & fail
    outOk = ok
    outFail = fail
    If Not silent Then _
        MsgBox "Backfill zavr" & ChrW(353) & "en." & vbCrLf & _
               "Kreirano prijemnica: " & ok & vbCrLf & _
               "Neuspesno: " & fail & _
               IIf(fail > 0, vbCrLf & "(vidi log; najcesce orphan paleta -> ocisti " & _
                                  "tblPaleta/tblPaletaStavka pa pokreni ponovo)", ""), _
               IIf(fail > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub
EH:
    LogErr SRC
    If Not silent Then _
        MsgBox "Gre" & ChrW(353) & "ka u backfill-u: " & Err.description, vbCritical, APP_NAME
End Sub

Private Function KeyZbrKlasa(ByVal zbr As String, ByVal klasa As String) As String
    KeyZbrKlasa = Trim$(zbr) & "|" & Trim$(klasa)
End Function

' Javni Nz (modHelpers) je String-tipa; ovde trebaju brojevi iz numerickih
' celija (NzD/NzL u modPaletniList su Private -> nedostupni). Minimalne lokalne
' koercije, IsNumeric-bezbedne (prazna/ne-broj celija -> 0).
Private Function AsDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then AsDbl = CDbl(v)
End Function

Private Function AsLng(ByVal v As Variant) As Long
    If IsNumeric(v) Then AsLng = CLng(v)
End Function

' Pravilo default klase zivi u modHelpers.KlasaOrDefault (deli ga i report sloj
' preko ZbirnaStavkaKljuc) -- ovde ostaje samo lokalno ime, bez druge kopije.
Private Function ClassOrDefault(ByVal v As Variant) As String
    ClassOrDefault = KlasaOrDefault(v)
End Function

