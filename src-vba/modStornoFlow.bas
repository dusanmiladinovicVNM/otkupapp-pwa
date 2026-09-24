Attribute VB_Name = "modStornoFlow"
Option Explicit

' ============================================================
' modStornoFlow - centralni orkestrator storno/ispravke (business sloj)
'
' Poslovni princip: storno NIJE Da/Ne. Prvo se bira STA storno poslovno znaci:
'   ISPRAVKA_ODMAH        - pogresan unos, isti fizicki dogadjaj -> storno stari,
'                           otvori novi, zapamti staro->novo, po snimanju relink
'                           + rekalkulisi zbirnu.
'   DUPLI_FANTOM          - dokument nikad nije trebalo da postoji -> skini/odvezi
'                           posledice bez naslednika (blokovi izgubljeni / otpremnice
'                           "ceka zbirnu" / saldo se ne duplira).
'   PONISTENJE_BEZ_ZAMENE - fizicki tok se ponistava, nema novog -> blokada ako
'                           postoje zavisni dokumenti (osim uz svesnu potvrdu).
'   RESI_KASNIJE          - persistent recovery zapis (ne samo MsgBox).
'
' Ovaj modul NE sadrzi MsgBox (UI je u frmDokumenta). Vraca rezultat kao
' Scripting.Dictionary: success/blocked/needsForm/correctionID/mode/message/valid.
'
' REUSE (bez dupliranja):
'   Storno: modStorno.StornoOtpremnica (core), StornoZbirna_TX,
'           StornoPrijemnicaByBroj_TX, StornoOMKoopByBrDok_TX, LookupActiveID.
'   Relink: modDokumenta.ReassignOtkupToOtpremnica_TX (prijemnica: Oporavak).
'   Zbirna: stavke se citaju KANONSKI (modDokumenta.ZbirnaPoKlasi). Stari
'           invarijantni racun po BrojZbirne je obrisan u S4-3a -- pod kanonom
'           su mu obe strane bile prazne, pa je uvek javljao "OK".
'   Context: modStornoContext.*  Ambalaza saldo vec iskljucuje stornirano.
' ============================================================

Private Const MOD_NAME As String = "modStornoFlow"

' Doc-type kljucevi za framework (interno; UI mapira sa combo vrednostima).
Public Const FLOW_DOC_OTPREMNICA As String = "Otpremnica"
Public Const FLOW_DOC_ZBIRNA As String = "Zbirna"
Public Const FLOW_DOC_REVERS As String = "Revers"
Public Const FLOW_DOC_PRIJEMNICA As String = "Prijemnica"

' ============================================================
' PREVIEW - multiline tekst za dijalog (UI ga prikaze u MsgBox-u).
' ============================================================
Public Function BuildStornoPreview(ByVal docType As String, ByVal broj As String, _
                                   Optional ByVal dokumentTip As String = "", _
                                   Optional ByVal docID As String = "") As String
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_ZBIRNA:      BuildStornoPreview = PreviewZbirna(broj, docID)
        Case FLOW_DOC_REVERS:      BuildStornoPreview = PreviewRevers(broj, dokumentTip, docID)
        Case FLOW_DOC_PRIJEMNICA:  BuildStornoPreview = PreviewPrijemnica(broj, docID)
        Case Else:                 BuildStornoPreview = "Dokument: " & docType & " " & broj
    End Select
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske). Pregled je
    ' kapija pre potvrde: operater mora da vidi ZASTO ga nema, ne samo da ga nema.
    Dim errDesc As String: errDesc = Err.description
    LogErr MOD_NAME & ".BuildStornoPreview"
    BuildStornoPreview = "Pregled nije dostupan (greska: " & errDesc & "). Dokument: " & docType & " " & broj
End Function

' Stavke JEDNE zbirne kao tekst za uvid: "I 400 kg / 20 amb; II 100 kg / 5 amb".
' Cita kanonski citalac po ID-u dokumenta -- isti koji koriste liste i stampa.
' Prazno clanstvo nije greska ovde: uvid pred storno sme da stoji nad nacrtom.
Private Function ZbirnaStavkeTekst(ByVal zbirnaID As String) As String
    Dim po As Object, k As Variant, m As String
    On Error GoTo EH

    If Len(Trim$(zbirnaID)) = 0 Then
        ZbirnaStavkeTekst = "(identitet nije razresen)"
        Exit Function
    End If

    Set po = modDokumenta.ZbirnaPoKlasi(zbirnaID)
    For Each k In po.Keys
        If Len(m) > 0 Then m = m & "; "
        m = m & CStr(k) & " " & Fmt(po(k)(0)) & " kg / " & CStr(po(k)(1)) & " amb"
    Next k
    If Len(m) = 0 Then m = "(nema stavki)"

    ZbirnaStavkeTekst = m
    Exit Function
EH:
    ' Greska se IMENUJE, ne guta: prazan tekst bi izgledao kao "nema stavki",
    ' a to je druga tvrdnja.
    ZbirnaStavkeTekst = "(stavke nisu citljive: " & Err.description & ")"
End Function

Private Function PreviewZbirna(ByVal broj As String, _
Optional ByVal docID As String = "") As String
    Dim s As Object: Set s = ScanZbirna(broj, docID)
    Dim m As String
    m = "ZBIRNA " & broj & vbCrLf
    m = m & "Aktivne otpremnice: " & CStr(s("otpCount")) & vbCrLf

    ' KG DOLAZE SA STAVKI, NE IZ ZAGLAVLJA (S4-3a).
    '
    ' Ovde je do sada stajao invarijantni racun: zbir otpremnica po BrojZbirne
    ' naspram zaglavlja tblZbirna. Pod kanonom pisac NE upisuje ni jedno ni
    ' drugo -- clanstvo je zapis (tblZbirnaIzvori), sadrzaj je na stavkama -- pa
    ' su obe strane bile nule i red "Invarijanta: OK" je stajao nad svakom
    ' zbirnom. Uvid koji uvek kaze OK je gori od uvida kog nema.
    m = m & "Stavke zbirne: " & ZbirnaStavkeTekst(CStr(s("zbrID"))) & vbCrLf
    m = m & "Prijemnica: " & YesNo(CBool(s("hasPrijemnica"))) & " (" & CStr(s("prijCount")) & ")" & vbCrLf
    m = m & "Paletizovano: " & YesNo(CBool(s("hasPalete"))) & " (" & CStr(s("paleteCount")) & ")"
    PreviewZbirna = m
End Function

' Pregled reversa pre potvrde -- za KLIKNUTI revers (docID = AmbID), ne za sve redove
' broja: isti broj legalno nosi revers druge stanice ili drugog dana, pa bi pregled
' po (broj, tip) pokazao tudj dokument bas tamo gde operater odlucuje. ReversID i
' noge bira isti kod kao pisac (ReversIDRazresi, ReversRedoviRID).
Private Function PreviewRevers(ByVal broj As String, ByVal dokumentTip As String, _
                               Optional ByVal docID As String = "") As String
    Dim s As Object: Set s = ScanRevers(broj, dokumentTip, docID)
    Dim m As String
    m = "REVERS " & broj & " [" & dokumentTip & "]" & vbCrLf
    If Len(CStr(s("razlog"))) > 0 Then
        PreviewRevers = m & "(revers nije jednoznacan: " & CStr(s("razlog")) & ")"
        Exit Function
    End If
    If Not CBool(s("exists")) Then
        PreviewRevers = m & "(nije pronadjen aktivan revers)"
        Exit Function
    End If
    m = m & "Stanica: " & CStr(s("stanica")) & " | dan: " & _
            Format$(CDate(CLng(s("dan"))), "dd.mm.yyyy") & vbCrLf
    m = m & "Kooperant/Stanica: " & CStr(s("entitet")) & vbCrLf
    m = m & "Tip ambalaze: " & CStr(s("tip")) & vbCrLf
    m = m & "Kolicina: " & CStr(s("kolicina")) & " (knjiznih redova: " & CStr(s("redova")) & ")" & vbCrLf
    m = m & "Smer: " & CStr(s("smer")) & vbCrLf
    m = m & "Uticaj na saldo: storno iskljucuje ovaj revers iz salda (bez duple stavke)."
    PreviewRevers = m
End Function

Private Function PreviewPrijemnica(ByVal broj As String, _
Optional ByVal docID As String = "") As String
    Dim s As Object: Set s = ScanPrijemnica(broj, docID)
    Dim m As String
    m = "PRIJEMNICA " & broj & vbCrLf
    If Not CBool(s("exists")) Then
        PreviewPrijemnica = m & "(nije pronadjena aktivna prijemnica)"
        Exit Function
    End If
    m = m & "Broj zbirne: " & IIf(Len(CStr(s("brojZbirne"))) > 0, CStr(s("brojZbirne")), "(nema)") & vbCrLf
    m = m & "Fakturisana: " & YesNo(CBool(s("fakturisano"))) & _
            IIf(CBool(s("fakturisano")), " (faktura/stavke se oslobadjaju)", "") & vbCrLf
    m = m & "Palete preko prijemnice: " & YesNo(CBool(s("hasPalete"))) & _
            " (" & CStr(s("paleteCount")) & ")" & vbCrLf
    m = m & "Otkupni blokovi (preko zbirne, samostalni): " & CStr(s("blockCount")) & vbCrLf
    m = m & "Rizik ambalaza: storno vraca ambalazu ove prijemnice (auto)."
    PreviewPrijemnica = m
End Function

' ============================================================
' CHAIN FLAGS - UI koristi da odluci koje opcije nudi / da li je PONISTENJE
' blokirano. Vraca dict: hasDependents, dependentsText, canPonistenjeClean.
' ============================================================
' strict = citanje koje NE SME da propadne u tisini. Prazan rezultat tada znaci
' iskljucivo "uspesno sam proverio i nema ih"; sve ostalo (schema drift,
' necitljiva tabela, greska u prolazu) DIZE gresku. Trazi ga samo
' modStornoImpact: model uvida se posle oznacava kao valid, a "ne znam" ne sme
' da prodje kao "nema". Podrazumevano False -- zatecenim pozivaocima (legacy
' frmDokumenta, paneli) ponasanje ostaje isto.
Public Function GetChainFlags(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal dokumentTip As String = "", _
                              Optional ByVal docID As String = "", _
                              Optional ByVal strict As Boolean = False) As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")
    Set GetChainFlags = r
    On Error GoTo EH
    r("hasDependents") = False
    r("dependentsText") = ""
    r("canPonistenjeClean") = True

    Select Case docType
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj, docID, strict)
            Dim depz As Boolean
            depz = (CLng(sz("otpCount")) > 0) Or CBool(sz("hasPrijemnica")) Or CBool(sz("hasPalete"))
            r("hasDependents") = depz
            r("canPonistenjeClean") = Not depz
            r("dependentsText") = "otpremnice=" & CStr(sz("otpCount")) & _
                ", prijemnica=" & YesNo(CBool(sz("hasPrijemnica"))) & _
                ", palete=" & YesNo(CBool(sz("hasPalete")))
        Case FLOW_DOC_REVERS
            r("hasDependents") = False       ' revers je list (nema nizvodni lanac)
            r("canPonistenjeClean") = True
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj, docID, strict)
            Dim depp As Boolean
            depp = CBool(sp("hasPalete")) Or CBool(sp("fakturisano"))
            r("hasDependents") = depp
            r("canPonistenjeClean") = Not depp
            r("dependentsText") = "palete=" & YesNo(CBool(sp("hasPalete"))) & _
                ", fakturisana=" & YesNo(CBool(sp("fakturisano")))
    End Select
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".GetChainFlags"
    If strict Then Err.Raise errNum, MOD_NAME & ".GetChainFlags", errDesc
End Function

' ============================================================
' PONISTENJE posledice: PUN spisak zavisnih dokumenata koje poistenje gasi.
' UI ga prikaze PRE nego sto bilo sta uradi (to je ono sto PONISTENJE cini
' razlicitim od DUPLI -- DUPLI tiho pocisti, PONISTENJE prvo pokaze ceo lanac).
' ============================================================
Public Function BuildPonistenjePosledice(ByVal docType As String, ByVal broj As String, _
                                         Optional ByVal dokumentTip As String = "", _
                                         Optional ByVal docID As String = "") As String
    On Error GoTo EH
    Dim m As String
    Select Case docType
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj, docID)
            Dim owz As Boolean: owz = ZbirnaOwnsExternalChain(broj)
            m = "PONISTENJE zbirne " & broj & " gasi interni tok (STORNO)." & vbCrLf & "Pogodjeno:" & vbCrLf
            m = m & " - aktivne otpremnice (storniraju se): " & CStr(sz("otpCount")) & vbCrLf
            m = m & " - prijemnice: " & CStr(sz("prijCount")) & _
                    IIf(owz, " (storniraju se)", " (eksterni kupac -> NETAKNUTE)") & vbCrLf
            m = m & " - paletne stavke: " & CStr(sz("paleteCount")) & _
                    IIf(owz, " (skidaju se sa paleta; prazna paleta stornirana)", " (NETAKNUTE)") & vbCrLf
            m = m & " - otkupni blokovi (OSLOBADJAJU se za reveze, NE storniraju)"
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj, docID)
            m = "PONISTENJE prijemnice " & broj & " gasi CEO tok (STORNO)." & vbCrLf & "Pogodjeno:" & vbCrLf
            m = m & " - zbirna: " & IIf(Len(CStr(sp("brojZbirne"))) > 0, CStr(sp("brojZbirne")), "(nema)") & _
                    " (rekalk; storno ako kg padne na 0)" & vbCrLf
            m = m & " - otpremnice te zbirne (storniraju se): " & CStr(sp("otpCount")) & vbCrLf
            m = m & " - faktura: " & IIf(CBool(sp("fakturisano")), "oslobadja se (stavke osirocene)", "(nije fakturisana)") & vbCrLf
            m = m & " - paletne stavke: " & CStr(sp("paleteCount")) & _
                    IIf(CBool(sp("hasPalete")), " (skidaju se sa paleta; prazna paleta stornirana)", "") & vbCrLf
            m = m & " - otkupni blokovi (samostalni; NE diraju se osim cekiranih za storno): " & CStr(sp("blockCount"))
        Case Else
            m = "PONISTENJE dokumenta " & broj & "."
    End Select
    BuildPonistenjePosledice = m
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildPonistenjePosledice"
    BuildPonistenjePosledice = "PONISTENJE dokumenta " & broj & " (spisak posledica nedostupan)."
End Function

' ============================================================
' SMART TRIGGER: da li storno TRAZI poslovni dijalog (4 moda)?
' Da SAMO kad postoji NIZVODNI tok koji trazi odluku operatera (prijemnica ili
' palete). Inace je obican storno + tiha rekalkulacija/odvezivanje dovoljan
' (motor cuva invarijantu bez ceremonije). Revers je list -> nikad dijalog.
' ============================================================
Public Function CorrectionNeedsDialog(ByVal docType As String, ByVal broj As String, _
                                      Optional ByVal dokumentTip As String = "", _
                                      Optional ByVal docID As String = "") As Boolean
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj, docID)
            CorrectionNeedsDialog = CBool(sz("hasPrijemnica")) Or CBool(sz("hasPalete"))
        Case FLOW_DOC_REVERS
            CorrectionNeedsDialog = False
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj, docID)
            ' Panel (pun dijalog) kad ima palete, fakture ILI otkupnih blokova (multiselect).
            CorrectionNeedsDialog = CBool(sp("hasPalete")) Or CBool(sp("fakturisano")) _
                                    Or (CLng(sp("blockCount")) > 0)
    End Select
    Exit Function
EH:
    ' Na gresku budi konzervativan -> ponudi pun dijalog.
    LogErr MOD_NAME & ".CorrectionNeedsDialog"
    CorrectionNeedsDialog = True
End Function

' ============================================================
' SIMPLE STORNO (bez dijaloga/context-a): obican storno + tiha zastita invarijante.
' Koristi se kad CorrectionNeedsDialog = False. Reuse postojecih storno funkcija.
' ============================================================

' Zbirna: storno + odvezivanje otpremnica ("ceka zbirnu") u JEDNOJ transakciji ->
' ne ostaje zbirna koja nije zbir svojih otpremnica (nema tihog mismatch-a).
Public Function RunSimpleStornoZbirna(ByVal broj As String, _
                                       Optional ByVal docID As String = "") As Object
    Const SRC As String = MOD_NAME & ".RunSimpleStornoZbirna"
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoZbirna = r
    On Error GoTo EH
    broj = Trim$(broj)
    If Not ZbirnaPostoji(broj) Then r("message") = "Aktivna zbirna nije pronadjena: " & broj: Exit Function

    ' Atomarno (jedna TX): storno zbirne + odvezivanje otpremnica ("ceka zbirnu").
    Dim det As Long
    If Not StornoZbirnaIDetach_TX(broj, det, docID) Then r("message") = "Storno zbirne nije uspeo.": Exit Function

    r("success") = True
    r("message") = "Zbirna " & broj & " stornirana." & _
                   IIf(det > 0, " Otpremnice vracene u 'ceka zbirnu': " & det & ".", "")
    MonitorSimple "Zbirna", broj, CStr(r("message"))
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr SRC
    r("message") = "Greska: " & errDescEH
End Function

' Revers: obican storno (saldo vec iskljucuje stornirano -> auto koreguje).
' ambID = identitet reda (iz njega se cita ReversID); bez njega ReversID po (broj,
' tip) mora biti jednoznacan.
Public Function RunSimpleStornoRevers(ByVal brDok As String, ByVal dokumentTip As String, _
                                      Optional ByVal ambID As String = "") As Object
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoRevers = r
    On Error GoTo EH
    brDok = Trim$(brDok)
    Dim revID As String, revRaz As String
    revRaz = ReversIDRazresi(ambID, brDok, dokumentTip, revID, False)
    If Len(revRaz) = 0 Then revRaz = ReversIDGranica(revID)
    If Len(revRaz) > 0 Then r("message") = revRaz: Exit Function
    If ReversRedoviRID(revID, False).count = 0 Then
        r("message") = "Aktivan revers nije pronadjen: " & brDok & " [" & dokumentTip & "]"
        Exit Function
    End If
    If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip, ambID) Then r("message") = "Storno reversa nije uspeo.": Exit Function
    r("success") = True
    r("message") = "Revers " & brDok & " storniran. Saldo azuriran (bez duple/kontra stavke)."
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr MOD_NAME & ".RunSimpleStornoRevers"
    r("message") = "Greska: " & errDescEH
End Function

' Prijemnica: obican storno (nema paleta/fakture/blokova -> nema odluke). Reuse
' StornoPrijemnicaByBroj_TX (oslobadja fakturu + ambalazu ako ih ima).
Public Function RunSimpleStornoPrijemnica(ByVal broj As String, _
Optional ByVal docID As String = "") As Object
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoPrijemnica = r
    On Error GoTo EH
    broj = Trim$(broj)
    Dim s As Object: Set s = ScanPrijemnica(broj, docID)
    If Not CBool(s("exists")) Then r("message") = "Aktivna prijemnica nije pronadjena: " & broj: Exit Function
    If Not StornoPrijemnicaByBroj_TX(broj, docID) Then r("message") = "Storno prijemnice nije uspeo.": Exit Function
    r("success") = True
    r("message") = "Prijemnica " & broj & " stornirana."
    MonitorSimple "Prijemnica", broj, CStr(r("message"))
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr MOD_NAME & ".RunSimpleStornoPrijemnica"
    r("message") = "Greska: " & errDescEH
End Function

Private Sub MonitorSimple(ByVal entityType As String, ByVal id As String, ByVal msg As String)
    On Error Resume Next
    Monitor_Event eventType:="STORNO_SIMPLE_" & UCase$(entityType), severity:="INFO", _
        message:=msg, moduleName:=MOD_NAME, procedureName:="RunSimpleStorno", _
        entityType:=entityType, entityID:=id, correlationId:=id
End Sub

' ============================================================
' ZBIRNA - dispatch po modu
' ============================================================
Public Function RunZbirnaCorrection(ByVal broj As String, ByVal mode As String, _
                                    Optional ByVal forceConfirm As Boolean = False, _
                                    Optional ByVal docID As String = "") As Object
    Const SRC As String = MOD_NAME & ".RunZbirnaCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunZbirnaCorrection = r
    On Error GoTo EH

    broj = Trim$(broj)
    If Not ZbirnaPostoji(broj) Then
        r("message") = "Aktivna zbirna nije pronadjena: " & broj
        Exit Function
    End If
    Dim s As Object: Set s = ScanZbirna(broj, docID)

    ' PK aktivne zbirne PRE storna -> OldDocID u context-u. Prefill ispravke polazi
    ' od njega (broj dokumenta nije globalno jedinstven identitet).
    ' MODOVI KOJI DIRAJU DECU STAJU PRE ICEGA kad broj nije jednoznacan.
    '
    ' Zaglavlje se moze stornirati po generaciji -- to je tacno. Ali DUPLI i
    ' PONISTENJE diraju decu po BrojZbirne, jer drugog kljuca u semi nema. Kod
    ' dva aktivna dokumenta istog broja to znaci: storniram TACNO svoje
    ' zaglavlje, pa TUDJOJ zbirni odnesem decu. Tiho.
    '
    ' OD S4-3c SCOPED IZBOR POSTOJI, pa kapija mora da ga vidi (review #384).
    ' Racuna se ISTIM telom koje akter koristi, sa decom BAS te operacije --
    ' inace kapija odbija ono sto primitiv ume bezbedno da uradi.
    Dim scopeID As String, razZC As String
    If mode <> SV_MODE_RESI_KASNIJE Then
        razZC = ZbirnaScopeRazlog(broj, docID, _
                                  (mode = SV_MODE_PONISTENJE) And ZbirnaOwnsExternalChain(broj), _
                                  scopeID)
        If Len(razZC) > 0 Then
            r("message") = ZbirnaMutPoruka(razZC, "zbirne", broj, _
                "Zamena bi prevezala decu OBE zbirne, jer se otpremnice i " & _
                "prijemnice vezuju BROJEM")
            Exit Function
        End If
    End If

    Dim zbrOldID As String
    zbrOldID = CStr(s("zbrID"))

    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, zbrOldID, broj, _
                , , , , , , "Zbirna parkirana za kasnije.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE)."

        Case SV_MODE_DUPLI
            ' DUPLI = razvezi: ATOMARNO (jedna TX) storno zbirne + odvezi otpremnice
            ' ("ceka zbirnu") -> isti obrazac kao RunSimpleStornoZbirna. Otpremnice
            ' (+blokovi) PREZIVLJAVAJU nevezane. Prijemnica/palete se NE storniraju
            ' (to je PONISTENJE) -> ostaju osirocene za reveze (recovery zabeleska).
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, zbrOldID, broj, , , , , , , "Dupli/fantom zbirna.")
            r("correctionID") = cidD
            ' Bez context-a nema recovery reda ni MANUAL flag-a -> ne diraj podatke.
            If Len(cidD) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            Dim expOtpD As Long: expOtpD = CLng(s("otpCount"))
            Dim detD As Long
            If Not StornoZbirnaIDetach_TX(broj, detD, docID) Then
                FailCorrectionContext cidD, "Storno/odvezivanje zbirne (dupli) nije uspelo."
                r("message") = "Storno zbirne nije uspeo.": Exit Function
            End If
            ' Bilo otpremnica a nijedna nije odvezana -> nedosledno: MANUAL (ne lazi COMPLETED).
            If expOtpD > 0 And detD = 0 Then
                MarkCorrectionManual cidD, "Otpremnice NISU odvezane sa stornirane zbirne -> proveri rucno.", _
                    "Zbirna stornirana ali odvezano 0 od " & expOtpD & " otpremnica."
                r("success") = True
                r("message") = "Zbirna stornirana, ali otpremnice (" & expOtpD & ") NISU odvezane. Proveri Osiroceni dokumenti."
                Exit Function
            End If
            If CBool(s("hasPrijemnica")) Or CBool(s("hasPalete")) Then
                MarkCorrectionManual cidD, "Odluci o osirocenoj prijemnici/paletama (reveze ili storno).", _
                    "Fantom zbirna stornirana; odvezano otpremnica: " & detD & "; prijemnica/palete osirocene."
                r("message") = "Zbirna stornirana (fantom); " & detD & " otpremnica vraceno u 'ceka zbirnu'. " & _
                               "Prijemnica/palete osirocene (Osiroceni dokumenti)."
            Else
                CompleteCorrectionContext cidD, , , "Fantom zbirna stornirana; otpremnice odvezane: " & detD & "."
                r("message") = "Zbirna stornirana (fantom); " & detD & " otpremnica vraceno u 'ceka zbirnu'."
            End If
            r("success") = True

        Case SV_MODE_PONISTENJE
            ' PONISTENJE = UVEK prvo pun spisak posledica + svesna potvrda (forceConfirm).
            ' Zatim KASKADA: storno zbirne + svih otpremnica (+oslobodi blokove) +
            ' (hladnjaca kupac) prijemnica + paletne stavke. Eksterni kupac -> prijemnica
            ' ostaje netaknuta (zbirna je poslednji interni dok). Razlika od DUPLI koji
            ' samo odvezuje (deca prezivljavaju).
            If Not forceConfirm Then
                r("blocked") = True
                r("message") = BuildPonistenjePosledice(FLOW_DOC_ZBIRNA, broj, "")
                Exit Function
            End If
            Dim cidP As String
            cidP = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, zbrOldID, broj, , , , , , , "Ponistenje zbirne bez zamene.")
            r("correctionID") = cidP
            ' Bez context-a nema recovery reda ni MANUAL flag-a -> ne diraj podatke.
            If Len(cidP) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            Dim ownsZ As Boolean: ownsZ = ZbirnaOwnsExternalChain(broj)
            Dim cascZ As Object: Set cascZ = PonistiZbirnaChain_TX(broj, ownsZ, docID)
            If Not CBool(cascZ("ok")) Then
                ' RAZLOG iz kaskade ide dalje. Bez ovoga operater vidi samo
                ' "nije uspelo", pa mu specificna kapija ne znaci nista.
                Dim razlogK As String: razlogK = ""
                If cascZ.Exists("message") Then razlogK = Trim$(CStr(cascZ("message")))
                FailCorrectionContext cidP, "Kaskadno ponistenje zbirne nije uspelo."
                r("message") = "Ponistenje zbirne nije uspelo (kaskada)."
                If Len(razlogK) > 0 Then r("message") = razlogK
                Exit Function
            End If
            CompleteCorrectionContext cidP, , , "Ponistena zbirna " & broj & " + ceo interni tok."
            r("success") = True
            r("message") = "Zbirna " & broj & " ponistena sa celim tokom. Otpremnice: " & CStr(cascZ("otp")) & _
                ", prijemnice: " & CStr(cascZ("prij")) & ", paletne stavke: " & CStr(cascZ("pals")) & _
                ", blokovi oslobodjeni: " & CStr(cascZ("blok")) & _
                IIf(ownsZ, "", " (eksterni kupac: prijemnica netaknuta).")

        Case Else
            r("message") = "Nepoznat mod: " & mode
    End Select
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr SRC
    r("message") = "Greska: " & errDescEH
End Function

' ============================================================
' REVERS AMBALAZE - dispatch po modu (saldo vec iskljucuje stornirano ->
' storno = uklanjanje iz salda; bez kontra-stavke, bez duplog salda).
' ============================================================
'
' ambID = AmbID kliknutog reda (ekran Storno). Broj reversa je labela, pa se
' ReversID razresava PRE konteksta i storna: iz reda kad je dat, inace po (broj,
' tip) -- i tada mora biti jednoznacan, inace odbijeno.
Public Function RunReversCorrection(ByVal brDok As String, ByVal dokumentTip As String, _
                                    ByVal mode As String, _
                                    Optional ByVal ambID As String = "") As Object
    Const SRC As String = MOD_NAME & ".RunReversCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunReversCorrection = r
    On Error GoTo EH

    brDok = Trim$(brDok)
    Dim revID As String, revRaz As String
    revRaz = ReversIDRazresi(ambID, brDok, dokumentTip, revID, False)
    If Len(revRaz) = 0 Then revRaz = ReversIDGranica(revID)
    If Len(revRaz) > 0 Then
        r("message") = "Revers nije jednoznacan: " & revRaz
        Exit Function
    End If
    If ReversRedoviRID(revID, False).count = 0 Then
        r("message") = "Aktivan revers nije pronadjen: " & brDok & " [" & dokumentTip & "]"
        Exit Function
    End If

    ' Trag ispravke nosi ReversID (REV-IDENT-01), ne broj: isti broj legalno nosi
    ' revers druge stanice ili drugog dana, pa OldDocID po broju ne bi mogao da
    ' kaze KOJI je revers zamenjen. Broj ide u OldBroj -- labela.
    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_REVERS, revID, brDok, _
                , , , dokumentTip, , , "Revers parkiran za kasnije.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE)."

        Case SV_MODE_ISPRAVKA
            Dim cid As String
            cid = CreateCorrectionContext(mode, FLOW_DOC_REVERS, revID, brDok, FLOW_DOC_REVERS, , , _
                dokumentTip, , , "Ispravka reversa: storno stari, ceka novi.")
            r("correctionID") = cid
            ' Bez context-a nema recovery reda ni MANUAL flag-a -> ne diraj podatke.
            If Len(cid) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip, ambID) Then
                FailCorrectionContext cid, "Storno starog reversa nije uspeo."
                r("message") = "Storno reversa nije uspeo.": Exit Function
            End If
            r("needsForm") = True
            r("success") = True
            r("message") = "Stari revers storniran (uklonjen iz salda). Unesi NOVI revers; " & _
                           "saldo racuna samo novi (bez duple stavke)."

        Case SV_MODE_DUPLI, SV_MODE_PONISTENJE
            Dim cidX As String
            cidX = CreateCorrectionContext(mode, FLOW_DOC_REVERS, revID, brDok, , , , _
                dokumentTip, , , IIf(mode = SV_MODE_DUPLI, "Dupli/fantom revers.", "Ponistenje reversa."))
            r("correctionID") = cidX
            ' Bez context-a nema recovery reda ni MANUAL flag-a -> ne diraj podatke.
            If Len(cidX) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip, ambID) Then
                FailCorrectionContext cidX, "Storno reversa nije uspeo."
                r("message") = "Storno reversa nije uspeo.": Exit Function
            End If
            CompleteCorrectionContext cidX, , , "Revers storniran; saldo koriguje storno (bez kontra-stavke)."
            r("success") = True
            r("message") = "Revers storniran. Saldo azuriran (bez duple/kontra stavke)."

        Case Else
            r("message") = "Nepoznat mod: " & mode
    End Select
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr SRC
    r("message") = "Greska: " & errDescEH
End Function

' Zavrsi ISPRAVKA reversa: veze novi revers u context -- NewDocID je ReversID
' novog reversa, NewBroj njegov broj (labela). Saldo je vec tacan (stari
' storniran, novi aktivan) -> nema dupliranja. Context postaje COMPLETED SAMO ako
' novi revers stvarno postoji kao AKTIVAN (inace MANUAL_REQUIRED).
' dokumentTip se cita iz konteksta (upisan u ParentDocType pri RunReversCorrection).
'
' newStanicaID / newDatum: stanica i dan NOVOG reversa iz snimanja
' (ZavrsiIspravkuAko <- ReversUpisi). Pisac vraca samo uspeh, pa se ReversID
' zamene nalazi po (broj, stanica, dan); zamena sme na drugu stanicu ili drugi
' dan, pa se proverava bas taj revers, ne "ima li aktivan red pod ovim brojem".
' Bez njih (legacy/test poziv) ReversID se trazi po (broj, tip) i mora biti
' jednoznacan, inace MANUAL.
Public Function CompleteReversIspravka(ByVal correctionID As String, ByVal newBrDok As String, _
                                       Optional ByVal newStanicaID As String = "", _
                                       Optional ByVal newDatum As Variant = Empty) As Object
    Dim r As Object: Set r = NewRes(SV_MODE_ISPRAVKA)
    Set CompleteReversIspravka = r
    On Error GoTo EH
    newBrDok = Trim$(newBrDok)
    r("correctionID") = correctionID

    Dim dokTip As String
    dokTip = GetCorrectionField(correctionID, COL_SV_PARENT_DOCTYPE)
    If Len(dokTip) = 0 Then
        MarkCorrectionManual correctionID, "Nedostaje tip reversa u kontekstu -> zavrsi ispravku rucno.", _
            "Context reversa nema DokumentTip (ParentDocType prazan)."
        r("message") = "Ne mogu da odredim tip reversa iz konteksta. Oznaceno za recovery."
        Exit Function
    End If

    ' NewDocID = ReversID NOVOG reversa; NewBroj = labela.
    '
    ' Smer zamene se NE pretpostavlja iz starog reversa: pogresan smer je upravo
    ' jedan od razloga za ispravku, pa bi trazenje pod starim smerom ostavilo
    ' ispravku MANUAL iako je zamena snimljena. Uz stanicu i dan iz snimanja zamenu
    ' nosi JEDAN ReversID preko sva cetiri smera (ReversIDStanice sa praznim tipom).
    ' Bez stanice (legacy/test poziv) vazi smer iz konteksta (ParentDocType).
    ' Oba puta gledaju samo AKTIVNE redove, pa nadjen ReversID jeste aktivan revers.
    Dim revRaz As String, newID As String
    If Len(Trim$(newStanicaID)) > 0 And IsDate(newDatum) Then
        newID = ReversIDStanice(newBrDok, "", Trim$(newStanicaID), Int(CDbl(CDate(newDatum))), False)
        If Len(newID) = 0 Then _
            revRaz = "Novi revers " & newBrDok & " nema jednoznacan aktivan ReversID na toj stanici " & _
                     "tog dana -> identitet zamene nije poznat."
    Else
        revRaz = ReversIDRazresi("", newBrDok, dokTip, newID, False)
    End If
    ' Trag ne sme da pokazuje na ReversID cija granica dokumenta nije cista.
    If Len(revRaz) = 0 Then revRaz = ReversIDGranica(newID)

    If Len(revRaz) > 0 Then
        MarkCorrectionManual correctionID, "Snimi novi revers pa ponovi zavrsetak ispravke.", _
            revRaz
        r("message") = "Novi revers nije pronadjen kao aktivan. Snimi novi revers pa ponovi zavrsetak ispravke."
        Exit Function
    End If

    CompleteCorrectionContext correctionID, newID, newBrDok, "Ispravka reversa: novi revers " & newBrDok & "."
    r("success") = True
    r("message") = "Ispravka reversa zavrsena. Saldo racuna samo novi revers."
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr MOD_NAME & ".CompleteReversIspravka"
    r("message") = "Greska: " & errDescEH
End Function

' ============================================================
' PRIJEMNICA - dispatch po modu. Prijemnica je skoro-list: nizvodni tok = paletne
' stavke (DetachOsirocenePaletaStavke) + faktura (oslobadja se u StornoPrijemnica).
' Otkupni blokovi su SAMOSTALNI (vezani preko BrojZbirne) -> NE diraju se automatski;
' dodatni storno blokova je obrisan u S1e (StornoSelectedBlocks_TX; vraca S3).
' ISPRAVKA: storno + needsForm; prevezivanje paleta radi save-putanja prijemnice
' (ReassignPaleteToPrijemnica_TX) -> ovde se samo pravi context i stornira stara.
' ============================================================
' #3: zavrsetak prijemnica-correctiona po ISHODU detach-a. Ne lazi COMPLETED ako
' palete nisu stvarno skinute:
'   skipPalete=True (Ne diraj palete) -> MANUAL (palete namerno osirocene; recovery JESTE potreban)
'   expected>0 i detached<>expected -> MANUAL (ostatak stavki -> recovery)
'   inace -> COMPLETED.
' expected = broj AKTIVNIH paletnih stavki PRE storna (ScanPrijemnica.paleteCount).
Private Sub CompletePrijemnicaByDetach(ByVal cid As String, ByVal hasPalete As Boolean, _
        ByVal skipPalete As Boolean, ByVal expected As Long, ByVal detached As Long, _
        ByVal what As String, ByVal r As Object)
    If hasPalete And skipPalete Then
        MarkCorrectionManual cid, "Prevezi ili skini palete rucno (Osiroceni dokumenti -> Palete).", _
            what & ": palete OSTAVLJENE osirocene (izbor operatera) -> recovery potreban."
        r("message") = what & ". Palete ostavljene osirocene -> Osiroceni dokumenti (Mod: Palete)."
    ElseIf hasPalete And expected > 0 And detached <> expected Then
        MarkCorrectionManual cid, "Skini preostale paletne stavke (Osiroceni dokumenti -> Palete).", _
            what & ": skinuto " & detached & " od " & expected & " paletnih stavki (ostatak -> recovery)."
        r("message") = what & ", ali skinuto " & detached & "/" & expected & " paletnih stavki -> Osiroceni dokumenti."
    Else
        CompleteCorrectionContext cid, , , what & ": paletne stavke skinute: " & detached & "."
        r("message") = what & ". Paletne stavke skinute: " & detached & "."
    End If
    r("success") = True
End Sub

Public Function RunPrijemnicaCorrection(ByVal broj As String, ByVal mode As String, _
                                        Optional ByVal forceConfirm As Boolean = False, _
                                        Optional ByVal skipPalete As Boolean = False, _
                                        Optional ByVal docID As String = "") As Object
    Const SRC As String = MOD_NAME & ".RunPrijemnicaCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunPrijemnicaCorrection = r
    On Error GoTo EH

    broj = Trim$(broj)
    Dim s As Object: Set s = ScanPrijemnica(broj, docID)
    If Not CBool(s("exists")) Then
        r("message") = "Aktivna prijemnica nije pronadjena: " & broj
        Exit Function
    End If
    Dim parentZbirna As String: parentZbirna = CStr(s("brojZbirne"))
    Dim prijID As String: prijID = CStr(s("prijID"))

    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_PRIJEMNICA, prijID, broj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Prijemnica parkirana za kasnije.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE). Vidljiv u: Osiroceni dokumenti."

        Case SV_MODE_ISPRAVKA
            Dim cid As String
            cid = CreateCorrectionContext(mode, FLOW_DOC_PRIJEMNICA, prijID, broj, _
                FLOW_DOC_PRIJEMNICA, , , FLOW_DOC_ZBIRNA, , parentZbirna, _
                "Ispravka prijemnice: storno stare, ceka snimanje nove (palete se prevezu).")
            If Len(cid) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            If Not StornoPrijemnicaByBroj_TX(broj, docID) Then
                FailCorrectionContext cid, "Storno stare prijemnice nije uspeo."
                r("correctionID") = cid: r("message") = "Storno prijemnice nije uspeo."
                Exit Function
            End If
            r("correctionID") = cid
            r("needsForm") = True
            r("success") = True
            r("message") = "Stara prijemnica stornirana. Popuni i snimi NOVU prijemnicu; " & _
                           "palete se prevezuju automatski po snimanju."

        Case SV_MODE_DUPLI
            ' DUPLI = dupli unos: storno prijemnice + skini paletne stavke (roba nije
            ' primljena 2x). Blokovi ostaju (samostalni; cekirani se storniraju van).
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_PRIJEMNICA, prijID, broj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Dupli/fantom prijemnica.")
            If Len(cidD) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            r("correctionID") = cidD
            If Not StornoPrijemnicaByBroj_TX(broj, docID) Then
                FailCorrectionContext cidD, "Storno prijemnice (dupli) nije uspeo."
                r("message") = "Storno prijemnice nije uspeo.": Exit Function
            End If
            Dim detD As Long, infoD As String
            If CBool(s("hasPalete")) And Not skipPalete Then detD = DetachOsirocenePaletaStavke_TX(broj, infoD)
            CompletePrijemnicaByDetach cidD, CBool(s("hasPalete")), skipPalete, _
                CLng(s("paleteCount")), detD, "Prijemnica stornirana (dupli)", r

        Case SV_MODE_PONISTENJE
            ' PONISTENJE = ceo tok otpada. Prijemnica je 1:1 sa zbirnom -> reuse ISTE
            ' kaskade kao zbirna PONISTENJE (PonistiZbirnaChain_TX): storno zbirne +
            ' otpremnica (+oslobodi blokove) + prijemnice + palete. Zbirna: sve otpremnice
            ' odlaze -> kg 0 -> storno (rekalk je unutar kaskade ako bi neka ostala).
            ' UVEK prvo pun spisak posledica + svesna potvrda (forceConfirm).
            If Not forceConfirm Then
                r("blocked") = True
                r("message") = BuildPonistenjePosledice(FLOW_DOC_PRIJEMNICA, broj, "")
                Exit Function
            End If
            Dim cidP As String
            cidP = CreateCorrectionContext(mode, FLOW_DOC_PRIJEMNICA, prijID, broj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Ponistenje prijemnice bez zamene.")
            r("correctionID") = cidP
            ' Bez context-a nema recovery reda ni MANUAL flag-a -> ne diraj podatke.
            If Len(cidP) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function

            If Len(parentZbirna) > 0 And ZbirnaPostoji(parentZbirna) Then
                Dim ownsP As Boolean: ownsP = ZbirnaOwnsExternalChain(parentZbirna)
                ' ZBR-CHILD-01: v. isti obrazac u otpremnickoj grani -- dete zna
                ' roditelja, pa se ne pogadja po broju.
                ' Dete nosi GENERACIJU roditelja (ZBR-CHILD-01), a kaskada od
                ' S4-2 radi po ZbirnaID-u -- prevod je fail-closed: jedna legacy
                ' generacija legitimno pokriva dva zaglavlja (Klasa I i II), a
                ' storno po ID-u obara tacno jedno.
                ' TRAG NA DETETU JE ZbirnaID (S4-3c), pa PREVODA NEMA.
                '
                ' Do ovog reza je trag citan kao GENERACIJA i prevodjen kroz
                ' IdoviGeneracije. Otkad ZbirnaIDZaBroj vraca identitet, pisci
                ' (SavePrijemnica, ReassignPrijemnicaToZbirna_TX, paletni relink)
                ' u taj trag upisuju ID -- pa bi prevod trazio ID medju
                ' generacijama, nikad ga ne nasao i TIHO vratio prazno. Identitet
                ' bi se izgubio, a kaskada pala nazad na broj: tacno ona klasa
                ' kvara zbog koje ceo ovaj refaktor postoji.
                Dim zbrIdP As String
                zbrIdP = NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, _
                                              prijID, COL_DETE_ZBIRNA_ROD))
                If Len(zbrIdP) = 0 Then zbrIdP = ZbirnaIDZaBroj(parentZbirna)

                Dim cascP As Object: Set cascP = PonistiZbirnaChain_TX(parentZbirna, ownsP, zbrIdP)
                If Not CBool(cascP("ok")) Then
                    ' RAZLOG iz kaskade ide dalje -- isto kao u zbirna grani.
                    Dim razlogP As String: razlogP = ""
                    If cascP.Exists("message") Then razlogP = Trim$(CStr(cascP("message")))
                    FailCorrectionContext cidP, "Kaskadno ponistenje toka (zbirna " & parentZbirna & ") nije uspelo."
                    r("message") = "Ponistenje nije uspelo (kaskada zbirne)."
                    If Len(razlogP) > 0 Then r("message") = razlogP
                    Exit Function
                End If
                ' Eksterni kupac (zbirna ne poseduje prijemnicu u kaskadi) -> prijemnicu
                ' + njene palete storniramo ovde (retko: prijemnica ~ hladnjaca = internal).
                Dim detX As Long, extRemainder As Boolean
                If Not ownsP Then
                    If Len(LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_ID)) > 0 Then
                        StornoPrijemnicaByBroj_TX broj
                        Dim dInfoX As String
                        If CBool(s("hasPalete")) Then detX = DetachOsirocenePaletaStavke_TX(broj, dInfoX)
                        If CBool(s("hasPalete")) And CLng(s("paleteCount")) > 0 And detX <> CLng(s("paleteCount")) Then _
                            extRemainder = True
                    End If
                End If
                If extRemainder Then
                    MarkCorrectionManual cidP, "Skini preostale paletne stavke (Osiroceni dokumenti -> Palete).", _
                        "Ponistenje toka: skinuto " & detX & " od " & CLng(s("paleteCount")) & " paletnih stavki prijemnice (ostatak -> recovery)."
                Else
                    CompleteCorrectionContext cidP, , , "Ponistena prijemnica sa CELIM tokom zbirne " & parentZbirna & "."
                End If
                r("success") = True
                r("message") = "Prijemnica ponistena sa CELIM tokom. Zbirna " & parentZbirna & _
                    " (rekalk/storno), otpremnice: " & CStr(cascP("otp")) & ", prijemnice: " & CStr(cascP("prij")) & _
                    ", paletne stavke: " & CStr(cascP("pals")) & ", blokovi oslobodjeni: " & CStr(cascP("blok")) & "."
            Else
                ' Nema zbirne (prijemnica bez BrojZbirne) -> leaf: storno prijemnice + palete.
                If Not StornoPrijemnicaByBroj_TX(broj, docID) Then
                    FailCorrectionContext cidP, "Storno prijemnice (ponistenje) nije uspeo."
                    r("message") = "Storno prijemnice nije uspeo.": Exit Function
                End If
                Dim detP As Long, infoP As String
                If CBool(s("hasPalete")) Then detP = DetachOsirocenePaletaStavke_TX(broj, infoP)
                CompletePrijemnicaByDetach cidP, CBool(s("hasPalete")), False, _
                    CLng(s("paleteCount")), detP, "Prijemnica ponistena", r
            End If

        Case Else
            r("message") = "Nepoznat mod: " & mode
    End Select
    Exit Function
EH:
    Dim errDescEH As String: errDescEH = Err.description
    LogErr SRC
    r("message") = "Greska: " & errDescEH
End Function

' ============================================================
' PK dokumenta iz KANONSKOG IDENTITETA
' ============================================================
' Zamena za LookupActiveID(tbl, brojCol, broj, idCol), koji uzima PRVI aktivan
' red tog broja. Broj je labela: BrojPrijemnice se racuna PO KUPCU, broj zbirne
' PO KUPCU i bez provere jedinstvenosti -- prvi red tog broja ne mora biti
' dokument koji je operater izabrao. (Kod zbirne generator broj drzi
' jedinstvenim; tamo je ovo pojas za rucni unos.)
'
' Kad je generacija poznata, bira se BAS taj dokument. Kad nije (zatecen zapis),
' pad na broj je dozvoljen tek posto se dokaze da broj nosi JEDNOG vlasnika;
' inace se vraca prazno, pa pozivalac vidi exists=False i staje. To je vaznije
' nego sto izgleda: kod moda RESI_KASNIJE se guarded writer uopste ne zove, pa
' bi se inace napravio TRAJAN recovery zapis nad tudjim dokumentom.
' strict: prazan PK tada znaci iskljucivo "nema takvog dokumenta", ne "nisam
' umeo da ga nadjem".
' --- identitet zbirne u okviru: ZbirnaID ulazi, generacija se IZVODI ---------
'
' Review #371 (P1): okvir je od F8 dobijao ZbirnaID a prosledjivao ga dalje kao
' `gen`, pa ga je PkPoIdentitetu tumacio kao generaciju. Ista vrednost je u dva
' sloja imala dva znacenja -- presecen ugovor, ne rubni slucaj.
'
' Smer je sada jedan:
'
'     ZbirnaID --> mutacija zaglavlja (direktno)
'              --> legacy scoping dece (izvedena generacija)
'
' a NE obrnuto (ZbirnaID tretiran kao generacija, pa trazen nazad ZbirnaID) --
' to bi vratilo sekundarni identitet kao autoritet.

' BROJ SE IZVODI IZ IDENTITETA, NE OBRNUTO (review #371, drugi krug).
'
' Posle prvog kruga je okvir imao ispravne TIPOVE (broj = labela, ID = identitet)
' ali par niko nije proveravao. Bio je moguc poziv (broj = A, zbirnaID = B), a
' posledica nije teorijska: StornoZbirnaIDetach_TX bi stornirao ZAGLAVLJE B i
' odvezao DECU A -- identitet i clanstvo se opet razidju, tacno ono sto refaktor
' uklanja.
'
' Zato ova funkcija vraca KANONSKI broj procitan iz zaglavlja, i nizvodno se
' koristi ON. Prosledjen broj je samo provera zastarelog/pokvarenog izbora:
' ako se ne slaze, staje se -- ne bira se "neki".
'
'     zbirnaID -> zaglavlje -> kanonski BrojZbirne -> scoping dece
'
' Prazan prosledjen broj je dozvoljen (pozivalac ga nema), i tada se samo cita.
Private Function RequireZbirnaPar(ByVal zbirnaID As String, ByVal broj As String, _
                                  ByVal src As String) As String
    Dim zid As String
    zid = Trim$(zbirnaID)

    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA, COL_ZBR_ID, zid)

    Dim n As Long
    If Not redovi Is Nothing Then n = redovi.count

    If Len(zid) = 0 Or n <> 1 Then
        Err.Raise ERR_STORNO_FW_BASE + 67, src, _
                  "ZbirnaID " & zbirnaID & " se nalazi " & CStr(n) & " puta. " & _
                  "Radnja ne sme da dira dokument koji ne zna."
    End If

    Dim kanon As String
    kanon = Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zid, COL_ZBR_BROJ)))

    If Len(Trim$(broj)) > 0 Then
        If StrComp(kanon, Trim$(broj), vbTextCompare) <> 0 Then
            Err.Raise ERR_STORNO_FW_BASE + 68, src, _
                      "Broj i identitet ne pripadaju istom dokumentu: " & _
                      "prosledjen broj " & broj & ", a ZbirnaID " & zbirnaID & _
                      " nosi broj " & kanon & "."
        End If
    End If

    RequireZbirnaPar = kanon
End Function

' Broj -> ZbirnaID, fail-closed.
'
' Postoji zbog POZIVALACA koji jos nose samo broj -- okvir ispravke, koji se
' brise u S4-3 -- a ne zbog zatecenih podataka: produkcije i legacy sveski nema
' (v. "Pravila koja vaze" u docs/STANJE_REFAKTORA.md). Kad okvir nestane,
' nestaje i ova funkcija; dotle je jedini bezbedan prevod onaj koji staje kad
' broj nije jednoznacan.
Private Function ZbrIdPoBroju(ByVal broj As String, ByVal src As String) As String
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If Not IsArray(data) Then
        Err.Raise ERR_STORNO_FW_BASE + 66, src, "Tabela zbirnih nije citljiva."
    End If

    Dim cBroj As Long, cId As Long, cSt As Long
    cBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, src)
    cId = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, src)
    cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)

    Dim i As Long, nadjen As String, n As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cBroj))), Trim$(broj), vbTextCompare) = 0 Then
            Dim jeStorn As Boolean: jeStorn = False
            If cSt > 0 Then jeStorn = (UCase$(Trim$(NzToText(data(i, cSt)))) = "DA")
            If Not jeStorn Then
                n = n + 1
                nadjen = Trim$(NzToText(data(i, cId)))
            End If
        End If
    Next i

    If n <> 1 Then
        Err.Raise ERR_STORNO_FW_BASE + 64, src, _
                  "Broj " & broj & " nosi " & CStr(n) & " aktivnih zbirnih. " & _
                  "Storno po broju nije bezbedan -- identitet je ZbirnaID."
    End If
    ZbrIdPoBroju = nadjen
End Function

Private Function PkPoIdentitetu(ByVal tblName As String, ByVal brojCol As String, _
                                ByVal idCol As String, ByVal broj As String, _
                                ByVal gen As String, ByVal vlasnikCols As Variant, _
                                Optional ByVal strict As Boolean = False) As String
    Const SRC As String = "modStornoFlow.PkPoIdentitetu"
    On Error GoTo EH

    If Len(Trim$(gen)) > 0 Then
        Dim ids As Object: Set ids = IdoviGeneracije(tblName, idCol, gen)
        ' ZADATA generacija koja se ne razresava je greska, ne poziv na fallback.
        ' Do sada je komentar to tvrdio, a kod je svejedno vracao prazno -- pa je
        ' nizvodno izgledalo kao "dokument ne postoji" umesto "ne mogu da ga
        ' razresim". U strict rezimu je to greska, jer se uvid posle oznacava kao
        ' valid. Van strict-a ostaje prazno, zbog zatecenih zapisa bez generacije.
        If ids.count = 0 Then
            If strict Then
                Err.Raise ERR_UI_BASE + 41, SRC, _
                          "Identitet dokumenta se ne moze razresiti u " & tblName & "."
            End If
            Exit Function
        End If
        PkPoIdentitetu = CStr(ids.Keys()(0))
        Exit Function
    End If

    ' Vlasnik moze biti KOMPOZIT -- zbirna je vozac + kupac. Sa jednom kolonom
    ' je ovaj racun bio u kontradikciji sa ScanZbirna, koji ambiguity meri sa
    ' oba.
    Dim vc As Variant
    If IsArray(vlasnikCols) Then vc = vlasnikCols Else vc = Array(vlasnikCols)
    If VlasniciPoBroju(tblName, brojCol, broj, SRC, False, vc).count > 1 Then
        Exit Function
    End If
    PkPoIdentitetu = LookupActiveID(tblName, brojCol, broj, idCol)
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr SRC
    If strict Then Err.Raise errNum, SRC, errDesc
End Function

' gen: kanonski identitet dokumenta koji je operater izabrao u F8. Opcion je
' zbog legacy forme i zatecenih zapisa; bez njega vazi kapija nad brojem.
' strict = citanje koje NE SME da propadne u tisini. Prazan rezultat tada znaci
' iskljucivo "uspesno sam proverio i nema ih"; sve ostalo (schema drift,
' necitljiva tabela, greska u prolazu) DIZE gresku. Trazi ga samo
' modStornoImpact: model uvida se posle oznacava kao valid, a "ne znam" ne sme
' da prodje kao "nema". Podrazumevano False -- zatecenim pozivaocima (legacy
' frmDokumenta, paneli) ponasanje ostaje isto.
Private Function ScanPrijemnica(ByVal broj As String, _
                                Optional ByVal gen As String = "", _
                                Optional ByVal strict As Boolean = False) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanPrijemnica = d
    On Error GoTo EH
    broj = Trim$(broj)
    d("broj") = broj
    Dim prijID As String
    prijID = PkPoIdentitetu(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_ID, broj, gen, COL_PRJ_KUPAC, strict)
    d("prijID") = prijID
    d("exists") = (Len(prijID) > 0)
    If Len(prijID) = 0 Then
        d("brojZbirne") = "": d("fakturisano") = False
        d("hasPalete") = False: d("paleteCount") = 0&: d("blockCount") = 0&: d("otpCount") = 0&
        Exit Function
    End If
    Dim bz As String: bz = NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_BROJ_ZBIRNE))
    d("brojZbirne") = bz
    ' Otpremnice te zbirne (PONISTENJE prijemnice ih stornira; zbirna se rekalk/storno).
    d("otpCount") = OtpCountZbirnePoBroju(bz, strict)
    d("fakturisano") = (UCase$(NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_FAKTURISANO))) = "DA")
    ' Palete se broje po PrijemnicaID kad je dokument razresen: broj bi uracunao
    ' i palete tudjeg dokumenta iste oznake, pa bi pregled lagao operatera.
    Dim palc As Long
    palc = CountActive(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, prijID, strict)
    d("paleteCount") = palc
    d("hasPalete") = (palc > 0)
    ' bz je vec procitan iz TACNOG prijID -- roditelj se ne trazi ponovo po
    ' poslovnom broju prijemnice, koji nije globalno jedinstven.
    d("blockCount") = ActiveOtkupIDsByZbirna(bz).count
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ScanPrijemnica"
    If strict Then Err.Raise errNum, MOD_NAME & ".ScanPrijemnica", errDesc
End Function

' ============================================================
' PANEL DATA - strukturirani podaci za "Storno / potvrda" overlay (frmDokumenta).
' Zamenjuju MsgBox-preview: chain rows (dotaknuti dokumenti) + block rows (multiselect).
' ============================================================

' Aktivni otkup blokovi (samostalni) vezani za flow dokument. Otpremnica: preko
' OtpremnicaID; Zbirna/Prijemnica: preko BrojZbirne. Za multiselect dodatni storno.
' docID (ZbirnaID izabranog dokumenta) NIJE kozmetika: rezultat ove funkcije
' ide u dodatni storno blokova, dakle u MUTACIJU. Bez njega su blokovi svih
' dokumenata istog poslovnog broja u istoj korpi -- a citanje otpremnice po broju
' namerno ukljucuje i STORNIRANE otpremnice, jer njihovi blokovi jos mogu da
' pokazuju na njih.
'
' Kapija BlockStornoDriftReason ovo ne hvata: prva linija joj je
' "If ModeStornoBlokParent(docType, mode) Then Exit Function", a to je True za
' svaki PONISTENJE i za OTPREMNICA+DUPLI/ISPRAVKA -- to jest za tacno one modove
' koji jedini i stizu do dodatnog storna blokova. Pretpostavka "roditelj umire,
' pa je blok-storno bezbedan" vazi samo za blokove IZABRANOG dokumenta.
' strict = citanje koje NE SME da propadne u tisini. Prazan rezultat tada znaci
' iskljucivo "uspesno sam proverio i nema ih"; sve ostalo (schema drift,
' necitljiva tabela, greska u prolazu) DIZE gresku. Trazi ga samo
' modStornoImpact: model uvida se posle oznacava kao valid, a "ne znam" ne sme
' da prodje kao "nema". Podrazumevano False -- zatecenim pozivaocima (legacy
' frmDokumenta, paneli) ponasanje ostaje isto.
Public Function ActiveBlocksForFlow(ByVal docType As String, ByVal broj As String, _
                                    Optional ByVal dokumentTip As String = "", _
                                    Optional ByVal docID As String = "", _
                                    Optional ByVal strict As Boolean = False) As Collection
    Dim result As New Collection
    Set ActiveBlocksForFlow = result
    On Error GoTo EH
    broj = Trim$(broj)
    Select Case docType
        Case FLOW_DOC_ZBIRNA
            ' SEMA: tblOtkup nosi denormalizovan BrojZbirne, ne ZbirnaID -- deca
            ' se po generaciji zbirne ne mogu razdvojiti. Zato ovde nema sta da se
            ' suzi; put je zasticen uzvodno (kapije nad dvosmislenim brojem
            ' zbirne obore mode operaciju, a dodatni storno blokova ide samo posle
            ' uspesne). Ako se te kapije ikad suze, ovo mesto se otvara.
            Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(broj, strict)
        Case FLOW_DOC_PRIJEMNICA
            ' BrojPrijemnice NIJE globalno jedinstven (sekvenca po kupcu), pa je
            ' roditeljska zbirna morala da se cita iz TACNOG dokumenta, ne iz
            ' prvog reda tog broja.
            Dim prijID As String
            prijID = PkPoIdentitetu(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_ID, broj, _
                                    docID, COL_PRJ_KUPAC, strict)
            If Len(prijID) = 0 Then Exit Function
            Dim bz As String
            bz = NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_BROJ_ZBIRNE))
            If Len(bz) > 0 Then Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(bz, strict)
    End Select
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ActiveBlocksForFlow"
    If strict Then Err.Raise errNum, MOD_NAME & ".ActiveBlocksForFlow", errDesc
End Function

' Aktivni OtkupID-jevi za dati BrojZbirne (denormalizovani otkup.BrojZbirne).
' strict: v. GetStornoBlockRows. Ova grana hrani spisak blokova za ZBIRNU i za
' PRIJEMNICU (preko njene zbirne). Dok strict nije stizao dovde, drift nad
' tblOtkup je davao prazan skup, GetStornoBlockRows bi izasao jos na
' "ids.count = 0" -- dakle PRE svoje kapije -- i uvid bi zavrsio kao valid sa
' praznim spiskom blokova.
' Aktivni blokovi (otkupi) te zbirne -- kroz DVA zapisa clanstva (S5-3b).
'
' Zatecen prolaz je citao denormalizovanu kolonu Otkup.BrojZbirne. Nju od S5-3
' vise niko ne pise, pa je broj blokova bio 0 i za zbirnu punu robe -- a taj broj
' operater cita u pregledu PRE nepovratne radnje.
'
' Kanonski put: tblZbirnaIzvori daje otpremnice zbirne, tblOtpremnicaIzvori
' blokove svake otpremnice. Dupli blok ne moze da nastane (jedan blok je na
' najvise jednoj otpremnici), ali se skup svejedno filtrira -- brojka u pregledu
' se ne sme oslanjati na tudju invarijantu.
Private Function ActiveOtkupIDsByZbirna(ByVal brojZbirne As String, _
                                        Optional ByVal strict As Boolean = False) As Collection
    Dim result As New Collection
    Set ActiveOtkupIDsByZbirna = result
    On Error GoTo EH

    Dim zid As String
    zid = ZbrIdPoBrojuMeko(brojZbirne, strict)
    If Len(zid) = 0 Then Exit Function

    Dim vidjeni As Object
    Set vidjeni = CreateObject("Scripting.Dictionary")
    vidjeni.CompareMode = vbTextCompare

    Dim otp As Variant, blok As Variant, bid As String
    For Each otp In KolekcijaUNiz(modDokumenta.ZbrClanoviPoStanju(zid))
        If Not OtpremnicaStornirana(CStr(otp)) Then
            For Each blok In KolekcijaUNiz(modDokumenta.IzvoriOtpremnice(CStr(otp)))
                bid = Trim$(NzToText(blok))
                If Len(bid) > 0 Then
                    If Not vidjeni.Exists(bid) Then
                        vidjeni.Add bid, 1
                        If Not JeStorniranRed(TBL_OTKUP, COL_OTK_ID, bid) Then result.Add bid
                    End If
                End If
            Next blok
        End If
    Next otp
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ActiveOtkupIDsByZbirna"
    If strict Then Err.Raise errNum, MOD_NAME & ".ActiveOtkupIDsByZbirna", errDesc
End Function

' Broj aktivnih otpremnica zbirne zadate BROJEM. Broj je labela, pa se prvo
' prevodi u identitet; neprevodiv broj znaci "nema takve zbirne", dakle 0.
Private Function OtpCountZbirnePoBroju(ByVal broj As String, _
                                       Optional ByVal strict As Boolean = False) As Long
    OtpCountZbirnePoBroju = OtpCountZbirnePoID(ZbrIdPoBrojuMeko(broj, strict))
End Function

' Broj aktivnih otpremnica zbirne zadate IDENTITETOM.
Private Function OtpCountZbirnePoID(ByVal zbirnaID As String) As Long
    If Len(Trim$(zbirnaID)) = 0 Then Exit Function

    Dim otp As Variant, n As Long
    For Each otp In KolekcijaUNiz(modDokumenta.ZbrClanoviPoStanju(zbirnaID))
        If Not OtpremnicaStornirana(CStr(otp)) Then n = n + 1
    Next otp
    OtpCountZbirnePoID = n
End Function

' Broj -> ZbirnaID, bez rusenja pregleda kad broja nema. U strict rezimu greska
' i dalje putuje gore: citanje koje ne sme da propadne u tisini je i dalje takvo.
Private Function ZbrIdPoBrojuMeko(ByVal broj As String, ByVal strict As Boolean) As String
    If Len(Trim$(broj)) = 0 Then Exit Function
    If strict Then
        ZbrIdPoBrojuMeko = ZbrIdPoBroju(broj, MOD_NAME & ".ZbrIdPoBrojuMeko")
        Exit Function
    End If
    On Error Resume Next
    ZbrIdPoBrojuMeko = ZbrIdPoBroju(broj, MOD_NAME & ".ZbrIdPoBrojuMeko")
    On Error GoTo 0
End Function

Private Function OtpremnicaStornirana(ByVal otpremnicaID As String) As Boolean
    OtpremnicaStornirana = JeStorniranRed(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID)
End Function

Private Function JeStorniranRed(ByVal tbl As String, ByVal idCol As String, _
                                ByVal id As String) As Boolean
    If Len(Trim$(id)) = 0 Then Exit Function
    JeStorniranRed = (StrComp(Trim$(NzToText(LookupValue(tbl, idCol, id, COL_STORNIRANO))), _
                              "Da", vbTextCompare) = 0)
End Function

' Dotaknuti dokumenti (pregled u panelu). Collection nizova(0..2): Dokument|Info|Napomena.
' strict = citanje koje NE SME da propadne u tisini. Prazan rezultat tada znaci
' iskljucivo "uspesno sam proverio i nema ih"; sve ostalo (schema drift,
' necitljiva tabela, greska u prolazu) DIZE gresku. Trazi ga samo
' modStornoImpact: model uvida se posle oznacava kao valid, a "ne znam" ne sme
' da prodje kao "nema". Podrazumevano False -- zatecenim pozivaocima (legacy
' frmDokumenta, paneli) ponasanje ostaje isto.
Public Function GetStornoChainRows(ByVal docType As String, ByVal broj As String, _
                                   Optional ByVal dokumentTip As String = "", _
                                   Optional ByVal docID As String = "", _
                                   Optional ByVal strict As Boolean = False) As Collection
    Dim result As New Collection
    Set GetStornoChainRows = result
    On Error GoTo EH
    ' Jedinstven stil: UVEK Dupli unos pa Ponistenje. Isti efekat -> jedan prefiks;
    ' razlicit -> oba, razdvojena crtom. Ispravka i Odlozeno resavanje su
    ' celo-dokumentni (uniformni), pa se objasnjavaju uz samo dugme.
    '
    ' Tekstovi idu kroz katalog, ne kao literali: poslovna recenica trazi
    ' dijakritiku, a VBA izvor mora ostati ASCII. SAM_BLOK zato vise nije Const --
    ' Poruka() nije konstantan izraz.
    Dim SAM_BLOK As String: SAM_BLOK = Poruka("STEF_BLOK_SAM")
    Select Case docType
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj, docID, strict)
            AddChainRow result, "Zbirna", broj, ChainEff(Poruka("STEF_STORNIRA"), Poruka("STEF_STORNIRA"))
            AddChainRow result, "Otpremnice", "(" & CStr(sz("otpCount")) & ")", ChainEff(Poruka("STEF_OTP_ODVEZ"), Poruka("STEF_STORNIRAJU"))
            If CBool(sz("hasPrijemnica")) Then AddChainRow result, "Prijemnica", "(" & CStr(sz("prijCount")) & ")", ChainEff(Poruka("STEF_PRJ_SIROCE"), Poruka("STEF_STORNIRA"))
            If CBool(sz("hasPalete")) Then AddChainRow result, "Paletne stavke", "(" & CStr(sz("paleteCount")) & ")", ChainEff(Poruka("STEF_PAL_SIROCE"), Poruka("STEF_PAL_ODVEZ"))
            AddChainRow result, "Otkupni blokovi", "", SAM_BLOK
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj, docID, strict)
            AddChainRow result, "Prijemnica", broj, ChainEff(Poruka("STEF_STORNO_AMB"), Poruka("STEF_STORNO_AMB"))
            If Len(CStr(sp("brojZbirne"))) > 0 Then _
                AddChainRow result, "Zbirna", CStr(sp("brojZbirne")), ChainEff(Poruka("STEF_NEPROM"), Poruka("STEF_ZBR_NULA"))
            If CLng(sp("otpCount")) > 0 Then _
                AddChainRow result, "Otpremnice", "(" & CStr(sp("otpCount")) & ")", ChainEff(Poruka("STEF_NEPROM_MN"), Poruka("STEF_STORNIRAJU"))
            If CBool(sp("fakturisano")) Then AddChainRow result, "Faktura", "(vezana)", ChainEff(Poruka("STEF_FAK_OSLOB"), Poruka("STEF_FAK_OSLOB"))
            If CBool(sp("hasPalete")) Then AddChainRow result, "Paletne stavke", "(" & CStr(sp("paleteCount")) & ")", ChainEff(Poruka("STEF_PAL_ODVEZ"), Poruka("STEF_PAL_ODVEZ"))
            AddChainRow result, "Otkupni blokovi", "(" & CStr(sp("blockCount")) & ")", SAM_BLOK
        Case FLOW_DOC_REVERS
            AddChainRow result, "Revers", broj & " [" & dokumentTip & "]", Poruka("STEF_REVERS")
    End Select
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".GetStornoChainRows"
    If strict Then Err.Raise errNum, MOD_NAME & ".GetStornoChainRows", errDesc
End Function

Private Sub AddChainRow(ByRef col As Collection, ByVal dok As String, ByVal info As String, ByVal nap As String)
    Dim row(0 To 2) As Variant
    row(0) = dok: row(1) = info: row(2) = nap
    col.Add row
End Sub

' Jedinstven format posledice: Dupli unos UVEK prvo, pa Ponistenje. Isti efekat za
' oba -> jedan spojen prefiks (da se ne ponavlja). Razlicit -> oba, redom.
'
' Prefiksi se zovu ISTO kao dugmad odluke. Do v6-ui-148 je pisalo DUPLIKAT /
' PONISTENJE, a dugmad su glasila "Duplikat" i "Nista se nije desilo" -- operater
' je morao sam da poveze redak u tabeli sa dugmetom na koje se odnosi.
Private Function ChainEff(ByVal dup As String, ByVal pon As String) As String
    If StrComp(Trim$(dup), Trim$(pon), vbTextCompare) = 0 Then
        ChainEff = Poruka("STEF_PRE_OBA") & dup
    Else
        ChainEff = Poruka("STEF_PRE_DUPLI") & dup & "   |   " & _
                   Poruka("STEF_PRE_PONIST") & pon
    End If
End Function

' Bezbedno citanje celije po indeksu kolone (0 = kolona ne postoji -> "").
'
' `data` je ByRef iz istog razloga kao modPaletniList.SafeCell: ByVal bi kopirao
' CEO niz pri svakom pozivu, a ovo je citac PO CELIJI. Tamo je taj obrazac merio
' 1.8 ms po redu; ovde su tabele manje pa se ne vidi, ali je greska ista.
Private Function NzTxC(ByRef data As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then NzTxC = NzTx(data(r, c))
End Function

' ============================================================
' BROWSE za Storno centar (Faza 2b): aktivni dokumenti framework-tipova
' (Prijemnica/Otpremnica/Zbirna) za "Nadji" listu. Distinct po broju (Klasa I/II
' dele broj). Imena razresena preko O(1) dict-ova (BuildLookupDict), otkupna mesta
' iz pre-izgradjene mape (zbirna -> stanice). Namena: pozvati JEDNOM (kes u formi),
' pa filtrirati u memoriji -> nema citanja tabela po tasteru.
' Red = niz(0..7): tip, broj, datum, brojZbirne, kupac, vozac, otkupnaMesta, kolicina.
' ============================================================
Public Function GetActiveDocumentsForStorno(ByVal tipFilter As String, _
                                            ByVal textFilter As String) As Collection
    Const SRC As String = MOD_NAME & ".GetActiveDocumentsForStorno"
    Dim result As New Collection
    Set GetActiveDocumentsForStorno = result
    On Error GoTo EH
    tipFilter = Trim$(tipFilter)
    Dim tf As String: tf = LCase$(Trim$(textFilter))

    ' Name-dict-ovi + otkupna mesta po zbirni (jednom, O(n)).
    Dim kupci As Object: Set kupci = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim vozaci As Object: Set vozaci = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim stByZbr As Object: Set stByZbr = BuildStationsByZbirna()
    ' Otpremnica vise ne nosi BrojZbirne kao kolonu, pa se labela za prikaz
    ' racuna iz clanstva: OtpremnicaID -> broj njene aktivne zbirne.
    Dim zbrPoOtp As Object: Set zbrPoOtp = BuildBrojZbirnePoOtpremnici()

    If WantTip(tipFilter, FLOW_DOC_PRIJEMNICA) Then _
        AddStornoDocs2 result, TBL_PRIJEMNICA, FLOW_DOC_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_DATUM, _
            COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KUPAC, COL_PRJ_VOZAC, COL_PRJ_KOLICINA, tf, kupci, vozaci, stByZbr
    ' Otpremnica: kolicina je na STAVKAMA (S3b), pa umesto kolone zaglavlja ide
    ' zbir po dokumentu. Meko: lista za storno ne sme da nestane zbog jednog
    ' pokvarenog dokumenta -- taj red ostaje, bez brojke.
    Dim zbirOtp As Object
    On Error Resume Next
    Set zbirOtp = modDokumenta.ZbirStavkiPoOtpremnici()
    On Error GoTo EH

    If WantTip(tipFilter, FLOW_DOC_OTPREMNICA) Then _
        AddStornoDocs2 result, TBL_OTPREMNICA, FLOW_DOC_OTPREMNICA, COL_OTP_BROJ, COL_OTP_DATUM, _
            "", "", COL_OTP_VOZAC, "", tf, kupci, vozaci, stByZbr, _
            zbirOtp, COL_OTP_ID, zbrPoOtp
    If WantTip(tipFilter, FLOW_DOC_ZBIRNA) Then _
        AddStornoDocs2 result, TBL_ZBIRNA, FLOW_DOC_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_DATUM, _
            COL_ZBR_BROJ, COL_ZBR_KUPAC, COL_ZBR_VOZAC, COL_ZBR_KOLICINA, tf, kupci, vozaci, stByZbr
    Exit Function
EH:
    LogErr SRC
End Function

Private Function WantTip(ByVal tipFilter As String, ByVal tip As String) As Boolean
    WantTip = (Len(tipFilter) = 0 Or StrComp(tipFilter, "Svi", vbTextCompare) = 0 _
               Or StrComp(tipFilter, tip, vbTextCompare) = 0)
End Function

' zbirnaCol: za Zbirnu = njen broj; za Prijemnicu/Otpremnicu = njihov BrojZbirne.
' kupacCol/vozacCol: "" -> preskace (otpremnica nema kupca). Imena preko dict-ova.
Private Sub AddStornoDocs2(ByRef result As Collection, ByVal tbl As String, ByVal tip As String, _
        ByVal brojCol As String, ByVal datumCol As String, ByVal zbirnaCol As String, _
        ByVal kupacCol As String, ByVal vozacCol As String, ByVal kolCol As String, _
        ByVal tf As String, ByVal kupci As Object, ByVal vozaci As Object, ByVal stByZbr As Object, _
        Optional ByVal zbirStavki As Object, Optional ByVal idCol As String = "", _
        Optional ByVal zbrPoId As Object = Nothing)
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Sub
    Dim cBr As Long, cDa As Long, cZb As Long, cKu As Long, cVo As Long, cKo As Long, cSt As Long
    Dim cId As Long
    cBr = GetColumnIndex(tbl, brojCol)
    cDa = GetColumnIndex(tbl, datumCol)
    cZb = GetColumnIndex(tbl, zbirnaCol)
    If Len(kupacCol) > 0 Then cKu = GetColumnIndex(tbl, kupacCol)
    If Len(vozacCol) > 0 Then cVo = GetColumnIndex(tbl, vozacCol)
    If Len(kolCol) > 0 Then cKo = GetColumnIndex(tbl, kolCol)
    If Len(idCol) > 0 Then cId = GetColumnIndex(tbl, idCol)
    cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    If cBr = 0 Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            Dim broj As String: broj = Trim$(CStr(data(i, cBr)))
            If Len(broj) > 0 Then
                If Not seen.Exists(broj) Then
                    seen(broj) = True
                    ' Labela zbirne: sa reda kad je tabela jos nosi (prijemnica,
                    ' sama zbirna), inace iz clanstva po PK-u (otpremnica od S5-3b).
                    Dim zbr As String
                    If cZb > 0 Then
                        zbr = NzTxC(data, i, cZb)
                    ElseIf Not zbrPoId Is Nothing And cId > 0 Then
                        zbr = DictGet2(zbrPoId, UCase$(Trim$(NzTxC(data, i, cId))), "")
                    End If
                    Dim kup As String: kup = ""
                    If cKu > 0 Then kup = DictGet2(kupci, NzTxC(data, i, cKu), NzTxC(data, i, cKu))
                    Dim voz As String: voz = ""
                    If cVo > 0 Then voz = DictGet2(vozaci, NzTxC(data, i, cVo), "")
                    Dim mesta As String: mesta = DictGet2(stByZbr, zbr, "")
                    Dim datum As String: datum = FmtDatum(NzTxC(data, i, cDa))
                    ' Kolicina: sa zaglavlja kad je tamo, sa STAVKI kad nije
                    ' (otpremnica od S3b). Dokument bez stavki ostaje u listi sa
                    ' praznom kolonom -- storno se radi po identitetu, a nula bi
                    ' rekla da nema sta da se stornira.
                    Dim kol As String: kol = ""
                    If cKo > 0 Then
                        kol = NzTxC(data, i, cKo)
                    ElseIf Not zbirStavki Is Nothing And cId > 0 Then
                        Dim oid As String: oid = Trim$(NzTxC(data, i, cId))
                        If zbirStavki.Exists(oid) Then
                            Dim rec As Variant: rec = zbirStavki(oid)
                            kol = modStornoDok.KgTekst(CDbl(rec(0)))
                        End If
                    End If
                    If Len(tf) = 0 Or _
                       InStr(LCase$(broj & " " & zbr & " " & kup & " " & mesta & " " & datum), tf) > 0 Then
                        Dim row(0 To 7) As Variant
                        row(0) = tip: row(1) = broj: row(2) = datum: row(3) = zbr
                        row(4) = kup: row(5) = voz: row(6) = mesta: row(7) = kol
                        result.Add row
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Mapa: brojZbirne -> ";"-spojena distinct otkupna mesta (stanice) te zbirne, iz
' aktivnih otpremnica. Jednoprolazno; stanice imena preko dict-a.
Private Function BuildStationsByZbirna() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildStationsByZbirna = d
    On Error GoTo EH

    ' OTKUPNA MESTA IDU ZA CLANSTVOM (S5-3b).
    '
    ' Zatecen prolaz je grupisao tblOtpremnica po koloni BrojZbirne. Nju vise
    ' niko ne pise, pa je mapa bila prazna i kolona "otkupna mesta" u listi za
    ' storno je stajala prazna za svaku zbirnu.
    '
    ' Kljuc ostaje BROJ, jer ga takvog trazi prikaz; menja se samo odakle veza
    ' dolazi -- iz tblZbirnaIzvori umesto sa deteta.
    Dim stanice As Object: Set stanice = BuildLookupDict(TBL_STANICE, "StanicaID", "Naziv")
    Dim clanstvo As Object: Set clanstvo = modDokumenta.AktivnoZbrClanstvoPoKanonu()
    If clanstvo Is Nothing Then Exit Function
    If clanstvo.count = 0 Then Exit Function

    Dim brojPoZbr As Object: Set brojPoZbr = BuildLookupDict(TBL_ZBIRNA, COL_ZBR_ID, COL_ZBR_BROJ)

    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cId As Long, cSt As Long, cStorno As Long
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
    If cId = 0 Or cSt = 0 Then Exit Function

    Dim seenPair As Object: Set seenPair = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> "DA" Then
            Dim otpId As String: otpId = UCase$(Trim$(CStr(data(i, cId))))
            If clanstvo.Exists(otpId) Then
                Dim zbr As String
                zbr = DictGet2(brojPoZbr, CStr(clanstvo(otpId)), "")
                If Len(zbr) > 0 Then
                    Dim stId As String: stId = Trim$(CStr(data(i, cSt)))
                    Dim stNm As String: stNm = DictGet2(stanice, stId, stId)
                    If Len(stNm) > 0 Then
                        Dim pk As String: pk = zbr & "|" & stNm
                        If Not seenPair.Exists(pk) Then
                            seenPair(pk) = True
                            If d.Exists(zbr) Then d(zbr) = CStr(d(zbr)) & ";" & stNm Else d(zbr) = stNm
                        End If
                    End If
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildStationsByZbirna"
End Function

' Mapa: UCase(OtpremnicaID) -> BROJ njene aktivne zbirne. Prikaz trazi labelu,
' a labela od S5-3b zivi samo na zaglavlju zbirne.
Private Function BuildBrojZbirnePoOtpremnici() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set BuildBrojZbirnePoOtpremnici = d
    On Error GoTo EH

    Dim clanstvo As Object: Set clanstvo = modDokumenta.AktivnoZbrClanstvoPoKanonu()
    If clanstvo Is Nothing Then Exit Function

    Dim brojPoZbr As Object: Set brojPoZbr = BuildLookupDict(TBL_ZBIRNA, COL_ZBR_ID, COL_ZBR_BROJ)

    Dim k As Variant, broj As String
    For Each k In clanstvo.Keys
        broj = DictGet2(brojPoZbr, CStr(clanstvo(k)), "")
        If Len(broj) > 0 Then d(CStr(k)) = broj
    Next k
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildBrojZbirnePoOtpremnici"
End Function

' Dict lookup sa fallback-om (kljuc prazan -> ""; nema u dict -> fb).
Private Function DictGet2(ByVal d As Object, ByVal key As String, ByVal fb As String) As String
    If Len(key) = 0 Then Exit Function
    If Not d Is Nothing Then
        If d.Exists(key) Then DictGet2 = CStr(d(key)) Else DictGet2 = fb
    Else
        DictGet2 = fb
    End If
End Function

Private Function FmtDatum(ByVal v As String) As String
    On Error Resume Next
    If IsDate(v) Then FmtDatum = Format$(CDate(v), "dd.mm.yyyy") Else FmtDatum = v
End Function

' ============================================================
' GUARD C (ADR-0001): blok-storno nad ZIVOM otpremnicom pravi tihi disbalans
' (otpremnica/zbirna precenjene). Dozvoljeno je samo kad ova operacija i sama
' stornira roditeljsku otpremnicu bloka (PONISTENJE kaskada; ili otpremnica-nivo
' DUPLI/ISPRAVKA). Inace: odbij + preusmeri na otpremnica ISPRAVKA. Unbound blok
' (bez otpremnice) je uvek bezbedan (ne precenjuje nista).
' Vraca "" ako je bezbedno; inace razlog odbijanja (za MsgBox).
' ============================================================
Public Function BlockStornoDriftReason(ByVal docType As String, ByVal mode As String, _
                                       ByVal blkIds As Collection) As String
    On Error GoTo EH
    If blkIds Is Nothing Then Exit Function
    If blkIds.count = 0 Then Exit Function
    If ModeStornoBlokParent(docType, mode) Then Exit Function     ' roditelj umire -> ok
    Dim liveOtp As String: liveOtp = FirstLiveOtpremnicaForBlocks(blkIds)
    If Len(liveOtp) > 0 Then
        BlockStornoDriftReason = _
            "Cekiran otkupni blok je vezan za AKTIVNU otpremnicu " & liveOtp & "." & vbCrLf & _
            "Storno bloka bi ostavio otpremnicu i zbirnu precenjene (ADR-0001: izdati " & _
            "dokument se ne menja u mestu)." & vbCrLf & vbCrLf & _
            "Skini cekiranje bloka, ILI koristi otpremnica ISPRAVKA (storno cele otpremnice + reizdaj)."
    End If
    Exit Function
EH:
    LogErr MOD_NAME & ".BlockStornoDriftReason"
End Function

' True = ova (docType, mode) i sama stornira roditeljsku otpremnicu bloka, pa je
' dodatni blok-storno bezbedan (nema zive otpremnice da precenjuje).
Private Function ModeStornoBlokParent(ByVal docType As String, ByVal mode As String) As Boolean
    ' Otpremnica vise nema modove (S3c): ispravka je jedan potez u F1, a DUPLI
    ' i PONISTENJE se vracaju sa S4/S6, kad nizvodni lanac opet postoji.
    If mode = SV_MODE_PONISTENJE Then ModeStornoBlokParent = True
End Function

' Prvi (citljiv) broj AKTIVNE otpremnice na koju je vezan neki od datih blokova;
' "" ako nijedan blok nije u sastavu aktivne otpremnice.
'
' KANON, NE STARA VEZA (S3c): pripadnost zivi u tblOtpremnicaIzvori i cita se
' kroz modDokumenta.OtpremnicaZaOtkup. Do ovog koraka je citana kolona
' Otkup.OtpremnicaID, koju od S3a ne pise nijedan zivi put -- kapija je zato
' UVEK vracala "" i odbijanje nikad nije stizalo do operatera.
'
' FAIL-CLOSED: strog citac clanstva pada na korupciji, i tada se ne sme
' odgovoriti "bezbedno je". Razlog se vraca, pa panel odbije.
Private Function FirstLiveOtpremnicaForBlocks(ByVal blkIds As Collection) As String
    Dim k As Long, otpID As String, br As String
    On Error GoTo EH
    For k = 1 To blkIds.count
        otpID = modDokumenta.OtpremnicaZaOtkup(Trim$(CStr(blkIds(k))))
        If Len(otpID) > 0 Then
            br = LookupActiveID(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)
            If Len(br) = 0 Then br = otpID
            FirstLiveOtpremnicaForBlocks = br
            Exit Function
        End If
    Next k
    Exit Function
EH:
    LogErr MOD_NAME & ".FirstLiveOtpremnicaForBlocks"
    FirstLiveOtpremnicaForBlocks = "(clanstvo se ne moze procitati)"
End Function

' ============================================================
' Sledljivost (ADR-0002 / Faza 7 korak 2): utisni na dokument-redove da je NOVI red
' ispravka STAROG -> IspravkaOd + CorrectionID na AKTIVNOM novom redu; ZamenjenSa na
' STORNIRANOM starom redu. Best-effort, guarded na postojanje kolona (schema-drift
' safe). NIJE agregat -> ne menja ponasanje; samo vidljiv audit trag NA dokumentu.
' newBroj == oldBroj (in-place, bez zamene) -> nema sta da se utisne.
' ============================================================
Public Sub StampIspravkaTrace(ByVal tbl As String, ByVal brojCol As String, _
        ByVal newBroj As String, ByVal oldBroj As String, ByVal correctionID As String)
    On Error GoTo EH
    newBroj = Trim$(newBroj): oldBroj = Trim$(oldBroj)
    If Len(newBroj) = 0 Then Exit Sub
    If StrComp(newBroj, oldBroj, vbTextCompare) = 0 Then Exit Sub
    Dim cBr As Long: cBr = GetColumnIndex(tbl, brojCol)
    If cBr = 0 Then Exit Sub
    Dim cIsp As Long: cIsp = GetColumnIndex(tbl, COL_TRACE_ISPRAVKA_OD)
    Dim cZam As Long: cZam = GetColumnIndex(tbl, COL_TRACE_ZAMENJEN_SA)
    Dim cCid As Long: cCid = GetColumnIndex(tbl, COL_TRACE_CORRECTION_ID)
    Dim cSt As Long: cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    If cIsp = 0 And cZam = 0 And cCid = 0 Then Exit Sub          ' schema jos nije zdrava
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Sub
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim b As String: b = Trim$(CStr(data(i, cBr)))
        Dim isStorno As Boolean
        isStorno = (cSt > 0 And UCase$(Trim$(CStr(data(i, cSt)))) = "DA")
        If b = newBroj And Not isStorno Then
            If cIsp > 0 And Len(oldBroj) > 0 Then UpdateCell tbl, i, COL_TRACE_ISPRAVKA_OD, oldBroj
            If cCid > 0 And Len(correctionID) > 0 Then UpdateCell tbl, i, COL_TRACE_CORRECTION_ID, correctionID
        ElseIf b = oldBroj And isStorno Then
            If cZam > 0 Then UpdateCell tbl, i, COL_TRACE_ZAMENJEN_SA, newBroj
        End If
    Next i
    Exit Sub
EH:
    LogErr MOD_NAME & ".StampIspravkaTrace"
End Sub
' Atomarno (JEDNA TX): storno zbirne (core) + odvezivanje otpremnica ("ceka
' zbirnu") + otkup denorm. Jedan izvor istine za "storno+detach zbirne" -> koriste
' ga i RunSimpleStornoZbirna i DUPLI grana (ne dve odvojene transakcije). Vraca
' True na uspeh; outDet = broj odvezanih otpremnica.
Private Function StornoZbirnaIDetach_TX(ByVal broj As String, ByRef outDet As Long, _
                                        Optional ByVal zbirnaID As String = "") As Boolean
    Const SRC As String = MOD_NAME & ".StornoZbirnaIDetach_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    outDet = 0
    broj = Trim$(broj)
    If Len(broj) = 0 Then Exit Function
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP
    ' Zaglavlje po generaciji. DetachOtpremniceInline nize ide po BROJU jer
    ' otpremnica zbirnu i nosi kao broj -- zato kapija: dva aktivna dokumenta
    ' istog broja delila bi otpremnice, pa bi se odvezale i tudje.
    '
    ' Do v6-ui-225 je ova kapija brojala VLASNIKE, a komentar iznad nje govorio o
    ' DOKUMENTIMA. Dva aktivna dokumenta ISTOG vlasnika (A17, sto MasterSync od
    ' v6-ui-224 ispravno pravi) prolazila su: zaglavlje bi bilo stornirano tacno,
    ' po generaciji, a onda bi Detach nize odvezao decu OBA dokumenta.
    ' Storniran vlasnik i dalje moze imati AKTIVNU decu -- v. ScanZbirna.
    ' Faza 4: odluka o rezimu se racuna PRE kapije i deli sa akterom. Detach je
    ' do sada odlucivao sam, ispod kapije -- pa je kapija branila i ono sto akter
    ' vise ne moze da pogresi. Isti izraz sada vide oboje.
    ' Identitet ulazi kao ZbirnaID. Do S4-3b se iz njega izvodila GENERACIJA, za
    ' "legacy scoping dece" -- ali generaciju vise ne pise nijedan ziv pisac, pa
    ' je taj scoping bio mrtva grana koja je uvek davala prazno.
    Dim zbrID As String
    zbrID = Trim$(zbirnaID)
    If Len(zbrID) = 0 Then zbrID = ZbrIdPoBroju(broj, SRC)

    ' Deca se odvezuju po BROJU, pa broj mora doci iz ISTOG dokumenta ciji se
    ' header stornira. Inace: header B storniran, deca A odvezana.
    broj = RequireZbirnaPar(zbrID, broj, SRC)

    ' SCOPING DECE IDE PO ZbirnaID-u (S4-3c).
    '
    ' Sposobnost je ista kao pre: kad SVA aktivna deca nose trag roditelja, smem
    ' uze -- diram samo svoju decu, pa me kapija dvosmislenog broja ne mora
    ' zaustaviti. Promenilo se samo STA je scope: do ovog reza generacija
    ' zaglavlja, sada njegov IDENTITET.
    '
    ' Nista drugo nije trebalo dirati: i SvaAktivnaDecaNoseZbirnaID i
    ' SuziDecuNaZbirnu rade nad TRAGOM NA DETETU, a taj trag od S4-3c nosi
    ' ZbirnaID. Prvo sam ceo scoping obrisao kao "mrtav kod" -- nije bio mrtav
    ' nego pogresno hranjen.
    Dim scopeID As String
    Dim razMut As String
    razMut = ZbirnaScopeRazlog(broj, zbrID, False, scopeID)
    If Len(razMut) > 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 62, SRC, _
                  ZbirnaMutPoruka(razMut, "zbirne", broj, _
                                  "Otpremnice se vezuju BROJEM, pa se ne mogu odvezati samo za jedan")
    End If

    ' KANONSKO CLANSTVO SE BROJI PRE STORNA (S4-3b).
    '
    ' Posle storna ga AktivnoClanstvoPoKanonu vise ne vidi -- i to je bas ono sto
    ' ga oslobadja -- pa bi brojanje posle uvek dalo nulu. Poruka operateru je do
    ' ovog reza govorila "0 otpremnica vraceno" i za zbirnu koja ih je imala:
    ' brojala je samo staru vezu po BrojZbirne, koju kanonska otpremnica ne nosi.
    Dim clanova As Long
    ' ZbrClanoviPoStanju, ne ZbrClanovi (review #389, P2).
    '
    ' Prvi pokusaj je bio IzvoriZbirne (strog uvek) i srusio je transakciju nad
    ' NACRTOM, koji legitimno nema nijedan red clanstva -- palo je 9 storno
    ' provera. Popravka je tada bila "onda uvek permisivan", i time je SIMPLE/
    ' DUPLI ostao bez kapije koju PONISTENJE ima: izdata zbirna sa izgubljenim
    ' clanstvom je prolazila, a dupli red se brojao kao druga otpremnica i tako
    ' prijavljivao operateru.
    '
    ' Tacan odgovor nije ni "uvek strog" ni "uvek permisivan" nego PO STANJU
    ' DOKUMENTA, i sada ga daje jedno telo za sve ulaze.
    '
    ' Poziv je PRE StornoZbirna, pa greska staje bez ijedne mutacije.
    clanova = modDokumenta.ZbrClanoviPoStanju(zbrID).count

    If Not StornoZbirna(zbrID) Then _
        Err.Raise ERR_STORNO_FW_BASE + 60, SRC, "StornoZbirna nije uspeo."

    ' ODVEZIVANJA VISE NEMA -- STORNO ZAGLAVLJA JESTE ODVEZIVANJE (S5-3b).
    '
    ' AktivnoClanstvoPoKanonu izbacuje stornirane zbirne, pa otpremnica prestaje
    ' da bude clan istog trena kad zaglavlje padne i odmah se vraca u
    ' NevezaneOtpremnice. Nista se ne brise ni ne prazni.
    '
    ' DetachOtpremniceInline je brisao labelu na detetu. Otkad tu labelu niko ne
    ' pise, brisao je prazno polje i vracao 0, pa se njegov rezultat SABIRAO sa
    ' kanonskim brojem clanova -- zbir u kom jedan sabirak nije mogao biti razlicit
    ' od nule. Sad je ostao samo broj koji nesto meri.
    outDet = clanova
    tx.CommitTx
    Set tx = Nothing
    StornoZbirnaIDetach_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    StornoZbirnaIDetach_TX = False
End Function

' ============================================================
' PRIVATE - DUPLI (razvezi) / PONISTENJE (kaskada) primitive
'
' Ownership pravilo: dokument sme da kaskadira/razveze SAMO ono sto ekskluzivno
' poseduje. Zbirna poseduje otpremnice + prijemnicu + palete (preko BrojZbirne);
' otpremnica poseduje otkup blokove + zbirnu SAMO ako je jedina otpremnica te
' zbirne. U normalnom modu nizvodni tok (prijemnica/palete) pripada zbirni SAMO
' za hladnjaca-kupca; za eksternog kupca je zbirna poslednji interni dokument.
' ============================================================

' Da li nizvodni tok (prijemnica/palete) PRIPADA zbirni -> sme kaskada. Interni
' hladnjaca-tok (kupac == CFG_MALINA_DEFAULT_KUPAC / malina): DA. Eksterni kupac:
' NE (prijemnica je eksterna, ide svojim faktura-mehanizmom). Detekcija = kao u
' frmDokumenta.RefreshBrojPrijSuggestion (modAutoHladnjaca.IsHladnjacaKupac).
Private Function ZbirnaOwnsExternalChain(ByVal brojZbirne As String) As Boolean
    On Error Resume Next
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function
    Dim kup As String
    kup = NzTx(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KUPAC))
    ZbirnaOwnsExternalChain = IsHladnjacaKupac(kup)
End Function
' TEST SEAM: DistinctActiveValues je Private, a ZBR-NORM-02 trazi da se i ona
' meri -- inace bi test dokazao samo dva od tri odlucivaca, a treci bi mogao da
' ostane na starom poredjenju bez ijedne crvene tvrdnje.
Public Function DistinctActiveValues_Test(ByVal tblName As String, ByVal valueCol As String, _
                                          ByVal filterCol As String, _
                                          ByVal filterVal As String) As Long
    If Not IsTestMode() Then Exit Function
    DistinctActiveValues_Test = DistinctActiveValues(tblName, valueCol, filterCol, filterVal).count
End Function

' TEST SEAM: ZbirnaBrojJeDvosmislenIkad je Private, a njeno ponasanje NA
' SOPSTVENU GRESKU je poslovna odluka -- fail-open kapija je gora od nikakve.
' Kroz ponasanje se to ne moze izmeriti jednoznacno: pod schema drift-om pada i
' sve ostalo, pa bi operacija stala iz drugog razloga i test bio placebo.
' Tvrdo gejtovano -- van test-rezima ne radi nista.
Public Function ZbirnaDvosmislenaIkad_Test(ByVal broj As String) As Boolean
    If Not IsTestMode() Then Exit Function
    ZbirnaDvosmislenaIkad_Test = ZbirnaBrojJeDvosmislenIkad(broj)
End Function
' Rekalkulisi zbirnu iz preostalih aktivnih otpremnica; ako ih VISE NEMA -> STORNO
' zbirne (nikad aktivna 0/0 -> to je bio "nuliranje" bug). NE dira prijemnicu/palete
' (mod odlucuje: DUPLI ostavlja osiroceno; PONISTENJE kaskadira zasebno). True=uspeh.
' Sme li se po BROJU mutirati ono sto visi o zbirni? Boolean oblik, za dva
' pozivaoca kojima ne treba poruka; ostali zovu ZbirnaMutRazlog i dobiju UZROK.
'
' Otpremnica flow mutira RODITELJSKU zbirnu -- rekalkulise je, stornira, ili joj
' prevezuje prijemnice -- a sve to ide PO BrojZbirne. Dok deca nemaju generaciju,
' nerazresen broj roditelja mora da zaustavi operaciju.
'
' IME JE ZATECENO i uze od znacenja: od v6-ui-225 ovo nije samo "vise vlasnika"
' nego i "vise aktivnih dokumenata istog vlasnika". Nije preimenovano da se u
' istom koraku ne bi menjala i mera i imena na osam mesta.
'
' Fail-closed na sopstvenu gresku je sada u DVA sloja, oba u jezgru:
' ZbirnaIdentResolve na gresku vraca INTEGRITY_ERROR (ne prazan DTO), a
' ZbirnaMutacijaPoBrojuRazlog to pretvara u blokadu. Za kapiju je "ne mogu da
' dokazem jednoznacnost" isto sto i "ne mutiraj".
Private Function ZbirnaBrojJeDvosmislenIkad(ByVal broj As String) As Boolean
    ZbirnaBrojJeDvosmislenIkad = (Len(ZbirnaMutRazlog(broj)) > 0)
End Function

' ZBR-MUT-01: JEDNA definicija "sme li se mutirati po broju", u jezgru identiteta.
'
' Do v6-ui-225 je ova kapija brojala VLASNIKE (VlasniciPoBroju .count > 1). Svaki
' komentar uz njenih sest poziva je opisivao DOKUMENTNU dvosmislenost ("dva
' aktivna dokumenta istog broja delila bi otpremnice"), a mera je bila vlasnicka
' -- dva pojma koja se poklapaju samo dok jedan vlasnik ne moze da ima dva
' dokumenta pod istim brojem. MasterSync od v6-ui-224 bas to ispravno pravi
' (KR-001, dva uredjaja offline), pa se pojmovi razilaze i mera vise ne vazi.
'
' Sta se NIJE promenilo: vlasnicka grana ostaje, i dalje IKAD (storniran vlasnik
' ima aktivnu decu). Dodata je samo dokumentna.
' JEDAN RACUN SCOPE-A ZA KAPIJU I AKTERA (review #384, P2).
'
' Kapija dispecera je racunala NESCOPED (ZbirnaMutRazlog(broj)), a akter scoped
' -- pa je RunZbirnaCorrection odbijao radnju koju primitiv ume bezbedno da
' uradi: dva aktivna dokumenta istog broja, svako sa svojom decom, i tacan
' ZbirnaID u ruci. Sposobnost je postojala i nije se mogla dosegnuti iz F8.
'
' To je ista klasa greske koju ovaj refaktor vise puta sece: kapija i akter
' odgovaraju na ISTO pitanje, a odgovor im nije isto telo.
'
' diraPrijemnice opisuje KOJU DECU ce operacija mutirati -- scope vazi samo ako
' BAS TA deca nose identitet roditelja. Zato je parametar, a ne fiksan skup:
' PONISTENJE kaskadira na prijemnice i palete kad lanac ide do njih, DUPLI ne.
'
' OTPREMNICE I BLOKOVI SU ISPALI IZ RACUNA (S5-3b), i to ne kao popustanje.
' Pitanje "nose li sva aktivna deca identitet roditelja" imalo je smisla dok su
' se birala po BROJU, pa je trag na detetu bio jedino sto ih je razdvajalo. Od
' S5-3b se oba sprata biraju iz clanstva po ZbirnaID-u, gde tudje dete ne moze ni
' da udje u skup -- uslov je postao tautologija nad praznom kolonom, a ne kapija.
'
' Prijemnice i palete se JOS UVEK biraju po broju (njihov most pada u S6), pa
' njihov uslov ostaje netaknut.
'
' outScopeID: "" = operacija ide po broju; inace identitet po kom se deca
' suzavaju. Vraca RAZLOG odbijanja, "" = sme.
Private Function ZbirnaScopeRazlog(ByVal broj As String, ByVal zbirnaID As String, _
                                   ByVal diraPrijemnice As Boolean, _
                                   ByRef outScopeID As String) As String
    outScopeID = ""

    If Len(Trim$(zbirnaID)) > 0 Then
        Dim ok As Boolean
        ok = True
        If diraPrijemnice Then
            ok = modDokumenta.SvaAktivnaDecaNoseZbirnaID(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, broj) _
                 And modDokumenta.SvaAktivnaDecaNoseZbirnaID(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, broj)
        End If
        If ok Then outScopeID = Trim$(zbirnaID)
    End If

    ZbirnaScopeRazlog = ZbirnaMutRazlog(broj, Len(outScopeID) > 0)
End Function

Private Function ZbirnaMutRazlog(ByVal broj As String, _
                                 Optional ByVal scopedPoGeneraciji As Boolean = False) As String
    ZbirnaMutRazlog = modDokumenta.ZbirnaMutacijaPoBrojuRazlogZaBroj(broj, scopedPoGeneraciji)
End Function

' Poruka za operatera. Uzrok se NE stapa u jednu recenicu: "dva vlasnika" i "dva
' unosa istog vlasnika" traze razlicit potez, a pokvaren identitet treci.
Private Function ZbirnaMutPoruka(ByVal razlog As String, ByVal uloga As String, _
                                 ByVal broj As String, ByVal posledica As String) As String
    Dim uzrok As String, savet As String
    Select Case razlog
        Case ZBR_MUT_VISE_VLASNIKA
            uzrok = "je pripadao VISE vlasnika"
            savet = "Razdvoj brojeve pa ponovi."
        Case ZBR_MUT_VISE_DOKUMENATA
            uzrok = "nosi VISE aktivnih dokumenata (isti vlasnik, dva odvojena unosa)"
            savet = "Storniraj visak ili razdvoj brojeve pa ponovi."
        Case Else
            uzrok = "ima aktivnu zbirnu bez identiteta (ZbirnaID)"
            savet = "Pokreni Provere integriteta (B9) pa ponovi."
    End Select
    ZbirnaMutPoruka = "Broj " & uloga & " '" & broj & "' " & uzrok & "."
    If Len(posledica) > 0 Then ZbirnaMutPoruka = ZbirnaMutPoruka & " " & posledica & "."
    ZbirnaMutPoruka = ZbirnaMutPoruka & " " & savet
End Function

' Aktivne otpremnice zbirne -- CLANSTVO JE JEDINI IZVOR (S5-3b).
'
' Most sa rokom iz review-a #384 je istekao. Do S5-3 je uz kanonsko clanstvo
' stajala i stara veza Otpremnica.BrojZbirne, pa je skup bio unija dva izvora:
' kanon plus sve sto je pauziran PWA uvoz ostavio za sobom. Sam komentar mosta
' je nosio rok -- "umire u S5, kad uvoz predje na kanon".
'
' S5-3 je taj uvoz preveo na CreateZbirnaIzIzvora_TX, pa je poslednji pisac te
' kolone nestao, a S5-3b je kolonu uklonio iz kanona. Unija je time postala
' kanon plus prazan skup, a citalac koji sabira nesto sa praznim skupom laze o
' tome odakle mu podatak.
'
' Parametar gen je otisao sa njom: generacijsko suzavanje je postojalo da razdvoji
' dva dokumenta pod ISTIM brojem. Clanstvo je po ZbirnaID-u, pa dvosmislenosti
' nema -- suziti skup koji je vec tacan moze samo da izbaci tacan red.
Private Function ActiveOtpIDsByZbirna(ByVal SRC As String, _
                                      ByVal zbirnaID As String) As Collection
    Dim result As New Collection
    Set ActiveOtpIDsByZbirna = result

    Dim vidjeni As Object
    Set vidjeni = CreateObject("Scripting.Dictionary")
    vidjeni.CompareMode = vbTextCompare
    '
    ' Izbor citaoca po stanju dokumenta zivi u modDokumenta.ZbrClanoviPoStanju --
    ' ovde je do review-a #389 stajala If-grana, pa je isto pravilo vazilo samo za
    ' PONISTENJE dok su SIMPLE, DUPLI i strog uvid zvali permisivan citac.
    '
    ' Poziv je PRE BeginTx, pa greska staje bez ijedne mutacije.
    If Len(Trim$(zbirnaID)) > 0 Then
        Dim clan As Variant, clanId As String
        For Each clan In KolekcijaUNiz(modDokumenta.ZbrClanoviPoStanju(zbirnaID))
            clanId = Trim$(NzToText(clan))
            If Len(clanId) > 0 Then
                If StrComp(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                             clanId, COL_STORNIRANO))), "Da", vbTextCompare) <> 0 Then
                    If Not vidjeni.Exists(clanId) Then
                        vidjeni.Add clanId, 1
                        result.Add clanId
                    End If
                End If
            End If
        Next clan
    End If
End Function

' Collection -> Variant niz, da For Each ne zavisi od tipa kolekcije.
Private Function KolekcijaUNiz(ByVal c As Collection) As Variant
    If c Is Nothing Then KolekcijaUNiz = Array(): Exit Function
    If c.count = 0 Then KolekcijaUNiz = Array(): Exit Function

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    KolekcijaUNiz = a
End Function

' Aktivni PrijemnicaID-jevi za dati BrojZbirne (svi redovi, obe klase).
Private Function ActivePrijIDsByZbirna(ByVal brojZbirne As String, ByVal gen As String, _
                                       ByVal SRC As String) As Collection
    Dim result As New Collection
    Set ActivePrijIDsByZbirna = result
    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cId As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    cSt = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, SRC)
    Dim kand As Collection: Set kand = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            kand.Add i
        End If
    Next i
    Set kand = SuziDecuNaZbirnu(TBL_PRIJEMNICA, data, kand, gen)
    For i = 1 To kand.count
        result.Add Trim$(CStr(data(CLng(kand(i)), cId)))
    Next i
End Function

' Oslobodi (razvezi) otkup blokove datih otpremnica ID-jeva: OtpremnicaID="" i
' Veza Otkup.OtpremnicaID="" na AKTIVNIM otkup redovima -> vracaju se u pool
' (za reveze). Labela BrojZbirne je otisla sa kolonom u S5-3b. Bez TX
' (unutar otvorene transakcije). Otkup se NIKAD ne stornira (realne kupovine).
Private Function FreeOtkupBloksInline(ByVal otpIDs As Collection, ByVal SRC As String) As Long
    If otpIDs Is Nothing Then Exit Function
    If otpIDs.count = 0 Then Exit Function
    Dim idSet As Object: Set idSet = CreateObject("Scripting.Dictionary")
    Dim x As Long
    For x = 1 To otpIDs.count
        idSet(CStr(otpIDs(x))) = True
    Next x
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cOtp As Long, cSt As Long
    cOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If idSet.Exists(Trim$(CStr(data(i, cOtp)))) Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                RequireUpdateCell TBL_OTKUP, i, COL_OTK_OTPREMNICA_ID, "", SRC
                SetOtkupBrojOtpremnice i, ""      ' Faza 7 korak 5: ocisti denorm kljuc (unbind)
                n = n + 1
            End If
        End If
    Next i
    FreeOtkupBloksInline = n
End Function

' TX wrapper za oslobadjanje blokova (DUPLI / deljena zbirna).
Private Function FreeOtkupBloksByOtpIDs_TX(ByVal otpIDs As Collection) As Long
    Const SRC As String = MOD_NAME & ".FreeOtkupBloksByOtpIDs_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    If otpIDs Is Nothing Then Exit Function
    If otpIDs.count = 0 Then Exit Function
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    Dim n As Long: n = FreeOtkupBloksInline(otpIDs, SRC)
    tx.CommitTx
    Set tx = Nothing
    FreeOtkupBloksByOtpIDs_TX = n
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
End Function

' PONISTENJE kaskada. ownsChain = da li prijemnica/palete pripadaju zbirni
' (hladnjaca kupac / malina); eksterni kupac -> prijemnica/palete se NE diraju.
' Faza A (jedna TX): zbirna -> otpremnice (+oslobodi blokove) -> prijemnice
' (faktura osirocena kroz StornoPrijemnica). Faza B: paletne stavke idu kroz
' PALETNI MOTOR (DetachOsirocenePaletaStavke_TX po prijemnici) -> skida gajbe/neto/
' amb sa palete (reopen ispod kapaciteta), PRAZNA paleta se stornira, su-stanari
' (druge prijemnice/zbirne na istoj paleti) NETAKNUTI. Motor se samo poziva (isti
' put kao recovery panel "Skini stavke"), ne dira se. Vraca: ok/otp/prij/pals/blok.
' gen bira ZAGLAVLJE zbirne. Decu bira BROJ -- drugog kljuca u semi nema -- pa
' kad broj nose dve aktivne zbirne kaskada staje: ponistavanje bi odvezalo i
' tudje otpremnice i prijemnice.
Private Function PonistiZbirnaChain_TX(ByVal brojZbirne As String, ByVal ownsChain As Boolean, _
                                       Optional ByVal zbirnaID As String = "") As Object
    Const SRC As String = MOD_NAME & ".PonistiZbirnaChain_TX"
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    res("ok") = False: res("otp") = 0&: res("prij") = 0&: res("pals") = 0&: res("blok") = 0&
    Set PonistiZbirnaChain_TX = res
    Dim tx As clsTransaction
    On Error GoTo EH

    ' FAIL-CLOSED: deca se biraju po BrojZbirne, pa dva aktivna dokumenta istog
    ' broja dele decu iz ugla ove rutine. Ponistavanje bi odvezalo i tudje.
    ' Storniran vlasnik i dalje moze imati AKTIVNU decu -- v. ScanZbirna.
    ' Do v6-ui-225 je i ovde mera bila vlasnicka, a opasnost dokumentna.
    ' Faza 4: rezim se racuna PRE kapije i deli sa akterom (v. isti obrazac u
    ' StornoZbirnaIDetach_TX).
    ' Kao u StornoZbirnaIDetach_TX: ZbirnaID je identitet, generacija je izvedena
    ' i sluzi samo za scoping dece.
    ' ID se razresava SAMO kad je zadat. Kad nije, razresenje ceka da prodje
    ' kapija dvosmislenosti ispod -- inace bi fail-closed prevod progutao
    ' informativnu poruku ("broj je pripadao vise vlasnika") i operater bi dobio
    ' genericki neuspeh.
    Dim zbrID As String
    zbrID = Trim$(zbirnaID)

    If Len(zbrID) > 0 Then
        brojZbirne = RequireZbirnaPar(zbrID, brojZbirne, SRC)
    End If

    ' SCOPING DECE IDE PO ZbirnaID-u (S4-3c) -- v. isti obrazac u
    ' StornoZbirnaIDetach_TX. Prijemnice i palete ulaze u odluku samo kad lanac
    ' stvarno ide do njih (ownsChain).
    Dim scopeID As String
    Dim razPon As String
    razPon = ZbirnaScopeRazlog(brojZbirne, zbrID, ownsChain, scopeID)
    If Len(razPon) > 0 Then
        res("message") = ZbirnaMutPoruka(razPon, "zbirne", brojZbirne, _
                                         "Deca se u semi vezuju BROJEM, pa se lanac ne moze ponistiti samo za jedan")
        Exit Function
    End If
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    ' ID SE RAZRESAVA PRE IZBORA DECE (S5-3b).
    '
    ' Dok se biralo po broju, razresenje je smelo da ceka kapiju iznad -- izbor
    ' ga nije trazio. Kanon bira po ZbirnaID-u, pa bi prazan ID ovde dao prazan
    ' skup otpremnica i tiho "ponisteno, 0 otpremnica" nad zbirnom koja ih ima.
    ' To je ista klasa laznog uspeha koju je review #384 zatvorio na drugom kraju.
    '
    ' Razresenje je i dalje POSLE kapije, pa informativna poruka o dvosmislenom
    ' broju i dalje stize do operatera umesto generickog neuspeha.
    If Len(zbrID) = 0 Then
        If ZbirnaPostoji(brojZbirne) Then zbrID = ZbrIdPoBroju(brojZbirne, SRC)
    End If

    ' Rezim je izracunat IZNAD kapije: kaskada bira po broju iz tri skupa, pa bi
    ' nezavisna odluka po tabeli mogla da stornira otpremnice samo GEN-B a
    ' prijemnice svih generacija. Prijemnice i palete ulaze u odluku samo kad
    ' ownsChain -- kad ih kaskada ne dira, njihov legacy red nema zasto da obori
    ' suzavanje otpremnica.

    ' ID-jeve + prijemnica-brojeve-sa-paletama skupi PRE mutacije.
    Dim otpIDs As Collection
    Set otpIDs = ActiveOtpIDsByZbirna(SRC, zbrID)
    Dim prijIDs As Collection, prijBrPalete As Collection
    If ownsChain Then
        Set prijIDs = ActivePrijIDsByZbirna(brojZbirne, scopeID, SRC)
        Set prijBrPalete = DistinctActiveValues(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, _
                                                COL_PALS_BROJ_ZBIRNE, brojZbirne, scopeID)
    Else
        Set prijIDs = New Collection: Set prijBrPalete = New Collection
    End If

    ' --- Faza A: dokument kaskada (zbirna + otpremnice + blokovi + prijemnice) ---
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_OTKUP
    If ownsChain Then
        tx.AddTableSnapshot TBL_PRIJEMNICA
        tx.AddTableSnapshot TBL_FAKTURE
        tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    End If

    ' gen je do v6-ui-225 bio MRTAV PARAMETAR: pozivalac ga je slao (linija sa
    ' docID), a StornoZbirna se zvao bez njega -- pa je zaglavlje biralo po BROJU,
    ' dakle sve redove tog broja. Kapija iznad sada ne pusta dva aktivna dokumenta,
    ' ali izbor svejedno mora da bude po identitetu: ispravljena zbirna pod istim
    ' brojem ima i storniranu generaciju, i nju ne treba ponovo dirati.
    If Len(zbrID) > 0 Then
        If Not StornoZbirna(zbrID) Then _
            Err.Raise ERR_STORNO_FW_BASE + 50, SRC, "StornoZbirna (ponistenje) nije uspeo."
    End If
    Dim k As Long
    ' BLOKOVI SE BROJE PRE STORNA (review #384).
    '
    ' FreeOtkupBloksInline broji samo staru vezu Otkup.OtpremnicaID, koju
    ' kanonski pisac ne pise -- pa je poruka javljala "blokovi oslobodjeni: 0"
    ' i za otpremnicu koja ih je imala. Sami blokovi JESU oslobodjeni:
    ' StornoOtpremnica to radi kroz clanstvo. Lagao je samo broj, a broj koji
    ' operater cita posle nepovratne radnje ne sme da laze.
    Dim blokKanon As Long
    For k = 1 To otpIDs.count
        blokKanon = blokKanon + modDokumenta.IzvoriOtpremnice(CStr(otpIDs(k))).count
    Next k

    For k = 1 To otpIDs.count
        If Not StornoOtpremnica(CStr(otpIDs(k))) Then _
            Err.Raise ERR_STORNO_FW_BASE + 51, SRC, "StornoOtpremnica (ponistenje) nije uspeo: " & CStr(otpIDs(k))
    Next k
    res("otp") = otpIDs.count
    res("blok") = FreeOtkupBloksInline(otpIDs, SRC) + blokKanon
    If ownsChain Then
        For k = 1 To prijIDs.count
            If Not StornoPrijemnica(CStr(prijIDs(k))) Then _
                Err.Raise ERR_STORNO_FW_BASE + 52, SRC, "StornoPrijemnica (ponistenje) nije uspeo: " & CStr(prijIDs(k))
        Next k
        res("prij") = prijIDs.count
    End If
    tx.CommitTx
    Set tx = Nothing

    ' --- Faza B: paletne stavke kroz paletni motor (header/reopen/storno-prazne) ---
    If ownsChain Then
        Dim info As String, b As Long
        For b = 1 To prijBrPalete.count
            res("pals") = CLng(res("pals")) + DetachOsirocenePaletaStavke_TX(CStr(prijBrPalete(b)), info)
        Next b
    End If

    res("ok") = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
End Function

' Distinktni OtpremnicaID-jevi za dati BrojOtpremnice (ukljucuje i stornirane,
' jer blokovi mogu jos pokazivati na storniran ID).
' Red pripada IZABRANOM dokumentu: po generaciji kad je poznata, inace po
' broju. Isto pravilo kao RedJeIzabranogDokumenta u modStorno -- ovde zaseban
' jer modStornoFlow radi nad svojim ucitanim nizovima.
Private Function RedJeGeneracije(ByRef data As Variant, ByVal i As Long, _
                                 ByVal cBroj As Long, ByVal cGen As Long, _
                                 ByVal broj As String, ByVal gen As String) As Boolean
    If Len(Trim$(gen)) = 0 Then
        RedJeGeneracije = (Trim$(CStr(data(i, cBroj))) = broj)
        Exit Function
    End If
    ' Zadata generacija a kolone nema: tih pad na broj bi znacio da se dira
    ' nesto drugo. Isto pravilo kao RedJeIzabranogDokumenta u modStorno.
    If cGen = 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 63, MOD_NAME & ".RedJeGeneracije", _
                  "Zadata je generacija dokumenta, a tabela nema kolonu " & _
                  COL_GENERACIJA_ID & ". Pokreni EnsureRuntimeSchema pa ponovi."
    End If
    RedJeGeneracije = (Trim$(NzToText(data(i, cGen))) = Trim$(gen))
End Function

' Distinktni AKTIVNI OtkupID-jevi vezani (OtpremnicaID) za dati skup otp ID-jeva.
' strict: v. GetStornoBlockRows. Prazan spisak sme da znaci samo "proverio sam i
' nema blokova", nikad "ne umem da proverim" -- inace uvid tvrdi da nema
' pogodjenih blokova nad odlukom koja blokove stornira.
Private Function GetBlokOtkupIDs(ByVal otpIDs As Collection, _
                                 Optional ByVal strict As Boolean = False) As Collection
    Dim result As New Collection
    Set GetBlokOtkupIDs = result
    On Error GoTo EH
    If otpIDs Is Nothing Then Exit Function
    If otpIDs.count = 0 Then Exit Function

    Dim idSet As Object: Set idSet = CreateObject("Scripting.Dictionary")
    Dim x As Long
    For x = 1 To otpIDs.count
        idSet(CStr(otpIDs(x))) = True
    Next x

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        If strict Then
            If Not modUiData.TabelaCitljiva(TBL_OTKUP) Then
                Err.Raise ERR_UI_BASE + 34, MOD_NAME & ".GetBlokOtkupIDs", _
                          "Tabela " & TBL_OTKUP & " nije nadjena."
            End If
        End If
        Exit Function
    End If
    Dim cOtp As Long, cId As Long, cSt As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    If cOtp = 0 Or cId = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 35, MOD_NAME & ".GetBlokOtkupIDs", _
                      "Kolona " & COL_OTK_OTPREMNICA_ID & " ili " & COL_OTK_ID & _
                      " ne postoji u " & TBL_OTKUP & "."
        End If
        Exit Function
    End If

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, oid As String
    For i = 1 To UBound(data, 1)
        If idSet.Exists(Trim$(CStr(data(i, cOtp)))) Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                oid = Trim$(CStr(data(i, cId)))
                If Len(oid) > 0 And Not seen.Exists(oid) Then
                    seen(oid) = True
                    result.Add oid
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".GetBlokOtkupIDs"
    If strict Then Err.Raise errNum, MOD_NAME & ".GetBlokOtkupIDs", errDesc
End Function

' ============================================================
' PRIVATE - chain scan + generic helpers
' ============================================================

Private Function ScanZbirna(ByVal broj As String, _
                            Optional ByVal zbirnaID As String = "", _
                            Optional ByVal strict As Boolean = False) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanZbirna = d
    On Error GoTo EH
    broj = Trim$(broj)
    d("broj") = broj
    ' PK izabrane zbirne -- correction context polazi od njega.
    '
    ' Od S4-2 ga ljuska SALJE (nevidljiva kolona reda je ZbirnaID), pa se ovde
    ' vise ne razresava iz broja i generacije. Zatecene putanje koje nose samo
    ' broj ga izvode fail-closed: dva aktivna dokumenta istog broja su greska,
    ' ne "uzmi prvi".
    If Len(Trim$(zbirnaID)) > 0 Then
        ' Par se proverava, a broj se dalje cita IZ ZAGLAVLJA: nizvodno brojanje
        ' dece ne sme da veruje labeli koju je pozivalac poslao.
        broj = RequireZbirnaPar(Trim$(zbirnaID), broj, MOD_NAME & ".ScanZbirna")
        d("broj") = broj
        d("zbrID") = Trim$(zbirnaID)
    ElseIf strict Then
        d("zbrID") = ZbrIdPoBroju(broj, MOD_NAME & ".ScanZbirna")
    Else
        d("zbrID") = ""
        On Error Resume Next
        d("zbrID") = ZbrIdPoBroju(broj, MOD_NAME & ".ScanZbirna")
        On Error GoTo EH
    End If
    ' Deca (otpremnice, prijemnice, palete) vezuju zbirnu KOLONOM BrojZbirne --
    ' ZbirnaID im nije strani kljuc nigde u semi. Zato se broje po broju, a kad
    ' broj nose DVE aktivne zbirne, brojke opisuju oba dokumenta. To se ne moze
    ' razdvojiti podatkom koji postoji, pa se ne pravimo da moze -- putanje koje
    ' bi na osnovu toga menjale decu staju (v. PonistiZbirnaChain_TX).
    ' ZBR-MUT-01: JEDAN racun, isti koji koriste sve ostale putanje po broju.
    '
    ' Do v6-ui-225 su ovde stajala DVA vlasnicka brojaca. Oba su merila vlasnike,
    ' ne dokumente, pa dva aktivna dokumenta ISTOG vlasnika (A17) nisu videla --
    ' a bas po ovom kljucu je RunZbirnaCorrection odlucivao da li sme da dira
    ' decu. Jedan od ta dva (brojDvosmislen, samo aktivni) nije imao nijednog
    ' citaoca, pa je i sam bio poziv na pogresan racun.
    '
    ' Racun UKLJUCUJE STORNIRANE vlasnike, namerno. StornoZbirna_TX stornira SAMO
    ' redove tblZbirna -- otpremnice, prijemnice i palete ne dira. Zato je ovo
    ' potpuno legitimno stanje:
    '
    '   Zbirna A  broj Z-10  STORNIRANA   ali OTP-A i PRJ-A jos AKTIVNI
    '   Zbirna B  broj Z-10  AKTIVNA
    '
    ' Sa brojanjem samo AKTIVNIH, izbor B daje "broj je jednoznacan" -- pa
    ' DetachOtpremniceInline i kaskada, koje idu PO BROJU, odvezu i decu
    ' stornirane A. Storniran vlasnik nestaje iz racuna, njegova deca ne.
    d("mutRazlog") = ZbirnaMutRazlog(broj)
    d("otpCount") = OtpCountZbirnePoID(CStr(d("zbrID")))
    Dim pc As Long: pc = CountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, broj, strict)
    d("prijCount") = pc
    d("hasPrijemnica") = (pc > 0)
    Dim palc As Long: palc = CountActive(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, broj, strict)
    d("paleteCount") = palc
    d("hasPalete") = (palc > 0)
    Exit Function
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ScanZbirna"
    If strict Then Err.Raise errNum, MOD_NAME & ".ScanZbirna", errDesc
End Function

' Uvid u JEDAN revers -- ReversID iz ambID-a kliknutog reda, ili jednoznacno po
' (broj, tip). Noge ovog dokumenta bira isti kod kao pisac (ReversIDRazresi,
' ReversRedoviRID). razlog <> "" = identitet nije razresen (dvosmislen broj, red
' bez ReversID-a, ambalaza uz otkup, noge Stanica bez jedne stanice i dana): uvid
' tada ne prikazuje NIJEDAN dokument.
' FAIL-CLOSED: greska se DIZE (BuildStornoPreview je prikaze), ne pretvara se u
' "nije pronadjen".
Private Function ScanRevers(ByVal brDok As String, ByVal dokumentTip As String, _
                            Optional ByVal ambID As String = "") As Object
    Const SRC As String = MOD_NAME & ".ScanRevers"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanRevers = d
    Dim errNum As Long, errDesc As String
    On Error GoTo EH
    brDok = Trim$(brDok)
    d("broj") = brDok
    d("exists") = False: d("razlog") = "": d("stanica") = "": d("dan") = 0&
    d("tip") = "": d("kolicina") = 0&: d("smer") = "": d("entitet") = "": d("redova") = 0&

    Dim revID As String, st As String, dan As Long, razlog As String
    razlog = ReversIDRazresi(ambID, brDok, dokumentTip, revID, False)
    If Len(razlog) = 0 Then razlog = ReversIDGranica(revID)
    If Len(razlog) > 0 Then
        d("razlog") = razlog
        Exit Function
    End If
    Dim redovi As Collection, v As Variant, i As Long
    Set redovi = ReversRedoviRID(revID, False)
    d("exists") = (redovi.count > 0)
    If Not CBool(d("exists")) Then Exit Function
    razlog = ReversStanicaDan(revID, st, dan)
    If Len(razlog) > 0 Then
        d("exists") = False
        d("razlog") = razlog
        Exit Function
    End If
    d("stanica") = st
    d("dan") = dan

    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function
    Dim cTip As Long, cKol As Long, cSmer As Long, cEnt As Long
    cTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    cKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
    cSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
    cEnt = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)

    For Each v In redovi
        i = CLng(v)
        d("tip") = NzTx(data(i, cTip))
        d("smer") = NzTx(data(i, cSmer))
        d("entitet") = NzTx(data(i, cEnt))
        d("redova") = CLng(d("redova")) + 1
        ' Revers = dvojni upis (Kooperant + Stanica, isti broj/tip) -> NE sabiraj
        ' obe noge; kolicina dokumenta = jedna noga (reprezentativna/veca).
        If IsNumeric(data(i, cKol)) Then
            If CLng(data(i, cKol)) > CLng(d("kolicina")) Then d("kolicina") = CLng(data(i, cKol))
        End If
    Next v
    Exit Function
EH:
    errNum = Err.Number: errDesc = Err.description
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' Broj AKTIVNIH redova gde filterCol = value.
' strict: nula tada znaci iskljucivo "prebrojao sam i nema ih". Bez toga je
' Scan* bio strict spolja a slep iznutra: nestane tblPrijemnica.BrojZbirne ->
' CountActive vrati 0 -> ekran kaze hasPrijemnica = False, i uvid je i dalje
' valid.
Private Function CountActive(ByVal tblName As String, ByVal filterCol As String, _
                             ByVal value As String, _
                             Optional ByVal strict As Boolean = False) As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(tblName)
    If IsEmpty(data) Then
        If strict Then
            If Not modUiData.TabelaCitljiva(tblName) Then
                Err.Raise ERR_UI_BASE + 38, MOD_NAME & ".CountActive", _
                          "Tabela " & tblName & " nije nadjena."
            End If
        End If
        Exit Function
    End If
    Dim cF As Long, cSt As Long
    cF = GetColumnIndex(tblName, filterCol)
    cSt = GetColumnIndex(tblName, COL_STORNIRANO)
    If cF = 0 Then
        If strict Then
            Err.Raise ERR_UI_BASE + 39, MOD_NAME & ".CountActive", _
                      "Kolona " & filterCol & " ne postoji u " & tblName & "."
        End If
        Exit Function
    End If
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cF))) = value Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then n = n + 1
        End If
    Next i
    CountActive = n
    Exit Function
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".CountActive"
    If strict Then Err.Raise errNum, MOD_NAME & ".CountActive", errDesc
End Function

' Distinktne AKTIVNE vrednosti valueCol gde filterCol = filterVal.
Private Function DistinctActiveValues(ByVal tblName As String, ByVal valueCol As String, _
                                      ByVal filterCol As String, ByVal filterVal As String, _
                                      Optional ByVal gen As String = "") As Collection
    Dim result As New Collection
    Set DistinctActiveValues = result
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(tblName)
    If IsEmpty(data) Then Exit Function
    Dim cV As Long, cF As Long, cSt As Long
    cV = GetColumnIndex(tblName, valueCol)
    cF = GetColumnIndex(tblName, filterCol)
    cSt = GetColumnIndex(tblName, COL_STORNIRANO)
    If cV = 0 Or cF = 0 Then Exit Function

    ' ZBR-CHILD-01 faza 3: kandidati pa suzavanje, da bi dedup radio nad decom
    ' JEDNOG dokumenta. Dedup pre suzavanja bi spojio vrednosti dva dokumenta pod
    ' istim brojem i suzavanje vise ne bi imalo sta da razdvoji.
    Dim kand As Collection: Set kand = New Collection
    Dim c As Long
    For c = 1 To UBound(data, 1)
        ' ZBR-NORM-02: filterVal je do sada poredjen NETRIMOVAN, dok je celija
        ' bila trimovana -- asimetrija koja bi netrimovanom pozivaocu tiho
        ' vratila prazan skup. Sva tri zatecena pozivaoca salju trimovanu
        ' vrednost, pa nije bilo ziv kvar, ali zamka jeste.
        If BrojJednak(data(c, cF), filterVal) Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(c, cSt)))) <> "DA" Then kand.Add c
        End If
    Next c
    Set kand = SuziDecuNaZbirnu(tblName, data, kand, gen)

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, v As String
    For i = 1 To kand.count
        v = Trim$(CStr(data(CLng(kand(i)), cV)))
        If Len(v) > 0 And Not seen.Exists(v) Then
            seen(v) = True
            result.Add v
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".DistinctActiveValues"
End Function

Private Function NewRes(ByVal mode As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("success") = False
    d("blocked") = False
    d("needsForm") = False
    d("correctionID") = ""
    d("mode") = mode
    d("message") = ""
    Set NewRes = d
End Function

Private Function NzTx(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzTx = ""
    Else
        NzTx = Trim$(CStr(v))
    End If
End Function

Private Function YesNo(ByVal b As Boolean) As String
    YesNo = IIf(b, "Da", "Ne")
End Function

Private Function Fmt(ByVal v As Variant) As String
    On Error Resume Next
    Fmt = Format$(CDbl(v), "0.##")
End Function
