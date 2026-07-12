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
'   Relink: modDokumenta.ReassignOtkupToOtpremnica_TX, ReassignPrijemnicaToZbirna_TX.
'   Invariant: modDokumentInvariant.RecalculateZbirnaFromOtpremnice_TX,
'              ValidateZbirnaInvariant, ValidateOtpremnicaZbirnaImpact.
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
                                   Optional ByVal dokumentTip As String = "") As String
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_OTPREMNICA:  BuildStornoPreview = PreviewOtpremnica(broj)
        Case FLOW_DOC_ZBIRNA:      BuildStornoPreview = PreviewZbirna(broj)
        Case FLOW_DOC_REVERS:      BuildStornoPreview = PreviewRevers(broj, dokumentTip)
        Case FLOW_DOC_PRIJEMNICA:  BuildStornoPreview = PreviewPrijemnica(broj)
        Case Else:                 BuildStornoPreview = "Dokument: " & docType & " " & broj
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildStornoPreview"
    BuildStornoPreview = "Pregled nije dostupan (greska). Dokument: " & docType & " " & broj
End Function

Private Function PreviewOtpremnica(ByVal broj As String) As String
    Dim s As Object: Set s = ScanOtpremnica(broj)
    Dim m As String
    m = "OTPREMNICA " & broj & vbCrLf
    If Not CBool(s("exists")) Then
        PreviewOtpremnica = m & "(nije pronadjena aktivna otpremnica)"
        Exit Function
    End If
    m = m & "Stanica: " & CStr(s("stanica")) & vbCrLf
    m = m & "Otkupni blokovi: " & CStr(s("blockCount")) & vbCrLf
    m = m & "Broj zbirne: " & IIf(Len(CStr(s("brojZbirne"))) > 0, CStr(s("brojZbirne")), "(nema)") & vbCrLf
    m = m & "Prijemnica preko zbirne: " & YesNo(CBool(s("hasPrijemnica"))) & _
            " (" & CStr(s("prijCount")) & ")" & vbCrLf
    m = m & "Palete preko prijemnice: " & YesNo(CBool(s("hasPalete"))) & _
            " (" & CStr(s("paleteCount")) & ")" & vbCrLf
    m = m & "Rizik ambalaza: storno vraca ambalazu ove otpremnice (auto)."
    PreviewOtpremnica = m
End Function

Private Function PreviewZbirna(ByVal broj As String) As String
    Dim s As Object: Set s = ScanZbirna(broj)
    Dim inv As Object: Set inv = s("invariant")
    Dim m As String
    m = "ZBIRNA " & broj & vbCrLf
    m = m & "Aktivne otpremnice: " & CStr(s("otpCount")) & vbCrLf
    m = m & "Suma otpremnica  KG I: " & Fmt(inv("kgOtpI")) & " | KG II: " & Fmt(inv("kgOtpII")) & _
            " | AMB: " & CStr(inv("ambOtpTotal")) & vbCrLf
    m = m & "Redovi zbirne    KG I: " & Fmt(inv("kgZbrI")) & " | KG II: " & Fmt(inv("kgZbrII")) & _
            " | AMB: " & CStr(inv("ambZbrTotal")) & vbCrLf
    m = m & "Invarijanta: " & IIf(CBool(inv("isValid")), "OK", "MISMATCH") & vbCrLf
    m = m & "Prijemnica: " & YesNo(CBool(s("hasPrijemnica"))) & " (" & CStr(s("prijCount")) & ")" & vbCrLf
    m = m & "Paletizovano: " & YesNo(CBool(s("hasPalete"))) & " (" & CStr(s("paleteCount")) & ")"
    PreviewZbirna = m
End Function

Private Function PreviewRevers(ByVal broj As String, ByVal dokumentTip As String) As String
    Dim s As Object: Set s = ScanRevers(broj, dokumentTip)
    Dim m As String
    m = "REVERS " & broj & " [" & dokumentTip & "]" & vbCrLf
    If Not CBool(s("exists")) Then
        PreviewRevers = m & "(nije pronadjen aktivan revers)"
        Exit Function
    End If
    m = m & "Kooperant/Stanica: " & CStr(s("entitet")) & vbCrLf
    m = m & "Tip ambalaze: " & CStr(s("tip")) & vbCrLf
    m = m & "Kolicina: " & CStr(s("kolicina")) & " (knjiznih redova: " & CStr(s("redova")) & ")" & vbCrLf
    m = m & "Smer: " & CStr(s("smer")) & vbCrLf
    m = m & "Uticaj na saldo: storno iskljucuje ovaj revers iz salda (bez duple stavke)."
    PreviewRevers = m
End Function

Private Function PreviewPrijemnica(ByVal broj As String) As String
    Dim s As Object: Set s = ScanPrijemnica(broj)
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
Public Function GetChainFlags(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal dokumentTip As String = "") As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")
    Set GetChainFlags = r
    On Error GoTo EH
    r("hasDependents") = False
    r("dependentsText") = ""
    r("canPonistenjeClean") = True

    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            Dim so As Object: Set so = ScanOtpremnica(broj)
            Dim dep As Boolean
            dep = CBool(so("hasZbirna")) Or CBool(so("hasPrijemnica")) Or CBool(so("hasPalete"))
            r("hasDependents") = dep
            r("canPonistenjeClean") = Not dep
            r("dependentsText") = "zbirna=" & YesNo(CBool(so("hasZbirna"))) & _
                ", prijemnica=" & YesNo(CBool(so("hasPrijemnica"))) & _
                ", palete=" & YesNo(CBool(so("hasPalete")))
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj)
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
            Dim sp As Object: Set sp = ScanPrijemnica(broj)
            Dim depp As Boolean
            depp = CBool(sp("hasPalete")) Or CBool(sp("fakturisano"))
            r("hasDependents") = depp
            r("canPonistenjeClean") = Not depp
            r("dependentsText") = "palete=" & YesNo(CBool(sp("hasPalete"))) & _
                ", fakturisana=" & YesNo(CBool(sp("fakturisano")))
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".GetChainFlags"
End Function

' ============================================================
' PONISTENJE posledice: PUN spisak zavisnih dokumenata koje poistenje gasi.
' UI ga prikaze PRE nego sto bilo sta uradi (to je ono sto PONISTENJE cini
' razlicitim od DUPLI -- DUPLI tiho pocisti, PONISTENJE prvo pokaze ceo lanac).
' ============================================================
Public Function BuildPonistenjePosledice(ByVal docType As String, ByVal broj As String, _
                                         Optional ByVal dokumentTip As String = "") As String
    On Error GoTo EH
    Dim m As String
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            Dim so As Object: Set so = ScanOtpremnica(broj)
            Dim owo As Boolean: owo = CBool(so("hasZbirna"))
            If owo Then owo = ZbirnaOwnsExternalChain(CStr(so("brojZbirne")))
            m = "PONISTENJE otpremnice " & broj & " gasi tok (STORNO)." & vbCrLf & "Pogodjeno:" & vbCrLf
            m = m & " - zbirna: " & IIf(CBool(so("hasZbirna")), CStr(so("brojZbirne")), "(nema)") & _
                    " (stornira se samo ako je ovo jedina otpremnica; inace ostaje + rekalk)" & vbCrLf
            m = m & " - prijemnice preko zbirne: " & CStr(so("prijCount")) & _
                    IIf(CBool(so("hasZbirna")) And Not owo, " (eksterni kupac -> NETAKNUTE)", "") & vbCrLf
            m = m & " - paletne stavke: " & CStr(so("paleteCount")) & _
                    IIf(CBool(so("hasZbirna")) And owo, " (skidaju se sa paleta; prazna paleta stornirana)", "") & vbCrLf
            m = m & " - otkupni blokovi (OSLOBADJAJU se za reveze, NE storniraju): " & CStr(so("blockCount"))
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj)
            Dim owz As Boolean: owz = ZbirnaOwnsExternalChain(broj)
            m = "PONISTENJE zbirne " & broj & " gasi interni tok (STORNO)." & vbCrLf & "Pogodjeno:" & vbCrLf
            m = m & " - aktivne otpremnice (storniraju se): " & CStr(sz("otpCount")) & vbCrLf
            m = m & " - prijemnice: " & CStr(sz("prijCount")) & _
                    IIf(owz, " (storniraju se)", " (eksterni kupac -> NETAKNUTE)") & vbCrLf
            m = m & " - paletne stavke: " & CStr(sz("paleteCount")) & _
                    IIf(owz, " (skidaju se sa paleta; prazna paleta stornirana)", " (NETAKNUTE)") & vbCrLf
            m = m & " - otkupni blokovi (OSLOBADJAJU se za reveze, NE storniraju)"
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj)
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
                                      Optional ByVal dokumentTip As String = "") As Boolean
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            Dim so As Object: Set so = ScanOtpremnica(broj)
            CorrectionNeedsDialog = CBool(so("hasPrijemnica")) Or CBool(so("hasPalete"))
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj)
            CorrectionNeedsDialog = CBool(sz("hasPrijemnica")) Or CBool(sz("hasPalete"))
        Case FLOW_DOC_REVERS
            CorrectionNeedsDialog = False
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj)
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

' Otpremnica: postojeci StornoOtpremnicaByBroj_TX (u malina modu kaskadira zbirnu)
' + rekalkulacija zbirne AKO je prezivela (non-malina / multi-otpremnica) -> nema
' tihog mismatch-a. Bez context-a (obican storno nema staro->novo).
Public Function RunSimpleStornoOtpremnica(ByVal broj As String) As Object
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoOtpremnica = r
    On Error GoTo EH
    broj = Trim$(broj)
    Dim s As Object: Set s = ScanOtpremnica(broj)
    If Not CBool(s("exists")) Then r("message") = "Aktivna otpremnica nije pronadjena: " & broj: Exit Function
    Dim pz As String: pz = CStr(s("brojZbirne"))

    If Not StornoOtpremnicaByBroj_TX(broj) Then r("message") = "Storno otpremnice nije uspeo.": Exit Function

    ' Zbirna: rekalk na preostale otpremnice; PRAZNA (jedina otpremnica) -> STORNO,
    ' NE aktivna 0/0 -> dosledno DUPLI/PONISTENJE grani (RecalcOrStornoEmptyZbirna_TX).
    ' Malina: StornoOtpremnicaByBroj_TX je vec oborio zbirnu -> helper je tada no-op.
    Dim zbrRek As Boolean, recOk As Boolean: recOk = True
    zbrRek = (Len(pz) > 0 And ZbirnaPostoji(pz))
    If Len(pz) > 0 Then recOk = RecalcOrStornoEmptyZbirna_TX(pz)
    Dim zbrStorn As Boolean: zbrStorn = (zbrRek And Not ZbirnaPostoji(pz))

    r("success") = True
    r("message") = "Otpremnica " & broj & " stornirana." & _
        IIf(zbrStorn, " Zbirna " & pz & " stornirana (bez otpremnica).", _
            IIf(zbrRek, " Zbirna " & pz & " rekalkulisana.", ""))
    If Not recOk Then r("message") = r("message") & " UPOZORENJE: rekalkulacija/storno zbirne nije uspela (vidi Monitor)."
    MonitorSimple "Otpremnica", broj, CStr(r("message"))
    Exit Function
EH:
    LogErr MOD_NAME & ".RunSimpleStornoOtpremnica"
    r("message") = "Greska: " & Err.description
End Function

' Zbirna: storno + odvezivanje otpremnica ("ceka zbirnu") u JEDNOJ transakciji ->
' ne ostaje zbirna koja nije zbir svojih otpremnica (nema tihog mismatch-a).
Public Function RunSimpleStornoZbirna(ByVal broj As String) As Object
    Const SRC As String = MOD_NAME & ".RunSimpleStornoZbirna"
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoZbirna = r
    On Error GoTo EH
    broj = Trim$(broj)
    If Not ZbirnaPostoji(broj) Then r("message") = "Aktivna zbirna nije pronadjena: " & broj: Exit Function

    ' Atomarno (jedna TX): storno zbirne + odvezivanje otpremnica ("ceka zbirnu").
    Dim det As Long
    If Not StornoZbirnaIDetach_TX(broj, det) Then r("message") = "Storno zbirne nije uspeo.": Exit Function

    r("success") = True
    r("message") = "Zbirna " & broj & " stornirana." & _
                   IIf(det > 0, " Otpremnice vracene u 'ceka zbirnu': " & det & ".", "")
    MonitorSimple "Zbirna", broj, CStr(r("message"))
    Exit Function
EH:
    LogErr SRC
    r("message") = "Greska: " & Err.description
End Function

' Revers: obican storno (saldo vec iskljucuje stornirano -> auto koreguje).
Public Function RunSimpleStornoRevers(ByVal brDok As String, ByVal dokumentTip As String) As Object
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoRevers = r
    On Error GoTo EH
    brDok = Trim$(brDok)
    If Not ActiveAmbalazaDokExists(brDok, dokumentTip) Then
        r("message") = "Aktivan revers nije pronadjen: " & brDok & " [" & dokumentTip & "]"
        Exit Function
    End If
    If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip) Then r("message") = "Storno reversa nije uspeo.": Exit Function
    r("success") = True
    r("message") = "Revers " & brDok & " storniran. Saldo azuriran (bez duple/kontra stavke)."
    Exit Function
EH:
    LogErr MOD_NAME & ".RunSimpleStornoRevers"
    r("message") = "Greska: " & Err.description
End Function

' Prijemnica: obican storno (nema paleta/fakture/blokova -> nema odluke). Reuse
' StornoPrijemnicaByBroj_TX (oslobadja fakturu + ambalazu ako ih ima).
Public Function RunSimpleStornoPrijemnica(ByVal broj As String) As Object
    Dim r As Object: Set r = NewRes("SIMPLE")
    Set RunSimpleStornoPrijemnica = r
    On Error GoTo EH
    broj = Trim$(broj)
    Dim s As Object: Set s = ScanPrijemnica(broj)
    If Not CBool(s("exists")) Then r("message") = "Aktivna prijemnica nije pronadjena: " & broj: Exit Function
    If Not StornoPrijemnicaByBroj_TX(broj) Then r("message") = "Storno prijemnice nije uspeo.": Exit Function
    r("success") = True
    r("message") = "Prijemnica " & broj & " stornirana."
    MonitorSimple "Prijemnica", broj, CStr(r("message"))
    Exit Function
EH:
    LogErr MOD_NAME & ".RunSimpleStornoPrijemnica"
    r("message") = "Greska: " & Err.description
End Function

Private Sub MonitorSimple(ByVal entityType As String, ByVal id As String, ByVal msg As String)
    On Error Resume Next
    Monitor_Event eventType:="STORNO_SIMPLE_" & UCase$(entityType), severity:="INFO", _
        message:=msg, moduleName:=MOD_NAME, procedureName:="RunSimpleStorno", _
        entityType:=entityType, entityID:=id, correlationId:=id
End Sub

' ============================================================
' OTPREMNICA - dispatch po modu
' ============================================================
Public Function RunOtpremnicaCorrection(ByVal oldBroj As String, ByVal mode As String, _
                                        Optional ByVal forceConfirm As Boolean = False) As Object
    Const SRC As String = MOD_NAME & ".RunOtpremnicaCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunOtpremnicaCorrection = r
    On Error GoTo EH

    oldBroj = Trim$(oldBroj)
    Dim s As Object: Set s = ScanOtpremnica(oldBroj)
    If Not CBool(s("exists")) Then
        r("message") = "Aktivna otpremnica nije pronadjena: " & oldBroj
        Exit Function
    End If
    Dim parentZbirna As String: parentZbirna = CStr(s("brojZbirne"))

    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, _
                CStr(s("otpID")), oldBroj, , , , FLOW_DOC_ZBIRNA, , parentZbirna, _
                "Parkirano za kasnije resavanje.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE). Vidljiv u: Osiroceni dokumenti."

        Case SV_MODE_ISPRAVKA
            Dim cid As String
            cid = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, CStr(s("otpID")), oldBroj, _
                FLOW_DOC_OTPREMNICA, , , FLOW_DOC_ZBIRNA, , parentZbirna, _
                "Ispravka otpremnice: storno stare, ceka snimanje nove.")
            If Len(cid) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            If Not StornoOtpremnicaBrojAtomic_TX(oldBroj) Then
                FailCorrectionContext cid, "Storno stare otpremnice nije uspeo."
                r("correctionID") = cid: r("message") = "Storno stare otpremnice nije uspeo."
                Exit Function
            End If
            r("correctionID") = cid
            r("needsForm") = True
            r("success") = True
            r("message") = "Stara otpremnica stornirana. Popuni i snimi NOVU otpremnicu; " & _
                           "blokovi i zbirna se prevezuju/rekalkulisu po snimanju."

        Case SV_MODE_DUPLI
            ' DUPLI = fantom: storno otpremnice, OSLOBODI (razvezi) otkup blokove za
            ' reveze, rekalkulisi zbirnu (prazna -> STORNO, nikad aktivna 0/0). NE
            ' kaskadira prijemnicu/palete (to je PONISTENJE) -> ako postoje, ostaju
            ' osirocene uz recovery zabelesku (ne blokira).
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, CStr(s("otpID")), oldBroj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Dupli/fantom otpremnica.")
            If Len(cidD) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            r("correctionID") = cidD
            Dim otpIDsD As Collection: Set otpIDsD = GetOtpremnicaIDsByBroj(oldBroj)
            If Not StornoOtpremnicaBrojAtomic_TX(oldBroj) Then
                FailCorrectionContext cidD, "Storno otpremnice (dupli) nije uspeo."
                r("message") = "Storno otpremnice nije uspeo."
                Exit Function
            End If
            Dim freedD As Long: freedD = FreeOtkupBloksByOtpIDs_TX(otpIDsD)
            Dim recOkD As Boolean: recOkD = True
            If Len(parentZbirna) > 0 Then recOkD = RecalcOrStornoEmptyZbirna_TX(parentZbirna)
            ' Bilo blokova a nijedan nije oslobodjen -> ne lazi clean success (blok bi
            ' ostao na storniranoj otpremnici -> "izgubljen"). MANUAL + Monitor.
            If CLng(s("blockCount")) > 0 And freedD = 0 Then
                MarkCorrectionManual cidD, "Otkupni blokovi NISU oslobodjeni -> prevezi rucno (Osiroceni dokumenti).", _
                    "Otpremnica stornirana; " & CLng(s("blockCount")) & " blokova nije oslobodjeno (free=0)."
                r("success") = True
                r("message") = "Otpremnica stornirana, ali blokovi (" & CLng(s("blockCount")) & ") NISU oslobodjeni. Proveri Osiroceni dokumenti."
                If Not recOkD Then r("message") = r("message") & " UPOZORENJE: rekalkulacija/storno zbirne nije uspela."
                Exit Function
            End If
            If CBool(s("hasPrijemnica")) Or CBool(s("hasPalete")) Then
                MarkCorrectionManual cidD, "Odluci o osirocenoj prijemnici/paletama (reveze ili storno).", _
                    "Fantom otpremnica stornirana; blokovi oslobodjeni: " & freedD & "; prijemnica/palete osirocene."
                r("message") = "Otpremnica stornirana (fantom). Blokovi oslobodjeni: " & freedD & _
                               ". Prijemnica/palete osirocene (Osiroceni dokumenti)."
            Else
                CompleteCorrectionContext cidD, , , "Fantom otpremnica stornirana; blokovi oslobodjeni: " & freedD & "."
                r("message") = "Otpremnica stornirana (fantom). Blokovi oslobodjeni: " & freedD & "."
            End If
            r("success") = True
            If Not recOkD Then r("message") = r("message") & " UPOZORENJE: rekalkulacija/storno zbirne nije uspela."

        Case SV_MODE_PONISTENJE
            ' PONISTENJE = UVEK prvo pun spisak posledica + svesna potvrda (forceConfirm).
            ' Zatim: ako otpremnica EKSKLUZIVNO drzi zbirnu (jedina) -> kaskada celog
            ' toka; deljena zbirna -> ne sme da obori zbirnu (sestre) -> storno otpremnice
            ' + oslobodi blokove + rekalk (za deljenu zbirnu = isto kao DUPLI).
            If Not forceConfirm Then
                r("blocked") = True
                r("message") = BuildPonistenjePosledice(FLOW_DOC_OTPREMNICA, oldBroj, "")
                Exit Function
            End If
            Dim cidP As String
            cidP = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, CStr(s("otpID")), oldBroj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Ponistenje otpremnice bez zamene.")
            r("correctionID") = cidP
            If Len(parentZbirna) > 0 And ZbirnaPostoji(parentZbirna) And OtpremnicaIsSoleOwner(parentZbirna, oldBroj) Then
                Dim ownsP As Boolean: ownsP = ZbirnaOwnsExternalChain(parentZbirna)
                Dim cascP As Object: Set cascP = PonistiZbirnaChain_TX(parentZbirna, ownsP)
                If Not CBool(cascP("ok")) Then
                    FailCorrectionContext cidP, "Kaskadno ponistenje toka zbirne nije uspelo."
                    r("message") = "Ponistenje nije uspelo (kaskada zbirne).": Exit Function
                End If
                CompleteCorrectionContext cidP, , , "Ponistena otpremnica + ceo tok zbirne " & parentZbirna & "."
                r("success") = True
                r("message") = "Otpremnica ponistena sa celim tokom (zbirna " & parentZbirna & "). Otpremnice: " & _
                    CStr(cascP("otp")) & ", prijemnice: " & CStr(cascP("prij")) & ", paletne stavke: " & _
                    CStr(cascP("pals")) & ", blokovi oslobodjeni: " & CStr(cascP("blok")) & _
                    IIf(ownsP, "", " (eksterni kupac: prijemnica netaknuta).")
            Else
                Dim otpIDsP As Collection: Set otpIDsP = GetOtpremnicaIDsByBroj(oldBroj)
                If Not StornoOtpremnicaBrojAtomic_TX(oldBroj) Then
                    FailCorrectionContext cidP, "Storno otpremnice (ponistenje) nije uspeo."
                    r("message") = "Storno otpremnice nije uspeo.": Exit Function
                End If
                Dim freedP As Long: freedP = FreeOtkupBloksByOtpIDs_TX(otpIDsP)
                Dim recPok As Boolean: recPok = True
                If Len(parentZbirna) > 0 Then recPok = RecalcOrStornoEmptyZbirna_TX(parentZbirna)
                If CLng(s("blockCount")) > 0 And freedP = 0 Then
                    MarkCorrectionManual cidP, "Otkupni blokovi NISU oslobodjeni -> prevezi rucno (Osiroceni dokumenti).", _
                        "Otpremnica ponistena; " & CLng(s("blockCount")) & " blokova nije oslobodjeno (free=0)."
                    r("success") = True
                    r("message") = "Otpremnica ponistena, ali blokovi NISU oslobodjeni. Proveri Osiroceni dokumenti.": Exit Function
                End If
                CompleteCorrectionContext cidP, , , "Ponistena otpremnica (deljena zbirna); blokovi oslobodjeni: " & freedP & "."
                r("success") = True
                r("message") = "Otpremnica ponistena. Blokovi oslobodjeni: " & freedP & _
                    "; zbirna " & parentZbirna & " rekalkulisana (deljena -> nije oborena)."
                If Not recPok Then r("message") = r("message") & " UPOZORENJE: rekalkulacija/storno zbirne nije uspela."
            End If

        Case Else
            r("message") = "Nepoznat mod: " & mode
    End Select
    Exit Function
EH:
    LogErr SRC
    r("message") = "Greska: " & Err.description
End Function

' Zavrsi ISPRAVKA_ODMAH otpremnice: relink blokova na novu + rekalkulacija stare
' i nove zbirne. Poziva se posle sto operater snimi NOVU otpremnicu.
Public Function CompleteOtpremnicaIspravka(ByVal correctionID As String, _
                                           ByVal newBroj As String) As Object
    Const SRC As String = MOD_NAME & ".CompleteOtpremnicaIspravka"
    Dim r As Object: Set r = NewRes(SV_MODE_ISPRAVKA)
    Set CompleteOtpremnicaIspravka = r
    On Error GoTo EH

    newBroj = Trim$(newBroj)
    r("correctionID") = correctionID
    Dim oldBroj As String, oldZbirna As String
    oldBroj = GetCorrectionField(correctionID, COL_SV_OLD_BROJ)
    oldZbirna = GetCorrectionField(correctionID, COL_SV_PARENT_BROJ)

    Dim newOtpID As String
    newOtpID = LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, newBroj, COL_OTP_ID)
    If Len(newOtpID) = 0 Then
        MarkCorrectionManual correctionID, "Snimi novu otpremnicu pa ponovi prevezivanje.", _
            "Nova otpremnica " & newBroj & " nije pronadjena kao aktivna."
        r("message") = "Nova otpremnica nije pronadjena: " & newBroj
        Exit Function
    End If
    Dim newZbirna As String
    newZbirna = NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, newOtpID, COL_OTP_BROJ_ZBIRNE))

    ' 1) Relink otkupnih blokova: svi blokovi vezani za ID-jeve stare otpremnice.
    Dim oldIDs As Collection: Set oldIDs = GetOtpremnicaIDsByBroj(oldBroj)
    Dim blokovi As Collection: Set blokovi = GetBlokOtkupIDs(oldIDs)
    Dim k As Long
    For k = 1 To blokovi.count
        If Not ReassignOtkupToOtpremnica_TX(CStr(blokovi(k)), newOtpID) Then
            MarkCorrectionManual correctionID, "Prevezi otkupne blokove rucno (Osiroceni dokumenti).", _
                "Relink bloka " & CStr(blokovi(k)) & " na " & newBroj & " nije uspeo."
            r("message") = "Relink bloka nije uspeo: " & CStr(blokovi(k))
            Exit Function
        End If
    Next k

    ' 2) Zbirna. Dva slucaja:
    '  (a) ISTA zbirna (multi-otpremnica, non-malina) -> samo rekalkulisi.
    '  (b) NOVA zbirna (malina 1:1: nova otpremnica nosi novu zbirnu) -> preseli
    '      PRIJEMNICU (+ paleta-stavke, kroz ReassignPrijemnicaToZbirna_TX) sa stare
    '      na novu, rekalkulisi novu, a staru STORNIRAJ ako je ostala prazna
    '      (NE nuliraj je -> to je bio bug: stara zbirna 0 kg + prijemnica/palete
    '       zaglavljene na njoj).
    Dim recOk As Boolean: recOk = True
    Dim prijMoved As Long: prijMoved = 0
    Dim staraStornirana As Boolean: staraStornirana = False

    If StrComp(oldZbirna, newZbirna, vbTextCompare) = 0 Then
        If Len(newZbirna) > 0 And ZbirnaPostoji(newZbirna) Then _
            recOk = RecalculateZbirnaFromOtpremnice_TX(newZbirna)
    Else
        ' Preseli nizvodni tok (prijemnica + paleta-stavke) sa stare na novu zbirnu.
        If Len(newZbirna) > 0 And ZbirnaPostoji(newZbirna) And Len(oldZbirna) > 0 Then
            Dim prijBrojevi As Collection
            Set prijBrojevi = DistinctActiveValues(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE, oldZbirna)
            Dim p As Long
            For p = 1 To prijBrojevi.count
                If Not ReassignPrijemnicaToZbirna_TX(CStr(prijBrojevi(p)), newZbirna) Then
                    MarkCorrectionManual correctionID, "Prevezi prijemnicu na novu zbirnu rucno (Osiroceni dokumenti).", _
                        "Relink prijemnice " & CStr(prijBrojevi(p)) & " na " & newZbirna & " nije uspeo."
                    r("message") = "Relink prijemnice nije uspeo: " & CStr(prijBrojevi(p))
                    Exit Function
                End If
                prijMoved = prijMoved + 1
            Next p
        End If
        ' Rekalkulacija nove zbirne.
        If Len(newZbirna) > 0 And ZbirnaPostoji(newZbirna) Then recOk = RecalculateZbirnaFromOtpremnice_TX(newZbirna)
        ' Stara zbirna: prazna (nema otpremnica ni prijemnica) -> STORNO; inace rekalkulisi.
        If Len(oldZbirna) > 0 And ZbirnaPostoji(oldZbirna) Then
            If CountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, oldZbirna) > 0 Then
                RecalculateZbirnaFromOtpremnice_TX oldZbirna
            ElseIf CountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, oldZbirna) = 0 Then
                staraStornirana = StornoZbirna_TX(oldZbirna)
            Else
                RecalculateZbirnaFromOtpremnice_TX oldZbirna    ' prijemnica bez cilja -> ne orphanuj
            End If
        End If
    End If

    If Not recOk Then
        MarkCorrectionManual correctionID, "Rekalkulisi zbirnu rucno / proveri Monitor.", _
            "Rekalkulacija zbirne posle ispravke otpremnice nije uspela."
        r("message") = "Rekalkulacija zbirne nije uspela."
        Exit Function
    End If

    ' 3) Validacija NOVE zbirne (mora biti = zbir svojih otpremnica).
    If Len(newZbirna) > 0 And ZbirnaPostoji(newZbirna) Then
        Dim inv As Object: Set inv = ValidateZbirnaInvariant(newZbirna)
        If Not CBool(inv("isValid")) Then
            MarkCorrectionManual correctionID, "Proveri novu zbirnu (mismatch posle ispravke).", CStr(inv("message"))
            r("message") = "Nova zbirna nije konzistentna: " & CStr(inv("message"))
            Exit Function
        End If
    End If

    CompleteCorrectionContext correctionID, newOtpID, newBroj, _
        "Ispravka otpremnice: blokovi prevezani, prijemnica/palete preseljene, stara zbirna " & _
        IIf(staraStornirana, "stornirana", "rekalkulisana") & "."
    r("success") = True
    r("message") = "Ispravka zavrsena. Blokovi prevezani na " & newBroj & _
        IIf(prijMoved > 0, ", prijemnica/palete preseljene na " & newZbirna, "") & _
        IIf(staraStornirana, ", stara zbirna " & oldZbirna & " stornirana", "") & "."
    Exit Function
EH:
    LogErr SRC
    On Error Resume Next
    FailCorrectionContext correctionID, "Greska u CompleteOtpremnicaIspravka: " & Err.description
    r("message") = "Greska: " & Err.description
End Function

' ============================================================
' ZBIRNA - dispatch po modu
' ============================================================
Public Function RunZbirnaCorrection(ByVal broj As String, ByVal mode As String, _
                                    Optional ByVal forceConfirm As Boolean = False) As Object
    Const SRC As String = MOD_NAME & ".RunZbirnaCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunZbirnaCorrection = r
    On Error GoTo EH

    broj = Trim$(broj)
    If Not ZbirnaPostoji(broj) Then
        r("message") = "Aktivna zbirna nije pronadjena: " & broj
        Exit Function
    End If
    Dim s As Object: Set s = ScanZbirna(broj)

    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, _
                , , , , , , "Zbirna parkirana za kasnije.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE)."

        Case SV_MODE_ISPRAVKA
            Dim cid As String
            cid = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, FLOW_DOC_ZBIRNA, , , , , , _
                "Ispravka zbirne: storno stare, ceka snimanje nove.")
            If Len(cid) = 0 Then r("message") = "Ne mogu da kreiram context.": Exit Function
            If Not StornoZbirna_TX(broj) Then
                FailCorrectionContext cid, "Storno stare zbirne nije uspeo."
                r("correctionID") = cid: r("message") = "Storno zbirne nije uspeo."
                Exit Function
            End If
            r("correctionID") = cid
            r("needsForm") = True
            r("success") = True
            r("message") = "Stara zbirna stornirana. Snimi NOVU zbirnu (iz agregata otpremnica); " & _
                           "otpremnice i prijemnica se prevezuju po snimanju."

        Case SV_MODE_DUPLI
            ' DUPLI = razvezi: ATOMARNO (jedna TX) storno zbirne + odvezi otpremnice
            ' ("ceka zbirnu") -> isti obrazac kao RunSimpleStornoZbirna. Otpremnice
            ' (+blokovi) PREZIVLJAVAJU nevezane. Prijemnica/palete se NE storniraju
            ' (to je PONISTENJE) -> ostaju osirocene za reveze (recovery zabeleska).
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, , , , , , , "Dupli/fantom zbirna.")
            r("correctionID") = cidD
            Dim expOtpD As Long: expOtpD = CLng(s("otpCount"))
            Dim detD As Long
            If Not StornoZbirnaIDetach_TX(broj, detD) Then
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
            cidP = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, , , , , , , "Ponistenje zbirne bez zamene.")
            r("correctionID") = cidP
            Dim ownsZ As Boolean: ownsZ = ZbirnaOwnsExternalChain(broj)
            Dim cascZ As Object: Set cascZ = PonistiZbirnaChain_TX(broj, ownsZ)
            If Not CBool(cascZ("ok")) Then
                FailCorrectionContext cidP, "Kaskadno ponistenje zbirne nije uspelo."
                r("message") = "Ponistenje zbirne nije uspelo (kaskada).": Exit Function
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
    LogErr SRC
    r("message") = "Greska: " & Err.description
End Function

' Zavrsi ISPRAVKA_ODMAH zbirne: prevezi otpremnice i prijemnicu (+palete) na novu
' zbirnu, pa rekalkulisi i validiraj novu.
Public Function CompleteZbirnaIspravka(ByVal correctionID As String, _
                                       ByVal newBroj As String) As Object
    Const SRC As String = MOD_NAME & ".CompleteZbirnaIspravka"
    Dim r As Object: Set r = NewRes(SV_MODE_ISPRAVKA)
    Set CompleteZbirnaIspravka = r
    On Error GoTo EH

    newBroj = Trim$(newBroj)
    r("correctionID") = correctionID
    Dim oldBroj As String
    oldBroj = GetCorrectionField(correctionID, COL_SV_OLD_BROJ)
    If Not ZbirnaPostoji(newBroj) Then
        MarkCorrectionManual correctionID, "Snimi novu zbirnu pa ponovi prevezivanje.", _
            "Nova zbirna " & newBroj & " nije aktivna."
        r("message") = "Nova zbirna nije pronadjena: " & newBroj
        Exit Function
    End If

    ' Ako je broj promenjen -> prevezi otpremnice(+otkup) i prijemnice(+palete).
    If StrComp(oldBroj, newBroj, vbTextCompare) <> 0 Then
        RelinkOtpremniceToZbirna_TX oldBroj, newBroj
        Dim prijBrojevi As Collection
        Set prijBrojevi = DistinctActiveValues(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE, oldBroj)
        Dim k As Long
        For k = 1 To prijBrojevi.count
            If Not ReassignPrijemnicaToZbirna_TX(CStr(prijBrojevi(k)), newBroj) Then
                MarkCorrectionManual correctionID, "Prevezi prijemnicu na novu zbirnu rucno (Osiroceni dokumenti).", _
                    "Relink prijemnice " & CStr(prijBrojevi(k)) & " na " & newBroj & " nije uspeo."
                r("message") = "Relink prijemnice nije uspeo: " & CStr(prijBrojevi(k))
                Exit Function
            End If
        Next k
    End If

    ' Rekalkulacija nove zbirne iz (sada prevezanih) otpremnica.
    If Not RecalculateZbirnaFromOtpremnice_TX(newBroj) Then
        MarkCorrectionManual correctionID, "Rekalkulisi novu zbirnu rucno / proveri Monitor.", _
            "Rekalkulacija nove zbirne " & newBroj & " nije uspela."
        r("message") = "Rekalkulacija nove zbirne nije uspela."
        Exit Function
    End If

    Dim inv As Object: Set inv = ValidateZbirnaInvariant(newBroj)
    If Not CBool(inv("isValid")) Then
        MarkCorrectionManual correctionID, "Proveri novu zbirnu (mismatch).", CStr(inv("message"))
        r("message") = "Nova zbirna nije konzistentna: " & CStr(inv("message"))
        Exit Function
    End If

    CompleteCorrectionContext correctionID, "", newBroj, _
        "Ispravka zbirne zavrsena: otpremnice/prijemnica prevezane, zbirna rekalkulisana."
    r("success") = True
    r("message") = "Ispravka zbirne zavrsena. Sve prevezano na " & newBroj & "."
    Exit Function
EH:
    LogErr SRC
    On Error Resume Next
    FailCorrectionContext correctionID, "Greska u CompleteZbirnaIspravka: " & Err.description
    r("message") = "Greska: " & Err.description
End Function

' ============================================================
' REVERS AMBALAZE - dispatch po modu (saldo vec iskljucuje stornirano ->
' storno = uklanjanje iz salda; bez kontra-stavke, bez duplog salda).
' ============================================================
Public Function RunReversCorrection(ByVal brDok As String, ByVal dokumentTip As String, _
                                    ByVal mode As String) As Object
    Const SRC As String = MOD_NAME & ".RunReversCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunReversCorrection = r
    On Error GoTo EH

    brDok = Trim$(brDok)
    If Not ActiveAmbalazaDokExists(brDok, dokumentTip) Then
        r("message") = "Aktivan revers nije pronadjen: " & brDok & " [" & dokumentTip & "]"
        Exit Function
    End If

    Select Case mode
        Case SV_MODE_RESI_KASNIJE
            r("correctionID") = CreateCorrectionContext(mode, FLOW_DOC_REVERS, brDok, brDok, _
                , , , dokumentTip, , , "Revers parkiran za kasnije.")
            r("success") = (Len(CStr(r("correctionID"))) > 0)
            r("message") = "Kreiran recovery zapis (RESI_KASNIJE)."

        Case SV_MODE_ISPRAVKA
            Dim cid As String
            cid = CreateCorrectionContext(mode, FLOW_DOC_REVERS, brDok, brDok, FLOW_DOC_REVERS, , , _
                dokumentTip, , , "Ispravka reversa: storno stari, ceka novi.")
            r("correctionID") = cid
            If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip) Then
                FailCorrectionContext cid, "Storno starog reversa nije uspeo."
                r("message") = "Storno reversa nije uspeo.": Exit Function
            End If
            r("needsForm") = True
            r("success") = True
            r("message") = "Stari revers storniran (uklonjen iz salda). Unesi NOVI revers; " & _
                           "saldo racuna samo novi (bez duple stavke)."

        Case SV_MODE_DUPLI, SV_MODE_PONISTENJE
            Dim cidX As String
            cidX = CreateCorrectionContext(mode, FLOW_DOC_REVERS, brDok, brDok, , , , _
                dokumentTip, , , IIf(mode = SV_MODE_DUPLI, "Dupli/fantom revers.", "Ponistenje reversa."))
            r("correctionID") = cidX
            If Not StornoOMKoopByBrDok_TX(brDok, dokumentTip) Then
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
    LogErr SRC
    r("message") = "Greska: " & Err.description
End Function

' Zavrsi ISPRAVKA reversa: veze novi revers broj u context. Saldo je vec tacan
' (stari storniran, novi aktivan) -> nema dupliranja. Context postaje COMPLETED
' SAMO ako novi revers stvarno postoji kao AKTIVAN (inace MANUAL_REQUIRED).
' dokumentTip se cita iz konteksta (upisan u ParentDocType pri RunReversCorrection).
Public Function CompleteReversIspravka(ByVal correctionID As String, ByVal newBrDok As String) As Object
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

    If Not ActiveAmbalazaDokExists(newBrDok, dokTip) Then
        MarkCorrectionManual correctionID, "Snimi novi revers pa ponovi zavrsetak ispravke.", _
            "Novi revers " & newBrDok & " [" & dokTip & "] nije aktivan."
        r("message") = "Novi revers nije pronadjen kao aktivan. Snimi novi revers pa ponovi zavrsetak ispravke."
        Exit Function
    End If

    CompleteCorrectionContext correctionID, newBrDok, newBrDok, "Ispravka reversa: novi revers " & newBrDok & "."
    r("success") = True
    r("message") = "Ispravka reversa zavrsena. Saldo racuna samo novi revers."
    Exit Function
EH:
    LogErr MOD_NAME & ".CompleteReversIspravka"
    r("message") = "Greska: " & Err.description
End Function

' ============================================================
' PRIJEMNICA - dispatch po modu. Prijemnica je skoro-list: nizvodni tok = paletne
' stavke (DetachOsirocenePaletaStavke) + faktura (oslobadja se u StornoPrijemnica).
' Otkupni blokovi su SAMOSTALNI (vezani preko BrojZbirne) -> NE diraju se automatski;
' operater ih cekira u panelu za dodatni storno (StornoSelectedBlocks_TX, van ovog).
' ISPRAVKA: storno + needsForm; prevezivanje paleta radi save-putanja prijemnice
' (ReassignPaleteToPrijemnica_TX) -> ovde se samo pravi context i stornira stara.
' ============================================================
Public Function RunPrijemnicaCorrection(ByVal broj As String, ByVal mode As String, _
                                        Optional ByVal forceConfirm As Boolean = False, _
                                        Optional ByVal skipPalete As Boolean = False) As Object
    Const SRC As String = MOD_NAME & ".RunPrijemnicaCorrection"
    Dim r As Object: Set r = NewRes(mode)
    Set RunPrijemnicaCorrection = r
    On Error GoTo EH

    broj = Trim$(broj)
    Dim s As Object: Set s = ScanPrijemnica(broj)
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
            If Not StornoPrijemnicaByBroj_TX(broj) Then
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
            If Not StornoPrijemnicaByBroj_TX(broj) Then
                FailCorrectionContext cidD, "Storno prijemnice (dupli) nije uspeo."
                r("message") = "Storno prijemnice nije uspeo.": Exit Function
            End If
            Dim detD As Long, infoD As String
            If CBool(s("hasPalete")) And Not skipPalete Then detD = DetachOsirocenePaletaStavke_TX(broj, infoD)
            If skipPalete And CBool(s("hasPalete")) Then
                CompleteCorrectionContext cidD, , , "Dupli prijemnica stornirana; palete OSTAVLJENE osirocene (izbor operatera)."
                r("message") = "Prijemnica stornirana (dupli). Palete ostavljene osirocene -> Osiroceni dokumenti (Mod: Palete)."
            Else
                CompleteCorrectionContext cidD, , , "Dupli prijemnica stornirana; paletne stavke skinute: " & detD & "."
                r("message") = "Prijemnica stornirana (dupli). Paletne stavke skinute: " & detD & "."
            End If
            r("success") = True

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

            If Len(parentZbirna) > 0 And ZbirnaPostoji(parentZbirna) Then
                Dim ownsP As Boolean: ownsP = ZbirnaOwnsExternalChain(parentZbirna)
                Dim cascP As Object: Set cascP = PonistiZbirnaChain_TX(parentZbirna, ownsP)
                If Not CBool(cascP("ok")) Then
                    FailCorrectionContext cidP, "Kaskadno ponistenje toka (zbirna " & parentZbirna & ") nije uspelo."
                    r("message") = "Ponistenje nije uspelo (kaskada zbirne).": Exit Function
                End If
                ' Eksterni kupac (zbirna ne poseduje prijemnicu u kaskadi) -> prijemnicu
                ' + njene palete storniramo ovde (retko: prijemnica ~ hladnjaca = internal).
                If Not ownsP Then
                    If Len(LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_ID)) > 0 Then
                        StornoPrijemnicaByBroj_TX broj
                        Dim dInfoX As String
                        If CBool(s("hasPalete")) Then DetachOsirocenePaletaStavke_TX broj, dInfoX
                    End If
                End If
                CompleteCorrectionContext cidP, , , "Ponistena prijemnica sa CELIM tokom zbirne " & parentZbirna & "."
                r("success") = True
                r("message") = "Prijemnica ponistena sa CELIM tokom. Zbirna " & parentZbirna & _
                    " (rekalk/storno), otpremnice: " & CStr(cascP("otp")) & ", prijemnice: " & CStr(cascP("prij")) & _
                    ", paletne stavke: " & CStr(cascP("pals")) & ", blokovi oslobodjeni: " & CStr(cascP("blok")) & "."
            Else
                ' Nema zbirne (prijemnica bez BrojZbirne) -> leaf: storno prijemnice + palete.
                If Not StornoPrijemnicaByBroj_TX(broj) Then
                    FailCorrectionContext cidP, "Storno prijemnice (ponistenje) nije uspeo."
                    r("message") = "Storno prijemnice nije uspeo.": Exit Function
                End If
                Dim detP As Long, infoP As String
                If CBool(s("hasPalete")) Then detP = DetachOsirocenePaletaStavke_TX(broj, infoP)
                CompleteCorrectionContext cidP, , , "Ponistena prijemnica (bez zbirne); paletne stavke skinute: " & detP & "."
                r("success") = True
                r("message") = "Prijemnica ponistena. Paletne stavke skinute: " & detP & "."
            End If

        Case Else
            r("message") = "Nepoznat mod: " & mode
    End Select
    Exit Function
EH:
    LogErr SRC
    r("message") = "Greska: " & Err.description
End Function

Private Function ScanPrijemnica(ByVal broj As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanPrijemnica = d
    On Error GoTo EH
    broj = Trim$(broj)
    d("broj") = broj
    Dim prijID As String
    prijID = LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_ID)
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
    d("otpCount") = IIf(Len(bz) > 0, CountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, bz), 0&)
    d("fakturisano") = (UCase$(NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_FAKTURISANO))) = "DA")
    Dim palc As Long: palc = CountActive(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, broj)
    d("paleteCount") = palc
    d("hasPalete") = (palc > 0)
    d("blockCount") = ActiveBlocksForFlow(FLOW_DOC_PRIJEMNICA, broj).count
    Exit Function
EH:
    LogErr MOD_NAME & ".ScanPrijemnica"
End Function

' ============================================================
' PANEL DATA - strukturirani podaci za "Storno / potvrda" overlay (frmDokumenta).
' Zamenjuju MsgBox-preview: chain rows (dotaknuti dokumenti) + block rows (multiselect).
' ============================================================

' Aktivni otkup blokovi (samostalni) vezani za flow dokument. Otpremnica: preko
' OtpremnicaID; Zbirna/Prijemnica: preko BrojZbirne. Za multiselect dodatni storno.
Public Function ActiveBlocksForFlow(ByVal docType As String, ByVal broj As String, _
                                    Optional ByVal dokumentTip As String = "") As Collection
    Dim result As New Collection
    Set ActiveBlocksForFlow = result
    On Error GoTo EH
    broj = Trim$(broj)
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            Set ActiveBlocksForFlow = GetBlokOtkupIDs(GetOtpremnicaIDsByBroj(broj))
        Case FLOW_DOC_ZBIRNA
            Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(broj)
        Case FLOW_DOC_PRIJEMNICA
            Dim bz As String
            bz = NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_BROJ_ZBIRNE))
            If Len(bz) > 0 Then Set ActiveBlocksForFlow = ActiveOtkupIDsByZbirna(bz)
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".ActiveBlocksForFlow"
End Function

' Aktivni OtkupID-jevi za dati BrojZbirne (denormalizovani otkup.BrojZbirne).
Private Function ActiveOtkupIDsByZbirna(ByVal brojZbirne As String) As Collection
    Dim result As New Collection
    Set ActiveOtkupIDsByZbirna = result
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cId As Long, cSt As Long
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    If cZbr = 0 Or cId = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                result.Add Trim$(CStr(data(i, cId)))
            End If
        End If
    Next i
End Function

' Dotaknuti dokumenti (pregled u panelu). Collection nizova(0..2): Dokument|Info|Napomena.
Public Function GetStornoChainRows(ByVal docType As String, ByVal broj As String, _
                                   Optional ByVal dokumentTip As String = "") As Collection
    Dim result As New Collection
    Set GetStornoChainRows = result
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            Dim so As Object: Set so = ScanOtpremnica(broj)
            AddChainRow result, "Otpremnica", broj, "Stornira se (uz ambalazu)"
            If CBool(so("hasZbirna")) Then
                Dim zEff As String
                If OtpremnicaIsSoleOwner(CStr(so("brojZbirne")), broj) Then
                    zEff = "Preracun; storno ako ostane prazna (jedini vlasnik)"
                Else
                    zEff = "Preracun; NE pada (deljena zbirna - sestre ostaju)"
                End If
                AddChainRow result, "Zbirna", CStr(so("brojZbirne")), zEff
            End If
            If CBool(so("hasPrijemnica")) Then AddChainRow result, "Prijemnica", "(" & CStr(so("prijCount")) & ")", "DUPLIKAT: ostaje osirocena (rucno) | PONISTENJE: stornira se"
            If CBool(so("hasPalete")) Then AddChainRow result, "Paletne stavke", "(" & CStr(so("paleteCount")) & ")", "DUPLIKAT: ostaju osirocene (rucno) | PONISTENJE: skidaju se"
            AddChainRow result, "Otkupni blokovi", "(" & CStr(so("blockCount")) & ")", "Oslobadjaju se za reveze (cekiraj za storno)"
        Case FLOW_DOC_ZBIRNA
            Dim sz As Object: Set sz = ScanZbirna(broj)
            AddChainRow result, "Zbirna", broj, "Stornira se"
            AddChainRow result, "Otpremnice", "(" & CStr(sz("otpCount")) & ")", "PONISTENJE: storniraju se | DUPLIKAT: odvezuju se (prezivljavaju)"
            If CBool(sz("hasPrijemnica")) Then AddChainRow result, "Prijemnica", "(" & CStr(sz("prijCount")) & ")", "DUPLIKAT: ostaje osirocena (rucno) | PONISTENJE: stornira se"
            If CBool(sz("hasPalete")) Then AddChainRow result, "Paletne stavke", "(" & CStr(sz("paleteCount")) & ")", "DUPLIKAT: ostaju osirocene (rucno) | PONISTENJE: skidaju se"
            AddChainRow result, "Otkupni blokovi", "", "Oslobadjaju se za reveze (cekiraj za storno)"
        Case FLOW_DOC_PRIJEMNICA
            Dim sp As Object: Set sp = ScanPrijemnica(broj)
            AddChainRow result, "Prijemnica", broj, "Stornira se (uz ambalazu)"
            If Len(CStr(sp("brojZbirne"))) > 0 Then _
                AddChainRow result, "Zbirna", CStr(sp("brojZbirne")), "Preracun; storno ako padne na 0 (ponistenje)"
            If CLng(sp("otpCount")) > 0 Then _
                AddChainRow result, "Otpremnice", "(" & CStr(sp("otpCount")) & ")", "Storniraju se (ponistenje)"
            If CBool(sp("fakturisano")) Then AddChainRow result, "Faktura", "(vezana)", "Oslobadja se (stavke ostaju osirocene)"
            If CBool(sp("hasPalete")) Then AddChainRow result, "Paletne stavke", "(" & CStr(sp("paleteCount")) & ")", "Skidaju se (Duplikat / Ponistenje)"
            AddChainRow result, "Otkupni blokovi", "(" & CStr(sp("blockCount")) & ")", "Samostalni - cekiraj za dodatni storno"
        Case FLOW_DOC_REVERS
            AddChainRow result, "Revers", broj & " [" & dokumentTip & "]", "Stornira se (saldo se koriguje, bez kontra-stavke)"
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".GetStornoChainRows"
End Function

Private Sub AddChainRow(ByRef col As Collection, ByVal dok As String, ByVal info As String, ByVal nap As String)
    Dim row(0 To 2) As Variant
    row(0) = dok: row(1) = info: row(2) = nap
    col.Add row
End Sub

' Otkupni blokovi za multiselect listu. Collection nizova(0..4):
' OtkupID | BrojDokumenta | Kolicina | Klasa | Kooperant.
Public Function GetStornoBlockRows(ByVal docType As String, ByVal broj As String, _
                                   Optional ByVal dokumentTip As String = "") As Collection
    Dim result As New Collection
    Set GetStornoBlockRows = result
    On Error GoTo EH
    Dim ids As Collection: Set ids = ActiveBlocksForFlow(docType, broj, dokumentTip)
    If ids Is Nothing Then Exit Function
    If ids.count = 0 Then Exit Function

    ' Jedan scan tblOtkup + indeks OtkupID->red (umesto 4x LookupValue po bloku).
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cId As Long, cBr As Long, cKol As Long, cKl As Long, cKoop As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cKl = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    If cId = 0 Then Exit Function
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        idx(Trim$(CStr(data(i, cId)))) = i
    Next i

    Dim pdict As Object: Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Dim k As Long
    For k = 1 To ids.count
        Dim oid As String: oid = CStr(ids(k))
        If idx.Exists(oid) Then
            Dim rr As Long: rr = CLng(idx(oid))
            Dim row(0 To 4) As Variant
            row(0) = oid
            row(1) = NzTxC(data, rr, cBr)
            row(2) = NzTxC(data, rr, cKol)
            row(3) = NzTxC(data, rr, cKl)
            Dim koopID As String: koopID = NzTxC(data, rr, cKoop)
            If Not pdict Is Nothing Then
                If pdict.Exists(koopID) Then row(4) = CStr(pdict(koopID)) Else row(4) = koopID
            Else
                row(4) = koopID
            End If
            result.Add row
        End If
    Next k
    Exit Function
EH:
    LogErr MOD_NAME & ".GetStornoBlockRows"
End Function

' Bezbedno citanje celije po indeksu kolone (0 = kolona ne postoji -> "").
Private Function NzTxC(ByVal data As Variant, ByVal r As Long, ByVal c As Long) As String
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

    If WantTip(tipFilter, FLOW_DOC_PRIJEMNICA) Then _
        AddStornoDocs2 result, TBL_PRIJEMNICA, FLOW_DOC_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_DATUM, _
            COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KUPAC, COL_PRJ_VOZAC, COL_PRJ_KOLICINA, tf, kupci, vozaci, stByZbr
    If WantTip(tipFilter, FLOW_DOC_OTPREMNICA) Then _
        AddStornoDocs2 result, TBL_OTPREMNICA, FLOW_DOC_OTPREMNICA, COL_OTP_BROJ, COL_OTP_DATUM, _
            COL_OTP_BROJ_ZBIRNE, "", COL_OTP_VOZAC, COL_OTP_KOLICINA, tf, kupci, vozaci, stByZbr
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
        ByVal tf As String, ByVal kupci As Object, ByVal vozaci As Object, ByVal stByZbr As Object)
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Sub
    Dim cBr As Long, cDa As Long, cZb As Long, cKu As Long, cVo As Long, cKo As Long, cSt As Long
    cBr = GetColumnIndex(tbl, brojCol)
    cDa = GetColumnIndex(tbl, datumCol)
    cZb = GetColumnIndex(tbl, zbirnaCol)
    If Len(kupacCol) > 0 Then cKu = GetColumnIndex(tbl, kupacCol)
    If Len(vozacCol) > 0 Then cVo = GetColumnIndex(tbl, vozacCol)
    cKo = GetColumnIndex(tbl, kolCol)
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
                    Dim zbr As String: zbr = NzTxC(data, i, cZb)
                    Dim kup As String: kup = ""
                    If cKu > 0 Then kup = DictGet2(kupci, NzTxC(data, i, cKu), NzTxC(data, i, cKu))
                    Dim voz As String: voz = ""
                    If cVo > 0 Then voz = DictGet2(vozaci, NzTxC(data, i, cVo), "")
                    Dim mesta As String: mesta = DictGet2(stByZbr, zbr, "")
                    Dim datum As String: datum = FmtDatum(NzTxC(data, i, cDa))
                    Dim kol As String: kol = NzTxC(data, i, cKo)
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
    Dim stanice As Object: Set stanice = BuildLookupDict(TBL_STANICE, "StanicaID", "Naziv")
    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZb As Long, cSt As Long, cStorno As Long
    cZb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
    If cZb = 0 Or cSt = 0 Then Exit Function
    Dim seenPair As Object: Set seenPair = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> "DA" Then
            Dim zbr As String: zbr = Trim$(CStr(data(i, cZb)))
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
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildStationsByZbirna"
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

' Storniraj cekirane otkupne blokove (samostalne realne kupovine) u JEDNOJ TX.
' Reuse modStorno.StornoOtkup (core). Vraca broj storniranih; -1 na gresku.
Public Function StornoSelectedBlocks_TX(ByVal ids As Collection) As Long
    Const SRC As String = MOD_NAME & ".StornoSelectedBlocks_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    If ids Is Nothing Then Exit Function
    If ids.count = 0 Then Exit Function
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    Dim k As Long, n As Long
    For k = 1 To ids.count
        If Not StornoOtkup(CStr(ids(k))) Then
            Err.Raise ERR_STORNO_FW_BASE + 70, SRC, "StornoOtkup (blok) nije uspeo: " & CStr(ids(k))
        End If
        n = n + 1
    Next k
    tx.CommitTx
    Set tx = Nothing
    StornoSelectedBlocks_TX = n
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    StornoSelectedBlocks_TX = -1
End Function

' ============================================================
' PRIVATE - storno / relink / detach TX helpers (reuse core-a, bez malina kaskade)
' ============================================================

' Storniraj SVE aktivne redove otpremnice za broj u JEDNOJ transakciji, preko
' javnog non-TX core-a modStorno.StornoOtpremnica (koji stornira i ambalazu).
' Namerno NE koristi StornoOtpremnicaByBroj_TX (izbegava malina zbirna-kaskadu).
Private Function StornoOtpremnicaBrojAtomic_TX(ByVal broj As String) As Boolean
    Const SRC As String = MOD_NAME & ".StornoOtpremnicaBrojAtomic_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then Exit Function

    Dim ids As Collection: Set ids = New Collection
    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long, cId As Long, cSt As Long
    cBr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, SRC)
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = broj And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            ids.Add Trim$(CStr(data(i, cId)))
        End If
    Next i
    If ids.count = 0 Then Exit Function

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    Dim k As Long
    For k = 1 To ids.count
        If Not StornoOtpremnica(CStr(ids(k))) Then
            Err.Raise ERR_STORNO_FW_BASE + 20, SRC, "StornoOtpremnica nije uspeo: " & CStr(ids(k))
        End If
    Next k
    tx.CommitTx
    Set tx = Nothing
    StornoOtpremnicaBrojAtomic_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    StornoOtpremnicaBrojAtomic_TX = False
End Function

' Prevezi sve aktivne otpremnice (i denormalizovani otkup.BrojZbirne) sa stare na
' novu zbirnu. Vraca broj prevezanih otpremnica redova.
Private Function RelinkOtpremniceToZbirna_TX(ByVal oldZbirna As String, ByVal newZbirna As String) As Long
    Const SRC As String = MOD_NAME & ".RelinkOtpremniceToZbirna_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    oldZbirna = Trim$(oldZbirna): newZbirna = Trim$(newZbirna)
    If Len(oldZbirna) = 0 Or Len(newZbirna) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)

    Dim otpRows As Collection: Set otpRows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = oldZbirna And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            otpRows.Add i
        End If
    Next i
    If otpRows.count = 0 Then Exit Function

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP
    Dim k As Long
    For k = 1 To otpRows.count
        RequireUpdateCell TBL_OTPREMNICA, CLng(otpRows(k)), COL_OTP_BROJ_ZBIRNE, newZbirna, SRC
    Next k
    ' Denormalizovani otkup.BrojZbirne (za aktivne blokove sa starom zbirnom).
    Dim od As Variant: od = GetTableData(TBL_OTKUP)
    If Not IsEmpty(od) Then
        Dim ocZbr As Long, ocSt As Long
        ocZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
        ocSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
        If ocZbr > 0 Then
            Dim j As Long
            For j = 1 To UBound(od, 1)
                If Trim$(CStr(od(j, ocZbr))) = oldZbirna Then
                    If ocSt = 0 Or UCase$(Trim$(CStr(od(j, ocSt)))) <> "DA" Then
                        RequireUpdateCell TBL_OTKUP, j, COL_OTK_BROJ_ZBIRNE, newZbirna, SRC
                    End If
                End If
            Next j
        End If
    End If
    tx.CommitTx
    Set tx = Nothing
    RelinkOtpremniceToZbirna_TX = otpRows.count
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    RelinkOtpremniceToZbirna_TX = 0
End Function

' Telo odvezivanja (bez TX; koristi se unutar vec otvorene transakcije). Aktivne
' otpremnice sa datom zbirnom -> BrojZbirne = "" ("ceka zbirnu"), + otkup denorm.
Private Function DetachOtpremniceInline(ByVal brojZbirne As String, ByVal SRC As String) As Long
    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            RequireUpdateCell TBL_OTPREMNICA, i, COL_OTP_BROJ_ZBIRNE, "", SRC
            n = n + 1
        End If
    Next i
    ' Denormalizovani otkup.BrojZbirne -> takodje prazno.
    Dim od As Variant: od = GetTableData(TBL_OTKUP)
    If Not IsEmpty(od) Then
        Dim ocZbr As Long, ocSt As Long
        ocZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
        ocSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
        If ocZbr > 0 Then
            Dim j As Long
            For j = 1 To UBound(od, 1)
                If Trim$(CStr(od(j, ocZbr))) = brojZbirne Then
                    If ocSt = 0 Or UCase$(Trim$(CStr(od(j, ocSt)))) <> "DA" Then
                        RequireUpdateCell TBL_OTKUP, j, COL_OTK_BROJ_ZBIRNE, "", SRC
                    End If
                End If
            Next j
        End If
    End If
    DetachOtpremniceInline = n
End Function

' Atomarno (JEDNA TX): storno zbirne (core) + odvezivanje otpremnica ("ceka
' zbirnu") + otkup denorm. Jedan izvor istine za "storno+detach zbirne" -> koriste
' ga i RunSimpleStornoZbirna i DUPLI grana (ne dve odvojene transakcije). Vraca
' True na uspeh; outDet = broj odvezanih otpremnica.
Private Function StornoZbirnaIDetach_TX(ByVal broj As String, ByRef outDet As Long) As Boolean
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
    If Not StornoZbirna(broj) Then Err.Raise ERR_STORNO_FW_BASE + 60, SRC, "StornoZbirna nije uspeo."
    outDet = DetachOtpremniceInline(broj, SRC)
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

' Da li je otpremnica JEDINA (aktivna) otpremnica svoje zbirne -> ekskluzivno je
' poseduje (malina 1:1 ili poslednja). Tada PONISTENJE sme da obori ceo tok zbirne;
' deljena zbirna -> ne sme (oborio bi sestre) -> samo rekalk.
Private Function OtpremnicaIsSoleOwner(ByVal parentZbirna As String, ByVal oldBroj As String) As Boolean
    On Error GoTo EH
    parentZbirna = Trim$(parentZbirna): oldBroj = Trim$(oldBroj)
    If Len(parentZbirna) = 0 Then Exit Function
    Dim brojevi As Collection
    Set brojevi = DistinctActiveValues(TBL_OTPREMNICA, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, parentZbirna)
    Dim k As Long
    For k = 1 To brojevi.count
        If StrComp(Trim$(CStr(brojevi(k))), oldBroj, vbTextCompare) <> 0 Then Exit Function
    Next k
    OtpremnicaIsSoleOwner = True
    Exit Function
EH:
    LogErr MOD_NAME & ".OtpremnicaIsSoleOwner"
End Function

' Rekalkulisi zbirnu iz preostalih aktivnih otpremnica; ako ih VISE NEMA -> STORNO
' zbirne (nikad aktivna 0/0 -> to je bio "nuliranje" bug). NE dira prijemnicu/palete
' (mod odlucuje: DUPLI ostavlja osiroceno; PONISTENJE kaskadira zasebno). True=uspeh.
Private Function RecalcOrStornoEmptyZbirna_TX(ByVal broj As String) As Boolean
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then RecalcOrStornoEmptyZbirna_TX = True: Exit Function
    If Not ZbirnaPostoji(broj) Then RecalcOrStornoEmptyZbirna_TX = True: Exit Function
    If CountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, broj) > 0 Then
        RecalcOrStornoEmptyZbirna_TX = RecalculateZbirnaFromOtpremnice_TX(broj)
    Else
        RecalcOrStornoEmptyZbirna_TX = StornoZbirna_TX(broj)
    End If
    Exit Function
EH:
    LogErr MOD_NAME & ".RecalcOrStornoEmptyZbirna_TX"
End Function

' Aktivni OtpremnicaID-jevi za dati BrojZbirne.
Private Function ActiveOtpIDsByZbirna(ByVal brojZbirne As String, ByVal SRC As String) As Collection
    Dim result As New Collection
    Set ActiveOtpIDsByZbirna = result
    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cId As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            result.Add Trim$(CStr(data(i, cId)))
        End If
    Next i
End Function

' Aktivni PrijemnicaID-jevi za dati BrojZbirne (svi redovi, obe klase).
Private Function ActivePrijIDsByZbirna(ByVal brojZbirne As String, ByVal SRC As String) As Collection
    Dim result As New Collection
    Set ActivePrijIDsByZbirna = result
    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cId As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    cSt = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, SRC)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            result.Add Trim$(CStr(data(i, cId)))
        End If
    Next i
End Function

' Oslobodi (razvezi) otkup blokove datih otpremnica ID-jeva: OtpremnicaID="" i
' BrojZbirne="" na AKTIVNIM otkup redovima -> vracaju se u pool (za reveze). Bez TX
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
    Dim cOtp As Long, cSt As Long, cZbr As Long
    cOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If idSet.Exists(Trim$(CStr(data(i, cOtp)))) Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                RequireUpdateCell TBL_OTKUP, i, COL_OTK_OTPREMNICA_ID, "", SRC
                If cZbr > 0 Then RequireUpdateCell TBL_OTKUP, i, COL_OTK_BROJ_ZBIRNE, "", SRC
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
Private Function PonistiZbirnaChain_TX(ByVal brojZbirne As String, ByVal ownsChain As Boolean) As Object
    Const SRC As String = MOD_NAME & ".PonistiZbirnaChain_TX"
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    res("ok") = False: res("otp") = 0&: res("prij") = 0&: res("pals") = 0&: res("blok") = 0&
    Set PonistiZbirnaChain_TX = res
    Dim tx As clsTransaction
    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    ' ID-jeve + prijemnica-brojeve-sa-paletama skupi PRE mutacije.
    Dim otpIDs As Collection: Set otpIDs = ActiveOtpIDsByZbirna(brojZbirne, SRC)
    Dim prijIDs As Collection, prijBrPalete As Collection
    If ownsChain Then
        Set prijIDs = ActivePrijIDsByZbirna(brojZbirne, SRC)
        Set prijBrPalete = DistinctActiveValues(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, brojZbirne)
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

    If ZbirnaPostoji(brojZbirne) Then
        If Not StornoZbirna(brojZbirne) Then _
            Err.Raise ERR_STORNO_FW_BASE + 50, SRC, "StornoZbirna (ponistenje) nije uspeo."
    End If
    Dim k As Long
    For k = 1 To otpIDs.count
        If Not StornoOtpremnica(CStr(otpIDs(k))) Then _
            Err.Raise ERR_STORNO_FW_BASE + 51, SRC, "StornoOtpremnica (ponistenje) nije uspeo: " & CStr(otpIDs(k))
    Next k
    res("otp") = otpIDs.count
    res("blok") = FreeOtkupBloksInline(otpIDs, SRC)
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
Private Function GetOtpremnicaIDsByBroj(ByVal broj As String) As Collection
    Dim result As New Collection
    Set GetOtpremnicaIDsByBroj = result
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then Exit Function
    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long, cId As Long
    cBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    If cBr = 0 Or cId = 0 Then Exit Function
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, id As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = broj Then
            id = Trim$(CStr(data(i, cId)))
            If Len(id) > 0 And Not seen.Exists(id) Then
                seen(id) = True
                result.Add id
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".GetOtpremnicaIDsByBroj"
End Function

' Distinktni AKTIVNI OtkupID-jevi vezani (OtpremnicaID) za dati skup otp ID-jeva.
Private Function GetBlokOtkupIDs(ByVal otpIDs As Collection) As Collection
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
    If IsEmpty(data) Then Exit Function
    Dim cOtp As Long, cId As Long, cSt As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cSt = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    If cOtp = 0 Or cId = 0 Then Exit Function

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
    LogErr MOD_NAME & ".GetBlokOtkupIDs"
End Function

' ============================================================
' PRIVATE - chain scan + generic helpers
' ============================================================

Private Function ScanOtpremnica(ByVal broj As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanOtpremnica = d
    On Error GoTo EH
    broj = Trim$(broj)
    d("broj") = broj
    Dim otpID As String
    otpID = LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_ID)
    d("otpID") = otpID
    d("exists") = (Len(otpID) > 0)
    If Len(otpID) = 0 Then
        d("stanica") = "": d("brojZbirne") = "": d("blockCount") = 0&
        d("hasZbirna") = False: d("hasPrijemnica") = False: d("prijCount") = 0&
        d("hasPalete") = False: d("paleteCount") = 0&
        Exit Function
    End If
    d("stanica") = NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))
    Dim bz As String: bz = NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))
    d("brojZbirne") = bz

    Dim allIDs As Collection: Set allIDs = GetOtpremnicaIDsByBroj(broj)
    d("blockCount") = GetBlokOtkupIDs(allIDs).count

    d("hasZbirna") = (Len(bz) > 0 And ZbirnaPostoji(bz))
    Dim pc As Long: pc = 0
    If Len(bz) > 0 Then pc = CountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, bz)
    d("prijCount") = pc
    d("hasPrijemnica") = (pc > 0)
    Dim palc As Long: palc = 0
    If Len(bz) > 0 Then palc = CountActive(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, bz)
    d("paleteCount") = palc
    d("hasPalete") = (palc > 0)
    Exit Function
EH:
    LogErr MOD_NAME & ".ScanOtpremnica"
End Function

Private Function ScanZbirna(ByVal broj As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanZbirna = d
    On Error GoTo EH
    broj = Trim$(broj)
    d("broj") = broj
    d("otpCount") = CountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, broj)
    Dim pc As Long: pc = CountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, broj)
    d("prijCount") = pc
    d("hasPrijemnica") = (pc > 0)
    Dim palc As Long: palc = CountActive(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, broj)
    d("paleteCount") = palc
    d("hasPalete") = (palc > 0)
    Set d("invariant") = ValidateZbirnaInvariant(broj)
    Exit Function
EH:
    LogErr MOD_NAME & ".ScanZbirna"
    If Not d.Exists("invariant") Then Set d("invariant") = ValidateZbirnaInvariant(broj)
End Function

Private Function ScanRevers(ByVal brDok As String, ByVal dokumentTip As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ScanRevers = d
    On Error GoTo EH
    brDok = Trim$(brDok)
    d("broj") = brDok
    d("exists") = ActiveAmbalazaDokExists(brDok, dokumentTip)
    d("tip") = "": d("kolicina") = 0&: d("smer") = "": d("entitet") = "": d("redova") = 0&
    If Not CBool(d("exists")) Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function
    Dim cDok As Long, cTip As Long, cKol As Long, cSmer As Long, cEnt As Long, cSt As Long, cDokTip As Long
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    cKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    cSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    cEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cSt = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cDok))) = brDok And Trim$(CStr(data(i, cDokTip))) = dokumentTip Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                d("tip") = NzTx(data(i, cTip))
                d("smer") = NzTx(data(i, cSmer))
                d("entitet") = NzTx(data(i, cEnt))
                d("redova") = CLng(d("redova")) + 1
                ' Revers = dvojni upis (Kooperant + Stanica, isti broj/tip) -> NE sabiraj
                ' obe noge; kolicina dokumenta = jedna noga (reprezentativna/veca).
                If IsNumeric(data(i, cKol)) Then
                    If CLng(data(i, cKol)) > CLng(d("kolicina")) Then d("kolicina") = CLng(data(i, cKol))
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".ScanRevers"
End Function

' Broj AKTIVNIH redova gde filterCol = value.
Private Function CountActive(ByVal tblName As String, ByVal filterCol As String, ByVal value As String) As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(tblName)
    If IsEmpty(data) Then Exit Function
    Dim cF As Long, cSt As Long
    cF = GetColumnIndex(tblName, filterCol)
    cSt = GetColumnIndex(tblName, COL_STORNIRANO)
    If cF = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cF))) = value Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then n = n + 1
        End If
    Next i
    CountActive = n
    Exit Function
EH:
    LogErr MOD_NAME & ".CountActive"
End Function

' Distinktne AKTIVNE vrednosti valueCol gde filterCol = filterVal.
Private Function DistinctActiveValues(ByVal tblName As String, ByVal valueCol As String, _
                                      ByVal filterCol As String, ByVal filterVal As String) As Collection
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
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, v As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cF))) = filterVal Then
            If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
                v = Trim$(CStr(data(i, cV)))
                If Len(v) > 0 And Not seen.Exists(v) Then
                    seen(v) = True
                    result.Add v
                End If
            End If
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
