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

' ============================================================
' PREVIEW - multiline tekst za dijalog (UI ga prikaze u MsgBox-u).
' ============================================================
Public Function BuildStornoPreview(ByVal docType As String, ByVal broj As String, _
                                   Optional ByVal dokumentTip As String = "") As String
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_OTPREMNICA: BuildStornoPreview = PreviewOtpremnica(broj)
        Case FLOW_DOC_ZBIRNA:     BuildStornoPreview = PreviewZbirna(broj)
        Case FLOW_DOC_REVERS:     BuildStornoPreview = PreviewRevers(broj, dokumentTip)
        Case Else:                BuildStornoPreview = "Dokument: " & docType & " " & broj
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
    m = m & "Kolicina: " & CStr(s("kolicina")) & vbCrLf
    m = m & "Smer: " & CStr(s("smer")) & vbCrLf
    m = m & "Uticaj na saldo: storno iskljucuje ovaj revers iz salda (bez duple stavke)."
    PreviewRevers = m
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
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".GetChainFlags"
End Function

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
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, CStr(s("otpID")), oldBroj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Dupli/fantom otpremnica.")
            If Len(cidD) = 0 Then r("message") = "Ne mogu da kreiram correction context.": Exit Function
            r("correctionID") = cidD
            If Not StornoOtpremnicaBrojAtomic_TX(oldBroj) Then
                FailCorrectionContext cidD, "Storno otpremnice (dupli) nije uspeo."
                r("message") = "Storno otpremnice nije uspeo."
                Exit Function
            End If
            ' Zbirna mora ostati = zbir preostalih aktivnih otpremnica.
            Dim recOk As Boolean: recOk = True
            If Len(parentZbirna) > 0 And ZbirnaPostoji(parentZbirna) Then
                recOk = RecalculateZbirnaFromOtpremnice_TX(parentZbirna)
            End If
            ' Otkupni blokovi ostaju vezani za storniran(u) otpremnicu -> izgubljeni,
            ' cekaju svesno prevezivanje. To NIJE tihi mismatch: MANUAL_REQUIRED + Monitor.
            If CLng(s("blockCount")) > 0 Then
                MarkCorrectionManual cidD, "Prevezi otkupne blokove na aktivnu otpremnicu (Osiroceni dokumenti).", _
                    "Otpremnica stornirana; " & CStr(s("blockCount")) & " otkupnih blokova cekaju prevezivanje."
                r("message") = "Otpremnica stornirana. Blokovi (" & CStr(s("blockCount")) & _
                               ") oznaceni za prevezivanje (Osiroceni dokumenti)."
            Else
                CompleteCorrectionContext cidD, , , "Fantom otpremnica stornirana; nema blokova."
                r("message") = "Otpremnica stornirana (fantom)."
            End If
            r("success") = True
            If Not recOk Then r("message") = r("message") & " UPOZORENJE: rekalkulacija zbirne nije uspela."

        Case SV_MODE_PONISTENJE
            Dim dep As Boolean
            dep = CBool(s("hasZbirna")) Or CBool(s("hasPrijemnica")) Or CBool(s("hasPalete"))
            If dep And Not forceConfirm Then
                r("blocked") = True
                r("message") = "BLOKADA: postoje zavisni dokumenti (" & _
                    "zbirna=" & YesNo(CBool(s("hasZbirna"))) & ", prijemnica=" & YesNo(CBool(s("hasPrijemnica"))) & _
                    ", palete=" & YesNo(CBool(s("hasPalete"))) & "). Ponistenje bez zamene trazi svesnu potvrdu."
                Exit Function
            End If
            Dim cidP As String
            cidP = CreateCorrectionContext(mode, FLOW_DOC_OTPREMNICA, CStr(s("otpID")), oldBroj, _
                , , , FLOW_DOC_ZBIRNA, , parentZbirna, "Ponistenje otpremnice bez zamene.")
            r("correctionID") = cidP
            If Not StornoOtpremnicaBrojAtomic_TX(oldBroj) Then
                FailCorrectionContext cidP, "Storno otpremnice (ponistenje) nije uspeo."
                r("message") = "Storno otpremnice nije uspeo."
                Exit Function
            End If
            Dim recP As Boolean: recP = True
            If Len(parentZbirna) > 0 And ZbirnaPostoji(parentZbirna) Then
                recP = RecalculateZbirnaFromOtpremnice_TX(parentZbirna)
            End If
            If dep Then
                MarkCorrectionManual cidP, "Proveri zbirnu/prijemnicu/palete (svesno ponistenje uz zavisni tok).", _
                    "Ponistena otpremnica sa aktivnim zavisnim tokom; zbirna rekalkulisana."
                r("message") = "Otpremnica ponistena (uz potvrdu). Zbirna rekalkulisana; proveri prijemnicu/palete."
            Else
                CompleteCorrectionContext cidP, , , "Ponistena otpremnica bez zavisnog toka."
                r("message") = "Otpremnica ponistena."
            End If
            r("success") = True
            If Not recP Then r("message") = r("message") & " UPOZORENJE: rekalkulacija zbirne nije uspela."

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

    ' 2) Rekalkulacija stare i nove zbirne (izvor istine = otpremnice).
    Dim recOk As Boolean: recOk = True
    Dim done As Object: Set done = CreateObject("Scripting.Dictionary")
    recOk = RecalcIfNeeded(oldZbirna, done) And recOk
    recOk = RecalcIfNeeded(newZbirna, done) And recOk
    If Not recOk Then
        MarkCorrectionManual correctionID, "Rekalkulisi zbirnu rucno / proveri Monitor.", _
            "Rekalkulacija zbirne posle ispravke otpremnice nije uspela."
        r("message") = "Rekalkulacija zbirne nije uspela."
        Exit Function
    End If

    ' 3) Validacija OBE zbirne (stara i nova ne smeju ostati u mismatch-u).
    Dim impact As Object: Set impact = ValidateOtpremnicaZbirnaImpact(oldZbirna, newZbirna)
    If Not CBool(impact("bothValid")) Then
        MarkCorrectionManual correctionID, "Proveri zbirnu (mismatch posle ispravke).", _
            "Posle ispravke otpremnice zbirna nije = zbir otpremnica."
        r("message") = "Zbirna nije konzistentna posle ispravke. Oznaceno za recovery."
        Exit Function
    End If

    CompleteCorrectionContext correctionID, newOtpID, newBroj, _
        "Ispravka otpremnice zavrsena: blokovi prevezani, zbirna rekalkulisana."
    r("success") = True
    r("message") = "Ispravka zavrsena. Blokovi prevezani na " & newBroj & ", zbirna rekalkulisana."
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
            Dim cidD As String
            cidD = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, , , , , , , "Dupli/fantom zbirna.")
            r("correctionID") = cidD
            If Not StornoZbirna_TX(broj) Then
                FailCorrectionContext cidD, "Storno zbirne (dupli) nije uspeo."
                r("message") = "Storno zbirne nije uspeo.": Exit Function
            End If
            ' Odvezi otpremnice: vrati u stanje "ceka zbirnu" (BrojZbirne = "").
            Dim det As Long: det = DetachOtpremniceFromZbirna_TX(broj)
            If CBool(s("hasPrijemnica")) Then
                MarkCorrectionManual cidD, "Prevezi/odvezi prijemnicu sa stornirane zbirne (Osiroceni dokumenti).", _
                    "Fantom zbirna stornirana; odvezano otpremnica: " & det & "; prijemnica ceka odluku."
                r("message") = "Zbirna stornirana; " & det & " otpremnica vraceno u 'ceka zbirnu'. Prijemnica: proveri."
            Else
                CompleteCorrectionContext cidD, , , "Fantom zbirna stornirana; otpremnice odvezane: " & det & "."
                r("message") = "Zbirna stornirana; " & det & " otpremnica vraceno u 'ceka zbirnu'."
            End If
            r("success") = True

        Case SV_MODE_PONISTENJE
            Dim depz As Boolean
            depz = (CLng(s("otpCount")) > 0) Or CBool(s("hasPrijemnica")) Or CBool(s("hasPalete"))
            If depz And Not forceConfirm Then
                r("blocked") = True
                r("message") = "BLOKADA: zbirna ima zavisni tok (otpremnice=" & CStr(s("otpCount")) & _
                    ", prijemnica=" & YesNo(CBool(s("hasPrijemnica"))) & ", palete=" & YesNo(CBool(s("hasPalete"))) & _
                    "). Trazi svesnu potvrdu o zavisnim dokumentima."
                Exit Function
            End If
            Dim cidP As String
            cidP = CreateCorrectionContext(mode, FLOW_DOC_ZBIRNA, "", broj, , , , , , , "Ponistenje zbirne bez zamene.")
            r("correctionID") = cidP
            If Not StornoZbirna_TX(broj) Then
                FailCorrectionContext cidP, "Storno zbirne (ponistenje) nije uspeo."
                r("message") = "Storno zbirne nije uspeo.": Exit Function
            End If
            ' Svesno ponistenje uz zavisni tok: odvezi otpremnice + oznaci prijemnicu/palete za recovery.
            Dim detP As Long: detP = DetachOtpremniceFromZbirna_TX(broj)
            If depz Then
                MarkCorrectionManual cidP, "Odluci o prijemnici/paletama stornirane zbirne (Osiroceni dokumenti).", _
                    "Ponistena zbirna sa zavisnim tokom; otpremnice odvezane: " & detP & "."
                r("message") = "Zbirna ponistena (uz potvrdu). Otpremnice odvezane; prijemnica/palete za recovery."
            Else
                CompleteCorrectionContext cidP, , , "Ponistena prazna zbirna."
                r("message") = "Zbirna ponistena."
            End If
            r("success") = True

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

' Zavrsi ISPRAVKA reversa: povezi novi revers broj u context. Saldo je vec tacan
' (stari storniran, novi aktivan) -> nema dupliranja.
Public Function CompleteReversIspravka(ByVal correctionID As String, ByVal newBrDok As String) As Object
    Dim r As Object: Set r = NewRes(SV_MODE_ISPRAVKA)
    Set CompleteReversIspravka = r
    On Error GoTo EH
    r("correctionID") = correctionID
    CompleteCorrectionContext correctionID, newBrDok, newBrDok, "Ispravka reversa: novi revers " & newBrDok & "."
    r("success") = True
    r("message") = "Ispravka reversa zavrsena. Saldo racuna samo novi revers."
    Exit Function
EH:
    LogErr MOD_NAME & ".CompleteReversIspravka"
    r("message") = "Greska: " & Err.description
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

' Odvezi otpremnice sa zbirne -> "ceka zbirnu" (BrojZbirne = ""), + otkup denorm.
' Zaseban od RelinkOtpremniceToZbirna_TX jer taj ima guard na prazan cilj.
' Vraca broj odvezanih otpremnica redova.
Private Function DetachOtpremniceFromZbirna_TX(ByVal brojZbirne As String) As Long
    Const SRC As String = MOD_NAME & ".DetachOtpremniceFromZbirna_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cZbr As Long, cSt As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)

    Dim otpRows As Collection: Set otpRows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne And UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
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
        RequireUpdateCell TBL_OTPREMNICA, CLng(otpRows(k)), COL_OTP_BROJ_ZBIRNE, "", SRC
    Next k
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
    tx.CommitTx
    Set tx = Nothing
    DetachOtpremniceFromZbirna_TX = otpRows.count
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    DetachOtpremniceFromZbirna_TX = 0
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
    d("tip") = "": d("kolicina") = 0&: d("smer") = "": d("entitet") = ""
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
                If IsNumeric(data(i, cKol)) Then d("kolicina") = CLng(d("kolicina")) + CLng(data(i, cKol))
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

Private Function RecalcIfNeeded(ByVal broj As String, ByRef done As Object) As Boolean
    RecalcIfNeeded = True
    broj = Trim$(broj)
    If Len(broj) = 0 Then Exit Function
    If done.Exists(broj) Then Exit Function
    done(broj) = True
    If Not ZbirnaPostoji(broj) Then Exit Function
    RecalcIfNeeded = RecalculateZbirnaFromOtpremnice_TX(broj)
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
