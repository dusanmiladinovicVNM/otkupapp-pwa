Attribute VB_Name = "modBrojevi"
'Attribute VB_Name = "modBrojevi"
Option Explicit

' ============================================================
' modBrojevi -- broj-allocation helperi za OTK, OTP, ZBR.
'
' Format (kanon v6.15):
'   x/ddmmyy[-rb]
'   - "x"      = numericki deo entityID-a bez vodecih nula
'   - "ddmmyy" = lokalni poslovni datum
'   - "-rb"    = -2, -3, ... za drugi i dalje istog dana
'   - prvi u danu: bez "-"
'
' Javni API:
'   SuggestNextBroj(kind, entityID, datum)   -- VBA forma prefill
'   GenerateBrojDokumenta(stanicaID, datum)  -- VBA fallback za ImportRowToTblOtkup
'   GenerateBrojOtpremnice(stanicaID, datum) -- VBA jedinstveni generator za OTP
'   ExtractNumericFromEntityID(entityID)     -- "ST-00001" -> 1
'   ExtractSeqFromBroj(broj)                 -- "1/220526-3" -> 3
'   IsValidBrojFormat(broj)                  -- regex check kanonskog formata
'   FormatBroj(entityID, datum, seq)         -- kompozit
'   ClearSpreadsheetIDCache                  -- reset session cache (retko)
' ============================================================

Private gSheetIDCache As Object

Public Const KIND_OTK As String = "OTK"
Public Const KIND_OTP As String = "OTP"
Public Const KIND_ZBR As String = "ZBR"
Public Const KIND_REV As String = "REV"   ' OM<->koop revers (izdavanje/povrat ambalaze)

' ============================================================
' PUBLIC -- forma prefill
' ============================================================
Public Function SuggestNextBroj(ByVal kind As String, _
                                ByVal entityID As String, _
                                ByVal datum As Date, _
                                Optional ByVal checkRemote As Boolean = True) As String
    Const SRC As String = "SuggestNextBroj"

    On Error GoTo EH

    ' Toggle: kad je auto-generisanje brojeva iskljuceno (Podesavanja), forma ne
    ' dobija predlog -> operater unosi svoj broj. Default ON (modConfig).
    If Not IsAutoBrojDokumenta() Then
        SuggestNextBroj = ""
        Exit Function
    End If

    If Len(Trim$(entityID)) = 0 Then
        SuggestNextBroj = ""
        Exit Function
    End If

    Dim maxLocal As Long
    Dim maxRemote As Long
    Dim nextSeq As Long

    Select Case UCase$(kind)
        Case KIND_OTK
            maxLocal = MaxSeqFromTable(TBL_OTKUP, COL_OTK_BR_DOK, _
                                       COL_OTK_DATUM, COL_OTK_STANICA, _
                                       entityID, datum)
            If checkRemote Then
                maxRemote = MaxSeqFromGoogleSheet("OTK-" & entityID, _
                                                  "BrojDokumenta", datum)
            End If
        Case KIND_OTP
            maxLocal = MaxSeqFromTable(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                       COL_OTP_DATUM, COL_OTP_STANICA, _
                                       entityID, datum)
            maxRemote = 0
        Case KIND_ZBR
            maxLocal = MaxSeqFromTable(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                       COL_ZBR_DATUM, COL_ZBR_VOZAC, _
                                       entityID, datum)
            If checkRemote Then
                maxRemote = MaxSeqFromGoogleSheet("VOZ-" & entityID, _
                                                  "BrojZbirne", datum)
            End If
        Case KIND_REV
            ' Revers (OM<->koop ambalaza): sopstveni dnevni niz po stanici;
            ' scan tblAmbalaza (OM-Izlaz-Koop / OM-Ulaz-Koop, Stanica noga).
            maxLocal = MaxSeqReversAmbalaza(entityID, datum)
            maxRemote = 0
        Case Else
            LogError SRC, "Nepoznata kind vrednost: " & kind
            SuggestNextBroj = ""
            Exit Function
    End Select

    If maxLocal > maxRemote Then
        nextSeq = maxLocal + 1
    Else
        nextSeq = maxRemote + 1
    End If

    SuggestNextBroj = FormatBroj(entityID, datum, nextSeq)

    ' ZBR: mirror-stanica (VozacID==StanicaID) dobija "S" prefiks (S1/ddmmyy) da se
    ' ne sudara sa realnim vozacem istog numerickog dela. Plus bump sekvence dok
    ' predlozeni broj (string) ne bude slobodan u tblZbirna (mreza za legacy).
    If UCase$(kind) = KIND_ZBR Then
        SuggestNextBroj = ApplyMirrorPrefix(entityID, FormatBroj(entityID, datum, nextSeq))
        Do While BrojZbirneExists(SuggestNextBroj)
            nextSeq = nextSeq + 1
            SuggestNextBroj = ApplyMirrorPrefix(entityID, FormatBroj(entityID, datum, nextSeq))
        Loop
    End If
    Exit Function

EH:
    LogErr SRC, "kind=" & kind & " entity=" & entityID
    SuggestNextBroj = ""
End Function

' ============================================================
' PUBLIC -- fallback generatori (PWA broj nedostaje ili je VBA-only)
' ============================================================

' Fallback za OTK kad ImportRowToTblOtkup primi prazan brojDokumenta
' iz PWA recorda (legacy/pre-rollout). Scan samo tblOtkup.
Public Function GenerateBrojDokumenta(ByVal stanicaID As String, _
                                       ByVal datum As Date) As String
    Const SRC As String = "GenerateBrojDokumenta"
    
    On Error GoTo EH

    If ExtractNumericFromEntityID(stanicaID) = 0 Then
        LogError SRC, "Nevazeci stanicaID (bez cifara): " & stanicaID
        GenerateBrojDokumenta = ""
        Exit Function
    End If
    
    Dim maxSeq As Long
    maxSeq = MaxSeqFromTable(TBL_OTKUP, COL_OTK_BR_DOK, _
                             COL_OTK_DATUM, COL_OTK_STANICA, _
                             stanicaID, datum)
    
    GenerateBrojDokumenta = FormatBroj(stanicaID, datum, maxSeq + 1)
    Exit Function

EH:
    LogErr SRC, "stanica=" & stanicaID
    GenerateBrojDokumenta = ""
End Function

' Jedinstveni generator za BrojOtpremnice. Otpremnica je VBA-only entity
' (PWA je ne pravi), scan samo lokalno. Koristi se u:
'   - AutoCreateOtpremniceFromPWA (zamenjuje inline format generaciju)
'   - frmDokumenta manual otpremnica unos
Public Function GenerateBrojOtpremnice(ByVal stanicaID As String, _
                                        ByVal datum As Date) As String
    Const SRC As String = "GenerateBrojOtpremnice"
    
    On Error GoTo EH

    If ExtractNumericFromEntityID(stanicaID) = 0 Then
        LogError SRC, "Nevazeci stanicaID (bez cifara): " & stanicaID
        GenerateBrojOtpremnice = ""
        Exit Function
    End If
    
    Dim maxSeq As Long
    maxSeq = MaxSeqFromTable(TBL_OTPREMNICA, COL_OTP_BROJ, _
                             COL_OTP_DATUM, COL_OTP_STANICA, _
                             stanicaID, datum)
    
    GenerateBrojOtpremnice = FormatBroj(stanicaID, datum, maxSeq + 1)
    Exit Function

EH:
    LogErr SRC, "stanica=" & stanicaID
    GenerateBrojOtpremnice = ""
End Function

' Jedinstveni generator za BrojPrijemnice. Prijemnica je VBA-only entity
' (PWA je ne pravi), scan samo lokalno. Auto-numeracija vazi SAMO za hladnjaca-
' kupca (CFG_MALINA_DEFAULT_KUPAC); ostali kupci nose eksterni, nezavisni broj
' koji se unosi rucno. x-deo je fiksno "1" (konvencija za hladnjacu), NE iz
' KupacID broja; kupacID se koristi samo da ogranici dnevni brojac na hladnjaca-
' kupca (da eksterni "1/..." drugih kupaca ne naduvaju niz). Robustno preko
' MaxSeqFromTable (MAX sekvence), ne brojanjem redova -> dvoklasna prijemnica
' (Kl I + Kl II, isti broj) ne pomera brojac za 2.
' Koristi se u: AutoChainHladnjaca (modAutoHladnjaca).
Public Function GenerateBrojPrijemnice(ByVal kupacID As String, _
                                        ByVal datum As Date) As String
    Const SRC As String = "GenerateBrojPrijemnice"

    On Error GoTo EH

    Dim maxSeq As Long
    maxSeq = MaxSeqFromTable(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                             COL_PRJ_DATUM, COL_PRJ_KUPAC, _
                             kupacID, datum)

    GenerateBrojPrijemnice = FormatBroj("1", datum, maxSeq + 1)
    Exit Function

EH:
    ' AUD-041(a): EH NE SME da vrati validan-looking broj. "1/ddmmyy" je izgledao
    ' kao regularan prvi broj dana, pa je posle greske u skenu (schema drift,
    ' nedostupna tabela) prijemnica dobijala broj koji vec postoji. Prazan string
    ' je jedini bezbedan izlaz -- isto kao GenerateBrojDokumenta /
    ' GenerateBrojOtpremnice; caller (AutoChainHladnjaca) ga vidi kao pad koraka.
    LogErr SRC, "kupac=" & kupacID
    GenerateBrojPrijemnice = ""
End Function

' ============================================================
' PUBLIC -- utility (drugi moduli ih koriste)
' ============================================================

' "VOZ-00004" -> 4 ; "ST-00001" -> 1 ; "ST-103" -> 103 ; "garbage" -> 0
Public Function ExtractNumericFromEntityID(ByVal entityID As String) As Long
    Dim i As Long, ch As String, digits As String
    
    For i = 1 To Len(entityID)
        ch = Mid$(entityID, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i
    
    If Len(digits) = 0 Then
        ExtractNumericFromEntityID = 0
    Else
        ExtractNumericFromEntityID = CLng(digits)
    End If
End Function

' "1/220526" -> 1 ; "1/220526-2" -> 2 ; "" -> 0 ; "garbage" -> 0
Public Function ExtractSeqFromBroj(ByVal broj As String) As Long
    Dim s As String: s = Trim$(broj)
    If Len(s) = 0 Then
        ExtractSeqFromBroj = 0
        Exit Function
    End If
    
    Dim slashPos As Long: slashPos = InStr(s, "/")
    If slashPos = 0 Then
        ExtractSeqFromBroj = 0
        Exit Function
    End If
    
    Dim dashPos As Long: dashPos = InStrRev(s, "-")
    
    If dashPos = 0 Or dashPos < slashPos Then
        ExtractSeqFromBroj = 1   ' bare "x/ddmmyy" forma
        Exit Function
    End If
    
    Dim tail As String: tail = Mid$(s, dashPos + 1)
    If IsNumeric(tail) Then
        ExtractSeqFromBroj = CLng(tail)
    Else
        ExtractSeqFromBroj = 0
    End If
End Function

' Regex check kanonskog formata. Reuse ako se vracas na modMasterSync
' IsValidBrojZbirneFormat -- ista regex pattern.
Public Function IsValidBrojFormat(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "^\d+/\d{6}(-\d+)?$"
    re.Global = False
    IsValidBrojFormat = re.Test(s)
End Function

' Formatuje broj prema kanonu:
'   seq <= 1 -> "X/ddmmyy"
'   seq >= 2 -> "X/ddmmyy-N"
Public Function FormatBroj(ByVal entityID As String, _
                            ByVal datum As Date, _
                            ByVal seq As Long) As String
    Dim numPart As String
    numPart = CStr(ExtractNumericFromEntityID(entityID))
    
    Dim ddmmyy As String
    ddmmyy = Format$(datum, "ddmmyy")
    
    If seq <= 1 Then
        FormatBroj = numPart & "/" & ddmmyy
    Else
        FormatBroj = numPart & "/" & ddmmyy & "-" & seq
    End If
End Function

' Da li je ovaj "vozac" zapravo mirror stanice (VozacID == StanicaID)?
' U malina modu par-vozac ima isti ID kao stanica (npr. "ST-00001").
Public Function IsStanicaMirrorVozac(ByVal vozacID As String) As Boolean
    On Error Resume Next
    If Len(Trim$(vozacID)) = 0 Then Exit Function
    IsStanicaMirrorVozac = _
        (Len(Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", vozacID, "StanicaID"), ""))) > 0)
End Function

' BrojZbirne za mirror-stanicu dobija "S" prefiks (S1/ddmmyy) da se NE sudara sa
' realnim vozacem istog numerickog dela (ST-00001 i VOZ-00001 oba daju "1").
' Realni vozaci ostaju bez prefiksa. Idempotentno (ne dodaje "S" dvaput).
Public Function ApplyMirrorPrefix(ByVal vozacID As String, ByVal broj As String) As String
    ApplyMirrorPrefix = broj
    If Len(broj) = 0 Then Exit Function
    If Left$(broj, 1) = "S" Then Exit Function
    If IsStanicaMirrorVozac(vozacID) Then ApplyMirrorPrefix = "S" & broj
End Function

' Reset sheet ID cache. Zovi ako se OTK-* / VOZ-* sheet rucno preimenuje
' ili obrise tokom rada workbook-a (retko).
Public Sub ClearSpreadsheetIDCache()
    Set gSheetIDCache = Nothing
End Sub

' ============================================================
' PRIVATE -- scan helperi
' ============================================================

' Max sekvenca broja za stanicu+datum nad CELOM tblAmbalaza (svi tipovi/noge).
' Revers deli "x/ddmmyy" namespace sa ostalim ambalaza dokumentima, a btnUnosOMUlaz
' radi CheckDuplicate nad celom tblAmbalaza -> broji se po PREFIKSU broja da auto-broj
' nikad ne kolidira (npr. sa rucnim OM-Ulaz brojem istog prefiksa). OTP-/PRJ-/OTK-
' ID-evi drugih tokova ne odgovaraju prefiksu, pa ne uticu; bare "x/ddmmyy" = seq 1.
Private Function MaxSeqReversAmbalaza(ByVal stanicaID As String, _
                                     ByVal datum As Date) As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long: iBroj = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    If iBroj = 0 Then Exit Function

    Dim base As String
    base = CStr(ExtractNumericFromEntityID(stanicaID)) & "/" & Format$(datum, "ddmmyy")
    Dim baseDash As String: baseDash = base & "-"
    Dim nDash As Long: nDash = Len(baseDash)

    Dim r As Long, best As Long, broj As String, seq As Long
    For r = 1 To UBound(data, 1)
        broj = Trim$(CStr(data(r, iBroj)))
        If broj = base Or (Len(broj) > nDash And Left$(broj, nDash) = baseDash) Then
            seq = ExtractSeqFromBroj(broj)
            If seq > best Then best = seq
        End If
    Next r
    MaxSeqReversAmbalaza = best
    Exit Function
EH:
    LogErr "modBrojevi.MaxSeqReversAmbalaza", "stanica=" & stanicaID
End Function

Private Function MaxSeqFromTable(ByVal tblName As String, _
                                  ByVal colBroj As String, _
                                  ByVal colDatum As String, _
                                  ByVal colEntity As String, _
                                  ByVal entityID As String, _
                                  ByVal datum As Date) As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        MaxSeqFromTable = 0
        Exit Function
    End If
    
    Dim iBroj As Long, iDatum As Long, iEntity As Long
    iBroj = RequireColumnIndex(tblName, colBroj, "MaxSeqFromTable")
    iDatum = RequireColumnIndex(tblName, colDatum, "MaxSeqFromTable")
    iEntity = RequireColumnIndex(tblName, colEntity, "MaxSeqFromTable")
    
    Dim datumStr As String: datumStr = Format$(datum, "ddmmyy")
    Dim maxSeq As Long: maxSeq = 0
    
    Dim r As Long
    For r = 1 To UBound(data, 1)
        If CStr(data(r, iEntity)) = entityID Then
            Dim rowDatum As String
            rowDatum = ""
            On Error Resume Next
            rowDatum = Format$(CDate(data(r, iDatum)), "ddmmyy")
            On Error GoTo EH
            
            If rowDatum = datumStr Then
                Dim broj As String: broj = CStr(data(r, iBroj))
                Dim seq As Long: seq = ExtractSeqFromBroj(broj)
                If seq > maxSeq Then maxSeq = seq
            End If
        End If
    Next r
    
    MaxSeqFromTable = maxSeq
    Exit Function

EH:
    LogErr "MaxSeqFromTable", "tbl=" & tblName & " entity=" & entityID
    MaxSeqFromTable = 0
End Function

' True ako BrojZbirne (tacan string) vec postoji u tblZbirna (bilo koji vozac).
' Koristi se da predlog (KIND_ZBR) ne ponudi vec zauzet broj kada se malina
' zbirne (po StanicaID) i normalne zbirne (po VozacID) preklope u numerickom delu.
Private Function BrojZbirneExists(ByVal broj As String) As Boolean
    On Error GoTo EH

    Dim b As String: b = Trim$(broj)
    If Len(b) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long
    iBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "BrojZbirneExists")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(r, iBroj))), b, vbTextCompare) = 0 Then
            BrojZbirneExists = True
            Exit Function
        End If
    Next r
    Exit Function

EH:
    LogErr "BrojZbirneExists", "broj=" & broj
End Function

Private Function MaxSeqFromGoogleSheet(ByVal sheetName As String, _
                                        ByVal brojColHeader As String, _
                                        ByVal datum As Date) As Long
    On Error GoTo EH
    
    ' DODATO: desktop-only -- ne idemo na Google. Lokal scan je dovoljan.
    If Not IsCloudSyncEnabled() Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If
    
    ' AUD-001: "sheet ne postoji" i "lookup nije uspeo" moraju biti razdvojeni.
    ' Neuspeo Drive lookup ne sme da izgleda kao "remote nema brojeve".
    Dim spreadsheetID As String
    If Not TryResolveSpreadsheetIDByName(sheetName, spreadsheetID) Then
        Err.Raise vbObjectError + 8403, "MaxSeqFromGoogleSheet", _
                  "Drive lookup nije uspeo. Broj se ne predlaze da se ne bi dodelio duplikat. Sheet=" & sheetName
    End If

    If Len(spreadsheetID) = 0 Then
        ' Lookup je prosao -- remote sheet stvarno ne postoji, nema brojeva.
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If

    ' AUD-001: read/parse greska NE SME da izgleda kao "na Google-u nema redova".
    ' Remote scan postoji upravo da bi uhvatio jos neuvezene PWA dokumente;
    ' ako padne, tihi 0 bi mogao da ponudi vec zauzet broj (duplikat).
    ' Zato podizemo gresku van lokalnog handlera -- SuggestNextBroj je hvata
    ' i bezbedno vraca "" (operater unosi broj rucno).
    Dim data As Variant
    If Not TryReadSheetData(spreadsheetID, "Sheet1", data) Then
        Err.Raise vbObjectError + 8402, "MaxSeqFromGoogleSheet", _
                  "Remote scan brojeva nije uspeo (HTTP/JSON). Sheet=" & sheetName
    End If

    If IsEmpty(data) Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If

    Dim iBroj As Long, iDatum As Long
    iBroj = FindHeaderIndexInData(data, brojColHeader)
    iDatum = FindHeaderIndexInData(data, "Datum")

    ' AUD-001: sheet je procitan, ali nema ocekivane kolone -> remote scan
    ' nije obavljen. Tih 0 bi ovde takode mogao dati vec zauzet broj.
    If iBroj = 0 Or iDatum = 0 Then
        Err.Raise vbObjectError + 8404, "MaxSeqFromGoogleSheet", _
                  "Remote sheet nema ocekivane headere (" & brojColHeader & "/Datum). Sheet=" & sheetName
    End If

    Dim datumStr As String: datumStr = Format$(datum, "ddmmyy")
    Dim maxSeq As Long: maxSeq = 0
    
    Dim r As Long
    For r = 2 To UBound(data, 1)
        Dim rowDatum As String
        rowDatum = ""
        On Error Resume Next
        rowDatum = Format$(CDate(data(r, iDatum)), "ddmmyy")
        On Error GoTo EH
        
        If rowDatum = datumStr Then
            Dim broj As String: broj = CStr(data(r, iBroj))
            Dim seq As Long: seq = ExtractSeqFromBroj(broj)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next r
    
    MaxSeqFromGoogleSheet = maxSeq
    Exit Function

EH:
    ' AUD-001: NIJEDNA greska ne sme da se pretvori u 0.
    ' 0 znaci "remote nema brojeve", pa bi svaka progutana greska (config,
    ' cache, neocekivani podaci, buduce izmene) mogla da ponudi vec zauzet
    ' broj. Greska se loguje i propagira -- SuggestNextBroj je hvata i vraca "".
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    If errNum = 0 Then errNum = vbObjectError + 8405
    If Len(errSrc) = 0 Then errSrc = "MaxSeqFromGoogleSheet"

    ' On Error GoTo 0 brise Err, zato su vrednosti vec sacuvane iznad.
    On Error GoTo 0

    LogError "MaxSeqFromGoogleSheet", _
             "Remote scan prekinut. Broj se ne predlaze. Sheet=" & sheetName & _
             "; Err=" & CStr(errNum) & "; Desc=" & errDesc & "; Src=" & errSrc, _
             errNum

    Err.Raise errNum, errSrc, errDesc
End Function

' AUD-001 FAIL-CLOSED lookup.
'
'   True  + outID = ID -> sheet pronaden
'   True  + outID = "" -> lookup je prosao, sheet stvarno ne postoji
'   False              -> lookup NIJE uspeo (nema foldera/tokena, HTTP greska,
'                         necitljiv Drive JSON, exception)
'
' Kesira se SAMO neprazan ID. Prazan rezultat se NE kesira -- ni posle neuspelog
' lookup-a (prolazni Drive problem bi trajno izgledao kao "nema brojeva"), ni
' posle uspesnog "ne postoji" (PWA moze kreirati sheet dok je Excel otvoren, pa
' bi kesirano "" sakrilo nove remote brojeve do restarta).
Private Function TryResolveSpreadsheetIDByName(ByVal sheetName As String, _
                                               ByRef outID As String) As Boolean
    Dim folderID As String
    Dim spreadsheetID As String

    outID = ""

    If gSheetIDCache Is Nothing Then
        Set gSheetIDCache = CreateObject("Scripting.Dictionary")
    End If

    If gSheetIDCache.Exists(sheetName) Then
        outID = CStr(gSheetIDCache(sheetName))
        TryResolveSpreadsheetIDByName = True
        Exit Function
    End If

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(folderID) = 0 Then
        LogError "TryResolveSpreadsheetIDByName", _
                 "GOOGLE_PWA_FOLDER_ID nije postavljen, a cloud sync je ukljucen -- remote scan nije moguc."
        Exit Function
    End If

    If Not TryGetSpreadsheetID(sheetName, folderID, spreadsheetID) Then Exit Function

    If Len(spreadsheetID) > 0 Then gSheetIDCache(sheetName) = spreadsheetID

    outID = spreadsheetID
    TryResolveSpreadsheetIDByName = True
End Function

' ============================================================
' DEV/SMOKE TEST HOOKS
' ============================================================

Public Function TestHook_MaxSeqFromGoogleSheet(ByVal sheetName As String, _
                                               ByVal brojColHeader As String, _
                                               ByVal datum As Date) As Long
    ' DEV/SMOKE TEST HOOK ONLY.
    ' Drzi MaxSeqFromGoogleSheet privatnim za produkciju, ali dozvoljava
    ' regresioni test da greska propagira umesto da postane 0.

    TestHook_MaxSeqFromGoogleSheet = MaxSeqFromGoogleSheet(sheetName, brojColHeader, datum)
End Function

Public Function TestHook_SpreadsheetIDCacheContains(ByVal sheetName As String) As Boolean
    ' DEV/SMOKE TEST HOOK ONLY -- provera da nema negativnog kesiranja.

    If gSheetIDCache Is Nothing Then Exit Function

    TestHook_SpreadsheetIDCacheContains = gSheetIDCache.Exists(sheetName)
End Function

Private Function FindHeaderIndexInData(ByVal data As Variant, _
                                        ByVal headerName As String) As Long
    If IsEmpty(data) Then
        FindHeaderIndexInData = 0
        Exit Function
    End If
    
    Dim target As String: target = Trim$(headerName)
    
    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If CStr(nz(data(LBound(data, 1), c), "")) = target Then
            FindHeaderIndexInData = c
            Exit Function
        End If
    Next c
    
    FindHeaderIndexInData = 0
End Function

' Lokalna Nz (postojeca je Private u modMasterSync)
Private Function nz(ByVal v As Variant, _
                     Optional ByVal Fallback As Variant = "") As Variant
    If IsNull(v) Then
        nz = Fallback
    ElseIf IsEmpty(v) Then
        nz = Fallback
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        nz = Fallback
    Else
        nz = v
    End If
End Function

