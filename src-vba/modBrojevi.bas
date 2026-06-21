Attribute VB_Name = "modBrojevi"
'Attribute VB_Name = "modBrojevi"
Option Explicit

' ============================================================
' modBrojevi — broj-allocation helperi za OTK, OTP, ZBR.
'
' Format (kanon v6.15):
'   x/ddmmyy[-rb]
'   - "x"      = numericki deo entityID-a bez vodecih nula
'   - "ddmmyy" = lokalni poslovni datum
'   - "-rb"    = -2, -3, ... za drugi i dalje istog dana
'   - prvi u danu: bez "-"
'
' Javni API:
'   SuggestNextBroj(kind, entityID, datum)   — VBA forma prefill
'   GenerateBrojDokumenta(stanicaID, datum)  — VBA fallback za ImportRowToTblOtkup
'   GenerateBrojOtpremnice(stanicaID, datum) — VBA jedinstveni generator za OTP
'   ExtractNumericFromEntityID(entityID)     — "ST-00001" -> 1
'   ExtractSeqFromBroj(broj)                 — "1/220526-3" -> 3
'   IsValidBrojFormat(broj)                  — regex check kanonskog formata
'   FormatBroj(entityID, datum, seq)         — kompozit
'   ClearSpreadsheetIDCache                  — reset session cache (retko)
' ============================================================

Private gSheetIDCache As Object

Public Const KIND_OTK As String = "OTK"
Public Const KIND_OTP As String = "OTP"
Public Const KIND_ZBR As String = "ZBR"

' ============================================================
' PUBLIC — forma prefill
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
' PUBLIC — fallback generatori (PWA broj nedostaje ili je VBA-only)
' ============================================================

' Fallback za OTK kad ImportRowToTblOtkup primi prazan brojDokumenta
' iz PWA recorda (legacy/pre-rollout). Scan samo tblOtkup.
Public Function GenerateBrojDokumenta(ByVal stanicaID As String, _
                                       ByVal datum As Date) As String
    Const SRC As String = "GenerateBrojDokumenta"
    
    On Error GoTo EH

    If ExtractNumericFromEntityID(stanicaID) = 0 Then
        LogError SRC, "Nevažeci stanicaID (bez cifara): " & stanicaID
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
        LogError SRC, "Nevažeci stanicaID (bez cifara): " & stanicaID
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
    LogErr SRC, "kupac=" & kupacID
    GenerateBrojPrijemnice = "1/" & Format$(datum, "ddmmyy")
End Function

' ============================================================
' PUBLIC — utility (drugi moduli ih koriste)
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

' Regex check kanonskog formata. Reuse ako se vracaš na modMasterSync
' IsValidBrojZbirneFormat — ista regex pattern.
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
        (Len(Trim$(Nz(LookupValue(TBL_STANICE, "StanicaID", vozacID, "StanicaID"), ""))) > 0)
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
' ili obriše tokom rada workbook-a (retko).
Public Sub ClearSpreadsheetIDCache()
    Set gSheetIDCache = Nothing
End Sub

' ============================================================
' PRIVATE — scan helperi
' ============================================================

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
    
    ' DODATO: desktop-only — ne idemo na Google. Lokal scan je dovoljan.
    If Not IsCloudSyncEnabled() Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If
    
    Dim spreadsheetID As String
    spreadsheetID = ResolveSpreadsheetIDByName(sheetName)
    If Len(spreadsheetID) = 0 Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If
    
    Dim data As Variant
    data = ReadSheetData(spreadsheetID, "Sheet1")
    If IsEmpty(data) Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
    End If
    
    Dim iBroj As Long, iDatum As Long
    iBroj = FindHeaderIndexInData(data, brojColHeader)
    iDatum = FindHeaderIndexInData(data, "Datum")
    
    If iBroj = 0 Or iDatum = 0 Then
        MaxSeqFromGoogleSheet = 0
        Exit Function
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
    LogErr "MaxSeqFromGoogleSheet", "sheet=" & sheetName
    MaxSeqFromGoogleSheet = 0
End Function

Private Function ResolveSpreadsheetIDByName(ByVal sheetName As String) As String
    If gSheetIDCache Is Nothing Then
        Set gSheetIDCache = CreateObject("Scripting.Dictionary")
    End If
    
    If gSheetIDCache.Exists(sheetName) Then
        ResolveSpreadsheetIDByName = CStr(gSheetIDCache(sheetName))
        Exit Function
    End If
    
    Dim folderID As String
    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(folderID) = 0 Then
        ResolveSpreadsheetIDByName = ""
        Exit Function
    End If
    
    Dim spreadsheetID As String
    spreadsheetID = GetSpreadsheetID(sheetName, folderID)
    
    gSheetIDCache(sheetName) = spreadsheetID
    ResolveSpreadsheetIDByName = spreadsheetID
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
        If CStr(Nz(data(LBound(data, 1), c), "")) = target Then
            FindHeaderIndexInData = c
            Exit Function
        End If
    Next c
    
    FindHeaderIndexInData = 0
End Function

' Lokalna Nz (postojeca je Private u modMasterSync)
Private Function Nz(ByVal v As Variant, _
                     Optional ByVal Fallback As Variant = "") As Variant
    If IsNull(v) Then
        Nz = Fallback
    ElseIf IsEmpty(v) Then
        Nz = Fallback
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        Nz = Fallback
    Else
        Nz = v
    End If
End Function

