Option Explicit

' ============================================================
' modIsplatePregled v6.18+
'
' Read-only helper sloj za pregled otvorenih otkup blokova
' (per-kooperant aggregation + already-paid awareness + TR check).
'
' Bazira se na postojecim helperima:
'   - GetOpenOtkupi (extended 7-col shape)
'   - GetKooperantNaziv (modBankaMapiranje)
'   - LookupValue za TekuciRacun
'
' NE pise nista u tblNovac. CSV generisanje + selection ide u
' kasnijim commit-ima.
' ============================================================

'======================================================================
' BuildBlokIsplataList
'
' Vraca Collection of BlokIsplata, sve sistemski-otvoreni blokovi
' enrich-ovani sa kooperant TR + KooperantNaziv.
'
' Filteri:
'   - datumOd / datumDo: opciono date range (#1/1/1900# = no filter)
'   - stanicaIDFilter: opciono stanica filter ("" = sve stanice)
'
' Performance: jedan poziv GetOpenOtkupi() (single-pass),
' jedan poziv BuildIsplataDictByOtkup() (sadrzi se u GetOpenOtkupi),
' N lookup-a za TR i KooperantNaziv (cached preko Dictionary-ja).
'======================================================================
Public Function BuildBlokIsplataList( _
    Optional ByVal datumOd As Date, _
    Optional ByVal datumDo As Date, _
    Optional ByVal stanicaIDFilter As String = "" _
) As Collection
    
    Dim result As New Collection
    Dim openOtkupi As Variant
    
    openOtkupi = GetOpenOtkupi("")
    
    If IsEmpty(openOtkupi) Then
        Set BuildBlokIsplataList = result
        Exit Function
    End If
    
    Dim trCache As Object
    Dim nazivCache As Object
    Set trCache = BuildKooperantTekuciRacunCache()
    Set nazivCache = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(openOtkupi, 1)
        Dim brojDok As String
        Dim otkupID As String
        Dim ukupan As Double
        Dim isplaceno As Double
        Dim otvoren As Double
        Dim datumVal As Date
        Dim stanicaID As String
        
        brojDok = CStr(openOtkupi(i, 1))
        otkupID = CStr(openOtkupi(i, 2))
        ukupan = CDbl(openOtkupi(i, 3))
        isplaceno = CDbl(openOtkupi(i, 4))
        otvoren = CDbl(openOtkupi(i, 5))
        
        If IsDate(openOtkupi(i, 6)) Then
            datumVal = CDate(openOtkupi(i, 6))
        Else
            On Error Resume Next
            datumVal = CDate(CStr(openOtkupi(i, 6)))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0
        End If
        
        stanicaID = CStr(openOtkupi(i, 7))
        
        If stanicaIDFilter <> "" And stanicaID <> stanicaIDFilter Then GoTo NextRow
        If datumOd > #1/1/1900# And datumVal < datumOd Then GoTo NextRow
        If datumDo > #1/1/1900# And datumVal > datumDo Then GoTo NextRow
        
        Dim kooperantID As String
        kooperantID = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_KOOPERANT))
        If LenB(Trim$(kooperantID)) = 0 Then GoTo NextRow
        
        Dim kooperantNaziv As String
        If nazivCache.Exists(kooperantID) Then
            kooperantNaziv = nazivCache(kooperantID)
        Else
            kooperantNaziv = GetKooperantNaziv(kooperantID)
            nazivCache.Add kooperantID, kooperantNaziv
        End If
        
        Dim tr As String
        If trCache.Exists(kooperantID) Then
            tr = trCache(kooperantID)
        Else
            tr = ""
        End If
        
        ' --- promenjeno: class instance umesto UDT ---
        Dim blk As clsBlokIsplata
        Set blk = New clsBlokIsplata
        blk.otkupID = otkupID
        blk.Datum = datumVal
        blk.kooperantID = kooperantID
        blk.kooperantNaziv = kooperantNaziv
        blk.stanicaID = stanicaID
        blk.TekuciRacun = tr
        blk.BrojDokumenta = brojDok
        blk.UkupanIznos = ukupan
        blk.VecIsplaceno = isplaceno
        blk.OtvorenIznos = otvoren
        blk.HasTekuciRacun = (LenB(Trim$(tr)) > 0)
        
        result.Add blk
        
NextRow:
    Next i
    
    Set BuildBlokIsplataList = result
End Function
'======================================================================
' BuildKooperantTekuciRacunCache
' Single-pass cache KooperantID -> TekuciRacun
'======================================================================
Private Function BuildKooperantTekuciRacunCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        Set BuildKooperantTekuciRacunCache = dict
        Exit Function
    End If
    
    Const SRC As String = "BuildKooperantTekuciRacunCache"
    Dim colID As Long, colTR As Long
    colID = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID, SRC)
    colTR = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_TEKUCI_RACUN, SRC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim kID As String
        kID = Trim$(CStr(data(i, colID)))
        If LenB(kID) > 0 Then
            If Not dict.Exists(kID) Then
                dict.Add kID, Trim$(CStr(data(i, colTR)))
            End If
        End If
    Next i
    
    Set BuildKooperantTekuciRacunCache = dict
End Function

'======================================================================
' SummarizeBlokList - status bar string
'======================================================================
Public Function SummarizeBlokList(ByVal blokovi As Collection) As String
    If blokovi Is Nothing Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    If blokovi.count = 0 Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    
    Dim totalOpen As Double
    Dim missingTR As Long
    Dim kooperantSet As Object
    Set kooperantSet = CreateObject("Scripting.Dictionary")
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v               ' Set jer je v object reference
        totalOpen = totalOpen + blk.OtvorenIznos
        If Not blk.HasTekuciRacun Then missingTR = missingTR + 1
        If Not kooperantSet.Exists(blk.kooperantID) Then
            kooperantSet.Add blk.kooperantID, True
        End If
    Next v
    
    Dim baseMsg As String
    baseMsg = blokovi.count & " blokova | " & _
              kooperantSet.count & " kooperanata | " & _
              "Otvoreno: " & Format$(totalOpen, "#,##0.00") & " RSD"
    
    If missingTR > 0 Then
        baseMsg = baseMsg & " | Bez TR: " & missingTR
    End If
    
    SummarizeBlokList = baseMsg
End Function

'======================================================================
' ExportBlokListAsTSV - clipboard-ready TSV za Excel paste
'======================================================================
Public Function ExportBlokListAsTSV(ByVal blokovi As Collection) As String
    Dim s As String
    s = "Datum" & vbTab & "Kooperant" & vbTab & "StanicaID" & vbTab & _
        "BrojDok" & vbTab & "Ukupan" & vbTab & "Isplaceno" & vbTab & _
        "Otvoren" & vbTab & "TekuciRacun" & vbCrLf
    
    If blokovi Is Nothing Then
        ExportBlokListAsTSV = s
        Exit Function
    End If
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        s = s & Format$(blk.Datum, "yyyy-mm-dd") & vbTab & _
                blk.kooperantNaziv & vbTab & _
                blk.stanicaID & vbTab & _
                blk.BrojDokumenta & vbTab & _
                Format$(blk.UkupanIznos, "0.00") & vbTab & _
                Format$(blk.VecIsplaceno, "0.00") & vbTab & _
                Format$(blk.OtvorenIznos, "0.00") & vbTab & _
                blk.TekuciRacun & vbCrLf
    Next v
    
    ExportBlokListAsTSV = s
End Function

