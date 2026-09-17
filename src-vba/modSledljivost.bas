Attribute VB_Name = "modSledljivost"
Option Explicit

' ============================================================
' modSledljivost v1.1 - Automatische Zuordnung Otkup ? Otpremnica
'
' Hardening only:
'   - same matching logic
'   - same return shapes
'   - no PWA/trace flow change
'   - column guards
'   - checked updates
'   - optional TX wrapper for batch auto-link
' ============================================================

Private Function BuildAutoLinkKey(ByVal stanicaID As Variant, _
                                  ByVal datumValue As Variant, _
                                  ByVal vozacID As Variant, _
                                  ByVal klasa As Variant, _
                                  ByVal brojZbirne As String, _
                                  ByVal includeZbirna As Boolean) As String
    If Not IsDate(datumValue) Then Exit Function

    BuildAutoLinkKey = _
        UCase$(Trim$(CStr(stanicaID))) & "|" & _
        Format$(CDate(datumValue), "yyyy-mm-dd") & "|" & _
        UCase$(Trim$(CStr(vozacID))) & "|" & _
        UCase$(Trim$(CStr(klasa)))

    If includeZbirna Then
        BuildAutoLinkKey = BuildAutoLinkKey & "|" & _
                           UCase$(Trim$(brojZbirne))
    End If
End Function

Private Sub AddAutoLinkCandidate(ByVal index As Object, _
                                 ByVal key As String, _
                                 ByVal otpID As String)
    If Len(Trim$(key)) = 0 Then Exit Sub
    If Len(Trim$(otpID)) = 0 Then Exit Sub

    Dim bucket As Collection

    If Not index.Exists(key) Then
        Set bucket = New Collection
        index.Add key, bucket
    End If

    index(key).Add otpID
End Sub

Private Function GetUniqueAutoLinkTarget(ByVal index As Object, _
                                         ByVal key As String) As String
    If Len(Trim$(key)) = 0 Then Exit Function

    If Not index.Exists(key) Then Exit Function

    If index(key).count = 1 Then
        GetUniqueAutoLinkTarget = CStr(index(key)(1))
    End If
End Function

' Kandidati otpremnica za RUCNO povezivanje nepovezanog otkupa -- ISTO
' pravilo koje koristi frmSledljivost.lstNepovezani_Click od prvog dana:
' nestornirane otpremnice sa ISTE stanice i ISTOG datuma kao otkup.
' Izdvojeno za ekran Sledljivost (v6-ui-187, smoke krug 2): forma zadrzava
' svoju kopiju i ne menja se (CLAUDE.md par. 5 / Faza B). Samo citanje;
' upis ide iskljucivo kroz ReassignOtkupToOtpremnica_TX (kapije cilja).
'
' Returns: 2D Array (1..N, 1..5) ili Empty:
'   1 OtpremnicaID  2 BrojOtpremnice  3 BrojZbirne  4 Kolicina  5 Klasa
Public Function GetOtpremnicaKandidatiZaOtkup(ByVal otkupID As String) As Variant
    On Error GoTo EH

    Const SRC As String = "modSledljivost.GetOtpremnicaKandidatiZaOtkup"

    If Len(Trim$(otkupID)) = 0 Then Exit Function

    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkupData) Then Exit Function
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    If IsEmpty(otkupData) Then Exit Function

    Dim cOtkId As Long, cOtkSt As Long, cOtkDat As Long
    cOtkId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cOtkSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)

    Dim stanicaID As String, datum As Date, nasao As Boolean
    Dim i As Long
    For i = 1 To UBound(otkupData, 1)
        If Trim$(CStr(otkupData(i, cOtkId))) = Trim$(otkupID) Then
            stanicaID = Trim$(CStr(otkupData(i, cOtkSt)))
            If IsDate(otkupData(i, cOtkDat)) Then
                datum = CDate(otkupData(i, cOtkDat))
                nasao = True
            End If
            Exit For
        End If
    Next i
    If Not nasao Then Exit Function

    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Function
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Function

    Dim cId As Long, cSt As Long, cDat As Long
    Dim cBr As Long, cZbr As Long, cKol As Long, cKl As Long
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, SRC)
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cBr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, SRC)
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cKl = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)

    Dim count As Long
    For i = 1 To UBound(otpData, 1)
        If Trim$(CStr(otpData(i, cSt))) = stanicaID And IsDate(otpData(i, cDat)) Then
            If CDate(otpData(i, cDat)) = datum Then count = count + 1
        End If
    Next i
    If count = 0 Then Exit Function

    Dim result() As Variant, idx As Long
    ReDim result(1 To count, 1 To 5)
    For i = 1 To UBound(otpData, 1)
        If Trim$(CStr(otpData(i, cSt))) = stanicaID And IsDate(otpData(i, cDat)) Then
            If CDate(otpData(i, cDat)) = datum Then
                idx = idx + 1
                result(idx, 1) = CStr(otpData(i, cId))
                result(idx, 2) = CStr(otpData(i, cBr))
                result(idx, 3) = CStr(otpData(i, cZbr))
                result(idx, 4) = otpData(i, cKol)
                result(idx, 5) = CStr(otpData(i, cKl))
            End If
        End If
    Next i

    GetOtpremnicaKandidatiZaOtkup = result
    Exit Function

EH:
    LogErr "modSledljivost.GetOtpremnicaKandidatiZaOtkup"
    GetOtpremnicaKandidatiZaOtkup = Empty
End Function

