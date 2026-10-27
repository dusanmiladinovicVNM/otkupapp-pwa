Attribute VB_Name = "modQA_AmbKarticaSaldos"
Option Explicit

' READ-ONLY temporary acceptance check for the *installed VBA report*.
' No table writes, no workbook Save, no file IO and no dependency on migration.
' Run QA_AmbKartice_SviKooperanti from a copy of the QA workbook.

Private Const QA_EPS As Double = 0.000001

Public Sub QA_AmbKartice_SviKooperanti()
    Const SRC As String = "modQA_AmbKarticaSaldos.QA_AmbKartice_SviKooperanti"
    On Error GoTo Fatal

    Dim kooperanti As Variant
    kooperanti = GetTableData(TBL_KOOPERANTI)
    If Not IsArray(kooperanti) Then Err.Raise vbObjectError + 4710, SRC, "tblKooperanti nije ucitana"

    Dim cID As Long
    cID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    If cID < 1 Then Err.Raise vbObjectError + 4711, SRC, "Nedostaje kolona KooperantID"

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim i As Long, id As String, why As String
    Dim qty As Long, subjects As Long, withMoves As Long, examined As Long
    Dim failed As Long, failuresShown As Long

    Debug.Print "QA_AMB_START -- READ ONLY"

    For i = 1 To UBound(kooperanti, 1)
        id = Trim$(CStr(kooperanti(i, cID)))
        If Len(id) > 0 Then
            If Not seen.Exists(id) Then
                seen.Add id, True
                subjects = subjects + 1
                why = vbNullString
                qty = 0
                If Not QA_CheckOne(id, qty, why) Then
                    failed = failed + 1
                    If failuresShown < 25 Then
                        Debug.Print "QA_AMB_FAIL " & id & " : " & why
                        failuresShown = failuresShown + 1
                    End If
                Else
                    If qty > 0 Then withMoves = withMoves + 1
                    examined = examined + qty
                End If
            End If
        End If
    Next i

    Debug.Print "QA_AMB_RESULT=" & IIf(failed = 0, "PASS", "FAIL") & _
                " subjects=" & subjects & _
                " with_moves=" & withMoves & _
                " report_rows_checked=" & examined & _
                " failed_subjects=" & failed
    If failed > failuresShown Then Debug.Print "Additional failures suppressed=" & (failed - failuresShown)
    Exit Sub

Fatal:
    Debug.Print "QA_AMB_FATAL " & Err.Number & " : " & Err.Description
End Sub

Private Function QA_CheckOne(ByVal kooperantID As String, _
                             ByRef rowsChecked As Long, _
                             ByRef why As String) As Boolean
    On Error GoTo EH

    Dim flows As Variant, card As Variant
    flows = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, kooperantID)

    Dim sourceRows As Long, n As Long, signedTotal As Double
    sourceRows = 0: signedTotal = 0#
    If IsArray(flows) Then
        For n = 1 To UBound(flows, 1)
            ' A reader silently skipping one malformed row must FAIL the check.
            If Len(Trim$(CStr(flows(n, 2)))) = 0 Then
                Err.Raise vbObjectError + 4712, , "Ledger row " & n & ": empty DokumentID"
            End If
            If Not IsDate(flows(n, 1)) Then
                Err.Raise vbObjectError + 4713, , "Ledger row " & n & ": invalid Datum"
            End If
            If Not IsNumeric(flows(n, 5)) Then
                Err.Raise vbObjectError + 4714, , "Ledger row " & n & ": invalid signed quantity"
            End If
            sourceRows = sourceRows + 1
            signedTotal = signedTotal + CDbl(flows(n, 5))
        Next n
    End If

    card = modIzvestaj.ReportKarticaAmbalaze(kooperantID, _
                                             DateSerial(1900, 1, 1), _
                                             DateSerial(9999, 12, 31))
    If sourceRows = 0 Then
        If IsArray(card) Then Err.Raise vbObjectError + 4715, , "Report rows with empty ledger"
        QA_CheckOne = True
        Exit Function
    End If
    If Not IsArray(card) Then Err.Raise vbObjectError + 4716, , "Report is empty despite ledger rows"

    Dim rows As Long
    rows = UBound(card, 1)
    If rows <> sourceRows + 1 Then
        Err.Raise vbObjectError + 4717, , "Report row count " & rows & " != ledger " & sourceRows & " + UKUPNO"
    End If

    Dim running As Double, sumU As Double, sumI As Double
    Dim inbound As Double, outbound As Double, reported As Double
    running = 0#: sumU = 0#: sumI = 0#

    For n = 1 To sourceRows
        If Not IsDate(card(n, 1)) Then Err.Raise vbObjectError + 4718, , "Report row " & n & ": no date"
        If Len(Trim$(CStr(card(n, 2)))) = 0 Then Err.Raise vbObjectError + 4719, , "Report row " & n & ": no document"
        inbound = QA_Number(card(n, 4), n, "Ulaz")
        outbound = QA_Number(card(n, 5), n, "Izlaz")
        reported = QA_Number(card(n, 6), n, "Saldo")
        If inbound < 0 Or outbound < 0 Then Err.Raise vbObjectError + 4720, , "Report row " & n & ": negative Ulaz/Izlaz"
        If inbound > QA_EPS And outbound > QA_EPS Then Err.Raise vbObjectError + 4721, , "Report row " & n & ": both Ulaz and Izlaz"
        running = running + inbound - outbound
        sumU = sumU + inbound
        sumI = sumI + outbound
        If Abs(running - reported) > QA_EPS Then
            Err.Raise vbObjectError + 4722, , "Report row " & n & ": expected saldo=" & CStr(running) & ", actual=" & CStr(reported)
        End If
    Next n

    If Trim$(CStr(card(rows, 3))) <> "UKUPNO" Then Err.Raise vbObjectError + 4723, , "Missing UKUPNO footer"
    If Abs(QA_Number(card(rows, 4), rows, "TotalUlaz") - sumU) > QA_EPS Then Err.Raise vbObjectError + 4724, , "UKUPNO Ulaz differs"
    If Abs(QA_Number(card(rows, 5), rows, "TotalIzlaz") - sumI) > QA_EPS Then Err.Raise vbObjectError + 4725, , "UKUPNO Izlaz differs"
    If Abs(QA_Number(card(rows, 6), rows, "TotalSaldo") - running) > QA_EPS Then Err.Raise vbObjectError + 4726, , "UKUPNO Saldo differs"
    If Abs(signedTotal - running) > QA_EPS Then Err.Raise vbObjectError + 4727, , "Ledger signed sum != report final saldo"

    rowsChecked = sourceRows
    QA_CheckOne = True
    Exit Function

EH:
    why = "Error " & Err.Number & " : " & Err.Description
    QA_CheckOne = False
End Function

Private Function QA_Number(ByVal v As Variant, ByVal rowNum As Long, ByVal fieldName As String) As Double
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 4728, , "Report row " & rowNum & " invalid " & fieldName
    End If
    QA_Number = CDbl(v)
End Function
