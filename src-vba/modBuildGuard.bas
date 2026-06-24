Attribute VB_Name = "modBuildGuard"
Option Explicit

' ============================================================
' modBuildGuard - BUILD-ONLY provera "blanko" build-a.
'
' VAZNO: ovo se NIKADA ne poziva automatski (ni Workbook_Open, ni
' Ensure*Schema). Pokrece se RUCNO na build masini PRE
' "Save As builds\AgriX_x.x.x.xlsm". Klijenti imaju ovaj kod, ali ga nikad
' ne pozovu -> kod njih NE okida.
'
' NEDESTRUKTIVNO: samo CITA broj redova po tabeli i prijavi sta nije prazno.
' Ne brise nista. Cilj: blanko fajl koji se distribuira SVIMA ne sme da nosi
' nicije podatke (vidi R3 u docs/RELEASE_PROCEDURE.md).
'
' Upotreba (build masina, pre Save As):  Alt+F8 -> AssertBlankBuild
' ============================================================

' Vraca True ako je fajl blanko (nijedna tabela nema podatke), inace False
' uz spisak tabela sa brojem redova.
Public Function AssertBlankBuild() As Boolean
    Const SRC As String = "modBuildGuard.AssertBlankBuild"
    On Error GoTo EH

    Dim report As String
    Dim nonEmpty As Long, totalRows As Long
    Dim ws As Worksheet, lo As ListObject, rows As Long

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            rows = 0
            If Not lo.DataBodyRange Is Nothing Then rows = lo.DataBodyRange.rows.count
            If rows > 0 Then
                nonEmpty = nonEmpty + 1
                totalRows = totalRows + rows
                report = report & "  - " & lo.name & " (" & ws.name & "): " & rows & " redova" & vbCrLf
            End If
        Next lo
    Next ws

    If nonEmpty = 0 Then
        AssertBlankBuild = True
        MsgBox "BLANKO OK: nijedna tabela nema podatke." & vbCrLf & _
               "Fajl je bezbedan za distribuciju (samo kod, prazne tabele).", _
               vbInformation, APP_NAME
    Else
        AssertBlankBuild = False
        MsgBox "PAZNJA: fajl NIJE blanko - " & nonEmpty & " tabela ima podatke (" & _
               totalRows & " redova ukupno):" & vbCrLf & vbCrLf & report & vbCrLf & _
               "Ako ovo distribuiras, SVI klijenti dobijaju ove podatke." & vbCrLf & _
               "Isprazni tabele (ili gradi iz praznog build-mastera) pa ponovi proveru." & vbCrLf & vbCrLf & _
               "(Napomena: ako NAMERNO saljes sifarnike kao seed, te tabele ce biti u listi - " & _
               "to je ocekivano; transakcione tabele - Otkup/Novac/Fakture/Otpremnice/Prijemnice - " & _
               "moraju biti prazne.)", _
               vbExclamation, APP_NAME
    End If
    Exit Function

EH:
    LogErr SRC
    AssertBlankBuild = False
    MsgBox "Greska pri proveri blanko build-a: " & Err.description, vbCritical, APP_NAME
End Function
