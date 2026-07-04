Attribute VB_Name = "modIntegritet"
Option Explicit

' ============================================================
' modIntegritet - konsolidovana revizija integriteta (tabela <-> tabela).
'
' Read-only. NE menja poslovne podatke. Rezultat je sheet
' INTEGRITET_PROVERE sa punim listama neuskladjenih zapisa (filter/
' sort/print direktno u Excelu) + zbirni MsgBox.
'
' Ulaz: RunIntegritetProvere (Alt+F8; Admin dugme se dodaje kasnije).
'
' Etapa 1 (ova): A1 (otpremnica vs zbirna kg), A2 (manjak/visak
' zbirna->prijemnica), B1 (verwaist otpremnice/prijemnice),
' B2 (otkupi bez otpremnice). Reuse postojecih funkcija:
'   - ValidateZbirna            (modDokumenta)
'   - GetVerwaisteDokumente     (modDokumenta)
'   - GetUnlinkedOtkupi         (modSledljivost)
' ============================================================

Private Const INTEGRITET_SHEET As String = "INTEGRITET_PROVERE"
Private Const PRAG_MANJAK_PCT As Double = 10#   ' manjak% iznad ovoga = "za proveru"

Private m_ws As Worksheet
Private m_row As Long
Private m_summary As String
Private m_totalIssues As Long

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================

Public Sub RunIntegritetProvere()
    On Error GoTo EH

    Application.ScreenUpdating = False

    InitIntegritetSheet
    m_summary = ""
    m_totalIssues = 0

    Chk_A1_OtpremnicaVsZbirna
    Chk_A2_ManjakAnomalije
    Chk_B1_Verwaiste
    Chk_B2_UnlinkedOtkupi
    Chk_B4_DanglingBrojZbirne
    Chk_B5_PrijemnicaBezZbirne
    Chk_C1_C4_StavkaPrijemnica
    Chk_C2_StavkaBezZbirne
    Chk_C3_PaletaBezStavke
    Chk_C5_DupliBrojPalete
    Chk_A3_StavkeVsPrijemnica
    Chk_A4_PaletaHeaderVsStavke

    WriteLine "UKUPNO neuskladjenih zapisa: " & CStr(m_totalIssues), True
    FinishIntegritetSheet

    Application.ScreenUpdating = True

    MsgBox "Integritet provere zavrsene." & vbCrLf & vbCrLf & _
           m_summary & vbCrLf & _
           "UKUPNO: " & CStr(m_totalIssues) & " neuskladjenih zapisa." & vbCrLf & vbCrLf & _
           "Detalji: sheet '" & INTEGRITET_SHEET & "'.", _
           IIf(m_totalIssues > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Integritet provere - greska: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' CHECK A1: OTPREMNICA vs ZBIRNA (Sigma kg po BrojZbirne)
' ============================================================
' Reuse ValidateZbirna po svakom BrojZbirne (agregira Klasa I+II).
' Flag gde ValidKg = False (|SumaOtpremnica - ZbirnaUkupno| >= 0.01).

Private Sub Chk_A1_OtpremnicaVsZbirna()
    On Error GoTo EH

    Dim brojevi As Object: Set brojevi = CreateObject("Scripting.Dictionary")
    CollectBrojZbirne brojevi, TBL_ZBIRNA, COL_ZBR_BROJ
    CollectBrojZbirne brojevi, TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE

    Dim bad As Collection: Set bad = New Collection
    Dim kk As Variant, v As Variant
    For Each kk In brojevi.keys
        v = ValidateZbirna(CStr(kk))
        If IsArray(v) Then
            If v(3) = False Then
                bad.Add Array(CStr(kk), v(0), v(1), v(2), v(6))
            End If
        End If
    Next kk

    WriteBlock "A1", "Otpremnica vs Zbirna (Sigma kg po BrojZbirne; ocekivano 0)", _
               Array("BrojZbirne", "SumaOtpremnicaKg", "ZbirnaUkupnoKg", "RazlikaKg", "RazlikaAmb"), _
               CollToArray(bad, 5)
    Exit Sub

EH:
    WriteErr "A1", Err.description
End Sub

' ============================================================
' CHECK A2: MANJAK / VISAK zbirna -> prijemnica
' ============================================================
' Semantika kao ReportManjak: manjak = Sigma zbirna.UkupnoKolicina
' - Sigma prijemnica.Kolicina po BrojZbirne. Anomalije:
'   - VISAK   (prijemnica > zbirna; manjak < 0),
'   - NISTA PRIMLJENO (zbirna > 0, prijemnica = 0),
'   - MANJAK > prag% (moguca greska unosa / veliki kalo).

Private Sub Chk_A2_ManjakAnomalije()
    On Error GoTo EH

    Dim zbrDict As Object: Set zbrDict = AggByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KOLICINA)
    Dim prijDict As Object: Set prijDict = AggByBroj(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KOLICINA)

    Dim bad As Collection: Set bad = New Collection
    Dim kk As Variant
    Dim zbrKg As Double, prijKg As Double, manjak As Double, pct As Double
    Dim razlog As String

    For Each kk In zbrDict.keys
        zbrKg = zbrDict(kk)
        prijKg = 0
        If prijDict.Exists(kk) Then prijKg = prijDict(kk)
        manjak = zbrKg - prijKg
        If zbrKg <> 0 Then pct = manjak / zbrKg * 100 Else pct = 0

        razlog = ""
        If manjak < -0.01 Then
            razlog = "VISAK (prijemnica > zbirna)"
        ElseIf prijKg <= 0.01 And zbrKg > 0.01 Then
            razlog = "NISTA PRIMLJENO"
        ElseIf pct > PRAG_MANJAK_PCT Then
            razlog = "MANJAK > " & CStr(PRAG_MANJAK_PCT) & "%"
        End If

        If Len(razlog) > 0 Then
            bad.Add Array(CStr(kk), zbrKg, prijKg, manjak, Format$(pct, "0.0") & "%", razlog)
        End If
    Next kk

    WriteBlock "A2", "Manjak/visak zbirna->prijemnica (anomalije; prag " & CStr(PRAG_MANJAK_PCT) & "%)", _
               Array("BrojZbirne", "ZbirnaKg", "PrijemnicaKg", "ManjakKg", "Manjak%", "Razlog"), _
               CollToArray(bad, 6)
    Exit Sub

EH:
    WriteErr "A2", Err.description
End Sub

' ============================================================
' CHECK B1: VERWAIST otpremnice / prijemnice
' ============================================================
' Reuse GetVerwaisteDokumente: ziv dokument ciji je BrojZbirne
' potpuno storniran (svi redovi te zbirne stornirani).

Private Sub Chk_B1_Verwaiste()
    On Error GoTo EH

    WriteBlock "B1a", "Verwaist OTPREMNICE (ziva otpremnica, zbirna potpuno stornirana)", _
               Array("OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "VrstaVoca", "Kolicina"), _
               GetVerwaisteDokumente("Otpremnica")

    WriteBlock "B1b", "Verwaist PRIJEMNICE (ziva prijemnica, zbirna potpuno stornirana)", _
               Array("PrijemnicaID", "BrojPrijemnice", "BrojZbirne", "Kupac", "Kolicina"), _
               GetVerwaisteDokumente("Prijemnica")
    Exit Sub

EH:
    WriteErr "B1", Err.description
End Sub

' ============================================================
' CHECK B2: OTKUPI bez otpremnice (unlinked)
' ============================================================
' Reuse GetUnlinkedOtkupi: aktivan otkup bez OtpremnicaID
' (prava "razlika otkup vs otpremnica").

Private Sub Chk_B2_UnlinkedOtkupi()
    On Error GoTo EH

    WriteBlock "B2", "Otkupi bez otpremnice (unlinked)", _
               Array("OtkupID", "Datum", "StanicaID", "VozacID", "KooperantID", "Kolicina", "VrstaVoca"), _
               GetUnlinkedOtkupi()
    Exit Sub

EH:
    WriteErr "B2", Err.description
End Sub

' ============================================================
' CHECK B4: DANGLING BrojZbirne (zbirna uopste ne postoji)
' ============================================================
' Ziv dokument (otpremnica/prijemnica) sa BrojZbirne koji NE postoji ni u
' jednom redu tblZbirna - ni aktivnom ni storniranom. Razlicito od verwaist
' (B1), gde zbirna postoji ali je potpuno stornirana.

Private Sub Chk_B4_DanglingBrojZbirne()
    On Error GoTo EH

    Dim zbrSet As Object: Set zbrSet = AllBrojeviInZbirna()

    WriteBlock "B4a", "Otpremnice sa BrojZbirne koji ne postoji u tblZbirna", _
               Array("OtpremnicaID", "BrojOtpremnice", "BrojZbirne", "Kolicina"), _
               DanglingDocs(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, COL_OTP_KOLICINA, zbrSet)

    WriteBlock "B4b", "Prijemnice sa BrojZbirne koji ne postoji u tblZbirna", _
               Array("PrijemnicaID", "BrojPrijemnice", "BrojZbirne", "Kolicina"), _
               DanglingDocs(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KOLICINA, zbrSet)
    Exit Sub

EH:
    WriteErr "B4", Err.description
End Sub

' ============================================================
' CHECK B5: PRIJEMNICA bez BrojZbirne (obavezna veza)
' ============================================================

Private Sub Chk_B5_PrijemnicaBezZbirne()
    On Error GoTo EH

    Dim arr As Variant: arr = Empty
    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)

    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PRIJEMNICA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cBroj As Long, cZbr As Long, cKol As Long
            cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, "modIntegritet.Chk_B5")
            cBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, "modIntegritet.Chk_B5")
            cZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "modIntegritet.Chk_B5")
            cKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIntegritet.Chk_B5")

            Dim bad As Collection: Set bad = New Collection
            Dim i As Long
            For i = 1 To UBound(data, 1)
                If Len(Trim$(CStr(data(i, cZbr)))) = 0 Then
                    bad.Add Array(CStr(data(i, cId)), CStr(data(i, cBroj)), data(i, cKol))
                End If
            Next i
            arr = CollToArray(bad, 3)
        End If
    End If

    WriteBlock "B5", "Prijemnice bez BrojZbirne (obavezna veza)", _
               Array("PrijemnicaID", "BrojPrijemnice", "Kolicina"), arr
    Exit Sub

EH:
    WriteErr "B5", Err.description
End Sub

' ============================================================
' CHECK C1 + C4: PALETA-STAVKA -> PRIJEMNICA
' ============================================================
' C1 = aktivna stavka bez zive prijemnice (prazan ili nepostojeci PrijemnicaID).
' C4 = aktivna stavka ka prijemnici koja postoji ali je STORNIRANA
'      (kaskadni storno prijemnice ne dira tblPaletaStavka).

Private Sub Chk_C1_C4_StavkaPrijemnica()
    On Error GoTo EH

    Dim allPrij As Object: Set allPrij = IdSet(TBL_PRIJEMNICA, COL_PRJ_ID, False)
    Dim actPrij As Object: Set actPrij = IdSet(TBL_PRIJEMNICA, COL_PRJ_ID, True)

    Dim c1 As Collection: Set c1 = New Collection
    Dim c4 As Collection: Set c4 = New Collection

    Dim data As Variant: data = GetTableData(TBL_PALETA_STAVKA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA_STAVKA)
        If Not IsEmpty(data) Then
            Dim cS As Long, cPal As Long, cPrij As Long, cBrPrij As Long
            cS = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_ID, "modIntegritet.C1C4")
            cPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, "modIntegritet.C1C4")
            cPrij = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, "modIntegritet.C1C4")
            cBrPrij = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, "modIntegritet.C1C4")

            Dim i As Long, p As String
            For i = 1 To UBound(data, 1)
                p = Trim$(CStr(data(i, cPrij)))
                If Len(p) = 0 Then
                    c1.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), "(prazno)", CStr(data(i, cBrPrij)))
                ElseIf Not allPrij.Exists(p) Then
                    c1.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), p, CStr(data(i, cBrPrij)))
                ElseIf Not actPrij.Exists(p) Then
                    c4.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), p, CStr(data(i, cBrPrij)))
                End If
            Next i
        End If
    End If

    WriteBlock "C1", "Paleta-stavke bez zive prijemnice (prazan/nepostojeci PrijemnicaID)", _
               Array("StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice"), CollToArray(c1, 4)
    WriteBlock "C4", "Paleta-stavke ka storniranoj prijemnici", _
               Array("StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice"), CollToArray(c4, 4)
    Exit Sub

EH:
    WriteErr "C1/C4", Err.description
End Sub

' ============================================================
' CHECK C2: PALETA-STAVKA bez ispravne zbirne
' ============================================================

Private Sub Chk_C2_StavkaBezZbirne()
    On Error GoTo EH

    Dim zbrSet As Object: Set zbrSet = AllBrojeviInZbirna()
    Dim bad As Collection: Set bad = New Collection

    Dim data As Variant: data = GetTableData(TBL_PALETA_STAVKA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA_STAVKA)
        If Not IsEmpty(data) Then
            Dim cS As Long, cPal As Long, cZbr As Long
            cS = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_ID, "modIntegritet.C2")
            cPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, "modIntegritet.C2")
            cZbr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, "modIntegritet.C2")

            Dim i As Long, b As String, razlog As String
            For i = 1 To UBound(data, 1)
                b = Trim$(CStr(data(i, cZbr)))
                razlog = ""
                If Len(b) = 0 Then
                    razlog = "prazan BrojZbirne"
                ElseIf Not zbrSet.Exists(b) Then
                    razlog = "BrojZbirne ne postoji u tblZbirna"
                End If
                If Len(razlog) > 0 Then
                    bad.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), b, razlog)
                End If
            Next i
        End If
    End If

    WriteBlock "C2", "Paleta-stavke bez ispravne zbirne", _
               Array("StavkaID", "PaletaID", "BrojZbirne", "Razlog"), CollToArray(bad, 4)
    Exit Sub

EH:
    WriteErr "C2", Err.description
End Sub

' ============================================================
' CHECK C3: PALETA (header) bez ijedne aktivne stavke
' ============================================================

Private Sub Chk_C3_PaletaBezStavke()
    On Error GoTo EH

    Dim refPal As Object: Set refPal = AggByBroj(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, COL_PALS_NETO)
    Dim bad As Collection: Set bad = New Collection

    Dim data As Variant: data = GetTableData(TBL_PALETA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cBroj As Long, cGod As Long, cNeto As Long
            cId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, "modIntegritet.C3")
            cBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, "modIntegritet.C3")
            cGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, "modIntegritet.C3")
            cNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, "modIntegritet.C3")

            Dim i As Long, pid As String
            For i = 1 To UBound(data, 1)
                pid = Trim$(CStr(data(i, cId)))
                If Len(pid) > 0 Then
                    If Not refPal.Exists(pid) Then
                        bad.Add Array(pid, CStr(data(i, cBroj)), CStr(data(i, cGod)), data(i, cNeto))
                    End If
                End If
            Next i
        End If
    End If

    WriteBlock "C3", "Palete (header) bez ijedne aktivne stavke", _
               Array("PaletaID", "BrojPalete", "Godina", "NetoKg"), CollToArray(bad, 4)
    Exit Sub

EH:
    WriteErr "C3", Err.description
End Sub

' ============================================================
' CHECK C5: dupli BrojPalete unutar iste Godine
' ============================================================

Private Sub Chk_C5_DupliBrojPalete()
    On Error GoTo EH

    Dim bad As Collection: Set bad = New Collection
    Dim data As Variant: data = GetTableData(TBL_PALETA)

    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cBroj As Long, cGod As Long
            cId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, "modIntegritet.C5")
            cBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, "modIntegritet.C5")
            cGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, "modIntegritet.C5")

            Dim cnt As Object: Set cnt = CreateObject("Scripting.Dictionary")
            Dim i As Long, key As String, br As String
            For i = 1 To UBound(data, 1)
                br = Trim$(CStr(data(i, cBroj)))
                If Len(br) > 0 Then
                    key = Trim$(CStr(data(i, cGod))) & "|" & br
                    If cnt.Exists(key) Then cnt(key) = cnt(key) + 1 Else cnt.Add key, 1
                End If
            Next i

            For i = 1 To UBound(data, 1)
                br = Trim$(CStr(data(i, cBroj)))
                If Len(br) > 0 Then
                    key = Trim$(CStr(data(i, cGod))) & "|" & br
                    If cnt(key) > 1 Then
                        bad.Add Array(CStr(data(i, cId)), br, CStr(data(i, cGod)), CStr(cnt(key)))
                    End If
                End If
            Next i
        End If
    End If

    WriteBlock "C5", "Dupli BrojPalete unutar iste Godine", _
               Array("PaletaID", "BrojPalete", "Godina", "BrojDuplikata"), CollToArray(bad, 4)
    Exit Sub

EH:
    WriteErr "C5", Err.description
End Sub

' ============================================================
' CHECK A3: Sigma paleta-stavke NetoKg vs prijemnica.Kolicina
' ============================================================
' Samo paletizovane prijemnice (koje imaju bar jednu aktivnu stavku).

Private Sub Chk_A3_StavkeVsPrijemnica()
    On Error GoTo EH

    Dim stByPrij As Object: Set stByPrij = AggByBroj(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, COL_PALS_NETO)
    Dim prijKol As Object: Set prijKol = AggByBroj(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_KOLICINA)

    Dim bad As Collection: Set bad = New Collection
    Dim kk As Variant, sumSt As Double, kol As Double, diff As Double
    For Each kk In stByPrij.keys
        If prijKol.Exists(kk) Then
            sumSt = stByPrij(kk)
            kol = prijKol(kk)
            diff = sumSt - kol
            If Abs(diff) > 0.5 Then
                bad.Add Array(CStr(kk), sumSt, kol, diff)
            End If
        End If
    Next kk

    WriteBlock "A3", "Paleta-stavke (Sigma NetoKg) vs prijemnica.Kolicina", _
               Array("PrijemnicaID", "SumaStavkeKg", "PrijemnicaKg", "RazlikaKg"), CollToArray(bad, 4)
    Exit Sub

EH:
    WriteErr "A3", Err.description
End Sub

' ============================================================
' CHECK A4: paleta header (NetoKg, BrojGajbica) vs Sigma stavke
' ============================================================

Private Sub Chk_A4_PaletaHeaderVsStavke()
    On Error GoTo EH

    Dim netoBy As Object: Set netoBy = AggByBroj(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, COL_PALS_NETO)
    Dim gajbeBy As Object: Set gajbeBy = AggByBroj(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, COL_PALS_BR_GAJBICA)

    Dim bad As Collection: Set bad = New Collection
    Dim data As Variant: data = GetTableData(TBL_PALETA)

    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cNeto As Long, cGajb As Long
            cId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, "modIntegritet.A4")
            cNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, "modIntegritet.A4")
            cGajb = RequireColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA, "modIntegritet.A4")

            Dim i As Long, pid As String
            Dim hNeto As Double, hGajb As Double, sNeto As Double, sGajb As Double
            For i = 1 To UBound(data, 1)
                pid = Trim$(CStr(data(i, cId)))
                If Len(pid) > 0 Then
                    If netoBy.Exists(pid) Then
                        hNeto = 0: If IsNumeric(data(i, cNeto)) Then hNeto = CDbl(data(i, cNeto))
                        hGajb = 0: If IsNumeric(data(i, cGajb)) Then hGajb = CDbl(data(i, cGajb))
                        sNeto = netoBy(pid)
                        sGajb = 0: If gajbeBy.Exists(pid) Then sGajb = gajbeBy(pid)
                        If Abs(hNeto - sNeto) > 0.5 Or Abs(hGajb - sGajb) > 0.001 Then
                            bad.Add Array(pid, hNeto, sNeto, hGajb, sGajb)
                        End If
                    End If
                End If
            Next i
        End If
    End If

    WriteBlock "A4", "Paleta header vs Sigma stavke (NetoKg, BrojGajbica)", _
               Array("PaletaID", "HeaderNetoKg", "StavkeNetoKg", "HeaderGajbica", "StavkeGajbica"), CollToArray(bad, 5)
    Exit Sub

EH:
    WriteErr "A4", Err.description
End Sub

' ============================================================
' SHARED HELPERS
' ============================================================

' Skupi jedinstvene ne-prazne BrojZbirne (bez storniranih) iz tabele.
Private Sub CollectBrojZbirne(ByRef dict As Object, ByVal tbl As String, ByVal col As String)
    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then Exit Sub
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then Exit Sub

    Dim ci As Long: ci = RequireColumnIndex(tbl, col, "modIntegritet.CollectBrojZbirne")

    Dim i As Long, b As String
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, ci)))
        If Len(b) > 0 Then
            If Not dict.Exists(b) Then dict.Add b, True
        End If
    Next i
End Sub

' Agregat kg po BrojZbirne (bez storniranih). Vraca Dictionary(broj -> Double).
Private Function AggByBroj(ByVal tbl As String, ByVal brojCol As String, ByVal kgCol As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then Set AggByBroj = d: Exit Function
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then Set AggByBroj = d: Exit Function

    Dim cb As Long, ck As Long
    cb = RequireColumnIndex(tbl, brojCol, "modIntegritet.AggByBroj")
    ck = RequireColumnIndex(tbl, kgCol, "modIntegritet.AggByBroj")

    Dim i As Long, b As String, kg As Double
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, cb)))
        If Len(b) > 0 Then
            kg = 0
            If IsNumeric(data(i, ck)) Then kg = CDbl(data(i, ck))
            If d.Exists(b) Then d(b) = d(b) + kg Else d.Add b, kg
        End If
    Next i

    Set AggByBroj = d
End Function

' Collection ciji su elementi 0-bazni Array(...) -> 2D (1..n, 1..cols). Empty ako prazno.
Private Function CollToArray(ByVal c As Collection, ByVal cols As Long) As Variant
    If c.count = 0 Then CollToArray = Empty: Exit Function

    Dim r() As Variant: ReDim r(1 To c.count, 1 To cols)
    Dim i As Long, k As Long, item As Variant
    For i = 1 To c.count
        item = c(i)
        For k = 1 To cols
            r(i, k) = item(k - 1)
        Next k
    Next i

    CollToArray = r
End Function

' Skup SVIH BrojZbirne u tblZbirna (i stornirani se racunaju kao "postoji").
Private Function AllBrojeviInZbirna() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    If Not IsArray(data) Then Set AllBrojeviInZbirna = d: Exit Function

    Dim cb As Long: cb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modIntegritet.AllBrojeviInZbirna")

    Dim i As Long, b As String
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, cb)))
        If Len(b) > 0 Then
            If Not d.Exists(b) Then d.Add b, True
        End If
    Next i

    Set AllBrojeviInZbirna = d
End Function

' Zivi dokumenti sa BrojZbirne koji nije u zbrSet. Vraca 2D(1..n,1..4) ili Empty.
Private Function DanglingDocs(ByVal tbl As String, ByVal idCol As String, _
                              ByVal brojCol As String, ByVal zbrCol As String, _
                              ByVal kolCol As String, ByVal zbrSet As Object) As Variant
    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then DanglingDocs = Empty: Exit Function
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then DanglingDocs = Empty: Exit Function

    Dim cId As Long, cBroj As Long, cZbr As Long, cKol As Long
    cId = RequireColumnIndex(tbl, idCol, "modIntegritet.DanglingDocs")
    cBroj = RequireColumnIndex(tbl, brojCol, "modIntegritet.DanglingDocs")
    cZbr = RequireColumnIndex(tbl, zbrCol, "modIntegritet.DanglingDocs")
    cKol = RequireColumnIndex(tbl, kolCol, "modIntegritet.DanglingDocs")

    Dim bad As Collection: Set bad = New Collection
    Dim i As Long, b As String
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, cZbr)))
        If Len(b) > 0 Then
            If Not zbrSet.Exists(b) Then
                bad.Add Array(CStr(data(i, cId)), CStr(data(i, cBroj)), b, data(i, cKol))
            End If
        End If
    Next i

    DanglingDocs = CollToArray(bad, 4)
End Function

' Skup ID-jeva iz tabele (opciono samo aktivni). Vraca Dictionary(id -> True).
Private Function IdSet(ByVal tbl As String, ByVal idCol As String, ByVal activeOnly As Boolean) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then Set IdSet = d: Exit Function
    If activeOnly Then data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then Set IdSet = d: Exit Function

    Dim ci As Long: ci = RequireColumnIndex(tbl, idCol, "modIntegritet.IdSet")

    Dim i As Long, s As String
    For i = 1 To UBound(data, 1)
        s = Trim$(CStr(data(i, ci)))
        If Len(s) > 0 Then
            If Not d.Exists(s) Then d.Add s, True
        End If
    Next i

    Set IdSet = d
End Function

' ============================================================
' SHEET WRITER
' ============================================================

Private Sub InitIntegritetSheet()
    On Error Resume Next
    Set m_ws = ThisWorkbook.Worksheets(INTEGRITET_SHEET)
    If m_ws Is Nothing Then
        Set m_ws = ThisWorkbook.Worksheets.Add( _
                       after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        m_ws.name = INTEGRITET_SHEET
    End If
    m_ws.cells.Clear
    On Error GoTo 0

    m_row = 1
    WriteLine "INTEGRITET PROVERE  --  " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
              "  (" & Environ$("Username") & ")", True
    m_row = m_row + 1
End Sub

Private Sub FinishIntegritetSheet()
    On Error Resume Next
    m_ws.columns("A:H").AutoFit
    m_ws.Activate
    On Error GoTo 0
End Sub

' Jedna linija u kolonu A (opciono bold).
Private Sub WriteLine(ByVal text As String, ByVal boldRow As Boolean)
    On Error Resume Next
    m_ws.cells(m_row, 1).value = text
    m_ws.rows(m_row).Font.Bold = boldRow
    On Error GoTo 0
    m_row = m_row + 1
End Sub

' Blok: naslov + (header + data) ili "OK - nema". Azurira summary + total.
Private Sub WriteBlock(ByVal code As String, ByVal title As String, _
                       ByVal headers As Variant, ByVal dataArr As Variant)
    Dim n As Long
    n = 0
    If IsArray(dataArr) Then
        If Not IsEmpty(dataArr) Then n = UBound(dataArr, 1)
    End If

    WriteLine "[" & code & "] " & title & "  --  " & CStr(n) & " zapis(a)", True

    If n = 0 Then
        WriteLine "    OK - nema neuskladjenih zapisa.", False
    Else
        Dim c As Long
        For c = LBound(headers) To UBound(headers)
            m_ws.cells(m_row, c - LBound(headers) + 1).value = headers(c)
        Next c
        m_ws.rows(m_row).Font.Italic = True
        m_row = m_row + 1

        Dim i As Long, k As Long
        For i = 1 To n
            For k = 1 To UBound(dataArr, 2)
                m_ws.cells(m_row, k).value = dataArr(i, k)
            Next k
            m_row = m_row + 1
        Next i
    End If

    m_row = m_row + 1   ' prazan separator

    m_summary = m_summary & "  " & code & ": " & CStr(n) & vbCrLf
    m_totalIssues = m_totalIssues + n
End Sub

Private Sub WriteErr(ByVal code As String, ByVal desc As String)
    WriteLine "[" & code & "] GRESKA: " & desc, True
    m_summary = m_summary & "  " & code & ": GRESKA" & vbCrLf
    m_row = m_row + 1
End Sub
