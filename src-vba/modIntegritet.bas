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
Private Const PRAG_VISAK_PCT As Double = 5#     ' visak do ovoga se NE prijavljuje

Private m_ws As Worksheet
Private m_row As Long
Private m_summary As String
Private m_totalIssues As Long
Private m_rows As Collection      ' ravan spisak nalaza: Array(code, detalj) -- za ListBox
Private m_toSheet As Boolean      ' True = pisi i u INTEGRITET sheet; False = samo memorija

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================

Public Sub RunIntegritetProvere(Optional ByVal showSummary As Boolean = True)
    On Error GoTo EH

    Application.ScreenUpdating = False

    m_toSheet = True
    InitIntegritetSheet
    Set m_rows = New Collection
    m_summary = ""
    m_totalIssues = 0

    RunAllChecks

    WriteLine "UKUPNO neuskladjenih zapisa: " & CStr(m_totalIssues), True
    FinishIntegritetSheet

    Application.ScreenUpdating = True

    If showSummary Then
        MsgBox "Integritet provere zavrsene." & vbCrLf & vbCrLf & _
               m_summary & vbCrLf & _
               "UKUPNO: " & CStr(m_totalIssues) & " neuskladjenih zapisa." & vbCrLf & vbCrLf & _
               "Detalji: sheet '" & INTEGRITET_SHEET & "'.", _
               IIf(m_totalIssues > 0, vbExclamation, vbInformation), APP_NAME
    End If
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Integritet provere - greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Ravan spisak nalaza (samo problemi) kao 2D(1..n,1..2): Provera | Detalj.
' Ne dira sheet. Za in-app ListBox prikaz (frmOtkupAPP overlay).
Public Function GetIntegritetRows() As Variant
    On Error GoTo EH

    m_toSheet = False
    Set m_rows = New Collection
    m_summary = ""
    m_totalIssues = 0

    RunAllChecks

    GetIntegritetRows = CollToArray(m_rows, 2)
    Exit Function

EH:
    LogErr "modIntegritet.GetIntegritetRows"
    GetIntegritetRows = Empty
End Function

' Broj problema iz poslednjeg prolaza.
Public Function IntegritetUkupno() As Long
    IntegritetUkupno = m_totalIssues
End Function

Private Sub RunAllChecks()
    Chk_A1_OtpremnicaVsZbirna
    Chk_A2_ManjakAnomalije
    Chk_B1_Verwaiste
    Chk_B2_UnlinkedOtkupi
    Chk_B3_IzgubljeniBlokovi
    Chk_B4_DanglingBrojZbirne
    Chk_B5_PrijemnicaBezZbirne
    Chk_B5b_OtpremnicaBezZbirne
    Chk_B6_ZbirnaCaseMismatch
    Chk_B7_ZbirnaNulaKg
    Chk_C1_C4_StavkaPrijemnica
    Chk_C2_StavkaBezZbirne
    Chk_C3_PaletaBezStavke
    Chk_C5_DupliBrojPalete
    Chk_A3_StavkeVsPrijemnica
    Chk_A4_PaletaHeaderVsStavke
    Chk_D1_PreradjenoBezPrerade
    Chk_D2_PreradaNesvezaPaleta
    Chk_A5_PreradaKg
    Chk_P4_PreradaLagerJedinica
    Chk_P5_StavkeLagerJedinica
    Chk_P6_LagerJedinicaOsnov
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
        If manjak < 0 And Abs(manjak) > zbrKg * (PRAG_VISAK_PCT / 100#) Then
            razlog = "VISAK > " & CStr(PRAG_VISAK_PCT) & "% (prijemnica > zbirna)"
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
' CHECK B3: izgubljeni otkup blokovi
' ============================================================
' Aktivan otkup ciji OtpremnicaID pokazuje na storniranu/nepostojecu
' otpremnicu (reuse GetLostOtkupBlokovi; frmOtkupAPP baner ovo prikazuje).

Private Sub Chk_B3_IzgubljeniBlokovi()
    On Error GoTo EH
    WriteBlock "B3", "Izgubljeni otkup blokovi (OtpremnicaID -> stornirana/nepostojeca otpremnica)", _
               Array("OtkupID", "BrojDokumenta", "KooperantID", "Datum", "Kolicina", "OtpremnicaID", "StaraOtpremnica"), _
               GetLostOtkupBlokovi()
    Exit Sub

EH:
    WriteErr "B3", Err.description
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
    WriteBlock "B5", "Prijemnice bez BrojZbirne (obavezna veza)", _
               Array("PrijemnicaID", "BrojPrijemnice", "Kolicina"), _
               DocsBezZbirne(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KOLICINA)
    Exit Sub

EH:
    WriteErr "B5", Err.description
End Sub

' ============================================================
' CHECK B5b: OTPREMNICA bez BrojZbirne
' ============================================================
' (frmOtkupAPP startup baner ovo prikazuje -- audit je sada nadskup.)

Private Sub Chk_B5b_OtpremnicaBezZbirne()
    On Error GoTo EH
    WriteBlock "B5b", "Otpremnice bez BrojZbirne (nije vezana za zbirnu)", _
               Array("OtpremnicaID", "BrojOtpremnice", "Kolicina"), _
               DocsBezZbirne(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, COL_OTP_KOLICINA)
    Exit Sub

EH:
    WriteErr "B5b", Err.description
End Sub

' ============================================================
' CHECK B7: ZBIRNA sa 0 (ili prazan) UkupnoKolicina
' ============================================================
' Zbirna bez kolicine je sama po sebi anomalija. Komplementarno sa A2:
' A2 hvata zbirne-sa-kg-bez-prijema, B7 hvata prazne zbirne.

Private Sub Chk_B7_ZbirnaNulaKg()
    On Error GoTo EH

    Dim zbrDict As Object: Set zbrDict = AggByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KOLICINA)
    Dim bad As Collection: Set bad = New Collection

    Dim kk As Variant
    For Each kk In zbrDict.keys
        If zbrDict(kk) <= 0.005 Then
            bad.Add Array(CStr(kk), zbrDict(kk))
        End If
    Next kk

    WriteBlock "B7", "Zbirne sa 0 (ili prazan) UkupnoKolicina", _
               Array("BrojZbirne", "ZbirnaUkupnoKg"), CollToArray(bad, 2)
    Exit Sub

EH:
    WriteErr "B7", Err.description
End Sub

' ============================================================
' CHECK B6: BrojZbirne case-mismatch (za normalizaciju)
' ============================================================
' Zapisi ciji se BrojZbirne poklapa sa zbirnom SAMO ignorisuci veliko/malo
' slovo (npr. dok "s5/..." vs zbirna "S5/..."). Nije "ne postoji" (to je B4/
' C2), ali je latentan problem: case-senzitivni delovi app-a (ReportManjak,
' GetVerwaisteDokumente) ih tretiraju kao razlicite -> tiha greska. Advisory:
' listaj sve za normalizaciju na tacan zapis iz tblZbirna.

Private Sub Chk_B6_ZbirnaCaseMismatch()
    On Error GoTo EH

    Dim exactSet As Object, ciMap As Object
    BuildZbirnaCaseMaps exactSet, ciMap

    Dim bad As Collection: Set bad = New Collection
    CollectCaseMismatch bad, "Otpremnica", TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ_ZBIRNE, exactSet, ciMap
    CollectCaseMismatch bad, "Prijemnica", TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ_ZBIRNE, exactSet, ciMap
    CollectCaseMismatch bad, "PaletaStavka", TBL_PALETA_STAVKA, COL_PALS_ID, COL_PALS_BROJ_ZBIRNE, exactSet, ciMap
    CollectCaseMismatch bad, "Otkup", TBL_OTKUP, COL_OTK_ID, COL_OTK_BROJ_ZBIRNE, exactSet, ciMap

    WriteBlock "B6", "BrojZbirne se poklapa samo do velikog/malog slova (za normalizaciju)", _
               Array("Tabela", "DokumentID", "BrojZbirne (napisano)", "Zbirna (kanon)"), CollToArray(bad, 4)
    Exit Sub

EH:
    WriteErr "B6", Err.description
End Sub

' ============================================================
' CHECK C1 + C4: PALETA-STAVKA -> PRIJEMNICA
' ============================================================
' C1 = aktivna stavka bez zive prijemnice (prazan ili nepostojeci PrijemnicaID).
' C4 = aktivna stavka ka prijemnici koja postoji ali je STORNIRANA
'      (kaskadni storno prijemnice ne dira tblPaletaStavka).
'      + dokumentni tok te prijemnice: BrojZbirne (i sa stornirane prijemnice),
'      Sigma kg/amb aktivne zbirne, naziv otkupnog mesta (preko otpremnica
'      zbirne; fallback mirror-stanica zbirna preko tblZbirna.VozacID) i
'      proizvodjac(i) = kooperanti otkupa te zbirne. Vise OM / proizvodjaca
'      po zbirnoj se odvaja sa "; ".

Private Sub Chk_C1_C4_StavkaPrijemnica()
    On Error GoTo EH

    Dim allPrij As Object: Set allPrij = IdSet(TBL_PRIJEMNICA, COL_PRJ_ID, False)
    Dim actPrij As Object: Set actPrij = IdSet(TBL_PRIJEMNICA, COL_PRJ_ID, True)

    ' C4 dokumentni tok: prijemnica -> zbirna (Sigma kg/amb) -> OM + proizvodjaci.
    Dim prijZbr As Object: Set prijZbr = PrijemnicaZbirnaMap()
    Dim zbrKg As Object: Set zbrKg = AggByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KOLICINA)
    Dim zbrAmb As Object: Set zbrAmb = AggByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KOL_AMB)
    Dim omByZbr As Object: Set omByZbr = OtkupnoMestoByZbirna()
    Dim przByZbr As Object: Set przByZbr = ProizvodjacByZbirna()

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
            Dim bz As String, kgV As Variant, ambV As Variant, om As String, prz As String
            For i = 1 To UBound(data, 1)
                p = Trim$(CStr(data(i, cPrij)))
                If Len(p) = 0 Then
                    c1.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), "(prazno)", CStr(data(i, cBrPrij)))
                ElseIf Not allPrij.Exists(p) Then
                    c1.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), p, CStr(data(i, cBrPrij)))
                ElseIf Not actPrij.Exists(p) Then
                    bz = "": If prijZbr.Exists(p) Then bz = CStr(prijZbr(p))
                    kgV = "": ambV = "": om = "": prz = ""
                    If Len(bz) > 0 Then
                        If zbrKg.Exists(bz) Then kgV = zbrKg(bz)
                        If zbrAmb.Exists(bz) Then ambV = zbrAmb(bz)
                        If omByZbr.Exists(bz) Then om = CStr(omByZbr(bz))
                        If przByZbr.Exists(bz) Then prz = CStr(przByZbr(bz))
                    End If
                    c4.Add Array(CStr(data(i, cS)), CStr(data(i, cPal)), p, CStr(data(i, cBrPrij)), _
                                 bz, kgV, ambV, om, prz)
                End If
            Next i
        End If
    End If

    WriteBlock "C1", "Paleta-stavke bez zive prijemnice (prazan/nepostojeci PrijemnicaID)", _
               Array("StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice"), CollToArray(c1, 4)
    WriteBlock "C4", "Paleta-stavke ka storniranoj prijemnici", _
               Array("StavkaID", "PaletaID", "PrijemnicaID", "BrojPrijemnice", _
                     "BrojZbirne", "ZbirnaKg", "ZbirnaAmb", "OtkupnoMesto", "Proizvodjac"), CollToArray(c4, 9)
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
' CHECK D1: PALETA Preradjeno=Da bez aktivne prerada-stavke
' ============================================================
' Paleta markirana kao preradjena, ali nijedna aktivna tblPreradaStavka je ne
' referencira (prerada stornirana/obrisana bez reseta Preradjeno flaga).

Private Sub Chk_D1_PreradjenoBezPrerade()
    On Error GoTo EH

    Dim refPal As Object: Set refPal = AggByBroj(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID, COL_PRES_NETO)
    Dim bad As Collection: Set bad = New Collection

    Dim data As Variant: data = GetTableData(TBL_PALETA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PALETA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cBroj As Long, cPre As Long
            cId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, "modIntegritet.D1")
            cBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, "modIntegritet.D1")
            cPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, "modIntegritet.D1")

            Dim i As Long, pid As String
            For i = 1 To UBound(data, 1)
                If UCase$(Trim$(CStr(data(i, cPre)))) = "DA" Then
                    pid = Trim$(CStr(data(i, cId)))
                    If Len(pid) > 0 Then
                        If Not refPal.Exists(pid) Then
                            bad.Add Array(pid, CStr(data(i, cBroj)))
                        End If
                    End If
                End If
            Next i
        End If
    End If

    WriteBlock "D1", "Palete Preradjeno=Da bez ijedne aktivne prerada-stavke", _
               Array("PaletaID", "BrojPalete"), CollToArray(bad, 2)
    Exit Sub

EH:
    WriteErr "D1", Err.description
End Sub

' ============================================================
' CHECK D2: PRERADA-STAVKA ka nevalidnoj paleti
' ============================================================
' Aktivna prerada-stavka ciji PaletaID ne postoji / stornirana paleta /
' paleta nije Preradjeno=Da. Ocekivano prazno (storno guard), ali orphan
' put (reset tblPrijemnica) moze proizvesti izuzetke.

Private Sub Chk_D2_PreradaNesvezaPaleta()
    On Error GoTo EH

    Dim palInfo As Object: Set palInfo = PaletaInfoMap()
    Dim bad As Collection: Set bad = New Collection

    Dim data As Variant: data = GetTableData(TBL_PRERADA_STAVKA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PRERADA_STAVKA)
        If Not IsEmpty(data) Then
            Dim cS As Long, cPre As Long, cPal As Long, cBrPal As Long
            cS = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_ID, "modIntegritet.D2")
            cPre = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID, "modIntegritet.D2")
            cPal = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID, "modIntegritet.D2")
            cBrPal = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_BROJ_PALETE, "modIntegritet.D2")

            Dim i As Long, pid As String, razlog As String, info As Variant
            For i = 1 To UBound(data, 1)
                pid = Trim$(CStr(data(i, cPal)))
                razlog = ""
                If Len(pid) = 0 Then
                    razlog = "prazan PaletaID"
                ElseIf Not palInfo.Exists(pid) Then
                    razlog = "paleta ne postoji"
                Else
                    info = palInfo(pid)
                    If info(0) = True Then
                        razlog = "paleta stornirana"
                    ElseIf UCase$(Trim$(CStr(info(1)))) <> "DA" Then
                        razlog = "paleta nije Preradjeno=Da"
                    End If
                End If
                If Len(razlog) > 0 Then
                    bad.Add Array(CStr(data(i, cS)), CStr(data(i, cPre)), pid, CStr(data(i, cBrPal)), razlog)
                End If
            Next i
        End If
    End If

    WriteBlock "D2", "Prerada-stavke ka nevalidnoj paleti (nesveza/stornirana/nepostojeca)", _
               Array("StavkaID", "PreradaID", "PaletaID", "BrojPalete", "Razlog"), CollToArray(bad, 5)
    Exit Sub

EH:
    WriteErr "D2", Err.description
End Sub

' ============================================================
' BLOK P: PRERADA 2.0 -- lager jedinica (Faza A). Invarijante I6/I7
' iz docs/PRERADA_2_MODEL_I_PLAN.md par. 5.
' ============================================================

' CHECK P4: svaka prerada ima TACNO JEDNU lager jedinicu (IzvorTip=PRERADA),
' KgPocetno = NetoIzlazKg, aktivna prerada sa tipom ima ProizvodID, obrnuti
' pokazivac tblPrerada.LagerJedinicaID pokazuje na tu jedinicu.
Private Sub Chk_P4_PreradaLagerJedinica()
    On Error GoTo EH
    Dim bad As Collection: Set bad = New Collection
    If GetTable(TBL_LAGER_JEDINICE) Is Nothing Then
        WriteBlock "P4", "Prerada <-> lager jedinica", _
                   Array("PreradaID", "BrojPrerade", "LagerJedinicaID", "Razlog"), Empty
        Exit Sub
    End If

    ' Mapa IzvorID -> (ljID, broj LJ, KgPocetno, ProizvodID) jednim prolazom.
    Dim lj As Variant: lj = GetTableData(TBL_LAGER_JEDINICE)
    Dim mapa As Object: Set mapa = CreateObject("Scripting.Dictionary")
    mapa.CompareMode = vbTextCompare
    If IsArray(lj) Then
        Dim lId As Long, lT As Long, lIz As Long, lKg As Long, lPr As Long, j As Long, k As String
        lId = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ID, "modIntegritet.P4")
        lT = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP, "modIntegritet.P4")
        lIz = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID, "modIntegritet.P4")
        lKg = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_KG_POCETNO, "modIntegritet.P4")
        lPr = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_PROIZVOD, "modIntegritet.P4")
        For j = 1 To UBound(lj, 1)
            If UCase$(Trim$(CStr(nz(lj(j, lT))))) = LJ_IZVOR_PRERADA Then
                k = Trim$(CStr(nz(lj(j, lIz))))
                If Len(k) > 0 Then
                    If mapa.Exists(k) Then
                        mapa(k) = Array(CStr(mapa(k)(0)), CLng(mapa(k)(1)) + 1, mapa(k)(2), mapa(k)(3))
                    Else
                        mapa.Add k, Array(Trim$(CStr(nz(lj(j, lId)))), 1, lj(j, lKg), _
                                          Trim$(CStr(nz(lj(j, lPr)))))
                    End If
                End If
            End If
        Next j
    End If

    Dim d As Variant: d = GetTableData(TBL_PRERADA)
    If IsArray(d) Then
        Dim cId As Long, cBroj As Long, cNeto As Long, cTip As Long, cSt As Long, cLj As Long
        cId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, "modIntegritet.P4")
        cBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, "modIntegritet.P4")
        cNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, "modIntegritet.P4")
        cTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, "modIntegritet.P4")
        cSt = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, "modIntegritet.P4")
        cLj = GetColumnIndex(TBL_PRERADA, COL_PRE_LJ_ID)

        Dim i As Long, pid As String, aktivna As Boolean, ljID As String, razlog As String
        Dim neto As Double, kg As Double
        For i = 1 To UBound(d, 1)
            pid = Trim$(CStr(nz(d(i, cId))))
            If Len(pid) > 0 Then
                aktivna = (UCase$(Trim$(CStr(nz(d(i, cSt))))) <> "DA")
                razlog = "": ljID = ""
                If Not mapa.Exists(pid) Then
                    razlog = "bez lager jedinice"
                ElseIf CLng(mapa(pid)(1)) > 1 Then
                    razlog = "vise lager jedinica (" & CStr(mapa(pid)(1)) & ")"
                    ljID = CStr(mapa(pid)(0))
                Else
                    ljID = CStr(mapa(pid)(0))
                    neto = 0: If IsNumeric(d(i, cNeto)) Then neto = CDbl(d(i, cNeto))
                    kg = 0: If IsNumeric(mapa(pid)(2)) Then kg = CDbl(mapa(pid)(2))
                    If aktivna And Abs(neto - kg) > 0.01 Then
                        razlog = "KgPocetno " & Format$(kg, "0.00") & " != NetoIzlazKg " & Format$(neto, "0.00")
                    ElseIf aktivna And Len(Trim$(CStr(nz(d(i, cTip))))) > 0 _
                           And Len(CStr(mapa(pid)(3))) = 0 Then
                        razlog = "LJ bez ProizvodID (tip '" & Trim$(CStr(nz(d(i, cTip)))) & "' nije u tblProizvodi)"
                    ElseIf cLj > 0 Then
                        If StrComp(Trim$(CStr(nz(d(i, cLj)))), ljID, vbTextCompare) <> 0 Then
                            razlog = "obrnuti pokazivac '" & Trim$(CStr(nz(d(i, cLj)))) & "' != LJ"
                        End If
                    End If
                End If
                If Len(razlog) > 0 Then bad.Add Array(pid, CStr(d(i, cBroj)), ljID, razlog)
            End If
        Next i
    End If

    WriteBlock "P4", "Prerada <-> lager jedinica (bez LJ / vise LJ / kg / proizvod / pokazivac)", _
               Array("PreradaID", "BrojPrerade", "LagerJedinicaID", "Razlog"), CollToArray(bad, 4)
    Exit Sub
EH:
    WriteErr "P4", Err.description
End Sub

' CHECK P5: aktivna utovarna/fakturna GP stavka nosi LagerJedinicaID koji
' postoji i cije poreklo (IzvorID) je bas ta prerada.
Private Sub Chk_P5_StavkeLagerJedinica()
    On Error GoTo EH
    Dim bad As Collection: Set bad = New Collection
    If Not GetTable(TBL_LAGER_JEDINICE) Is Nothing Then
        Dim izvorPoLj As Object: Set izvorPoLj = CreateObject("Scripting.Dictionary")
        izvorPoLj.CompareMode = vbTextCompare
        Dim lj As Variant: lj = GetTableData(TBL_LAGER_JEDINICE)
        If IsArray(lj) Then
            Dim lId As Long, lIz As Long, j As Long, k As String
            lId = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ID, "modIntegritet.P5")
            lIz = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID, "modIntegritet.P5")
            For j = 1 To UBound(lj, 1)
                k = Trim$(CStr(nz(lj(j, lId))))
                If Len(k) > 0 Then
                    If Not izvorPoLj.Exists(k) Then izvorPoLj.Add k, Trim$(CStr(nz(lj(j, lIz))))
                End If
            Next j
        End If
        P5Tabela bad, izvorPoLj, TBL_UTOVAR_STAVKE, COL_UTS_ID, COL_UTS_PRERADA_ID, COL_UTS_LJ_ID
        P5Tabela bad, izvorPoLj, TBL_FAKTURA_STAVKE, COL_FS_ID, COL_FS_PRERADA_ID, COL_FS_LJ_ID
    End If
    WriteBlock "P5", "GP stavke bez / sa pogresnom lager jedinicom", _
               Array("Tabela", "StavkaID", "PreradaID", "LagerJedinicaID", "Razlog"), CollToArray(bad, 5)
    Exit Sub
EH:
    WriteErr "P5", Err.description
End Sub

Private Sub P5Tabela(ByVal bad As Collection, ByVal izvorPoLj As Object, _
                     ByVal tbl As String, ByVal colId As String, _
                     ByVal colPre As String, ByVal colLj As String)
    If GetTable(tbl) Is Nothing Then Exit Sub
    Dim cId As Long, cPre As Long, cLj As Long
    cId = GetColumnIndex(tbl, colId)
    cPre = GetColumnIndex(tbl, colPre)
    cLj = GetColumnIndex(tbl, colLj)
    If cId = 0 Or cPre = 0 Or cLj = 0 Then Exit Sub
    Dim d As Variant: d = GetTableData(tbl)
    If Not IsArray(d) Then Exit Sub
    d = ExcludeStornirano(d, tbl)
    If Not IsArray(d) Then Exit Sub
    Dim i As Long, pre As String, ljID As String, razlog As String
    For i = 1 To UBound(d, 1)
        pre = Trim$(CStr(nz(d(i, cPre))))
        If Len(pre) > 0 Then
            ljID = Trim$(CStr(nz(d(i, cLj))))
            razlog = ""
            If Len(ljID) = 0 Then
                razlog = "bez LagerJedinicaID"
            ElseIf Not izvorPoLj.Exists(ljID) Then
                razlog = "LJ ne postoji"
            ElseIf StrComp(CStr(izvorPoLj(ljID)), pre, vbTextCompare) <> 0 Then
                razlog = "LJ pripada preradi '" & CStr(izvorPoLj(ljID)) & "'"
            End If
            If Len(razlog) > 0 Then bad.Add Array(tbl, CStr(nz(d(i, cId))), pre, ljID, razlog)
        End If
    Next i
End Sub

' CHECK P6: osnovne cinjenice lager jedinice -- jedinstven ID, poznat
' IzvorTip, neprazan IzvorID, StanicaID na aktivnoj jedinici.
Private Sub Chk_P6_LagerJedinicaOsnov()
    On Error GoTo EH
    Dim bad As Collection: Set bad = New Collection
    If Not GetTable(TBL_LAGER_JEDINICE) Is Nothing Then
        Dim d As Variant: d = GetTableData(TBL_LAGER_JEDINICE)
        If IsArray(d) Then
            Dim cId As Long, cT As Long, cIz As Long, cSt As Long, cStan As Long, cBroj As Long, cGod As Long
            cId = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_ID, "modIntegritet.P6")
            cT = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_TIP, "modIntegritet.P6")
            cIz = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_IZVOR_ID, "modIntegritet.P6")
            cSt = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_STORNIRANO, "modIntegritet.P6")
            cStan = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_STANICA, "modIntegritet.P6")
            cBroj = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_BROJ, "modIntegritet.P6")
            cGod = RequireColumnIndex(TBL_LAGER_JEDINICE, COL_LJ_GODINA, "modIntegritet.P6")
            Dim brojac As Object: Set brojac = CreateObject("Scripting.Dictionary")
            brojac.CompareMode = vbTextCompare
            Dim i As Long, id As String, tip As String, oznaka As String
            For i = 1 To UBound(d, 1)
                id = Trim$(CStr(nz(d(i, cId))))
                If brojac.Exists(id) Then brojac(id) = CLng(brojac(id)) + 1 Else brojac.Add id, 1
            Next i
            For i = 1 To UBound(d, 1)
                id = Trim$(CStr(nz(d(i, cId))))
                oznaka = CStr(nz(d(i, cBroj))) & "/" & CStr(nz(d(i, cGod)))
                tip = UCase$(Trim$(CStr(nz(d(i, cT)))))
                If Len(id) = 0 Then
                    bad.Add Array(id, oznaka, "prazan LagerJedinicaID")
                ElseIf CLng(brojac(id)) > 1 Then
                    bad.Add Array(id, oznaka, "dupli LagerJedinicaID")
                End If
                If tip <> LJ_IZVOR_SARZA And tip <> LJ_IZVOR_PALETA And tip <> LJ_IZVOR_PRERADA Then
                    bad.Add Array(id, oznaka, "nepoznat IzvorTip '" & tip & "'")
                ElseIf Len(Trim$(CStr(nz(d(i, cIz))))) = 0 Then
                    bad.Add Array(id, oznaka, "prazan IzvorID")
                End If
                If UCase$(Trim$(CStr(nz(d(i, cSt))))) <> "DA" Then
                    If Len(Trim$(CStr(nz(d(i, cStan))))) = 0 Then
                        bad.Add Array(id, oznaka, "aktivna jedinica bez StanicaID")
                    End If
                End If
            Next i
        End If
    End If
    WriteBlock "P6", "Lager jedinica -- osnovne cinjenice (ID, izvor, stanica)", _
               Array("LagerJedinicaID", "Oznaka", "Razlog"), CollToArray(bad, 3)
    Exit Sub
EH:
    WriteErr "P6", Err.description
End Sub

' ============================================================
' CHECK A5: prerada kg (ulaz vs Sigma stavke; izlaz <= ulaz)
' ============================================================

Private Sub Chk_A5_PreradaKg()
    On Error GoTo EH

    Dim stByPre As Object: Set stByPre = AggByBroj(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID, COL_PRES_NETO)
    Dim bad As Collection: Set bad = New Collection

    Dim data As Variant: data = GetTableData(TBL_PRERADA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_PRERADA)
        If Not IsEmpty(data) Then
            Dim cId As Long, cBroj As Long, cUlaz As Long, cIzlaz As Long
            cId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, "modIntegritet.A5")
            cBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, "modIntegritet.A5")
            cUlaz = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_ULAZ, "modIntegritet.A5")
            cIzlaz = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, "modIntegritet.A5")

            Dim i As Long, pid As String
            Dim ulaz As Double, izlaz As Double, sumSt As Double, razlog As String
            For i = 1 To UBound(data, 1)
                pid = Trim$(CStr(data(i, cId)))
                If Len(pid) > 0 Then
                    ulaz = 0: If IsNumeric(data(i, cUlaz)) Then ulaz = CDbl(data(i, cUlaz))
                    izlaz = 0: If IsNumeric(data(i, cIzlaz)) Then izlaz = CDbl(data(i, cIzlaz))
                    sumSt = 0: If stByPre.Exists(pid) Then sumSt = stByPre(pid)

                    razlog = ""
                    If Abs(ulaz - sumSt) > 0.5 Then
                        razlog = "NetoUlaz != Sigma stavke"
                    ElseIf izlaz > ulaz + 0.5 Then
                        razlog = "NetoIzlaz > NetoUlaz"
                    End If

                    If Len(razlog) > 0 Then
                        bad.Add Array(pid, CStr(data(i, cBroj)), ulaz, sumSt, izlaz, razlog)
                    End If
                End If
            Next i
        End If
    End If

    WriteBlock "A5", "Prerada kg (NetoUlaz vs Sigma stavke; NetoIzlaz <= NetoUlaz)", _
               Array("PreradaID", "BrojPrerade", "NetoUlazKg", "SumaStavkeKg", "NetoIzlazKg", "Razlog"), CollToArray(bad, 6)
    Exit Sub

EH:
    WriteErr "A5", Err.description
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
    d.CompareMode = vbTextCompare   ' kljucevi (BrojZbirne/ID) match case-insensitive

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
    d.CompareMode = vbTextCompare   ' BrojZbirne match case-insensitive (s5 == S5)

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

' Gradi dve mape iz tblZbirna: exactSet (case-SENZITIVAN skup tacnih brojeva)
' i ciMap (UCASE(broj) -> tacan zapis, prvi vidjeni). Za B6.
Private Sub BuildZbirnaCaseMaps(ByRef exactSet As Object, ByRef ciMap As Object)
    Set exactSet = CreateObject("Scripting.Dictionary")   ' default = case-senzitivan
    Set ciMap = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    If Not IsArray(data) Then Exit Sub

    Dim cb As Long: cb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modIntegritet.BuildZbirnaCaseMaps")

    Dim i As Long, b As String, u As String
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, cb)))
        If Len(b) > 0 Then
            If Not exactSet.Exists(b) Then exactSet.Add b, True
            u = UCase$(b)
            If Not ciMap.Exists(u) Then ciMap.Add u, b
        End If
    Next i
End Sub

' Za jednu tabelu: dodaj u 'bad' zive redove ciji BrojZbirne NIJE tacan case-match
' ali postoji ignorisuci case. Array(Tabela, ID, napisano, kanon).
Private Sub CollectCaseMismatch(ByRef bad As Collection, ByVal labela As String, _
                                ByVal tbl As String, ByVal idCol As String, ByVal zbrCol As String, _
                                ByVal exactSet As Object, ByVal ciMap As Object)
    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then Exit Sub
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then Exit Sub

    Dim cId As Long, cZbr As Long
    cId = RequireColumnIndex(tbl, idCol, "modIntegritet.CollectCaseMismatch")
    cZbr = RequireColumnIndex(tbl, zbrCol, "modIntegritet.CollectCaseMismatch")

    Dim i As Long, b As String, u As String
    For i = 1 To UBound(data, 1)
        b = Trim$(CStr(data(i, cZbr)))
        If Len(b) > 0 Then
            If Not exactSet.Exists(b) Then
                u = UCase$(b)
                If ciMap.Exists(u) Then
                    bad.Add Array(labela, CStr(data(i, cId)), b, ciMap(u))
                End If
            End If
        End If
    Next i
End Sub

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

' Zivi redovi tabele sa praznim BrojZbirne. Vraca 2D(1..n,1..3): ID, Broj, Kolicina.
Private Function DocsBezZbirne(ByVal tbl As String, ByVal idCol As String, _
                               ByVal brojCol As String, ByVal zbrCol As String, _
                               ByVal kolCol As String) As Variant
    Dim data As Variant: data = GetTableData(tbl)
    If Not IsArray(data) Then DocsBezZbirne = Empty: Exit Function
    data = ExcludeStornirano(data, tbl)
    If IsEmpty(data) Then DocsBezZbirne = Empty: Exit Function

    Dim cId As Long, cBroj As Long, cZbr As Long, cKol As Long
    cId = RequireColumnIndex(tbl, idCol, "modIntegritet.DocsBezZbirne")
    cBroj = RequireColumnIndex(tbl, brojCol, "modIntegritet.DocsBezZbirne")
    cZbr = RequireColumnIndex(tbl, zbrCol, "modIntegritet.DocsBezZbirne")
    cKol = RequireColumnIndex(tbl, kolCol, "modIntegritet.DocsBezZbirne")

    Dim bad As Collection: Set bad = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, cZbr)))) = 0 Then
            bad.Add Array(CStr(data(i, cId)), CStr(data(i, cBroj)), data(i, cKol))
        End If
    Next i

    DocsBezZbirne = CollToArray(bad, 3)
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

' PaletaID -> Array(storniranoBool, preradjenoStr). Iz SVIH redova tblPaleta.
Private Function PaletaInfoMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(TBL_PALETA)
    If Not IsArray(data) Then Set PaletaInfoMap = d: Exit Function

    Dim cId As Long, cStor As Long, cPre As Long
    cId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, "modIntegritet.PaletaInfoMap")
    cStor = RequireColumnIndex(TBL_PALETA, COL_STORNIRANO, "modIntegritet.PaletaInfoMap")
    cPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, "modIntegritet.PaletaInfoMap")

    Dim i As Long, pid As String, stor As Boolean, pre As String
    For i = 1 To UBound(data, 1)
        pid = Trim$(CStr(data(i, cId)))
        If Len(pid) > 0 Then
            stor = (UCase$(Trim$(CStr(data(i, cStor)))) = "DA")
            pre = Trim$(CStr(data(i, cPre)))
            If Not d.Exists(pid) Then d.Add pid, Array(stor, pre)
        End If
    Next i

    Set PaletaInfoMap = d
End Function

' PrijemnicaID -> BrojZbirne. Iz SVIH redova tblPrijemnica (i storniranih --
' C4 cilja bas prijemnice koje su stornirane).
Private Function PrijemnicaZbirnaMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(data) Then Set PrijemnicaZbirnaMap = d: Exit Function

    Dim cId As Long, cZbr As Long
    cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, "modIntegritet.PrijemnicaZbirnaMap")
    cZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "modIntegritet.PrijemnicaZbirnaMap")

    Dim i As Long, pid As String
    For i = 1 To UBound(data, 1)
        pid = Trim$(CStr(data(i, cId)))
        If Len(pid) > 0 Then
            If Not d.Exists(pid) Then d.Add pid, Trim$(CStr(data(i, cZbr)))
        End If
    Next i

    Set PrijemnicaZbirnaMap = d
End Function

' BrojZbirne -> naziv otkupnog mesta (distinct; "; " join ako ih je vise).
' Primarno preko AKTIVNIH otpremnica te zbirne (StanicaID -> tblStanice.Naziv;
' reuse BuildIdNameDict iz modDokumenta). Fallback za mirror-stanica zbirne
' (malina mod, VozacID == StanicaID, "S" prefiks): tblZbirna.VozacID direktno --
' pokriva i zbirne bez otpremnica / potpuno stornirane lance.
Private Function OtkupnoMestoByZbirna() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare   ' BrojZbirne match case-insensitive (s5 == S5)

    Dim stName As Object: Set stName = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")

    Dim data As Variant
    Dim i As Long, b As String, sid As String, nm As String

    ' 1) aktivne otpremnice zbirne: BrojZbirne -> distinct nazivi stanica
    data = GetTableData(TBL_OTPREMNICA)
    If IsArray(data) Then
        data = ExcludeStornirano(data, TBL_OTPREMNICA)
        If Not IsEmpty(data) Then
            Dim cZbr As Long, cSt As Long
            cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, "modIntegritet.OtkupnoMestoByZbirna")
            cSt = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, "modIntegritet.OtkupnoMestoByZbirna")

            For i = 1 To UBound(data, 1)
                b = Trim$(CStr(data(i, cZbr)))
                sid = Trim$(CStr(data(i, cSt)))
                If Len(b) > 0 And Len(sid) > 0 Then
                    nm = sid
                    If stName.Exists(sid) Then nm = CStr(stName(sid))
                    AddDistinctJoin d, b, nm
                End If
            Next i
        End If
    End If

    ' 2) fallback: zbirne bez pogotka gore, mirror-stanica (VozacID u tblStanice)
    data = GetTableData(TBL_ZBIRNA)
    If IsArray(data) Then
        Dim cBr As Long, cVoz As Long
        cBr = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modIntegritet.OtkupnoMestoByZbirna")
        cVoz = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, "modIntegritet.OtkupnoMestoByZbirna")

        For i = 1 To UBound(data, 1)
            b = Trim$(CStr(data(i, cBr)))
            If Len(b) > 0 Then
                If Not d.Exists(b) Then
                    sid = Trim$(CStr(data(i, cVoz)))
                    If stName.Exists(sid) Then d.Add b, CStr(stName(sid))
                End If
            End If
        Next i
    End If

    Set OtkupnoMestoByZbirna = d
End Function

' BrojZbirne -> proizvodjac(i) te zbirne (kooperanti; distinct, "; " join).
' Preko AKTIVNIH otkupa: primarno otkup.BrojZbirne (GetColumnIndex -- kolona
' je opciona, schema drift), fallback otkup.OtpremnicaID -> otpremnica.BrojZbirne
' (SVI redovi otpremnice, i stornirani -- trag porekla). Imena preko
' BuildIdNameDict (Ime+Prezime); nepoznat KooperantID ostaje kao ID.
Private Function ProizvodjacByZbirna() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare   ' BrojZbirne match case-insensitive (s5 == S5)

    Dim koopName As Object
    Set koopName = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")

    Dim data As Variant
    Dim i As Long, b As String, oid As String, kid As String, nm As String

    ' OtpremnicaID -> BrojZbirne (svi redovi) za fallback put
    Dim otpZbr As Object: Set otpZbr = CreateObject("Scripting.Dictionary")
    data = GetTableData(TBL_OTPREMNICA)
    If IsArray(data) Then
        Dim cOtpId As Long, cOtpZbr As Long
        cOtpId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, "modIntegritet.ProizvodjacByZbirna")
        cOtpZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, "modIntegritet.ProizvodjacByZbirna")
        For i = 1 To UBound(data, 1)
            oid = Trim$(CStr(data(i, cOtpId)))
            If Len(oid) > 0 Then
                If Not otpZbr.Exists(oid) Then otpZbr.Add oid, Trim$(CStr(data(i, cOtpZbr)))
            End If
        Next i
    End If

    data = GetTableData(TBL_OTKUP)
    If Not IsArray(data) Then Set ProizvodjacByZbirna = d: Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Set ProizvodjacByZbirna = d: Exit Function

    Dim cKoop As Long, cZbr As Long, cOtp As Long
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "modIntegritet.ProizvodjacByZbirna")
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)

    For i = 1 To UBound(data, 1)
        b = ""
        If cZbr > 0 Then b = Trim$(CStr(data(i, cZbr)))
        If Len(b) = 0 And cOtp > 0 Then
            oid = Trim$(CStr(data(i, cOtp)))
            If otpZbr.Exists(oid) Then b = CStr(otpZbr(oid))
        End If
        If Len(b) > 0 Then
            kid = Trim$(CStr(data(i, cKoop)))
            If Len(kid) > 0 Then
                nm = kid
                If koopName.Exists(kid) Then nm = CStr(koopName(kid))
                AddDistinctJoin d, b, nm
            End If
        End If
    Next i

    Set ProizvodjacByZbirna = d
End Function

' U d(key) upisi/dopisi nm kao "; "-odvojenu distinct listu (case-insensitive).
Private Sub AddDistinctJoin(ByRef d As Object, ByVal key As String, ByVal nm As String)
    If Len(nm) = 0 Then Exit Sub
    If Not d.Exists(key) Then
        d.Add key, nm
    ElseIf InStr(1, "; " & d(key) & "; ", "; " & nm & "; ", vbTextCompare) = 0 Then
        d(key) = d(key) & "; " & nm
    End If
End Sub

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
    m_ws.columns("A:I").AutoFit   ' C4 sada ima 9 kolona
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

' Levo poravnaj string na sirinu w (za monospace poravnanje kolona u ListBox-u).
Private Function PadR(ByVal s As String, ByVal w As Long) As String
    If Len(s) >= w Then PadR = s Else PadR = s & Space$(w - Len(s))
End Function

' Blok: naslov + (header + data) ili "OK - nema". Azurira summary + total.
Private Sub WriteBlock(ByVal code As String, ByVal title As String, _
                       ByVal headers As Variant, ByVal dataArr As Variant)
    Dim n As Long
    n = 0
    If IsArray(dataArr) Then
        If Not IsEmpty(dataArr) Then n = UBound(dataArr, 1)
    End If

    ' --- memorijski sink (samo blokovi sa problemima): podnaslov + kolone + nalazi ---
    ' Polja su padovana na max sirinu kolone u bloku -> poravnato uz monospace font.
    Dim i As Long, k As Long
    If n > 0 Then
        If m_rows.count > 0 Then m_rows.Add Array("", "")     ' prazan red izmedju oblasti

        m_rows.Add Array(code, "=== " & title & "  (" & CStr(n) & " zapisa) ===")

        Dim nc As Long: nc = UBound(dataArr, 2)
        Dim wds() As Long: ReDim wds(1 To nc)
        For k = 1 To nc
            wds(k) = Len(CStr(headers(LBound(headers) + k - 1)))
        Next k
        For i = 1 To n
            For k = 1 To nc
                If Len(CStr(dataArr(i, k))) > wds(k) Then wds(k) = Len(CStr(dataArr(i, k)))
            Next k
        Next i

        Dim line As String
        line = ""
        For k = 1 To nc
            If k > 1 Then line = line & " | "
            line = line & PadR(CStr(headers(LBound(headers) + k - 1)), wds(k))
        Next k
        m_rows.Add Array("", line)

        For i = 1 To n
            line = ""
            For k = 1 To nc
                If k > 1 Then line = line & " | "
                line = line & PadR(CStr(dataArr(i, k)), wds(k))
            Next k
            m_rows.Add Array("", line)
        Next i
    End If

    m_summary = m_summary & "  " & code & ": " & CStr(n) & vbCrLf
    m_totalIssues = m_totalIssues + n

    ' --- sheet sink ---
    If Not m_toSheet Then Exit Sub

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

        For i = 1 To n
            For k = 1 To UBound(dataArr, 2)
                m_ws.cells(m_row, k).value = dataArr(i, k)
            Next k
            m_row = m_row + 1
        Next i
    End If

    m_row = m_row + 1   ' prazan separator
End Sub

Private Sub WriteErr(ByVal code As String, ByVal desc As String)
    m_rows.Add Array(code, "GRESKA: " & desc)
    m_summary = m_summary & "  " & code & ": GRESKA" & vbCrLf
    If Not m_toSheet Then Exit Sub
    WriteLine "[" & code & "] GRESKA: " & desc, True
    m_row = m_row + 1
End Sub
