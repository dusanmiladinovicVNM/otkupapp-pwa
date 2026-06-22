Attribute VB_Name = "modAmbalazaDijagnostika"
Option Explicit

' ============================================================
' modAmbalazaDijagnostika - READ-ONLY revizija tblAmbalaza
'
' Cilj: uporediti postojeci (potencijalno pre-refaktor) tblAmbalaza
' ledger sa izvornim dokumentima (tblOtkup, tblPrijemnica) i prijaviti
' gde podaci ne prate AKTUELNI entitetski-relativan model:
'
'   PREGLED  - broj redova po DokumentTip / EntitetTip (odmah vidi
'              da li uopste postoje Kooperant / Otkup noge).
'   SEKCIJA A - Prijemnica: orijentacija smera (novi vs stari obrnuti).
'               Novi: Kupac Ulaz = KolAmbalaze, Kupac Izlaz = KolAmbVracena.
'               Stari: obrnuto. KolAmbalaze == KolAmbVracena = NEUTRALNO
'               (net 0, orijentacija se ne moze utvrditi iz ledgera).
'   SEKCIJA B - Otkup: nedostajuce dvojne noge.
'               KolAmbalaze>0  -> ocekuj DokumentTip="Otkup":
'                 Kooperant Izlaz + Stanica Ulaz (oba = KolAmbalaze).
'               KolAmbIzdata>0 -> ocekuj DokumentTip="OM-Izlaz-Koop":
'                 Kooperant Ulaz + Stanica Izlaz (oba = KolAmbIzdata).
'   SEKCIJA C - Orphan / legacy redovi (nepoznat DokumentTip / EntitetTip,
'               ili Prijemnica/Otkup red bez izvornog dokumenta).
'
' NE MENJA PODATKE (cista revizija). Storno redovi se preskacu
' (ExcludeStornirano) kao i u produkcionim citanjima.
'
' Pokretanje: Alt+F8 -> DijagnostikaAmbalaze
' Izvestaj ide u Immediate prozor (Ctrl+G) + kratak MsgBox rezime.
' ============================================================

Private Const MAX_PRINT As Long = 60   ' max detaljnih redova po sekciji

' ------------------------------------------------------------
' JAVNI ULAZ
' ------------------------------------------------------------
Public Sub DijagnostikaAmbalaze()
    On Error GoTo EH

    Dim amb As Variant
    amb = LoadActive(TBL_AMBALAZA)
    If Not IsArray(amb) Then
        MsgBox "tblAmbalaza je prazna ili ne postoji (ili su svi redovi stornirani).", _
               vbExclamation, "Dijagnostika ambalaze"
        Exit Sub
    End If

    ' --- indeksi kolona tblAmbalaza ---
    Dim cAmbID As Long, cTip As Long, cKol As Long, cSmer As Long
    Dim cEntID As Long, cEntTip As Long, cDok As Long, cDokTip As Long
    cAmbID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ID)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    cKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    cSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    cEntID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)

    If cKol = 0 Or cSmer = 0 Or cEntTip = 0 Or cDok = 0 Or cDokTip = 0 Then
        MsgBox "tblAmbalaza nema ocekivane kolone (Kolicina/Smer/EntitetTip/DokumentID/DokumentTip)." & vbCrLf & _
               "Pokreni DebugKoloneTabele za stvarne nazive kolona.", _
               vbCritical, "Dijagnostika ambalaze"
        Exit Sub
    End If

    Debug.Print String(70, "=")
    Debug.Print "DIJAGNOSTIKA AMBALAZE  (read-only)   " & Format(Now, "yyyy-mm-dd hh:nn")
    Debug.Print String(70, "=")

    ' ------------------------------------------------------------
    ' PREGLED: broj aktivnih redova po DokumentTip i EntitetTip
    ' ------------------------------------------------------------
    Dim byDokTip As Object, byEntTip As Object
    Set byDokTip = CreateObject("Scripting.Dictionary"): byDokTip.CompareMode = vbTextCompare
    Set byEntTip = CreateObject("Scripting.Dictionary"): byEntTip.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To UBound(amb, 1)
        Inc byDokTip, NzKey(TxtV(amb(i, cDokTip)))
        Inc byEntTip, NzKey(TxtV(amb(i, cEntTip)))
    Next i

    Debug.Print ""
    Debug.Print "PREGLED  (ukupno aktivnih redova: " & UBound(amb, 1) & ")"
    Debug.Print "  Po DokumentTip:"
    PrintDict byDokTip
    Debug.Print "  Po EntitetTip:"
    PrintDict byEntTip

    Dim koopRedova As Long, otkupRedova As Long
    koopRedova = DictGet(byEntTip, "Kooperant")
    otkupRedova = DictGet(byDokTip, DOK_TIP_OTKUP)

    ' Index tblAmbalaza po DokumentID -> Collection redova
    Dim ambByDoc As Object
    Set ambByDoc = BuildAmbByDoc(amb, cDok, cDokTip, cSmer, cKol, cEntTip, cEntID, cAmbID)

    ' ------------------------------------------------------------
    ' SEKCIJA A - Prijemnica orijentacija
    ' ------------------------------------------------------------
    Dim aRev As Long, aOk As Long, aNeutral As Long, aMiss As Long, aMismatch As Long
    DiagPrijemnica ambByDoc, aRev, aOk, aNeutral, aMiss, aMismatch

    ' ------------------------------------------------------------
    ' SEKCIJA B - Otkup noge
    ' ------------------------------------------------------------
    Dim bOtkMiss As Long, bOtkPart As Long, bIzdMiss As Long, bIzdPart As Long, bOtkOk As Long
    DiagOtkup ambByDoc, bOtkOk, bOtkMiss, bOtkPart, bIzdMiss, bIzdPart

    ' ------------------------------------------------------------
    ' SEKCIJA C - Orphan / legacy
    ' ------------------------------------------------------------
    Dim cOrphan As Long
    DiagOrphan amb, cDok, cDokTip, cEntTip, cKol, cAmbID, cOrphan

    ' ------------------------------------------------------------
    ' REZIME
    ' ------------------------------------------------------------
    Dim msg As String
    msg = "DIJAGNOSTIKA AMBALAZE (detalji u Immediate / Ctrl+G)" & vbCrLf & vbCrLf
    msg = msg & "Aktivnih redova: " & UBound(amb, 1) & vbCrLf
    msg = msg & "  EntitetTip 'Kooperant': " & koopRedova & vbCrLf
    msg = msg & "  DokumentTip 'Otkup':    " & otkupRedova & vbCrLf & vbCrLf
    msg = msg & "A) PRIJEMNICA:" & vbCrLf
    msg = msg & "   OK (novi model): " & aOk & vbCrLf
    msg = msg & "   OBRNUTO (stari): " & aRev & vbCrLf
    msg = msg & "   Neutralno (=, net 0): " & aNeutral & vbCrLf
    msg = msg & "   Nedostaje ledger: " & aMiss & vbCrLf
    msg = msg & "   Iznos ne valja: " & aMismatch & vbCrLf & vbCrLf
    msg = msg & "B) OTKUP:" & vbCrLf
    msg = msg & "   OK: " & bOtkOk & " | nedostaju obe noge: " & bOtkMiss & " | delimicno: " & bOtkPart & vbCrLf
    msg = msg & "   Izdata (OM-Izlaz-Koop) nedostaje: " & bIzdMiss & " | delimicno: " & bIzdPart & vbCrLf & vbCrLf
    msg = msg & "C) Orphan / legacy redova: " & cOrphan

    Debug.Print ""
    Debug.Print String(70, "=")
    Debug.Print "ZAVRSENO."
    Debug.Print String(70, "=")

    MsgBox msg, vbInformation, "Dijagnostika ambalaze"
    Exit Sub

EH:
    MsgBox "Greska u dijagnostici: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Dijagnostika ambalaze"
End Sub

' ============================================================
' SEKCIJA A - Prijemnica orijentacija
' ============================================================
Private Sub DiagPrijemnica(ByVal ambByDoc As Object, _
                           ByRef nRev As Long, ByRef nOk As Long, _
                           ByRef nNeutral As Long, ByRef nMiss As Long, _
                           ByRef nMismatch As Long)
    nRev = 0: nOk = 0: nNeutral = 0: nMiss = 0: nMismatch = 0

    Debug.Print ""
    Debug.Print String(70, "-")
    Debug.Print "SEKCIJA A - PRIJEMNICA orijentacija smera"
    Debug.Print "  (novi: Kupac Ulaz=KolAmbalaze, Kupac Izlaz=KolAmbVracena)"
    Debug.Print String(70, "-")

    Dim pr As Variant
    pr = LoadActive(TBL_PRIJEMNICA)
    If Not IsArray(pr) Then
        Debug.Print "  tblPrijemnica prazna / ne postoji - preskoceno."
        Exit Sub
    End If

    Dim cID As Long, cKol As Long, cVrac As Long
    cID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    cVrac = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA)
    If cID = 0 Or cKol = 0 Then
        Debug.Print "  tblPrijemnica nema PrijemnicaID / KolAmbalaze - preskoceno."
        Exit Sub
    End If

    Dim shown As Long: shown = 0
    Dim i As Long
    For i = 1 To UBound(pr, 1)
        Dim pid As String, kolAmb As Double, kolVrac As Double
        pid = TxtV(pr(i, cID))
        kolAmb = NumV(pr(i, cKol))
        If cVrac > 0 Then kolVrac = NumV(pr(i, cVrac)) Else kolVrac = 0

        ' Prijemnica bez ambalaze - nije predmet revizije
        If kolAmb = 0 And kolVrac = 0 Then GoTo NextPr

        Dim ulaz As Double, izlaz As Double, nRows As Long
        SumKupac ambByDoc, pid, ulaz, izlaz, nRows

        Dim status As String
        status = ClassifyPrijemnica(kolAmb, kolVrac, ulaz, izlaz, nRows)

        Select Case status
            Case "OK":       nOk = nOk + 1
            Case "OBRNUTO":  nRev = nRev + 1
            Case "NEUTRALNO": nNeutral = nNeutral + 1
            Case "NEDOSTAJE": nMiss = nMiss + 1
            Case Else:       nMismatch = nMismatch + 1
        End Select

        ' Detaljno stampaj samo probleme (OBRNUTO/NEDOSTAJE/MISMATCH)
        If status <> "OK" And status <> "NEUTRALNO" Then
            If shown < MAX_PRINT Then
                Debug.Print "  [" & status & "] " & pid & _
                            "  KolAmb=" & kolAmb & " Vrac=" & kolVrac & _
                            "  | ledger Kupac Ulaz=" & ulaz & " Izlaz=" & izlaz & _
                            " (redova=" & nRows & ")"
                shown = shown + 1
            End If
        End If
NextPr:
    Next i

    Debug.Print "  --> OK=" & nOk & "  OBRNUTO=" & nRev & "  NEUTRALNO=" & nNeutral & _
                "  NEDOSTAJE=" & nMiss & "  MISMATCH=" & nMismatch
    If shown >= MAX_PRINT Then Debug.Print "  (prikaz skracen na prvih " & MAX_PRINT & ")"
End Sub

Private Function ClassifyPrijemnica(ByVal kolAmb As Double, ByVal kolVrac As Double, _
                                    ByVal ulaz As Double, ByVal izlaz As Double, _
                                    ByVal nRows As Long) As String
    If nRows = 0 Then
        ClassifyPrijemnica = "NEDOSTAJE"
        Exit Function
    End If

    ' Oba iznosa > 0 i jednaka -> orijentacija se ne moze utvrditi (net 0)
    If kolAmb > 0 And kolVrac > 0 And kolAmb = kolVrac Then
        If ulaz = kolAmb And izlaz = kolAmb Then
            ClassifyPrijemnica = "NEUTRALNO"
        Else
            ClassifyPrijemnica = "MISMATCH"
        End If
        Exit Function
    End If

    ' Novi model: Ulaz = KolAmbalaze, Izlaz = KolAmbVracena
    If ulaz = kolAmb And izlaz = kolVrac Then
        ClassifyPrijemnica = "OK"
        Exit Function
    End If

    ' Stari obrnuti: Izlaz = KolAmbalaze, Ulaz = KolAmbVracena
    If izlaz = kolAmb And ulaz = kolVrac Then
        ClassifyPrijemnica = "OBRNUTO"
        Exit Function
    End If

    ClassifyPrijemnica = "MISMATCH"
End Function

Private Sub SumKupac(ByVal ambByDoc As Object, ByVal docID As String, _
                     ByRef ulaz As Double, ByRef izlaz As Double, ByRef nRows As Long)
    ulaz = 0: izlaz = 0: nRows = 0
    If Len(docID) = 0 Then Exit Sub
    If Not ambByDoc.Exists(docID) Then Exit Sub

    Dim it As Variant
    For Each it In ambByDoc(docID)
        ' it = Array(smer, kol, entTip, entID, dokTip, ambID)
        If StrComp(CStr(it(2)), "Kupac", vbTextCompare) = 0 Then
            nRows = nRows + 1
            Select Case LCase$(CStr(it(0)))
                Case "ulaz":  ulaz = ulaz + CDbl(it(1))
                Case "izlaz": izlaz = izlaz + CDbl(it(1))
            End Select
        End If
    Next it
End Sub

' ============================================================
' SEKCIJA B - Otkup noge
' ============================================================
Private Sub DiagOtkup(ByVal ambByDoc As Object, _
                      ByRef nOk As Long, ByRef nMiss As Long, ByRef nPart As Long, _
                      ByRef nIzdMiss As Long, ByRef nIzdPart As Long)
    nOk = 0: nMiss = 0: nPart = 0: nIzdMiss = 0: nIzdPart = 0

    Debug.Print ""
    Debug.Print String(70, "-")
    Debug.Print "SEKCIJA B - OTKUP dvojne noge"
    Debug.Print "  (KolAmbalaze -> Kooperant Izlaz + Stanica Ulaz, DokumentTip='Otkup')"
    Debug.Print "  (KolAmbIzdata -> Kooperant Ulaz + Stanica Izlaz, 'OM-Izlaz-Koop')"
    Debug.Print "  NAPOMENA: ako otkupni blokovi nisu unoseni, ova sekcija je"
    Debug.Print "  OCEKIVANO prazna - odsustvo Kooperant/Otkup nogu nije greska."
    Debug.Print String(70, "-")

    Dim ot As Variant
    ot = LoadActive(TBL_OTKUP)
    If Not IsArray(ot) Then
        Debug.Print "  tblOtkup prazna / ne postoji - preskoceno."
        Exit Sub
    End If

    Dim cID As Long, cKol As Long, cIzd As Long
    cID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    cIzd = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA)
    If cID = 0 Or cKol = 0 Then
        Debug.Print "  tblOtkup nema OtkupID / KolAmbalaze - preskoceno."
        Exit Sub
    End If

    Dim shown As Long: shown = 0
    Dim i As Long
    For i = 1 To UBound(ot, 1)
        Dim oid As String, kolAmb As Double, kolIzd As Double
        oid = TxtV(ot(i, cID))
        kolAmb = NumV(ot(i, cKol))
        If cIzd > 0 Then kolIzd = NumV(ot(i, cIzd)) Else kolIzd = 0

        ' --- glavna noga (Otkup) ---
        If kolAmb > 0 Then
            Dim kIz As Double, sUl As Double
            kIz = SumLeg(ambByDoc, oid, "Kooperant", "Izlaz", DOK_TIP_OTKUP)
            sUl = SumLeg(ambByDoc, oid, "Stanica", "Ulaz", DOK_TIP_OTKUP)

            Dim st As String
            st = ClassifyLeg(kolAmb, kIz, sUl)
            Select Case st
                Case "OK": nOk = nOk + 1
                Case "NEDOSTAJE": nMiss = nMiss + 1
                Case Else: nPart = nPart + 1
            End Select
            If st <> "OK" Then
                If shown < MAX_PRINT Then
                    Debug.Print "  [OTKUP " & st & "] " & oid & "  KolAmb=" & kolAmb & _
                                "  | KoopIzlaz=" & kIz & " StanicaUlaz=" & sUl
                    shown = shown + 1
                End If
            End If
        End If

        ' --- izdata prazna (OM-Izlaz-Koop) ---
        If kolIzd > 0 Then
            Dim kUl As Double, sIz As Double
            kUl = SumLeg(ambByDoc, oid, "Kooperant", "Ulaz", DOK_TIP_OM_IZLAZ_KOOP)
            sIz = SumLeg(ambByDoc, oid, "Stanica", "Izlaz", DOK_TIP_OM_IZLAZ_KOOP)

            Dim st2 As String
            st2 = ClassifyLeg(kolIzd, kUl, sIz)
            Select Case st2
                Case "OK": ' ok
                Case "NEDOSTAJE": nIzdMiss = nIzdMiss + 1
                Case Else: nIzdPart = nIzdPart + 1
            End Select
            If st2 <> "OK" Then
                If shown < MAX_PRINT Then
                    Debug.Print "  [IZDATA " & st2 & "] " & oid & "  KolIzdata=" & kolIzd & _
                                "  | KoopUlaz=" & kUl & " StanicaIzlaz=" & sIz
                    shown = shown + 1
                End If
            End If
        End If
    Next i

    Debug.Print "  --> Otkup: OK=" & nOk & " nedostaju obe=" & nMiss & " delimicno=" & nPart
    Debug.Print "  --> Izdata: nedostaju obe=" & nIzdMiss & " delimicno=" & nIzdPart
    If shown >= MAX_PRINT Then Debug.Print "  (prikaz skracen na prvih " & MAX_PRINT & ")"
End Sub

Private Function ClassifyLeg(ByVal ocekivano As Double, _
                             ByVal nogaA As Double, ByVal nogaB As Double) As String
    If nogaA = ocekivano And nogaB = ocekivano Then
        ClassifyLeg = "OK"
    ElseIf nogaA = 0 And nogaB = 0 Then
        ClassifyLeg = "NEDOSTAJE"
    Else
        ClassifyLeg = "DELIMICNO"
    End If
End Function

Private Function SumLeg(ByVal ambByDoc As Object, ByVal docID As String, _
                        ByVal entTip As String, ByVal smer As String, _
                        ByVal dokTip As String) As Double
    SumLeg = 0
    If Len(docID) = 0 Then Exit Function
    If Not ambByDoc.Exists(docID) Then Exit Function

    Dim it As Variant
    For Each it In ambByDoc(docID)
        ' it = Array(smer, kol, entTip, entID, dokTip, ambID)
        If StrComp(CStr(it(2)), entTip, vbTextCompare) = 0 _
           And StrComp(CStr(it(0)), smer, vbTextCompare) = 0 _
           And StrComp(CStr(it(4)), dokTip, vbTextCompare) = 0 Then
            SumLeg = SumLeg + CDbl(it(1))
        End If
    Next it
End Function

' ============================================================
' SEKCIJA C - Orphan / legacy
' ============================================================
Private Sub DiagOrphan(ByVal amb As Variant, ByVal cDok As Long, ByVal cDokTip As Long, _
                       ByVal cEntTip As Long, ByVal cKol As Long, ByVal cAmbID As Long, _
                       ByRef nOrphan As Long)
    nOrphan = 0

    Debug.Print ""
    Debug.Print String(70, "-")
    Debug.Print "SEKCIJA C - Orphan / legacy redovi"
    Debug.Print String(70, "-")

    ' poznati skupovi
    Dim knownDok As Object
    Set knownDok = CreateObject("Scripting.Dictionary"): knownDok.CompareMode = vbTextCompare
    knownDok(DOK_TIP_OM_ULAZ) = True
    knownDok(DOK_TIP_OTPREMNICA) = True
    knownDok(DOK_TIP_PRIJEMNICA) = True
    knownDok(DOK_TIP_OTKUP) = True
    knownDok(DOK_TIP_IZLAZ_KUPCI) = True
    knownDok(DOK_TIP_OM_IZLAZ_KOOP) = True

    Dim knownEnt As Object
    Set knownEnt = CreateObject("Scripting.Dictionary"): knownEnt.CompareMode = vbTextCompare
    knownEnt("Stanica") = True
    knownEnt("Kupac") = True
    knownEnt("Kooperant") = True

    ' izvorni ID setovi za proveru veze
    Dim prIDs As Object, otIDs As Object
    Set prIDs = BuildIDSet(TBL_PRIJEMNICA, COL_PRJ_ID)
    Set otIDs = BuildIDSet(TBL_OTKUP, COL_OTK_ID)

    Dim shown As Long: shown = 0
    Dim i As Long
    For i = 1 To UBound(amb, 1)
        Dim dt As String, et As String, did As String
        dt = TxtV(amb(i, cDokTip))
        et = TxtV(amb(i, cEntTip))
        did = TxtV(amb(i, cDok))

        Dim reason As String: reason = ""
        If Len(dt) = 0 Then
            reason = "prazan DokumentTip"
        ElseIf Not knownDok.Exists(dt) Then
            reason = "nepoznat DokumentTip '" & dt & "'"
        End If

        If Len(reason) = 0 Then
            If Len(et) = 0 Then
                reason = "prazan EntitetTip"
            ElseIf Not knownEnt.Exists(et) Then
                reason = "nepoznat EntitetTip '" & et & "'"
            End If
        End If

        If Len(reason) = 0 Then
            If StrComp(dt, DOK_TIP_PRIJEMNICA, vbTextCompare) = 0 Then
                If prIDs.count > 0 And Not prIDs.Exists(did) Then reason = "Prijemnica bez izvora (" & did & ")"
            ElseIf StrComp(dt, DOK_TIP_OTKUP, vbTextCompare) = 0 _
                Or StrComp(dt, DOK_TIP_OM_IZLAZ_KOOP, vbTextCompare) = 0 Then
                If otIDs.count > 0 And Not otIDs.Exists(did) Then reason = "Otkup-noga bez izvora (" & did & ")"
            End If
        End If

        If Len(reason) > 0 Then
            nOrphan = nOrphan + 1
            If shown < MAX_PRINT Then
                Debug.Print "  [" & TxtV(amb(i, cAmbID)) & "] " & reason & _
                            "  (Kol=" & NumV(amb(i, cKol)) & ")"
                shown = shown + 1
            End If
        End If
    Next i

    Debug.Print "  --> Orphan / legacy redova: " & nOrphan
    If shown >= MAX_PRINT Then Debug.Print "  (prikaz skracen na prvih " & MAX_PRINT & ")"
End Sub

' ============================================================
' HELPERI
' ============================================================
Private Function LoadActive(ByVal tbl As String) As Variant
    Dim d As Variant
    d = GetTableData(tbl)
    If IsEmpty(d) Then
        LoadActive = Empty
        Exit Function
    End If
    d = ExcludeStornirano(d, tbl)
    LoadActive = d
End Function

Private Function BuildAmbByDoc(ByVal amb As Variant, ByVal cDok As Long, ByVal cDokTip As Long, _
                               ByVal cSmer As Long, ByVal cKol As Long, ByVal cEntTip As Long, _
                               ByVal cEntID As Long, ByVal cAmbID As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = 1 To UBound(amb, 1)
        key = TxtV(amb(i, cDok))
        If Len(key) > 0 Then
            Dim col As Collection
            If dict.Exists(key) Then
                Set col = dict(key)
            Else
                Set col = New Collection
                dict.Add key, col
            End If
            col.Add Array( _
                TxtV(amb(i, cSmer)), NumV(amb(i, cKol)), _
                TxtV(amb(i, cEntTip)), TxtV(amb(i, cEntID)), _
                TxtV(amb(i, cDokTip)), TxtV(amb(i, cAmbID)))
        End If
    Next i
    Set BuildAmbByDoc = dict
End Function

Private Function BuildIDSet(ByVal tbl As String, ByVal colName As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim d As Variant
    d = LoadActive(tbl)
    If Not IsArray(d) Then
        Set BuildIDSet = dict
        Exit Function
    End If

    Dim c As Long
    c = GetColumnIndex(tbl, colName)
    If c = 0 Then
        Set BuildIDSet = dict
        Exit Function
    End If

    Dim i As Long, k As String
    For i = 1 To UBound(d, 1)
        k = TxtV(d(i, c))
        If Len(k) > 0 Then dict(k) = True
    Next i
    Set BuildIDSet = dict
End Function

Private Sub Inc(ByVal dict As Object, ByVal key As String)
    If dict.Exists(key) Then
        dict(key) = dict(key) + 1
    Else
        dict.Add key, 1
    End If
End Sub

Private Function DictGet(ByVal dict As Object, ByVal key As String) As Long
    If dict.Exists(key) Then DictGet = dict(key) Else DictGet = 0
End Function

Private Sub PrintDict(ByVal dict As Object)
    Dim k As Variant
    For Each k In dict.keys
        Debug.Print "      " & k & " : " & dict(k)
    Next k
End Sub

Private Function NzKey(ByVal s As String) As String
    If Len(s) = 0 Then NzKey = "(prazno)" Else NzKey = s
End Function

Private Function TxtV(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        TxtV = ""
    Else
        TxtV = Trim$(CStr(v))
    End If
End Function

Private Function NumV(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumV = CDbl(v) Else NumV = 0
End Function
