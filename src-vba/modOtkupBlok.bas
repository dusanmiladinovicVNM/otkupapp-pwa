Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok - rang kooperanata po iznosu otkupnih listova (KoopRangRows).
'
' S1b-3: sve sto je radilo nad vezom Otkup.OtpremnicaID je obrisano -- bilans
' otpremnice (SumKolByOtp, SumAmbByOtp, BuildNapisanoByOtp, ExistingBlokCena,
' ExistingBlokZbirna), specifikacija blokova po otpremnici (PrintSpecifikacija)
' i vezivanje bloka na otpremnicu (LinkOtkupIDsToOtpremnica). Taj model S3
' zamenjuje sa tblOtpremnicaIzvori; sposobnost vraca S3. Stari panel i clsBlokUI
' otisli su u S1b-2.
'
' Rang cita iznos i kilazu sa STAVKI (modOtkup.ZbirStavkiPoOtkupu).
' ============================================================

Private Function BuildLookup(ByVal tbl As String, ByVal keyName As String, _
                             ByVal valName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildLookup = d

    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ck As Long, cv As Long
    ck = GetColumnIndex(tbl, keyName)
    cv = GetColumnIndex(tbl, valName)
    If ck = 0 Or cv = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, ck)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, CStr(data(i, cv))
    Next i
End Function

Private Function BuildKoopNames() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildKoopNames = d

    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cIme As Long, cPr As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    If cId = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cId)))
        If Len(k) > 0 And Not d.Exists(k) Then
            d.Add k, Trim$(CStr(data(i, cIme)) & " " & CStr(data(i, cPr)))
        End If
    Next i
End Function

Private Function DictVal(ByVal d As Object, ByVal k As String) As String
    If d Is Nothing Then Exit Function
    k = Trim$(k)
    If d.Exists(k) Then DictVal = CStr(d(k))
End Function

' KooperantID -> naziv maticne stanice (OM), za listu kooperanata. Prazno ako
' kooperant nema stanicu ili stanica nema naziv (tada padne na StanicaID).
Private Function BuildKoopOM() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildKoopOM = d

    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cSt As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)
    cSt = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_STANICA)
    If cId = 0 Then Exit Function

    Dim dSt As Object: Set dSt = BuildLookup(TBL_STANICE, "StanicaID", "Naziv")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cId)))
        If Len(k) > 0 And Not d.Exists(k) Then
            Dim stId As String: stId = ""
            If cSt > 0 Then stId = Trim$(CStr(data(i, cSt)))
            Dim nm As String: nm = DictVal(dSt, stId)
            If Len(nm) = 0 Then nm = stId
            d.Add k, nm
        End If
    Next i
End Function

Private Function RowYear(ByVal v As Variant) As Integer
    On Error Resume Next
    If IsDate(v) Then RowYear = Year(CDate(v))
End Function

' Napuni rang: svi kooperanti firme sortirani opadajuce po ukupnom iznosu
' otkupnih listova (Sum Kolicina*Cena) u tekucoj godini; bez storniranih.
' RACUN ranga, bez ijedne kontrole. Vraca 1-bazirani 2D niz (n x 4):
'   1 KooperantID | 2 ime | 3 otkupno mesto | 4 iznos
' sortiran opadajuce po iznosu, plus kontrolne sume kroz izlazne parametre.
'
' Izdvojeno iz LoadKoopRang da isti racun mogu da koriste i legacy panel i novi
' ekran (modScrDokumenti), umesto da se agregacija prepisuje na dva mesta.
' Opsezne granice (odN/doN, serijski dani) su Optional: bez njih vazi staro
' pravilo "tekuca godina" (legacy panel i Dokumenti nepromenjeni); ekran
' Izvestaji salje svoj Od-Do, pa rang postuje isti period kao ostale liste.
Public Function KoopRangRows(ByRef rawKg As Double, ByRef rawVal As Double, _
                             ByRef emptyKg As Double, ByRef emptyVal As Double, _
                             Optional ByVal odN As Double = 0, _
                             Optional ByVal doN As Double = 0) As Variant
    On Error GoTo EH
    rawKg = 0: rawVal = 0: emptyKg = 0: emptyVal = 0
    Dim yr As Integer: yr = Year(Date)

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cKoop As Long, cId As Long, cDat As Long
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    If cKoop = 0 Or cId = 0 Then Exit Function

    ' Iznos i kilaza dokumenta su na STAVKAMA: CreateOtkup_TX ih na zaglavlju
    ' ostavlja prazne, pa je rang za nov dokument sabirao nulu (REFAKTOR S14.7).
    Dim stavkeZbir As Object: Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    Dim i As Long, uKrug As Boolean, dSer As Double
    For i = 1 To UBound(data, 1)
        If odN > 0 Or doN > 0 Then
            uKrug = False
            If cDat > 0 Then
                If IsDate(data(i, cDat)) Then
                    dSer = Int(CDbl(CDate(data(i, cDat))))
                    uKrug = (odN = 0 Or dSer >= odN) And (doN = 0 Or dSer <= doN)
                End If
            End If
        Else
            uKrug = (cDat = 0 Or RowYear(data(i, cDat)) = yr)
        End If
        If uKrug Then
            Dim kg As Double, vred As Double, zRang As Variant
            zRang = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, CStr(data(i, cId)), _
                        "modOtkupBlok.KoopRangRows")
            kg = CDbl(zRang(0))
            vred = CDbl(zRang(1))
            rawKg = rawKg + kg
            rawVal = rawVal + vred
            Dim k As String: k = Trim$(CStr(data(i, cKoop)))
            If Len(k) > 0 Then
                If agg.Exists(k) Then
                    agg(k) = CDbl(agg(k)) + vred
                Else
                    agg.Add k, vred
                End If
            Else
                emptyKg = emptyKg + kg
                emptyVal = emptyVal + vred
            End If
        End If
    Next i
    If agg.count = 0 Then Exit Function

    Dim n As Long: n = agg.count
    Dim koopIDs() As String: ReDim koopIDs(1 To n)
    Dim iznosi() As Double: ReDim iznosi(1 To n)
    Dim kk As Variant, j As Long: j = 0
    For Each kk In agg.keys
        j = j + 1
        koopIDs(j) = CStr(kk)
        iznosi(j) = CDbl(agg(kk))
    Next kk
    SortDescByVal koopIDs, iznosi

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim dOM As Object: Set dOM = BuildKoopOM()
    Dim outA() As Variant: ReDim outA(1 To n, 1 To 4)
    For i = 1 To n
        Dim nm2 As String: nm2 = DictVal(dKo, koopIDs(i))
        If Len(nm2) = 0 Then nm2 = koopIDs(i)
        outA(i, 1) = koopIDs(i)
        outA(i, 2) = nm2
        outA(i, 3) = DictVal(dOM, koopIDs(i))
        outA(i, 4) = iznosi(i)
    Next i
    KoopRangRows = outA
    Exit Function
EH:
    LogErr "modOtkupBlok.KoopRangRows"
End Function

' Opadajuci sort paralelnih nizova po iznosu (umeren broj kooperanata -> bubble).
Private Sub SortDescByVal(ByRef ids() As String, ByRef vals() As Double)
    Dim i As Long, j As Long, n As Long
    n = UBound(vals)
    For i = 1 To n - 1
        For j = 1 To n - i
            If vals(j) < vals(j + 1) Then
                Dim tv As Double: tv = vals(j): vals(j) = vals(j + 1): vals(j + 1) = tv
                Dim ts As String: ts = ids(j): ids(j) = ids(j + 1): ids(j + 1) = ts
            End If
        Next j
    Next i
End Sub
