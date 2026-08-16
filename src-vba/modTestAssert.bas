Attribute VB_Name = "modTestAssert"
Option Explicit

' ============================================================
' modTestAssert -- jedan assertion API za sve suite-ove.
'
' Do sada je isti posao postojao u pet-sest kopija: Chk/ChkEq/ChkEqD u
' modTestPalete i modTestStorno, AssertTrue/AssertEq u modLicenseTests, jos
' jedan par u modGoogleSyncSmokeTests, treci u modSEFTests, a modTest ima svoj
' AssertEq koji umesto brojanja PODIZE gresku. Razlike izmedju kopija nisu bile
' namerne, samo su nastale u razlicito vreme.
'
' Svaki assert ovde broji kroz modTestRunner (TR_Pass/TR_Fail) i NE prekida
' izvrsavanje. To je namerno: jedan pao assert ne sme da sakrije preostalih
' dvadeset u istom testu. Suite pada na kraju, kroz TR_EndSuite.
'
' PROVERA DA JE NESTO ODBIJENO -- AssertRaised
'     Velik deo pravila u ovom projektu su zabrane ("ovo mora da bude odbijeno"),
'     pa je do sada svaki modul rucno hvatao Err.Number. VBA nema nacin da se
'     poziv prosledi kao vrednost (nema delegata), pa se obrazac ne moze sakriti
'     do kraja -- ali moze da se svede na tri linije koje su svuda iste:
'
'         On Error Resume Next
'         Err.Clear
'         SaveNesto "nevalidno"                  ' ovo MORA da pukne
'         AssertRaised "prazan kooperant se odbija"
'         On Error GoTo 0
'
'     `AssertRaised` cita Err.Number kao PRVU stvar i odmah ga cisti. Zato izmedju
'     poziva koji treba da pukne i AssertRaised ne sme da stoji nijedna druga
'     naredba -- ubacena naredba koja uspe ostavlja Err.Number = 0 i provera bi
'     lagala u smeru "nije podignuto".
' ============================================================

' ============================================================
' Osnovni
' ============================================================

Public Sub AssertTrue(ByVal cond As Boolean, ByVal label As String)
    If cond Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivano True"
    End If
End Sub

Public Sub AssertFalse(ByVal cond As Boolean, ByVal label As String)
    If Not cond Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivano False"
    End If
End Sub

' Poredi se kao TEKST. Razlog je VBA: `1` i `"1"` i `1#` su tri tipa koja u
' ovom kodu dolaze iz iste celije, zavisno od toga da li je vrednost citana iz
' tabele, iz kontrole ili izracunata. Poredjenje po tipu bi davalo padove koji
' nemaju veze sa pravilom koje se testira. Za brojeve sa decimalama postoji
' AssertNear.
Public Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, _
                       ByVal label As String)
    If SafeStr(actual) = SafeStr(expected) Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivano [" & SafeStr(expected) & _
                "], dobijeno [" & SafeStr(actual) & "]"
    End If
End Sub

Public Sub AssertNotEqual(ByVal actual As Variant, ByVal notExpected As Variant, _
                          ByVal label As String)
    If SafeStr(actual) <> SafeStr(notExpected) Then
        TR_Pass label
    Else
        TR_Fail label & " -- vrednost NE sme da bude [" & SafeStr(notExpected) & "]"
    End If
End Sub

' Novac i kilogrami. Podrazumevana tolerancija 0.001 je ista kao u ChkEqD iz
' modTestPalete -- ne uvodi se nova konvencija. Stoji kao literal u potpisu jer
' VBA za Optional default trazi literal.
Public Sub AssertNear(ByVal actual As Double, ByVal expected As Double, _
                      ByVal label As String, Optional ByVal tol As Double = 0.001)
    If Abs(actual - expected) <= tol Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivano " & Format$(expected, "0.####") & _
                " +-" & Format$(tol, "0.####") & ", dobijeno " & Format$(actual, "0.####")
    End If
End Sub

' Empty, Null, "" i sami razmaci su ista stvar za poslovnu logiku ovog projekta
' (prazna celija ume da dodje kao bilo sta od toga).
Public Sub AssertEmpty(ByVal value As Variant, ByVal label As String)
    If Len(Trim$(SafeStr(value))) = 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivano prazno, dobijeno [" & SafeStr(value) & "]"
    End If
End Sub

Public Sub AssertNotEmpty(ByVal value As Variant, ByVal label As String)
    If Len(Trim$(SafeStr(value))) > 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- ocekivana vrednost, dobijeno prazno"
    End If
End Sub

Public Sub AssertContains(ByVal haystack As Variant, ByVal needle As String, _
                          ByVal label As String)
    If InStr(1, SafeStr(haystack), needle, vbTextCompare) > 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- [" & SafeStr(haystack) & "] ne sadrzi [" & needle & "]"
    End If
End Sub

' ============================================================
' Greske -- "ovo mora da bude odbijeno"
' ============================================================

Public Sub AssertRaised(ByVal label As String)
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    Err.Clear

    If errNum <> 0 Then
        TR_Pass label & " (odbijeno: " & CStr(errNum) & ")"
    Else
        TR_Fail label & " -- ocekivana greska, a nista nije podignuto"
    End If
End Sub

Public Sub AssertRaisedNumber(ByVal expectedNum As Long, ByVal label As String)
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    Err.Clear

    If errNum = expectedNum Then
        TR_Pass label
    ElseIf errNum <> 0 Then
        TR_Fail label & " -- podignuta greska " & CStr(errNum) & " (" & errDesc & _
                "), a ocekivana je " & CStr(expectedNum)
    Else
        TR_Fail label & " -- ocekivana greska " & CStr(expectedNum) & ", a nista nije podignuto"
    End If
End Sub

Public Sub AssertNoError(ByVal label As String)
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    Err.Clear

    If errNum = 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- neocekivana greska " & CStr(errNum) & ": " & errDesc
    End If
End Sub

' ============================================================
' Tabele
' ============================================================

' FAIL-CLOSED. Provere nad tabelama moraju da razlikuju TRI ishoda, ne dva:
'
'   nema reda            -> validan poslovni rezultat
'   ima reda             -> validan poslovni rezultat
'   nema tabele/kolone   -> greska u TESTU, ne odgovor o poslovnoj logici
'
' Prva verzija je trecu spajala sa prvom: `FindRows` pod `On Error Resume Next`
' je vracao Nothing, brojac je davao 0, pa je `AssertTableRowMissing` prijavljivao
' PASS i kad tabela uopste ne postoji. Test infrastruktura koja na pogresno ime
' tabele odgovara "provera prosla" je gora od one koje nema.

Public Sub AssertRowCount(ByVal tblName As String, ByVal expected As Long, _
                          ByVal label As String)
    Dim n As Long
    n = TableRowCount(tblName)
    If n < 0 Then
        TR_Fail label & " -- INFRASTRUKTURA: tabela " & tblName & " ne postoji"
        Exit Sub
    End If
    AssertEqual n, expected, label
End Sub

Public Sub AssertTableRowExists(ByVal tblName As String, ByVal colName As String, _
                                ByVal value As Variant, ByVal label As String)
    Dim hits As Long
    Dim okFlag As Boolean
    hits = CountMatching(tblName, colName, value, okFlag)
    If Not okFlag Then
        TR_Fail label & " -- " & InfraReason(tblName, colName)
    ElseIf hits > 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- nema reda u " & tblName & " sa " & colName & _
                "=[" & SafeStr(value) & "]"
    End If
End Sub

Public Sub AssertTableRowMissing(ByVal tblName As String, ByVal colName As String, _
                                 ByVal value As Variant, ByVal label As String)
    Dim hits As Long
    Dim okFlag As Boolean
    hits = CountMatching(tblName, colName, value, okFlag)
    If Not okFlag Then
        ' Bez ovoga bi "cela tabela ne postoji" prolazilo kao "reda nema".
        TR_Fail label & " -- " & InfraReason(tblName, colName)
    ElseIf hits = 0 Then
        TR_Pass label
    Else
        TR_Fail label & " -- " & tblName & " i dalje ima " & CStr(hits) & _
                " red(ova) sa " & colName & "=[" & SafeStr(value) & "]"
    End If
End Sub

' Broj redova tabele NA ULASKU vs NA IZLASKU. Za provere tipa "ova putanja ne sme
' nista da upise" i za dokaz da je rollback stvarno vratio stanje -- transakcija
' koja "bi trebalo" da rollback-uje nije dokaz da jeste.
'
' -1 znaci NEMA TABELE i nikad ne sme da se poredi sa drugom -1 kao "isto".
Public Function TableRowCount(ByVal tblName As String) As Long
    Dim lo As ListObject
    On Error Resume Next
    Set lo = GetTable(tblName)
    On Error GoTo 0
    If lo Is Nothing Then
        TableRowCount = -1
    ElseIf lo.ListRows Is Nothing Then
        TableRowCount = 0
    Else
        TableRowCount = lo.ListRows.count
    End If
End Function

Public Sub AssertUnchanged(ByVal tblName As String, ByVal countBefore As Long, _
                           ByVal label As String)
    Dim nowCount As Long
    nowCount = TableRowCount(tblName)

    ' Dve -1 vrednosti bi se poklopile i dale PASS ("nista se nije promenilo"),
    ' a znace da tabele nema ni pre ni posle -- provera tada ne meri nista.
    If countBefore < 0 Or nowCount < 0 Then
        TR_Fail label & " -- INFRASTRUKTURA: tabela " & tblName & " ne postoji " & _
                "(pre=" & CStr(countBefore) & ", posle=" & CStr(nowCount) & ")"
        Exit Sub
    End If

    If nowCount = countBefore Then
        TR_Pass label
    Else
        TR_Fail label & " -- " & tblName & " je imala " & CStr(countBefore) & _
                " redova, sada ima " & CStr(nowCount)
    End If
End Sub

' ============================================================
' Pomocno
' ============================================================

' Prazna kontrola ume da vrati Null umesto "", a CStr(Null) puca ("Invalid use
' of Null") -- provera bi pala na toj gresci umesto na ponasanju koje meri.
' Isti oblik kao modTest.SafeStr; tamo je Private jer modTest ima svoj
' raise-on-fail model.
Public Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Then
        SafeStr = ""
    ElseIf IsObject(v) Then
        SafeStr = "<object>"
    ElseIf IsArray(v) Then
        SafeStr = "<array>"
    Else
        SafeStr = CStr(v)
    End If
    On Error GoTo 0
End Function

' okFlag = False znaci da odgovor NE VREDI (nema tabele, nema kolone, FindRows
' pukao). Pozivalac to mora da prijavi kao pad, ne kao "nula pogodaka".
Private Function CountMatching(ByVal tblName As String, ByVal colName As String, _
                               ByVal value As Variant, ByRef okFlag As Boolean) As Long
    okFlag = False
    CountMatching = 0

    If GetTable(tblName) Is Nothing Then Exit Function
    If GetColumnIndex(tblName, colName) = 0 Then Exit Function

    Dim hitRows As Collection
    On Error Resume Next
    Err.Clear
    Set hitRows = FindRows(tblName, colName, value)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If hitRows Is Nothing Then Exit Function

    okFlag = True
    CountMatching = hitRows.count
End Function

' Zasto odgovor ne vredi -- tacan razlog, da se ne pogadja.
Private Function InfraReason(ByVal tblName As String, ByVal colName As String) As String
    If GetTable(tblName) Is Nothing Then
        InfraReason = "INFRASTRUKTURA: tabela " & tblName & " ne postoji"
    ElseIf GetColumnIndex(tblName, colName) = 0 Then
        InfraReason = "INFRASTRUKTURA: " & tblName & " nema kolonu " & colName
    Else
        InfraReason = "INFRASTRUKTURA: citanje " & tblName & "." & colName & " je puklo"
    End If
End Function
