Attribute VB_Name = "modLicenseTests"
Option Explicit

' ============================================================
' modLicenseTests -- unit testovi za cistu logiku modLicense.
'
' Pokreni: Alt+F8 -> TestLicense_All  (rezultati u Immediate / Ctrl+G)
'
' Testira fuzzy-match logiku otiska (LicSplitParts / LicPartsMatch /
' LicNonEmptyParts) koja odlucuje da li je masina "ista" -- to je deo gde
' off-by-one greska direktno znaci lazni lockout ili rupu. Mrezni put i
' GAS bind se testiraju server-side: gas/Code.gs -> runLicenseSelfTest.
' ============================================================

Private mPass As Long
Private mFail As Long

' Gate: bez ovoga runner vidi suite kao "blind" -- proslo bez greske, sto NIJE
' isto sto i sve provere prosle. Konvencija: modTestBanka.ERR_BIT_SUITE_FAILED.
Private Const ERR_LICENSE_SUITE_FAILED As Long = vbObjectError + 2967

Public Sub TestLicense_All()
    mPass = 0: mFail = 0
    Debug.Print String(60, "-")
    Debug.Print "LICENSE TEST SUITE START  " & Now
    Debug.Print String(60, "-")

    TestLicense_SplitParts
    TestLicense_PartsMatch
    TestLicense_NonEmptyParts
    TestLicense_DeviceFingerprint

    Debug.Print String(60, "-")
    Debug.Print "LICENSE TEST SUITE END  PASS=" & mPass & "  FAIL=" & mFail & _
                IIf(mFail = 0, "  (SVE PROSLO)", "  (IMA PADOVA!)")
    Debug.Print String(60, "-")

    If mFail > 0 Then
        Err.Raise ERR_LICENSE_SUITE_FAILED, "modLicenseTests.TestLicense_All", _
            "TestLicense_All: " & CStr(mFail) & " provera palo (PASS=" & _
            CStr(mPass) & "). Detalji u Immediate prozoru."
    End If
End Sub

' LicSplitParts uvek mora da vrati >= 3 elementa (stiti od index out of range).
Public Sub TestLicense_SplitParts()
    Debug.Print vbCrLf & "[1] LicSplitParts -- padding na 3"
    Dim p() As String

    p = LicSplitParts("A|B|C")
    AssertEq "puno: p(0)", p(0), "A"
    AssertEq "puno: p(2)", p(2), "C"

    p = LicSplitParts("")
    AssertTrue "prazno: UBound>=2", (UBound(p) >= 2)
    AssertEq "prazno: p(0)", p(0), ""
    AssertEq "prazno: p(2)", p(2), ""

    p = LicSplitParts("X")
    AssertEq "jedan: p(0)", p(0), "X"
    AssertEq "jedan: p(1)", p(1), ""
    AssertEq "jedan: p(2)", p(2), ""
End Sub

' Srce fuzzy matcha: koliko se od 3 komponente poklapa.
Public Sub TestLicense_PartsMatch()
    Debug.Print vbCrLf & "[2] LicPartsMatch -- broj poklapanja"

    AssertEq "identicno = 3", LicPartsMatch("A|B|C", "A|B|C"), 3
    AssertEq "2 od 3 (drift) = 2", LicPartsMatch("A|B|C", "A|B|Z"), 2
    AssertEq "1 od 3 = 1", LicPartsMatch("A|B|C", "A|Y|Z"), 1
    AssertEq "0 od 3 = 0", LicPartsMatch("A|B|C", "X|Y|Z"), 0
    AssertEq "case-insensitive", LicPartsMatch("aaa|bbb|ccc", "AAA|BBB|CCC"), 3
    AssertEq "prazne se NE racunaju kao poklapanje", LicPartsMatch("||", "||"), 0
    AssertEq "prazna pozicija ne broji", LicPartsMatch("A||C", "A||C"), 2

    ' Granica odluke: prag je 2 (>=2 => ista masina; <2 => drugi racunar)
    AssertTrue "prag: 2/3 prolazi (>=2)", (LicPartsMatch("A|B|C", "A|B|Z") >= 2)
    AssertTrue "prag: 1/3 pada (<2)", (LicPartsMatch("A|B|C", "A|Y|Z") < 2)
End Sub

' Koliko ne-praznih komponenti imamo (otisak slabiji od 2 => preskoci proveru).
Public Sub TestLicense_NonEmptyParts()
    Debug.Print vbCrLf & "[3] LicNonEmptyParts -- ne-prazne komponente"

    AssertEq "sve tri", LicNonEmptyParts("A|B|C"), 3
    AssertEq "dve", LicNonEmptyParts("A||C"), 2
    AssertEq "jedna", LicNonEmptyParts("A||"), 1
    AssertEq "nijedna", LicNonEmptyParts("||"), 0
    AssertTrue "2 komponente je dovoljno (>=2)", (LicNonEmptyParts("A||C") >= 2)
End Sub

' Integraciona provera: na realnoj masini otisak treba da da bar 2 komponente.
Public Sub TestLicense_DeviceFingerprint()
    Debug.Print vbCrLf & "[4] Otisak ovog racunara (integracija)"
    Dim parts As String
    parts = modLicense.GetDeviceParts()
    Debug.Print "  Otisak: " & parts
    AssertTrue "ovaj racunar daje >=2 komponente", (LicNonEmptyParts(parts) >= 2)
End Sub

' ============================================================
' Mini assert helperi (Debug.Print bazirani, kao ostali *Tests moduli)
' ============================================================

Private Sub AssertEq(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant)
    If actual = expected Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & label & "  (= " & expected & ")"
    Else
        mFail = mFail + 1
        Debug.Print "  FAIL  " & label & "  ocekivano=" & expected & "  dobijeno=" & actual
    End If
End Sub

Private Sub AssertTrue(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & label
    Else
        mFail = mFail + 1
        Debug.Print "  FAIL  " & label & "  (ocekivano True)"
    End If
End Sub
