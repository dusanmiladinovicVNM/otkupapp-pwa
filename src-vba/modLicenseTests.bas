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
'
' KANARINAC MIGRACIJE. Ovo je prvi suite prebacen sa sopstvenog mini-frameworka
' (mPass/mFail + privatni AssertEq/AssertTrue + rucni Err.Raise) na zajednicki:
' modTestRunner za lifecycle i brojanje, modTestAssert za provere,
' clsTestContext za vracanje stanja Excela. Izabran je prvi jer je najmanji i
' nema tabele -- ako se nesto u frameworku ponasa drugacije nego sto je
' zamisljeno, ovde se vidi bez rizika po podatke.
'
' Sta je time nestalo iz ovog modula:
'   - Private mPass / mFail                  -> modTestRunner
'   - Private AssertEq / AssertTrue          -> modTestAssert (AssertEqual/AssertTrue)
'   - Private Const ERR_LICENSE_SUITE_FAILED -> TR_EndSuite podize jedinstvenu gresku
'   - rucni Debug.Print rezime               -> TR_BeginSuite / TR_EndSuite
' Ono sto je DOBIO: broj provera zavrsi u suite_results.txt, pa `min_asserts: 23`
' iz tests/suite_manifest.json postaje kapija, a ne samo zapisano ocekivanje.
'
' PAZNJA na redosled argumenata: stari privatni AssertEq je bio
' (label, actual, expected), a zajednicki AssertEqual je (actual, expected, label)
' -- isti oblik kao modTest.AssertEq, koji je stariji i vec javan. Svi pozivi
' ispod su preuredjeni; ovo je jedina rucna izmena u telu testova.
' ============================================================

Public Sub TestLicense_All()
    Dim ctx As clsTestContext

    On Error GoTo EH

    ' Snapshot stanja Excela se uzima u Class_Initialize, a vraca u
    ' Class_Terminate -- i na urednom izlazu i posle greske. Nema rucnog cleanup-a.
    Set ctx = New clsTestContext
    ctx.Quiet

    TR_BeginSuite "TestLicense_All"

    TestLicense_SplitParts
    TestLicense_PartsMatch
    TestLicense_NonEmptyParts
    TestLicense_DeviceFingerprint

    On Error GoTo 0
    TR_EndSuite          ' upisuje izvestaj i PODIZE gresku ako je nesto palo
    Exit Sub

EH:
    ' Prekid nije "nije se desilo" nego pad: broji se kao pala provera, pa
    ' TR_EndSuite podize gresku sa punim spiskom.
    Dim errDesc As String
    errDesc = Err.description
    On Error Resume Next
    TR_Fail "SUITE prekinut greskom: " & errDesc
    On Error GoTo 0
    TR_EndSuite
End Sub

' LicSplitParts uvek mora da vrati >= 3 elementa (stiti od index out of range).
Public Sub TestLicense_SplitParts()
    TR_BeginTest "LicSplitParts"
    Dim p() As String

    p = LicSplitParts("A|B|C")
    AssertEqual p(0), "A", "puno: p(0)"
    AssertEqual p(2), "C", "puno: p(2)"

    p = LicSplitParts("")
    AssertTrue (UBound(p) >= 2), "prazno: UBound>=2"
    AssertEqual p(0), "", "prazno: p(0)"
    AssertEqual p(2), "", "prazno: p(2)"

    p = LicSplitParts("X")
    AssertEqual p(0), "X", "jedan: p(0)"
    AssertEqual p(1), "", "jedan: p(1)"
    AssertEqual p(2), "", "jedan: p(2)"
    TR_EndTest
End Sub

' Srce fuzzy matcha: koliko se od 3 komponente poklapa.
Public Sub TestLicense_PartsMatch()
    TR_BeginTest "LicPartsMatch"

    AssertEqual LicPartsMatch("A|B|C", "A|B|C"), 3, "identicno = 3"
    AssertEqual LicPartsMatch("A|B|C", "A|B|Z"), 2, "2 od 3 (drift) = 2"
    AssertEqual LicPartsMatch("A|B|C", "A|Y|Z"), 1, "1 od 3 = 1"
    AssertEqual LicPartsMatch("A|B|C", "X|Y|Z"), 0, "0 od 3 = 0"
    AssertEqual LicPartsMatch("aaa|bbb|ccc", "AAA|BBB|CCC"), 3, "case-insensitive"
    AssertEqual LicPartsMatch("||", "||"), 0, "prazne se NE racunaju kao poklapanje"
    AssertEqual LicPartsMatch("A||C", "A||C"), 2, "prazna pozicija ne broji"

    ' Granica odluke: prag je 2 (>=2 => ista masina; <2 => drugi racunar)
    AssertTrue (LicPartsMatch("A|B|C", "A|B|Z") >= 2), "prag: 2/3 prolazi (>=2)"
    AssertTrue (LicPartsMatch("A|B|C", "A|Y|Z") < 2), "prag: 1/3 pada (<2)"
    TR_EndTest
End Sub

' Koliko ne-praznih komponenti imamo (otisak slabiji od 2 => preskoci proveru).
Public Sub TestLicense_NonEmptyParts()
    TR_BeginTest "LicNonEmptyParts"

    AssertEqual LicNonEmptyParts("A|B|C"), 3, "sve tri"
    AssertEqual LicNonEmptyParts("A||C"), 2, "dve"
    AssertEqual LicNonEmptyParts("A||"), 1, "jedna"
    AssertEqual LicNonEmptyParts("||"), 0, "nijedna"
    AssertTrue (LicNonEmptyParts("A||C") >= 2), "2 komponente je dovoljno (>=2)"
    TR_EndTest
End Sub

' Integraciona provera: na realnoj masini otisak treba da da bar 2 komponente.
Public Sub TestLicense_DeviceFingerprint()
    TR_BeginTest "Otisak ovog racunara"
    Dim parts As String
    parts = modLicense.GetDeviceParts()
    Debug.Print "  Otisak: " & parts
    AssertTrue (LicNonEmptyParts(parts) >= 2), "ovaj racunar daje >=2 komponente"
    TR_EndTest
End Sub
