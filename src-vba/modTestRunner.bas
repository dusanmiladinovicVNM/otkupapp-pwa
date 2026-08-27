Attribute VB_Name = "modTestRunner"
Option Explicit

' ============================================================
' modTestRunner -- jedan lifecycle i jedan izvestaj za sve test suite-ove.
'
' Do sada je svaki modul nosio svoj mPass/mFail, svoj Chk/AssertTrue, svoj
' ReportResults i svoj ERR_*_SUITE_FAILED offset. Pet kopija istog posla, sa pet
' malih razlika -- i sa jednom velikom posledicom: broj provera je zivio samo u
' Immediate prozoru, koji runner ne moze da procita. Zato pad pokrivenosti sa
' 189 na 120 provera nije bio vidljiv nigde.
'
' Ovde su tri stvari:
'   1) brojac provera koji zovu assert-i iz modTestAssert
'   2) lifecycle: BeginSuite -> BeginTest/EndTest -> EndSuite, sa eksplicitnim
'      NOT_RUN umesto tihog Exit Sub
'   3) masinski citljiv izvestaj: jedan red po suite-u u suite_results.txt PORED
'      SVESKE, koji cita tools/run_vba.py i poredi sa min_asserts iz
'      tests/suite_manifest.json
'
' MIGRACIJA JE POSTEPENA. Suite koja jos nije prebacena moze da dobije zube
' jednim pozivom na kraju svog postojeceg izvestaja:
'
'     TR_Report "RunPaleteTestSuite", mPass, mFail
'
' Time zadrzava svoj harness, a broj provera postaje vidljiv runneru. Pun
' lifecycle (TR_BeginSuite/TR_EndSuite + modTestAssert) je sledeci korak, jedan
' modul po jedan -- primer je modLicenseTests.
'
' ZASTO SE PISE PORED SVESKE: isto mesto gde modTest.RunAllTests pise
' last_run.txt. run_vba.py radi nad temp kopijom fixture-a, pa je
' ThisWorkbook.Path taj temp folder i izvestaj ne prlja repo.
'
' ZASTO SE DOPUNJUJE (append) A NE PISE ODJEDNOM: run koji pukne na sedmoj suite
' i dalje nosi brojeve prvih sest. Prvi TR_ poziv u procesu brise stari fajl
' (v. mFileReset), pa se izvestaji dva razlicita run-a ne mesaju.
' ============================================================

Private Const RESULT_FILE As String = "suite_results.txt"
Private Const ERR_TR_SUITE_FAILED As Long = vbObjectError + 2970
Private Const ERR_TR_NOT_RUN As Long = vbObjectError + 2971

Private mSuiteId As String
Private mTestName As String
Private mPass As Long
Private mFail As Long
Private mFails As String
Private mReport As String
Private mStarted As Double
Private mFileReset As Boolean

' ============================================================
' Lifecycle
' ============================================================

Public Sub TR_BeginSuite(ByVal suiteId As String)
    mSuiteId = suiteId
    mTestName = ""
    mPass = 0
    mFail = 0
    mFails = ""
    mReport = ""
    mStarted = Timer
    Debug.Print String(60, "-")
    Debug.Print "SUITE START  " & suiteId & "  " & Now
End Sub

Public Sub TR_BeginTest(ByVal testName As String)
    mTestName = testName
End Sub

Public Sub TR_EndTest()
    mTestName = ""
End Sub

' Zatvara suite: upisuje red u izvestaj, ispisuje rezime i PODIZE gresku ako je
' ijedna provera pala. Podizanje je ono sto suite cini "gate"-om -- bez njega
' runner vidi samo "proslo bez greske", sto NIJE isto sto i "sve provere prosle".
Public Sub TR_EndSuite()
    Dim res As clsTestResult
    Set res = NewResult(mSuiteId, mPass, mFail, False, "")

    Debug.Print "SUITE END    " & mSuiteId & "  PASS=" & mPass & "  FAIL=" & mFail & _
                IIf(mFail = 0, "  (SVE PROSLO)", "  (IMA PADOVA)")
    Debug.Print String(60, "-")
    WriteResult res
    WriteRunFile mSuiteId, mPass + mFail, mFail, mFails

    If mFail > 0 Then
        Err.Raise ERR_TR_SUITE_FAILED, "modTestRunner." & mSuiteId, _
            mSuiteId & ": " & CStr(mFail) & " provera palo (PASS=" & CStr(mPass) & ")." & _
            vbCrLf & mFails
    End If
End Sub

' Suite koja se NIJE pokrenula. Tih `Exit Sub` je runner video kao OK -- cetiri
' takve putanje su bile zive u ovom projektu (paletiranje iskljuceno, zatecen
' TST- ostatak, operater odustao, dev-guard odbijen). "Nije pokrenuto" nije
' "proslo", pa ovo i upisuje NOT_RUN u izvestaj i podize gresku.
Public Sub TR_NotRun(ByVal suiteId As String, ByVal reason As String)
    Dim res As clsTestResult
    Set res = NewResult(suiteId, 0, 0, True, reason)
    Debug.Print "SUITE NOT_RUN  " & suiteId & "  " & reason
    WriteResult res
    Err.Raise ERR_TR_NOT_RUN, "modTestRunner." & suiteId, _
        "suite NIJE pokrenut: " & reason
End Sub

' Suite koja jos nije prebacena na ovaj lifecycle, ali hoce da joj se broj
' provera vidi. Jedan poziv na kraju postojeceg ReportResults.
Public Sub TR_Report(ByVal suiteId As String, ByVal passCount As Long, _
                     ByVal failCount As Long)
    Dim res As clsTestResult
    Set res = NewResult(suiteId, passCount, failCount, False, "")
    WriteResult res
End Sub

' ============================================================
' Brojac -- ovde stizu svi assert-i iz modTestAssert
' ============================================================

Public Sub TR_Pass(ByVal label As String)
    mPass = mPass + 1
    mReport = mReport & "OK    " & Prefixed(label) & vbLf
End Sub

Public Sub TR_Fail(ByVal label As String)
    mFail = mFail + 1
    mReport = mReport & "PAO   " & Prefixed(label) & vbLf

    ' Oblik "<ime testa> -- <tvrdnja>" nije kozmetika: tools/dokaz.py iz njega
    ' cita KOJA je tvrdnja pala, da bi razlikovao "pao je taj test" od "pala je
    ' bas ta tvrdnja". Isti oblik pise i modTest u last_run.txt.
    If Len(mFails) > 0 Then mFails = mFails & vbCrLf
    mFails = mFails & FailLine(label)

    Debug.Print "  FAIL  " & Prefixed(label)
End Sub

' Ime u redu je IME SUITE, ne unutrasnjeg testa: katalog sabotaza
' (tools/sabotaza.py) framework-sabotaze imenuje suite-om, pa se to mora
' poklopiti sa onim sto tools/dokaz.py procita. Ime unutrasnjeg testa nije
' izgubljeno -- stoji na pocetku poruke, kroz Prefixed.
Private Function FailLine(ByVal label As String) As String
    FailLine = "FAIL " & mSuiteId & " -- " & Prefixed(label)
End Function

Public Function TR_CurrentSuite() As String
    TR_CurrentSuite = mSuiteId
End Function

Public Function TR_PassCount() As Long
    TR_PassCount = mPass
End Function

Public Function TR_FailCount() As Long
    TR_FailCount = mFail
End Function

Public Function TR_ReportText() As String
    TR_ReportText = mReport
End Function

' ============================================================
' Pomocno
' ============================================================

Private Function Prefixed(ByVal label As String) As String
    If Len(mTestName) > 0 Then
        Prefixed = mTestName & ": " & label
    Else
        Prefixed = label
    End If
End Function

Private Function NewResult(ByVal suiteId As String, ByVal passCount As Long, _
                           ByVal failCount As Long, ByVal notRun As Boolean, _
                           ByVal message As String) As clsTestResult
    Dim r As clsTestResult
    Set r = New clsTestResult
    r.SuiteId = suiteId
    r.PassCount = passCount
    r.FailCount = failCount
    r.NotRun = notRun
    r.Message = message
    If mStarted > 0 Then r.Seconds = Timer - mStarted
    Set NewResult = r
End Function

' Ime palog testa mora da prezivi COM granicu. `xl.Run` vraca opis greske koji
' se preko COM-a cesto svede na golo "Exception occurred", pa od celog pada
' ostane samo "nesto je palo". Zato suite pise i fajl PORED SVESKE, u formatu
' koji modTest vec koristi (`TESTS=n FAIL=m`, pa red po padu) -- isti razlog
' zbog kog ga pise i modTestBanka, i isti citaoci: run_vba.py (result_file u
' manifestu) i tools/dokaz.py.
Private Sub WriteRunFile(ByVal suiteId As String, ByVal total As Long, _
                         ByVal failed As Long, ByVal fails As String)
    On Error Resume Next

    Dim path As String
    path = ThisWorkbook.path & Application.PathSeparator & _
           "last_run_" & suiteId & ".txt"

    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, "TESTS=" & CStr(total) & " FAIL=" & CStr(failed)
    If Len(fails) > 0 Then Print #fnum, fails
    Close #fnum

    On Error GoTo 0
End Sub

' Upis ne sme da obori suite: izvestaj je dijagnostika, a ne provera. Ako fajl
' ne moze da se otvori (read-only folder), suite i dalje daje svoj verdikt kroz
' Err.Raise -- samo broj provera nece biti vidljiv runneru, sto se u ispisu
' runnera vidi kao "nema broja".
Private Sub WriteResult(ByVal res As clsTestResult)
    On Error Resume Next

    Dim path As String
    path = ThisWorkbook.path & Application.PathSeparator & RESULT_FILE

    Dim fnum As Integer
    fnum = FreeFile

    If Not mFileReset Then
        ' Prvi upis u ovom procesu brise izvestaj prethodnog run-a. Bez ovoga bi
        ' se dva run-a nad istom temp kopijom (--repeat, --keep) spojila i broj
        ' provera bi izgledao dvostruko veci nego sto jeste.
        Open path For Output As #fnum
        mFileReset = True
    Else
        Open path For Append As #fnum
    End If

    Print #fnum, res.ToLine()
    Close #fnum

    On Error GoTo 0
End Sub
