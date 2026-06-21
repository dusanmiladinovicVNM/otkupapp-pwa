Attribute VB_Name = "modTrial"
Option Explicit

' ============================================================
' modTrial — vremenski ogranicen rad (trial / demo lock)
'
' Cita SISTEMSKI datum i BLOKIRA pokretanje posle TRIAL_DAYS dana od
' zadatog pocetnog datuma (TRIAL_START_*). Dodatno detektuje vracanje
' sistemskog sata unazad (anti-rollback), inace se lock trivijalno zaobidje.
'
' Podesavanje je CONFIG-DRIVEN preko tblSEFConfig (bez recompile/re-sign):
'   TRIAL_ENABLED = YES/NO
'   TRIAL_START   = "yyyy-mm-dd"   (pocetni "zadati" datum)
'   TRIAL_DAYS    = broj dana
' Ako kljuc nije u Config-u, koristi se Const default ispod.
'
' Orkestrira ga modLicense.AccessGateOrQuit (pravilo "trial samo ako NIJE
' licenciran"); ne poziva se direktno iz StartApp.
'
' NAPOMENA: deterrent. Ko otvori VBE moze da promeni datum ili iskljuci kod.
' Za jaci nivo videti modLicense (per-uredjaj licenca). Takodje: legitimno
' POGRESNO podesen sistemski sat (npr. prazna baterija -> 2000) moze da
' okine anti-rollback; za kratak trial to je prihvatljiv kompromis.
' ============================================================

' --- Config kljucevi (tblSEFConfig) ---
Private Const CFG_TRIAL_ENABLED As String = "TRIAL_ENABLED"
Private Const CFG_TRIAL_START   As String = "TRIAL_START"   ' ISO "yyyy-mm-dd"
Private Const CFG_TRIAL_DAYS    As String = "TRIAL_DAYS"

' --- Const DEFAULT-i (fallback kad Config kljuc ne postoji) ---
' Default OFF: trial je dormant dok ga ne upalis (Config TRIAL_ENABLED=YES).
' Dok je OFF, nelicencirana masina pada na license gate (ne propusta se).
Private Const TRIAL_ENABLED_DEFAULT As Boolean = False
Private Const TRIAL_START_Y As Integer = 2026
Private Const TRIAL_START_M As Integer = 6
Private Const TRIAL_START_D As Integer = 18
Private Const TRIAL_DAYS_DEFAULT As Long = 10

' Config kljuc (tblSEFConfig) za high-water-mark = najkasniji vidjeni datum.
Private Const TRIAL_KEY_HWM As String = "TRIAL_HWM"

' ============================================================
' Vraca True ako sme da nastavi; inace blokira i zatvara svesku.
' ============================================================
Public Function TrialGateOrQuit() As Boolean
    On Error GoTo EH

    If Not TrialEnabled() Then
        TrialGateOrQuit = True
        Exit Function
    End If

    Dim deadline As Date, today As Date
    deadline = TrialStartDate() + TrialDays()
    today = Date

    ' 1) PRIMARNO: istekao rok? Cista matematika datuma, bez zavisnosti ->
    '    ovaj deo uvek radi cak i ako config citanje zakaze.
    If today > deadline Then
        TrialBlock "Probni period je istekao.", deadline
        TrialGateOrQuit = False
        Exit Function
    End If

    ' 2) ANTI-ROLLBACK (best-effort): da li je sat vracen ispod ranije
    '    vidjenog datuma? Ako jeste -> blokiraj.
    Dim hwm As String
    On Error Resume Next
    hwm = Trim$(GetConfigValue(TRIAL_KEY_HWM))
    On Error GoTo EH

    If Len(hwm) > 0 Then
        If IsDate(hwm) Then
            If today < CDate(hwm) Then
                TrialBlock "Sistemski datum je vracen unazad. Rad je blokiran.", deadline
                TrialGateOrQuit = False
                Exit Function
            End If
        End If
    End If

    ' Azuriraj high-water-mark na najkasniji vidjeni datum.
    If Len(hwm) = 0 Or (IsDate(hwm) And today > CDate(hwm)) Then
        On Error Resume Next
        SetConfigValue TRIAL_KEY_HWM, Format$(today, "yyyy-mm-dd")
        On Error GoTo EH
    End If

    TrialGateOrQuit = True
    Exit Function

EH:
    ' Rok je vec proveren gore bez zavisnosti; ovde fail-OPEN da eventualni
    ' bag u anti-rollback delu ne zakljuca korisnika koji je UNUTAR roka.
    LogErr "modTrial.TrialGateOrQuit"
    TrialGateOrQuit = True
End Function

' Da li je trial ukljucen (Config TRIAL_ENABLED, fallback Const). Koristi ga
' modLicense.AccessGateOrQuit da odluci da li nelicencirana masina dobija trial
' ili pada na license gate.
Public Function TrialEnabled() As Boolean
    Dim v As String
    On Error Resume Next
    v = UCase$(Trim$(GetConfigValue(CFG_TRIAL_ENABLED)))
    On Error GoTo 0
    Select Case v
        Case "YES", "TRUE", "1", "ON", "ENABLED": TrialEnabled = True
        Case "NO", "FALSE", "0", "OFF", "DISABLED": TrialEnabled = False
        Case Else: TrialEnabled = TRIAL_ENABLED_DEFAULT   ' nema kljuca -> Const default
    End Select
End Function

' Da li je trial TRENUTNO aktivan (ukljucen, u roku, sat nije vracen unazad).
' CISTA provera bez side-efekata (ne blokira, ne pise HWM) — ogledalo allow-uslova
' iz TrialGateOrQuit. Koristi je modLicense.AccessGateOrQuit da bira izmedju
' trial-a i inline unosa licence.
Public Function TrialActive() As Boolean
    On Error Resume Next
    If Not TrialEnabled() Then Exit Function

    Dim today As Date: today = Date
    If today > TrialStartDate() + TrialDays() Then Exit Function     ' istekao

    Dim hwm As String: hwm = Trim$(GetConfigValue(TRIAL_KEY_HWM))
    If Len(hwm) > 0 Then
        If IsDate(hwm) Then
            If today < CDate(hwm) Then Exit Function                  ' sat vracen unazad
        End If
    End If

    TrialActive = True
End Function

' Efektivni pocetni datum: Config TRIAL_START (ISO), fallback na Const datum.
Private Function TrialStartDate() As Date
    Dim s As String
    On Error Resume Next
    s = Trim$(GetConfigValue(CFG_TRIAL_START))
    On Error GoTo 0
    If Len(s) > 0 Then
        s = Replace(s, "T", " ")          ' tolerisi ISO 'T' (VBA CDate ne razume)
        If IsDate(s) Then
            TrialStartDate = CDate(s)
            Exit Function
        End If
    End If
    TrialStartDate = DateSerial(TRIAL_START_Y, TRIAL_START_M, TRIAL_START_D)
End Function

' Efektivni broj dana: Config TRIAL_DAYS (>0), fallback na Const.
Private Function TrialDays() As Long
    Dim n As Long
    On Error Resume Next
    n = CLng(val(Trim$(GetConfigValue(CFG_TRIAL_DAYS))))
    On Error GoTo 0
    If n > 0 Then TrialDays = n Else TrialDays = TRIAL_DAYS_DEFAULT
End Function

' ============================================================
Private Sub TrialBlock(ByVal reason As String, ByVal deadline As Date)
    On Error Resume Next
    Application.Visible = True
    MsgBox reason & vbCrLf & _
           "Vazenje do: " & Format$(deadline, "dd.mm.yyyy") & "." & vbCrLf & vbCrLf & _
           "Kontaktirajte dobavljaca za nastavak rada.", vbCritical, APP_NAME

    ' Pouzdano zatvaranje preko deljenog mehanizma (modLicense). NE zatvarati
    ' sinhrono iz Workbook_Open toka — Excel Close tu odlaze/ignorise.
    DenyAccessAndScheduleClose
End Sub
