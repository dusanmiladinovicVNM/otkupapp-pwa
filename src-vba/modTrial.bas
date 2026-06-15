Option Explicit

' ============================================================
' modTrial — vremenski ogranicen rad (trial / demo lock)
'
' Cita SISTEMSKI datum i BLOKIRA pokretanje posle TRIAL_DAYS dana od
' zadatog pocetnog datuma (TRIAL_START_*). Dodatno detektuje vracanje
' sistemskog sata unazad (anti-rollback), inace se lock trivijalno zaobidje.
'
' >>> PODESI DOLE: TRIAL_START_* (pocetni "zadati datum") i TRIAL_DAYS. <<<
' Iskljucivanje cele provere: TRIAL_ENABLED = False.
'
' Poziva se na vrhu modMain.StartApp:  If Not TrialGateOrQuit() Then Exit Sub
'
' NAPOMENA: deterrent. Ko otvori VBE moze da promeni datum ili iskljuci kod.
' Za jaci nivo videti modLicense (potpisana licenca). Takodje: legitimno
' POGRESNO podesen sistemski sat (npr. prazna baterija -> 2000) moze da
' okine anti-rollback; za 10-dnevni trial to je prihvatljiv kompromis.
' ============================================================

Private Const TRIAL_ENABLED As Boolean = True

' >>> POCETNI ("zadati") DATUM — godina, mesec, dan <<<
Private Const TRIAL_START_Y As Integer = 2026
Private Const TRIAL_START_M As Integer = 6
Private Const TRIAL_START_D As Integer = 15

' Broj dana vazenja od pocetnog datuma.
Private Const TRIAL_DAYS As Long = 10

' Config kljuc (tblSEFConfig) za high-water-mark = najkasniji vidjeni datum.
Private Const TRIAL_KEY_HWM As String = "TRIAL_HWM"

' ============================================================
' Vraca True ako sme da nastavi; inace blokira i zatvara svesku.
' ============================================================
Public Function TrialGateOrQuit() As Boolean
    On Error GoTo EH

    If Not TRIAL_ENABLED Then
        TrialGateOrQuit = True
        Exit Function
    End If

    Dim deadline As Date, today As Date
    deadline = DateSerial(TRIAL_START_Y, TRIAL_START_M, TRIAL_START_D) + TRIAL_DAYS
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

' ============================================================
Private Sub TrialBlock(ByVal reason As String, ByVal deadline As Date)
    On Error Resume Next
    Application.Visible = True
    MsgBox reason & vbCrLf & _
           "Vazenje do: " & Format$(deadline, "dd.mm.yyyy") & "." & vbCrLf & vbCrLf & _
           "Kontaktirajte dobavljaca za nastavak rada.", vbCritical, APP_NAME

    ' Zatvori SAMO ovu svesku (bez Application.Quit). Saved=True spreci prompt.
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
End Sub
