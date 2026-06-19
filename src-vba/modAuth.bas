Attribute VB_Name = "modAuth"
Option Explicit

' ============================================================
' modAuth – per-user prijava + prava po oblasti (Faza 1)
'
' Cilj: admin sa svim pravima + korisnici kojima admin odobrava
' pristup PO OBLASTI (Otkup, Dokumenta, Agrohemija, ...). Model A:
' jedan red u tblKorisnici = jedan korisnik, kolona po oblasti = "DA"/"NE".
'
' REUSE (bez novog sloja):
'   - modConfig:     GetConfigValue (tblSEFConfig), TBL_KORISNICI, COL_KOR_*, OBL_*, ULOGA_*
'   - modDataAccess: LookupValue
'   - Monitor_Event: audit (kao u Workbook_Open / modMain)
'   - LogErr:        logovanje gresaka
'
' OPT-IN: CFG_KEY_AUTH_ENABLED u tblSEFConfig. Dok nije "YES" -> sve radi
' kao pre (bez prijave, bez restrikcija) -> nema rizika lockout-a.
'
' Stanje (mirror modStanicaLock gActiveStanica obrasca): modul-level globali.
' ============================================================

Private gCurrentUser As String
Private gCurrentUserUloga As String
Private gLoggedIn As Boolean

' ------------------------------------------------------------
' Da li je prijava ukljucena (opt-in). Prazno/missing = NE.
' ------------------------------------------------------------
Public Function AuthEnabled() As Boolean
    On Error GoTo EH
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_AUTH_ENABLED)))
    AuthEnabled = (v = "YES" Or v = "DA" Or v = "TRUE" Or v = "1")
    Exit Function
EH:
    AuthEnabled = False
End Function

' ------------------------------------------------------------
' Prikazi login formu i vrati True ako je prijava uspela.
' Poziva se iz modMain.StartApp (posle AccessGateOrQuit, pre frmSplash).
' ------------------------------------------------------------
Public Function Login() As Boolean
    On Error GoTo EH
    Logout

    frmLogin.LoginOK = False
    frmLogin.Show                       ' modal (default)
    Login = frmLogin.LoginOK
    Unload frmLogin
    Exit Function
EH:
    LogErr "modAuth.Login"
    Login = False
End Function

' ------------------------------------------------------------
' Validacija kredencijala (poziva frmLogin posle OK).
' Na uspeh postavlja globalno stanje (gCurrentUser/uloga) + audit.
' ------------------------------------------------------------
Public Function ValidateLogin(ByVal username As String, ByVal pin As String) As Boolean
    On Error GoTo EH

    Dim u As String
    u = Trim$(username)
    If Len(u) = 0 Or Len(Trim$(pin)) = 0 Then
        ValidateLogin = False
        Exit Function
    End If

    Dim storedPin As Variant
    storedPin = LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_PIN)
    If IsEmpty(storedPin) Then            ' nepoznat korisnik
        AuditAuth "AUTH_LOGIN_FAIL", "WARN", u & " (nepoznat)"
        ValidateLogin = False
        Exit Function
    End If

    ' Aktivan? Prazno = aktivan; samo "NE" blokira (drift-safe).
    Dim aktiv As String
    aktiv = UCase$(SafeStr(LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_AKTIVAN)))
    If aktiv = "NE" Then
        AuditAuth "AUTH_LOGIN_FAIL", "WARN", u & " (deaktiviran)"
        ValidateLogin = False
        Exit Function
    End If

    If StrComp(SafeStr(storedPin), Trim$(pin), vbBinaryCompare) <> 0 Then
        AuditAuth "AUTH_LOGIN_FAIL", "WARN", u & " (pogresan PIN)"
        ValidateLogin = False
        Exit Function
    End If

    gCurrentUser = u
    gCurrentUserUloga = SafeStr(LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_ULOGA))
    gLoggedIn = True
    AuditAuth "AUTH_LOGIN", "INFO", u & " (" & gCurrentUserUloga & ")"
    ValidateLogin = True
    Exit Function
EH:
    LogErr "modAuth.ValidateLogin"
    ValidateLogin = False
End Function

Public Function GetCurrentUser() As String
    GetCurrentUser = gCurrentUser
End Function

Public Function CurrentUserIsAdmin() As Boolean
    CurrentUserIsAdmin = (StrComp(Trim$(gCurrentUserUloga), ULOGA_ADMIN, vbTextCompare) = 0)
End Function

' ------------------------------------------------------------
' Glavna provera prava po oblasti. Pravila:
'   - AUTH iskljucen   -> True (sve dozvoljeno, kao pre)
'   - nije prijavljen  -> False
'   - Admin            -> True (bypass)
'   - prazna oblast    -> True (nemapirana/buduca sekcija se ne blokira)
'   - inace            -> celija oblasti = "DA"
' ------------------------------------------------------------
Public Function KorisnikImaPravo(ByVal oblast As String) As Boolean
    On Error GoTo EH

    If Not AuthEnabled() Then
        KorisnikImaPravo = True
        Exit Function
    End If
    If Not gLoggedIn Then
        KorisnikImaPravo = False
        Exit Function
    End If
    If CurrentUserIsAdmin() Then
        KorisnikImaPravo = True
        Exit Function
    End If
    If Len(Trim$(oblast)) = 0 Then
        KorisnikImaPravo = True
        Exit Function
    End If

    Dim v As String
    v = UCase$(SafeStr(LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, gCurrentUser, oblast)))
    KorisnikImaPravo = (v = "DA" Or v = "YES" Or v = "TRUE" Or v = "1" Or v = "X")
    Exit Function
EH:
    LogErr "modAuth.KorisnikImaPravo"
    KorisnikImaPravo = False
End Function

Public Sub Logout()
    gLoggedIn = False
    gCurrentUser = vbNullString
    gCurrentUserUloga = vbNullString
End Sub

' ------------------------------------------------------------
' Mapiranje forma (po imenu) -> oblast. Koristi guard u
' frmOtkupAPP.OpenContentForm. Nepoznata forma -> "" (ne blokira).
' ------------------------------------------------------------
Public Function OblastZaFormu(ByVal formName As String) As String
    Select Case LCase$(Trim$(formName))
        Case "frmotkup":               OblastZaFormu = OBL_OTKUP
        Case "frmdokumenta":           OblastZaFormu = OBL_DOKUMENTA
        Case "frmagrohemija":          OblastZaFormu = OBL_AGROHEMIJA
        Case "frmizvestaj":            OblastZaFormu = OBL_IZVESTAJI
        Case "frmfakturisanje":        OblastZaFormu = OBL_FAKTURISANJE
        Case "frmbankaimport", "frmbankaexportpregled": OblastZaFormu = OBL_BANKA
        Case "frmmarza":               OblastZaFormu = OBL_MARZA
        Case "frmsledljivost":         OblastZaFormu = OBL_SLEDLJIVOST
        Case "frmmaticnipodaci":       OblastZaFormu = OBL_MATICNI
        Case Else:                     OblastZaFormu = vbNullString
    End Select
End Function

' ------------------------------------------------------------
' Zakazano gasenje posle neuspele prijave (mirror license gate:
' zatvaranje se ne radi unutar Workbook_Open lanca, vec na sledeci tick).
' Poziva se preko Application.OnTime iz modMain.StartApp.
' ------------------------------------------------------------
Public Sub QuitAfterFailedLogin()
    On Error Resume Next
    MsgBox "Prijava neuspešna. Aplikacija se zatvara.", vbCritical, APP_NAME
    ThisWorkbook.Close SaveChanges:=False
End Sub

' ============================================================
' Interni helperi
' ============================================================
Private Function SafeStr(ByVal v As Variant) As String
    On Error GoTo EH
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        SafeStr = vbNullString
    Else
        SafeStr = Trim$(CStr(v))
    End If
    Exit Function
EH:
    SafeStr = vbNullString
End Function

Private Sub AuditAuth(ByVal evt As String, ByVal sev As String, ByVal msg As String)
    On Error Resume Next
    Monitor_Event _
        eventType:=evt, _
        severity:=sev, _
        message:=msg, _
        userId:=gCurrentUser, _
        moduleName:="modAuth", _
        procedureName:="ValidateLogin", _
        entityType:="Auth", _
        entityID:=gCurrentUser, _
        correlationId:="VBA-AUTH"
End Sub
