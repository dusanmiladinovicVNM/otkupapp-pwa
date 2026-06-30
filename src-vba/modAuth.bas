Attribute VB_Name = "modAuth"
Option Explicit

' ============================================================
' modAuth - per-user prijava + prava po oblasti (Faza 1)
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
Private gCurrentUserIme As String
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
' Prikazi login formu (frmLogin) i vrati True ako je prijava uspela.
' Validaciju radi frmLogin.btnOK_Click preko modAuth.ValidateLogin; forma
' postavi Me.LoginOK. Poziva se iz modMain.StartApp (posle AccessGateOrQuit,
' pre frmSplash). frmLogin se pravi u dizajneru (kontrole: txtUser, txtPin
' [PasswordChar], lblErr, btnOK, btnCancel) i izvozi sa svojim .frx parom.
' ------------------------------------------------------------
Public Function Login() As Boolean
    On Error GoTo EH
    Logout

    frmLogin.LoginOK = False
    frmLogin.Show                 ' modal (default)
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

    If Not VerifyPin(SafeStr(storedPin), pin) Then
        AuditAuth "AUTH_LOGIN_FAIL", "WARN", u & " (pogresan PIN)"
        ValidateLogin = False
        Exit Function
    End If

    ' Faza 3: transparentna migracija legacy plaintext -> hash (kad je hash ukljucen).
    If PinHashEnabled() Then
        If LCase$(Left$(SafeStr(storedPin), 7)) <> "sha256$" Then
            MigratePinToHash u, pin
        End If
    End If

    gCurrentUser = u
    gCurrentUserUloga = SafeStr(LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_ULOGA))
    gCurrentUserIme = SafeStr(LookupValue(TBL_KORISNICI, COL_KOR_USERNAME, u, COL_KOR_IME))
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

' Ime i prezime prijavljenog app-korisnika (za prikaz u top baru / userstamp).
' Prazno ako AUTH nije ukljucen ili niko nije prijavljen -> pozivalac fallback-uje
' na Windows nalog.
Public Function GetCurrentUserIme() As String
    GetCurrentUserIme = gCurrentUserIme
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

' Da li trenutni kontekst sme administraciju korisnika:
'  - AUTH iskljucen -> da (priprema korisnika pre ukljucenja prijave)
'  - AUTH ukljucen   -> samo Admin
Public Function MozeAdministraciju() As Boolean
    MozeAdministraciju = (Not AuthEnabled()) Or CurrentUserIsAdmin()
End Function

' Jedinstven izvor liste oblasti (= nazivi kolona prava u tblKorisnici).
Public Function OblastiList() As Variant
    OblastiList = Array(OBL_OTKUP, OBL_DOKUMENTA, OBL_AGROHEMIJA, OBL_IZVESTAJI, _
                        OBL_FAKTURISANJE, OBL_BANKA, OBL_MARZA, OBL_SLEDLJIVOST, OBL_MATICNI, _
                        OBL_PALETE, OBL_OTVORI_EXCEL, OBL_SYNC_PWA)
End Function

Public Sub Logout()
    gLoggedIn = False
    gCurrentUser = vbNullString
    gCurrentUserUloga = vbNullString
    gCurrentUserIme = vbNullString
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
        Case "frmpalete":              OblastZaFormu = OBL_PALETE
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
    MsgBox Poruka("AUTH_MSG_PRIJAVA_NEUSPESNA"), vbCritical, APP_NAME
    ThisWorkbook.Close SaveChanges:=False
End Sub

' ============================================================
' Faza 3 - PIN hashing (opt-in: PIN_HASH_ENABLED, default NO).
' Format sacuvanog PIN-a: "sha256$<salt>$<hexHash>" (hash) ili plaintext (legacy).
' SHA-256 preko .NET (System.Security.Cryptography). PRE ukljucenja: Alt+F8 ->
' TestPinHash (mora PASS). Migracija plaintext->hash je transparentna pri prijavi.
' Ako SHA nije dostupan -> sve pada nazad na plaintext (bez lockout-a).
' ============================================================
Public Function PinHashEnabled() As Boolean
    On Error GoTo EH
    ' Podrazumevano UKLJUCENO (opt-out): samo eksplicitno NO/NE/FALSE/0 gasi hash.
    ' Bezbedno: ako SHA (.NET) nije dostupan, PreparePin/VerifyPin padaju nazad na
    ' plaintext (bez rizika od lockout-a), a postojeci plaintext PIN-ovi i dalje rade.
    Dim v As String
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_PIN_HASH_ENABLED)))
    PinHashEnabled = Not (v = "NO" Or v = "NE" Or v = "FALSE" Or v = "0")
    Exit Function
EH:
    PinHashEnabled = True
End Function

Public Function Sha256Hex(ByVal text As String) As String
    On Error GoTo EH
    Dim enc As Object, sha As Object
    Dim bytes() As Byte, hash() As Byte
    Set enc = CreateObject("System.Text.UTF8Encoding")
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    bytes = enc.GetBytes_4(text)
    hash = sha.ComputeHash_2((bytes))

    Dim i As Long, s As String
    For i = LBound(hash) To UBound(hash)
        s = s & Right$("0" & Hex$(hash(i) And &HFF), 2)
    Next i
    Sha256Hex = LCase$(s)
    Exit Function
EH:
    Sha256Hex = vbNullString      ' prazno = SHA nedostupan -> pozivalac fallback-uje na plaintext
End Function

Public Function NewSalt() As String
    Static seeded As Boolean
    If Not seeded Then Randomize: seeded = True
    Dim s As String, i As Long
    For i = 1 To 16
        s = s & Mid$("0123456789abcdef", Int(Rnd() * 16) + 1, 1)
    Next i
    NewSalt = s
End Function

Public Function HashPin(ByVal pin As String, ByVal salt As String) As String
    HashPin = Sha256Hex(salt & "|" & Trim$(pin))
End Function

' Vrednost za upis u PIN kolonu: hash ako je ukljucen (i SHA radi), inace plaintext.
Public Function PreparePin(ByVal pin As String) As String
    If PinHashEnabled() Then
        Dim salt As String, h As String
        salt = NewSalt()
        h = HashPin(pin, salt)
        If Len(h) > 0 Then
            PreparePin = "sha256$" & salt & "$" & h
            Exit Function
        End If
    End If
    PreparePin = Trim$(pin)
End Function

' Provera unetog PIN-a protiv sacuvanog (auto-detekcija: hash vs plaintext).
Public Function VerifyPin(ByVal stored As String, ByVal inputPin As String) As Boolean
    Dim s As String
    s = Trim$(stored)

    If LCase$(Left$(s, 7)) = "sha256$" Then
        Dim parts() As String
        parts = Split(s, "$")
        If UBound(parts) < 2 Then Exit Function
        Dim h As String
        h = HashPin(inputPin, parts(1))
        VerifyPin = (Len(h) > 0 And StrComp(h, parts(2), vbTextCompare) = 0)
    Else
        VerifyPin = (StrComp(s, Trim$(inputPin), vbBinaryCompare) = 0)
    End If
End Function

' Alt+F8 self-test: SHA-256("abc") mora biti poznati vektor.
Public Sub TestPinHash()
    Dim got As String
    got = Sha256Hex("abc")
    If StrComp(got, "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", vbTextCompare) = 0 Then
        MsgBox Poruka("AUTH_MSG_PINHASH_RADI"), vbInformation, APP_NAME
    Else
        MsgBox Poruka("AUTH_MSG_PINHASH_NE_RADI") & got & "'", vbExclamation, APP_NAME
    End If
End Sub

Private Function FindUserRow(ByVal username As String) As Long
    On Error GoTo EH
    Dim rws As Collection
    Set rws = FindRows(TBL_KORISNICI, COL_KOR_USERNAME, Trim$(username))
    If Not rws Is Nothing Then
        If rws.count > 0 Then FindUserRow = rws(1)
    End If
    Exit Function
EH:
    FindUserRow = 0
End Function

Private Sub MigratePinToHash(ByVal username As String, ByVal pin As String)
    On Error Resume Next
    Dim r As Long
    r = FindUserRow(username)
    If r > 0 Then UpdateCell TBL_KORISNICI, r, COL_KOR_PIN, PreparePin(pin)
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
