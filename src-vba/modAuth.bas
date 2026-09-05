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

' TEST: sledeci PrikaziPrijavu vraca False BEZ otvaranja forme. Sme samo da
' obori prijavu, nikad da je odobri -- forma za prijavu je modalna i u harnessu
' se ne moze odigrati, a bas neuspela prijava je ono sto se meri.
Private mPrijavaOdbijTest As Boolean
' Test seam: AUTH ukljucen BEZ diranja tblSEFConfig. Postoji zato sto se prava
' drugacije ne mogu izmeriti -- fixture mora da ostane sa AUTH=NE (inace bi
' svaka postojeca suite trazila prijavu), a bas grane "nema prava" su one koje
' su do sada bile nemerene. Upis u config bi menjao svesku i curio u sledeci
' test; promenljiva se vraca u istom testu.
' Tvrdo gejtovan: van test rezima ne radi nista (isti obrazac kao
' modScrDokumenti.Scr_OtpTestSet).
Private mAuthTestOn As Boolean

' ------------------------------------------------------------
' Da li je prijava ukljucena (opt-in). Prazno/missing = NE.
' ------------------------------------------------------------
' Ukljuci/iskljuci AUTH samo za trajanje testa. Vraca prethodno stanje, da ga
' test moze vratiti bez pamcenja u svojoj promenljivoj.
Public Function AuthTestUkljuci(ByVal ukljuceno As Boolean) As Boolean
    If Not IsTestMode() Then Exit Function
    AuthTestUkljuci = mAuthTestOn
    mAuthTestOn = ukljuceno
End Function

Public Function AuthEnabled() As Boolean
    On Error GoTo EH
    Dim v As String
    ' Seam ide PRVI i samo u test rezimu; u produkciji je mAuthTestOn uvek False.
    If mAuthTestOn Then
        AuthEnabled = True
        Exit Function
    End If
    v = UCase$(Trim$(GetConfigValue(CFG_KEY_AUTH_ENABLED)))
    AuthEnabled = (v = "YES" Or v = "DA" Or v = "TRUE" Or v = "1")
    Exit Function
EH:
    AuthEnabled = False
End Function

' ------------------------------------------------------------
' Prikazi prijavu i vrati True ako je uspela.
' Prijava je od v6-ui-214 FAZA ljuske (modUiFaze), ne svoja forma: kartica se
' crta runtime-om preko frmOtkupUI, a validaciju i dalje radi ValidateLogin.
' Poziva se iz modMain.StartApp (posle AccessGateOrQuit, pre splash faze) i iz
' ljuske pri zameni operatera (modOtkupUI.DoSwitchOperater).
' ------------------------------------------------------------
Public Function Login() As Boolean
    Dim biUser As String, biUloga As String, biIme As String, biLog As Boolean
    On Error GoTo EH

    ' Prethodna sesija se PAMTI pa tek onda gasi.
    '
    ' Do v6-ui-203 je ovde stajao go Logout: klik na "Operater" pa "Otkazi"
    ' (ili tri promasena PIN-a) je Login vracao False, a stara sesija je vec
    ' bila obrisana. Ljuska je tada javljala "operater ostao isti" i zadrzavala
    ' njegovo ime, sidebar i podatke -- dok je auth kontekst bio PRAZAN.
    ' Prikaz je tvrdio jedno, prava su bila drugo.
    '
    ' Otkazivanje znaci "ne menjam operatera", pa se vraca tacno stanje pre
    ' klika. Ne dobija se nista novo: ta sesija je vec bila prijavljena i vec
    ' je bila na ekranu.
    biUser = gCurrentUser
    biUloga = gCurrentUserUloga
    biIme = gCurrentUserIme
    biLog = gLoggedIn

    Logout

    Login = PrikaziPrijavu()

    If Not Login Then VratiSesiju biUser, biUloga, biIme, biLog
    Exit Function
EH:
    LogErr "modAuth.Login"
    Login = False
    VratiSesiju biUser, biUloga, biIme, biLog
End Function

' Sam dijalog prijave. Izdvojen iz Login-a zato sto je to JEDINI korak koji se
' u harnessu ne moze odigrati -- a sve oko njega (pamcenje sesije, gasenje,
' vracanje) je bas ono sto je bilo pokvareno. Sa ovim se test vozi kroz PRAVI
' Login, pa sabotaza nad vracanjem sesije stvarno obara tvrdnju.
'
' Ljuska je bez modalnosti, pa cekanje pravi modUiFaze rukom (DoEvents dok
' dugme ne postavi ishod). Ugovor prema pozivaocu je nepromenjen: True samo
' kad je operater prijavljen.
Private Function PrikaziPrijavu() As Boolean
    If mPrijavaOdbijTest Then Exit Function      ' False, bez prikaza
    PrikaziPrijavu = modUiFaze.FazaPrijava()
End Function

' Vracanje zapamcene sesije posle NEUSPELE prijave. Odvojeno zato sto se zove
' sa dva mesta (normalan izlaz i EH), a preskoceno vracanje je bas onaj kvar
' zbog koga ovo i postoji.
'
' Prazna sesija se ne vraca: to je slucaj prve prijave (StartApp), gde nema
' cega da se vrati -- i gde bi "vracanje" bilo tiho ponistavanje odjave.
Private Sub VratiSesiju(ByVal usr As String, ByVal uloga As String, _
                        ByVal ime As String, ByVal biLog As Boolean)
    If Not biLog Then Exit Sub
    If Len(usr) = 0 Then Exit Sub
    gCurrentUser = usr
    gCurrentUserUloga = uloga
    gCurrentUserIme = ime
    gLoggedIn = True
    AuditAuth "AUTH_LOGIN_OTKAZ", "INFO", usr & " (prijava nije promenjena)"
End Sub

' Da li je BILO KO prijavljen. Javno zato sto ljuska posle neuspele zamene mora
' da zna da li je sesija vracena -- ako nije, prikaz ne sme da tvrdi da je stari
' operater i dalje tu.
Public Function JePrijavljen() As Boolean
    JePrijavljen = gLoggedIn
End Function

' REGRESIJA (samo test): "Otkazi" u prijavi ne sme da odjavi zatecenog operatera.
'
' OVO NIJE SETTER SESIJE, i to je cela poenta. U v6-ui-204 su ovde stajala dva
' javna seam-a: jedan je obarao prijavu, drugi je POSTAVLJAO korisnika, ulogu i
' gLoggedIn. Drugi je bio auth bypass: IsTestMode nije brana jer je SetTestMode
' javan (modTestMode), pa je bilo koji drugi workbook, add-in ili makro mogao
'
'     SetTestMode True
'     AuthSesijaTest "bilo_ko", ULOGA_ADMIN, "Admin"
'
' i dobiti administratorsku sesiju bez PIN-a, bez reda u tblKorisnici i bez
' ijednog audit traga.
'
' Zato procedura NE PRIMA nista i NE OSTAVLJA nista: sama napravi privremeno
' stanje, provoza PRAVI Login (odbijen iznutra, bez forme), izmeri i ODJAVI se
' pre izlaska -- i na uspesnom putu i kroz EH. Pozivalac dobija samo tekst
' nalaza. Van test rezima ne radi nista, a zatecenu sesiju ne dira.
'
' Vraca "" kad je sve kako treba, inace opis razlike.
Public Function AuthRegresijaOtkaz() As String
    Dim uspelo As Boolean, korPosle As String, prijavljenPosle As Boolean
    Const PROBNI As String = "regresija.otkaz"

    AuthRegresijaOtkaz = "regresija nije izvrsena (nije test rezim)"
    If Not IsTestMode() Then Exit Function
    ' Tudja sesija nije materijal za merenje -- ako je neko stvarno prijavljen,
    ' ne dira se nista i tvrdnja se ne postavlja.
    If gLoggedIn Then
        AuthRegresijaOtkaz = ""
        Exit Function
    End If

    On Error GoTo EH
    gCurrentUser = PROBNI
    gCurrentUserUloga = "operater"
    gCurrentUserIme = PROBNI
    gLoggedIn = True

    mPrijavaOdbijTest = True
    uspelo = Login()
    mPrijavaOdbijTest = False

    korPosle = gCurrentUser
    prijavljenPosle = gLoggedIn
    Logout                          ' NISTA se ne ostavlja iza

    AuthRegresijaOtkaz = ""
    If uspelo Then AuthRegresijaOtkaz = "odbijena prijava je vratila True."
    If korPosle <> PROBNI Then AuthRegresijaOtkaz = AuthRegresijaOtkaz & _
        " Otkaz nije vratio operatera (ostalo: '" & korPosle & "')."
    If Not prijavljenPosle Then AuthRegresijaOtkaz = AuthRegresijaOtkaz & _
        " Otkaz je ostavio odjavljenu sesiju."
    AuthRegresijaOtkaz = Trim$(AuthRegresijaOtkaz)
    Exit Function
EH:
    mPrijavaOdbijTest = False
    Logout
    AuthRegresijaOtkaz = "regresija je pukla: " & Err.description
End Function

' ------------------------------------------------------------
' Validacija kredencijala (zove je kartica prijave posle "Prijavi se").
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

' OblastZaFormu (mapa 'ime legacy forme -> oblast prava') je obrisana u koraku 7
' (docs/UI_MIGRACIJA_KATALOG.md par.27.17). Imala je TACNO JEDNOG pozivaoca --
' frmOtkupAPP.OpenContentForm -- i nestala je zajedno s njim.
'
' Prava se od tada traze SAMO kroz registar ekrana (modUiScreens.ScrDozvoljen,
' polje SCR_OBLAST). Jedna mapa manje znaci i jedno mesto manje na kome se
' oblast moze razici sa ekranom koji je trosi.

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
        ' RF-23 (item 5): provera se oslanja na plaintext PIN fallback -- PIN nije
        ' hash-ovan (legacy zapis ili SHA nedostupan pa PreparePin nije mogao da hash-uje).
        ' Signaliziramo (fail-soft log), BEZ menjanja logike provere. Vidi FM-0053 #55.9.
        On Error Resume Next
        LogWarn "modAuth.VerifyPin", "Plaintext PIN fallback u upotrebi (PIN nije hash-ovan)."
        On Error GoTo 0
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
