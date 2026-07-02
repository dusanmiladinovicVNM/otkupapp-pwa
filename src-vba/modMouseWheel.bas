Option Explicit
' ============================================================
' modMouseWheel - skrolovanje MSForms.ListBox tockicem misa.
'
' MSForms kontrole ne primaju WM_MOUSEWHEEL po defaultu (poznato VBA
' ogranicenje) - traka za skrolovanje radi, ali tockic ne. Ovaj modul
' hvata tockic preko Windows mouse hook-a i skroluje listu pod misem.
'
' DIZAJN JE PODredjen jednom cilju: NIKAD ne srusiti / usporiti / cudno
' ponasati app. Konkretne mere:
'
'  1) LOW-LEVEL hook (WH_MOUSE_LL). Hvata tockic na sistemskom nivou, pouzdano
'     bez obzira na fokus/foreground i trenutak instalacije. (Thread WH_MOUSE
'     hook NIJE hvatao tockic dok se foreground ne osvezi -> radilo "tek posle
'     minimize/restore".) Ako Windows skine hook (timeout), posledica je samo
'     "tockic privremeno ne radi" (fail-safe), nikad crash; sledeci hover ga
'     ponovo instalira.
'
'  2) Hook zivi SAMO dok forma ima fokus I dok je VBE ZATVOREN:
'     - Instalira se LENJIVO, tek na prvi MouseMove nad listom (EnsureHook),
'       NE u Activate -> ne kvari iscrtavanje novootvorene forme (beli ekran).
'     - EnsureHook NE dize hook dok je VBE otvoren -> tockic se SAM ugasi cim
'       otvoris VBE i sam upali kad ga zatvoris (bez rucnog gasenja).
'     - Detach na UserForm_Deactivate/QueryClose: odlazak sa forme (npr. na
'       VBE) odmah skine hook, PRE nego sto VBA reset moze da ostavi "mrtav"
'       AddressOf pointer (jedini realan nacin da mouse hook srusi Excel).
'
'  3) Callback (MouseProc) je neprobojan: "On Error Resume Next" prva
'     linija, TVRDI re-entrancy guard (mInHook) oko CELOG callback-a,
'     jeftina common-path grana, i UVEK CallNextHookEx (lanac se ne prekida).
'
'  4) OFF PO DEFAULTU (mArmed=False): na import se NE instalira nikakav hook;
'     Attach/Register su no-op dok se ne "naoruza". Ukljucuje se SAMO eksplicitno:
'     MouseWheel_SetEnabled True (pa otvori/re-fokusiraj formu). Razlog: mouse
'     hook uz OTVOREN VBE ume da zaledi Excel, pa feature ne sme da se pali sam
'     na startu. Iskljucivanje u hodu: MouseWheel_SetEnabled False.
'
'  5) Koja lista se skroluje bira se preko MouseMove (clsWheelList),
'     bez ijedne geometrijske/DPI/HWND racunice.
'
' KORISCENJE: forme vec zovu Attach/Detach (idempotentno), ali dok se ne
' "naoruza" sve je no-op. Paljenje je KONFIGURISANO, ne rucno:
'     - StartApp cita MOUSEWHEEL_SCROLL (tblLocalConfig; prazno = DA) i zove
'       MouseWheel_On/Off; Podesavanja ("Interfejs / lokalno") primenjuju odmah.
'     - Rucno (Alt+F8, dijagnostika): MouseWheel_On / MouseWheel_Off.
' Po formi (vec oziceno):
'     Private Sub UserForm_Activate()
'         MouseWheel_Attach Me
'         ...
'     Private Sub UserForm_Deactivate()
'         MouseWheel_Detach
'     Private Sub UserForm_QueryClose(...)
'         MouseWheel_Detach
'         ...
' ============================================================

#If VBA7 Then

' --- Win32 (isti PtrSafe/LongPtr obrazac kao modWindow/modClipboard) ---
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal SOURCE As LongPtr, ByVal Length As LongPtr)

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WHEEL_STEP As Long = 3      ' redova po jednom "kliku" tockica
Private Const ALIVE_MOVES As Long = 24    ' koliko pokreta misa hook "zivi" posto mis sidje s liste

' Ofset polja mouseData u MSLLHOOKSTRUCT: POINT pt (x,y = 2x4 = 8 bajta) pa
' mouseData -> UVEK offset 8 (isto na 32 i 64-bit; nema pointera pre njega).
' Citamo SAMO ta 4 bajta -> bez zavisnosti od poravnanja celog Type-a.
Private Const MD_OFFSET As LongPtr = 8

' --- Stanje modula ---
Private mHook As LongPtr                  ' handle hook-a; 0 = nije instaliran
Private mHot As MSForms.ListBox           ' lista trenutno pod misem
Private mWrappers As Collection           ' clsWheelList instance (drzi ih zivim)
Private mInHook As Boolean                ' TVRDI re-entrancy guard oko celog callback-a
Private mAlive As Long                    ' keep-alive: hook aktivan SAMO dok je mis nad listom
Private mArmed As Boolean                 ' OFF po defaultu (False)! Hook se NE instalira dok
                                          ' ga MouseWheel_SetEnabled True eksplicitno ne "naoruza".
                                          ' Razlog: mouse hook + otvoren VBE ume da zaledi Excel;
                                          ' zato feature ne sme da se pali sam od sebe na import.

' ------------------------------------------------------------
' Attach - poziva se iz UserForm_Activate. Idempotentno.
' ------------------------------------------------------------
Public Sub MouseWheel_Attach(ByVal frm As Object)
    On Error Resume Next
    If Not mArmed Then Exit Sub

    ' Obmotaj liste ove forme. UVEK (ne samo kad je mWrappers prazno) jer se
    ' MouseWheel_On moze pozvati DOK je forma vec otvorena (njen Activate je
    ' prosao neoruzan pa nista nije obmotao). AddWrapper dedupe-uje.
    If mWrappers Is Nothing Then Set mWrappers = New Collection
    BuildWrappers frm

    ' Hook se NE instalira ovde! Instalacija bas u trenutku iscrtavanja forme
    ' (Activate) izgladnjuje WM_PAINT (najnizi prioritet) -> beli/neiscrtan
    ' ekran dok minimize/restore ne forsira repaint. Zato ide LENJIVO, na prvi
    ' MouseMove nad listom (EnsureHook iz SetHot), kad je forma vec iscrtana.
    ' Tako se navigacija (otvaranje ekrana) uvek desava bez instaliranog hook-a.
End Sub

' ------------------------------------------------------------
' Detach - poziva se iz UserForm_Deactivate / QueryClose / Terminate.
' Idempotentno; bezbedno pozvati vise puta.
' ------------------------------------------------------------
Public Sub MouseWheel_Detach()
    On Error Resume Next
    If mHook <> 0 Then
        UnhookWindowsHookEx mHook
        mHook = 0
    End If
    Set mWrappers = Nothing
    Set mHot = Nothing
End Sub

' ------------------------------------------------------------
' Master prekidac - MouseWheel_SetEnabled False ugasi sve odmah.
' ------------------------------------------------------------
Public Sub MouseWheel_SetEnabled(ByVal onOff As Boolean)
    mArmed = onOff
    If Not mArmed Then MouseWheel_Detach
End Sub

' Bezargumentni obmotaci - vidljivi u Alt+F8, pa mogu iz Excela BEZ VBE:
'   MouseWheel_On  -> naoruzaj (pa otvori/re-fokusiraj formu sa listom)
'   MouseWheel_Off -> ugasi i skini hook
Public Sub MouseWheel_On()
    MouseWheel_SetEnabled True
    ' Obmotaj liste na SVIM trenutno otvorenim formama, da tockic radi ODMAH
    ' (bez cekanja na sledeci UserForm_Activate, tj. bez minimize/restore).
    On Error Resume Next
    Dim i As Long
    For i = 0 To UserForms.Count - 1
        MouseWheel_Attach UserForms(i)
    Next i
End Sub

Public Sub MouseWheel_Off()
    MouseWheel_SetEnabled False
End Sub

' Poziva clsWheelList.lst_MouseMove: lista pod misem postaje "aktivna".
' Ovde se hook i instalira (lenjivo) - forma je do ovog trenutka vec iscrtana.
' Svaki pokret nad listom osvezava mAlive (keep-alive) -> hook ostaje.
Public Sub MouseWheel_SetHot(ByVal lb As MSForms.ListBox)
    On Error Resume Next
    Set mHot = lb
    mAlive = ALIVE_MOVES     ' mis je NAD listom -> drzi hook zivim
    EnsureHook
End Sub

' Instaliraj hook tek kad zaista treba (mis dosao nad listu), da instalacija
' NE padne u kriticni trenutak iscrtavanja forme (beli ekran). Idempotentno.
Private Sub EnsureHook()
    On Error Resume Next
    If Not mArmed Then Exit Sub
    If mHook <> 0 Then Exit Sub          ' vec instaliran -> izlaz (bez per-move provera)

    ' BRANA: NE dizi hook dok je VBE (editor) otvoren. Mouse hook uz otvoren VBE
    ' ume da zaledi Excel. Ovim se tockic AUTOMATSKI gasi cim je VBE otvoren
    ' (bilo gde, makar u pozadini) i sam upali kad zatvoris VBE - bez rucnog
    ' iskljucivanja. (Uz to, hook se ionako skida na UserForm_Deactivate, koje
    ' okine cim odes sa forme na VBE.)
    If VbeIsOpen Then Exit Sub

    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, GetModuleHandle(vbNullString), 0)
End Sub

' True ako je VBE prozor otvoren/vidljiv. Na gresku (npr. iskljucen "Trust access
' to the VBA project object model") vraca True = tretiraj kao otvoren (bezbednije:
' ne dizi hook). Poziva se retko (samo kad bi se hook dizao), pa COM poziv ne smeta.
Private Function VbeIsOpen() As Boolean
    On Error GoTo assumeOpen
    VbeIsOpen = Application.VBE.MainWindow.Visible
    Exit Function
assumeOpen:
    VbeIsOpen = True
End Function

' ------------------------------------------------------------
' Register - registruj JEDNU listu na zahtev. Za dinamicke liste koje ne
' postoje u trenutku MouseWheel_Attach (npr. paneli koji se grade lazy,
' modOtkupBlok liste). Pozovi odmah posle Controls.Add. No-op-safe.
' ------------------------------------------------------------
Public Sub MouseWheel_Register(ByVal lb As MSForms.ListBox)
    On Error Resume Next
    If Not mArmed Then Exit Sub
    If lb Is Nothing Then Exit Sub

    AddWrapper lb
    ' Hook se NE instalira ovde - lenjivo, na prvi MouseMove nad listom
    ' (EnsureHook iz SetHot), da se ne pokvari iscrtavanje forme.
End Sub

' Napravi omotac za JEDNU listu i dodaj ga u mWrappers.
' Deljeno: BuildWrappers + MouseWheel_Register (anti-duplikacija).
Private Sub AddWrapper(ByVal lb As MSForms.ListBox)
    On Error Resume Next
    If lb Is Nothing Then Exit Sub
    If mWrappers Is Nothing Then Set mWrappers = New Collection

    ' Dedupe: preskoci ako je ova ista lista vec obmotana (Attach se sada
    ' zove vise puta / iz vise izvora).
    Dim wx As clsWheelList
    For Each wx In mWrappers
        If wx.lst Is lb Then Exit Sub
    Next wx

    Dim w As clsWheelList
    Set w = New clsWheelList
    Set w.lst = lb
    mWrappers.Add w
End Sub

' ------------------------------------------------------------
' BuildWrappers - rekurzivno nadje sve ListBox-ove (i one u Frame /
' MultiPage / Page) i obmota ih (AddWrapper). Sve guardovano.
' ------------------------------------------------------------
Private Sub BuildWrappers(ByVal container As Object)
    On Error Resume Next
    Dim c As Object, pg As Object
    For Each c In container.Controls
        Select Case TypeName(c)
            Case "ListBox"
                AddWrapper c
            Case "Frame"
                BuildWrappers c
            Case "MultiPage"
                For Each pg In c.Pages
                    BuildWrappers pg
                Next pg
        End Select
    Next c
End Sub

' ------------------------------------------------------------
' MouseProc - Windows callback. MORA biti neprobojan i jeftin.
' Vraca CallNextHookEx u SVIM granama (nikad ne prekida lanac).
' ------------------------------------------------------------
Public Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    On Error Resume Next

    ' Re-entrancy guard + brzi izlaz. (mArmed se proverava po grani.)
    If mInHook Or nCode < 0 Or mHook = 0 Then
        MouseProc = CallNextHookEx(mHook, nCode, wParam, lParam)
        Exit Function
    End If

    If wParam = WM_MOUSEWHEEL Then
        ' Tockic: skroluj aktivnu listu.
        mInHook = True
        If mArmed And Not mHot Is Nothing Then
            Dim md As Long
            md = 0
            CopyMemory md, lParam + MD_OFFSET, 4   ' samo mouseData (smer)
            If md <> 0 Then ScrollHot (md > 0)      ' md>0 = gore, md<0 = dole
        End If
        mInHook = False
    Else
        ' Bilo koji drugi dogadjaj (uglavnom POKRET misa). Dok je mis nad
        ' listom, clsWheelList MouseMove stalno vraca mAlive na pun iznos.
        ' Cim mis SIDJE s liste, mAlive padne na 0 i globalni LL hook se SAM
        ' skine -> kursor van liste vise NE secka. Re-hover ga ponovo digne.
        mAlive = mAlive - 1
        If mAlive <= 0 Then
            UnhookWindowsHookEx mHook
            mHook = 0
            Set mHot = Nothing
        End If
    End If

    MouseProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

' Skroluj aktivnu listu za WHEEL_STEP redova, sa clamp-om.
Private Sub ScrollHot(ByVal up As Boolean)
    On Error Resume Next
    Dim n As Long, t As Long
    n = mHot.ListCount
    If n <= 0 Then Exit Sub

    t = mHot.TopIndex
    If up Then
        t = t - WHEEL_STEP
    Else
        t = t + WHEEL_STEP
    End If
    If t < 0 Then t = 0
    If t > n - 1 Then t = n - 1

    mHot.TopIndex = t
End Sub

#Else
' ============================================================
' Pre-VBA7 (Office 2007 i stariji): mouse hook nije podrzan ovde;
' sve su no-op stubovi da projekat kompajlira i radi bez tockica.
' ============================================================
Public Sub MouseWheel_Attach(ByVal frm As Object)
End Sub

Public Sub MouseWheel_Detach()
End Sub

Public Sub MouseWheel_SetEnabled(ByVal onOff As Boolean)
End Sub

Public Sub MouseWheel_SetHot(ByVal lb As Object)
End Sub

Public Sub MouseWheel_Register(ByVal lb As Object)
End Sub

Public Sub MouseWheel_On()
End Sub

Public Sub MouseWheel_Off()
End Sub
#End If
