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
'  1) THREAD hook (WH_MOUSE), NE globalni WH_MOUSE_LL. Hvata samo mouse
'     poruke Excel-ove UI niti -> minimalan trosak; ne dira sistemsko
'     ponasanje misa. Ako Windows ikad skine hook (timeout pod opterece-
'     njem), posledica je samo "tockic privremeno ne radi" (fail-safe),
'     nikad crash; sledeci UserForm_Activate ga ponovo instalira.
'
'  2) Hook zivi SAMO dok forma ima fokus: Attach na UserForm_Activate,
'     Detach na UserForm_Deactivate/QueryClose. Posto su forme modeless,
'     odlazak na VBE odmah okine Deactivate -> hook se skine PRE nego sto
'     bilo kakav VBA reset moze da ostavi "mrtav" AddressOf pointer (to
'     je jedini realan nacin da mouse hook srusi Excel).
'
'  3) Callback (MouseProc) je neprobojan: "On Error Resume Next" prva
'     linija, re-entrancy guard (mBusy), jeftina common-path grana za
'     ne-wheel poruke, i UVEK CallNextHookEx (lanac se nikad ne prekida).
'
'  4) Master prekidac: MouseWheel_SetEnabled False iskljuci sve u hodu.
'
'  5) Koja lista se skroluje bira se preko MouseMove (clsWheelList),
'     bez ijedne geometrijske/DPI/HWND racunice.
'
' KORISCENJE (po formi koja ima liste, sve idempotentno / no-op-safe):
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
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal SOURCE As LongPtr, ByVal Length As LongPtr)

Private Const WH_MOUSE As Long = 7
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WHEEL_STEP As Long = 3      ' redova po jednom "kliku" tockica

' Ofset polja mouseData u MOUSEHOOKSTRUCTEX (razlicit 32/64-bit zbog
' velicine pointera). Citamo SAMO ta 4 bajta -> bez zavisnosti od
' poravnanja (packing) celog Type-a.
#If Win64 Then
Private Const MD_OFFSET As LongPtr = 32
#Else
Private Const MD_OFFSET As LongPtr = 20
#End If

' --- Stanje modula ---
Private mHook As LongPtr                  ' handle hook-a; 0 = nije instaliran
Private mHot As MSForms.ListBox           ' lista trenutno pod misem
Private mWrappers As Collection           ' clsWheelList instance (drzi ih zivim)
Private mBusy As Boolean                  ' re-entrancy guard u callback-u
Private mDisabled As Boolean              ' master kill-switch (False = ukljuceno)

' ------------------------------------------------------------
' Attach - poziva se iz UserForm_Activate. Idempotentno.
' ------------------------------------------------------------
Public Sub MouseWheel_Attach(ByVal frm As Object)
    On Error Resume Next
    If mDisabled Then Exit Sub

    ' Omotaci za liste ove forme - gradimo samo ako ih nema (posle Detach-a).
    If mWrappers Is Nothing Then
        Set mWrappers = New Collection
        BuildWrappers frm, mWrappers
    End If

    ' Hook instaliramo jednom, thread-scoped na tekucu (Excel UI) nit.
    If mHook = 0 Then
        mHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, GetCurrentThreadId())
    End If
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
    mDisabled = Not onOff
    If mDisabled Then MouseWheel_Detach
End Sub

' Poziva clsWheelList.lst_MouseMove: lista pod misem postaje "aktivna".
Public Sub MouseWheel_SetHot(ByVal lb As MSForms.ListBox)
    On Error Resume Next
    Set mHot = lb
End Sub

' ------------------------------------------------------------
' BuildWrappers - rekurzivno nadje sve ListBox-ove (i one u Frame /
' MultiPage / Page) i obmota ih u clsWheelList. Sve guardovano.
' ------------------------------------------------------------
Private Sub BuildWrappers(ByVal container As Object, ByVal coll As Collection)
    On Error Resume Next
    Dim c As Object, w As clsWheelList, pg As Object
    For Each c In container.Controls
        Select Case TypeName(c)
            Case "ListBox"
                Set w = New clsWheelList
                Set w.lst = c
                coll.Add w
            Case "Frame"
                BuildWrappers c, coll
            Case "MultiPage"
                For Each pg In c.Pages
                    BuildWrappers pg, coll
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

    ' Common path: sve sto nije nas slucaj prosledi dalje i izadji.
    If nCode < 0 Or mDisabled Or mBusy Or mHook = 0 Then
        MouseProc = CallNextHookEx(mHook, nCode, wParam, lParam)
        Exit Function
    End If

    If wParam = WM_MOUSEWHEEL Then
        If Not mHot Is Nothing Then
            Dim md As Long
            md = 0
            ' Procitaj samo mouseData (hi-word = smer tockica).
            CopyMemory md, lParam + MD_OFFSET, 4
            If md <> 0 Then
                mBusy = True
                ScrollHot (md > 0)        ' md>0 = tockic gore, md<0 = dole
                mBusy = False
            End If
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
#End If
