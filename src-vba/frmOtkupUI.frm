VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtkupUI 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmOtkupUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtkupUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================================
' Kod-behind za PRAZNU UserForm nazvanu frmOtkupUI.
' Nijedna kontrola ne postoji u dizajneru - sve gradi modOtkupUI.
' Properties: (Name)=frmOtkupUI, ShowModal=False.
'
' PAZNJA: forma se NE sme zvati frmOtkup - to ime zauzima postojeca
' produkciona forma (src-vba/frmOtkup.frm, 1294 linije).
'
' Ovde NEMA nijedne module-level deklaracije - ni WithEvents, ni
' "As MSForms.*". To je namerno: IsHardModuleBody (modSelfUpdate.bas)
' bi inace proglasio formu tvrdom i svaka njena izmena bi trazila
' REINSTALL umesto self-update-a. Sav UI i svi eventi zive u
' modOtkupUI / clsFlatBtn.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Private Sub UserForm_Initialize()
    ' ScreenUpdating se MORA vratiti i kad gradnja pukne - inace Excel ostaje
    ' zamrznut ekran bez ijedne poruke, a operater misli da je aplikacija pala.
    On Error GoTo EH
    Application.ScreenUpdating = False
    modOtkupUI.BuildOtkupScreen Me
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    Application.ScreenUpdating = True
    modOtkupUI.OtkupUI_BuildFailed errNum, errDesc
End Sub

Private Sub UserForm_Activate()
    ' Ceo ekran bez Windows naslovne trake: GoFullScreen postavlja velicinu,
    ' skida WS_CAPTION i dodaje WS_THICKFRAME (forma se i dalje hvata za
    ' ivicu). Posto sistemskog X vise nema, zatvaranje ide preko btnClose.
    ' MakeResizable se NE zove - prazna je; stil menja iskljucivo GoFullScreen,
    ' jer jedino ono preko FormHwnd zna koji je prozor NAS.
    modOtkupUI.GoFullScreen Me
    modOtkupUI.LayoutOtkup Me
    On Error Resume Next
    Me.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").SetFocus
    ' Mreza se puni TEK sada: hrom je vec na ekranu, pa citanje tabele ne
    ' produzava vreme do prvog prikaza.
    modOtkupUI.EnsureGridLoaded
End Sub

' Zatvaranje preko X ne unload-uje formu nego je sakrije: sledece otvaranje je
' trenutno jer se kontrole ne grade ponovo. Formu stvarno otpusta
' PrepareRuntimeForSelfUpdate (Unload svih formi) i zatvaranje radne sveske.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        ' NE "Me.Hide": posto se forma ne unload-uje, Terminate ne puca, pa bi
        ' zakljucana stanica ostala zauzeta na serveru. OtkupUI_Sakrij pusta
        ' lock pa sakriva - isto sto legacy radi u frmOtkup.UserForm_QueryClose.
        modOtkupUI.OtkupUI_Sakrij
    End If
End Sub

Private Sub UserForm_Resize()
    Static busy As Boolean
    If busy Then Exit Sub
    busy = True
    On Error Resume Next
    If Me.Height < 560 Then Me.Height = 560
    If Me.width < 660 Then Me.width = 660
    modOtkupUI.LayoutOtkup Me
    busy = False
End Sub

' rezerva: hover se gasi i kad pokazivac izadje na golu povrsinu forme
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modOtkupUI.ResetAllBtnVisuals
End Sub

' MSForms UserForm nema KeyPreview - ovaj event radi samo dok fokus NIJE
' u nekoj kontroli. Isti KeyCode se hvata i u clsFlatBtn (TextBox /
' ComboBox KeyDown) i salje u istu rutinu.
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If (Shift And 2) <> 0 Then
        Select Case KeyCode
            Case vbKeyK: modOtkupUI.FocusKontekst: KeyCode = 0: Exit Sub
            Case vbKeyF: modOtkupUI.FocusPretraga: KeyCode = 0: Exit Sub
        End Select
    End If
    If modOtkupUI.HandleGlobalKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub UserForm_Terminate()
    modOtkupUI.OtkupUI_FormClosed
End Sub

