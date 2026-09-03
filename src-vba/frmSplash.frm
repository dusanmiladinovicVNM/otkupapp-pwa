VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "UserForm1"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' frmSplash / startup splash
' Responsibility:
'   - show branding briefly, in the look of the new UI header:
'     forest gradient, "AX" gold + "OtkupApp" cream, display font
'   - then open the app shell (modOtkupUI.ShowOtkupUI)
'   - no business logic
'
' Kontrole iz dizajnera (lblApp, lblVersion, lblBy) se samo stilizuju i
' premestaju; sve ostalo nastaje u runtime-u (modUiKit.NewLbl), pa se .frx
' ne dira. Nema module-level MSForms deklaracija -- forma ostaje "meka" za
' self-update (modSelfUpdate.IsHardModuleBody). Sve mere su u tackama.
' ============================================================

Private Const SPL_W  As Single = 400
Private Const SPL_H  As Single = 236
Private Const STRIPS As Long = 24        ' isti gradijent kao zaglavlje ekrana
Private Const LOGO_Y As Single = 62      ' vrh "AX"; "OtkupApp" deli donju ivicu

Private mChromeRemoved As Boolean
Private m_Started As Boolean
Private m_IsNavigating As Boolean

' Prigusen tekst na forest podlozi -- ista vrednost kao hdrStat u zaglavlju.
Private Function MutedOnForest() As Long
    MutedOnForest = RGB(178, 190, 172)
End Function

Private Sub UserForm_Initialize()
    On Error GoTo EH

    mChromeRemoved = False
    m_Started = False
    m_IsNavigating = False

    BuildSplash

    Exit Sub

EH:
    LogErr "frmSplash.UserForm_Initialize"
End Sub

Private Sub BuildSplash()
    Dim i As Long, fnt As String, sw As Single
    fnt = DisplayFont()

    Me.width = SPL_W
    Me.Height = SPL_H
    Me.BackColor = C_FOREST

    ' gradijent ide IZA kontrola iz dizajnera (ZOrder 1 = pozadi)
    sw = SPL_W / STRIPS
    For i = 0 To STRIPS - 1
        NewLbl Me, "spGr" & i, "", i * sw, 0, sw + 1, SPL_H, 8, False, 0, _
               Lerp(C_FOREST, C_FOREST_DK, i / (STRIPS - 1))
        Me.Controls("spGr" & i).ZOrder 1
    Next i

    ' zlatna nit na vrhu -- isti akcenat kao aktivna stavka sidebara
    NewLbl Me, "spLine", "", 0, 0, SPL_W, 3, 8, False, 0, C_GOLD

    ' logo: "AX" zlatno (40pt) + "OtkupApp" krem (30pt), display font
    NewLbl Me, "spAX", "AX", PAD + 8, LOGO_Y, 74, TxtH(40), 40, True, C_GOLD, -1, fmTextAlignLeft, fnt

    With lblApp
        .caption = "OtkupApp"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = C_CREAM
        .Font.name = fnt
        .Font.Size = 30
        .Font.bold = True
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Left = PAD + 84
        .top = LOGO_Y + TxtH(40) - TxtH(30)
        .width = 260
        .Height = TxtH(30)
        .ZOrder 0
    End With

    With lblVersion
        .caption = "v" & APP_VERSION
        .BackStyle = fmBackStyleTransparent
        .ForeColor = MutedOnForest()
        .Font.name = F_UI
        .Font.Size = TS_META
        .Font.bold = False
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Left = PAD + 86
        .top = LOGO_Y + TxtH(40) + 2
        .width = 200
        .Height = TxtH(TS_META)
        .ZOrder 0
    End With

    ' podnozje: tanka linija, "Powered by AgriX" levo, status desno
    NewLbl Me, "spDiv", "", PAD, SPL_H - 42, SPL_W - 2 * PAD, 1, 8, False, 0, C_HDR_EDGE

    With lblBy
        .caption = "Powered by AgriX"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = MutedOnForest()
        .Font.name = F_UI
        .Font.Size = TS_MICRO
        .Font.bold = False
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Left = PAD
        .top = SPL_H - 30
        .width = 160
        .Height = TxtH(TS_MICRO)
        .ZOrder 0
    End With

    ' zlatna tacka + "Pokrecem aplikaciju..." -- kao hdrDot/hdrStat u zaglavlju
    NewLbl Me, "spDot", "", SPL_W - PAD - 160, SPL_H - 27, 6, 6, 8, False, 0, C_GOLD
    NewLbl Me, "spStat", Poruka("OTKUI_SPLASH_POKRECEM"), SPL_W - PAD - 150, SPL_H - 30, 150, _
           TxtH(TS_META), TS_META, False, MutedOnForest(), -1, fmTextAlignRight
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    EnsureUserFormChromeRemoved Me, mChromeRemoved

    If m_Started Then Exit Sub
    m_Started = True

    WaitSeconds 2
    OpenAppShell

    Exit Sub

EH:
    LogErr "frmSplash.UserForm_Activate"
    OpenAppShell
End Sub

Private Sub WaitSeconds(ByVal secondsToWait As Double)
    On Error GoTo EH

    Dim endTime As Date
    endTime = DateAdd("s", secondsToWait, Now)

    Do While Now < endTime
        DoEvents
    Loop

    Exit Sub

EH:
    LogErr "frmSplash.WaitSeconds"
End Sub

Private Sub OpenAppShell()
    On Error GoTo EH

    If m_IsNavigating Then Exit Sub
    m_IsNavigating = True

    Unload Me
    modOtkupUI.ShowOtkupUI

    Exit Sub

EH:
    LogErr "frmSplash.OpenAppShell"

    On Error Resume Next
    modOtkupUI.ShowOtkupUI
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OpenAppShell
    End If
End Sub
