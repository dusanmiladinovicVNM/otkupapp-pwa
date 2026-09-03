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
'   - full-screen brand moment in the shell palette, then
'     open the app shell (modOtkupUI.ShowOtkupUI)
'   - no business logic
'
' JEDAN ZNAK, NE TRI. .frx nosi DVA logotipa (Image12 = AX|OtkupApp,
' Image25 = AgriX). Ranija verzija je preko njih crtala jos i tekstualni
' "AX OtkupApp", pa se marka videla tri puta. Sada se koristi PRAVI logotip
' (Image12), a Image25 i tekstualni znak se gase: dole ostaje samo tiha
' linija "Powered by AgriX".
'
' Kontrole iz dizajnera se samo premestaju i gase; pozadina je runtime
' (modUiKit.NewLbl + Lerp), pa se .frx ne dira. Nema module-level MSForms
' deklaracija -- forma ostaje "meka" za self-update. Mere su u tackama.
' ============================================================

Private Const BANDS   As Long = 40      ' trake vertikalnog gradijenta
Private Const LOGO_W  As Single = 340   ' okvir logotipa; Zoom cuva odnos stranica
Private Const LOGO_H  As Single = 86
Private Const FOOT_H  As Single = 52    ' podnozje: linija + dva reda teksta

Private mChromeRemoved As Boolean
Private m_Started As Boolean
Private m_IsNavigating As Boolean

' Prigusen tekst na forest podlozi -- ista vrednost kao hdrStat u zaglavlju ljuske.
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
    Dim i As Long, w As Single, h As Single, bh As Single, cx As Single, Y As Single

    ' Ceo ekran: isti racun kao modOtkupUI.GoFullScreen (ScreenWidthPoints /
    ' ScreenHeightPoints iz modWindow), jer je Excel u ovom trenutku sakriven
    ' pa Application.Width ne vredi.
    w = ScreenWidthPoints()
    h = ScreenHeightPoints()
    If w < 600 Then w = 600
    If h < 400 Then h = 400

    Me.StartUpPosition = 0
    Me.Left = 0
    Me.top = 0
    Me.width = w
    Me.Height = h
    Me.BackColor = C_FOREST

    ' Vertikalni gradijent preko celog ekrana. Trake idu IZA svega (ZOrder 1);
    ' ne preklapaju se, pa im medjusobni redosled nije bitan.
    bh = h / BANDS
    For i = 0 To BANDS - 1
        NewLbl Me, "spGr" & i, "", 0, i * bh, w, bh + 1, 8, False, 0, _
               Lerp(C_FOREST, C_FOREST_DK, i / (BANDS - 1))
        Me.Controls("spGr" & i).ZOrder 1
    Next i

    ' zlatna nit na vrhu -- isti akcenat kao aktivna stavka sidebara
    NewLbl Me, "spLine", "", 0, 0, w, 3, 8, False, 0, C_GOLD

    ' LOGOTIP iz .frx, centriran u gornjoj trecini. Zoom cuva odnos stranica,
    ' pa okvir sme da bude fiksan a da se slika ne izoblici.
    cx = (w - LOGO_W) / 2
    Y = h * 0.34 - LOGO_H / 2
    With Image12
        .PictureSizeMode = fmPictureSizeModeZoom
        .PictureAlignment = fmPictureAlignmentCenter
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Left = cx: .top = Y: .width = LOGO_W: .Height = LOGO_H
        .Visible = True
        .ZOrder 0
    End With

    ' Tekstualni znak i drugi logotip se GASE -- v. zaglavlje modula.
    lblApp.Visible = False
    Image25.Visible = False

    ' verzija ispod logotipa, centrirano
    With lblVersion
        .caption = "v" & APP_VERSION
        .BackStyle = fmBackStyleTransparent
        .ForeColor = MutedOnForest()
        .Font.name = F_UI
        .Font.Size = TS_META
        .Font.bold = False
        .TextAlign = fmTextAlignCenter
        .WordWrap = False
        .Left = cx: .top = Y + LOGO_H + 10: .width = LOGO_W: .Height = TxtH(TS_META)
        .ZOrder 0
    End With

    ' podnozje: tanka linija, "Powered by AgriX" levo, status desno
    NewLbl Me, "spDiv", "", PAD, h - FOOT_H, w - 2 * PAD, 1, 8, False, 0, C_HDR_EDGE

    With lblBy
        .caption = "Powered by AgriX"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = MutedOnForest()
        .Font.name = F_UI
        .Font.Size = TS_MICRO
        .Font.bold = False
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Left = PAD: .top = h - FOOT_H + 16: .width = 200: .Height = TxtH(TS_MICRO)
        .ZOrder 0
    End With

    ' zlatna tacka + status -- kao hdrDot/hdrStat u zaglavlju ljuske
    NewLbl Me, "spDot", "", w - PAD - 168, h - FOOT_H + 19, 6, 6, 8, False, 0, C_GOLD
    NewLbl Me, "spStat", Poruka("OTKUI_SPLASH_POKRECEM"), w - PAD - 158, h - FOOT_H + 15, 158, _
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
