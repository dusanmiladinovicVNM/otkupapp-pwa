Attribute VB_Name = "modWindow"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function GetActiveWindow Lib "user32" () As Long
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Public Const GWL_STYLE As Long = -16
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_BORDER As Long = &H800000
Public Const SM_CXSCREEN As Long = 0
Public Const SM_CYSCREEN As Long = 1
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90

Public Function ScreenWidthPoints() As Single
    ScreenWidthPoints = PixelsToPointsX(GetSystemMetrics(SM_CXSCREEN))
End Function

Public Function ScreenHeightPoints() As Single
    ScreenHeightPoints = PixelsToPointsY(GetSystemMetrics(SM_CYSCREEN))
End Function

Public Function PixelsToPointsX(ByVal pixels As Long) As Single
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If

    Dim dpiX As Long

    hdc = GetDC(0)
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc

    If dpiX <= 0 Then dpiX = 96

    PixelsToPointsX = pixels * 72 / dpiX
End Function

Public Function PixelsToPointsY(ByVal pixels As Long) As Single
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If

    Dim dpiY As Long

    hdc = GetDC(0)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc

    If dpiY <= 0 Then dpiY = 96

    PixelsToPointsY = pixels * 72 / dpiY
End Function

Public Sub RemoveUserFormChrome(ByVal frm As Object, _
                                Optional ByVal removeBorder As Boolean = False)
    On Error GoTo EH

#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If

    Dim style As Long

    ' Bitno: ne koristiti privremeni Caption.
    ' Ovo prati stari pattern koji vec radi u formama.
    frm.caption = vbNullString

    hwnd = GetActiveWindow()

    If hwnd = 0 Then
        hwnd = FindWindow("ThunderDFrame", frm.caption)
    End If

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)

        style = style And Not WS_CAPTION

        If removeBorder Then
            style = style And Not WS_BORDER
        End If

        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If

    Exit Sub

EH:
    On Error Resume Next
    frm.caption = vbNullString
End Sub

Public Sub EnsureUserFormChromeRemoved(ByVal frm As Object, _
                                       ByRef chromeRemoved As Boolean, _
                                       Optional ByVal removeBorder As Boolean = False)
    If chromeRemoved Then Exit Sub

    RemoveUserFormChrome frm, removeBorder
    chromeRemoved = True
End Sub

