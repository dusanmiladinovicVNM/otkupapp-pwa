Attribute VB_Name = "modClipboard"
'MOD Clipboard
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As LongPtr, ByVal lpString2 As String) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
    Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
    Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const GMEM_MOVEABLE As Long = &H2
Private Const CF_TEXT As Long = 1

Public Sub CopyToClipboard(ByVal txt As String)

    Dim hGlobal As LongPtr
    Dim lpGlobal As LongPtr

    hGlobal = GlobalAlloc(GMEM_MOVEABLE, Len(txt) + 1)
    lpGlobal = GlobalLock(hGlobal)

    lstrcpy lpGlobal, txt

    GlobalUnlock hGlobal

    OpenClipboard 0
    EmptyClipboard
    SetClipboardData CF_TEXT, hGlobal
    CloseClipboard
End Sub

Public Function GetClipboardText() As String
    
    Dim hData As LongPtr
    Dim lpData As LongPtr
    Dim txtLen As Long
    Dim buffer() As Byte
    
    GetClipboardText = vbNullString
    
    If OpenClipboard(0) = 0 Then Exit Function
    
    If IsClipboardFormatAvailable(CF_TEXT) = 0 Then
        CloseClipboard
        Exit Function
    End If
    
    hData = GetClipboardData(CF_TEXT)
    If hData = 0 Then
        CloseClipboard
        Exit Function
    End If
    
    lpData = GlobalLock(hData)
    If lpData = 0 Then
        CloseClipboard
        Exit Function
    End If
    
    txtLen = lstrlenA(lpData)
    If txtLen > 0 Then
        ReDim buffer(0 To txtLen - 1) As Byte
        RtlMoveMemory buffer(0), lpData, txtLen
        GetClipboardText = StrConv(buffer, vbUnicode)
    End If
    
    GlobalUnlock hData
    CloseClipboard

End Function
