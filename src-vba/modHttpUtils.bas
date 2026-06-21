Attribute VB_Name = "modHttpUtils"
Option Explicit

' ============================================================
' modHttpUtils
'
' Canonical helpers for outbound HTTP request building, shared by
' all desktop HTTP clients (Google Auth, Google Sheets/Drive, SEF,
' future endpoints such as Viber Business API).
'
' Owns:
'   - UrlEncode(s)   RFC 3986 percent-encoding with UTF-8 byte expansion
'   - JsonEscape(s)  Basic JSON string escape (backslash, quote, newlines)
'
' Replaces the duplicated helpers:
'   - modGoogleAuth.UrlEncodeGoogle
'   - modGoogleSheets.JsonEscapeGoogle
'   - modSEFClient.UrlEncode (private)
'   - modSEFClient.JsonEscape (private)
'
' All four had the same shape and the same UTF-8 bug
' (Asc(ch) returns ANSI codepoint, not UTF-8 byte sequence).
' Serbian diacritics (c, c, s, z, dj, Dj) produced 400 Bad Request
' from SEF whenever they appeared in requestId or invoice fields.
' This module fixes the bug once for the whole desktop.
'
' Out of scope (intentional):
'   - JSON read/extract helpers (Google and SEF parsers differ;
'     consolidation will happen as part of the VBA-JSON migration
'     in audit P1.1, not now).
' ============================================================

' ============================================================
' PUBLIC API
' ============================================================

Public Function UrlEncode(ByVal s As String) As String
    ' RFC 3986 percent-encoding.
    ' Unreserved characters (A-Z a-z 0-9 - . _ ~) pass through unchanged.
    ' All other bytes are percent-encoded using their UTF-8 byte sequence.
    
    Dim bytes() As Byte
    Dim i As Long
    Dim b As Long
    Dim result As String
    
    If Len(s) = 0 Then
        UrlEncode = ""
        Exit Function
    End If
    
    bytes = Utf8Bytes(s)
    
    ' Defensive: empty array after BOM strip
    If UBound(bytes) < LBound(bytes) Then
        UrlEncode = ""
        Exit Function
    End If
    
    For i = LBound(bytes) To UBound(bytes)
        ' Widen Byte to Long with explicit unsigned mask.
        ' Guards against any sign-extension surprises.
        b = CLng(bytes(i)) And &HFF&
        
        Select Case b
            Case Asc("0") To Asc("9"), _
                 Asc("A") To Asc("Z"), _
                 Asc("a") To Asc("z")
                result = result & Chr$(b)
            Case Asc("-"), Asc("."), Asc("_"), Asc("~")
                result = result & Chr$(b)
            Case Else
                result = result & "%" & Right$("0" & Hex$(b), 2)
        End Select
    Next i
    
    UrlEncode = result
End Function

Public Function JsonEscape(ByVal s As String) As String
    ' Basic JSON string escape for embedding inside JSON string literal.
    ' Caller is responsible for surrounding quotes.
    '
    ' Note: this matches the pre-existing behavior of modSEFClient.JsonEscape
    ' and modGoogleSheets.JsonEscapeGoogle. It does not handle control
    ' characters below 0x20 individually; that is acceptable for current
    ' SEF/Google use because such characters do not appear in our payloads.
    
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    JsonEscape = t
End Function

' ============================================================
' PRIVATE HELPERS
' ============================================================

Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim stream As Object
    Dim bytes() As Byte
    Dim result() As Byte
    Dim i As Long
    Dim j As Long
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText s
    stream.Position = 0
    stream.Type = 1          ' adTypeBinary
    bytes = stream.Read
    stream.Close
    
    ' ADODB.Stream prepends UTF-8 BOM (EF BB BF) on text-mode writes.
    ' Strip it so the encoded output does not carry junk header bytes.
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            If UBound(bytes) = 2 Then
                Utf8Bytes = bytes
                Exit Function
            End If

            ReDim result(0 To UBound(bytes) - 3)
            j = 0
            For i = 3 To UBound(bytes)
                result(j) = bytes(i)
                j = j + 1
            Next i

            Utf8Bytes = result
            Exit Function
        End If
    End If
    
    Utf8Bytes = bytes
End Function
