Attribute VB_Name = "modDrive"
Option Explicit

' ============================================================
' modDrive - tanki Drive REST helperi (download / find / upload).
'
' Reuse: GetAccessToken (modGoogleAuth), UrlEncode (modHttpUtils),
'        ExtractJsonStringGoogle (modGoogleAuth), LogErr (modLogError).
'
' Koriste ga:
'   - self-update (povlacenje koda iz AgriX_Release) - modSelfUpdate
'   - PublishReleaseToDrive (build) - modRelease
'
' SVE binarno (ADODB.Stream + responseBody/Send bytes) -> cuva izvorne
' bajtove .bas/.cls/.frm (Windows-1250) bez transkodiranja.
' ASCII-only modul (vidi CLAUDE.md, sekcija 4 - encoding).
' ============================================================

Private Const DRIVE_FILES As String = "https://www.googleapis.com/drive/v3/files"
Private Const DRIVE_UPLOAD As String = "https://www.googleapis.com/upload/drive/v3/files"

' Skini fajl sa Drive-a (alt=media) i snimi kao sirove bajtove na destPath.
Public Function DriveDownloadToFile(ByVal fileID As String, ByVal destPath As String) As Boolean
    Const SRC As String = "modDrive.DriveDownloadToFile"
    On Error GoTo EH

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "GET", DRIVE_FILES & "/" & fileID & "?alt=media&supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.Send

    If http.status <> 200 Then
        LogErr SRC, "HTTP " & http.status & " za " & fileID
        Exit Function
    End If

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                 ' adTypeBinary
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile destPath, 2   ' adSaveCreateOverWrite
    stm.Close

    DriveDownloadToFile = True
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Vrati fileID istoimenog fajla u datom folderu, ili "" ako ga nema.
Public Function DriveFindInFolder(ByVal folderID As String, ByVal fileName As String) As String
    Const SRC As String = "modDrive.DriveFindInFolder"
    On Error GoTo EH

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    Dim q As String
    q = "name='" & DriveEscapeQ(fileName) & "' and '" & folderID & "' in parents and trashed=false"

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "GET", DRIVE_FILES & "?q=" & UrlEncode(q) & _
              "&fields=files(id,name)&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.Send

    If http.status = 200 Then
        DriveFindInFolder = ExtractJsonStringGoogle(http.responseText, "id")
    Else
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
    End If
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Upload fajla u folder (create-or-update po imenu). Vrati fileID ili "".
' Dvostepeno (metadata POST -> media PATCH) da se izbegne multipart boundary.
Public Function DriveUploadFile(ByVal folderID As String, _
                                ByVal localPath As String, _
                                ByVal fileName As String) As String
    Const SRC As String = "modDrive.DriveUploadFile"
    On Error GoTo EH

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    Dim fileID As String: fileID = DriveFindInFolder(folderID, fileName)
    If Len(fileID) = 0 Then
        fileID = DriveCreateEmpty(folderID, fileName, token)
        If Len(fileID) = 0 Then Exit Function
    End If

    Dim bytes() As Byte: bytes = DriveLoadFileBytes(localPath)

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "PATCH", DRIVE_UPLOAD & "/" & fileID & "?uploadType=media&supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/octet-stream"
    http.Send bytes

    If http.status = 200 Then
        DriveUploadFile = fileID
    Else
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
    End If
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Premesti fajl u Trash (trashed=true). Koristi PublishReleaseToDrive za ciscenje
' zastarelih (obrisanih) code fajlova iz AgriX_Release-a. Meko brisanje (Trash),
' ne hard-delete -> greska pri objavi ne unistava fajl trajno. Vrati True na uspeh.
Public Function DriveTrashFile(ByVal fileID As String) As Boolean
    Const SRC As String = "modDrive.DriveTrashFile"
    On Error GoTo EH

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "PATCH", DRIVE_FILES & "/" & fileID & "?supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{""trashed"":true}"

    If http.status = 200 Then
        DriveTrashFile = True
    Else
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
    End If
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' SHA-256 (hex, lowercase) SIROVIH BAJTOVA fajla, ili "" ako SHA nedostupan/greska.
' Isti .NET primitiv kao modAuth.Sha256Hex (dokazan u produkciji za PIN hash), samo
' nad bajtovima fajla (ADODB.Stream binarno, isti obrazac kao DriveDownloadToFile).
' Koristi ga: PublishReleaseToDrive (manifest) + modSelfUpdate (verifikacija skinutog).
Public Function Sha256File(ByVal path As String) As String
    On Error GoTo EH
    Dim stm As Object, sha As Object, bytes() As Byte, hash() As Byte
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    bytes = stm.Read                   ' svi bajtovi (adReadAll)
    stm.Close

    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha.ComputeHash_2((bytes))

    Dim i As Long, s As String
    For i = LBound(hash) To UBound(hash)
        s = s & Right$("0" & Hex$(hash(i) And &HFF), 2)
    Next i
    Sha256File = LCase$(s)
    Exit Function
EH:
    Sha256File = vbNullString          ' "" = SHA nedostupan / greska -> pozivalac odlucuje
End Function

' F0 self-test (RUCNO, Alt+F8): potvrdi da Sha256File radi na OVOJ masini. Hesuje
' fajl sa "abc" (3 bajta) i poredi sa standardnim vektorom SHA-256("abc").
Public Sub Test_Sha256File()
    On Error Resume Next
    Dim p As String: p = Environ$("TEMP") & "\AgriX_shatest.txt"
    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff
    Print #ff, "abc";                  ' tacno 3 bajta (bez CRLF/BOM)
    Close #ff
    Dim got As String: got = Sha256File(p)
    Kill p
    Const WANT As String = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    If LCase$(got) = WANT Then
        MsgBox "SHA-256 RADI na ovoj masini." & vbCrLf & "Sha256File(""abc"") = " & got, _
               vbInformation, APP_NAME
    Else
        MsgBox "SHA-256 NE radi / nesklad!" & vbCrLf & "Dobijeno: '" & got & "'" & vbCrLf & _
               "Ocekivano: " & WANT & vbCrLf & vbCrLf & _
               "Release verifikacija ce pasti na fallback (broj/prisustvo).", _
               vbExclamation, APP_NAME
    End If
End Sub

' Nadji podfolder po imenu u parent-u; ako ne postoji, kreiraj ga. Vrati folderID
' (ili "" na gresku). Za versioned release layout (AgriX_Release/releases/<verzija>/).
Public Function DriveEnsureFolder(ByVal parentID As String, ByVal name As String) As String
    Const SRC As String = "modDrive.DriveEnsureFolder"
    On Error GoTo EH

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    ' 1) find (mimeType = folder)
    Dim q As String
    q = "name='" & DriveEscapeQ(name) & "' and '" & parentID & "' in parents and " & _
        "mimeType='application/vnd.google-apps.folder' and trashed=false"
    Dim http As Object: Set http = DriveNewHttp()
    http.Open "GET", DRIVE_FILES & "?q=" & UrlEncode(q) & _
              "&fields=files(id,name)&pageSize=1&supportsAllDrives=true&includeItemsFromAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.Send
    If http.status = 200 Then
        Dim fid As String: fid = ExtractJsonStringGoogle(http.responseText, "id")
        If Len(fid) > 0 Then DriveEnsureFolder = fid: Exit Function
    End If

    ' 2) create
    Dim body As String
    body = "{""name"":""" & DriveJsonEsc(name) & """,""mimeType"":""application/vnd.google-apps.folder""," & _
           """parents"":[""" & parentID & """]}"
    Set http = DriveNewHttp()
    http.Open "POST", DRIVE_FILES & "?supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.Send body
    If http.status = 200 Then
        DriveEnsureFolder = ExtractJsonStringGoogle(http.responseText, "id")
    Else
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
    End If
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Vrati Dictionary (ime fajla -> fileID) svih fajlova u folderu.
' Jedna strana (pageSize=1000); AgriX_Release ima ~100 src-vba fajlova.
Public Function DriveListFolder(ByVal folderID As String) As Object
    Const SRC As String = "modDrive.DriveListFolder"
    On Error GoTo EH

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Set DriveListFolder = dict

    Dim token As String: token = GetAccessToken()
    If Len(token) = 0 Then Exit Function

    Dim q As String
    q = "'" & folderID & "' in parents and trashed=false"

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "GET", DRIVE_FILES & "?q=" & UrlEncode(q) & _
              "&fields=files(id,name)&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.Send

    If http.status <> 200 Then
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
        Exit Function
    End If

    ' Parsiraj files(id,name) - po jedan objekat izmedju "}" granica
    ' (isti pristup kao ExtractJson* u ostatku koda; bez JSON biblioteke).
    Dim parts() As String, i As Long, id As String, nm As String
    parts = Split(http.responseText, "}")
    For i = LBound(parts) To UBound(parts)
        id = ExtractJsonStringGoogle(parts(i), "id")
        nm = ExtractJsonStringGoogle(parts(i), "name")
        If Len(id) > 0 And Len(nm) > 0 Then
            If Not dict.Exists(nm) Then dict.Add nm, id
        End If
    Next i
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Dijagnostika (RUCNO, Alt+F8): pokazuje zasto upload pada - duzinu tokena,
' status LIST GET-a (citanje) i status + Google error JSON za CREATE POST (upis).
Public Sub DriveSelfTest()
    Const SRC As String = "modDrive.DriveSelfTest"
    On Error GoTo EH

    Dim msg As String, token As String
    token = GetAccessToken()
    msg = "REL_FOLDER_ID: " & REL_FOLDER_ID & vbCrLf & _
          "Token duzina:  " & Len(token) & vbCrLf & vbCrLf

    If Len(token) = 0 Then
        MsgBox msg & "GetAccessToken() je PRAZAN -> Google nije autentifikovan na OVOM fajlu." & vbCrLf & _
               "Uradi Google login/sync u ovom fajlu (kao za obican sync) pa ponovi.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    ' 1) LIST (GET) - isti obrazac kao postojeci radni Drive read
    Dim http As Object, q As String
    q = "'" & REL_FOLDER_ID & "' in parents and trashed=false"
    Set http = DriveNewHttp()
    http.Open "GET", DRIVE_FILES & "?q=" & UrlEncode(q) & _
              "&fields=files(id,name)&pageSize=5&supportsAllDrives=true&includeItemsFromAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.Send
    msg = msg & "LIST (GET) status: " & http.status & vbCrLf & _
          "  " & Left$(http.responseText, 250) & vbCrLf & vbCrLf

    ' 2) CREATE (POST metadata) - novi upload pipeline; Google error JSON kaze sve
    Dim body As String
    body = "{""name"":""agrix_selftest.txt"",""parents"":[""" & REL_FOLDER_ID & """]}"
    Set http = DriveNewHttp()
    http.Open "POST", DRIVE_FILES & "?supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.Send body
    msg = msg & "CREATE (POST) status: " & http.status & vbCrLf & _
          "  " & Left$(http.responseText, 450)

    MsgBox msg, vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "DriveSelfTest greska: " & Err.description, vbCritical, APP_NAME
End Sub

' ---------------- private ----------------

' Kreiraj prazan fajl (samo metadata) u folderu -> vrati novi fileID.
Private Function DriveCreateEmpty(ByVal folderID As String, _
                                  ByVal fileName As String, _
                                  ByVal token As String) As String
    Const SRC As String = "modDrive.DriveCreateEmpty"
    On Error GoTo EH

    Dim body As String
    body = "{""name"":""" & DriveJsonEsc(fileName) & """,""parents"":[""" & folderID & """]}"

    Dim http As Object: Set http = DriveNewHttp()
    http.Open "POST", DRIVE_FILES & "?supportsAllDrives=true", False
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.Send body

    If http.status = 200 Then
        DriveCreateEmpty = ExtractJsonStringGoogle(http.responseText, "id")
    Else
        LogErr SRC, "HTTP " & http.status & ": " & Left$(http.responseText, 300)
    End If
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

Private Function DriveNewHttp() As Object
    Set DriveNewHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' resolve/connect kratko (da offline ne blokira Workbook_Open self-update
    ' proveru); send/receive dugo (download .xlsm/koda moze potrajati).
    DriveNewHttp.SetTimeouts 5000, 5000, 30000, 120000
End Function

Private Function DriveLoadFileBytes(ByVal path As String) As Byte()
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                 ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    ' adReadAll na praznom stream-u vraca Null -> Type mismatch u Byte(); na 0
    ' bajtova ostavi neinicijalizovan Byte() (validan prazan sadrzaj).
    If stm.size > 0 Then DriveLoadFileBytes = stm.Read
    stm.Close
End Function

' Drive query escape: ' -> \'
Private Function DriveEscapeQ(ByVal s As String) As String
    DriveEscapeQ = Replace$(s, "'", "\'")
End Function

Private Function DriveJsonEsc(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace$(t, "\", "\\")
    t = Replace$(t, """", "\""")
    DriveJsonEsc = t
End Function
