Attribute VB_Name = "modKooperant"
'Attribute VB_Name = "modKooperant"
Option Explicit

' ============================================================
' modKooperant – razresavanje/kreiranje kooperanta po imenu.
'
' frmOtkup.cmbKooperant je sada 2-kolonski: kol0 = "Ime Prezime"
' (vidljivo, filtrira se po imenu), kol1 = KooperantID (skriveno).
' Kada operater ukuca ime koje nije u bazi, kreira se novi
' tblKooperanti red (samo Ime/Prezime + StanicaID; BPG/racun se
' dopunjavaju kasnije u Maticnim podacima). Bez odmah sync-a –
' PWA/Google ce povuci pri sledecoj sinhronizaciji.
' ============================================================

' Vrati KooperantID za trenutni izbor/unos u combo-u:
'   izabran red iz liste -> bound ID (kol1);
'   slobodan tekst       -> nadji po imenu (prioritet ista stanica),
'                           inace kreiraj novog i vrati novi ID.
Public Function ResolveKooperantByName(ByVal cmb As MSForms.ComboBox, _
                                       ByVal stanicaID As String) As String
    On Error GoTo EH

    Dim boundID As String: boundID = GetComboID(cmb)
    If Len(Trim$(boundID)) > 0 Then
        ResolveKooperantByName = boundID
        Exit Function
    End If

    Dim nm As String: nm = Trim$(cmb.value)
    If Len(nm) = 0 Then Exit Function

    Dim existing As String: existing = FindKooperantIDByName(nm, stanicaID)
    If Len(existing) > 0 Then
        ResolveKooperantByName = existing
        Exit Function
    End If

    ' Toggle (Podesavanja): auto-kreiranje nepoznatog kooperanta. OFF -> vrati ""
    ' pa frmOtkup javlja da kooperant nije pronadjen (bez tihog kreiranja).
    If Not KoopAutoCreate() Then Exit Function

    ResolveKooperantByName = CreateKooperantByName(nm, stanicaID)
    Exit Function
EH:
    LogErr "modKooperant.ResolveKooperantByName"
End Function

' Nadji KooperantID po punom imenu (Ime + Prezime), case-insensitive.
' Prioritet: isto otkupno mesto; fallback bilo koje.
Private Function FindKooperantIDByName(ByVal nm As String, ByVal stanicaID As String) As String
    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cIme As Long, cPr As Long, cSt As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    cSt = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_STANICA)

    Dim want As String: want = LCase$(Trim$(nm))
    Dim i As Long, Fallback As String
    For i = 1 To UBound(data, 1)
        Dim full As String
        full = LCase$(Trim$(Trim$(CStr(data(i, cIme))) & " " & Trim$(CStr(data(i, cPr)))))
        If full = want Then
            If CStr(data(i, cSt)) = stanicaID Then
                FindKooperantIDByName = CStr(data(i, cId))
                Exit Function
            ElseIf Len(Fallback) = 0 Then
                Fallback = CStr(data(i, cId))
            End If
        End If
    Next i
    FindKooperantIDByName = Fallback
End Function

' Kreiraj novi tblKooperanti red: samo Ime/Prezime + StanicaID + novi ID.
Private Function CreateKooperantByName(ByVal nm As String, ByVal stanicaID As String) As String
    Dim tx As clsTransaction
    On Error GoTo EH

    Dim ime As String, prezime As String
    SplitName nm, ime, prezime
    If Len(ime) = 0 Then Exit Function

    Dim newID As String: newID = GetNextID(TBL_KOOPERANTI, COL_KOOP_ID, "KOOP-")
    If Len(newID) = 0 Then
        Err.Raise vbObjectError + 3010, "modKooperant.CreateKooperantByName", _
                  "GetNextID nije vratio KooperantID."
    End If

    Dim lo As ListObject: Set lo = GetTable(TBL_KOOPERANTI)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 3011, "modKooperant.CreateKooperantByName", _
                  "tblKooperanti nije pronadjena."
    End If

    ' Red u redosledu kolona (sve prazno osim ID/Ime/Prezime/Stanica/Aktivan).
    Dim nCols As Long: nCols = lo.ListColumns.count
    Dim row() As Variant: ReDim row(1 To nCols)
    Dim c As Long
    For c = 1 To nCols: row(c) = "": Next c
    row(GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)) = newID
    row(GetColumnIndex(TBL_KOOPERANTI, "Ime")) = ime
    row(GetColumnIndex(TBL_KOOPERANTI, "Prezime")) = prezime
    row(GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_STANICA)) = stanicaID
    row(GetColumnIndex(TBL_KOOPERANTI, "Aktivan")) = STATUS_AKTIVAN

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_KOOPERANTI

    If AppendRow(TBL_KOOPERANTI, row) <= 0 Then
        Err.Raise vbObjectError + 3012, "modKooperant.CreateKooperantByName", _
                  "AppendRow nije uspeo za tblKooperanti."
    End If

    tx.CommitTx
    Set tx = Nothing

    CreateKooperantByName = newID
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modKooperant.CreateKooperantByName"
    MsgBox "Greska pri kreiranju kooperanta: " & Err.description, vbCritical, APP_NAME
End Function

' "Marko MARKOVIC" -> Ime="Marko", Prezime="MARKOVIC"
' "Petar Petar Petrovic" -> Ime="Petar", Prezime="Petar Petrovic"
' "Marko" -> Ime="Marko", Prezime=""
Private Sub SplitName(ByVal nm As String, ByRef ime As String, ByRef prezime As String)
    nm = Trim$(nm)
    Dim p As Long: p = InStr(nm, " ")
    If p > 0 Then
        ime = Trim$(Left$(nm, p - 1))
        prezime = Trim$(Mid$(nm, p + 1))
    Else
        ime = nm
        prezime = ""
    End If
End Sub

