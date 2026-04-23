Attribute VB_Name = "modStorno"
Option Explicit

' ============================================================
' modStorno v3.0 – Einfaches Soft-Delete
'
' Jedes Dokument wird einzeln storniert.
' Keine Kaskade zwischen Dokumenten.
' Ambalaza-Gegenbuchung wo relevant.
' Faktura: Prijemnice freigeben + Novac lösen.
' ============================================================

Private Const STORNO_DA As String = "Da"

' ============================================================
' OTKUP
' ============================================================

Public Function StornoOtkup_TX(ByVal otkupID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    
    StornoOtkup_TX = StornoOtkup(otkupID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoOtkup_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoOtkup_TX = False
End Function

Public Function StornoOtkup(ByVal otkupID As String) As Boolean
    If Not CanStorno(TBL_OTKUP, otkupID, COL_OTK_ID) Then
        StornoOtkup = False
        Exit Function
    End If
    
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    UpdateCell TBL_OTKUP, rows(1), COL_STORNIRANO, STORNO_DA
    StornoAmbalazaByDokument otkupID, DOK_TIP_OTKUP
    ResetNovacOtkupLink otkupID
    
    StornoOtkup = True
End Function

' ============================================================
' OTPREMNICA
' ============================================================

Public Function StornoOtpremnica_TX(ByVal otpremnicaID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    
    StornoOtpremnica_TX = StornoOtpremnica(otpremnicaID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoOtpremnica_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoOtpremnica_TX = False
End Function

Public Function StornoOtpremnica(ByVal otpremnicaID As String) As Boolean
    If Not CanStorno(TBL_OTPREMNICA, otpremnicaID, COL_OTP_ID) Then
        StornoOtpremnica = False
        Exit Function
    End If
    
    Dim rows As Collection
    Set rows = FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID)
    UpdateCell TBL_OTPREMNICA, rows(1), COL_STORNIRANO, STORNO_DA
    StornoAmbalazaByDokument otpremnicaID, DOK_TIP_OTPREMNICA
    
    StornoOtpremnica = True
End Function

' ============================================================
' ZBIRNA
' ============================================================

Public Function StornoZbirna_TX(ByVal brojZbirne As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    
    StornoZbirna_TX = StornoZbirna(brojZbirne)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoZbirna_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoZbirna_TX = False
End Function

Public Function StornoZbirna(ByVal brojZbirne As String) As Boolean
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        StornoZbirna = False
        Exit Function
    End If
    
    Dim colBroj As Long, colStorno As Long
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colStorno = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    
    Dim found As Boolean
    Dim i As Long
    For i = 1 To UBound(zbrData, 1)
        If CStr(zbrData(i, colBroj)) = brojZbirne Then
            If CStr(zbrData(i, colStorno)) = STORNO_DA Then GoTo NextZbr
            UpdateCell TBL_ZBIRNA, i, COL_STORNIRANO, STORNO_DA
            found = True
        End If
NextZbr:
    Next i
    
    If Not found Then
        MsgBox "Zbirna nije pronadena: " & brojZbirne, vbExclamation, APP_NAME
    End If
    
    StornoZbirna = found
End Function

' ============================================================
' PRIJEMNICA
' ============================================================

Public Function StornoPrijemnica_TX(ByVal prijemnicaID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    
    StornoPrijemnica_TX = StornoPrijemnica(prijemnicaID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoPrijemnica_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoPrijemnica_TX = False
End Function

Public Function StornoPrijemnica(ByVal prijemnicaID As String) As Boolean
    If Not CanStorno(TBL_PRIJEMNICA, prijemnicaID, COL_PRJ_ID) Then
        StornoPrijemnica = False
        Exit Function
    End If
    
    Dim rows As Collection
    Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)
    Dim r As Long: r = rows(1)
    
    UpdateCell TBL_PRIJEMNICA, r, COL_STORNIRANO, STORNO_DA
    
    ' Fakturisano-Flag lösen (Faktura bleibt!)
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If CStr(prijData(r, GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO))) = "Da" Then
        UpdateCell TBL_PRIJEMNICA, r, COL_PRJ_FAKTURISANO, ""
        
        Dim fakturaID As String
        fakturaID = CStr(prijData(r, GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)))
        UpdateCell TBL_PRIJEMNICA, r, COL_PRJ_FAKTURA_ID, ""
        
        If fakturaID <> "" Then
            ' Faktura als verwaist markieren
            Dim fakRows As Collection
            Set fakRows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)
            If fakRows.count > 0 Then
                UpdateCell TBL_FAKTURE, fakRows(1), COL_OSIROCENO_OD, prijemnicaID
            End If
            
            ' FakturaStavke als verwaist markieren  ? NEU
            Dim stavkeData As Variant
            stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
            If Not IsEmpty(stavkeData) Then
                Dim colSPrijID As Long, colSFakID As Long
                colSPrijID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID)
                colSFakID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID)
                
                Dim j As Long
                For j = 1 To UBound(stavkeData, 1)
                    If CStr(stavkeData(j, colSPrijID)) = prijemnicaID And _
                       CStr(stavkeData(j, colSFakID)) = fakturaID Then
                        UpdateCell TBL_FAKTURA_STAVKE, j, COL_OSIROCENO_OD, prijemnicaID
                    End If
                Next j
            End If
        End If
    End If
    
    StornoAmbalazaByDokument prijemnicaID, DOK_TIP_PRIJEMNICA
    
    StornoPrijemnica = True
End Function

' ============================================================
' FAKTURA
' ============================================================

Public Function StornoFaktura_TX(ByVal fakturaID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_NOVAC
    
    StornoFaktura_TX = StornoFaktura(fakturaID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoFaktura_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoFaktura_TX = False
End Function

Public Function StornoFaktura(ByVal fakturaID As String) As Boolean
    If Not CanStorno(TBL_FAKTURE, fakturaID, COL_FAK_ID) Then
        StornoFaktura = False
        Exit Function
    End If
    
    Dim fakRows As Collection
    Set fakRows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)
    UpdateCell TBL_FAKTURE, fakRows(1), COL_STORNIRANO, STORNO_DA
    UpdateCell TBL_FAKTURE, fakRows(1), COL_FAK_STATUS, "Stornirano"
    
    ' Stavke stornieren + Prijemnice freigeben
    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If Not IsEmpty(stavkeData) Then
        Dim colFakID As Long, colPrijID As Long
        colFakID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID)
        colPrijID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID)
        
        Dim i As Long
        For i = 1 To UBound(stavkeData, 1)
            If CStr(stavkeData(i, colFakID)) = fakturaID Then
                UpdateCell TBL_FAKTURA_STAVKE, i, COL_STORNIRANO, STORNO_DA
                
                Dim prijID As String
                prijID = CStr(stavkeData(i, colPrijID))
                Dim prijRows As Collection
                Set prijRows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijID)
                If prijRows.count > 0 Then
                    UpdateCell TBL_PRIJEMNICA, prijRows(1), COL_PRJ_FAKTURISANO, ""
                    UpdateCell TBL_PRIJEMNICA, prijRows(1), COL_PRJ_FAKTURA_ID, ""
                End If
            End If
        Next i
    End If
    
    ' Novac: FakturaID lösen
    ResetNovacFakturaLink fakturaID
    
    StornoFaktura = True
End Function

' ============================================================
' NOVAC
' ============================================================

Public Function StornoNovac_TX(ByVal novacID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    
    StornoNovac_TX = StornoNovac(novacID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "StornoNovac_TX"
    tx.RollbackTx
    MsgBox "Greska, promene vracene: " & Err.Description, vbCritical, APP_NAME
    StornoNovac_TX = False
End Function

Public Function StornoNovac(ByVal novacID As String) As Boolean
    If Not CanStorno(TBL_NOVAC, novacID, COL_NOV_ID) Then
        StornoNovac = False
        Exit Function
    End If
    
    Dim rows As Collection
    Set rows = FindRows(TBL_NOVAC, COL_NOV_ID, novacID)
    Dim r As Long: r = rows(1)
    
    UpdateCell TBL_NOVAC, r, COL_STORNIRANO, STORNO_DA
    
    ' Faktura-Status neu berechnen
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    Dim fakturaID As String
    fakturaID = CStr(novData(r, GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)))
    If fakturaID <> "" Then
        ResetFakturaStatus fakturaID
    End If
    
    StornoNovac = True
End Function

' ============================================================
' PRIVATE HELPERS
' ============================================================

Private Sub ResetNovacFakturaLink(ByVal fakturaID As String)
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    If IsEmpty(novData) Then Exit Sub
    
    Dim colFakID As Long
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    
    Dim i As Long
    For i = 1 To UBound(novData, 1)
        If CStr(novData(i, colFakID)) = fakturaID Then
            UpdateCell TBL_NOVAC, i, COL_NOV_FAKTURA_ID, ""
        End If
    Next i
End Sub

Private Sub ResetFakturaStatus(ByVal fakturaID As String)
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    If IsEmpty(novData) Then Exit Sub
    
    Dim colFakID As Long, colUplata As Long, colStorno As Long
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    colStorno = GetColumnIndex(TBL_NOVAC, COL_STORNIRANO)
    
    Dim uplaceno As Double
    Dim i As Long
    For i = 1 To UBound(novData, 1)
        If CStr(novData(i, colFakID)) = fakturaID Then
            If CStr(novData(i, colStorno)) = STORNO_DA Then GoTo NextNov
            If IsNumeric(novData(i, colUplata)) Then
                uplaceno = uplaceno + CDbl(novData(i, colUplata))
            End If
        End If
NextNov:
    Next i
    
    Dim fakIznos As Double
    fakIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS))
    
    Dim fakRows As Collection
    Set fakRows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)
    If fakRows.count > 0 Then
        If uplaceno >= fakIznos Then
            UpdateCell TBL_FAKTURE, fakRows(1), COL_FAK_STATUS, STATUS_PLACENO
        Else
            UpdateCell TBL_FAKTURE, fakRows(1), COL_FAK_STATUS, STATUS_NEPLACENO
            UpdateCell TBL_FAKTURE, fakRows(1), COL_FAK_DATUM_PLACANJA, ""
        End If
    End If
End Sub

Private Sub StornoAmbalazaByDokument(ByVal dokumentID As String, _
                                      ByVal dokumentTip As String)
    Dim ambData As Variant
    ambData = GetTableData(TBL_AMBALAZA)
    If IsEmpty(ambData) Then Exit Sub
    
    Dim colDokID As Long, colDokTip As Long, colStorno As Long
    colDokID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    colDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    colStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If colStorno = 0 Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(ambData, 1)
        If CStr(ambData(i, colDokID)) = dokumentID And _
           CStr(ambData(i, colDokTip)) = dokumentTip Then
            If CStr(ambData(i, colStorno)) <> STORNO_DA Then
                UpdateCell TBL_AMBALAZA, i, COL_STORNIRANO, STORNO_DA
            End If
        End If
    Next i
End Sub
Public Function CanStorno(ByVal tblName As String, _
                          ByVal recordID As String, _
                          ByVal idColumn As String) As Boolean
    Dim rows As Collection
    Set rows = FindRows(tblName, idColumn, recordID)
    If rows.count = 0 Then
        MsgBox "Stavka nije pronadjena!", vbExclamation, APP_NAME
        CanStorno = False
        Exit Function
    End If
    
    Dim data As Variant
    data = GetTableData(tblName)
    Dim colStorno As Long
    colStorno = GetColumnIndex(tblName, COL_STORNIRANO)
    
    If colStorno > 0 Then
        If CStr(data(rows(1), colStorno)) = STORNO_DA Then
            MsgBox "Vec stornirano!", vbExclamation, APP_NAME
            CanStorno = False
            Exit Function
        End If
    End If
    
    CanStorno = True
End Function

Public Function LookupActiveID(ByVal tblName As String, _
                                ByVal brojColName As String, _
                                ByVal brojValue As String, _
                                ByVal idColName As String) As String
    ' Wie LookupValue, aber überspringt Stornirano="Da"
    ' Findet den LETZTEN nicht-stornierten Treffer
    
    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then
        LookupActiveID = ""
        Exit Function
    End If
    
    Dim colBroj As Long, colID As Long, colStorno As Long
    colBroj = GetColumnIndex(tblName, brojColName)
    colID = GetColumnIndex(tblName, idColName)
    colStorno = GetColumnIndex(tblName, COL_STORNIRANO)
    
    Dim resultID As String
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colBroj)) = brojValue Then
            If colStorno > 0 Then
                If CStr(data(i, colStorno)) = "Da" Then GoTo NextRow
            End If
            resultID = CStr(data(i, colID))
        End If
NextRow:
    Next i
    
    LookupActiveID = resultID
End Function

