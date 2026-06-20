'Attribute VB_Name = "modStorno"
Option Explicit

' ============================================================
' modStorno v4.0 – Hardened Soft-Delete
'
' Stil uskladjen sa modNovac/modFaktura:
' - fail-fast schema guards
' - RequireColumnIndex / RequireUpdateCell
' - stroga provera single-row dokumenata
' - transakcioni rollback u *_TX wrapperima
' - monitoring success/fail eventa
' - bez MsgBox u business sloju
'
' Business pravila:
' - Svaki dokument se stornira pojedinacno.
' - Nema automatske kaskade izmedju dokumenata.
' - Ambalaza se stornira za dokument gde postoji.
' - Faktura: stavke se storniraju, prijemnice se oslobadjaju,
'   novac se odvezuje od fakture.
' - Prijemnica: ako je bila fakturisana, oslobadja se i faktura/stavke
'   se oznacavaju kao osirocene.
' ============================================================

Private Const MOD_NAME As String = "modStorno"
Private Const STORNO_DA As String = "Da"
Private Const STATUS_STORNIRANO As String = "Stornirano"
Private Const ERR_STORNO_BASE As Long = vbObjectError + 2400

' ============================================================
' OTKUP
' ============================================================

Public Function StornoOtkup_TX(ByVal otkupID As String) As Boolean
    Const SRC As String = "StornoOtkup_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    If Not StornoOtkup(otkupID) Then
        Err.Raise ERR_STORNO_BASE + 1, SRC, _
                  "StornoOtkup nije uspeo. OtkupID=" & otkupID
    End If

    tx.CommitTx

    StornoOtkup_TX = True
    MonitorStornoSuccess SRC, "Otkup", otkupID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Otkup", otkupID, tx
    StornoOtkup_TX = False
End Function

Public Function StornoOtkup(ByVal otkupID As String) As Boolean
    Const SRC As String = "StornoOtkup"

    On Error GoTo EH

    Dim rowOtkup As Long
    rowOtkup = RequireStornoAllowed(TBL_OTKUP, otkupID, COL_OTK_ID, SRC)

    MarkRowStornirano TBL_OTKUP, rowOtkup, SRC
    StornoAmbalazaByDokument otkupID, DOK_TIP_OTKUP
    StornoAmbalazaByDokument otkupID, DOK_TIP_OM_IZLAZ_KOOP   ' izdata ambalaza (OM->kooperant) uz otkup
    ResetNovacOtkupLink otkupID

    StornoOtkup = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' OTPREMNICA
' ============================================================

Public Function StornoOtpremnica_TX(ByVal otpremnicaID As String) As Boolean
    Const SRC As String = "StornoOtpremnica_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    If Not StornoOtpremnica(otpremnicaID) Then
        Err.Raise ERR_STORNO_BASE + 2, SRC, _
                  "StornoOtpremnica nije uspeo. OtpremnicaID=" & otpremnicaID
    End If

    tx.CommitTx

    StornoOtpremnica_TX = True
    MonitorStornoSuccess SRC, "Otpremnica", otpremnicaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Otpremnica", otpremnicaID, tx
    StornoOtpremnica_TX = False
End Function

Public Function StornoOtpremnica(ByVal otpremnicaID As String) As Boolean
    Const SRC As String = "StornoOtpremnica"

    On Error GoTo EH

    Dim rowOtp As Long
    rowOtp = RequireStornoAllowed(TBL_OTPREMNICA, otpremnicaID, COL_OTP_ID, SRC)

    MarkRowStornirano TBL_OTPREMNICA, rowOtp, SRC
    StornoAmbalazaByDokument otpremnicaID, DOK_TIP_OTPREMNICA

    StornoOtpremnica = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' ZBIRNA
' ============================================================

Public Function StornoZbirna_TX(ByVal brojZbirne As String) As Boolean
    Const SRC As String = "StornoZbirna_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    If Not StornoZbirna(brojZbirne) Then
        Err.Raise ERR_STORNO_BASE + 3, SRC, _
                  "StornoZbirna nije uspeo. BrojZbirne=" & brojZbirne
    End If

    tx.CommitTx

    StornoZbirna_TX = True
    MonitorStornoSuccess SRC, "Zbirna", brojZbirne

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Zbirna", brojZbirne, tx
    StornoZbirna_TX = False
End Function

Public Function StornoZbirna(ByVal brojZbirne As String) As Boolean
    Const SRC As String = "StornoZbirna"

    On Error GoTo EH

    RequireNonBlank brojZbirne, "BrojZbirne", SRC

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_BASE + 20, SRC, _
                  "Tabela je prazna: " & TBL_ZBIRNA
    End If

    Dim colBroj As Long
    Dim colStorno As Long

    colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    colStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, SRC)

    Dim foundAny As Boolean
    Dim changedCount As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colBroj))) = Trim$(brojZbirne) Then
            foundAny = True

            If Not IsStorniranoValue(data(i, colStorno)) Then
                MarkRowStornirano TBL_ZBIRNA, i, SRC
                changedCount = changedCount + 1
            End If
        End If
    Next i

    If Not foundAny Then
        Err.Raise ERR_STORNO_BASE + 21, SRC, _
                  "Zbirna nije pronadjena. BrojZbirne=" & brojZbirne
    End If

    If changedCount = 0 Then
        Err.Raise ERR_STORNO_BASE + 22, SRC, _
                  "Zbirna je vec stornirana. BrojZbirne=" & brojZbirne
    End If

    StornoZbirna = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PRIJEMNICA
' ============================================================

Public Function StornoPrijemnica_TX(ByVal prijemnicaID As String) As Boolean
    Const SRC As String = "StornoPrijemnica_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE

    If Not StornoPrijemnica(prijemnicaID) Then
        Err.Raise ERR_STORNO_BASE + 4, SRC, _
                  "StornoPrijemnica nije uspeo. PrijemnicaID=" & prijemnicaID
    End If

    tx.CommitTx

    StornoPrijemnica_TX = True
    MonitorStornoSuccess SRC, "Prijemnica", prijemnicaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Prijemnica", prijemnicaID, tx
    StornoPrijemnica_TX = False
End Function

Public Function StornoPrijemnica(ByVal prijemnicaID As String) As Boolean
    Const SRC As String = "StornoPrijemnica"

    On Error GoTo EH

    Dim rowPrij As Long
    rowPrij = RequireStornoAllowed(TBL_PRIJEMNICA, prijemnicaID, COL_PRJ_ID, SRC)

    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, SRC
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, SRC
    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, SRC
    RequireColumnIndex TBL_FAKTURE, COL_OSIROCENO_OD, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, SRC

    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(prijData) Then
        Err.Raise ERR_STORNO_BASE + 30, SRC, _
                  "Tabela prijemnica je prazna."
    End If

    Dim colFakturisano As Long
    Dim colFakturaID As Long
    Dim fakturaID As String

    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, SRC)
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, SRC)

    fakturaID = Trim$(CStr(prijData(rowPrij, colFakturaID)))

    MarkRowStornirano TBL_PRIJEMNICA, rowPrij, SRC

    If UCase$(Trim$(CStr(prijData(rowPrij, colFakturisano)))) = "DA" Then
        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURISANO, "", SRC
        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURA_ID, "", SRC

        If Len(fakturaID) > 0 Then
            MarkFakturaOrphaned fakturaID, prijemnicaID
            MarkFakturaStavkeOrphaned fakturaID, prijemnicaID
        End If
    End If

    StornoAmbalazaByDokument prijemnicaID, DOK_TIP_PRIJEMNICA

    StornoPrijemnica = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' FAKTURA
' ============================================================

Public Function StornoFaktura_TX(ByVal fakturaID As String) As Boolean
    Const SRC As String = "StornoFaktura_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_NOVAC

    If Not StornoFaktura(fakturaID) Then
        Err.Raise ERR_STORNO_BASE + 5, SRC, _
                  "StornoFaktura nije uspeo. FakturaID=" & fakturaID
    End If

    tx.CommitTx

    StornoFaktura_TX = True
    MonitorStornoSuccess SRC, "Faktura", fakturaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Faktura", fakturaID, tx
    StornoFaktura_TX = False
End Function

Public Function StornoFaktura(ByVal fakturaID As String) As Boolean
    Const SRC As String = "StornoFaktura"

    On Error GoTo EH

    Dim rowFak As Long
    rowFak = RequireStornoAllowed(TBL_FAKTURE, fakturaID, COL_FAK_ID, SRC)

    RequireColumnIndex TBL_FAKTURE, COL_FAK_STATUS, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_STORNIRANO, SRC
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_ID, SRC
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, SRC
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, SRC

    MarkRowStornirano TBL_FAKTURE, rowFak, SRC
    RequireUpdateCell TBL_FAKTURE, rowFak, COL_FAK_STATUS, STATUS_STORNIRANO, SRC

    StornoFakturaStavkeAndReleasePrijemnice fakturaID
    ResetNovacFakturaLink fakturaID

    StornoFaktura = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' NOVAC
' ============================================================

Public Function StornoNovac_TX(ByVal novacID As String) As Boolean
    Const SRC As String = "StornoNovac_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    If Not StornoNovac(novacID) Then
        Err.Raise ERR_STORNO_BASE + 6, SRC, _
                  "StornoNovac nije uspeo. NovacID=" & novacID
    End If

    tx.CommitTx

    StornoNovac_TX = True
    MonitorStornoSuccess SRC, "Novac", novacID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Novac", novacID, tx
    StornoNovac_TX = False
End Function

Public Function StornoNovac(ByVal novacID As String) As Boolean
    Const SRC As String = "StornoNovac"

    On Error GoTo EH

    Dim rowNov As Long
    rowNov = RequireStornoAllowed(TBL_NOVAC, novacID, COL_NOV_ID, SRC)

    RequireColumnIndex TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC

    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)

    If IsEmpty(novData) Then
        Err.Raise ERR_STORNO_BASE + 40, SRC, _
                  "Tabela novac je prazna."
    End If

    Dim fakturaID As String
    fakturaID = Trim$(CStr(novData(rowNov, _
                    RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC))))

    MarkRowStornirano TBL_NOVAC, rowNov, SRC

    If Len(fakturaID) > 0 Then
        UpdateFakturaStatus fakturaID
    End If

    StornoNovac = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PALETA
' ============================================================

Public Function StornoPaleta_TX(ByVal palID As String) As Boolean
    Const SRC As String = "StornoPaleta_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    If Not StornoPaleta(palID) Then
        Err.Raise ERR_STORNO_BASE + 40, SRC, _
                  "StornoPaleta nije uspeo. PaletaID=" & palID
    End If

    tx.CommitTx

    StornoPaleta_TX = True
    MonitorStornoSuccess SRC, "Paleta", palID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Paleta", palID, tx
    StornoPaleta_TX = False
End Function

Public Function StornoPaleta(ByVal palID As String) As Boolean
    Const SRC As String = "StornoPaleta"

    On Error GoTo EH

    Dim rowPal As Long
    rowPal = RequireStornoAllowed(TBL_PALETA, palID, COL_PAL_ID, SRC)

    ' preradjenu paletu ne stornira se direktno -> prvo storno prerade
    Dim colPre As Long
    colPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
    Dim palData As Variant: palData = GetTableData(TBL_PALETA)
    If UCase$(Trim$(CStr(palData(rowPal, colPre)))) = "DA" Then
        Err.Raise ERR_STORNO_BASE + 42, SRC, _
                  "Paleta je preradjena - prvo stornirajte preradu."
    End If

    MarkRowStornirano TBL_PALETA, rowPal, SRC

    ' storniraj stavke palete (prijemnice se time oslobadjaju)
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(s) Then
        Dim sPal As Long, sStorno As Long
        sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
        sStorno = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)
        Dim r As Long
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(s(r, sPal))) = Trim$(palID) _
               And Not IsStorniranoValue(s(r, sStorno)) Then
                MarkRowStornirano TBL_PALETA_STAVKA, r, SRC
            End If
        Next r
    End If

    StornoPaleta = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PRERADA
' ============================================================

Public Function StornoPrerada_TX(ByVal preradaID As String) As Boolean
    Const SRC As String = "StornoPrerada_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRERADA
    tx.AddTableSnapshot TBL_PRERADA_STAVKA
    tx.AddTableSnapshot TBL_PALETA

    If Not StornoPrerada(preradaID) Then
        Err.Raise ERR_STORNO_BASE + 45, SRC, _
                  "StornoPrerada nije uspeo. PreradaID=" & preradaID
    End If

    tx.CommitTx

    StornoPrerada_TX = True
    MonitorStornoSuccess SRC, "Prerada", preradaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Prerada", preradaID, tx
    StornoPrerada_TX = False
End Function

Public Function StornoPrerada(ByVal preradaID As String) As Boolean
    Const SRC As String = "StornoPrerada"

    On Error GoTo EH

    Dim rowPre As Long
    rowPre = RequireStornoAllowed(TBL_PRERADA, preradaID, COL_PRE_ID, SRC)

    MarkRowStornirano TBL_PRERADA, rowPre, SRC

    ' storniraj stavke + vrati prerađene palete u lager (Preradjeno = "")
    Dim s As Variant: s = GetTableData(TBL_PRERADA_STAVKA)
    If Not IsEmpty(s) Then
        Dim sPre As Long, sPalID As Long, sStorno As Long
        sPre = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID, SRC)
        sPalID = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID, SRC)
        sStorno = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_STORNIRANO, SRC)
        Dim r As Long
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(s(r, sPre))) = Trim$(preradaID) _
               And Not IsStorniranoValue(s(r, sStorno)) Then
                MarkRowStornirano TBL_PRERADA_STAVKA, r, SRC

                Dim palID As String: palID = Trim$(CStr(s(r, sPalID)))
                Dim c As Collection: Set c = FindRows(TBL_PALETA, COL_PAL_ID, palID)
                If Not c Is Nothing Then
                    If c.count > 0 Then
                        RequireUpdateCell TBL_PALETA, CLng(c(1)), COL_PAL_PRERADJENO, "", SRC
                    End If
                End If
            End If
        Next r
    End If

    StornoPrerada = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PUBLIC HELPERS / COMPATIBILITY
' ============================================================

Public Function CanStorno(ByVal tblName As String, _
                          ByVal recordID As String, _
                          ByVal idColumn As String) As Boolean
    Const SRC As String = "CanStorno"

    On Error GoTo EH

    CanStorno = (RequireStornoAllowed(tblName, recordID, idColumn, SRC) > 0)
    Exit Function

EH:
    On Error Resume Next
    LogErr SRC
    Debug.Print SRC & " failed. Table=" & tblName & _
                " ID=" & recordID & _
                " Err=" & CStr(Err.Number) & _
                " Desc=" & Err.description
    On Error GoTo 0

    CanStorno = False
End Function

Public Function LookupActiveID(ByVal tblName As String, _
                               ByVal brojColName As String, _
                               ByVal brojValue As String, _
                               ByVal idColName As String) As String
    Const SRC As String = "LookupActiveID"

    On Error GoTo EH

    RequireNonBlank tblName, "TableName", SRC
    RequireNonBlank brojColName, "BrojColumn", SRC
    RequireNonBlank idColName, "IdColumn", SRC

    Dim data As Variant
    data = GetTableData(tblName)

    If IsEmpty(data) Then
        LookupActiveID = ""
        Exit Function
    End If

    Dim colBroj As Long
    Dim colID As Long
    Dim colStorno As Long

    colBroj = RequireColumnIndex(tblName, brojColName, SRC)
    colID = RequireColumnIndex(tblName, idColName, SRC)
    colStorno = RequireColumnIndex(tblName, COL_STORNIRANO, SRC)

    Dim resultId As String
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colBroj))) = Trim$(brojValue) Then
            If Not IsStorniranoValue(data(i, colStorno)) Then
                resultId = CStr(data(i, colID))
            End If
        End If
    Next i

    LookupActiveID = resultId
    Exit Function

EH:
    On Error Resume Next
    LogErr SRC
    On Error GoTo 0
    LookupActiveID = ""
End Function

' ============================================================
' PRIVATE BUSINESS HELPERS
' ============================================================

Private Sub StornoFakturaStavkeAndReleasePrijemnice(ByVal fakturaID As String)
    Const SRC As String = "StornoFakturaStavkeAndReleasePrijemnice"

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colFakID As Long
    Dim colPrijID As Long

    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC)

    Dim i As Long

    For i = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(i, colFakID))) = Trim$(fakturaID) Then
            MarkRowStornirano TBL_FAKTURA_STAVKE, i, SRC

            Dim prijID As String
            prijID = Trim$(CStr(stavkeData(i, colPrijID)))

            If Len(prijID) > 0 Then
                ReleasePrijemnicaFromFaktura prijID, fakturaID
            End If
        End If
    Next i
End Sub

Private Sub ReleasePrijemnicaFromFaktura(ByVal prijemnicaID As String, _
                                         ByVal fakturaID As String)
    Const SRC As String = "ReleasePrijemnicaFromFaktura"

    Dim rows As Collection
    Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)

    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 50, SRC, _
                  "Dupla PrijemnicaID vrednost: " & prijemnicaID
    End If

    Dim rowPrij As Long
    rowPrij = CLng(rows(1))

    RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURISANO, "", SRC
    RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURA_ID, "", SRC
End Sub

Private Sub MarkFakturaOrphaned(ByVal fakturaID As String, _
                                ByVal prijemnicaID As String)
    Const SRC As String = "MarkFakturaOrphaned"

    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)

    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 51, SRC, _
                  "Dupla FakturaID vrednost: " & fakturaID
    End If

    RequireUpdateCell TBL_FAKTURE, CLng(rows(1)), COL_OSIROCENO_OD, _
                      prijemnicaID, SRC
End Sub

Private Sub MarkFakturaStavkeOrphaned(ByVal fakturaID As String, _
                                      ByVal prijemnicaID As String)
    Const SRC As String = "MarkFakturaStavkeOrphaned"

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colFakID As Long
    Dim colPrijID As Long

    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC)

    Dim i As Long

    For i = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(i, colPrijID))) = Trim$(prijemnicaID) And _
           Trim$(CStr(stavkeData(i, colFakID))) = Trim$(fakturaID) Then
            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, _
                              prijemnicaID, SRC
        End If
    Next i
End Sub

Private Sub ResetNovacFakturaLink(ByVal fakturaID As String)
    Const SRC As String = "ResetNovacFakturaLink"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Sub

    Dim colFakID As Long
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakID))) = Trim$(fakturaID) Then
            RequireUpdateCell TBL_NOVAC, i, COL_NOV_FAKTURA_ID, "", SRC
        End If
    Next i
End Sub

Private Sub ResetNovacOtkupLink(ByVal otkupID As String)
    Const SRC As String = "ResetNovacOtkupLink"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Sub

    Dim colOtkupID As Long
    colOtkupID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colOtkupID))) = Trim$(otkupID) Then
            RequireUpdateCell TBL_NOVAC, i, COL_NOV_OTKUP_ID, "", SRC
        End If
    Next i
End Sub

Private Sub StornoAmbalazaByDokument(ByVal dokumentID As String, _
                                     ByVal dokumentTip As String)
    Const SRC As String = "StornoAmbalazaByDokument"

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)

    If IsEmpty(data) Then Exit Sub

    Dim colDokID As Long
    Dim colDokTip As Long
    Dim colStorno As Long

    colDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)
    colDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, SRC)
    colStorno = RequireColumnIndex(TBL_AMBALAZA, COL_STORNIRANO, SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colDokID))) = Trim$(dokumentID) And _
           Trim$(CStr(data(i, colDokTip))) = Trim$(dokumentTip) Then

            If Not IsStorniranoValue(data(i, colStorno)) Then
                MarkRowStornirano TBL_AMBALAZA, i, SRC
            End If
        End If
    Next i
End Sub

' ============================================================
' PRIVATE GUARDS / LOW-LEVEL HELPERS
' ============================================================

Private Function RequireStornoAllowed(ByVal tblName As String, _
                                      ByVal recordID As String, _
                                      ByVal idColumn As String, _
                                      ByVal sourceName As String) As Long
    RequireNonBlank tblName, "TableName", sourceName
    RequireNonBlank recordID, "RecordID", sourceName
    RequireNonBlank idColumn, "IdColumn", sourceName

    RequireColumnIndex tblName, idColumn, sourceName
    RequireColumnIndex tblName, COL_STORNIRANO, sourceName

    Dim rows As Collection
    Set rows = FindRows(tblName, idColumn, recordID)

    If rows Is Nothing Then
        Err.Raise ERR_STORNO_BASE + 60, sourceName, _
                  "FindRows je vratio Nothing. Table=" & tblName & _
                  " ID=" & recordID
    End If

    If rows.count = 0 Then
        Err.Raise ERR_STORNO_BASE + 61, sourceName, _
                  "Stavka nije pronadjena. Table=" & tblName & _
                  " ID=" & recordID
    End If

    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 62, sourceName, _
                  "ID nije jedinstven. Table=" & tblName & _
                  " ID=" & recordID & _
                  " Count=" & CStr(rows.count)
    End If

    Dim rowIndex As Long
    rowIndex = CLng(rows(1))

    Dim data As Variant
    data = GetTableData(tblName)

    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_BASE + 63, sourceName, _
                  "Tabela je prazna posle pronalaska reda. Table=" & tblName
    End If

    Dim colStorno As Long
    colStorno = RequireColumnIndex(tblName, COL_STORNIRANO, sourceName)

    If IsStorniranoValue(data(rowIndex, colStorno)) Then
        Err.Raise ERR_STORNO_BASE + 64, sourceName, _
                  "Vec stornirano. Table=" & tblName & _
                  " ID=" & recordID
    End If

    RequireStornoAllowed = rowIndex
End Function

Private Sub MarkRowStornirano(ByVal tblName As String, _
                              ByVal rowIndex As Long, _
                              ByVal sourceName As String)
    RequireUpdateCell tblName, rowIndex, COL_STORNIRANO, STORNO_DA, sourceName
End Sub

Private Sub RequireNonBlank(ByVal value As String, _
                            ByVal fieldName As String, _
                            ByVal sourceName As String)
    If Len(Trim$(value)) = 0 Then
        Err.Raise ERR_STORNO_BASE + 70, sourceName, _
                  fieldName & " je obavezan."
    End If
End Sub

Private Function IsStorniranoValue(ByVal value As Variant) As Boolean
    IsStorniranoValue = (UCase$(Trim$(CStr(value))) = UCase$(STORNO_DA))
End Function

Private Sub LogAndReraise(ByVal sourceName As String)
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr sourceName
    On Error GoTo 0

    Err.Raise errNum, sourceName, "Source=" & errSrc & " | " & errDesc
End Sub

' ============================================================
' MONITORING / TX ERROR HANDLING
' ============================================================

Private Sub HandleStornoTxError(ByVal procedureName As String, _
                                ByVal entityType As String, _
                                ByVal entityId As String, _
                                ByRef tx As clsTransaction)
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr procedureName

    If Not tx Is Nothing Then tx.RollbackTx

    Monitor_Error _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityId:=entityId, _
        correlationId:=entityId, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="STORNO_" & UCase$(entityType) & "_FAIL", _
        severity:="ERROR", _
        message:=entityType & " storno failed. ID=" & entityId & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityId:=entityId, _
        correlationId:=entityId

    Debug.Print procedureName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc

    On Error GoTo 0
End Sub

Private Sub MonitorStornoSuccess(ByVal procedureName As String, _
                                 ByVal entityType As String, _
                                 ByVal entityId As String)
    On Error Resume Next

    Monitor_Event _
        eventType:="STORNO_" & UCase$(entityType) & "_SUCCESS", _
        severity:="INFO", _
        message:=entityType & " stornirano. ID=" & entityId, _
        userId:="Operator", _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityId:=entityId, _
        correlationId:=entityId

    On Error GoTo 0
End Sub

