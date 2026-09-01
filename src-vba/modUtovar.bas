Attribute VB_Name = "modUtovar"
Option Explicit

' ============================================================
' modUtovar (krug 5 revizije #248) -- UTOVARNA LISTA: dokument
' FIZICKE isporuke gotove robe.
'
' Grain: prerada je proizvodni lot (koliko je PROIZVEDENO); utovarna
' stavka (prerada + kg) je prodajna jedinica -- parcijalna prodaja
' (500 kg od 2.000) je legalna, "na stanju" = NetoIzlazKg - SUM
' aktivnih utovarenih kg. Prerada se NIKAD ne zakljucava fakturom.
'
' v1 ugovor: JEDAN utovar = JEDNA GP faktura, prave se u ISTOJ
' transakciji (CreateUtovarSaFakturom_TX); DatumUtovara je datum
' izrade i on ide na SEF kao datum isporuke. Poseban ekran utovara i
' stampani obrazac utovarne liste su sledeci korak (dokument vec
' postoji podatkovno i stampa ce citati ove tabele).
'
' Storno simetrija (modStorno): storno FAKTURE oslobadja utovar
' (roba ostaje utovarena); storno UTOVARA vraca robu na stanje i
' dozvoljen je samo nad nefakturisanim utovarom.
' ============================================================

' Sledeci broj utovara -- maxN+1 unutar tekuce godine (isti obrazac
' kao GenerateBrojPalete).
Public Function GenerateBrojUtovara() As Long
    Const SRC As String = "modUtovar.GenerateBrojUtovara"
    Dim d As Variant, i As Long, maxN As Long
    Dim cBr As Long, cGod As Long
    d = GetTableData(TBL_UTOVAR)
    If IsArray(d) Then
        cBr = RequireColumnIndex(TBL_UTOVAR, COL_UT_BROJ, SRC)
        cGod = RequireColumnIndex(TBL_UTOVAR, COL_UT_GODINA, SRC)
        For i = 1 To UBound(d, 1)
            If IsNumeric(d(i, cGod)) And IsNumeric(d(i, cBr)) Then
                If CLng(d(i, cGod)) = Year(Date) Then
                    If CLng(d(i, cBr)) > maxN Then maxN = CLng(d(i, cBr))
                End If
            End If
        Next i
    End If
    GenerateBrojUtovara = maxN + 1
End Function

' Mapa AKTIVNO utovarenih kg po preradi -- jedan prolaz (S5 pravilo).
' Aktivna stavka = stavka nije stornirana I njen utovar nije storniran.
' Meko: sveska pre nadogradnje nema tabele -> prazna mapa (sve na
' stanju). Public: dele je grid, writer i storno kapija -- JEDNO
' pravilo, ne tri kopije (pouka kruga 3).
Public Function UtovarenoPoPreradi() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set UtovarenoPoPreradi = d

    If GetTable(TBL_UTOVAR_STAVKE) Is Nothing Then Exit Function
    If GetTable(TBL_UTOVAR) Is Nothing Then Exit Function

    ' Aktivni utovari (dict utovarID -> True).
    Dim ut As Variant, i As Long
    Dim aktivni As Object: Set aktivni = CreateObject("Scripting.Dictionary")
    aktivni.CompareMode = vbTextCompare
    ut = GetTableData(TBL_UTOVAR)
    If IsArray(ut) Then
        ut = ExcludeStornirano(ut, TBL_UTOVAR)
        If IsArray(ut) Then
            Dim cUtId As Long
            cUtId = RequireColumnIndex(TBL_UTOVAR, COL_UT_ID, "modUtovar.UtovarenoPoPreradi")
            For i = 1 To UBound(ut, 1)
                aktivni(Trim$(CStr(nz(ut(i, cUtId))))) = True
            Next i
        End If
    End If

    Dim s As Variant
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Function
    s = ExcludeStornirano(s, TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Function

    Dim cUt As Long, cPre As Long, cKol As Long, k As String
    cUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, "modUtovar.UtovarenoPoPreradi")
    cPre = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_PRERADA_ID, "modUtovar.UtovarenoPoPreradi")
    cKol = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KOLICINA, "modUtovar.UtovarenoPoPreradi")
    For i = 1 To UBound(s, 1)
        If aktivni.Exists(Trim$(CStr(nz(s(i, cUt))))) Then
            k = Trim$(CStr(nz(s(i, cPre))))
            If Len(k) > 0 And IsNumeric(s(i, cKol)) Then
                If d.Exists(k) Then
                    d(k) = CDbl(d(k)) + CDbl(s(i, cKol))
                Else
                    d.Add k, CDbl(s(i, cKol))
                End If
            End If
        End If
    Next i
End Function

' Aktivno utovareno kg jedne prerade (kapija storna prerade).
Public Function UtovarenoKgPrerade(ByVal preradaID As String) As Double
    Dim d As Object: Set d = UtovarenoPoPreradi()
    If d.Exists(Trim$(preradaID)) Then _
        UtovarenoKgPrerade = CDbl(d(Trim$(preradaID)))
End Function

' ============================================================
' UTOVAR + GP FAKTURA u jednoj transakciji (v1: 1 utovar = 1 faktura).
' stavke: Collection of Array(preradaID, kolicinaKg, cena).
' Kapije u BASE, pod TX: kupac postoji tacno jednom; prerada postoji
' tacno jednom, nije stornirana, ima IMENOVAN proizvod; kolicina > 0 i
' <= na stanju; dupla prerada u istoj listi zabranjena; cena > 0.
' ============================================================
Public Function CreateUtovarSaFakturom_TX(ByVal kupacID As String, _
                                          ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR
    tx.AddTableSnapshot TBL_UTOVAR_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_NOVAC

    CreateUtovarSaFakturom_TX = CreateUtovarSaFakturom(kupacID, stavke)

    If CreateUtovarSaFakturom_TX = "" Then
        Err.Raise vbObjectError + 1730, "CreateUtovarSaFakturom_TX", _
                  "CreateUtovarSaFakturom nije uspeo."
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="UTOVAR_FAKTURA_GP_SUCCESS", _
        severity:="INFO", _
        message:="Utovar + GP faktura created successfully", _
        userId:="Operator", _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovarSaFakturom_TX", _
        entityType:="Faktura", _
        entityID:=CreateUtovarSaFakturom_TX, _
        correlationId:=CreateUtovarSaFakturom_TX
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "CreateUtovarSaFakturom_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovarSaFakturom_TX", _
        entityType:="Faktura", _
        entityID:="", _
        correlationId:="CreateUtovarSaFakturom", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateUtovarSaFakturom_TX = ""
End Function

' Base -- NE zovi je spolja (pola upisa bez transakcije).
Private Function CreateUtovarSaFakturom(ByVal kupacID As String, _
                                        ByVal stavke As Collection) As String
    Const SRC As String = "CreateUtovarSaFakturom"
    On Error GoTo EH

    If Trim$(kupacID) = "" Then
        Err.Raise vbObjectError + 1731, SRC, "KupacID je obavezan."
    End If
    ' Writer je samostalna granica: GP nema prijemnicu za implicitnu
    ' proveru vlasnistva, kupac se proverava ovde.
    Dim kupRows As Collection
    Set kupRows = FindRows(TBL_KUPCI, COL_KUP_ID, Trim$(kupacID))
    If kupRows Is Nothing Then
        Err.Raise vbObjectError + 1751, SRC, _
                  "Kupac ne postoji u tblKupci: " & kupacID
    ElseIf kupRows.count <> 1 Then
        Err.Raise vbObjectError + 1751, SRC, _
                  "Kupac ne postoji jednoznacno u tblKupci: " & kupacID & _
                  "; Count=" & CStr(kupRows.count)
    End If
    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1732, SRC, "Stavke nisu prosledjene."
    End If
    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1733, SRC, _
                  "Utovar mora imati bar jednu stavku."
    End If

    ' Fail-fast schema guards -- bez utovar tabela se staje ODMAH.
    RequireColumnIndex TBL_UTOVAR, COL_UT_ID, SRC
    RequireColumnIndex TBL_UTOVAR_STAVKE, COL_UTS_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_UTOVAR_ID, SRC

    Dim preData As Variant
    preData = GetTableData(TBL_PRERADA)
    If IsEmpty(preData) Then
        Err.Raise vbObjectError + 1734, SRC, "Tabela prerada je prazna."
    End If

    Dim colNetoIzlaz As Long, colBroj As Long, colGodina As Long
    Dim colStorno As Long, colTipGp As Long
    colNetoIzlaz = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
    colBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
    colGodina = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
    colStorno = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, SRC)
    colTipGp = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)

    ' Na stanju = proizvedeno - vec utovareno (jedno pravilo za sve).
    Dim utovareno As Object: Set utovareno = UtovarenoPoPreradi()

    ' Pre-validacija SVIH stavki pre ijednog upisa.
    Dim s As Variant, preradaID As String, kolicina As Double, cena As Double
    Dim rows As Collection, rowPre As Long, raspolozivo As Double
    Dim preRows As Object, preValues As Object
    Set preRows = CreateObject("Scripting.Dictionary")
    Set preValues = CreateObject("Scripting.Dictionary")

    For Each s In stavke
        preradaID = Trim$(CStr(s(0)))
        If Len(preradaID) = 0 Then
            Err.Raise vbObjectError + 1735, SRC, "PreradaID je obavezan."
        End If
        If Not IsNumeric(s(1)) Then
            Err.Raise vbObjectError + 1752, SRC, _
                      "Kolicina nije numericka. PreradaID=" & preradaID
        End If
        kolicina = CDbl(s(1))
        If kolicina <= 0 Then
            Err.Raise vbObjectError + 1752, SRC, _
                      "Kolicina mora biti veca od nule. PreradaID=" & preradaID
        End If
        If Not IsNumeric(s(2)) Then
            Err.Raise vbObjectError + 1736, SRC, _
                      "Cena nije numericka. PreradaID=" & preradaID
        End If
        cena = CDbl(s(2))
        If cena <= 0 Then
            Err.Raise vbObjectError + 1736, SRC, _
                      "Cena mora biti veca od nule. PreradaID=" & preradaID
        End If
        If preRows.Exists(preradaID) Then
            Err.Raise vbObjectError + 1737, SRC, _
                      "Dupla prerada u izboru: " & preradaID
        End If

        Set rows = FindRows(TBL_PRERADA, COL_PRE_ID, preradaID)
        If rows Is Nothing Then
            Err.Raise vbObjectError + 1738, SRC, _
                      "Prerada nije pronadjena: " & preradaID
        End If
        If rows.count = 0 Then
            Err.Raise vbObjectError + 1738, SRC, _
                      "Prerada nije pronadjena: " & preradaID
        End If
        If rows.count > 1 Then
            Err.Raise vbObjectError + 1739, SRC, _
                      "Duplikat PreradaID=" & preradaID & _
                      "; Count=" & CStr(rows.count)
        End If
        rowPre = CLng(rows(1))

        ' Inline "DA" provera (IsStorniranoValue je Private u modStorno).
        If UCase$(Trim$(CStr(nz(preData(rowPre, colStorno))))) = "DA" Then
            Err.Raise vbObjectError + 1740, SRC, _
                      "Prerada je stornirana: " & preradaID
        End If
        ' Stavka prodajne fakture mora imenovati proizvod.
        If Len(Trim$(CStr(nz(preData(rowPre, colTipGp))))) = 0 Then
            Err.Raise vbObjectError + 1750, SRC, _
                      "TipGotovogProizvoda je prazan -- faktura mora imenovati proizvod: " & preradaID
        End If
        If Not IsNumeric(preData(rowPre, colNetoIzlaz)) Then
            Err.Raise vbObjectError + 1742, SRC, _
                      "NetoIzlazKg nije numericki. PreradaID=" & preradaID
        End If
        ' KLJUCNA kapija graina: kolicina <= na stanju. Parcijalna
        ' prodaja je legalna; prekoracenje stanja nije.
        raspolozivo = CDbl(preData(rowPre, colNetoIzlaz))
        If utovareno.Exists(preradaID) Then _
            raspolozivo = raspolozivo - CDbl(utovareno(preradaID))
        If kolicina > raspolozivo + 0.0001 Then
            Err.Raise vbObjectError + 1753, SRC, _
                      "Kolicina " & CStr(kolicina) & " kg prelazi stanje (" & _
                      CStr(raspolozivo) & " kg). PreradaID=" & preradaID
        End If

        preRows.Add preradaID, rowPre
        preValues.Add preradaID, Array( _
            kolicina, cena, _
            Trim$(CStr(nz(preData(rowPre, colBroj)))) & "/" & _
            Trim$(CStr(nz(preData(rowPre, colGodina)))))
    Next s

    Dim ukupno As Double, key As Variant, preVals As Variant
    For Each key In preValues.keys
        preVals = preValues(CStr(key))
        ukupno = ukupno + (CDbl(preVals(0)) * CDbl(preVals(1)))
    Next key
    If ukupno <= 0 Then
        Err.Raise vbObjectError + 1743, SRC, _
                  "Ukupan iznos fakture mora biti veci od nule."
    End If

    ' --- UPIS 1: utovarna lista (fizicka isporuka, danasnji datum).
    Dim utovarID As String
    utovarID = GetNextID(TBL_UTOVAR, COL_UT_ID, "UT-")
    If utovarID = "" Then
        Err.Raise vbObjectError + 1754, SRC, "GetNextID nije vratio UtovarID."
    End If

    ' Positional AppendRow je ovde bezbedan: tblUtovar/tblUtovarStavke
    ' pravi EnsureUtovarSchemaCore pa je redosled kolona nas (v. Array
    ' u modSetup); svaka BUDUCA kolona ide na kraj (EnsureDataTable).
    If AppendRow(TBL_UTOVAR, Array( _
        utovarID, GenerateBrojUtovara(), Year(Date), Date, kupacID, _
        "", "", "", "")) <= 0 Then
        Err.Raise vbObjectError + 1755, SRC, "AppendRow nije uspeo za tblUtovar."
    End If

    Dim stavkaNum As Long, rowUts As Long
    For Each s In stavke
        preradaID = Trim$(CStr(s(0)))
        preVals = preValues(preradaID)
        stavkaNum = stavkaNum + 1
        rowUts = AppendRow(TBL_UTOVAR_STAVKE, Array( _
            utovarID & "-" & Format$(stavkaNum, "00"), utovarID, preradaID, _
            CStr(preVals(2)), CDbl(preVals(0)), ""))
        If rowUts <= 0 Then
            Err.Raise vbObjectError + 1756, SRC, _
                      "AppendRow nije uspeo za tblUtovarStavke."
        End If
    Next s

    ' --- UPIS 2: faktura iz utovara (v1: 1 utovar = 1 faktura).
    Dim fakturaID As String
    fakturaID = GetNextID(TBL_FAKTURE, COL_FAK_ID, "FAK-")
    If fakturaID = "" Then
        Err.Raise vbObjectError + 1744, SRC, "GetNextID nije vratio FakturaID."
    End If

    Dim brojFakture As String
    brojFakture = GenerateBrojFakture()
    If brojFakture = "" Then
        Err.Raise vbObjectError + 1745, SRC, _
                  "GenerateBrojFakture nije vratio broj fakture."
    End If

    ' Header -- ISTI pozicioni oblik kao CreateFaktura (21 kolona).
    Dim fakturaRow As Variant
    fakturaRow = Array( _
        fakturaID, _
        brojFakture, _
        Date, _
        kupacID, _
        ukupno, _
        STATUS_NEPLACENO, _
        Empty, _
        "", _
        "", _
        WF_LOCAL_FINALIZED, _
        "", _
        "", _
        "", _
        Empty, _
        Empty, _
        "", _
        "", _
        "", _
        0, _
        "Ne", _
        "" _
    )
    If AppendRow(TBL_FAKTURE, fakturaRow) <= 0 Then
        Err.Raise vbObjectError + 1746, SRC, _
                  "AppendRow nije uspeo za tblFakture."
    End If

    ' Stavke: pozicioni deo isti kao sveze (PrijemnicaID/BrojPrijemnice
    ' PRAZNI); GP identitet (prerada + utovar) ide PO IMENU u kolone na
    ' kraju tabele (podaci-i-config pravilo).
    Dim stavkaID As String, stavkaRow As Variant, rowStavke As Long
    stavkaNum = 0
    For Each s In stavke
        stavkaNum = stavkaNum + 1
        stavkaID = fakturaID & "-" & Format$(stavkaNum, "00")
        preradaID = Trim$(CStr(s(0)))
        preVals = preValues(preradaID)

        stavkaRow = Array( _
            stavkaID, _
            fakturaID, _
            "", _
            CDbl(preVals(0)), _
            CDbl(preVals(1)), _
            "", _
            "", _
            "", _
            "" _
        )
        rowStavke = AppendRow(TBL_FAKTURA_STAVKE, stavkaRow)
        If rowStavke <= 0 Then
            Err.Raise vbObjectError + 1747, SRC, _
                      "AppendRow nije uspeo za tblFakturaStavke."
        End If
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_PRERADA_ID, _
                          preradaID, SRC
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_BROJ_PRERADE, _
                          CStr(preVals(2)), SRC
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_UTOVAR_ID, _
                          utovarID, SRC
    Next s

    ' Utovar markiran svojom fakturom (1:1 -- utovar je dokument
    ' isporuke i taj invariant JESTE validan, za razliku od prerade).
    Dim utRows As Collection
    Set utRows = FindRows(TBL_UTOVAR, COL_UT_ID, utovarID)
    RequireUpdateCell TBL_UTOVAR, CLng(utRows(1)), COL_UT_FAKTURISANO, "Da", SRC
    RequireUpdateCell TBL_UTOVAR, CLng(utRows(1)), COL_UT_FAKTURA_ID, fakturaID, SRC

    ' Avans se prebija isto kao kod svezih faktura.
    ApplyAvansToFaktura kupacID, fakturaID

    CreateUtovarSaFakturom = fakturaID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function
