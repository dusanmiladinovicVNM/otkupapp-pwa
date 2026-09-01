Attribute VB_Name = "modFaktura"

Option Explicit

' ============================================================
' modFaktura v2.1 - Rechnungserstellung
' GEAeNDERT: Basiert auf tblPrijemnica statt tblIsporuka
' Faktura-Betrag = Prijemnica.Kolicina x Prijemnica.Cena
' ============================================================

Public Function CreateFaktura_TX(ByVal kupacID As String, _
                                  ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_NOVAC

    CreateFaktura_TX = CreateFaktura(kupacID, stavke)

    If CreateFaktura_TX = "" Then
        Err.Raise vbObjectError + 1701, "CreateFaktura_TX", _
                  "CreateFaktura fehlgeschlagen"
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="FAKTURA_CREATE_SUCCESS", _
        severity:="INFO", _
        message:="Faktura created successfully", _
        userId:="Operator", _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX
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

    LogErr "CreateFaktura_TX"
    On Error Resume Next


    Monitor_Error _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="FAKTURA_CREATE_FAIL", _
        severity:="ERROR", _
        message:=errDesc, _
        userId:="Operator", _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX

    If Not tx Is Nothing Then tx.RollbackTx

    On Error GoTo 0

    CreateFaktura_TX = ""

    Debug.Print " CreateFaktura_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function

' Base funkcija -- NE zovi je spolja. Jedini ulaz je CreateFaktura_TX, koji drzi
' snapshot transakciju; direktan poziv bi kod greske ostavio pola upisa
' (faktura header bez stavki, prijemnice markirane bez fakture) -- AUD-011 /
' FM-0034 #3.
Private Function CreateFaktura(ByVal kupacID As String, _
                               ByVal stavke As Collection) As String
    On Error GoTo EH

    If Trim$(kupacID) = "" Then
        Err.Raise vbObjectError + 1702, "CreateFaktura", _
                  "KupacID je obavezan."
    End If

    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1703, "CreateFaktura", _
                  "Stavke nisu prosledjene."
    End If

    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1704, "CreateFaktura", _
                  "Faktura mora imati bar jednu stavku."
    End If

    ' Fail-fast schema guards
    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_BROJ, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_DATUM, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_KUPAC, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_IZNOS, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_STATUS, "CreateFaktura"

    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_KOLICINA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_CENA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_KLASA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRIJEMNICE, "CreateFaktura"

    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_ID, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KUPAC, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_STORNIRANO, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_CENA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KLASA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_BROJ, "CreateFaktura"

    Dim fakturaID As String
    fakturaID = GetNextID(TBL_FAKTURE, COL_FAK_ID, "FAK-")

    If fakturaID = "" Then
        Err.Raise vbObjectError + 1705, "CreateFaktura", _
                  "GetNextID nije vratio FakturaID."
    End If

    Dim brojFakture As String
    brojFakture = GenerateBrojFakture()

    If brojFakture = "" Then
        Err.Raise vbObjectError + 1706, "CreateFaktura", _
                  "GenerateBrojFakture nije vratio broj fakture."
    End If
    
        Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(prijData) Then
        Err.Raise vbObjectError + 1717, "CreateFaktura", _
                  "Tabela prijemnica je prazna."
    End If

    Dim colPrjKol As Long
    Dim colPrjCena As Long
    Dim colPrjKlasa As Long
    Dim colPrjBroj As Long
    Dim colPrjKupac As Long

    colPrjKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "CreateFaktura")
    colPrjCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "CreateFaktura")
    colPrjKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, "CreateFaktura")
    colPrjBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, "CreateFaktura")
    colPrjKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, "CreateFaktura")

    ' Pre-validacija svih prijemnica pre bilo kog upisa.
    ' Business module trusts only PrijemnicaID from caller.
    ' Kolicina/Cena/Klasa/BrojPrijemnice are derived from tblPrijemnica.
    Dim s As Variant
    Dim prijemnicaID As String
    Dim rows As Collection
    Dim prijRows As Object
    Dim prijValues As Object

    Set prijRows = CreateObject("Scripting.Dictionary")
    Set prijValues = CreateObject("Scripting.Dictionary")

    For Each s In stavke

        prijemnicaID = GetPrijemnicaIDFromFakturaStavka(s, "CreateFaktura")

        Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)

        If rows Is Nothing Or rows.count = 0 Then
            Err.Raise vbObjectError + 1708, "CreateFaktura", _
                      "Prijemnica nije pronadena: " & prijemnicaID
        End If

        If prijRows.Exists(prijemnicaID) Then
            Err.Raise vbObjectError + 1710, "CreateFaktura", _
                      "Dupla prijemnica u izboru: " & prijemnicaID
        End If

        ' Fail-closed kod duplog PrijemnicaID: bez ovoga se tiho uzima prvi
        ' pogodak, pa se fakturise kolicina/cena pogresnog reda (AUD-011 /
        ' FM-0034 #2). Isti obrazac kao RequireSingleFakturaRow nad tblFakture.
        If rows.count > 1 Then
            Err.Raise vbObjectError + 1707, "CreateFaktura", _
                      "Duplikat PrijemnicaID=" & prijemnicaID & _
                      "; Count=" & CStr(rows.count)
        End If

        Dim rowPrijValidate As Long
        rowPrijValidate = CLng(rows(1))

        ' Vlasnistvo: prijemnica mora da pripada kupcu fakture. Bez ove provere
        ' se prijemnica drugog kupca moze zavuci u fakturu (UI filter po kupcu
        ' nije sigurnosna granica) -- AUD-011 / FM-0034 #1.
        Dim prjKupac As String
        prjKupac = Trim$(CStr(prijData(rowPrijValidate, colPrjKupac)))

        If prjKupac <> Trim$(kupacID) Then
            Err.Raise vbObjectError + 1721, "CreateFaktura", _
                      "Prijemnica pripada drugom kupcu. PrijemnicaID=" & prijemnicaID & _
                      "; Prijemnica.KupacID=" & prjKupac & _
                      "; Faktura.KupacID=" & Trim$(kupacID)
        End If

        If Not IsPrijemnicaAvailableForFaktura(rowPrijValidate, prijemnicaID) Then
            Err.Raise vbObjectError + 1709, "CreateFaktura", _
                      "Prijemnica je ve" & ChrW(263) & " fakturisana ili stornirana: " & prijemnicaID
        End If

        If Not IsNumeric(prijData(rowPrijValidate, colPrjKol)) Then
            Err.Raise vbObjectError + 1711, "CreateFaktura", _
                      "Koli" & ChrW(269) & "ina nije numericka za prijemnicu: " & prijemnicaID
        End If

        If Not IsNumeric(prijData(rowPrijValidate, colPrjCena)) Then
            Err.Raise vbObjectError + 1711, "CreateFaktura", _
                      "Cena nije numericka za prijemnicu: " & prijemnicaID
        End If

        Dim prjKolicina As Double
        Dim prjCena As Double
        Dim prjKlasa As String
        Dim prjBroj As String

        prjKolicina = CDbl(prijData(rowPrijValidate, colPrjKol))
        prjCena = CDbl(prijData(rowPrijValidate, colPrjCena))
        prjKlasa = CStr(prijData(rowPrijValidate, colPrjKlasa))
        prjBroj = CStr(prijData(rowPrijValidate, colPrjBroj))

        If prjKolicina <= 0 Then
            Err.Raise vbObjectError + 1712, "CreateFaktura", _
                      "Koli" & ChrW(269) & "ina mora biti veca od nule. PrijemnicaID=" & prijemnicaID
        End If

        If prjCena < 0 Then
            Err.Raise vbObjectError + 1713, "CreateFaktura", _
                      "Cena ne sme biti negativna. PrijemnicaID=" & prijemnicaID
        End If

        prijRows.Add prijemnicaID, rowPrijValidate
        prijValues.Add prijemnicaID, Array(prjKolicina, prjCena, prjKlasa, prjBroj)

    Next s

    ' Ukupan iznos se racuna iz canonical tblPrijemnica vrednosti.
    Dim ukupno As Double
    Dim key As Variant
    Dim prjVals As Variant

    For Each key In prijValues.keys
        prjVals = prijValues(CStr(key))
        ukupno = ukupno + (CDbl(prjVals(0)) * CDbl(prjVals(1)))
    Next key

    If ukupno <= 0 Then
        Err.Raise vbObjectError + 1714, "CreateFaktura", _
                  "Ukupan iznos fakture mora biti veci od nule."
    End If

    ' Faktura header
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
        Err.Raise vbObjectError + 1715, "CreateFaktura", _
                  "AppendRow fehlgeschlagen fuer tblFakture."
    End If

    ' Faktura stavke + markiranje prijemnica
    Dim stavkaID As String
    Dim stavkaNum As Long
    Dim stavkaRow As Variant
    Dim rowPrij As Long

    For Each s In stavke
        stavkaNum = stavkaNum + 1
        stavkaID = fakturaID & "-" & Format$(stavkaNum, "00")

        prijemnicaID = GetPrijemnicaIDFromFakturaStavka(s, "CreateFaktura")
        rowPrij = CLng(prijRows(prijemnicaID))

        prjVals = prijValues(prijemnicaID)

        stavkaRow = Array( _
            stavkaID, _
            fakturaID, _
            prijemnicaID, _
            CDbl(prjVals(0)), _
            CDbl(prjVals(1)), _
            CStr(prjVals(2)), _
            CStr(prjVals(3)), _
            "", _
            "" _
        )

        If AppendRow(TBL_FAKTURA_STAVKE, stavkaRow) <= 0 Then
            Err.Raise vbObjectError + 1716, "CreateFaktura", _
                      "AppendRow fehlgeschlagen fuer tblFakturaStavke."
        End If

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURISANO, _
                          "Da", "CreateFaktura"

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURA_ID, _
                          fakturaID, "CreateFaktura"
    Next s

    ' Avans automatisch verrechnen.
    ' Ovo mora biti base funkcija, ne ApplyAvansToFaktura_TX,
    ' jer CreateFaktura_TX vec drzi siru transakciju.
    ApplyAvansToFaktura kupacID, fakturaID

    CreateFaktura = fakturaID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "CreateFaktura"
    On Error Resume Next
    On Error GoTo 0

    Err.Raise errNum, "CreateFaktura", _
              "Source=" & errSrc & " | " & errDesc
End Function

' GP lista za ekran Fakturisanje: NESTORNIRANE prerade. 1-bazirano:
'   1 PreradaID | 2 Broj (b/g) | 3 TipGP | 4 Datum | 5 NetoIzlazKg
'   6 Kutije | 7 Kese | 8 Dostupna (nije fakturisana) | 9 BrojFakture
' GetColumnIndex za GP kolone fakturisanosti: sveska PRE EnsureSchema
' nadogradnje ih nema -- tada su sve prerade "dostupne" a broj fakture
' prazan, sto je i stvarno stanje te sveske.
Public Function GetGPZaFakturisanjeForGrid() As Variant
    Const SRC As String = "GetGPZaFakturisanjeForGrid"
    On Error GoTo EH

    Dim pd As Variant
    pd = GetTableData(TBL_PRERADA)
    If IsEmpty(pd) Then Exit Function

    Dim cId As Long, cBroj As Long, cGod As Long, cTip As Long
    Dim cDat As Long, cNeto As Long, cKut As Long, cKes As Long
    Dim cStorno As Long, cFakt As Long, cFakID As Long
    cId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
    cBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
    cGod = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
    cTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)
    cDat = RequireColumnIndex(TBL_PRERADA, COL_PRE_DATUM, SRC)
    cNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
    cKut = RequireColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE, SRC)
    cKes = RequireColumnIndex(TBL_PRERADA, COL_PRE_KESE, SRC)
    cStorno = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, SRC)
    cFakt = GetColumnIndex(TBL_PRERADA, COL_PRE_FAKTURISANO)
    cFakID = GetColumnIndex(TBL_PRERADA, COL_PRE_FAKTURA_ID)

    ' P1 (revizija #248): dupli PreradaID (korupcija) ne sme da izgleda
    ' kao normalan red -- isti guard kao prijemnice/fakture: identitet
    ' se prazni pa radnja nema nad cim da radi (fail-closed u UI).
    Dim brojac As Object: Set brojac = BrojacIdova(TBL_PRERADA, COL_PRE_ID)
    ' B2: dostupnost deli isti kanonski contract kao writer.
    Dim stAkt As Object: Set stAkt = GPAktivneStavkePoPreradi()

    ' Mapa AKTIVNIH faktura za prikaz broja -- PRE petlje (par. 23.11/S5).
    Dim fakBroj As Object: Set fakBroj = CreateObject("Scripting.Dictionary")
    fakBroj.CompareMode = vbTextCompare
    Dim fd As Variant, cFId As Long, cFBr As Long
    fd = GetTableData(TBL_FAKTURE)
    If IsArray(fd) Then fd = ExcludeStornirano(fd, TBL_FAKTURE)
    If IsArray(fd) Then
        cFId = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
        cFBr = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        Dim j As Long
        For j = 1 To UBound(fd, 1)
            If Not fakBroj.Exists(Trim$(CStr(nz(fd(j, cFId))))) Then _
                fakBroj.Add Trim$(CStr(nz(fd(j, cFId)))), Trim$(CStr(nz(fd(j, cFBr))))
        Next j
    End If

    Dim outA() As Variant, i As Long, n As Long
    Dim fakturisana As Boolean, fid As String
    ReDim outA(1 To UBound(pd, 1), 1 To 9)
    For i = 1 To UBound(pd, 1)
        If UCase$(Trim$(CStr(nz(pd(i, cStorno))))) = "DA" Then GoTo Sledeci
        fakturisana = False
        fid = ""
        If cFakt > 0 Then _
            fakturisana = (UCase$(Trim$(CStr(nz(pd(i, cFakt))))) = "DA")
        If cFakID > 0 Then fid = Trim$(CStr(nz(pd(i, cFakID))))
        n = n + 1
        outA(n, 1) = IdIliPrazno(brojac, Trim$(CStr(nz(pd(i, cId)))))
        outA(n, 2) = Trim$(CStr(nz(pd(i, cBroj)))) & "/" & _
                     Trim$(CStr(nz(pd(i, cGod))))
        outA(n, 3) = Trim$(CStr(nz(pd(i, cTip))))
        outA(n, 4) = pd(i, cDat)
        ' R5 (spoljna revizija #248): NIKAD Val(CStr(...)) za kolicine --
        ' nz radi CStr, pa 50.5 na srpskom locale-u postane "50,5" i Val
        ' procita 50: grid/korpa pokazu jedno, writer upise drugo. Isti
        ' IsNumeric obrazac kao FakD.
        outA(n, 5) = 0#
        outA(n, 6) = 0#
        outA(n, 7) = 0#
        If IsNumeric(pd(i, cNeto)) Then outA(n, 5) = CDbl(pd(i, cNeto))
        If IsNumeric(pd(i, cKut)) Then outA(n, 6) = CDbl(pd(i, cKut))
        If IsNumeric(pd(i, cKes)) Then outA(n, 7) = CDbl(pd(i, cKes))
        ' B2 kanonski contract: dostupna = nefakturisana AND bez veze
        ' na fakturu AND bez aktivne stavke -- stale FakturaID ili
        ' zaostala stavka NE sme ponovo u prodaju. Krug 4 P1: i bez
        ' imena proizvoda nije dostupna -- writer bi je ionako odbio,
        ' pa UI ne sme da kaze "moze" a finalni klik "ne moze".
        outA(n, 8) = (Not fakturisana) And Len(fid) = 0 _
                     And Not stAkt.Exists(Trim$(CStr(nz(pd(i, cId))))) _
                     And Len(Trim$(CStr(nz(pd(i, cTip))))) > 0
        If fakturisana And Len(fid) > 0 And fakBroj.Exists(fid) Then
            outA(n, 9) = CStr(fakBroj(fid))
        Else
            outA(n, 9) = ""
        End If
Sledeci:
    Next i
    If n = 0 Then Exit Function

    ' Isecanje na n redova radi pozivalac po n koji dobija? Ne -- vrati
    ' tacno n (obrazac IzvIseci nije ovde; ekran cita UBound).
    Dim res() As Variant, r As Long, c As Long
    ReDim res(1 To n, 1 To 9)
    For r = 1 To n
        For c = 1 To 9
            res(r, c) = outA(r, c)
        Next c
    Next r
    GetGPZaFakturisanjeForGrid = res
    Exit Function

EH:
    LogErr "modFaktura.GetGPZaFakturisanjeForGrid"
End Function

' Aktivne stavke faktura po preradi (B2 revizije #248): "aktivna" =
' nije stornirana i nije osirocena -- ISTI filter kao SEF mapper.
' Kolone meke: sveska pre nadogradnje nema GP stavke -> prazna mapa.
' Vraca dict preradaID -> broj aktivnih stavki.
Private Function GPAktivneStavkePoPreradi() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set GPAktivneStavkePoPreradi = d

    Dim colPre As Long
    colPre = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID)
    If colPre = 0 Then Exit Function

    Dim sd As Variant, i As Long, k As String
    sd = GetTableData(TBL_FAKTURA_STAVKE)
    If Not IsArray(sd) Then Exit Function

    Dim colSt As Long, colOs As Long
    colSt = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_STORNIRANO)
    colOs = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD)

    For i = 1 To UBound(sd, 1)
        k = Trim$(CStr(nz(sd(i, colPre))))
        If Len(k) > 0 Then
            If colSt > 0 Then
                If UCase$(Trim$(CStr(nz(sd(i, colSt))))) = "DA" Then GoTo Dalje
            End If
            If colOs > 0 Then
                If Len(Trim$(CStr(nz(sd(i, colOs))))) > 0 Then GoTo Dalje
            End If
            If d.Exists(k) Then
                d(k) = CLng(d(k)) + 1
            Else
                d.Add k, 1
            End If
        End If
Dalje:
    Next i
End Function

' ============================================================
' FAKTURA GOTOVE ROBE (GP grana). Stavka nosi PRERADU umesto
' prijemnice -- time se lanac sledljivosti zavrsava fakturom gotovog
' proizvoda kroz PODATKOVNU vezu (FakturaStavka.PreradaID), ne
' pogadjanjem. Novi writer, ne grana u CreateFaktura: izvori vrednosti
' su potpuno razliciti (kolicina = NetoIzlazKg prerade; CENA je unos
' operatera pri fakturisanju -- gotova roba nema evidentiranu cenu
' nigde u podacima, pa bi "izvedena" cena bila izmisljanje).
'
' stavke: Collection of Array(preradaID, cena). Kapije (u base, pod
' TX): prerada postoji i jedinstvena, nije stornirana, nije vec
' fakturisana; NetoIzlazKg > 0; cena > 0. Prerada se markira
' Fakturisano=Da + FakturaID -- ISTI obrazac kao prijemnica, pa storno
' fakture i sledljivost rade istim pravilima nad obe grane.
' ============================================================
Public Function CreateFakturaGP_TX(ByVal kupacID As String, _
                                   ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRERADA
    tx.AddTableSnapshot TBL_NOVAC

    CreateFakturaGP_TX = CreateFakturaGP(kupacID, stavke)

    If CreateFakturaGP_TX = "" Then
        Err.Raise vbObjectError + 1730, "CreateFakturaGP_TX", _
                  "CreateFakturaGP nije uspeo."
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="FAKTURA_GP_CREATE_SUCCESS", _
        severity:="INFO", _
        message:="GP faktura created successfully", _
        userId:="Operator", _
        moduleName:="modFaktura", _
        procedureName:="CreateFakturaGP_TX", _
        entityType:="Faktura", _
        entityID:=CreateFakturaGP_TX, _
        correlationId:=CreateFakturaGP_TX
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

    LogErr "CreateFakturaGP_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modFaktura", _
        procedureName:="CreateFakturaGP_TX", _
        entityType:="Faktura", _
        entityID:="", _
        correlationId:="CreateFakturaGP", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateFakturaGP_TX = ""
End Function

' Base -- NE zovi je spolja (isti razlog kao CreateFaktura: pola upisa
' bez transakcije).
Private Function CreateFakturaGP(ByVal kupacID As String, _
                                 ByVal stavke As Collection) As String
    Const SRC As String = "CreateFakturaGP"
    On Error GoTo EH

    If Trim$(kupacID) = "" Then
        Err.Raise vbObjectError + 1731, SRC, "KupacID je obavezan."
    End If
    ' Krug 4 P1: writer je samostalna granica -- GP nema prijemnicu
    ' cijim bi se vlasnistvom kupac implicitno proverio, pa se kupac
    ' proverava OVDE: mora postojati tacno jednom u tblKupci.
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
                  "Faktura mora imati bar jednu stavku."
    End If

    ' Fail-fast schema guards -- GP kolone su EnsurePaletniListSchema
    ' dodatak; bez njih se staje ODMAH, ne na pola upisa.
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE, SRC
    RequireColumnIndex TBL_PRERADA, COL_PRE_FAKTURISANO, SRC
    RequireColumnIndex TBL_PRERADA, COL_PRE_FAKTURA_ID, SRC

    Dim preData As Variant
    preData = GetTableData(TBL_PRERADA)
    If IsEmpty(preData) Then
        Err.Raise vbObjectError + 1734, SRC, "Tabela prerada je prazna."
    End If

    Dim colNetoIzlaz As Long, colBroj As Long, colGodina As Long
    Dim colFakturisano As Long, colStorno As Long
    Dim colFakVeza As Long, colTipGp As Long
    colNetoIzlaz = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
    colBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
    colGodina = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
    colFakturisano = RequireColumnIndex(TBL_PRERADA, COL_PRE_FAKTURISANO, SRC)
    colStorno = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, SRC)
    colFakVeza = RequireColumnIndex(TBL_PRERADA, COL_PRE_FAKTURA_ID, SRC)
    colTipGp = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)

    ' Kanonski GP contract (B2 revizije #248): DOSTUPNO za fakturisanje
    ' = Fakturisano <> Da AND FakturaID prazan AND nema aktivne stavke
    ' fakture sa ovim PreradaID. Mapa aktivnih stavki PRE petlje.
    Dim stavkeAkt As Object: Set stavkeAkt = GPAktivneStavkePoPreradi()

    ' Pre-validacija SVIH stavki pre ijednog upisa (obrazac CreateFaktura).
    Dim s As Variant, preradaID As String, cena As Double
    Dim rows As Collection, rowPre As Long
    Dim preRows As Object, preValues As Object
    Set preRows = CreateObject("Scripting.Dictionary")
    Set preValues = CreateObject("Scripting.Dictionary")

    For Each s In stavke
        preradaID = Trim$(CStr(s(0)))
        If Len(preradaID) = 0 Then
            Err.Raise vbObjectError + 1735, SRC, "PreradaID je obavezan."
        End If
        If Not IsNumeric(s(1)) Then
            Err.Raise vbObjectError + 1736, SRC, _
                      "Cena nije numericka. PreradaID=" & preradaID
        End If
        cena = CDbl(s(1))
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

        ' Inline "DA" provera: IsStorniranoValue je Private u modStorno
        ' (ne vidi se odavde -- ista klasa zamke kao NzS, v. cb8d4f5b).
        If UCase$(Trim$(CStr(nz(preData(rowPre, colStorno))))) = "DA" Then
            Err.Raise vbObjectError + 1740, SRC, _
                      "Prerada je stornirana: " & preradaID
        End If
        If UCase$(Trim$(CStr(nz(preData(rowPre, colFakturisano))))) = "DA" Then
            Err.Raise vbObjectError + 1741, SRC, _
                      "Prerada je vec fakturisana: " & preradaID
        End If
        ' B2: stale FakturaID uz Fakturisano=Ne bi nova faktura tiho
        ' PREGAZILA -- veza mora biti prazna, ne samo marker.
        If Len(Trim$(CStr(nz(preData(rowPre, colFakVeza))))) > 0 Then
            Err.Raise vbObjectError + 1748, SRC, _
                      "Prerada ima zaostalu vezu na fakturu (FakturaID nije prazan): " & preradaID
        End If
        If stavkeAkt.Exists(preradaID) Then
            Err.Raise vbObjectError + 1749, SRC, _
                      "Prerada vec ima aktivnu stavku fakture: " & preradaID
        End If
        ' P1: stavka PRODAJNE fakture mora imenovati proizvod -- prazan
        ' TipGotovogProizvoda bi dao fakturu bez naziva robe (print) i
        ' fail-open SEF fallback.
        If Len(Trim$(CStr(nz(preData(rowPre, colTipGp))))) = 0 Then
            Err.Raise vbObjectError + 1750, SRC, _
                      "TipGotovogProizvoda je prazan -- faktura mora imenovati proizvod: " & preradaID
        End If
        If Not IsNumeric(preData(rowPre, colNetoIzlaz)) Then
            Err.Raise vbObjectError + 1742, SRC, _
                      "NetoIzlazKg nije numericki. PreradaID=" & preradaID
        End If
        If CDbl(preData(rowPre, colNetoIzlaz)) <= 0 Then
            Err.Raise vbObjectError + 1742, SRC, _
                      "NetoIzlazKg mora biti veci od nule. PreradaID=" & preradaID
        End If

        preRows.Add preradaID, rowPre
        preValues.Add preradaID, Array( _
            CDbl(preData(rowPre, colNetoIzlaz)), cena, _
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

    ' Stavke: pozicioni deo je isti kao kod svezih (PrijemnicaID i
    ' BrojPrijemnice PRAZNI), a GP identitet ide PO IMENU u kolone koje
    ' je EnsurePaletniListSchema dodao NA KRAJ tabele -- pozicioni upis
    ' preko zatecenog kraja bi zavisio od redosleda (podaci-i-config).
    Dim stavkaID As String, stavkaNum As Long, stavkaRow As Variant
    Dim rowStavke As Long
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

        rowPre = CLng(preRows(preradaID))
        RequireUpdateCell TBL_PRERADA, rowPre, COL_PRE_FAKTURISANO, "Da", SRC
        RequireUpdateCell TBL_PRERADA, rowPre, COL_PRE_FAKTURA_ID, fakturaID, SRC
    Next s

    ' Avans se prebija isto kao kod svezih faktura.
    ApplyAvansToFaktura kupacID, fakturaID

    CreateFakturaGP = fakturaID
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

Private Function GenerateBrojFakture() As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    Dim currentYear As Long
    currentYear = Year(Date)

    Dim maxNum As Long

    If Not IsEmpty(data) Then
        Dim colBroj As Long
        colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                     "GenerateBrojFakture")

        Dim i As Long
        Dim broj As String
        Dim parts As Variant
        Dim num As Long
        Dim yr As Long

        For i = 1 To UBound(data, 1)
            broj = Trim$(CStr(data(i, colBroj)))

            If InStr(broj, "/") > 0 Then
                parts = Split(broj, "/")

                If UBound(parts) >= 1 Then
                    num = 0
                    yr = 0

                    If IsNumeric(Trim$(parts(0))) Then num = CLng(Trim$(parts(0)))
                    If IsNumeric(Trim$(parts(1))) Then yr = CLng(Trim$(parts(1)))

                    If yr = currentYear Then
                        If num > maxNum Then maxNum = num
                    End If
                End If
            End If
        Next i
    End If

    GenerateBrojFakture = CStr(maxNum + 1) & "/" & CStr(currentYear)
    Exit Function

EH:
    LogErr "GenerateBrojFakture"
    GenerateBrojFakture = ""
End Function

Public Sub PrintFaktura(ByVal fakturaID As String)
    On Error GoTo EH

    Dim fRow As Long
    fRow = RequireSingleFakturaRow(fakturaID, "PrintFaktura")

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        Err.Raise vbObjectError + 1732, "PrintFaktura", _
                  "Tabela faktura je prazna."
    End If

    Dim colFakBroj As Long
    Dim colFakDatum As Long
    Dim colFakKupac As Long
    Dim colFakIznos As Long
    Dim colFakStornirano As Long

    colFakBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                    "PrintFaktura")
    colFakDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, _
                                     "PrintFaktura")
    colFakKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, _
                                     "PrintFaktura")
    colFakIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, _
                                     "PrintFaktura")
    colFakStornirano = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, _
                                     "PrintFaktura")

    If UCase$(Trim$(CStr(data(fRow, colFakStornirano)))) = "DA" Then
        Err.Raise vbObjectError + 1736, "PrintFaktura", _
              Poruka("FAK_ERR_STORNIRANA_FAKTURA_MOZE") & fakturaID
    End If

    Dim kupacID As String
    kupacID = Trim$(CStr(data(fRow, colFakKupac)))

    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, COL_KUP_NAZIV))

    If kupacNaziv = "" Then kupacNaziv = kupacID

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then
        Err.Raise vbObjectError + 1734, "PrintFaktura", _
                  "Faktura nema stavke: " & fakturaID
    End If

    Dim colStFakID As Long, colStBrojPrij As Long, colStKlasa As Long
    Dim colStKol As Long, colStCena As Long
    colStFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, "PrintFaktura")
    colStBrojPrij = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRIJEMNICE, "PrintFaktura")
    colStKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, "PrintFaktura")
    colStKol = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KOLICINA, "PrintFaktura")
    colStCena = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_CENA, "PrintFaktura")

    ' R1 (revizija #248): GP stavka nosi PreradaID/BrojPrerade a
    ' prijemnicka polja su joj prazna -- bez ove grane bi GP faktura na
    ' papiru imala prazan "Broj prijemnice" i prazan "Klasa", bez imena
    ' proizvoda. GP kolone se citaju MEKO (sveska pre nadogradnje nema
    ' GP kolone ni GP fakture, pa je stari put netaknut).
    Dim colStPreID As Long, colStBrPre As Long
    colStPreID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID)
    colStBrPre = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE)

    ' Tip gotovog proizvoda po PreradaID -- mapa PRE petlje (S5).
    Dim gpTip As Object: Set gpTip = CreateObject("Scripting.Dictionary")
    gpTip.CompareMode = vbTextCompare
    If colStPreID > 0 Then
        Dim gpD As Variant, cGpId As Long, cGpTip As Long, g As Long
        gpD = GetTableData(TBL_PRERADA)
        If IsArray(gpD) Then
            cGpId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, "PrintFaktura")
            cGpTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, "PrintFaktura")
            For g = 1 To UBound(gpD, 1)
                If Not gpTip.Exists(Trim$(CStr(nz(gpD(g, cGpId))))) Then _
                    gpTip.Add Trim$(CStr(nz(gpD(g, cGpId)))), Trim$(CStr(nz(gpD(g, cGpTip))))
            Next g
        End If
    End If

    Dim stavke() As Variant
    ReDim stavke(1 To UBound(stavkeData, 1), 1 To 5)
    Dim outRow As Long, j As Long
    Dim kolicina As Double, cena As Double
    Dim preID As String, gpFaktura As Boolean
    For j = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(j, colStFakID))) = fakturaID Then
            outRow = outRow + 1
            kolicina = 0: cena = 0
            If IsNumeric(stavkeData(j, colStKol)) Then kolicina = CDbl(stavkeData(j, colStKol))
            If IsNumeric(stavkeData(j, colStCena)) Then cena = CDbl(stavkeData(j, colStCena))
            preID = ""
            If colStPreID > 0 Then preID = Trim$(CStr(nz(stavkeData(j, colStPreID))))
            If Len(preID) > 0 Then
                ' GP: dokument = broj prerade, proizvod = TipGotovogProizvoda.
                gpFaktura = True
                stavke(outRow, 1) = stavkeData(j, colStBrPre)
                If gpTip.Exists(preID) Then
                    stavke(outRow, 2) = CStr(gpTip(preID))
                Else
                    stavke(outRow, 2) = ""
                End If
            Else
                stavke(outRow, 1) = stavkeData(j, colStBrojPrij)
                stavke(outRow, 2) = stavkeData(j, colStKlasa)
            End If
            stavke(outRow, 3) = kolicina
            stavke(outRow, 4) = cena
            stavke(outRow, 5) = kolicina * cena
        End If
    Next j

    If outRow = 0 Then
        Err.Raise vbObjectError + 1735, "PrintFaktura", _
                  "Nisu pronadene stavke za fakturu: " & fakturaID
    End If

    Dim ukupno As Double
    If IsNumeric(data(fRow, colFakIznos)) Then ukupno = CDbl(data(fRow, colFakIznos))

    Dim ws As Worksheet
    Set ws = FillFakturaSablon(CStr(data(fRow, colFakBroj)), data(fRow, colFakDatum), _
                               kupacNaziv, stavke, outRow, ukupno, gpFaktura)
    If ws Is Nothing Then Exit Sub

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_FAKTURA_PRINT_MODE), "PRINT")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, ThisWorkbook.path & "\Faktura_" & _
                         Replace(CStr(data(fRow, colFakBroj)), "/", "-") & ".pdf", True
        ' OFF -> bez izlaza
    End Select
    Exit Sub

EH:
    ' Isti razlog kao kod citaca: LogErr brise Err. Zateceno, ali ulazi u
    ' popravku jer stampu sada zove radnja ekrana (fkprint), koja operateru
    ' prikazuje bas ovaj opis -- prazan opis ne kaze nista.
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    LogErr "PrintFaktura"
    Err.Raise errNum, "PrintFaktura", errDesc
End Sub

Public Sub UpdateFakturaStatus(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "UpdateFakturaStatus"

    If Trim$(fakturaID) = "" Then
        Err.Raise vbObjectError + 1723, SRC, _
                "FakturaID je obavezan."
    End If

    Dim colID As Long
    Dim colIznos As Long
    Dim colStatus As Long
    Dim colDatumPlacanja As Long
    Dim colStornirano As Long

    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, SRC)
    colDatumPlacanja = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM_PLACANJA, SRC)
    colStornirano = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, SRC)

    Dim r As Long
    r = RequireSingleFakturaRow(fakturaID, SRC)

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        Err.Raise vbObjectError + 1722, SRC, _
                  "Tabela faktura je prazna."
    End If

    If UCase$(Trim$(CStr(data(r, colStornirano)))) = "DA" Then
        Exit Sub
    End If

    If Not IsNumeric(data(r, colIznos)) Then
        Err.Raise vbObjectError + 1720, SRC, _
                  "Iznos fakture nije numericki: " & fakturaID
    End If

    Dim fakturaIznos As Double
    fakturaIznos = CDbl(data(r, colIznos))

    Dim uplaceno As Double
    uplaceno = GetUplataForFaktura(fakturaID)

    Dim currentStatus As String
    Dim currentDatumPlacanja As String

    currentStatus = Trim$(CStr(data(r, colStatus)))
    currentDatumPlacanja = Trim$(CStr(data(r, colDatumPlacanja)))

    If uplaceno >= fakturaIznos And fakturaIznos > 0 Then

        If currentStatus <> STATUS_PLACENO Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_STATUS, _
                              STATUS_PLACENO, SRC
        End If

        If Len(currentDatumPlacanja) = 0 Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_DATUM_PLACANJA, _
                              Date, SRC
        End If

    Else

        If currentStatus <> STATUS_NEPLACENO Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_STATUS, _
                              STATUS_NEPLACENO, SRC
        End If

        If Len(currentDatumPlacanja) > 0 Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_DATUM_PLACANJA, _
                              Empty, SRC
        End If

    End If

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

Private Function RequireSingleFakturaRow(ByVal fakturaID As String, _
                                         ByVal sourceName As String) As Long
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise vbObjectError + 1740, sourceName, _
                  "FakturaID je obavezan."
    End If

    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, sourceName

    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)

    If rows Is Nothing Then
        Err.Raise vbObjectError + 1741, sourceName, _
                  "FindRows je vratio Nothing za FakturaID=" & fakturaID
    End If

    If rows.count = 0 Then
        Err.Raise vbObjectError + 1742, sourceName, _
                  "Faktura nije pronadena: " & fakturaID
    End If

    If rows.count > 1 Then
        Err.Raise vbObjectError + 1743, sourceName, _
                  "Duplicate FakturaID. FakturaID=" & fakturaID & _
                  "; Count=" & CStr(rows.count)
    End If

    RequireSingleFakturaRow = CLng(rows(1))
End Function

Private Function IsPrijemnicaAvailableForFaktura(ByVal rowIndex As Long, _
                                                 ByVal prijemnicaID As String) As Boolean
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then Exit Function

    Dim colFakturisano As Long
    Dim colFakturaID As Long
    Dim colStorno As Long

    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                                        "IsPrijemnicaAvailableForFaktura")
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, _
                                      "IsPrijemnicaAvailableForFaktura")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "IsPrijemnicaAvailableForFaktura")

    If rowIndex <= 0 Or rowIndex > UBound(data, 1) Then Exit Function

    ' Pravilo je JEDNO i zivi u PrijemnicaDostupna. Do v6-ui-175 je stajalo
    ' samo ovde, pa ga citac mreze novog UI-ja nije imao odakle da uzme -- a
    ' prepisana kopija u ekranu bi se razisla sa kapijom koja stvarno odlucuje.
    IsPrijemnicaAvailableForFaktura = PrijemnicaDostupna( _
        CStr(data(rowIndex, colStorno)), _
        CStr(data(rowIndex, colFakturisano)), _
        CStr(data(rowIndex, colFakturaID)))
    Exit Function

EH:
    LogErr "IsPrijemnicaAvailableForFaktura"
    IsPrijemnicaAvailableForFaktura = False
End Function

Private Function GetPrijemnicaIDFromFakturaStavka(ByVal stavka As Variant, _
                                                  ByVal sourceName As String) As String
    On Error GoTo EH

    GetPrijemnicaIDFromFakturaStavka = Trim$(CStr(stavka(0)))

    If Len(GetPrijemnicaIDFromFakturaStavka) = 0 Then
        Err.Raise vbObjectError + 1718, sourceName, _
                  "Stavka nema PrijemnicaID."
    End If

    Exit Function

EH:
    Err.Raise vbObjectError + 1719, sourceName, _
              "Neispravan oblik stavke fakture. Ocekuje se da stavka(0) bude PrijemnicaID."
End Function

' ============================================================
' CITACI ZA MREZU NOVOG UI-JA (v6-ui-176, Faza E/16)
'
' Ekran modScrFakture ne cita tabele sam -- isto pravilo i isti oblik kao
' modAgrohemija.GetMagacinPrometForGrid / GetAgroDugoviForGrid.
'
' IDENTITET IDE U RED. Svaki citac vraca ID u PRVOJ koloni; ekran ga stavlja u
' skrivenu kolonu mreze (prioritet 4, mreza crta do 3). Prazan ID znaci
' DVOSMISLENO -- dva reda istog ID-a u tabeli -- i tada radnja ODBIJA da bira.
' To nije teorija: RequireSingleFakturaRow i CreateFaktura vec fail-close-uju
' na duplikat, pa bi radnja nad takvim redom svakako pukla; ovako pukne sa
' porukom operateru umesto sa greskom transakcije.
' ============================================================

' Pravilo "sme li ova prijemnica u fakturu", izdvojeno iz
' IsPrijemnicaAvailableForFaktura da bi ga imao i citac mreze. Prima VREDNOSTI
' celija, ne red -- pa ne mora da cita tabelu po redu (citac je vec ima).
Public Function PrijemnicaDostupna(ByVal stornirano As String, _
                                   ByVal fakturisano As String, _
                                   ByVal fakturaID As String) As Boolean
    If Trim$(stornirano) = "Da" Then Exit Function
    If Trim$(fakturisano) = "Da" Then Exit Function
    If Trim$(fakturaID) <> "" Then Exit Function
    PrijemnicaDostupna = True
End Function

' Koliko puta se svaki ID pojavljuje u koloni. Duplikat = dvosmislen identitet.
' JAVNA je zato sto je pravilo, ne pomocna rutina: bez ulaza se "dvosmislen
' prikaz nosi prazan identitet" moze izmeriti samo tako sto se u fixture
' namerno ubaci duplikat -- a duplikat PrijemnicaID obara kapije koje o njemu
' nista ne znaju (RequireSingleFakturaRow, CreateFaktura), pa bi jedan test
' kupio crvenilo tuceti drugih.
' Broji se nad SIROVOM tabelom, ne nad filtriranom: FindRows (koji na kraju
' odlucuje) gleda ceo list, pa i storniran red istog ID-a cini ID dvosmislenim.
Public Function BrojacIdova(ByVal tbl As String, ByVal colName As String) As Object
    Dim d As Object, data As Variant, c As Long, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    Set BrojacIdova = d
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    c = GetColumnIndex(tbl, colName)
    If c <= 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        k = Trim$(CStr(data(i, c)))
        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = CLng(d(k)) + 1
            Else
                d(k) = 1
            End If
        End If
    Next i
End Function

' Identitet reda: prazan kad se ID ponavlja u tabeli.
Public Function IdIliPrazno(ByVal brojac As Object, ByVal iD As String) As String
    If Len(iD) = 0 Then Exit Function
    If brojac Is Nothing Then Exit Function
    If Not brojac.Exists(iD) Then Exit Function
    If CLng(brojac(iD)) <> 1 Then Exit Function
    IdIliPrazno = iD
End Function

Private Function FakD(ByVal v As Variant) As Double
    If IsNumeric(v) Then FakD = CDbl(v)
End Function

' NEFAKTURISANE PRIJEMNICE JEDNOG KUPCA -- korpa ovog ekrana.
' Filter po kupcu radi POSTOJECI modDokumenta.GetPrijemniceByKupac (on vec
' izbacuje stornirane); ovde se samo prevodi u redove mreze.
'
'   1 PrijemnicaID (identitet) | 2 BrojPrijemnice | 3 BrojZbirne | 4 Datum
'   5 Klasa | 6 Kolicina | 7 Cena | 8 Vrednost | 9 Dostupna (Boolean)
'   10 BrojFakture (prazno kad nije fakturisana)
Public Function GetPrijemniceZaFakturisanjeForGrid(ByVal kupacID As String) As Variant
    On Error GoTo EH

    If Len(Trim$(kupacID)) = 0 Then
        GetPrijemniceZaFakturisanjeForGrid = Empty
        Exit Function
    End If

    Dim data As Variant
    data = modDokumenta.GetPrijemniceByKupac(kupacID)
    If Not IsArray(data) Then
        GetPrijemniceZaFakturisanjeForGrid = Empty
        Exit Function
    End If

    Const SRC As String = "GetPrijemniceZaFakturisanjeForGrid"

    Dim cID As Long, cBroj As Long, cZb As Long, cDat As Long, cKlasa As Long
    Dim cKol As Long, cCena As Long, cFakt As Long, cFakID As Long, cStorno As Long
    cID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    cBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
    cZb = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    cDat = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, SRC)
    cKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
    cKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
    cCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, SRC)
    cFakt = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, SRC)
    cFakID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, SRC)
    cStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, SRC)

    Dim brojac As Object
    Set brojac = BrojacIdova(TBL_PRIJEMNICA, COL_PRJ_ID)

    Dim brFak As Object
    Set brFak = BuildLookupDict(TBL_FAKTURE, COL_FAK_ID, COL_FAK_BROJ)

    Dim outA() As Variant, i As Long, n As Long, iD As String, fakID As String
    ReDim outA(1 To UBound(data, 1), 1 To 10)

    For i = 1 To UBound(data, 1)
        iD = Trim$(CStr(data(i, cID)))
        fakID = Trim$(CStr(data(i, cFakID)))
        n = n + 1
        outA(n, 1) = IdIliPrazno(brojac, iD)
        outA(n, 2) = CStr(data(i, cBroj))
        outA(n, 3) = CStr(data(i, cZb))
        outA(n, 4) = data(i, cDat)
        outA(n, 5) = CStr(data(i, cKlasa))
        outA(n, 6) = FakD(data(i, cKol))
        outA(n, 7) = FakD(data(i, cCena))
        outA(n, 8) = FakD(data(i, cKol)) * FakD(data(i, cCena))
        outA(n, 9) = PrijemnicaDostupna(CStr(data(i, cStorno)), _
                                        CStr(data(i, cFakt)), fakID)
        outA(n, 10) = ""
        If Len(fakID) > 0 Then
            If Not brFak Is Nothing Then
                If brFak.Exists(fakID) Then outA(n, 10) = CStr(brFak(fakID))
            End If
        End If
    Next i

    If n = 0 Then
        GetPrijemniceZaFakturisanjeForGrid = Empty
    Else
        GetPrijemniceZaFakturisanjeForGrid = outA
    End If
    Exit Function

EH:
    ' Err se cita PRE LogErr-a: LogError pocinje sa `On Error Resume Next`,
    ' a svaka On Error naredba brise Err. Bez ovoga bi Err.Raise dobio nulu i
    ' prazan opis, pa bi pad seme stigao do ekrana kao 'nema redova'.
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' IZDATE FAKTURE. Stornirane izbacuje ExcludeStornirano, uplate sabira
' modNovac.BuildUplataDictByFaktura -- isti primitiv koji koristi i
' GetOpenFakture, pa se "uplaceno" na dva mesta ne moze razici.
'
'   1 FakturaID (identitet) | 2 BrojFakture | 3 Datum | 4 KupacNaziv
'   5 Iznos | 6 Uplaceno | 7 Preostalo | 8 Status | 9 KupacID
'
' Kupac ide i kao NAZIV i kao ID: prikaz trazi naziv, a svako poredjenje
' (npr. slaganje sa GetOpenFakture, koji radi po kupcu) trazi identitet.
' Naziv nije identitet -- dva kupca smeju da se zovu isto.
Public Function GetFaktureForGrid() As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then
        GetFaktureForGrid = Empty
        Exit Function
    End If

    Dim brojac As Object
    Set brojac = BrojacIdova(TBL_FAKTURE, COL_FAK_ID)

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then
        GetFaktureForGrid = Empty
        Exit Function
    End If

    Const SRC As String = "GetFaktureForGrid"

    Dim cID As Long, cBroj As Long, cDat As Long, cKup As Long
    Dim cIznos As Long, cStatus As Long
    cID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    cBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, SRC)
    cDat = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, SRC)
    cKup = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, SRC)
    cIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)
    cStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, SRC)

    Dim uplate As Object
    Set uplate = BuildUplataDictByFaktura()

    Dim kupci As Object
    Set kupci = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)

    Dim outA() As Variant, i As Long, n As Long
    Dim iD As String, kupID As String, iznos As Double, upl As Double
    ReDim outA(1 To UBound(data, 1), 1 To 9)

    For i = 1 To UBound(data, 1)
        iD = Trim$(CStr(data(i, cID)))
        kupID = Trim$(CStr(data(i, cKup)))
        iznos = FakD(data(i, cIznos))
        upl = 0
        If Not uplate Is Nothing Then
            If uplate.Exists(iD) Then upl = FakD(uplate(iD))
        End If
        n = n + 1
        outA(n, 1) = IdIliPrazno(brojac, iD)
        outA(n, 2) = CStr(data(i, cBroj))
        outA(n, 3) = data(i, cDat)
        outA(n, 4) = kupID
        If Not kupci Is Nothing Then
            If kupci.Exists(kupID) Then outA(n, 4) = CStr(kupci(kupID))
        End If
        outA(n, 5) = iznos
        outA(n, 6) = upl
        outA(n, 7) = iznos - upl
        outA(n, 8) = CStr(data(i, cStatus))
        outA(n, 9) = kupID
    Next i

    If n = 0 Then
        GetFaktureForGrid = Empty
    Else
        GetFaktureForGrid = outA
    End If
    Exit Function

EH:
    ' Err se cita PRE LogErr-a: LogError pocinje sa `On Error Resume Next`,
    ' a svaka On Error naredba brise Err. Bez ovoga bi Err.Raise dobio nulu i
    ' prazan opis, pa bi pad seme stigao do ekrana kao 'nema redova'.
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' Je li OVA instalacija uopste povezana na SEF. Bez baze i kljuca svaka SEF
' radnja moze samo da padne, pa ekran listu SEF-a tada i ne nudi. Config je
' izvor istine, ne prisustvo modula -- moduli su u svakom buildu.
Public Function SEFKonfigurisan() As Boolean
    On Error Resume Next
    SEFKonfigurisan = (Len(Trim$(GetConfigValue("SEF_BASE_URL"))) > 0) And _
                      (Len(Trim$(GetConfigValue("SEF_API_KEY"))) > 0)
    Err.Clear
End Function

' STANJE ELEKTRONSKIH FAKTURA. SEF kolone se citaju MEKO (GetColumnIndex, ne
' RequireColumnIndex): sema je izvor istine, a instalacija bez SEF kolona sme
' da vidi prazan SEF a ne rusenje ekrana. Nazivi su isti literali koje koristi
' modSEFPersistance -- taj modul se po zadatku ne dira, pa se ni konstante
' odatle ne mogu uvesti.
'
'   1 FakturaID (identitet) | 2 BrojFakture | 3 KupacNaziv | 4 Iznos
'   5 SEFWorkflowState | 6 SEFDocumentId | 7 SEFSentAt | 8 SEFLastErrorMessage
Public Function GetFaktureSEFForGrid() As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then
        GetFaktureSEFForGrid = Empty
        Exit Function
    End If

    Dim brojac As Object
    Set brojac = BrojacIdova(TBL_FAKTURE, COL_FAK_ID)

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then
        GetFaktureSEFForGrid = Empty
        Exit Function
    End If

    Const SRC As String = "GetFaktureSEFForGrid"

    Dim cID As Long, cBroj As Long, cKup As Long, cIznos As Long
    cID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    cBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, SRC)
    cKup = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, SRC)
    cIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)

    Dim cWf As Long, cDoc As Long, cSent As Long, cErr As Long
    cWf = GetColumnIndex(TBL_FAKTURE, "SEFWorkflowState")
    cDoc = GetColumnIndex(TBL_FAKTURE, "SEFDocumentId")
    cSent = GetColumnIndex(TBL_FAKTURE, "SEFSentAt")
    cErr = GetColumnIndex(TBL_FAKTURE, "SEFLastErrorMessage")

    ' Bez kolone stanja lista nema sta da pokaze -- prazno, ne izmisljeno.
    If cWf <= 0 Then
        GetFaktureSEFForGrid = Empty
        Exit Function
    End If

    Dim kupci As Object
    Set kupci = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)

    Dim outA() As Variant, i As Long, n As Long, iD As String, kupID As String
    ReDim outA(1 To UBound(data, 1), 1 To 8)

    For i = 1 To UBound(data, 1)
        iD = Trim$(CStr(data(i, cID)))
        kupID = Trim$(CStr(data(i, cKup)))
        n = n + 1
        outA(n, 1) = IdIliPrazno(brojac, iD)
        outA(n, 2) = CStr(data(i, cBroj))
        outA(n, 3) = kupID
        If Not kupci Is Nothing Then
            If kupci.Exists(kupID) Then outA(n, 3) = CStr(kupci(kupID))
        End If
        outA(n, 4) = FakD(data(i, cIznos))
        outA(n, 5) = Trim$(CStr(data(i, cWf)))
        outA(n, 6) = ""
        outA(n, 7) = ""
        outA(n, 8) = ""
        If cDoc > 0 Then outA(n, 6) = Trim$(CStr(data(i, cDoc)))
        If cSent > 0 Then outA(n, 7) = data(i, cSent)
        If cErr > 0 Then outA(n, 8) = Trim$(CStr(data(i, cErr)))
        ' Faktura koja jos nije ni pripremljena za SEF nosi lokalno stanje;
        ' prazno polje se prikazuje kao ono sto jeste -- neposlata.
        If Len(CStr(outA(n, 5))) = 0 Then outA(n, 5) = WF_LOCAL_FINALIZED
    Next i

    If n = 0 Then
        GetFaktureSEFForGrid = Empty
    Else
        GetFaktureSEFForGrid = outA
    End If
    Exit Function

EH:
    ' Err se cita PRE LogErr-a: LogError pocinje sa `On Error Resume Next`,
    ' a svaka On Error naredba brise Err. Bez ovoga bi Err.Raise dobio nulu i
    ' prazan opis, pa bi pad seme stigao do ekrana kao 'nema redova'.
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function
