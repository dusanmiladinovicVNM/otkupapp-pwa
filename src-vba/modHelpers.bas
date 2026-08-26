Attribute VB_Name = "modHelpers"
'Attribute VB_Name = "modHelpers"
Option Explicit

' Format kolicine (kg) za izvestaje: UVEK tacno dve decimale
' (1000 -> "1.000,00"; 1234.5 -> "1.234,50"). Fiksni "#,##0.00" je
' lokalno-bezbedan (uvek 2 cifre iza zareza, bez praznog decimalnog zareza
' "500," koji je pravio opcioni ".##"). Jedinstveni izvor istine za kg prikaz.
Public Function FmtKolicina(ByVal X As Double) As String
    FmtKolicina = Format$(X, "#,##0.00")
End Function

Public Function ExtractIDFromDisplay(ByVal displayText As String) As String
    ' Unterst?tzt: "ID - Name" und "(ID) Name"
    Dim dashPos As Long
    dashPos = InStr(displayText, " - ")
    If dashPos > 0 Then
        ExtractIDFromDisplay = Left$(displayText, dashPos - 1)
        Exit Function
    End If
    
    Dim startPos As Long, endPos As Long
    startPos = InStr(displayText, "(")
    endPos = InStr(displayText, ")")
    If startPos > 0 And endPos > startPos Then
        ExtractIDFromDisplay = Mid$(displayText, startPos + 1, endPos - startPos - 1)
        Exit Function
    End If
    
    ExtractIDFromDisplay = displayText
End Function

Public Function GetVozacDisplayList() As Variant
    Dim data As Variant
    data = GetTableData(TBL_VOZACI)
    If IsEmpty(data) Then
        GetVozacDisplayList = Array()
        Exit Function
    End If
    
    Dim colID As Long, colIme As Long, colPrezime As Long, colAktivan As Long
    colID = GetColumnIndex(TBL_VOZACI, "VozacID")
    colIme = GetColumnIndex(TBL_VOZACI, "Ime")
    colPrezime = GetColumnIndex(TBL_VOZACI, "Prezime")
    colAktivan = GetColumnIndex(TBL_VOZACI, "Aktivan")
    
    Dim result() As String
    Dim count As Long
    ReDim result(0 To UBound(data, 1) - 1)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colAktivan)) = STATUS_AKTIVAN Then
            result(count) = CStr(data(i, colIme)) & " " & _
                           CStr(data(i, colPrezime)) & " (" & _
                           CStr(data(i, colID)) & ")"
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        GetVozacDisplayList = Array()
    Else
        ReDim Preserve result(0 To count - 1)
        GetVozacDisplayList = result
    End If
End Function

Public Sub FillCmb(ByRef cmb As MSForms.ComboBox, ByVal items As Variant)
    cmb.Clear
    If IsEmpty(items) Then Exit Sub
    If Not IsArray(items) Then Exit Sub
    Dim i As Long
    For i = LBound(items) To UBound(items)
        If CStr(items(i)) <> "" Then cmb.AddItem CStr(items(i))
    Next i
End Sub

Public Sub FillComboKooperantiByStanica(ByRef cmb As MSForms.ComboBox, ByVal stanicaID As String)
    cmb.Clear

    ' 2-kolonski: kol0 = "Ime Prezime" (vidljivo, filter po imenu),
    ' kol1 = KooperantID (skriveno). BoundColumn=1 -> .Value je ime
    ' (omogucava i slobodan unos novog imena). GetComboID cita kol1.
    On Error Resume Next
    cmb.ColumnCount = 2
    cmb.ColumnWidths = ";0"
    cmb.TextColumn = 1
    cmb.BoundColumn = 1
    cmb.MatchEntry = fmMatchEntryComplete
    cmb.MatchRequired = False
    On Error GoTo 0

    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Sub

    Dim colID As Long, colIme As Long, colPrezime As Long, colStanica As Long
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    colStanica = GetColumnIndex(TBL_KOOPERANTI, "StanicaID")
    Dim colAkt As Long: colAkt = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")

    Dim names() As String, ids() As String
    ReDim names(1 To UBound(data, 1))
    ReDim ids(1 To UBound(data, 1))
    Dim n As Long: n = 0

    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' stanicaID = "" -> svi kooperanti (toggle KOOP_FILTER_BY_OM = OFF)
        Dim aktOk As Boolean: aktOk = True
        If colAkt > 0 Then
            If StrComp(Trim$(CStr(data(i, colAkt))), STATUS_NEAKTIVAN, vbTextCompare) = 0 Then aktOk = False
        End If
        If (stanicaID = "" Or CStr(data(i, colStanica)) = stanicaID) And aktOk Then
            n = n + 1
            names(n) = Trim$(CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime)))
            ids(n) = CStr(data(i, colID))
        End If
    Next i
    If n = 0 Then Exit Sub

    ' insertion sort po imenu (case-insensitive)
    Dim a As Long, b As Long
    For a = 2 To n
        Dim kn As String: kn = names(a)
        Dim ki As String: ki = ids(a)
        b = a - 1
        Do While b >= 1
            If LCase$(names(b)) <= LCase$(kn) Then Exit Do
            names(b + 1) = names(b): ids(b + 1) = ids(b): b = b - 1
        Loop
        names(b + 1) = kn: ids(b + 1) = ki
    Next a

    For a = 1 To n
        cmb.AddItem names(a)
        cmb.List(cmb.ListCount - 1, 1) = ids(a)
    Next a
End Sub

Public Function ExcludeStornirano(ByVal data As Variant, _
                                  ByVal tblName As String) As Variant
    ' Filtert Stornirano="Da" Zeilen raus, gibt bereinigtes Array zur?ck
    '
    ' KLASIFIKACIJA I OBAVEZNA KOLONA IDU PRE SVEGA -- i pre IsEmpty.
    '
    ' Prazna tabela iz registra kojoj nedostaje kolona je i dalje DRIFT, pa mora
    ' da padne: ugovor je "tabela iz STORNO_TABELE bez kolone znaci drift", a ne
    ' "drift se prijavljuje samo dok tabela ima redova". Ranije je IsEmpty izlazio
    ' prvi, pa se schema guard nad praznom tabelom nikad nije ni pitao.
    RequireStornoKlasifikaciju tblName, "modHelpers.ExcludeStornirano"

    If Not modSchemaGuard.TabelaNosiStorno(tblName) Then
        ' EKSPLICITNO BEZ_STORNA -- maticni podaci. Prolaz je TACAN ishod, a ne
        ' posledica toga sto tabela nije prepoznata: to je kapija iznad vec
        ' razdvojila.
        ExcludeStornirano = data
        Exit Function
    End If

    ' Tabela JESTE iz registra: kolona mora da postoji, bez obzira na broj redova.
    ' Trazi se JEDNOM, ovde, pa se indeks nosi dalje -- nula ispod vise nije
    ' moguca, jer RequireColumnIndex pada umesto da je vrati.
    Dim colStorno As Long
    colStorno = RequireColumnIndex(tblName, COL_STORNIRANO, _
                                   "modHelpers.ExcludeStornirano")

    If IsEmpty(data) Then
        ExcludeStornirano = data
        Exit Function
    End If

    ' Request-scoped kes: samo kad je 'data' PUNA tabela (isti broj redova kao
    ' kesiran GetTableData) -> filtrirani podskupovi se NE kesiraju po tblName.
    Dim canCache As Boolean: canCache = False
    If ExclCacheActive() And IsArray(data) Then
        Dim fullData As Variant: fullData = GetTableData(tblName)   ' iz kesa (mTableCache)
        If IsArray(fullData) Then
            If UBound(data, 1) = UBound(fullData, 1) Then canCache = True
        End If
    End If
    If canCache Then
        Dim cached As Variant
        If ExclCacheTryGet(tblName, cached) Then
            ExcludeStornirano = cached
            Exit Function
        End If
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init colStorno, "<>", "Da"
    filters.Add fp

    ExcludeStornirano = FilterArray(data, filters)
    If canCache Then ExclCachePut tblName, ExcludeStornirano
End Function

Public Function SafeGetTable(ByVal tableName As String) As ListObject
    On Error Resume Next
    Set SafeGetTable = GetTable(tableName)
    On Error GoTo 0
End Function


Public Function nz(ByVal val As Variant, Optional ByVal default As String = "") As String
    If IsEmpty(val) Or IsNull(val) Then
        nz = default
    Else
        nz = CStr(val)
    End If
End Function

' Null-safe pretvaranje vrednosti iz tabele/celije u tekst.
' Vraca "" za Null/Empty/Error; inace CStr(v). Pozivaoci sami rade Trim$ gde treba.
Public Function NzToText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzToText = ""
    ElseIf isError(v) Then
        NzToText = ""
    Else
        NzToText = CStr(v)
    End If
End Function

' Klasa dokumenta sa defaultom. Prazna kolona (starija sema / rucni unos) znaci
' Klasa I -- ista konvencija koju vec primenjuje auto-lanac hladnjace.
Public Function KlasaOrDefault(ByVal v As Variant) As String
    KlasaOrDefault = Trim$(NzToText(v))
    If Len(KlasaOrDefault) = 0 Then KlasaOrDefault = KLASA_I
End Function

' VLASNIK zbirne (ko je nosilac dokumenta) = BrojZbirne + VozacID + KupacID.
' ISTA definicija koju koristi modStorno.RequireJedanVlasnikPoBroju (AUD-052 /
' RF-05: broj se generise po vozacu, a dokument pripada kupcu). NE uvoditi drugu.
Public Function ZbirnaVlasnikKljuc(ByVal brojZbirne As String, _
                                   ByVal vozacID As String, _
                                   ByVal kupacID As String) As String
    ZbirnaVlasnikKljuc = Trim$(brojZbirne) & "|" & Trim$(vozacID) & "|" & Trim$(kupacID)
End Function

' STAVKA dokumenta = vlasnik + Klasa. Klasa I i II istog dokumenta dele broj,
' vozaca i kupca, ali imaju ZASEBNU otpremnicu, zbirnu i prijemnicu (auto-lanac
' hladnjace ih vodi kroz ceo lanac odvojeno), pa je (BrojZbirne + Klasa) prava
' granularnost za racun manjka -- isti kljuc koji auto-lanac vec koristi za
' idempotenciju (modAutoHladnjaca.KeyZbrKlasa). Bez Klase u kljucu se prijem
' obe klase sabere pa dodeli SVAKOJ klasi ponaosob.
Public Function ZbirnaStavkaKljuc(ByVal vlasnikKljuc As String, _
                                  ByVal klasa As String) As String
    ZbirnaStavkaKljuc = vlasnikKljuc & "|" & KlasaOrDefault(klasa)
End Function

Public Function BuildManjakDict() As Object
    ' Agregacija zbirna/prijemnica za racun manjka -- po VLASNIKU, ne po broju.
    '
    ' RF-06/AUD-023: raniji oblik je grupisao ISKLJUCIVO po `BrojZbirne`, a
    ' RF-05/AUD-052 je vec dokazao da poslovni broj NIJE identitet -- dve aktivne
    ' zbirne mogu deliti isti broj (razliciti vozac/kupac). Posledica je bila
    ' sabiranje tudje prijemne kolicine, pa pogresan manjak i procenat.
    '
    ' Granularnost je STAVKA (vlasnik + Klasa), ne dokument: Klasa I i II dele
    ' broj/vozaca/kupca ali imaju zasebnu otpremnicu, zbirnu i prijemnicu, pa bi
    ' kljuc bez Klase sabrao prijem obe klase i taj zbir dodelio svakoj ponaosob.
    '
    ' Returns: Dictionary sa dve vrste kljuceva
    '   <stavkaKljuc>               -> Array(ZbirnaKg, PrijemnicaKg, CntPrijem)
    '                                  stavkaKljuc = broj|vozac|kupac|klasa
    '   "#V|" & broj                -> Long: koliko RAZLICITIH aktivnih VLASNIKA
    '                                  (vozac+kupac, BEZ klase) nosi taj BrojZbirne;
    '                                  1 = broj je pouzdan. Dve klase istog
    '                                  dokumenta NISU dva vlasnika.
    '   "#1|" & broj                -> vlasnikKljuc (samo kad je vlasnik jedinstven)
    '   "#O|" & broj & "|" & vozac  -> vlasnikKljuc (samo kad je po tom vozacu
    '                                  razresenje jednoznacno)
    '   "#N|" & broj                -> Long: prijemnice tog broja koje NEMAJU
    '                                  kompletnog vlasnika (prazan vozac ili kupac,
    '                                  ili kolona ne postoji) -> fail-closed signal
    '   "#C|" & broj & "|" & klasa  -> Long/Double: prijem agregiran po (broj, klasa)
    '   "#K|" & broj & "|" & klasa     BEZ vlasnika. Sme se koristiti iskljucivo kad
    '                                  je "#V|" = 1, tj. kad je dokazano da broj ima
    '                                  jednog vlasnika; pokriva starije prijemnice
    '                                  kojima vlasnik nije popunjen. Isti (broj,
    '                                  klasa) kljuc koji auto-lanac hladnjace vec
    '                                  koristi za idempotenciju.
    ' Prefiks "#" se ne moze pojaviti u BrojZbirne (`x/ddmmyy[-rb]`), pa nema
    ' kolizije sa stavka kljucevima.

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' --- Zbirna ---
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If Not IsArray(zbrData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If Not IsArray(zbrData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If

    Dim colBroj As Long, colZbrKol As Long, colZbrVoz As Long, colZbrKup As Long
    Dim colZbrKla As Long
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    colZbrVoz = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC)
    colZbrKup = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC)
    colZbrKla = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KLASA)

    ' broj -> Dictionary(vlasnikKljuc -> vozacID); sluzi za brojanje VLASNIKA
    ' (bez klase!) i za razresenje po vozacu (otpremnica nema KupacID).
    Dim vlasniciPoBroju As Object
    Set vlasniciPoBroju = CreateObject("Scripting.Dictionary")

    Dim z As Long
    For z = 1 To UBound(zbrData, 1)
        Dim brZbr As String
        brZbr = Trim$(NzToText(zbrData(z, colBroj)))

        Dim zVoz As String, zKup As String, zKla As String
        zVoz = ""
        zKup = ""
        zKla = KLASA_I
        If colZbrVoz > 0 Then zVoz = Trim$(NzToText(zbrData(z, colZbrVoz)))
        If colZbrKup > 0 Then zKup = Trim$(NzToText(zbrData(z, colZbrKup)))
        If colZbrKla > 0 Then zKla = KlasaOrDefault(zbrData(z, colZbrKla))

        Dim vk As String
        vk = ZbirnaVlasnikKljuc(brZbr, zVoz, zKup)

        Dim sk As String
        sk = ZbirnaStavkaKljuc(vk, zKla)

        If Not dict.Exists(sk) Then dict.Add sk, Array(0#, 0#, 0&)
        Dim vals As Variant
        vals = dict(sk)
        If IsNumeric(zbrData(z, colZbrKol)) Then vals(0) = vals(0) + CDbl(zbrData(z, colZbrKol))
        dict(sk) = vals

        If Not vlasniciPoBroju.Exists(brZbr) Then
            vlasniciPoBroju.Add brZbr, CreateObject("Scripting.Dictionary")
        End If
        Dim vSet As Object
        Set vSet = vlasniciPoBroju(brZbr)
        If Not vSet.Exists(vk) Then vSet.Add vk, zVoz
    Next z

    ' --- Indeks vlasnika po broju (#V / #1 / #O) ---
    Dim bKey As Variant
    For Each bKey In vlasniciPoBroju.keys
        Set vSet = vlasniciPoBroju(bKey)
        dict("#V|" & CStr(bKey)) = CLng(vSet.count)

        If vSet.count = 1 Then
            Dim onlyKey As Variant
            For Each onlyKey In vSet.keys
                dict("#1|" & CStr(bKey)) = CStr(onlyKey)
            Next onlyKey
        Else
            ' Razresenje po vozacu: upisi samo ako je za tog vozaca JEDINSTVENO.
            Dim perVoz As Object
            Set perVoz = CreateObject("Scripting.Dictionary")
            Dim vKey As Variant
            For Each vKey In vSet.keys
                Dim vz As String
                vz = CStr(vSet(vKey))
                If Not perVoz.Exists(vz) Then
                    perVoz.Add vz, CStr(vKey)
                Else
                    perVoz(vz) = ""          ' vise vlasnika istog vozaca -> nejednoznacno
                End If
            Next vKey
            For Each vKey In perVoz.keys
                If Len(CStr(perVoz(vKey))) > 0 Then
                    dict("#O|" & CStr(bKey) & "|" & CStr(vKey)) = CStr(perVoz(vKey))
                End If
            Next vKey
        End If
    Next bKey

    ' --- Prijemnica ---
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then
        Set BuildManjakDict = dict
        Exit Function
    End If

    Dim colPBrZbr As Long, colPKol As Long, colPVoz As Long, colPKup As Long
    Dim colPKla As Long
    colPBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    colPKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colPVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    colPKup = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    colPKla = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)

    Dim p As Long
    For p = 1 To UBound(prijData, 1)
        Dim pZbr As String
        pZbr = Trim$(NzToText(prijData(p, colPBrZbr)))
        If Len(pZbr) = 0 Then GoTo NextPrijem
        If Not vlasniciPoBroju.Exists(pZbr) Then GoTo NextPrijem

        Dim pVoz As String, pKup As String, pKla As String
        pVoz = ""
        pKup = ""
        pKla = KLASA_I
        If colPVoz > 0 Then pVoz = Trim$(NzToText(prijData(p, colPVoz)))
        If colPKup > 0 Then pKup = Trim$(NzToText(prijData(p, colPKup)))
        If colPKla > 0 Then pKla = KlasaOrDefault(prijData(p, colPKla))

        Dim pKg As Double
        pKg = 0
        If IsNumeric(prijData(p, colPKol)) Then pKg = CDbl(prijData(p, colPKol))

        ' Agregat po (broj, klasa) bez vlasnika -- validan samo kad broj ima
        ' jednog vlasnika. Klasa MORA biti u kljucu, inace se prijem Klase I i II
        ' sabere pa taj zbir dodeli svakoj klasi ponaosob.
        Dim bkKey As String
        bkKey = pZbr & "|" & pKla
        If Not dict.Exists("#C|" & bkKey) Then dict.Add "#C|" & bkKey, 0&
        dict("#C|" & bkKey) = CLng(dict("#C|" & bkKey)) + 1
        If Not dict.Exists("#K|" & bkKey) Then dict.Add "#K|" & bkKey, 0#
        dict("#K|" & bkKey) = CDbl(dict("#K|" & bkKey)) + pKg

        ' Prijemnica bez kompletnog vlasnika se NE sme pripisati nijednoj zbirnoj
        ' kad broj nosi vise vlasnika -- broji se posebno kao fail-closed signal.
        If Len(pVoz) = 0 Or Len(pKup) = 0 Then
            If Not dict.Exists("#N|" & pZbr) Then dict.Add "#N|" & pZbr, 0&
            dict("#N|" & pZbr) = CLng(dict("#N|" & pZbr)) + 1
        End If

        Dim psk As String
        psk = ZbirnaStavkaKljuc(ZbirnaVlasnikKljuc(pZbr, pVoz, pKup), pKla)
        If dict.Exists(psk) Then
            vals = dict(psk)
            vals(1) = vals(1) + pKg
            vals(2) = CLng(vals(2)) + 1
            dict(psk) = vals
        End If
NextPrijem:
    Next p

    Set BuildManjakDict = dict
End Function


Public Function CheckVerwaisteDokumente() As String
    Dim warnings As String
    
    ' 1. Otkup ohne OtpremnicaID
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsArray(otkupData) Then
        otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
        If IsArray(otkupData) Then
            Dim colOtpID As Long, colOtkID As Long, colOtkKol As Long, colOtkBrDok As Long
            colOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
            colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
            colOtkKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
            colOtkBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
            
            Dim cntOtkup As Long, detailOtkup As String, i As Long
            For i = 1 To UBound(otkupData, 1)
                If CStr(otkupData(i, colOtpID)) = "" Then
                    cntOtkup = cntOtkup + 1
                    If cntOtkup <= 40 Then
                        detailOtkup = detailOtkup & "  " & CStr(otkupData(i, colOtkID)) & _
                                      " (" & CStr(otkupData(i, colOtkBrDok)) & ") " & _
                                      Format$(CDbl(otkupData(i, colOtkKol)), "#,##0") & "kg" & vbCrLf
                    End If
                End If
            Next i
            If cntOtkup > 0 Then
                warnings = warnings & cntOtkup & " otkup(a) bez otpremnice:" & vbCrLf & detailOtkup
                If cntOtkup > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    ' 2. Otpremnice ohne BrojZbirne
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsArray(otpData) Then
        otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
        If IsArray(otpData) Then
            Dim colOtpZbr As Long, colOtpBroj As Long, colOtpKol As Long, colOtpAmb As Long
            colOtpZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
            colOtpBroj = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
            colOtpKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
            colOtpAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB)
            
            Dim cntOtp As Long, detailOtp As String
            For i = 1 To UBound(otpData, 1)
                If CStr(otpData(i, colOtpZbr)) = "" Then
                    cntOtp = cntOtp + 1
                    If cntOtp <= 40 Then
                        detailOtp = detailOtp & "  " & CStr(otpData(i, colOtpBroj)) & " " & _
                                    Format$(CDbl(otpData(i, colOtpKol)), "#,##0") & "kg " & _
                                    Format$(CLng(otpData(i, colOtpAmb)), "#,##0") & " amb" & vbCrLf
                    End If
                End If
            Next i
            If cntOtp > 0 Then
                warnings = warnings & cntOtp & " otpremnica(e) bez zbirne:" & vbCrLf & detailOtp
                If cntOtp > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    ' 2. Verwaiste Otpremnice (stornierte Zbirna)
    Dim verwOtp As Variant
    verwOtp = GetVerwaisteDokumente("Otpremnica")
    If IsArray(verwOtp) Then
        warnings = warnings & UBound(verwOtp, 1) & " otpremnica(e) sa storniranom zbirnom:" & vbCrLf
        Dim o As Long
        For o = 1 To IIf(UBound(verwOtp, 1) > 5, 5, UBound(verwOtp, 1))
            warnings = warnings & "  " & CStr(verwOtp(o, 2)) & " (Zbr:" & CStr(verwOtp(o, 3)) & ") " & _
                       Format$(CDbl(verwOtp(o, 5)), "#,##0") & "kg" & vbCrLf
        Next o
        If UBound(verwOtp, 1) > 40 Then warnings = warnings & "  ..." & vbCrLf
    End If
    
    ' 3. Verwaiste Prijemnice (stornierte Zbirna)
    Dim verwPrij As Variant
    verwPrij = GetVerwaisteDokumente("Prijemnica")
    If IsArray(verwPrij) Then
        warnings = warnings & UBound(verwPrij, 1) & " prijemnica(e) sa storniranom zbirnom:" & vbCrLf
        Dim pr As Long
        For pr = 1 To IIf(UBound(verwPrij, 1) > 5, 5, UBound(verwPrij, 1))
            warnings = warnings & "  " & CStr(verwPrij(pr, 2)) & " (Zbr:" & CStr(verwPrij(pr, 3)) & ") " & _
                       Format$(CDbl(verwPrij(pr, 5)), "#,##0") & "kg" & vbCrLf
        Next pr
        If UBound(verwPrij, 1) > 40 Then warnings = warnings & "  ..." & vbCrLf
    End If
    
    ' 4. Zbirna ohne Prijemnica
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then
        zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
        If IsArray(zbrData) Then
            Dim prijData As Variant
            prijData = GetTableData(TBL_PRIJEMNICA)
            Dim prijDict As Object
            Set prijDict = CreateObject("Scripting.Dictionary")
            If IsArray(prijData) Then
                prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
                If IsArray(prijData) Then
                    Dim colPZbr As Long
                    colPZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
                    Dim p As Long
                    For p = 1 To UBound(prijData, 1)
                        Dim pKey As String
                        pKey = CStr(prijData(p, colPZbr))
                        If Not prijDict.Exists(pKey) Then prijDict.Add pKey, True
                    Next p
                End If
            End If
            
            Dim colZBroj As Long, colZKol As Long, colZAmb As Long
            colZBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
            colZKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
            colZAmb = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
            
            Dim cntZbr As Long, detailZbr As String
            Dim z As Long
            For z = 1 To UBound(zbrData, 1)
                Dim zBroj As String
                zBroj = CStr(zbrData(z, colZBroj))
                If Not prijDict.Exists(zBroj) Then
                    cntZbr = cntZbr + 1
                    If cntZbr <= 40 Then
                        Dim zKg As String: zKg = ""
                        If IsNumeric(zbrData(z, colZKol)) Then zKg = Format$(CDbl(zbrData(z, colZKol)), "#,##0") & "kg"
                        Dim zAmb As String: zAmb = ""
                        If IsNumeric(zbrData(z, colZAmb)) Then zAmb = Format$(CLng(zbrData(z, colZAmb)), "#,##0") & " amb"
                        detailZbr = detailZbr & "  " & zBroj & " " & zKg & " " & zAmb & vbCrLf
                    End If
                End If
            Next z
            
            If cntZbr > 0 Then
                warnings = warnings & cntZbr & " zbirna(e) bez prijemnice:" & vbCrLf & detailZbr
                If cntZbr > 40 Then warnings = warnings & "  ..." & vbCrLf
            End If
        End If
    End If
    
    ' 5. Izgubljeni otkup blokovi (OtpremnicaID -> stornirana/nepostojeca otpremnica)
    Dim lostBlk As Variant
    lostBlk = GetLostOtkupBlokovi()
    If IsArray(lostBlk) Then
        warnings = warnings & UBound(lostBlk, 1) & " blok(ova) vezano za nepostojecu otpremnicu:" & vbCrLf
        Dim lb As Long
        For lb = 1 To IIf(UBound(lostBlk, 1) > 5, 5, UBound(lostBlk, 1))
            Dim lbKg As String: lbKg = ""
            If IsNumeric(lostBlk(lb, 5)) Then lbKg = Format$(CDbl(lostBlk(lb, 5)), "#,##0") & "kg"
            warnings = warnings & "  blok " & CStr(lostBlk(lb, 2)) & " " & lbKg & _
                       " -> otp " & CStr(lostBlk(lb, 7)) & vbCrLf
        Next lb
        If UBound(lostBlk, 1) > 5 Then warnings = warnings & "  ..." & vbCrLf
    End If

    CheckVerwaisteDokumente = warnings
End Function

