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
' ZBR-NORM-02: JEDINSTVENA SEMANTIKA poredjenja poslovnih brojeva.
'
' Nije jedini komparator u projektu i ne treba da bude: ZbirnaPostoji,
' BrojZbirneExists, ZbirnaIdentResolve i AktivniBrojeviZbirne imaju svoja
' inline poredjenja sa ISTOM semantikom (Trim + vbTextCompare). Prebacivati i
' njih samo radi jednog izvora znacilo bi dirati stabilan kod bez poslovne
' koristi. Ova funkcija postoji za ODLUCIVACE koji su radili DRUGACIJE.
'
' Poslovni broj (BrojZbirne, BrojPrijemnice, BrDok...) je LABELA koju kuca
' covek ili generise generator, pa se ista vrednost javlja u vise oblika:
' vodeci/prateci razmak iz Excel celije, i razlicit case kod rucnog unosa
' (format zbirne dozvoljava prefiks "S" -- v. IsValidBrojZbirneFormat).
'
' Do v6-ui-226 je isti kljuc imao TRI normalizacije: pun (Trim + vbTextCompare)
' u ZbirnaPostoji/BrojZbirneExists/ZbirnaIdentResolve, samo Trim u
' VlasniciPoBroju/LookupActiveID, i sirov u CheckDuplicate. Kapija i akter su
' tako mogli da racunaju razlicito nad istim brojem.
'
' Pravilo: KAPIJA sme da bude SIRA od aktera, nikad uza. Zato ovo koriste
' odlucivaci; mutatori (DetachOtpremniceInline, RedJeIzabranogDokumenta)
' namerno ostaju uzi -- kapija >= akter je time
' garantovano, a prosirivanje aktera bi diralo redove koje danas ne dira.
' CheckDuplicate ostaje sirov, imenovano u ugovoru par.3. Od 14.09.2026 ga zove
' jos samo prijemnica (do svog PR-a); OTK, OTP, ZBR i REV drzi
' modBrojevi.BrojZauzetUNizu.
Public Function BrojJednak(ByVal a As Variant, ByVal b As Variant) As Boolean
    BrojJednak = (StrComp(Trim$(NzToText(a)), Trim$(NzToText(b)), vbTextCompare) = 0)
End Function

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

' KLJUC STAVKE ZBIRNE za racun manjka: dokument + klasa (S9-5, S95-ODL-03).
' Klasa I i II istog dokumenta imaju ZASEBNU otpremnicu, zbirnu i prijemnicu u
' hladnjackom lancu, pa je (ZbirnaID + Klasa) prava granularnost: bez klase se
' prijem obe klase sabere pa dodeli SVAKOJ klasi ponaosob.
Public Function ManjakKljuc(ByVal zbirnaID As String, ByVal klasa As String) As String
    ManjakKljuc = UCase$(Trim$(zbirnaID)) & "|" & KlasaOrDefault(klasa)
End Function

Public Function BuildManjakDict() As Object
    ' Agregacija zbirna/prijemnica za racun manjka -- po DOKUMENTU (ZbirnaID).
    '
    ' S9-5 (S95-ODL-03, KI-097): kljuc je ZbirnaID, ne vlasnik. Do S9-5 je kljuc
    ' bio vlasnik (broj + vozac + kupac + klasa, RF-06/AUD-023), pa su se dve
    ' AKTIVNE zbirne istog broja i vlasnika (PWA KR-001: dva uredjaja) sabirale u
    ' jedan red, a ROBA (OM) je prijem prve delila i na otpremnicu druge. Od
    ' S9-3 se prijemnica pripisuje svojoj zbirnoj po ivici E4, a od S9-2
    ' otpremnica svojoj po clanstvu E3 -- oba kraja znaju ZbirnaID, pa vlasnik
    ' vise ne mora da se razresava po broju (#V/#1/#O/#K/#C kljucevi su obrisani).
    '
    ' Kolicina i klasa se citaju iz STAVKI oba dokumenta (S6 korak 3). Prijemnica
    ' bez aktivne zbirne (siroce) ne ulazi -- nju imenuje integritet E4.
    '
    ' Returns: Dictionary
    '   ManjakKljuc(ZbirnaID, klasa) -> Array(ZbirnaKg, PrijemnicaKg, CntPrijem)
    '   "#Z|" & UCase(ZbirnaID)      -> Array(ZbirnaKg, PrijemnicaKg, CntPrijem,
    '                                   ZbirnaAmb) -- ceo dokument, preko klasa
    '                                   koje zbirna nosi
    '   "#N|" & UCase(ZbirnaID)      -> Long: aktivne prijemnice te zbirne ciji se
    '                                   skup ne zna (kvar E4) -> fail-closed signal
    ' Prefiks "#" se ne moze pojaviti u ZbirnaID-u, pa nema kolizije.
    Const SRC As String = "modHelpers.BuildManjakDict"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set BuildManjakDict = dict

    ' --- Zbirna ---
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If Not IsArray(zbrData) Then Exit Function
    zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If Not IsArray(zbrData) Then Exit Function

    Dim colZbrID As Long
    colZbrID = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, SRC)

    Dim stavkeZbr As Object
    Set stavkeZbr = modDokumenta.StavkeZbirnePoDokumentu()

    Dim z As Long, zid As String, sk As String, vals As Variant, zs As Variant
    For z = 1 To UBound(zbrData, 1)
        zid = UCase$(Trim$(NzToText(zbrData(z, colZbrID))))
        If Not dict.Exists("#Z|" & zid) Then dict.Add "#Z|" & zid, Array(0#, 0#, 0&, 0&)
        For Each zs In modDokumenta.StavkeZaZbirnu(stavkeZbr, _
                           CStr(zbrData(z, colZbrID)), SRC)
            sk = ManjakKljuc(zid, CStr(zs(3)))
            If Not dict.Exists(sk) Then dict.Add sk, Array(0#, 0#, 0&)
            vals = dict(sk)
            vals(0) = vals(0) + CDbl(zs(4))
            dict(sk) = vals
            vals = dict("#Z|" & zid)
            vals(0) = vals(0) + CDbl(zs(4))
            vals(3) = CLng(vals(3)) + CLng(zs(5))
            dict("#Z|" & zid) = vals
        Next zs
    Next z

    ' --- Prijemnica ---
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then Exit Function
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If Not IsArray(prijData) Then Exit Function

    Dim colPID As Long
    colPID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)

    Dim stavkePrj As Object
    Set stavkePrj = modDokumenta.StavkePrijemnicePoDokumentu()

    ' Ivica E4 u smeru prijemnica -> zbirna, i kvar po zbirni.
    Dim rodPrj As Object, e4Kvar As Object, kz As Variant
    Set rodPrj = modAutoHladnjaca.ZbirneAktivnihPrijemnica(e4Kvar)
    For Each kz In e4Kvar.keys
        zid = UCase$(Trim$(CStr(kz)))
        If dict.Exists("#Z|" & zid) Then
            If Not dict.Exists("#N|" & zid) Then dict.Add "#N|" & zid, 0&
            dict("#N|" & zid) = CLng(dict("#N|" & zid)) + 1
        End If
    Next kz

    Dim p As Long, pid As String, pRod As String, ps As Variant
    Dim pKg As Double, psk As String
    For p = 1 To UBound(prijData, 1)
        pid = UCase$(Trim$(NzToText(prijData(p, colPID))))
        pRod = ""
        If rodPrj.Exists(pid) Then pRod = UCase$(Trim$(CStr(rodPrj(pid))))
        If dict.Exists("#Z|" & pRod) Then
            For Each ps In modDokumenta.StavkeZaPrijemnicu(stavkePrj, _
                               CStr(prijData(p, colPID)), SRC)
                pKg = CDbl(ps(4))
                psk = ManjakKljuc(pRod, CStr(ps(3)))
                ' Prijem se pripisuje samo klasi koju zbirna nosi -- isto pravilo
                ' koje je izvestaj imao i pre S9-5 (red sabira klase zbirne).
                If dict.Exists(psk) Then
                    vals = dict(psk)
                    vals(1) = vals(1) + pKg
                    vals(2) = CLng(vals(2)) + 1
                    dict(psk) = vals
                    vals = dict("#Z|" & pRod)
                    vals(1) = vals(1) + pKg
                    vals(2) = CLng(vals(2)) + 1
                    dict("#Z|" & pRod) = vals
                End If
            Next ps
        End If
    Next p
End Function


