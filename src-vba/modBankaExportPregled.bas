Attribute VB_Name = "modBankaExportPregled"
Option Explicit

' ============================================================
' modIsplatePregled v6.18+
'
' Read-only helper sloj za pregled otvorenih otkup blokova
' (per-kooperant aggregation + already-paid awareness + TR check)
' + izlazi za pripremu isplata:
'   - GenerisiNalogeCSV: CSV naloga za prenos (uvoz u e-banking);
'     iznos po bloku = clsBlokIsplata.IsplatitiIznos (operater unos).
'   - PrintIsplataSpecifikacija: specifikacija isplata (PDF/stampa,
'     ISPLATA_SPEC_PRINT_MODE), isti blokovi i iznosi kao CSV.
'
' Bazira se na postojecim helperima:
'   - GetOpenOtkupi (extended 7-col shape)
'   - GetKooperantNaziv (modBankaMapiranje)
'   - LookupValue za TekuciRacun
'
' NE pise nista u tblNovac: isplata se knjizi tek kroz uvoz izvoda
' (modBankaImport/modBankaMapiranje, auto-map poziv na broj = broj bloka).
' ============================================================

' Koliko problematicnih blokova se nabroji u poruci odbijanja (ostatak se
' sazme u "..."), da poruka ostane citljiva i kad je lista velika.
Private Const MAX_SALDO_PROBLEMA As Long = 10

' Integritet kanonskog kljuca otvorenih blokova (vidi BuildOpenAmountDict).
Public Const ERR_ISPLATA_DUPLI_OTKUPID As Long = vbObjectError + 2620
Public Const ERR_ISPLATA_PRAZAN_OTKUPID As Long = vbObjectError + 2621

' ------------------------------------------------------------------
' PRAVILO ZA POREDJENJE NOVCA NA OVOJ PUTANJI (AUD-026)
'
' Iznosi se NE porede sa epsilon tolerancijom nego u cent-domenu: oba se
' zaokruze postojecim modNovac.ZaokruziNovac (half-up) pa se porede TACNO.
' Prag tipa "> otvoreno + 0.01" je propustao preplatu od PUNOG centa
' (600.01 nije vece od 600.00 + 0.01), a to je iznos koji banka stvarno
' naplati -- CSV izvozi dve decimale, pa sub-cent razlike u fajlu ionako
' ne postoje.
'
' Isto pravilo vazi za klamp override-a, finalnu revalidaciju, CSV writer
' (CsvIznos normalizuje PRE Format$) i operaterski unos u formi
' (frmBankaExportPregled.txtIsplatiti_Exit). NE uvoditi novu granicu ni
' novi helper -- ZaokruziNovac je jedina normalizacija.
' ------------------------------------------------------------------

'======================================================================
' BuildBlokIsplataList
'
' Vraca Collection of BlokIsplata, sve sistemski-otvoreni blokovi
' enrich-ovani sa kooperant TR + KooperantNaziv.
'
' Filteri:
'   - datumOd / datumDo: opciono date range (#1/1/1900# = no filter)
'   - stanicaIDFilter: opciono stanica filter ("" = sve stanice)
'
' Performance: jedan poziv GetOpenOtkupi() (single-pass),
' jedan poziv BuildIsplataDictByOtkup() (sadrzi se u GetOpenOtkupi),
' N lookup-a za TR i KooperantNaziv (cached preko Dictionary-ja).
'======================================================================
Public Function BuildBlokIsplataList( _
    Optional ByVal datumOd As Date, _
    Optional ByVal datumDo As Date, _
    Optional ByVal stanicaIDFilter As String = "" _
) As Collection
    
    Dim result As New Collection
    Dim openOtkupi As Variant
    
    openOtkupi = GetOpenOtkupi("")
    
    If IsEmpty(openOtkupi) Then
        Set BuildBlokIsplataList = result
        Exit Function
    End If
    
    Dim trCache As Object
    Dim nazivCache As Object
    Dim avansCache As Object
    Set trCache = BuildKooperantTekuciRacunCache()
    Set nazivCache = CreateObject("Scripting.Dictionary")
    Set avansCache = BuildKooperantUnallocatedAvansDict()
    ' KooperantID po OtkupID: O(n) mapa JEDNOM (umesto LookupValue O(n) po redu
    ' u petlji, sto je bilo O(n^2) i glavni uzrok laga pri ucitavanju).
    ' AUD-026: NE preko BuildLookupDict -- vidi BuildOtkupOwnerIndex.
    Dim koopByOtkup As Object
    Dim dupliOtkup As Object
    Set koopByOtkup = BuildOtkupOwnerIndex(dupliOtkup)
    
    Dim i As Long
    For i = 1 To UBound(openOtkupi, 1)
        Dim brojDok As String
        Dim otkupID As String
        Dim ukupan As Double
        Dim isplaceno As Double
        Dim otvoren As Double
        Dim datumVal As Date
        Dim stanicaID As String
        
        brojDok = CStr(openOtkupi(i, 1))
        otkupID = Trim$(CStr(openOtkupi(i, 2)))
        ukupan = CDbl(openOtkupi(i, 3))
        isplaceno = CDbl(openOtkupi(i, 4))
        otvoren = CDbl(openOtkupi(i, 5))

        ' AUD-026: vlasnik OTVORENOG bloka mora biti jednoznacan. Provera ide
        ' pre filtera (i pre citanja vlasnika), da se ekran ponasa isto kao
        ' finalna kapija, koja listu gradi bez filtera.
        If dupliOtkup.Exists(otkupID) Then
            Err.Raise ERR_ISPLATA_DUPLI_OTKUPID, "modBankaExportPregled.BuildBlokIsplataList", _
                      "Dupli OtkupID u tblOtkup: " & otkupID & " (otvoren blok: " & brojDok & "). " & _
                      "Vlasnik i tekuci racun se ne mogu pouzdano utvrditi -- pokrenite proveru " & _
                      "integriteta podataka."
        End If

        If IsDate(openOtkupi(i, 6)) Then
            datumVal = CDate(openOtkupi(i, 6))
        Else
            On Error Resume Next
            datumVal = CDate(CStr(openOtkupi(i, 6)))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0
        End If
        
        stanicaID = CStr(openOtkupi(i, 7))
        
        If stanicaIDFilter <> "" And stanicaID <> stanicaIDFilter Then GoTo NextRow
        If datumOd > #1/1/1900# And datumVal < datumOd Then GoTo NextRow
        If datumDo > #1/1/1900# And datumVal > datumDo Then GoTo NextRow
        
        Dim kooperantID As String
        If koopByOtkup.Exists(otkupID) Then kooperantID = Trim$(CStr(koopByOtkup(otkupID)))
        If LenB(kooperantID) = 0 Then GoTo NextRow
        
        Dim kooperantNaziv As String
        If nazivCache.Exists(kooperantID) Then
            kooperantNaziv = nazivCache(kooperantID)
        Else
            kooperantNaziv = GetKooperantNaziv(kooperantID)
            nazivCache.Add kooperantID, kooperantNaziv
        End If
        
        Dim tr As String
        If trCache.Exists(kooperantID) Then
            tr = trCache(kooperantID)
        Else
            tr = ""
        End If
        
        ' --- promenjeno: class instance umesto UDT ---
        Dim blk As clsBlokIsplata
        Set blk = New clsBlokIsplata
        blk.otkupID = otkupID
        blk.datum = datumVal
        blk.kooperantID = kooperantID
        blk.kooperantNaziv = kooperantNaziv
        blk.stanicaID = stanicaID
        blk.TekuciRacun = tr
        blk.brojDokumenta = brojDok
        blk.UkupanIznos = ukupan
        blk.VecIsplaceno = isplaceno
        blk.OtvorenIznos = otvoren
        blk.HasTekuciRacun = (LenB(Trim$(tr)) > 0)
        If avansCache.Exists(kooperantID) Then
            blk.KooperantAvansSaldo = CDbl(avansCache(kooperantID))
        Else
            blk.KooperantAvansSaldo = 0
        End If
        
        result.Add blk
        
NextRow:
    Next i
    
    Set BuildBlokIsplataList = result
End Function
'======================================================================
' BuildOtkupOwnerIndex - OtkupID -> KooperantID iz SIROVE tblOtkup, uz
' evidenciju dupliranih ID-eva (AUD-026).
'
' Zasto ne BuildLookupDict: on je "prvi pojav pobedjuje" nad SIROVOM
' tabelom, a GetOpenOtkupi preskace stornirane i vec isplacene redove. Kod
' dupliranog OtkupID-a te dve semantike se razilaze:
'
'   red A: OTK-777, kooperant K-A, blok BLOK-A, Isplaceno = "Da"
'   red B: OTK-777, kooperant K-B, blok BLOK-B, otvoren
'
' GetOpenOtkupi vrati SAMO B (A je isplacen), pa u otvorenoj listi postoji
' jedan jedini OTK-777 i stroga saldo-mapa nema sta da prijavi. Ali
' first-wins vlasnik dolazi sa reda A -- blok B bi dobio naziv i TEKUCI
' RACUN kooperanta A, i nalog bi otisao POGRESNOM PRIMAOCU. To je gore od
' pogresnog prikaza, pa se identitet primaoca ne sme resavati "prvim
' pogotkom".
'
' Zato: vlasnik se pamti samo za JEDNOZNACNE ID-eve, a dupli se vracaju
' kroz outDupli. Pozivalac pada tek ako je bas taj ID medju OTVORENIMA --
' istorijski duplikat koji ne proizvodi nalog ne obara ceo ekran.
'======================================================================
Private Function BuildOtkupOwnerIndex(ByRef outDupli As Object) As Object
    Dim ownerDict As Object
    Dim dupliDict As Object
    Set ownerDict = CreateObject("Scripting.Dictionary")
    Set dupliDict = CreateObject("Scripting.Dictionary")

    Set outDupli = dupliDict
    Set BuildOtkupOwnerIndex = ownerDict

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Const SRC As String = "BuildOtkupOwnerIndex"
    Dim colID As Long, colKoop As Long
    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)

    Dim i As Long
    Dim oid As String
    For i = 1 To UBound(data, 1)
        oid = Trim$(CStr(data(i, colID)))
        If LenB(oid) > 0 Then
            If ownerDict.Exists(oid) Then
                dupliDict(oid) = True
            Else
                ownerDict.Add oid, Trim$(CStr(data(i, colKoop)))
            End If
        End If
    Next i
End Function

'======================================================================
' BuildKooperantTekuciRacunCache
' Single-pass cache KooperantID -> TekuciRacun
'======================================================================
Private Function BuildKooperantTekuciRacunCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        Set BuildKooperantTekuciRacunCache = dict
        Exit Function
    End If
    
    Const SRC As String = "BuildKooperantTekuciRacunCache"
    Dim colID As Long, colTR As Long
    colID = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID, SRC)
    colTR = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_TEKUCI_RACUN, SRC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim kID As String
        kID = Trim$(CStr(data(i, colID)))
        If LenB(kID) > 0 Then
            If Not dict.Exists(kID) Then
                dict.Add kID, Trim$(CStr(data(i, colTR)))
            End If
        End If
    Next i
    
    Set BuildKooperantTekuciRacunCache = dict
End Function

'======================================================================
' SummarizeBlokList - status bar string
'======================================================================
Public Function SummarizeBlokList(ByVal blokovi As Collection) As String
    If blokovi Is Nothing Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    If blokovi.count = 0 Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    
    Dim totalOpen As Double
    Dim missingTR As Long
    Dim kooperantSet As Object
    Set kooperantSet = CreateObject("Scripting.Dictionary")
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v               ' Set jer je v object reference
        totalOpen = totalOpen + blk.OtvorenIznos
        If Not blk.HasTekuciRacun Then missingTR = missingTR + 1
        If Not kooperantSet.Exists(blk.kooperantID) Then
            kooperantSet.Add blk.kooperantID, True
        End If
    Next v
    
    Dim baseMsg As String
    baseMsg = blokovi.count & " blokova | " & _
              kooperantSet.count & " kooperanata | " & _
              "Otvoreno: " & Format$(totalOpen, "#,##0.00") & " RSD"
    
    If missingTR > 0 Then
        baseMsg = baseMsg & " | Bez TR: " & missingTR
    End If
    
    SummarizeBlokList = baseMsg
End Function

'======================================================================
' PrintIsplataSpecifikacija - specifikacija isplata po blokovima (PDF/
' stampa po ISPLATA_SPEC_PRINT_MODE, default PDF -> otvori se).
' Ocekuje kolekciju blokova SA postavljenim IsplatitiIznos
' (CollectIsplataBlokovi u frmBankaExportPregled). Render u modPrint
' (EnsureIsplataSpecSablon/FillIsplataSpecSablon, house-style).
'======================================================================
Public Sub PrintIsplataSpecifikacija(ByVal blokovi As Collection, _
                                     Optional ByVal platilacRacun As String = "")
    On Error GoTo EH
    If blokovi Is Nothing Then Exit Sub
    If blokovi.count = 0 Then Exit Sub

    platilacRacun = NormalizujRacun(platilacRacun)
    If LenB(platilacRacun) = 0 Then platilacRacun = NormalizujRacun(DocConfigOr("SELLER_ACCOUNT", ""))

    Dim n As Long: n = blokovi.count
    Dim spec() As Variant: ReDim spec(1 To n, 1 To 9)
    Dim sumUkupan As Double, sumIsplaceno As Double
    Dim sumOtvoreno As Double, sumIsplatiti As Double

    Dim i As Long
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        i = i + 1
        spec(i, 1) = i
        spec(i, 2) = Format$(blk.datum, "d.m.yyyy")
        spec(i, 3) = blk.brojDokumenta
        spec(i, 4) = blk.kooperantNaziv
        spec(i, 5) = blk.TekuciRacun
        spec(i, 6) = blk.UkupanIznos
        spec(i, 7) = blk.VecIsplaceno
        spec(i, 8) = blk.OtvorenIznos
        spec(i, 9) = blk.IsplatitiIznos
        sumUkupan = sumUkupan + blk.UkupanIznos
        sumIsplaceno = sumIsplaceno + blk.VecIsplaceno
        sumOtvoreno = sumOtvoreno + blk.OtvorenIznos
        sumIsplatiti = sumIsplatiti + blk.IsplatitiIznos
    Next v

    Dim subtitle As String
    subtitle = "Datum: " & Format$(Date, "d.m.yyyy") & _
               "     Platilac: " & DocConfigOr("SELLER_NAME", "") & _
               " (" & platilacRacun & ")"

    Dim ws As Worksheet
    Set ws = FillIsplataSpecSablon(spec, n, subtitle, _
                                   sumUkupan, sumIsplaceno, sumOtvoreno, sumIsplatiti)
    If ws Is Nothing Then Exit Sub

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_ISPLATA_SPEC_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            Dim pdfPath As String
            pdfPath = EnsureDocFolder(PDF_DIR_SPECIFIKACIJE) & "\Specifikacija_isplata_" & _
                      IsplataFileTag(platilacRacun) & "_" & Format$(Now, "hhnnss") & ".pdf"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    Dim errDesc As String
    errDesc = Err.description
    LogErr "modBankaExportPregled.PrintIsplataSpecifikacija"
    MsgBox "Gre" & ChrW(353) & "ka pri izradi specifikacije: " & errDesc, vbCritical, APP_NAME
End Sub

'======================================================================
' BankaNalogRacuniCSV - efektivni ";"-spisak racuna firme za isplate.
'
' Izvor istine: tri zasebna polja iz Podesavanja (grupa "Banka / nalozi")
' CFG_BANKA_NALOG_RACUN_1/2/3 -- spajaju se (preskacuci prazna) u ";"-listu.
' Fallback (stare instalacije): ako su sva tri prazna -> legacy jedinstveni
' ";"-spisak CFG_BANKA_NALOG_RACUNI; ako i to prazno -> SELLER_ACCOUNT (firma
' sa jednim racunom). Vraca "" ako nista nije uneto (forma tada javlja poruku).
'
' Koristi frmBankaExportPregled.PopulateRacunCombo (combo "Sa racuna").
'======================================================================
Public Function BankaNalogRacuniCSV() As String
    Dim res As String, v As String, i As Long
    Dim keys As Variant
    keys = Array(CFG_BANKA_NALOG_RACUN_1, CFG_BANKA_NALOG_RACUN_2, _
                 CFG_BANKA_NALOG_RACUN_3, CFG_BANKA_NALOG_RACUN_4)
    For i = LBound(keys) To UBound(keys)
        v = Trim$(GetConfigValue(CStr(keys(i))))
        If LenB(v) > 0 Then
            If LenB(res) > 0 Then res = res & ";"
            res = res & v
        End If
    Next i
    If LenB(res) = 0 Then res = Trim$(GetConfigValue(CFG_BANKA_NALOG_RACUNI))
    If LenB(res) = 0 Then res = Trim$(GetConfigValue("SELLER_ACCOUNT"))
    BankaNalogRacuniCSV = res
End Function

'======================================================================
' BuildOpenAmountDict - OtkupID -> OtvorenIznos iz kolekcije blokova.
'
' Koriste je i clamp override-a i finalna revalidacija naloga, pa obe
' gledaju iznos iz ISTOG izvora (clsBlokIsplata.OtvorenIznos, koji dolazi
' iz GetOpenOtkupi -> vrednost minus knjizene isplate). Nema pozicijskog
' citanja tabela ovde: saldo se ne racuna ponovo, samo se indeksira.
'
' STRIKTNO na kanonskom kljucu (fail-closed): dupli ili prazan OtkupID
' PODIZE gresku umesto da jedan red pobedi drugi. GetOpenOtkupi ne agregira
' po OtkupID nego vraca red po red, pa bi kod korumpiranog PK-a (dva reda,
' isti OtkupID, RAZLICIT BrojDokumenta i otvoreni iznos) poslednji upis
' odlucio saldo -- i saldo bloka B bi odobrio nalog za blok A. Ranija
' verzija je ovde svesno pisala assignment-om "da korumpirani podatak ne
' obori pregled"; to je na finansijskoj kapiji fail-OPEN i zato je vraceno
' na tvrd pad.
'
' Posledica je namerna i na obe strane:
'   - pregled (ClampOverridesToOpen) prijavi gresku umesto da "nekako"
'     izabere koji duplikat vazi;
'   - GenerisiNalogeCSV pukne u svom EH pre nego sto payload uopste nastane,
'     pa nema ni fajla za banku.
' Duplikat OtkupID-a inace hvata i RunProductionHealthCheck, ali health-check
' nije preduslov svakog klika na "Generisi naloge".
'======================================================================
Public Function BuildOpenAmountDict(ByVal blokovi As Collection) As Object
    Const SRC As String = "modBankaExportPregled.BuildOpenAmountDict"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not blokovi Is Nothing Then
        Dim blk As clsBlokIsplata
        Dim v As Variant
        Dim oid As String
        For Each v In blokovi
            Set blk = v
            oid = Trim$(blk.otkupID)

            If LenB(oid) = 0 Then
                Err.Raise ERR_ISPLATA_PRAZAN_OTKUPID, SRC, _
                          "Prazan OtkupID u otvorenim blokovima (dokument: " & _
                          blk.brojDokumenta & "). Saldo se ne moze pouzdano utvrditi."
            End If

            If dict.Exists(oid) Then
                Err.Raise ERR_ISPLATA_DUPLI_OTKUPID, SRC, _
                          "Dupli OtkupID u otvorenim blokovima: " & oid & _
                          " (dokument: " & blk.brojDokumenta & "). Saldo se ne moze " & _
                          "pouzdano utvrditi -- pokrenite proveru integriteta podataka."
            End If

            dict.Add oid, blk.OtvorenIznos
        Next v
    End If

    Set BuildOpenAmountDict = dict
End Function

'======================================================================
' ClampOverridesToOpen - AUD-026 / FM-0020 #1.
'
' Per-blok "Isplatiti" override zivi u formi (Dictionary OtkupID -> iznos)
' i PREZIVLJAVA reload liste. Ako se u medjuvremenu deo bloka isplati,
' veze avans ili se otkup stornira, zaostao override moze biti VECI od
' onoga sto je jos otvoreno -> nalog bi narucio preplatu. Zato se pri
' svakom rebuild-u liste override uskladi sa trenutnim stanjem:
'   - blok kog vise nema u otvorenima (zatvoren/storniran) -> override se BRISE
'   - blok bez otvorenog iznosa                            -> override se BRISE
'   - override veci od otvorenog                           -> SPUSTA se na otvoreno
'
' Override MANJI od otvorenog se ne dira: to je namerna operater odluka
' (delimicna isplata) i ona i dalje vazi.
'
' Menja prosledjeni Dictionary na mestu. Vraca broj promenjenih unosa
' (clamp-ovanih + obrisanih), da forma moze da javi operateru; 0 = nista.
'======================================================================
Public Function ClampOverridesToOpen(ByVal overrideDict As Object, _
                                     ByVal blokovi As Collection) As Long
    ClampOverridesToOpen = 0

    ' Mapa se gradi PRE svih ranih izlaza: ona je i provera integriteta
    ' kanonskog kljuca (BuildOpenAmountDict pada na dupli/prazan OtkupID).
    ' Ako bi se preskocila kad override-a nema, prvi otvor forme -- gde je
    ' dictionary uvek prazan -- pustio bi korumpiran OtkupID na ekran, a onda
    ' i u akcije koje sa tog ekrana mutiraju podatke.
    Dim openByOtkup As Object
    If Not blokovi Is Nothing Then Set openByOtkup = BuildOpenAmountDict(blokovi)

    If overrideDict Is Nothing Then Exit Function

    ' Nema liste blokova = nema dokaza da je ijedan override jos vazeci.
    If blokovi Is Nothing Then
        ClampOverridesToOpen = overrideDict.count
        overrideDict.RemoveAll
        Exit Function
    End If

    If overrideDict.count = 0 Then Exit Function

    Dim toRemove As Collection
    Set toRemove = New Collection

    Dim changed As Long
    Dim k As Variant
    Dim otkupKey As String
    Dim otvoreno As Double
    Dim trenutni As Double

    For Each k In overrideDict.keys
        otkupKey = CStr(k)
        If Not openByOtkup.Exists(otkupKey) Then
            toRemove.Add otkupKey
        Else
            otvoreno = ZaokruziNovac(CDbl(openByOtkup(otkupKey)))
            If otvoreno <= 0 Then
                toRemove.Add otkupKey
            Else
                trenutni = ZaokruziNovac(CDbl(overrideDict(otkupKey)))
                If trenutni > otvoreno Then
                    overrideDict(otkupKey) = otvoreno
                    changed = changed + 1
                End If
            End If
        End If
    Next k

    Dim s As Variant
    For Each s In toRemove
        overrideDict.Remove CStr(s)
        changed = changed + 1
    Next s

    ClampOverridesToOpen = changed
End Function

'======================================================================
' ValidateNalogSaldo - AUD-026 / FM-0020 #2 + FM-0021 #1.
'
' Finalna kapija pred upis CSV-a: nijedan nalog ne sme da naruci vise nego
' sto je NA TAJ TRENUTAK otvoreno. Clamp pri reload-u nije dovoljan, jer se
' podaci mogu promeniti izmedju prikaza i klika na "Generisi" (uvoz izvoda,
' vezivanje avansa, storno u drugom delu aplikacije).
'
' Proverava tacno one blokove koje bi GenerisiNalogeCSV i upisao
' (HasTekuciRacun + ZAOKRUZEN iznos > 0, isti prag koji koristi writer) --
' inace bi blok koji uopste ne ide u fajl mogao da obori generisanje.
'
' Blok kog vise NEMA u openByOtkup se tretira kao otvoreno = 0, dakle svaki
' iznos na njemu je preplata (fail-closed: nestao je jer je zatvoren/storniran).
'
' Vraca "" kada je sve u redu, inace poruku za operatera (koji blok, koliko
' trazi vs koliko je otvoreno).
'======================================================================
Public Function ValidateNalogSaldo(ByVal blokovi As Collection, _
                                   ByVal openByOtkup As Object) As String
    ValidateNalogSaldo = ""
    If blokovi Is Nothing Then Exit Function
    If blokovi.count = 0 Then Exit Function

    If openByOtkup Is Nothing Then
        ValidateNalogSaldo = "Nalozi nisu generisani: nije mogu" & ChrW(263) & "e proveriti " & _
                             "otvorene iznose blokova."
        Exit Function
    End If

    Dim problemi As String
    Dim brojProblema As Long
    Dim blk As clsBlokIsplata
    Dim v As Variant
    Dim otvoreno As Double
    Dim trazeno As Double

    For Each v In blokovi
        Set blk = v
        ' Poredi se ono sto CSV stvarno nosi: iznos u cent-domenu, bez
        ' tolerancije (vidi pravilo na vrhu modula). Normalizacija ide i PRE
        ' praga "> 0" -- sirov ostatak (0.004) inace prodje kao nalog, a u
        ' fajlu zavrsi kao "0.00".
        trazeno = ZaokruziNovac(blk.IsplatitiIznos)
        If blk.HasTekuciRacun And trazeno > 0 Then
            otvoreno = 0
            If openByOtkup.Exists(blk.otkupID) Then otvoreno = ZaokruziNovac(CDbl(openByOtkup(blk.otkupID)))

            If trazeno > otvoreno Then
                brojProblema = brojProblema + 1
                If brojProblema <= MAX_SALDO_PROBLEMA Then
                    problemi = problemi & vbCrLf & "  " & blk.brojDokumenta & _
                               " (" & blk.kooperantNaziv & "): tra" & ChrW(382) & "i " & _
                               Format$(trazeno, "#,##0.00") & _
                               " / otvoreno " & Format$(otvoreno, "#,##0.00")
                End If
            End If
        End If
    Next v

    If brojProblema = 0 Then Exit Function

    If brojProblema > MAX_SALDO_PROBLEMA Then
        problemi = problemi & vbCrLf & "  ... i jo" & ChrW(353) & " " & _
                   CStr(brojProblema - MAX_SALDO_PROBLEMA) & " blok(ova)"
    End If

    ValidateNalogSaldo = "Nalozi NISU generisani: " & CStr(brojProblema) & _
                         " blok(ova) tra" & ChrW(382) & "i vi" & ChrW(353) & "e nego " & _
                         ChrW(353) & "to je otvoreno." & vbCrLf & problemi & vbCrLf & vbCrLf & _
                         "Otvoreni iznosi su se promenili u me" & ChrW(273) & "uvremenu. " & _
                         "Osve" & ChrW(382) & "ite pregled i ponovite."
End Function

'======================================================================
' GenerisiNalogeCSV - CSV naloga za prenos za uvoz u e-banking.
'
' Jedan red = jedan nalog za prenos po otkupnom bloku:
'   - platilac: SELLER_NAME / SELLER_ACCOUNT (config, grupa "Prodavac (firma)")
'   - primalac: kooperant (naziv + TekuciRacun iz tblKooperanti)
'   - iznos:    blk.IsplatitiIznos (postavlja forma: operater unos ili otvoreno)
'   - poziv na broj (odobrenje) = broj otkupnog bloka -> jaki kljuc za
'     auto-map pri kasnijem uvozu izvoda (frmBankaImport)
'
' Ocekuje blokove SA tekucim racunom (filtrira ih forma). Blokove bez TR
' preskace i tiho broji (defenzivno). Vraca punu putanju CSV fajla,
' "" ako nema nijednog reda ili upis ne uspe.
'
' AUD-026: pre upisa se saldo SVAKOG naloga revalidira protiv SVEZE
' ocitanih otvorenih iznosa (ValidateNalogSaldo). Ako ijedan nalog trazi
' vise nego sto je otvoreno, fajl se NE pise -- razlog se vraca kroz
' outOdbijeno (forma ga pokazuje operateru). Fail-closed: i greska pri
' ocitavanju zavrsi u EH, dakle bez fajla.
'
' Format: ";" separator, UTF-8 (BOM - vidi WriteAllTextUtf8), decimalna
' tacka u iznosu (deterministicki, nezavisno od Windows locale-a).
' Kolone drzati stabilnim: e-banking uvozi se mapiraju po pozicijama.
'======================================================================
Public Function GenerisiNalogeCSV(ByVal blokovi As Collection, _
                                  Optional ByVal platilacRacun As String = "", _
                                  Optional ByRef outOdbijeno As String) As String
    On Error GoTo EH

    GenerisiNalogeCSV = ""
    outOdbijeno = ""
    If blokovi Is Nothing Then Exit Function
    If blokovi.count = 0 Then Exit Function

    ' Platilac racun: prosledjen iz forme (combo "Sa racuna"); prazno ->
    ' SELLER_ACCOUNT (backward-compatible kada racuna ima samo jedan).
    platilacRacun = NormalizujRacun(platilacRacun)
    If LenB(platilacRacun) = 0 Then platilacRacun = NormalizujRacun(DocConfigOr("SELLER_ACCOUNT", ""))
    If LenB(platilacRacun) = 0 Then Exit Function   ' forma validira i javlja poruku

    ' AUD-026: sadrzaj fajla se gradi TEK kroz BuildNalogCsvPayload, koji u
    ' sebi nosi finalnu saldo kapiju -- nema putanje do WriteAllTextUtf8 koja
    ' bi zaobisla proveru. Svez saldo se cita ovde (a ne iz prikaza) bas zato
    ' sto se stanje moglo promeniti otkad je lista ucitana.
    Dim s As String
    s = BuildNalogCsvPayload(blokovi, platilacRacun, _
                             BuildOpenAmountDict(BuildBlokIsplataList()), outOdbijeno)
    If LenB(s) = 0 Then Exit Function

    Dim csvPath As String
    csvPath = EnsureDocFolder(CSV_DIR_BANKA_NALOZI) & "\Nalozi_za_prenos_" & _
              IsplataFileTag(platilacRacun) & "_" & Format$(Now, "hhnnss") & ".csv"
    WriteAllTextUtf8 csvPath, s

    GenerisiNalogeCSV = csvPath
    Exit Function

EH:
    ' Razlog se hvata PRE LogErr-a (BUG-1/AUD-054) i vraca kroz outOdbijeno:
    ' ovde zavrsava i tvrd pad integriteta kljuca (dupli/prazan OtkupID), pa
    ' bi generican tekst "pogledajte log" ostavio operatera bez ideje sta se
    ' desilo. Fajl u svakom slucaju NIJE napisan.
    Dim errDesc As String
    errDesc = Err.description

    LogErr "modBankaExportPregled.GenerisiNalogeCSV"
    outOdbijeno = "Nalozi NISU generisani (gre" & ChrW(353) & "ka): " & errDesc
    GenerisiNalogeCSV = ""
End Function

'======================================================================
' BuildNalogCsvPayload - AUD-026: sadrzaj CSV fajla (zaglavlje + redovi),
' sa finalnom saldo kapijom UNUTAR iste funkcije.
'
' Razlog za izdvajanje: kapija i pisanje fajla ne smeju da budu dva koraka
' koja neko kasnije moze da razdvoji. Ovde je gradnja sadrzaja fizicki iza
' provere -- ako ijedan nalog prelazi otvoreno, funkcija vraca "" i razlog
' kroz outOdbijeno, pa GenerisiNalogeCSV nema sta da upise. Uz to je cela
' putanja (validacija + payload) testabilna bez diranja diska.
'
' Vraca "" i kada nema nijednog naloga za upis (bez outOdbijeno -- to nije
' greska, nego prazan izbor).
'======================================================================
Public Function BuildNalogCsvPayload(ByVal blokovi As Collection, _
                                     ByVal platilacRacun As String, _
                                     ByVal openByOtkup As Object, _
                                     Optional ByRef outOdbijeno As String) As String
    BuildNalogCsvPayload = ""
    outOdbijeno = ""
    If blokovi Is Nothing Then Exit Function
    If blokovi.count = 0 Then Exit Function

    Dim odbijeno As String
    odbijeno = ValidateNalogSaldo(blokovi, openByOtkup)
    If LenB(odbijeno) > 0 Then
        outOdbijeno = odbijeno
        Exit Function
    End If

    Dim platilacNaziv As String
    platilacNaziv = DocConfigOr("SELLER_NAME", "")

    Dim sifra As String, svrhaBase As String
    sifra = DocConfigOr(CFG_BANKA_NALOG_SIFRA, BANKA_NALOG_SIFRA_DEFAULT)
    svrhaBase = DocConfigOr(CFG_BANKA_NALOG_SVRHA, BANKA_NALOG_SVRHA_DEFAULT)

    Dim datumValute As String
    datumValute = Format$(Date, "dd.mm.yyyy")

    Dim s As String
    s = "RacunPlatioca;NazivPlatioca;RacunPrimaoca;NazivPrimaoca;Iznos;Valuta;" & _
        "SifraPlacanja;Model;PozivNaBroj;SvrhaPlacanja;DatumValute" & vbCrLf

    Dim rows As Long
    Dim trazeno As Double
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        ' Isti normalizovan iznos koji je kapija odobrila -- i za prag "> 0"
        ' i za upis. Sirov prag bi propustio ostatak koji se u fajlu prikaze
        ' kao "0.00" (nalog na nula dinara).
        trazeno = ZaokruziNovac(blk.IsplatitiIznos)
        If blk.HasTekuciRacun And trazeno > 0 Then
            s = s & CsvField(NormalizujRacun(platilacRacun)) & ";" & _
                    CsvField(platilacNaziv) & ";" & _
                    CsvField(NormalizujRacun(blk.TekuciRacun)) & ";" & _
                    CsvField(blk.kooperantNaziv) & ";" & _
                    CsvIznos(trazeno) & ";" & _
                    "RSD;" & _
                    CsvField(sifra) & ";" & _
                    ";" & _
                    CsvField(blk.brojDokumenta) & ";" & _
                    CsvField(svrhaBase & " " & blk.brojDokumenta) & ";" & _
                    datumValute & vbCrLf
            rows = rows + 1
        End If
    Next v

    If rows = 0 Then Exit Function
    BuildNalogCsvPayload = s
End Function

' Iznos za CSV: uvek decimalna TACKA, bez hiljada separatora, 2 decimale.
' Format$ "0.00" na sr locale daje zarez -> normalizuj deterministicki.
' Iznos se PRE formatiranja normalizuje u cent-domen (ZaokruziNovac, half-up),
' da fajl nosi tacno onu vrednost koju je kapija odobrila -- Format$ sam
' zaokruzuje po svom pravilu, pa bi se inace mogli razlikovati.
Private Function CsvIznos(ByVal amt As Double) As String
    CsvIznos = Replace(Format$(ZaokruziNovac(amt), "0.00"), ",", ".")
End Function

' CSV polje: trim + quote ako sadrzi separator/navodnik/novi red.
Private Function CsvField(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If InStr(t, ";") > 0 Or InStr(t, """") > 0 Or _
       InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then
        t = """" & Replace(t, """", """""") & """"
    End If
    CsvField = t
End Function

' Tekuci racun za nalog: skini razmake (format "160-xxxx-xx" ili 18 cifara
' ostaje kako je unet u maticne podatke / config).
Private Function NormalizujRacun(ByVal racun As String) As String
    NormalizujRacun = Replace(Trim$(racun), " ", "")
End Function

' Deo naziva fajla za izlaze isplata: "<yyyy-mm-dd>_<Banka>" -- datum placanja
' (danas, = DatumValute u nalozima) + naziv banke platioca, radi lakseg
' snalazenja u folderu. Banka iz BankaNazivZaRacun (sanitizovana za ime fajla);
' nepoznat/prazan racun -> samo datum.
Private Function IsplataFileTag(ByVal platilacRacun As String) As String
    Dim tag As String
    tag = Format$(Date, "yyyy-mm-dd")
    Dim banka As String
    banka = SanitizeFileNamePart(BankaNazivZaRacun(platilacRacun))
    If LenB(banka) > 0 Then tag = tag & "_" & banka
    IsplataFileTag = tag
End Function

' Sanitizuj string za ime fajla: SR dijakritika -> ASCII, razmaci -> "-", ukloni
' znakove nedozvoljene u Windows imenima. Vraca ASCII-safe deo (npr. "Banca
' Intesa" -> "Banca-Intesa", "Postanska stedionica" ostaje bez dijakritike).
Private Function SanitizeFileNamePart(ByVal s As String) As String
    Dim t As String: t = Trim$(s)
    t = Replace(t, ChrW(352), "S"): t = Replace(t, ChrW(353), "s")
    t = Replace(t, ChrW(381), "Z"): t = Replace(t, ChrW(382), "z")
    t = Replace(t, ChrW(268), "C"): t = Replace(t, ChrW(269), "c")
    t = Replace(t, ChrW(262), "C"): t = Replace(t, ChrW(263), "c")
    t = Replace(t, ChrW(272), "Dj"): t = Replace(t, ChrW(273), "dj")
    t = Replace(t, " ", "-")
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch Like "[A-Za-z0-9]") Or ch = "." Or ch = "_" Or ch = "-" Then out = out & ch
    Next i
    SanitizeFileNamePart = out
End Function

'======================================================================
' BankaNazivZaRacun - ime banke iz vodeceg NBS koda racuna (prve 3
' cifre / deo pre prve crtice). Za prikaz u combo-u "Sa racuna"
' (frmBankaExportPregled) kada firma ima racune u vise banaka.
' Nepoznat kod -> "" (prikaze se samo racun).
'======================================================================
Public Function BankaNazivZaRacun(ByVal racun As String) As String
    Dim r As String
    r = Replace(Trim$(racun), " ", "")

    Dim kod As String
    If InStr(r, "-") > 0 Then
        kod = Left$(r, InStr(r, "-") - 1)
    ElseIf Len(r) >= 3 Then
        kod = Left$(r, 3)
    End If

    Select Case kod
        Case "105": BankaNazivZaRacun = "AIK"
        Case "155": BankaNazivZaRacun = "Halkbank"
        Case "160": BankaNazivZaRacun = "Banca Intesa"
        Case "165": BankaNazivZaRacun = "Addiko"
        Case "170": BankaNazivZaRacun = "UniCredit"
        Case "200": BankaNazivZaRacun = "Po" & ChrW(353) & "tanska " & ChrW(353) & "tedionica"
        Case "205": BankaNazivZaRacun = "NLB Komercijalna"
        Case "220": BankaNazivZaRacun = "ProCredit"
        Case "265": BankaNazivZaRacun = "Raiffeisen"
        Case "275", "325": BankaNazivZaRacun = "OTP"
        Case "340": BankaNazivZaRacun = "Erste"
        Case Else: BankaNazivZaRacun = ""
    End Select
End Function

