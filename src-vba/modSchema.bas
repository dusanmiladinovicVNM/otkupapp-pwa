Attribute VB_Name = "modSchema"

Option Explicit

'=====================================================================
' modSchema -- REGISTAR SEME TABELA
'
' Izvor istine za STRUKTURU tabela je schema/schema.json u gitu. Ovaj modul
' je GENERISAN ARTEFAKT te datoteke, a sveska je posledica -- ne izvor.
' Do PR1 je vazilo obrnuto: spiskovi kolona ziveli su iskljucivo u .xlsm, pa
' se prazna sveska nije mogla rekonstruisati, a obrisana kolona se videla tek
' kao pad upisa satima kasnije.
'
' GENERISAN FAJL -- ne menjaj rukom:
'     (izmeni schema/schema.json, pa:)
'     python tools/gen_schema_module.py
'
' REDOSLED KOLONA JE DEO SEME. modDataAccess.AppendRow pise POZICIONO, a
' pisci poput modOtkup.SaveOtkup grade goli Array(...) -- kolona ubacena u
' sredinu tiho pomera sve iza sebe u pogresne kolone. Zato se nove kolone
' dodaju NA KRAJ, a otisak (SchemaFingerprint) racuna nad UREDJENIM kolonama.
'
' Javni ulazi:
'   EnsureAllTables      kreira sto fali, dopunjava kolone. Idempotentno.
'   VerifySchema         prijavljuje odstupanja; NISTA ne menja.
'   SchemaReadyOrFail    tvrda kapija pred upis, nad zadatim tabelama.
'   SchemaCheckOnStart   jeftina provera pri pokretanju (otisak).
'   SchemaRegistry       registar, za testove i alate.
'
' NE zvati EnsureAllTables unutar otvorene transakcije: clsTransaction.RestoreTable
' pada na neslaganje broja kolona, pa bi promena seme onemogucila rollback.
'=====================================================================

' Vrste odstupanja koje VerifySchema vraca (prefiks stavke u kolekciji).
Public Const SCHEMA_DRIFT_TABELA As String = "TABELA"
Public Const SCHEMA_DRIFT_KOLONA As String = "KOLONA"
Public Const SCHEMA_DRIFT_REDOSLED As String = "REDOSLED"

' Otisak kanonske seme (FNV-1a 32 nad "tbl|kol|kol;..." REDOM). Generisan
' zajedno sa registrom -- ne menjati rukom.
Public Const SCHEMA_FINGERPRINT As String = "3B702DE1"

' Kes registra. Registar je DEKLARACIJA, ne snimak sveske, pa se ne menja
' u toku rada -- kesiranje je bezbedno.
Private mReg As Object


'=====================================================================
' JAVNI API
'=====================================================================

Public Function SchemaRegistry() As Object
    If mReg Is Nothing Then Set mReg = BuildRegistry()
    Set SchemaRegistry = mReg
End Function

' Kolone jedne tabele, REDOM iz kanona. Prazna kolekcija = tabela nije u
' registru (pozivalac razlikuje "nema je" od "nema kolona").
Public Function SchemaTableColumns(ByVal tblName As String) As Collection
    Dim out As Collection
    Dim delovi() As String
    Dim i As Long

    Set out = New Collection
    delovi = Split(RegKolone(tblName), "|")
    For i = LBound(delovi) To UBound(delovi)
        If Len(delovi(i)) > 0 Then out.Add delovi(i)
    Next i

    Set SchemaTableColumns = out
End Function

Public Function SchemaTableSheet(ByVal tblName As String) As String
    Dim reg As Object
    Dim v As String
    Dim p As Long

    Set reg = SchemaRegistry()
    If Not reg.Exists(tblName) Then Exit Function

    v = CStr(reg(tblName))
    p = InStr(1, v, vbTab, vbBinaryCompare)
    If p > 0 Then SchemaTableSheet = Left$(v, p - 1)
End Function

' Kreira tabele kojih nema i dopunjava kolone koje fale. Idempotentno.
' Pad JEDNE tabele se zapise i NE zaustavlja ostale -- isti razlog kao
' EnsureKolonaSaTragom u modSetup: blanket "On Error Resume Next" je cutke
' preskakao ostatak posla, pa se posledica videla tek kao pad upisa.
'
' NE popravlja REDOSLED: premestanje kolone u postojecoj tabeli bi pomerilo
' podatke. Pogresan redosled je nalaz za coveka, ne nesto sto se leci u prolazu.
Public Sub EnsureAllTables()
    Dim reg As Object
    Dim tblName As Variant

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        EnsureJednuTabelu CStr(tblName)
    Next tblName
End Sub

' Odstupanja sveske od kanona. NISTA ne menja -- dijagnostika ne sme da
' zameni gresku koju opisuje.
' Stavka je "VRSTA|tabela|detalj".
Public Function VerifySchema() As Collection
    Dim out As Collection
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim imena As Object
    Dim lc As ListColumn
    Dim delovi() As String
    Dim i As Long
    Dim ocekivano As String
    Dim neslaganje As String

    Set out = New Collection
    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        Set lo = Nothing
        On Error Resume Next
        Set lo = modDataAccess.GetTable(CStr(tblName))
        On Error GoTo 0

        If lo Is Nothing Then
            out.Add SCHEMA_DRIFT_TABELA & "|" & CStr(tblName) & "|"
        Else
            Set imena = CreateObject("Scripting.Dictionary")
            imena.CompareMode = vbTextCompare
            For Each lc In lo.ListColumns
                If Not imena.Exists(lc.name) Then imena.Add lc.name, True
            Next lc

            ocekivano = RegKolone(CStr(tblName))
            delovi = Split(ocekivano, "|")
            For i = LBound(delovi) To UBound(delovi)
                If Len(delovi(i)) > 0 Then
                    If Not imena.Exists(delovi(i)) Then
                        out.Add SCHEMA_DRIFT_KOLONA & "|" & CStr(tblName) & "|" & _
                                delovi(i)
                    End If
                End If
            Next i

            ' Redosled: kanon mora biti PREFIKS stvarnog zaglavlja, po INDEKSU
            ' KOLONE. Visak na kraju je dozvoljen (kolona koju kanon jos ne zna),
            ' ali svako razilazenje PRE kraja znaci da je pozicion upis promasen.
            neslaganje = PrefiksNeslaganje(lo, SchemaTableColumns(CStr(tblName)))
            If Len(neslaganje) > 0 Then
                out.Add SCHEMA_DRIFT_REDOSLED & "|" & CStr(tblName) & _
                        "|" & neslaganje
            End If
        End If
    Next tblName

    Set VerifySchema = out
End Function

' Jeftina provera pri pokretanju: otisak sveske vs otisak kanona.
' Vraca "" kad je sve u redu, inace kratak opis (pozivalac odlucuje sta s tim).
'
' Otisak se racuna nad UREDJENIM kolonama, pa hvata i nedostatak i preraspored.
' Cena je ~590 citanja .Name -- reda velicine 10 ms, zanemarljivo po startu.
' (EnsureAllTables je skup jer PISE; ovo samo cita.)
Public Function SchemaCheckOnStart() As String
    Dim stvarni As String
    Dim odstupanja As Collection
    Dim i As Long
    Dim n As Long
    Dim prikaz As String

    On Error GoTo EH

    stvarni = SchemaFingerprintActual()
    If StrComp(stvarni, SCHEMA_FINGERPRINT, vbBinaryCompare) = 0 Then Exit Function

    ' Otisak se razlikuje -- tek sada se placa pun prolaz, da bi poruka rekla STA.
    Set odstupanja = VerifySchema()
    If odstupanja.count = 0 Then
        ' Otisak i VerifySchema mere isto (kanonski prefiks), pa ovo znaci da
        ' im se implementacije razisle -- kvar u kodu, ne u svesci.
        SchemaCheckOnStart = "Otisak (" & stvarni & " vs " & SCHEMA_FINGERPRINT & _
                             ") se ne slaze sa VerifySchema, koja ne vidi nijedno " & _
                             "odstupanje. Dve provere su se razisle -- prijavi kao bug."
        Exit Function
    End If

    n = odstupanja.count
    If n > 5 Then n = 5
    For i = 1 To n
        If Len(prikaz) > 0 Then prikaz = prikaz & "; "
        prikaz = prikaz & CStr(odstupanja(i))
    Next i
    If odstupanja.count > 5 Then
        prikaz = prikaz & "; ... (+" & CStr(odstupanja.count - 5) & ")"
    End If

    SchemaCheckOnStart = "Sema odstupa od kanona (" & CStr(odstupanja.count) & _
                         "): " & prikaz
    Exit Function

EH:
    SchemaCheckOnStart = "Provera seme nije izvrsena: " & Err.description
End Function

' FNV-1a 32 nad "tbl|kol|kol;..." REDOM, po registru. Isti algoritam kao
' tools/gen_schema_module.py.otisak -- kad se jedan menja, mora i drugi.
'
' Meri se KANONSKI PREFIKS zaglavlja, ne celo zaglavlje. Dva razloga:
'
'   1. modSetup.EnsureRuntimeSchema dodaje kolone na svakom startu i one idu
'      NA KRAJ. One su legitimne (samo ih kanon jos ne drzi), pa bi otisak nad
'      celim zaglavljem prijavljivao trajan lazan drift -- a provera koju
'      operater nauci da ignorise ne stiti nista.
'   2. Pozicion upis (AppendRow) zavisi TACNO od prefiksa: sve iza kanonske
'      poslednje kolone ne moze da pomeri nijednu vrednost koju pisci salju.
'
' Time se otisak i VerifySchema poklapaju po konstrukciji: oba mere isto.
' Kolona koja FALI ili je PREMESTENA unutar prefiksa menja otisak; rep ne.
Public Function SchemaFingerprintActual() As String
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim delovi As String
    Dim red As String
    Dim koliko As Long
    Dim i As Long

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        Set lo = Nothing
        On Error Resume Next
        Set lo = modDataAccess.GetTable(CStr(tblName))
        On Error GoTo 0

        red = CStr(tblName)
        If Not lo Is Nothing Then
            koliko = UBound(Split(RegKolone(CStr(tblName)), "|"))
            If koliko > lo.ListColumns.count Then koliko = lo.ListColumns.count
            For i = 1 To koliko
                red = red & "|" & lo.ListColumns(i).name
            Next i
        End If

        If Len(delovi) > 0 Then delovi = delovi & ";"
        delovi = delovi & red
    Next tblName

    SchemaFingerprintActual = Fnv1a32(delovi)
End Function

' Tvrda kapija pred upis. tblList je "tblA|tblB|tblC" -- proverava se SAMO
' to, jer pun prolaz po svakom upisu je preskup.
'
' Proverava se i REDOSLED, ne samo prisustvo: AppendRow pise poziciono, pa
' tabela sa svim kolonama u pogresnom rasporedu tiho upisuje u pogresna polja.
' To je gore od pada upisa.
Public Sub SchemaReadyOrFail(ByVal sourceName As String, ByVal tblList As String)
    Dim reg As Object
    Dim delovi() As String
    Dim tblName As String
    Dim lo As ListObject
    Dim i As Long
    Dim neslaganje As String

    Set reg = SchemaRegistry()
    delovi = Split(tblList, "|")

    For i = LBound(delovi) To UBound(delovi)
        tblName = Trim$(delovi(i))
        If Len(tblName) > 0 Then
            If Not reg.Exists(tblName) Then
                Err.Raise vbObjectError + 9403, sourceName, _
                          "Tabela '" & tblName & "' nije u kanonskoj semi " & _
                          "(schema/schema.json). Dopuni kanon pa regenerisi."
            End If

            Set lo = Nothing
            On Error Resume Next
            Set lo = modDataAccess.GetTable(tblName)
            On Error GoTo 0

            If lo Is Nothing Then
                Err.Raise vbObjectError + 9404, sourceName, _
                          "Tabela '" & tblName & "' ne postoji u svesci. " & _
                          "Pokreni modSchema.EnsureAllTables pa ponovi."
            End If

            neslaganje = PrefiksNeslaganje(lo, SchemaTableColumns(tblName))
            If Len(neslaganje) > 0 Then
                Err.Raise vbObjectError + 9405, sourceName, _
                          "Tabela '" & tblName & "' ne odgovara kanonskoj semi (" & _
                          neslaganje & "). Upis je POZICION, pa bi vrednosti " & _
                          "otisle u pogresne kolone. Pokreni " & _
                          "modSchema.EnsureAllTables; ako i posle toga odstupa, " & _
                          "redosled se mora popraviti rucno."
            End If
        End If
    Next i
End Sub


'=====================================================================
' INTERNO
'=====================================================================

' FNV-1a 32: h = (h Xor bajt) * 16777619 mod 2^32.
'
' Double, ne Long: VBA Long je SA ZNAKOM, pa bi i medjurezultat i konacna
' vrednost preticali. Tri zamke koje su ovde vec placene:
'
'   1. Xor nad celom 32-bitnom vrednoscu ne treba. Bajt je < 256, pa menja
'      SAMO donjih 8 bita -- radi se Xor nad tim bajtom, u Long-u gde je
'      bezbedan. (Prva verzija je vrtela petlju po 32 bita: tacno, ali
'      ~320.000 iteracija za ovaj string.)
'   2. Mul32 mora da svede visu polovinu po modulu PRE mnozenja sa 65536,
'      inace medjurezultat predje 2^53 i Double pocne da gubi cifre.
'   3. Hex$ prima Long (sa znakom), pa preti preko 2^31. Ispis ide iz dve
'      16-bitne polovine.
Public Function Fnv1a32(ByVal s As String) As String
    Dim h As Double
    Dim i As Long
    Dim b As Long
    Dim nizak As Double
    Dim visa As Long
    Dim niza As Long

    h = 2166136261#

    For i = 1 To Len(s)
        b = Asc(Mid$(s, i, 1)) And &HFF&
        nizak = h - Int(h / 256#) * 256#
        h = h - nizak + CDbl(CLng(nizak) Xor b)
        h = Mul32(h, 16777619#)
    Next i

    visa = CLng(Int(h / 65536#))
    niza = CLng(h - CDbl(visa) * 65536#)
    Fnv1a32 = Right$("0000" & Hex$(visa), 4) & Right$("0000" & Hex$(niza), 4)
End Function

Private Function Mul32(ByVal a As Double, ByVal b As Double) As Double
    ' (a * b) mod 2^32. a se deli na dve 16-bitne polovine da nijedan
    ' medjurezultat ne predje 2^53 (granica tacnosti Double-a).
    Dim lo As Double
    Dim hi As Double

    hi = Int(a / 65536#)
    lo = a - hi * 65536#

    lo = Modulo32(lo * b)
    hi = Modulo32(Modulo32(hi * b) * 65536#)

    Mul32 = Modulo32(lo + hi)
End Function

Private Function Modulo32(ByVal v As Double) As Double
    Modulo32 = v - Int(v / 4294967296#) * 4294967296#
End Function

' Da li zaglavlje tabele pocinje TACNO kanonskim kolonama, po INDEKSU.
'
' Poredjenje po stringu ("|A|B" u "|A|BExtra") daje LAZAN prolaz: InStr vrati 1,
' a kolona B ne postoji. Ranjiva je bas poslednja kanonska kolona, jer iza nje
' nema delimitera. Zato se poredi kolona po kolona.
'
' Vraca "" kad je sve u redu, inace opis PRVOG neslaganja -- pozivalac odlucuje
' da li ga prijavljuje ili dize gresku.
'
' Jedan helper za VerifySchema i SchemaReadyOrFail: dve kapije ne smeju da
' razviju razlicite definicije "ispravnog prefiksa".
Private Function PrefiksNeslaganje(ByVal lo As ListObject, _
                                   ByVal kolone As Collection) As String
    Dim i As Long
    Dim stvarno As String

    If lo.ListColumns.count < kolone.count Then
        PrefiksNeslaganje = "tabela ima " & CStr(lo.ListColumns.count) & _
                            " kolona, kanon trazi " & CStr(kolone.count)
        Exit Function
    End If

    For i = 1 To kolone.count
        stvarno = lo.ListColumns(i).name
        If StrComp(stvarno, CStr(kolone(i)), vbTextCompare) <> 0 Then
            PrefiksNeslaganje = "pozicija " & CStr(i) & ": ocekivano '" & _
                                CStr(kolone(i)) & "', stvarno '" & stvarno & "'"
            Exit Function
        End If
    Next i
End Function

Private Sub EnsureJednuTabelu(ByVal tblName As String)
    Dim kolone As Collection
    Dim arr() As String
    Dim i As Long

    On Error GoTo EH

    Set kolone = SchemaTableColumns(tblName)
    If kolone.count = 0 Then
        Err.Raise vbObjectError + 9406, "modSchema.EnsureJednuTabelu", _
                  "Kanon nema nijednu kolonu za '" & tblName & "'."
    End If

    ReDim arr(1 To kolone.count)
    For i = 1 To kolone.count
        arr(i) = CStr(kolone(i))
    Next i

    modSetup.EnsureDataTable tblName, SchemaTableSheet(tblName), arr
    Exit Sub

EH:
    LogError "modSchema.EnsureAllTables", _
             "Tabela '" & tblName & "' nije obezbedjena: " & Err.description, _
             Err.Number
End Sub

' Ime NIJE "Reg": parametar svake Spec* procedure se zove "reg", a VBA je
' case-insensitive -- parametar bi zaklonio proceduru, pa bi "Reg reg, ..."
' postalo pozivanje Dictionary-ja sa cetiri argumenta (runtime 438). Ista klasa
' greske koju hvata vba_check pravilo ZAKLONJENO, ali ono namerno preskace
' objekte, pa je ovo nasao tek test.
'
' Registar cuva STRING, ne ugnjezden objekat: "sheet" & vbTab & "|kol|kol".
'
' Prva verzija je drzala Dictionary po tabeli sa Collection kolona. Ne radi:
' "Set d("kolone") = obj" nad late-bound Scripting.Dictionary trazi Property Set,
' koji Dictionary ne izlaze -- runtime 438 ("Object doesn't support this property
' or method"). vba_check to ne moze da vidi; uhvatila su ga tek dva testa.
'
' Uz to je brze: otisak i poredjenje redosleda ionako rade nad spojenim stringom,
' pa se nista ne sklapa dvaput.
Private Sub RegistrujTabelu(ByVal reg As Object, ByVal tblName As String, _
                ByVal sheetName As String, ByVal kolone As Collection)
    Dim spojeno As String
    Dim i As Long

    If reg.Exists(tblName) Then
        Err.Raise vbObjectError + 9401, "modSchema.RegistrujTabelu", _
                  "Dupla tabela u registru: " & tblName
    End If

    For i = 1 To kolone.count
        spojeno = spojeno & "|" & CStr(kolone(i))
    Next i

    reg.Add tblName, sheetName & vbTab & spojeno
End Sub

' "|kol|kol" iz registra -- oblik koji otisak i poredjenje redosleda traze.
Private Function RegKolone(ByVal tblName As String) As String
    Dim reg As Object
    Dim v As String
    Dim p As Long

    Set reg = SchemaRegistry()
    If Not reg.Exists(tblName) Then Exit Function

    v = CStr(reg(tblName))
    p = InStr(1, v, vbTab, vbBinaryCompare)
    If p > 0 Then RegKolone = Mid$(v, p + 1)
End Function


'=====================================================================
' REGISTAR -- GENERISANO iz schema/schema.json, ne menjaj rukom
'=====================================================================

Private Function BuildRegistry() As Object
    Dim reg As Object
    Set reg = CreateObject("Scripting.Dictionary")
    reg.CompareMode = vbTextCompare

    SpecAmbalaza reg
    SpecArtikli reg
    SpecBankaImport reg
    SpecCenovnik reg
    SpecConfig reg
    SpecFakturaStavke reg
    SpecFakture reg
    SpecKese reg
    SpecKooperanti reg
    SpecKorisnici reg
    SpecKulture reg
    SpecKupci reg
    SpecKutije reg
    SpecLocalConfig reg
    SpecMagacin reg
    SpecMGMT reg
    SpecNovac reg
    SpecOtkup reg
    SpecOtpremnica reg
    SpecPaleta reg
    SpecPaletaStavka reg
    SpecParcele reg
    SpecPartnerMap reg
    SpecPoruke reg
    SpecPrerada reg
    SpecPreradaStavka reg
    SpecPrevoznici reg
    SpecPrijemnica reg
    SpecSEFConfig reg
    SpecSEFEventLog reg
    SpecSEFSubmission reg
    SpecStanice reg
    SpecStornoVeze reg
    SpecStornoZurnal reg
    SpecTipAmbalaze reg
    SpecTipPalete reg
    SpecUtovar reg
    SpecUtovarStavke reg
    SpecVozaci reg
    SpecVrstaGotovihProizvoda reg
    SpecZbirna reg

    Set BuildRegistry = reg
End Function


Private Sub SpecAmbalaza(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "AmbID"
    k.Add "Datum"
    k.Add "TipAmbalaze"
    k.Add "Kolicina"
    k.Add "Smer"
    k.Add "EntitetID"
    k.Add "EntitetTip"
    k.Add "VozacID"
    k.Add "DokumentID"
    k.Add "DokumentTIP"
    k.Add "Stornirano"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_AMBALAZA, "Ambalaza", k
End Sub

Private Sub SpecArtikli(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ArtikalID"
    k.Add "Naziv"
    k.Add "Tip"
    k.Add "JedinicaMere"
    k.Add "CenaPoJedinici"
    k.Add "DozaPoHa"
    k.Add "Kultura"
    k.Add "Pakovanje"
    k.Add "BarKod"
    k.Add "KarencaDana"
    k.Add "Aktivan"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_ARTIKLI, "Artikli", k
End Sub

Private Sub SpecBankaImport(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "BankaImportID"
    k.Add "BrojDokumenta"
    k.Add "DatumIzvoda"
    k.Add "BrojRacuna"
    k.Add "DatumTransakcije"
    k.Add "Partner"
    k.Add "PartnerKonto"
    k.Add "Opis"
    k.Add "Uplata"
    k.Add "Isplata"
    k.Add "Valuta"
    k.Add "PozivNaBroj"
    k.Add "SvrhaPlacanja"
    k.Add "BankaReferenz"
    k.Add "IzvorFajl"
    k.Add "ImportVreme"
    k.Add "Obradjeno"
    k.Add "Stornirano"
    k.Add "PocetnoStanje"
    k.Add "ZavrsnoStanje"
    k.Add "UkupanDuguje"
    k.Add "UkupanPotrazuje"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_BANKA_IMPORT, "BankaImport", k
End Sub

Private Sub SpecCenovnik(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "CenaID"
    k.Add "Datum"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "Klasa"
    k.Add "Cena"
    k.Add "CreatedAt"
    k.Add "Stornirano"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_CENOVNIK, "Cenovnik", k
End Sub

Private Sub SpecConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Vrednost"
    k.Add "Opis"
    RegistrujTabelu reg, TBL_CONFIG, "Config", k
End Sub

Private Sub SpecFakturaStavke(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "StavkaID"
    k.Add "FakturaID"
    k.Add "PrijemnicaID"
    k.Add "Kolicina"
    k.Add "Cena"
    k.Add "Klasa"
    k.Add "BrojPrijemnice"
    k.Add "Stornirano"
    k.Add "OsirocenoOd"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "PreradaID"
    k.Add "BrojPrerade"
    k.Add "UtovarID"
    RegistrujTabelu reg, TBL_FAKTURA_STAVKE, "FakturaStavke", k
End Sub

Private Sub SpecFakture(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "FakturaID"
    k.Add "BrojFakture"
    k.Add "Datum"
    k.Add "KupacID"
    k.Add "Iznos"
    k.Add "Status"
    k.Add "DatumPlacanja"
    k.Add "Stornirano"
    k.Add "OsirocenoOd"
    k.Add "SEFWorkflowState"
    k.Add "SEFStatus"
    k.Add "SEFDocumentId"
    k.Add "SEFInvoiceNumber"
    k.Add "SEFSentAt"
    k.Add "SEFLastSyncAt"
    k.Add "SEFLastErrorCode"
    k.Add "SEFLastErrorMessage"
    k.Add "SEFPayloadHash"
    k.Add "SEFVersionNo"
    k.Add "PoslatNaSEF"
    k.Add "SEFSubmissionIDLast"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "GeneracijaID"
    RegistrujTabelu reg, TBL_FAKTURE, "Fakture", k
End Sub

Private Sub SpecKese(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipKese"
    k.Add "TezinaKg"
    k.Add "Aktivan"
    RegistrujTabelu reg, TBL_KESE, "Kese", k
End Sub

Private Sub SpecKooperanti(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "KooperantID"
    k.Add "Ime"
    k.Add "Prezime"
    k.Add "Mesto"
    k.Add "Telefon"
    k.Add "StanicaID"
    k.Add "Aktivan"
    k.Add "BPGBroj"
    k.Add "TekuciRacun"
    k.Add "Pin"
    k.Add "Adresa"
    k.Add "JMBG"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_KOOPERANTI, "Kooperanti", k
End Sub

Private Sub SpecKorisnici(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "KorisnikID"
    k.Add "Username"
    k.Add "ImePrezime"
    k.Add "PIN"
    k.Add "Uloga"
    k.Add "Aktivan"
    k.Add "StanicaID"
    k.Add "CreatedAt"
    k.Add "Otkup"
    k.Add "Dokumenta"
    k.Add "Agrohemija"
    k.Add "Izvestaji"
    k.Add "Fakturisanje"
    k.Add "Banka"
    k.Add "Marza"
    k.Add "Sledljivost"
    k.Add "MaticniPodaci"
    k.Add "Palete"
    k.Add "OtvoriExcel"
    k.Add "SyncPWA"
    RegistrujTabelu reg, TBL_KORISNICI, "Korisnici", k
End Sub

Private Sub SpecKulture(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "KulturaID"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "GajbicaPoPaleti"
    k.Add "Aktivan"
    k.Add "TipAmbalaze"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "PragProsekUpoz"
    k.Add "PragProsekBlok"
    RegistrujTabelu reg, TBL_KULTURE, "Kulture", k
End Sub

Private Sub SpecKupci(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "KupacID"
    k.Add "Naziv"
    k.Add "Ulica"
    k.Add "Mesto"
    k.Add "PostanskiBroj"
    k.Add "Drzava"
    k.Add "PIB"
    k.Add "MaticniBroj"
    k.Add "Email"
    k.Add "Hladnjaca"
    k.Add "Aktivan"
    k.Add "TekuciRacun"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_KUPCI, "Kupci", k
End Sub

Private Sub SpecKutije(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipKutije"
    k.Add "TezinaKg"
    k.Add "Aktivan"
    RegistrujTabelu reg, TBL_KUTIJE, "Kutije", k
End Sub

Private Sub SpecLocalConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Vrednost"
    k.Add "Opis"
    RegistrujTabelu reg, TBL_LOCAL_CONFIG, "LocalConfig", k
End Sub

Private Sub SpecMagacin(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "MagacinID"
    k.Add "Datum"
    k.Add "ArtikalID"
    k.Add "Tip"
    k.Add "Kolicina"
    k.Add "KooperantID"
    k.Add "ParcelaID"
    k.Add "BrojDokumenta"
    k.Add "CenaPoJedinici"
    k.Add "Vrednost"
    k.Add "Napomena"
    k.Add "Stornirano"
    k.Add "DobavljacID"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_MAGACIN, "Magacin", k
End Sub

Private Sub SpecMGMT(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "MGMTID"
    k.Add "Ime"
    k.Add "Prezime"
    k.Add "Mesto"
    k.Add "Adresa"
    k.Add "JMBG"
    k.Add "Telefon"
    k.Add "Email"
    k.Add "Aktivan"
    k.Add "Pozicija"
    k.Add "TekuciRacun"
    k.Add "PIN"
    RegistrujTabelu reg, TBL_MGMT, "MGMT", k
End Sub

Private Sub SpecNovac(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "NovacID"
    k.Add "BrojDokumenta"
    k.Add "Datum"
    k.Add "Partner"
    k.Add "PartnerID"
    k.Add "EntitetTip"
    k.Add "OMID"
    k.Add "KooperantID"
    k.Add "FakturaID"
    k.Add "VrstaVoca"
    k.Add "Tip"
    k.Add "Uplata"
    k.Add "Isplata"
    k.Add "Napomena"
    k.Add "Stornirano"
    k.Add "OtkupID"
    k.Add "OsirocenoOd"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "GeneracijaID"
    RegistrujTabelu reg, TBL_NOVAC, "Novac", k
End Sub

Private Sub SpecOtkup(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "OtkupID"
    k.Add "Datum"
    k.Add "KooperantID"
    k.Add "StanicaID"
    k.Add "KulturaID"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "Kolicina"
    k.Add "Cena"
    k.Add "TipAmbalaze"
    k.Add "KolAmbalaze"
    k.Add "VozacID"
    k.Add "BrojDokumenta"
    k.Add "Novac"
    k.Add "PrimalacNovca"
    k.Add "Klasa"
    k.Add "Stornirano"
    k.Add "BrojZbirne"
    k.Add "Isplaceno"
    k.Add "DatumIsplate"
    k.Add "OtpremnicaID"
    k.Add "ParcelaID"
    k.Add "ClientRecordID"
    k.Add "SyncSource"
    k.Add "KolAmbIzdata"
    k.Add "VremeUnosa"
    k.Add "BrutoKg"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "BrojOtpremnice"
    k.Add "GeneracijaID"
    k.Add "ZbirnaGeneracijaID"
    RegistrujTabelu reg, TBL_OTKUP, "Otkup", k
End Sub

Private Sub SpecOtpremnica(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "OtpremnicaID"
    k.Add "Datum"
    k.Add "StanicaID"
    k.Add "VozacID"
    k.Add "BrojOtpremnice"
    k.Add "BrojZbirne"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "Kolicina"
    k.Add "Cena"
    k.Add "TipAmbalaze"
    k.Add "KolAmbalaze"
    k.Add "Klasa"
    k.Add "Stornirano"
    k.Add "BrutoKg"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "GeneracijaID"
    k.Add "ZbirnaGeneracijaID"
    RegistrujTabelu reg, TBL_OTPREMNICA, "Otpremnica", k
End Sub

Private Sub SpecPaleta(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "PaletaID"
    k.Add "BrojPalete"
    k.Add "Godina"
    k.Add "Datum"
    k.Add "VrstaVoca"
    k.Add "TipPalete"
    k.Add "BrojGajbica"
    k.Add "NetoKg"
    k.Add "AmbalazaKg"
    k.Add "PaletaKg"
    k.Add "BrutoKg"
    k.Add "Status"
    k.Add "Stornirano"
    k.Add "Preradjeno"
    k.Add "SortaVoca"
    k.Add "Klasa"
    k.Add "TipAmbalaze"
    k.Add "KapacitetGajbica"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "Istorija"
    RegistrujTabelu reg, TBL_PALETA, "Paleta", k
End Sub

Private Sub SpecPaletaStavka(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "StavkaID"
    k.Add "PaletaID"
    k.Add "BrojPrijemnice"
    k.Add "BrojZbirne"
    k.Add "BrojGajbica"
    k.Add "NetoKg"
    k.Add "AmbalazaKg"
    k.Add "PrijemnicaID"
    k.Add "Klasa"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "CreatedAt"
    k.Add "Stornirano"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "ZbirnaGeneracijaID"
    RegistrujTabelu reg, TBL_PALETA_STAVKA, "PaletaStavka", k
End Sub

Private Sub SpecParcele(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ParcelaID"
    k.Add "KooperantID"
    k.Add "KatBroj"
    k.Add "KatOpstina"
    k.Add "Kultura"
    k.Add "PovrsinaHa"
    k.Add "GGAPStatus"
    k.Add "Aktivna"
    k.Add "GeoStatus"
    k.Add "GeoSource"
    k.Add "N_Coord"
    k.Add "E_Coord"
    k.Add "Lat"
    k.Add "Lng"
    k.Add "PolygonGeoJSON"
    k.Add "MeteoEnabled"
    k.Add "RizikStatus"
    k.Add "DatumGeoUnosa"
    k.Add "DatumAzuriranja"
    k.Add "Napomena"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_PARCELE, "Parcele", k
End Sub

Private Sub SpecPartnerMap(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "BankaName"
    k.Add "PartnerID"
    k.Add "EntitetTip"
    k.Add "OMID"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_PARTNER_MAP, "PartnerMap", k
End Sub

Private Sub SpecPoruke(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Tekst"
    RegistrujTabelu reg, TBL_PORUKE, "sPoruke", k
End Sub

Private Sub SpecPrerada(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "PreradaID"
    k.Add "BrojPrerade"
    k.Add "Godina"
    k.Add "Datum"
    k.Add "NetoKolicina"
    k.Add "BrojKutija"
    k.Add "BrojKesa"
    k.Add "Stornirano"
    k.Add "NetoUlazKg"
    k.Add "NetoIzlazKg"
    k.Add "Napomena"
    k.Add "CreatedAt"
    k.Add "TezinaPaleteKg"
    k.Add "BrutoKg"
    k.Add "AmbalazaKg"
    k.Add "TipKutije"
    k.Add "TipKese"
    k.Add "TipGotovogProizvoda"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "DatumIsteka"
    RegistrujTabelu reg, TBL_PRERADA, "Prerada", k
End Sub

Private Sub SpecPreradaStavka(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "StavkaID"
    k.Add "PreradaID"
    k.Add "PaletaID"
    k.Add "BrojPalete"
    k.Add "NetoKg"
    k.Add "CreatedAt"
    k.Add "Stornirano"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_PRERADA_STAVKA, "PreradaStavke", k
End Sub

Private Sub SpecPrevoznici(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "PrevoznikID"
    k.Add "Naziv"
    k.Add "Vozac"
    k.Add "Registracija"
    k.Add "Aktivan"
    RegistrujTabelu reg, TBL_PREVOZNICI, "Prevoznici", k
End Sub

Private Sub SpecPrijemnica(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "PrijemnicaID"
    k.Add "Datum"
    k.Add "KupacID"
    k.Add "VozacID"
    k.Add "BrojPrijemnice"
    k.Add "BrojZbirne"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "Kolicina"
    k.Add "Cena"
    k.Add "TipAmbalaze"
    k.Add "KolAmbalaze"
    k.Add "KolAmbVracena"
    k.Add "Klasa"
    k.Add "Fakturisano"
    k.Add "FakturaID"
    k.Add "Stornirano"
    k.Add "BrutoKg"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "GeneracijaID"
    k.Add "ZbirnaGeneracijaID"
    RegistrujTabelu reg, TBL_PRIJEMNICA, "Prijemnica", k
End Sub

Private Sub SpecSEFConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ConfigKey"
    k.Add "ConfigValue"
    k.Add "Opis"
    k.Add "Aktivan"
    RegistrujTabelu reg, TBL_SEF_CONFIG, "SEFConfig", k
End Sub

Private Sub SpecSEFEventLog(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "SEFEventID"
    k.Add "FakturaID"
    k.Add "SEFSubmissionID"
    k.Add "EventTime"
    k.Add "EventType"
    k.Add "Message"
    k.Add "Details"
    k.Add "OperatorName"
    k.Add "Stornirano"
    RegistrujTabelu reg, TBL_SEF_EVENT_LOG, "SEFEventLog", k
End Sub

Private Sub SpecSEFSubmission(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "SEFSubmissionID"
    k.Add "FakturaID"
    k.Add "VersionNo"
    k.Add "WorkflowStateAtSubmit"
    k.Add "CreatedAt"
    k.Add "SubmittedAt"
    k.Add "SubmissionStatus"
    k.Add "PayloadHash"
    k.Add "RequestFormat"
    k.Add "RequestBody"
    k.Add "ResponseBody"
    k.Add "HttpStatus"
    k.Add "ApiStatus"
    k.Add "CorrelationId"
    k.Add "SEFDocumentId"
    k.Add "ErrorCode"
    k.Add "ErrorMessage"
    k.Add "OperatorName"
    k.Add "Stornirano"
    k.Add "FinishedAt"
    RegistrujTabelu reg, TBL_SEF_SUBMISSION, "SEFSubmission", k
End Sub

Private Sub SpecStanice(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "StanicaID"
    k.Add "Naziv"
    k.Add "Mesto"
    k.Add "Kontakt"
    k.Add "Aktivan"
    k.Add "Ime"
    k.Add "Prezime"
    k.Add "PIN"
    k.Add "JeHladnjaca"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_STANICE, "Otkupna Mesta", k
End Sub

Private Sub SpecStornoVeze(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "CorrectionID"
    k.Add "Mode"
    k.Add "Status"
    k.Add "OldDocType"
    k.Add "OldDocID"
    k.Add "OldBroj"
    k.Add "NewDocType"
    k.Add "NewDocID"
    k.Add "NewBroj"
    k.Add "ParentDocType"
    k.Add "ParentDocID"
    k.Add "ParentBroj"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "CompletedAt"
    k.Add "Message"
    k.Add "NeedsRecovery"
    k.Add "RecoveryAction"
    RegistrujTabelu reg, TBL_STORNO_VEZE, "StornoVeze", k
End Sub

Private Sub SpecStornoZurnal(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ZurnalID"
    k.Add "OperationID"
    k.Add "Timestamp"
    k.Add "DocType"
    k.Add "Broj"
    k.Add "Tabela"
    k.Add "RowID"
    k.Add "Kolona"
    k.Add "StaraVrednost"
    k.Add "NovaVrednost"
    RegistrujTabelu reg, TBL_STORNO_ZURNAL, "StornoZurnal", k
End Sub

Private Sub SpecTipAmbalaze(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipAmbalaze"
    k.Add "TezinaGajbiceKg"
    k.Add "Aktivan"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_TIP_AMBALAZE, "TipAmbalaze", k
End Sub

Private Sub SpecTipPalete(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipPalete"
    k.Add "TezinaKg"
    k.Add "Aktivan"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_TIP_PALETE, "TipPalete", k
End Sub

Private Sub SpecUtovar(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "UtovarID"
    k.Add "BrojUtovara"
    k.Add "Godina"
    k.Add "DatumUtovara"
    k.Add "KupacID"
    k.Add "Fakturisano"
    k.Add "FakturaID"
    k.Add "Napomena"
    k.Add "Stornirano"
    k.Add "Prevoznik"
    k.Add "Vozac"
    k.Add "Registracija"
    k.Add "Plomba"
    k.Add "TemperaturniRezim"
    k.Add "MestoIstovara"
    k.Add "VremeUtovara"
    k.Add "BrojNarudzbenice"
    RegistrujTabelu reg, TBL_UTOVAR, "Utovar", k
End Sub

Private Sub SpecUtovarStavke(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "UtovarStavkaID"
    k.Add "UtovarID"
    k.Add "PreradaID"
    k.Add "BrojPrerade"
    k.Add "KolicinaKg"
    k.Add "Stornirano"
    k.Add "BrojKutija"
    k.Add "BrojKesa"
    k.Add "CenaKg"
    RegistrujTabelu reg, TBL_UTOVAR_STAVKE, "UtovarStavke", k
End Sub

Private Sub SpecVozaci(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "VozacID"
    k.Add "Ime"
    k.Add "Prezime"
    k.Add "Telefon"
    k.Add "Aktivan"
    k.Add "PIN"
    k.Add "KapacitetKG"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    RegistrujTabelu reg, TBL_VOZACI, "Vozaci", k
End Sub

Private Sub SpecVrstaGotovihProizvoda(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipGotovogProizvoda"
    k.Add "Aktivan"
    k.Add "RokMeseci"
    RegistrujTabelu reg, TBL_VRSTA_GP, "VrstaGotProizvoda", k
End Sub

Private Sub SpecZbirna(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ZbirnaID"
    k.Add "Datum"
    k.Add "VozacID"
    k.Add "BrojZbirne"
    k.Add "KupacID"
    k.Add "Hladnjaca"
    k.Add "Pogon"
    k.Add "VrstaVoca"
    k.Add "SortaVoca"
    k.Add "UkupnoKolicina"
    k.Add "TipAmbalaze"
    k.Add "UkupnoAmbalaze"
    k.Add "Klasa"
    k.Add "Stornirano"
    k.Add "ClientRecordID"
    k.Add "SyncSource"
    k.Add "CreatedAt"
    k.Add "CreatedBy"
    k.Add "ModifiedAt"
    k.Add "ModifiedBy"
    k.Add "IspravkaOd"
    k.Add "ZamenjenSa"
    k.Add "CorrectionID"
    k.Add "IzdatoStatus"
    k.Add "GeneracijaID"
    RegistrujTabelu reg, TBL_ZBIRNA, "Zbirna", k
End Sub

