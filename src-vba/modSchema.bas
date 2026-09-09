Attribute VB_Name = "modSchema"

Option Explicit

'=====================================================================
' modSchema -- REGISTAR SEME TABELA
'
' Od ovog modula je izvor istine za STRUKTURU tabela KOD, ne sveska.
' Do sada je vazilo obrnuto: spiskovi kolona osnovnih tabela ziveli su
' iskljucivo u .xlsm, pa se prazna sveska nije mogla rekonstruisati, a
' obrisana kolona se videla tek kao pad upisa satima kasnije.
'
' GENERISAN FAJL -- ne menjaj rukom:
'     python tools/dump_schema.py <sveska> --json <put.json>
'     python tools/gen_schema_module.py --json <put.json>
'
' Javni ulazi:
'   EnsureAllTables      kreira sto fali, dopunjava kolone. Idempotentno.
'   VerifySchema         prijavljuje odstupanja; NISTA ne menja.
'   SchemaReadyOrFail    tvrda kapija pred upis, nad zadatim tabelama.
'   SchemaRegistry       registar, za testove i alate.
'
' NE zvati unutar otvorene transakcije: clsTransaction.RestoreTable pada
' na neslaganje broja kolona, pa bi promena seme onemogucila rollback.
'
' EnsureAllTables se NE zove sa svakog starta -- prolaz kroz sve tabele i
' sve kolone je skup. Zove se iz setup-a i iz testova. Na startu ide samo
' VerifySchema (citanje), kroz health check.
'=====================================================================

' Vrste odstupanja koje VerifySchema vraca (prefiks stavke u koleciji).
Public Const SCHEMA_DRIFT_TABELA As String = "TABELA"
Public Const SCHEMA_DRIFT_KOLONA As String = "KOLONA"

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

' Kolone jedne tabele, redosledom iz registra. Prazna kolekcija = tabela
' nije u registru (pozivalac razlikuje "nema je" od "nema kolona").
Public Function SchemaTableColumns(ByVal tblName As String) As Collection
    Dim reg As Object
    Set reg = SchemaRegistry()
    If reg.Exists(tblName) Then
        Set SchemaTableColumns = reg(tblName)("kolone")
    Else
        Set SchemaTableColumns = New Collection
    End If
End Function

Public Function SchemaTableSheet(ByVal tblName As String) As String
    Dim reg As Object
    Set reg = SchemaRegistry()
    If reg.Exists(tblName) Then SchemaTableSheet = CStr(reg(tblName)("sheet"))
End Function

' Kreira tabele kojih nema i dopunjava kolone koje fale. Idempotentno.
' Pad JEDNE tabele se zapise i NE zaustavlja ostale -- isti razlog kao
' EnsureKolonaSaTragom u modSetup: blanket "On Error Resume Next" je cutke
' preskakao ostatak posla, pa se posledica videla tek kao pad upisa.
Public Sub EnsureAllTables()
    Dim reg As Object
    Dim tblName As Variant

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        EnsureJednuTabelu CStr(tblName), reg(tblName)
    Next tblName
End Sub

' Odstupanja sveske od registra. NISTA ne menja -- dijagnostika ne sme da
' zameni gresku koju opisuje.
' Stavka je "VRSTA|tabela|kolona"; kolona je prazna kad fali cela tabela.
Public Function VerifySchema() As Collection
    Dim out As Collection
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim imena As Object
    Dim lc As ListColumn
    Dim kolone As Collection
    Dim i As Long

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

            Set kolone = reg(tblName)("kolone")
            For i = 1 To kolone.count
                If Not imena.Exists(CStr(kolone(i))) Then
                    out.Add SCHEMA_DRIFT_KOLONA & "|" & CStr(tblName) & "|" & _
                            CStr(kolone(i))
                End If
            Next i
        End If
    Next tblName

    Set VerifySchema = out
End Function

' Tvrda kapija pred upis. tblList je "tblA|tblB|tblC" -- proverava se SAMO
' to, jer pun prolaz po svakom upisu je preskup.
'
' Postoji zato sto nedostajuca tabela inace pukne tek na AppendRow-u, sa
' porukom koja o uzroku ne kaze nista. Sa novim modelom dokumenta to vise
' nije degradacija nego "model dokumenta ne postoji".
Public Sub SchemaReadyOrFail(ByVal sourceName As String, ByVal tblList As String)
    Dim reg As Object
    Dim delovi() As String
    Dim tblName As String
    Dim lo As ListObject
    Dim kolone As Collection
    Dim imena As Object
    Dim lc As ListColumn
    Dim i As Long
    Dim j As Long

    Set reg = SchemaRegistry()
    delovi = Split(tblList, "|")

    For i = LBound(delovi) To UBound(delovi)
        tblName = Trim$(delovi(i))
        If Len(tblName) > 0 Then
            If Not reg.Exists(tblName) Then
                Err.Raise vbObjectError + 9403, sourceName, _
                          "Tabela '" & tblName & "' nije u registru seme " & _
                          "(modSchema). Registar je izvor istine -- dopuni ga."
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

            Set imena = CreateObject("Scripting.Dictionary")
            imena.CompareMode = vbTextCompare
            For Each lc In lo.ListColumns
                If Not imena.Exists(lc.name) Then imena.Add lc.name, True
            Next lc

            Set kolone = reg(tblName)("kolone")
            For j = 1 To kolone.count
                If Not imena.Exists(CStr(kolone(j))) Then
                    Err.Raise vbObjectError + 9405, sourceName, _
                              "Tabeli '" & tblName & "' fali kolona '" & _
                              CStr(kolone(j)) & "'. Pokreni " & _
                              "modSchema.EnsureAllTables pa ponovi."
                End If
            Next j
        End If
    Next i
End Sub


'=====================================================================
' INTERNO
'=====================================================================

Private Sub EnsureJednuTabelu(ByVal tblName As String, ByVal spec As Object)
    Dim kolone As Collection
    Dim arr() As String
    Dim i As Long

    On Error GoTo EH

    Set kolone = spec("kolone")
    If kolone.count = 0 Then
        Err.Raise vbObjectError + 9406, "modSchema.EnsureJednuTabelu", _
                  "Registar nema nijednu kolonu za '" & tblName & "'."
    End If

    ReDim arr(1 To kolone.count)
    For i = 1 To kolone.count
        arr(i) = CStr(kolone(i))
    Next i

    modSetup.EnsureDataTable tblName, CStr(spec("sheet")), arr
    Exit Sub

EH:
    LogError "modSchema.EnsureAllTables", _
             "Tabela '" & tblName & "' nije obezbedjena: " & Err.description, _
             Err.Number
End Sub

Private Sub Reg(ByVal reg As Object, ByVal tblName As String, _
                ByVal sheetName As String, ByVal kolone As Collection)
    Dim d As Object

    If reg.Exists(tblName) Then
        Err.Raise vbObjectError + 9401, "modSchema.Reg", _
                  "Dupla tabela u registru: " & tblName
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("sheet") = sheetName
    Set d("kolone") = kolone
    reg.Add tblName, d
End Sub


'=====================================================================
' REGISTAR -- GENERISANO, ne menjaj rukom
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
    Reg reg, TBL_AMBALAZA, "Ambalaza", k
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
    Reg reg, TBL_ARTIKLI, "Artikli", k
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
    Reg reg, TBL_BANKA_IMPORT, "BankaImport", k
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
    Reg reg, TBL_CENOVNIK, "Cenovnik", k
End Sub

Private Sub SpecConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Vrednost"
    k.Add "Opis"
    Reg reg, TBL_CONFIG, "Config", k
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
    Reg reg, TBL_FAKTURA_STAVKE, "FakturaStavke", k
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
    Reg reg, TBL_FAKTURE, "Fakture", k
End Sub

Private Sub SpecKese(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipKese"
    k.Add "TezinaKg"
    k.Add "Aktivan"
    Reg reg, TBL_KESE, "Kese", k
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
    Reg reg, TBL_KOOPERANTI, "Kooperanti", k
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
    Reg reg, TBL_KORISNICI, "Korisnici", k
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
    Reg reg, TBL_KULTURE, "Kulture", k
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
    Reg reg, TBL_KUPCI, "Kupci", k
End Sub

Private Sub SpecKutije(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipKutije"
    k.Add "TezinaKg"
    k.Add "Aktivan"
    Reg reg, TBL_KUTIJE, "Kutije", k
End Sub

Private Sub SpecLocalConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Vrednost"
    k.Add "Opis"
    Reg reg, TBL_LOCAL_CONFIG, "LocalConfig", k
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
    Reg reg, TBL_MAGACIN, "Magacin", k
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
    Reg reg, TBL_MGMT, "MGMT", k
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
    Reg reg, TBL_NOVAC, "Novac", k
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
    Reg reg, TBL_OTKUP, "Otkup", k
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
    Reg reg, TBL_OTPREMNICA, "Otpremnica", k
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
    Reg reg, TBL_PALETA, "Paleta", k
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
    Reg reg, TBL_PALETA_STAVKA, "PaletaStavka", k
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
    Reg reg, TBL_PARCELE, "Parcele", k
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
    Reg reg, TBL_PARTNER_MAP, "PartnerMap", k
End Sub

Private Sub SpecPoruke(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "Kljuc"
    k.Add "Tekst"
    Reg reg, TBL_PORUKE, "sPoruke", k
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
    Reg reg, TBL_PRERADA, "Prerada", k
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
    Reg reg, TBL_PRERADA_STAVKA, "PreradaStavke", k
End Sub

Private Sub SpecPrevoznici(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "PrevoznikID"
    k.Add "Naziv"
    k.Add "Vozac"
    k.Add "Registracija"
    k.Add "Aktivan"
    Reg reg, TBL_PREVOZNICI, "Prevoznici", k
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
    Reg reg, TBL_PRIJEMNICA, "Prijemnica", k
End Sub

Private Sub SpecSEFConfig(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "ConfigKey"
    k.Add "ConfigValue"
    k.Add "Opis"
    k.Add "Aktivan"
    Reg reg, TBL_SEF_CONFIG, "SEFConfig", k
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
    Reg reg, TBL_SEF_EVENT_LOG, "SEFEventLog", k
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
    Reg reg, TBL_SEF_SUBMISSION, "SEFSubmission", k
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
    Reg reg, TBL_STANICE, "Otkupna Mesta", k
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
    Reg reg, TBL_STORNO_VEZE, "StornoVeze", k
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
    Reg reg, TBL_STORNO_ZURNAL, "StornoZurnal", k
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
    Reg reg, TBL_TIP_AMBALAZE, "TipAmbalaze", k
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
    Reg reg, TBL_TIP_PALETE, "TipPalete", k
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
    Reg reg, TBL_UTOVAR, "Utovar", k
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
    Reg reg, TBL_UTOVAR_STAVKE, "UtovarStavke", k
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
    Reg reg, TBL_VOZACI, "Vozaci", k
End Sub

Private Sub SpecVrstaGotovihProizvoda(ByVal reg As Object)
    Dim k As Collection
    Set k = New Collection
    k.Add "TipGotovogProizvoda"
    k.Add "Aktivan"
    k.Add "RokMeseci"
    Reg reg, TBL_VRSTA_GP, "VrstaGotProizvoda", k
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
    Reg reg, TBL_ZBIRNA, "Zbirna", k
End Sub

