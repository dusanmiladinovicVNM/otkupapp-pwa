Attribute VB_Name = "modOtkup"
Option Explicit

' ============================================================
' modOtkup - Aufkauf-Geschaeftslogik
' Kernmodul: Erfassung Lieferant zu Station
' ============================================================

' ============================================================
' OTKUP -- header + stavke (skela)
' ============================================================
'
' Jedan javni ulaz koji vraca JEDAN OtkupID. Obrazac je CreateZbirna_TX: _TX
' drzi transakciju i monitoring, Private core radi posao, kompletna
' prevalidacija PRE ijednog upisa.
'
' RAZLIKA U ODNOSU NA ZBIRNU: otkup je PRIMARNA CINJENICA.
'
' Zbirna i otpremnica su izvedeni dokumenti, pa njihovi writeri primaju IZVORE i
' racunaju stavke. Otkup nema izvor ispod sebe -- kolicine nastaju neposrednim
' unosom. Zato prima stavke, i zato nema ni "ocekivano" ni drugi ulaz.
'
' Sta se menja u odnosu na SaveOtkupMulti_TX:
'
'   staro:  dva reda u tblOtkup (Klasa I i Klasa II), dva ID-a, pa string
'           "OTK-1 + OTK-2" koji pozivalac posle parsira na devet mesta
'   novo:   JEDAN header + N stavki, jedan ID
'
' KulturaID SE PRIMA, NE RAZRESAVA.
'
' Zatecen kod ga fabrikuje na dva mesta -- modOtkup.bas:556 i
' modMasterSync.bas:1959 -- tako sto trazi samo po VrstaVoca, a kad ne nadje
' sklopi "vrsta-sorta" string koji izgleda kao FK a ne pokazuje ni na sta.
' Razresavanje (Vrsta, Sorta) -> KulturaID je posao ADAPTERA; writer proverava
' da FK postoji i da se snapshot vrsta/sorta slaze sa tom kulturom.
'
' Skela je ADITIVNA: produkcioni pozivaoci i dalje idu starim putem
' (modOtkupUnos.bas:275), a golden scenariji to dokazuju nepromenjeni. Zato
' header NAMERNO ostavlja Kolicina / Cena / Klasa / KolAmbalaze / BrutoKg /
' VozacID / Isplaceno / DatumIsplate / VremeUnosa prazne -- to su kolone koje u
' ciljnoj semi ne postoje (DOCUMENT_HEADER_LINES S4.1).
'
' NIJE u skeli: tblAmbalaza (ide u Otkup cutover) i tblNovac (kes ne ulazi kroz
' otkupni list, S4.1b).
'
' Header (h) -- Scripting.Dictionary, obavezni kljucevi:
'   Datum, KooperantID, StanicaID, KulturaID, VrstaVoca, SortaVoca,
'   TipAmbalaze, BrojDokumenta
' opcioni:
'   ParcelaID, KolAmbIzdata, ClientRecordID, SyncSource, SourceCreatedAt
'
' Stavke -- Collection diktova, svaki:
'   Klasa (I ili II), Kolicina (> 0, NETO), Cena (> 0), KolAmbalaze (>= 0),
'   BrutoKg (opciono; > 0 samo kad je unos bio bruto)
Public Function CreateOtkup_TX(ByVal h As Object, _
                               ByVal stavke As Collection, _
                               Optional ByRef outGreska As String) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO. Ide PRE BeginTx -- kapija sme da
    ' digne gresku, a nema smisla otvarati transakciju koja se odmah rollback-uje.
    modSchema.SchemaReadyOrFail "CreateOtkup_TX", _
        TBL_OTKUP & "|" & TBL_OTKUP_STAVKE

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE

    CreateOtkup_TX = CreateOtkup(h, stavke)

    If CreateOtkup_TX = "" Then
        Err.Raise vbObjectError + 1860, "CreateOtkup_TX", _
                  "CreateOtkup nije vratio OtkupID."
    End If

    tx.CommitTx

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "CreateOtkup_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modOtkup", _
        procedureName:="CreateOtkup_TX", _
        entityType:="Otkup", _
        entityID:=CreateOtkup_TX, _
        correlationId:=CreateOtkup_TX, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="CreateOtkup_TX failed. Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="CreateOtkup_TX", _
        entityType:="Otkup", _
        entityID:=CreateOtkup_TX, _
        correlationId:=CreateOtkup_TX

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateOtkup_TX = ""
    outGreska = errDesc

    PrintOtkupTxFailure "CreateOtkup_TX", errSrc, errNum, errDesc
End Function

' Core -- NE zovi spolja. Jedini ulaz je CreateOtkup_TX, koji drzi snapshot
' transakciju; direktan poziv bi kod greske ostavio header bez stavki.
Private Function CreateOtkup(ByVal h As Object, _
                             ByVal stavke As Collection) As String
    Const SRC As String = "CreateOtkup"

    On Error GoTo EH

    If h Is Nothing Then
        Err.Raise vbObjectError + 1861, SRC, "Header nije prosledjen."
    End If

    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1862, SRC, "Stavke nisu prosledjene."
    End If

    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1863, SRC, _
                  "Otkup mora imati bar jednu stavku."
    End If

    ' Fail-fast nad semom pre ijednog upisa.
    RequireColumnIndex TBL_OTKUP, COL_OTK_ID, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_DATUM, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_KOOPERANT, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_STANICA, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_KULTURA, SRC
    RequireColumnIndex TBL_OTKUP, COL_OTK_BR_DOK, SRC

    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_ID, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_RB, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KLASA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_CENA, SRC
    RequireColumnIndex TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, SRC

    OtkHdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim kooperantID As String, stanicaID As String, kulturaID As String
    Dim vrstaVoca As String, sortaVoca As String, tipAmb As String
    Dim brDok As String, parcelaID As String

    datum = OtkHdrDatum(h, "Datum", SRC)
    kooperantID = OtkHdrObavezan(h, "KooperantID", SRC)
    stanicaID = OtkHdrObavezan(h, "StanicaID", SRC)
    kulturaID = OtkHdrObavezan(h, "KulturaID", SRC)
    vrstaVoca = OtkHdrObavezan(h, "VrstaVoca", SRC)
    sortaVoca = OtkHdrObavezan(h, "SortaVoca", SRC)
    tipAmb = OtkHdrObavezan(h, "TipAmbalaze", SRC)
    brDok = OtkHdrObavezan(h, "BrojDokumenta", SRC)
    parcelaID = OtkHdrOpcion(h, "ParcelaID")

    RequireKulturaSeSlaze kulturaID, vrstaVoca, sortaVoca, SRC
    RequireParcelaKooperanta parcelaID, kooperantID, SRC

    ' Prevalidacija SVIH stavki pre bilo kog upisa: dokument sa dve stavke od
    ' kojih druga ne valja ne sme da ostavi prvu u tabeli.
    Dim vidjeneKlase As Object
    Set vidjeneKlase = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim s As Object
    Dim klasa As String
    Dim kolicina As Double, cena As Double, kolAmb As Double, bruto As Double
    Dim imaAmbalaze As Boolean

    For i = 1 To stavke.count
        If Not IsObject(stavke(i)) Then
            Err.Raise vbObjectError + 1864, SRC, _
                      "Stavka " & CStr(i) & " nije Dictionary."
        End If

        Set s = stavke(i)

        klasa = Trim$(NzToText(OtkStavkaVrednost(s, "Klasa", i, SRC)))
        RequireValidOtkupClass klasa, SRC

        ' Dokument ima najvise jednu stavku po klasi -- dve iste klase su bas
        ' bug koji header+stavke uklanja: jedan logicki dokument rasut po redovima.
        If vidjeneKlase.Exists(UCase$(klasa)) Then
            Err.Raise vbObjectError + 1865, SRC, _
                      "Dve stavke iste klase: " & klasa
        End If
        vidjeneKlase.Add UCase$(klasa), True

        kolicina = OtkStavkaBroj(s, "Kolicina", i, SRC)
        If kolicina <= 0 Then
            Err.Raise vbObjectError + 1866, SRC, _
                      "Kolicina mora biti veca od nule. Stavka " & CStr(i) & _
                      ", klasa " & klasa & "."
        End If

        ' Cenovnik je PREDLOG; writer trazi samo da cena postoji. Override je
        ' legitiman -- sacuvana cena je istorijska cinjenica dokumenta.
        cena = OtkStavkaBroj(s, "Cena", i, SRC)
        If cena <= 0 Then
            Err.Raise vbObjectError + 1867, SRC, _
                      "Cena mora biti veca od nule. Stavka " & CStr(i) & _
                      ", klasa " & klasa & "."
        End If

        kolAmb = OtkStavkaBroj(s, "KolAmbalaze", i, SRC)
        If kolAmb < 0 Then
            Err.Raise vbObjectError + 1868, SRC, _
                      "Kolicina ambalaze ne sme biti negativna. Stavka " & _
                      CStr(i) & ", klasa " & klasa & "."
        End If
        RequireCeoBrojOtk kolAmb, "Ambalaza na stavci " & CStr(i), SRC
        If kolAmb > 0 Then imaAmbalaze = True

        ' Bruto se cuva SAMO kad je unos bio bruto. Kolicina je uvek neto, pa
        ' bruto koji je manji od nje znaci zamenjene vrednosti, ne rubni slucaj.
        bruto = OtkStavkaBrojOpcion(s, "BrutoKg", i, SRC)
        If bruto < 0 Then
            Err.Raise vbObjectError + 1869, SRC, _
                      "BrutoKg ne sme biti negativan. Stavka " & CStr(i) & "."
        End If
        If bruto > 0 And bruto < kolicina Then
            Err.Raise vbObjectError + 1870, SRC, _
                      "BrutoKg (" & Fmt2Otk(bruto) & ") je manji od neto kolicine (" & _
                      Fmt2Otk(kolicina) & "). Stavka " & CStr(i) & "."
        End If
    Next i

    If imaAmbalaze And Len(tipAmb) = 0 Then
        Err.Raise vbObjectError + 1871, SRC, _
                  "Tip ambalaze je obavezan kada postoji ambalaza."
    End If

    Dim kolAmbIzdata As Double
    kolAmbIzdata = OtkHdrBrojOpcion(h, "KolAmbIzdata", SRC)
    If kolAmbIzdata < 0 Then
        Err.Raise vbObjectError + 1872, SRC, _
                  "Izdata ambalaza ne sme biti negativna."
    End If
    RequireCeoBrojOtk kolAmbIzdata, "Izdata ambalaza", SRC

    ' --- upis ----------------------------------------------------------------
    Dim otkupID As String
    otkupID = NewEntityID("OTK-")

    If otkupID = "" Then
        Err.Raise vbObjectError + 1873, SRC, _
                  "NewEntityID nije vratio OtkupID."
    End If

    Dim rowData As Variant
    rowData = BuildOtkupHeaderRowData(otkupID, datum, kooperantID, stanicaID, _
                                      kulturaID, vrstaVoca, sortaVoca, tipAmb, _
                                      brDok, parcelaID, kolAmbIzdata, _
                                      OtkHdrOpcion(h, "ClientRecordID"), _
                                      OtkHdrOpcion(h, "SyncSource"), _
                                      OtkHdrOpcion(h, "SourceCreatedAt"))

    If AppendRow(TBL_OTKUP, rowData) <= 0 Then
        Err.Raise vbObjectError + 1874, SRC, _
                  "AppendRow nije upisao header u tblOtkup."
    End If

    Dim stavkaID As String
    For i = 1 To stavke.count
        Set s = stavke(i)

        ' Fail-closed: NewEntityID vraca "" kad CoCreateGuid ne uspe. Red bez
        ' identiteta je gori od pada -- niko ga posle ne moze ni naci ni vezati.
        stavkaID = NewEntityID("OKS-")
        If stavkaID = "" Then
            Err.Raise vbObjectError + 1875, SRC, _
                      "NewEntityID nije vratio OtkupStavkaID za stavku " & CStr(i) & "."
        End If

        rowData = BuildOtkupStavkaRowData(stavkaID, otkupID, i, _
                    Trim$(NzToText(s("Klasa"))), _
                    OtkStavkaBroj(s, "Kolicina", i, SRC), _
                    OtkStavkaBroj(s, "Cena", i, SRC), _
                    OtkStavkaBroj(s, "KolAmbalaze", i, SRC), _
                    OtkStavkaBrojOpcion(s, "BrutoKg", i, SRC))

        If AppendRow(TBL_OTKUP_STAVKE, rowData) <= 0 Then
            Err.Raise vbObjectError + 1876, SRC, _
                      "AppendRow nije upisao stavku " & CStr(i) & "."
        End If
    Next i

    CreateOtkup = otkupID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError SRC, errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

' KulturaID se NE razresava ovde -- proverava se.
'
' Mora postojati tacno jednom, i vrsta/sorta koje dokument nosi kao snapshot
' moraju odgovarati toj kulturi. Time fabrikovan "vrsta-sorta" string pada odmah:
' takvog reda u tblKulture nema.
Private Sub RequireKulturaSeSlaze(ByVal kulturaID As String, _
                                  ByVal vrstaVoca As String, _
                                  ByVal sortaVoca As String, _
                                  ByVal src As String)
    ' FindRows uvek vraca Collection (svaki izlaz radi Set) -- provera
    ' "Is Nothing" bi bila mrtav kod koji samo izgleda kao paznja.
    Dim redovi As Collection
    Set redovi = FindRows(TBL_KULTURE, COL_KUL_ID, kulturaID)

    If redovi.count = 0 Then
        Err.Raise vbObjectError + 1877, src, _
                  "KulturaID ne postoji: " & kulturaID
    End If
    If redovi.count > 1 Then
        Err.Raise vbObjectError + 1878, src, _
                  "KulturaID nije jednoznacan: " & kulturaID & _
                  "; Count=" & CStr(redovi.count)
    End If

    Dim kVrsta As String, kSorta As String
    kVrsta = Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_VRSTA)))
    kSorta = Trim$(NzToText(LookupValue(TBL_KULTURE, COL_KUL_ID, kulturaID, COL_KUL_SORTA)))

    If StrComp(kVrsta, vrstaVoca, vbTextCompare) <> 0 Or _
       StrComp(kSorta, sortaVoca, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1879, src, _
                  "Vrsta/sorta se ne slazu sa kulturom " & kulturaID & _
                  ": dokument nosi '" & vrstaVoca & "/" & sortaVoca & _
                  "', kultura je '" & kVrsta & "/" & kSorta & "'."
    End If
End Sub

' Parcela mora pripadati kooperantu dokumenta. Tudja parcela ne prolazi
' kanonski writer (DOCUMENT_HEADER_LINES S4.1f).
'
' Neslaganje KULTURE parcele ostaje stvar ekrana (warning uz override) -- za
' tvrdo pravilo tu nema dovoljno osnova.
Private Sub RequireParcelaKooperanta(ByVal parcelaID As String, _
                                     ByVal kooperantID As String, _
                                     ByVal src As String)
    If Len(parcelaID) = 0 Then Exit Sub

    Dim redovi As Collection
    Set redovi = FindRows(TBL_PARCELE, COL_PAR_ID, parcelaID)

    If redovi.count = 0 Then
        Err.Raise vbObjectError + 1880, src, "ParcelaID ne postoji: " & parcelaID
    End If
    If redovi.count > 1 Then
        Err.Raise vbObjectError + 1881, src, _
                  "ParcelaID nije jednoznacan: " & parcelaID
    End If

    Dim vlasnik As String
    vlasnik = Trim$(NzToText(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KOOP)))

    If StrComp(vlasnik, kooperantID, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1882, src, _
                  "Parcela " & parcelaID & " pripada kooperantu " & vlasnik & _
                  ", a otkup je za " & kooperantID & "."
    End If
End Sub

' Header ciljne seme. Kolone koje u ciljnom modelu ne postoje ostaju PRAZNE --
' Kolicina, Cena, Klasa, KolAmbalaze, BrutoKg (stavka), VozacID (otpremnica),
' Isplaceno / DatumIsplate (read-model), VremeUnosa (CreatedAt/SourceCreatedAt),
' Novac / PrimalacNovca, veze po broju i GeneracijaID.
Private Function BuildOtkupHeaderRowData(ByVal otkupID As String, _
                                         ByVal datum As Date, _
                                         ByVal kooperantID As String, _
                                         ByVal stanicaID As String, _
                                         ByVal kulturaID As String, _
                                         ByVal vrstaVoca As String, _
                                         ByVal sortaVoca As String, _
                                         ByVal tipAmb As String, _
                                         ByVal brDok As String, _
                                         ByVal parcelaID As String, _
                                         ByVal kolAmbIzdata As Double, _
                                         ByVal clientRecordID As String, _
                                         ByVal syncSource As String, _
                                         ByVal sourceCreatedAt As String) As Variant
    Const SRC As String = "BuildOtkupHeaderRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTKUP)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1883, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtkup."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_ID, otkupID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KOOPERANT, kooperantID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_STANICA, stanicaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KULTURA, kulturaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_BR_DOK, brDok, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_PARCELA, parcelaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA, kolAmbIzdata, SRC
    SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_STORNIRANO, "", SRC

    OtkPostaviAkoPostoji rowData, COL_OTK_CLIENT_RECORD_ID, clientRecordID, SRC
    OtkPostaviAkoPostoji rowData, COL_OTK_SYNC_SOURCE, syncSource, SRC
    OtkPostaviAkoPostoji rowData, COL_OTK_SOURCE_CREATED_AT, sourceCreatedAt, SRC

    ' IZDATO se pise EKSPLICITNO -- nov model se ne oslanja na legacy konvenciju
    ' "prazno = IZDATO". Otkup nema persistentan DRAFT: forma je njegov draft, a
    ' dokument nastaje vec izdat (DOCUMENT_HEADER_LINES S4.1e).
    OtkPostaviAkoPostoji rowData, COL_TRACE_IZDATO_STATUS, IZDATO_IZDATO, SRC

    BuildOtkupHeaderRowData = rowData
End Function

Private Function BuildOtkupStavkaRowData(ByVal stavkaID As String, _
                                         ByVal otkupID As String, _
                                         ByVal redniBroj As Long, _
                                         ByVal klasa As String, _
                                         ByVal kolicina As Double, _
                                         ByVal cena As Double, _
                                         ByVal kolAmb As Double, _
                                         ByVal bruto As Double) As Variant
    Const SRC As String = "BuildOtkupStavkaRowData"

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_OTKUP_STAVKE)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1884, SRC, _
                  "Ne mogu da odredim broj kolona za tblOtkupStavke."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_ID, stavkaID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkupID, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_RB, redniBroj, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_CENA, cena, SRC
    SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_KOL_AMB, kolAmb, SRC

    ' BrutoKg ostaje PRAZAN kad je unos bio neto -- prazno je podatak, ne nula.
    If bruto > 0 Then
        SetRowValueByColumn rowData, TBL_OTKUP_STAVKE, COL_OKS_BRUTO, bruto, SRC
    End If

    BuildOtkupStavkaRowData = rowData
End Function

' Kolona koja u zatecenoj svesci mozda jos ne postoji ne sme da obori upis.
Private Sub OtkPostaviAkoPostoji(ByRef rowData() As Variant, _
                                 ByVal columnName As String, _
                                 ByVal value As Variant, _
                                 ByVal src As String)
    If GetColumnIndex(TBL_OTKUP, columnName) > 0 Then
        SetRowValueByColumn rowData, TBL_OTKUP, columnName, value, src
    End If
End Sub

Private Sub RequireCeoBrojOtk(ByVal v As Double, ByVal opis As String, _
                              ByVal src As String)
    If Abs(v - Fix(v)) > 0.0000001 Then
        Err.Raise vbObjectError + 1886, src, _
                  opis & " mora biti ceo broj, a nije: " & Fmt2Otk(v)
    End If
End Sub

' Broj u poruku, nezavisno od Windows locale-a (decimalna tacka uvek).
Private Function Fmt2Otk(ByVal v As Double) As String
    Fmt2Otk = Replace(Format$(v, "0.00"), ",", ".")
End Function

' --- citanje DTO-a ----------------------------------------------------------
'
' Nedostajuci kljuc je GRESKA, ne prazna vrednost: Dictionary(k) nad nepostojecim
' kljucem tiho vraca Empty i doda kljuc, pa bi tipfeler prosao kao "nije uneto".
'
' Sve sto nije na spisku je greska -- tipfeler u OPCIONOM polju je inace
' nevidljiv. VozacID / Isplaceno / Kolicina i drustvo NISU na spisku namerno:
' pozivalac koji ih salje radi po starom modelu i mora to da cuje.
Private Function OtkHdrKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "datum", "kooperantid", "stanicaid", "kulturaid", _
             "vrstavoca", "sortavoca", "tipambalaze", "brojdokumenta", _
             "parcelaid", "kolambizdata", _
             "clientrecordid", "syncsource", "sourcecreatedat"
            OtkHdrKljucPoznat = True
    End Select
End Function

Private Sub OtkHdrProveriKljuceve(ByVal h As Object, ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In h.Keys
        If Not OtkHdrKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1887, src, _
                      "Header ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". Kolicina/Cena/Klasa/KolAmbalaze/BrutoKg idu na STAVKU, " & _
                      "VozacID na otpremnicu, Isplaceno je izvedeno."
        End If
    Next kljuc
End Sub

Private Function OtkHdrObavezan(ByVal h As Object, ByVal kljuc As String, _
                                ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1888, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    OtkHdrObavezan = Trim$(NzToText(h(kljuc)))

    If Len(OtkHdrObavezan) = 0 Then
        Err.Raise vbObjectError + 1889, src, _
                  "Header polje je prazno: " & kljuc
    End If
End Function

Private Function OtkHdrOpcion(ByVal h As Object, ByVal kljuc As String) As String
    If h.Exists(kljuc) Then OtkHdrOpcion = Trim$(NzToText(h(kljuc)))
End Function

Private Function OtkHdrBrojOpcion(ByVal h As Object, ByVal kljuc As String, _
                                  ByVal src As String) As Double
    If Not h.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = h(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1890, src, _
                  "Header polje " & kljuc & " nije broj: " & NzToText(v)
    End If

    OtkHdrBrojOpcion = CDbl(v)
End Function

Private Function OtkHdrDatum(ByVal h As Object, ByVal kljuc As String, _
                             ByVal src As String) As Date
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1891, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    If Not IsDate(h(kljuc)) Then
        Err.Raise vbObjectError + 1892, src, _
                  "Header polje nije datum: " & kljuc
    End If

    OtkHdrDatum = CDate(h(kljuc))
End Function

Private Function OtkStavkaVrednost(ByVal s As Object, ByVal kljuc As String, _
                                   ByVal idx As Long, ByVal src As String) As Variant
    If Not s.Exists(kljuc) Then
        Err.Raise vbObjectError + 1893, src, _
                  "Stavka " & CStr(idx) & " nema kljuc: " & kljuc
    End If

    OtkStavkaVrednost = s(kljuc)
End Function

Private Function OtkStavkaBroj(ByVal s As Object, ByVal kljuc As String, _
                               ByVal idx As Long, ByVal src As String) As Double
    Dim v As Variant
    v = OtkStavkaVrednost(s, kljuc, idx, src)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1894, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    OtkStavkaBroj = CDbl(v)
End Function

' BrutoKg je jedino polje stavke koje sme da izostane -- neto unos ga nema.
Private Function OtkStavkaBrojOpcion(ByVal s As Object, ByVal kljuc As String, _
                                     ByVal idx As Long, ByVal src As String) As Double
    If Not s.Exists(kljuc) Then Exit Function

    Dim v As Variant
    v = s(kljuc)
    If IsEmpty(v) Then Exit Function
    If Len(Trim$(NzToText(v))) = 0 Then Exit Function

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1895, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    OtkStavkaBrojOpcion = CDbl(v)
End Function

Public Function SaveOtkup_TX(ByVal datum As Date, ByVal kooperantID As String, _
                              ByVal stanicaID As String, ByVal vrstaVoca As String, _
                              ByVal sortaVoca As String, ByVal kolicina As Double, _
                              ByVal cena As Double, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, ByVal vozacID As String, _
                              ByVal brDok As String, ByVal novac As Double, _
                              ByVal primalac As String, _
                              Optional ByVal klasa As String = "I", _
                              Optional ByVal parcelaID As String = "", _
                              Optional ByVal brojZbirne As String = "") As String

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. SaveOtkupMulti_TX).
    modSchema.SchemaReadyOrFail "SaveOtkup_TX", _
        TBL_OTKUP & "|" & TBL_AMBALAZA & "|" & TBL_NOVAC

tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtkup_TX = SaveOtkup(datum, kooperantID, stanicaID, vrstaVoca, _
                              sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                              vozacID, brDok, novac, primalac, klasa, _
                              parcelaID, brojZbirne)

    If SaveOtkup_TX = "" Then
        Err.Raise vbObjectError + 1801, "SaveOtkup_TX", _
                  "SaveOtkup fehlgeschlagen"
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="OTKUP_SAVE_SUCCESS", _
        severity:="INFO", _
        message:="Otkup saved. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; Vrsta=" & vrstaVoca & _
                 "; Koli" & ChrW(269) & "ina=" & CStr(kolicina), _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=SaveOtkup_TX
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

    LogErr "SaveOtkup_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=brDok, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="OTKUP_SAVE_FAIL", _
        severity:="ERROR", _
        message:="Otkup save failed. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; BrDok=" & brDok & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkup_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkup_TX, _
        correlationId:=brDok

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtkup_TX = ""

    PrintOtkupTxFailure "SaveOtkup_TX", errSrc, errNum, errDesc
End Function

Public Function SaveOtkupMulti_TX(ByVal datum As Date, _
                                   ByVal kooperantID As String, _
                                   ByVal stanicaID As String, _
                                   ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, _
                                   ByVal kolicinaI As Double, _
                                   ByVal cenaI As Double, _
                                   ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, _
                                   ByVal vozacID As String, _
                                   ByVal brDok As String, _
                                   ByVal novac As Double, _
                                   ByVal primalac As String, _
                                   ByVal parcelaID As String, _
                                   ByVal brojZbirne As String, _
                                   Optional ByVal hasKlasaII As Boolean = False, _
                                   Optional ByVal kolicinaII As Double = 0, _
                                   Optional ByVal cenaII As Double = 0, _
                                   Optional ByVal kolAmbIzdata As Long = 0, _
                                   Optional ByVal brutoKgI As Double = 0, _
                                   Optional ByVal kolAmbII As Long = 0, _
                                   Optional ByVal brutoKgII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO, pa tabela sa kolonom manje ili
    ' u pogresnom rasporedu tiho salje vrednosti u pogresna polja. To je gore od
    ' pada upisa -- greska nastaje u podacima, ne u logu.
    '
    ' Ide PRE BeginTx: kapija sme da digne gresku, a nema smisla otvarati
    ' transakciju koja se odmah rollback-uje.
    modSchema.SchemaReadyOrFail "SaveOtkupMulti_TX", _
        TBL_OTKUP & "|" & TBL_AMBALAZA & "|" & TBL_NOVAC

    If Trim$(kooperantID) = "" Then
        Err.Raise vbObjectError + 1810, "SaveOtkupMulti_TX", _
                  "KooperantID je obavezan."
    End If

    If Trim$(stanicaID) = "" Then
        Err.Raise vbObjectError + 1811, "SaveOtkupMulti_TX", _
                  "StanicaID je obavezan."
    End If

    ' Klasa I je opciona (kolicinaI = 0 -> unosi se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)

    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1812, "SaveOtkupMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    If hasKlasaI And cenaI <= 0 Then
        Err.Raise vbObjectError + 1812, "SaveOtkupMulti_TX", _
                  "Cena za Klasu I mora biti veca od nule."
    End If

    If hasKlasaII Then
        If kolicinaII <= 0 Or cenaII <= 0 Then
            Err.Raise vbObjectError + 1813, "SaveOtkupMulti_TX", _
                      "Koli" & ChrW(269) & "ina i cena za Klasu II moraju biti vece od nule."
        End If
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1814, "SaveOtkupMulti_TX", _
                  "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmbIzdata < 0 Then
        Err.Raise vbObjectError + 1841, "SaveOtkupMulti_TX", _
                  "Koli" & ChrW(269) & "ina izdate ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmbII < 0 Then
        Err.Raise vbObjectError + 1842, "SaveOtkupMulti_TX", _
                  "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e (Klasa II) ne sme biti negativna."
    End If

    If novac < 0 Then
        Err.Raise vbObjectError + 1815, "SaveOtkupMulti_TX", _
                  "Iznos novca ne sme biti negativan."
    End If

    If (kolAmb > 0 Or kolAmbII > 0) And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1816, "SaveOtkupMulti_TX", _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    Dim resultI As String
    If hasKlasaI Then
        resultI = SaveOtkup( _
            datum:=datum, _
            kooperantID:=kooperantID, _
            stanicaID:=stanicaID, _
            vrstaVoca:=vrstaVoca, _
            sortaVoca:=sortaVoca, _
            kolicina:=kolicinaI, _
            cena:=cenaI, _
            tipAmb:=tipAmb, _
            kolAmb:=kolAmb, _
            vozacID:=vozacID, _
            brDok:=brDok, _
            novac:=novac, _
            primalac:=primalac, _
            klasa:=KLASA_I, _
            parcelaID:=parcelaID, _
            brojZbirne:=brojZbirne, _
            kolAmbIzdata:=kolAmbIzdata, _
            brutoKg:=brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1817, "SaveOtkupMulti_TX", _
                      "SaveOtkup Klasa I fehlgeschlagen"
        End If
    End If

    Dim resultII As String

    ' Kes I izdata ambalaza se belezi na red Klase I; ako Klase I nema (samo II),
    ' belezi se na Klasu II -- inace bi se izgubili (SaveOtkup Klase I se preskace).
    Dim novacII As Double
    Dim kolAmbIzdataII As Long
    If Not hasKlasaI Then
        novacII = novac
        kolAmbIzdataII = kolAmbIzdata
    End If

    If hasKlasaII Then
        resultII = SaveOtkup( _
            datum:=datum, _
            kooperantID:=kooperantID, _
            stanicaID:=stanicaID, _
            vrstaVoca:=vrstaVoca, _
            sortaVoca:=sortaVoca, _
            kolicina:=kolicinaII, _
            cena:=cenaII, _
            tipAmb:=tipAmb, _
            kolAmb:=kolAmbII, _
            vozacID:=vozacID, _
            brDok:=brDok, _
            novac:=novacII, _
            primalac:=primalac, _
            klasa:=KLASA_II, _
            parcelaID:=parcelaID, _
            brojZbirne:=brojZbirne, _
            kolAmbIzdata:=kolAmbIzdataII, _
            brutoKg:=brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1818, "SaveOtkupMulti_TX", _
                      "SaveOtkup Klasa II fehlgeschlagen"
        End If
    End If

    ' Primarni OtkupID dokumenta (za kes/avans veze): Klasa I ako postoji, inace II.
    Dim primaryID As String
    If hasKlasaI Then primaryID = resultI Else primaryID = resultII

    If novac > 0 Then
        Dim koopNaziv As String
        koopNaziv = GetKooperantNazivForNovac(kooperantID)

        Dim novacID As String
        novacID = SaveNovac( _
            brojDok:=brDok, _
            datum:=datum, _
            partner:=koopNaziv, _
            partnerId:=kooperantID, _
            entitetTip:="Kooperant", _
            omID:=stanicaID, _
            kooperantID:=kooperantID, _
            fakturaID:="", _
            vrstaVoca:=vrstaVoca, _
            tip:=NOV_KES_OTKUPAC_KOOP, _
            uplata:=0, _
            isplata:=novac, _
            napomena:=primalac, _
            otkupID:=primaryID)

        If novacID = "" Then
            Err.Raise vbObjectError + 1819, "SaveOtkupMulti_TX", _
                      "SaveNovac fehlgeschlagen"
        End If
    End If

    If hasKlasaI Then ApplyAvansToOtkup kooperantID, resultI
    If hasKlasaII Then ApplyAvansToOtkup kooperantID, resultII

    tx.CommitTx
    Set tx = Nothing

    If hasKlasaI And hasKlasaII Then
        SaveOtkupMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SaveOtkupMulti_TX = resultI
    Else
        SaveOtkupMulti_TX = resultII
    End If

    On Error Resume Next
    Monitor_Event _
        eventType:="OTKUP_MULTI_SAVE_SUCCESS", _
        severity:="INFO", _
        message:="Otkup multi saved. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; Vrsta=" & vrstaVoca & _
                 "; ResultI=" & resultI & _
                 "; ResultII=" & resultII & _
                 "; HasKlasaII=" & CStr(hasKlasaII), _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkupMulti_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkupMulti_TX, _
        correlationId:=resultI
    On Error GoTo 0

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "SaveOtkupMulti_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkupMulti_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkupMulti_TX, _
        correlationId:=brDok, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="OTKUP_MULTI_SAVE_FAIL", _
        severity:="ERROR", _
        message:="Otkup multi save failed. KooperantID=" & kooperantID & _
                 "; StanicaID=" & stanicaID & _
                 "; BrDok=" & brDok & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modOtkup", _
        procedureName:="SaveOtkupMulti_TX", _
        entityType:="Otkup", _
        entityID:=SaveOtkupMulti_TX, _
        correlationId:=brDok

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtkupMulti_TX = ""

    PrintOtkupTxFailure "SaveOtkupMulti_TX", errSrc, errNum, errDesc
End Function

' ============================================================
' Kontrola proseka neto kg po gajbici (Kolicina / KolAmbalaze).
' Pragovi se citaju iz tblKulture po VrstaVoca (PragProsekUpoz/PragProsekBlok).
' Prazno / 0 -> provera se preskace (opt-in po kulturi; fail-safe i kad kolone
' jos ne postoje na klijentu -- LookupValue vrati Empty -> 0).
'   prosek > PragProsekBlok -> tvrda blokada (False, bez override-a)
'   prosek > PragProsekUpoz -> upozorenje (vbYesNo; False samo ako operater odustane)
' Poziva se iz frmOtkup.btnUnos_Click POSLE bruto->neto konverzije (Kolicina = neto).
' Vraca True kad je unos dozvoljen (ili potvrdjen), False kad treba prekinuti.
' ============================================================
Public Function OtkupProsekGajbiceOK(ByVal vrstaVoca As String, _
        ByVal kolicinaI As Double, ByVal kolAmbI As Long, _
        ByVal kolicinaII As Double, ByVal kolAmbII As Long) As Boolean

    OtkupProsekGajbiceOK = True
    On Error GoTo EH

    Dim pragUpoz As Double, pragBlok As Double
    pragUpoz = KulturaProsekPrag(vrstaVoca, COL_KUL_PRAG_PROSEK_UPOZ)
    pragBlok = KulturaProsekPrag(vrstaVoca, COL_KUL_PRAG_PROSEK_BLOK)

    ' Nijedan prag nije podesen za ovu kulturu -> nema provere.
    If pragUpoz <= 0 And pragBlok <= 0 Then Exit Function

    ' Najveci prosek po klasama (svaka klasa ima svoje gajbe).
    Dim maxProsek As Double, klasaLbl As String
    ProsekKlase kolicinaI, kolAmbI, "I", maxProsek, klasaLbl
    ProsekKlase kolicinaII, kolAmbII, "II", maxProsek, klasaLbl

    If maxProsek <= 0 Then Exit Function

    Dim poruka As String
    poruka = "Prosek po gajbici" & IIf(Len(klasaLbl) > 0, " (klasa " & klasaLbl & ")", "") & _
             " je " & Format$(maxProsek, "0.00") & " kg."

    ' Tvrda blokada.
    If pragBlok > 0 And maxProsek > pragBlok Then
        MsgBox poruka & vbCrLf & _
               "Dozvoljeni maksimum je " & Format$(pragBlok, "0.00") & " kg po gajbici." & vbCrLf & _
               "Unos je blokiran -- proverite neto kila" & ChrW(382) & "u i broj gajbi.", _
               vbCritical, APP_NAME
        OtkupProsekGajbiceOK = False
        Exit Function
    End If

    ' Upozorenje uz mogucnost nastavka.
    If pragUpoz > 0 And maxProsek > pragUpoz Then
        If MsgBox(poruka & vbCrLf & _
                  "Preporu" & ChrW(269) & "eni maksimum je " & Format$(pragUpoz, "0.00") & " kg po gajbici." & vbCrLf & _
                  "Da li ipak " & ChrW(382) & "elite da nastavite?", _
                  vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            OtkupProsekGajbiceOK = False
        End If
    End If
    Exit Function

EH:
    ' Fail-safe: greska u proveri ne sme da obori normalan unos.
    LogErr "modOtkup.OtkupProsekGajbiceOK"
    OtkupProsekGajbiceOK = True
End Function

' Prag proseka za kulturu (po VrstaVoca) iz tblKulture; Empty/ne-broj -> 0.
Private Function KulturaProsekPrag(ByVal vrstaVoca As String, ByVal colName As String) As Double
    Dim v As Variant
    v = LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, colName)
    If IsNumeric(v) Then KulturaProsekPrag = CDbl(v)
End Function

' Prosek jedne klase (neto/gajbe); azurira maxProsek + labelu ako je veci od dosad.
Private Sub ProsekKlase(ByVal kolicina As Double, ByVal kolAmb As Long, _
        ByVal klasa As String, ByRef maxProsek As Double, ByRef klasaLbl As String)
    If kolicina <= 0 Or kolAmb <= 0 Then Exit Sub
    Dim p As Double: p = kolicina / kolAmb
    If p > maxProsek Then
        maxProsek = p
        klasaLbl = klasa
    End If
End Sub

Public Function SaveOtkup(ByVal datum As Date, ByVal kooperantID As String, _
                          ByVal stanicaID As String, ByVal vrstaVoca As String, _
                          ByVal sortaVoca As String, ByVal kolicina As Double, _
                          ByVal cena As Double, ByVal tipAmb As String, _
                          ByVal kolAmb As Long, ByVal vozacID As String, _
                          ByVal brDok As String, ByVal novac As Double, _
                          ByVal primalac As String, _
                          Optional ByVal klasa As String = "I", _
                          Optional ByVal parcelaID As String = "", _
                          Optional ByVal brojZbirne As String = "", _
                          Optional ByVal kolAmbIzdata As Long = 0, _
                          Optional ByVal brutoKg As Double = 0) As String
    On Error GoTo EH

    If Trim$(kooperantID) = "" Then
        Err.Raise vbObjectError + 1820, "SaveOtkup", _
                  "Kooperant mora biti izabran."
    End If

    If Trim$(stanicaID) = "" Then
        Err.Raise vbObjectError + 1821, "SaveOtkup", _
                  "Stanica mora biti izabrana."
    End If

    If Trim$(vrstaVoca) = "" Then
        Err.Raise vbObjectError + 1822, "SaveOtkup", _
                  "Vrsta vo" & ChrW(263) & "a je obavezna."
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1823, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena <= 0 Then
        Err.Raise vbObjectError + 1824, "SaveOtkup", _
                  "Cena mora biti veca od nule."
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1825, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If novac < 0 Then
        Err.Raise vbObjectError + 1826, "SaveOtkup", _
                  "Novac ne sme biti negativan."
    End If

    If kolAmb > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1827, "SaveOtkup", _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    If kolAmbIzdata < 0 Then
        Err.Raise vbObjectError + 1831, "SaveOtkup", _
                  "Koli" & ChrW(269) & "ina izdate ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmbIzdata > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1832, "SaveOtkup", _
                  "Tip ambala" & ChrW(382) & "e je obavezan kada postoji izdata ambala" & ChrW(382) & "a."
    End If
    
    Call RequireValidOtkupClass(klasa, "SaveOtkup")

    RequireColumns TBL_OTKUP, "SaveOtkup", _
                   COL_OTK_ID, _
                   COL_OTK_DATUM, _
                   COL_OTK_KOOPERANT, _
                   COL_OTK_STANICA, _
                   COL_OTK_KULTURA, _
                   COL_OTK_VRSTA, _
                   COL_OTK_SORTA, _
                   COL_OTK_KOLICINA, _
                   COL_OTK_CENA, _
                   COL_OTK_TIP_AMB, _
                   COL_OTK_KOL_AMB, _
                   COL_OTK_VOZAC, _
                   COL_OTK_BR_DOK, _
                   COL_OTK_NOVAC, _
                   COL_OTK_PRIMALAC, _
                   COL_OTK_KLASA, _
                   COL_OTK_STORNIRANO, _
                   COL_OTK_BROJ_ZBIRNE, _
                   COL_OTK_ISPLACENO, _
                   COL_OTK_DATUM_ISPLATE, _
                   COL_OTK_OTPREMNICA_ID, _
                   COL_OTK_PARCELA

    Dim newID As String
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")

    If newID = "" Then
        Err.Raise vbObjectError + 1828, "SaveOtkup", _
                  "GetNextID nije vratio OtkupID."
    End If

    Dim kulturaID As String
    kulturaID = CStr(LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID"))

    If kulturaID = "" Then
        kulturaID = vrstaVoca & "-" & sortaVoca
    End If

    Dim rowData As Variant
    rowData = Array( _
        newID, _
        datum, _
        kooperantID, _
        stanicaID, _
        kulturaID, _
        vrstaVoca, _
        sortaVoca, _
        kolicina, _
        cena, _
        tipAmb, _
        kolAmb, _
        vozacID, _
        brDok, _
        novac, _
        primalac, _
        klasa, _
        "", _
        brojZbirne, _
        "", _
        Empty, _
        "", _
        parcelaID _
    )

    Dim newRow As Long
    newRow = AppendRow(TBL_OTKUP, rowData)
    If newRow <= 0 Then
        Err.Raise vbObjectError + 1829, "SaveOtkup", _
                  "AppendRow fehlgeschlagen fuer tblOtkup."
    End If

    ' Izdata ambalaza (OM->kooperant) -> upis u kolonu PO IMENU (kolona je na kraju
    ' tblOtkup; pozicijski rowData se ne dira). Kolona postoji posle EnsureDoradeSchema.
    If kolAmbIzdata > 0 Then
        UpdateCell TBL_OTKUP, newRow, COL_OTK_KOL_AMB_IZDATA, kolAmbIzdata
    End If

    ' Vreme snimanja otkupa (Now()) -> upis po imenu (kolona na kraju tblOtkup).
    UpdateCell TBL_OTKUP, newRow, COL_OTK_VREME_UNOSA, Now

    ' Bruto tezina (kad je unet bruto pa oduzeta ambalaza) -> upis po imenu; prazno = neto.
    If brutoKg > 0 Then UpdateCell TBL_OTKUP, newRow, COL_OTK_BRUTO, brutoKg

    If kolAmb > 0 Then
        ' Kooperant predaje pune gajbe na OM -> DVOJNI upis (otkup nema vozaca na OM-strani):
        '   1) Kooperant IZLAZ (kooperant se razduzuje),
        '   2) OM/Stanica ULAZ (OM se zaduzuje za isti iznos).
        TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", _
                      kooperantID, "Kooperant", vozacID, _
                      newID, DOK_TIP_OTKUP
        TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", _
                      stanicaID, "Stanica", "", _
                      newID, DOK_TIP_OTKUP
    End If

    If kolAmbIzdata > 0 Then
        ' OM izdaje prazne gajbe kooperantu (uz otkup) -> DVOJNI upis (bez vozaca):
        '   1) Kooperant ULAZ (dobija prazne),
        '   2) OM/Stanica IZLAZ (OM se razduzuje).
        ' Isti DokumentID (otkupID) -> storno otkupa hvata i ovu nogu (modStorno).
        TrackAmbalaza datum, tipAmb, kolAmbIzdata, "Ulaz", _
                      kooperantID, "Kooperant", "", _
                      newID, DOK_TIP_OM_IZLAZ_KOOP
        TrackAmbalaza datum, tipAmb, kolAmbIzdata, "Izlaz", _
                      stanicaID, "Stanica", "", _
                      newID, DOK_TIP_OM_IZLAZ_KOOP
    End If

    SaveOtkup = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "SaveOtkup"
    On Error Resume Next
    On Error GoTo 0

    Err.Raise errNum, "SaveOtkup", _
              "Source=" & errSrc & " | " & errDesc
End Function

Private Function GetKooperantNazivForNovac(ByVal kooperantID As String) As String
    On Error GoTo EH

    Dim ime As String
    Dim prezime As String

    ime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Ime")))
    prezime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Prezime")))

    GetKooperantNazivForNovac = Trim$(ime & " " & prezime)

    If GetKooperantNazivForNovac = "" Then
        GetKooperantNazivForNovac = kooperantID
    End If

    Exit Function

EH:
    LogErr "GetKooperantNazivForNovac"
    GetKooperantNazivForNovac = kooperantID
End Function

Public Function GetOtkupByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, _
            "modOtkup.GetOtkupByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByStation"
    GetOtkupByStation = Empty
End Function

Public Function GetOtkupByKooperant(ByVal kooperantID As String, _
                                    Optional ByVal datumOd As Date = 0, _
                                    Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
            "modOtkup.GetOtkupByKooperant"), "=", kooperantID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByKooperant"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByKooperant = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByKooperant"
    GetOtkupByKooperant = Empty
End Function
Public Function GetSaldoByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(otkupData) Then
        Dim colKoop As Long
        Dim colKol As Long
        Dim colNovac As Long
        Dim colAmb As Long

        colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
                                     "modOtkup.GetSaldoByStation")
        colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, _
                                    "modOtkup.GetSaldoByStation")
        colNovac = RequireColumnIndex(TBL_OTKUP, COL_OTK_NOVAC, _
                                      "modOtkup.GetSaldoByStation")
        colAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, _
                                    "modOtkup.GetSaldoByStation")

        Dim i As Long
        Dim key As String
        Dim vals As Variant

        For i = 1 To UBound(otkupData, 1)
            key = CStr(otkupData(i, colKoop))

            If key <> "" Then
                If Not dict.Exists(key) Then
                    dict.Add key, Array(0#, 0#, 0#)
                End If

                vals = dict(key)

                If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
                If IsNumeric(otkupData(i, colNovac)) Then vals(1) = vals(1) + CDbl(otkupData(i, colNovac))
                If IsNumeric(otkupData(i, colAmb)) Then vals(2) = vals(2) + CLng(otkupData(i, colAmb))

                dict(key) = vals
            End If
        Next i
    End If

    ' TODO:
    ' Ovaj helper trenutno racuna samo bruto saldo iz tblOtkup.
    ' Banka/Novac/Isporuka korekcije treba resiti u posebnom report modulu,
    ' ne siriti ovaj core save modul bez jasnog accounting pravila.

    If dict.count = 0 Then
        GetSaldoByStation = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 4)

    Dim keys As Variant
    keys = dict.keys

    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)

        vals = dict(keys(i))
        result(i + 1, 2) = vals(0)
        result(i + 1, 3) = vals(1)
        result(i + 1, 4) = vals(2)
    Next i

    GetSaldoByStation = result
    Exit Function

EH:
    LogErr "modOtkup.GetSaldoByStation"
    GetSaldoByStation = Empty
End Function


Private Sub PrintOtkupTxFailure(ByVal sourceName As String, _
                                ByVal errSrc As String, _
                                ByVal errNum As Long, _
                                ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidOtkupClass(ByVal klasa As String, _
                                   ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1830, sourceName, _
              "Neispravna klasa otkupa: " & klasa
End Sub

