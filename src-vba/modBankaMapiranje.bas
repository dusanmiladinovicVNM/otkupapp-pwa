Attribute VB_Name = "modBankaMapiranje"
Option Explicit

' ============================================================
' modBankaMapiranje
'
' Mapiranje iz tblBankaImport -> tblNovac
'
' Logika uskladjena sa stvarnim unosom iz frmDokumenta:
'
' 1) Kupac uplata:
'    - Partner = Naziv kupca
'    - PartnerID = KupacID
'    - EntitetTip = "Kupac"
'    - Tip = NOV_KUPCI_UPLATA ili NOV_KUPCI_AVANS
'
' 2) Isplata kooperantu preko banke:
'    - Partner = Naziv stanice / OM
'    - PartnerID = StanicaID
'    - EntitetTip = "OM"
'    - OMID = StanicaID
'    - KooperantID = KooperantID
'    - Tip = NOV_VIRMAN_FIRMA_KOOP ili NOV_VIRMAN_AVANS_KOOP
'
' 3) Uplata OM / dopuna OM:
'    - Partner = Naziv stanice
'    - PartnerID = StanicaID
'    - EntitetTip = "OM"
'    - OMID = StanicaID
'    - Tip = NOV_VIRMAN_FIRMA_OTKUPAC (izvod je bezgotovinski; KES tip se ovde
'      NE upisuje - vidi modConfig, sekcija Novac Tipovi)
'
' Obradjeno u tblBankaImport:
'   ""      = nije obradjeno
'   "Da"    = obradjeno
'   "Skip"  = preskoceno
'   "Error" = nije moguce automatski mapirati
'
' Napomena u tblNovac:
'   BIM:<id>; Ref:<...>; Konto:<...>; Opis:<...>; Svrha:<...>
' ============================================================

' ============================================================
' PATCH: P1-4 BankaMapiranje exact-row guards
' File: src-vba/modBankaMapiranje.bas
'
' Goals:
' - NovacID, BankaImportID, OtkupID, FakturaID use Count = 1.
' - Count = 0 => missing link error.
' - Count > 1 => duplicate key error.
' - Link/status updates use RequireUpdateCell.
' - Manual and auto-map share same integrity standard.
' - Any link error bubbles to *_TX wrapper and rollback runs.
' ============================================================

Private Const ERR_BMAP_BASE As Long = vbObjectError + 2900

' AUD-025: "za ovaj red je potrebno RUCNO mapiranje" - jedina greska koju batch
' petlje (AutoMapAll / strong pass) smeju da progutaju po redu umesto da obore ceo
' batch. Dize se PRE ijednog upisa za taj red, pa progutan slucaj ne moze ostaviti
' parcijalno stanje. Public je da ga test suite moze proveriti po broju.
' Vrednost = ERR_BMAP_BASE + 50 (ERR_BMAP_BASE je Private, pa se pise razvijeno).
Public Const ERR_BMAP_MANUAL_REQUIRED As Long = vbObjectError + 2950

' Blok otkupa po modelu ima najvise 2 otvorene klase (vrste voca). Treci kandidat
' je anomalija podataka (recikliran broj bloka, duplirani unos) - AUTOMATSKA
' raspodela se tu ne pogadja, nego se red salje operateru na rucno mapiranje
' (rucni put sme preko granice, ali samo uz izricitu potvrdu - vidi
' MapBankaImportAsKooperantBlockManual_TX allowManyCandidates).
Public Const MAX_BLOK_KANDIDATA As Long = 2

' Smer stavke izvoda - JEDAN klasifikator za citaoce (preview u formi) i pisce
' (RequireBimSmer). Bez zajednickog koda su preview i writer imali razlicit
' ugovor: preview je gledao samo "ima li uplate/isplate", pa je red sa OBA iznosa
' izgledao validno a writer bi ga odbio.
Public Const BIM_SMER_UPLATA As String = "UPLATA"
Public Const BIM_SMER_ISPLATA As String = "ISPLATA"
Public Const BIM_SMER_NEJASAN As String = "NEJASAN"

' Kad je True, batch auto-map (strong pass / Auto sve) ne prikazuje blokirajuci
' MsgBox po redu (npr. "Kooperant nema StanicaID"). Postavlja ga pozivalac oko petlje.
Public gBankaSilentBatch As Boolean

' ============================================================
' PUBLIC
' ============================================================

Public Function GetBankaImportOpen() As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim colObr As Long
    Dim i As Long, j As Long
    Dim outRow As Long
    
    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        GetBankaImportOpen = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        GetBankaImportOpen = Empty
        Exit Function
    End If
    
    colObr = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO, _
                               "GetBankaImportOpen")
    
    ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))
    
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(NzBIM(data(i, colObr), ""))) <> "Da" _
           And Trim$(CStr(NzBIM(data(i, colObr), ""))) <> "Skip" Then
            outRow = outRow + 1
            For j = 1 To UBound(data, 2)
                result(outRow, j) = data(i, j)
            Next j
        End If
    Next i
    
    If outRow = 0 Then
        GetBankaImportOpen = Empty
        Exit Function
    End If

    Dim finalResult() As Variant
    Dim r As Long, c As Long

    ReDim finalResult(1 To outRow, 1 To UBound(data, 2))

    For r = 1 To outRow
        For c = 1 To UBound(data, 2)
            finalResult(r, c) = result(r, c)
        Next c
    Next r

    GetBankaImportOpen = finalResult
End Function

Public Function AutoMapBankaImportRow(ByVal bankaImportID As String, _
                                      Optional ByVal strongOnly As Boolean = False, _
                                      Optional ByVal markErrorOnFail As Boolean = True) As String
    Dim bim As Variant
    Dim uplata As Double
    Dim isplata As Double
    Dim partnerName As String

    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        AutoMapBankaImportRow = ""
        Exit Function
    End If

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        If Not gBankaSilentBatch Then MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        AutoMapBankaImportRow = ""
        Exit Function
    End If

    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))
    partnerName = CStr(bim(1, 3))

    If uplata > 0 And isplata = 0 Then
        AutoMapBankaImportRow = AutoMapIncomingKupac(bankaImportID, strongOnly)
        If AutoMapBankaImportRow = "" And markErrorOnFail Then UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If

    If isplata > 0 And uplata = 0 Then
        AutoMapBankaImportRow = AutoMapOutgoingKooperantOrOM(bankaImportID, strongOnly)
        If AutoMapBankaImportRow = "" And markErrorOnFail Then UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If

    ' Nema cist smer uplata/isplata. U strong-only rezimu preskoci tiho (bez Error/MsgBox).
    If strongOnly Then
        AutoMapBankaImportRow = ""
        Exit Function
    End If

    If Not gBankaSilentBatch Then MsgBox "Stavka nema cist smer uplata/isplata: " & partnerName, vbExclamation, APP_NAME
    If markErrorOnFail Then UpdateBankaImportStatus bankaImportID, "Error"
    AutoMapBankaImportRow = ""
End Function

' AUD-025: batch-bezbedan omotac oko AutoMapBankaImportRow.
'
' Guta SAMO ERR_BMAP_MANUAL_REQUIRED ("ovaj red mora rucno") - svaka druga greska
' ide dalje i batch pada/rollback-uje se kao i pre. To je bezbedno jer se ta greska
' dize PRE ijednog upisa za taj red (resolver kandidata bloka), pa progutan slucaj
' ne moze ostaviti pola knjizenja. Ranije je jedan takav red (3+ otvorenih stavki u
' bloku) obarao CEO AutoMapAll batch i ponistavao sve vec mapirane redove.
Private Function AutoMapBankaImportRowBatch(ByVal bankaImportID As String, _
                                            ByVal strongOnly As Boolean, _
                                            ByVal markErrorOnFail As Boolean, _
                                            ByRef manualRequiredCount As Long) As String
    On Error GoTo EH

    AutoMapBankaImportRowBatch = AutoMapBankaImportRow(bankaImportID, strongOnly, markErrorOnFail)
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    If errNum <> ERR_BMAP_MANUAL_REQUIRED Then
        Err.Raise errNum, errSrc, errDesc
    End If

    manualRequiredCount = manualRequiredCount + 1

    ' Monitoring je best-effort (telemetrija sme da padne).
    On Error Resume Next

    Monitor_Event _
        eventType:="BANKA_AUTOMAP_MANUAL_REQUIRED", _
        severity:="WARN", _
        message:="Red trazi rucno mapiranje. BankaImportID=" & bankaImportID & _
                 "; Razlog=" & errDesc, _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:="AutoMapBankaImportRowBatch", _
        entityType:="BankaImport", _
        entityID:=bankaImportID, _
        correlationId:=bankaImportID

    On Error GoTo 0

    ' Poslovni status NIJE best-effort: ako se "Error" ne moze upisati, red bi
    ' ostao otvoren dok batch prijavljuje da je poslat na rucno. Pad ovde ide
    ' dalje i obara batch (rollback), umesto tihog neslaganja izvestaja i stanja.
    If markErrorOnFail Then UpdateBankaImportStatus bankaImportID, "Error"

    AutoMapBankaImportRowBatch = ""
End Function

Public Function AutoMapBankaImportRow_TX(ByVal bankaImportID As String) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_FAKTURE

    AutoMapBankaImportRow_TX = AutoMapBankaImportRow(bankaImportID)
    
    tx.CommitTx

    If Len(Trim$(AutoMapBankaImportRow_TX)) > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="AutoMapBankaImportRow_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=AutoMapBankaImportRow_TX, _
            partnerType:="AUTO"
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "AutoMapBankaImportRow_TX"

    Monitor_BankaMapFail _
        procedureName:="AutoMapBankaImportRow_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="AUTO"

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri automatskom mapiranju banke, promene vra" & ChrW(263) & "ene: " & errDesc, _
           vbCritical, APP_NAME

    AutoMapBankaImportRow_TX = ""
End Function

Public Function MapBankaImportAsKupac(ByVal bankaImportID As String, _
                                      ByVal kupacID As String, _
                                      Optional ByVal fakturaID As String = "", _
                                      Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim kupacNaziv As String
    Dim uplataUkupno As Double
    Dim otvorenoNaFakturi As Double
    Dim naFakturu As Double
    Dim avans As Double

    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsKupac = ""
        Exit Function
    End If
    
    If Trim$(kupacID) = "" Then
        MsgBox "KupacID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsKupac = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsKupac = ""
        Exit Function
    End If

    ' Kupac = novac ULAZI u firmu. Isplata na ovom tipu je odbijena (AUD-025).
    RequireBimSmer bim, "UPLATA", "MapBankaImportAsKupac"

    kupacNaziv = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    If kupacNaziv = "" Then kupacNaziv = CStr(bim(1, 3))

    uplataUkupno = CDbl(NzBIM(bim(1, 5), 0#))
    naFakturu = 0
    avans = uplataUkupno

    If fakturaID <> "" Then
        RequireSingleRow TBL_FAKTURE, COL_FAK_ID, fakturaID, _
                         "MapBankaImportAsKupac"
        RequireFakturaZaKupca fakturaID, kupacID, "MapBankaImportAsKupac"

        ' Preplata: na fakturu ide najvise njen OTVOREN iznos, visak je avans
        ' kupca. Ranije je ceo iznos izvoda isao na fakturu sa `NOV_KUPCI_UPLATA`,
        ' pa je faktura mogla da ostane preplacena, a stvaran avans neevidentiran.
        ' Saldo se cita ovde (a ne iz liste u formi) jer se izmedju prikaza i
        ' klika mogao promeniti.
        otvorenoNaFakturi = GetOtvorenoNaFakturi(fakturaID)

        If otvorenoNaFakturi <= 0.009 Then
            Err.Raise ERR_BMAP_BASE + 53, "MapBankaImportAsKupac", _
                "Faktura " & fakturaID & " nema otvoren iznos (vec je placena). " & _
                "Osvezi listu pa izaberi drugu fakturu ili knjizi kao avans."
        End If

        If uplataUkupno <= otvorenoNaFakturi Then
            naFakturu = uplataUkupno
        Else
            naFakturu = otvorenoNaFakturi
        End If

        avans = uplataUkupno - naFakturu
    End If

    If naFakturu > 0 Then
        MapBankaImportAsKupac = SaveNovac( _
            RequireIzvodBroj(bim, "MapBankaImportAsKupac"), _
            CDate(bim(1, 2)), _
            kupacNaziv, _
            kupacID, _
            "Kupac", _
            "", _
            "", _
            fakturaID, _
            "", _
            NOV_KUPCI_UPLATA, _
            naFakturu, _
            0, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kupac") _
        )

        If MapBankaImportAsKupac = "" Then
            UpdateBankaImportStatus bankaImportID, "Error"
            Exit Function
        End If
    End If

    If avans > 0.009 Then
        Dim avansID As String

        avansID = SaveNovac( _
            RequireIzvodBroj(bim, "MapBankaImportAsKupac"), _
            CDate(bim(1, 2)), _
            kupacNaziv, _
            kupacID, _
            "Kupac", _
            "", _
            "", _
            "", _
            "", _
            NOV_KUPCI_AVANS, _
            avans, _
            0, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kupac-visak") _
        )

        If avansID = "" Then
            ' Deo iznosa bi ostao neproknjizen, a stavka bi bila zatvorena kao
            ' obradjena -- fail-closed (TX omotac vraca i eventualni prvi red).
            Err.Raise ERR_BMAP_BASE + 54, "MapBankaImportAsKupac", _
                "Avansni deo uplate (" & Format$(avans, "#,##0.00") & _
                ") nije proknjizen za BankaImportID=" & bankaImportID
        End If

        If MapBankaImportAsKupac = "" Then MapBankaImportAsKupac = avansID
    End If

    If MapBankaImportAsKupac = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If

    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), kupacID, "Kupac", "")
    End If
    
    If fakturaID <> "" Then
        RequireSingleRow TBL_FAKTURE, COL_FAK_ID, fakturaID, _
                         "MapBankaImportAsKupac"
        UpdateFakturaStatus fakturaID
    End If
End Function

Public Function MapBankaImportAsKupac_TX(ByVal bankaImportID As String, _
                                         ByVal kupacID As String, _
                                         Optional ByVal fakturaID As String = "", _
                                         Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKupac_TX = MapBankaImportAsKupac(bankaImportID, kupacID, fakturaID, savePartnerMapFlag)
    
    tx.CommitTx

    If Len(Trim$(MapBankaImportAsKupac_TX)) > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="MapBankaImportAsKupac_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=MapBankaImportAsKupac_TX, _
            partnerType:="Kupac", _
            partnerId:=kupacID, _
            linkedEntityId:=fakturaID
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "MapBankaImportAsKupac_TX"

    Monitor_BankaMapFail _
        procedureName:="MapBankaImportAsKupac_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="Kupac", _
        partnerId:=kupacID, _
        linkedEntityId:=fakturaID

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri mapiranju kupca, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    MapBankaImportAsKupac_TX = ""
End Function

Public Function MapBankaImportAsKooperant(ByVal bankaImportID As String, _
                                          ByVal kooperantID As String, _
                                          Optional ByVal otkupID As String = "", _
                                          Optional ByVal vrstaVoca As String = "", _
                                          Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim omID As String
    Dim omNaziv As String
    Dim tip As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    If Trim$(kooperantID) = "" Then
        MsgBox "KooperantID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    ' Kooperant = novac IZLAZI iz firme. Uplata na ovom tipu je odbijena (AUD-025).
    RequireBimSmer bim, "ISPLATA", "MapBankaImportAsKooperant"

    omID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    If omID = "" Then
        MsgBox "Kooperant nema StanicaID!", vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID
    
    If otkupID <> "" Then
        tip = NOV_VIRMAN_FIRMA_KOOP
    Else
        tip = NOV_VIRMAN_AVANS_KOOP
    End If
    
    If otkupID <> "" Then
        RequireSingleRow TBL_OTKUP, COL_OTK_ID, otkupID, _
                         "MapBankaImportAsKooperant"
    End If
    
    MapBankaImportAsKooperant = SaveNovac( _
        RequireIzvodBroj(bim, "MapBankaImportAsKooperant"), _
        CDate(bim(1, 2)), _
        omNaziv, _
        omID, _
        "OM", _
        omID, _
        kooperantID, _
        "", _
        vrstaVoca, _
        tip, _
        CDbl(NzBIM(bim(1, 5), 0#)), _
        CDbl(NzBIM(bim(1, 6), 0#)), _
        BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
    )
    
    If MapBankaImportAsKooperant = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
    End If
    
    If otkupID <> "" Then
        LinkNovacToOtkupStrict MapBankaImportAsKooperant, otkupID, _
                                "MapBankaImportAsKooperant"
    End If
End Function

Public Function MapBankaImportAsKooperant_TX(ByVal bankaImportID As String, _
                                             ByVal kooperantID As String, _
                                             Optional ByVal otkupID As String = "", _
                                             Optional ByVal vrstaVoca As String = "", _
                                             Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperant_TX = MapBankaImportAsKooperant(bankaImportID, kooperantID, otkupID, vrstaVoca, savePartnerMapFlag)
    
    tx.CommitTx

    If Len(Trim$(MapBankaImportAsKooperant_TX)) > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="MapBankaImportAsKooperant_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=MapBankaImportAsKooperant_TX, _
            partnerType:="Kooperant", _
            partnerId:=kooperantID, _
            linkedEntityId:=otkupID
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "MapBankaImportAsKooperant_TX"

    Monitor_BankaMapFail _
        procedureName:="MapBankaImportAsKooperant_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="Kooperant", _
        partnerId:=kooperantID, _
        linkedEntityId:=otkupID

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri mapiranju kooperanta, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    MapBankaImportAsKooperant_TX = ""
End Function

Public Function MapBankaImportAsOM(ByVal bankaImportID As String, _
                                   ByVal omID As String, _
                                   Optional ByVal vrstaVoca As String = "", _
                                   Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim omNaziv As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    If Trim$(omID) = "" Then
        MsgBox "OMID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    ' OM promet ide u oba smera (dopuna OM / povracaj), pa se trazi samo CIST smer:
    ' red sa oba iznosa ili bez iznosa nije knjizljiv (AUD-025).
    RequireBimSmer bim, "", "MapBankaImportAsOM"

    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID

    MapBankaImportAsOM = SaveNovac( _
        RequireIzvodBroj(bim, "MapBankaImportAsOM"), _
        CDate(bim(1, 2)), _
        omNaziv, _
        omID, _
        "OM", _
        omID, _
        "", _
        "", _
        vrstaVoca, _
        NOV_VIRMAN_FIRMA_OTKUPAC, _
        CDbl(NzBIM(bim(1, 5), 0#)), _
        CDbl(NzBIM(bim(1, 6), 0#)), _
        BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "OM") _
    )
    
    If MapBankaImportAsOM = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), omID, "OM", omID)
    End If
End Function

Public Function MapBankaImportAsOM_TX(ByVal bankaImportID As String, _
                                      ByVal omID As String, _
                                      Optional ByVal vrstaVoca As String = "", _
                                      Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsOM_TX = MapBankaImportAsOM(bankaImportID, omID, vrstaVoca, savePartnerMapFlag)
    
    tx.CommitTx

    If Len(Trim$(MapBankaImportAsOM_TX)) > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="MapBankaImportAsOM_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=MapBankaImportAsOM_TX, _
            partnerType:="OM", _
            partnerId:=omID
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "MapBankaImportAsOM_TX"

    Monitor_BankaMapFail _
        procedureName:="MapBankaImportAsOM_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="OM", _
        partnerId:=omID

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri mapiranju OM, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    MapBankaImportAsOM_TX = ""
End Function

Public Function MapBankaImportAsKooperantBlock(ByVal bankaImportID As String, _
                                               ByVal kooperantID As String, _
                                               Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim blockNo As String
    
    blockNo = Trim$(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    MapBankaImportAsKooperantBlock = MapBankaImportAsKooperantBlockCore( _
        bankaImportID, kooperantID, blockNo, savePartnerMapFlag)
End Function

Public Function MapBankaImportAsKooperantBlock_TX(ByVal bankaImportID As String, _
                                                  ByVal kooperantID As String, _
                                                  Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperantBlock_TX = MapBankaImportAsKooperantBlock(bankaImportID, kooperantID, savePartnerMapFlag)
    
    tx.CommitTx

    If MapBankaImportAsKooperantBlock_TX > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="MapBankaImportAsKooperantBlock_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=CStr(MapBankaImportAsKooperantBlock_TX), _
            partnerType:="KooperantBlock", _
            partnerId:=kooperantID, _
            linkedEntityId:=CStr(MapBankaImportAsKooperantBlock_TX)
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "MapBankaImportAsKooperantBlock_TX"

    Monitor_BankaMapFail _
        procedureName:="MapBankaImportAsKooperantBlock_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="KooperantBlock", _
        partnerId:=kooperantID

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri mapiranju kooperanta po bloku, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    MapBankaImportAsKooperantBlock_TX = 0
End Function

' allowManyCandidates: operater je video kandidate i predlozenu podelu pa je
' potvrdio (frmBankaImport). Samo tako blok sa 3+ otvorenih stavki moze da se
' zavrsi - automatski put ga i dalje odbija (ERR_BMAP_MANUAL_REQUIRED).
Public Function MapBankaImportAsKooperantBlockManual(ByVal bankaImportID As String, _
                                                     ByVal kooperantID As String, _
                                                     ByVal brojBloka As String, _
                                                     Optional ByVal savePartnerMapFlag As Boolean = True, _
                                                     Optional ByVal allowManyCandidates As Boolean = False) As Long
    MapBankaImportAsKooperantBlockManual = MapBankaImportAsKooperantBlockCore( _
        bankaImportID, kooperantID, brojBloka, savePartnerMapFlag, allowManyCandidates)
End Function

Public Function MapBankaImportAsKooperantBlockManual_TX(ByVal bankaImportID As String, _
                                                        ByVal kooperantID As String, _
                                                        ByVal brojBloka As String, _
                                                        Optional ByVal savePartnerMapFlag As Boolean = True, _
                                                        Optional ByVal allowManyCandidates As Boolean = False) As Long
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperantBlockManual_TX = MapBankaImportAsKooperantBlockManual( _
        bankaImportID, kooperantID, brojBloka, savePartnerMapFlag, allowManyCandidates)
    
    tx.CommitTx

    If MapBankaImportAsKooperantBlockManual_TX > 0 Then
        Monitor_BankaMapSuccess _
            procedureName:="MapBankaImportAsKooperantBlockManual_TX", _
            bankaImportID:=bankaImportID, _
            resultId:=CStr(MapBankaImportAsKooperantBlockManual_TX), _
            partnerType:="KooperantBlockManual", _
            partnerId:=kooperantID, _
            linkedEntityId:=brojBloka
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "MapBankaImportAsKooperantBlockManual_TX"

    Monitor_BankaMapFail _
        procedureName:="MapBankaImportAsKooperantBlockManual_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="KooperantBlockManual", _
        partnerId:=kooperantID, _
        linkedEntityId:=brojBloka

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri rucnom mapiranju kooperanta po bloku, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    MapBankaImportAsKooperantBlockManual_TX = 0
End Function

Private Function MapBankaImportAsKooperantBlockCore(ByVal bankaImportID As String, _
                                                    ByVal kooperantID As String, _
                                                    ByVal blockNo As String, _
                                                    Optional ByVal savePartnerMapFlag As Boolean = True, _
                                                    Optional ByVal allowManyCandidates As Boolean = False) As Long
    Dim bim As Variant
    Dim omID As String
    Dim omNaziv As String
    Dim isplataUkupno As Double
    Dim preostaloZaRaspodelu As Double
    Dim kandidati As Variant
    Dim plan As Variant
    Dim planCount As Long
    Dim i As Long

    If Not ValidateBankaImportNotProcessed(bankaImportID) Then Exit Function
    
    If Trim$(kooperantID) = "" Then
        MsgBox "KooperantID je obavezan!", vbExclamation, APP_NAME
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        Exit Function
    End If
    
    ' Blok kooperanta = isplata. Raniji tihi `Exit Function` na uplati je iz UI
    ' izgledao kao "nista se nije desilo"; sada je odbijanje glasno (AUD-025).
    RequireBimSmer bim, "ISPLATA", "MapBankaImportAsKooperantBlockCore"

    omID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    If omID = "" Then
        If Not gBankaSilentBatch Then MsgBox "Kooperant nema StanicaID!", vbExclamation, APP_NAME
        Exit Function
    End If

    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID

    isplataUkupno = CDbl(NzBIM(bim(1, 6), 0#))
    
    kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo, allowManyCandidates)
    preostaloZaRaspodelu = isplataUkupno

    If IsEmpty(kandidati) Then
        If SaveNovac( _
            RequireIzvodBroj(bim, "MapBankaImportAsKooperantBlockCore"), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            "", _
            NOV_VIRMAN_AVANS_KOOP, _
            0, _
            preostaloZaRaspodelu, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
        ) <> "" Then
            MapBankaImportAsKooperantBlockCore = 1
            UpdateBankaImportStatus bankaImportID, "Da"
            If savePartnerMapFlag Then
                Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
            End If
        Else
            UpdateBankaImportStatus bankaImportID, "Error"
        End If
        Exit Function
    End If
    
    ' Podela je izracunata na JEDNOM mestu (PlanBlokRaspodela) koje koristi i
    ' preview u formi -- prikaz i knjizenje ne mogu da se razidju.
    plan = PlanBlokRaspodela(kandidati, isplataUkupno)

    If IsEmpty(plan) Then
        planCount = 0
    Else
        planCount = UBound(plan, 1)
    End If

    For i = 1 To planCount
        If preostaloZaRaspodelu <= 0 Then Exit For

        Dim otkupID As String
        Dim iznosZaRed As Double
        Dim vrstaVoca As String
        Dim novID As String

        otkupID = CStr(plan(i, 1))
        RequireSingleRow TBL_OTKUP, COL_OTK_ID, otkupID, _
                 "MapBankaImportAsKooperantBlockCore"
        iznosZaRed = CDbl(NzBIM(plan(i, 2), 0#))
        vrstaVoca = CStr(plan(i, 3))

        If iznosZaRed <= 0 Then GoTo NextCandidate

        novID = SaveNovac( _
            RequireIzvodBroj(bim, "MapBankaImportAsKooperantBlockCore"), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            vrstaVoca, _
            NOV_VIRMAN_FIRMA_KOOP, _
            0, _
            iznosZaRed, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
        )
        
        If Len(Trim$(novID)) = 0 Then
            Err.Raise ERR_BMAP_BASE + 40, "MapBankaImportAsKooperantBlockCore", _
              "SaveNovac nije vratio NovacID za OtkupID=" & otkupID
        End If

        LinkNovacToOtkupStrict novID, otkupID, _
                        "MapBankaImportAsKooperantBlockCore"

        MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
        preostaloZaRaspodelu = preostaloZaRaspodelu - iznosZaRed

NextCandidate:
    Next i

    If preostaloZaRaspodelu > 0 Then
        If SaveNovac( _
            RequireIzvodBroj(bim, "MapBankaImportAsKooperantBlockCore"), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            "", _
            NOV_VIRMAN_AVANS_KOOP, _
            0, _
            preostaloZaRaspodelu, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant-visak") _
        ) <> "" Then
            MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
        End If
    End If
    
    If MapBankaImportAsKooperantBlockCore > 0 Then
        UpdateBankaImportStatus bankaImportID, "Da"
        If savePartnerMapFlag Then
            Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
        End If
    Else
        UpdateBankaImportStatus bankaImportID, "Error"
    End If
End Function


Public Function SkipBankaImportRow(ByVal bankaImportID As String) As Boolean
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        SkipBankaImportRow = False
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Skip"
    SkipBankaImportRow = True
End Function

Public Function SkipBankaImportRow_TX(ByVal bankaImportID As String) As Boolean
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    
    SkipBankaImportRow_TX = SkipBankaImportRow(bankaImportID)
    
    tx.CommitTx

    If SkipBankaImportRow_TX Then
        On Error Resume Next
        Monitor_Event _
            eventType:="BANKA_IMPORT_SKIP", _
            severity:="WARN", _
            message:="Bank import row skipped. BankaImportID=" & bankaImportID, _
            userId:="Operator", _
            moduleName:="modBankaMapiranje", _
            procedureName:="SkipBankaImportRow_TX", _
            entityType:="BankaImport", _
            entityID:=bankaImportID, _
            correlationId:=bankaImportID
        On Error GoTo 0
    End If

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "SkipBankaImportRow_TX"

    Monitor_BankaMapFail _
        procedureName:="SkipBankaImportRow_TX", _
        bankaImportID:=bankaImportID, _
        errNum:=errNum, _
        errDesc:=errDesc, _
        errSrc:=errSrc, _
        partnerType:="Skip"

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri preskakanju bank stavke, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    SkipBankaImportRow_TX = False
End Function

Public Function AutoMapAllBankaImport_TX(Optional ByRef manualRequiredCount As Long) As Long
    Dim tx As clsTransaction
    Dim data As Variant
    Dim colID As Long
    Dim i As Long
    Dim novID As String
    Dim openCount As Long
    Dim failedCount As Long

    manualRequiredCount = 0

    On Error GoTo EH
    
    data = GetBankaImportOpen()
    If IsEmpty(data) Then
        AutoMapAllBankaImport_TX = 0
        Exit Function
    End If
    
    openCount = UBound(data, 1)

    On Error Resume Next
    Monitor_Event _
        eventType:="BANKA_AUTOMAP_ALL_START", _
        severity:="INFO", _
        message:="Started auto mapping all bank import rows. OpenRows=" & CStr(openCount), _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:="AutoMapAllBankaImport_TX", _
        entityType:="BankaImport", _
        entityID:="Batch", _
        correlationId:="BANKA-AUTOMAP-ALL"
    On Error GoTo EH
    
    colID = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, _
                           "AutoMapAllBankaImport_TX")
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_FAKTURE

    For i = 1 To UBound(data, 1)
        novID = AutoMapBankaImportRowBatch(CStr(data(i, colID)), False, True, manualRequiredCount)

        If novID <> "" Then
            AutoMapAllBankaImport_TX = AutoMapAllBankaImport_TX + 1
        Else
            failedCount = failedCount + 1
        End If
    Next i

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="BANKA_AUTOMAP_ALL_SUMMARY", _
        severity:="INFO", _
        message:="Bank auto mapping completed. OpenRows=" & CStr(openCount) & _
                 "; Mapped=" & CStr(AutoMapAllBankaImport_TX) & _
                 "; NotMapped=" & CStr(failedCount) & _
                 "; ManualRequired=" & CStr(manualRequiredCount), _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:="AutoMapAllBankaImport_TX", _
        entityType:="BankaImport", _
        entityID:="Batch", _
        correlationId:="BANKA-AUTOMAP-ALL"
    On Error GoTo 0

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "AutoMapAllBankaImport_TX"

    Monitor_Error _
        moduleName:="modBankaMapiranje", _
        procedureName:="AutoMapAllBankaImport_TX", _
        entityType:="BankaImport", _
        entityID:="Batch", _
        correlationId:="BANKA-AUTOMAP-ALL", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="BANKA_AUTOMAP_ALL_FAIL", _
        severity:="CRITICAL", _
        message:="Auto mapping all bank import rows failed. OpenRows=" & CStr(openCount) & _
                 "; MappedBeforeFail=" & CStr(AutoMapAllBankaImport_TX) & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:="AutoMapAllBankaImport_TX", _
        entityType:="BankaImport", _
        entityID:="Batch", _
        correlationId:="BANKA-AUTOMAP-ALL"

    If Not tx Is Nothing Then tx.RollbackTx

    MsgBox "Gre" & ChrW(353) & "ka pri automatskom mapiranju svih bank stavki, promene vra" & ChrW(263) & "ene: " & errDesc, vbCritical, APP_NAME

    AutoMapAllBankaImport_TX = 0
End Function


' ============================================================
' PRIVATE - AUTO MAP
' ============================================================

' ============================================================
' AUTO-MAP SAMO PREKO JAKIH KLJUCEVA (poziv->otkup/faktura, tekuci racun).
' Ne koristi ime/PartnerMap heuristiku i NE markira Error na promasaj -
' dvosmislene stavke ostaju otvorene za rucno mapiranje. Namenjeno auto-pokretanju
' pri otvaranju frmBankaImport (i kao Immediate/backfill jednim pozivom).
' Vraca broj auto-mapiranih stavki.
' ============================================================
Public Function AutoMapStrongKeysBankaImport_TX() As Long
    Const SRC As String = "AutoMapStrongKeysBankaImport_TX"

    Dim tx As clsTransaction
    Dim data As Variant
    Dim colID As Long
    Dim i As Long
    Dim res As String
    Dim mappedCount As Long
    Dim manualRequired As Long

    On Error GoTo EH

    data = GetBankaImportOpen()
    If IsEmpty(data) Then
        AutoMapStrongKeysBankaImport_TX = 0
        Exit Function
    End If

    colID = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, SRC)

    gBankaSilentBatch = True

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_FAKTURE

    ' markErrorOnFail = False: dvosmislena stavka ostaje OTVORENA za rucno mapiranje.
    ' manualRequired se ovde samo broji (red ostaje otvoren, ne dobija status Error).
    For i = 1 To UBound(data, 1)
        res = AutoMapBankaImportRowBatch(CStr(data(i, colID)), True, False, manualRequired)
        If res <> "" Then mappedCount = mappedCount + 1
    Next i

    tx.CommitTx
    Set tx = Nothing

    gBankaSilentBatch = False

    AutoMapStrongKeysBankaImport_TX = mappedCount
    Exit Function

EH:
    ' AUD-014: greska se vise NE guta. Pozivalac (dugme u frmBankaImport) mora da
    ' vidi da pass nije prosao - tiho "0 mapirano" je izgledalo kao "nema sta da se
    ' mapira", a u stvari je ceo pass bio rollback-ovan.
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    If Not tx Is Nothing Then tx.RollbackTx
    gBankaSilentBatch = False
    On Error GoTo 0

    AutoMapStrongKeysBankaImport_TX = 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

' ============================================================
' JAKI KLJUCEVI - CISTO CITANJE (bez upisa)
'
' AUD-014: otvaranje frmBankaImport je ranije KNJIZILO novac (auto-map na
' Activate). Da bi forma mogla da PRIKAZE sta bi se mapiralo, a da knjizi tek na
' klik, odluka o jakim kljucevima mora da postoji kao read-only funkcija. Zato je
' izdvojena ovde i koriste je oba puta (auto-map i prebrojavanje) - jedan izvor
' istine, bez druge kopije pravila.
' ============================================================

' Vraca KupacID po jakim kljucevima ili "" (ukljucujuci konflikt racun<->poziv).
' fakturaID je izlazni parametar: popunjen samo kad poziv na broj jednoznacno
' pogodi fakturu.
Private Function ResolveStrongKupac(ByVal bankaImportID As String, _
                                    ByRef fakturaID As String) As String
    Dim bim As Variant
    Dim kupacByKonto As Variant
    Dim fakByPoziv As Variant
    Dim idDoc As String
    Dim idKonto As String

    fakturaID = ""

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function

    kupacByKonto = TryResolveKupacByKonto(CStr(NzBIM(bim(1, 4), "")))
    fakByPoziv = TryResolveFakturaByPoziv(CStr(NzBIM(bim(1, 10), "")))

    If Not IsEmpty(fakByPoziv) Then
        fakturaID = CStr(fakByPoziv(0))
        idDoc = CStr(fakByPoziv(1))
    End If

    If Not IsEmpty(kupacByKonto) Then idKonto = CStr(kupacByKonto(0))

    ' Konflikt: racun i poziv na broj pokazuju na razlicite kupce -> rucno, ne auto.
    If idDoc <> "" And idKonto <> "" And UCase$(idDoc) <> UCase$(idKonto) Then
        fakturaID = ""
        Exit Function
    End If

    If idDoc <> "" Then
        ResolveStrongKupac = idDoc
    Else
        ResolveStrongKupac = idKonto
    End If
End Function

' Vraca KooperantID po jakim kljucevima ili "" (ukljucujuci konflikt poziv<->racun).
Private Function ResolveStrongKooperant(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim koopByKonto As Variant
    Dim idDoc As String
    Dim idKonto As String

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function

    idDoc = TryResolveKooperantByOtkupPoziv(CStr(NzBIM(bim(1, 10), "")))
    koopByKonto = TryResolveKooperantByKonto(CStr(NzBIM(bim(1, 4), "")))

    If Not IsEmpty(koopByKonto) Then idKonto = CStr(koopByKonto(0))

    ' Konflikt: poziv na broj i racun pokazuju na razlicite kooperante -> rucno.
    If idDoc <> "" And idKonto <> "" And UCase$(idDoc) <> UCase$(idKonto) Then
        Exit Function
    End If

    If idDoc <> "" Then
        ResolveStrongKooperant = idDoc
    Else
        ResolveStrongKooperant = idKonto
    End If
End Function

' Koliko OTVORENIH stavki bi jaki kljucevi mapirali - bez ijednog upisa.
' Koristi ga frmBankaImport pri otvaranju (prikaz umesto tihog knjizenja).
'
' NE guta gresku: schema/read problem bi kao "0" izgledao isto kao "nema sta da
' se mapira" - a AUD-014 je upravo o tome da greske jakog prolaza budu vidljive.
' Pozivalac (forma) hvata i prikazuje stanje greske.
Public Function CountStrongKeyReadyBankaImport() As Long
    Const SRC As String = "CountStrongKeyReadyBankaImport"

    On Error GoTo EH

    Dim data As Variant
    Dim colID As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim i As Long
    Dim uplata As Double
    Dim isplata As Double
    Dim bimID As String
    Dim fakturaID As String

    data = GetBankaImportOpen()
    If IsEmpty(data) Then Exit Function

    colID = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, SRC)
    colUplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA, SRC)
    colIsplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA, SRC)

    For i = 1 To UBound(data, 1)
        bimID = CStr(data(i, colID))
        uplata = CDbl(NzBIM(data(i, colUplata), 0#))
        isplata = CDbl(NzBIM(data(i, colIsplata), 0#))

        If uplata > 0 And isplata = 0 Then
            If ResolveStrongKupac(bimID, fakturaID) <> "" Then
                CountStrongKeyReadyBankaImport = CountStrongKeyReadyBankaImport + 1
            End If
        ElseIf isplata > 0 And uplata = 0 Then
            If ResolveStrongKooperant(bimID) <> "" Then
                CountStrongKeyReadyBankaImport = CountStrongKeyReadyBankaImport + 1
            End If
        End If
    Next i

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    On Error GoTo 0

    CountStrongKeyReadyBankaImport = 0
    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Private Function AutoMapIncomingKupac(ByVal bankaImportID As String, Optional ByVal strongOnly As Boolean = False) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim mapped As Variant
    Dim fakturaID As String
    Dim resolvedKupac As String
    Dim learn As Boolean

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function

    partnerName = CStr(bim(1, 3))
    learn = (Len(Trim$(partnerName)) > 0)

    ' PRIORITET: jaki kljucevi (tekuci racun -> kupac, poziv na broj -> faktura).
    ' Sama odluka zivi u ResolveStrongKupac da bi isti kod mogao i SAMO da prebroji
    ' sta bi se mapiralo, bez knjizenja (CountStrongKeyReadyBankaImport).
    resolvedKupac = ResolveStrongKupac(bankaImportID, fakturaID)

    If resolvedKupac <> "" Then
        If fakturaID = "" Then
            fakturaID = TryResolveFakturaForKupac(bankaImportID, resolvedKupac)
        End If
        AutoMapIncomingKupac = MapBankaImportAsKupac(bankaImportID, resolvedKupac, fakturaID, learn)
        Exit Function
    End If

    ' Strong-only rezim (auto na otvaranje): ne idi na ime/PartnerMap heuristiku.
    If strongOnly Then
        AutoMapIncomingKupac = ""
        Exit Function
    End If

    ' FALLBACK (nepromenjeno): PartnerMap -> egzaktno ime -> OM.
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        If CStr(mapped(1)) = "Kupac" Then
            fakturaID = TryResolveFakturaForKupac(bankaImportID, CStr(mapped(0)))
            AutoMapIncomingKupac = MapBankaImportAsKupac(bankaImportID, CStr(mapped(0)), fakturaID, False)
            Exit Function
        End If
    End If

    mapped = TryResolveKupacBIM(partnerName)
    If Not IsEmpty(mapped) Then
        fakturaID = TryResolveFakturaForKupac(bankaImportID, CStr(mapped(0)))
        AutoMapIncomingKupac = MapBankaImportAsKupac(bankaImportID, CStr(mapped(0)), fakturaID, False)
        Exit Function
    End If

    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        AutoMapIncomingKupac = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
        Exit Function
    End If

    AutoMapIncomingKupac = ""
End Function

Private Function AutoMapOutgoingKooperantOrOM(ByVal bankaImportID As String, Optional ByVal strongOnly As Boolean = False) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim mapped As Variant
    Dim createdCount As Long
    Dim resolvedKoop As String
    Dim learn As Boolean

    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function

    partnerName = CStr(bim(1, 3))
    learn = (Len(Trim$(partnerName)) > 0)

    ' PRIORITET: jaki kljucevi (poziv na broj -> otkupni list, tekuci racun).
    ' Odluka je izdvojena u ResolveStrongKooperant (vidi ResolveStrongKupac).
    resolvedKoop = ResolveStrongKooperant(bankaImportID)

    If resolvedKoop <> "" Then
        createdCount = MapBankaImportAsKooperantBlock(bankaImportID, resolvedKoop, learn)
        If createdCount > 0 Then AutoMapOutgoingKooperantOrOM = "OK"
        Exit Function
    End If

    ' Strong-only rezim (auto na otvaranje): ne idi na ime/PartnerMap heuristiku.
    If strongOnly Then
        AutoMapOutgoingKooperantOrOM = ""
        Exit Function
    End If

    ' FALLBACK (nepromenjeno): PartnerMap -> egzaktno ime -> OM.
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        Select Case CStr(mapped(1))
            Case "Kooperant"
                createdCount = MapBankaImportAsKooperantBlock(bankaImportID, CStr(mapped(0)), False)
                If createdCount > 0 Then AutoMapOutgoingKooperantOrOM = "OK"
                Exit Function
            Case "OM"
                AutoMapOutgoingKooperantOrOM = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
                Exit Function
        End Select
    End If

    mapped = TryResolveKooperantBIM(partnerName)
    If Not IsEmpty(mapped) Then
        createdCount = MapBankaImportAsKooperantBlock(bankaImportID, CStr(mapped(0)), False)
        If createdCount > 0 Then AutoMapOutgoingKooperantOrOM = "OK"
        Exit Function
    End If

    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        AutoMapOutgoingKooperantOrOM = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
        Exit Function
    End If

    AutoMapOutgoingKooperantOrOM = ""
End Function



' ============================================================
' PUBLIC - FACTURA MATCH
' ============================================================

Public Function TryResolveFakturaForKupac(ByVal bankaImportID As String, ByVal kupacID As String) As String
    Dim bim As Variant
    Dim faktData As Variant
    Dim colFID As Long, colBroj As Long, colKup As Long, colIznos As Long, colStatus As Long
    Dim poziv As String, svrha As String, uplata As Double
    Dim i As Long, hitCount As Long, hitID As String
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function
    
    poziv = NormalizeLooseBIM(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    svrha = NormalizeLooseBIM(CStr(bim(1, 8)))
    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    
    faktData = GetTableData(TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    faktData = ExcludeStornirano(faktData, TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colIznos = GetColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS)
    colStatus = GetColumnIndex(TBL_FAKTURE, COL_FAK_STATUS)
    
    For i = 1 To UBound(faktData, 1)
        If CStr(faktData(i, colKup)) <> kupacID Then GoTo NextI
        
        Dim brojFak As String
        brojFak = NormalizeLooseBIM(CStr(faktData(i, colBroj)))
        
        If brojFak <> "" Then
            If poziv <> "" Then
                If InStr(1, poziv, brojFak, vbTextCompare) > 0 Or InStr(1, brojFak, poziv, vbTextCompare) > 0 Then
                    hitCount = hitCount + 1
                    hitID = CStr(faktData(i, colFID))
                    GoTo NextI
                End If
            End If
            
            If svrha <> "" Then
                If InStr(1, svrha, brojFak, vbTextCompare) > 0 Then
                    hitCount = hitCount + 1
                    hitID = CStr(faktData(i, colFID))
                    GoTo NextI
                End If
            End If
        End If
        
        If Abs(CDbl(NzBIM(faktData(i, colIznos), 0#)) - uplata) < 0.01 Then
            hitCount = hitCount + 1
            hitID = CStr(faktData(i, colFID))
        End If
        
NextI:
    Next i
    
    If hitCount = 1 Then TryResolveFakturaForKupac = hitID
End Function

Private Function TryResolveOtkupForKooperant(ByVal bankaImportID As String, _
                                             ByVal kooperantID As String) As String
    Dim pozivNaBroj As String
    Dim otkData As Variant
    Dim colBrDok As Long, colOtkID As Long, colKoop As Long
    Dim i As Long
    Dim hitCount As Long
    Dim hitID As String
    
    pozivNaBroj = Trim$(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    If pozivNaBroj = "" Then Exit Function
    
    otkData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkData) Then Exit Function
    
    otkData = ExcludeStornirano(otkData, TBL_OTKUP)
    If IsEmpty(otkData) Then Exit Function
    
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    
    For i = 1 To UBound(otkData, 1)
        If CStr(otkData(i, colKoop)) <> kooperantID Then GoTo NextI
        
        If NormalizePozivKey(CStr(otkData(i, colBrDok))) = NormalizePozivKey(pozivNaBroj) Then
            hitCount = hitCount + 1
            hitID = CStr(otkData(i, colOtkID))
        End If
        
NextI:
    Next i
    
    If hitCount = 1 Then
        TryResolveOtkupForKooperant = hitID
    End If
End Function

Public Function GetOtkupCandidatesForKooperantBlock(ByVal kooperantID As String, _
                                                     ByVal brojBloka As String, _
                                                     Optional ByVal allowOverMax As Boolean = False) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim colOtkID As Long
    Dim colKoop As Long
    Dim colBrDok As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colVrsta As Long
    Dim i As Long
    Dim count As Long
    
    If Trim$(kooperantID) = "" Then Exit Function
    If Trim$(brojBloka) = "" Then Exit Function
    
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    
    ' AUD-025: bafer je velicine ulaza, ne fiksnih 2. Raniji `ReDim result(1 To 2)`
    ' + `If count > 2 Then Exit For` je ostavljao count=3 uz bafer od 2 reda, pa je
    ' kopiranje u finalResult padalo na "Subscript out of range" - a taj pad je iz
    ' AutoMapAll rollback-ovao CEO batch (svi vec mapirani redovi se ponistavali
    ' zbog jednog anomalnog bloka). Sada je 3+ kandidata eksplicitna, uhvatljiva
    ' greska (ERR_BMAP_MANUAL_REQUIRED) koja obara SAMO taj red.
    ReDim result(1 To UBound(data, 1), 1 To 3)

    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) <> kooperantID Then GoTo NextI

        If NormalizePozivKey(CStr(data(i, colBrDok))) = NormalizePozivKey(brojBloka) Then
            Dim vrednost As Double
            Dim uplaceno As Double
            Dim otvoreno As Double

            vrednost = 0
            If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
                vrednost = CDbl(data(i, colKol)) * CDbl(data(i, colCena))
            End If

            uplaceno = GetUplataForOtkup(CStr(data(i, colOtkID)))
            otvoreno = vrednost - uplaceno

            If otvoreno > 0.009 Then
                count = count + 1
                result(count, 1) = CStr(data(i, colOtkID))
                result(count, 2) = otvoreno
                result(count, 3) = CStr(data(i, colVrsta))
            End If
        End If

NextI:
    Next i

    If count = 0 Then Exit Function

    ' Granica vazi za AUTOMATSKU raspodelu. Rucni put je prosledjuje kao
    ' allowOverMax:=True tek posto je operater video kandidate i potvrdio podelu
    ' (inace bi red oznacen "za rucno" bio trajno nezavrsiv - ista greska bi ga
    ' docekala i na rucnom dugmetu).
    If count > MAX_BLOK_KANDIDATA And Not allowOverMax Then
        Err.Raise ERR_BMAP_MANUAL_REQUIRED, "GetOtkupCandidatesForKooperantBlock", _
            "Blok " & brojBloka & " (kooperant " & kooperantID & ") ima " & CStr(count) & _
            " otvorene otkupne stavke, a automatska raspodela dozvoljava najvise " & _
            CStr(MAX_BLOK_KANDIDATA) & ". Stavka izvoda se mapira RUCNO (izaberi " & _
            "kooperanta i blok, pa potvrdi predlozenu podelu)."
    End If

    Dim finalResult() As Variant
    Dim r As Long, c As Long

    ReDim finalResult(1 To count, 1 To 3)

    For r = 1 To count
        For c = 1 To 3
            finalResult(r, c) = result(r, c)
        Next c
    Next r

    SortKandidatiPoOtvorenomDesc finalResult

    GetOtkupCandidatesForKooperantBlock = finalResult
End Function

' Veci otvoreni iznos prvi (selection sort - lista je po prirodi kratka).
' Ranije je ovo bio swap koji je radio samo za tacno 2 reda.
Private Sub SortKandidatiPoOtvorenomDesc(ByRef kandidati As Variant)
    Dim i As Long, j As Long, best As Long, c As Long
    Dim tmp As Variant

    For i = LBound(kandidati, 1) To UBound(kandidati, 1) - 1
        best = i

        For j = i + 1 To UBound(kandidati, 1)
            If CDbl(NzBIM(kandidati(j, 2), 0#)) > CDbl(NzBIM(kandidati(best, 2), 0#)) Then best = j
        Next j

        If best <> i Then
            For c = 1 To 3
                tmp = kandidati(i, c)
                kandidati(i, c) = kandidati(best, c)
                kandidati(best, c) = tmp
            Next c
        End If
    Next i
End Sub

' Raspodela iznosa isplate po kandidatima bloka (greedy: veci otvoreni prvi, do
' iscrpljenja iznosa). JEDAN izvor istine - koriste ga i pisac
' (MapBankaImportAsKooperantBlockCore) i preview u frmBankaImport, pa operater
' pre klika vidi TACNO onu podelu koja ce biti proknjizena.
' Vraca (1 To n, 1 To 3): OtkupID | iznos za taj otkup | VrstaVoca. Redovi sa
' iznosom 0 se ne vracaju. Visak (iznos - suma plana) racuna pozivalac.
Public Function PlanBlokRaspodela(ByVal kandidati As Variant, ByVal iznos As Double) As Variant
    If IsEmpty(kandidati) Then Exit Function
    If iznos <= 0 Then Exit Function

    Dim plan() As Variant
    Dim i As Long
    Dim n As Long
    Dim preostalo As Double
    Dim otvoreno As Double
    Dim zaRed As Double

    ReDim plan(1 To UBound(kandidati, 1), 1 To 3)
    preostalo = iznos

    For i = 1 To UBound(kandidati, 1)
        If preostalo <= 0 Then Exit For

        otvoreno = CDbl(NzBIM(kandidati(i, 2), 0#))
        If otvoreno > 0 Then
            If preostalo >= otvoreno Then
                zaRed = otvoreno
            Else
                zaRed = preostalo
            End If

            n = n + 1
            plan(n, 1) = CStr(kandidati(i, 1))
            plan(n, 2) = zaRed
            plan(n, 3) = CStr(kandidati(i, 3))

            preostalo = preostalo - zaRed
        End If
    Next i

    If n = 0 Then Exit Function

    Dim finalPlan() As Variant
    Dim r As Long, c As Long

    ReDim finalPlan(1 To n, 1 To 3)

    For r = 1 To n
        For c = 1 To 3
            finalPlan(r, c) = plan(r, c)
        Next c
    Next r

    PlanBlokRaspodela = finalPlan
End Function

' ============================================================
' PUBLIC - ACCESS
' ============================================================

Public Function GetBankaImportRowByID(ByVal bankaImportID As String) As Variant
    Const SRC As String = "GetBankaImportRowByID"

    On Error GoTo EH

    Dim rowBim As Long
    rowBim = RequireSingleRow(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, SRC)

    Dim data As Variant
    data = GetTableData(TBL_BANKA_IMPORT)

    If IsEmpty(data) Then
        Err.Raise ERR_BMAP_BASE + 30, SRC, _
                  "Tabela je prazna: " & TBL_BANKA_IMPORT
    End If

    Dim colBrojDok As Long
    Dim colDatumTx As Long
    Dim colPartner As Long
    Dim colPartnerKonto As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colOpis As Long
    Dim colSvrha As Long
    Dim colRef As Long
    Dim colPozivNaBroj As Long

    colBrojDok = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, SRC)
    colDatumTx = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE, SRC)
    colPartner = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER, SRC)
    colPartnerKonto = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER_KONTO, SRC)
    colUplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA, SRC)
    colIsplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA, SRC)
    colOpis = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OPIS, SRC)
    colSvrha = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_SVRHA_PLACANJA, SRC)
    colRef = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ, SRC)
    colPozivNaBroj = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ, SRC)
    
    Dim result() As Variant
    ReDim result(1 To 1, 1 To 10)

    result(1, 1) = data(rowBim, colBrojDok)
    result(1, 2) = data(rowBim, colDatumTx)
    result(1, 3) = data(rowBim, colPartner)
    result(1, 4) = data(rowBim, colPartnerKonto)
    result(1, 5) = data(rowBim, colUplata)
    result(1, 6) = data(rowBim, colIsplata)
    result(1, 7) = data(rowBim, colOpis)
    result(1, 8) = data(rowBim, colSvrha)
    result(1, 9) = data(rowBim, colRef)
    result(1, 10) = data(rowBim, colPozivNaBroj)

    GetBankaImportRowByID = result
    Exit Function

EH:
    LogAndReraiseBankaMap SRC
End Function

Private Function FindBankaImportRowIndex(ByVal bankaImportID As String) As Long
    FindBankaImportRowIndex = RequireSingleRow( _
        TBL_BANKA_IMPORT, _
        COL_BIM_ID, _
        bankaImportID, _
        "FindBankaImportRowIndex" _
    )
End Function

Private Function ValidateBankaImportNotProcessed(ByVal bankaImportID As String) As Boolean
    Const SRC As String = "ValidateBankaImportNotProcessed"

    On Error GoTo EH

    Dim rowBim As Long
    rowBim = RequireSingleRow(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, SRC)

    Dim data As Variant
    data = GetTableData(TBL_BANKA_IMPORT)

    If IsEmpty(data) Then
        Err.Raise ERR_BMAP_BASE + 20, SRC, _
                  "Tabela je prazna: " & TBL_BANKA_IMPORT
    End If

    Dim colObr As Long
    Dim colStorno As Long
    Dim obrValue As String
    Dim stornoValue As String

    colObr = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO, SRC)
    colStorno = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO, SRC)

    obrValue = UCase$(Trim$(CStr(NzBIM(data(rowBim, colObr), ""))))
    stornoValue = UCase$(Trim$(CStr(NzBIM(data(rowBim, colStorno), ""))))

    If stornoValue = "DA" Then
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If

    If obrValue = "DA" Then
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If
    
    If obrValue = "SKIP" Then
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If

    ValidateBankaImportNotProcessed = True
    Exit Function

EH:
    LogAndReraiseBankaMap SRC
End Function

Private Sub UpdateBankaImportStatus(ByVal bankaImportID As String, _
                                    ByVal statusValue As String)
    Const SRC As String = "UpdateBankaImportStatus"

    On Error GoTo EH

    Dim rowBim As Long
    rowBim = RequireSingleRow(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, SRC)

    RequireUpdateCell TBL_BANKA_IMPORT, rowBim, COL_BIM_OBRADJENO, _
                      statusValue, SRC

    Exit Sub

EH:
    LogAndReraiseBankaMap SRC
End Sub


' ============================================================
' PUBLIC - ENTITY RESOLUTION
' ============================================================

Public Function TryResolveKupacBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colNaziv As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_KUPCI)
    If IsEmpty(data) Then
        TryResolveKupacBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KUPCI, "KupacID")
    colNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLooseBIM(CStr(data(i, colNaziv))) = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveKupacBIM = Array(hitID, "Kupac", "")
End Function

Public Function TryResolveKooperantBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colIme As Long
    Dim colPrezime As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    Dim hitOMID As String
    Dim fullName As String
    
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        TryResolveKooperantBIM = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_KOOPERANTI)
    If IsEmpty(data) Then
        TryResolveKooperantBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    
    For i = 1 To UBound(data, 1)
        fullName = NormalizeLooseBIM(CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime)))
        If fullName = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
            hitOMID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", hitID, COL_KOOP_STANICA), ""))
        End If
    Next i
    
    If hits = 1 Then TryResolveKooperantBIM = Array(hitID, "Kooperant", hitOMID)
End Function

Public Function TryResolveOMBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colNaziv As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then
        TryResolveOMBIM = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_STANICE)
    If IsEmpty(data) Then
        TryResolveOMBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLooseBIM(CStr(data(i, colNaziv))) = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveOMBIM = Array(hitID, "OM", hitID)
End Function


' ============================================================
' PRIVATE/PUBLIC - HELPERS
' ============================================================

' Broj izvoda je OBAVEZAN identitet novca nastalog iz izvoda: po njemu se grupise
' storno celog izvoda (izvod se ne stornira parcijalno). Banka uvek daje broj, pa
' je prazan broj greska uvoza/parsera - ne slucaj za fallback. Raniji fallback je
' sve stavke bez broja slivao u literal "IZVOD" i time spajao nepovezane izvode u
' jednu storno grupu (tiha korupcija identiteta).
' Otvoren (jos neplacen) iznos fakture = iznos - aktivne uplate. JEDAN izvor
' istine: koristi ga i lista faktura u frmBankaImport (prikaz) i writer
' (raspodela uplata/avans), pa prikazan i knjizen saldo ne mogu da se razidju.
Public Function GetOtvorenoNaFakturi(ByVal fakturaID As String) As Double
    Dim iznos As Double

    If Trim$(fakturaID) = "" Then Exit Function

    iznos = CDbl(nz(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS), "0"))
    GetOtvorenoNaFakturi = iznos - GetUplataForFaktura(fakturaID)
End Function

' FM-0023 #2: uplata se sme vezati SAMO za fakturu tog kupca i to nestorniranu.
' Do RF-09 je ovo bilo latentno (UI je uvek slao prazan FakturaID); sada
' frmBankaImport nudi izbor fakture, pa veza mora da se proveri na upisu, ne samo
' u listi koja se nudi.
Private Sub RequireFakturaZaKupca(ByVal fakturaID As String, _
                                  ByVal kupacID As String, _
                                  ByVal sourceName As String)
    Dim vlasnik As String
    Dim storno As String

    vlasnik = Trim$(CStr(NzBIM(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_KUPAC), "")))

    If vlasnik <> Trim$(kupacID) Then
        Err.Raise ERR_BMAP_BASE + 51, sourceName, _
            "Faktura " & fakturaID & " pripada kupcu '" & vlasnik & "', a uplata se " & _
            "vezuje za kupca '" & kupacID & "'. Mapiranje je odbijeno."
    End If

    storno = UCase$(Trim$(CStr(NzBIM(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_STORNIRANO), ""))))

    If storno = "DA" Then
        Err.Raise ERR_BMAP_BASE + 52, sourceName, _
            "Faktura " & fakturaID & " je stornirana - uplata se ne moze vezati za nju."
    End If
End Sub

' AUD-025 (FM-0023 #8) / AUD-014: smer stavke izvoda je deo identiteta knjizenja,
' a rucni put (frmBankaImport -> cmbMapTip) sme da posalje bilo koji tip na bilo
' koji red. Bez ove kapije se isplata kooperantu moze knjiziti kao uplata kupca
' (i obrnuto): iznos ode u pogresnu kolonu tblNovac, saldo partnera se pomeri u
' pogresnom smeru, a u izvodu nema traga da je red pogresno knjizen. Auto put vec
' grana po smeru (AutoMapBankaImportRow), pa mu guard ne menja ponasanje.
' expectedSmer: "UPLATA" | "ISPLATA" | "" (bilo koji, ali cist smer).
' Klasifikacija smera - jedini izvor istine (koriste ga i RequireBimSmer i preview
' u frmBankaImport). Red sa OBA iznosa (ili bez ijednog) nije knjizljiv.
Public Function ClassifyBimSmer(ByVal uplata As Double, ByVal isplata As Double) As String
    If uplata > 0 And isplata > 0 Then
        ClassifyBimSmer = BIM_SMER_NEJASAN
    ElseIf uplata > 0 Then
        ClassifyBimSmer = BIM_SMER_UPLATA
    ElseIf isplata > 0 Then
        ClassifyBimSmer = BIM_SMER_ISPLATA
    Else
        ClassifyBimSmer = BIM_SMER_NEJASAN
    End If
End Function

Private Sub RequireBimSmer(ByVal bim As Variant, _
                           ByVal expectedSmer As String, _
                           ByVal sourceName As String)
    Dim uplata As Double
    Dim isplata As Double
    Dim smer As String

    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))
    smer = ClassifyBimSmer(uplata, isplata)

    If smer = BIM_SMER_NEJASAN Then
        If uplata > 0 And isplata > 0 Then
            Err.Raise ERR_BMAP_BASE + 46, sourceName, _
                "Stavka izvoda ima i uplatu i isplatu - nema cist smer, mapiranje je odbijeno."
        End If

        Err.Raise ERR_BMAP_BASE + 47, sourceName, _
            "Stavka izvoda nema ni uplatu ni isplatu - mapiranje je odbijeno."
    End If

    Select Case UCase$(Trim$(expectedSmer))
        Case BIM_SMER_UPLATA
            If smer <> BIM_SMER_UPLATA Then
                Err.Raise ERR_BMAP_BASE + 48, sourceName, _
                    "Smer ne odgovara: stavka izvoda je ISPLATA (" & _
                    Format$(isplata, "#,##0.00") & "), a izabrano mapiranje ocekuje UPLATU. " & _
                    "Mapiranje je odbijeno."
            End If

        Case BIM_SMER_ISPLATA
            If smer <> BIM_SMER_ISPLATA Then
                Err.Raise ERR_BMAP_BASE + 49, sourceName, _
                    "Smer ne odgovara: stavka izvoda je UPLATA (" & _
                    Format$(uplata, "#,##0.00") & "), a izabrano mapiranje ocekuje ISPLATU. " & _
                    "Mapiranje je odbijeno."
            End If
    End Select
End Sub

Private Function RequireIzvodBroj(ByVal bim As Variant, ByVal sourceName As String) As String
    Dim broj As String
    broj = Trim$(CStr(NzBIM(bim(1, 1), "")))
    If Len(broj) = 0 Then
        Err.Raise ERR_BMAP_BASE + 45, sourceName, _
                  "Izvod nema broj dokumenta - mapiranje je odbijeno. " & _
                  "Proveri uvoz/parser izvoda (BankaImportID bez BrojDokumenta)."
    End If
    RequireIzvodBroj = broj
End Function

Private Function BuildBIMNapomena(ByVal bankaImportID As String, _
                                  ByVal bankaRef As String, _
                                  ByVal partnerKonto As String, _
                                  ByVal opis As String, _
                                  ByVal svrha As String, _
                                  ByVal reason As String) As String
    Dim s As String
    
    ' Prefiks je Public Const (modConfig): storno sloj po njemu prepoznaje da je
    ' novac red nastao iz izvoda -> ne sme se dirati bez tog citaoca.
    s = NOV_NAPOMENA_BIM_PREFIX & bankaImportID
    If Trim$(bankaRef) <> "" Then s = s & "; Ref:" & bankaRef
    If Trim$(partnerKonto) <> "" Then s = s & "; Konto:" & partnerKonto
    If Trim$(opis) <> "" Then s = s & "; Opis:" & Left$(opis, 80)
    If Trim$(svrha) <> "" Then s = s & "; Svrha:" & Left$(svrha, 120)
    If Trim$(reason) <> "" Then s = s & "; Match:" & Left$(reason, 50)
    
    BuildBIMNapomena = s
End Function

' Citac markera koji pise BuildBIMNapomena: vraca BankaImportID iz Napomene
' ("BIM:<id>; Ref:...") ili "" ako red nije nastao iz izvoda. Zivi uz pisca da
' format ima JEDNO mesto istine (citaju ga modStorno i modNovac).
Public Function BimIdFromNapomena(ByVal napomena As String) As String
    Dim s As String: s = LTrim$(CStr(napomena))
    Dim pfx As String: pfx = NOV_NAPOMENA_BIM_PREFIX
    If Len(s) < Len(pfx) Then Exit Function
    If UCase$(Left$(s, Len(pfx))) <> UCase$(pfx) Then Exit Function
    s = Mid$(s, Len(pfx) + 1)
    Dim p As Long: p = InStr(s, ";")
    If p > 0 Then s = Left$(s, p - 1)
    BimIdFromNapomena = Trim$(s)
End Function

' Napomena za avans-split red: nasledi BIM marker roditelja da split ostane
' pripadan SVOM izvodu (bez toga se poreklo gubi i storno izvoda mora da ga
' pogadja po broju+partneru -> vidi TL-007). Rucni avans nema marker -> samo opis.
Public Function BuildAvansSplitNapomena(ByVal parentNapomena As String, _
                                        ByVal opis As String) As String
    Dim bimID As String: bimID = BimIdFromNapomena(parentNapomena)
    If Len(bimID) = 0 Then
        BuildAvansSplitNapomena = opis
    Else
        BuildAvansSplitNapomena = NOV_NAPOMENA_BIM_PREFIX & bimID & "; " & opis
    End If
End Function

Public Function NormalizeLooseBIM(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, vbTab, " ")
    s = Replace(s, ",", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeLooseBIM = Trim$(s)
End Function

Public Function NzBIM(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If isError(v) Then
        NzBIM = Fallback
    ElseIf IsNull(v) Then
        NzBIM = Fallback
    ElseIf IsEmpty(v) Then
        NzBIM = Fallback
    ElseIf Trim$(CStr(v)) = "" Then
        NzBIM = Fallback
    Else
        NzBIM = v
    End If
End Function

' ============================================================
' PRIORITETNI RESOLVERI: tekuci racun + poziv na broj
' Jaki poslovni kljucevi za auto-map. Sve je single-hit
' (0 ili >1 pogodaka -> ne auto). Reuse Map*/SaveNovac putanja.
' ============================================================

Public Function TryResolveKupacByKonto(ByVal partnerKonto As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colRac As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    Dim target As String

    TryResolveKupacByKonto = Empty

    target = NormalizeKonto(partnerKonto)
    If target = "" Then Exit Function

    colRac = GetColumnIndex(TBL_KUPCI, COL_KUP_TEKUCI_RACUN)
    If colRac = 0 Then Exit Function

    data = GetTableData(TBL_KUPCI)
    If IsEmpty(data) Then Exit Function

    colID = GetColumnIndex(TBL_KUPCI, COL_KUP_ID)

    For i = 1 To UBound(data, 1)
        If NormalizeKonto(CStr(data(i, colRac))) = target Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i

    If hits = 1 Then TryResolveKupacByKonto = Array(hitID, "Kupac", "")
End Function

Public Function TryResolveKooperantByKonto(ByVal partnerKonto As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colRac As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    Dim hitOMID As String
    Dim target As String

    TryResolveKooperantByKonto = Empty

    target = NormalizeKonto(partnerKonto)
    If target = "" Then Exit Function

    colRac = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_TEKUCI_RACUN)
    If colRac = 0 Then Exit Function

    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    colID = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)

    For i = 1 To UBound(data, 1)
        If NormalizeKonto(CStr(data(i, colRac))) = target Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i

    If hits = 1 Then
        hitOMID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, hitID, COL_KOOP_STANICA), ""))
        TryResolveKooperantByKonto = Array(hitID, "Kooperant", hitOMID)
    End If
End Function

' Poziv na broj -> otkupni list (BrDok) -> KooperantID.
' Vise otkupa istog bloka pripada istom kooperantu = i dalje jedan kooperant.
Public Function TryResolveKooperantByOtkupPoziv(ByVal pozivNaBroj As String) As String
    Dim data As Variant
    Dim colBrDok As Long
    Dim colKoop As Long
    Dim i As Long
    Dim hits As Long
    Dim hitKoop As String
    Dim curKoop As String
    Dim target As String

    target = NormalizePozivKey(pozivNaBroj)
    If target = "" Then Exit Function

    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)

    For i = 1 To UBound(data, 1)
        If NormalizePozivKey(CStr(data(i, colBrDok))) = target Then
            curKoop = CStr(data(i, colKoop))
            If hitKoop = "" Then
                hits = 1
                hitKoop = curKoop
            ElseIf UCase$(curKoop) <> UCase$(hitKoop) Then
                hits = hits + 1
            End If
        End If
    Next i

    If hits = 1 Then TryResolveKooperantByOtkupPoziv = hitKoop
End Function

' Poziv na broj -> Broj fakture -> Array(FakturaID, KupacID); Empty ako nije jednoznacno.
Public Function TryResolveFakturaByPoziv(ByVal pozivNaBroj As String) As Variant
    Dim data As Variant
    Dim colFID As Long
    Dim colBroj As Long
    Dim colKup As Long
    Dim i As Long
    Dim hits As Long
    Dim hitFID As String
    Dim hitKup As String
    Dim target As String
    Dim k As String

    TryResolveFakturaByPoziv = Empty

    target = NormalizePozivKey(pozivNaBroj)
    If target = "" Then Exit Function

    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then Exit Function

    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)

    For i = 1 To UBound(data, 1)
        k = NormalizePozivKey(CStr(data(i, colBroj)))
        If k <> "" And k = target Then
            hits = hits + 1
            hitFID = CStr(data(i, colFID))
            hitKup = CStr(data(i, colKup))
        End If
    Next i

    If hits = 1 Then TryResolveFakturaByPoziv = Array(hitFID, hitKup)
End Function

' Kanonski tekuci racun: 3 (banka) + 13 (partija, vodece nule) + 2 (kontrola) = 18 cifara.
' Banka salje 205-0000000420902-31; operater u sifarnik moze uneti sa/bez crtica i nula.
Public Function NormalizeKonto(ByVal s As String) As String
    Dim raw As String
    Dim onlyDigits As String
    Dim ch As String
    Dim i As Long
    Dim groups As Collection
    Dim g1 As String
    Dim g2 As String
    Dim g3 As String

    raw = Trim$(s)
    If Len(raw) = 0 Then Exit Function

    Set groups = SplitDigitGroups(raw)

    If groups.count = 3 Then
        g1 = CStr(groups(1))
        g2 = CStr(groups(2))
        g3 = CStr(groups(3))
        If Len(g1) <= 3 And Len(g2) <= 13 And Len(g3) <= 2 Then
            g1 = Right$("000" & g1, 3)
            g2 = Right$(String$(13, "0") & g2, 13)
            g3 = Right$("00" & g3, 2)
            NormalizeKonto = g1 & g2 & g3
            Exit Function
        End If
    End If

    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch Like "[0-9]" Then onlyDigits = onlyDigits & ch
    Next i
    NormalizeKonto = onlyDigits
End Function

Private Function SplitDigitGroups(ByVal s As String) As Collection
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim cur As String

    Set result = New Collection

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9]" Then
            cur = cur & ch
        Else
            If Len(cur) > 0 Then
                result.Add cur
                cur = ""
            End If
        End If
    Next i
    If Len(cur) > 0 Then result.Add cur

    Set SplitDigitGroups = result
End Function

' Kljuc za poredjenje poziva-na-broj sa BrDok otkupa / Broj fakture.
' Skida [NN] model (npr. [97]) i sve ne-alfanumericke znakove; UPPERCASE.
' "[97]182106264" -> "182106264"; "18/210626-4" -> "182106264".
Public Function NormalizePozivKey(ByVal s As String) As String
    Dim r As String
    Dim out As String
    Dim i As Long
    Dim ch As String
    Dim p1 As Long
    Dim p2 As Long

    r = UCase$(Trim$(s))
    If Len(r) = 0 Then Exit Function

    Do
        p1 = InStr(r, "[")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, r, "]")
        If p2 = 0 Then Exit Do
        r = Left$(r, p1 - 1) & Mid$(r, p2 + 1)
    Loop

    For i = 1 To Len(r)
        ch = Mid$(r, i, 1)
        If ch Like "[0-9A-Z]" Then out = out & ch
    Next i

    NormalizePozivKey = out
End Function

Public Function GetKooperantNaziv(ByVal kooperantID As String) As String
    GetKooperantNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime")) & _
                        " " & _
                        CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
End Function

Private Sub Monitor_BankaMapSuccess(ByVal procedureName As String, _
                                    ByVal bankaImportID As String, _
                                    ByVal resultId As String, _
                                    ByVal partnerType As String, _
                                    Optional ByVal partnerId As String = "", _
                                    Optional ByVal linkedEntityId As String = "")
    On Error Resume Next

    Monitor_Event _
        eventType:="BANKA_MAP_SUCCESS", _
        severity:="INFO", _
        message:="Bank import mapping succeeded. Type=" & partnerType & _
                 "; BankaImportID=" & bankaImportID & _
                 "; ResultID=" & resultId & _
                 "; PartnerID=" & partnerId & _
                 "; LinkedEntityID=" & linkedEntityId, _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:=procedureName, _
        entityType:="BankaImport", _
        entityID:=bankaImportID, _
        correlationId:=bankaImportID
End Sub

Private Sub Monitor_BankaMapFail(ByVal procedureName As String, _
                                 ByVal bankaImportID As String, _
                                 ByVal errNum As Long, _
                                 ByVal errDesc As String, _
                                 ByVal errSrc As String, _
                                 Optional ByVal partnerType As String = "", _
                                 Optional ByVal partnerId As String = "", _
                                 Optional ByVal linkedEntityId As String = "")
    On Error Resume Next

    Monitor_Error _
        moduleName:="modBankaMapiranje", _
        procedureName:=procedureName, _
        entityType:="BankaImport", _
        entityID:=bankaImportID, _
        correlationId:=bankaImportID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="BANKA_MAP_FAIL", _
        severity:="ERROR", _
        message:="Bank import mapping failed. Type=" & partnerType & _
                 "; BankaImportID=" & bankaImportID & _
                 "; PartnerID=" & partnerId & _
                 "; LinkedEntityID=" & linkedEntityId & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modBankaMapiranje", _
        procedureName:=procedureName, _
        entityType:="BankaImport", _
        entityID:=bankaImportID, _
        correlationId:=bankaImportID
End Sub

Private Sub LinkNovacToOtkupStrict(ByVal novacID As String, _
                                   ByVal otkupID As String, _
                                   ByVal sourceName As String)
    Const SRC As String = "LinkNovacToOtkupStrict"

    On Error GoTo EH

    If Len(Trim$(sourceName)) = 0 Then sourceName = SRC

    Dim rowNov As Long
    Dim rowOtk As Long

    rowNov = RequireSingleRow(TBL_NOVAC, COL_NOV_ID, novacID, sourceName)
    rowOtk = RequireSingleRow(TBL_OTKUP, COL_OTK_ID, otkupID, sourceName)

    RequireUpdateCell TBL_NOVAC, rowNov, COL_NOV_OTKUP_ID, otkupID, sourceName

    UpdateOtkupStatus otkupID

    Exit Sub

EH:
    LogAndReraiseBankaMap SRC
End Sub

Private Function RequireSingleRow(ByVal tblName As String, _
                                  ByVal idColumn As String, _
                                  ByVal idValue As String, _
                                  ByVal sourceName As String) As Long
    If Len(Trim$(tblName)) = 0 Then
        Err.Raise ERR_BMAP_BASE + 1, sourceName, "TableName je obavezan."
    End If

    If Len(Trim$(idColumn)) = 0 Then
        Err.Raise ERR_BMAP_BASE + 2, sourceName, "IdColumn je obavezan."
    End If

    If Len(Trim$(idValue)) = 0 Then
        Err.Raise ERR_BMAP_BASE + 3, sourceName, _
                  "ID vrednost je obavezna. Table=" & tblName & _
                  " Column=" & idColumn
    End If

    RequireColumnIndex tblName, idColumn, sourceName

    Dim rows As Collection
    Set rows = FindRows(tblName, idColumn, idValue)

    If rows Is Nothing Then
        Err.Raise ERR_BMAP_BASE + 4, sourceName, _
                  "FindRows je vratio Nothing. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue
    End If
    
    If rows.count = 0 Then
        Err.Raise ERR_BMAP_BASE + 5, sourceName, _
                  "Missing link. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue
    End If

    If rows.count > 1 Then
        Err.Raise ERR_BMAP_BASE + 6, sourceName, _
                  "Duplicate key. Table=" & tblName & _
                  " Column=" & idColumn & _
                  " ID=" & idValue & _
                  " Count=" & CStr(rows.count)
    End If

    RequireSingleRow = CLng(rows(1))
End Function

Private Sub LogAndReraiseBankaMap(ByVal sourceName As String)
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
' TESTS
' ============================================================

Public Sub Test_GetBankaImportOpen()
    Dim data As Variant
    Dim i As Long
    Dim colID As Long
    
    data = GetBankaImportOpen()
    If IsEmpty(data) Then
        Debug.Print "No open rows."
        Exit Sub
    End If
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    
    For i = 1 To UBound(data, 1)
        Debug.Print data(i, colID)
    Next i
End Sub

Public Sub Test_MapBankaImportAsKooperantBlock_TX()
    Dim bankaImportID As String
    Dim kooperantID As String
    Dim n As Long
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kooperantID = InputBox$("KooperantID:")
    If Trim$(kooperantID) = "" Then Exit Sub
    
    n = MapBankaImportAsKooperantBlock_TX(bankaImportID, kooperantID, True)
    Debug.Print "Created rows: " & n
End Sub

Public Sub Test_AutoMapBankaImportRow_TX()
    Dim bankaImportID As String
    Dim novacID As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    novacID = AutoMapBankaImportRow_TX(bankaImportID)
    Debug.Print "Created NovacID: " & novacID
End Sub

Public Sub Test_MapBankaImportAsKupac_TX()
    Dim bankaImportID As String
    Dim kupacID As String
    Dim fakturaID As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kupacID = InputBox$("KupacID:")
    If Trim$(kupacID) = "" Then Exit Sub
    
    fakturaID = InputBox$("FakturaID (optional):")
    
    Debug.Print MapBankaImportAsKupac_TX(bankaImportID, kupacID, fakturaID, True)
End Sub

Public Sub Test_MapBankaImportAsKooperant_TX()
    Dim bankaImportID As String
    Dim kooperantID As String
    Dim otkupID As String
    Dim vrsta As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kooperantID = InputBox$("KooperantID:")
    If Trim$(kooperantID) = "" Then Exit Sub
    
    otkupID = InputBox$("OtkupID (optional):")
    vrsta = InputBox$("VrstaVoca (optional):")
    
    Debug.Print MapBankaImportAsKooperant_TX(bankaImportID, kooperantID, otkupID, vrsta, True)
End Sub

Public Sub Test_MapBankaImportAsOM_TX()
    Dim bankaImportID As String
    Dim omID As String
    Dim vrsta As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    omID = InputBox$("OMID:")
    If Trim$(omID) = "" Then Exit Sub
    
    vrsta = InputBox$("VrstaVoca (optional):")
    
    Debug.Print MapBankaImportAsOM_TX(bankaImportID, omID, vrsta, True)
End Sub

Public Sub Test_SkipBankaImportRow_TX()
    Dim bankaImportID As String
     
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    If SkipBankaImportRow_TX(bankaImportID) Then
        Debug.Print "Skipped: " & bankaImportID
    End If
End Sub

