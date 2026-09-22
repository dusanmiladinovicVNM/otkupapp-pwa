Attribute VB_Name = "modDokumentInvariant"
Option Explicit

' ============================================================
' modDokumentInvariant - centralni invariant engine
'
' Kljucni poslovni invariant:
' S4-3a JE ODAVDE SKLONIO CEO ZBIRNA INVARIJANT.
'
' Merio je "zbirna = zbir svojih aktivnih otpremnica PO BrojZbirne" naspram
' zaglavlja tblZbirna. Pod kanonom nijedna strana ne postoji: clanstvo je zapis
' (tblZbirnaIzvori, po ZbirnaID), sadrzaj je na stavkama, a pisac zaglavlje
' NAMERNO ostavlja prazno. Obe strane su bile nule, pa je racun uvek javljao
' "OK" -- provera koja ne moze da padne nije provera.
'
' Sa njim je otisla i rekalkulacija u mestu: po ZBR-KANON-03 se izmenjena zbirna
' ne prepravlja nego dobija NOVU VERZIJU (storno + nova iz istih izvora).
'
' Ostalo je ono sto i dalje meri nesto: SumOtpremniceByKlasa (cita ga golden i
' BFP), FindSingleActiveRow, DocIsIssued, SetIzdatoStatus.
'
' Stil: reuse modDataAccess (GetTableData/GetColumnIndex/RequireUpdateCell),
' clsTransaction za mutacije, LogErr/Monitor_Event za greske. Bez MsgBox
' (business sloj). Sve mutacije u *_TX funkciji.
' ============================================================

Private Const MOD_NAME As String = "modDokumentInvariant"
Private Const EPS_KG As Double = 0.01

' Test-observability seam: Monitor_Event je HTTP (nema lokalni red) i moze biti
' iskljucen, pa se emisija audita ne moze asertovati direktno. AuditIssuedZbirnaChange
' usput postavlja ovaj marker (delta poslednjeg audita izdate zbirne) da regres-test
' potvrdi da je gate-putanja (izdato + promena) stvarno prosla. Ne utice na ponasanje.
Private mLastIssuedZbirnaAudit As String

' ============================================================
' Per-klasa suma AKTIVNIH otpremnica za dati BrojZbirne.
' Vraca Scripting.Dictionary sa kljucevima:
'   kgI, kgII, kgOther, kgTotal
'   ambI, ambII, ambOther, ambTotal
'   nRows, nRowsI, nRowsII
'   vrstaI, sortaI, tipAmbI (reprezentativna zaglavlja za klasu I)
'   vrstaII, sortaII, tipAmbII
' ============================================================
Public Function SumOtpremniceByKlasa(ByVal brojZbirne As String) As Object
    Const SRC As String = MOD_NAME & ".SumOtpremniceByKlasa"
    Dim d As Object
    Set d = NewSumDict()
    Set SumOtpremniceByKlasa = d

    On Error GoTo EH
    brojZbirne = Trim$(brojZbirne)
    If Len(brojZbirne) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Function

    Dim cZbr As Long, cId As Long, cStorno As Long
    Dim cVrsta As Long, cSorta As Long, cTipAmb As Long
    cZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    cVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    cSorta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA)
    cTipAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB)

    ' Kilaza i gajbe dolaze sa STAVKI otpremnice (S3b). Zaglavlje ih od S3a ne
    ' nosi -- invariant koji bi ih odatle citao poredio bi zbirnu sa NULOM i
    ' proglasavao svaku zbirnu neispravnom. Vrsta, sorta i tip ambalaze ostaju
    ' na zaglavlju, pa CaptureHeader i dalje dobija red zaglavlja.
    '
    ' nRows broji STAVKE, ne zaglavlja: druga strana poredjenja (tblZbirna) je
    ' i dalje jedan red po klasi, pa je par "stavka <-> red zbirne".
    Dim stavkeDok As Object
    Set stavkeDok = modDokumenta.StavkeOtpremnicePoDokumentu()

    Dim i As Long, s As Long, klasa As String
    Dim kol As Double, amb As Long
    Dim stavke As Collection, stavka As Variant
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = brojZbirne Then
            If Not IsDaFlag(data(i, cStorno)) Then
                Set stavke = modDokumenta.StavkeZaOtpremnicu(stavkeDok, _
                                 Trim$(NzToText(data(i, cId))), SRC)
                For s = 1 To stavke.count
                    stavka = stavke(s)
                    klasa = Trim$(CStr(stavka(3)))
                    kol = CDbl(stavka(4))
                    amb = CLng(stavka(6))

                    d("kgTotal") = CDbl(d("kgTotal")) + kol
                    d("ambTotal") = CLng(d("ambTotal")) + amb
                    d("nRows") = CLng(d("nRows")) + 1

                    If klasa = KLASA_I Then
                        d("kgI") = CDbl(d("kgI")) + kol
                        d("ambI") = CLng(d("ambI")) + amb
                        d("nRowsI") = CLng(d("nRowsI")) + 1
                        CaptureHeader d, "I", data, i, cVrsta, cSorta, cTipAmb
                    ElseIf klasa = KLASA_II Then
                        d("kgII") = CDbl(d("kgII")) + kol
                        d("ambII") = CLng(d("ambII")) + amb
                        d("nRowsII") = CLng(d("nRowsII")) + 1
                        CaptureHeader d, "II", data, i, cVrsta, cSorta, cTipAmb
                    Else
                        d("kgOther") = CDbl(d("kgOther")) + kol
                        d("ambOther") = CLng(d("ambOther")) + amb
                    End If
                Next s
            End If
        End If
    Next i
    Exit Function
EH:
    ' Greska se PROPUSTA, ne guta (S3b-1). Progutana je ostavljala recnik sa
    ' nulama, a pozivalac ga je uzimao kao tacan zbir: "0 kg" bez ijedne greske.
    ' Pozivalac koji je na tome upisivao u zbirnu je obrisan u S4-3a, ali
    ' pravilo ostaje -- prazan recnik i procitan recnik ne smeju da izgledaju
    ' isto.
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' ============================================================
' HELPERS
' ============================================================

Private Function NewSumDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("kgI") = 0#: d("kgII") = 0#: d("kgOther") = 0#: d("kgTotal") = 0#
    d("ambI") = 0&: d("ambII") = 0&: d("ambOther") = 0&: d("ambTotal") = 0&
    d("nRows") = 0&: d("nRowsI") = 0&: d("nRowsII") = 0&
    d("vrstaI") = "": d("sortaI") = "": d("tipAmbI") = ""
    d("vrstaII") = "": d("sortaII") = "": d("tipAmbII") = ""
    Set NewSumDict = d
End Function

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Sub CaptureHeader(ByRef d As Object, ByVal klasa As String, ByRef data As Variant, _
                          ByVal rowIdx As Long, ByVal cVrsta As Long, _
                          ByVal cSorta As Long, ByVal cTipAmb As Long)
    ' Zapamti prvu ne-praznu vrednost zaglavlja po klasi (za kreiranje reda u recalc).
    If Len(CStr(d("vrsta" & klasa))) = 0 And cVrsta > 0 Then d("vrsta" & klasa) = NzTx(data(rowIdx, cVrsta))
    If Len(CStr(d("sorta" & klasa))) = 0 And cSorta > 0 Then d("sorta" & klasa) = NzTx(data(rowIdx, cSorta))
    If Len(CStr(d("tipAmb" & klasa))) = 0 And cTipAmb > 0 Then d("tipAmb" & klasa) = NzTx(data(rowIdx, cTipAmb))
End Sub

Private Function IsDaFlag(ByVal v As Variant) As Boolean
    IsDaFlag = (UCase$(Trim$(CStr(v))) = "DA")
End Function

Private Function NzTx(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzTx = ""
    Else
        NzTx = Trim$(CStr(v))
    End If
End Function

' ============================================================
' Faza 7 (3.0) - KANONSKO ADRESIRANJE append-only modela.
' Linijski model: jedan poslovni broj ima vise redova (po klasi). Identitet reda =
' (broj, klasa) -> AKTIVAN red. Vrati indeks tog reda:
'   0  = nema aktivnog reda za (broj, klasa)
'   -1 = VISE aktivnih (integritet povreda; u append-only sme najvise jedan)
' klasa == "" -> ignorisi klasu (match samo po broju; ambiguo kad ima vise klasa).
' Osnov za: PWA sync migraciju (3.1), append-only re-verziju (3.2), citace (3.3).
' ============================================================
Public Function FindSingleActiveRow(ByVal tbl As String, ByVal brojCol As String, _
        ByVal broj As String, ByVal klasaCol As String, ByVal klasa As String) As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long: cBr = GetColumnIndex(tbl, brojCol)
    If cBr = 0 Then Exit Function
    Dim cSt As Long: cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    Dim cKl As Long: cKl = 0
    If Len(klasaCol) > 0 Then cKl = GetColumnIndex(tbl, klasaCol)
    broj = Trim$(broj): klasa = Trim$(klasa)
    Dim i As Long, found As Long, cnt As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = broj Then
            If cSt = 0 Or Not IsDaFlag(data(i, cSt)) Then
                Dim klMatch As Boolean: klMatch = True
                If Len(klasa) > 0 And cKl > 0 Then klMatch = (Trim$(CStr(data(i, cKl))) = klasa)
                If klMatch Then found = i: cnt = cnt + 1
            End If
        End If
    Next i
    If cnt > 1 Then FindSingleActiveRow = -1 Else FindSingleActiveRow = found
    Exit Function
EH:
    LogErr MOD_NAME & ".FindSingleActiveRow"
End Function

' ============================================================
' Faza 7 - IzdatoStatus gate (ADR-0001 granica: izdat/prosledjen dokument je
' nepromenljiv -> koriguje se storno+reizdaj, ne in-place).
' Ova app NEMA "draft" fazu za chain dokumente -> IzdatoStatus je podrazumevano
' IZDATO; prazno = IZDATO (konzervativno). DRAFT je rezervisan (buduci parkiran/
' held dokument), PROSLEDJENO za buduci sync-push ka PWA/kupcu.
' DocIsIssued: True ako je IZDAT/PROSLEDJEN, False SAMO ako eksplicitno DRAFT.
' ============================================================
' #7: IzdatoStatus se cita sa AKTIVNOG reda (LookupActiveID), ne sa bilo kog reda
' (LookupValue je mogao pokupiti STORNIRAN red istog broja i procitati njegov status).
Public Function DocIsIssued(ByVal tbl As String, ByVal brojCol As String, ByVal broj As String) As Boolean
    On Error Resume Next
    DocIsIssued = True                                  ' default: izdato (konzervativno)
    If GetColumnIndex(tbl, COL_TRACE_IZDATO_STATUS) = 0 Then Exit Function
    Dim v As String
    v = UCase$(Trim$(LookupActiveID(tbl, brojCol, broj, COL_TRACE_IZDATO_STATUS)))
    DocIsIssued = (v <> UCase$(IZDATO_DRAFT))
End Function

' Postavi IzdatoStatus na jedan red (buduci prelazi: PROSLEDJENO pri sync-u ka PWA).
' Guarded na kolonu (schema-drift safe).
Public Sub SetIzdatoStatus(ByVal tbl As String, ByVal rowIndex As Long, ByVal status As String)
    On Error Resume Next
    If GetColumnIndex(tbl, COL_TRACE_IZDATO_STATUS) = 0 Then Exit Sub
    UpdateCell tbl, rowIndex, COL_TRACE_IZDATO_STATUS, status
End Sub

' ============================================================
' TEST (Alt+F8) - rollback-safe (clsTransaction snapshot + rollback; fixture ne
' ostaje). Automatske asertacije -> Debug.Print (Ctrl+G). Faza 7 (3.0).
' ============================================================
Public Sub Test_FindSingleActiveRow()
    Dim tx As clsTransaction
    Dim ok As Boolean: ok = True
    On Error GoTo EH
    Const B As String = "SVT-FSAR-Z"
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    ' fixture: KlasaI aktivan, KlasaII aktivan, KlasaI storniran (3 reda, isti broj).
    FsarSeed B, "I", ""
    FsarSeed B, "II", ""
    FsarSeed B, "I", "Da"

    Dim rI As Long: rI = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "I")
    Dim rII As Long: rII = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "II")
    Dim rAmb As Long: rAmb = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "")
    Dim rNone As Long: rNone = FindSingleActiveRow(TBL_ZBIRNA, COL_ZBR_BROJ, B, COL_ZBR_KLASA, "III")

    ok = FsarChk(rI > 0, "KlasaI -> jedan aktivan red (" & rI & ")") And ok
    ok = FsarChk(rII > 0, "KlasaII -> jedan aktivan red (" & rII & ")") And ok
    ok = FsarChk(rI <> rII, "KlasaI != KlasaII (razliciti redovi)") And ok
    Dim dchk As Variant: dchk = GetTableData(TBL_ZBIRNA)
    Dim cSt As Long: cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    ok = FsarChk(rI > 0 And Not IsDaFlag(dchk(rI, cSt)), "KlasaI red je AKTIVAN (ne storniran)") And ok
    ok = FsarChk(rAmb = -1, "klasa='' + 2 aktivna -> -1 (ambiguous)") And ok
    ok = FsarChk(rNone = 0, "nepostojeca klasa -> 0") And ok

    tx.RollbackTx: Set tx = Nothing
    Debug.Print "=== Test_FindSingleActiveRow: " & IIf(ok, "PROSAO", "PAO") & " (fixture rollback-ovan) ==="
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "Test_FindSingleActiveRow GRESKA: " & Err.description
End Sub

Private Sub FsarSeed(ByVal broj As String, ByVal klasa As String, ByVal storno As String)
    Dim lo As ListObject: Set lo = GetTable(TBL_ZBIRNA)
    If lo Is Nothing Then Exit Sub
    Dim nr As ListRow: Set nr = lo.ListRows.Add
    Dim ri As Long: ri = nr.Index
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_ID, "SVT-FSAR-" & klasa & "-" & IIf(Len(storno) > 0, "S", "A")
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_BROJ, broj
    UpdateCell TBL_ZBIRNA, ri, COL_ZBR_KLASA, klasa
    If Len(storno) > 0 Then UpdateCell TBL_ZBIRNA, ri, COL_STORNIRANO, storno
End Sub

Private Function FsarChk(ByVal cond As Boolean, ByVal nm As String) As Boolean
    Debug.Print IIf(cond, "OK   ", "FAIL ") & nm
    FsarChk = cond
End Function
