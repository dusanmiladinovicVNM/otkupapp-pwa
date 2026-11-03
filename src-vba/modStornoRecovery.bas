Attribute VB_Name = "modStornoRecovery"
Option Explicit

' Vrsta reda liste Nedovrseno kad se izvor ne moze procitati (S8R1B-REV-01): red
' se PRIKAZUJE, umesto da lista neprimetno bude kraca.
Public Const KIND_GRESKA_CITANJA As String = "GRESKA_CITANJA"

' TEST SEAM (S8 rez 1b, S8R1B-REV-01): "ID_KOLONA" -- lista osirotelih paleta dobija
' necitljivu kolonu identiteta. Postavljac ga pali samo u test rezimu, kuka ponovo
' pita IsTestMode(), a gasi ga modTest.ResetSeamova.
Private mNedovrsenoKvarTest As String

' ============================================================
' modStornoRecovery - "Nedovrseno" (jedan pregled) + "Vrati storno" (reverzija).
' Faza 5.
'
' GetNedovrseno (READ-ONLY): ujedinjen pregled onoga sto ceka -> PENDING/MANUAL
'   contexti (tblStornoVeze, strukturirano) + brojevi osirocenih (prijemnice /
'   palete / izgubljeni blokovi; detalj i akcije su u recovery panelu).
'
' UndoStorno_TX (MUTIRA): "Vrati storno" po (tip, broj) -- SAMO zurnal-put:
'   poslednja operacija za taj broj (LatestOpFor) ide kroz UndoOperation_TX, sa
'   istim gardama kao ekran Oporavak. Zovu ga Alt+F8 hook Test_UndoStorno i testovi.
'   Legacy put bez zurnala (Otkup i Revers po nogama) je obrisan u 10e-2: skidao je
'   Stornirano bez provere kontra-stava u knjizi (AMB-10-ODL-25).
'
' Ekran Oporavak cilja konkretan OperationID (modScrOporavak.VratiStorno ->
' modStornoZurnal.UndoOperation_TX); garde su fail-closed.
' ============================================================

Private Const MOD_NAME As String = "modStornoRecovery"
Private Const ERR_REC_BASE As Long = vbObjectError + 2900

' ============================================================
' NEDOVRSENO - ujedinjen read-only pregled.
' ============================================================
' Svaki red nosi PUN strukturirani zapis (ne samo tekst) -> panel moze da deluje po
' redu. Kljucevi: kind, ref (poslovni broj), status, opis, akcija (tekst), correctionID,
' docType, newBroj, mode, actionCode (CONTEXT | PRIJ | PAL | BLOK | LANAC -> UI
' dispecuje), otkupID (samo LANAC: identitet za radnju "Ponovi auto-lanac").
' Dedup PO IDENTITETU (S8 rez 1b, N9): osirocen red se izostavi samo kad PENDING/
' MANUAL context nosi BAS njegov dokument (StorniraniDocID), ili kad je isti
' dokument vec u listi. Do 1b se dedup radio po poslovnom BROJU, a broj nose
' dokumenti vise kupaca (i zbirna i prijemnica imaju oblik x/ddmmyy): context
' ispravke jednog kupca je sakrivao osirotele palete drugog kupca istog broja, a
' drugi osiroteli dokument istog broja nije ni dolazio u listu. Kad pripadnost
' nije dokaziva po ID-u, red se PRIKAZUJE -- isti problem dvaput je bezbedniji od
' sakrivenog tudjeg.
Public Function GetNedovrseno() As Collection
    Dim result As New Collection
    Set GetNedovrseno = result
    On Error GoTo EH

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    ' 1) Persistentni contexti (tblStornoVeze) - strukturirano, sa CorrectionID.
    Dim ctx As Collection: Set ctx = GetPendingCorrections()
    If Not ctx Is Nothing Then
        Dim i As Long
        For i = 1 To ctx.count
            Dim c As Object: Set c = ctx(i)
            Dim oB As String: oB = CStr(c("oldBroj"))
            AddNedRowFull result, "CONTEXT/" & CStr(c("mode")), oB, CStr(c("status")), _
                CStr(c("message")), _
                CStr(c("recoveryAction")), _
                CStr(c("id")), CStr(c("oldDocType")), CStr(c("newBroj")), CStr(c("mode")), "CONTEXT"
            ' dokument koji context pokriva -> njegov osiroceni red preskoci (dedup).
            Dim oId As String: oId = StorniraniDocIDTiho(CStr(c("id")))
            If Len(oId) > 0 Then seen("ID|" & oId) = True
        Next i
    End If

    ' 2) Osirocene - PER-STAVKA red (poslovni broj + akcija), deduplikovano protiv contexta.
    '    Akcija svih: "Osiroceni dokumenti" panel (postojeci re-point / skidanje).
    AddOsiroceneRows result, GetOsirocenePrijemnice(), "OSIROCENA_PRIJEMNICA", 1, 8, _
        "Prijemnica", "PRIJ", "Osiroceni dokumenti (prevezi prijemnicu)", seen
    AddOsiroceneRows result, GetPrijemniceSaOsirocenimPaletama(), "OSIROCENE_PALETE", 1, _
        IIf(NedovrsenoKvarAkoTest("ID_KOLONA"), 0, 7), _
        "Prijemnica (palete)", "PAL", "Osiroceni dokumenti (prevezi ili skini palete)", seen
    ' 3) Izgubljeni blokovi (B-041) -- KANON, ne stara veza.
    AddIzgubljeniBlokovi result
    ' 4) Nedovrseni hladnjacki lanci (S6 korak 7, odluka 8).
    AddNedovrseniLanci result
    Exit Function
EH:
    LogErr MOD_NAME & ".GetNedovrseno"
End Function

' Blokovi koje je STORNO otpremnice ostavio nevezane (B-041).
'
' "Izgubljen" nije svaki blok bez otpremnice. Blok upisan bez izabrane otpremnice
' je normalno stanje -- ceka na radnom stolu i vidi se u F1, lista "Bez
' otpremnice". Ovde je samo onaj koji je BIO u otpremnici pa ga je njen storno
' oslobodio: posao koji je neko zapoceo i ostavio. Zato se broji bas zapis
' istorije ("bila u"), a ne sama nevezanost.
'
' Skup racuna KANON (modDokumenta.NevezaniOtkupi) -- isto pravilo po kome radni
' sto zna sta je slobodno, pa lista ne moze da ponudi blok koji je u stvari
' zauzet, ni da sakrije onaj koji nije.
'
' Dedup `seen` se NAMERNO ne deli sa osirocenim dokumentima: broj bloka i broj
' prijemnice su dva razlicita niza istog oblika ("1/ddmmgg"), pa bi zajednicki
' spisak sakrio red zbog tudjeg broja.
'
' Strog citac clanstva PADA na korupciji, i to se ovde NE guta: red sa statusom
' GRESKA je vidljiv ishod, a tiho kraca lista nije -- ovo je ekran koji postoji
' da nabroji ono sto nije u redu.
Private Sub AddIzgubljeniBlokovi(ByRef result As Collection)
    Dim nevezani As Object, brojevi As Object
    Dim k As Variant, bilaU As String, ref As String
    Dim errDesc As String

    On Error GoTo EH

    Set nevezani = modDokumenta.NevezaniOtkupi()
    If nevezani Is Nothing Then Exit Sub
    If nevezani.count = 0 Then Exit Sub

    Set brojevi = BrojeviOtkupa()

    For Each k In nevezani.keys
        bilaU = Trim$(CStr(nevezani(k)))
        If Len(bilaU) > 0 Then
            ref = ""
            If brojevi.Exists(CStr(k)) Then ref = Trim$(CStr(brojevi(CStr(k))))
            If Len(ref) = 0 Then ref = CStr(k)
            AddNedRowFull result, "IZGUBLJEN_BLOK", ref, "IZGUBLJEN", _
                Poruka("OTKUI_OPO_BLOK_OPIS") & " " & bilaU, _
                Poruka("OTKUI_OPO_BLOK_AKCIJA"), _
                "", "Otkup", "", "", "BLOK"
        End If
    Next k
    Exit Sub
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr MOD_NAME & ".AddIzgubljeniBlokovi"
    AddNedRowFull result, "IZGUBLJEN_BLOK", "", "GRESKA", _
        Poruka("OTKUI_OPO_BLOK_GRESKA") & " " & errDesc, _
        Poruka("OTKUI_OPO_BLOK_AKCIJA_GRESKA"), "", "Otkup", "", "", "BLOK"
End Sub

' Nedovrseni hladnjacki lanci (S6 korak 7, odluka 8, HLD-LANAC-01).
'
' Blok JE u otpremnici, a lancu fali zbirna ili prijemnica. Nijedna druga lista
' ga ne vidi: "Bez otpremnice" cita nevezane blokove, a ovaj je vezan. Skup
' racuna vlasnik lanca (modAutoHladnjaca.NedovrseniLanci), po kanonskom
' roditelju -- ista pitanja koja lanac postavlja kad nastavlja.
'
' Red nosi OtkupID: radnja "Ponovi auto-lanac" bira blok po njemu, ne po broju
' (broj bloka je labela stanice, pa ga nosi i blok druge stanice). Pad citanja
' nije kraca lista nego red sa statusom GRESKA -- isti obrazac kao
' AddIzgubljeniBlokovi.
Private Sub AddNedovrseniLanci(ByRef result As Collection)
    Dim lanci As Collection, i As Long, r As Variant
    Dim ref As String, opis As String, errDesc As String

    On Error GoTo EH

    Set lanci = modAutoHladnjaca.NedovrseniLanci()
    For i = 1 To lanci.count
        r = lanci(i)
        ref = Trim$(CStr(r(1)))
        If Len(ref) = 0 Then ref = CStr(r(0))
        If Len(CStr(r(3))) = 0 Then
            opis = Poruka("OTKUI_OPO_LANAC_OPIS_OTP") & " " & _
                   modAutoHladnjaca.LanacBrojDokumenta("OTP", CStr(r(2)))
        Else
            opis = Poruka("OTKUI_OPO_LANAC_OPIS_ZBR") & " " & _
                   modAutoHladnjaca.LanacBrojDokumenta("ZBR", CStr(r(3)))
        End If
        AddNedRowFull result, "LANAC", ref, "NEDOVRSEN", opis, _
            Poruka("OTKUI_OPO_LANAC_AKCIJA"), "", "Otkup", "", "", "LANAC", CStr(r(0))
    Next i
    Exit Sub
EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr MOD_NAME & ".AddNedovrseniLanci"
    AddNedRowFull result, "LANAC", "", "GRESKA", _
        Poruka("OTKUI_OPO_LANAC_GRESKA") & " " & errDesc, _
        Poruka("OTKUI_OPO_LANAC_AKCIJA_GRESKA"), "", "Otkup", "", "", "LANAC"
End Sub

' UCase(OtkupID) -> BrojDokumenta, jednim prolazom kroz tabelu. NevezaniOtkupi
' vraca kljuceve u UCase, pa se i ovde tako kljuca; LookupValue po bloku bi bio
' jos jedan prolaz po svakom redu liste.
Private Function BrojeviOtkupa() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set BrojeviOtkupa = d

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If Not IsArray(data) Then Exit Function

    Dim cId As Long, cBr As Long, i As Long, oid As String
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    If cId = 0 Or cBr = 0 Then Exit Function

    For i = 1 To UBound(data, 1)
        oid = UCase$(Trim$(CStr(data(i, cId))))
        If Len(oid) > 0 Then d(oid) = Trim$(CStr(data(i, cBr)))
    Next i
End Function

Private Sub AddNedRowFull(ByRef col As Collection, ByVal kind As String, ByVal ref As String, _
                          ByVal status As String, ByVal opis As String, ByVal akcija As String, _
                          ByVal correctionID As String, ByVal docType As String, _
                          ByVal newBroj As String, ByVal mode As String, ByVal actionCode As String, _
                          Optional ByVal otkupID As String = "")
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("kind") = kind: d("ref") = ref: d("status") = status
    d("opis") = opis: d("akcija") = akcija
    d("correctionID") = correctionID: d("docType") = docType
    d("newBroj") = newBroj: d("mode") = mode: d("actionCode") = actionCode
    d("otkupID") = otkupID
    col.Add d
End Sub

' Prosiri osirocene (2D niz iz Get*-funkcija) u pojedinacne redove. brojCol /
' idCol = 1-based kolone sa poslovnim brojem (labela) i PrijemnicaID-om
' (identitet). Preskace DOKUMENT koji je vec u `seen` (context dedup + interni),
' po ID-u; red bez ID-a se ne deduplikuje.
'
' Greska citanja se NE guta (S8R1B-REV-01): necitljiva kolona identiteta daje red
' GRESKA i redove bez deduplikacije, a pad usred liste daje red GRESKA umesto
' neprimetno krace liste. Do 1b je "On Error GoTo done" tiho prekidao dodavanje.
Private Sub AddOsiroceneRows(ByRef result As Collection, ByVal v As Variant, ByVal kind As String, _
                             ByVal brojCol As Long, ByVal idCol As Long, ByVal docLabel As String, _
                             ByVal actionCode As String, ByVal akcija As String, ByRef seen As Object)
    On Error GoTo EH
    If Not Is2DArray(v) Then Exit Sub
    Dim idOk As Boolean
    idOk = (idCol >= LBound(v, 2) And idCol <= UBound(v, 2))
    If Not idOk Then
        AddNedRowFull result, KIND_GRESKA_CITANJA, docLabel, "GRESKA", _
            "Lista " & docLabel & ": kolona identiteta (" & idCol & ") nije citljiva -- redovi su " & _
            "prikazani bez deduplikacije.", "", "", docLabel, "", "", "GRESKA"
    End If
    Dim r As Long
    For r = 1 To UBound(v, 1)
        Dim b As String: b = Trim$(CStr(v(r, brojCol)))
        Dim id As String: id = ""
        If idOk Then id = Trim$(CStr(v(r, idCol)))
        If Len(b) > 0 Then
            Dim kljuc As String: kljuc = ""
            If Len(id) > 0 Then kljuc = "ID|" & id
            If Len(kljuc) = 0 Or Not seen.Exists(kljuc) Then
                If Len(kljuc) > 0 Then seen(kljuc) = True
                AddNedRowFull result, kind, b, "OSIROCENO", docLabel & " " & b & " -- ceka doradu", _
                    akcija, "", docLabel, "", "", actionCode
            End If
        End If
    Next r
    Exit Sub
EH:
    AddNedRowFull result, KIND_GRESKA_CITANJA, docLabel, "GRESKA", _
        "Lista " & docLabel & " nije procitana do kraja: " & Err.description, "", "", docLabel, "", "", "GRESKA"
End Sub

' Pali/gasi seam kvara liste Nedovrseno (v. mNedovrsenoKvarTest): "ID_KOLONA" ili ""
' (gasi). Paljenje van test rezima ne radi nista; gasenje radi uvek.
Public Sub NedovrsenoKvarTestSet(ByVal rezim As String)
    rezim = UCase$(Trim$(rezim))
    If Len(rezim) > 0 And Not IsTestMode() Then Exit Sub
    mNedovrsenoKvarTest = rezim
End Sub

Private Function NedovrsenoKvarAkoTest(ByVal rezim As String) As Boolean
    If mNedovrsenoKvarTest <> rezim Then Exit Function
    NedovrsenoKvarAkoTest = IsTestMode()
End Function

' ID storniranog dokumenta iz contexta; prazno kad ga context ne nosi ili citanje
' padne. Prazno znaci "context ne pokriva nijedan osiroceni red" -- red se tada
' prikazuje (fail-visible), umesto da greska obori celu listu Nedovrseno.
Private Function StorniraniDocIDTiho(ByVal cid As String) As String
    On Error Resume Next
    StorniraniDocIDTiho = Trim$(modStornoDok.StorniraniDocID(cid))
End Function

Private Function Is2DArray(ByVal v As Variant) As Boolean
    On Error GoTo no
    If IsEmpty(v) Then Exit Function
    If Not IsArray(v) Then Exit Function
    Dim probe As Long: probe = UBound(v, 2)   ' baci gresku ako nije 2D
    Is2DArray = (UBound(v, 1) >= 1)
    Exit Function
no:
End Function

' ============================================================
' VRATI STORNO PO BROJU -- samo zurnal-put. MUTIRA.
' ============================================================
' Storno bez zurnal operacije se NE vraca. Legacy put (do 10e-2) je za Otkup
' skidao Stornirano sa bloka i zastavicu sa reda knjige, a kontra-stav je ostajao
' -- aktivan otkup bez ambalaze, protiv AMB-10-ODL-25. Dosezao ga je samo storno
' bez zurnala, a StornoOtkup ga uvek otvara (BeginStornoOp). Revers vise nije u
' zurnalu: storno mu je kontra-stav nad AmbDokID-em.
Public Function UndoStorno_TX(ByVal docType As String, ByVal broj As String) As Boolean
    Const SRC As String = MOD_NAME & ".UndoStorno_TX"
    Dim opID As String
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then Err.Raise ERR_REC_BASE + 1, SRC, "Broj dokumenta je obavezan."

    opID = LatestOpFor(docType, broj)
    If Len(opID) = 0 Then Err.Raise ERR_REC_BASE + 3, SRC, _
        "Nema zurnal operacije storna za " & docType & " " & broj & " -> nema sta da se vrati."
    UndoStorno_TX = UndoOperation_TX(opID)
    Exit Function
EH:
    LogErr SRC
    UndoStorno_TX = False
End Function

' Mrtav roditelj PO REDU: za KONKRETAN OtkupID, ne za sve redove istog broja.
' Koristi je zurnal-put (UndoOperation_TX) da mrtav roditelj DRUGE, nepovezane
' istorijske generacije istog broja ne preblokira undo bas ove operacije. Verzija
' po broju je otisla u 10e-2, sa legacy granom UndoStorno_TX koja ju je jedina zvala.
' FAIL-CLOSED: greska/nedostajuca sema -> BLOKIRAJUCI marker. "" = bezbedno.
Public Function OtkupBlockDeadParentByID(ByVal otkupID As String) As String
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cId As Long, cOtp As Long, cZbr As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cZbr = GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE)
    If cId = 0 Then OtkupBlockDeadParentByID = "(provera roditelja nije moguca - sema)": Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cId))) = Trim$(otkupID) Then
            If cOtp > 0 Then
                Dim otpID As String: otpID = Trim$(CStr(data(i, cOtp)))
                If Len(otpID) > 0 Then
                    If Len(LookupActiveID(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_ID)) = 0 Then
                        OtkupBlockDeadParentByID = "otpremnica (ID " & otpID & ")": Exit Function
                    End If
                End If
            End If
            If cZbr > 0 Then
                Dim brZbr As String: brZbr = Trim$(CStr(data(i, cZbr)))
                If Len(brZbr) > 0 Then
                    If Len(LookupActiveID(TBL_ZBIRNA, COL_ZBR_BROJ, brZbr, COL_ZBR_BROJ)) = 0 Then
                        OtkupBlockDeadParentByID = "zbirna " & brZbr: Exit Function
                    End If
                End If
            End If
            Exit For
        End If
    Next i
    Exit Function
EH:
    LogErr MOD_NAME & ".OtkupBlockDeadParentByID"
    OtkupBlockDeadParentByID = "(provera roditelja nije uspela)"
End Function

' ============================================================
' TEST HOOK-ovi (Alt+F8) - verifikacija PRE vezivanja UI-a.
' ============================================================
Public Sub Test_GetNedovrseno()
    Dim c As Collection: Set c = GetNedovrseno()
    Debug.Print "=== NEDOVRSENO (" & IIf(c Is Nothing, 0, c.count) & ") ==="
    Dim i As Long
    If Not c Is Nothing Then
        For i = 1 To c.count
            Dim d As Object: Set d = c(i)
            Debug.Print d("kind") & " | ref=" & d("ref") & " | " & d("status") & " | " & d("opis")
        Next i
    End If
    Debug.Print "=== kraj ==="
End Sub

Public Sub Test_UndoStorno()
    Dim tip As String
    tip = Trim$(InputBox("Tip za reverziju (zurnal operacija, npr. 'Otkup'):", _
        "Vrati storno - TEST"))
    If Len(tip) = 0 Then Exit Sub
    Dim broj As String
    broj = Trim$(InputBox("Broj STORNIRANOG dokumenta:", "Vrati storno - TEST"))
    If Len(broj) = 0 Then Exit Sub

    Dim ok As Boolean: ok = UndoStorno_TX(tip, broj)
    If ok Then
        MsgBox "Vrati storno USPEO: " & tip & " " & broj, vbInformation, "Vrati storno"
    Else
        MsgBox "Vrati storno NIJE uspeo za " & tip & " " & broj & "." & vbCrLf & _
               "(Guard/greska -> vidi Immediate / Monitor. Storno bez zurnala se ne vraca.)", _
               vbExclamation, "Vrati storno"
    End If
End Sub
