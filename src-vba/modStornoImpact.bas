Attribute VB_Name = "modStornoImpact"
Option Explicit

' ============================================================
' modStornoImpact - agregator UVIDA za Storno centar (Faza 1).
'
' READ-ONLY: sklapa pun model uticaja storna iz postojecih citaca; NE mutira
' nista. Jedan poziv (BuildStornoImpact) daje UI-u sve za "Uvid" ekran:
'   header  - {tip, broj, partnerID, datum, kolicina}
'   chain   - Collection redova [dok, info, napomena]  (modStornoFlow.GetStornoChainRows)
'   blocks  - Collection redova [otkupID, brDok, kg, klasa, koop]  (GetStornoBlockRows)
'   flags   - {hasDependents, canPonistenjeClean, dependentsText}  (GetChainFlags)
'   palete  - Collection dicta po paleti (agregat + detach delta)  (GetPaleteImpactByField)
'   faktura - {hasFaktura, fakturaID}
'   summary - {blockCount, paleteCount, detachGajb, detachNeto, detachAmb}  (traka uticaja)
'
' Reuse: modStornoFlow (Public citaci), modPaletniList.GetPaleteImpactByField,
' modDataAccess.LookupValue. Bez MsgBox, bez TX.
' ============================================================

Private Const MOD_NAME As String = "modStornoImpact"

' docID = KANONSKI IDENTITET izabranog dokumenta (GeneracijaID za robna
' dokumenta). Do v6-ui-142 ga ovaj sloj nije imao, pa je ceo uvid isao po
' BROJU -- a broj nije jedinstven (GenerateBrojPrijemnice nema proveru
' jedinstvenosti). Posledica nije bila teorijska: pod kolizijom broja je
' pregled pokazivao lanac i blokove TUDJEG dokumenta, a operater bi na osnovu
' toga doneo odluku o stornu. Tri citaca ispod su vec primala docID od
' v6-ui-136/140 -- nedostajao je samo ovaj sloj koji ih spaja.
'
' Prazan docID i dalje prolazi: zatecen zapis bez generacije nema identitet, i
' tada nizvodno vazi fail-closed kapija nad jednoznacnoscu broja.
Public Function BuildStornoImpact(ByVal docType As String, ByVal broj As String, _
                                  Optional ByVal dokumentTip As String = "", _
                                  Optional ByVal docID As String = "") As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildStornoImpact = d
    On Error GoTo EH
    broj = Trim$(broj)

    Set d("header") = ImpactHeader(docType, broj, docID)
    Set d("chain") = GetStornoChainRows(docType, broj, dokumentTip, docID)
    Set d("blocks") = GetStornoBlockRows(docType, broj, dokumentTip, docID)
    Set d("flags") = GetChainFlags(docType, broj, dokumentTip, docID)
    Set d("palete") = ImpactPalete(docType, broj)
    Set d("faktura") = ImpactFaktura(docType, broj)
    Set d("summary") = ImpactSummary(d)
    Exit Function
EH:
    LogErr MOD_NAME & ".BuildStornoImpact"
End Function

' ------------------------------------------------------------
' Header po tipu (samo potvrdjene kolone; partner ostaje ID -> UI resolvira ime).
' ------------------------------------------------------------
Private Function ImpactHeader(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal docID As String = "") As Object
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    Set ImpactHeader = h
    On Error GoTo EH
    h("tip") = docType: h("broj") = broj
    h("partnerID") = "": h("partner") = "": h("datum") = "": h("kolicina") = ""
    h("ispravkaOd") = "": h("zamenjenSa") = ""
    Dim tTbl As String, tCol As String: tTbl = "": tCol = ""
    Select Case docType
        ' kolicina = SUMA aktivnih redova broja (Klasa I + II dele broj; HL bi vratio
        ' samo prvu klasu -> potceni dvoklasni dokument u uvidu). partner/datum su
        ' isti po klasama pa im je dovoljan jedan red -- ali BAS njegov, pa idu
        ' kroz HLI, koji uz broj postuje i identitet.
        Case FLOW_DOC_OTPREMNICA
            tTbl = TBL_OTPREMNICA: tCol = COL_OTP_BROJ
        Case FLOW_DOC_ZBIRNA
            tTbl = TBL_ZBIRNA: tCol = COL_ZBR_BROJ
        Case FLOW_DOC_PRIJEMNICA
            tTbl = TBL_PRIJEMNICA: tCol = COL_PRJ_BROJ
    End Select
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            h("partnerID") = HLI(tTbl, tCol, broj, COL_OTP_STANICA, docID)
            h("datum") = HLI(tTbl, tCol, broj, COL_OTP_DATUM, docID)
            h("kolicina") = SumActiveNum(tTbl, tCol, broj, COL_OTP_KOLICINA, docID)
        Case FLOW_DOC_ZBIRNA
            h("partnerID") = HLI(tTbl, tCol, broj, COL_ZBR_KUPAC, docID)
            h("datum") = HLI(tTbl, tCol, broj, COL_ZBR_DATUM, docID)
            h("kolicina") = SumActiveNum(tTbl, tCol, broj, COL_ZBR_KOLICINA, docID)
        Case FLOW_DOC_PRIJEMNICA
            h("partnerID") = HLI(tTbl, tCol, broj, COL_PRJ_KUPAC, docID)
            h("datum") = HLI(tTbl, tCol, broj, COL_PRJ_DATUM, docID)
            h("kolicina") = SumActiveNum(tTbl, tCol, broj, COL_PRJ_KOLICINA, docID)
    End Select
    ' Razresi ID -> naziv (otpremnica = stanica; zbirna/prijemnica = kupac). Fallback ID.
    h("partner") = ResolvePartnerName(docType, CStr(h("partnerID")))
    ' Sledljivost (Faza 7): da li je ovaj dokument ispravka drugog / zamenjen drugim.
    If Len(tTbl) > 0 Then
        h("ispravkaOd") = HLI(tTbl, tCol, broj, COL_TRACE_ISPRAVKA_OD, docID)
        h("zamenjenSa") = HLI(tTbl, tCol, broj, COL_TRACE_ZAMENJEN_SA, docID)
    End If
    Exit Function
EH:
    LogErr MOD_NAME & ".ImpactHeader"
End Function

' ID partnera -> citljiv naziv. Otpremnica: tblStanice (StanicaID -> Naziv);
' zbirna/prijemnica: tblKupci (KupacID -> Naziv). Ako naziv nema -> vrati ID.
Private Function ResolvePartnerName(ByVal docType As String, ByVal partnerID As String) As String
    On Error Resume Next
    ResolvePartnerName = partnerID
    If Len(Trim$(partnerID)) = 0 Then Exit Function
    Dim nm As String
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            nm = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", partnerID, "Naziv")))
        Case FLOW_DOC_ZBIRNA, FLOW_DOC_PRIJEMNICA
            nm = Trim$(CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, partnerID, COL_KUP_NAZIV)))
    End Select
    If Len(nm) > 0 Then ResolvePartnerName = nm
End Function

' ------------------------------------------------------------
' Palete po tipu: prijemnica preko BrojPrij; zbirna preko BrojZbirne;
' otpremnica preko svoje zbirne. Reuse modPaletniList.GetPaleteImpactByField.
' ------------------------------------------------------------
Private Function ImpactPalete(ByVal docType As String, ByVal broj As String) As Collection
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_PRIJEMNICA
            Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_PRIJ, broj)
        Case FLOW_DOC_ZBIRNA
            Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_ZBIRNE, broj)
        Case FLOW_DOC_OTPREMNICA
            Dim bz As String: bz = HL(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_BROJ_ZBIRNE)
            If Len(bz) > 0 Then
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_ZBIRNE, bz)
            Else
                Set ImpactPalete = New Collection
            End If
        Case Else
            Set ImpactPalete = New Collection
    End Select
    Exit Function
EH:
    LogErr MOD_NAME & ".ImpactPalete"
    Set ImpactPalete = New Collection
End Function

' ------------------------------------------------------------
' Faktura (samo prijemnica ima direktan link u ovom sloju).
' ------------------------------------------------------------
Private Function ImpactFaktura(ByVal docType As String, ByVal broj As String) As Object
    Dim f As Object: Set f = CreateObject("Scripting.Dictionary")
    Set ImpactFaktura = f
    f("hasFaktura") = False: f("fakturaID") = ""
    On Error GoTo EH
    If docType = FLOW_DOC_PRIJEMNICA Then
        f("hasFaktura") = (UCase$(HL(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_FAKTURISANO)) = "DA")
        f("fakturaID") = HL(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_FAKTURA_ID)
    End If
    Exit Function
EH:
    LogErr MOD_NAME & ".ImpactFaktura"
End Function

' ------------------------------------------------------------
' Summary (traka uticaja): brojevi + detach delta zbir preko svih paleta.
' detach* = koliko bi DUPLI/PONISTENJE skinulo (thisGajb/thisNeto/thisAmb).
' ------------------------------------------------------------
Private Function ImpactSummary(ByVal d As Object) As Object
    Dim sm As Object: Set sm = CreateObject("Scripting.Dictionary")
    Set ImpactSummary = sm
    sm("blockCount") = 0: sm("paleteCount") = 0
    sm("detachGajb") = 0&: sm("detachNeto") = 0#: sm("detachAmb") = 0#
    On Error Resume Next

    Dim blocks As Collection: Set blocks = d("blocks")
    If Not blocks Is Nothing Then sm("blockCount") = blocks.count

    Dim palete As Collection: Set palete = d("palete")
    If Not palete Is Nothing Then
        sm("paleteCount") = palete.count
        Dim pg As Long, pk As Double, pa As Double, i As Long
        For i = 1 To palete.count
            pg = pg + CLng(palete(i)("thisGajb"))
            pk = pk + CDbl(palete(i)("thisNeto"))
            pa = pa + CDbl(palete(i)("thisAmb"))
        Next i
        sm("detachGajb") = pg: sm("detachNeto") = pk: sm("detachAmb") = pa
    End If
End Function

' Safe lookup -> Trim$ string ("" na gresku/prazno).
Private Function HL(ByVal tbl As String, ByVal keyCol As String, _
                    ByVal keyVal As String, ByVal valCol As String) As String
    On Error Resume Next
    HL = Trim$(CStr(LookupValue(tbl, keyCol, keyVal, valCol)))
End Function

' Isto, ali uz IDENTITET. Prazan docID -> obican lookup po broju (zatecen zapis
' bez generacije). Neprazan docID -> red koji ima i taj broj i tu generaciju;
' ako ga nema, vraca prazno umesto tudje vrednosti.
'
' LookupValue vraca PRVI red po broju, pa je pod kolizijom umeo da prikaze
' partnera i datum drugog dokumenta -- zaglavlje uvida bi opisivalo dokument
' koji se ne stornira.
Private Function HLI(ByVal tbl As String, ByVal keyCol As String, _
                     ByVal keyVal As String, ByVal valCol As String, _
                     ByVal docID As String) As String
    If Len(Trim$(docID)) = 0 Then
        HLI = HL(tbl, keyCol, keyVal, valCol)
        Exit Function
    End If
    On Error GoTo done
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cKey As Long, cVal As Long, cGen As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cVal = GetColumnIndex(tbl, valCol)
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cKey = 0 Or cVal = 0 Then Exit Function
    ' Tabela bez kolone generacije (zatecena instalacija) -> vrati se na broj.
    If cGen = 0 Then
        HLI = HL(tbl, keyCol, keyVal, valCol)
        Exit Function
    End If
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cKey))) = Trim$(keyVal) Then
            If Trim$(CStr(data(i, cGen))) = Trim$(docID) Then
                HLI = Trim$(CStr(data(i, cVal)))
                Exit Function
            End If
        End If
    Next i
done:
End Function

' Suma numericke kolone preko AKTIVNIH (ne-storniranih) redova istog kljuca -> ukupno
' po dokumentu (Klasa I + II). Prazan string ako nema aktivnog reda.
'
' docID suzava na redove IZABRANOG dokumenta. Klasa I i II iz istog upisa dele
' generaciju, pa suzavanje ne gubi drugu klasu -- odbacuje samo tudji dokument
' istog broja.
Private Function SumActiveNum(ByVal tbl As String, ByVal keyCol As String, _
                             ByVal keyVal As String, ByVal sumCol As String, _
                             Optional ByVal docID As String = "") As String
    On Error GoTo done
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cKey As Long, cSum As Long, cSt As Long, cGen As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cSum = GetColumnIndex(tbl, sumCol)
    cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cKey = 0 Or cSum = 0 Then Exit Function
    Dim uzmiID As Boolean
    uzmiID = (Len(Trim$(docID)) > 0 And cGen > 0)
    Dim i As Long, total As Double, found As Boolean
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cKey))) = Trim$(keyVal) Then
            Dim uzmi As Boolean: uzmi = True
            If uzmiID Then uzmi = (Trim$(CStr(data(i, cGen))) = Trim$(docID))
            If uzmi Then
                Dim isStor As Boolean: isStor = False
                If cSt > 0 Then isStor = (UCase$(Trim$(CStr(data(i, cSt)))) = "DA")
                If Not isStor Then
                    total = total + SafeDblZ(data(i, cSum))
                    found = True
                End If
            End If
        End If
    Next i
    If found Then SumActiveNum = Format$(total, "#,##0.##")
done:
End Function

Private Function SafeDblZ(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then SafeDblZ = CDbl(v) Else SafeDblZ = Val(CStr(v))
End Function

' ============================================================
' TEST HOOK (Alt+F8): ispisi model uticaja u Immediate (Ctrl+G). Read-only.
' ============================================================
Public Sub Test_BuildStornoImpact()
    Dim broj As String
    broj = Trim$(InputBox("Broj dokumenta:", "Impact test"))
    If Len(broj) = 0 Then Exit Sub
    Dim tip As String
    tip = Trim$(InputBox("Tip (Prijemnica / Otpremnica / Zbirna):", "Impact test", "Prijemnica"))
    If Len(tip) = 0 Then Exit Sub

    Dim d As Object: Set d = BuildStornoImpact(tip, broj)
    Dim h As Object: Set h = d("header")
    Dim sm As Object: Set sm = d("summary")
    Debug.Print "=== IMPACT " & tip & " " & broj & " ==="
    Debug.Print "Header: partnerID=" & h("partnerID") & " datum=" & h("datum") & " kol=" & h("kolicina")
    Debug.Print "Summary: blocks=" & sm("blockCount") & " palete=" & sm("paleteCount") & _
                " | detach gajb=" & sm("detachGajb") & " neto=" & sm("detachNeto") & " amb=" & sm("detachAmb")

    Dim pal As Collection: Set pal = d("palete")
    Dim i As Long
    For i = 1 To pal.count
        Dim p As Object: Set p = pal(i)
        Debug.Print "  Paleta " & p("label") & ": " & p("used") & "/" & p("cap") & " gajb" & _
                    " | this=" & p("thisGajb") & " (neto " & p("thisNeto") & ", amb " & p("thisAmb") & ")" & _
                    IIf(CBool(p("preradjena")), " [PRERADJENA]", "")
    Next i

    Dim fk As Object: Set fk = d("faktura")
    Debug.Print "Faktura: has=" & fk("hasFaktura") & " id=" & fk("fakturaID")
    Debug.Print "=== kraj ==="
End Sub
