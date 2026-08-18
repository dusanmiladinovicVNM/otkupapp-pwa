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
' dokumenta). Do v6-ui-143 ga ovaj sloj nije imao, pa je ceo uvid isao po
' BROJU -- a broj nije jedinstven (GenerateBrojPrijemnice nema proveru
' jedinstvenosti). Posledica nije bila teorijska: pod kolizijom broja je
' pregled pokazivao lanac i blokove TUDJEG dokumenta, a operater bi na osnovu
' toga doneo odluku o stornu. Tri citaca ispod su vec primala docID od
' v6-ui-136/140 -- nedostajao je samo ovaj sloj koji ih spaja.
'
' Prazan docID i dalje prolazi: zatecen zapis bez generacije nema identitet, i
' tada nizvodno vazi fail-closed kapija nad jednoznacnoscu broja.
' "valid" JE DEO UGOVORA. Model se do v6-ui-143 vracao pozivaocu PRE nego sto je
' izgradjen (Set BuildStornoImpact = d na pocetku), pa je pad na pola davao
' PARCIJALAN recnik koji spolja izgleda kao ispravan uvid. Ekran na osnovu njega
' crta posledice i nudi dugmad za mutaciju -- dakle tacno suprotno od svrhe ovog
' ekrana ("prvo vidi posledice, pa odluci"). Sada svaka sekcija mora da se
' izgradi, i tek onda se "valid" postavlja na True; pozivalac koji ga ne proveri
' dobija recnik ciji je valid False.
Public Function BuildStornoImpact(ByVal docType As String, ByVal broj As String, _
                                  Optional ByVal dokumentTip As String = "", _
                                  Optional ByVal docID As String = "") As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("valid") = False
    d("greska") = ""
    Set BuildStornoImpact = d
    On Error GoTo EH
    broj = Trim$(broj)

    Set d("header") = ImpactHeader(docType, broj, docID)
    Set d("chain") = GetStornoChainRows(docType, broj, dokumentTip, docID)
    Set d("blocks") = GetStornoBlockRows(docType, broj, dokumentTip, docID)
    Set d("flags") = GetChainFlags(docType, broj, dokumentTip, docID)
    Set d("palete") = ImpactPalete(docType, broj, docID)
    Set d("faktura") = ImpactFaktura(docType, broj, docID)
    Set d("summary") = ImpactSummary(d)
    d("valid") = True
    Exit Function
EH:
    d("greska") = Err.description
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
    ' Greska se PROPUSTA dalje, ne guta. Ovaj sloj je deo modela koji se posle
    ' oznacava kao valid -- progutana greska bi dala prazno polje koje spolja
    ' izgleda kao "podatka nema", a znaci "ne znam".
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ImpactHeader"
    Err.Raise errNum, MOD_NAME & ".ImpactHeader", errDesc
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
' Palete po tipu, SUZENO NA IZABRAN DOKUMENT gde god sema to dozvoljava.
'
' Do v6-ui-143 je ovde stajao goli broj, i to je bila prava rupa: zaglavlje,
' lanac i blokovi su isli po identitetu, a palete i faktura po BROJU -- pa je
' pod kolizijom broja uvid pokazivao palete OBA dokumenta, dok writer nizvodno
' mutira samo jedan. Ekran je time tvrdio posledice koje ne odgovaraju radnji.
'
' Dokle sema dozvoljava suzavanje:
'   PRIJEMNICA  tblPaletaStavka nosi PrijemnicaID -> puno suzavanje. Jedan
'               logicki dokument ima VISE redova (Klasa I i II dele generaciju,
'               a imaju razlicit PrijemnicaID), pa se salje SKUP ID-jeva.
'   OTPREMNICA  stavke ne nose otpremnicu nego BrojZbirne; identitet se ipak
'               koristi da se nadje zbirna BAS te otpremnice (HLI, ne HL).
'   ZBIRNA      stavke nose BrojZbirne, ne ZbirnaID -- suzavanje po identitetu
'               nije moguce. Ista granica sheme kao kod FLOW_DOC_ZBIRNA u
'               ActiveBlocksForFlow (tblOtkup nosi denormalizovan BrojZbirne).
'               Prijavljuje se kao granica, ne krpi se pogadjanjem.
' ------------------------------------------------------------
Private Function ImpactPalete(ByVal docType As String, ByVal broj As String, _
                              Optional ByVal docID As String = "") As Collection
    Dim ids As Object, bz As String
    On Error GoTo EH
    Select Case docType
        Case FLOW_DOC_PRIJEMNICA
            Set ids = PrijemniceIDPoIdentitetu(broj, docID)
            If ids Is Nothing Then
                ' IDENTITET SE NE ODUSTAJE. Kad je docID zadat a ne moze da se
                ' razresi (schema drift, nema reda pod tom generacijom), povratak
                ' na broj bi vratio tacno ono sto je #198 vadio -- i to unutar
                ' modela koji se posle oznacava kao valid. Prazan docID je druga
                ' prica: zatecen zapis nema identitet, pa je broj sve sto postoji.
                If Len(Trim$(docID)) > 0 Then
                    Err.Raise ERR_UI_BASE + 26, MOD_NAME & ".ImpactPalete", _
                              "Identitet prijemnice se ne moze razresiti -- palete se ne mogu suziti."
                End If
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_PRIJ, broj, Nothing, True)
            Else
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_PRIJEMNICA_ID, "", ids, True)
            End If
        Case FLOW_DOC_ZBIRNA
            Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_ZBIRNE, broj, Nothing, True)
        Case FLOW_DOC_OTPREMNICA
            bz = HLI(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_BROJ_ZBIRNE, docID)
            If Len(bz) > 0 Then
                Set ImpactPalete = GetPaleteImpactByField(COL_PALS_BROJ_ZBIRNE, bz, Nothing, True)
            Else
                Set ImpactPalete = New Collection
            End If
        Case Else
            Set ImpactPalete = New Collection
    End Select
    Exit Function
EH:
    ' Greska se PODIZE, ne guta. Prazna kolekcija je legitiman odgovor (dokument
    ' nema palete) i ne sme da znaci isto sto i neuspelo citanje -- inace bi uvid
    ' tvrdio da posledica nema, a operater bi na osnovu toga stornirao.
    LogErr MOD_NAME & ".ImpactPalete"
    Err.Raise ERR_UI_BASE + 23, MOD_NAME & ".ImpactPalete", _
              "Palete izabranog dokumenta se ne mogu procitati."
End Function

' PrijemnicaID-jevi koji pripadaju IZABRANOJ generaciji. Nothing = nema identiteta
' ili sema nema kolonu generacije -> pozivalac se vraca na broj. Fail-open je ovde
' ispravan: zatecen zapis bez generacije nema identitet, a uvid mora nesto da
' pokaze; nizvodne KAPIJE su te koje su fail-closed.
Private Function PrijemniceIDPoIdentitetu(ByVal broj As String, _
                                          ByVal docID As String) As Object
    If Len(Trim$(docID)) = 0 Then Exit Function
    On Error GoTo done
    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function
    Dim cBr As Long, cId As Long, cGen As Long
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    cGen = GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID)
    If cBr = 0 Or cId = 0 Or cGen = 0 Then Exit Function
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, pid As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBr))) = Trim$(broj) Then
            If Trim$(CStr(data(i, cGen))) = Trim$(docID) Then
                pid = Trim$(CStr(data(i, cId)))
                If Len(pid) > 0 Then
                    If Not d.Exists(pid) Then d.Add pid, 1
                End If
            End If
        End If
    Next i
    ' Prazan skup NIJE "nema paleta" nego "identitet nije nadjen" -- vraca se
    ' Nothing, da pozivalac padne na broj umesto da tvrdi da paleta nema.
    If d.count > 0 Then Set PrijemniceIDPoIdentitetu = d
done:
End Function

' Faktura (samo prijemnica ima direktan link u ovom sloju), po IDENTITETU.
' HL vraca PRVI red po broju -- pod kolizijom je uvid umeo da prijavi fakturu
' tudje prijemnice, pa i da tvrdi "fakturisano" za dokument koji to nije.
Private Function ImpactFaktura(ByVal docType As String, ByVal broj As String, _
                               Optional ByVal docID As String = "") As Object
    Dim f As Object: Set f = CreateObject("Scripting.Dictionary")
    Set ImpactFaktura = f
    f("hasFaktura") = False: f("fakturaID") = ""
    On Error GoTo EH
    If docType = FLOW_DOC_PRIJEMNICA Then
        f("hasFaktura") = (UCase$(HLI(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_FAKTURISANO, docID)) = "DA")
        f("fakturaID") = HLI(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_FAKTURA_ID, docID)
    End If
    Exit Function
EH:
    ' Greska se PROPUSTA dalje, ne guta. Ovaj sloj je deo modela koji se posle
    ' oznacava kao valid -- progutana greska bi dala prazno polje koje spolja
    ' izgleda kao "podatka nema", a znaci "ne znam".
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    LogErr MOD_NAME & ".ImpactFaktura"
    Err.Raise errNum, MOD_NAME & ".ImpactFaktura", errDesc
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
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then
        Err.Raise ERR_UI_BASE + 29, MOD_NAME & ".HLI", _
                  "Tabela " & tbl & " nije citljiva."
    End If
    Dim cKey As Long, cVal As Long, cGen As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cVal = GetColumnIndex(tbl, valCol)
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cKey = 0 Or cVal = 0 Then
        Err.Raise ERR_UI_BASE + 28, MOD_NAME & ".HLI", _
                  "Kolona " & keyCol & " ili " & valCol & " ne postoji u " & tbl & "."
    End If
    ' Tabela bez kolone generacije, a identitet JE zadat -> ne moze da se sazna
    ' o kom je dokumentu rec. Povratak na broj bi ovde bio tiha degradacija
    ' unutar modela koji se posle oznacava kao valid, pa se umesto toga dize
    ' greska i ceo uvid pada.
    If cGen = 0 Then
        Err.Raise ERR_UI_BASE + 27, MOD_NAME & ".HLI", _
                  "Tabela " & tbl & " nema kolonu " & COL_GENERACIJA_ID & "."
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
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then
        Err.Raise ERR_UI_BASE + 30, MOD_NAME & ".SumActiveNum", _
                  "Tabela " & tbl & " nije citljiva."
    End If
    Dim cKey As Long, cSum As Long, cSt As Long, cGen As Long
    cKey = GetColumnIndex(tbl, keyCol)
    cSum = GetColumnIndex(tbl, sumCol)
    cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cKey = 0 Or cSum = 0 Then
        Err.Raise ERR_UI_BASE + 31, MOD_NAME & ".SumActiveNum", _
                  "Kolona " & keyCol & " ili " & sumCol & " ne postoji u " & tbl & "."
    End If
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
