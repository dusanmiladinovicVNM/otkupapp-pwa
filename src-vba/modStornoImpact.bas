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

Public Function BuildStornoImpact(ByVal docType As String, ByVal broj As String, _
                                  Optional ByVal dokumentTip As String = "") As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildStornoImpact = d
    On Error GoTo EH
    broj = Trim$(broj)

    Set d("header") = ImpactHeader(docType, broj)
    Set d("chain") = GetStornoChainRows(docType, broj, dokumentTip)
    Set d("blocks") = GetStornoBlockRows(docType, broj, dokumentTip)
    Set d("flags") = GetChainFlags(docType, broj, dokumentTip)
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
Private Function ImpactHeader(ByVal docType As String, ByVal broj As String) As Object
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    Set ImpactHeader = h
    On Error GoTo EH
    h("tip") = docType: h("broj") = broj
    h("partnerID") = "": h("partner") = "": h("datum") = "": h("kolicina") = ""
    Select Case docType
        Case FLOW_DOC_OTPREMNICA
            h("partnerID") = HL(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_STANICA)
            h("datum") = HL(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_DATUM)
            h("kolicina") = HL(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_KOLICINA)
        Case FLOW_DOC_ZBIRNA
            h("partnerID") = HL(TBL_ZBIRNA, COL_ZBR_BROJ, broj, COL_ZBR_KUPAC)
            h("datum") = HL(TBL_ZBIRNA, COL_ZBR_BROJ, broj, COL_ZBR_DATUM)
            h("kolicina") = HL(TBL_ZBIRNA, COL_ZBR_BROJ, broj, COL_ZBR_KOLICINA)
        Case FLOW_DOC_PRIJEMNICA
            h("partnerID") = HL(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_KUPAC)
            h("datum") = HL(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_DATUM)
            h("kolicina") = HL(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_KOLICINA)
    End Select
    ' Razresi ID -> naziv (otpremnica = stanica; zbirna/prijemnica = kupac). Fallback ID.
    h("partner") = ResolvePartnerName(docType, CStr(h("partnerID")))
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
