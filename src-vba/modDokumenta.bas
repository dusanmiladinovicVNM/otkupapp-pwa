Attribute VB_Name = "modDokumenta"
'Attribute VB_Name = "modDokumenta"

Option Explicit

' ============================================================
' modDokumenta - Otpremnica, Zbirna, Prijemnica
' Dokumentenfluss: Otkup zu Otpremnica zu Zbirna zu Prijemnica zu Faktura
' ============================================================

' ============================================================
' ZBR-IDENT-01 -- IDENTITET ZBIRNE
'
' Ugovor i acceptance testovi: docs/DOMEN/ZBR_IDENTITET.md
'
' Broj zbirne je LABELA, ne identitet. Identitet logickog dokumenta je
' GeneracijaID; broj + VozacID + KupacID je SCOPE u kome se generacija nalazi
' ili kuje. Dvoklasna zbirna je JEDAN dokument na DVA reda -- SaveZbirnaMulti_TX
' zove SaveZbirna dvaput sa ISTIM brojem, vozacem i kupcem, pa oba reda nose
' istu generaciju.
'
' Resolver odgovara na DVA NEZAVISNA pitanja, i to je cela poenta razdvajanja:
'   resolutionStatus     -- da li je broj SADA jednoznacno razresiv (bez storniranih)
'   historicalOwnerCount -- da li je broj IKAD pripadao vise od jednog vlasnika
' UNIQUE uz historicalOwnerCount = 2 je VALIDNO stanje, ne kontradikcija: broj je
' danas jednoznacan, a u proslosti ga je drzao i drugi vlasnik.
'
' activeLogicalCount i activeOwnerCount se NE izvode jedan iz drugog: dvoklasna
' zbirna daje 1/1 iz dva reda, dva zasebna unosa istog vozaca daju 2/1.
' ============================================================
Public Const ZBR_INT_OK As String = "OK"
Public Const ZBR_INT_ERROR As String = "INTEGRITY_ERROR"

Public Const ZBR_RES_NONE As String = "NONE"
Public Const ZBR_RES_UNIQUE As String = "UNIQUE"
Public Const ZBR_RES_OWNER_MISMATCH As String = "OWNER_MISMATCH"
Public Const ZBR_RES_AMBIGUOUS As String = "CURRENT_AMBIGUOUS"

' Razlog odbijanja novog unosa. Kod, ne poruka: modDokumenta je sloj podataka i
' ne nosi korisnicki tekst (prevod je u modDokUnos.ZbirnaGatePoruka).
Public Const ZBR_GATE_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_GATE_AKTIVNA As String = "AKTIVNA"
Public Const ZBR_GATE_TUDJ As String = "TUDJ"
Public Const ZBR_GATE_SIROCE As String = "SIROCE"

' Razlog odbijanja RODITELJA (F4). Zaseban skup od ZBR_GATE_*: tamo je pitanje
' "sme li NOV broj", ovde "sme li se prijemnica vezati na POSTOJECI dokument".
Public Const ZBR_PARENT_NEMA As String = "NEMA"
Public Const ZBR_PARENT_DVOSMISLEN As String = "DVOSMISLEN"
Public Const ZBR_PARENT_TUDJ As String = "TUDJ_VLASNIK"
Public Const ZBR_PARENT_ISTORIJA As String = "ISTORIJA"

' ZBR-MUT-01: razlozi zbog kojih se po BROJU ne sme mutirati.
'
' Deca zbirne (otpremnica, prijemnica, paletna stavka, denormalizovan otkup) od
' faze 1 nose i ZbirnaGeneracijaID -- ali on sme biti PRAZAN (roditelj jos nije
' razresen), pa se na njega ne moze racunati bez provere.
'
' Rutina koja decu bira po broju zato i dalje zahvata SVE dokumente tog broja --
' osim kad je izbor scoped po generaciji (faza 3), a to zna samo pozivalac.
' Otud `scopedPoGeneraciji` parametar nize: kapija popusta tek kad akter dokaze
' da bira po generaciji, i to za CELU svoju operaciju.
'
' Dva razloga su RAZLICITA i ne smeju se stopiti:
'   VLASNIK   -- broj je IKAD pripadao vise od jednog (VozacID, KupacID).
'                Storniran vlasnik se broji: deca mu ostaju aktivna.
'   DOKUMENTI -- broj SADA nosi vise od jednog aktivnog logickog dokumenta,
'                makar bili istog vlasnika (A17). Do v6-ui-224 to stanje uvoz
'                nije ni pravio, jer je stapao generacije; sada ga pravi
'                ispravno, pa mora i da se vidi.
'
' historicalLogicalCount se MERI i prijavljuje, ali NE blokira: ispravka pod
' istim brojem (ZBR-ACTIVE-NUMBER-01 ALLOW grana) je zakuje na 2 kod svakog
' redovnog re-entry-ja, a DetachOtpremniceInline pri tom prazni broj na deci
' stornirane generacije -- pa ona za mutaciju po broju vise i ne konkurisu.
Public Const ZBR_MUT_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_MUT_VISE_VLASNIKA As String = "VISE_VLASNIKA"
Public Const ZBR_MUT_VISE_DOKUMENATA As String = "VISE_DOKUMENATA"

' ZBR-CHILD-01 faza 2: ZASTO backfill preskace broj. Tri razloga, ne jedan --
' prvo merenje nad pravim podacima je dalo 85 popunjenih i 4572 preskocena, a
' poruka je sva tri prijavljivala kao "broj je IKAD nosio vise dokumenata".
' Migracija koja ne ume da kaze STA je zatekla ne moze da vodi sledeci korak.
Public Const ZBR_BF_INTEGRITET As String = "INTEGRITET"
Public Const ZBR_BF_NEMA_GENERACIJE As String = "NEMA_GENERACIJE"
Public Const ZBR_BF_VISE_GENERACIJA As String = "VISE_GENERACIJA"

Public Type ZbirnaIdent
    normalizedBroj As String
    integrityStatus As String
    activeLogicalCount As Long
    activeOwnerCount As Long
    historicalOwnerCount As Long
    historicalLogicalCount As Long
    historicalOnlyGeneracijaID As String   ' popunjeno samo kad je historicalLogicalCount = 1
    scopeProvided As Boolean
    historicalOwnerIsScope As Boolean
    matchingScopeActiveLogicalCount As Long
    resolutionStatus As String
    selectedGeneracijaID As String
    selectedVozacID As String
    selectedKupacID As String
    brojUAktivnojPrijemnici As Boolean
End Type

' ============================================================
' OTPREMNICA - Station gibt Ware an Fahrer
' ============================================================
Public Function SaveOtpremnicaMulti_TX(ByVal datum As Date, _
                                       ByVal stanicaID As String, _
                                       ByVal vozacID As String, _
                                       ByVal brojOtp As String, _
                                       ByVal brojZbirne As String, _
                                       ByVal vrsta As String, _
                                       ByVal sorta As String, _
                                       ByVal kolicinaI As Double, _
                                       ByVal cenaI As Double, _
                                       ByVal tipAmb As String, _
                                       ByVal kolAmb As Long, _
                                       Optional ByVal hasKlasaII As Boolean = False, _
                                       Optional ByVal kolicinaII As Double = 0, _
                                       Optional ByVal cenaII As Double = 0, _
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
    modSchema.SchemaReadyOrFail "SaveOtpremnicaMulti_TX", _
        TBL_OTPREMNICA & "|" & TBL_AMBALAZA

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    ' Klasa I je opciona (kolicinaI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1103, "SaveOtpremnicaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
        resultI = SaveOtpremnica( _
            datum, _
            stanicaID, _
            vozacID, _
            brojOtp, _
            brojZbirne, _
            vrsta, _
            sorta, _
            kolicinaI, _
            cenaI, _
            tipAmb, _
            kolAmb, _
            KLASA_I, _
            brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1101, "SaveOtpremnicaMulti_TX", _
                      "SaveOtpremnica Klasa I fehlgeschlagen"
        End If
    End If

    Dim resultII As String
    If hasKlasaII Then
        resultII = SaveOtpremnica( _
            datum, _
            stanicaID, _
            vozacID, _
            brojOtp, _
            brojZbirne, _
            vrsta, _
            sorta, _
            kolicinaII, _
            cenaII, _
            tipAmb, _
            kolAmbII, _
            KLASA_II, _
            brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1102, "SaveOtpremnicaMulti_TX", _
                      "SaveOtpremnica Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SaveOtpremnicaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SaveOtpremnicaMulti_TX = resultI
    Else
        SaveOtpremnicaMulti_TX = resultII
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

    ' Od ovog reda Err je MRTAV: 'On Error Resume Next' ga resetuje (dokazano
    ' testom 68). Zato se ne sme zvati LogErr, koji pise samo kad je
    ' Err.Number <> 0 -- tako je pad upisa godinama zavrsavao u PRAZNOM logu, a
    ' operater dobijao poruku bez razloga. Opis se predaje IZRICITO.
    On Error Resume Next

    LogError "SaveOtpremnicaMulti_TX", errDesc, errNum

    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnicaMulti_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnicaMulti_TX, _
        correlationId:=brojOtp, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveOtpremnicaMulti_TX failed. BrojOtp=" & brojOtp & _
                 "; BrojZbirne=" & brojZbirne & _
                 "; StanicaID=" & stanicaID & _
                 "; VozacID=" & vozacID & _
                 "; HasKlasaII=" & CStr(hasKlasaII) & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnicaMulti_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnicaMulti_TX, _
        correlationId:=brojOtp

    If Not tx Is Nothing Then tx.RollbackTx

    On Error GoTo 0

    SaveOtpremnicaMulti_TX = ""
End Function

Public Function SaveOtpremnica_TX(ByVal datum As Date, ByVal stanicaID As String, _
                                   ByVal vozacID As String, ByVal brojOtp As String, _
                                   ByVal brojZbirne As String, ByVal vrsta As String, _
                                   ByVal sorta As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, _
                                   Optional ByVal klasa As String = "I", _
                                   Optional ByVal brutoKg As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. SaveOtkupMulti_TX).
    modSchema.SchemaReadyOrFail "SaveOtpremnica_TX", _
        TBL_OTPREMNICA & "|" & TBL_AMBALAZA

tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtpremnica_TX = SaveOtpremnica(datum, stanicaID, vozacID, brojOtp, _
                                        brojZbirne, vrsta, sorta, kolicina, _
                                        cena, tipAmb, kolAmb, klasa, brutoKg)

    If SaveOtpremnica_TX = "" Then
        Err.Raise vbObjectError + 1001, "SaveOtpremnica_TX", _
                  "SaveOtpremnica fehlgeschlagen"
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
    LogError "SaveOtpremnica_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnica_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnica_TX, _
        correlationId:=brojOtp, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveOtpremnica_TX failed. BrojOtp=" & brojOtp & _
                "; BrojZbirne=" & brojZbirne & _
                "; StanicaID=" & stanicaID & _
                "; VozacID=" & vozacID & _
                "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnica_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnica_TX, _
        correlationId:=brojOtp

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtpremnica_TX = ""

    PrintTxFailure "SaveOtpremnica_TX", errSrc, errNum, errDesc
End Function

Public Function SaveOtpremnica(ByVal datum As Date, ByVal stanicaID As String, _
                               ByVal vozacID As String, ByVal brojOtp As String, _
                               ByVal brojZbirne As String, ByVal vrsta As String, _
                               ByVal sorta As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, _
                               Optional ByVal klasa As String = "I", _
                               Optional ByVal brutoKg As Double = 0) As String

    On Error GoTo EH

    Call ValidateOtpremnicaInput(stanicaID, vozacID, brojOtp, brojZbirne, _
                             kolicina, cena, tipAmb, kolAmb, klasa)

    Dim newID As String
    newID = GetNextID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-")

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 1002, "SaveOtpremnica", _
                "GetNextID nije vratio OtpremnicaID."
    End If

    Dim rowData As Variant
    rowData = Array(newID, datum, stanicaID, vozacID, brojOtp, _
                    brojZbirne, vrsta, sorta, kolicina, cena, tipAmb, kolAmb, klasa)

    Dim newRow As Long
    newRow = AppendRow(TBL_OTPREMNICA, rowData)
    If newRow > 0 Then
        ' Generacija: nasledjuje se od aktivnih redova istog broja (Klasa I <-> II),
        ' inace nova. Vazi za sve pozivaoce -- Multi_TX i pojedinacne _TX putanje.
        ' ZBR-CHILD-01: generacija RODITELJSKE zbirne na ovom detetu. Prazna je
        ' pravilo, ne izuzetak: u malina/hladnjaca lancu otpremnica nastaje PRE
        ' zbirne. Ide kroz PoveziDeteNaZbirnu iako je broj vec u rowData: da
        ' JEDINI PUT bude stvarno jedini -- prepis iste vrednosti je jeftin,
        ' a druga putanja bi znacila da se par moze raziciti.
        ' zbirne (modAutoHladnjaca), pa roditelj tada jos ne postoji -- popunice
        ' je LinkZbirnaToOtkupAndOtpremnica ili backfill.
        PoveziDeteNaZbirnu TBL_OTPREMNICA, newRow, COL_OTP_BROJ_ZBIRNE, brojZbirne, _
                           ZbirnaGeneracijaZaBroj(brojZbirne), "modDokumenta.SaveOtpremnica"
        ApplyGeneracijaID TBL_OTPREMNICA, newRow, COL_OTP_BROJ, brojOtp, _
                          COL_OTP_STANICA, stanicaID

        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", stanicaID, "Stanica", vozacID, newID, DOK_TIP_OTPREMNICA
        End If
        ' Bruto (kad je unet bruto pa oduzeta ambalaza) -> upis po imenu; prazno = neto.
        If brutoKg > 0 Then UpdateCell TBL_OTPREMNICA, newRow, COL_OTP_BRUTO, brutoKg
        SaveOtpremnica = newID
    Else
        Err.Raise vbObjectError + 1003, "SaveOtpremnica", _
                  "AppendRow fehlgeschlagen fuer tblOtpremnica"
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
    LogError "SaveOtpremnica", errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, "SaveOtpremnica", _
              "Source=" & errSrc & " | " & errDesc
End Function

Public Function GetOtpremniceByZbirna(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
            "modDokumenta.GetOtpremniceByZbirna"), "=", brojZbirne
    filters.Add fp

    GetOtpremniceByZbirna = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByZbirna"
    GetOtpremniceByZbirna = Empty
End Function

' True ako zbirna sa datim brojem postoji u tblZbirna (iskljucujuci stornirane).
' Storno-aware namerno: prijemnica ne sme da referencira storniranu zbirnu.
' (modBrojevi.BrojZbirneExists je Private i broji i stornirane -> drugi scenario.)
Public Function ZbirnaPostoji(ByVal brojZbirne As String) As Boolean
    On Error GoTo EH

    Dim b As String: b = Trim$(brojZbirne)
    If Len(b) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long
    iBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modDokumenta.ZbirnaPostoji")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(r, iBroj))), b, vbTextCompare) = 0 Then
            ZbirnaPostoji = True
            Exit Function
        End If
    Next r
    Exit Function

EH:
    LogErr "modDokumenta.ZbirnaPostoji", "broj=" & brojZbirne
End Function

' Normalizacija broja zbirne -- JEDNO mesto za ceo ZBR-IDENT lanac.
'
' CheckDuplicate poredi SIROVO (CStr(...) = searchValue, bez Trim, case-sensitive),
' pa " 5/070926 " prolazi pored "5/070926". Ta rupa se ne zatvara u njemu (D2 --
' preskakanje storniranih nosi ispravka-workflow na svih 6 tipova), nego OVDE.
Public Function ZbirnaBrojNorm(ByVal broj As String) As String
    ZbirnaBrojNorm = Trim$(NzToText(broj))
End Function

' Vlasnik zbirne je KOMPOZITAN: VozacID + KupacID. Isti spisak koji koriste
' ApplyGeneracijaID i modStorno.VlasniciPoBroju.
Private Function ZbirnaVlasnikKljuc(ByVal vozacID As Variant, _
                                    ByVal kupacID As Variant) As String
    ZbirnaVlasnikKljuc = UCase$(Trim$(NzToText(vozacID))) & "|" & _
                         UCase$(Trim$(NzToText(kupacID)))
End Function

' Razresava BROJ zbirne u LOGICKI DOKUMENT. Puni ceo DTO iz jednog citanja
' tabele: sirov niz nosi IKAD, ExcludeStornirano nad njim nosi SADA.
'
' Zasto ExcludeStornirano a ne sopstveni test na kolonu Stornirano: to je ISTI
' filtar koji zovu ZbirnaPostoji i GeneracijaIDZaBrojArr. Resolver, picker i
' writer time gledaju istu definiciju "aktivne" i ne mogu da se raziidju.
'
' Zasto NE modStorno.VlasniciPoBroju, koja je istog oblika: ona poredi broj
' case-sensitive (ZBR-NORM-02) i broji samo vlasnike, ne generacije. Ovde su
' potrebna oba, pa i normalizacija mora da bude sopstvena.
'
' FAIL-CLOSED: kad integrityStatus nije OK, resolutionStatus NIKAD nije NONE.
' NONE je jedina vrednost koja negde znaci "sme se" (kapija za nov unos), pa se
' ne sme dobiti iz greske. Aktivan red bez generacije zato daje CURRENT_AMBIGUOUS
' uz INTEGRITY_ERROR -- oba znace tvrdu blokadu, a integrityStatus kaze zasto.
' Identitet se NE pogadja iz broja i vlasnika (D4: nema legacy fallbacka).
Public Function ZbirnaIdentResolve(ByVal broj As String, _
                                   Optional ByVal vozacID As String = "", _
                                   Optional ByVal kupacID As String = "") As ZbirnaIdent
    Const SRC As String = "modDokumenta.ZbirnaIdentResolve"

    Dim res As ZbirnaIdent
    res.integrityStatus = ZBR_INT_OK
    res.resolutionStatus = ZBR_RES_NONE
    res.normalizedBroj = ZbirnaBrojNorm(broj)
    res.scopeProvided = (Len(Trim$(NzToText(vozacID))) > 0) And _
                        (Len(Trim$(NzToText(kupacID))) > 0)

    On Error GoTo EH

    If Len(res.normalizedBroj) = 0 Then GoTo XIT

    ' I1: da li broj vec drzi AKTIVNA prijemnica. Cita se iz kanonskog
    ' read-modela, da postoji JEDNA definicija "zauzetog broja".
    Dim unija As Object
    Set unija = AktivniBrojeviZbirne()
    If unija.Exists(res.normalizedBroj) Then
        res.brojUAktivnojPrijemnici = _
            (InStr(1, CStr(unija(res.normalizedBroj)), "P", vbBinaryCompare) > 0)
    End If

    Dim sirovo As Variant
    sirovo = GetTableData(TBL_ZBIRNA)
    If Not IsArray(sirovo) Then GoTo XIT

    Dim cBr As Long, cVoz As Long, cKup As Long, cGen As Long
    cBr = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    cVoz = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, SRC)
    cKup = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, SRC)
    cGen = RequireColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID, SRC)

    Dim scopeKljuc As String
    If res.scopeProvided Then scopeKljuc = ZbirnaVlasnikKljuc(vozacID, kupacID)

    ' --- IKAD: sirov niz, stornirani se BROJE ---
    Dim ikadVl As Object: Set ikadVl = CreateObject("Scripting.Dictionary")
    Dim ikadGen As Object: Set ikadGen = CreateObject("Scripting.Dictionary")
    Dim r As Long, vl As String, ikadG As String
    For r = 1 To UBound(sirovo, 1)
        If StrComp(Trim$(NzToText(sirovo(r, cBr))), res.normalizedBroj, vbTextCompare) = 0 Then
            vl = ZbirnaVlasnikKljuc(sirovo(r, cVoz), sirovo(r, cKup))
            If Not ikadVl.Exists(vl) Then ikadVl.Add vl, 1
            ikadG = Trim$(NzToText(sirovo(r, cGen)))
            If Len(ikadG) > 0 Then
                If Not ikadGen.Exists(ikadG) Then ikadGen.Add ikadG, 1
            End If
        End If
    Next r
    res.historicalOwnerCount = ikadVl.Count
    ' MERI SE, NE BLOKIRA -- v. komentar uz ZBR_MUT_* konstante.
    res.historicalLogicalCount = ikadGen.Count
    If res.historicalLogicalCount = 1 Then res.historicalOnlyGeneracijaID = ikadGen.Keys()(0)

    If res.scopeProvided And res.historicalOwnerCount = 1 Then
        res.historicalOwnerIsScope = (StrComp(ikadVl.Keys()(0), scopeKljuc, vbTextCompare) = 0)
    End If

    ' --- SADA: isti niz kroz ExcludeStornirano ---
    Dim akt As Variant
    akt = ExcludeStornirano(sirovo, TBL_ZBIRNA)
    If Not IsArray(akt) Then GoTo XIT

    Dim aktGen As Object: Set aktGen = CreateObject("Scripting.Dictionary")
    Dim aktVl As Object: Set aktVl = CreateObject("Scripting.Dictionary")
    Dim genVoz As Object: Set genVoz = CreateObject("Scripting.Dictionary")
    Dim genKup As Object: Set genKup = CreateObject("Scripting.Dictionary")
    Dim scopeGen As Object: Set scopeGen = CreateObject("Scripting.Dictionary")

    Dim gen As String, prazneAktivne As Long
    For r = 1 To UBound(akt, 1)
        If StrComp(Trim$(NzToText(akt(r, cBr))), res.normalizedBroj, vbTextCompare) = 0 Then
            vl = ZbirnaVlasnikKljuc(akt(r, cVoz), akt(r, cKup))
            If Not aktVl.Exists(vl) Then aktVl.Add vl, 1

            gen = Trim$(NzToText(akt(r, cGen)))
            If Len(gen) = 0 Then
                prazneAktivne = prazneAktivne + 1
            Else
                If Not aktGen.Exists(gen) Then
                    aktGen.Add gen, 1
                    genVoz.Add gen, Trim$(NzToText(akt(r, cVoz)))
                    genKup.Add gen, Trim$(NzToText(akt(r, cKup)))
                End If
                If res.scopeProvided Then
                    If StrComp(vl, scopeKljuc, vbTextCompare) = 0 Then
                        If Not scopeGen.Exists(gen) Then scopeGen.Add gen, 1
                    End If
                End If
            End If
        End If
    Next r

    res.activeLogicalCount = aktGen.Count
    res.activeOwnerCount = aktVl.Count
    res.matchingScopeActiveLogicalCount = scopeGen.Count

    ' ZBR-IDENT-01: aktivan red MORA da nosi generaciju. Prazna nije alternativni
    ' oblik identiteta nego integritetska greska -- sva tri writer-a u tblZbirna
    ' (SaveZbirna, modMasterSync, modDokumentInvariant) odmah PECATE validan
    ' GeneracijaID: prva dva ga NASLEDJUJU u svom scope-u, MasterSync ga KUJE.
    If prazneAktivne > 0 Then
        res.integrityStatus = ZBR_INT_ERROR
        res.resolutionStatus = ZBR_RES_AMBIGUOUS
        GoTo XIT
    End If

    If res.activeLogicalCount = 0 Then
        res.resolutionStatus = ZBR_RES_NONE
    ElseIf res.activeLogicalCount > 1 Then
        res.resolutionStatus = ZBR_RES_AMBIGUOUS
    ElseIf res.scopeProvided And res.matchingScopeActiveLogicalCount = 0 Then
        res.resolutionStatus = ZBR_RES_OWNER_MISMATCH
    Else
        res.resolutionStatus = ZBR_RES_UNIQUE
        gen = aktGen.Keys()(0)
        res.selectedGeneracijaID = gen
        res.selectedVozacID = genVoz(gen)
        res.selectedKupacID = genKup(gen)
    End If

XIT:
    ZbirnaIdentResolve = res
    Exit Function

EH:
    LogErr SRC, "broj=" & broj
    res.integrityStatus = ZBR_INT_ERROR
    res.resolutionStatus = ZBR_RES_AMBIGUOUS
    res.selectedGeneracijaID = ""
    res.selectedVozacID = ""
    res.selectedKupacID = ""
    ZbirnaIdentResolve = res
End Function

' Kapija ZBR-ACTIVE-NUMBER-01 (ugovor par.5) kao TABELA, na jednom mestu.
'
' Strogo, BEZ izuzetka za istog vlasnika: aktivan logicki dokument pod tim brojem
' znaci NE, ma ciji bio. Isti vlasnik sme tek posle storna, kad aktivnih nema.
' To nista ne lomi: ZbirnaValidiraj se zove TACNO jednom (modScrDokumenti
' Scr_Save), pre SaveZbirnaMulti_TX, pa validator nikad ne vidi red koji je sam
' upravo napisao; izmene zbirne u mestu nema -- ispravka je storno pa nov unos.
Public Function ZbirnaNovUnosRazlog(ByRef id As ZbirnaIdent) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaNovUnosRazlog = ZBR_GATE_INTEGRITET
        Exit Function
    End If

    If id.activeLogicalCount > 0 Then
        ZbirnaNovUnosRazlog = ZBR_GATE_AKTIVNA
        Exit Function
    End If

    If id.historicalOwnerCount = 0 Then
        ' Nijedna zbirna IKAD pod tim brojem -- slobodno, OSIM ako ga vec drzi
        ' aktivna prijemnica (I1). Tada bi nova zbirna tiho postala njen
        ' roditelj: prijemnica se vezuje samo brojem i ne bi ni primetila.
        If id.brojUAktivnojPrijemnici Then ZbirnaNovUnosRazlog = ZBR_GATE_SIROCE
        Exit Function
    End If

    ' Aktivnih nema, a istorija postoji: ispravku sme SAMO isti vlasnik.
    If Not id.historicalOwnerIsScope Then ZbirnaNovUnosRazlog = ZBR_GATE_TUDJ
End Function

' Tanak omotac nad JEDNOM tabelom iznad -- da druga kopija pravila ne odluta.
Public Function ZbirnaSmeNovUnos(ByRef id As ZbirnaIdent) As Boolean
    ZbirnaSmeNovUnos = (Len(ZbirnaNovUnosRazlog(id)) = 0)
End Function

' Ugovor par.6: prijemnica se vezuje SAMO na jednoznacno razresen dokument.
' Kod CURRENT_AMBIGUOUS / OWNER_MISMATCH / INTEGRITY_ERROR se kanonski roditelj
' NE trazi -- biranje "najverovatnijeg" iz dvosmislenog skupa je tiho pogadjanje.
'
' ISTORIJA JE DEO BEZBEDNOSTI, ne samo sadasnje stanje. UNIQUE danas uz broj koji
' je IKAD drzalo vise vlasnika i dalje nije bezbedan roditelj: nizvodna operacija
' koja ide po broju moze da zahvati i tudje. Prijemnica od faze 1 ima FK na
' generaciju, ali on sme biti prazan, pa sam po sebi ne ukida ovu kapiju.
' Zato postoji i modStorno.RequireJedanVlasnikIkadPoBroju -- ista kapija za
' mutaciju po broju.
' ZBR-MUT-01 -- sme li se po BROJU mutirati ono sto visi o zbirni.
'
' Trece pitanje, uz kapiju za kreiranje (F3) i za roditelja (F4). Ovde se ne pita
' "sme li nov unos" ni "koji je roditelj", nego: SME LI RUTINA KOJA DECU BIRA PO
' BROJU da radi. Izbor po broju zahvata sve dokumente tog broja; generacija na
' detetu postoji od faze 1, ali sme biti prazna, pa je popustanje uslovljeno --
' v. scopedPoGeneraciji.
'
' Vraca prazno kad sme. Razlog se ne stapa u jednu poruku: vise vlasnika i vise
' dokumenata istog vlasnika su dva razlicita poteza za operatera.
'
' Namerno NE gleda historicalLogicalCount -- v. komentar uz ZBR_MUT_* konstante.
' `scopedPoGeneraciji` je faza 4: kapija SME da pusti dva aktivna dokumenta pod
' istim brojem, ali samo tamo gde akter vise ne bira decu po broju.
'
' Default je False i to nije opreznost nego nuznost: kapiju zove DEVET mesta, a
' faza 3 je na generaciju prebacila TRI. Ostali i dalje biraju po broju --
' RecalculateZbirnaFromOtpremnice_TX preko SumOtpremniceByKlasa sabira SVE
' otpremnice pod brojem. Da je ova grana bezuslovna, zbirna bi dobila zbir tudjeg
' dokumenta u zaglavlje, tiho i u kilogramima. To je ZBR-MUT-01 naopako: ne
' sirenjem aktera nego suzavanjem kapije.
'
' Pozivalac NE sme da salje "postoji generacija" nego BAS onu odluku koju vec
' racuna za svoju selekciju (`genEff <> ""`). Kapija i akter tako gledaju isti
' izraz, ne dva slicna.
Public Function ZbirnaMutacijaPoBrojuRazlog(ByRef id As ZbirnaIdent, _
                                            Optional ByVal scopedPoGeneraciji As Boolean = False) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_INTEGRITET
        Exit Function
    End If

    ' IKAD, ne samo sada: storniran vlasnik i dalje moze imati AKTIVNU decu, a
    ' ona nose isti broj.
    If id.historicalOwnerCount > 1 Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_VISE_VLASNIKA
        Exit Function
    End If

    ' SADA: samo aktivni dokumenti konkurisu za decu. Dva aktivna istog vlasnika
    ' (A17) owner-brojac ne vidi -- zbog toga ova grana i postoji.
    '
    ' Faza 4: kad akter bira decu po generaciji, "dva aktivna dokumenta" mu vise
    ' nije opasnost -- dira samo svoje. Grana tada nema sta da brani.
    If scopedPoGeneraciji Then Exit Function

    If id.activeLogicalCount > 1 Then
        ZbirnaMutacijaPoBrojuRazlog = ZBR_MUT_VISE_DOKUMENATA
    End If
End Function

' Isto, ali od samog broja -- za pozivaoce koji nemaju razresen DTO.
'
' FAIL-CLOSED na sopstvenu gresku: "ne mogu da dokazem jednoznacnost" je za
' kapiju isto sto i "ne mutiraj". Prazan broj nije nerazresen nego "nema
' roditelja" -- nema sta da se mutira.
Public Function ZbirnaMutacijaPoBrojuRazlogZaBroj(ByVal broj As String, _
                                                 Optional ByVal scopedPoGeneraciji As Boolean = False) As String
    Dim id As ZbirnaIdent
    On Error GoTo EH
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    id = ZbirnaIdentResolve(broj)
    ZbirnaMutacijaPoBrojuRazlogZaBroj = ZbirnaMutacijaPoBrojuRazlog(id, scopedPoGeneraciji)
    Exit Function
EH:
    LogErr "modDokumenta.ZbirnaMutacijaPoBrojuRazlogZaBroj"
    ZbirnaMutacijaPoBrojuRazlogZaBroj = ZBR_MUT_INTEGRITET
End Function

Public Function ZbirnaRoditeljRazlog(ByRef id As ZbirnaIdent) As String
    If id.integrityStatus <> ZBR_INT_OK Then
        ZbirnaRoditeljRazlog = ZBR_GATE_INTEGRITET
        Exit Function
    End If

    Select Case id.resolutionStatus
        Case ZBR_RES_UNIQUE
            ' Jednoznacan DANAS jos nije dovoljno -- v. komentar iznad.
            If id.historicalOwnerCount > 1 Then
                ZbirnaRoditeljRazlog = ZBR_PARENT_ISTORIJA
            End If
        Case ZBR_RES_AMBIGUOUS
            ZbirnaRoditeljRazlog = ZBR_PARENT_DVOSMISLEN
        Case ZBR_RES_OWNER_MISMATCH
            ZbirnaRoditeljRazlog = ZBR_PARENT_TUDJ
        Case Else
            ' NONE: pod tim brojem nema aktivnog dokumenta, pa nema ni roditelja.
            ' Politiku za taj slucaj (BLOK ili UPOZORENJE, po
            ' PRIJEMNICA_ZBIRNA_PROVERA) drzi ZbirnaPostoji u validatoru IZNAD
            ' ove kapije -- ovde stoji samo da funkcija ne laze kad se pozove sama.
            ZbirnaRoditeljRazlog = ZBR_PARENT_NEMA
    End Select
End Function

' Tanak omotac nad JEDNOM tabelom iznad -- da druga kopija pravila ne odluta.
Public Function ZbirnaRoditeljOK(ByRef id As ZbirnaIdent) As Boolean
    ZbirnaRoditeljOK = (Len(ZbirnaRoditeljRazlog(id)) = 0)
End Function

' I1 -- KANONSKI READ-MODEL AKTIVNIH BROJEVA ZBIRNE (ugovor par.8).
'
' Skup = normalizovani brojevi AKTIVNIH zbirnih UNIJA normalizovani BrojZbirne
' iz AKTIVNIH prijemnica.
'
' Unija nije opreznost nego ispravka rupe: broj koji referencira aktivna
' prijemnica je ZAUZET i onda kad mu je zbirna-red storniran ili nikad nije
' upisan. Bez tog drugog izvora takav broj izgleda slobodan.
'
' Vrednost kaze ODAKLE broj dolazi -- "Z", "P" ili "ZP" -- da pozivalac moze da
' razlikuje siroce (samo prijemnica) od normalnog para.
Public Function AktivniBrojeviZbirne() As Object
    Const SRC As String = "modDokumenta.AktivniBrojeviZbirne"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AktivniBrojeviZbirne = res

    On Error GoTo EH

    DodajBrojeve res, TBL_ZBIRNA, COL_ZBR_BROJ, "Z", SRC
    DodajBrojeve res, TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "P", SRC
    Exit Function

EH:
    LogErr SRC
End Function

' Jedan izvor u recnik. Oznaka se DODAJE na postojecu, pa broj koji nose obe
' tabele dobije "ZP" -- razlika siroce/par se ne gubi.
Private Sub DodajBrojeve(ByRef res As Object, ByVal tblName As String, _
                         ByVal brojCol As String, ByVal oznaka As String, _
                         ByVal sourceName As String)
    Dim data As Variant
    data = GetTableData(tblName)
    If Not IsArray(data) Then Exit Sub
    data = ExcludeStornirano(data, tblName)
    If Not IsArray(data) Then Exit Sub

    Dim cBr As Long
    cBr = RequireColumnIndex(tblName, brojCol, sourceName)

    Dim r As Long, b As String
    For r = 1 To UBound(data, 1)
        b = ZbirnaBrojNorm(NzToText(data(r, cBr)))
        If Len(b) > 0 Then
            If res.Exists(b) Then
                If InStr(1, CStr(res(b)), oznaka, vbBinaryCompare) = 0 Then
                    res(b) = CStr(res(b)) & oznaka
                End If
            Else
                res.Add b, oznaka
            End If
        End If
    Next r
End Sub

Public Function GetOtpremniceByStation(ByVal stanicaID As String, _
                                       Optional ByVal datumOd As Date = 0, _
                                       Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, _
            "modDokumenta.GetOtpremniceByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetOtpremniceByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtpremniceByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByStation"
    GetOtpremniceByStation = Empty
End Function

' ============================================================
' ZBIRNA - Gesamtdokument Fahrer
' ============================================================
Public Function SaveZbirnaMulti_TX(ByVal datum As Date, _
                                   ByVal vozacID As String, _
                                   ByVal brojZbirne As String, _
                                   ByVal kupacID As String, _
                                   ByVal hladnjaca As String, _
                                   ByVal pogon As String, _
                                   ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, _
                                   ByVal ukupnoKolI As Double, _
                                   ByVal tipAmb As String, _
                                   ByVal ukupnoAmb As Long, _
                                   Optional ByVal hasKlasaII As Boolean = False, _
                                   Optional ByVal ukupnoKolII As Double = 0, _
                                   Optional ByVal ukupnoAmbII As Long = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO, pa tabela sa kolonom manje ili
    ' u pogresnom rasporedu tiho salje vrednosti u pogresna polja. To je gore od
    ' pada upisa -- greska nastaje u podacima, ne u logu.
    '
    ' Ide PRE BeginTx: kapija sme da digne gresku, a nema smisla otvarati
    ' transakciju koja se odmah rollback-uje.
    modSchema.SchemaReadyOrFail "SaveZbirnaMulti_TX", TBL_ZBIRNA

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    ' Klasa I je opciona (ukupnoKolI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (ukupnoKolI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1203, "SaveZbirnaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
        resultI = SaveZbirna( _
            datum, _
            vozacID, _
            brojZbirne, _
            kupacID, _
            hladnjaca, _
            pogon, _
            vrstaVoca, _
            sortaVoca, _
            ukupnoKolI, _
            tipAmb, _
            ukupnoAmb, _
            KLASA_I)

        If resultI = "" Then
            Err.Raise vbObjectError + 1201, "SaveZbirnaMulti_TX", _
                      "SaveZbirna Klasa I fehlgeschlagen"
        End If
    End If

    Dim resultII As String
    If hasKlasaII Then
        resultII = SaveZbirna( _
            datum, _
            vozacID, _
            brojZbirne, _
            kupacID, _
            hladnjaca, _
            pogon, _
            vrstaVoca, _
            sortaVoca, _
            ukupnoKolII, _
            tipAmb, _
            ukupnoAmbII, _
            KLASA_II)

        If resultII = "" Then
            Err.Raise vbObjectError + 1202, "SaveZbirnaMulti_TX", _
                      "SaveZbirna Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SaveZbirnaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SaveZbirnaMulti_TX = resultI
    Else
        SaveZbirnaMulti_TX = resultII
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
    LogError "SaveZbirnaMulti_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirnaMulti_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirnaMulti_TX, _
        correlationId:=brojZbirne, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveZbirnaMulti_TX failed. BrojZbirne=" & brojZbirne & _
             "; VozacID=" & vozacID & _
             "; KupacID=" & kupacID & _
             "; HasKlasaII=" & CStr(hasKlasaII) & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirnaMulti_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirnaMulti_TX, _
        correlationId:=brojZbirne

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveZbirnaMulti_TX = ""

    PrintTxFailure "SaveZbirnaMulti_TX", errSrc, errNum, errDesc
End Function

Public Function SaveZbirna_TX(ByVal datum As Date, ByVal vozacID As String, _
                               ByVal brojZbirne As String, ByVal kupacID As String, _
                               ByVal hladnjaca As String, ByVal pogon As String, _
                               ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                               ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                               ByVal ukupnoAmb As Long, _
                               Optional ByVal klasa As String = "I") As String
    Dim tx As New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. SaveOtkupMulti_TX).
    modSchema.SchemaReadyOrFail "SaveZbirna_TX", _
        TBL_ZBIRNA

tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    SaveZbirna_TX = SaveZbirna(datum, vozacID, brojZbirne, kupacID, _
                                hladnjaca, pogon, vrstaVoca, sortaVoca, _
                                ukupnoKol, tipAmb, ukupnoAmb, klasa)

    If SaveZbirna_TX = "" Then
        Err.Raise vbObjectError + 1004, "SaveZbirna_TX", _
                  "SaveZbirna fehlgeschlagen"
    End If

    tx.CommitTx

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "SaveZbirna_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirna_TX, _
        correlationId:=brojZbirne, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveZbirna_TX failed. BrojZbirne=" & brojZbirne & _
             "; VozacID=" & vozacID & _
             "; KupacID=" & kupacID & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirna_TX, _
        correlationId:=brojZbirne

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveZbirna_TX = ""

    PrintTxFailure "SaveZbirna_TX", errSrc, errNum, errDesc
End Function

Public Function SaveZbirna(ByVal datum As Date, ByVal vozacID As String, _
                           ByVal brojZbirne As String, ByVal kupacID As String, _
                           ByVal hladnjaca As String, ByVal pogon As String, _
                           ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                           ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                           ByVal ukupnoAmb As Long, _
                           Optional ByVal klasa As String = "I") As String
    On Error GoTo EH

    Call ValidateZbirnaInput(vozacID, brojZbirne, kupacID, ukupnoKol, _
                         tipAmb, ukupnoAmb, klasa)

    Dim newID As String
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")

    If newID = "" Then
        Err.Raise vbObjectError + 1009, "SaveZbirna", _
                  "GetNextID nije vratio ZbirnaID."
    End If

    ' AUD-003: upis PO IMENU kolone (BuildZbirnaRowData), ne pozicijski Array(...).
    ' Presence guard (RequireColumns) dokazuje samo da kolone postoje; promenjen
    ' redosled ili kolona umetnuta u sredinu bi kroz pozicijski upis tiho iskrivili
    ' red. Isti obrazac kao BuildPrijemnicaRowData.
    Dim rowData As Variant
    rowData = BuildZbirnaRowData(newID, datum, vozacID, brojZbirne, kupacID, _
                                 hladnjaca, pogon, vrstaVoca, sortaVoca, _
                                 ukupnoKol, tipAmb, ukupnoAmb, klasa)

    Dim newRowZbr As Long
    newRowZbr = AppendRow(TBL_ZBIRNA, rowData)

    If newRowZbr > 0 Then
        ApplyGeneracijaID TBL_ZBIRNA, newRowZbr, COL_ZBR_BROJ, brojZbirne, _
                          COL_ZBR_VOZAC, vozacID, COL_ZBR_KUPAC, kupacID

        SaveZbirna = newID
    Else
        Err.Raise vbObjectError + 1010, "SaveZbirna", _
                  "AppendRow fehlgeschlagen fuer tblZbirna."
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
    LogError "SaveZbirna", errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, "SaveZbirna", _
              "Source=" & errSrc & " | " & errDesc
End Function

' ============================================================
' ZBIRNA -- header + stavke (PR3)
' ============================================================
'
' Jedan javni ulaz koji vraca JEDAN ID, i koji obuhvata CELU poslovnu operaciju.
' Obrazac je CreateFaktura_TX (modFaktura.bas:11): _TX drzi transakciju i
' monitoring, Private core radi posao, kompletna prevalidacija PRE ijednog upisa.
'
' Sta se menja u odnosu na SaveZbirnaMulti_TX:
'
'   staro:  dva reda u tblZbirna (Klasa I i Klasa II), dva ID-a, pa string
'           "ZBR-1 + ZBR-2" koji pozivalac posle parsira
'   novo:   JEDAN header u tblZbirna + N redova u tblZbirnaStavke, jedan ID
'
' KOLICINE SE NE PRIMAJU -- IZVODE SE.
'
' Zbirna je agregat svojih otpremnica (DOCUMENT_HEADER_LINES.md S4.3: stavke su
' kes, otpremnice su izvor istine). Zato writer prima IZVORNE OTPREMNICE, a
' stavke racuna iz njih. Prva verzija je primala gotove stavke, pa je bilo
' legalno napraviti zbirnu od izmisljenih 400+600 kg bez ijedne otpremnice --
' kes koji se ne slaze sa izvorom, i to kroz kanonski writer.
'
' Iz istog razloga VrstaVoca / SortaVoca / TipAmbalaze NISU u headeru: dolaze iz
' otpremnica, koje moraju biti saglasne. Dokument ima jednu vrstu, jednu sortu i
' jedan tip ambalaze (REFAKTOR_DOKUMENT_HEADER_STAVKE.md S2).
'
' CLANSTVO SE UPISUJE U ISTOJ TRANSAKCIJI.
'
' Kanonski sastav je tblZbirnaIzvori -- "od kojih je otpremnica ova VERZIJA
' sastavljena" (A15). Otpremnica.ZbirnaID je samo POKAZIVAC na trenutno aktivnu
' zbirnu, imenovan kes u smislu A5. Jedan FK ne moze da nosi istoriju: posle
' ispravke jedne otpremnice sestre koje se nisu menjale pripadaju i staroj i
' novoj verziji.
'
' Oboje ide kroz JEDAN snapshot -- ili sve, ili nista. Da writer ne upisuje
' clanstvo, cutover bi morao "commit; pa povezi", a pad drugog koraka ostavlja
' zbirnu bez izvora.
'
' ZbirnaID je OPAQUE (NewEntityID), ne GetNextID: broj vise nije identitet, pa
' ni ID ne sme da bude brojac po kome se pogadja "sledeci".
'
' GeneracijaID se NE pise. Ta masinerija je kompenzacija za nepostojeci header i
' brise se u PR4; nov pisac je ne sme ozivljavati.
'
' PR3 je ADITIVAN. Produkcioni pozivaoci (modDokUnos.bas:533, modAutoHladnjaca,
' modMasterSync) i dalje idu starim putem; cutover citalaca, invarijante i storna
' je PR4. Zato header NAMERNO ostavlja UkupnoKolicina / UkupnoAmbalaze / Klasa
' prazne -- to su kolone koje u ciljnoj semi ne postoje jer kolicina zivi na
' stavci. Ko ih procita dobija prazno, i to je tacan odgovor: nije "nula
' kilograma", nego "ne pitaj header za kolicinu".
'
' Iz istog razloga se NE pise ni Otpremnica.BrojZbirne: veza je ZbirnaID. Stara
' broj-veza ostaje netaknuta na dokumentima koje je napravio stari put.
'
' OVAJ WRITER PRAVI I ODMAH FINALIZUJE DOKUMENT (IzdatoStatus = IZDATO).
'
' Zato trazi bar jednu izvornu otpremnicu: zbirna bez izvora nije izdat dokument.
' Draft-first tok -- gde operater prvo otvori zbirnu pa vezuje otpremnice -- je
' zasebna funkcija koja jos ne postoji (CreateZbirnaDraft_TX). Dok je nema,
' kanonski tok je JEDAN: otpremnice postoje, pa se zbirna napravi i izda.
'
' Argumenti:
'   h                 Scripting.Dictionary. Obavezno: Datum, VozacID,
'                     BrojZbirne, KupacID. Opciono: Hladnjaca, Pogon.
'                     Nepoznat kljuc je GRESKA (v. HdrProveriKljuceve).
'   izvorOtpremnice   Collection OtpremnicaID-jeva. Bar jedan.
'   outGreska         RAZLOG odbijanja, ne samo cinjenica.
'   ocekivano         Opciono: Collection diktova {Klasa, Kolicina, KolAmbalaze}
'                     -- ono sto je operater OTKUCAO. Sluzi kao unakrsna
'                     provera protiv izvedenog; neslaganje obara upis.
'
' outGreska postoji jer se bez njega ne moze razlikovati "kapija je odbila upis"
' od "upis je pukao iz drugog razloga pa je ispalo isto". Nije teorijska
' razlika: sabotaza koja je iskljucila proveru duple klase ostavila je suite
' ZELEN, jer je posao preuzeo Dictionary.Add svojom greskom o duplom kljucu.
Public Function CreateZbirna_TX(ByVal h As Object, _
                                ByVal izvorOtpremnice As Collection, _
                                Optional ByRef outGreska As String, _
                                Optional ByVal ocekivano As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO, pa tabela sa kolonom manje ili u
    ' pogresnom rasporedu tiho salje vrednosti u pogresna polja. Ide PRE BeginTx:
    ' kapija sme da digne gresku, a nema smisla otvarati transakciju koja se
    ' odmah rollback-uje.
    modSchema.SchemaReadyOrFail "CreateZbirna_TX", _
        TBL_ZBIRNA & "|" & TBL_ZBIRNA_STAVKE & "|" & TBL_ZBIRNA_IZVORI & _
        "|" & TBL_OTPREMNICA

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI
    tx.AddTableSnapshot TBL_OTPREMNICA

    CreateZbirna_TX = CreateZbirna(h, izvorOtpremnice, ocekivano)

    If CreateZbirna_TX = "" Then
        Err.Raise vbObjectError + 1220, "CreateZbirna_TX", _
                  "CreateZbirna nije vratio ZbirnaID."
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
    LogError "CreateZbirna_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="CreateZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=CreateZbirna_TX, _
        correlationId:=CreateZbirna_TX, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="CreateZbirna_TX failed. Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="CreateZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=CreateZbirna_TX, _
        correlationId:=CreateZbirna_TX

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateZbirna_TX = ""
    outGreska = errDesc

    PrintTxFailure "CreateZbirna_TX", errSrc, errNum, errDesc
End Function

' Core -- NE zovi spolja. Jedini ulaz je CreateZbirna_TX, koji drzi snapshot
' transakciju; direktan poziv bi kod greske ostavio pola dokumenta.
Private Function CreateZbirna(ByVal h As Object, _
                              ByVal izvorOtpremnice As Collection, _
                              ByVal ocekivano As Collection) As String
    Const SRC As String = "CreateZbirna"

    On Error GoTo EH

    If h Is Nothing Then
        Err.Raise vbObjectError + 1221, SRC, "Header nije prosledjen."
    End If

    If izvorOtpremnice Is Nothing Then
        Err.Raise vbObjectError + 1222, SRC, _
                  "Izvorne otpremnice nisu prosledjene."
    End If

    If izvorOtpremnice.count = 0 Then
        Err.Raise vbObjectError + 1223, SRC, _
                  "Zbirna mora imati bar jednu izvornu otpremnicu."
    End If

    ' Fail-fast nad semom pre ijednog upisa.
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_ID, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_DATUM, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_VOZAC, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_BROJ, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KUPAC, SRC
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_TIP_AMB, SRC

    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_RB, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, SRC
    RequireColumnIndex TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, SRC

    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, SRC
    RequireColumnIndex TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, SRC

    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_ID, SRC
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_ZBIRNA_ID, SRC

    HdrProveriKljuceve h, SRC

    Dim datum As Date
    Dim vozacID As String
    Dim brojZbirne As String
    Dim kupacID As String

    datum = HdrDatum(h, "Datum", SRC)
    vozacID = HdrObavezan(h, "VozacID", SRC)
    brojZbirne = HdrObavezan(h, "BrojZbirne", SRC)
    kupacID = HdrObavezan(h, "KupacID", SRC)

    ' --- izvor: procitaj, proveri, izvedi ------------------------------------
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(data) Then
        Err.Raise vbObjectError + 1224, SRC, "Tabela otpremnica je prazna."
    End If

    Dim cID As Long, cKlasa As Long, cKol As Long, cAmb As Long
    Dim cVrsta As Long, cSorta As Long, cTip As Long
    Dim cStorno As Long, cZbrID As Long, cZbrBroj As Long
    Dim cVozac As Long

    cID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, SRC)
    cVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    cSorta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_SORTA, SRC)
    cTip = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_TIP_AMB, SRC)
    cVozac = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, SRC)
    cZbrID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ZBIRNA_ID, SRC)
    cZbrBroj = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)

    Dim redPoID As Object
    Dim kolPoKlasi As Object
    Dim ambPoKlasi As Object
    Set redPoID = CreateObject("Scripting.Dictionary")
    Set kolPoKlasi = CreateObject("Scripting.Dictionary")
    Set ambPoKlasi = CreateObject("Scripting.Dictionary")

    Dim vrsta As String, sorta As String, tipAmb As String
    Dim prvi As Boolean
    Dim i As Long, r As Long
    Dim otpID As String, klasa As String
    Dim kolRed As Double, ambRed As Double

    prvi = True

    For i = 1 To izvorOtpremnice.count
        otpID = Trim$(NzToText(izvorOtpremnice(i)))
        If Len(otpID) = 0 Then
            Err.Raise vbObjectError + 1225, SRC, _
                      "Prazan OtpremnicaID na poziciji " & CStr(i) & "."
        End If

        If redPoID.Exists(UCase$(otpID)) Then
            Err.Raise vbObjectError + 1226, SRC, _
                      "Ista otpremnica navedena dvaput: " & otpID
        End If

        r = NadjiJedanRedOtpremnice(data, cID, otpID, SRC)
        redPoID.Add UCase$(otpID), r

        ' IsStorniranoValue je Private u modStorno -- odavde nevidljiv.
        ' Poredjenje je case-insensitive, kao tamo.
        If StrComp(Trim$(NzToText(data(r, cStorno))), "Da", _
                   vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1227, SRC, _
                      "Otpremnica je stornirana: " & otpID
        End If

        ' Fail-closed na vec vezanu otpremnicu. BrojZbirne se gleda samo dok
        ' traje prelazni period: dokumenti koje je napravio stari put nose
        ' broj-vezu, i ne smeju se tiho preuzeti. PR4 brise tu polovinu provere
        ' zajedno sa broj-vezom.
        If Len(Trim$(NzToText(data(r, cZbrID)))) > 0 Then
            Err.Raise vbObjectError + 1228, SRC, _
                      "Otpremnica je vec na zbirnoj: " & otpID
        End If
        If cZbrBroj > 0 Then
            If Len(Trim$(NzToText(data(r, cZbrBroj)))) > 0 Then
                Err.Raise vbObjectError + 1229, SRC, _
                          "Otpremnica je vec na zbirnoj (broj): " & otpID
            End If
        End If

        ' Zbirna je JEDAN transport JEDNOG vozaca. Otpremnica drugog vozaca u
        ' istoj zbirnoj je korupcija domena, ne rubni slucaj: BrojZbirne je
        ' scoped po vozacu, pa bi takav dokument bio i nepretraziv.
        RequireIstoPolje vozacID, Trim$(NzToText(data(r, cVozac))), _
                         "VozacID", otpID, SRC

        ' Dokument ima JEDNU vrstu, sortu i tip ambalaze. Otpremnica koja se ne
        ' slaze ne pripada ovoj zbirnoj -- i to je strukturno, ne stvar ukusa.
        If prvi Then
            vrsta = Trim$(NzToText(data(r, cVrsta)))
            sorta = Trim$(NzToText(data(r, cSorta)))
            tipAmb = Trim$(NzToText(data(r, cTip)))
            prvi = False
        Else
            RequireIstoPolje vrsta, Trim$(NzToText(data(r, cVrsta))), _
                             "VrstaVoca", otpID, SRC
            RequireIstoPolje sorta, Trim$(NzToText(data(r, cSorta))), _
                             "SortaVoca", otpID, SRC
            RequireIstoPolje tipAmb, Trim$(NzToText(data(r, cTip))), _
                             "TipAmbalaze", otpID, SRC
        End If

        klasa = Trim$(NzToText(data(r, cKlasa)))
        RequireValidDocumentClass klasa, SRC
        klasa = UCase$(klasa)

        If Not IsNumeric(data(r, cKol)) Then
            Err.Raise vbObjectError + 1230, SRC, _
                      "Kolicina nije broj na otpremnici: " & otpID
        End If
        kolRed = CDbl(data(r, cKol))

        ambRed = 0
        If IsNumeric(data(r, cAmb)) Then ambRed = CDbl(data(r, cAmb))

        If Not kolPoKlasi.Exists(klasa) Then
            kolPoKlasi.Add klasa, 0#
            ambPoKlasi.Add klasa, 0#
        End If
        kolPoKlasi(klasa) = CDbl(kolPoKlasi(klasa)) + kolRed
        ambPoKlasi(klasa) = CDbl(ambPoKlasi(klasa)) + ambRed
    Next i

    ' --- izvedene stavke moraju biti smislene --------------------------------
    Dim klase As Collection
    Set klase = KlaseUKanonskomRedu(kolPoKlasi)

    Dim k As Long
    For k = 1 To klase.count
        klasa = CStr(klase(k))
        If CDbl(kolPoKlasi(klasa)) <= 0 Then
            Err.Raise vbObjectError + 1231, SRC, _
                      "Zbir kolicine za klasu " & klasa & " nije veci od nule."
        End If
        If CDbl(ambPoKlasi(klasa)) < 0 Then
            Err.Raise vbObjectError + 1232, SRC, _
                      "Zbir ambalaze za klasu " & klasa & " je negativan."
        End If
        ' Ambalaza je BROJ KOMADA, ne tezina. Legacy ValidateZbirnaInput je to
        ' drzao tipom (ukupnoAmb As Long); rec/nista drugo to vise ne cuva, pa
        ' se trazi ovde. Odbija se, ne zaokruzuje: "20.5 gajbica" je kvar u
        ' izvoru, a tiha ispravka ga sakriva.
        RequireCeoBroj CDbl(ambPoKlasi(klasa)), _
                       "Ambalaza za klasu " & klasa, SRC
    Next k

    RequireOcekivanoSeSlaze ocekivano, kolPoKlasi, ambPoKlasi, SRC

    ' --- upis ----------------------------------------------------------------
    Dim zbirnaID As String
    zbirnaID = NewEntityID("ZBR-")

    If zbirnaID = "" Then
        Err.Raise vbObjectError + 1233, SRC, _
                  "NewEntityID nije vratio ZbirnaID."
    End If

    Dim rowData As Variant
    rowData = BuildZbirnaHeaderRowData(zbirnaID, datum, vozacID, brojZbirne, _
                                       kupacID, HdrOpcion(h, "Hladnjaca"), _
                                       HdrOpcion(h, "Pogon"), vrsta, sorta, _
                                       tipAmb)

    If AppendRow(TBL_ZBIRNA, rowData) <= 0 Then
        Err.Raise vbObjectError + 1234, SRC, _
                  "AppendRow nije upisao header u tblZbirna."
    End If

    Dim stavkaID As String
    For k = 1 To klase.count
        klasa = CStr(klase(k))

        ' Fail-closed: NewEntityID vraca "" kad CoCreateGuid ne uspe. Red bez
        ' identiteta je gori od pada -- niko ga posle ne moze ni naci ni vezati.
        stavkaID = NewEntityID("ZBS-")
        If stavkaID = "" Then
            Err.Raise vbObjectError + 1235, SRC, _
                      "NewEntityID nije vratio ZbirnaStavkaID za klasu " & klasa & "."
        End If

        rowData = BuildZbirnaStavkaRowData(stavkaID, zbirnaID, k, klasa, _
                                           CDbl(kolPoKlasi(klasa)), _
                                           CDbl(ambPoKlasi(klasa)))

        If AppendRow(TBL_ZBIRNA_STAVKE, rowData) <= 0 Then
            Err.Raise vbObjectError + 1236, SRC, _
                      "AppendRow nije upisao stavku za klasu " & klasa & "."
        End If
    Next k

    ' --- clanstvo, u ISTOJ transakciji kao header i stavke -------------------
    '
    ' Dva zapisa, i NISU isti pojam:
    '
    '   tblZbirnaIzvori        od kojih je otpremnica ova VERZIJA sastavljena.
    '                          Nepromenljivo. Posle ispravke jedne otpremnice
    '                          nastaje nova verzija zbirne, a sestre koje se nisu
    '                          menjale pripadaju i staroj i novoj -- jedan FK to
    '                          ne moze da pokaze.
    '   Otpremnica.ZbirnaID    na kojoj je AKTIVNOJ zbirnoj otpremnica sada.
    '                          Izvedeno iz gornjeg (A5: imenovan kes), drzi se
    '                          zbog jeftine provere "vec vezana" i citalaca.
    Dim kljuc As Variant
    Dim izvorID As String

    For Each kljuc In redPoID.Keys
        izvorID = NewEntityID("ZBI-")
        If izvorID = "" Then
            Err.Raise vbObjectError + 1256, SRC, _
                      "NewEntityID nije vratio ZbirnaIzvorID."
        End If

        rowData = BuildZbirnaIzvorRowData(izvorID, zbirnaID, _
                                          IDIzRedaOtpremnice(data, cID, _
                                                             CLng(redPoID(kljuc))))
        If AppendRow(TBL_ZBIRNA_IZVORI, rowData) <= 0 Then
            Err.Raise vbObjectError + 1257, SRC, _
                      "AppendRow nije upisao clanstvo otpremnice."
        End If

        RequireUpdateCell TBL_OTPREMNICA, CLng(redPoID(kljuc)), _
                          COL_OTP_ZBIRNA_ID, zbirnaID, SRC
    Next kljuc

    CreateZbirna = zbirnaID
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

' Tacno jedan red za dati OtpremnicaID. Nula i vise od jednog su oba greska:
' tihi "uzmi prvi pogodak" je klasa buga zbog koje CreateFaktura ima
' RequireSingleFakturaRow.
Private Function NadjiJedanRedOtpremnice(ByRef data As Variant, _
                                         ByVal colID As Long, _
                                         ByVal otpID As String, _
                                         ByVal src As String) As Long
    Dim i As Long, nadjen As Long, koliko As Long

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, colID))), otpID, vbTextCompare) = 0 Then
            nadjen = i
            koliko = koliko + 1
        End If
    Next i

    If koliko = 0 Then
        Err.Raise vbObjectError + 1237, src, _
                  "Otpremnica nije pronadjena: " & otpID
    End If
    If koliko > 1 Then
        Err.Raise vbObjectError + 1238, src, _
                  "Duplikat OtpremnicaID=" & otpID & "; Count=" & CStr(koliko)
    End If

    NadjiJedanRedOtpremnice = nadjen
End Function

Private Sub RequireIstoPolje(ByVal ocekivano As String, ByVal stvarno As String, _
                             ByVal polje As String, ByVal otpID As String, _
                             ByVal src As String)
    If StrComp(ocekivano, stvarno, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1239, src, _
                  "Otpremnica " & otpID & " ima drugo polje " & polje & _
                  ": ocekivano '" & ocekivano & "', naslo '" & stvarno & "'."
    End If
End Sub

Private Sub RequireCeoBroj(ByVal v As Double, ByVal opis As String, _
                           ByVal src As String)
    If Abs(v - Fix(v)) > 0.0000001 Then
        Err.Raise vbObjectError + 1240, src, _
                  opis & " mora biti ceo broj, a nije: " & Fmt2Zbr(v)
    End If
End Sub

' Klase u kanonskom redu: I, pa II, pa sve ostalo azbucno. RedniBroj mora biti
' determinisan -- Dictionary.Keys cuva redosled ubacivanja, a on zavisi od
' redosleda otpremnica u pozivu.
Private Function KlaseUKanonskomRedu(ByVal kolPoKlasi As Object) As Collection
    Dim c As Collection
    Set c = New Collection

    If kolPoKlasi.Exists(UCase$(KLASA_I)) Then c.Add UCase$(KLASA_I)
    If kolPoKlasi.Exists(UCase$(KLASA_II)) Then c.Add UCase$(KLASA_II)

    Dim kljuc As Variant
    Dim ostatak As Collection
    Set ostatak = New Collection

    For Each kljuc In kolPoKlasi.Keys
        If StrComp(CStr(kljuc), UCase$(KLASA_I), vbTextCompare) <> 0 And _
           StrComp(CStr(kljuc), UCase$(KLASA_II), vbTextCompare) <> 0 Then
            ostatak.Add CStr(kljuc)
        End If
    Next kljuc

    Dim i As Long, j As Long, tmp As String
    Dim arr() As String
    If ostatak.count > 0 Then
        ReDim arr(1 To ostatak.count)
        For i = 1 To ostatak.count
            arr(i) = CStr(ostatak(i))
        Next i
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            Next j
        Next i
        For i = 1 To UBound(arr)
            c.Add arr(i)
        Next i
    End If

    Set KlaseUKanonskomRedu = c
End Function

' Unakrsna provera: ono sto je operater otkucao mora da se slaze sa onim sto je
' izvedeno iz otpremnica. Neslaganje je danas tiha greska -- ekran prikaze svoje
' brojeve, tabela nosi druge.
Private Sub RequireOcekivanoSeSlaze(ByVal ocekivano As Collection, _
                                    ByVal kolPoKlasi As Object, _
                                    ByVal ambPoKlasi As Object, _
                                    ByVal src As String)
    If ocekivano Is Nothing Then Exit Sub
    If ocekivano.count = 0 Then Exit Sub

    Dim vidjene As Object
    Set vidjene = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim s As Object
    Dim klasa As String
    Dim kol As Double, amb As Double

    For i = 1 To ocekivano.count
        If Not IsObject(ocekivano(i)) Then
            Err.Raise vbObjectError + 1241, src, _
                      "Ocekivana stavka " & CStr(i) & " nije Dictionary."
        End If
        Set s = ocekivano(i)

        klasa = UCase$(Trim$(NzToText(StavkaVrednost(s, "Klasa", i, src))))
        If vidjene.Exists(klasa) Then
            Err.Raise vbObjectError + 1242, src, _
                      "Dve ocekivane stavke iste klase: " & klasa
        End If
        vidjene.Add klasa, True

        If Not kolPoKlasi.Exists(klasa) Then
            Err.Raise vbObjectError + 1243, src, _
                      "Ocekivana klasa " & klasa & " ne postoji na izvornim otpremnicama."
        End If

        kol = StavkaBroj(s, "Kolicina", i, src)
        amb = StavkaBroj(s, "KolAmbalaze", i, src)

        If Abs(kol - CDbl(kolPoKlasi(klasa))) > 0.001 Then
            Err.Raise vbObjectError + 1244, src, _
                      "Kolicina za klasu " & klasa & " se ne slaze sa otpremnicama: " & _
                      "uneto " & Fmt2Zbr(kol) & ", izvedeno " & Fmt2Zbr(CDbl(kolPoKlasi(klasa))) & "."
        End If
        If Abs(amb - CDbl(ambPoKlasi(klasa))) > 0.001 Then
            Err.Raise vbObjectError + 1245, src, _
                      "Ambalaza za klasu " & klasa & " se ne slaze sa otpremnicama: " & _
                      "uneto " & Fmt2Zbr(amb) & ", izvedeno " & Fmt2Zbr(CDbl(ambPoKlasi(klasa))) & "."
        End If
    Next i

    If vidjene.count <> kolPoKlasi.count Then
        Err.Raise vbObjectError + 1246, src, _
                  "Ocekivano ima " & CStr(vidjene.count) & " klasa, a otpremnice daju " & _
                  CStr(kolPoKlasi.count) & "."
    End If
End Sub

' Broj u poruku, nezavisno od Windows locale-a (decimalna tacka uvek).
Private Function Fmt2Zbr(ByVal v As Double) As String
    Fmt2Zbr = Replace(Format$(v, "0.00"), ",", ".")
End Function

' Header ciljne seme: BEZ UkupnoKolicina / UkupnoAmbalaze / Klasa. Te kolone jos
' postoje u tabeli (brisu se u PR4) i ostaju prazne namerno -- v. komentar iznad
' CreateZbirna_TX. GeneracijaID se ne pise.
Private Function BuildZbirnaHeaderRowData(ByVal zbirnaID As String, _
                                          ByVal datum As Date, _
                                          ByVal vozacID As String, _
                                          ByVal brojZbirne As String, _
                                          ByVal kupacID As String, _
                                          ByVal hladnjaca As String, _
                                          ByVal pogon As String, _
                                          ByVal vrstaVoca As String, _
                                          ByVal sortaVoca As String, _
                                          ByVal tipAmb As String) As Variant
    Const SRC As String = "BuildZbirnaHeaderRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_ZBIRNA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1247, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirna."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_HLADNJACA, hladnjaca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_POGON, pogon, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_TIP_AMB, tipAmb, SRC

    If GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_ZBIRNA, COL_STORNIRANO, "", SRC
    End If

    ' IZDATO se pise EKSPLICITNO.
    '
    ' Legacy konvencija je "prazno = IZDATO" (modDokumentInvariant.DocIsIssued,
    ' konzervativno). Nov model se na nju ne oslanja: ovaj writer pravi i odmah
    ' FINALIZUJE dokument, pa to i kaze. Kad se pojavi draft-first tok, on ce
    ' imati svoj ulaz (CreateZbirnaDraft_TX) i pisati DRAFT -- a razlika izmedju
    ' "niko nije upisao" i "izdato" tada vise nije pretpostavka.
    If GetColumnIndex(TBL_ZBIRNA, COL_TRACE_IZDATO_STATUS) > 0 Then
        SetRowValueByColumn rowData, TBL_ZBIRNA, COL_TRACE_IZDATO_STATUS, _
                            IZDATO_IZDATO, SRC
    End If

    BuildZbirnaHeaderRowData = rowData
End Function

' Clanstvo se pise ORIGINALNIM OtpremnicaID-em iz tabele, ne kljucem recnika:
' kljuc je UCase$ normalizovan da bi duplikat bio uhvatljiv, a u tabelu mora da
' ode ono sto tamo stvarno stoji.
Private Function IDIzRedaOtpremnice(ByRef data As Variant, ByVal colID As Long, _
                                    ByVal red As Long) As String
    IDIzRedaOtpremnice = Trim$(NzToText(data(red, colID)))
End Function

Private Function BuildZbirnaIzvorRowData(ByVal izvorID As String, _
                                         ByVal zbirnaID As String, _
                                         ByVal otpremnicaID As String) As Variant
    Const SRC As String = "BuildZbirnaIzvorRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_ZBIRNA_IZVORI)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1258, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirnaIzvori."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_ID, izvorID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, _
                        otpremnicaID, SRC

    BuildZbirnaIzvorRowData = rowData
End Function

Private Function BuildZbirnaStavkaRowData(ByVal stavkaID As String, _
                                          ByVal zbirnaID As String, _
                                          ByVal redniBroj As Long, _
                                          ByVal klasa As String, _
                                          ByVal kolicina As Double, _
                                          ByVal kolAmb As Double) As Variant
    Const SRC As String = "BuildZbirnaStavkaRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_ZBIRNA_STAVKE)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1248, SRC, _
                  "Ne mogu da odredim broj kolona za tblZbirnaStavke."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_ID, stavkaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_RB, redniBroj, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA_STAVKE, COL_ZBS_KOL_AMB, kolAmb, SRC

    BuildZbirnaStavkaRowData = rowData
End Function

' --- citanje DTO-a ----------------------------------------------------------
'
' Nedostajuci kljuc je GRESKA, ne prazna vrednost: Dictionary(k) nad nepostojecim
' kljucem tiho vraca Empty i doda kljuc, pa bi tipfeler u imenu polja prosao kao
' "korisnik nije uneo".

' Sve sto nije na spisku je greska.
'
' Provera obaveznih kljuceva hvata samo pola problema: tipfeler u OPCIONOM polju
' ("Hladnjca") ne obara nista, samo tiho ostavi prazno. Zato spisak.
'
' VrstaVoca / SortaVoca / TipAmbalaze NISU na spisku namerno -- izvode se iz
' izvornih otpremnica. Pozivalac koji ih salje pravi drugi izvor istine.
Private Function HdrKljucPoznat(ByVal kljuc As String) As Boolean
    Select Case LCase$(Trim$(kljuc))
        Case "datum", "vozacid", "brojzbirne", "kupacid", "hladnjaca", "pogon"
            HdrKljucPoznat = True
    End Select
End Function

Private Sub HdrProveriKljuceve(ByVal h As Object, ByVal src As String)
    Dim kljuc As Variant

    For Each kljuc In h.Keys
        If Not HdrKljucPoznat(CStr(kljuc)) Then
            Err.Raise vbObjectError + 1249, src, _
                      "Header ima nepoznat kljuc: " & CStr(kljuc) & _
                      ". Dozvoljeni: Datum, VozacID, BrojZbirne, KupacID, " & _
                      "Hladnjaca, Pogon. Vrsta/sorta/tip ambalaze dolaze iz otpremnica."
        End If
    Next kljuc
End Sub

Private Function HdrObavezan(ByVal h As Object, ByVal kljuc As String, _
                             ByVal src As String) As String
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1250, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    HdrObavezan = Trim$(NzToText(h(kljuc)))

    If Len(HdrObavezan) = 0 Then
        Err.Raise vbObjectError + 1251, src, _
                  "Header polje je prazno: " & kljuc
    End If
End Function

Private Function HdrOpcion(ByVal h As Object, ByVal kljuc As String) As String
    If h.Exists(kljuc) Then HdrOpcion = Trim$(NzToText(h(kljuc)))
End Function

Private Function HdrDatum(ByVal h As Object, ByVal kljuc As String, _
                          ByVal src As String) As Date
    If Not h.Exists(kljuc) Then
        Err.Raise vbObjectError + 1252, src, _
                  "Header nema obavezan kljuc: " & kljuc
    End If

    If Not IsDate(h(kljuc)) Then
        Err.Raise vbObjectError + 1253, src, _
                  "Header polje nije datum: " & kljuc
    End If

    HdrDatum = CDate(h(kljuc))
End Function

Private Function StavkaVrednost(ByVal s As Object, ByVal kljuc As String, _
                                ByVal idx As Long, ByVal src As String) As Variant
    If Not s.Exists(kljuc) Then
        Err.Raise vbObjectError + 1254, src, _
                  "Stavka " & CStr(idx) & " nema kljuc: " & kljuc
    End If

    StavkaVrednost = s(kljuc)
End Function

Private Function StavkaBroj(ByVal s As Object, ByVal kljuc As String, _
                            ByVal idx As Long, ByVal src As String) As Double
    Dim v As Variant
    v = StavkaVrednost(s, kljuc, idx, src)

    If Not IsNumeric(v) Then
        Err.Raise vbObjectError + 1255, src, _
                  "Stavka " & CStr(idx) & ", polje " & kljuc & _
                  " nije broj: " & NzToText(v)
    End If

    StavkaBroj = CDbl(v)
End Function

Public Function GetZbirnaByKupac(ByVal kupacID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_ZBIRNA)

    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, _
            "modDokumenta.GetZbirnaByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, _
                "modDokumenta.GetZbirnaByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetZbirnaByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetZbirnaByKupac"
    GetZbirnaByKupac = Empty
End Function

' ============================================================
' ZBIRNA VALIDIERUNG
' ============================================================
Public Function ValidateZbirna(ByVal brojZbirne As String) As Variant
    ' Prueft Summe Otpremnice vs Zbirna
    ' Returns: Array(SumaOtpKg, ZbirnaKg, RazlikaKg, ValidKg,
    '                SumaOtpAmb, ZbirnaAmb, RazlikaAmb)
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaOtpKg As Double
    Dim sumaOtpAmb As Long

    If Not IsEmpty(otpData) Then
        Dim colKol As Long
        Dim colAmb As Long

        colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                    "modDokumenta.ValidateZbirna")
        colAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                                    "modDokumenta.ValidateZbirna")

        Dim i As Long
        For i = 1 To UBound(otpData, 1)
            If IsNumeric(otpData(i, colKol)) Then sumaOtpKg = sumaOtpKg + CDbl(otpData(i, colKol))
            If IsNumeric(otpData(i, colAmb)) Then sumaOtpAmb = sumaOtpAmb + CLng(otpData(i, colAmb))
        Next i
    End If

    Dim zbirnaKg As Double
    Dim zbirnaAmb As Long

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If Not IsEmpty(zbrData) Then
        Dim colZbrBroj As Long
        Dim colZbrKol As Long
        Dim colZbrAmb As Long

        colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                        "modDokumenta.ValidateZbirna")
        colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                       "modDokumenta.ValidateZbirna")
        colZbrAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                                       "modDokumenta.ValidateZbirna")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colZbrBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colZbrKol))
                If IsNumeric(zbrData(i, colZbrAmb)) Then zbirnaAmb = zbirnaAmb + CLng(zbrData(i, colZbrAmb))
            End If
        Next i
    End If

    ValidateZbirna = Array(sumaOtpKg, zbirnaKg, sumaOtpKg - zbirnaKg, _
                           (Abs(sumaOtpKg - zbirnaKg) < 0.01), _
                           sumaOtpAmb, zbirnaAmb, sumaOtpAmb - zbirnaAmb)

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirna"
    ValidateZbirna = Array(0#, 0#, 0#, False, 0&, 0&, 0&)
End Function

Public Function ValidateZbirnaPreUnosa(ByVal brojZbirne As String, _
                                      ByVal inputKgKlI As Double, _
                                      ByVal inputKgKlII As Double, _
                                      ByVal inputAmb As Long) As Variant
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaKgKlI As Double
    Dim sumaKgKlII As Double
    Dim sumaAmb As Long

    If IsArray(otpData) Then
        Dim colKol As Long
        Dim colAmb As Long
        Dim colKlasa As Long
        Dim i As Long

        colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                    "modDokumenta.ValidateZbirnaPreUnosa")
        colAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                                    "modDokumenta.ValidateZbirnaPreUnosa")
        colKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, _
                                      "modDokumenta.ValidateZbirnaPreUnosa")

        Dim rowKlasa As String

        For i = 1 To UBound(otpData, 1)
            rowKlasa = Trim$(CStr(otpData(i, colKlasa)))

            If rowKlasa = KLASA_I Then
                If IsNumeric(otpData(i, colKol)) Then sumaKgKlI = sumaKgKlI + CDbl(otpData(i, colKol))
            ElseIf rowKlasa = KLASA_II Then
                If IsNumeric(otpData(i, colKol)) Then sumaKgKlII = sumaKgKlII + CDbl(otpData(i, colKol))
            End If

            If IsNumeric(otpData(i, colAmb)) Then sumaAmb = sumaAmb + CLng(otpData(i, colAmb))
        Next i
    End If

    ValidateZbirnaPreUnosa = Array( _
        sumaKgKlI, inputKgKlI, sumaKgKlI - inputKgKlI, (Abs(sumaKgKlI - inputKgKlI) < 0.01), _
        sumaKgKlII, inputKgKlII, sumaKgKlII - inputKgKlII, (Abs(sumaKgKlII - inputKgKlII) < 0.01), _
        sumaAmb, inputAmb, sumaAmb - inputAmb _
    )

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirnaPreUnosa"
    ValidateZbirnaPreUnosa = Array(0#, inputKgKlI, -inputKgKlI, False, _
                                   0#, inputKgKlII, -inputKgKlII, False, _
                                   0&, inputAmb, -inputAmb)
End Function

' Da li izvorne (nestornirane) otpremnice date zbirne imaju Klasu II.
' Koristi frmDokumenta (hard-blokada unosa zbirne kad je "Dve klase" iskljuceno --
' inace bi SaveZbirnaMulti_TX dobio hasKlasaII:=False i Kl.II bi se tiho izgubila).
' Isti izvor kao validacija u formi: ValidateZbirnaPreUnosa -> val(4) = sumaKgKlII.
Public Function ZbirnaIzvorImaKlasuII(ByVal brojZbirne As String) As Boolean
    On Error GoTo EH

    If Trim$(brojZbirne) = "" Then Exit Function

    Dim v As Variant
    v = ValidateZbirnaPreUnosa(brojZbirne, 0, 0, 0)

    If IsArray(v) Then
        If UBound(v) >= 4 Then
            If IsNumeric(v(4)) Then ZbirnaIzvorImaKlasuII = (CDbl(v(4)) > 0)
        End If
    End If

    Exit Function

EH:
    LogErr "modDokumenta.ZbirnaIzvorImaKlasuII"
    ZbirnaIzvorImaKlasuII = False
End Function

' ============================================================
' Generacija dokumenta (COL_GENERACIJA_ID)
'
' PRAVILO (jedno mesto, svi writer-i): red nasledjuje generaciju od AKTIVNIH
' (nestorniranih) redova ISTOG broja dokumenta; ako takvih nema -> nova generacija.
'
' Time bez ijednog dodatnog parametra vazi:
'   - Klasa I i Klasa II istog upisa dele generaciju (druga nasledjuje od prve),
'     bez obzira da li ih pise Multi_TX, AutoChainHladnjaca (dva zasebna _TX
'     poziva), malina auto-zbirna (red po red) ili backfill prijemnice;
'   - ISPRAVKA posle storna dobija NOVU generaciju -- svi stariji redovi tog broja
'     su tada stornirani, pa se nema sta naslediti;
'   - dopuna aktivnog dokumenta ostaje u njegovoj generaciji.
'
' Kolona je obavezna (RequireColumnIndex): stvara je modSetup.EnsureSledljivostSchema
' na svakom startu. Bez nje upis pada umesto da tiho nastane red bez generacije.
' ============================================================

' Generacija za red koji se upisuje pod datim brojem dokumenta.
' scopePairs: parovi (kolona, vrednost) koji uz broj cine IDENTITET dokumenta --
' broj sam po sebi NIJE globalno jedinstven (npr. GenerateBrojPrijemnice racuna
' sekvencu po kupcu, pa dva kupca istog dana oba dobiju "1/050826").
Public Function GeneracijaIDZaBroj(ByVal tableName As String, _
                                   ByVal brojCol As String, _
                                   ByVal broj As String, _
                                   ParamArray scopePairs() As Variant) As String
    GeneracijaIDZaBroj = GeneracijaIDZaBrojArr(tableName, brojCol, broj, _
                                               ScopePairsToArray(scopePairs, "modDokumenta.GeneracijaIDZaBroj"))
End Function

' Radna varijanta (scope kao obican niz) -- ParamArray se ne prosledjuje dalje.
Private Function GeneracijaIDZaBrojArr(ByVal tableName As String, _
                                       ByVal brojCol As String, _
                                       ByVal broj As String, _
                                       ByVal sc As Variant) As String
    Const SRC As String = "modDokumenta.GeneracijaIDZaBroj"

    Dim cGen As Long
    cGen = RequireColumnIndex(tableName, COL_GENERACIJA_ID, SRC)

    Dim target As String
    target = Trim$(broj)

    If Len(target) > 0 Then
        Dim data As Variant
        data = GetTableData(tableName)

        If IsArray(data) Then data = ExcludeStornirano(data, tableName)

        If IsArray(data) Then
            Dim cBr As Long
            cBr = RequireColumnIndex(tableName, brojCol, SRC)

            Dim scCols() As Long
            Dim i As Long
            If IsArray(sc) Then
                ReDim scCols(0 To UBound(sc))
                For i = 0 To UBound(sc) Step 2
                    scCols(i) = RequireColumnIndex(tableName, CStr(sc(i)), SRC)
                Next i
            End If

            Dim best As Double, bestGen As String
            Dim r As Long, g As String, rank As Double
            Dim inScope As Boolean
            best = -1

            For r = 1 To UBound(data, 1)
                If Trim$(NzToText(data(r, cBr))) = target Then
                    inScope = True
                    If IsArray(sc) Then
                        For i = 0 To UBound(sc) Step 2
                            If Trim$(NzToText(data(r, scCols(i)))) <> Trim$(NzToText(sc(i + 1))) Then
                                inScope = False
                                Exit For
                            End If
                        Next i
                    End If

                    If inScope Then
                        g = Trim$(NzToText(data(r, cGen)))
                        If Len(g) > 0 Then
                            rank = IdRank(g)
                            If rank > best Then
                                best = rank
                                bestGen = g
                            End If
                        End If
                    End If
                End If
            Next r

            If Len(bestGen) > 0 Then
                GeneracijaIDZaBrojArr = bestGen
                Exit Function
            End If
        End If
    End If

    GeneracijaIDZaBrojArr = NewGeneracijaID(tableName)
End Function

' ParamArray -> obican niz (parni: kolona, neparni: vrednost). Prazan -> Empty.
Private Function ScopePairsToArray(ByVal scopePairs As Variant, _
                                   ByVal sourceName As String) As Variant
    If Not IsArray(scopePairs) Then Exit Function
    If UBound(scopePairs) < LBound(scopePairs) Then Exit Function

    Dim n As Long
    n = UBound(scopePairs) - LBound(scopePairs) + 1

    If (n Mod 2) <> 0 Then
        Err.Raise vbObjectError + 1017, sourceName, _
                  "scopePairs mora imati parove (kolona, vrednost)."
    End If

    Dim out() As Variant
    ReDim out(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        out(i) = scopePairs(LBound(scopePairs) + i)
    Next i

    ScopePairsToArray = out
End Function

' Nova generacija za dokument-tabelu ("GEN-00042").
Public Function NewGeneracijaID(ByVal tableName As String) As String
    Const SRC As String = "modDokumenta.NewGeneracijaID"

    RequireColumnIndex tableName, COL_GENERACIJA_ID, SRC

    NewGeneracijaID = GetNextID(tableName, COL_GENERACIJA_ID, "GEN-")

    If Len(Trim$(NewGeneracijaID)) = 0 Then
        Err.Raise vbObjectError + 1015, SRC, _
                  "GetNextID nije vratio GeneracijaID za " & tableName & "."
    End If
End Function

' Izracunaj i upisi generaciju na upravo dodat red. Koriste je i writer-i koji
' ne idu kroz Save* (PWA import, invariant rekalkulacija).
Public Sub ApplyGeneracijaID(ByVal tableName As String, ByVal rowIndex As Long, _
                             ByVal brojCol As String, ByVal broj As String, _
                             ParamArray vlasnikPairs() As Variant)
    Const SRC As String = "modDokumenta.ApplyGeneracijaID"

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 1016, SRC, _
                  "Neispravan red za upis generacije (" & tableName & ")."
    End If

    ' Identitet dokumenta = broj + vlasnik: broj sam nije jedinstven. Vlasnik je
    ' otpremnica -> StanicaID, prijemnica -> KupacID, zbirna -> VozacID + KupacID
    ' (sekvenca se broji po vozacu, a dokument pripada kupcu; sam BROJ je
    ' jedinstven -- SuggestNextBroj za ZBR vrti petlju dok ne nadje slobodan).
    RequireUpdateCell tableName, rowIndex, COL_GENERACIJA_ID, _
                      GeneracijaIDZaBrojArr(tableName, brojCol, broj, _
                                            ScopePairsToArray(vlasnikPairs, SRC)), SRC
End Sub

' Pecati NOVU generaciju na upravo dodat red -- bez gledanja na to sta jos stoji
' pod istim brojem.
'
' ApplyGeneracijaID resava DRUGI problem: tamo dva fizicka reda cine JEDAN
' dokument (Kl.I + Kl.II), pa drugi red mora da nasledi generaciju prvog. Nasledje
' je tacno samo za pisca koji jedan dokument deli na vise redova.
'
' Pisac koji svaki red pise kao ZASEBAN dokument mora ovo. Za njega je nasledje
' aktivno stetno: dva odvojena dokumenta bi dobila isti GeneracijaID, pa bi
'   - ZbirnaIdentResolve izbrojao activeLogicalCount = 1 (broji GENERACIJE),
'     dakle UNIQUE -- F4 pusta, B8 cuti, kolizija se ne prijavljuje;
'   - StornoZbirna, koji redove bira po generaciji (RedJeIzabranogDokumenta),
'     stornirao OBA dokumenta na jedan storno;
'   - modScrStorno vise ne bi mogao ni da ih razlikuje: skrivena kolona
'     identiteta u gridu je bas COL_GENERACIJA_ID (modScrDokumenti.IdKolonaTipa).
'
' Prvi korisnik je modMasterSync: PWA import pise JEDAN red po ClientRecordID-u
' (obe klase sabrane u "I/II"), a IsDuplicateZbirnaInMaster odbija ponovljen CRID
' pre upisa -- svaki red koji stigne do AppendRow je dokument koji nikad nije
' vidjen.
' ZBR-CHILD-01: generacija roditeljske zbirne za dati BROJ, ili prazno.
'
' FAIL-CLOSED: prazno se vraca za sve sto nije jednoznacno -- nema aktivne zbirne,
' dvosmislen broj, pokvaren identitet. Dete tada ostaje bez generacije, sto je
' legitimno stanje i znaci "citaj po broju, kao i do sada". Pogadjati identitet
' iz broja je tacno ono protiv cega cela ZBR-IDENT celina i postoji.
'
' Zove se JEDNOM PO BROJU, ne po redu: ZbirnaIdentResolve cita celu tblZbirna, pa
' bi poziv u petlji nad decom bio O(n*m). Petlje zato uzimaju gen jednom i salju
' ga u PoveziDeteNaZbirnu.
Public Function ZbirnaGeneracijaZaBroj(ByVal broj As String) As String
    Dim id As ZbirnaIdent
    On Error GoTo EH
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    id = ZbirnaIdentResolve(broj)
    If id.integrityStatus <> ZBR_INT_OK Then Exit Function
    If id.resolutionStatus <> ZBR_RES_UNIQUE Then Exit Function
    ZbirnaGeneracijaZaBroj = id.selectedGeneracijaID
    Exit Function
EH:
    LogErr "modDokumenta.ZbirnaGeneracijaZaBroj", "broj=" & broj
End Function

' ZBR-CHILD-01: identitet roditelja kad je broj IKAD imao samo JEDNU generaciju.
'
' Za BACKFILL, ne za pisce. Razlika je sustinska:
'
'   ZbirnaGeneracijaZaBroj  pita "ko je roditelj SADA" -- tacno za red koji se
'                           upravo vezuje, jer se vezuje za tekuci dokument.
'   ova funkcija            pita "ko je IKAD bio pod ovim brojem" -- tacno za
'                           stari red kome se identitet naknadno rekonstruise.
'
' Zasto backfill ne sme "sada": ugovor par.5 IZRICITO dozvoljava re-entry istog
' vlasnika posle storna, pa je ovo legitimno stanje:
'
'   GEN-A | ZB-10 | vlasnik X | STORNIRANO
'   GEN-B | ZB-10 | vlasnik X | AKTIVNO      <- resolver kaze UNIQUE = GEN-B
'   OTP-A | BrojZbirne = ZB-10 | generacija prazna
'
' OTP-A je istorijski dete GEN-A. "Sada" bi mu upisalo GEN-B i napravilo LAZNU
' SLEDLJIVOST -- gore od prazne kolone, jer prazna bar ne tvrdi nista.
'
' Stornirana JEDINA generacija se sme upisati: ako je pod tim brojem ikad
' postojala samo jedna, identitet je poznat bez obzira na danasnje stanje.
' ZBR-CHILD-01: da li generacija ZAISTA pripada tom poslovnom broju.
'
' `RedJeIzabranogDokumenta` (modStorno) kad dobije generaciju bira red ISKLJUCIVO
' po njoj -- broj se vise ne gleda. Zato `StornoZbirna("X", "GEN-C")` stornira
' GEN-C i kad GEN-C pripada broju Y. `IdoviGeneracije` isto poredi samo
' generaciju, pa ni ona nije branila par.
'
' Rupa je STARIJA od faze 4: i kad broj X nosi jedan dokument, nespojiv par bi
' prosao. Faza 4 je samo uklonila slucajnu barijeru -- kapiju koja je taj poziv
' zaustavljala kad je X dvosmislen -- i time je ucinila dohvatljivom u vise
' slucajeva. A njena premisa ("akter zna identitet") bez ove provere ne stoji:
' neprazan GeneracijaID nije dokaz da akter zna dokument POD TIM BROJEM.
'
' Gleda i STORNIRANE redove namerno: `CompleteZbirnaIspravka` legitimno radi sa
' identitetom stare, vec stornirane zbirne.
Public Function ZbirnaGeneracijaPripadaBroju(ByVal broj As String, _
                                             ByVal gen As String) As Boolean
    On Error GoTo EH
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    If Len(Trim$(NzToText(gen))) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim cBroj As Long: cBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    Dim cGen As Long: cGen = GetColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID)
    If cBroj = 0 Or cGen = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, cGen))), Trim$(NzToText(gen)), vbTextCompare) = 0 Then
            If BrojJednak(data(i, cBroj), broj) Then
                ZbirnaGeneracijaPripadaBroju = True
                Exit Function
            End If
        End If
    Next i
    Exit Function
EH:
    LogErr "modDokumenta.ZbirnaGeneracijaPripadaBroju", "broj=" & broj & " gen=" & gen
End Function

' `outRazlog` je Optional ByRef: zatecenim pozivaocima se nista ne menja, a
' migracija dobija RAZLOG. Odluka ostaje na JEDNOM mestu -- da backfill sam
' racuna razlog, imao bi drugu kopiju pravila, sto je tacno ono sto faza 2
' nije htela.
Public Function ZbirnaJedinaGeneracijaIkadZaBroj(ByVal broj As String, _
                                                 Optional ByRef outRazlog As String) As String
    Dim id As ZbirnaIdent
    On Error GoTo EH
    outRazlog = ""
    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function
    id = ZbirnaIdentResolve(broj)
    If id.integrityStatus <> ZBR_INT_OK Then
        outRazlog = ZBR_BF_INTEGRITET
        Exit Function
    End If
    ' 0 i >1 oba padaju na <> 1, ali traze RAZLICIT potez: nula znaci da nijedna
    ' zbirna pod tim brojem nema GeneracijaID (stari red pre uvodjenja kolone),
    ' vise od jedne znaci stvarnu dvosmislenost. Prvo se resava migracijom
    ' roditelja, drugo se ne resava uopste.
    If id.historicalLogicalCount = 0 Then
        outRazlog = ZBR_BF_NEMA_GENERACIJE
        Exit Function
    End If
    If id.historicalLogicalCount > 1 Then
        outRazlog = ZBR_BF_VISE_GENERACIJA
        Exit Function
    End If
    ZbirnaJedinaGeneracijaIkadZaBroj = id.historicalOnlyGeneracijaID
    Exit Function
EH:
    outRazlog = ZBR_BF_INTEGRITET
    LogErr "modDokumenta.ZbirnaJedinaGeneracijaIkadZaBroj", "broj=" & broj
End Function

' ZBR-CHILD-01: JEDINI put kojim dete dobija zbirnu u DVA upisa.
'
' Broj i generacija se upisuju ZAJEDNO. Dva odvojena upisa bi se pre ili kasnije
' razisla: neko doda putanju koja postavlja broj a zaboravi generaciju, i dete
' ostane sa TUDJOM generacijom -- gore od prazne, jer prazna bar znaci "ne znam".
'
' Zovu ga i SaveOtpremnica/SavePrijemnica posle AppendRow, iako je broj vec u
' rowData: prepis iste vrednosti je jeftin, a druga putanja bi znacila da se par
' moze raziciti. Jedini upis koji NE ide ovuda je PalAppendRow u modPaletniList,
' i tamo rizika nema -- oba polja idu u ISTOM append pozivu, pa ih nema sta da
' razdvoji.
'
' gen se prosledjuje, ne racuna ovde: pozivaoci su cesto petlje nad decom istog
' broja (v. ZbirnaGeneracijaZaBroj).
Public Sub PoveziDeteNaZbirnu(ByVal tableName As String, ByVal rowIndex As Long, _
                              ByVal brojCol As String, ByVal brojZbirne As String, _
                              ByVal gen As String, ByVal sourceName As String)
    RequireUpdateCell tableName, rowIndex, brojCol, brojZbirne, sourceName
    RequireUpdateCell tableName, rowIndex, COL_DETE_ZBIRNA_GEN, gen, sourceName
End Sub

' ZBR-CHILD-01: dete se odvezuje od zbirne -- oba polja, u istom potezu.
'
' Bez ovoga bi odvezano dete zadrzalo generaciju stornirane zbirne i izgledalo kao
' da joj i dalje pripada, dok mu je broj prazan.
Public Sub OdveziDeteOdZbirne(ByVal tableName As String, ByVal rowIndex As Long, _
                              ByVal brojCol As String, ByVal sourceName As String)
    PoveziDeteNaZbirnu tableName, rowIndex, brojCol, "", "", sourceName
End Sub

' ZBR-CHILD-01 faza 3: da li se CELA operacija sme suzavati.
'
' `SuziDecuNaGeneraciju` odlucuje po JEDNOM skupu, a poslovna mutacija dira vise
' tabela. Kad svaka odlucuje sama, jedna kaskada zna da bude pola scoped a pola
' po broju: otpremnice suzene na GEN-B, a prijemnice -- jer je jedna legacy --
' vracene na broj, pa se stornira i prijemnica GEN-A. Sve-ili-nista mora da vazi
' nad CELOM operacijom, ne nad tabelom.
'
' Odluka se zato racuna JEDNOM, nad svim tabelama koje ta operacija bira po broju,
' pa se svim selektorima prosledi ista: generacija (suzavaj) ili prazno (ne suzavaj).
'
' Poredi kroz BrojJednak namerno, iako cetiri pozivaoca porede tacno. BrojJednak
' je siri, pa je ovde skup kandidata NADSKUP stvarnog: ako svi u nadskupu nose
' generaciju, nosi je i svaki podskup. Greska ide samo u stranu "ne suzavaj",
' nikad u "suzi pogresno".
Public Function SvaAktivnaDecaNoseGeneraciju(ByVal tableName As String, _
                                             ByVal brojCol As String, _
                                             ByVal broj As String) As Boolean
    SvaAktivnaDecaNoseGeneraciju = True

    If Len(Trim$(NzToText(broj))) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim cBroj As Long: cBroj = GetColumnIndex(tableName, brojCol)
    If cBroj = 0 Then Exit Function
    Dim cSt As Long: cSt = GetColumnIndex(tableName, COL_STORNIRANO)
    Dim cGen As Long: cGen = GetColumnIndex(tableName, COL_DETE_ZBIRNA_GEN)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If BrojJednak(data(i, cBroj), broj) Then
            If cSt = 0 Or UCase$(Trim$(NzToText(data(i, cSt)))) <> "DA" Then
                ' Kandidat postoji, a tabela nema kolonu -> ne moze se scope-ovati.
                If cGen = 0 Then
                    SvaAktivnaDecaNoseGeneraciju = False
                    Exit Function
                End If
                If Len(Trim$(NzToText(data(i, cGen)))) = 0 Then
                    SvaAktivnaDecaNoseGeneraciju = False
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ZBR-CHILD-01 faza 3: JEDAN put kojim se skup dece suzava na JEDAN dokument.
'
' Pandan write choke point-u iznad. Namerno NE preuzima i izbor po broju: cetiri
' zatecena odlucivaca porede broj TACNO (Trim$(CStr(..)) = broj), a peti
' (DistinctActiveValues) kanonski preko BrojJednak. Da je ovaj put preuzeo i to
' poredjenje, ona cetiri bi se tiho PROSIRILA -- akter siri od zatecenog je bas
' smer koji ZBR-NORM-02 zove opasnim. Zato svaki pozivalac zadrzava svoje
' poredjenje, a ovde se centralizuje samo ono sto je faza 3: generacija.
'
' PRAVILO (sve-ili-nista, ne hibrid):
'   1) trazena generacija prazna       -> vrati kandidate nepromenjeno
'   2) bilo koji kandidat bez generacije -> vrati kandidate nepromenjeno
'   3) inace                            -> vrati samo one koje se poklapaju
'
' Zasto NE hibrid "poklapa se ILI je prazno": pod jednim brojem mogu stajati dva
' dokumenta, pa bi isti prazan red upao u skup OBA -- dupli detach, pogresan
' racun. Fallback je bit-identican zatecenom ponasanju, sto je i uslov da faza 3
' ne pomeri nijednu zatecenu brojku.
'
' `data` ide ByRef i mora biti BAS onaj snimak nad kojim pozivalac vrti petlju:
' drugo citanje unutar iste transakcije moglo bi da vidi drugo stanje, pa bi
' indeksi redova pokazivali na tudje redove. ByRef i zbog KOPIJA_NIZA.
Public Function SuziDecuNaGeneraciju(ByVal tableName As String, ByRef data As Variant, _
                                     ByVal kandidati As Collection, _
                                     ByVal gen As String) As Collection
    Set SuziDecuNaGeneraciju = kandidati

    If kandidati Is Nothing Then Exit Function
    If kandidati.count = 0 Then Exit Function
    If Len(Trim$(NzToText(gen))) = 0 Then Exit Function
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim cGen As Long: cGen = GetColumnIndex(tableName, COL_DETE_ZBIRNA_GEN)
    If cGen = 0 Then Exit Function

    Dim k As Long
    For k = 1 To kandidati.count
        If Len(Trim$(NzToText(data(CLng(kandidati(k)), cGen)))) = 0 Then Exit Function
    Next k

    Dim suzeno As New Collection
    For k = 1 To kandidati.count
        If StrComp(Trim$(NzToText(data(CLng(kandidati(k)), cGen))), _
                   Trim$(NzToText(gen)), vbTextCompare) = 0 Then
            suzeno.Add CLng(kandidati(k))
        End If
    Next k

    Set SuziDecuNaGeneraciju = suzeno
End Function

Public Sub ApplyNovaGeneracijaID(ByVal tableName As String, ByVal rowIndex As Long)
    Const SRC As String = "modDokumenta.ApplyNovaGeneracijaID"

    If rowIndex <= 0 Then
        Err.Raise vbObjectError + 1018, SRC, _
                  "Neispravan red za upis generacije (" & tableName & ")."
    End If

    RequireUpdateCell tableName, rowIndex, COL_GENERACIJA_ID, _
                      NewGeneracijaID(tableName), SRC
End Sub

' Bira redove dokumenta za prefill ispravke (frmDokumenta.Prefill*FromStornirana).
'
' Polazi od ANCHOR reda -- konkretnog PK-a stornirane (OldDocID iz correction
' context-a). Broj dokumenta se NE koristi kao identitet: nije globalno jedinstven
' (GenerateBrojPrijemnice racuna sekvencu po kupcu, pa dva kupca istog dana dobiju
' isti broj), pa bi pretraga po broju mogla da prefiluje tudji dokument.
'
' Kad anchor PK nije poznat (stariji context bez OldDocID), fallback je poslednje
' upisan red datog broja -- i tada se ostaje unutar generacije tog reda, pa se
' redovi dva vlasnika ne mesaju.
'
' Iz anchor-a se cita generacija (COL_GENERACIJA_ID) i uzimaju Kl.I i Kl.II SAMO
' iz nje. Bez generacije (red stariji od uvodjenja kolone) prefiluje se samo sam
' anchor -- konzervativno, jer pripadnost druge klase nije dokaziva.
' ============================================================
' IDENTITET DOKUMENTA NA GRANICI PREVEZIVANJA
'
' Broj dokumenta je LABELA, ne identitet: BrojPrijemnice se racuna po kupcu
' (GenerateBrojPrijemnice), broj zbirne po vozacu. Dva dokumenta lako dele broj.
'
' Identitet logickog dokumenta (Klasa I + II zajedno) je GeneracijaID -- kolonu
' pravi modSetup.EnsureSledljivostSchema na SVAKOM startu, a pecate je svi
' writer-i kroz ApplyGeneracijaID, sa vec tacnim kompozitnim vlasnistvom po
' tipu (otpremnica StanicaID, prijemnica KupacID, zbirna VozacID + KupacID).
'
' Rutine prevezivanja su do sada primale samo broj i skenirale tabelu po njemu.
' Ove dve funkcije im daju identitet.
' ============================================================

' Da li red tabele pripada izvornom dokumentu. Sa poznatom generacijom odlucuje
' iskljucivo PK; bez nje broj (pozivalac je pre toga dokazao jednoznacnost).
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Public Function PripadaIzvoru(ByRef data As Variant, ByVal rowIdx As Long, _
                              ByVal cBroj As Long, ByVal cId As Long, _
                              ByVal broj As String, ByVal srcIds As Object) As Boolean
    If Not srcIds Is Nothing Then
        If srcIds.count > 0 Then
            If cId > 0 Then
                PripadaIzvoru = srcIds.Exists(Trim$(NzToText(data(rowIdx, cId))))
                Exit Function
            End If
        End If
    End If
    If cBroj > 0 Then PripadaIzvoru = (Trim$(NzToText(data(rowIdx, cBroj))) = Trim$(broj))
End Function

' GeneracijaID dokumenta kome pripada dati red (po PK-u). Prazno = red ne
' postoji ili nema generaciju (stari zapis, pre uvodjenja kolone).
Public Function GeneracijaPoID(ByVal tableName As String, ByVal idCol As String, _
                               ByVal docID As String) As String
    On Error Resume Next
    docID = Trim$(docID)
    If Len(docID) = 0 Then Exit Function
    GeneracijaPoID = Trim$(NzToText(LookupValue(tableName, idCol, docID, COL_GENERACIJA_ID)))
End Function

' Svi PK-evi koji pripadaju datoj generaciji. To je "koji su redovi OVAJ
' dokument" - i Klasa I i Klasa II, i nijedan tudji.
'
' Prazna generacija vraca prazan recnik, NE sve redove: pozivalac tada mora da
' padne na kapiju nad brojem, ne na tiho zahvatanje svega.
Public Function IdoviGeneracije(ByVal tableName As String, ByVal idCol As String, _
                                ByVal gen As String) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set IdoviGeneracije = res

    On Error Resume Next
    gen = Trim$(gen)
    If Len(gen) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cGen As Long
    cId = GetColumnIndex(tableName, idCol)
    cGen = GetColumnIndex(tableName, COL_GENERACIJA_ID)
    If cId = 0 Or cGen = 0 Then Exit Function

    Dim i As Long, k As String
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cGen))) = gen Then
            k = Trim$(NzToText(data(i, cId)))
            If Len(k) > 0 Then res(k) = True
        End If
    Next i
End Function

Public Sub PickPrefillRows(ByVal data As Variant, _
                           ByVal cBroj As Long, ByVal cKlasa As Long, _
                           ByVal cId As Long, ByVal cGen As Long, _
                           ByVal broj As String, ByVal oldDocID As String, _
                           ByRef outRowI As Long, ByRef outRowII As Long)
    outRowI = 0
    outRowII = 0

    If Not IsArray(data) Then Exit Sub

    Dim anchor As Long
    anchor = FindAnchorRow(data, cBroj, cId, broj, oldDocID)
    If anchor = 0 Then Exit Sub

    Dim genTop As String
    If cGen > 0 Then genTop = Trim$(NzToText(data(anchor, cGen)))

    If Len(genTop) = 0 Then
        If RowKlasaII(data, anchor, cKlasa) Then
            outRowII = anchor
        Else
            outRowI = anchor
        End If
        Exit Sub
    End If

    Dim bestI As Double, bestII As Double
    Dim r As Long, rank As Double
    bestI = -1
    bestII = -1

    For r = 1 To UBound(data, 1)
        If Trim$(NzToText(data(r, cGen))) = genTop Then
            rank = RowRank(data, r, cId)

            If RowKlasaII(data, r, cKlasa) Then
                If rank > bestII Then
                    bestII = rank
                    outRowII = r
                End If
            Else
                If rank > bestI Then
                    bestI = rank
                    outRowI = r
                End If
            End If
        End If
    Next r
End Sub

' Anchor: red sa datim PK-om; ako PK nije poznat -> poslednje upisan red datog broja.
Private Function FindAnchorRow(ByVal data As Variant, ByVal cBroj As Long, _
                               ByVal cId As Long, ByVal broj As String, _
                               ByVal oldDocID As String) As Long
    Dim r As Long

    If Len(Trim$(oldDocID)) > 0 And cId > 0 Then
        For r = 1 To UBound(data, 1)
            If Trim$(NzToText(data(r, cId))) = Trim$(oldDocID) Then
                FindAnchorRow = r
                Exit Function
            End If
        Next r
        ' PK je ZADAT ali ga u tabeli NEMA -> vrati prazno, bez fallback-a.
        '
        ' Fallback po broju postoji samo za STARE kontekste koji OldDocID
        ' uopste nemaju. Kad kontekst tvrdi konkretan dokument a njega nema,
        ' "uzmi poslednji red istog broja" znaci: prefiluj TUDJI dokument.
        ' Broj nije jedinstven (GenerateBrojPrijemnice broji po kupcu), pa je
        ' to realan scenario, a ne teorijski.
        Exit Function
    End If

    If cBroj <= 0 Then Exit Function

    Dim target As String
    target = Trim$(broj)
    If Len(target) = 0 Then Exit Function

    Dim best As Double, rank As Double
    best = -1

    For r = 1 To UBound(data, 1)
        If Trim$(NzToText(data(r, cBroj))) = target Then
            rank = RowRank(data, r, cId)
            If rank >= best Then
                best = rank
                FindAnchorRow = r
            End If
        End If
    Next r
End Function

' Rang reda: numericki sufiks ID-a (GetNextID je monoton), inace indeks reda.
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowRank(ByRef data As Variant, ByVal rowIndex As Long, _
                         ByVal cId As Long) As Double
    RowRank = -1

    If cId > 0 Then RowRank = IdRank(NzToText(data(rowIndex, cId)))
    If RowRank < 0 Then RowRank = CDbl(rowIndex)
End Function

' Klasa reda je II (prazna/nepoznata klasa se tretira kao I, kao u ostatku koda).
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowKlasaII(ByRef data As Variant, ByVal rowIndex As Long, _
                            ByVal cKlasa As Long) As Boolean
    If cKlasa <= 0 Then Exit Function
    RowKlasaII = (UCase$(Trim$(NzToText(data(rowIndex, cKlasa)))) = "II")
End Function

' Numericki sufiks ID-a ("OTP-00042" -> 42); -1 ako ga nema. GetNextID je monoton
' po tabeli, pa je veci sufiks = kasnije upisan red.
Private Function IdRank(ByVal idText As String) As Double
    Dim s As String
    s = Trim$(idText)

    Dim i As Long, ch As String, digits As String
    For i = Len(s) To 1 Step -1
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = ch & digits
        Else
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then
        IdRank = -1
    Else
        IdRank = Val(digits)
    End If
End Function

' ============================================================
' PRIJEMNICA - Kunde wiegt bei Annahme
' ============================================================

Public Function SavePrijemnicaMulti_TX(ByVal datum As Date, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal brojPrij As String, _
                                       ByVal brojZbirne As String, _
                                       ByVal vrstaVoca As String, _
                                       ByVal sortaVoca As String, _
                                       ByVal kolicinaI As Double, _
                                       ByVal cenaI As Double, _
                                       ByVal tipAmb As String, _
                                       ByVal kolAmb As Long, _
                                       ByVal kolAmbVracena As Long, _
                                       Optional ByVal hasKlasaII As Boolean = False, _
                                       Optional ByVal kolicinaII As Double = 0, _
                                       Optional ByVal cenaII As Double = 0, _
                                       Optional ByVal kolAmbII As Long = 0, _
                                       Optional ByVal brutoKgI As Double = 0, _
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
    modSchema.SchemaReadyOrFail "SavePrijemnicaMulti_TX", _
        TBL_PRIJEMNICA & "|" & TBL_AMBALAZA & "|" & TBL_FAKTURA_STAVKE & _
        "|" & TBL_FAKTURE & "|" & TBL_PALETA & "|" & TBL_PALETA_STAVKA

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' Klasa I je opciona (kolicinaI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1303, "SavePrijemnicaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
        resultI = SavePrijemnica( _
            datum, _
            kupacID, _
            vozacID, _
            brojPrij, _
            brojZbirne, _
            vrstaVoca, _
            sortaVoca, _
            kolicinaI, _
            cenaI, _
            tipAmb, _
            kolAmb, _
            kolAmbVracena, _
            KLASA_I, _
            brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1301, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa I fehlgeschlagen"
        End If
    End If

    Dim resultII As String
    If hasKlasaII Then
        resultII = SavePrijemnica( _
            datum, _
            kupacID, _
            vozacID, _
            brojPrij, _
            brojZbirne, _
            vrstaVoca, _
            sortaVoca, _
            kolicinaII, _
            cenaII, _
            tipAmb, _
            kolAmbII, _
            0, _
            KLASA_II, _
            brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1302, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SavePrijemnicaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SavePrijemnicaMulti_TX = resultI
    Else
        SavePrijemnicaMulti_TX = resultII
    End If

    ' Paletizacija Klase I UNUTAR transakcije -> atomicno sa prijemnicom.
    Dim closedPal As Collection
    Set closedPal = New Collection
    If hasKlasaI And kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=resultI, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_I, netoKg:=kolicinaI, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    ' Paletizacija Klase II (zasebne gajbe) -> u istu kolekciju zatvorenih paleta.
    If hasKlasaII And kolAmbII > 0 Then
        PaletizePrijemnica prijemnicaID:=resultII, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_II, netoKg:=kolicinaII, brGajbica:=kolAmbII, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

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
    LogError "SavePrijemnicaMulti_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnicaMulti_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; HasKlasaII=" & CStr(hasKlasaII) & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnicaMulti_TX = ""

    PrintTxFailure "SavePrijemnicaMulti_TX", errSrc, errNum, errDesc
End Function

Public Function SavePrijemnica_TX(ByVal datum As Date, ByVal kupacID As String, _
                                   ByVal vozacID As String, ByVal brojPrij As String, _
                                   ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                                   Optional ByVal klasa As String = "I", _
                                   Optional ByVal brutoKg As Double = 0) As String
    Dim tx As New clsTransaction

    On Error GoTo EH

        ' Sema pre upisa: AppendRow pise POZICIONO (v. SaveOtkupMulti_TX).
    modSchema.SchemaReadyOrFail "SavePrijemnica_TX", _
        TBL_PRIJEMNICA & "|" & TBL_AMBALAZA

tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    SavePrijemnica_TX = SavePrijemnica(datum, kupacID, vozacID, brojPrij, _
                                        brojZbirne, vrstaVoca, sortaVoca, _
                                        kolicina, cena, tipAmb, kolAmb, _
                                        kolAmbVracena, klasa, brutoKg)

    If SavePrijemnica_TX = "" Then
        Err.Raise vbObjectError + 1011, "SavePrijemnica_TX", _
                  "SavePrijemnica fehlgeschlagen"
    End If

    ' Paletizacija UNUTAR transakcije -> atomicno sa prijemnicom (gajbice bilo koje klase).
    Dim closedPal As Collection
    Set closedPal = New Collection
    If kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=SavePrijemnica_TX, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=klasa, netoKg:=kolicina, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "SavePrijemnica_TX", errDesc, errNum
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnica_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnica_TX = ""

    PrintTxFailure "SavePrijemnica_TX", errSrc, errNum, errDesc
End Function
    
Public Function SavePrijemnica(ByVal datum As Date, ByVal kupacID As String, _
                               ByVal vozacID As String, ByVal brojPrij As String, _
                               ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                               ByVal sortaVoca As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                               Optional ByVal klasa As String = "I", _
                               Optional ByVal brutoKg As Double = 0) As String
    On Error GoTo EH

    Call ValidatePrijemnicaInput(kupacID, vozacID, brojPrij, brojZbirne, _
                             kolicina, cena, tipAmb, kolAmb, _
                             kolAmbVracena, klasa)

    Dim newID As String
    newID = GetNextID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-")

    If newID = "" Then
        Err.Raise vbObjectError + 1013, "SavePrijemnica", _
                  "GetNextID nije vratio PrijemnicaID."
    End If

    Dim rowData As Variant
    rowData = BuildPrijemnicaRowData( _
                newID, datum, kupacID, vozacID, brojPrij, brojZbirne, _
                vrstaVoca, sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                kolAmbVracena, klasa)


    Dim appendedRow As Long
    appendedRow = AppendRow(TBL_PRIJEMNICA, rowData)

    If appendedRow <= 0 Then
        Err.Raise vbObjectError + 1014, "SavePrijemnica", _
                "AppendRow fehlgeschlagen fuer tblPrijemnica."
    End If

    ' ZBR-CHILD-01: generacija roditeljske zbirne (v. isti komentar u
    ' SaveOtpremnica). Prijemnica roditelja obicno IMA, pa je ovde retko prazna.
    PoveziDeteNaZbirnu TBL_PRIJEMNICA, appendedRow, COL_PRJ_BROJ_ZBIRNE, brojZbirne, _
                       ZbirnaGeneracijaZaBroj(brojZbirne), "modDokumenta.SavePrijemnica"
    ApplyGeneracijaID TBL_PRIJEMNICA, appendedRow, COL_PRJ_BROJ, brojPrij, _
                      COL_PRJ_KUPAC, kupacID

    ' Bruto tezina (preneto iz otkupa kad je OTKUP_BRUTO_UNOS) -> upis po imenu;
    ' prazno = neto. Kolona postoji posle EnsureDoradeSchema (na kraju tblPrijemnica).
    If brutoKg > 0 Then UpdateCell TBL_PRIJEMNICA, appendedRow, COL_PRJ_BRUTO, brutoKg

    ' Ambalaza je ENTITETSKI-relativna (smer iz ugla hladnjace / Kupca):
    ' 1. txt = pune gajbe koje hladnjaca PRIMA od zbirne -> Kupac ULAZ.
    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    ' 2. txt = zamena: prazne gajbe koje hladnjaca VRACA (daje vozacu) -> Kupac IZLAZ.
    If kolAmbVracena > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmbVracena, "Izlaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    RelinkFakturaStavke newID, brojPrij, klasa

    SavePrijemnica = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogError "SavePrijemnica", errDesc, errNum
    On Error GoTo 0

    Err.Raise errNum, "SavePrijemnica", _
              "Source=" & errSrc & " | " & errDesc
End Function

' Gradi red tblZbirna PO IMENU kolone (AUD-003). Otporno na promenu redosleda
' kolona i na kolonu umetnutu u sredinu -- za razliku od pozicijskog Array(...).
Private Function BuildZbirnaRowData(ByVal zbirnaID As String, _
                                    ByVal datum As Date, _
                                    ByVal vozacID As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kupacID As String, _
                                    ByVal hladnjaca As String, _
                                    ByVal pogon As String, _
                                    ByVal vrstaVoca As String, _
                                    ByVal sortaVoca As String, _
                                    ByVal ukupnoKol As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal ukupnoAmb As Long, _
                                    ByVal klasa As String) As Variant
    Const SRC As String = "BuildZbirnaRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_ZBIRNA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1011, SRC, _
                  "Could not resolve tblZbirna column count."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_HLADNJACA, hladnjaca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_POGON, pogon, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KOLICINA, ukupnoKol, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KOL_AMB, ukupnoAmb, SRC
    SetRowValueByColumn rowData, TBL_ZBIRNA, COL_ZBR_KLASA, klasa, SRC

    If GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_ZBIRNA, COL_STORNIRANO, "", SRC
    End If

    BuildZbirnaRowData = rowData
End Function

Private Function BuildPrijemnicaRowData(ByVal prijemnicaID As String, _
                                        ByVal datum As Date, _
                                        ByVal kupacID As String, _
                                        ByVal vozacID As String, _
                                        ByVal brojPrij As String, _
                                        ByVal brojZbirne As String, _
                                        ByVal vrstaVoca As String, _
                                        ByVal sortaVoca As String, _
                                        ByVal kolicina As Double, _
                                        ByVal cena As Double, _
                                        ByVal tipAmb As String, _
                                        ByVal kolAmb As Long, _
                                        ByVal kolAmbVracena As Long, _
                                        ByVal klasa As String) As Variant
    Const SRC As String = "BuildPrijemnicaRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_PRIJEMNICA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1430, SRC, _
                  "Could not resolve tblPrijemnica column count."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ, brojPrij, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_CENA, cena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, kolAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA, kolAmbVracena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "Ne", SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "", SRC

    If GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_STORNIRANO, "", SRC
    End If

    BuildPrijemnicaRowData = rowData
End Function

Public Function GetPrijemniceByKupac(ByVal kupacID As String, _
                                      Optional ByVal datumOd As Date = 0, _
                                      Optional ByVal datumDo As Date = 0, _
                                      Optional ByVal samoNefakturisano As Boolean = False) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
            "modDokumenta.GetPrijemniceByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, _
                "modDokumenta.GetPrijemniceByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    If samoNefakturisano Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                "modDokumenta.GetPrijemniceByKupac"), "<>", "Da"
        filters.Add fp
    End If

    GetPrijemniceByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetPrijemniceByKupac"
    GetPrijemniceByKupac = Empty
End Function

Public Function SaveKupciIzlaz_TX(ByVal datum As Date, _
                                  ByVal brojDok As String, _
                                  ByVal kupacNaziv As String, _
                                  ByVal kupacID As String, _
                                  ByVal vozacID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolAmb As Long, _
                                  ByVal vrstaVoca As String, _
                                  ByVal novac As Double, _
                                  ByVal fakturaID As String, _
                                  ByVal napomena As String, _
                                  ByVal tipNovca As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If kupacID = "" Then
        Err.Raise vbObjectError + 1601, "SaveKupciIzlaz_TX", _
                  "KupacID je obavezan."
    End If

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1602, "SaveKupciIzlaz_TX", _
                  Poruka("DOK_ERR_NEMA_AMBALAZE_NOVCA")
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, _
                      "Izlaz", kupacID, "Kupac", _
                      vozacID, brojDok, DOK_TIP_IZLAZ_KUPCI
    End If

    If novac > 0 Then
        Dim novacID As String

        ' Ista kapija kao u SaveOMUlaz_TX, samo nad fakturom: vlasnistvo, storno
        ' stanje i TRENUTNO preostalo se citaju ovde, ne uzimaju iz parametara.
        Dim fakErr As String
        fakErr = UplataFakturaProblem(fakturaID, kupacID, novac)
        If Len(fakErr) > 0 Then
            Err.Raise vbObjectError + 1604, "SaveKupciIzlaz_TX", fakErr
        End If

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=kupacNaziv, _
            partnerId:=kupacID, _
            entitetTip:="Kupac", _
            omID:="", _
            kooperantID:="", _
            fakturaID:=fakturaID, _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=novac, _
            isplata:=0, _
            napomena:=napomena, _
            otkupID:="")

        If novacID = "" Then
            Err.Raise vbObjectError + 1603, "SaveKupciIzlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

        If fakturaID <> "" Then
            UpdateFakturaStatus fakturaID
        End If
    End If

    tx.CommitTx
    Set tx = Nothing

    SaveKupciIzlaz_TX = True
    Exit Function

EH:
    LogErr "SaveKupciIzlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveKupciIzlaz_TX = False
End Function

' ============================================================
' MANJAK - Schwundberechnung
' ============================================================

Public Function CalculateManjak(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim zbirnaKg As Double
    Dim zbrData As Variant

    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then
        zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

        Dim colBroj As Long
        Dim colZbrKol As Long
        Dim j As Long

        colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                     "modDokumenta.CalculateManjak")
        colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                       "modDokumenta.CalculateManjak")

        For j = 1 To UBound(zbrData, 1)
            If CStr(zbrData(j, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(j, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(j, colZbrKol))
            End If
        Next j
    End If

    Dim prijKg As Double
    Dim prijData As Variant

    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)

    If Not IsEmpty(prijData) Then
        Dim colBrZbr As Long
        Dim colKol As Long
        Dim i As Long

        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                      "modDokumenta.CalculateManjak")
        colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                    "modDokumenta.CalculateManjak")

        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colKol)) Then prijKg = prijKg + CDbl(prijData(i, colKol))
            End If
        Next i
    End If

    Dim manjakKg As Double
    Dim manjakPct As Double

    manjakKg = zbirnaKg - prijKg

    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjak = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjak"
    CalculateManjak = Array(0#, 0#, 0#, 0#)
End Function

Public Function CalculateManjakPreview(ByVal brojZbirne As String, _
                                      ByVal pendingKgKlI As Double, _
                                      ByVal pendingKgKlII As Double) As Variant
    On Error GoTo EH

    Dim zbirnaKg As Double

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If IsArray(zbrData) Then
        Dim colBroj As Long
        Dim colKol As Long
        Dim i As Long

        colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                     "modDokumenta.CalculateManjakPreview")
        colKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                    "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colKol))
            End If
        Next i
    End If

    Dim prijKg As Double
    Dim prijData As Variant

    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)

    If IsArray(prijData) Then
        Dim colBrZbr As Long
        Dim colPrijKol As Long

        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                      "modDokumenta.CalculateManjakPreview")
        colPrijKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                        "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colPrijKol)) Then prijKg = prijKg + CDbl(prijData(i, colPrijKol))
            End If
        Next i
    End If

    prijKg = prijKg + pendingKgKlI + pendingKgKlII

    Dim manjakKg As Double
    Dim manjakPct As Double

    manjakKg = zbirnaKg - prijKg

    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjakPreview = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjakPreview"
    CalculateManjakPreview = Array(0#, pendingKgKlI + pendingKgKlII, 0#, 0#)
End Function

Public Function CalculateManjakByOtpremnica(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim manjak As Variant
    manjak = CalculateManjak(brojZbirne)

    Dim zbirnaKg As Double
    Dim manjakKg As Double
    Dim manjakPct As Double

    zbirnaKg = CDbl(manjak(0))
    manjakKg = CDbl(manjak(2))
    manjakPct = CDbl(manjak(3))

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsEmpty(otpData) Then
        CalculateManjakByOtpremnica = Empty
        Exit Function
    End If

    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        CalculateManjakByOtpremnica = Empty
        Exit Function
    End If

    Dim colBroj As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colStan As Long

    colBroj = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                 "modDokumenta.CalculateManjakByOtpremnica")
    colStan = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, _
                                 "modDokumenta.CalculateManjakByOtpremnica")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                "modDokumenta.CalculateManjakByOtpremnica")
    colCena = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, _
                                 "modDokumenta.CalculateManjakByOtpremnica")

    Dim rowCount As Long
    rowCount = UBound(otpData, 1)

    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)

    Dim i As Long
    For i = 1 To rowCount
        Dim kol As Double
        Dim udeo As Double
        Dim cena As Double

        kol = 0
        udeo = 0
        cena = 0

        If IsNumeric(otpData(i, colKol)) Then kol = CDbl(otpData(i, colKol))
        If zbirnaKg > 0 Then udeo = kol / zbirnaKg
        If IsNumeric(otpData(i, colCena)) Then cena = CDbl(otpData(i, colCena))

        result(i, 1) = CStr(otpData(i, colBroj))
        result(i, 2) = CStr(otpData(i, colStan))
        result(i, 3) = kol
        result(i, 4) = udeo
        result(i, 5) = udeo * manjakKg
        result(i, 6) = manjakPct
        result(i, 7) = udeo * manjakKg * cena
    Next i

    CalculateManjakByOtpremnica = result
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjakByOtpremnica"
    CalculateManjakByOtpremnica = Empty
End Function

' ============================================================
' PROSEK GAJBE - Durchschnittsgewicht pro Kaestchen
' ============================================================

Public Function CalculateProsekGajbe(ByVal brojOtp As String) As Double
    On Error GoTo EH

    If Trim$(brojOtp) = "" Then Exit Function

    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_BROJ, _
                       "modDokumenta.CalculateProsekGajbe"
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                       "modDokumenta.CalculateProsekGajbe"
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                       "modDokumenta.CalculateProsekGajbe"

    ' Dvoklasni doc: sumiraj kol i amb preko SVIH redova broja (ne samo prvi red).
    Dim kol As Double, amb As Double
    kol = SumByBroj(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOLICINA)
    amb = SumByBroj(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOL_AMB)

    If amb > 0 Then
        CalculateProsekGajbe = kol / amb
    Else
        CalculateProsekGajbe = 0
    End If

    Exit Function

EH:
    LogErr "modDokumenta.CalculateProsekGajbe"
    CalculateProsekGajbe = 0
End Function

Public Function CalculateProsekGajbeByZbirna(ByVal brojZbirne As String) As Double
    On Error GoTo EH

    If Trim$(brojZbirne) = "" Then Exit Function

    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_BROJ, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"

    ' Dvoklasni doc: sumiraj kol i amb preko SVIH redova broja (ne samo prvi red).
    Dim kol As Double, amb As Double
    kol = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOLICINA)
    amb = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOL_AMB)

    If amb > 0 Then
        CalculateProsekGajbeByZbirna = kol / amb
    Else
        CalculateProsekGajbeByZbirna = 0
    End If

    Exit Function

EH:
    LogErr "modDokumenta.CalculateProsekGajbeByZbirna"
    CalculateProsekGajbeByZbirna = 0
End Function

' Sumira valCol preko SVIH redova gde brojCol = broj (obe klase dvorednog doc-a).
Private Function SumByBroj(ByVal tbl As String, ByVal brojCol As String, _
                           ByVal broj As String, ByVal valCol As String) As Double
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    ' Stornirani redovi ne ulaze u sumu (prosek gajbe je racunao i njih).
    If IsArray(data) Then data = ExcludeStornirano(data, tbl)
    If Not IsArray(data) Then Exit Function
    Dim cB As Long, cV As Long
    cB = GetColumnIndex(tbl, brojCol)
    cV = GetColumnIndex(tbl, valCol)
    If cB = 0 Or cV = 0 Then Exit Function
    Dim i As Long, s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cB))) = Trim$(broj) Then
            If IsNumeric(data(i, cV)) Then s = s + CDbl(data(i, cV))
        End If
    Next i
    SumByBroj = s
End Function

' ============================================================
' Storno Verweiste
' ============================================================

Public Function GetVerwaisteDokumente(ByVal dokumentTip As String) As Variant
    ' Returns: 2D Array der Dokumente deren BrojZbirne auf eine
    '          stornierte Zbirna zeigt, die selbst aber NICHT storniert sind.
    '
    ' dokumentTip: "Otpremnica" oder "Prijemnica"
    '
    ' Otpremnica Returns: (OtpremnicaID, BrojOtp, BrojZbirne, VrstaVoca, Kolicina)
    ' Prijemnica Returns: (PrijemnicaID, BrojPrij, BrojZbirne, KupacNaziv, Kolicina)
    ' oder Empty
    On Error GoTo EH

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsEmpty(zbrData) Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Dim colZbrBroj As Long
    Dim colZbrStorno As Long

    colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                    "modDokumenta.GetVerwaisteDokumente")
    colZbrStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, _
                                      "modDokumenta.GetVerwaisteDokumente")

    Dim storniraneBrojevi As Object
    Set storniraneBrojevi = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim brz As String

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) = "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If brz <> "" Then
                If Not storniraneBrojevi.Exists(brz) Then
                    storniraneBrojevi.Add brz, True
                End If
            End If
        End If
    Next i

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) <> "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If storniraneBrojevi.Exists(brz) Then
                storniraneBrojevi.Remove brz
            End If
        End If
    Next i

    If storniraneBrojevi.count = 0 Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Select Case dokumentTip
        Case "Otpremnica"
            GetVerwaisteDokumente = GetVerwaisteOtpremnice(storniraneBrojevi)

        Case "Prijemnica"
            GetVerwaisteDokumente = GetVerwaistePrijemnice(storniraneBrojevi)

        Case Else
            GetVerwaisteDokumente = Empty
    End Select

    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteDokumente"
    GetVerwaisteDokumente = Empty
End Function

Private Function GetVerwaisteOtpremnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrOtp As Long
    Dim colBrZbr As Long
    Dim colVrsta As Long
    Dim colKol As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                               "modDokumenta.GetVerwaisteOtpremnice")
    colBrOtp = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                "modDokumenta.GetVerwaisteOtpremnice")
    colStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaisteOtpremnice")

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            kol = 0
            If IsNumeric(data(i, colKol)) Then kol = CDbl(data(i, colKol))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrOtp))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = NzToText(data(i, colVrsta))
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaisteOtpremnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteOtpremnice"
    GetVerwaisteOtpremnice = Empty
End Function

Private Function GetVerwaistePrijemnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrPrij As Long
    Dim colBrZbr As Long
    Dim colKupac As Long
    Dim colKol As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, _
                               "modDokumenta.GetVerwaistePrijemnice")
    colBrPrij = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                                   "modDokumenta.GetVerwaistePrijemnice")
    colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                "modDokumenta.GetVerwaistePrijemnice")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaistePrijemnice")

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kupacNaziv As String
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, _
                                          NzToText(data(i, colKupac)), COL_KUP_NAZIV))

            kol = 0
            If IsNumeric(data(i, colKol)) Then kol = CDbl(data(i, colKol))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrPrij))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = kupacNaziv
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaistePrijemnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaistePrijemnice"
    GetVerwaistePrijemnice = Empty
End Function

Public Sub RelinkFakturaStavke(ByVal newPrijemnicaID As String, _
                               ByVal brojPrijemnice As String, _
                               Optional ByVal klasaFilter As String = "")
    ' Sucht verwaiste FakturaStavke die auf eine stornierte Prijemnica
    ' mit gleichem BrojPrijemnice zeigen, und verlinkt sie auf die neue.
    On Error GoTo EH

    If Trim$(newPrijemnicaID) = "" Or Trim$(brojPrijemnice) = "" Then Exit Sub

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colPrijID As Long
    Dim colOsir As Long
    Dim colFakID As Long
    Dim colFsKlasa As Long

    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, _
                                   "modDokumenta.RelinkFakturaStavke")
    colOsir = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, _
                                 "modDokumenta.RelinkFakturaStavke")
    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, _
                                  "modDokumenta.RelinkFakturaStavke")
    colFsKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, _
                                "modDokumenta.RelinkFakturaStavke")

    Dim i As Long
    Dim oldPrijID As String
    Dim oldBroj As String
    Dim fakID As String

    For i = 1 To UBound(stavkeData, 1)
    
        Dim stavkaKlasa As String
        stavkaKlasa = Trim$(NzToText(stavkeData(i, colFsKlasa)))
        If Len(Trim$(klasaFilter)) > 0 Then
            If UCase$(stavkaKlasa) <> UCase$(Trim$(klasaFilter)) Then GoTo NextStavka
        End If

        If Trim$(NzToText(stavkeData(i, colOsir))) = "" Then GoTo NextStavka

        oldPrijID = Trim$(NzToText(stavkeData(i, colPrijID)))

        If oldPrijID = "" Then GoTo NextStavka

        oldBroj = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, oldPrijID, COL_PRJ_BROJ))

        If oldBroj = brojPrijemnice Then
            fakID = Trim$(NzToText(stavkeData(i, colFakID)))
            
            If Len(Trim$(fakID)) = 0 Then
            Err.Raise vbObjectError + 7413, "modDokumenta.RelinkFakturaStavke", _
                    "FakturaStavka nema FakturaID za relink. BrojPrijemnice=" & _
                    brojPrijemnice & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_FS_PRIJEMNICA_ID, _
                              newPrijemnicaID, "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, _
                              "", "modDokumenta.RelinkFakturaStavke"

            Dim newPrijRow As Long
            newPrijRow = FindPrijemnicaRowByIDAndKlasa(newPrijemnicaID, stavkaKlasa, _
                                           "modDokumenta.RelinkFakturaStavke")

            If newPrijRow <= 0 Then
                Err.Raise vbObjectError + 7410, "modDokumenta.RelinkFakturaStavke", _
                        "Nova prijemnica nije pronadena za relink. PrijemnicaID=" & _
                        newPrijemnicaID & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURISANO, _
                            "Da", "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURA_ID, _
                            fakID, "modDokumenta.RelinkFakturaStavke"


            If Len(Trim$(fakID)) > 0 Then
                UpdateFakturaStatus fakID
            End If

        End If

NextStavka:
    Next i

    Exit Sub

EH:
    LogErr "modDokumenta.RelinkFakturaStavke"
    Err.Raise Err.Number, "modDokumenta.RelinkFakturaStavke", Err.description
End Sub

' ============================================================
' HELPER - Vozac-Report (ersetzt alten modTransport)
' ============================================================

Public Function GetVozacDokumenta(ByVal vozacID As String, _
                                   Optional ByVal datumOd As Date = 0, _
                                   Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If

    data = ExcludeStornirano(data, TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, _
            "modDokumenta.GetVozacDokumenta"), "=", vozacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetVozacDokumenta"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetVozacDokumenta = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetVozacDokumenta"
    GetVozacDokumenta = Empty
End Function

Public Function BuildZbirnaVrstaCache() As Object
    On Error GoTo EH

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If
    
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If

    Dim colBrZbr As Long
    Dim colVrsta As Long

    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.BuildZbirnaVrstaCache")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.BuildZbirnaVrstaCache")

    Dim i As Long
    Dim brz As String
    Dim vrsta As String

    For i = 1 To UBound(otpData, 1)
        brz = Trim$(NzToText(otpData(i, colBrZbr)))
        vrsta = Trim$(NzToText(otpData(i, colVrsta)))

        If brz <> "" Then
            If Not dict.Exists(brz) Then
                dict.Add brz, vrsta
            End If
        End If
    Next i

    Set BuildZbirnaVrstaCache = dict
    Exit Function

EH:
    LogErr "modDokumenta.BuildZbirnaVrstaCache"

    Dim emptyDict As Object
    Set emptyDict = CreateObject("Scripting.Dictionary")
    Set BuildZbirnaVrstaCache = emptyDict
End Function

Public Function GetVrstaFromCache(ByVal dict As Object, _
                                  ByVal brojZbirne As String) As String
    On Error GoTo EH

    If dict Is Nothing Then
        GetVrstaFromCache = ""
    ElseIf dict.Exists(brojZbirne) Then
        GetVrstaFromCache = CStr(dict(brojZbirne))
    Else
        GetVrstaFromCache = ""
    End If

    Exit Function

EH:
    LogErr "modDokumenta.GetVrstaFromCache"
    GetVrstaFromCache = ""
End Function

Private Sub PrintTxFailure(ByVal sourceName As String, _
                           ByVal errSrc As String, _
                           ByVal errNum As Long, _
                           ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidDocumentClass(ByVal klasa As String, _
                                      ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1400, sourceName, _
              "Neispravna klasa dokumenta: " & klasa
End Sub

Private Sub ValidateOtpremnicaInput(ByVal stanicaID As String, _
                                    ByVal vozacID As String, _
                                    ByVal brojOtp As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal kolAmb As Long, _
                                    ByVal klasa As String)

    Const SRC As String = "ValidateOtpremnicaInput"

    If Len(Trim$(stanicaID)) = 0 Then
        Err.Raise vbObjectError + 1401, SRC, "StanicaID je obavezan."
    End If

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1402, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojOtp)) = 0 Then
        Err.Raise vbObjectError + 1403, SRC, "Broj otpremnice je obavezan."
    End If

    'If Len(Trim$(brojZbirne)) = 0 Then
    '    Err.Raise vbObjectError + 1404, SRC, "Broj zbirne je obavezan."
    'End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1405, SRC, "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena < 0 Then
        Err.Raise vbObjectError + 1406, SRC, "Cena ne sme biti negativna."
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1407, SRC, "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1408, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub ValidateZbirnaInput(ByVal vozacID As String, _
                                ByVal brojZbirne As String, _
                                ByVal kupacID As String, _
                                ByVal ukupnoKol As Double, _
                                ByVal tipAmb As String, _
                                ByVal ukupnoAmb As Long, _
                                ByVal klasa As String)

    Const SRC As String = "ValidateZbirnaInput"

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1410, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise vbObjectError + 1411, SRC, "Broj zbirne je obavezan."
    End If

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise vbObjectError + 1412, SRC, "KupacID je obavezan."
    End If

    If ukupnoKol <= 0 Then
        Err.Raise vbObjectError + 1413, SRC, "Ukupna koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If ukupnoAmb < 0 Then
        Err.Raise vbObjectError + 1414, SRC, "Ukupna ambala" & ChrW(382) & "a ne sme biti negativna."
    End If

    If ukupnoAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1415, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub ValidatePrijemnicaInput(ByVal kupacID As String, _
                                    ByVal vozacID As String, _
                                    ByVal brojPrij As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal kolAmb As Long, _
                                    ByVal kolAmbVracena As Long, _
                                    ByVal klasa As String)

    Const SRC As String = "ValidatePrijemnicaInput"

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise vbObjectError + 1420, SRC, "KupacID je obavezan."
    End If

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1421, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojPrij)) = 0 Then
        Err.Raise vbObjectError + 1422, SRC, "Broj prijemnice je obavezan."
    End If

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise vbObjectError + 1423, SRC, "Broj zbirne je obavezan."
    End If

    ' Referencijalni integritet: broj zbirne mora da postoji u tblZbirna.
    ' Samo u BLOK modu -- u UPOZORENJE modu forma trazi potvrdu, pa backend ne
    ' sme tvrdo da padne. Ujedno je bezbedna mreza i za ne-form pozivaoce.
    If PrijemnicaZbirnaBlokira() Then
        If Not ZbirnaPostoji(brojZbirne) Then
            Err.Raise vbObjectError + 1428, SRC, _
                "Zbirna '" & brojZbirne & "' ne postoji u sistemu."
        End If
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1424, SRC, "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena < 0 Then
        Err.Raise vbObjectError + 1425, SRC, "Cena ne sme biti negativna."
    End If

    If kolAmb < 0 Or kolAmbVracena < 0 Then
        Err.Raise vbObjectError + 1426, SRC, "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If (kolAmb > 0 Or kolAmbVracena > 0) And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1427, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub SetRowValueByColumn(ByRef rowData() As Variant, _
                                ByVal tableName As String, _
                                ByVal columnName As String, _
                                ByVal value As Variant, _
                                ByVal sourceName As String)
    Dim colIndex As Long
    colIndex = RequireColumnIndex(tableName, columnName, sourceName)

    rowData(colIndex - 1) = value
End Sub

Private Function GetDokumentaTableColumnCount(ByVal tableName As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                GetDokumentaTableColumnCount = lo.ListColumns.count
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindPrijemnicaRowByIDAndKlasa(ByVal prijemnicaID As String, _
                                               ByVal klasa As String, _
                                               ByVal sourceName As String) As Long
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colKlasa As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, sourceName)
    colKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, sourceName)

    Dim i As Long
    Dim foundRow As Long
    Dim foundCount As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colID))) = Trim$(prijemnicaID) And _
           Trim$(NzToText(data(i, colKlasa))) = Trim$(klasa) Then

            foundRow = i
            foundCount = foundCount + 1
        End If
    Next i

    If foundCount > 1 Then
        Err.Raise vbObjectError + 7412, sourceName, _
                  "Dupli redovi za PrijemnicaID + Klasa. PrijemnicaID=" & _
                  prijemnicaID & "; Klasa=" & klasa
    End If

    FindPrijemnicaRowByIDAndKlasa = foundRow
End Function

' ============================================================
' STORNO PREGLED (read-only) -- agregira stornirane dokumente po tipu za
' prikaz u panelu unutar frmDokumenta (dugme "Pregled storniranih").
' Soft-delete: red je storniran kad je COL_STORNIRANO = "Da" (modStorno).
' Jedinstven (unifikovan) skup korisnih kolona za sve tipove:
'   Broj | Datum | Partner | Vrsta | Sorta | Klasa | Kolicina | Cena | Iznos
'   | Zbirna | Otpremnica | Faktura  (poslednje 3 = lanac zavisnih dokumenata)
' Partner se razresava na naziv/ime (best-effort; fallback = ID).
' ============================================================

' Tipovi storno dokumenata u redosledu prikaza (isti nazivi kao cmbStornoDokument).
Public Function StorniraniTipovi() As Variant
    StorniraniTipovi = Array("Otkup", "Otpremnica", "Zbirna", "Prijemnica", "Faktura", "Novac", _
                         "Revers izdavanje koop.", "Revers povrat koop.", _
                         "Revers izdato OM (firma).", "Revers prijem od OM (firma).")
End Function

' Zaglavlja unifikovanih kolona (0-bazni niz, 12 kolona).
Public Function StorniraniHeaders() As Variant
    StorniraniHeaders = Array("Broj", "Datum", "Partner", "Vrsta", "Sorta", _
                              "Klasa", "Koli" & ChrW(269) & "ina", "Cena", "Iznos (RSD)", _
                              "Zbirna", "Otpremnica", "Faktura")
End Function

' Vrati 2D niz (1..n, 1..12) storniranih redova za dati tip u unifikovanom
' rasporedu kolona (osnovne + lanac: Zbirna/Otpremnica/Faktura). Empty ako nema.
' chainIdx: pre-izgradjen indeks lanca (BuildChainIndex); ako fali, gradi se sam.
Public Function GetStorniraniByTip(ByVal tip As String, _
                                   Optional ByVal chainIdx As Object = Nothing) As Variant
    On Error GoTo EH
    If chainIdx Is Nothing Then Set chainIdx = BuildChainIndex()

    ' OM<->koop revers (ambalaza) ima drugaciju strukturu (2 noge, EntitetID/Tip) ->
    ' poseban scan tblAmbalaze.
    Select Case tip
        Case "Revers izdavanje koop."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_IZLAZ_KOOP)
            Exit Function
        Case "Revers povrat koop."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_ULAZ_KOOP)
            Exit Function
        Case "Revers izdato OM (firma)."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_ULAZ_FIRMA, "Stanica")
            Exit Function
        Case "Revers prijem od OM (firma)."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_IZLAZ_FIRMA, "Stanica")
            Exit Function
    End Select

    Dim tbl As String
    Dim iBroj As Long, iDat As Long, iPart As Long, iVrsta As Long, iSorta As Long
    Dim iKlasa As Long, iKol As Long, iCena As Long, iIzn As Long, iIzn2 As Long
    Dim iZbr As Long, iFak As Long
    Dim partnerIsName As Boolean
    Dim pdict As Object

    Select Case tip
        Case "Otkup"
            tbl = TBL_OTKUP
            Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
            iBroj = GetColumnIndex(tbl, COL_OTK_BR_DOK)
            iDat = GetColumnIndex(tbl, COL_OTK_DATUM)
            iPart = GetColumnIndex(tbl, COL_OTK_KOOPERANT)
            iVrsta = GetColumnIndex(tbl, COL_OTK_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_OTK_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_OTK_KLASA)
            iKol = GetColumnIndex(tbl, COL_OTK_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_OTK_CENA)
            iIzn = GetColumnIndex(tbl, COL_OTK_NOVAC)
            iZbr = GetColumnIndex(tbl, COL_OTK_BROJ_ZBIRNE)

        Case "Otpremnica"
            tbl = TBL_OTPREMNICA
            Set pdict = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")
            iBroj = GetColumnIndex(tbl, COL_OTP_BROJ)
            iDat = GetColumnIndex(tbl, COL_OTP_DATUM)
            iPart = GetColumnIndex(tbl, COL_OTP_STANICA)
            iVrsta = GetColumnIndex(tbl, COL_OTP_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_OTP_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_OTP_KLASA)
            iKol = GetColumnIndex(tbl, COL_OTP_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_OTP_CENA)
            iZbr = GetColumnIndex(tbl, COL_OTP_BROJ_ZBIRNE)

        Case "Zbirna"
            tbl = TBL_ZBIRNA
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_ZBR_BROJ)
            iDat = GetColumnIndex(tbl, COL_ZBR_DATUM)
            iPart = GetColumnIndex(tbl, COL_ZBR_KUPAC)
            iVrsta = GetColumnIndex(tbl, COL_ZBR_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_ZBR_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_ZBR_KLASA)
            iKol = GetColumnIndex(tbl, COL_ZBR_KOLICINA)
            iZbr = GetColumnIndex(tbl, COL_ZBR_BROJ)        ' sopstveni broj = pivot zbirne

        Case "Prijemnica"
            tbl = TBL_PRIJEMNICA
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_PRJ_BROJ)
            iDat = GetColumnIndex(tbl, COL_PRJ_DATUM)
            iPart = GetColumnIndex(tbl, COL_PRJ_KUPAC)
            iVrsta = GetColumnIndex(tbl, COL_PRJ_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_PRJ_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_PRJ_KLASA)
            iKol = GetColumnIndex(tbl, COL_PRJ_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_PRJ_CENA)
            iZbr = GetColumnIndex(tbl, COL_PRJ_BROJ_ZBIRNE)
            iFak = GetColumnIndex(tbl, COL_PRJ_FAKTURA_ID)

        Case "Faktura"
            tbl = TBL_FAKTURE
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_FAK_BROJ)
            iDat = GetColumnIndex(tbl, COL_FAK_DATUM)
            iPart = GetColumnIndex(tbl, COL_FAK_KUPAC)
            iIzn = GetColumnIndex(tbl, COL_FAK_IZNOS)
            iFak = GetColumnIndex(tbl, COL_FAK_ID)          ' sopstveni FakturaID

        Case "Novac"
            tbl = TBL_NOVAC
            partnerIsName = True       ' Partner je vec naziv (ne ID)
            iBroj = GetColumnIndex(tbl, COL_NOV_BROJ_DOK)
            iDat = GetColumnIndex(tbl, COL_NOV_DATUM)
            iPart = GetColumnIndex(tbl, COL_NOV_PARTNER)
            iVrsta = GetColumnIndex(tbl, COL_NOV_VRSTA)
            iIzn = GetColumnIndex(tbl, COL_NOV_UPLATA)
            iIzn2 = GetColumnIndex(tbl, COL_NOV_ISPLATA)    ' Iznos = Uplata - Isplata
            iFak = GetColumnIndex(tbl, COL_NOV_FAKTURA_ID)

        Case Else
            Exit Function
    End Select

    Dim data As Variant
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim colStorno As Long
    colStorno = GetColumnIndex(tbl, COL_STORNIRANO)
    If colStorno = 0 Then Exit Function        ' tabela bez storno markera -> nista

    Dim otpByZbr As Object: Set otpByZbr = chainIdx("otpByZbr")
    Dim fakByZbr As Object: Set fakByZbr = chainIdx("fakByZbr")
    Dim fakById As Object:  Set fakById = chainIdx("fakById")
    Dim zbrByFak As Object: Set zbrByFak = chainIdx("zbrByFak")

    Dim rows As Collection
    Set rows = New Collection

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(NzToText(data(i, colStorno)))) = "DA" Then
            Dim part As String
            If partnerIsName Then
                part = StornoCellText(data, i, iPart)
            Else
                part = ResolveNameFromDict(pdict, StornoCellRaw(data, i, iPart))
            End If

            ' Lanac zavisnih dokumenata preko BrojZbirne / FakturaID.
            Dim zbr As String, fakId As String, otp As String, fak As String
            zbr = StornoCellText(data, i, iZbr)
            fakId = StornoCellText(data, i, iFak)
            If Len(zbr) = 0 And Len(fakId) > 0 Then zbr = DictGet(zbrByFak, fakId)
            otp = DictGet(otpByZbr, zbr)
            If Len(fakId) > 0 Then fak = DictGet(fakById, fakId) Else fak = DictGet(fakByZbr, zbr)

            ' Iznos: stored (Faktura=Iznos, Novac=Uplata-Isplata) ili izracunat
            ' Kolicina x Cena (Otkup/Otpremnica/Prijemnica nemaju zaseban iznos).
            Dim iznos As String
            iznos = StornoIznosText(StornoCellRaw(data, i, iIzn), StornoCellRaw(data, i, iIzn2))
            If Len(iznos) = 0 Then _
                iznos = StornoMnozi(StornoCellRaw(data, i, iKol), StornoCellRaw(data, i, iCena))

            rows.Add Array( _
                StornoCellText(data, i, iBroj), _
                StornoDateText(StornoCellRaw(data, i, iDat)), _
                part, _
                StornoCellText(data, i, iVrsta), _
                StornoCellText(data, i, iSorta), _
                StornoCellText(data, i, iKlasa), _
                StornoNumText(StornoCellRaw(data, i, iKol), "#,##0.00"), _
                StornoNumText(StornoCellRaw(data, i, iCena), "#,##0.00"), _
                iznos, _
                zbr, otp, fak)
        End If
    Next i

    GetStorniraniByTip = StornoRowsTo2D(rows, 12)
    Exit Function

EH:
    LogErr "modDokumenta.GetStorniraniByTip(" & tip & ")"
    GetStorniraniByTip = Empty
End Function

' Stornirani OM<->koop revers (dokTip) iz tblAmbalaza -> unifikovane 12 kolona.
' Jedan red po dokumentu: Kooperant noga (nosi partnera); Stanica noga se preskace.
' Mapiranje: Vrsta=TipAmbalaze, Kolicina=broj gajbica; ostalo prazno.
' Preskace OM-Izlaz-Koop noge knjizene UZ otkup (DokumentID = otkupID "OTK-...";
' vec se vide pod grupom 'Otkup') -> ostaju samo standalone reversi (broj x/ddmmyy).
Private Function GetStorniraniRevers(ByVal dokTip As String, _
                                     Optional ByVal entType As String = "Kooperant") As Variant
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long, iDat As Long, iEntID As Long, iEntTip As Long
    Dim iDokTip As Long, iTipAmb As Long, iKol As Long, iStorno As Long
    iBroj = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    iDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    iEntID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    iEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    iDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    iTipAmb = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    iKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    iStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If iBroj = 0 Or iDokTip = 0 Or iStorno = 0 Then Exit Function

    Dim pdict As Object
    If entType = "Stanica" Then
        Set pdict = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")
    Else
        Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    End If
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(NzToText(data(i, iStorno)))) = "DA" Then
            If Trim$(CStr(data(i, iDokTip))) = dokTip And _
               Trim$(CStr(data(i, iEntTip))) = entType And _
               UCase$(Left$(Trim$(CStr(data(i, iBroj))), 3)) <> "OTK" Then
                rows.Add Array( _
                    StornoCellText(data, i, iBroj), _
                    StornoDateText(StornoCellRaw(data, i, iDat)), _
                    ResolveNameFromDict(pdict, StornoCellRaw(data, i, iEntID)), _
                    StornoCellText(data, i, iTipAmb), _
                    "", "", _
                    StornoNumText(StornoCellRaw(data, i, iKol), "#,##0"), _
                    "", "", "", "", "")
            End If
        End If
    Next i
    GetStorniraniRevers = StornoRowsTo2D(rows, 12)
    Exit Function
EH:
    LogErr "modDokumenta.GetStorniraniRevers(" & dokTip & ")"
    GetStorniraniRevers = Empty
End Function

' Svi stornirani grupisano: Collection of Array(tipNaziv, rows2D(1..n,1..12)).
' Indeks lanca se gradi JEDNOM (reverzni cross-table lookup-i) i deli svim tipovima.
Public Function GetStorniraniGrupisano() As Collection
    On Error GoTo EH
    Dim res As Collection
    Set res = New Collection
    Set GetStorniraniGrupisano = res

    Dim idx As Object: Set idx = BuildChainIndex()
    Dim tipovi As Variant: tipovi = StorniraniTipovi()
    Dim k As Long
    For k = LBound(tipovi) To UBound(tipovi)
        Dim tip As String: tip = CStr(tipovi(k))
        Dim rowsv As Variant: rowsv = GetStorniraniByTip(tip, idx)
        If Not IsEmpty(rowsv) Then res.Add Array(tip, rowsv)
    Next k
    Exit Function
EH:
    LogErr "modDokumenta.GetStorniraniGrupisano"
End Function

' id -> "Naziv" (ili "Ime Prezime") recnik; prazan recnik ako tabela/kolone fale.
' Public: reuse i iz modIntegritet (OtkupnoMestoByZbirna).
Public Function BuildIdNameDict(ByVal tbl As String, ByVal idCol As String, _
                                ByVal nameCol1 As String, _
                                Optional ByVal nameCol2 As String = "") As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set BuildIdNameDict = d

    Dim data As Variant
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ci As Long, c1 As Long, c2 As Long
    ci = GetColumnIndex(tbl, idCol)
    c1 = GetColumnIndex(tbl, nameCol1)
    If Len(nameCol2) > 0 Then c2 = GetColumnIndex(tbl, nameCol2)
    If ci = 0 Or c1 = 0 Then Exit Function

    Dim i As Long, idv As String, nm As String
    For i = 1 To UBound(data, 1)
        idv = Trim$(NzToText(data(i, ci)))
        If Len(idv) > 0 Then
            nm = Trim$(NzToText(data(i, c1)))
            If c2 > 0 Then nm = Trim$(nm & " " & NzToText(data(i, c2)))
            If Not d.Exists(idv) Then d.Add idv, nm
        End If
    Next i
End Function

Private Function ResolveNameFromDict(ByVal d As Object, ByVal idRaw As Variant) As String
    Dim s As String
    s = Trim$(NzToText(idRaw))
    If Len(s) = 0 Then Exit Function
    If Not d Is Nothing Then
        If d.Exists(s) Then
            If Len(Trim$(NzToText(d(s)))) > 0 Then
                ResolveNameFromDict = d(s)
                Exit Function
            End If
        End If
    End If
    ResolveNameFromDict = s          ' fallback: prikazi ID ako nema imena
End Function

' Sirova vrednost celije (Empty ako kolona ne postoji -> idx=0).
' `data` je ByRef namerno. ByVal na Variantu koji SADRZI niz kopira ceo niz pri
' svakom pozivu, a ovo je citac PO CELIJI -- u petlji se zove vise puta po redu.
' Mereno na istom obrascu u modPaletniList.SafeCell: 1063 stavke, 1883 ms, to jest
' 1.8 ms po redu za citanje dva polja iz niza koji je vec u memoriji.
'
' Funkcija niz samo CITA, nikad ne pise, pa je razlika iskljucivo u tome sto se
' niz ne umnozava.
Private Function StornoCellRaw(ByRef data As Variant, ByVal r As Long, ByVal idx As Long) As Variant
    If idx = 0 Then Exit Function
    StornoCellRaw = data(r, idx)
End Function

' ByRef iz istog razloga kao StornoCellRaw: ByVal bi kopirao ceo niz po pozivu.
Private Function StornoCellText(ByRef data As Variant, ByVal r As Long, ByVal idx As Long) As String
    If idx = 0 Then Exit Function
    StornoCellText = Trim$(NzToText(data(r, idx)))
End Function

Private Function StornoDateText(ByVal v As Variant) As String
    If IsDate(v) Then
        StornoDateText = Format$(CDate(v), "d.m.yyyy")
    Else
        StornoDateText = Trim$(NzToText(v))
    End If
End Function

Private Function StornoNumText(ByVal v As Variant, ByVal fmt As String) As String
    Dim d As Double
    If TryParseDouble(Trim$(NzToText(v)), d) Then
        If d <> 0 Then StornoNumText = Format$(d, fmt)
    End If
End Function

' Iznos = v1 - v2 (Novac: Uplata - Isplata; ostali: v2 prazno -> samo v1).
Private Function StornoIznosText(ByVal v1 As Variant, ByVal v2 As Variant) As String
    Dim u As Double, s As Double
    If Not TryParseDouble(Trim$(NzToText(v1)), u) Then u = 0
    If Not TryParseDouble(Trim$(NzToText(v2)), s) Then s = 0
    Dim net As Double: net = u - s
    If net <> 0 Then StornoIznosText = Format$(net, "#,##0.00")
End Function

' Iznos = Kolicina x Cena (prazno ako je proizvod 0).
Private Function StornoMnozi(ByVal vKol As Variant, ByVal vCena As Variant) As String
    Dim kol As Double, cena As Double
    If Not TryParseDouble(Trim$(NzToText(vKol)), kol) Then kol = 0
    If Not TryParseDouble(Trim$(NzToText(vCena)), cena) Then cena = 0
    Dim p As Double: p = kol * cena
    If p <> 0 Then StornoMnozi = Format$(p, "#,##0.00")
End Function

' --- Indeks lanca dokumenata (reverzni lookup-i preko BrojZbirne / FakturaID) ---

' Prazan recnik sa case-insensitive poredjenjem kljuceva.
Private Function NewDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewDict = d
End Function

Private Function DictGet(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Len(key) = 0 Then Exit Function
    If d.Exists(key) Then DictGet = CStr(d(key))
End Function

' Dodaj val u listu pod key (", "-spojeno, bez duplikata).
Private Sub DictAppend(ByVal d As Object, ByVal key As String, ByVal val As String)
    If d Is Nothing Then Exit Sub
    If Len(key) = 0 Or Len(val) = 0 Then Exit Sub
    If Not d.Exists(key) Then
        d.Add key, val
    Else
        Dim cur As String: cur = CStr(d(key))
        If InStr(1, ", " & cur & ", ", ", " & val & ", ", vbTextCompare) = 0 Then
            d(key) = cur & ", " & val
        End If
    End If
End Sub

' Izgradi indeks lanca: otpByZbr, fakByZbr, fakById, zbrByFak (sve case-insensitive).
Private Function BuildChainIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Set BuildChainIndex = idx

    Dim fakById As Object:  Set fakById = NewDict()
    Dim otpByZbr As Object: Set otpByZbr = NewDict()
    Dim fakByZbr As Object: Set fakByZbr = NewDict()
    Dim zbrByFak As Object: Set zbrByFak = NewDict()
    idx.Add "fakById", fakById
    idx.Add "otpByZbr", otpByZbr
    idx.Add "fakByZbr", fakByZbr
    idx.Add "zbrByFak", zbrByFak

    Dim d As Variant, i As Long, k As String, v As String

    ' Faktura: FakturaID -> BrojFakture
    d = GetTableData(TBL_FAKTURE)
    If Not IsEmpty(d) Then
        Dim cfi As Long, cfb As Long
        cfi = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
        cfb = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        If cfi > 0 And cfb > 0 Then
            For i = 1 To UBound(d, 1)
                k = Trim$(NzToText(d(i, cfi))): v = Trim$(NzToText(d(i, cfb)))
                If Len(k) > 0 And Not fakById.Exists(k) Then fakById.Add k, v
            Next i
        End If
    End If

    ' Otpremnica: BrojZbirne -> lista BrojOtpremnice
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(d) Then
        Dim coz As Long, cob As Long
        coz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
        cob = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        If coz > 0 And cob > 0 Then
            For i = 1 To UBound(d, 1)
                DictAppend otpByZbr, Trim$(NzToText(d(i, coz))), Trim$(NzToText(d(i, cob)))
            Next i
        End If
    End If

    ' Prijemnica: BrojZbirne <-> FakturaID  (i BrojZbirne -> BrojFakture preko fakById)
    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(d) Then
        Dim cpz As Long, cpf As Long
        cpz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        cpf = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)
        If cpz > 0 And cpf > 0 Then
            For i = 1 To UBound(d, 1)
                Dim z As String, fid As String
                z = Trim$(NzToText(d(i, cpz))): fid = Trim$(NzToText(d(i, cpf)))
                If Len(fid) > 0 Then
                    DictAppend zbrByFak, fid, z
                    If fakById.Exists(fid) Then DictAppend fakByZbr, z, CStr(fakById(fid))
                End If
            Next i
        End If
    End If
End Function

' Kolekcija 0-baznih 1D nizova -> 2D (1..n, 1..ncol). Empty ako je prazna.
Private Function StornoRowsTo2D(ByVal rows As Collection, ByVal ncol As Long) As Variant
    If rows Is Nothing Then Exit Function
    If rows.count = 0 Then Exit Function

    Dim arr() As Variant
    ReDim arr(1 To rows.count, 1 To ncol)
    Dim r As Long, c As Long, one As Variant
    For r = 1 To rows.count
        one = rows(r)
        For c = 1 To ncol
            arr(r, c) = one(c - 1)
        Next c
    Next r
    StornoRowsTo2D = arr
End Function

' ============================================================
' IZGUBLJENI OTKUP BLOKOVI (read + bezbedan re-point)
' Blok je "izgubljen" kad mu OtpremnicaID pokazuje na storniranu ili
' nepostojecu otpremnicu (a sam blok nije storniran). Koriste: dashboard
' dijagnostika (modHelpers.CheckVerwaisteDokumente) i panel Otkupni blokovi.
' ============================================================

' Vrati 2D (1..n, 1..7): OtkupID | BrojDokumenta | KooperantID | Datum |
' Kolicina | StaraOtpremnicaID | StaraOtpremnicaBroj. Empty ako nema.
Public Function GetLostOtkupBlokovi() As Variant
    On Error GoTo EH

    Dim otk As Variant: otk = GetTableData(TBL_OTKUP)
    If IsEmpty(otk) Then Exit Function

    ' mape otpremnica po ID-u: stornirane i aktivne (ID -> BrojOtpremnice)
    Dim storMap As Object: Set storMap = CreateObject("Scripting.Dictionary")
    Dim aktMap As Object: Set aktMap = CreateObject("Scripting.Dictionary")
    storMap.CompareMode = vbTextCompare: aktMap.CompareMode = vbTextCompare

    Dim otp As Variant: otp = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(otp) Then
        Dim oId As Long, oBr As Long, oSt As Long
        oId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
        oBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        oSt = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
        Dim j As Long, idv As String
        For j = 1 To UBound(otp, 1)
            idv = Trim$(NzToText(otp(j, oId)))
            If Len(idv) > 0 Then
                If UCase$(Trim$(NzToText(otp(j, oSt)))) = "DA" Then
                    If Not storMap.Exists(idv) Then storMap.Add idv, Trim$(NzToText(otp(j, oBr)))
                ElseIf Not aktMap.Exists(idv) Then
                    aktMap.Add idv, Trim$(NzToText(otp(j, oBr)))
                End If
            End If
        Next j
    End If

    Dim cId As Long, cBr As Long, cKoop As Long, cDat As Long
    Dim cKol As Long, cOtp As Long, cStor As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cStor = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(otk, 1)
        If UCase$(Trim$(NzToText(otk(i, cStor)))) = "DA" Then GoTo NextRow
        Dim otpId As String: otpId = Trim$(NzToText(otk(i, cOtp)))
        If Len(otpId) = 0 Then GoTo NextRow            ' nije vezan -> nije "izgubljen"

        Dim staleBroj As String: staleBroj = ""
        Dim lost As Boolean: lost = False
        If storMap.Exists(otpId) Then
            lost = True: staleBroj = CStr(storMap(otpId))
        ElseIf Not aktMap.Exists(otpId) Then
            lost = True: staleBroj = "(nepostojeca)"
        End If

        If lost Then
            rows.Add Array( _
                Trim$(NzToText(otk(i, cId))), _
                Trim$(NzToText(otk(i, cBr))), _
                Trim$(NzToText(otk(i, cKoop))), _
                otk(i, cDat), _
                otk(i, cKol), _
                otpId, _
                staleBroj)
        End If
NextRow:
    Next i

    GetLostOtkupBlokovi = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetLostOtkupBlokovi"
    GetLostOtkupBlokovi = Empty
End Function

' Osirocene prijemnice: AKTIVNE prijemnice cija zbirna (BrojZbirne) vise nije
' aktivna (stornirana ili ne postoji). Jedan red po BrojPrijemnice (Klasa I+II
' dele broj). Za re-point UI (frmDokumenta recovery panel).
' Kolone 1..7: BrojPrijemnice|Datum|Vrsta|Sorta|Kolicina|StaraZbirna|Status
Public Function GetOsirocenePrijemnice() As Variant
    On Error GoTo EH

    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojZbirne -> ima li bar jednu AKTIVNU zbirnu; i postoji li uopste.
    Dim aktZbr As Object: Set aktZbr = CreateObject("Scripting.Dictionary")
    Dim allZbr As Object: Set allZbr = CreateObject("Scripting.Dictionary")
    aktZbr.CompareMode = vbTextCompare: allZbr.CompareMode = vbTextCompare
    Dim zd As Variant: zd = GetTableData(TBL_ZBIRNA)
    If Not IsEmpty(zd) Then
        Dim zBr As Long, zSt As Long, zr As Long, zk As String
        zBr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        zSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
        For zr = 1 To UBound(zd, 1)
            zk = Trim$(NzToText(zd(zr, zBr)))
            If Len(zk) > 0 Then
                allZbr(zk) = True
                If UCase$(Trim$(NzToText(zd(zr, zSt)))) <> "DA" Then aktZbr(zk) = True
            End If
        Next zr
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long
    Dim cKol As Long, cZbr As Long, cSt As Long, cGen As Long
    cGen = GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID)
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Or cZbr = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim isStor As Boolean: isStor = False
        If cSt > 0 Then isStor = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not isStor Then
            Dim bz As String: bz = Trim$(NzToText(prj(i, cZbr)))
            Dim brp As String: brp = Trim$(NzToText(prj(i, cBr)))
            If Len(bz) > 0 And Len(brp) > 0 Then
                If Not aktZbr.Exists(bz) Then
                    ' Grupise se po BROJ + GENERACIJA, ne po samom broju: dva
                    ' kupca istog dana dele "1/ddmmyy", pa bi dedup po broju
                    ' spojio DVA dokumenta u jedan red - i recovery bi posle
                    ' prevezao onaj koji zatekne.
                    Dim gen As String: gen = Trim$(NzToText(prj(i, cGen)))
                    Dim kljuc As String: kljuc = brp & Chr$(1) & gen
                    If Not seen.Exists(kljuc) Then
                        seen(kljuc) = True
                        Dim st As String
                        If allZbr.Exists(bz) Then st = "zbirna stornirana" Else st = "zbirna ne postoji"
                        ' 8. kolona je GENERACIJA - dodata na KRAJ, pa legacy
                        ' panel (cita fiksno 1..7) ostaje netaknut.
                        rows.Add Array(brp, prj(i, cDat), Trim$(NzToText(prj(i, cVr))), _
                            Trim$(NzToText(prj(i, cSo))), StornoNumText(prj(i, cKol), "#,##0.00"), bz, st, gen)
                    End If
                End If
            End If
        End If
    Next i

    GetOsirocenePrijemnice = StornoRowsTo2D(rows, 8)
    Exit Function
EH:
    LogErr "modDokumenta.GetOsirocenePrijemnice"
    GetOsirocenePrijemnice = Empty
End Function

' Aktivne (ne-stornirane) zbirne, jedan red po BrojZbirne (Klasa I+II dele broj).
' Za izbor cilja u re-point UI. Kolone 1..5: BrojZbirne|Datum|Vrsta|Sorta|Kolicina
Public Function GetAktivneZbirne() As Variant
    On Error GoTo EH

    Dim zd As Variant: zd = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zd) Then Exit Function

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long, cKol As Long, cSt As Long
    cBr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    cDat = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM)
    cVr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VRSTA)
    cSo = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_SORTA)
    cKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(zd, 1)
        Dim isStor As Boolean: isStor = False
        If cSt > 0 Then isStor = (UCase$(Trim$(NzToText(zd(i, cSt)))) = "DA")
        If Not isStor Then
            Dim bz As String: bz = Trim$(NzToText(zd(i, cBr)))
            If Len(bz) > 0 Then
                If Not seen.Exists(bz) Then
                    seen(bz) = True
                    rows.Add Array(bz, zd(i, cDat), Trim$(NzToText(zd(i, cVr))), _
                        Trim$(NzToText(zd(i, cSo))), StornoNumText(zd(i, cKol), "#,##0.00"))
                End If
            End If
        End If
    Next i

    GetAktivneZbirne = StornoRowsTo2D(rows, 5)
    Exit Function
EH:
    LogErr "modDokumenta.GetAktivneZbirne"
    GetAktivneZbirne = Empty
End Function

' Stornirane prijemnice koje imaju AKTIVNE (osirocene) paleta-stavke. Za P1 pallet
' re-point. Kolone 1..6: BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica|StavkiAktivnih
Public Function GetPrijemniceSaOsirocenimPaletama() As Variant
    On Error GoTo EH
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(ps) Then Exit Function

    Dim stPrij As Object: Set stPrij = CreateObject("Scripting.Dictionary"): stPrij.CompareMode = vbTextCompare
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prj) Then
        Dim qId As Long, qSt As Long, q As Long
        qId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
        qSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
        For q = 1 To UBound(prj, 1)
            If qSt > 0 Then
                If UCase$(Trim$(NzToText(prj(q, qSt)))) = "DA" Then stPrij(Trim$(NzToText(prj(q, qId)))) = True
            End If
        Next q
    End If

    Dim cBr As Long, cPid As Long, cVr As Long, cSo As Long, cGajb As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    cVr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_VRSTA)
    cSo = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_SORTA)
    cGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cBr = 0 Or cPid = 0 Then Exit Function

    ' Grupise se po GENERACIJI storniranog dokumenta, ne po njegovom broju.
    ' Dve stornirane prijemnice razlicitih kupaca lako dele broj (racuna se po
    ' kupcu), pa bi grupisanje po broju spojilo DVA dokumenta u jedan recovery
    ' red - a prevezivanje bi posle pomerilo palete oba.
    ' Generacija se cita sa PRIJEMNICE, preko PrijemnicaID koji stavka nosi.
    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim sCnt As Object: Set sCnt = CreateObject("Scripting.Dictionary"): sCnt.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim brD As Object: Set brD = CreateObject("Scripting.Dictionary"): brD.CompareMode = vbTextCompare
    Dim genD As Object: Set genD = CreateObject("Scripting.Dictionary"): genD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(ps, 1)
        Dim stx As Boolean: stx = False
        If cSt > 0 Then stx = (UCase$(Trim$(NzToText(ps(i, cSt)))) = "DA")
        If Not stx Then
            Dim pid As String: pid = Trim$(NzToText(ps(i, cPid)))
            Dim br As String: br = Trim$(NzToText(ps(i, cBr)))
            If stPrij.Exists(pid) And Len(br) > 0 Then
                Dim gen As String: gen = GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, pid)
                Dim kl As String: kl = br & Chr$(1) & gen
                If Not gSum.Exists(kl) Then
                    gSum(kl) = 0&: sCnt(kl) = 0&: order.Add kl
                    vrD(kl) = Trim$(NzToText(ps(i, cVr))): soD(kl) = Trim$(NzToText(ps(i, cSo)))
                    brD(kl) = br: genD(kl) = gen
                End If
                If IsNumeric(ps(i, cGajb)) Then gSum(kl) = CLng(gSum(kl)) + CLng(ps(i, cGajb))
                sCnt(kl) = CLng(sCnt(kl)) + 1
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim k2 As String: k2 = CStr(v)
        ' 7. kolona je GENERACIJA - na KRAJ, pa legacy panel (cita 1..6) ostaje.
        rows.Add Array(CStr(brD(k2)), _
                       LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ, CStr(brD(k2)), COL_PRJ_DATUM), _
                       CStr(vrD(k2)), CStr(soD(k2)), CLng(gSum(k2)), CLng(sCnt(k2)), CStr(genD(k2)))
    Next v
    GetPrijemniceSaOsirocenimPaletama = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetPrijemniceSaOsirocenimPaletama"
    GetPrijemniceSaOsirocenimPaletama = Empty
End Function

' Aktivne (ne-stornirane) prijemnice = ciljevi za pallet re-point. Ukljucuje i
' paletizovane i NEpaletizovane (motor ReassignPaleteToPrijemnica_TX podrzava cilj
' bez svezih stavki: STEP1 tada ne undo-uje nista). Kolone 1..7:
'   BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica(KolAmb)|Zbirna|Paletizovana(Da/Ne)
Public Function GetAktivnePrijemnice() As Variant
    On Error GoTo EH
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojPrijemnice koje imaju bar jednu AKTIVNU paleta-stavku (-> Paletizovana=Da).
    Dim palBr As Object: Set palBr = CreateObject("Scripting.Dictionary"): palBr.CompareMode = vbTextCompare
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim xBr As Long, xSt As Long, x As Long
        xBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        xSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        For x = 1 To UBound(ps, 1)
            Dim sx As Boolean: sx = False
            If xSt > 0 Then sx = (UCase$(Trim$(NzToText(ps(x, xSt)))) = "DA")
            If Not sx Then palBr(Trim$(NzToText(ps(x, xBr)))) = True
        Next x
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long, cAmb As Long, cZbr As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function

    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim datD As Object: Set datD = CreateObject("Scripting.Dictionary"): datD.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim zbD As Object: Set zbD = CreateObject("Scripting.Dictionary"): zbD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim stp As Boolean: stp = False
        If cSt > 0 Then stp = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not stp Then
            Dim br As String: br = Trim$(NzToText(prj(i, cBr)))
            If Len(br) > 0 Then
                If Not gSum.Exists(br) Then
                    gSum(br) = 0&: order.Add br
                    datD(br) = prj(i, cDat): vrD(br) = Trim$(NzToText(prj(i, cVr)))
                    soD(br) = Trim$(NzToText(prj(i, cSo))): zbD(br) = Trim$(NzToText(prj(i, cZbr)))
                End If
                If cAmb > 0 Then If IsNumeric(prj(i, cAmb)) Then gSum(br) = CLng(gSum(br)) + CLng(prj(i, cAmb))
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim br2 As String: br2 = CStr(v)
        rows.Add Array(br2, datD(br2), CStr(vrD(br2)), CStr(soD(br2)), CLng(gSum(br2)), _
                       CStr(zbD(br2)), IIf(palBr.Exists(br2), "Da", "Ne"))
    Next v
    GetAktivnePrijemnice = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetAktivnePrijemnice"
    GetAktivnePrijemnice = Empty
End Function

' Bezbedan re-point izgubljenog bloka na ciljnu (aktivnu) otpremnicu:
' menja SAMO OtpremnicaID + BrojZbirne (cuva OtkupID -> uplate/ambalaza ostaju).
' Transakciono. Vraca False ako cilj ne postoji/storniran ili upis padne.
' Faza 7 korak 5: dual-write BrojOtpremnice na blok (denorm poslovni kljuc, stabilan
' kroz re-verziju otpremnice). otpID prazan -> ocisti (unbind). Guarded na kolonu
' (schema-drift safe) -> non-breaking dok se citanje ne prebaci sa OtpremnicaID.
Public Sub SetOtkupBrojOtpremnice(ByVal rowIndex As Long, ByVal otpID As String)
    On Error Resume Next
    If GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE) = 0 Then Exit Sub
    Dim broj As String: broj = ""
    If Len(Trim$(otpID)) > 0 Then _
        broj = Trim$(CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))
    UpdateCell TBL_OTKUP, rowIndex, COL_OTK_BROJ_OTPREMNICE, broj
End Sub

Public Function ReassignOtkupToOtpremnica_TX(ByVal otkupID As String, _
                                             ByVal targetOtpID As String) As Boolean
    Const SRC As String = "modDokumenta.ReassignOtkupToOtpremnica_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    If Len(Trim$(otkupID)) = 0 Or Len(Trim$(targetOtpID)) = 0 Then Exit Function

    ' cilj mora postojati (BrojOtpremnice nikad nije blank) i biti aktivan.
    ' NB: Stornirano JE blank za aktivnu otpremnicu -> ne sme se koristiti
    '     za proveru postojanja (LookupValue bi vratio Empty = lazno "ne postoji").
    Dim tBroj As Variant
    tBroj = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_OTP_BROJ)
    If IsEmpty(tBroj) Then Exit Function                        ' cilj stvarno ne postoji
    Dim tStor As String
    tStor = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_STORNIRANO))
    If UCase$(Trim$(tStor)) = "DA" Then Exit Function           ' cilj storniran

    Dim tZbr As String
    tZbr = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_OTP_BROJ_ZBIRNE))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim rows As Collection: Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows.count = 0 Then Err.Raise vbObjectError + 2600, SRC, "Otkup nije nadjen: " & otkupID

    Dim hasZbr As Boolean: hasZbr = (GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE) > 0)
    Dim k As Long
    For k = 1 To rows.count
        RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_OTPREMNICA_ID, targetOtpID, SRC
        ' ZBR-CHILD-01: broj i generacija idu zajedno (v. PoveziDeteNaZbirnu).
        If hasZbr Then PoveziDeteNaZbirnu TBL_OTKUP, rows(k), COL_OTK_BROJ_ZBIRNE, _
                                          tZbr, ZbirnaGeneracijaZaBroj(tZbr), SRC
        SetOtkupBrojOtpremnice rows(k), targetOtpID
    Next k

    tx.CommitTx
    Set tx = Nothing
    ReassignOtkupToOtpremnica_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignOtkupToOtpremnica_TX = False
End Function

' Re-point prijemnice na drugu (aktivnu) zbirnu po BrojZbirne. Koristi se posle
' storna otpremnice+zbirne (van autohladnjace): operater napravi novu otpremnicu+
' zbirnu pa "osirocenu" prijemnicu prevezuje na nju. Handle = BrojPrijemnice
' (citljiv, eksteran kad nije autohladnjaca); stvarna veza lanca = BrojZbirne.
' PrijemnicaID se NE dira -> faktura-stavke i paleta-stavke ostaju ispravne.
' prijemnicaGeneracijaID / zbirnaGeneracijaID: identitet prijemnice koja se
' prevezuje i zbirne na koju ide. Kad su poznati, bira se BAS taj dokument, a
' broj ostaje labela. Oba su opciona zbog zatecenih pozivalaca - bez njih vazi
' kapija nad jednoznacnoscu broja, na svakoj strani zasebno.
Public Function ReassignPrijemnicaToZbirna_TX(ByVal brPrijemnice As String, _
                                              ByVal targetBrZbirne As String, _
                                              Optional ByVal prijemnicaGeneracijaID As String = "", _
                                              Optional ByVal zbirnaGeneracijaID As String = "") As Boolean
    Const SRC As String = "modDokumenta.ReassignPrijemnicaToZbirna_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    brPrijemnice = Trim$(brPrijemnice)
    targetBrZbirne = Trim$(targetBrZbirne)
    If Len(brPrijemnice) = 0 Or Len(targetBrZbirne) = 0 Then Exit Function

    ' Cilj mora biti AKTIVNA zbirna (postoji + nije stornirana). ZbirnaID je uvek
    ' popunjen -> dokaz postojanja; Stornirano je blank za aktivnu pa se NE sme
    ' koristiti kao dokaz postojanja.
    ' Sa generacijom cilj se bira po IDENTITETU: LookupValue po broju uzima prvi
    ' pogodak. Broj zbirne JESTE jedinstven kad ga da generator (SuggestNextBroj
    ' za ZBR bumpuje sekvencu dok BrojZbirneExists ne kaze da je slobodan), ali
    ' ne i kad je unet rucno -- auto-broj se moze iskljuciti u Podesavanjima.
    Dim tgtIds As Object
    Set tgtIds = IdoviGeneracije(TBL_ZBIRNA, COL_ZBR_ID, zbirnaGeneracijaID)

    If Len(Trim$(zbirnaGeneracijaID)) > 0 And tgtIds.count = 0 Then Exit Function

    Dim tId As Variant
    If tgtIds.count > 0 Then
        tId = tgtIds.Keys()(0)
        If UCase$(Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, CStr(tId), _
                                             COL_STORNIRANO)))) = "DA" Then Exit Function
        ' Labela koja se upisuje mora da opisuje BAS izabrani dokument. Kad
        ' generacija odlucuje, broj se cita iz nje - ne veruje se onome sto je
        ' pozivalac poslao, jer bi neuskladjen par tiho upisao tudji broj.
        targetBrZbirne = Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, CStr(tId), _
                                                    COL_ZBR_BROJ)))
        If Len(targetBrZbirne) = 0 Then Exit Function
    Else
        ' CILJ SE NE RAZRESAVA PO REDU KOJI JE SLUCAJNO PRVI.
        '
        ' LookupValue po broju vraca PRVI pogodak i ne gleda storno. Posle
        ' storna jednog vlasnika prvi red sa tim brojem moze biti storniran
        ' dok pod istim brojem stoji AKTIVNA zbirna -- a tada je legitimno
        ' prevezivanje TIHO stajalo: bez poruke, bez loga, samo False.
        '
        ' ZbirnaPostoji gleda samo AKTIVNE redove, pa jedno tacno pitanje
        ' ("ima li aktivnog cilja pod ovim brojem") zamenjuje dva pogresna
        ' ("postoji li ijedan red" + "da li je PRVI storniran"). Test 125.
        '
        ' RAZRESAVA SE ISTIM POREDJENJEM KOJIM RADI KAPIJA ISPOD.
        '
        ' ZbirnaPostoji ovde NE valja iako pita pravu stvar: on poredi bez
        ' obzira na velicinu slova (StrComp vbTextCompare), a VlasniciPoBroju
        ' -- koji stoji iza kapije -- poredi TACNO. Sa "zb-test-kask" bi
        ' postojanje reklo DA, kapija bi videla NULA vlasnika (a ona hvata samo
        ' n > 1, pa bi propustila), i u tblPrijemnica i tblPaletaStavka bi se
        ' upisala labela POZIVAOCA umesto one iz tabele. Postojanje i
        ' vlasnistvo bi govorili o dve razlicite stvari.
        '
        ' Jedan poziv daje oba odgovora: broj AKTIVNIH vlasnika pod tim brojem.
        ' Nula znaci "nema aktivnog cilja", a vise od jedan hvata kapija ispod
        ' -- koja ostaje zbog svoje poruke. Test 125.
        Dim aktivniVlasnici As Long
        aktivniVlasnici = VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                          targetBrZbirne, SRC, False, _
                                          Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count
        If aktivniVlasnici = 0 Then Exit Function

        ' Bez generacije CILJ MORA BITI JEDNOZNACAN, i to ISTORIJSKI.
        '
        ' Vlasnistvo zbirne je vozac + kupac, isti par koji koriste
        ' StornoZbirna i ApplyGeneracijaID. Broji se IKAD, ne samo aktivni:
        ' storniran vlasnik i dalje ima AKTIVNU decu, pa posle njegovog storna
        ' ostane jedan aktivan i broj IZGLEDA jednoznacan -- a nije.
        '
        ' Ovde je to obavezno, ne opciono: recovery panel u frmDokumenta zove
        ' ovu funkciju sa DVA argumenta (bez generacije) i BEZ ijedne spoljne
        ' kapije, pa se funkcija mora braniti sama. Putanje iz modStornoFlow
        ' imaju ZbirnaBrojJeDvosmislenIkad iznad sebe; ta nema nista.
        '
        ' Veza koja se upisuje je gola labela (COL_PRJ_BROJ_ZBIRNE,
        ' COL_PALS_BROJ_ZBIRNE), pa bi dete zavrsilo vezano za broj koji
        ' pripada dvama vlasnickim tokovima. Test 125.
        RequireJedanVlasnikIkadPoBroju TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, _
                                       SRC, COL_ZBR_VOZAC, COL_ZBR_KUPAC
    End If

    Dim cBrPrij As Long, cBrZbr As Long, cStorno As Long
    cBrPrij = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBrPrij = 0 Or cBrZbr = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function

    ' IZVOR SE BIRA PO IDENTITETU. Sa poznatom generacijom menjaju se samo redovi
    ' TE generacije (Klasa I i II zajedno, i nijedan tudji). Bez nje ostaje izbor
    ' po broju, ali tek posto se dokaze da je broj jednoznacan - inace bi se
    ' prevezala i prijemnica drugog kupca koja deli broj.
    Dim srcIds As Object
    Set srcIds = IdoviGeneracije(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaGeneracijaID)
    If Len(Trim$(prijemnicaGeneracijaID)) > 0 Then
        ' ZADATA generacija koja se ne razresava je GRESKA, ne poziv na fallback.
        ' Pad na broj bi znacio: pozivalac je rekao "bas ovaj dokument", nije
        ' nadjen, pa se dira nesto drugo. Prazan argument je nesto sasvim drugo
        ' (legacy zapis) i za njega fallback ostaje.
        If srcIds.count = 0 Then Exit Function
    Else
        ' AKTIVNI vlasnik mora biti jedan. NAMERNO nije IKAD -- probano pa
        ' povuceno, jer je bila sira zabrana bez dokaza pogresne mutacije.
        '
        ' Zastita ovde nije jedna kapija nego SLOJEVI, i svaki radi svoj posao:
        '   zaglavlje prijemnice  -> u targetRows ulaze samo AKTIVNI redovi,
        '                            pa storniran dokument drugog kupca ne
        '                            moze da se pomeri (test 126);
        '   paletna stavka sa ID  -> odlucuje IDENTITET (docIds.Exists);
        '   legacy stavka bez ID  -> brojDvosmislen racuna IKAD i puca, uz
        '                            rollback cele transakcije.
        '
        ' IKAD kapija ovde bi zabranila i potpuno resiv legacy oporavak: broj
        ' prijemnice je numerisan PO KUPCU, pa je kolizija ocekivana, a writer
        ' vec ume da razdvoji aktivan dokument. Sira zabrana bi operatera
        ' terala na rucni rad bez ijedne izmerene koristi.
        RequireJedanVlasnikPoBroju TBL_PRIJEMNICA, COL_PRJ_BROJ, brPrijemnice, SRC, _
                                   COL_PRJ_KUPAC
    End If
    Dim cPrjId As Long: cPrjId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)

    ' Aktivni redovi prijemnice (Klasa I i II dele broj I generaciju).
    Dim targetRows As Collection: Set targetRows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If PripadaIzvoru(data, i, cBrPrij, cPrjId, brPrijemnice, srcIds) Then
            If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> "DA" Then
                targetRows.Add i
            End If
        End If
    Next i
    If targetRows.count = 0 Then Exit Function                  ' nema aktivne prijemnice

    ' ZBR-CHILD-01: generacija CILJA, jednom za sve redove.
    '
    ' Kad je pozivalac zadao generaciju, ona JE identitet cilja -- iznad je vec
    ' dokazano da se razresava (tgtIds). Kad nije, izvodi se iz broja i prazna je
    ' ako broj nije jednoznacan; tada dete ostaje bez generacije, sto je isto
    ' stanje kao pre ove kolone.
    Dim genCilja As String
    genCilja = Trim$(NzToText(zbirnaGeneracijaID))
    If Len(genCilja) = 0 Then genCilja = ZbirnaGeneracijaZaBroj(targetBrZbirne)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    Dim k As Long
    For k = 1 To targetRows.count
        PoveziDeteNaZbirnu TBL_PRIJEMNICA, targetRows(k), COL_PRJ_BROJ_ZBIRNE, _
                           targetBrZbirne, genCilja, SRC
    Next k

    ' Sledljivost: paletne stavke te prijemnice moraju dobiti NOVU BrojZbirne, inace
    ' ostaju sa mrtvom zbirnom (paleta -> zbirna -> kooperanti pukne). Menja se samo
    ' veza (BrojZbirne); roba/pripadnost prijemnici ostaje.
    '
    ' PO PrijemnicaID, NE PO BROJU. Ovaj upis je ranije isao po BrojPrijemnice i
    ' time je ponistavao ceo izbor iznad: tblPrijemnica se menjala samo dokumentu
    ' A, a paletne stavke i dokumenta A i dokumenta B. Dokument B je tako
    ' postajao SAM SEBI PROTIVRECAN -- prijemnica na staroj zbirni, njena paleta
    ' na novoj. Redovi su vec izabrani u targetRows; odatle se citaju ID-evi.
    Dim docIds As Object: Set docIds = CreateObject("Scripting.Dictionary")
    docIds.CompareMode = vbTextCompare
    Dim q As Long, pidT As String
    For q = 1 To targetRows.count
        If cPrjId > 0 Then
            pidT = Trim$(NzToText(data(targetRows(q), cPrjId)))
            If Len(pidT) > 0 Then docIds(pidT) = True
        End If
    Next q

    ' Stavka bez PrijemnicaID (zatecen zapis) sme po broju SAMO ako taj broj nosi
    ' jedan dokument. Kad ga nose dva, ne moze se utvrditi cija je, pa se staje
    ' umesto da se pogodi -- transakcija se povlaci u celini.
    Dim brojDvosmislen As Boolean
    brojDvosmislen = (VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, brPrijemnice, SRC, _
                                      True, Array(COL_PRJ_KUPAC)).count > 1)

    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim pBr As Long, pZb As Long, pSt As Long, pPid As Long
        pBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        pZb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        pSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        pPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
        If pBr > 0 And pZb > 0 Then
            Dim r2 As Long, pidS As String, pripada As Boolean
            For r2 = 1 To UBound(ps, 1)
                If pSt = 0 Or UCase$(Trim$(CStr(ps(r2, pSt)))) <> "DA" Then
                    pidS = ""
                    If pPid > 0 Then pidS = Trim$(NzToText(ps(r2, pPid)))
                    If Len(pidS) > 0 Then
                        pripada = docIds.Exists(pidS)
                    Else
                        If Trim$(CStr(ps(r2, pBr))) = brPrijemnice And brojDvosmislen Then
                            Err.Raise vbObjectError + 7311, SRC, _
                                      "Paletna stavka broja '" & brPrijemnice & "' nema " & _
                                      "PrijemnicaID, a taj broj nose dokumenta vise kupaca. " & _
                                      "Ne moze se utvrditi cija je -- prevezivanje je " & _
                                      "prekinuto. Popuni PrijemnicaID pa ponovi."
                        End If
                        pripada = (Trim$(CStr(ps(r2, pBr))) = brPrijemnice)
                    End If
                    If pripada Then
                        PoveziDeteNaZbirnu TBL_PALETA_STAVKA, r2, COL_PALS_BROJ_ZBIRNE, _
                                           targetBrZbirne, genCilja, SRC
                    End If
                End If
            Next r2
        End If
    End If

    tx.CommitTx
    Set tx = Nothing
    ReassignPrijemnicaToZbirna_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignPrijemnicaToZbirna_TX = False
End Function

' ============================================================
' OM ULAZ (ambalaza + kes) -- servis bez UI-ja
' Premesteno iz frmDokumenta (RF-05): nema referenci na kontrole, a smesteno u
' modul moze da se testira bez instanciranja forme (core guard za smer ambalaze).
' ============================================================

Public Function SaveOMUlaz_TX(ByVal datum As Date, _
                              ByVal brojDok As String, _
                              ByVal stanicaNaziv As String, _
                              ByVal stanicaID As String, _
                              ByVal vozacID As String, _
                              ByVal tipAmb As String, _
                              ByVal kolAmb As Long, _
                              ByVal vrstaVoca As String, _
                              ByVal novac As Double, _
                              ByVal kooperantID As String, _
                              ByVal primalacDisplay As String, _
                              ByVal otkupID As String, _
                              ByVal tipNovca As String, _
                              ByVal koopSmer As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1501, "SaveOMUlaz_TX", _
                  Poruka("DOK_ERR_NEMA_AMBALAZE_NOVCA")
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP

    If kolAmb > 0 Then
        Select Case koopSmer
        Case "IZDAVANJE"
            ' OM IZDAJE prazne kooperantu -> DVOJNI upis (bez vozaca):
            '   1) Kooperant ULAZ (dobija prazne), 2) OM/Stanica IZLAZ (razduzenje OM).
            If Trim$(kooperantID) = "" Then
                Err.Raise vbObjectError + 1503, "SaveOMUlaz_TX", _
                          "Izdavanje kooperantu: kooperant je obavezan."
            End If
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1504, "SaveOMUlaz_TX", _
                          "Izdavanje kooperantu: OM (otkupno mesto) je obavezan za razdu" & ChrW(382) & "enje."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", kooperantID, "Kooperant", _
                          "", brojDok, DOK_TIP_OM_IZLAZ_KOOP
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", stanicaID, "Stanica", _
                          "", brojDok, DOK_TIP_OM_IZLAZ_KOOP
        Case "PRIJEM"
            ' KOOPERANT VRACA prazne na OM (povrat) -> DVOJNI upis, mirror izdavanja:
            '   1) Kooperant IZLAZ (predaje prazne), 2) OM/Stanica ULAZ (zaduzenje OM).
            If Trim$(kooperantID) = "" Then
                Err.Raise vbObjectError + 1505, "SaveOMUlaz_TX", _
                          "Prijem od kooperanta: kooperant je obavezan."
            End If
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1506, "SaveOMUlaz_TX", _
                          "Prijem od kooperanta: OM (otkupno mesto) je obavezan za zadu" & ChrW(382) & "enje."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", kooperantID, "Kooperant", _
                          "", brojDok, DOK_TIP_OM_ULAZ_KOOP
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", stanicaID, "Stanica", _
                          "", brojDok, DOK_TIP_OM_ULAZ_KOOP
        Case "IZDATO_OM"
            ' Vozac raspodeljuje prazne na OM (revers ide na OM): OM (Stanica) ULAZ +
            ' vozac (inverzno Izlaz = vozac se razduzuje). Vozac je prethodno zaduzen
            ' kod kupca (prijemnica-povrat / kupci-izlaz) -> hladnjaca se NE knjizi ovde.
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1507, "SaveOMUlaz_TX", _
                          "Izdato OM: OM (otkupno mesto) je obavezan."
            End If
            If Trim$(vozacID) = "" Then
                Err.Raise vbObjectError + 1509, "SaveOMUlaz_TX", _
                          "Izdato OM: vozac je obavezan (firma<->OM ide preko vozaca)."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Ulaz", stanicaID, "Stanica", _
                          vozacID, brojDok, DOK_TIP_OM_ULAZ_FIRMA
        Case "PRIJEM_OD_OM"
            ' OM vraca prazne vozacu (revers ide na OM): OM (Stanica) IZLAZ + vozac
            ' (inverzno Ulaz = vozac se zaduzuje). Vozac kasnije razduzuje firmi
            ' (hladnjaci) kroz postojece kupac tokove -> hladnjaca se NE knjizi ovde.
            If Trim$(stanicaID) = "" Then
                Err.Raise vbObjectError + 1508, "SaveOMUlaz_TX", _
                          "Prijem od OM: OM (otkupno mesto) je obavezan."
            End If
            If Trim$(vozacID) = "" Then
                Err.Raise vbObjectError + 1510, "SaveOMUlaz_TX", _
                          "Prijem od OM: vozac je obavezan (firma<->OM ide preko vozaca)."
            End If
            TrackAmbalaza datum, tipAmb, kolAmb, _
                          "Izlaz", stanicaID, "Stanica", _
                          vozacID, brojDok, DOK_TIP_OM_IZLAZ_FIRMA
        Case Else
            ' Smer je OBAVEZAN uz kolicinu ambalaze. Ranije je ovde tiho knjizen
            ' legacy "OM prima od vozaca" (Stanica ULAZ, DOK_TIP_OM_ULAZ), pa je
            ' prazan/nepoznat smer davao pogresan ledger red bez ijedne poruke.
            ' UI blokira prazan smer, ovo je core guard za sve ostale pozivaoce.
            Err.Raise vbObjectError + 1511, "SaveOMUlaz_TX", _
                      "Nepoznat smer ambalaze '" & koopSmer & "'. Dozvoljeni: " & _
                      "IZDAVANJE, PRIJEM, IZDATO_OM, PRIJEM_OD_OM."
        End Select
    End If

    If novac > 0 Then
        Dim novacID As String

        ' KAPIJA VLASNISTVA I TRENUTNOG OSTATKA (AUD-026 obrazac). UI je ovo vec
        ' proverio, ali nad snimkom iz trenutka kad je lista punjena -- a izmedju
        ' punjenja i potvrde stanje se moze promeniti. Writer zato ne veruje
        ' parametrima nego cita stanje SADA. Vazi za SVAKOG pozivaoca, pa i za
        ' legacy frmDokumenta.
        Dim blokErr As String
        blokErr = IsplataBlokProblem(otkupID, kooperantID, stanicaID, novac)
        If Len(blokErr) > 0 Then
            Err.Raise vbObjectError + 1512, "SaveOMUlaz_TX", blokErr
        End If

        ' ISTA KAPIJA ZA AVANS OTKUPNOG MESTA. Blok i faktura su vec bili
        ' zasticeni i u writer-u; avans je ostajao samo na UI sloju
        ' (modNovacUnos.IsplataValidiraj), pa je writer bio poslednja linija
        ' za dve od tri stvari koje isti dokument moze da prekoraci.
        '
        ' Kes isplata kooperantu TROSI avans koji je firma dala otkupnom mestu -
        ' GetOMAvansSaldo ga i racuna kao "avansi firma->OM minus vec izdate
        ' NOV_KES_OTKUPAC_KOOP isplate". Isplata preko tog salda pravi minus
        ' koji se nigde ne vidi do sledeceg obracuna.
        If tipNovca = NOV_KES_OTKUPAC_KOOP Then
            Dim avansSaldo As Double
            avansSaldo = ZaokruziNovac(GetOMAvansSaldo(stanicaID))
            If ZaokruziNovac(novac) > avansSaldo Then
                Err.Raise vbObjectError + 1513, "SaveOMUlaz_TX", _
                          Poruka("NOVAC_ERR_AVANS_PREKO") & " " & _
                          Format$(avansSaldo, "#,##0.00")
            End If
        End If

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=stanicaNaziv, _
            partnerId:=stanicaID, _
            entitetTip:="OM", _
            omID:=stanicaID, _
            kooperantID:=kooperantID, _
            fakturaID:="", _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=0, _
            isplata:=novac, _
            napomena:=primalacDisplay, _
            otkupID:=otkupID)

        If novacID = "" Then
            Err.Raise vbObjectError + 1502, "SaveOMUlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

        If otkupID <> "" Then
            UpdateOtkupStatus otkupID
        End If
    End If

    tx.CommitTx

    Set tx = Nothing

    SaveOMUlaz_TX = True
    Exit Function

EH:
    LogErr "SaveOMUlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOMUlaz_TX = False
End Function

