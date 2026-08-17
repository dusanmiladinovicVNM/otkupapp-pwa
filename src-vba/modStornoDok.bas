Attribute VB_Name = "modStornoDok"
'=====================================================================
' modStornoDok - STORNO DOKUMENTA PO TIPU I BROJU, bez ijedne kontrole.
'
' Cetvrti modul istog oblika: modOtkupUnos (F1), modDokUnos (F2-F4),
' modNovacUnos (F5-F7), modStornoDok (F8). Razlog je isti - poslovni
' posao ne sme da zivi u formi, jer ga onda drugi ekran ne moze pozvati
' bez prepisivanja.
'
' ODAKLE DOLAZE PRAVILA: frmDokumenta.btnStorno_Click. Tamo je jedan
' Select Case po tipu dokumenta radio tri stvari pomesane: razresavanje
' broja u ID, kapije koje odbijaju storno, i poziv pravog Storno*_TX.
' Ovde su razdvojene u tri javne rutine, pa ekran moze da pita "sme li"
' pre nego sto uopste ponudi dugme:
'
'   StornoRazlog(tip, broj, opcija)    "" = sme; inace RAZLOG odbijanja
'   StornoPregled(tip, broj, opcija)   tekst koji operater vidi pre potvrde
'   StornoIzvrsi(tip, broj, opcija, poruka)  poziva tacan Storno*_TX
'
' KAPIJE NISU OVDE - sve su vec javne u modStorno i dizu ih i legacy
' forma i svaki drugi pozivalac:
'   LookupActiveID            dokument postoji i nije vec storniran
'   ResolveNovacForStorno     izvod se ne stornira parcijalno; broj sa
'                             vise aktivnih redova trazi NovacID
'   ResolveIzvodZaStorno      "broj" ili "broj/racun" -> jedan izvod
'   GetIzvodStornoBlokade     preflight: razlog pre potvrde, ne tih pad
'   ActiveAmbalazaDokExists   revers po broju I po smeru
' Ovaj modul ih samo redja po tipu. Duplirane provere se ne pisu.
'
' TIP DOKUMENTA je kljuc rezima novog UI-ja (modScrDokumenti.modeKey),
' ne natpis iz legacy combo-a. Prevod postoji samo na jednom mestu -
' u Select Case-u ispod - pa ekran nikad ne salje tekst sa dijakritikom.
'
' OPCIJA nosi ono sto tip ne moze da izvede iz broja:
'   REVERSI -> DokumentTip (cetiri smera dele isti brojevni niz, pa broj
'              sam po sebi ne kaze koji je red u tblAmbalaza)
'   IZVOD   -> ishod (IZVOD_STORNO_REMAP / IZVOD_STORNO_REIMPORT); to je
'              ODLUKA operatera o PDF-u, ne pravilo, pa je pita ekran
'   ostali  -> prazno
'
' OD v6-ui-120 modul nosi jos dve stvari (Faza D/13):
'   framework ispravke  - prevod tipa u modStornoFlow i cetiri moda
'                         (odeljak 4); sam IZBOR moda je pitanje operateru,
'                         pa je u ekranu, ne ovde
'   prefill iz storniranog (Z10, odeljak 5) - opis vrednosti storniranog
'                         dokumenta, izdvojen iz cetiri legacy Prefill*
'                         rutine koje su bile vezane za kontrole
'
' STA OVAJ MODUL NAMERNO NE RADI: nema Undo operacije, "Nedovrseno" ni
' Recovery panela (stavka 14 Faze D) - oni i dalje zive samo u
' frmDokumenta. Nema ni multiselect otkupnih blokova uz DUPLI/PONISTENJE:
' to je deo legacy overlay panela, koji nije prenet.
'
' VAZNO: legacy frmDokumenta OSTAJE netaknut i potpuno operativan. Ovaj
' modul je drugi pozivalac istih poslovnih rutina, ne zamena za formu.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const STORNODOK_BUILD As String = "v6-ui-122"

' Kljucevi tipova - isti kao modScrDokumenti.modeKey, plus dva koja novi
' UI jos nema kao rezim (fakture i bankovni izvodi se u njemu ne kreiraju,
' ali se stornirati moraju - legacy ih ima u istom combo-u).
Public Const STIP_OTKUP      As String = "OTKUP"
Public Const STIP_OTPREMNICA As String = "OTPREMNICA"
Public Const STIP_ZBIRNA     As String = "ZBIRNA"
Public Const STIP_PRIJEMNICA As String = "PRIJEMNICA"
Public Const STIP_ISPLATE    As String = "AMB_ISPLATE"
Public Const STIP_UPLATE     As String = "AMB_UPLATE"
Public Const STIP_REVERSI    As String = "REVERSI"
Public Const STIP_FAKTURA    As String = "FAKTURA"
Public Const STIP_IZVOD      As String = "IZVOD"

'=====================================================================
' 1) SME LI SE STORNIRATI - "" znaci da sme
'
' Zove se PRE potvrde. Razlog je tekst za operatera; vecina razloga
' dolazi gotova iz modStorno (ResolveNovacForStorno, ResolveIzvodZaStorno,
' GetIzvodStornoBlokade), jer je tamo i pravilo koje ih pravi.
'=====================================================================
' docID: KANONSKI IDENTITET reda koji je operater izabrao u F8 --
' GeneracijaID za robna dokumenta, PK za novac i fakturu. Kad je poznat,
' dokument se NE trazi ponovo po broju. Opcion je zbog legacy forme i
' zatecenih zapisa bez generacije; tada vazi kapija nad jednoznacnoscu broja.
' Postoji li AKTIVAN dokument koji je operater izabrao?
'
' Sa docID se pita za BAS TAJ dokument. Bez njega ostaje provera po broju --
' ista koju je preflight oduvek radio.
'
' Zasto je bitno: preflight je do sada primao identitet pa ga ignorisao, i za
' novac je i dalje govorio "broj je dvosmislen, treba NovacID" iako mu je F8
' NovacID upravo poslao. StornoIzvrsi nize je vec bio ispravan, ali se do njega
' nije stizalo -- kapija iznad je zaustavljala operaciju.
Private Function AktivanPoIdentitetu(ByVal tblName As String, ByVal brojCol As String, _
                                     ByVal idCol As String, ByVal broj As String, _
                                     ByVal docID As String) As Boolean
    On Error Resume Next
    If Len(Trim$(docID)) = 0 Then
        AktivanPoIdentitetu = (Len(LookupActiveID(tblName, brojCol, broj, idCol)) > 0)
        Exit Function
    End If
    Dim ids As Object: Set ids = IdoviGeneracije(tblName, idCol, docID)
    If ids.count = 0 Then Exit Function
    AktivanPoIdentitetu = (UCase$(Trim$(NzToText( _
        LookupValue(tblName, idCol, CStr(ids.Keys()(0)), COL_STORNIRANO)))) <> "DA")
End Function

Public Function StornoRazlog(ByVal tip As String, ByVal broj As String, _
                             ByVal opcija As String, _
                             Optional ByVal docID As String = "") As String
    Dim razlog As String, izvBroj As String, izvRacun As String, errDesc As String
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then
        StornoRazlog = Poruka("STORNO_ERR_NEMA_BROJA")
        Exit Function
    End If

    Select Case tip
        Case STIP_OTKUP
            If Not AktivanPoIdentitetu(TBL_OTKUP, COL_OTK_BR_DOK, COL_OTK_ID, broj, docID) Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_OTPREMNICA
            If Not AktivanPoIdentitetu(TBL_OTPREMNICA, COL_OTP_BROJ, COL_OTP_ID, broj, docID) Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_ZBIRNA
            ' StornoZbirna_TX prima BROJ (ne ID) i sam razresava; provera
            ' postojanja je ista kao za ostale robne dokumente.
            If Not AktivanPoIdentitetu(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_ID, broj, docID) Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_PRIJEMNICA
            If Not AktivanPoIdentitetu(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_ID, broj, docID) Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_FAKTURA
            ' Identitet fakture JE njen PK, pa se proverava direktno.
            If Len(Trim$(docID)) > 0 Then
                If UCase$(Trim$(NzToText(LookupValue(TBL_FAKTURE, COL_FAK_ID, docID, _
                                                     COL_STORNIRANO)))) = "DA" Then _
                    StornoRazlog = NijePronadjen(broj)
            ElseIf Len(LookupActiveID(TBL_FAKTURE, COL_FAK_BROJ, broj, COL_FAK_ID)) = 0 Then
                StornoRazlog = NijePronadjen(broj)
            End If

        Case STIP_ISPLATE, STIP_UPLATE
            ' Ceo razlog dolazi iz modStorno: izvod se ne stornira
            ' parcijalno, a broj sa vise aktivnih redova (avans raspodela
            ' deli isti broj) trazi NovacID umesto tihog storna jednog reda.
            ' Sa poznatim NovacID-em nema sta da se razresava: pita se BAS taj
            ' red. ResolveNovacForStorno ostaje za legacy poziv bez ID-a -- i
            ' bas on je javljao "broj je dvosmislen" i kad ID postoji.
            If Len(Trim$(docID)) > 0 Then
                If UCase$(Trim$(NzToText(LookupValue(TBL_NOVAC, COL_NOV_ID, docID, _
                                                     COL_STORNIRANO)))) = "DA" Then _
                    StornoRazlog = NijePronadjen(broj)
            Else
                ResolveNovacForStorno broj, razlog
                StornoRazlog = razlog
            End If

        Case STIP_REVERSI
            If Len(Trim$(opcija)) = 0 Then
                StornoRazlog = Poruka("STORNO_ERR_NEMA_SMERA")
            ElseIf Not ActiveAmbalazaDokExists(broj, opcija) Then
                StornoRazlog = NijePronadjen(broj)
            End If

        Case STIP_IZVOD
            If Not ResolveIzvodZaStorno(broj, izvBroj, izvRacun, razlog) Then
                StornoRazlog = razlog
            Else
                ' Preflight: ako nesto blokira, operater vidi RAZLOG pre
                ' potvrde, a ne tihi neuspeh posle nje.
                StornoRazlog = GetIzvodStornoBlokade(izvBroj, izvRacun)
            End If

        Case Else
            StornoRazlog = Poruka("STORNO_ERR_NEPOZNAT_TIP") & " " & tip
    End Select
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modStornoDok.StornoRazlog"
    StornoRazlog = Poruka("STORNO_ERR_RAZRESENJE") & " " & errDesc
End Function

'=====================================================================
' 2) STA OPERATER VIDI PRE POTVRDE
'
' Za izvod je to pun pregled (koliko stavki, koji racun, koliko novca) -
' izvod pada u celosti, pa potvrda bez pregleda ne bi bila informisana.
' Za ostale tipove je dovoljno ime dokumenta i broj.
'=====================================================================
Public Function StornoPregled(ByVal tip As String, ByVal broj As String, _
                              ByVal opcija As String) As String
    Dim izvBroj As String, izvRacun As String, razlog As String
    On Error Resume Next
    If tip = STIP_IZVOD Then
        If ResolveIzvodZaStorno(broj, izvBroj, izvRacun, razlog) Then
            StornoPregled = GetIzvodPregled(izvBroj, izvRacun)
            Exit Function
        End If
    End If
    StornoPregled = TipNaziv(tip, opcija) & " " & broj
End Function

'=====================================================================
' 3) IZVRSI
'
' Poruka je ono sto operater vidi POSLE uspeha - nije uvek "Stornirano".
' Dva slucaja nose vise od potvrde:
'   ZBIRNA  - aktivna prijemnica ostaje vezana za storniranu zbirnu
'             (StornoZbirna namerno ne kaskadira), pa sledljivost visi
'             dok se prijemnica ne preveze. Legacy to isto upozorenje
'             ispisuje; bez njega operater ne zna da mu je ostao posao.
'   IZVOD   - StornoIzvod_TX sam vraca izvestaj (koliko redova, koji ishod).
'=====================================================================
Public Function StornoIzvrsi(ByVal tip As String, ByVal broj As String, _
                             ByVal opcija As String, ByRef poruka As String, _
                             Optional ByVal docID As String = "") As Boolean
    Dim ok As Boolean, novID As String, razlog As String, errDesc As String
    Dim izvBroj As String, izvRacun As String, izvInfo As String
    Dim fakID As String, vezPrij As String
    On Error GoTo EH
    poruka = ""
    broj = Trim$(broj)

    Select Case tip
        Case STIP_OTKUP
            ' Klasa I i II dele isti BrDok (zaseban red po klasi) -> stornira
            ' se ceo dokument, ne jedan red. Isto sto radi F1 lista.
            ok = StornoOtkupByBrDok_TX(broj, docID)

        Case STIP_OTPREMNICA
            ' i ovde klase dele broj
            ok = StornoOtpremnicaByBroj_TX(broj, docID)

        Case STIP_ZBIRNA
            ok = StornoZbirna_TX(broj, docID)
            If ok Then
                vezPrij = NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, broj, COL_PRJ_BROJ))
                ' KVALIFIKOVANO, i mora ostati: izlazni parametar se zove "poruka",
                ' a VBA je case-insensitive -- pa nekvalifikovan poziv te funkcije
                ' unutar ove procedure nije poziv funkcije nego indeksiranje tog
                ' String parametra. Compile error "Expected array", i to samo u
                ' Debug > Compile: nijedna suite ovu proceduru nije zvala, a VBA
                ' proceduru kompajlira tek kad se pozove.
                If Len(vezPrij) > 0 Then _
                    poruka = modPoruke.Poruka("STORNO_MSG_ZBIRNA_PRIJ") & " " & vezPrij
            End If

        Case STIP_PRIJEMNICA
            ok = StornoPrijemnicaByBroj_TX(broj, docID)

        Case STIP_FAKTURA
            ' Identitet fakture JE njen PK -- kad ga ekran posalje, ne trazi se
            ' ponovo po broju (LookupActiveID uzima prvi pogodak).
            fakID = docID
            If Len(fakID) = 0 Then _
                fakID = LookupActiveID(TBL_FAKTURE, COL_FAK_BROJ, broj, COL_FAK_ID)
            If Len(fakID) = 0 Then
                poruka = NijePronadjen(broj)
                Exit Function
            End If
            ok = StornoFaktura_TX(fakID)

        Case STIP_ISPLATE, STIP_UPLATE
            ' StornoNovac_TX ocekuje NovacID, a mreza pokazuje BROJ. Isto
            ' razresavanje kao u StornoRazlog - i ovde, jer se izmedju
            ' provere i potvrde stanje moglo promeniti.
            ' Isto i za novac: NovacID iz kliknutog reda. ResolveNovacForStorno
            ' ostaje za legacy pozivaoce koji salju samo broj.
            novID = docID
            If Len(novID) = 0 Then
                novID = ResolveNovacForStorno(broj, razlog)
                If Len(razlog) > 0 Then
                    poruka = razlog
                    Exit Function
                End If
            End If
            ok = StornoNovac_TX(novID)

        Case STIP_REVERSI
            ok = StornoOMKoopByBrDok_TX(broj, opcija)

        Case STIP_IZVOD
            If Not ResolveIzvodZaStorno(broj, izvBroj, izvRacun, razlog) Then
                poruka = razlog
                Exit Function
            End If
            If opcija <> IZVOD_STORNO_REMAP And opcija <> IZVOD_STORNO_REIMPORT Then
                poruka = modPoruke.Poruka("STORNO_ERR_NEMA_ISHODA")
                Exit Function
            End If
            ok = StornoIzvod_TX(izvBroj, izvRacun, opcija, izvInfo)
            If ok Then poruka = izvInfo

        Case Else
            poruka = modPoruke.Poruka("STORNO_ERR_NEPOZNAT_TIP") & " " & tip
            Exit Function
    End Select

    StornoIzvrsi = ok
    If ok And Len(poruka) = 0 Then poruka = modPoruke.Poruka("STORNO_MSG_OK")
    If Not ok And Len(poruka) = 0 Then poruka = modPoruke.Poruka("STORNO_ERR_NEUSPEH") & " " & broj
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modStornoDok.StornoIzvrsi"
    poruka = modPoruke.Poruka("STORNO_ERR_NEUSPEH") & " " & broj & ": " & errDesc
End Function

'=====================================================================
' 4) FRAMEWORK ISPRAVKE - storno NIJE Da/Ne
'
' Cetiri moda iz modStornoFlow; koji je izabran menja STA se desava sa
' nizvodnim tokom, a ne samo da li dokument pada:
'   ISPRAVKA_ODMAH        pogresan unos, isti fizicki dogadjaj -> storno
'                         stari, otvori novi, zapamti staro->novo, po
'                         snimanju relink + rekalkulacija
'   DUPLI_FANTOM          dokument nikad nije trebalo da postoji -> skini
'                         posledice bez naslednika
'   PONISTENJE_BEZ_ZAMENE fizicki tok se ponistava, nema novog -> blokada
'                         ako postoje zavisni dokumenti (osim uz svesnu
'                         potvrdu)
'   RESI_KASNIJE          persistentan recovery zapis, ne samo poruka
'
' Framework poznaje CETIRI tipa (otpremnica, zbirna, prijemnica, revers).
' Ostalih pet (otkup, isplate, uplate, faktura, izvod) nemaju nizvodni tok
' o kome se odlucuje, pa im je storno obican - kao i u legacy formi, gde
' TryRunCorrectionFramework za njih vraca False.
'
' Izbor moda NIJE ovde: to je pitanje operateru, a ovaj modul nema
' nijednu kontrolu i nijedan MsgBox. Ovde je samo prevod tipa u framework
' i prosledjivanje na modStornoFlow.
'=====================================================================

' Kljuc tipa -> framework docType. Prazno = tip nije framework tip.
Public Function TipUFlowDoc(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTPREMNICA: TipUFlowDoc = FLOW_DOC_OTPREMNICA
        Case STIP_ZBIRNA:     TipUFlowDoc = FLOW_DOC_ZBIRNA
        Case STIP_PRIJEMNICA: TipUFlowDoc = FLOW_DOC_PRIJEMNICA
        Case STIP_REVERSI:    TipUFlowDoc = FLOW_DOC_REVERS
    End Select
End Function

' Da li storno ovog dokumenta trazi da operater bira mod.
'
' SMART TRIGGER iz modStornoFlow: pun izbor SAMO kad postoji nizvodni tok
' (prijemnica, palete, faktura, blokovi) o kome se stvarno odlucuje.
' Inace je obican storno dovoljan - motor cuva invarijantu bez ceremonije.
' Revers je list u lancu, pa nikad ne trazi pun izbor; njegova jedina
' poslovna razlika (storno vs ispravka) resava se kratkim pitanjem u
' ekranu, isto kao u legacy RunReversStornoUI.
' FAIL-CLOSED: greska znaci "ponudi izbor". CorrectionNeedsDialog je i sam
' takav (na gresku vraca True), pa bi ovde "On Error Resume Next" ponistio
' njegovu zastitu - rezultat bi ostao False i storno bi presao u obican, bez
' pitanja o nizvodnom toku. Neizvesnost mora da vodi ka VISE pitanja, ne ka
' manje.
Public Function StornoTraziIzborModa(ByVal tip As String, ByVal broj As String, _
                                     ByVal opcija As String, _
                                     Optional ByVal docID As String = "") As Boolean
    Dim dt As String
    On Error GoTo EH
    dt = TipUFlowDoc(tip)
    If Len(dt) = 0 Then Exit Function      ' nije framework tip - to se zna pouzdano
    StornoTraziIzborModa = CorrectionNeedsDialog(dt, broj, opcija, docID)
    Exit Function
EH:
    LogErr "modStornoDok.StornoTraziIzborModa"
    StornoTraziIzborModa = True
End Function

' Pun pregled lanca za dijalog: sta sve visi o ovom dokumentu. Prazno =
' tip nije framework tip (tada se koristi StornoPregled iz odeljka 2).
Public Function StornoPregledLanca(ByVal tip As String, ByVal broj As String, _
                                   ByVal opcija As String, _
                                   Optional ByVal docID As String = "") As String
    Dim dt As String
    On Error Resume Next
    dt = TipUFlowDoc(tip)
    If Len(dt) = 0 Then Exit Function
    StornoPregledLanca = BuildStornoPreview(dt, broj, opcija, docID)
End Function

' Izvrsi izabrani mod. Vraca recnik iz modStornoFlow:
'   success / blocked / needsForm / correctionID / mode / message
'
' Nothing = tip nije framework tip. PONISTENJE prvo vraca blocked=True sa
' punim spiskom posledica; pozivalac tada pita operatera i, na potvrdu,
' zove ponovo sa forceConfirm=True. Ta dva koraka su namerno razdvojena -
' spisak posledica se pravi PRE nego sto se ista promeni.
Public Function StornoIzvrsiMod(ByVal tip As String, ByVal broj As String, _
                                ByVal opcija As String, ByVal mode As String, _
                                ByVal forceConfirm As Boolean, _
                                ByVal neDiraj As Boolean, _
                                Optional ByVal docID As String = "") As Object
    Dim dt As String
    On Error GoTo EH
    dt = TipUFlowDoc(tip)
    If Len(dt) = 0 Then Exit Function
    Select Case dt
        Case FLOW_DOC_OTPREMNICA: Set StornoIzvrsiMod = RunOtpremnicaCorrection(broj, mode, forceConfirm, docID)
        Case FLOW_DOC_ZBIRNA:     Set StornoIzvrsiMod = RunZbirnaCorrection(broj, mode, forceConfirm, docID)
        Case FLOW_DOC_REVERS:     Set StornoIzvrsiMod = RunReversCorrection(broj, opcija, mode)
        ' neDiraj = "ne diraj palete" (samo prijemnica, DUPLI/PONISTENJE):
        ' palete ostaju vezane za storniranu prijemnicu umesto da se odvezu.
        Case FLOW_DOC_PRIJEMNICA: Set StornoIzvrsiMod = RunPrijemnicaCorrection(broj, mode, forceConfirm, neDiraj, docID)
    End Select
    Exit Function
EH:
    LogErr "modStornoDok.StornoIzvrsiMod"
End Function

' PK stornirane iz correction context-a. Prefill mora da krene od njega, a
' ne od broja: broj dokumenta NIJE globalno jedinstven (GenerateBrojPrijemnice
' racuna sekvencu po kupcu, pa dva kupca istog dana dobiju "1/ddmmyy").
Public Function StorniraniDocID(ByVal correctionID As String) As String
    On Error Resume Next
    If Len(correctionID) = 0 Then Exit Function
    StorniraniDocID = modStornoContext.GetCorrectionField(correctionID, COL_SV_OLD_DOCID)
End Function

' Broj nadredjenog dokumenta iz context-a (za prijemnicu: broj zbirne).
Public Function StorniraniParentBroj(ByVal correctionID As String) As String
    On Error Resume Next
    If Len(correctionID) = 0 Then Exit Function
    StorniraniParentBroj = modStornoContext.GetCorrectionField(correctionID, COL_SV_PARENT_BROJ)
End Function

'=====================================================================
' 5) Z10 - PREFILL IZ STORNIRANOG DOKUMENTA
'
' Posle ISPRAVKE operater ne treba da kuca dokument iznova napamet:
' polja se pune iz upravo storniranog, pa on menja samo gresku. Legacy to
' radi u cetiri forme (PrefillOtkupFromStornirano u modOtkupBlok,
' PrefillOtpremnicaFromStornirana / PrefillZbirnaFromStornirana /
' PrefillPrijemnicaFromStornirana u frmDokumenta) - cetiri kopije istog
' racuna, svaka vezana za svoje kontrole.
'
' Ovde je taj racun IZDVOJEN IZ PRIKAZA: vraca se opis vrednosti u istom
' obliku koji ekran vec razume (modScrDokumenti.PrefillSpec ->
' modOtkupUI.ApplyPrefill), "kljuc=vrednost" spojeno sa "|". Ekran ne zna
' ni jednu kolonu tabele, a ovaj modul ne zna ni jednu kontrolu.
'
' TRI PRAVILA KOJA SE LAKO IZGUBE, a sva tri su iz legacy:
'   1) Polazi se od PK-a stornirane (OldDocID iz correction context-a), NE
'      od broja: broj nije globalno jedinstven (GenerateBrojPrijemnice
'      broji po kupcu, pa dva kupca istog dana dobiju "1/ddmmyy"). Posao
'      radi modDokumenta.PickPrefillRows - ista rutina koju zove i legacy.
'   2) DATUM se preuzima iz stornirane. Ispravka sutradan ne sme da dobije
'      danasnji datum - dokument bi promenio dan.
'   3) BROJ dokumenta se NE preuzima. Ispravka je NOV dokument sa novim
'      brojem; veza na stari cuva se u tblStornoVeza, ne u broju.
'
' BRUTO: kad je OTKUP_BRUTO_UNOS ukljucen, u polje kolicine ide BrutoKg
' (ono sto je operater zaista ukucao), a ne neto. Inace bi ispravka
' oduzela taru drugi put.
'=====================================================================
Public Function PrefillIzStorniranog(ByVal tip As String, ByVal brStorn As String, _
                                     ByVal oldDocID As String) As String
    Dim tbl As String, d As Variant, rI As Long, rII As Long, base As Long
    Dim cBroj As Long, cKlasa As Long, cId As Long, cGen As Long
    Dim cDat As Long, cVr As Long, cSo As Long, cTip As Long, cZbr As Long
    Dim cKol As Long, cCena As Long, cAmb As Long, cBruto As Long
    Dim res As String, brutoMode As Boolean
    On Error GoTo EH

    tbl = TabelaZaPrefill(tip)
    If Len(tbl) = 0 Then Exit Function
    d = GetTableData(tbl)
    If IsEmpty(d) Then Exit Function

    cBroj = GetColumnIndex(tbl, ColBrojZaPrefill(tip))
    cKlasa = GetColumnIndex(tbl, ColKlasaZaPrefill(tip))
    cId = GetColumnIndex(tbl, ColIdZaPrefill(tip))
    cGen = GetColumnIndex(tbl, COL_GENERACIJA_ID)
    If cBroj = 0 Then Exit Function

    modDokumenta.PickPrefillRows d, cBroj, cKlasa, cId, cGen, brStorn, oldDocID, rI, rII
    If rI = 0 And rII = 0 Then Exit Function
    base = rI
    If base = 0 Then base = rII

    cDat = GetColumnIndex(tbl, ColDatumZaPrefill(tip))
    cVr = GetColumnIndex(tbl, ColVrstaZaPrefill(tip))
    cSo = GetColumnIndex(tbl, ColSortaZaPrefill(tip))
    cTip = GetColumnIndex(tbl, ColTipAmbZaPrefill(tip))
    cZbr = GetColumnIndex(tbl, ColBrojZbirneZaPrefill(tip))
    cKol = GetColumnIndex(tbl, ColKolicinaZaPrefill(tip))
    cCena = GetColumnIndex(tbl, ColCenaZaPrefill(tip))
    cAmb = GetColumnIndex(tbl, ColKolAmbZaPrefill(tip))
    cBruto = GetColumnIndex(tbl, ColBrutoZaPrefill(tip))
    brutoMode = OtkupBrutoUnos()

    ' Datum PRVI: predlog broja dokumenta zavisi od njega.
    If cDat > 0 Then
        If IsDate(d(base, cDat)) Then _
            res = Spoji(res, "datum", Format$(CDate(d(base, cDat)), "dd.mm.yyyy"))
    End If
    res = Spoji(res, "vrsta", NzToText(d(base, cVr)))
    res = Spoji(res, "sorta", NzToText(d(base, cSo)))
    res = Spoji(res, "tipamb", NzToText(d(base, cTip)))
    ' Zbirna: kod same zbirne broj dokumenta JESTE broj zbirne, pa se ne
    ' preuzima (pravilo 3) - novi dokument dobija nov broj.
    If tip <> STIP_ZBIRNA Then res = Spoji(res, "brzbirne", NzToText(d(base, cZbr)))

    res = Spoji(res, "omid", CelijaAko(d, base, GetColumnIndex(tbl, ColStanicaZaPrefill(tip))))
    res = Spoji(res, "vozacid", CelijaAko(d, base, GetColumnIndex(tbl, ColVozacZaPrefill(tip))))
    res = Spoji(res, "partnerid", CelijaAko(d, base, GetColumnIndex(tbl, ColPartnerZaPrefill(tip))))
    res = Spoji(res, "parcela", CelijaAko(d, base, GetColumnIndex(tbl, ColParcelaZaPrefill(tip))))

    res = Spoji(res, "dveklase", IIf(rII > 0, "2", "1"))
    If rI > 0 Then
        res = Spoji(res, "kol1", BrojUTekst(KolicinaReda(d, rI, cKol, cBruto, brutoMode)))
        res = Spoji(res, "amb1", BrojUTekst(CeliBroj(d, rI, cAmb)))
        res = Spoji(res, "cena", BrojUTekst(CeliBrojD(d, rI, cCena)))
        res = Spoji(res, "ambpr", BrojUTekst(CeliBroj(d, rI, GetColumnIndex(tbl, ColAmbPrZaPrefill(tip)))))
    End If
    If rII > 0 Then
        res = Spoji(res, "kol2", BrojUTekst(KolicinaReda(d, rII, cKol, cBruto, brutoMode)))
        res = Spoji(res, "amb2", BrojUTekst(CeliBroj(d, rII, cAmb)))
        res = Spoji(res, "cena2", BrojUTekst(CeliBrojD(d, rII, cCena)))
    End If

    ' Partner je popunjen, pa fokus ide na kolicinu - to je polje koje se
    ' pri ispravci najcesce menja. (Kod prefilla sa otpremnice u F1 partner
    ' NIJE popunjen, pa tamo fokus ostaje na njemu.)
    res = Spoji(res, "fokus", "kolicina")
    PrefillIzStorniranog = res
    Exit Function
EH:
    LogErr "modStornoDok.PrefillIzStorniranog"
End Function

'--------------------------------------------------- mape kolona po tipu
' Ime kolone se NE pise kao literal ni kad je u vecini tabela isto. Zbirna
' je zivi razlog: njena kolicina je "UkupnoKolicina", a ambalaza
' "UkupnoAmbalaze" - literal "Kolicina" bi za nju tiho vratio nulu, pa bi
' prefill ispravke zbirne dosao prazan i nista ne bi prijavilo gresku.
' Sema tabela je izvor istine, ne kod (.claude/rules/podaci-i-config.md).
'
' Prazan povratak znaci "tip nema taj pojam" - GetColumnIndex tada vraca 0,
' a citaci ispod preskacu vrednost. Tako zbirna nema cenu ni bruto,
' otpremnica nema partnera, a samo otkup ima parcelu.
Private Function TabelaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      TabelaZaPrefill = TBL_OTKUP
        Case STIP_OTPREMNICA: TabelaZaPrefill = TBL_OTPREMNICA
        Case STIP_ZBIRNA:     TabelaZaPrefill = TBL_ZBIRNA
        Case STIP_PRIJEMNICA: TabelaZaPrefill = TBL_PRIJEMNICA
    End Select
End Function

Private Function ColBrojZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColBrojZaPrefill = COL_OTK_BR_DOK
        Case STIP_OTPREMNICA: ColBrojZaPrefill = COL_OTP_BROJ
        Case STIP_ZBIRNA:     ColBrojZaPrefill = COL_ZBR_BROJ
        Case STIP_PRIJEMNICA: ColBrojZaPrefill = COL_PRJ_BROJ
    End Select
End Function

Private Function ColIdZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColIdZaPrefill = COL_OTK_ID
        Case STIP_OTPREMNICA: ColIdZaPrefill = COL_OTP_ID
        Case STIP_ZBIRNA:     ColIdZaPrefill = COL_ZBR_ID
        Case STIP_PRIJEMNICA: ColIdZaPrefill = COL_PRJ_ID
    End Select
End Function

Private Function ColDatumZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColDatumZaPrefill = COL_OTK_DATUM
        Case STIP_OTPREMNICA: ColDatumZaPrefill = COL_OTP_DATUM
        Case STIP_ZBIRNA:     ColDatumZaPrefill = COL_ZBR_DATUM
        Case STIP_PRIJEMNICA: ColDatumZaPrefill = COL_PRJ_DATUM
    End Select
End Function

' Otkupno mesto postoji samo na otkupu i otpremnici; zbirna i prijemnica
' ga nemaju (njihov partner je kupac).
Private Function ColStanicaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColStanicaZaPrefill = COL_OTK_STANICA
        Case STIP_OTPREMNICA: ColStanicaZaPrefill = COL_OTP_STANICA
    End Select
End Function

' Partner novog UI-ja (polje cbKupac): kooperant na otkupu, kupac na
' zbirnoj i prijemnici. Otpremnica partnera nema - njen "kupac" je sama
' stanica, koja vec ide kroz omid.
Private Function ColPartnerZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColPartnerZaPrefill = COL_OTK_KOOPERANT
        Case STIP_ZBIRNA:     ColPartnerZaPrefill = COL_ZBR_KUPAC
        Case STIP_PRIJEMNICA: ColPartnerZaPrefill = COL_PRJ_KUPAC
    End Select
End Function

Private Function ColParcelaZaPrefill(ByVal tip As String) As String
    If tip = STIP_OTKUP Then ColParcelaZaPrefill = COL_OTK_PARCELA
End Function

Private Function ColVozacZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColVozacZaPrefill = COL_OTK_VOZAC
        Case STIP_OTPREMNICA: ColVozacZaPrefill = COL_OTP_VOZAC
        Case STIP_ZBIRNA:     ColVozacZaPrefill = COL_ZBR_VOZAC
        Case STIP_PRIJEMNICA: ColVozacZaPrefill = COL_PRJ_VOZAC
    End Select
End Function

Private Function ColKlasaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColKlasaZaPrefill = COL_OTK_KLASA
        Case STIP_OTPREMNICA: ColKlasaZaPrefill = COL_OTP_KLASA
        Case STIP_ZBIRNA:     ColKlasaZaPrefill = COL_ZBR_KLASA
        Case STIP_PRIJEMNICA: ColKlasaZaPrefill = COL_PRJ_KLASA
    End Select
End Function

Private Function ColVrstaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColVrstaZaPrefill = COL_OTK_VRSTA
        Case STIP_OTPREMNICA: ColVrstaZaPrefill = COL_OTP_VRSTA
        Case STIP_ZBIRNA:     ColVrstaZaPrefill = COL_ZBR_VRSTA
        Case STIP_PRIJEMNICA: ColVrstaZaPrefill = COL_PRJ_VRSTA
    End Select
End Function

Private Function ColSortaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColSortaZaPrefill = COL_OTK_SORTA
        Case STIP_OTPREMNICA: ColSortaZaPrefill = COL_OTP_SORTA
        Case STIP_ZBIRNA:     ColSortaZaPrefill = COL_ZBR_SORTA
        Case STIP_PRIJEMNICA: ColSortaZaPrefill = COL_PRJ_SORTA
    End Select
End Function

Private Function ColTipAmbZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColTipAmbZaPrefill = COL_OTK_TIP_AMB
        Case STIP_OTPREMNICA: ColTipAmbZaPrefill = COL_OTP_TIP_AMB
        Case STIP_ZBIRNA:     ColTipAmbZaPrefill = COL_ZBR_TIP_AMB
        Case STIP_PRIJEMNICA: ColTipAmbZaPrefill = COL_PRJ_TIP_AMB
    End Select
End Function

' Zbirna nema BrojZbirne - ona JESTE zbirna, pa se u telu i preskace.
Private Function ColBrojZbirneZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColBrojZbirneZaPrefill = COL_OTK_BROJ_ZBIRNE
        Case STIP_OTPREMNICA: ColBrojZbirneZaPrefill = COL_OTP_BROJ_ZBIRNE
        Case STIP_PRIJEMNICA: ColBrojZbirneZaPrefill = COL_PRJ_BROJ_ZBIRNE
    End Select
End Function

' PAZNJA: zbirna ne zove ovo "Kolicina" nego "UkupnoKolicina".
Private Function ColKolicinaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColKolicinaZaPrefill = COL_OTK_KOLICINA
        Case STIP_OTPREMNICA: ColKolicinaZaPrefill = COL_OTP_KOLICINA
        Case STIP_ZBIRNA:     ColKolicinaZaPrefill = COL_ZBR_KOLICINA
        Case STIP_PRIJEMNICA: ColKolicinaZaPrefill = COL_PRJ_KOLICINA
    End Select
End Function

' PAZNJA: zbirna ne zove ovo "KolAmbalaze" nego "UkupnoAmbalaze".
Private Function ColKolAmbZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColKolAmbZaPrefill = COL_OTK_KOL_AMB
        Case STIP_OTPREMNICA: ColKolAmbZaPrefill = COL_OTP_KOL_AMB
        Case STIP_ZBIRNA:     ColKolAmbZaPrefill = COL_ZBR_KOL_AMB
        Case STIP_PRIJEMNICA: ColKolAmbZaPrefill = COL_PRJ_KOL_AMB
    End Select
End Function

' Zbirna NEMA cenu (tblZbirna nema tu kolonu) - zbirna je zbir vec
' obracunatih otpremnica.
Private Function ColCenaZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColCenaZaPrefill = COL_OTK_CENA
        Case STIP_OTPREMNICA: ColCenaZaPrefill = COL_OTP_CENA
        Case STIP_PRIJEMNICA: ColCenaZaPrefill = COL_PRJ_CENA
    End Select
End Function

' Zbirna NEMA BrutoKg - ona zbraja vec netirane otpremnice.
Private Function ColBrutoZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColBrutoZaPrefill = COL_OTK_BRUTO
        Case STIP_OTPREMNICA: ColBrutoZaPrefill = COL_OTP_BRUTO
        Case STIP_PRIJEMNICA: ColBrutoZaPrefill = COL_PRJ_BRUTO
    End Select
End Function

' Prazna ambalaza: uz otkup se IZDAJE, uz prijemnicu se VRACA. Otpremnica
' i zbirna je ne dodiruju - isto pravilo koje drzi modOtkupUI.ApplyFormFields.
Private Function ColAmbPrZaPrefill(ByVal tip As String) As String
    Select Case tip
        Case STIP_OTKUP:      ColAmbPrZaPrefill = COL_OTK_KOL_AMB_IZDATA
        Case STIP_PRIJEMNICA: ColAmbPrZaPrefill = COL_PRJ_KOL_AMB_VRACENA
    End Select
End Function

'------------------------------------------------------ citanje celija
Private Function KolicinaReda(ByVal d As Variant, ByVal r As Long, _
                              ByVal cKol As Long, ByVal cBruto As Long, _
                              ByVal brutoMode As Boolean) As Double
    If brutoMode And cBruto > 0 Then
        If NzD(d(r, cBruto)) > 0 Then
            KolicinaReda = NzD(d(r, cBruto))
            Exit Function
        End If
    End If
    If cKol > 0 Then KolicinaReda = NzD(d(r, cKol))
End Function

Private Function CeliBroj(ByVal d As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c > 0 Then CeliBroj = NzL(d(r, c))
End Function

Private Function CeliBrojD(ByVal d As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c > 0 Then CeliBrojD = NzD(d(r, c))
End Function

Private Function CelijaAko(ByVal d As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then CelijaAko = Trim$(NzToText(d(r, c)))
End Function

' Nula se ne salje: prazno polje i "0" su dva razlicita stanja, a prefill
' ne sme da ukuca nulu tamo gde operater nije nista uneo.
Private Function BrojUTekst(ByVal v As Double) As String
    If v <> 0 Then BrojUTekst = CStr(v)
End Function

Private Function Spoji(ByVal res As String, ByVal k As String, ByVal v As String) As String
    Spoji = res
    If Len(v) = 0 Then Exit Function
    Spoji = res & IIf(Len(res) > 0, "|", "") & k & "=" & v
End Function

'=====================================================================
' POMOCNO
'=====================================================================

' Ime tipa za poruke. Za revers zavisi od SMERA - cetiri smera su cetiri
' razlicita dokumenta iako dele brojevni niz.
Public Function TipNaziv(ByVal tip As String, ByVal opcija As String) As String
    Select Case tip
        Case STIP_OTKUP:      TipNaziv = Poruka("STORNO_TIP_OTKUP")
        Case STIP_OTPREMNICA: TipNaziv = Poruka("STORNO_TIP_OTPREMNICA")
        Case STIP_ZBIRNA:     TipNaziv = Poruka("STORNO_TIP_ZBIRNA")
        Case STIP_PRIJEMNICA: TipNaziv = Poruka("STORNO_TIP_PRIJEMNICA")
        Case STIP_ISPLATE:    TipNaziv = Poruka("STORNO_TIP_ISPLATA")
        Case STIP_UPLATE:     TipNaziv = Poruka("STORNO_TIP_UPLATA")
        Case STIP_FAKTURA:    TipNaziv = Poruka("STORNO_TIP_FAKTURA")
        Case STIP_IZVOD:      TipNaziv = Poruka("STORNO_TIP_IZVOD")
        Case STIP_REVERSI:    TipNaziv = ReversNaziv(opcija)
        Case Else:            TipNaziv = tip
    End Select
End Function

Private Function ReversNaziv(ByVal dokTip As String) As String
    Select Case dokTip
        Case DOK_TIP_OM_IZLAZ_KOOP:  ReversNaziv = Poruka("STORNO_TIP_REV_IZD_KOOP")
        Case DOK_TIP_OM_ULAZ_KOOP:   ReversNaziv = Poruka("STORNO_TIP_REV_PRI_KOOP")
        Case DOK_TIP_OM_ULAZ_FIRMA:  ReversNaziv = Poruka("STORNO_TIP_REV_IZD_OM")
        Case DOK_TIP_OM_IZLAZ_FIRMA: ReversNaziv = Poruka("STORNO_TIP_REV_PRI_OM")
        Case Else:                   ReversNaziv = Poruka("STORNO_TIP_REVERS")
    End Select
End Function

Private Function NijePronadjen(ByVal broj As String) As String
    NijePronadjen = Poruka("STORNO_ERR_NEMA_DOK") & " " & broj
End Function
