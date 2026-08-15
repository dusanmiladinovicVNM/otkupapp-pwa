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
' STA OVAJ MODUL NAMERNO NE RADI: ne nudi ispravku ni dupli unos posle
' storna (modStornoFlow.TryRunCorrectionFramework, Z10 iz kataloga). To
' je stavka 13 Faze D i ide zasebno; dotle je storno iz novog UI-ja
' OBICAN storno, isti onaj koji legacy radi kad framework ne preuzme tip.
'
' VAZNO: legacy frmDokumenta OSTAJE netaknut i potpuno operativan. Ovaj
' modul je drugi pozivalac istih poslovnih rutina, ne zamena za formu.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const STORNODOK_BUILD As String = "v6-ui-119"

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
Public Function StornoRazlog(ByVal tip As String, ByVal broj As String, _
                             ByVal opcija As String) As String
    Dim razlog As String, izvBroj As String, izvRacun As String
    On Error GoTo EH
    broj = Trim$(broj)
    If Len(broj) = 0 Then
        StornoRazlog = Poruka("STORNO_ERR_NEMA_BROJA")
        Exit Function
    End If

    Select Case tip
        Case STIP_OTKUP
            If Len(LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, broj, COL_OTK_ID)) = 0 Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_OTPREMNICA
            If Len(LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_ID)) = 0 Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_ZBIRNA
            ' StornoZbirna_TX prima BROJ (ne ID) i sam razresava; provera
            ' postojanja je ista kao za ostale robne dokumente.
            If Len(LookupActiveID(TBL_ZBIRNA, COL_ZBR_BROJ, broj, COL_ZBR_BROJ)) = 0 Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_PRIJEMNICA
            If Len(LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, broj, COL_PRJ_BROJ)) = 0 Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_FAKTURA
            If Len(LookupActiveID(TBL_FAKTURE, COL_FAK_BROJ, broj, COL_FAK_ID)) = 0 Then _
                StornoRazlog = NijePronadjen(broj)

        Case STIP_ISPLATE, STIP_UPLATE
            ' Ceo razlog dolazi iz modStorno: izvod se ne stornira
            ' parcijalno, a broj sa vise aktivnih redova (avans raspodela
            ' deli isti broj) trazi NovacID umesto tihog storna jednog reda.
            ResolveNovacForStorno broj, razlog
            StornoRazlog = razlog

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
    LogErr "modStornoDok.StornoRazlog"
    StornoRazlog = Poruka("STORNO_ERR_RAZRESENJE") & " " & Err.description
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
                             ByVal opcija As String, ByRef poruka As String) As Boolean
    Dim ok As Boolean, novID As String, razlog As String
    Dim izvBroj As String, izvRacun As String, izvInfo As String
    Dim fakID As String, vezPrij As String
    On Error GoTo EH
    poruka = ""
    broj = Trim$(broj)

    Select Case tip
        Case STIP_OTKUP
            ' Klasa I i II dele isti BrDok (zaseban red po klasi) -> stornira
            ' se ceo dokument, ne jedan red. Isto sto radi F1 lista.
            ok = StornoOtkupByBrDok_TX(broj)

        Case STIP_OTPREMNICA
            ' i ovde klase dele broj
            ok = StornoOtpremnicaByBroj_TX(broj)

        Case STIP_ZBIRNA
            ok = StornoZbirna_TX(broj)
            If ok Then
                vezPrij = NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, broj, COL_PRJ_BROJ))
                If Len(vezPrij) > 0 Then poruka = Poruka("STORNO_MSG_ZBIRNA_PRIJ") & " " & vezPrij
            End If

        Case STIP_PRIJEMNICA
            ok = StornoPrijemnicaByBroj_TX(broj)

        Case STIP_FAKTURA
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
            novID = ResolveNovacForStorno(broj, razlog)
            If Len(razlog) > 0 Then
                poruka = razlog
                Exit Function
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
                poruka = Poruka("STORNO_ERR_NEMA_ISHODA")
                Exit Function
            End If
            ok = StornoIzvod_TX(izvBroj, izvRacun, opcija, izvInfo)
            If ok Then poruka = izvInfo

        Case Else
            poruka = Poruka("STORNO_ERR_NEPOZNAT_TIP") & " " & tip
            Exit Function
    End Select

    StornoIzvrsi = ok
    If ok And Len(poruka) = 0 Then poruka = Poruka("STORNO_MSG_OK")
    If Not ok And Len(poruka) = 0 Then poruka = Poruka("STORNO_ERR_NEUSPEH") & " " & broj
    Exit Function
EH:
    LogErr "modStornoDok.StornoIzvrsi"
    poruka = Poruka("STORNO_ERR_NEUSPEH") & " " & broj & ": " & Err.description
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
