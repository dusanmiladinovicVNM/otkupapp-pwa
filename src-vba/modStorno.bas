Attribute VB_Name = "modStorno"
'Attribute VB_Name = "modStorno"
Option Explicit

' ============================================================
' modStorno v4.0 - Hardened Soft-Delete
'
' Stil uskladjen sa modNovac/modFaktura:
' - fail-fast schema guards
' - RequireColumnIndex / RequireUpdateCell
' - stroga provera single-row dokumenata
' - transakcioni rollback u *_TX wrapperima
' - monitoring success/fail eventa
' - bez MsgBox u business sloju
'
' Business pravila:
' - Svaki dokument se stornira pojedinacno.
' - Nema automatske kaskade izmedju dokumenata.
' - Ambalaza se stornira za dokument gde postoji.
' - Faktura: stavke se storniraju, novac se odvezuje od fakture. Prijemnica
'   se time oslobadja SAMA: njena fakturisanost je izvedena iz aktivnih
'   linija (S6 korak 5, FAK-LIN-03), pa se nista na njoj ne pise.
' - Prijemnica: fakturisana (i delimicno) se ODBIJA -- prvo storno fakture
'   (S6 korak 4, odluka 1).
' ============================================================

Private Const MOD_NAME As String = "modStorno"
Private Const STORNO_DA As String = "Da"
Private Const STATUS_STORNIRANO As String = "Stornirano"
Private Const ERR_STORNO_BASE As Long = vbObjectError + 2400

' TEST SEAM (R3 revizije #248, rollback dokaz): namerna greska POSLE
' cross-table upisa u StornoFaktura -- jedini nacin da test dokaze da
' TX snapshot pokriva SVE tabele koje storno pise (tblPrerada je GP
' release). Tvrdo gejtovano: setter radi samo u test rezimu, flag je
' u produkciji uvek False.
Private mTestFailPosleRelease As Boolean

' TEST SEAM (S6 korak 4, R2/R4): namerna greska na KRAJU jezgra storna
' prijemnice, posle SVIH upisa (palete, zaglavlje, knjiga) -- jedini nacin da test
' dokaze da rollback pozivaoca vraca ceo dokument. Postavljac ga pali samo u test
' rezimu, kuka ponovo pita IsTestMode(), a gasi ga modTest.ResetSeamova.
Private mTestPrjStornoPadni As Boolean

' ============================================================
' OTKUP
' ============================================================

' KANONSKI ULAZ STORNA OTKUPA: po OtkupID-u (S1e). Otkup je jedno zaglavlje po
' dokumentu, pa nema "svih redova broja" koje treba skupiti -- ID iz reda koji je
' operater izabrao putuje do mutacije bez ponovnog trazenja po broju.
'
' Autohladnjaca: lanac (otpremnica + zbirna + prijemnica) je auto-generisan iz
' ovog otkupa i deli njegov BrojZbirne, pa storno otkupa povlaci i njih. Stanica
' i BrojZbirne se citaju sa ISTOG zaglavlja (po ID-u). Gejt je stanica-hladnjaca
' (struktura), ne toggle: bez lanca je kaskada no-op.
Public Function StornoOtkup_TX(ByVal otkupID As String) As Boolean
    Const SRC As String = "StornoOtkup_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    RequireNonBlank otkupID, "OtkupID", SRC

    ' HLADNJACKA KASKADA JE OBRISANA, JER JE NEDOSTIZNA (S5-3b).
    '
    ' Grana je kaskadirala sa storna BLOKA na otpremnicu, zbirnu i prijemnicu, a
    ' okidala se uslovom "blok ima zbirnu". Uslov je dugo bio mrtav zato sto ga
    ' je hranila kolona Otkup.BrojZbirne, koju od S5-3 niko ne pise -- pa je prvi
    ' potez ovog reza bio da se prehrani iz clanstva.
    '
    ' Mereno posle toga: grana se NE MOZE dosegnuti ni sa tacnim podatkom.
    ' StornoOtkup (jezgro, review #362) odbija blok koji je u sastavu AKTIVNE
    ' otpremnice -- A13/A15. A blok ima zbirnu samo preko svoje otpremnice:
    '
    '     zbirnaID = AktivnaZbirnaZaOtpremnicu(OtpremnicaZaOtkup(otkupID))
    '
    ' Neprazan zbirnaID dakle ZAHTEVA neprazan OtpremnicaZaOtkup, a bas to
    ' jezgro odbija jedan red iznad. Uslov kaskade i uslov prolaska se iskljucuju
    ' po konstrukciji: kaskada bi se izvrsila tacno onda kad se do nje ne stize.
    '
    ' Sposobnost nije izgubljena nego ZAMENJENA: lanac se rusi sa nivoa
    ' otpremnice (ISPRAVKA / PONISTENJE), koji kaskadira nadole. To je isto
    ' pravilo koje ADR-0001 trazi -- izdat dokument se ne menja u mestu.

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_STORNO_ZURNAL    ' zurnal upisi teku u istoj TX -> rollback ih povlaci

    ' StornoOtkup otvara i zatvara SVOJ zurnal op, pa kaskada ispod NE ulazi u
    ' operaciju otkupa.
    If Not StornoOtkup(otkupID, tx) Then
        Err.Raise ERR_STORNO_BASE + 1, SRC, _
                  "StornoOtkup nije uspeo. OtkupID=" & otkupID
    End If

    tx.CommitTx

    StornoOtkup_TX = True
    MonitorStornoSuccess SRC, "Otkup", otkupID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Otkup", otkupID, tx
    StornoOtkup_TX = False
End Function

' tx je OBAVEZAN, ne opcion: storno knjige je kontra-stav koji mora da legne u
' ISTU transakciju (AMB-INV-08). Opcion tx bio bi fail-open seam -- pozivalac
' koji ga zaboravi tiho preskoci kontra-stavove, a zastavica bi ostavila zeleno.
'
' odloziObavezu: SAMO ispravka otkupa (modAmbalaza.StornirajAmbalazuDokumenta) --
' AMB-INV-09 tada meri pozivalac, posle zamene.
Public Function StornoOtkup(ByVal otkupID As String, _
                            ByVal tx As clsTransaction, _
                            Optional ByVal odloziObavezu As Boolean = False) As Boolean
    Const SRC As String = "StornoOtkup"
    Dim owns As Boolean

    On Error GoTo EH

    ' IZVOR AKTIVNE OTPREMNICE SE NE STORNIRA (review #362, A13/A15) -- ista
    ' kapija koju StornoOtpremnica drzi jedan nivo vise, prema zbirnoj.
    '
    ' Clan otpremnice (tblOtpremnicaIzvori) je osnova njenog sastava, i to u OBA
    ' stanja roditelja:
    '   DRAFT  -- storno bi ostavio clanstvo nacrta koje pokazuje na storniran
    '             otkup (izdavanje ga tek kasnije odbija, a do tada nacrt laze);
    '   IZDATO -- izdat dokument bi ostao zasnovan na storniranom izvoru.
    ' Isto pravilo drzi IspravkaOtkupa_TX (modOtkup) za ispravku.
    '
    ' Kapija je u JEZGRU, a ne samo u F1 i F8: jezgro zovu i StornoOtkup_TX i
    ' ispravka otkupa, pa bi kapija na jednom ulazu ostavila ostale otvorene.
    Dim otpID As String
    otpID = modDokumenta.OtpremnicaZaOtkup(otkupID)
    If Len(otpID) > 0 Then
        Err.Raise ERR_STORNO_BASE + 72, SRC, _
                  "Otkup " & otkupID & " je u sastavu aktivne otpremnice " & otpID & _
                  ". Storno izvora nije dozvoljen dok je otkup u njenom sastavu."
    End If

    Dim rowOtkup As Long
    rowOtkup = RequireStornoAllowed(TBL_OTKUP, otkupID, COL_OTK_ID, SRC)

    ' Zurnal: otvori op po broju bloka (lossless undo). Journal stare vrednosti PRE
    ' mutacije; primitiv ResetNovacOtkupLink usput belezen.
    Dim brDok As String: brDok = NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_BR_DOK))
    owns = BeginStornoOp(DOK_TIP_OTKUP, brDok)
    JournalCell TBL_OTKUP, otkupID, COL_STORNIRANO, _
        NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO)), "Da"

    MarkRowStornirano TBL_OTKUP, rowOtkup, SRC

    ' AMB-INV-08: dokument se vezuje TEK POSLE sto je OVA transakcija stvarno
    ' promenila zaglavlje. Vezivanje iz ambalaznog primitiva bilo bi samopotvrda --
    ' ledger-storno nije pisac izvornog dokumenta, pa je mogao da anulira efekat
    ' AKTIVNOG otkupa i prodje sve kapije (review 03.10.2026, P1 #1). Zato je
    ' modStorno.StornoOtkup na AMB_BIND_DOZVOLJENI, a primitiv nije.
    tx.BindSourceDocument DOK_TIP_OTKUP, otkupID

    ' Storno ambalaze je KONTRA-STAV u knjizi (AMB-10-ODL-16, 6.12d): knjiga je
    ' nepromenljiva, a citalac vidi red po Od_*/Na_*/VrstaKretanja/StornoOd.
    ' Zastavica Stornirano na redu knjige je bila mehanizam STAROG modela i
    ' obrisana je u 10e-2 -- nov citalac je nikad nije gledao.
    modAmbalaza.StornirajAmbalazuDokumenta tx, DOK_TIP_OTKUP, otkupID, odloziObavezu
    ResetNovacOtkupLink otkupID

    EndStornoOp owns
    StornoOtkup = True
    Exit Function

EH:
    EndStornoOp owns
    LogAndReraise SRC
End Function

' ============================================================
' OTPREMNICA
' ============================================================

Public Function StornoOtpremnica_TX(ByVal otpremnicaID As String) As Boolean
    Const SRC As String = "StornoOtpremnica_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    If Not StornoOtpremnica(otpremnicaID, tx) Then
        Err.Raise ERR_STORNO_BASE + 2, SRC, _
                  "StornoOtpremnica nije uspeo. OtpremnicaID=" & otpremnicaID
    End If

    tx.CommitTx

    StornoOtpremnica_TX = True
    MonitorStornoSuccess SRC, "Otpremnica", otpremnicaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Otpremnica", otpremnicaID, tx
    StornoOtpremnica_TX = False
End Function

' tx je OBAVEZAN -- isti razlog kao kod StornoOtkup: storno knjige je
' kontra-stav koji mora da legne u ISTU transakciju (AMB-INV-08). Sva tri
' pozivna mesta (StornoOtpremnica_TX, OtpIspravi, PonistiZbirnaChain_TX) tx vec
' imaju, uz snapshot TBL_OTPREMNICA i TBL_AMBALAZA.
Public Function StornoOtpremnica(ByVal otpremnicaID As String, _
                                 ByVal tx As clsTransaction) As Boolean
    Const SRC As String = "StornoOtpremnica"

    On Error GoTo EH

    ' IZVOR AKTIVNE ZBIRNE SE NE STORNIRA (review #362, A13/A15).
    '
    ' Clan kanonske zbirne (tblZbirnaIzvori) je osnova izvedenog dokumenta:
    ' storno ispod nje bi ostavio aktivnu zbirnu zasnovanu na storniranom izvoru,
    ' sa clanstvom koje i dalje pokazuje na njega. Sta storno tada znaci -- zamenu
    ' zbirne, novu verziju, kaskadu -- odlucuje S4; do tada se odbija.
    '
    ' Kapija je u JEZGRU, a ne samo u F8 i u StornoOtpremnica_TX: jezgro zovu i
    ' put po broju i kaskade starog okvira, pa bi kapija samo na jednom ulazu
    ' ostavila ostale otvorene. Stari lanac (veza BrojZbirne) je ne dotice --
    ' clanstvo se cita iskljucivo iz kanona.
    Dim zbrID As String
    zbrID = modDokumenta.AktivnaZbirnaZaOtpremnicu(otpremnicaID)
    If Len(zbrID) > 0 Then
        Err.Raise ERR_STORNO_BASE + 71, SRC, _
                  "Otpremnica " & otpremnicaID & " je izvor aktivne zbirne " & zbrID & _
                  ". Storno izvora je pauziran do S4 (zbirna na novom modelu)."
    End If

    Dim rowOtp As Long
    rowOtp = RequireStornoAllowed(TBL_OTPREMNICA, otpremnicaID, COL_OTP_ID, SRC)

    MarkRowStornirano TBL_OTPREMNICA, rowOtp, SRC

    ' AMB-INV-08: dokument se vezuje POSLE sto ga je ova transakcija promenila.
    ' Vezivanje iz ambalaznog primitiva bilo bi samopotvrda (v. AMB-10-ODL-15),
    ' pa je modStorno.StornoOtpremnica na AMB_BIND_DOZVOLJENI.
    tx.BindSourceDocument DOK_TIP_OTPREMNICA, otpremnicaID

    ' Storno ambalaze je kontra-stav u knjizi (AMB-10-ODL-16), kao kod otkupa.
    modAmbalaza.StornirajAmbalazuDokumenta tx, DOK_TIP_OTPREMNICA, otpremnicaID

    StornoOtpremnica = True
    Exit Function

EH:
    LogAndReraise SRC
End Function
' generacijaID: identitet izabrane zbirne. Sa njim se storniraju SAMO redovi te
' generacije.
'
' Napomena: broj zbirne generator drzi jedinstvenim (SuggestNextBroj za ZBR
' bumpuje dok BrojZbirneExists ne kaze da je slobodan), pa je generacija ovde
' pojas za RUCNI UNOS -- ne za redovan tok.
'
' OGRANICENJE SEME: otpremnice, prijemnice i paletne stavke vezuju zbirnu
' KOLONOM BrojZbirne -- ZbirnaID im nije strani kljuc nigde.
'
' Ovo NIJE tvrdnja da podataka nema: kaskade vec umeju BrojZbirne + VozacID
' (otpremnica) i + KupacID (prijemnica), a palete nose PrijemnicaID. Znaci
' scope se moze izvesti -- samo child mutacije jos nisu sve dovedene dotle.
' Do tada je fail-closed bezbedan izbor, ne dokaz nemogucnosti. Kad broj nose dve
' aktivne zbirne, kaskada bi dirala i tudju decu, pa te putanje staju
' (RequireJedanVlasnikPoBroju). Sam storno zaglavlja je i tada tacan.
Public Function StornoZbirna_TX(ByVal zbirnaID As String) As Boolean
    Const SRC As String = "StornoZbirna_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    If Not StornoZbirna(zbirnaID) Then
        Err.Raise ERR_STORNO_BASE + 3, SRC, _
                  "StornoZbirna nije uspeo. ZbirnaID=" & zbirnaID
    End If

    tx.CommitTx

    StornoZbirna_TX = True
    MonitorStornoSuccess SRC, "Zbirna", zbirnaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Zbirna", zbirnaID, tx
    StornoZbirna_TX = False
End Function

' STORNO ZBIRNE IDE PO ZbirnaID-u (S4-2).
'
' Do ovog koraka je jezgro primalo (BrojZbirne, GeneracijaID): broj nije
' identitet -- isti broj sme da nose dva vozaca -- pa je generacija bila drugi,
' paralelni identitet koji je razlikovao dokumente pod istim brojem.
'
' Kanonski pisac (CreateZbirna_TX) generaciju NE PISE: identitet kanonske zbirne
' je ZbirnaID. Da je F3 odmrznut nad starim potpisom, operater bi napravio
' dokument bez generacije, ljuska mu ne bi nasla stabilan kljuc i radnje bi pale
' nazad na broj -- dakle na "poslovni broj = identitet", tacno ono sto ceo
' refaktor uklanja. Isti rez je S1e uradio za otkup, a #362 za otpremnicu.
'
' Stari okvir ispravke (modStornoFlow) jos radi po (broj, generacija) i sam
' prevodi u ID pre poziva; on se brise u S4-3.
Public Function StornoZbirna(ByVal zbirnaID As String) As Boolean
    Const SRC As String = "StornoZbirna"

    On Error GoTo EH

    RequireNonBlank zbirnaID, "ZbirnaID", SRC

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_BASE + 20, SRC, _
                  "Tabela je prazna: " & TBL_ZBIRNA
    End If

    Dim colId As Long
    Dim colStorno As Long

    colId = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, SRC)
    colStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, SRC)

    ' Nema vise kapije nad vlasnicima broja: ID bira TACNO jedan red, pa
    ' dvosmislenost broja ovde ne postoji. Kapija koja je to cuvala
    ' (RequireJedanVlasnikPoBroju) stitila je od izbora po broju -- izbora kog
    ' vise nema.
    Dim kljuc As String
    kljuc = UCase$(Trim$(zbirnaID))

    Dim foundAny As Boolean
    Dim changedCount As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(NzToText(data(i, colId))), kljuc, vbTextCompare) = 0 Then
            If foundAny Then
                Err.Raise ERR_STORNO_BASE + 23, SRC, _
                          "Dva reda nose isti ZbirnaID: " & zbirnaID & _
                          ". ZbirnaID je identitet dokumenta."
            End If
            foundAny = True

            If Not IsStorniranoValue(data(i, colStorno)) Then
                MarkRowStornirano TBL_ZBIRNA, i, SRC
                changedCount = changedCount + 1
            End If
        End If
    Next i

    If Not foundAny Then
        Err.Raise ERR_STORNO_BASE + 21, SRC, _
                  "Zbirna nije pronadjena. ZbirnaID=" & zbirnaID
    End If

    If changedCount = 0 Then
        Err.Raise ERR_STORNO_BASE + 22, SRC, _
                  "Zbirna je ve" & ChrW(263) & " stornirana. ZbirnaID=" & zbirnaID
    End If

    StornoZbirna = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PRIJEMNICA
' ============================================================

' STORNO PRIJEMNICE IDE PO PrijemnicaID-u (S6 korak 4, odluka 2) -- jedini ulaz.
'
' Od koraka 2 je dokument JEDAN red zaglavlja sa 1..2 stavke, pa "svih redova
' broja" vise nema; StornoPrijemnicaByBroj_TX je obrisan. Broj nose dva kupca i
' dva dokumenta istog kupca, a sa generacijom je stari ulaz broj ignorisao.
'
' skiniPalete = False samo kad operater to IZRICITO bira: ISPRAVKA (palete cekaju
' novu prijemnicu) i DUPLI "ne diraj palete" (MANUAL recovery). Neuspeh nikad ne
' vodi u tu granu -- tada pada ceo storno (R3).
'
' outRazlog: ZASTO je storno odbijen, citljivo operateru (review #438: ekran je
' javljao samo "Storno prijemnice nije uspeo", pa operater nije znao da je, npr.,
' paleta preradjena). Prazno kad je storno prosao.
Public Function StornoPrijemnica_TX(ByVal prijemnicaID As String, _
                                    Optional ByVal skiniPalete As Boolean = True, _
                                    Optional ByRef outPaletnihStavki As Long, _
                                    Optional ByRef outRazlog As String) As Boolean
    Const SRC As String = "StornoPrijemnica_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    outPaletnihStavki = 0
    outRazlog = ""
    tx.BeginTx
    SnimiZaStornoPrijemnice tx

    If Not StornoPrijemnica(prijemnicaID, tx, skiniPalete, outPaletnihStavki) Then
        Err.Raise ERR_STORNO_BASE + 4, SRC, _
                  "StornoPrijemnica nije uspeo. PrijemnicaID=" & prijemnicaID
    End If

    tx.CommitTx

    StornoPrijemnica_TX = True
    MonitorStornoSuccess SRC, "Prijemnica", prijemnicaID

    Set tx = Nothing
    Exit Function

EH:
    outPaletnihStavki = 0
    outRazlog = StornoRazlogOperateru(Err.description)
    HandleStornoTxError SRC, "Prijemnica", prijemnicaID, tx
    StornoPrijemnica_TX = False
End Function

' Poruka greske za operatera: LogAndReraise na svakom nivou dodaje prefiks
' "Source=<procedura> | ", pa je izvorni razlog posle POSLEDNJEG " | ". Prazan
' opis ostaje prazan -- pozivalac tada javlja samo neuspeh.
Public Function StornoRazlogOperateru(ByVal opis As String) As String
    Dim k As Long
    k = InStrRev(opis, " | ")
    If k > 0 Then
        StornoRazlogOperateru = Trim$(Mid$(opis, k + 3))
    Else
        StornoRazlogOperateru = Trim$(opis)
    End If
End Function

' JEDAN skup snimaka za SVAKOG pozivaoca storna prijemnice (S6 korak 4, KI-031).
'
' Kaskada zbirne je tblPrijemnica snimala samo u jednoj grani, a subjekat je
' stornirala u drugoj -- rollback je vracao sve osim zaglavlja. Skup je SIRI od
' onoga sto jezgro pise (stavke prijemnice, faktura i njene stavke): rollback
' vraca ceo dokument i njegovu okolinu, a jezgro posebno TRAZI ono sto samo menja
' (RequireSnimakZaStornoPrijemnice), pa pozivalac koji preskoci ovaj poziv pada
' glasno, pre prvog upisa. AddTableSnapshot preskace vec snimljenu tabelu.
Public Sub SnimiZaStornoPrijemnice(ByVal tx As clsTransaction)
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
End Sub

' JEZGRO STORNA PRIJEMNICE (S6 korak 4). Storno ponistava TRI dogadjaja --
' palete, dokument i knjigu -- i svi padaju ZAJEDNO ili nijedan (R3):
'
'   1) snimci     tx MORA da snima svaku tabelu koju jezgro menja (KI-031);
'   2) provere    red postoji jednom i nije storniran; dokument nije fakturisan
'                 (odluka 1); preradjena paleta se proverava u paletnom motoru,
'                 pre njegove prve mutacije (odluka 3);
'   3) upisi      palete -> Stornirano -> bind -> kontra-stav knjige.
'
' Stornirano ne ide pre posledica van knjige (R4), a knjiga ide POSLE bind-a, jer
' AMB-10-ODL-15 trazi da dokument vezuje samo pisac koji ga je u ovoj transakciji
' stvarno promenio. Pad bilo gde vraca SVE rollback-om pozivaoca: zato tacka 1.
'
' Faktura se ovde vise ne pise: fakturisana prijemnica se odbija, pa je stara
' grana osirotelosti (MarkFakturaOrphaned) postala nedostizna i obrisana je.
Public Function StornoPrijemnica(ByVal prijemnicaID As String, _
                                 ByVal tx As clsTransaction, _
                                 Optional ByVal skiniPalete As Boolean = True, _
                                 Optional ByRef outPaletnihStavki As Long) As Boolean
    Const SRC As String = "StornoPrijemnica"

    On Error GoTo EH

    outPaletnihStavki = 0
    RequireSnimakZaStornoPrijemnice tx, SRC

    Dim rowPrij As Long
    rowPrij = RequireStornoAllowed(TBL_PRIJEMNICA, prijemnicaID, COL_PRJ_ID, SRC)

    Dim razlogFak As String
    razlogFak = FakturisanaPrijemnicaRazlog(prijemnicaID)
    If Len(razlogFak) > 0 Then
        Err.Raise ERR_STORNO_BASE + 74, SRC, razlogFak
    End If

    If skiniPalete Then
        outPaletnihStavki = modPaletniList.SkiniPaleteZaStornoPrijemnice(tx, prijemnicaID, SRC)
    End If

    MarkRowStornirano TBL_PRIJEMNICA, rowPrij, SRC

    ' Vezivanje stoji POSLE izmene zaglavlja (AMB-10-ODL-15): ova transakcija je
    ' dokument stvarno promenila, pa sme da tvrdi da ga poseduje.
    tx.BindSourceDocument DOK_TIP_PRIJEMNICA, prijemnicaID

    ' Storno ambalaze je kontra-stav u knjizi (AMB-10-ODL-16), kao kod otkupa.
    modAmbalaza.StornirajAmbalazuDokumenta tx, DOK_TIP_PRIJEMNICA, prijemnicaID

    If mTestPrjStornoPadni Then
        If IsTestMode() Then
            Err.Raise ERR_STORNO_BASE + 75, SRC, _
                      "TEST: namerna greska posle svih upisa storna prijemnice (rollback dokaz)"
        End If
    End If

    StornoPrijemnica = True
    Exit Function

EH:
    outPaletnihStavki = 0
    LogAndReraise SRC
End Function

' KI-031 PO KONSTRUKCIJI: jezgro TRAZI snimak svake tabele koju menja, pre prvog
' upisa. Zakrpa jednog pozivaoca bi ostavila sledeceg otvorenog -- invarijanta
' ide u jezgro, ne u omotac.
Private Sub RequireSnimakZaStornoPrijemnice(ByVal tx As clsTransaction, _
                                           ByVal sourceName As String)
    If tx Is Nothing Then
        Err.Raise ERR_STORNO_BASE + 73, sourceName, _
                  "Storno prijemnice trazi aktivnu transakciju."
    End If
    Dim t As Variant
    For Each t In Array(TBL_PRIJEMNICA, TBL_AMBALAZA, TBL_PALETA, TBL_PALETA_STAVKA)
        If Not tx.ImaSnapshot(CStr(t)) Then
            Err.Raise ERR_STORNO_BASE + 73, sourceName, _
                      "Storno prijemnice: transakcija ne snima " & CStr(t) & _
                      " -- rollback ne bi vratio ceo dokument (KI-031). " & _
                      "Pozivalac snima kroz SnimiZaStornoPrijemnice."
        End If
    Next t
End Sub

' FAKTURISANA PRIJEMNICA SE NE STORNIRA (S6 korak 4, odluka 1). Razlog odbijanja,
' citljiv operateru; prazno = dokument nije fakturisan. JEDAN tekst za jezgro
' storna i okvir ispravke (KI-040) -- dva teksta bi se razisla. Delimicno
' fakturisan dokument se odbija kao fakturisan.
'
' Fakturisanost se ne racuna ovde: od S6 koraka 5 je izvedena iz aktivnih linija
' fakture u JEDNOM telu, modFaktura.PrijemnicaJeFakturisana (FAK-LIN-03), koje
' pitaju i lista, mreza i uvid. Do tada je ovde stajao drugi predikat (kes
' Fakturisano ILI linija), a lista je citala treci izvor (KI-038).
Public Function FakturisanaPrijemnicaRazlog(ByVal prijemnicaID As String) As String
    Dim fak As String
    If Not modFaktura.PrijemnicaJeFakturisana(prijemnicaID, fak) Then Exit Function
    FakturisanaPrijemnicaRazlog = "Prijemnica " & Trim$(prijemnicaID) & _
        " je fakturisana (faktura " & fak & "). Prvo stornirajte fakturu, pa prijemnicu."
End Function

' ============================================================
' FAKTURA
' ============================================================

Public Function StornoFaktura_TX(ByVal fakturaID As String) As Boolean
    Const SRC As String = "StornoFaktura_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    ' Prijemnica se ne snima: od S6 koraka 5 je storno fakture ne pise --
    ' oslobadja se sama, jer joj je fakturisanost izvedena iz linija.
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_NOVAC
    ' R3/krug 5 (revizija #248): GP grana oslobadja UTOVAR
    ' (ReleaseUtovarFromFaktura pise tblUtovar) -- bez snapshota bi
    ' pukli rollback vratio fakturu u aktivnu, a utovar bi OSTAO
    ' slobodan za novo fakturisanje = dvostruka prodaja iste isporuke.
    ' USLOVNO (B3): stariji workbook legitimno NEMA tblUtovar, a
    ' storno obicne sveze fakture tamo mora da radi -- isti meki
    ' ugovor kao GetColumnIndex kapije u ostatku GP integracije.
    If Not GetTable(TBL_UTOVAR) Is Nothing Then
        tx.AddTableSnapshot TBL_UTOVAR
    End If

    If Not StornoFaktura(fakturaID) Then
        Err.Raise ERR_STORNO_BASE + 5, SRC, _
                  "StornoFaktura nije uspeo. FakturaID=" & fakturaID
    End If

    tx.CommitTx

    StornoFaktura_TX = True
    MonitorStornoSuccess SRC, "Faktura", fakturaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Faktura", fakturaID, tx
    StornoFaktura_TX = False
End Function

Public Function StornoFaktura(ByVal fakturaID As String) As Boolean
    Const SRC As String = "StornoFaktura"

    On Error GoTo EH

    Dim rowFak As Long
    rowFak = RequireStornoAllowed(TBL_FAKTURE, fakturaID, COL_FAK_ID, SRC)

    RequireColumnIndex TBL_FAKTURE, COL_FAK_STATUS, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_STORNIRANO, SRC

    MarkRowStornirano TBL_FAKTURE, rowFak, SRC
    RequireUpdateCell TBL_FAKTURE, rowFak, COL_FAK_STATUS, STATUS_STORNIRANO, SRC

    StornoFakturaStavke fakturaID
    ' R3 rollback dokaz: tacka je POSLE oslobadjanja prijemnica/prerada
    ' a PRE kraja -- pukne li ovde, TX mora da vrati i tblPrerada.
    If mTestFailPosleRelease Then
        Err.Raise ERR_STORNO_BASE + 52, SRC, _
                  "TEST: namerna greska posle release-a (rollback dokaz)"
    End If
    ResetNovacFakturaLink fakturaID

    StornoFaktura = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' Pali/gasi namernu gresku (v. mTestFailPosleRelease) -- van test
' rezima ne radi nista, pa produkcioni put ne moze da je upali.
Public Sub StornoTestFailPosleRelease(ByVal fail As Boolean)
    If Not IsTestMode() Then Exit Sub
    mTestFailPosleRelease = fail
End Sub

' Pali/gasi namernu gresku na kraju jezgra storna prijemnice (v.
' mTestPrjStornoPadni). Paljenje van test rezima ne radi nista; gasenje radi
' uvek, da ga ResetSeamova ugasi bez obzira na rezim.
Public Sub StornoPrijemnicePadniTest(ByVal pali As Boolean)
    If pali And Not IsTestMode() Then Exit Sub
    mTestPrjStornoPadni = pali
End Sub

' ============================================================
' NOVAC
' ============================================================

Public Function StornoNovac_TX(ByVal novacID As String) As Boolean
    Const SRC As String = "StornoNovac_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    ' Poslovno pravilo (i kad pozivalac nije forma): pojedinacni storno je samo za
    ' rucno unet novac. Novac iz bankovnog izvoda se stornira SAMO u celosti (ceo
    ' izvod), nikad red po red -> odbij pre BeginTx (RollbackTx je no-op van TX).
    ' Provera je po REDU (marker), ne po broju - rucni red koji deli broj sa izvodom
    ' mora ostati stornirljiv.
    If IsNovacFromBankaImport(novacID) Then
        Err.Raise ERR_STORNO_BASE + 41, SRC, _
                  "Novac iz bankovnog izvoda se ne stornira pojedinacno (samo ceo izvod). " & _
                  "NovacID=" & novacID
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_OTKUP        ' StornoNovac osvezava Isplaceno/DatumIsplate vezanog otkupa

    If Not StornoNovac(novacID) Then
        Err.Raise ERR_STORNO_BASE + 6, SRC, _
                  "StornoNovac nije uspeo. NovacID=" & novacID
    End If

    tx.CommitTx

    StornoNovac_TX = True
    MonitorStornoSuccess SRC, "Novac", novacID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Novac", novacID, tx
    StornoNovac_TX = False
End Function

Public Function StornoNovac(ByVal novacID As String) As Boolean
    Const SRC As String = "StornoNovac"

    On Error GoTo EH

    Dim rowNov As Long
    rowNov = RequireStornoAllowed(TBL_NOVAC, novacID, COL_NOV_ID, SRC)

    RequireColumnIndex TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC
    RequireColumnIndex TBL_NOVAC, COL_NOV_OTKUP_ID, SRC

    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)

    If IsEmpty(novData) Then
        Err.Raise ERR_STORNO_BASE + 40, SRC, _
                  "Tabela novac je prazna."
    End If

    Dim fakturaID As String
    fakturaID = Trim$(CStr(novData(rowNov, _
                    RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC))))

    ' Uplata/isplata moze biti vezana i za otkup (blok) - ne samo za fakturu.
    ' Bez ovoga bi storno isplate ostavio otkup "Isplaceno" + DatumIsplate ustajao.
    Dim otkupID As String
    otkupID = Trim$(CStr(novData(rowNov, _
                    RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC))))

    MarkRowStornirano TBL_NOVAC, rowNov, SRC

    If Len(fakturaID) > 0 Then
        UpdateFakturaStatus fakturaID
    End If

    ' Status isplate se vise ne odrzava kolonom -- izveden je iz razlike
    ' VrednostOtkupa - SUM(isplate), a GetIsplataForOtkup vec iskljucuje
    ' stornirane redove. Storno novca zato nema sta da 'rekalkulise'.

    StornoNovac = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' Marker porekla cita modBankaMapiranje.BimIdFromNapomena (zivi uz pisca).
' True ako je KONKRETAN novac red nastao iz izvoda (nosi BIM marker u Napomeni).
'
' Provera je po REDU, ne po BrojDokumenta: svaki izvodni red nosi marker (direktan
' upis od BuildBIMNapomena, split ga nasledjuje od roditelja), pa je red-provera
' potpuna. Provera po broju bi uz to zarobila i RUCNI red koji slucajno deli broj
' sa izvodom - ostao bi netaknut pri stornu izvoda, ali i nestornirljiv pojedinacno
' (vidi T36). Poslovno pravilo ostaje: izvod se ne stornira parcijalno.
Public Function IsNovacFromBankaImport(ByVal novacID As String) As Boolean
    Const SRC As String = "IsNovacFromBankaImport"
    On Error GoTo EH

    RequireNonBlank novacID, "NovacID", SRC

    Dim rows As Collection
    Set rows = FindRows(TBL_NOVAC, COL_NOV_ID, novacID)
    If rows Is Nothing Then
        Err.Raise ERR_STORNO_BASE + 51, SRC, "FindRows je vratio Nothing. NovacID=" & novacID
    End If
    If rows.count = 0 Then
        Err.Raise ERR_STORNO_BASE + 52, SRC, "Novac red nije pronadjen. NovacID=" & novacID
    End If
    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 53, SRC, _
                  "NovacID nije jedinstven: " & novacID & " (Count=" & rows.count & ")"
    End If

    Dim data As Variant: data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_BASE + 54, SRC, "Tabela novac je prazna posle pronalaska reda."
    End If

    Dim colNap As Long: colNap = RequireColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA, SRC)
    IsNovacFromBankaImport = (Len(BimIdFromNapomena(CStr(data(CLng(rows(1)), colNap)))) > 0)
    Exit Function
EH:
    ' Guard koji odlucuje da li je destruktivna operacija dozvoljena NE sme da
    ' padne "otvoreno": greska u citanju bi inace znacila "nije iz izvoda" i
    ' propustila pojedinacni storno izvodnog reda. Fail-closed.
    LogAndReraise SRC
End Function

' Razresi ono sto je operater ukucao (BrojDokumenta ili NovacID) u NovacID za
' POJEDINACNI storno i primeni poslovna pravila. Vraca "" i popunjen reason kad
' storno nije dozvoljen (UI samo prikaze reason - bez MsgBox-a u business sloju).
' Kriterijum je POREKLO, ne Tip: rucno unet virman nije deo nijednog izvoda pa
' mora ostati ispravljiv; iz izvoda ne sme nista pojedinacno (vidi modConfig,
' sekcija Novac Tipovi - kanal placanja).
' Pravila:
'   1) izvod (BIM) -> odbij: izvod se stornira samo u celosti, ne parcijalno
'   2) broj sa vise aktivnih redova (avans-split) -> trazi NovacID (bez tihog
'      storna samo poslednjeg reda; ukucan NovacID je jednoznacan pa prolazi)
Public Function ResolveNovacForStorno(ByVal ulaz As String, _
                                      ByRef reason As String) As String
    Const SRC As String = "ResolveNovacForStorno"
    On Error GoTo EH
    reason = ""
    ulaz = Trim$(ulaz)

    Dim novID As String, poBroju As Boolean
    novID = LookupActiveID(TBL_NOVAC, COL_NOV_BROJ_DOK, ulaz, COL_NOV_ID)
    poBroju = (Len(novID) > 0)
    If Not poBroju Then novID = LookupActiveID(TBL_NOVAC, COL_NOV_ID, ulaz, COL_NOV_ID)

    If Len(novID) = 0 Then
        reason = "Novac stavka '" & ulaz & "' nije pronadjena (ili je ve" & ChrW(263) & " stornirana)."
        Exit Function
    End If

    Dim brNov As String
    brNov = NzToText(LookupValue(TBL_NOVAC, COL_NOV_ID, novID, COL_NOV_BROJ_DOK))

    If IsNovacFromBankaImport(novID) Then
        reason = "Novac '" & ulaz & "' potice iz bankovnog izvoda " & brNov & "." & vbCrLf & vbCrLf & _
                 "Izvod se ne stornira parcijalno - samo u celosti (Banka / uvoz izvoda)." & vbCrLf & _
                 "Ovde se stornira samo rucno unet novac (ke" & ChrW(353) & " / virman)."
        ' Redovi uvezeni PRE razdvajanja kanala nose KES tip iako su bankovni ->
        ' operateru objasni zasto "ke" & ChrW(353) & " red" ipak nije za pojedinacni storno.
        If IsKesNovacTip(NzToText(LookupValue(TBL_NOVAC, COL_NOV_ID, novID, COL_NOV_TIP))) Then
            reason = reason & vbCrLf & vbCrLf & _
                     "(Red ima KES tip, ali je uvezen iz izvoda - stari zapis, pre razdvajanja kanala.)"
        End If
        Exit Function
    End If

    If poBroju Then
        Dim n As Long: n = CountActiveNovacByBroj(brNov)
        If n > 1 Then
            reason = "Broj '" & brNov & "' ima " & n & " aktivnih novac stavki " & _
                     "(avans raspodela deli isti broj)." & vbCrLf & vbCrLf & _
                     "Storno po broju bi stornirao samo jednu. Unesite NovacID (NOV-...) tacne stavke."
            Exit Function
        End If
    End If

    ResolveNovacForStorno = novID
    Exit Function
EH:
    LogErr "modStorno.ResolveNovacForStorno"
    reason = "Greska pri razresavanju novac stavke: " & Err.description
End Function

' Broj AKTIVNIH tblNovac redova za dati BrojDokumenta. BrojDokumenta NIJE
' jedinstven kljuc: uvoz izvoda upisuje sve stavke jednog izvoda pod istim brojem
' (stari redovi mogu nositi i literal "IZVOD" - taj fallback je uklonjen, vidi
' modBankaMapiranje.RequireIzvodBroj), a ApplyAvansToFaktura/ApplyAvansToOtkup
' split nasledjuje broj originalne stavke. Forma to koristi da NE stornira tiho
' samo jedan od vise redova sa istim brojem (jedini jedinstven kljuc je NovacID).
Public Function CountActiveNovacByBroj(ByVal brDok As String) As Long
    Const SRC As String = "CountActiveNovacByBroj"
    On Error GoTo EH
    If Trim$(brDok) = "" Then Exit Function
    Dim data As Variant: data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    Dim colBroj As Long, colStorno As Long
    colBroj = RequireColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK, SRC)
    colStorno = RequireColumnIndex(TBL_NOVAC, COL_STORNIRANO, SRC)
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colBroj))) = Trim$(brDok) Then
            If Not IsStorniranoValue(data(i, colStorno)) Then n = n + 1
        End If
    Next i
    CountActiveNovacByBroj = n
    Exit Function
EH:
    ' Fail-closed (isti razlog kao IsNovacFromBankaImport): tiha nula bi
    ' preskocila proveru vise aktivnih redova i pustila tih parcijalan storno.
    LogAndReraise SRC
End Function

' ============================================================
' IZVOD (bankovni) - storno u CELOSTI
'
' Izvod se nikad ne stornira parcijalno (4-nivo integritet uvoza). Storno obara
' SVE novac redove tog izvoda u jednoj transakciji, pa staging redove obradjuje
' po ishodu koji bira operater:
'   IZVOD_STORNO_REMAP    - PDF ispravan, mapiranje pogresno -> Obradjeno = ""
'                           (stavke nazad u "za obradu"; izvod ostaje uvezen)
'   IZVOD_STORNO_REIMPORT - PDF los/korumpiran -> Stornirano = "Da"
'                           (izvod se uvozi ponovo; IsDuplicateBankaImport i
'                            GetBankaImportOpen rade nad ExcludeStornirano)
'
' Novac i staging MORAJU pasti u istoj TX: ako se staging ugasi a novac ostane
' aktivan, ponovni uvoz + mapiranje daju dvostruko knjizenje.
' ============================================================

' Razresi ono sto je operater ukucao ("broj" ili "broj/racun") u jedan izvod.
' Vise banaka moze imati isti broj izvoda -> tada trazi "broj/racun" umesto da
' tiho uzme jedan. False + reason = ne diraj nista.
Public Function ResolveIzvodZaStorno(ByVal ulaz As String, ByRef brojIzvoda As String, _
                                     ByRef brojRacuna As String, ByRef reason As String) As Boolean
    Const SRC As String = "ResolveIzvodZaStorno"
    On Error GoTo EH
    reason = "": brojIzvoda = "": brojRacuna = ""

    Dim s As String: s = Trim$(ulaz)
    If Len(s) = 0 Then
        reason = "Unesite broj izvoda (ili 'broj/racun' ako isti broj postoji na vi" & ChrW(353) & "e racuna)."
        Exit Function
    End If

    ' Broj izvoda SME da sadrzi kosu crtu (npr. "42/2026") -> prvo probaj ceo unos
    ' kao broj; tek ako takvog izvoda nema, tumaci poslednju crtu kao broj/racun.
    Dim racuni As Object
    brojIzvoda = s: brojRacuna = ""
    Set racuni = IzvodRacuniZaBroj(brojIzvoda, "", SRC)

    If racuni.count = 0 Then
        Dim p As Long: p = InStrRev(s, "/")
        If p > 1 Then
            brojIzvoda = Trim$(Left$(s, p - 1))
            brojRacuna = Trim$(Mid$(s, p + 1))
            Set racuni = IzvodRacuniZaBroj(brojIzvoda, brojRacuna, SRC)
        End If
    End If

    If racuni.count = 0 Then
        reason = "Aktivan izvod nije prona" & ChrW(273) & "en: " & ulaz
        Exit Function
    End If

    If racuni.count > 1 Then
        Dim lst As String, k As Variant
        For Each k In racuni.Keys
            lst = lst & vbCrLf & "  - " & brojIzvoda & "/" & CStr(k)
        Next k
        reason = "Broj izvoda '" & brojIzvoda & "' postoji na vi" & ChrW(353) & "e racuna:" & lst & vbCrLf & vbCrLf & _
                 "Unesite 'broj/racun' da se zna koji se izvod stornira."
        Exit Function
    End If

    If Len(brojRacuna) = 0 Then
        Dim kk As Variant: kk = racuni.Keys
        brojRacuna = CStr(kk(0))
    End If
    ResolveIzvodZaStorno = True
    Exit Function
EH:
    LogErr "modStorno.ResolveIzvodZaStorno"
    reason = "Greska pri razresavanju izvoda: " & Err.description
End Function

' Numericka vrednost celije ili 0 (iznosi u staging/novac tabelama mogu biti prazni).
Private Function NumOrZero(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOrZero = CDbl(v)
End Function

' Tekuca suma iz akumulator-dictionary-ja (0 ako kljuc jos ne postoji).
Private Function NzNum(ByVal d As Object, ByVal key As String) As Double
    If d.Exists(key) Then NzNum = CDbl(d(key))
End Function

' Distinktni racuni koji imaju AKTIVAN staging red sa datim brojem izvoda
' (opciono suzeno na jedan racun). Prazan skup = takav izvod ne postoji.
Private Function IzvodRacuniZaBroj(ByVal brojIzvoda As String, ByVal brojRacuna As String, _
                                   ByVal sourceName As String) As Object
    Dim racuni As Object: Set racuni = CreateObject("Scripting.Dictionary")
    Set IzvodRacuniZaBroj = racuni
    If Len(Trim$(brojIzvoda)) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Function

    Dim cBroj As Long, cRac As Long, cSt As Long
    cBroj = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, sourceName)
    cRac = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA, sourceName)
    cSt = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO, sourceName)

    Dim i As Long, rac As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBroj))) = Trim$(brojIzvoda) Then
            If Not IsStorniranoValue(data(i, cSt)) Then
                rac = Trim$(CStr(data(i, cRac)))
                If Len(Trim$(brojRacuna)) = 0 Or rac = Trim$(brojRacuna) Then racuni(rac) = True
            End If
        End If
    Next i
End Function

' PREFLIGHT: sve sto mora da vazi PRE nego sto se dirne ijedan red. Vraca ""
' kad je storno bezbedan, inace razlog (UI ga prikaze; TX ga podize kao gresku).
' Obrazac je isti kao ResolveNovacForStorno - poslovna pravila u modulu, poruka
' u UI-ju. Pravila:
'   1) racun je obavezan (isti broj postoji na vise banaka)
'   2) staging red bez BankaImportID -> odbij (bez PK nema pouzdanog lineage-a)
'   3) REKONSILIJACIJA po stavci: za svaku OBRADJENU stavku zbir AKTIVNOG novca sa
'      njenim markerom mora biti jednak iznosu stavke (uplata i isplata zasebno).
'      Split umanji original i nosi ostatak pod istim markerom, pa zbir mora da se
'      poklopi. Nesklad = nesto je vec dirnuto -> odbij ceo posao (vracanje u
'      "za obradu" uz aktivan novac bi omogucilo dvostruko knjizenje).
Public Function GetIzvodStornoBlokade(ByVal brojIzvoda As String, _
                                      ByVal brojRacuna As String) As String
    Const SRC As String = "GetIzvodStornoBlokade"
    On Error GoTo EH

    If Len(Trim$(brojRacuna)) = 0 Then
        GetIzvodStornoBlokade = "Nije odre" & ChrW(273) & "en racun izvoda - identitet je dvosmislen " & _
            "(isti broj moze postojati na vise racuna/banaka)."
        Exit Function
    End If

    Dim bim As Variant: bim = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(bim) Then
        GetIzvodStornoBlokade = "Tabela uvoza izvoda je prazna."
        Exit Function
    End If

    Dim cId As Long, cBroj As Long, cRac As Long, cSt As Long, cObr As Long
    cId = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, SRC)
    cBroj = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, SRC)
    cRac = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA, SRC)
    cSt = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO, SRC)
    cObr = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO, SRC)

    Dim cUpl As Long, cIsp As Long
    cUpl = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA, SRC)
    cIsp = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA, SRC)

    ' Ocekivani iznosi po OBRADJENOJ stavci (uplata i isplata zasebno).
    Dim ocekUpl As Object: Set ocekUpl = CreateObject("Scripting.Dictionary")
    Dim ocekIsp As Object: Set ocekIsp = CreateObject("Scripting.Dictionary")
    Dim i As Long, n As Long, bid As String

    For i = 1 To UBound(bim, 1)
        If Trim$(CStr(bim(i, cBroj))) = Trim$(brojIzvoda) And Not IsStorniranoValue(bim(i, cSt)) Then
            If Trim$(CStr(bim(i, cRac))) = Trim$(brojRacuna) Then
                bid = Trim$(CStr(bim(i, cId)))
                If Len(bid) = 0 Then
                    GetIzvodStornoBlokade = "Staging red bez BankaImportID (red " & i & ") - " & _
                        "bez njega storno ne moze pouzdano da odredi svoj novac. Popravi red pa ponovi."
                    Exit Function
                End If
                n = n + 1
                If UCase$(Trim$(CStr(bim(i, cObr)))) = "DA" Then
                    ocekUpl(bid) = NzNum(ocekUpl, bid) + NumOrZero(bim(i, cUpl))
                    ocekIsp(bid) = NzNum(ocekIsp, bid) + NumOrZero(bim(i, cIsp))
                End If
            End If
        End If
    Next i

    If n = 0 Then
        GetIzvodStornoBlokade = "Aktivan izvod nije prona" & ChrW(273) & "en: " & brojIzvoda & "/" & brojRacuna
        Exit Function
    End If
    If ocekUpl.count = 0 Then Exit Function     ' nista nije mapirano -> nema sta da se sravnjuje

    ' Stvarni zbir AKTIVNOG novca po markeru. Split umanji original i nosi ostatak
    ' pod ISTIM markerom, pa zbir mora da se poklopi sa iznosom stavke.
    Dim stvarUpl As Object: Set stvarUpl = CreateObject("Scripting.Dictionary")
    Dim stvarIsp As Object: Set stvarIsp = CreateObject("Scripting.Dictionary")
    Dim nov As Variant: nov = GetTableData(TBL_NOVAC)
    If Not IsEmpty(nov) Then
        Dim nNap As Long, nSt As Long, nUpl As Long, nIsp As Long
        nNap = RequireColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA, SRC)
        nSt = RequireColumnIndex(TBL_NOVAC, COL_STORNIRANO, SRC)
        nUpl = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)
        nIsp = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)
        Dim rb As String
        For i = 1 To UBound(nov, 1)
            If Not IsStorniranoValue(nov(i, nSt)) Then
                rb = BimIdFromNapomena(CStr(nov(i, nNap)))
                If Len(rb) > 0 Then
                    If ocekUpl.Exists(rb) Then
                        stvarUpl(rb) = NzNum(stvarUpl, rb) + NumOrZero(nov(i, nUpl))
                        stvarIsp(rb) = NzNum(stvarIsp, rb) + NumOrZero(nov(i, nIsp))
                    End If
                End If
            End If
        Next i
    End If

    Dim k As Variant
    For Each k In ocekUpl.Keys
        If Not stvarUpl.Exists(CStr(k)) Then
            GetIzvodStornoBlokade = "Stavka " & CStr(k) & " je ozna" & ChrW(269) & "ena kao obra" & ChrW(273) & "ena, " & _
                "ali nema nijedan aktivan novac red. Vracanje u 'za obradu' bi omogucilo " & _
                "dvostruko knjizenje. Proveri stavku rucno pa ponovi."
            Exit Function
        End If
        ' Iznosi su na dve decimale -> poredi zaokruzene, sa epsilon samo za
        ' Double sabiranje. Tolerancija 0.01 bi propustila razliku od tacno 1 pare.
        If Abs(Round(CDbl(stvarUpl(CStr(k))), 2) - Round(CDbl(ocekUpl(CStr(k))), 2)) > 0.001 Or _
           Abs(Round(CDbl(stvarIsp(CStr(k))), 2) - Round(CDbl(ocekIsp(CStr(k))), 2)) > 0.001 Then
            GetIzvodStornoBlokade = "Iznosi se ne sla" & ChrW(382) & "u za stavku " & CStr(k) & ": izvod " & _
                Format$(CDbl(ocekUpl(CStr(k))), "#,##0.00") & " / " & Format$(CDbl(ocekIsp(CStr(k))), "#,##0.00") & _
                ", knjizeno " & Format$(CDbl(stvarUpl(CStr(k))), "#,##0.00") & " / " & _
                Format$(CDbl(stvarIsp(CStr(k))), "#,##0.00") & " (uplata/isplata). " & _
                "Storno je odbijen dok se razlika ne razjasni."
            Exit Function
        End If
    Next k
    Exit Function
EH:
    ' Fail-closed: neuspela provera znaci "ne znam", ne "sme".
    LogErr "modStorno.GetIzvodStornoBlokade"
    GetIzvodStornoBlokade = "Provera izvoda nije uspela: " & Err.description
End Function

' Pregled pre potvrde: sta ce tacno pasti (stavke uvoza + novac redovi + iznosi).
Public Function GetIzvodPregled(ByVal brojIzvoda As String, ByVal brojRacuna As String) As String
    Const SRC As String = "GetIzvodPregled"
    On Error GoTo EH

    Dim stRows As Collection: Set stRows = New Collection
    Dim bimIDs As Object: Set bimIDs = CreateObject("Scripting.Dictionary")
    CollectIzvodStaging brojIzvoda, brojRacuna, stRows, bimIDs, SRC

    Dim novIDs As Collection: Set novIDs = CollectIzvodNovacIDs(bimIDs, SRC)

    Dim upl As Double, isp As Double, k As Long
    Dim v As Variant
    For k = 1 To novIDs.count
        v = LookupValue(TBL_NOVAC, COL_NOV_ID, CStr(novIDs(k)), COL_NOV_UPLATA)
        If IsNumeric(v) Then upl = upl + CDbl(v)
        v = LookupValue(TBL_NOVAC, COL_NOV_ID, CStr(novIDs(k)), COL_NOV_ISPLATA)
        If IsNumeric(v) Then isp = isp + CDbl(v)
    Next k

    GetIzvodPregled = "IZVOD " & brojIzvoda & IIf(Len(brojRacuna) > 0, " / " & brojRacuna, "") & vbCrLf & _
        "Stavke uvoza: " & stRows.count & vbCrLf & _
        "Novac redovi (sa avans raspodelom): " & novIDs.count & vbCrLf & _
        "Uplate: " & Format$(upl, "#,##0.00") & "   Isplate: " & Format$(isp, "#,##0.00")
    Exit Function
EH:
    LogErr "modStorno.GetIzvodPregled"
    GetIzvodPregled = "IZVOD " & brojIzvoda & " (pregled nije dostupan - vidi Monitor)"
End Function

Public Function StornoIzvod_TX(ByVal brojIzvoda As String, ByVal brojRacuna As String, _
                               ByVal ishod As String, Optional ByRef info As String) As Boolean
    Const SRC As String = "StornoIzvod_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    RequireNonBlank brojIzvoda, "BrojIzvoda", SRC
    ' Racun je deo identiteta izvoda, ne UI detalj: invariant mora da vazi i kad
    ' pozivalac nije forma (test, servisni makro, migracija).
    RequireNonBlank brojRacuna, "BrojRacuna", SRC
    If ishod <> IZVOD_STORNO_REMAP And ishod <> IZVOD_STORNO_REIMPORT Then
        Err.Raise ERR_STORNO_BASE + 46, SRC, "Nepoznat ishod storna izvoda: " & ishod
    End If

    ' Preflight PRE bilo kakve mutacije: prazan PK, nerazresen lineage, dvosmislena
    ' pripadnost -> odbij ceo posao (bez delimicnog storna).
    Dim blokada As String
    blokada = GetIzvodStornoBlokade(brojIzvoda, brojRacuna)
    If Len(blokada) > 0 Then
        Err.Raise ERR_STORNO_BASE + 50, SRC, blokada
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_BANKA_IMPORT

    Dim stRows As Collection: Set stRows = New Collection
    Dim bimIDs As Object: Set bimIDs = CreateObject("Scripting.Dictionary")
    CollectIzvodStaging brojIzvoda, brojRacuna, stRows, bimIDs, SRC

    If stRows.count = 0 Then
        Err.Raise ERR_STORNO_BASE + 47, SRC, _
                  "Aktivan izvod nije pronadjen. Broj=" & brojIzvoda & " Racun=" & brojRacuna
    End If

    ' Novac: direktni redovi izvoda + avans-split naslednici (nose novac istog
    ' izvoda, a izgubili su BIM marker). Reuse StornoNovac po redu - markira red i
    ' osvezava status vezane fakture/otkupa.
    Dim novIDs As Collection: Set novIDs = CollectIzvodNovacIDs(bimIDs, SRC)
    Dim k As Long
    For k = 1 To novIDs.count
        If Not StornoNovac(CStr(novIDs(k))) Then
            Err.Raise ERR_STORNO_BASE + 48, SRC, _
                      "Storno novac reda nije uspeo. NovacID=" & CStr(novIDs(k))
        End If
    Next k

    ' Staging po ishodu. Indeksi redova su i dalje vazeci (novac storno ne dira
    ' tblBankaImport).
    Dim i As Long
    For i = 1 To stRows.count
        If ishod = IZVOD_STORNO_REMAP Then
            RequireUpdateCell TBL_BANKA_IMPORT, CLng(stRows(i)), COL_BIM_OBRADJENO, "", SRC
        Else
            RequireUpdateCell TBL_BANKA_IMPORT, CLng(stRows(i)), COL_BIM_STORNIRANO, STORNO_DA, SRC
        End If
    Next i

    tx.CommitTx

    If ishod = IZVOD_STORNO_REMAP Then
        info = "Izvod " & brojIzvoda & " storniran." & vbCrLf & _
               "Novac redova oboreno: " & novIDs.count & vbCrLf & _
               "Stavki vraceno u 'za obradu': " & stRows.count & vbCrLf & _
               "Izvod ostaje uvezen - mapiraj stavke ponovo (Banka / uvoz izvoda)."
    Else
        info = "Izvod " & brojIzvoda & " storniran i uga" & ChrW(353) & "en." & vbCrLf & _
               "Novac redova oboreno: " & novIDs.count & vbCrLf & _
               "Stavki uvoza ugaseno: " & stRows.count & vbCrLf & _
               "Uvezi izvod PONOVO iz ispravnog PDF-a."
    End If

    StornoIzvod_TX = True
    MonitorStornoSuccess SRC, "Izvod", brojIzvoda & "/" & brojRacuna & " [" & ishod & "]"

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Izvod", brojIzvoda, tx
    StornoIzvod_TX = False
End Function

' AKTIVNI staging redovi jednog izvoda: indeksi redova + skup BankaImportID-jeva.
Private Sub CollectIzvodStaging(ByVal brojIzvoda As String, ByVal brojRacuna As String, _
                                ByRef rowsOut As Collection, ByRef idsOut As Object, _
                                ByVal sourceName As String)
    Dim data As Variant: data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Sub

    Dim cId As Long, cBroj As Long, cRac As Long, cSt As Long
    cId = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, sourceName)
    cBroj = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, sourceName)
    cRac = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA, sourceName)
    cSt = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO, sourceName)

    Dim i As Long, bimID As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBroj))) = Trim$(brojIzvoda) Then
            If Trim$(CStr(data(i, cRac))) = Trim$(brojRacuna) Then
                If Not IsStorniranoValue(data(i, cSt)) Then
                    ' PK je OBAVEZAN: prazan BankaImportID bi u skup ubacio kljuc ""
                    ' i time (preko BimIdFromNapomena = "" za svaki rucni red)
                    ' prosirio storno na sav markerless novac. Fail-fast.
                    bimID = Trim$(CStr(data(i, cId)))
                    If Len(bimID) = 0 Then
                        Err.Raise ERR_STORNO_BASE + 49, sourceName, _
                                  "BankaImport red bez BankaImportID (red " & i & ") -> " & _
                                  "storno izvoda je odbijen. Popravi staging red pa ponovi."
                    End If
                    rowsOut.Add i
                    idsOut(bimID) = True
                End If
            End If
        End If
    Next i
End Sub

' AKTIVNI novac redovi jednog izvoda - ISKLJUCIVO po BIM markeru.
'
' Svaki red koji pripada izvodu nosi marker: direktni upis ga dobija od
' BuildBIMNapomena, a avans-split ga NASLEDJUJE (BuildAvansSplitNapomena), pa je
' pripadnost eksplicitna i ne pogadja se. Marker mora biti NEPRAZAN pre poredjenja
' sa skupom - prazan marker ima svaki rucno unet red, pa bi ga poredjenje "usvojilo"
' u izvod. Nema fallback-a po broju/partneru: taj put bi mogao da obori tudj rucni
' red, a ne resava nijedan stvaran slucaj (splitovi nose marker).
Private Function CollectIzvodNovacIDs(ByVal bimIDs As Object, _
                                      ByVal sourceName As String) As Collection
    Dim result As Collection: Set result = New Collection
    Set CollectIzvodNovacIDs = result

    Dim data As Variant: data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cNap As Long, cSt As Long
    cId = RequireColumnIndex(TBL_NOVAC, COL_NOV_ID, sourceName)
    cNap = RequireColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA, sourceName)
    cSt = RequireColumnIndex(TBL_NOVAC, COL_STORNIRANO, sourceName)

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, nid As String, rowBimID As String

    For i = 1 To UBound(data, 1)
        If Not IsStorniranoValue(data(i, cSt)) Then
            rowBimID = BimIdFromNapomena(CStr(data(i, cNap)))
            If Len(rowBimID) > 0 Then
                If bimIDs.Exists(rowBimID) Then
                    nid = Trim$(CStr(data(i, cId)))
                    If Not seen.Exists(nid) Then
                        result.Add nid
                        seen(nid) = True
                    End If
                End If
            End If
        End If
    Next i
End Function

' ============================================================
' PALETA
' ============================================================

Public Function StornoPaleta_TX(ByVal palID As String) As Boolean
    Const SRC As String = "StornoPaleta_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    If Not StornoPaleta(palID) Then
        Err.Raise ERR_STORNO_BASE + 40, SRC, _
                  "StornoPaleta nije uspeo. PaletaID=" & palID
    End If

    tx.CommitTx

    StornoPaleta_TX = True
    MonitorStornoSuccess SRC, "Paleta", palID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Paleta", palID, tx
    StornoPaleta_TX = False
End Function

Public Function StornoPaleta(ByVal palID As String) As Boolean
    Const SRC As String = "StornoPaleta"

    On Error GoTo EH

    Dim rowPal As Long
    rowPal = RequireStornoAllowed(TBL_PALETA, palID, COL_PAL_ID, SRC)

    ' preradjenu paletu ne stornira se direktno -> prvo storno prerade
    Dim colPre As Long
    colPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
    Dim palData As Variant: palData = GetTableData(TBL_PALETA)
    If UCase$(Trim$(CStr(palData(rowPal, colPre)))) = "DA" Then
        Err.Raise ERR_STORNO_BASE + 42, SRC, _
                  "Paleta je preradjena - prvo stornirajte preradu."
    End If

    MarkRowStornirano TBL_PALETA, rowPal, SRC

    ' storniraj stavke palete (prijemnice se time oslobadjaju)
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(s) Then
        Dim sPal As Long, sStorno As Long
        sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
        sStorno = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)
        Dim r As Long
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(s(r, sPal))) = Trim$(palID) _
               And Not IsStorniranoValue(s(r, sStorno)) Then
                MarkRowStornirano TBL_PALETA_STAVKA, r, SRC
            End If
        Next r
    End If

    StornoPaleta = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' UTOVAR (krug 5): storno utovarne liste VRACA robu na stanje.
' Kapija: fakturisan utovar se ne stornira -- prvo storno fakture
' (koji ga oslobadja), pa storno utovara. Stavke se markiraju zajedno
' sa headerom (utovar je jedan dokument).
' ============================================================

Public Function StornoUtovar_TX(ByVal utovarID As String) As Boolean
    Const SRC As String = "StornoUtovar_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR
    tx.AddTableSnapshot TBL_UTOVAR_STAVKE

    If Not StornoUtovar(utovarID) Then
        Err.Raise ERR_STORNO_BASE + 54, SRC, _
                  "StornoUtovar nije uspeo. UtovarID=" & utovarID
    End If

    tx.CommitTx

    StornoUtovar_TX = True
    MonitorStornoSuccess SRC, "Utovar", utovarID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Utovar", utovarID, tx
    StornoUtovar_TX = False
End Function

Public Function StornoUtovar(ByVal utovarID As String) As Boolean
    Const SRC As String = "StornoUtovar"

    On Error GoTo EH

    Dim rowUt As Long
    rowUt = RequireStornoAllowed(TBL_UTOVAR, utovarID, COL_UT_ID, SRC)

    Dim d As Variant
    d = GetTableData(TBL_UTOVAR)
    Dim cFakt As Long, cFid As Long
    cFakt = RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURISANO, SRC)
    cFid = RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)
    If UCase$(Trim$(CStr(nz(d(rowUt, cFakt))))) = "DA" _
       Or Len(Trim$(CStr(nz(d(rowUt, cFid))))) > 0 Then
        Err.Raise ERR_STORNO_BASE + 55, SRC, _
                  "Utovar je fakturisan -- prvo storniraj fakturu. UtovarID=" & utovarID
    End If

    ' Revizija #10 B2: header marker nije dovoljan -- AKTIVNA faktura-
    ' stavka koja tvrdi ovaj utovar znaci da finansijski dokument i
    ' dalje prodaje bas ovu robu; storno bi je vratio na stanje =
    ' dupla zaliha. Isto kanonsko pravilo kao CreateFakturaIzUtovara.
    If modUtovar.AktivnihFstZaUtovar(utovarID) > 0 Then
        Err.Raise ERR_STORNO_BASE + 56, SRC, _
                  "Aktivna faktura-stavka tvrdi ovaj utovar -- podaci su " & _
                  "neusaglaseni, storno utovara je blokiran. UtovarID=" & utovarID
    End If

    MarkRowStornirano TBL_UTOVAR, rowUt, SRC

    Dim s As Variant, r As Long
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If Not IsEmpty(s) Then
        Dim sUt As Long, sSt As Long
        sUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, SRC)
        sSt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_STORNIRANO, SRC)
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(s(r, sUt))) = Trim$(utovarID) _
               And Not IsStorniranoValue(s(r, sSt)) Then
                MarkRowStornirano TBL_UTOVAR_STAVKE, r, SRC
            End If
        Next r
    End If

    StornoUtovar = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PRERADA
' ============================================================

Public Function StornoPrerada_TX(ByVal preradaID As String) As Boolean
    Const SRC As String = "StornoPrerada_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRERADA
    tx.AddTableSnapshot TBL_PRERADA_STAVKA
    tx.AddTableSnapshot TBL_PALETA

    If Not StornoPrerada(preradaID) Then
        Err.Raise ERR_STORNO_BASE + 45, SRC, _
                  "StornoPrerada nije uspeo. PreradaID=" & preradaID
    End If

    tx.CommitTx

    StornoPrerada_TX = True
    MonitorStornoSuccess SRC, "Prerada", preradaID

    Set tx = Nothing
    Exit Function

EH:
    HandleStornoTxError SRC, "Prerada", preradaID, tx
    StornoPrerada_TX = False
End Function

Public Function StornoPrerada(ByVal preradaID As String) As Boolean
    Const SRC As String = "StornoPrerada"

    On Error GoTo EH

    Dim rowPre As Long
    rowPre = RequireStornoAllowed(TBL_PRERADA, preradaID, COL_PRE_ID, SRC)

    ' Krug 5: prodaja zivi na UTOVARNOJ LISTI, ne na markeru prerade.
    ' Prerada sa aktivnim utovarenim stavkama se NE stornira (roba je
    ' fizicki isporucena) -- prvo storno fakture pa storno utovara,
    ' tek onda prerada. Fail-closed umesto orphan kaskade.
    If UtovarenoKgPrerade(preradaID) > 0 Then
        Err.Raise ERR_STORNO_BASE + 53, SRC, _
                  "Prerada ima aktivne utovarene stavke -- prvo storniraj " & _
                  "utovar (i njegovu fakturu). PreradaID=" & preradaID
    End If

    MarkRowStornirano TBL_PRERADA, rowPre, SRC

    ' storniraj stavke + vrati preradene palete u lager (Preradjeno = "")
    Dim s As Variant: s = GetTableData(TBL_PRERADA_STAVKA)
    If Not IsEmpty(s) Then
        Dim sPre As Long, sPalID As Long, sStorno As Long
        sPre = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID, SRC)
        sPalID = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID, SRC)
        sStorno = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_STORNIRANO, SRC)
        Dim r As Long
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(s(r, sPre))) = Trim$(preradaID) _
               And Not IsStorniranoValue(s(r, sStorno)) Then
                MarkRowStornirano TBL_PRERADA_STAVKA, r, SRC

                Dim palID As String: palID = Trim$(CStr(s(r, sPalID)))
                Dim c As Collection: Set c = FindRows(TBL_PALETA, COL_PAL_ID, palID)
                If Not c Is Nothing Then
                    If c.count > 0 Then
                        RequireUpdateCell TBL_PALETA, CLng(c(1)), COL_PAL_PRERADJENO, "", SRC
                    End If
                End If
            End If
        Next r
    End If

    StornoPrerada = True
    Exit Function

EH:
    LogAndReraise SRC
End Function

' ============================================================
' PUBLIC HELPERS / COMPATIBILITY
' ============================================================

Public Function CanStorno(ByVal tblName As String, _
                          ByVal recordID As String, _
                          ByVal idColumn As String) As Boolean
    Const SRC As String = "CanStorno"

    On Error GoTo EH

    CanStorno = (RequireStornoAllowed(tblName, recordID, idColumn, SRC) > 0)
    Exit Function

EH:
    LogErr SRC
    On Error Resume Next
    Debug.Print SRC & " failed. Table=" & tblName & _
                " ID=" & recordID & _
                " Err=" & CStr(Err.Number) & _
                " Desc=" & Err.description
    On Error GoTo 0

    CanStorno = False
End Function

Public Function LookupActiveID(ByVal tblName As String, _
                               ByVal brojColName As String, _
                               ByVal brojValue As String, _
                               ByVal idColName As String) As String
    Const SRC As String = "LookupActiveID"

    On Error GoTo EH

    RequireNonBlank tblName, "TableName", SRC
    RequireNonBlank brojColName, "BrojColumn", SRC
    RequireNonBlank idColName, "IdColumn", SRC

    Dim data As Variant
    data = GetTableData(tblName)

    If IsEmpty(data) Then
        LookupActiveID = ""
        Exit Function
    End If

    Dim colBroj As Long
    Dim colID As Long
    Dim colStorno As Long

    colBroj = RequireColumnIndex(tblName, brojColName, SRC)
    colID = RequireColumnIndex(tblName, idColName, SRC)
    colStorno = RequireColumnIndex(tblName, COL_STORNIRANO, SRC)

    Dim resultId As String
    Dim i As Long

    For i = 1 To UBound(data, 1)
        ' ZBR-NORM-02: isto poredjenje kao VlasniciPoBroju -- odlucivac.
        If BrojJednak(data(i, colBroj), brojValue) Then
            If Not IsStorniranoValue(data(i, colStorno)) Then
                resultId = CStr(data(i, colID))
            End If
        End If
    Next i

    LookupActiveID = resultId
    Exit Function

EH:
    LogErr SRC
    On Error Resume Next
    On Error GoTo 0
    LookupActiveID = ""
End Function

' ============================================================
' PRIVATE BUSINESS HELPERS
' ============================================================

' Linije fakture se storniraju. Prijemnica se time OSLOBADJA SAMA -- njena
' fakturisanost je izvedena iz aktivnih linija (S6 korak 5, FAK-LIN-03), pa se na
' njoj nista ne pise. Do koraka 5 je ovde stajao ReleasePrijemnicaFromFaktura,
' koji je brisao kes Fakturisano/FakturaID na zaglavlju.
Private Sub StornoFakturaStavke(ByVal fakturaID As String)
    Const SRC As String = "StornoFakturaStavke"

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colFakID As Long

    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)

    Dim i As Long

    ' GP grana (krug 5): stavka gotove robe nosi UTOVAR umesto
    ' prijemnice -- storno fakture OSLOBADJA utovar (reset markera),
    ' roba ostaje utovarena dok se i utovar ne stornira zasebno.
    ' GetColumnIndex, ne Require: na svesci PRE EnsureSchema nadogradnje
    ' kolona ne postoji, a storno svezih faktura tamo mora da radi.
    Dim colUtID As Long
    colUtID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_UTOVAR_ID)
    ' Release JEDNOM po utovaru (revizija #11 P1): faktura sa vise
    ' stavki istog utovara bi drugim/trecim pozivom pravila lazne
    ' "tvrdi drugu fakturu" warninge na potpuno validnom stornu.
    Dim relUt As Object: Set relUt = CreateObject("Scripting.Dictionary")
    relUt.CompareMode = vbTextCompare

    For i = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(i, colFakID))) = Trim$(fakturaID) Then
            MarkRowStornirano TBL_FAKTURA_STAVKE, i, SRC

            If colUtID > 0 Then
                Dim utID As String
                utID = Trim$(CStr(nz(stavkeData(i, colUtID))))
                If Len(utID) > 0 Then
                    If Not relUt.Exists(utID) Then
                        relUt.Add utID, True
                        ReleaseUtovarFromFaktura utID, fakturaID
                    End If
                End If
            End If
        End If
    Next i
End Sub

' GP grana (krug 5): storno GP fakture oslobadja UTOVAR (njegov marker je
' cinjenica utovara, ne izvedena) -- fail-closed na dupli ID.
Private Sub ReleaseUtovarFromFaktura(ByVal utovarID As String, _
                                     ByVal fakturaID As String)
    Const SRC As String = "ReleaseUtovarFromFaktura"

    Dim rows As Collection
    Set rows = FindRows(TBL_UTOVAR, COL_UT_ID, utovarID)

    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 51, SRC, _
                  "Dupla UtovarID vrednost: " & utovarID
    End If

    Dim rowUt As Long
    rowUt = CLng(rows(1))

    ' Revizija #10 B3: utovar se oslobadja SAMO ako tvrdi bas fakturu
    ' koja se stornira. Korumpirana stavka tudje fakture ne sme da
    ' "oslobodi" utovar validne fakture -- resetom markera bi unistila
    ' vezu FAK-B iako se stornira FAK-A. Neusaglasenost se NE popravlja
    ' ovde (storno tekuce fakture mora da prodje) -- loguje se, a lanac
    ' je prijavljuje kao "faktura neusaglasena".
    Dim d As Variant, tvrdi As String
    d = GetTableData(TBL_UTOVAR)
    tvrdi = Trim$(CStr(nz(d(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)))))
    If StrComp(tvrdi, Trim$(fakturaID), vbTextCompare) <> 0 Then
        LogWarn SRC, "Utovar " & utovarID & " tvrdi fakturu '" & tvrdi & _
                     "', ne '" & Trim$(fakturaID) & _
                     "' -- marker se NE dira (neusaglasena stavka)."
        Exit Sub
    End If

    RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_FAKTURISANO, "", SRC
    RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_FAKTURA_ID, "", SRC
End Sub

' MarkFakturaOrphaned / MarkFakturaStavkeOrphaned su obrisani u S6 koraku 4:
' storno fakturisane prijemnice se odbija (odluka 1), pa je grana osirotelosti
' nedostizna. Korak 5 je nije vratio: put je storno fakture (linije se storniraju,
' prijemnica postaje nefakturisana sama), pa storno prijemnice.

Private Sub ResetNovacFakturaLink(ByVal fakturaID As String)
    Const SRC As String = "ResetNovacFakturaLink"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Sub

    Dim colFakID As Long
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakID))) = Trim$(fakturaID) Then
            RequireUpdateCell TBL_NOVAC, i, COL_NOV_FAKTURA_ID, "", SRC
        End If
    Next i
End Sub

' ResetNovacOtkupLink vise NE ZIVI OVDE. Do 13.09.2026. su postojale DVE kopije:
' privatna ovde (zurnalirana) i javna u modNovac (TIHA). Produkcija je bila
' bezbedna samo zato sto VBA prvo razresava modul-lokalno ime, pa je StornoOtkup
' pogadjao ovu. Prvi modul koji bi pozvao javnu dobio bi nepovratan storno bez
' ijedne poruke -- a vba_check to ne hvata, jer DUPLIKAT gleda dva PUBLIC imena.
' Spojene su u modNovac.ResetNovacOtkupLink, koja sada zurnalira i preskace
' stornirane redove. Poziv ispod se sada razresava tamo.


' ============================================================
' STORNO REVERSA -- nema ga u ovom modulu od 10e-2.
'
' Revers je ambalazni dokument (tblAmbalazaDokument, identitet AmbDokID), a
' storno je kontra-stav u knjizi: modAmbalaza.StornirajAmbDokument_TX, sa
' ekrana Storno kroz modStornoDok.StornoIzvrsi. Stari storno po nogama
' (ReversID, zastavica Stornirano na redu tblAmbalaza) je obrisan zajedno sa
' undo-om i ispravkom koji su ga zvali.
' ============================================================

' ============================================================
' PRIVATE GUARDS / LOW-LEVEL HELPERS
' ============================================================

Private Function RequireStornoAllowed(ByVal tblName As String, _
                                      ByVal recordID As String, _
                                      ByVal idColumn As String, _
                                      ByVal sourceName As String) As Long
    RequireNonBlank tblName, "TableName", sourceName
    RequireNonBlank recordID, "RecordID", sourceName
    RequireNonBlank idColumn, "IdColumn", sourceName

    RequireColumnIndex tblName, idColumn, sourceName
    RequireColumnIndex tblName, COL_STORNIRANO, sourceName

    Dim rows As Collection
    Set rows = FindRows(tblName, idColumn, recordID)

    If rows Is Nothing Then
        Err.Raise ERR_STORNO_BASE + 60, sourceName, _
                  "FindRows je vratio Nothing. Table=" & tblName & _
                  " ID=" & recordID
    End If

    If rows.count = 0 Then
        Err.Raise ERR_STORNO_BASE + 61, sourceName, _
                  "Stavka nije pronadjena. Table=" & tblName & _
                  " ID=" & recordID
    End If

    If rows.count > 1 Then
        Err.Raise ERR_STORNO_BASE + 62, sourceName, _
                  "ID nije jedinstven. Table=" & tblName & _
                  " ID=" & recordID & _
                  " Count=" & CStr(rows.count)
    End If

    Dim rowIndex As Long
    rowIndex = CLng(rows(1))

    Dim data As Variant
    data = GetTableData(tblName)

    If IsEmpty(data) Then
        Err.Raise ERR_STORNO_BASE + 63, sourceName, _
                  "Tabela je prazna posle pronalaska reda. Table=" & tblName
    End If

    Dim colStorno As Long
    colStorno = RequireColumnIndex(tblName, COL_STORNIRANO, sourceName)

    If IsStorniranoValue(data(rowIndex, colStorno)) Then
        Err.Raise ERR_STORNO_BASE + 64, sourceName, _
                  "Ve" & ChrW(263) & " stornirano. Table=" & tblName & _
                  " ID=" & recordID
    End If

    RequireStornoAllowed = rowIndex
End Function

Private Sub MarkRowStornirano(ByVal tblName As String, _
                              ByVal rowIndex As Long, _
                              ByVal sourceName As String)
    RequireUpdateCell tblName, rowIndex, COL_STORNIRANO, STORNO_DA, sourceName
End Sub

' Storno po BROJU dokumenta zahvata SVE aktivne redove tog broja (Klasa I + II
' dele broj). Broj medjutim nije globalno jedinstven: GenerateBrojPrijemnice
' racuna sekvencu po kupcu a x-deo je fiksno "1", pa dva kupca istog dana mogu
' imati "1/ddmmyy". Bez ove provere storno jednog dokumenta tiho stornira i tudji.
'
' Guard: ako aktivni redovi broja pripadaju vise od jednog vlasnika -> greska
' unutar transakcije (rollback, nijedan red nije promenjen). Za normalan slucaj
' (jedan vlasnik, obe klase) ponasanje je nepromenjeno.
'
' vlasnikCols: jedna ili vise kolona koje cine vlasnika (zbirna = VozacID+KupacID,
' jer se broj generise po vozacu a dokument pripada kupcu). Sve moraju postojati --
' RequireColumnIndex, NE fail-open: schema drift ne sme da ugasi safety guard.
'
' Poziva se iz SVAKE putanje koja mutira po broju (wrapperi u modStorno + core
' StornoZbirna + atomic varijante u modStornoFlow). Puni identitetski storno
' (po PK/GeneracijaID) je zaseban paket -- ovo je zastita od tihe destrukcije.
' Broj razlicitih vlasnika pod istim brojem. JEDAN racun za obe kapije ispod.
'
' Izdvojen zato sto se dve kopije istog brojanja neizbezno razidju, a razlika
' izmedju kapija je tacno JEDNA zastavica -- da li se stornirani broje.
Private Function BrojVlasnikaPoBroju(ByVal tblName As String, _
                                     ByVal brojCol As String, _
                                     ByVal broj As String, _
                                     ByVal sourceName As String, _
                                     ByVal ikad As Boolean, _
                                     ByVal vlasnikCols As Variant) As Long
    If UBound(vlasnikCols) < LBound(vlasnikCols) Then
        Err.Raise ERR_STORNO_BASE + 12, sourceName, _
                  "RequireJedanVlasnikPoBroju: nije zadata nijedna vlasnik kolona."
    End If
    BrojVlasnikaPoBroju = VlasniciPoBroju(tblName, brojCol, broj, sourceName, _
                                          ikad, vlasnikCols).count
End Function

Public Sub RequireJedanVlasnikPoBroju(ByVal tblName As String, _
                                      ByVal brojCol As String, _
                                      ByVal broj As String, _
                                      ByVal sourceName As String, _
                                      ParamArray vlasnikCols() As Variant)
    Dim n As Long
    n = BrojVlasnikaPoBroju(tblName, brojCol, broj, sourceName, False, _
                            ScopeColsToArray(vlasnikCols))
    If n > 1 Then
        Err.Raise ERR_STORNO_BASE + 11, sourceName, _
                  "Broj '" & broj & "' nije jedinstven: aktivni dokumenti pripadaju " & _
                  CStr(n) & " razlicita vlasnika. Storno po broju bi " & _
                  "zahvatio i tudji dokument. Storniraj pojedinacno (po ID-u dokumenta) " & _
                  "ili razdvoj brojeve."
    End If
End Sub

' ISTA kapija, jedna jedina razlika: STORNIRAN vlasnik se BROJI.
'
' Storniran vlasnik i dalje ima AKTIVNU decu, pa mutacija po broju zahvati i
' njih. Kapija koja broji samo aktivne to NE vidi: posle storna je aktivan
' jedan, pa broj izgleda jednoznacan. To je zapisana cena iz v6-ui-138.
'
' Ovo je kapija za MUTACIJU PO BROJU. Kad se radi po ID-u dokumenta, ne treba.
Public Sub RequireJedanVlasnikIkadPoBroju(ByVal tblName As String, _
                                          ByVal brojCol As String, _
                                          ByVal broj As String, _
                                          ByVal sourceName As String, _
                                          ParamArray vlasnikCols() As Variant)
    Dim n As Long
    n = BrojVlasnikaPoBroju(tblName, brojCol, broj, sourceName, True, _
                            ScopeColsToArray(vlasnikCols))
    If n > 1 Then
        Err.Raise ERR_STORNO_BASE + 14, sourceName, _
                  "Broj '" & broj & "' nije jedinstven: dokumenti pod njim su IKAD " & _
                  "pripadali " & CStr(n) & " razlicita vlasnika. Racunaju se i " & _
                  "stornirani, jer storniran vlasnik i dalje ima AKTIVNU decu -- " & _
                  "mutacija po broju bi zahvatila i tudje. Radi po ID-u dokumenta " & _
                  "ili razdvoj brojeve."
    End If
End Sub

' Recnik razlicitih VLASNIKA koji pod istim brojem imaju dokument.
'
' Jedan racun za dve upotrebe koje se razlikuju u jednoj jedinoj stvari - da li
' se stornirani redovi broje:
'
'   CILJ prevezivanja / storno  -> samo aktivni (ukljuciStornirane = False)
'   IZVOR prevezivanja          -> i stornirani (True), jer je izvor bas
'                                  storniran dokument; brojanje samo aktivnih
'                                  bi tu uvek dalo nulu i kapija ne bi radila
'
' Vlasnik je KOMPOZITAN i razlicit po tipu dokumenta - isti spisak koji vec
' koristi modDokumenta.ApplyGeneracijaID: otpremnica StanicaID, prijemnica
' KupacID, zbirna VozacID + KupacID.
Public Function VlasniciPoBroju(ByVal tblName As String, ByVal brojCol As String, _
                                ByVal broj As String, ByVal sourceName As String, _
                                ByVal ukljuciStornirane As Boolean, _
                                ByVal vlasnikCols As Variant) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Set VlasniciPoBroju = res

    Dim data As Variant
    data = GetTableData(tblName)
    If IsEmpty(data) Then Exit Function

    Dim cBr As Long, cSt As Long
    cBr = RequireColumnIndex(tblName, brojCol, sourceName)
    cSt = RequireColumnIndex(tblName, COL_STORNIRANO, sourceName)

    Dim cVl() As Long
    ReDim cVl(LBound(vlasnikCols) To UBound(vlasnikCols))
    Dim j As Long
    For j = LBound(vlasnikCols) To UBound(vlasnikCols)
        cVl(j) = RequireColumnIndex(tblName, CStr(vlasnikCols(j)), sourceName)
    Next j

    Dim i As Long, k As String
    For i = 1 To UBound(data, 1)
        ' ZBR-NORM-02: kapija poredi kroz BrojJednak (Trim + vbTextCompare).
        If BrojJednak(data(i, cBr), broj) Then
            If ukljuciStornirane Or Not IsStorniranoValue(data(i, cSt)) Then
                k = ""
                For j = LBound(cVl) To UBound(cVl)
                    k = k & "|" & Trim$(NzToText(data(i, cVl(j))))
                Next j
                If Not res.Exists(k) Then res.Add k, 1
            End If
        End If
    Next i
End Function

' ParamArray se ne prosledjuje dalje kao ParamArray - prepakuje se u obican niz.
Private Function ScopeColsToArray(ByVal p As Variant) As Variant
    Dim out() As Variant, i As Long
    ReDim out(LBound(p) To UBound(p))
    For i = LBound(p) To UBound(p)
        out(i) = p(i)
    Next i
    ScopeColsToArray = out
End Function

Private Sub RequireNonBlank(ByVal value As String, _
                            ByVal fieldName As String, _
                            ByVal sourceName As String)
    If Len(Trim$(value)) = 0 Then
        Err.Raise ERR_STORNO_BASE + 70, sourceName, _
                  fieldName & " je obavezan."
    End If
End Sub

Private Function IsStorniranoValue(ByVal value As Variant) As Boolean
    IsStorniranoValue = (UCase$(Trim$(CStr(value))) = UCase$(STORNO_DA))
End Function

Private Sub LogAndReraise(ByVal sourceName As String)
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
' MONITORING / TX ERROR HANDLING
' ============================================================

Private Sub HandleStornoTxError(ByVal procedureName As String, _
                                ByVal entityType As String, _
                                ByVal entityID As String, _
                                ByRef tx As clsTransaction)
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr procedureName

    If Not tx Is Nothing Then tx.RollbackTx

    Monitor_Error _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=entityID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="STORNO_" & UCase$(entityType) & "_FAIL", _
        severity:="ERROR", _
        message:=entityType & " storno failed. ID=" & entityID & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=entityID

    Debug.Print procedureName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc

    On Error GoTo 0
End Sub

Private Sub MonitorStornoSuccess(ByVal procedureName As String, _
                                 ByVal entityType As String, _
                                 ByVal entityID As String)
    On Error Resume Next

    Monitor_Event _
        eventType:="STORNO_" & UCase$(entityType) & "_SUCCESS", _
        severity:="INFO", _
        message:=entityType & " stornirano. ID=" & entityID, _
        userId:="Operator", _
        moduleName:=MOD_NAME, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=entityID

    On Error GoTo 0
End Sub

