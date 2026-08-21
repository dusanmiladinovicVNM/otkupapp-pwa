Attribute VB_Name = "modScrFakture"
'=====================================================================
' modScrFakture - ekran "Fakturisanje" (v6-ui-176). Faza E, stavka 16.
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da klijent
' kome ovaj modul nedostaje i dalje radi (zamka #19). Red u registru
' (modUiScreens.ScrRows) je postojao i pre ovog modula -- stavka menija se do
' sada crtala prigusena.
'
' ODAKLE DOLAZI: frmFakturisanje bira kupca, ucita njegove prijemnice u
' ListBox sa MultiSelect-om, pa od oznacenih redova napravi fakturu. Mreza
' ljuske bira JEDAN red, pa je multiselect postao KORPA -- red se dugmetom
' dodaje i uklanja, a sta je u korpi se vidi u koloni sa oznakom i u traci
' uz desnu ivicu zone.
'
' STA JE OVDE, A STA NIJE: ovde je REDOSLED i PRIKAZ. Nijedno poslovno
' pravilo, nijedna kapija i nijedan upis nisu ovde:
'   - izrada fakture (transakcija)   -> modFaktura.CreateFaktura_TX
'   - stampa                          -> modFaktura.PrintFaktura
'   - status placanja                 -> modFaktura.UpdateFakturaStatus
'   - redovi mreze                    -> modFaktura.Get*ForGrid
'   - nefakturisane prijemnice kupca  -> modDokumenta.GetPrijemniceByKupac
'   - otvorene fakture kupca          -> modNovac.GetOpenFakture
'   - SEF                             -> modSEFService / modSEFStatusSync
'
' TRI LISTE u deljenoj mrezi (prekidac iznad nje):
'   ZAFAKT   prijemnice izabranog kupca; radnje: dodaj u korpu / ukloni
'   FAKTURE  izdate fakture sa uplatama; radnje: stampaj / osvezi status
'   SEF      stanje elektronskih faktura; pet radnji nad redom
'
' ZASTO "NEPLACENE" NIJE LISTA nego CIP: to je lista FAKTURE sa filterom po
' statusu -- iste kolone, isti citac, isti identitet, iste radnje. Zasebna
' lista bi bila druga kopija istog citaca koja moze da se razidje.
'
' ZASTO SEF JESTE LISTA a ne radnje na listi FAKTURE: MAX_ACT je 5, a lista
' faktura vec nosi dve radnje. Pet SEF operacija bi dalo sedam ukupno, pa bi
' se visak TIHO odsekao (RefreshRowActions radi Exit For) -- operater bi
' dobio ekran kome fali dugme, bez ijedne poruke. Isti kvar je vec placen na
' listi paleta (v6-ui-162).
'
' frmSEF OSTAJE OPERATIVAN i nepromenjen: nosi event log po fakturi,
' PrepareResubmit i batch radnje (RecoverAllStuckSEFSendingInvoices, refresh
' pending) -- nista od toga nije radnja nad JEDNIM redom, pa ovde ne pripada.
'
' POLJA SU LJUSKINA, NE EKRANOVA. Sklop "natpis + shell + kontrola" pravi
' modOtkupUI.NewFieldG, raspored unutar polja modOtkupUI.LayoutFieldInner.
' Kombo u zoni MORA da bude polje (okvir 'nm' + kontrola 'nmT'), a ne gola
' kontrola: panel za izbor (modOtkupUI.FindCombo) trazi bas taj oblik.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRFAK_BUILD As String = "v6-ui-176"

' Visina zone. Manja je od agro zone jer ovaj ekran ima JEDNO polje (kupac) --
' sve ostalo bira mreza, a ne unos.
Private Const FK_ZONA_H   As Single = 148

' Redovi zone (Y koordinate)
Private Const FK_Y_CAP    As Single = 6
Private Const FK_Y_KPI_V  As Single = 18
Private Const FK_Y_LBL    As Single = 48
Private Const FK_Y_HINT   As Single = 98
Private Const FK_Y_BTN    As Single = 116
Private Const FK_BTN_H    As Single = 24
Private Const FK_KPI_W    As Single = 140
Private Const FK_FLD_W    As Single = 260

' DESNA TRAKA ZONE NOSI KORPU -- isti razlog i isti raspored kao na ekranu
' Agrohemija (PRE_DESNO na Paletama): bez nje se sadrzaj korpe vidi samo u
' koloni oznake, i to samo dok je prikazana lista prijemnica.
Private Const FK_KORPA_W  As Single = 300
Private Const FK_KORPA_N  As Long = 4
' Ispod ove sirine bi polje i dugmad ostali pretesni, pa traka nestaje.
Private Const FK_POLJA_MIN As Single = 460

' Kljucevi lista
Private Const FK_ZAFAKT  As String = "ZAFAKT"
Private Const FK_FAKTURE As String = "FAKTURE"
Private Const FK_SEF     As String = "SEF"

' SKRIVENA KOLONA IDENTITETA. Prioritet 4, a LayoutGrid crta do 3 -- vrednost
' postoji u modelu, celija se nikad ne pravi. Identitet ide U RED, ne pored
' njega: mreza redove sortira i deli na strane, pa bi svaka mapa "prikaz -> ID"
' koju ekran drzi sa strane zastarela na prvi klik po zaglavlju.
Private Const FK_ZAF_KOL_ID As Long = 10
' Dostupnost se NE izvodi iz prikaza. Prazna kolona broja fakture nije isto
' sto i "sme u fakturu": red sa Fakturisano="Da" i praznim FakturaID bi se iz
' prikaza citao kao slobodan, a CreateFaktura bi ga odbio. Pravilo racuna
' citac (modFaktura.PrijemnicaDostupna), red ga samo PRENOSI -- takodje
' prioriteta 4, pa se ne crta.
Private Const FK_ZAF_KOL_DOST As Long = 11
Private Const FK_FAK_KOL_ID As Long = 8
Private Const FK_SEF_KOL_ID As Long = 8

Private mLista As String            ' FK_ZAFAKT | FK_FAKTURE | FK_SEF

' KORPA: prijemnice koje cekaju da postanu faktura. Ovo NIJE podatak u tabeli
' nego prolazno stanje ekrana, pa ima svoj kanal ka znacki -- v. KorpaPromenjena.
' Svaka stavka je recnik: prijemnicaID / broj / kolicina / cena / vrednost.
'
' Korpa NE nosi poslovno pravilo i zato ne zivi u domenskom modulu (za razliku
' od agro korpe): CreateFaktura veruje SAMO PrijemnicaID-u i sve ostalo iznova
' izvodi iz tblPrijemnica. Korpa je ovde spisak onoga sto je operater pokazao.
Private mKorpa As Collection
' Kupac za koga je korpa napunjena. CreateFaktura odbija prijemnicu drugog
' kupca (greska 1721), pa korpa ne sme da prezivi promenu kupca.
Private mKorpaKupac As String

Private mCombosPunjeni As Boolean
Private mFill As Boolean            ' punjenje comboa okida Change - v. mPopMute u ljusci
Private mStep As String             ' korak za poruku o gresci

' KupacID -> Array(brojDostupnih, neplacenoUkupno). OsveziZonu se zove pri
' svakom citanju mreze, a obe brojke su pun prolaz kroz tabele. Kes cisti
' Scr_ResetCache, koju ljuska zove posle svakog upisa (RefreshFromData).
Private mKpiKes As Object

' Poslednji broj koji je znacka uz stavku menija dobila. Zona se u testu ne
' crta, pa je ovo jedini nacin da se pravilo "znacka prati korpu i kad korpa
' NIJE prikazana lista" izmeri bez forme.
Private mZnacka As Long

' Kupac koga je test postavio. Zone u testu nema, pa se combo ne moze
' procitati; bez ovoga bi lista prijemnica u svakom testu bila prazna.
' Vazi SAMO u test rezimu -- v. IzabraniKupacID.
Private mKupacTest As String

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=FAKTURE|naslov=OTKUI_NAV_FAKT|sub=OTKUI_SCRFK_SUB" & _
               "|lista=OTKUI_SCRFK_LISTA|oblik=zona+mreza|upis=zona"
End Function

' SEF LISTA POSTOJI UVEK, i na instalaciji koja na SEF nije povezana.
'
' Do prvog smoke-a je bila uslovna (SEFKonfigurisan), i to je bila pogresna
' procena: lista ima dva dela, a kapija je potrebna samo jednom. CITANJE
' stanja (SEFWorkflowState, SEF ID, poslato, greska) su kolone tblFakture --
' ne trebaju im ni baza ni kljuc, i operateru je to legitiman pregled
' ('sta je od mojih faktura poslato'). RADNJE jesu te koje traze podesen SEF,
' i one kapiju vec imaju (SefID -> OTKUI_ERR_FK_SEF_OFF).
'
' Skrivanje cele liste je novi UI cinilo UZIM od legacy-ja: frmFakturisanje
' otvara frmSEF bezuslovno, bez ijedne provere configa.
Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        FK_ZAFAKT & "|OTKUI_SEG_FK_ZAFAKT|OTKUI_GRID_TITLE_FK_ZAFAKT|108", _
        FK_FAKTURE & "|OTKUI_SEG_FK_FAKTURE|OTKUI_GRID_TITLE_FK_FAKTURE|64", _
        FK_SEF & "|OTKUI_SEG_FK_SEF|OTKUI_GRID_TITLE_FK_SEF|44")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = FK_ZAFAKT
    Scr_Lista = mLista
End Function

' Lista prijemnica je uvek lista JEDNOG kupca, pa naslov nosi koga -- inace se
' ne vidi cije se prijemnice gledaju.
Public Function Scr_NaslovDopuna() As String
    Dim naziv As String
    If Scr_Lista() <> FK_ZAFAKT Then Exit Function
    naziv = KupacNaziv()
    If Len(naziv) = 0 Then Exit Function
    Scr_NaslovDopuna = ChrW(8212) & " " & naziv
End Function

' Prvi cip je svuda "sve" -- ljuska na njega pada kad zatecen filter ne pripada
' listi na koju se upravo preslo (RefreshChipsForScreen). Zato prvi mora da
' bude NAJSIRI: povratak na uzi cip bi tiho sakrio redove.
Public Function Scr_Cipovi() As String
    Scr_Cipovi = FkCipoviZaListu(Scr_Lista())
End Function

' Cipovi PO KLJUCU LISTE. Odvojeno od Scr_Cipovi zato sto je Scr_Lista
' gate-ovana konfiguracijom SEF-a: bez ovog ulaza se ugovor SEF liste na
' instalaciji bez SEF-a ne bi mogao izmeriti, a fixture je donor-zavisan.
Public Function FkCipoviZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case FK_ZAFAKT
            FkCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                                "ceka:OTKUI_CIPF_CEKA:104|" & _
                                "fakt:OTKUI_CIPF_FAKT:96"
        Case FK_FAKTURE
            FkCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                                "nepl:OTKUI_CIPF_NEPLACENE:88|" & _
                                "plac:OTKUI_CIPF_PLACENE:76|" & _
                                "godina:OTKUI_CIPA_GODINA:84"
        Case FK_SEF
            FkCipoviZaListu = "sve:OTKUI_CHIP_SVE:40|" & _
                                "zaslanje:OTKUI_CIPF_SEF_ZA:80|" & _
                                "uslanju:OTKUI_CIPF_SEF_U:76|" & _
                                "odbijeno:OTKUI_CIPF_SEF_ODB:80|" & _
                                "greska:OTKUI_CIPF_SEF_GRESKA:70"
    End Select
End Function

' PRAVILA CIPOVA, odvojena od mreze da bi mogla da se izmere bez nje. Kljuc je
' EKRANOV -- ljuska ga je samo vratila onakvog kakvog ga je dobila iz Scr_Cipovi.
' Nepoznat i prazan kljuc PUSTAJU sve: ekran koji dobije filter koji ne poznaje
' pokazuje punu listu, ne praznu.
Public Function FkCipPrijemnica(ByVal filter As String, ByVal dostupna As Boolean) As Boolean
    Select Case filter
        Case "ceka": FkCipPrijemnica = dostupna
        Case "fakt": FkCipPrijemnica = Not dostupna
        Case Else:   FkCipPrijemnica = True
    End Select
End Function

' Faktura je OTVORENA kad je nesto od nje ostalo neplaceno. Isto pravilo koje
' modNovac.GetOpenFakture primenjuje nad jednim kupcem; slaganje sa njim tvrdi
' test, da se dve implementacije ne razidju.
Public Function FkCipFaktura(ByVal filter As String, ByVal status As String, _
                             ByVal iznos As Double, ByVal uplaceno As Double, _
                             ByVal datum As Variant) As Boolean
    Select Case filter
        Case "nepl"
            ' Ista dva uslova koja primenjuje modNovac.GetOpenFakture nad jednim
            ' kupcem: ZAPISAN status Neplaceno I nesto stvarno preostalo.
            FkCipFaktura = (StrComp(Trim$(status), STATUS_NEPLACENO, vbTextCompare) = 0) _
                           And (iznos - uplaceno > 0)
        Case "plac"
            ' Placena je ona koju paypill u istom redu crta kao placenu. Faktura
            ' iznosa 0 ima preostalo 0, ali nije placena nego prazna.
            FkCipFaktura = (FkPayKod(iznos, uplaceno) = PAY_PLACENO)
        Case "godina"
            If IsDate(datum) Then FkCipFaktura = (Year(CDate(datum)) = Year(Date))
        Case Else
            FkCipFaktura = True
    End Select
End Function

' Stanja koja ne padaju ni u jedan uzi cip (SENT, ACCEPTED, STORNO) vide se
' samo pod "sve" -- to su fakture nad kojima vise nema sta da se uradi.
Public Function FkCipSEF(ByVal filter As String, ByVal stanje As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(stanje))
    Select Case filter
        Case "zaslanje"
            FkCipSEF = (s = UCase$(WF_LOCAL_FINALIZED)) Or (s = UCase$(WF_SEF_READY))
        Case "uslanju"
            FkCipSEF = (s = UCase$(WF_SEF_SENDING))
        Case "odbijeno"
            FkCipSEF = (s = UCase$(WF_SEF_REJECTED))
        Case "greska"
            FkCipSEF = (s = UCase$(WF_SEF_TECH_FAILED)) Or _
                       (s = UCase$(WF_SEF_SYNC_ERROR)) Or _
                       (s = UCase$(WF_SEF_UNKNOWN))
        Case Else
            FkCipSEF = True
    End Select
End Function

' Radnje nad izabranim redom. Broj radnji po listi je namerno <= MAX_ACT (5):
' visak se tiho odseca, pa lista SEF-a stoji tacno na granici i to tvrdi test.
Public Function Scr_Radnje() As String
    Scr_Radnje = FkRadnjeZaListu(Scr_Lista())
End Function

' Radnje PO KLJUCU LISTE -- isti razlog kao FkCipoviZaListu.
Public Function FkRadnjeZaListu(ByVal kljuc As String) As String
    Select Case kljuc
        Case FK_ZAFAKT
            FkRadnjeZaListu = "fkadd:OTKUI_BTN_FK_DODAJ:132:soft:1|" & _
                         "fkdel:OTKUI_BTN_FK_UKLONI:124:ghost:1"
        Case FK_FAKTURE
            FkRadnjeZaListu = "fkprint:OTKUI_BTN_FK_STAMPAJ:104:ghost:1|" & _
                         "fkstat:OTKUI_BTN_FK_STATUS:132:soft:1"
        Case FK_SEF
            FkRadnjeZaListu = "sfsend:OTKUI_BTN_FK_SEF_POSALJI:96:primary:1|" & _
                         "sfstat:OTKUI_BTN_FK_SEF_STATUS:112:soft:1|" & _
                         "sfcancel:OTKUI_BTN_FK_SEF_OTKAZI:80:ghost:1|" & _
                         "sfstorno:OTKUI_BTN_FK_SEF_STORNO:80:danger:1|" & _
                         "sfrecov:OTKUI_BTN_FK_SEF_OPORAVI:88:ghost:1"
    End Select
End Function

' Koliko stavki ceka operatera. Korpa je jedino sto na ovom ekranu nije u
' tabeli: bez ove brojke neproknjizena korpa nestane bez traga cim se predje
' na drugi ekran.
Public Function Scr_Brojac() As Long
    If mKorpa Is Nothing Then Exit Function
    Scr_Brojac = mKorpa.count
End Function

Public Sub Scr_ResetCache()
    Set mKpiKes = Nothing
End Sub

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim errDesc As String
    On Error GoTo EH
    Scr_Event = ObradiKlik(tag)
    Err.Clear
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrFakture.Scr_Event"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
    Err.Clear
End Function

'=====================================================================
' KLIKOVI
'=====================================================================
Private Function ObradiKlik(ByVal tag As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        ObradiKlik = True
        Exit Function
    End If

    ' Izbor reda ne menja podatke ni u jednoj listi -- korpa se menja radnjom.
    If Left$(tag, 4) = "row:" Then Exit Function

    ' Promena u polju zone stize kao "chg:<tag kontrole>"; vrednost se cita iz
    ' same kontrole, ne iz taga.
    If Left$(tag, 4) = "chg:" Then
        ObradiKlik = ObradiPromenu(Mid$(tag, 5))
        Exit Function
    End If

    ' Dvoklik PREBACUJE red u korpu i iz nje -- jedan potez umesto trazenja
    ' dugmeta. To je najblizi parnjak legacy multiselect-u, gde se klikom po
    ' redu oznacavalo i odznacavalo.
    If Left$(tag, 4) = "dbl:" Then
        ObradiKlik = PrebaciRed(CLng(val(Mid$(tag, 5))))
        Exit Function
    End If

    ' Radnja nad redom stize kao "act:<kljuc>:<red>" -- ljuska u tag stavlja
    ' i BROJ reda, pa ekran ne mora da pita mrezu koji je red izabran.
    If Left$(tag, 4) = "act:" Then
        ObradiKlik = RadnjaNadRedom(Mid$(tag, 5))
        Exit Function
    End If

    Select Case tag
        Case "scrFkIzradi": ObradiKlik = IzradiFakturu()
        Case "scrFkOcisti": ObradiKlik = IsprazniKorpu()
    End Select
End Function

Private Function ObradiPromenu(ByVal tag As String) As Boolean
    Dim nov As String
    If mFill Then Exit Function
    Select Case tag
        Case "scrFkKupT"
            nov = IzabraniKupacID()
            ' "chg:" stize na SVAKI otkucaj u polju. Posla ima samo kad se
            ' ukucano razresilo u DRUGOG kupca -- inace bi svaki znak povukao
            ' pun prolaz kroz tabele.
            If nov = mKorpaKupac Then Exit Function
            PromeniKupca nov
            ' LJUSKA POVRATNU VREDNOST 'chg:' NE GLEDA -- UiClick zove ScrAct
            ' pa odmah Exit Sub. Ekranu cija lista zavisi od polja zone
            ' osvezavanje mreze zato mora da zatrazi SAM; bez ovoga lista
            ' prijemnica ostane na prethodnom kupcu do sledeceg klika bilo
            ' gde, sto izgleda kao da izbor kupca ne radi. Agrohemija ovo
            ' nema jer nijedna njena lista ne zavisi od polja zone.
            modOtkupUI.RefreshFromData
    End Select
End Function

' Korpa ne prezivljava promenu kupca. CreateFaktura odbija prijemnicu drugog
' kupca (greska 1721), pa bi takva korpa mogla samo da padne pri upisu --
' bolje odmah i uz poruku.
Private Sub PromeniKupca(ByVal nov As String)
    Dim bilo As Long
    If Not mKorpa Is Nothing Then bilo = mKorpa.count
    mKorpaKupac = nov
    If bilo = 0 Then
        OsveziZonu
        Exit Sub
    End If
    Set mKorpa = New Collection
    KorpaPromenjena
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_KORPA_DRUGI_KUPAC"), True
End Sub

Private Function RadnjaNadRedom(ByVal spec As String) As Boolean
    Dim p() As String, red As Long, kljuc As String
    p = Split(spec, ":")
    If UBound(p) < 1 Then Exit Function
    kljuc = p(0)
    red = CLng(val(p(1)))
    If red < 1 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    Select Case kljuc
        Case "fkadd":    RadnjaNadRedom = DodajRedUKorpu(red)
        Case "fkdel":    RadnjaNadRedom = UkloniRedIzKorpe(red)
        Case "fkprint":  RadnjaNadRedom = StampajFakturu(red)
        Case "fkstat":   RadnjaNadRedom = OsveziStatusFakture(red)
        Case "sfsend":   RadnjaNadRedom = SefPosalji(red)
        Case "sfstat":   RadnjaNadRedom = SefOsvezi(red)
        Case "sfcancel": RadnjaNadRedom = SefOtkazi(red)
        Case "sfstorno": RadnjaNadRedom = SefStorno(red)
        Case "sfrecov":  RadnjaNadRedom = SefOporavi(red)
    End Select
End Function

' Identitet iza prikazanog reda. PRAZNO znaci DVOSMISLENO -- dva reda istog
' ID-a u tabeli -- i tada radnja ODBIJA da bira umesto da pogodi. Citac je taj
' koji dvosmislenost prepoznaje (modFaktura.IdIliPrazno); ovde se samo ne
' pogadja. Isto pravilo kao "dvosmislen broj -> MANUAL" u storno okviru.
Private Function IdReda(ByVal red As Long, ByVal kol As Long) As String
    Dim iD As String
    If red < 1 Then Exit Function
    iD = Trim$(CStr(modOtkupUI.GridCell(red, kol)))
    If Len(iD) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_DVOSMISLEN"), True
        Exit Function
    End If
    IdReda = iD
End Function

'=====================================================================
' KORPA
'=====================================================================
Private Function Korpa() As Collection
    If mKorpa Is Nothing Then Set mKorpa = New Collection
    Set Korpa = mKorpa
End Function

' Jedno mesto za obe posledice promene korpe. Ljuska brojace uz stavke menija
' pita SAMO kroz RefreshFromData, a nju zove tek kad ekran na klik javi True =
' "podaci su promenjeni" -- a to ekran javlja samo za listu koja se stvarno
' promenila. "Podaci su promenjeni" i "korpa je promenjena" NISU ista stvar:
' bez ovog kanala bi operater koji gleda listu faktura dodao stavke, a znacka
' bi i dalje pisala nulu.
'
' Cena: OsveziNavBrojace pita SVAKI ekran, a vecina brojaca je prolaz kroz
' tabele. Zato ovo ide na KLIK (Dodaj / Ukloni / Isprazni / Izradi), a nikad iz
' OsveziZonu -- zonu osvezava i svako citanje mreze.
Private Sub KorpaPromenjena()
    mZnacka = Scr_Brojac()
    OsveziZonu
    modOtkupUI.OsveziNavBrojace
End Sub

Private Function UKorpi(ByVal prijemnicaID As String) As Long
    Dim i As Long
    If mKorpa Is Nothing Then Exit Function
    If Len(prijemnicaID) = 0 Then Exit Function
    For i = 1 To mKorpa.count
        If CStr(mKorpa(i)("prijemnicaID")) = prijemnicaID Then
            UKorpi = i
            Exit Function
        End If
    Next i
End Function

' Dodavanje ide preko ID-a, a sve ostale vrednosti reda stizu iz mreze -- iz
' istog reda iz kog je i ID. Sluze SAMO prikazu (traka i zbir); fakturu racuna
' CreateFaktura iznova iz tblPrijemnica.
Public Function FkDodaj(ByVal prijemnicaID As String, ByVal broj As String, _
                        ByVal kolicina As Double, ByVal cena As Double, _
                        ByVal dostupna As Boolean) As String
    Dim red As Object
    If Len(Trim$(prijemnicaID)) = 0 Then
        FkDodaj = Poruka("OTKUI_ERR_FK_DVOSMISLEN")
        Exit Function
    End If
    If Not dostupna Then
        FkDodaj = Poruka("OTKUI_ERR_FK_NIJE_DOSTUPNA")
        Exit Function
    End If
    If UKorpi(prijemnicaID) > 0 Then
        FkDodaj = Poruka("OTKUI_ERR_FK_VEC_U_KORPI")
        Exit Function
    End If
    Set red = CreateObject("Scripting.Dictionary")
    red("prijemnicaID") = prijemnicaID
    red("broj") = broj
    red("kolicina") = kolicina
    red("cena") = cena
    red("vrednost") = kolicina * cena
    Korpa().Add red
End Function

' Uklanjanje po IDENTITETU, ne po prikazu. Vraca True kad je nesto izbaceno.
Public Function FkUkloni(ByVal prijemnicaID As String) As Boolean
    Dim i As Long
    i = UKorpi(prijemnicaID)
    If i = 0 Then Exit Function
    Korpa().Remove i
    FkUkloni = True
End Function

Public Function FkZbirKorpe() As Double
    Dim i As Long, s As Double
    If mKorpa Is Nothing Then Exit Function
    For i = 1 To mKorpa.count
        s = s + CDbl(mKorpa(i)("vrednost"))
    Next i
    FkZbirKorpe = s
End Function

Private Function DodajRedUKorpu(ByVal red As Long) As Boolean
    Dim iD As String, greska As String
    If Scr_Lista() <> FK_ZAFAKT Then Exit Function
    iD = IdReda(red, FK_ZAF_KOL_ID)
    If Len(iD) = 0 Then Exit Function

    greska = FkDodaj(iD, Trim$(CStr(modOtkupUI.GridCell(red, 1))), _
                     RedD(red, 6), RedD(red, 7), RedDostupna(red))
    If Len(greska) > 0 Then
        modOtkupUI.ShowToast greska, True
        Exit Function
    End If
    KorpaPromenjena
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_DODATO"), False
    DodajRedUKorpu = True
End Function

Private Function UkloniRedIzKorpe(ByVal red As Long) As Boolean
    Dim iD As String
    If Scr_Lista() <> FK_ZAFAKT Then Exit Function
    iD = IdReda(red, FK_ZAF_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    If Not FkUkloni(iD) Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_NIJE_U_KORPI"), True
        Exit Function
    End If
    KorpaPromenjena
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_UKLONJENO"), False
    UkloniRedIzKorpe = True
End Function

Private Function PrebaciRed(ByVal red As Long) As Boolean
    Dim iD As String
    If Scr_Lista() <> FK_ZAFAKT Then Exit Function
    iD = IdReda(red, FK_ZAF_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    If UKorpi(iD) > 0 Then
        PrebaciRed = UkloniRedIzKorpe(red)
    Else
        PrebaciRed = DodajRedUKorpu(red)
    End If
End Function

Private Function IsprazniKorpu() As Boolean
    If mKorpa Is Nothing Then Exit Function
    If mKorpa.count = 0 Then Exit Function
    Set mKorpa = New Collection
    KorpaPromenjena
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_KORPA_OCISCENA"), False
    IsprazniKorpu = True
End Function

Private Function RedD(ByVal red As Long, ByVal kol As Long) As Double
    Dim v As Variant
    On Error Resume Next
    v = modOtkupUI.GridCell(red, kol)
    If IsNumeric(v) Then RedD = CDbl(v)
    Err.Clear
End Function

' Sme li red u fakturu. Cita se ono sto je CITAC izracunao i sto red NOSI, ne
' ono sto se u redu vidi -- v. FK_ZAF_KOL_DOST.
Private Function RedDostupna(ByVal red As Long) As Boolean
    RedDostupna = (Trim$(CStr(modOtkupUI.GridCell(red, FK_ZAF_KOL_DOST))) = "1")
End Function

'=====================================================================
' RADNJE NAD FAKTUROM
'=====================================================================
Private Function IzradiFakturu() As Boolean
    Dim kupID As String, stavke As Collection, i As Long
    Dim fakturaID As String, brojFakture As String, zbir As Double

    kupID = IzabraniKupacID()
    If Len(kupID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_NEMA_KUPCA"), True
        Exit Function
    End If
    If mKorpa Is Nothing Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_KORPA_PRAZNA"), True
        Exit Function
    End If
    If mKorpa.count = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_KORPA_PRAZNA"), True
        Exit Function
    End If

    zbir = FkZbirKorpe()
    If MsgBox(Poruka("OTKUI_ASK_FK_IZRADI") & vbCrLf & vbCrLf & _
              KupacNaziv() & vbCrLf & _
              mKorpa.count & " " & Poruka("OTKUI_LBL_AG_KORPA_STAVKI") & _
              "  " & ChrW(183) & "  " & Format$(zbir, "#,##0") & " RSD", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    ' Stavka nosi SAMO PrijemnicaID. CreateFaktura svaku drugu vrednost iznova
    ' izvodi iz tblPrijemnica i eksplicitno veruje samo stavka(0) -- dodatna
    ' polja bi bila mrtav teret koji navodi da se u njih veruje.
    Set stavke = New Collection
    For i = 1 To mKorpa.count
        stavke.Add Array(CStr(mKorpa(i)("prijemnicaID")))
    Next i

    fakturaID = modFaktura.CreateFaktura_TX(kupID, stavke)

    If Len(fakturaID) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_IZRADA"), True
        Exit Function
    End If

    Set mKorpa = New Collection
    Scr_ResetCache
    KorpaPromenjena

    brojFakture = NzToText(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ))
    If Len(brojFakture) = 0 Then brojFakture = fakturaID
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_IZRADJENA") & " " & brojFakture, False
    IzradiFakturu = True
End Function

Private Function StampajFakturu(ByVal red As Long) As Boolean
    Dim iD As String
    iD = IdReda(red, FK_FAK_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    modFaktura.PrintFaktura iD
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_STAMPA"), False
End Function

' Status placanja se ne racuna ovde -- UpdateFakturaStatus poredi uplate sa
' iznosom i sam odlucuje. Vraca True: red se u mrezi promenio.
Private Function OsveziStatusFakture(ByVal red As Long) As Boolean
    Dim iD As String
    iD = IdReda(red, FK_FAK_KOL_ID)
    If Len(iD) = 0 Then Exit Function
    modFaktura.UpdateFakturaStatus iD
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka("OTKUI_MSG_FK_STATUS"), False
    OsveziStatusFakture = True
End Function

'=====================================================================
' RADNJE NAD SEF-om
'
' Sve su pozivi POSTOJECIH javnih funkcija modSEFService / modSEFStatusSync --
' ti moduli se ne diraju. Komentar za otkazivanje i storno trazi se InputBox-om,
' isto kao u frmSEF: to je jedini slobodan tekst koji ove radnje traze, a polje
' u zoni bi imalo smisla za dve od pet radnji.
'=====================================================================
' Lista se vidi uvek, ali se nad njom RADI samo kad je SEF podesen. Ovde je
' jedina kapija -- svih pet radnji prolazi kroz nju.
Private Function SefID(ByVal red As Long) As String
    If Not modFaktura.SEFKonfigurisan() Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_FK_SEF_OFF"), True
        Exit Function
    End If
    SefID = IdReda(red, FK_SEF_KOL_ID)
End Function

' SendInvoiceToSEF_TX baca TIPIZIRANU gresku kad SEF odbije fakturu ili kad
' slanje tehnicki padne (AUD-032a). Stanje je u oba slucaja vec sacuvano, pa se
' ishod hvata i prikazuje kao ishod, a ne kao rusenje ekrana -- isto kao
' frmSEF.btnPosalji. Zato ovde stoji Resume Next, ne EH.
Private Function SefPosalji(ByVal red As Long) As Boolean
    Dim iD As String, submissionID As String, errNo As Long, errDesc As String
    iD = SefID(red)
    If Len(iD) = 0 Then Exit Function
    If MsgBox(Poruka("OTKUI_ASK_FK_SEF_POSALJI") & " " & BrojReda(red), _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Function

    On Error Resume Next
    submissionID = modSEFService.SendInvoiceToSEF_TX(iD)
    errNo = Err.Number
    errDesc = Err.description
    Err.Clear
    On Error GoTo 0

    Scr_ResetCache
    SefPosalji = True
    If errNo = 0 Then
        modOtkupUI.ShowToast modSEFService.SEFSendOutcomeMessage( _
            modSEFPersistance.GetFakturaSEFWorkflowState(iD), submissionID), False
    Else
        modOtkupUI.ShowToast errDesc, True
    End If
End Function

Private Function SefOsvezi(ByVal red As Long) As Boolean
    Dim iD As String, ok As Boolean
    iD = SefID(red)
    If Len(iD) = 0 Then Exit Function
    ok = modSEFStatusSync.RefreshSEFStatus_TX(iD)
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka(IIf(ok, "OTKUI_MSG_FK_SEF_STATUS", _
                                        "OTKUI_ERR_FK_SEF_STATUS")), Not ok
    SefOsvezi = True
End Function

Private Function SefOtkazi(ByVal red As Long) As Boolean
    Dim iD As String, kom As String, ok As Boolean
    iD = SefID(red)
    If Len(iD) = 0 Then Exit Function
    kom = Trim$(InputBox(Poruka("OTKUI_ASK_FK_SEF_OTKAZI_KOM"), APP_NAME))
    If Len(kom) = 0 Then Exit Function
    If MsgBox(Poruka("OTKUI_ASK_FK_SEF_OTKAZI") & " " & BrojReda(red), _
              vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Function
    ok = modSEFService.CancelInvoiceOnSEF_TX(iD, kom)
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka(IIf(ok, "OTKUI_MSG_FK_SEF_OTKAZANO", _
                                        "OTKUI_ERR_FK_SEF_OTKAZANO")), Not ok
    SefOtkazi = True
End Function

Private Function SefStorno(ByVal red As Long) As Boolean
    Dim iD As String, kom As String, brStorno As String, ok As Boolean
    iD = SefID(red)
    If Len(iD) = 0 Then Exit Function
    kom = Trim$(InputBox(Poruka("OTKUI_ASK_FK_SEF_STORNO_KOM"), APP_NAME))
    If Len(kom) = 0 Then Exit Function
    brStorno = Trim$(InputBox(Poruka("OTKUI_ASK_FK_SEF_STORNO_BROJ"), APP_NAME))
    If MsgBox(Poruka("OTKUI_ASK_FK_SEF_STORNO") & " " & BrojReda(red), _
              vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Function
    ok = modSEFService.StornoInvoiceOnSEF_TX(iD, kom, brStorno)
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka(IIf(ok, "OTKUI_MSG_FK_SEF_STORNIRANO", _
                                        "OTKUI_ERR_FK_SEF_STORNIRANO")), Not ok
    SefStorno = True
End Function

' Oporavak fakture koja je ostala u SEF_SENDING: proverava se sta je na SEF-u
' stvarno proslo. Racun radi RecoverStuckSEFSendingInvoice.
Private Function SefOporavi(ByVal red As Long) As Boolean
    Dim iD As String, ok As Boolean
    iD = SefID(red)
    If Len(iD) = 0 Then Exit Function
    ok = modSEFService.RecoverStuckSEFSendingInvoice(iD)
    Scr_ResetCache
    modOtkupUI.ShowToast Poruka(IIf(ok, "OTKUI_MSG_FK_SEF_OPORAVLJENO", _
                                        "OTKUI_ERR_FK_SEF_OPORAVLJENO")), Not ok
    SefOporavi = True
End Function

Private Function BrojReda(ByVal red As Long) As String
    On Error Resume Next
    BrojReda = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    Err.Clear
End Function

'=====================================================================
' REDOVI MREZE
'=====================================================================
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    ' Zona se puni odavde, kao i na ekranima Palete i Agrohemija: gradi se
    ' jednom, a podaci za nju postoje tek kad se lista cita.
    OsveziZonu
    Select Case Scr_Lista()
        Case FK_FAKTURE: Scr_Rows = RedoviFakture(filter, q): Exit Function
        Case FK_SEF:     Scr_Rows = RedoviSEF(filter, q): Exit Function
    End Select
    Scr_Rows = RedoviPrijemnice(filter, q)
End Function

' Opis kolona PO KLJUCU LISTE. Postoji da bi se pravilo "identitet je u redu
' i NE crta se" moglo tvrditi za SVAKU listu, i na instalaciji bez SEF-a.
Public Function FkKoloneZaListu(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case FK_FAKTURE: FkKoloneZaListu = FaktureKolone()
        Case FK_SEF:     FkKoloneZaListu = SEFKolone()
        Case Else:       FkKoloneZaListu = PrijemniceKolone()
    End Select
End Function

Private Function PrazanRezultat(ByVal kolone As Variant) As Variant
    PrazanRezultat = Array(kolone, Empty, 0, 0#, 0#, Array(0, 0, 0))
End Function

'--------------------------------------------------- LISTA: PRIJEMNICE
Private Function PrijemniceKolone() As Variant
    ' Prva kolona se uvek crta kao BROJ dokumenta (StyleGridCell, isBroj).
    ' Poslednja nosi identitet i ima prioritet 4 -- mreza crta do 3.
    PrijemniceKolone = Array( _
        "OTKUI_HD_BROJ||txt|104|1", _
        "OTKUI_HD_OZN||txt|32|1", _
        "OTKUI_HD_BROJ_ZBIRNE||txt|96|3", _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HD_KLASA||txt|48|2", _
        "OTKUI_HDA_KOLICINA||num|86|1", _
        "OTKUI_HD_CENA||rsd|80|2", _
        "OTKUI_HD_VREDNOST||rsd|94|1", _
        "OTKUI_HDF_FAKTURA||txt|0|1", _
        "OTKUI_HDF_PRJID||txt|1|4", _
        "OTKUI_HDF_DOSTUPNA||txt|1|4")
End Function

' Citac vraca 1-bazirano:
'   1 PrijemnicaID | 2 Broj | 3 BrojZbirne | 4 Datum | 5 Klasa | 6 Kolicina
'   7 Cena | 8 Vrednost | 9 Dostupna | 10 BrojFakture
Private Function RedoviPrijemnice(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String, iD As String, kupID As String
    Dim zbirKg As Double, zbirVal As Double, dostupna As Boolean
    On Error GoTo EH
    mStep = "prijemnice"

    kupID = IzabraniKupacID()
    If Len(kupID) = 0 Then
        RedoviPrijemnice = PrazanRezultat(PrijemniceKolone())
        Exit Function
    End If

    src = modFaktura.GetPrijemniceZaFakturisanjeForGrid(kupID)
    If Not IsArray(src) Then
        RedoviPrijemnice = PrazanRezultat(PrijemniceKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 11)
    For i = 1 To UBound(src, 1)
        iD = Trim$(CStr(src(i, 1)))
        dostupna = CBool(src(i, 9))
        If Not FkCipPrijemnica(filter, dostupna) Then GoTo Sledeci
        hay = CStr(src(i, 2)) & "|" & CStr(src(i, 3)) & "|" & CStr(src(i, 10))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i, 2))
        ' Kvacica se racuna iz KORPE, ne iz tabele -- korpa je prolazno stanje.
        outA(n, 2) = IIf(UKorpi(iD) > 0, ChrW(10003), "")
        outA(n, 3) = CStr(src(i, 3))
        outA(n, 4) = src(i, 4)
        outA(n, 5) = CStr(src(i, 5))
        outA(n, 6) = CDbl(src(i, 6))
        outA(n, 7) = CDbl(src(i, 7))
        outA(n, 8) = CDbl(src(i, 8))
        outA(n, 9) = CStr(src(i, 10))
        outA(n, 10) = iD
        outA(n, 11) = IIf(dostupna, "1", "")
        zbirKg = zbirKg + CDbl(src(i, 6))
        zbirVal = zbirVal + CDbl(src(i, 8))
Sledeci:
    Next i

    mStep = "OK"
    RedoviPrijemnice = Array(PrijemniceKolone(), outA, n, zbirKg, zbirVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrFakture.RedoviPrijemnice[" & mStep & "]", Err.description
End Function

'------------------------------------------------------ LISTA: FAKTURE
Private Function FaktureKolone() As Variant
    FaktureKolone = Array( _
        "OTKUI_HD_BROJ||txt|92|1", _
        "OTKUI_HD_DATUM||date|74|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_IZNOS||rsd|104|1", _
        "OTKUI_HD_PLACENO||rsd|96|2", _
        "OTKUI_HD_OSTATAK||rsd|96|1", _
        "OTKUI_HD_STATUS||paypill|92|1", _
        "OTKUI_HDF_FAKID||txt|1|4")
End Function

' Citac vraca 1-bazirano:
'   1 FakturaID | 2 Broj | 3 Datum | 4 KupacNaziv | 5 Iznos | 6 Uplaceno
'   7 Preostalo | 8 Status
Private Function RedoviFakture(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String, zbirVal As Double, preostalo As Double
    On Error GoTo EH
    mStep = "fakture"

    src = modFaktura.GetFaktureForGrid()
    If Not IsArray(src) Then
        RedoviFakture = PrazanRezultat(FaktureKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 8)
    For i = 1 To UBound(src, 1)
        preostalo = CDbl(src(i, 7))
        If Not FkCipFaktura(filter, CStr(src(i, 8)), CDbl(src(i, 5)), _
                            CDbl(src(i, 6)), src(i, 3)) Then GoTo Sledeci
        hay = CStr(src(i, 2)) & "|" & CStr(src(i, 4)) & "|" & CStr(src(i, 8))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i, 2))
        outA(n, 2) = src(i, 3)
        outA(n, 3) = CStr(src(i, 4))
        outA(n, 4) = CDbl(src(i, 5))
        outA(n, 5) = CDbl(src(i, 6))
        outA(n, 6) = preostalo
        ' Status je paypill: ljuska ga crta iz SIFRE, ne iz teksta -- iste tri
        ' vrednosti koje vec koristi lista dokumenata.
        outA(n, 7) = FkPayKod(CDbl(src(i, 5)), CDbl(src(i, 6)))
        outA(n, 8) = Trim$(CStr(src(i, 1)))
        zbirVal = zbirVal + CDbl(src(i, 5))
Sledeci:
    Next i

    mStep = "OK"
    RedoviFakture = Array(FaktureKolone(), outA, n, 0#, zbirVal, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrFakture.RedoviFakture[" & mStep & "]", Err.description
End Function

' Sifra za paypill. Odvojena da bi se izmerila bez mreze.
Public Function FkPayKod(ByVal iznos As Double, ByVal uplaceno As Double) As Long
    If iznos > 0 And uplaceno >= iznos Then
        FkPayKod = PAY_PLACENO
    ElseIf uplaceno > 0 Then
        FkPayKod = PAY_DELIM
    Else
        FkPayKod = PAY_NEPLAC
    End If
End Function

'---------------------------------------------------------- LISTA: SEF
Private Function SEFKolone() As Variant
    SEFKolone = Array( _
        "OTKUI_HD_BROJ||txt|92|1", _
        "OTKUI_HD_PARTNER||part|0|1", _
        "OTKUI_HD_IZNOS||rsd|104|2", _
        "OTKUI_HDF_SEF_STANJE||txt|136|1", _
        "OTKUI_HDF_SEF_ID||txt|112|2", _
        "OTKUI_HDF_SEF_POSLATO||date|84|2", _
        "OTKUI_HDF_SEF_GRESKA||txt|180|3", _
        "OTKUI_HDF_FAKID||txt|1|4")
End Function

' Citac vraca 1-bazirano:
'   1 FakturaID | 2 Broj | 3 KupacNaziv | 4 Iznos | 5 SEFWorkflowState
'   6 SEFDocumentId | 7 SEFSentAt | 8 SEFLastErrorMessage
Private Function RedoviSEF(ByVal filter As String, ByVal q As String) As Variant
    Dim src As Variant, i As Long, n As Long, outA() As Variant
    Dim hay As String, stanje As String
    On Error GoTo EH
    mStep = "sef"

    src = modFaktura.GetFaktureSEFForGrid()
    If Not IsArray(src) Then
        RedoviSEF = PrazanRezultat(SEFKolone())
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To 8)
    For i = 1 To UBound(src, 1)
        stanje = CStr(src(i, 5))
        If Not FkCipSEF(filter, stanje) Then GoTo Sledeci
        hay = CStr(src(i, 2)) & "|" & CStr(src(i, 3)) & "|" & stanje & "|" & CStr(src(i, 6))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(src(i, 2))
        outA(n, 2) = CStr(src(i, 3))
        outA(n, 3) = CDbl(src(i, 4))
        outA(n, 4) = stanje
        outA(n, 5) = CStr(src(i, 6))
        outA(n, 6) = src(i, 7)
        outA(n, 7) = CStr(src(i, 8))
        outA(n, 8) = Trim$(CStr(src(i, 1)))
Sledeci:
    Next i

    mStep = "OK"
    RedoviSEF = Array(SEFKolone(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrFakture.RedoviSEF[" & mStep & "]", Err.description
End Function

'=====================================================================
' ZONA
'=====================================================================
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long

    ' Bela podloga ispod reda polja. Zona je krem, a polja su bela -- bez
    ' podloge se izmedju njih vidi pozadina zone. MORA da bude LABELA, ne
    ' Frame: Frame je prozorska kontrola i crta se IZNAD bezprozorskih bez
    ' obzira na z-order. Napravljena PRVA, labela ostaje ispod svega.
    modUiKit.NewLbl z, "fkBg", "", 0, 0, 100, 10, 8, False, 0, C_WHITE

    modUiKit.NewLbl z, "fkCap", UCase$(Poruka("OTKUI_SCRFK_CAP")), PAD, FK_Y_CAP, _
                    260, 11, TS_MICRO, True, C_MUTED, -1

    ' Cetiri brojke desno -- ono sto je legacy drzao u statusnoj liniji ispod
    ' liste, plus dve koje legacy nije imao (korpa i neplaceno kupca).
    For i = 0 To 3
        modUiKit.NewLbl z, "fkKL" & i, "", 0, FK_Y_CAP, FK_KPI_W, 11, _
                        TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "fkKV" & i, ChrW(8212), 0, FK_Y_KPI_V, FK_KPI_W, 20, _
                        TS_KPI, True, C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i

    ' KORPA U ZONI: naslov, poslednje stavke i zbir. Sadrzaj puni
    ' OsveziKorpuPanel, mesto daje RasporediPolja.
    modUiKit.NewLbl z, "fkKorpaCap", "", 0, FK_Y_LBL, FK_KORPA_W, 11, _
                    TS_MICRO, True, C_MUTED, -1
    For i = 0 To FK_KORPA_N - 1
        modUiKit.NewLbl z, "fkKorpaR" & i, "", 0, FK_Y_LBL + 16 + i * 13, _
                        FK_KORPA_W, 12, TS_META, False, C_FOREST, -1
    Next i
    modUiKit.NewLbl z, "fkKorpaZ", "", 0, FK_Y_LBL + 18 + FK_KORPA_N * 13, _
                    FK_KORPA_W, 13, TS_META, True, C_GREEN, -1

    ' POLJE. Pravi ga ljuska (NewFieldG), ekran mu samo kaze gde stoji.
    ' Prefiks "scr" je OBAVEZAN: bez njega promena teksta ide ljusci, koja o
    ' ovom polju ne zna nista.
    modOtkupUI.NewFieldG z, "scrFkKup", Poruka("OTKUI_FLD_FK_KUPAC"), "cmb", "", _
                         1, False, False, "FK"

    ' BROJA FAKTURE OVDE NEMA, i to je namerno. Broj dodeljuje transakcija
    ' (CreateFaktura sam zove GenerateBrojFakture), operater ga ne bira. Polje
    ' sa "predlogom" koji transakcija ignorise bilo bi prikaz koji se garantovano
    ' razilazi sa upisanim. Broj stize u poruci posle upisa i u listi faktura.

    modUiKit.NewLbl z, "fkHint", "", PAD, FK_Y_HINT, 400, 12, TS_META, False, C_MUTED, -1

    ' "Isprazni korpu" stoji UZ dugme izrade, a ne uz desnu ivicu kao na ekranu
    ' Agrohemija: desnu ivicu ovde drzi traka korpe.
    modUiKit.BtnV z, "scrFkIzradi", Poruka("OTKUI_BTN_FK_IZRADI"), PAD, FK_Y_BTN, _
                  164, FK_BTN_H, "primary"
    modUiKit.BtnV z, "scrFkOcisti", Poruka("OTKUI_BTN_FK_OCISTI"), PAD + 172, FK_Y_BTN, _
                  132, FK_BTN_H, "ghost"

    modUiKit.NewLbl z, "fkLnB", "", 0, FK_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    RasporediPolja z, w
    Scr_Layout = FK_ZONA_H
End Function

Private Sub RasporediPolja(ByVal z As Object, ByVal w As Single)
    Dim i As Long, kx As Single, kxK As Single
    Dim wPolja As Single, korpaVidi As Boolean, capDesno As Single
    On Error Resume Next
    If z Is Nothing Then Exit Sub
    If w < 200 Then Exit Sub

    z.Controls("fkBg").Left = PAD - 10
    z.Controls("fkBg").top = FK_Y_LBL - 8
    z.Controls("fkBg").width = w - 2 * (PAD - 10)
    z.Controls("fkBg").Height = FK_Y_BTN - FK_Y_LBL + 2

    ' Desna traka (korpa) uzima svoje, polje i dugmad dele OSTATAK. Na uskom
    ' ekranu traka nestaje -- bolje bez trake nego sa dugmadima koja se ne vide.
    wPolja = w - FK_KORPA_W - PAD
    korpaVidi = (wPolja >= FK_POLJA_MIN)
    If Not korpaVidi Then wPolja = w
    kxK = w - FK_KORPA_W

    z.Controls("fkKorpaCap").Left = kxK
    z.Controls("fkKorpaCap").Visible = korpaVidi
    z.Controls("fkKorpaZ").Left = kxK
    z.Controls("fkKorpaZ").Visible = korpaVidi
    For i = 0 To FK_KORPA_N - 1
        z.Controls("fkKorpaR" & i).Left = kxK
        z.Controls("fkKorpaR" & i).Visible = korpaVidi
    Next i

    ' Brojke idu uz desnu ivicu; sakriva se ona koja bi nalegla na naslov zone.
    capDesno = PAD + 200
    For i = 0 To 3
        kx = w - PAD - (4 - i) * FK_KPI_W
        z.Controls("fkKL" & i).Left = kx
        z.Controls("fkKV" & i).Left = kx
        z.Controls("fkKL" & i).Visible = (kx > capDesno)
        z.Controls("fkKV" & i).Visible = (kx > capDesno)
    Next i

    PoljeX z, "scrFkKup", PAD, FK_FLD_W, FK_Y_LBL

    ' Objasnjenje se zaustavlja pred trakom -- Label ne prelama, samo istece.
    z.Controls("fkHint").width = wPolja - PAD * 2

    modUiKit.MoveBtn z, "scrFkIzradi", PAD, FK_Y_BTN
    modUiKit.MoveBtn z, "scrFkOcisti", PAD + 172, FK_Y_BTN

    z.Controls("fkLnB").width = w
End Sub

Private Sub PoljeX(ByVal z As Object, ByVal nm As String, ByVal X As Single, _
                   ByVal w As Single, ByVal yLbl As Single)
    On Error Resume Next
    z.Controls(nm).Left = X
    z.Controls(nm).top = yLbl
    z.Controls(nm).width = w
    modOtkupUI.LayoutFieldInner z.Controls(nm)
End Sub

Private Function Zona() As Object
    On Error Resume Next
    Set Zona = modOtkupUI.ScreenZone("FAKTURE")
End Function

Private Function Kontrola(ByVal nm As String) As Object
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Function
    Set Kontrola = z.Controls(nm).Controls(nm & "T")
End Function

Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    PuniCombos
    RasporediPolja z, z.width
    OsveziKorpuPanel z
    OsveziObjasnjenje z
    OsveziBrojke z
End Sub

Private Sub PuniCombos()
    Dim z As Object, CB As Object, mapa As Object, k As Variant
    On Error GoTo EH
    If mCombosPunjeni Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    mFill = True
    mStep = "kupci"
    Set CB = z.Controls("scrFkKup").Controls("scrFkKupT")
    CB.Clear
    CB.ColumnCount = 2
    CB.ColumnWidths = "180 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    Set mapa = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    For Each k In mapa.keys
        CB.AddItem Trim$(CStr(mapa(k)))
        CB.List(CB.ListCount - 1, 1) = CStr(k)
    Next k

    mCombosPunjeni = True
    mFill = False
    Exit Sub
EH:
    mFill = False
    ' Prazan combo bez traga je bio glavni razlog zasto je izgledalo da "nista
    ' nije povezano" -- isto kao u modOtkupUI.FillCombos.
    Debug.Print "modScrFakture.PuniCombos PAO na koraku [" & mStep & "]: " & _
                Err.Number & " " & Err.description
End Sub

Private Function IzabraniKupacID() As String
    Dim c As Object
    If IsTestMode() Then
        If Len(mKupacTest) > 0 Then
            IzabraniKupacID = mKupacTest
            Exit Function
        End If
    End If
    On Error Resume Next
    Set c = Kontrola("scrFkKup")
    If c Is Nothing Then Exit Function
    IzabraniKupacID = GetComboID(c)
    Err.Clear
End Function

Private Function KupacNaziv() As String
    Dim c As Object
    On Error Resume Next
    Set c = Kontrola("scrFkKup")
    If c Is Nothing Then Exit Function
    KupacNaziv = Trim$(CStr(c.value))
    Err.Clear
End Function

'------------------------------------------------------- PANEL KORPE
Private Sub OsveziKorpuPanel(ByVal z As Object)
    Dim i As Long, n As Long
    On Error Resume Next
    n = Scr_Brojac()
    z.Controls("fkKorpaCap").caption = UCase$(Poruka("OTKUI_LBL_FK_KORPA_CAP"))
    For i = 0 To FK_KORPA_N - 1
        z.Controls("fkKorpaR" & i).caption = TrakaRed(i)
    Next i
    If n = 0 Then
        z.Controls("fkKorpaZ").caption = Poruka("OTKUI_LBL_AG_KORPA_PRAZNA")
    Else
        z.Controls("fkKorpaZ").caption = n & " " & _
            Poruka("OTKUI_LBL_AG_KORPA_STAVKI") & "   " & ChrW(183) & "   " & _
            Format$(FkZbirKorpe(), "#,##0") & " RSD"
    End If
End Sub

' Tekst reda trake. NAJNOVIJE PRVO: operater upravo nesto doda, pa mu je
' potvrda ono sto trazi. PRELIV SE PRIJAVLJUJE: lista koja se tiho odseca
' izgleda kao cela -- isto pravilo koje ljuska nad sobom vec ima (BazenStaje).
' Racun je odvojen od crtanja, pa se meri bez forme.
Private Function TrakaRed(ByVal i As Long) As String
    Dim n As Long, sakriveno As Long
    If mKorpa Is Nothing Then Exit Function
    n = mKorpa.count
    If n = 0 Then Exit Function
    If i < 0 Or i > FK_KORPA_N - 1 Then Exit Function

    ' Sve staje: samo obrni redosled.
    If n <= FK_KORPA_N Then
        If i > n - 1 Then Exit Function
        TrakaRed = KorpaRedPrikaz(n - i)
        Exit Function
    End If

    ' Ne staje: poslednji red je prelivni.
    If i < FK_KORPA_N - 1 Then
        TrakaRed = KorpaRedPrikaz(n - i)
        Exit Function
    End If
    sakriveno = n - (FK_KORPA_N - 1)
    TrakaRed = ChrW(8230) & " " & Poruka("OTKUI_LBL_AG_KORPA_JOS") & " " & sakriveno
End Function

Private Function KorpaRedPrikaz(ByVal i As Long) As String
    Dim red As Object
    On Error Resume Next
    If mKorpa Is Nothing Then Exit Function
    If i < 1 Or i > mKorpa.count Then Exit Function
    Set red = mKorpa(i)
    KorpaRedPrikaz = CStr(red("broj")) & "   " & ChrW(183) & "   " & _
                     Format$(CDbl(red("vrednost")), "#,##0")
End Function

'---------------------------------------------------------- BROJKE
Private Sub OsveziObjasnjenje(ByVal z As Object)
    On Error Resume Next
    If Len(IzabraniKupacID()) = 0 Then
        z.Controls("fkHint").caption = Poruka("OTKUI_LBL_FK_HINT_KUPAC")
    Else
        z.Controls("fkHint").caption = Poruka("OTKUI_LBL_FK_HINT")
    End If
End Sub

Private Sub OsveziBrojke(ByVal z As Object)
    Dim kpi As Variant
    On Error Resume Next
    kpi = KpiZaKupca(IzabraniKupacID())

    z.Controls("fkKL0").caption = UCase$(Poruka("OTKUI_KPI_FK_CEKA"))
    z.Controls("fkKV0").caption = CStr(CLng(kpi(0)))
    z.Controls("fkKL1").caption = UCase$(Poruka("OTKUI_KPI_FK_KORPA"))
    z.Controls("fkKV1").caption = CStr(Scr_Brojac())
    z.Controls("fkKL2").caption = UCase$(Poruka("OTKUI_KPI_FK_IZNOS"))
    z.Controls("fkKV2").caption = Format$(FkZbirKorpe(), "#,##0")
    z.Controls("fkKL3").caption = UCase$(Poruka("OTKUI_KPI_FK_NEPLACENO"))
    z.Controls("fkKV3").caption = Format$(CDbl(kpi(1)), "#,##0")
End Sub

' Broj prijemnica koje cekaju fakturu i ukupno neotplaceno -- oboje za JEDNOG
' kupca. Neplaceno racuna modNovac.GetOpenFakture: to je jedini read-model
' otvorenih faktura kupca (izbacuje stornirane, trazi status Neplaceno, racuna
' stvarno uplaceno) i ovde se NE ponavlja.
Private Function KpiZaKupca(ByVal kupID As String) As Variant
    Dim src As Variant, i As Long, ceka As Long, nepl As Double
    On Error GoTo EH
    KpiZaKupca = Array(0, 0#)
    If Len(kupID) = 0 Then Exit Function
    If mKpiKes Is Nothing Then Set mKpiKes = CreateObject("Scripting.Dictionary")
    If mKpiKes.Exists(kupID) Then
        KpiZaKupca = mKpiKes(kupID)
        Exit Function
    End If

    src = modFaktura.GetPrijemniceZaFakturisanjeForGrid(kupID)
    If IsArray(src) Then
        For i = 1 To UBound(src, 1)
            If CBool(src(i, 9)) Then ceka = ceka + 1
        Next i
    End If

    src = GetOpenFakture(kupID)
    If IsArray(src) Then
        For i = 1 To UBound(src, 1)
            nepl = nepl + CDbl(src(i, 5))
        Next i
    End If

    KpiZaKupca = Array(ceka, nepl)
    mKpiKes(kupID) = KpiZaKupca
    Exit Function
EH:
    KpiZaKupca = Array(0, 0#)
End Function

'=====================================================================
' TEST SEAM
' Zona se u testu ne crta (forma se ne prikazuje), pa se stanje ekrana ne moze
' procitati iz kontrola. Isti oblik i ista kapija kao Scr_*Test u modScrAgro:
' seam koji MENJA stanje ekrana van test-rezima ne radi nista -- pozvan iz
' liste makroa bi inace tiho bacio operateru neproknjizenu korpu.
'=====================================================================
Public Sub Scr_FkListaTestSet(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mLista = kljuc
End Sub

Public Sub Scr_FkKupacTestSet(ByVal kupacID As String)
    If Not IsTestMode() Then Exit Sub
    mKupacTest = kupacID
    Scr_ResetCache
End Sub

Public Function Scr_FkKorpaBroj() As Long
    Scr_FkKorpaBroj = Scr_Brojac()
End Function

' Dodaj u korpu bez mreze -- ide kroz ISTU rutinu kao radnja nad redom.
' Sta seam NE pokriva: da je bas DodajRedUKorpu ta koja je zove. To se vidi u
' kodu (na svim mestima gde se korpa menja stoji KorpaPromenjena, nigde goli
' OsveziZonu) i mereno je sabotazom.
Public Function Scr_FkKorpaTestDodaj(ByVal prijemnicaID As String, _
                                   ByVal broj As String, _
                                   ByVal kolicina As Double, _
                                   ByVal cena As Double, _
                                   ByVal dostupna As Boolean) As String
    If Not IsTestMode() Then Exit Function
    Scr_FkKorpaTestDodaj = FkDodaj(prijemnicaID, broj, kolicina, cena, dostupna)
    If Len(Scr_FkKorpaTestDodaj) = 0 Then KorpaPromenjena
End Function

Public Function Scr_FkUkloniStavkuTest(ByVal prijemnicaID As String) As Boolean
    If Not IsTestMode() Then Exit Function
    Scr_FkUkloniStavkuTest = FkUkloni(prijemnicaID)
    If Scr_FkUkloniStavkuTest Then KorpaPromenjena
End Function

' Identitet i-te stavke korpe (1-bazirano).
Public Function Scr_FkStavkaIdTest(ByVal i As Long) As String
    On Error Resume Next
    If mKorpa Is Nothing Then Exit Function
    If i < 1 Or i > mKorpa.count Then Exit Function
    Scr_FkStavkaIdTest = CStr(mKorpa(i)("prijemnicaID"))
End Function

' Poslednji broj koji je znacka dobila. Vidi KorpaPromenjena.
Public Function Scr_FkZnackaTest() As Long
    Scr_FkZnackaTest = mZnacka
End Function

Public Function Scr_FkTrakaRedTest(ByVal i As Long) As String
    If Not IsTestMode() Then Exit Function
    Scr_FkTrakaRedTest = TrakaRed(i)
End Function

Public Sub Scr_FkKorpaTestReset()
    If Not IsTestMode() Then Exit Sub
    Set mKorpa = New Collection
    mKorpaKupac = ""
    mKupacTest = ""
    Scr_ResetCache
    KorpaPromenjena
End Sub
