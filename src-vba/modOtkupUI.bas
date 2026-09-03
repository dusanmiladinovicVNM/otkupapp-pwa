Attribute VB_Name = "modOtkupUI"
'Attribute VB_Name = "modOtkupUI"
'=====================================================================
' modOtkupUI - ekran "Otkup i prodaja", v6 (prezentacioni sloj prepravljen).
'
' Sta je novo u odnosu na v5 -----------------------------------------
' 1) NATIVNE KONTROLE SU "OBUCENE". TextBox/ComboBox vise ne crtaju svoju
'    ivicu; oko njih ide SHELL (dva Label-a: ivica + ispuna), a kontrola je
'    uvucena za INPUT_PAD. Tako se dobija pravi unutrasnji padding koji
'    MSForms nema, mirna 1px ivica i jasan focus/error state (ShellState).
'    Nativno dugme ComboBox-a maskira NewChevron (Label sa ZOrder=front);
'    klik na njega zove cmb.DropDown, pa se ponasanje ne menja.
' 2) MREZA VISE NIJE ListBox. Redovi su crtani Label-ovima (MAX_ROWS x 6),
'    naprave se JEDNOM a po renderu se samo pune. Time dobijamo status-pilule
'    u boji, hover na redu, tacan padding po koloni i fiksnu visinu reda -
'    nista od toga MSForms ListBox ne moze.
' 3) TIPOGRAFSKA SKALA (TS_*) - jedan skup velicina za ceo ekran; nema vise
'    ad-hoc brojeva po pozivima.
' 4) UJEDNACENE KOMPONENTE - svako dugme, cip, polje i kartica prolaze kroz
'    isti primitiv (NewShell / BtnV / ChipV / CardV), pa su ivice, padding i
'    stanja identicni svuda.
'
' Nepromenjeno iz v5: FillGrid i ceo data put, self-update higijena
' (OtkupUI_Release / StopOtkupUITimers), perf rad (zone se grade nevidljive,
' formatiranje samo za vidljivu stranu, indeksi kolona razreseni jednom).
'
' Modul MORA ostati "mek" za modSelfUpdate: nijedna MODULE-LEVEL deklaracija
' ne sme biti "As MSForms.*" ni "WithEvents" (IsHardModuleBody, modSelfUpdate
' .bas:902). Procedure-level "As MSForms.ComboBox" je dozvoljeno.
'
' Fajl mora ostati 100% ASCII; tekst ide kroz modPoruke.Poruka("OTKUI_*").
'=====================================================================
Option Explicit

'--- paleta ----------------------------------------------------------
' PAZI: VBA Long boja je &HBBGGRR, obrnuto od CSS-a. Ovde je stajalo &H140D1E
' (= #1E0D14, tamna sljiva) umesto &H142D1E - zelena komponenta 0D umesto 2D.
' Kod se uredno kompajlira, a zaglavlje, ploca VREDNOST, aktivna kartica i
' aktivna stavka sidebara su bili braon umesto forest zelene. Provera je sada
' u _colors.py: svaka C_* konstanta se poredi sa heksom u svom komentaru.
Public Const C_FOREST      As Long = &H142D1E    ' #1E2D14
Public Const C_FOREST_DK   As Long = &HD2014     ' #14200D
Public Const C_CREAM       As Long = &HEEF4F7    ' #F7F4EE
' Ploca alatke NA TAMNOM zaglavlju. Bez nje su "Otvori Excel" i "Snimi radnu
' svesku" bili tekst boje krem na skoro istoj podlozi - operater ih prosto
' nije video. Ivica je jos svetlija, da se dugme cita kao dugme.
Public Const C_HDR_SURF    As Long = &H26422E    ' #2E4226
Public Const C_HDR_EDGE    As Long = &H405C4A    ' #4A5C40
Public Const C_SAND        As Long = &HE1EBEF    ' #EFEBE1
Public Const C_SAND_LT     As Long = &HF5F9FB    ' #FBF9F5
Public Const C_GREEN       As Long = &H2A5E1A    ' #1A5E2A
' Primarno dugme je #1A5E2A - ISTO sto i C_GREEN. Ranije je ovde stajalo
' #2E4726 iz modTheme (BTN_BG); to je ispravno ZA STARE FORME, gde je dugme
' delilo vrednost sa tekstom #1E2D14. Ovaj ekran nikad nije imao taj problem
' (tekst je forest, dugme zeleno), pa je preuzimanje te vrednosti samo odvelo
' dugme i od teme i od vizije. Hover je svetliji, to modTheme nema.
Public Const C_BTN_HOVER   As Long = &H38684A    ' #4A6838 hover primarnog dugmeta
Public Const C_SOFT_BG     As Long = &HECF5F0    ' #F0F5EC svetla zelena podloga
Public Const C_GREEN_LT    As Long = &H35A15E    ' #5EA135
Public Const C_GOLD        As Long = &H4BA8C8    ' #C8A84B
Public Const C_RUST        As Long = &H2D52A0    ' #A0522D
Public Const C_WHITE       As Long = &HFFFFFF
Public Const C_BORDER      As Long = &HCED8D2    ' #D2D8CE
Public Const C_BORDER_LT   As Long = &HE2E8DE    ' mirnija ivica u mirovanju
' Ivica POLJA je zasebna od ivice linija i okvira: #D2D8CE je na belom polju
' pretanka i jedva se vidi, pa polja nose topliju sage ivicu. Linije sekcija,
' zaglavlje mreze i ivice kartica ostaju na C_BORDER / C_BORDER_LT.
Public Const C_INPUT_BORDER As Long = &H96B0A8   ' #A8B096
Public Const C_HEAD_BG     As Long = &HEDF2F4    ' #F4F2ED zaglavlje mreze
Public Const C_MUTED       As Long = &H667A6E    ' #6E7A66
Public Const C_KPI_OK      As Long = &HF0F7F3    ' #F3F7F0
Public Const C_DISABLED    As Long = &HEAEFF1    ' #F1EFEA
Public Const C_DISABLED_FG As Long = &HA4AFA8    ' #A8AFA4
Public Const C_ICON_OFF    As Long = &H84958A    ' #8A9584
Public Const C_ROW_ALT     As Long = &HFCFDFB    ' zebra red
Public Const C_ROW_HOVER   As Long = &HF2F7F0    ' hover red
Public Const C_ROW_LN      As Long = &HDEE5E8    ' #E8E5DE linija izmedju redova
Public Const C_AMBER       As Long = &H1B89C8    ' #C8891B delimicno placeno

' status-pilule (ono sto ListBox nije mogao)
Public Const C_PILL_OK_BG  As Long = &HE9F3ED    ' #EDF3E9
Public Const C_PILL_OK_FG  As Long = &H2A5E1A
Public Const C_PILL_SND_BG As Long = &HF7F2ED    ' #EDF2F7
Public Const C_PILL_SND_FG As Long = &H8C6B3E    ' #3E6B8C celicno plava - "Poslato"
Public Const C_PILL_ERR_BG As Long = &HEAF0FA    ' #FAF0EA
Public Const C_PILL_ERR_FG As Long = &H2D52A0

'--- tipografija ------------------------------------------------------
Public Const F_UI          As String = "Segoe UI"
Public Const F_NUM         As String = "Consolas"
Public Const F_DISPLAY     As String = "Cormorant Garamond"
Public Const F_DISPLAY_FB  As String = "Segoe UI Semibold"
Public Const F_ICON        As String = "Segoe MDL2 Assets"

' Sidebar ide na velicinu naslova mreze (TS_H1), a ne na TS_META - to je glavna
' navigacija, cita se iz daljine i ne sme biti sitnija od sadrzaja.
Public Const TS_NAV       As Single = 10.5    ' = TS_H1, kao "Postojece otpremnice"
Public Const TS_NAVGRP    As Single = 8.5     ' naslov grupe (OPERACIJE...)
Public Const TS_NAVICO    As Single = 13      ' glif uz stavku

' Pecat verzije - DiagOtkupUI ga ispisuje, pa se odmah vidi da li je u projektu
' uvezen pravi fajl (a ne neka ranija kopija).
Public Const OTKUI_BUILD   As String = "v6-ui-204"
' Ekran na kome se aplikacija otvara. Na jednom mestu, jer ga traze i gradnja i
' provera prava pri otvaranju (ShowOtkupUI).
Public Const SCR_POCETNI   As String = "DOKUMENTI"
' Najstarija verzija pratecih modula sa kojom ova ljuska radi. Nije jednakost:
' moduli se menjaju nezavisno (clsFlatBtn je namerno zamrznut), pa bi zahtev
' "sve verzije iste" prijavljivao neispravnu instalaciju i kad je ispravna.
' Podici SAMO kad se ugovor izmedju modula stvarno promeni.
Public Const OTKUI_MIN_BUILD As String = "v6-ui-107"

'--- TIPOGRAFSKA SKALA -----------------------------------------------
' Jedan izvor istine za velicine. Ako neka velicina nije ovde, ne koristi se.
Public Const TS_DISPLAY As Single = 17      ' naslov ekrana
Public Const TS_KPI     As Single = 14      ' KPI vrednost
Public Const TS_VAL     As Single = 16      ' iznos u ploci VREDNOST
Public Const TS_H1      As Single = 10.5    ' naziv kartice, naslov mreze
Public Const TS_ACTION  As Single = 10      ' primarno dugme
Public Const TS_BODY    As Single = 10      ' unos i redovi mreze
Public Const TS_META    As Single = 8.5     ' statusna traka, hint, cip
Public Const TS_LABEL   As Single = 8       ' label iznad polja, sekcija
Public Const TS_MICRO   As Single = 7.5     ' KPI caption, podnozje, eyebrow
' Mreza je gusca od forme: brojevi 8.5, datum 8, partner 9. Ispod 7.5 se ne ide -
' na 125% skaliranju postaje neCitko.
Public Const TS_CELL    As Single = 8.5     ' brojevi i kratki tekst u mrezi
Public Const TS_CELL_SM As Single = 8       ' datum, placeno
Public Const TS_CELL_LG As Single = 9       ' partner

'--- geometrija (points) - vise vazduha nego u v5 --------------------
Private Const HEADER_H    As Single = 42
Private Const SIDEBAR_W   As Single = 152
Private Const SIDEBAR_MIN As Single = 36
Public Const KPI_H       As Single = 56      ' tri reda: naslov, vrednost, kontekst
Private Const TITLE_H     As Single = 40
Private Const OTP_H       As Single = 50      ' traka otpremnice (F1)
' Zona konteksta ima isti ritam kao grupa polja u formi: eyebrow (2..13),
' pa natpis polja (18..30), pa polje (32..58). Sa 14 / 26 su se eyebrow
' "OSNOVNI PODACI" i natpis "VRSTA" preklapali za 5pt i citali se kao jedan
' zbijen blok. Zona zato raste za 2pt, ostatak je preraspodela unutar nje.
Private Const CTX_H       As Single = 60      ' eyebrow red + natpis + polje
Private Const CTX_LBL_Y   As Single = 18      ' natpis IZNAD polja
Private Const CTX_FLD_Y   As Single = 32      ' vrh polja u zoni konteksta
Private Const RIGHT_W     As Single = 188
' 28, a ne 26: na 26 su kutije za unos izgledale stisnuto (tekst od 20pt u
' kutiji od 26 ostavlja po 4pt gore i dole, ali je sama kutija bila niza od
' dugmeta pored nje). 28 je tacno visina dugmeta u akcionom redu, pa se ploca
' iznosa i dugme poklapaju bez pomeranja za pola tacke.
Private Const FIELD_H     As Single = 28
' natpis (0..12) + kutija (16..16+FIELD_H) + 2pt vazduha
Private Const FIELD_GRP_H As Single = 46
Private Const INPUT_PAD   As Single = 9      ' unutrasnji padding teksta
Private Const NAV_H       As Single = 27     ' /3, i prima TS_NAV od 10.5
' 21, a NE 19: visina reda mora biti deljiva sa 3. MSForms racuna u tackama, a
' Windows crta u pikselima - na 96 dpi je 1pt = 4/3 px, pa je red od 19pt tacno
' 25,333 px. Redovi tada padaju na razlicite podpiksel faze (0 / 0,33 / 0,67 px)
' i GDI im rasterizuje tekst u tri razlicita oblika, sto se cita kao "svaki
' drugi red ima drugi font". 21pt = 28 px tacno (i 35 px na 120 dpi, 42 px na
' 144 dpi), pa svi redovi imaju identican raster.
Private Const GRID_ROW_H  As Single = 21
' Ceo vertikalni niz mreze mora biti u koracima od 3pt (= ceo broj piksela na
' 96/120/144 dpi). Zaglavlje je zato 21, ne 20: sa 20 je telo mreze pocinjalo na
' 82pt = 109,33 px, pa je frame mreze bio u podpiksel fazi i prvih par redova se
' rasterizovalo drugacije od ostalih.
Private Const GRID_TOP    As Single = 57     ' naslov + pretraga + cipovi
Private Const GRID_HEAD_H As Single = 21
Private Const GRID_FOOT_H As Single = 24

' Koliko NOVCANIH SLOTOVA podnozje ume da nacrta. Lista izvoda trazi dva
' (uplate i isplate), jer je njihov zbir -- promet -- operateru manje koristan
' od dve brojke koje ume da uporedi sa izvodom u ruci.
'
' Bazen je konacan kao i svaki drugi u ljusci: ekran koji zatrazi vise ne gresi,
' ali se visak NE crta i BazenStaje to prijavi. Tiho odsecanje je vec dvaput
' kostalo pun krug (jedanaesti cip, sesta radnja).
Private Const MAX_FT_VAL  As Long = 2
' Traka poruke iznad podnozja mreze. Ista je na svim ekranima -- poruka nije
' deo liste pa ne sme da se menja sa njom.
Private Const TOAST_H     As Single = 26
' Devet, zbog F8: tamo prekidac bira TIP dokumenta, a tipova je devet
' (sedam rezima F1..F7 + fakture + izvodi). Ostali ekrani koriste manje i
' visak dugmadi ostaje sakriven (LayoutGrid). Natpisi u F8 su kratki bas
' zato da devet komada stane levo od polja za pretragu.
' Dugmadi prekidaca lista koje se PRAVE. Ekran Storno ima DESET lista (devet
' tipova + navigacioni "Svi"), pa je granica podignuta sa 9 -- na 9 je deseta
' tiho nestajala: LayoutGrid crta samo prvih MAX_SEG, bez ijedne poruke, pa se
' "Izvodi" nisu mogli izabrati ni na koji nacin.
' Bazeni ljuske. JAVNI su zato sto se o njima TVRDI: ekran koji zatrazi vise
' nego sto bazen ima izgubi visak, pa broj mora da bude dohvatljiv testu.
Public Const MAX_SEG      As Long = 11
' JAVNA iz istog razloga kao MAX_CHIP: lista koja prijavi vise radnji nego sto
' ima dugmadi izgubi visak BEZ ijedne poruke (RefreshRowActions radi Exit For).
' Tako je sesta radnja na listi paleta tiho izbacila 'Nepotpune palete'.
Public Const MAX_ACT      As Long = 5        ' dugmadi radnji nad redom koje se PRAVE
Private Const MAX_ROWS    As Long = 22       ' redova mreze koji se PRAVE
Public Const MAX_COLS     As Long = 14       ' kolona mreze koje se PRAVE

' Sifre stanja placanja (odvojene od 0/1/2 koje nosi STATUS pilula).
' MORAJU biti u deklaracionom delu: VBA module-level Const iza prve procedure
' nije vidljiv procedurama iznad njega ("Variable not defined").
Public Const PAY_PLACENO As Long = 10
Public Const PAY_DELIM   As Long = 11
Public Const PAY_NEPLAC  As Long = 12
Public Const PAY_NEFAKT  As Long = 13
Private Const POP_MAX     As Long = 14       ' stavki u sopstvenom dropdown-u
Private Const POP_ITEM_H  As Single = 21
Private Const FLT_W       As Single = 236     ' panel "Filteri"
Private Const FLT_H       As Single = 222
Private Const TIT_ICO_W   As Single = 26      ' marker modula uz naslov
' Eyebrow red grupe polja. Sam natpis je visok ~11pt, pa 18 ostavlja 7pt do
' natpisa prvog polja ispod. Sa 15 je taj razmak bio 4pt - naslov grupe je bio
' blizi redu IZNAD sebe nego svojim poljima, pa je delovao kao da pripada
' prethodnoj grupi. Razlika je uzeta iz vazduha izmedju grupa (3 -> 0), koji
' je radio isti posao sa pogresne strane natpisa; forma po visini ostaje ista.
Private Const GRP_H       As Single = 18
' Vertikalni razmak izmedju redova polja. Do sada je i tu isao GAP (10),
' isti koji razdvaja kolone - a po visini je 10pt cista rasipnja: tri reda
' polja i tri grupe su tako uzimali 40pt koje je mreza trazila.
' 7 je kompromis: sa 6 (i kutijom od 26) je razmak izmedju dna kutije i
' sledeceg natpisa padao na 6pt i unosni deo je izgledao zbijeno; sa 7 i
' kutijom od 28 taj razmak je 9pt, a mreza gubi samo 3pt po redu polja.
Private Const ROW_GAP     As Single = 7
Private Const GRP_LBL_W   As Single = 108     ' natpis grupe, pa linija do kraja
Private Const SEG_KL_W    As Single = 30      ' prekidac klase u polju KLASA I CENA
Private Const VAL_CALC_W  As Single = 150     ' racunica DESNO od ploce iznosa
Private Const VAL_KG_W    As Single = 250     ' kilogrami desno od racunice
Private Const VAL_PLATE_W As Single = 236     ' ploca iznosa (natpis + broj + RSD)
Private Const GAP         As Single = 10
Public Const ICO_INSET   As Single = 24     ' sirina zone ikonice u dugmetu
Public Const PAD         As Single = 16
Private Const STATUS_H    As Single = 24
' Donji blok sidebara (profil + podnozje) je prikovan za dno i NE deli visinu
' sa stavkama. Imenovan je zato sto ulazi u meru na kojoj stoji odluka o
' sekcijama (docs/UI_MIGRACIJA_KATALOG.md, 24.1) -- gola 62 u LayoutOtkup se
' ne bi mogla tvrditi u testu.
Private Const PROFIL_H    As Single = 62
Private Const MIN_H       As Single = 620
Private Const MIN_W       As Single = 900
Private Const BP_NARROW   As Single = 860

'--- MDL2 kodne tacke (Long sufiks & - iznad Integer opsega) ---------
' Komentar uz svaki kod je ZVANICNO ime glifa iz Segoe MDL2 Assets, a ne opis
' po izgledu. Prethodni set je bio pogodjen po opisu i vise njih je bilo
' promaseno (IC_PALETE je bio E8C0 = CalendarWeek, ne mreza; IC_OTKUP E7AC =
' OpenWith). Kad menjas kod, upisi ime iz dokumentacije - tako se promasaj vidi
' u kodu, bez pokretanja Excela.
' Svi kodovi su iz opsega ispod EC00 (originalni Windows 10 glifovi): noviji
' opsezi (F1xx+) fale na starijim verzijama fonta i iscrtaju se kao pravougaonik.
' Sidebar
Public Const IC_OTKUP    As Long = &HE70B&   ' QuickNote
Public Const IC_BLOKOVI  As Long = &HE8A5&   ' Document
Public Const IC_PALETE   As Long = &HE8F1&   ' Library
Public Const IC_AGRO     As Long = &HE8BE&   ' Leaf
Public Const IC_FAKT     As Long = &HE8C7&   ' PaymentCard
Private Const IC_BANKA    As Long = &HE825&   ' Bank
Public Const IC_UVOZ     As Long = &HEA90&   ' PDF - izvodi su PDF fajlovi
Public Const IC_NALOZI   As Long = &HE9D5&   ' CheckList - nalozi su spisak stavki
Public Const IC_MARZA    As Long = &HE9D2&   ' AreaChart
Public Const IC_IZVEST   As Long = &HE9F9&   ' ReportDocument
Public Const IC_SLEDLJ   As Long = &HE71B&   ' Link
' Akcije
Private Const IC_SAVE     As Long = &HE74E&   ' Save
Private Const IC_PRINT    As Long = &HE749&   ' Print
Private Const IC_CLOSE    As Long = &HE711&   ' Cancel
Private Const IC_FILTER   As Long = &HE71C&   ' Filter
Private Const IC_SEARCH   As Long = &HE71E&   ' Zoom
Private Const IC_MAX      As Long = &HE740&   ' FullScreen
Private Const IC_MIN      As Long = &HE73F&   ' BackToWindow
Private Const IC_SYNCBTN  As Long = &HE895&   ' Sync
Private Const IC_SNIMI    As Long = &HE78C&   ' SaveLocal (razlicit od IC_SAVE)
Private Const IC_EXCEL    As Long = &HE8A7&   ' OpenInNewWindow
Private Const IC_OPERATER As Long = &HE748&   ' SwitchUser
Public Const IC_STORNO   As Long = &HE7A7&   ' Undo - izbor korisnika
' PRIVREMENO: E7B8 (Package) se ne iscrtava u ovoj verziji fonta, kao ni druge
' kutije koje sam probao. Dok ne stigne kod iz DumpMdl2Sheet, stoji Library -
' isti glif kao "Palete" u sidebaru, ali bar nije prazan kvadrat.
Public Const IC_REVERS   As Long = &HE8F1&   ' Library - privremeno
Private Const IC_ENTER    As Long = &HE751&   ' ReturnKey - rezervisano
' Refresh: bio rezervisan, od v6-ui-121 nosi ekran "Oporavak" u sidebaru -
' taj ekran i jeste "vrati u ispravno stanje". Kod je iz verifikovanog
' opsega ispod EC00, kao i ostali.
Public Const IC_OPORAVAK As Long = &HE72C&   ' Refresh
' Stanja
Private Const IC_SYNC     As Long = &HE753&   ' Cloud
Private Const IC_HELP     As Long = &HE897&   ' Help
Private Const IC_MATICNI  As Long = &HE713&   ' Setting
' Ikonice maticne sekcije sidebara. Dve su PROVERENE (isti kod kao glif koji
' ekran vec crta), tri su izabrane iz MDL2 kataloga i NISU vizuelno proverene u
' ovoj sesiji -- provera je DumpMdl2Used (ime glifa stoji uz kod, pa se promasaj
' vidi i na papiru). Pogresan kod je kvadratic, ne pad: BuildNav crta ikonicu
' samo kad je kod razlicit od nule.
Public Const IC_MAT_PARTNERI  As Long = &HE716&   ' People       (za proveru)
Public Const IC_MAT_ROBA      As Long = &HE8EC&   ' Tag          (za proveru)
Public Const IC_MAT_PAKOVANJE As Long = &HE7B8&   ' Package      (za proveru)
Public Const IC_MAT_KORISNICI As Long = &HE748&   ' SwitchUser   (= IC_OPERATER)
Public Const IC_MAT_SISTEM    As Long = &HE713&   ' Setting      (= IC_MATICNI)
Public Const IC_MAT_ADMIN     As Long = &HE7EF&   ' Admin        (za proveru)
Private Const IC_NAZAD    As Long = &HE72B&   ' Back         (za proveru)
Private Const IC_KOPIRAJ  As Long = &HE8C8&   ' Copy - prepisi kolicinu
Private Const IC_PROFIL   As Long = &HE77B&   ' Contact - profil u sidebaru
' Novac ima SMER, pa isplata i uplata nose strelice, ne "novcanik" - iz
' ikonice se odmah vidi da li novac izlazi iz firme ili ulazi u nju.
Public Const IC_ISPLATA  As Long = &HE898&   ' Upload   - novac IZLAZI iz firme
Public Const IC_UPLATA   As Long = &HE896&   ' Download - novac ULAZI u firmu
Private Const IC_GORE     As Long = &HE70E&   ' ChevronUp - skrol panela
Private Const IC_DOLE     As Long = &HE70D&   ' ChevronDown - skrol panela

'--- Win32 -----------------------------------------------------------
Private Const WS_THICKFRAME As Long = &H40000

'--- stanje (SVE Object - modul mora ostati "mek") -------------------
Public Btns As Collection
Public ActiveMode As String

Private mFrm As Object
Private mHovered As clsFlatBtn
Private mView As Variant
Private mViewN As Long
Private mSumKg As Double
Private mSumVal As Double
Private mPage As Long
Private mPageSize As Long
Private mSortCol As Long
Private mSortLista As String     ' lista za koju vazi tekuci sort (v. SortZaListu)
Private mSortAsc As Boolean
Private mFilter As String
Private mSearch As String
Private mKlasa As Long
Private mToastPending As String
Private mDisplayFont As String
Private mBusyGrid As Boolean
Private mBuilding As Boolean
Private mColSpec As Object
Private mLastPageSize As Long
' Bazen cipova je ljuskin i ima MAX_CHIP slotova (ChipRow). Prvih MODE_CHIP
' pripada rezimima unosnog ekrana; ostale slobodno pozajmljuje ekran koji
' prijavi svoje cipove. mScrChipKey drzi KLJUC koji je ekran dao za slot --
' ljuska ga ne tumaci, samo ga vrati kroz Scr_Rows.
' JAVNA: ekran koji prijavi vise cipova nego sto bazen ima izgubio bi visak
' bez ijedne poruke. Test to meri bas kroz ovu konstantu.
Public Const MAX_CHIP   As Long = 7
Private Const MODE_CHIP As Long = 6
Private mScrChipN As Long
Private mScrChipKey(0 To 6) As String
Private mChipW(0 To 6) As Single
' Vec prijavljena prekoracenja bazena, po kljucu ekran/lista/vrsta.
Private mBazenPrijave As Object
Private mCntOtkaz As Long
Private mCntBezZb As Long
Private mCntNefakt As Long
' Otpremnice kojima ostatak (kg otpremnice - kg upisanih otkupnih listova)
' jos nije nula. Broj se racuna nad SVIM otpremnicama, ne nad filtriranom
' listom - cip mora da kaze koliko ih ukupno ceka, i kad je sam ugasen.
Private mCntOtvor As Long
' Stoji li na ekranu crveni toast koji je podigla mreza. Sledece uspelo
' citanje ga gasi; toast greske nema tajmer, pa bi inace ostao zauvek.
Private mGridToast As Boolean
' Polje na koje ide kursor pri novom dokumentu, i polje koje TRENUTNO ima
' zeleni okvir. Drugo se pamti da bi se okvir skinuo i kad MSForms preskoci
' dogadjaj napustanja (a preskace ga kad se forma sakrije ili polje ugasi).
Private mPrvoPolje As String
Private mFokusTag As String
' Oznacavanje vise redova (lista otpremnica -> specifikacija). Kljuc je PRVA
' KOLONA reda, ne njegov redni broj: sortiranje, pretraga i strane menjaju
' redni broj, a broj otpremnice ostaje isti.
Private mMarkOn As Boolean
Private mMark As Object
Private mChromeRemoved As Boolean
Private mCombosFilled As Boolean
Private mPartMap As Object
Private mSelRow As Long              ' izabran red mreze (1..mViewN), 0 = nijedan
Private mHoverRow As Long            ' red pod pokazivacem (0-bazirano), -1 = nijedan
Private mHoverHd As Long             ' zaglavlje kolone pod pokazivacem, -1 = nijedno
Private mBuildMs As Long             ' koliko je trajala gradnja ekrana (ms)
' nav tag ("nav01") -> kljuc ekrana iz registra. Sidebar se gradi iz
' modUiScreens, pa se veza tag->ekran vise ne moze citati iz rednog broja.
Private mNavKey As Object
' Brojaci uz stavke menija, po kljucu ekrana. Racunaju se na promenu podataka
' (RefreshFromData), ne pri crtanju -- izvor je prolaz kroz tabele.
Private mNavBroj As Object
' Aktivan ekran iz registra. "DOKUMENTI" je zatecen ekran koji ljuska jos
' crta po starom; svaki drugi se crta kroz ugovor, u svojoj zoni zScr_<kljuc>.
' Radna povrsina je USTUPLJENA panelu. Ljuska ne zna KOM panelu -- zna samo da
' zona ekrana, naslov i mreza ne smeju da se crtaju dok je ustupljena. Ko je u
' njoj i sta ume, zna modUiPanel; ljuska mu daje samo prazan okvir.
Private mPanelRezim As Boolean

Private mScreen As String
' SEKCIJA ljuske: RAD (dokumenti, finansije, analitika) ili MATICNI (sifarnici,
' sistem). Zlatno dugme u zaglavlju je menja; dva skupa stavki sidebara se nikad
' ne crtaju zajedno jer sidebar nema skrol (v. modUiScreens.SCR_SEKCIJA).
Private mSekcija As String
Private mNavUzak As Boolean          ' sidebar je suzen (uzak prozor)
Private mNavVisina As Single         ' visina koju je aktivna sekcija zauzela
Private mNavOff As Object            ' ime stavke sidebara -> stavka je prigusena
Private mZadnjiRad As String         ' poslednji ekran radne sekcije
Private mZadnjiMaticni As String     ' poslednji ekran maticne sekcije
Private mPopFor As String            ' combo za koji je otvoren nas dropdown ("" = zatvoren)
Private mColX(0 To MAX_COLS - 1) As Single   ' x i sirina kolona mreze; postavlja
Private mColW(0 To MAX_COLS - 1) As Single   ' LayoutGrid, RenderGrid samo puni

' Novcani slotovi podnozja, onako kako ih je vratio ekran: parovi
' Array(kljucPoruke, iznos). Prazno = ekran ih ne salje, pa podnozje crta zbir
' vrednosti kao i pre.
Private mFtSlot As Variant
Private mFtSlotN As Long
Private mCols As Variant             ' opis kolona AKTIVNOG rezima (GridCols)
Private mColN As Long                ' koliko ih je stvarno vidljivo
' Da li mColX / mColW jos odgovaraju mCols. Postavlja se kad se OPIS kolona
' promeni, brise kad ga LayoutGrid preracuna. V. OsveziGeometriju.
Private mGeomStara As Boolean
' Koliko celija u POSLEDNJEM crtanju nije moglo da se prikaze. Prijavljuje se
' jednom po crtanju, ne po celiji: pokvarena kolona bi inace dala red po redu
' istih poruka. v. CelijaTekst i GridKvarCelijaTest.
Private mKvarCelija As Long
' Prva kolona koja u poslednjem crtanju nije mogla da se prikaze -- "vrsta/indeks".
Private mKvarKind As String
Private mColsPotpis As String        ' po cemu se zna da je opis DRUGI
' Poslednja zbirna sa kojom se radilo. Preuzeto iz frmDokumenta: tamo je
' RefreshBrojZbirneSuggestion upisivala isti broj u sekciju Zbirne I u sekciju
' Otpremnice, a btnUnosZbr ga je posle snimanja gurao u Prijemnicu. Ovde je to
' jedan sticky broj koji prati prelazak izmedju rezima.
Private mAktivnaZbirna As String
' Traje dok FillZbirneCombo puni listu. ComboBox.Clear brise i UPISANU vrednost,
' pa polje javi promenu koju operater nije napravio - bez ovog garda bi svako
' osvezavanje posle snimanja obrisalo zapamcenu zbirnu i zaprljalo dokument.
Private mZbirnaFill As Boolean
' Isto za listu partnera: FillFormPartner radi CB.Clear, sto polje javi kao
' promenu partnera - a na tu promenu visi brisanje i predlog broja prijemnice.
Private mPartnerFill As Boolean
Private mGridMax As Boolean          ' mreza razvucena do ispod naslova dokumenta
' Izabran smer reversa (1..4); NULA znaci "operater jos nije izabrao" i tako se
' i predaje ekranu. Nijedan segment nije unapred obelezen namerno: legacy
' izricito trazi eksplicitan izbor, jer je prazan smer ranije tiho knjizio
' "OM prima od vozaca" (modNovacUnos.ReversValidiraj).
Private mSmerRev As Long
Private mIzAvansa As Boolean         ' F5: isplata ide iz OM avansa, ne virmanom
Private mPartnerFor As String        ' za koji rezim je partner lista vec napunjena
' Da li je lista USPESNO ucitana. Prazna lista sme da znaci "nema otvorenih"
' -- i time avans -- samo posle uspesnog citanja; pad ucitavanja ostavlja isti
' prazan combo, a odluka o novcu je suprotna. v. NovacListaSme.
Private mBlokListaOk As Boolean      ' lista otvorenih blokova (F5) ucitana
Private mBlokListaErr As String
Private mFakListaOk As Boolean       ' lista otvorenih faktura (F6) ucitana
Private mFakListaErr As String
Private mBlokIDs As Variant          ' OtkupID po indeksu combo-a otvorenih blokova
Private mBlokRest As Variant         ' neisplaceni ostatak po istom indeksu
Private mFakIDs As Variant           ' FakturaID po indeksu combo-a otvorenih faktura
Private mFakRest As Variant          ' preostali iznos fakture po istom indeksu
Private mPopSrc() As Long            ' indeksi stavki combo-a koje panel prikazuje
Private mPopN As Long                ' koliko ih je posle suzavanja kucanjem
Private mPopTop As Long              ' prva prikazana (skrol po stranama)
Private mPopMute As Boolean          ' programski upis u combo - ne otvaraj panel
Private mPopHover As Long            ' hover red panela (v6-ui-188: centralno bojenje)
Private mFltOpen As Boolean          ' panel "Filteri" otvoren
Private mFltVrsta As String          ' dodatni uslov: vrsta iz Osnovnih podataka
Private mFltPart As String           ' dodatni uslov: partner iz Osnovnih podataka
Private mDirty As Boolean            ' ima neupisanih izmena u formi
Private mLoading As Boolean          ' program puni polja - to NIJE izmena korisnika
'=====================================================================
' BUILD
'=====================================================================
Public Sub BuildOtkupScreen(frm As Object)
    Dim t0 As Double: t0 = Timer
    Dim errNum As Long, errDesc As String
    ' Gradnja pipa stotine runtime kontrola; ako bilo koja padne, mBuilding mora
    ' da se vrati na False (svi handleri ga proveravaju) pre nego sto greska ode
    ' dalje - inace ostaje "ziv" ekran koji ne reaguje ni na sta.
    On Error GoTo EH
    mBuilding = True
    Set mFrm = frm
    Set Btns = New Collection
    Set mHovered = Nothing
    modUiData.ResetCache
    Set mColSpec = CreateObject("Scripting.Dictionary")
    mColSpec.CompareMode = 1

    ActiveMode = "F1"
    mFilter = "danas"
    mSearch = ""
    mKlasa = 1
    mSortCol = 2
    mSortAsc = False
    mPage = 1
    mLastPageSize = 0
    mSelRow = 0
    mHoverRow = -1
    mHoverHd = -1
    mScreen = SCR_POCETNI
    mSekcija = SEK_RAD
    mNavUzak = False
    mZadnjiRad = ""
    mZadnjiMaticni = ""
    modUiKit.ResetNumFields             ' popunjava ga NewTxt tokom gradnje
    mDisplayFont = DisplayFont()

    With frm
        .caption = Poruka("OTKUI_APP_CAPTION")
        .BackColor = C_CREAM
        .width = 1100: .Height = 700
    End With

    BuildHeader frm
    BuildNav frm
    BuildKpi frm
    BuildTitle frm
    BuildOtpTraka frm
    BuildCtx frm
    BuildForm frm
    BuildGrid frm
    BuildPanelHost frm
    BuildRight frm
    BuildStatus frm

    BuildPopup frm
    BuildFilterPanel frm
    SetKlasa 1                      ' II klasa starta onemogucena
    SelectModeCore frm, "F1", False
    ShowZones frm
    mBuilding = False
    mBuildMs = CLng((Timer - t0) * 1000)
    Exit Sub
EH:
    errNum = Err.Number: errDesc = Err.description
    mBuilding = False
    Err.Raise errNum, "modOtkupUI.BuildOtkupScreen", errDesc
End Sub

'--------------------------------------------------------- HEADER ----
Private Sub BuildHeader(frm As Object)
    Dim z As Object, i As Long
    Set z = NewZone(frm, "zHdr", 0, 0, 1100, HEADER_H, C_FOREST)
    WireZone z

    ' gradijent 135deg se fejkuje trakama interpolirane boje (MSForms nema fill)
    For i = 0 To 23
        NewLbl z, "hgr" & i, "", i * 46, 0, 47, HEADER_H, 8, False, 0, _
               Lerp(C_FOREST, C_FOREST_DK, i / 23)
        ' podloga pokriva celu zonu, pa Frame ispod nikad ne bi video klik;
        ' bez ovoga klik u zaglavlje ne zatvara otvoreni panel
        WireBtn z.Controls("hgr" & i), "zHdr", "chev"
    Next i

    NewLbl z, "hdrAX", "AX", PAD, 10, 24, 20, TS_H1 + 3, True, C_GOLD, -1, fmTextAlignLeft, mDisplayFont
    NewLbl z, "hdrName", "OtkupApp", PAD + 24, 11, 76, 18, TS_H1 + 1, True, C_CREAM, -1, fmTextAlignLeft, mDisplayFont
    NewLbl z, "hdrDot", "", PAD + 108, 18, 6, 6, 8, False, 0, C_GREEN_LT
    NewLbl z, "hdrStat", HeaderStatusText(), PAD + 120, 13, 250, 14, TS_META, False, RGB(178, 190, 172), -1

    ' Znacka stanja: zeleno kvacica dok je sve upisano, zlatna tacka cim
    ' korisnik nesto promeni. Operater mora da vidi da mu unos jos nije snimljen.
    NewLbl z, "hdrSaveTick", ChrW(10003), 600, 13, 12, 14, TS_META, True, C_GREEN_LT, -1, fmTextAlignRight
    NewLbl z, "hdrSaved", Poruka("OTKUI_HDR_SACUVANO"), 616, 13, 160, 14, TS_META, False, RGB(188, 198, 182), -1
    ' Alatke aplikacije (ne dokumenta): operater, snimanje radne sveske, izlaz
    ' na radni list. Sve tri su "ghostdark" - na tamnom zaglavlju su tiho
    ' prisutne, a "Maticni podaci" ostaje jedino zlatno (primarno) dugme.
    BtnV z, "btnExcel", Poruka("OTKUI_BTN_EXCEL"), 400, 10, 112, 22, "ghostdark", IC_EXCEL
    ' "Snimi" je znacilo ThisWorkbook.Save - za operatera to nije poslovna
    ' radnja i mesalo se sa "Sacuvaj dokument". Natpis sada kaze sta radi.
    BtnV z, "btnSnimi", Poruka("OTKUI_BTN_SNIMI"), 520, 10, 132, 22, "ghostdark", IC_SNIMI
    ' operater vise nije u zaglavlju - presao je u dno sidebara (profil)
    BtnV z, "btnMatic", Poruka("OTKUI_BTN_MATICNI"), 780, 10, 118, 22, "gold", IC_MATICNI
    BtnV z, "btnClose", "", 900, 10, 26, 22, "ghostx", IC_CLOSE
End Sub

'------------------------------------------------------------ NAV ----
' Sidebar se GRADI za obe sekcije odjednom, a CRTA se samo aktivna.
'
' Zasto obe odjednom: stavka menija je ozicena kroz WireBtn -> clsFlatBtn
' (WithEvents). Pravljenje i rusenje kontrola pri svakoj promeni sekcije bi
' znacilo i ponovno ozicavanje, a ziv event-sink je u ovom projektu vec jednom
' skupo naplacen (modSelfUpdate zove MaticniMenu_Release pre importa). Ovako se
' kontrole naprave jednom i samo pale, gase i pomeraju -- LayoutNav.
'
' Grupe se broje kroz SVE sekcije (gi 0..4), pa imena stavki ostaju "nav" + dve
' cifre: PaintNav i NavCollapse racunaju na Len(name) = 5.
Private Sub BuildNav(frm As Object)
    Dim z As Object, i As Long, j As Long
    Dim grp As Variant, r As Variant, gk As String, kljuc As String
    Dim aktivan As Boolean, ico As Long, nm As String, cap As String
    Set z = NewZone(frm, "zNav", 0, HEADER_H, SIDEBAR_W, 400, C_SAND)
    WireZone z
    NewLbl z, "navEdge", "", SIDEBAR_W - 1, 0, 1, 400, 8, False, 0, C_BORDER

    ' Sidebar vise nije nabrojan ovde nego dolazi iz modUiScreens. Ekran ciji
    ' modul jos ne postoji (ili na koji korisnik nema pravo) crta se prigusen
    ' i klik na njega kaze zasto - umesto dosadasnjeg "nije vezano" za sve.
    Set mNavKey = CreateObject("Scripting.Dictionary")
    Set mNavOff = CreateObject("Scripting.Dictionary")
    modUiScreens.ScrResetCache
    grp = modUiScreens.ScrGroups()
    For i = 0 To UBound(grp)
        gk = modUiScreens.ScrField(CStr(grp(i)), SCRG_KLJUC)
        NewLbl z, "navGrp" & i, _
               UCase$(Poruka(modUiScreens.ScrField(CStr(grp(i)), SCRG_NASLOV))), 13, 0, 130, _
               TxtH(TS_NAVGRP), TS_NAVGRP, True, C_MUTED, -1
        j = 0
        For Each r In modUiScreens.ScrRows()
            If modUiScreens.ScrField(CStr(r), SCR_GRUPA) = gk Then
                kljuc = modUiScreens.ScrField(CStr(r), SCR_KLJUC)
                cap = Poruka(modUiScreens.ScrField(CStr(r), SCR_NASLOV))
                ico = CLng(val(modUiScreens.ScrField(CStr(r), SCR_IKONICA)))
                aktivan = modUiScreens.ScrAktivan(kljuc)
                nm = "nav" & i & j
                mNavKey(nm) = kljuc
                mNavOff(nm) = Not aktivan

                Dim sel As Boolean: sel = (kljuc = SCR_POCETNI)
                NewLbl z, "navBar" & i & j, "", 0, 0, 3, NAV_H, 8, False, 0, _
                       IIf(sel, C_GOLD, C_SAND)
                If ico <> 0 Then
                    NewLbl z, "navIco" & i & j, ChrW(ico), 12, 0, 20, TxtH(TS_NAVICO), _
                           TS_NAVICO, False, _
                           IIf(sel, C_GOLD, IIf(aktivan, C_ICON_OFF, C_DISABLED_FG)), -1, _
                           fmTextAlignLeft, F_ICON
                End If
                NewNavItem z, nm, cap, 0, 0, SIDEBAR_W - 1, NAV_H, sel
                ' Znacka sa brojem stavki koje cekaju. Prazna dok brojac ne stigne;
                ' ekran koji nema sta da broji je nikad ne popuni.
                NewLbl z, nm & "B", "", SIDEBAR_W - 42, 0, 30, TxtH(TS_MICRO), _
                       TS_MICRO, True, C_RUST, -1, fmTextAlignRight, F_NUM
                z.Controls(nm & "B").ZOrder 0
                If Not aktivan And Not sel Then _
                    z.Controls(nm & "X").ForeColor = C_DISABLED_FG
                ' stavka menija je neprozirna i preko pune sirine - bez ZOrder-a
                ' bi prekrila i zlatnu traku selekcije i MDL2 ikonicu
                z.Controls("navBar" & i & j).ZOrder 0
                On Error Resume Next
                z.Controls("navIco" & i & j).ZOrder 0     ' moze da ne postoji (kod 0)
                On Error GoTo 0
                z.Controls(nm & "X").ZOrder 0
                j = j + 1
            End If
        Next r
    Next i

    ' PROFIL - dno sidebara. Nije natpis nego prekidac operatera: klik otvara
    ' prijavu, isto sto je radilo dugme u zaglavlju (tag ostaje "btnOperater",
    ' pa se rukovalac klika ne dira). Umesto inicijala stoji MDL2 glif osobe.
    NewLbl z, "profLn", "", 0, 0, SIDEBAR_W - 1, 1, 8, False, 0, C_BORDER
    NewLbl z, "profBox", "", 9, 0, 26, 26, 8, False, 0, C_FOREST
    NewLbl z, "profIco", ChrW(IC_PROFIL), 9, 0, 26, TxtH(TS_NAVICO), _
           TS_NAVICO, False, C_GOLD, -1, fmTextAlignCenter, F_ICON
    NewLbl z, "profName", OperaterText(), 41, 0, SIDEBAR_W - 50, TxtH(TS_NAV), _
           TS_NAV, True, C_FOREST, -1
    NewLbl z, "profSub", "", 41, 0, SIDEBAR_W - 50, TxtH(TS_MICRO), _
           TS_MICRO, False, RGB(146, 158, 140), -1
    ' providna meta preko celog bloka - Label hvata klik i kad se pogodi razmak
    NewLbl z, "profHit", "", 0, 0, SIDEBAR_W - 1, 30, 8, False, 0, -1
    z.Controls("profHit").ZOrder 0
    WireBtn z.Controls("profHit"), "btnOperater", "chev"

    NewLbl z, "navFootLn", "", 0, 0, SIDEBAR_W - 1, 1, 8, False, 0, C_BORDER
    NewLbl z, "navSezona", Poruka("OTKUI_FOOT_SEZONA"), 13, 0, 70, 13, TS_MICRO, False, RGB(146, 158, 140), -1
    NewLbl z, "navVerzija", BuildTagOrBlank(), SIDEBAR_W - 68, 0, 55, 13, TS_MICRO, False, RGB(146, 158, 140), -1, fmTextAlignRight, F_NUM

    LayoutNav frm
End Sub

' Postavlja Y i vidljivost stavkama AKTIVNE sekcije, a stavke druge gasi.
' Zove se pri gradnji, pri promeni sekcije i iz LayoutOtkup (posle NavCollapse).
'
' Vidljivost je NAMERNO na jednom mestu. Da je razdeljena izmedju gradnje i
' NavCollapse, suzavanje prozora bi upalilo i stavke sekcije koja se ne crta --
' ista klasa kao traka zOtp koja je ostajala upaljena na tudjem ekranu.
Private Sub LayoutNav(frm As Object)
    Dim z As Object, grp As Variant, r As Variant
    Dim i As Long, j As Long, Y As Single, gk As String, nm As String
    Dim uSekciji As Boolean, brStavki As Long
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    Set z = frm.Controls("zNav")
    If z Is Nothing Then Exit Sub

    grp = modUiScreens.ScrGroups()
    Y = 12
    For i = 0 To UBound(grp)
        gk = modUiScreens.ScrField(CStr(grp(i)), SCRG_KLJUC)
        uSekciji = (modUiScreens.ScrGrupaSekcija(CStr(grp(i))) = mSekcija)
        brStavki = BrojStavkiGrupe(gk)

        z.Controls("navGrp" & i).Visible = uSekciji And (brStavki > 0) And Not mNavUzak
        If uSekciji And brStavki > 0 Then
            z.Controls("navGrp" & i).top = Y
            Y = Y + 20
        End If

        j = 0
        For Each r In modUiScreens.ScrRows()
            If modUiScreens.ScrField(CStr(r), SCR_GRUPA) = gk Then
                nm = "nav" & i & j
                If uSekciji Then
                    PostaviNavStavku z, nm, i, j, Y, True
                    Y = Y + NAV_H
                Else
                    PostaviNavStavku z, nm, i, j, 0, False
                End If
                j = j + 1
            End If
        Next r
        If uSekciji And brStavki > 0 Then Y = Y + 12
    Next i
    mNavVisina = Y
End Sub

Private Function BrojStavkiGrupe(ByVal gk As String) As Long
    Dim r As Variant
    For Each r In modUiScreens.ScrRows()
        If modUiScreens.ScrField(CStr(r), SCR_GRUPA) = gk Then _
            BrojStavkiGrupe = BrojStavkiGrupe + 1
    Next r
End Function

' Jedna stavka sidebara: pozadina, natpis, znacka, zlatna traka i ikonica.
' Sve pet kontrola nose ISTU odluku o vidljivosti - stavka koja se ne crta ne
' sme da ostavi za sobom traku ili glif druge sekcije.
Private Sub PostaviNavStavku(z As Object, ByVal nm As String, ByVal i As Long, _
                             ByVal j As Long, ByVal Y As Single, ByVal vis As Boolean)
    On Error Resume Next
    z.Controls(nm).top = Y
    z.Controls(nm).Visible = vis
    z.Controls(nm).width = IIf(mNavUzak, SIDEBAR_MIN - 1, SIDEBAR_W - 1)
    z.Controls(nm & "X").top = CenterY(Y, NAV_H, TS_NAV)
    z.Controls(nm & "X").Visible = vis And Not mNavUzak
    z.Controls(nm & "B").top = CenterY(Y, NAV_H, TS_MICRO)
    z.Controls(nm & "B").Visible = vis
    z.Controls("navBar" & i & j).top = Y
    z.Controls("navBar" & i & j).Visible = vis
    z.Controls("navIco" & i & j).top = CenterIco(Y, NAV_H, TS_NAVICO)
    z.Controls("navIco" & i & j).Left = IIf(mNavUzak, 11, 12)
    z.Controls("navIco" & i & j).Visible = vis
End Sub


Private Sub BuildKpi(frm As Object)
    Dim z As Object, i As Long, cap As Variant
    Set z = NewZone(frm, "zKpi", SIDEBAR_W, HEADER_H, 800, KPI_H, C_CREAM)
    WireZone z
    NewLbl z, "kpiLnB", "", 0, KPI_H - 1, 800, 1, 8, False, 0, C_BORDER

    cap = Array(Poruka("OTKUI_KPI_DANAS"), Poruka("OTKUI_KPI_OM_SALDO"), Poruka("OTKUI_KPI_OTVORENO"), _
                Poruka("OTKUI_KPI_DOKUMENATA"), Poruka("OTKUI_KPI_VALIDACIJA"))
    ' Akcentna linija po plocici - boja govori o CEMU je podatak, pre nego sto
    ' se procita naslov: kolicina, dugovanje, otvorene stavke, validacija.
    Dim acc As Variant
    acc = Array(C_GREEN_LT, C_RUST, C_GOLD, C_GREEN_LT, C_GREEN)
    For i = 0 To 4
        NewLbl z, "kpiBg" & i, "", 0, 0, 120, KPI_H - 1, 8, False, 0, IIf(i = 4, C_KPI_OK, C_CREAM)
        WireBtn z.Controls("kpiBg" & i), "zKpi", "chev"
        NewLbl z, "kpiAcc" & i, "", 0, 12, 2, KPI_H - 24, 8, False, 0, CLng(acc(i))
        NewSectionHdr z, "kpiC" & i, CStr(cap(i)), 0, 6, 120
        NewLbl z, "kpiV" & i, ChrW(8212), 0, 18, 120, TxtH(TS_KPI), IIf(i = 4, TS_H1, TS_KPI), True, C_FOREST, -1, _
               fmTextAlignLeft, IIf(i = 4, F_UI, F_NUM)
        ' treci red: kontekst uz vrednost; prazan kad nema sta da kaze
        NewLbl z, "kpiS" & i, "", 0, 38, 120, TxtH(TS_MICRO), TS_MICRO, False, RGB(146, 158, 140), -1
        If i < 4 Then NewLbl z, "kpiSep" & i, "", 0, 10, 1, KPI_H - 21, 8, False, 0, C_BORDER_LT
    Next i
End Sub

'---------------------------------------------------------- TITLE ----
Private Sub BuildTitle(frm As Object)
    Dim z As Object
    Set z = NewZone(frm, "zTitle", SIDEBAR_W, HEADER_H + KPI_H, 800, TITLE_H, C_WHITE)
    WireZone z
    NewLbl z, "titLnB", "", 0, TITLE_H - 1, 800, 1, 8, False, 0, C_BORDER

    ' Marker uz naslov: tamni kvadrat sa MDL2 ikonicom DOKUMENTA (ModeIco).
    ' MSForms nema radijus, pa ostaje pravougaonik - tako i stoji u delti.
    NewLbl z, "titIcoB", "", PAD, CenterY(0, TITLE_H, TS_DISPLAY) - 4, _
           TIT_ICO_W, TIT_ICO_W, 8, False, 0, C_FOREST
    NewLbl z, "titIco", ChrW(IC_OTKUP), PAD, 0, TIT_ICO_W, TxtH(TS_NAVICO), _
           TS_NAVICO, False, C_GOLD, -1, fmTextAlignCenter, F_ICON
    z.Controls("titIco").top = CenterIco(z.Controls("titIcoB").top, TIT_ICO_W, TS_NAVICO)
    z.Controls("titIco").ZOrder 0

    NewLbl z, "titName", "-", PAD + TIT_ICO_W + 10, 1, 230, TxtH(TS_DISPLAY), TS_DISPLAY, True, C_FOREST, -1, fmTextAlignLeft, mDisplayFont
    NewLbl z, "titSub", "-", PAD + TIT_ICO_W + 10, 25, 380, TxtH(TS_META), TS_META, False, C_MUTED, -1
    ' znacka rezima (F1..F8) - stoji uz naslov, ne u liniji sa podnaslovom
    NewLbl z, "titBadgeB", "", 0, 4, 26, 15, 8, False, 0, C_SOFT_BG
    NewLbl z, "titBadgeC", "", 0, 5, 26, TxtH(TS_MICRO), TS_MICRO, True, C_GREEN, -1, fmTextAlignCenter, F_NUM
    z.Controls("titBadgeC").ZOrder 0
    ' lenjir za merenje naslova - znacka mora da stane tacno iza teksta
    NewLbl z, "titRul", "", 0, -400, 10, 14, TS_DISPLAY, True, C_WHITE, -1, fmTextAlignLeft, mDisplayFont
    z.Controls("titRul").AutoSize = True
    NewLbl z, "titDatum", "-", 0, CenterY(0, TITLE_H, TS_META), 190, TxtH(TS_META), _
           TS_META, False, C_MUTED, -1, fmTextAlignRight
End Sub

'------------------------------------------------------------ CTX ----
' TRAKA OTPREMNICE (F1). Otpremnica je izvor robe; ovde stoji koja je aktivna
' i koliko je od nje jos neraspodeljeno. Vidi se samo u F1 - u ostalim
' rezimima zona je sakrivena i ne zauzima visinu.
Private Sub BuildOtpTraka(frm As Object)
    Dim z As Object, i As Long, X As Single
    Set z = NewZone(frm, "zOtp", SIDEBAR_W, 0, 800, OTP_H, C_SOFT_BG)
    WireZone z
    NewLbl z, "otpLnB", "", 0, OTP_H - 1, 800, 1, 8, False, 0, C_BORDER

    NewLbl z, "otpCap", UCase$(Poruka("OTKUI_OTP_TRAKA")), PAD, 5, 120, 11, _
           TS_MICRO, True, C_MUTED, -1
    NewLbl z, "otpBroj", ChrW(8212), PAD, 17, 260, 18, TS_H1, True, C_FOREST, -1
    NewLbl z, "otpSub", "", PAD, 34, 260, 13, TS_META, False, C_MUTED, -1

    ' tri meraca + cena; svaki ima natpis, veliku brojku i red za ambalazu
    For i = 0 To 3
        NewLbl z, "otpML" & i, "", 0, 5, 120, 11, TS_MICRO, True, C_MUTED, -1
        NewLbl z, "otpMV" & i, ChrW(8212), 0, 16, 120, 20, TS_KPI, True, C_FOREST, _
               -1, fmTextAlignLeft, F_NUM
        NewLbl z, "otpMA" & i, "", 0, 35, 120, 12, TS_MICRO, False, C_MUTED, _
               -1, fmTextAlignLeft, F_NUM
    Next i
    NewLbl z, "otpPrazno", Poruka("OTKUI_OTP_NEMA"), PAD, 20, 420, 16, _
           TS_META, False, C_MUTED, -1
End Sub

' Puni traku iz opisa koji daje ekran (Scr_OtpInfo). Ljuska ne zna sta je
' otpremnica - dobija gotove brojeve i samo ih crta.
Private Sub RefreshOtpTraka(frm As Object)
    Dim z As Object, info As String, p() As String, i As Long
    Dim ost As Double, ostA As Double, ima As Boolean
    On Error Resume Next
    Set z = frm.Controls("zOtp")
    If z Is Nothing Then Exit Sub
    info = ""
    If mScreen = "DOKUMENTI" Then _
        info = CStr(Application.Run("modScrDokumenti.Scr_OtpInfo"))
    Err.Clear
    ima = (Len(info) > 0)

    z.Controls("otpPrazno").Visible = Not ima
    z.Controls("otpBroj").Visible = ima
    z.Controls("otpSub").Visible = ima
    For i = 0 To 3
        z.Controls("otpML" & i).Visible = ima
        z.Controls("otpMV" & i).Visible = ima
        z.Controls("otpMA" & i).Visible = ima
    Next i
    If Not ima Then Exit Sub

    p = Split(info, "|")
    z.Controls("otpBroj").caption = p(0)
    z.Controls("otpSub").caption = p(1) & "  " & ChrW(183) & "  " & p(2)

    z.Controls("otpML0").caption = UCase$(Poruka("OTKUI_OTP_UKUPNO"))
    z.Controls("otpMV0").caption = FmtBroj(CDbl(Val(p(3))), 2)
    z.Controls("otpMA0").caption = Poruka("OTKUI_OTP_AMB") & " " & FmtBroj(CDbl(Val(p(6))), 0)
    z.Controls("otpML1").caption = UCase$(Poruka("OTKUI_OTP_UBLOK"))
    z.Controls("otpMV1").caption = FmtBroj(CDbl(Val(p(4))), 2)
    z.Controls("otpMA1").caption = Poruka("OTKUI_OTP_AMB") & " " & FmtBroj(CDbl(Val(p(7))), 0)

    ost = CDbl(Val(p(5))): ostA = CDbl(Val(p(8)))
    z.Controls("otpML2").caption = UCase$(Poruka("OTKUI_OTP_OSTATAK"))
    z.Controls("otpMV2").caption = FmtBroj(ost, 2)
    ' Ostatak je jedini broj zbog koga ova traka postoji - zato ima semafor.
    ' Negativan znaci da je u blokove upisano vise nego sto je otpremnica
    ' donela; isto pravilo kao u staroj formi (RefreshSummary).
    z.Controls("otpMV2").ForeColor = IIf(ost < -0.0001, C_RUST, C_GREEN)
    z.Controls("otpMA2").caption = Poruka("OTKUI_OTP_AMB") & " " & FmtBroj(ostA, 0)
    z.Controls("otpMA2").ForeColor = IIf(ostA < -0.0001, C_RUST, C_MUTED)

    z.Controls("otpML3").caption = UCase$(Poruka("OTKUI_OTP_CENA"))
    z.Controls("otpMV3").caption = FmtBroj(CDbl(Val(p(9))), 2)
    z.Controls("otpMA3").caption = Poruka("OTKUI_OTP_PO_OTP")
End Sub

' Koja je lista izabrana. Stanje drzi EKRAN; ljuska ga pita kasno vezano, pa
' ekran koji nema liste (ili modul koji nedostaje) daje prazno umesto greske.
Private Function ActiveLista() As String
    On Error Resume Next
    ActiveLista = modUiScreens.ScrLista(mScreen)
    Err.Clear
End Function

' Naslov mreze i cipovi prate IZABRANU listu, ne samo rezim. Naslov svake liste
' stoji u njenoj definiciji (treca kolona Scr_Liste) - ljuska ne poznaje nijedan
' kljuc liste. Cipovi ("Danas", "Bez zbirne"...) pripadaju listi dokumenata; u
' ostalim listama nemaju sta da suze pa bi stajali kao mrtva dugmad sa nulama.
Private Sub RefreshGridTitle(frm As Object)
    Dim z As Object, akt As String, k As String, ch As Variant
    Dim seg As Variant, i As Long, sp As Variant, dop As String
    On Error Resume Next
    Set z = frm.Controls("zGrid")
    akt = ActiveLista()
    seg = SegDefs()
    If Len(akt) > 0 And IsArray(seg) Then
        For i = 0 To UBound(seg)
            sp = Split(CStr(seg(i)), "|")
            If CStr(sp(0)) = akt Then
                ' naslov sme da nosi dopunu (broj aktivne otpremnice, broj
                ' palete) - ekran je daje kroz Scr_NaslovDopuna
                dop = ScrNaslovDopuna()
                z.Controls("grdTitle").caption = Poruka(CStr(sp(2))) & _
                                                 IIf(Len(dop) > 0, " " & dop, "")
                Exit For
            End If
        Next i
    ElseIf mScreen = "DOKUMENTI" Then
        z.Controls("grdTitle").caption = Poruka("OTKUI_GRID_TITLE_" & modeKey(ActiveMode))
    End If

    RefreshChipsForScreen frm
End Sub

' Cipovi AKTIVNE liste. Ekran ih prijavljuje kroz Scr_Cipovi, istim oblikom
' kao radnje: kljuc:KATALOG:sirina.
'
' Ovde je do sada stajalo 'ako je lista OTPREMNICE, pokazi chipSve i
' chipOtvorene' -- poslednje mesto na kom je ljuska znala jedan ekran po
' imenu. Sada zna samo svoj POCETNI, koji jedini ima cipove vezane za rezim
' dokumenta (njihovu vidljivost postavlja SelectMode, prema tome da li rezim
' uopste ima zbirnu i fakturu).
Private Sub RefreshChipsForScreen(frm As Object)
    Dim spec As String, e As Variant, p As Variant, i As Long, n As Long
    Dim ch As Variant, nm As String, imaAktivan As Boolean
    On Error Resume Next
    ch = ChipRow()
    mScrChipN = 0
    For i = 0 To MAX_CHIP - 1
        mScrChipKey(i) = ""
    Next i
    spec = ScrCipovi()

    If Len(spec) = 0 Then
        If mScreen = SCR_POCETNI Then
            ' vidljivost cipova rezima postavlja SelectMode -- ovde se sklanjaju
            ' samo slobodni slotovi bazena i vracaju natpisi koje je ekran menjao
            For i = MODE_CHIP To MAX_CHIP - 1
                ShowChip frm, Split(CStr(ch(i)), "|")(0), False
            Next i
            VratiNatpiseCipova frm
        Else
            For i = 0 To MAX_CHIP - 1
                ShowChip frm, Split(CStr(ch(i)), "|")(0), False
            Next i
        End If
        LayoutChips frm
        Exit Sub
    End If

    Dim staje As Long
    staje = BazenStaje(UBound(Split(spec, "|")) + 1, MAX_CHIP, "cipovi")
    For Each e In Split(spec, "|")
        p = Split(CStr(e), ":")
        If UBound(p) >= 2 And n < staje Then
            nm = Split(CStr(ch(n)), "|")(0)
            mScrChipKey(n) = CStr(p(0))
            mChipW(n) = CSng(val(p(2)))
            ' Natpis '~xxx' je SIROVA vrednost (dinamicki cipovi -- vrsta/
            ' sorta iz podataka nemaju kataloski kljuc; krug 18).
            If Left$(CStr(p(1)), 1) = "~" Then
                frm.Controls("zGrid").Controls(nm & "C").caption = Mid$(CStr(p(1)), 2)
            Else
                frm.Controls("zGrid").Controls(nm & "C").caption = Poruka(CStr(p(1)))
            End If
            ShowChip frm, nm, True
            If mScrChipKey(n) = mFilter Then imaAktivan = True
            n = n + 1
        End If
    Next e
    mScrChipN = n
    For i = n To MAX_CHIP - 1
        ShowChip frm, Split(CStr(ch(i)), "|")(0), False
    Next i
    ' Filter koji izabrana lista ne poznaje ostavio bi mrezu suzenu necim sto
    ' se ne vidi ni na jednom upaljenom cipu. Prvi cip je po dogovoru najsiri.
    If Not imaAktivan Then mFilter = mScrChipKey(0)
    LayoutChips frm
    ApplyChipVisual frm.Controls("zGrid")
End Sub

' Natpisi i sirine bazena se vracaju na ljuskine kad ekran vise ne pozajmljuje
' slotove -- inace bi na unosnom ekranu ostao natpis prethodnog ekrana.
Private Sub VratiNatpiseCipova(frm As Object)
    Dim ch As Variant, p As Variant, i As Long
    On Error Resume Next
    ch = ChipRow()
    For i = 0 To MAX_CHIP - 1
        p = Split(CStr(ch(i)), "|")
        frm.Controls("zGrid").Controls(CStr(p(0)) & "C").caption = Poruka(CStr(p(1)))
        mChipW(i) = CSng(val(p(2)))
    Next i
    RenderChipCounts frm.Controls("zGrid")
End Sub

' Opis cipova AKTIVNOG ekrana. Ekran koji ih nema vraca prazno.
Private Function ScrCipovi() As String
    On Error Resume Next
    ScrCipovi = modUiScreens.ScrCipovi(mScreen)
    Err.Clear
End Function

' Dopuna naslova mreze (npr. broj aktivne otpremnice). Opciona - ekran koji je
' nema vraca prazno.
Private Function ScrNaslovDopuna() As String
    Dim m As String
    On Error Resume Next
    m = modUiScreens.ScrField(modUiScreens.ScrRowByKey(mScreen), SCR_MODUL)
    If Len(m) > 0 Then ScrNaslovDopuna = CStr(Application.Run(m & ".Scr_NaslovDopuna"))
    Err.Clear
End Function

' Koji segment prekidaca je izabran - stanje drzi ekran.
Private Sub RefreshListSeg(frm As Object)
    Dim z As Object, akt As String, seg As Variant, i As Long, k As String
    On Error Resume Next
    Set z = frm.Controls("zGrid")
    akt = ActiveLista()
    seg = SegDefs()
    If Not IsArray(seg) Then Exit Sub
    For i = 0 To BazenStaje(UBound(seg) + 1, MAX_SEG, "liste") - 1
        k = Split(CStr(seg(i)), "|")(0)
        BoxState z, "lsSeg" & i, _
                 IIf(akt = k, C_FOREST, C_WHITE), _
                 IIf(akt = k, C_CREAM, C_MUTED), (akt = k)
        RebaseSink "lsSeg" & i
    Next i
End Sub

Private Sub BuildCtx(frm As Object)
    Dim z As Object, i As Long, nm As Variant, w As Variant, X As Single
    Set z = NewZone(frm, "zCtx", SIDEBAR_W, HEADER_H + KPI_H + TITLE_H, 620, CTX_H, C_SAND_LT)
    WireZone z
    NewLbl z, "ctxLnB", "", 0, CTX_H - 1, 620, 1, 8, False, 0, C_BORDER

    ' Naslov sekcije dobija isti tretman kao grupe polja: eyebrow levo, tanka
    ' linija preko sredine, hint desno - pa tek onda red polja. Ranije je stajao
    ' u istoj liniji sa poljima, pa je izgledao kao natpis prvog polja.
    NewSectionHdr z, "ctxHdr", Poruka("OTKUI_LBL_OSNOVNI"), PAD, 2, 96
    NewLbl z, "ctxLn", "", 0, 7, 100, 1, 8, False, 0, C_BORDER_LT
    ' Poslednje polje (cbKupac / ctxL4) je PARTNER dokumenta - onaj koga
    ' zaduzujemo ili razduzujemo. Natpis je uvek "Partner" jer tip nije vezan
    ' za dokument (revers ide i kooperantu i kupcu); lista nosi SVE partnere,
    ' redosled zavisi od rezima (FillFormPartner).
    nm = Array("cbVrsta", "cbSorta", "cbOM", "cbVozac", "cbKupac")
    ' Natpis stoji IZNAD polja, isto kao u ostatku forme - ovo je bio jedini
    ' red na ekranu sa natpisom sa strane. Sirinu deli LayoutCtx po tezinama.
    w = Array(78, 78, 96, 82, 194)
    Dim lb As Variant
    lb = Array("OTKUI_CTXL_VRSTA", "OTKUI_CTXL_SORTA", "OTKUI_CTXL_OM", _
               "OTKUI_CTXL_VOZAC", "OTKUI_CTXL_PARTNER")
    X = PAD
    For i = 0 To 4
        ' isti rez kao natpis polja u formi: 8pt Semibold, verzal, prigusen
        NewLbl z, "ctxL" & i, UCase$(Poruka(CStr(lb(i)))), X, CTX_LBL_Y, 100, TxtH(TS_LABEL), _
               TS_LABEL, True, C_MUTED, -1, fmTextAlignLeft
        NewCtxCombo z, CStr(nm(i)), X, CTX_FLD_Y, CSng(w(i))
        X = X + CSng(w(i)) + GAP
    Next i
    ' PALETA INFO - "jos N gajbi do zatvaranja palete br. X". Stoji u eyebrow
    ' redu, uz robu na koju se odnosi (vrsta/sorta su u istom redu). Legacy ga
    ' ima kao runtime labelu ispod dugmadi (UpdatePaletaInfo).
    NewLbl z, "ctxPal", "", 0, 2, 420, TxtH(TS_MICRO), TS_MICRO, False, C_GREEN, -1, fmTextAlignRight
    NewLbl z, "ctxHint", "Ctrl+K", 0, 2, 52, TxtH(TS_MICRO), TS_MICRO, False, C_MUTED, -1, fmTextAlignRight
End Sub

'----------------------------------------------------------- FORM ----
Private Sub BuildForm(frm As Object)
    Dim z As Object, fr As Object
    Set z = NewZone(frm, "zForm", SIDEBAR_W, 0, 620, 220, C_WHITE)
    WireZone z

    ' Naslovi grupa (eyebrow + linija do desne ivice). Prave se jednom;
    ' LayoutFields ih pozicionira ili sakriva kad grupa u rezimu nema polja.
    Dim gi As Long, gk As Variant
    gk = GrpKeys()
    For gi = 0 To UBound(gk)
        NewLbl z, "grpH" & gi, UCase$(Poruka("OTKUI_GRP_" & gk(gi))), PAD, 0, GRP_LBL_W, _
               TxtH(TS_MICRO), TS_MICRO, True, C_MUTED, -1
        NewLbl z, "grpLn" & gi, "", 0, 0, 100, 1, 8, False, 0, C_BORDER
    Next gi

    NewFieldG z, "fgBrOtpr", Poruka("OTKUI_FLD_BROJ_OTPR"), "txt", "", 1, False, True, "DOK"
    ' PARTNER nije ovde - on je poslednje polje u "Osnovni podaci" (cbKupac),
    ' jer je isto pitanje "s kim radimo" kao otkupno mesto i vozac.
    NewFieldG z, "fgBrZbir", Poruka("OTKUI_FLD_BROJ_ZBIRNE"), "cmb", "", 1, False, False, "DOK"
    ' DATUM dokumenta. Do sada je bio precutno "danas" - stajao je samo u
    ' naslovu ekrana, pa se zaostali dokument nije mogao uneti sa svojim
    ' datumom. Ide u grupu DOKUMENT, u trecu kolonu koja je u svakom rezimu
    ' ostajala prazna, i vidljiv je u SVIM rezimima - datum ima svaki dokument.
    NewFieldG z, "fgDatum", Poruka("OTKUI_FLD_DATUM"), "txt", "", 1, False, False, "DOK"
    NewFieldG z, "fgNovac", Poruka("OTKUI_FLD_NOVAC"), "txt", Poruka("OTKUI_UNIT_RSD"), 1, True, False, "NOV"
    NewFieldG z, "fgKgI", Poruka("OTKUI_FLD_KG_I"), "txt", Poruka("OTKUI_UNIT_KG"), 1, True, False, "ROBA"
    NewFieldG z, "fgKgII", Poruka("OTKUI_FLD_KG_II"), "txt", Poruka("OTKUI_UNIT_KG"), 1, True, False, "ROBA"
    ' KLASA I CENA - jedno polje: prekidac klase (I / II) pa cena za svaku.
    ' Klasa i cena su ista odluka ("koja roba, po kojoj ceni") i zajedno staju
    ' u prostor koji je ranije zauzimala sama cena - tako ceo blok robe
    ' (kolicine + klasa + cene) ide u JEDAN red.
    Set fr = NewFrame(z, "fgCena", 0, 0, 180, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:1:ROBA"
    NewLbl fr, "fgCenaL", UCase$(Poruka("OTKUI_FLD_KLASA_CENA")), _
           0, 0, 170, TxtH(TS_LABEL), TS_LABEL, True, C_MUTED, -1
    NewSegBtn fr, "segKlasa1", "I", 0, 17, SEG_KL_W, FIELD_H - 2, True
    NewSegBtn fr, "segKlasa2", "II", SEG_KL_W + 1, 17, SEG_KL_W, FIELD_H - 2, False
    NewCenaBox fr, "fgCena1", "I", 70, 54
    NewCenaBox fr, "fgCena2", "II", 128, 54

    NewFieldG z, "fgKolAmb", Poruka("OTKUI_FLD_KOL_AMB"), "txt", Poruka("OTKUI_UNIT_KOM"), 1, True, False, "AMB"

    NewFieldG z, "fgTipAmb", Poruka("OTKUI_FLD_TIP_AMB"), "cmb", "", 1, False, False, "AMB"
    ' Gajbe II klase - zasebne od gajbi I klase, jer nose svoju taru. Vide se
    ' samo kad je II klasa ukljucena (legacy runtime polje txtKolAmbalazeII).
    NewFieldG z, "fgKolAmbII", Poruka("OTKUI_FLD_KOL_AMB_II"), "txt", Poruka("OTKUI_UNIT_KOM"), 1, True, False, "AMB"

    ' PRAZNA AMBALAZA - izdata uz otkup (F1) ili vracena uz prijemnicu (F4).
    ' Natpis postavlja SelectModeCore. Malo dugme "=" prepisuje kolicinu pune
    ' ambalaze, jer je razmena jedan-za-jedan pravilo a ne izuzetak.
    NewFieldG z, "fgAmbPr", "", "txt", Poruka("OTKUI_UNIT_KOM"), 1, True, False, "AMB"
    Set fr = z.Controls("fgAmbPr")
    NewLbl fr, "fgAmbPrEQB", "", 0, 17, 26, FIELD_H - 2, 8, False, 0, C_SAND
    NewLbl fr, "fgAmbPrEQ", ChrW(IC_KOPIRAJ), 0, CenterIco(17, FIELD_H - 2, TS_META), _
           26, TxtH(TS_META), TS_META, False, C_GREEN, -1, fmTextAlignCenter, F_ICON
    fr.Controls("fgAmbPrEQ").ZOrder 0
    fr.Controls("fgAmbPrEQ").ControlTipText = Poruka("OTKUI_TIP_AMB_EQ")
    ' "chev": ikonica koja pokriva svoju podlogu ne sme da menja boju na hover
    ' (isti razlog kao strelica dropdown-a), a klik i dalje stize.
    WireBtn fr.Controls("fgAmbPrEQ"), "btnAmbEq", "chev"

    ' PARCELA - samo otkupni list i samo kad je pracenje parcela ukljuceno
    ' (isti flag koji gasi cmbParcela u frmOtkup). Lista zavisi od izabranog
    ' kooperanta, kao frmOtkup.cmbKooperant_Change.
    NewFieldG z, "fgParcela", Poruka("OTKUI_FLD_PARCELA"), "cmb", "", 1, False, False, "DOK"

    ' SMER REVERSA - cetiri medjusobno iskljuciva segmenta, isti skup kao
    ' frmDokumenta.SetupOMIzdavanjeToggle. Span 2 da sva cetiri stanu u red.
    Set fr = NewFrame(z, "fgSmerRev", 0, 0, 370, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:3:AMB"
    NewLbl fr, "fgSmerRevL", Poruka("OTKUI_FLD_SMER_REV"), 0, 0, 300, 12, TS_LABEL, True, C_MUTED, -1
    NewShell fr, "fgSmerRev", 0, 16, 370, FIELD_H, C_BORDER_LT, C_WHITE
    ' NIJEDAN segment nije unapred obelezen: smer se bira eksplicitno. Ranije je
    ' segRev1 izgledao izabrano a mSmerRev je bio 0, pa je forma pokazivala smer
    ' koji dokument nije imao.
    NewSegBtn fr, "segRev1", Poruka("OTKUI_SEG_REV_IZD_KOOP"), 1, 17, 91, FIELD_H - 2, False
    NewSegBtn fr, "segRev2", Poruka("OTKUI_SEG_REV_PRI_KOOP"), 93, 17, 91, FIELD_H - 2, False
    NewSegBtn fr, "segRev3", Poruka("OTKUI_SEG_REV_IZD_OM"), 185, 17, 91, FIELD_H - 2, False
    NewSegBtn fr, "segRev4", Poruka("OTKUI_SEG_REV_PRI_OM"), 277, 17, 91, FIELD_H - 2, False


    ' VREDNOST - forest ploca (isti shell, druga ispuna). Ploca ne zauzima celo
    ' polje: levo od nje stoji racunica po klasi, pa se vidi ODAKLE iznos.
    ' Ploca iznosa NEMA grupu: ne redja se sa poljima nego deli red sa dugmadima
    ' (iznos i potvrda su jedna radnja), pa ne trosi ceo red za sebe.
    ' Okvir ploce je visok TACNO kao ploca (FIELD_H) i pocinje na nuli. Ranije
    ' je pozajmljivao geometriju polja - okvir 46pt sa praznih 16pt gore, gde
    ' kod polja stoji natpis. Ploca natpis nema (nosi ga u sebi), pa je taj
    ' prazan BEO pojas ulazio 5pt u poslednji red polja i brisao mu donju
    ' ivicu. Slanje na dno z-reda to nije resilo - okvir i dalje prekriva ono
    ' preko cega stoji - pa praznog pojasa vise nema.
    Set fr = NewFrame(z, "fgVrednost", 0, 0, 180, FIELD_H, C_WHITE)
    fr.tag = "fld:1:"
    NewShell fr, "fgVrednost", 0, 0, 180, FIELD_H, C_FOREST, C_FOREST
    NewLbl fr, "fgVrednostA", "", 1, 1, 3, FIELD_H - 2, 8, False, 0, C_GOLD
    NewLbl fr, "fgVrednostL", Poruka("OTKUI_LBL_VREDNOST"), INPUT_PAD, CenterY(0, FIELD_H, TS_MICRO), _
           66, TxtH(TS_MICRO), TS_MICRO, False, RGB(166, 178, 160), -1
    NewLbl fr, "fgVrednostV", "0", 0, CenterY(0, FIELD_H, TS_VAL), 100, TxtH(TS_VAL), _
           TS_VAL, True, C_CREAM, -1, fmTextAlignRight, F_NUM
    NewLbl fr, "fgVrednostU", Poruka("OTKUI_UNIT_RSD"), 0, CenterY(0, FIELD_H, TS_MICRO), _
           28, TxtH(TS_MICRO), TS_MICRO, True, C_GOLD, -1, fmTextAlignRight
    ' dva reda racunice - po jedan za svaku klasu; stoje DESNO od ploce
    NewLbl fr, "fgVrednostK1", "", 0, 2, 120, TxtH(TS_MICRO), TS_MICRO, False, C_DISABLED_FG, -1
    NewLbl fr, "fgVrednostK2", "", 0, 15, 120, TxtH(TS_MICRO), TS_MICRO, False, C_DISABLED_FG, -1
    ' KILOGRAMI - desno od racunice po klasama. U bruto rezimu pokazuje neto
    ' posle tare, inace zbir klasa. Legacy je to imao kao lblUkupnoKG uz polje
    ' kolicine; ovde stoji uz iznos, jer se cita zajedno sa njim.
    NewLbl fr, "fgVrednostKg", "", 0, CenterY(0, FIELD_H, TS_META), 250, TxtH(TS_META), _
           TS_META, True, C_FOREST, -1, fmTextAlignLeft, F_NUM

    ' OTVORENI BLOK + NEISPLACENI OSTATAK (samo isplate). U frmDokumenta je to
    ' cmbOtkupBlok, poznat kao "preostali kes" - poslovno to nije kes nego
    ' neisplaceni ostatak po otkupnom bloku, pa polje tako i pise.
    NewFieldG z, "fgBlok", Poruka("OTKUI_FLD_BLOK"), "cmb", "", 1, False, False, "NOV"
    Set fr = NewFrame(z, "fgOstatak", 0, 0, 180, FIELD_H, C_WHITE)
    fr.tag = "fld:1:"
    NewShell fr, "fgOstatak", 0, 0, 180, FIELD_H, C_FOREST, C_FOREST
    NewLbl fr, "fgOstatakL", Poruka("OTKUI_LBL_OSTATAK"), INPUT_PAD, CenterY(0, FIELD_H, TS_MICRO), _
           86, TxtH(TS_MICRO), TS_MICRO, False, RGB(166, 178, 160), -1
    NewLbl fr, "fgOstatakV", "0", 0, CenterY(0, FIELD_H, TS_VAL), 100, TxtH(TS_VAL), _
           TS_VAL, True, C_CREAM, -1, fmTextAlignRight, F_NUM
    NewLbl fr, "fgOstatakU", Poruka("OTKUI_UNIT_RSD"), 0, CenterY(0, FIELD_H, TS_MICRO), _
           28, TxtH(TS_MICRO), TS_MICRO, True, C_GOLD, -1, fmTextAlignRight

    ' ISPLATA IZ (samo F5) - legacy tglIzOMAvansa. Nije kozmetika: prekidac
    ' bira TIP NOVCA (kes iz avansa otkupnog mesta vs virman firme), a od tipa
    ' zavisi da li se isplata skida sa avansa OM-a. Bez njega bi svaka isplata
    ' po otkupnom bloku bila virman, pa avans OM-a nikad ne bi bio razduzen.
    ' Dva medjusobno iskljuciva segmenta - isti obrazac kao klasa i smer reversa.
    Set fr = NewFrame(z, "fgAvans", 0, 0, 250, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:2:NOV"
    NewLbl fr, "fgAvansL", Poruka("OTKUI_FLD_AVANS"), 0, 0, 240, 12, TS_LABEL, True, C_MUTED, -1
    NewShell fr, "fgAvans", 0, 16, 250, FIELD_H, C_BORDER_LT, C_WHITE
    NewSegBtn fr, "segAvans1", Poruka("OTKUI_SEG_AVANS_VIRMAN"), 1, 17, 123, FIELD_H - 2, True
    NewSegBtn fr, "segAvans2", Poruka("OTKUI_SEG_AVANS_OM"), 125, 17, 123, FIELD_H - 2, False

    ' RASPOLOZIV AVANS OTKUPNOG MESTA (legacy UpdateOMAvansSaldo) NIJE zasebna
    ' ploca nego stoji u NATPISU ovog polja (RefreshSaldoOM). Dva razloga:
    ' akcioni red vec nosi jednu plocu (VisibleValFrame bira izmedju VREDNOST i
    ' NEISPLACENO, treca nema gde), a broj se ionako cita bas u trenutku izbora
    ' "kes iz OM avansa" - iz avansa se ne moze isplatiti vise nego sto ga ima.

    ' OTVORENA FAKTURA (samo F6) - legacy cmbFakturaIzlaz + FillOpenFakture.
    ' Izbor fakture je jedina razlika izmedju uplate po fakturi i avansa kupca:
    ' bez njega nijedna faktura iz novog UI-ja ne bi bila zatvorena.
    NewFieldG z, "fgFaktura", Poruka("OTKUI_FLD_FAKTURA"), "cmb", "", 2, False, False, "NOV"

    SetDatumDanas z

    ' akcije: sekundarno pa primarno, desno poravnato
    BtnV z, "btnOtkazi", Poruka("OTKUI_BTN_OTKAZI"), 0, 0, 72, 28, "plain"
    BtnV z, "btnSacuvajPrint", Poruka("OTKUI_BTN_SACUVAJ_PRINT"), 0, 0, 132, 28, "secondary", IC_PRINT
    ' Natpis je kontekstualan ("Sacuvaj otkupni list"...), postavlja ga
    ' SelectModeCore - zato je dugme sire nego generickim "Sacuvaj".
    BtnV z, "btnSacuvaj", Poruka("OTKUI_BTN_SACUVAJ_OTKUP"), 0, 0, 196, 28, "primary", IC_SAVE
End Sub

'---------------------------------------------------------- PANEL ----
' Prazan okvir preko radne povrsine. Ljuska ga samo pravi, meri i pali/gasi --
' nikad ne stavlja nista u njega. Sadrzaj gradi modul panela, kroz PanelHost.
'
' ZASTO OKVIR, a ne sama forma: graditelji panela (modPodesavanja, modAdmin)
' prvo sakriju SVE kontrole domacina. Nad formom bi to ugasilo celu ljusku;
' nad namenskim okvirom gase samo ono sto je u njemu, a to je njihovo.
Private Sub BuildPanelHost(frm As Object)
    NewZone frm, "zPanel", SIDEBAR_W, HEADER_H, 620, 260, C_SAND_LT
End Sub

' Okvir u koji modul panela gradi svoje kontrole. Vraca Nothing ako ljuska nije
' ziva -- pozivalac tada nema gde da gradi i to mora da vidi.
Public Function PanelHost() As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Function
    Set PanelHost = mFrm.Controls("zPanel")
    If Err.Number <> 0 Then Err.Clear
End Function

' Ustupanje i vracanje radne povrsine. Ljuska ne pita ZASTO -- samo prerasporedi
' i prekrije. Sadrzaj okvira ostaje netaknut: prazni ga onaj ko ga je i napunio.
Public Sub PanelRezim(ByVal ukljuci As Boolean)
    If mPanelRezim = ukljuci Then Exit Sub
    mPanelRezim = ukljuci
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    ShowZones mFrm
    LayoutOtkup mFrm
End Sub

Public Function PanelRezimAktivan() As Boolean
    PanelRezimAktivan = mPanelRezim
End Function

' Radna povrsina se VRATILA ekranu ispod panela. Zove ga modUiPanel.PanelZatvori.
'
' Otvaranje panela je taj ekran deaktiviralo (ScrDeaktiviraj), pa mu je stanje
' obrisano -- kod maticnih i mZonaEkran, koju modMaticniEkran.Zona() trazi za
' SVAKU radnju. Sam raspored to ne vraca: mreza bi se videla nepromenjena, a
' "Izmeni" bi tiho nista ne radila, jer bi Zona() vracala Nothing.
'
' Uz to se vraca i oznaka u sidebaru: dok je panel bio otvoren ona je stajala
' na njemu, a panela vise nema.
Public Sub PanelVracenNaEkran()
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If mPanelRezim Then Exit Sub          ' jos je ustupljena -- nema sta da se vraca
    PaintNav mFrm, NavTagFor(mScreen)
    LayoutOtkup mFrm
    If Len(mScreen) > 0 Then RefreshFromData
    Err.Clear
End Sub

'----------------------------------------------------------- GRID ----
Private Sub BuildGrid(frm As Object)
    Dim z As Object, i As Long, chip As Variant, X As Single, hd As Object, ft As Object
    Dim tf As Object
    Set z = NewZone(frm, "zGrid", SIDEBAR_W, 0, 620, 260, C_SAND_LT)
    WireZone z
    NewLbl z, "grdLnT", "", 0, 0, 620, 1, 8, False, 0, C_BORDER

    NewLbl z, "grdTitle", "-", PAD, 10, 180, 16, TS_H1, True, C_FOREST, -1
    NewLbl z, "grdSrc", "-", PAD + 176, 11, 120, 14, TS_MICRO, False, C_MUTED, C_SAND, fmTextAlignCenter, F_NUM
    ' PREKIDAC LISTA - prazan. Ko su liste i kako se zovu zna EKRAN (Scr_Liste),
    ' ne ljuska: F1 ima pet (svi listovi, otpremnice, blokovi, izgubljeni,
    ' kooperanti), Palete tri (palete, stavke, prerade), a vecina ekrana nijednu.
    ' Zato su dugmad bezimena - lsSeg0..4 - a natpis i kljuc dobijaju pri
    ' rasporedu. Tako se nijedan kljuc liste ne pominje u ljusci.
    For i = 0 To MAX_SEG - 1
        NewSegBtn z, "lsSeg" & i, "", 0, 9, 96, 20, (i = 0)
    Next i
    NewShell z, "srch", 0, 8, 200, 22, C_INPUT_BORDER, C_WHITE
    NewLbl z, "srchIco", ChrW(IC_SEARCH), 0, CenterIco(8, 22, TS_META), 16, TxtH(TS_META), _
           TS_META, False, C_MUTED, -1, fmTextAlignCenter, F_ICON
    NewTxt z, "txtSearch", "", 0, 12, 170, 16, False
    NewLbl z, "lblSearchPh", Poruka("OTKUI_PH_PRETRAGA"), 0, 12, 168, 14, TS_META, False, C_DISABLED_FG, -1
    BtnV z, "btnFilteri", Poruka("OTKUI_BTN_FILTERI"), 0, 8, 88, 22, "ghost", IC_FILTER
    BtnV z, "btnMax", "", 0, 8, 26, 22, "ghost", IC_MAX

    chip = ChipRow()
    X = PAD
    For i = 0 To UBound(chip)
        Dim p As Variant: p = Split(chip(i), "|")
        ChipV z, CStr(p(0)), Poruka(CStr(p(1))), X, 36, CSng(p(2)), 19, (i = 0), (i = 5)
        X = X + CSng(p(2)) + 6
    Next i

    ' RADNJE NAD IZABRANIM REDOM - u redu cipova, uz desnu ivicu. Ljuska ne zna
    ' sta rade niti koliko ih ima: ekran ih prijavi kroz Scr_Radnje, a klik se
    ' vraca kao "act:<kljuc>:<red>". Dugmad su zato bezimena - btnAct0..3.
    For i = 0 To MAX_ACT - 1
        BtnV z, "btnAct" & i, "", 0, 36, 100, 19, "ghost"
    Next i
    ' OD / DO za specifikaciju po datumu (legacy "Stampaj po datumu"). Ovo NISU
    ' polja dokumenta nego parametri stampe - zato ne prljaju formu (UiChange
    ' ih preskace pri MarkDirty).
    NewLbl z, "specOdL", Poruka("OTKUI_LBL_OD"), 0, 39, 18, 12, TS_MICRO, True, C_MUTED, -1
    NewShell z, "specOd", 0, 36, 74, 20, C_INPUT_BORDER, C_WHITE
    NewTxt z, "specOdT", Format$(Date, "dd.mm.yyyy"), 0, 39, 62, 14, False
    NewLbl z, "specDoL", Poruka("OTKUI_LBL_DO"), 0, 39, 18, 12, TS_MICRO, True, C_MUTED, -1
    NewShell z, "specDo", 0, 36, 74, 20, C_INPUT_BORDER, C_WHITE
    NewTxt z, "specDoT", Format$(Date, "dd.mm.yyyy"), 0, 39, 62, 14, False
    BtnV z, "btnRedSpecDat", Poruka("OTKUI_BTN_RED_SPECDAT"), 0, 36, 112, 19, "ghost"

    ' zaglavlje mreze - sami crtamo (ListBox zaglavlje se ne moze stilizovati)
    Set hd = NewFrame(z, "grdHead", 0, 0, 620, GRID_HEAD_H, C_HEAD_BG)
    For i = 0 To MAX_COLS - 1
        NewSortHdr hd, "hd" & i, "", 0, CenterY(0, GRID_HEAD_H, TS_MICRO), 60, TxtH(TS_MICRO), False
    Next i

    BuildGridRows z
    NewLbl z, "grdEmpty", Poruka("OTKUI_GRID_PRAZNO"), 0, 0, 320, 16, TS_META, False, C_MUTED, -1, fmTextAlignCenter
    ' izlaz iz praznog stanja u jednom kliku, bez trazenja cipa
    NewLbl z, "grdEmptyA", Poruka("OTKUI_GRID_PRIKAZI_SVE"), 0, 0, 320, 16, TS_META, _
           True, C_GREEN, -1, fmTextAlignCenter
    WireBtn z.Controls("grdEmptyA"), "grdEmptyA", "chev"

    Set ft = NewFrame(z, "grdFoot", 0, 0, 620, GRID_FOOT_H, C_SAND_LT)

    ' TOAST -- traka preko cele sirine mreze, tacno IZNAD podnozja.
    '
    ' Do v6-ui-145 je ziveo u zForm, zoni unosnog ekrana, koju ShowZones gasi na
    ' svakom ugovornom ekranu: Visible = True nad kontrolom u skrivenom roditelju
    ' ne prikazuje nista, pa poruke sa ekrana Storno, Palete i Oporavak nisu
    ' stizale do operatera. Prva popravka ga je stavila u naslovnu traku -- video
    ' se, ali je pokrivao naslov ekrana.
    '
    ' Mreza je bolji domacin od naslova: vidljiva je na svim ekranima isto kao i
    ' naslov, ali joj je dno PRAZNO kad lista ne popuni stranu, pa poruka nista ne
    ' zaklanja. A i odgovor na radnju nad redom stoji uz same redove.
    Set tf = NewFrame(z, "tstScr", 0, 0, 320, 26, C_SOFT_BG)
    tf.Visible = False
    NewLbl tf, "tstScrBar", "", 0, 0, 3, 26, 8, False, 0, C_GREEN_LT
    NewLbl tf, "tstScrMsg", "", 11, 6, 300, 15, TS_META, True, C_GREEN, -1
    NewLbl ft, "ftLnT", "", 0, 0, 620, 1, 8, False, 0, C_BORDER
    NewLbl ft, "ftRange", "-", PAD, 6, 170, 14, TS_META, False, C_MUTED, -1
    ' zbirovi su brojevi - Consolas, podebljano (Consolas nema Semibold rez)
    NewLbl ft, "ftKg", "-", 0, 6, 170, 14, TS_CELL_LG, True, C_FOREST, -1, fmTextAlignRight, F_NUM
    NewLbl ft, "ftVal", "-", 0, 6, 190, 14, TS_CELL_LG, True, C_FOREST, -1, fmTextAlignRight, F_NUM
    NewLbl ft, "ftVal2", "-", 0, 6, 190, 14, TS_CELL_LG, True, C_FOREST, -1, fmTextAlignRight, F_NUM
    ' visina 20 (ne 15): TxtH(TS_META) = 16, pa u 15pt dugme tekst nije stao
    BtnV ft, "pgPrev", ChrW(8249), 0, 2, 21, 20, "page"
    For i = 1 To 5
        BtnV ft, "pg" & i, CStr(i), 0, 2, 21, 20, "page"
        ft.Controls("pg" & i & "C").Font.name = F_NUM
        ft.Controls("pg" & i & "C").Font.Size = TS_CELL_SM
    Next i
    BtnV ft, "pgNext", ChrW(8250), 0, 2, 21, 20, "page"
End Sub

' Redovi mreze kao crtani Label-ovi. Prave se JEDNOM; render ih samo puni.
' Zato imamo status-pilule u boji, hover reda i tacan padding po koloni.
Private Sub BuildGridRows(z As Object)
    Dim body As Object, r As Long, c As Long, Y As Single
    Set body = NewFrame(z, "grdBody", 0, 0, 620, MAX_ROWS * GRID_ROW_H, C_WHITE)
    For r = 0 To MAX_ROWS - 1
        Y = r * GRID_ROW_H
        NewCell body, "rb" & r, 0, Y, 620, GRID_ROW_H, C_WHITE, C_FOREST, TS_BODY, F_UI, fmTextAlignLeft
        NewLbl body, "rl" & r, "", 0, Y + GRID_ROW_H - 1, 620, 1, 8, False, 0, C_ROW_LN
        ' celije se prave za MAX_COLS; koliko ih je vidljivo odlucuje rezim
        For c = 0 To MAX_COLS - 1
            NewCell body, "c" & r & "_" & c, 0, CenterY(Y, GRID_ROW_H, TS_BODY), 60, TxtH(TS_BODY), _
                    C_WHITE, C_FOREST, TS_BODY, F_UI, fmTextAlignLeft
        Next c
    Next r
End Sub

'---------------------------------------------------------- RIGHT ----
Private Sub BuildRight(frm As Object)
    Dim z As Object, i As Long
    Set z = NewZone(frm, "zRight", 0, 0, RIGHT_W, 320, C_SAND_LT)
    WireZone z
    NewLbl z, "rgtEdge", "", 0, 0, 1, 320, 8, False, 0, C_BORDER
    NewSectionHdr z, "rgtHdr", Poruka("OTKUI_LBL_REZIMI"), 13, 12, 150

    ' SEDAM kartica - prikazani su SVI unosni rezimi, ukljucujuci aktivni, a
    ' aktivni je obelezen (tamna ispuna + zlatna traka). Ranije se aktivni
    ' izostavljao, pa se spisak premetao pri svakoj promeni i nije se videlo
    ' gde si. Redosled je fiksan F1..F7, kartica nosi rezim u .Tag.
    '
    ' Osma kartica je do v6-ui-143 bila Storno, odvojena linijom jer nije
    ' unosni rezim. Storno je sada svoj EKRAN i stoji u sidebaru uz Dokumenta
    ' i Oporavak, pa ovde nema sta da trazi -- ostalo je sedam ravnopravnih.
    For i = 0 To 6
        Dim cy As Single: cy = 30 + i * 40
        CardV z, "mc" & i, 13, cy, RIGHT_W - 26, 36
        NewLbl z, "mcT" & i, "-", 21, cy + 5, 108, TxtH(TS_META), TS_META, True, C_FOREST, -1
        NewLbl z, "mcKey" & i, "", 0, cy + 6, 28, TxtH(TS_MICRO), TS_MICRO, False, C_MUTED, -1, fmTextAlignRight, F_NUM
        NewLbl z, "mcHint" & i, "-", 21, cy + 20, RIGHT_W - 40, TxtH(TS_MICRO), TS_MICRO, False, C_MUTED, -1
        ' CELA kartica je meta klika, ne samo traka naslova. Ranije je bio vezan
        ' samo mcT (108 x 12pt u kartici od 36pt), pa je klik cesto promasivao -
        ' a MSForms Label ne okida Click ako se mis izmedju pritiska i otpustanja
        ' pomeri na drugu kontrolu, sto je na tako uskoj meti stalno.
        ' Ispuna nosi "card", tekstovi "chev" (klik da, hover ne) - isti obrazac
        ' kao u BoxText za natpis dugmeta.
        WireBtn z.Controls("mc" & i & "F"), "mcT" & i, "card"
        WireBtn z.Controls("mcT" & i), "mcT" & i, "chev"
        WireBtn z.Controls("mcHint" & i), "mcT" & i, "chev"
        WireBtn z.Controls("mcKey" & i), "mcT" & i, "chev"
    Next i

    NewLbl z, "rgtSep1", "", 13, 0, RIGHT_W - 26, 1, 8, False, 0, C_BORDER
    NewSectionHdr z, "rgtHdr2", Poruka("OTKUI_LBL_POSLEDNJI_UNOSI"), 13, 0, 150
    ' Tri reda crtica su izgledala kao pokvaren podatak. Dok sekcija nije
    ' vezana za izvor, stoji jedan red teksta; redovi ostaju napravljeni da ih
    ' punjenje kasnije samo otkrije.
    For i = 0 To 2
        NewLbl z, "luK" & i, "", 13, 0, 82, 14, TS_META, False, C_FOREST, -1, fmTextAlignLeft, F_NUM
        NewLbl z, "luV" & i, "", 0, 0, 90, 14, TS_META, True, C_FOREST, -1, fmTextAlignRight, F_NUM
        NewLbl z, "luLn" & i, "", 13, 0, RIGHT_W - 26, 1, 8, False, 0, RGB(240, 243, 238)
        z.Controls("luK" & i).Visible = False
        z.Controls("luV" & i).Visible = False
        z.Controls("luLn" & i).Visible = False
    Next i
    NewLbl z, "luEmpty", Poruka("OTKUI_LU_PRAZNO"), 13, 0, RIGHT_W - 26, 14, _
           TS_MICRO, False, RGB(146, 158, 140), -1

    NewLbl z, "rgtSep2", "", 13, 0, RIGHT_W - 26, 1, 8, False, 0, C_BORDER
    NewSectionHdr z, "rgtHdr3", Poruka("OTKUI_LBL_SINHRONIZACIJA"), 13, 0, 150
    NewLbl z, "syncDot", ChrW(IC_SYNC), 13, 0, 16, TxtH(TS_META), TS_META, False, C_RUST, -1, fmTextAlignLeft, F_ICON
    NewLbl z, "syncTxt", Poruka("OTKUI_SYNC_NIJE"), 30, 0, RIGHT_W - 44, 14, _
           TS_META, False, RGB(146, 158, 140), -1
    BtnV z, "btnSync", Poruka("OTKUI_BTN_SYNC"), 13, 0, RIGHT_W - 26, 22, "soft", IC_SYNCBTN
End Sub

'--------------------------------------------------------- STATUS ----
Private Sub BuildStatus(frm As Object)
    Dim z As Object
    Set z = NewZone(frm, "zStatus", 0, 0, 1100, STATUS_H, C_SAND)
    WireZone z
    NewLbl z, "stLnT", "", 0, 0, 1100, 1, 8, False, 0, C_BORDER
    NewLbl z, "stLeft", StatusLeftText(""), PAD, 6, 420, 14, TS_MICRO, False, C_MUTED, -1
    NewLbl z, "stDot", "", 0, 9, 6, 6, 8, False, 0, C_GREEN_LT
    NewLbl z, "stSys", Poruka("OTKUI_ST_SISTEMI"), 0, 6, 170, 14, TS_META, True, C_GREEN, -1
    ' bez ikonice: vizija u tom uglu ima samo tekst, a ikonica je zbog desnog
    ' poravnanja teksta ostajala odvojena od njega
    NewLbl z, "stHelp", Poruka("OTKUI_ST_POMOC"), 0, 6, 70, 14, TS_MICRO, False, C_MUTED, -1, fmTextAlignRight
End Sub

'=====================================================================
' SOPSTVENI DROPDOWN
'
' Nativna lista ComboBox-a je Windows prozor: ivica, pozadina, visina reda i
' plava selekcija se NE mogu stilizovati (jedini deo koji MSForms uopste ne
' daje pod kontrolu). Zato je lista nasa - Frame sa Label redovima, iznad
' svih zona. ComboBox ostaje ispod i dalje nosi podatke (.List, .ListIndex,
' skrivenu ID kolonu iz FillComboDisplayID) i kucanje sa MatchEntry.
'
' Panel je JEDAN za ceo ekran i pravi se jednom; redovi se pune po potrebi.
'=====================================================================
Private Sub BuildPopup(frm As Object)
    Dim z As Object, i As Long
    Set z = NewFrame(frm, "zPop", 0, 0, 180, 40, C_WHITE)
    z.Visible = False
    ' lenjir za merenje teksta - MSForms nema drugi nacin da se sazna sirina
    ' stringa; Label sa AutoSize sam sebi postavi Width po sadrzaju
    NewLbl z, "popRul", "", 0, -400, 10, 14, TS_BODY, False, C_WHITE, -1
    z.Controls("popRul").AutoSize = True
    NewLbl z, "popBd", "", 0, 0, 180, 40, 8, False, 0, C_BORDER
    NewLbl z, "popBg", "", 1, 1, 178, 38, 8, False, 0, C_WHITE
    For i = 0 To POP_MAX - 1
        ' v6-ui-188: pozadina/hit ostaje PUN red (21pt), a TEKST ide u
        ' zaseban unutrasnji label visine TxtH(TS_BODY) centriran kroz
        ' CenterY -- ISTO rasterizacijsko pravilo kao mreza. Label pune
        ' visine je GDI-ju pomerao baseline fazu po redu, pa je npr. osmi
        ' red izgledao kao "veci font" iako je Font.Size svuda isti
        ' (Sledljivost smoke, S15). Oba nose ISTI tag: klik radi i preko
        ' teksta, a hover se boji CENTRALNO (PopHoverRed) jer su tekst i
        ' pozadina sada odvojene kontrole.
        NewLbl z, "pop" & i, "", 1, 1 + i * POP_ITEM_H, 178, POP_ITEM_H, _
               TS_BODY, False, C_FOREST, C_WHITE
        WireBtn z.Controls("pop" & i), "pop" & i, "row"
        NewLbl z, "popT" & i, "", 1, _
               CenterY(1 + i * POP_ITEM_H, POP_ITEM_H, TS_BODY), _
               178, TxtH(TS_BODY), TS_BODY, False, C_FOREST, -1
        WireBtn z.Controls("popT" & i), "pop" & i, "row"
    Next i
    ' Podnozje sa brojacem i strelicama - vidljivo samo kad lista ne staje u
    ' POP_MAX redova. Time panel pokriva i duge liste (kooperanti), pa nema
    ' vise vracanja na nativnu listu i njen Windows okvir.
    NewLbl z, "popFt", "", 1, 1, 178, POP_ITEM_H, 8, False, 0, C_SAND_LT
    NewLbl z, "popCnt", "", 8, 1, 110, TxtH(TS_MICRO), TS_MICRO, False, C_MUTED, -1
    NewLbl z, "popUp", ChrW(IC_GORE), 1, 1, 24, TxtH(TS_META), _
           TS_META, False, C_FOREST, -1, fmTextAlignCenter, F_ICON
    NewLbl z, "popDn", ChrW(IC_DOLE), 1, 1, 24, TxtH(TS_META), _
           TS_META, False, C_FOREST, -1, fmTextAlignCenter, F_ICON
    WireBtn z.Controls("popUp"), "popUp", "chev"
    WireBtn z.Controls("popDn"), "popDn", "chev"
End Sub

Private Sub OpenPopup(ByVal comboNm As String)
    Dim cmb As Object
    On Error Resume Next
    Set cmb = FindCombo(comboNm)
    If cmb Is Nothing Then Exit Sub
    If cmb.ListCount = 0 Then
        ClosePopup
        Exit Sub
    End If
    mPopFor = comboNm
    mPopTop = 0
    mPopHover = -1
    PopIndex cmb
    PopRender
End Sub

' Koje stavke panel prikazuje. Suzavanje je po PODNIZU, ne po pocetku reci -
' "pilica" nalazi i "Milosavljevic Dragan - Pilica". Zato je MatchEntry na
' combo-u ugasen: nativno dopunjavanje teksta bi se tuklo sa ovim filterom.
Private Sub PopIndex(cmb As Object)
    Dim i As Long, f As String
    mPopN = 0
    If cmb.ListCount = 0 Then Exit Sub
    ReDim mPopSrc(0 To cmb.ListCount - 1)
    f = Trim$(CStr(cmb.text))
    For i = 0 To cmb.ListCount - 1
        If Len(f) = 0 Then
            mPopSrc(mPopN) = i
            mPopN = mPopN + 1
        ElseIf InStr(1, CStr(cmb.List(i, 0)), f, vbTextCompare) > 0 Then
            mPopSrc(mPopN) = i
            mPopN = mPopN + 1
        End If
    Next i
End Sub

Private Sub PopRender()
    Dim cmb As Object, z As Object, i As Long, vis As Long
    Dim ax As Single, ay As Single, h As Single, tw As Single
    Dim scroll As Boolean
    On Error Resume Next
    If Len(mPopFor) = 0 Then Exit Sub
    Set cmb = FindCombo(mPopFor)
    If cmb Is Nothing Then Exit Sub
    Set z = mFrm.Controls("zPop")
    If mPopN = 0 Then
        ClosePopup
        Exit Sub
    End If
    vis = mPopN
    If vis > POP_MAX Then vis = POP_MAX
    scroll = (mPopN > POP_MAX)
    If mPopTop > mPopN - vis Then mPopTop = mPopN - vis
    If mPopTop < 0 Then mPopTop = 0

    AbsPos cmb, ax, ay
    ' Panel mora da stane najduza stavka, ne sirina polja - "Willamette teren"
    ' je bio isecen jer je Sorta usko polje. Sire od polja je u redu; nas panel
    ' nije nativna lista pa nema ni horizontalnog skrolera da to prikrije.
    z.width = cmb.width + INPUT_PAD * 2 + 20
    tw = PopTextW(z, cmb) + 26
    If tw > z.width Then z.width = tw
    If z.width > 320 Then z.width = 320
    h = vis * POP_ITEM_H + 2
    If scroll Then h = h + POP_ITEM_H
    z.Height = h
    z.Controls("popBd").width = z.width: z.Controls("popBd").Height = h
    z.Controls("popBg").width = z.width - 2: z.Controls("popBg").Height = h - 2
    For i = 0 To POP_MAX - 1
        If i < vis Then
            ' Tekst nosi unutrasnji label (v6-ui-188, v. BuildPopup);
            ' pozadinski red ostaje prazan hit/hover sloj.
            z.Controls("pop" & i).caption = ""
            z.Controls("pop" & i).BackColor = C_WHITE
            z.Controls("pop" & i).top = 1 + i * POP_ITEM_H
            z.Controls("pop" & i).width = z.width - 2
            z.Controls("pop" & i).Visible = True
            z.Controls("popT" & i).caption = "   " & CStr(cmb.List(mPopSrc(mPopTop + i), 0))
            z.Controls("popT" & i).top = CenterY(1 + i * POP_ITEM_H, POP_ITEM_H, TS_BODY)
            z.Controls("popT" & i).width = z.width - 2
            z.Controls("popT" & i).Visible = True
            z.Controls("popT" & i).ZOrder 0
        Else
            z.Controls("pop" & i).Visible = False
            z.Controls("popT" & i).Visible = False
        End If
    Next i
    PopFooter z, vis, scroll
    z.Left = ax - INPUT_PAD
    z.top = ay + cmb.Height + 5
    ' panel bi ispod donjih polja izasao iz forme - tada se otvara nagore
    If z.top + h > mFrm.InsideHeight Then z.top = ay - h - 5
    If z.top < 0 Then z.top = 0
    z.Visible = True
    z.ZOrder 0
End Sub

' Sirina najduze stavke u trenutnom (suzenom) skupu. Meri se JEDNA stavka -
' ona sa najvise znakova - pa sirina ne skace pri skrolovanju.
Private Function PopTextW(z As Object, cmb As Object) As Single
    Dim i As Long, best As Long, bestLen As Long, t As String
    On Error Resume Next
    If mPopN = 0 Then Exit Function
    best = mPopSrc(0)
    For i = 0 To mPopN - 1
        t = CStr(cmb.List(mPopSrc(i), 0))
        If Len(t) > bestLen Then
            bestLen = Len(t)
            best = mPopSrc(i)
        End If
    Next i
    z.Controls("popRul").caption = "   " & CStr(cmb.List(best, 0)) & "   "
    PopTextW = z.Controls("popRul").width
End Function

Private Sub PopFooter(z As Object, ByVal vis As Long, ByVal scroll As Boolean)
    Dim Y As Single
    On Error Resume Next
    z.Controls("popFt").Visible = scroll
    z.Controls("popCnt").Visible = scroll
    z.Controls("popUp").Visible = scroll
    z.Controls("popDn").Visible = scroll
    If Not scroll Then Exit Sub
    Y = 1 + vis * POP_ITEM_H
    z.Controls("popFt").top = Y
    z.Controls("popFt").width = z.width - 2
    z.Controls("popCnt").top = CenterY(Y, POP_ITEM_H, TS_MICRO)
    z.Controls("popCnt").caption = (mPopTop + 1) & "-" & (mPopTop + vis) & " / " & mPopN
    z.Controls("popUp").top = CenterIco(Y, POP_ITEM_H, TS_META)
    z.Controls("popUp").Left = z.width - 56
    z.Controls("popDn").top = CenterIco(Y, POP_ITEM_H, TS_META)
    z.Controls("popDn").Left = z.width - 28
End Sub

' Skrol ide STRANU po stranu - sa dve strelice je to jedini raspored koji ne
' trazi desetine klikova na listi od stotinu imena.
' Kucanje u combo: prvi znak otvara panel, sledeci ga suzavaju. Tag promene
' je ime same kontrole (zCtx) ili ime polja + "T" (zForm), pa se oba oblika
' svode na ime koje FindCombo razume.
Private Sub PopFromTyping(ByVal tag As String)
    Dim cmb As Object, nm As String
    On Error Resume Next
    nm = tag
    Set cmb = FindCombo(nm)
    If cmb Is Nothing And Len(tag) > 1 And Right$(tag, 1) = "T" Then
        nm = Left$(tag, Len(tag) - 1)
        Set cmb = FindCombo(nm)
    End If
    If cmb Is Nothing Then Exit Sub
    If TypeName(cmb) <> "ComboBox" Then Exit Sub
    If Len(mPopFor) = 0 Then
        If Len(Trim$(CStr(cmb.text))) = 0 Then Exit Sub
        OpenPopup nm
        Exit Sub
    End If
    If mPopFor <> nm Then Exit Sub
    PopIndex cmb
    mPopTop = 0
    PopRender
End Sub

Private Sub PopScroll(ByVal d As Long)
    If Len(mPopFor) = 0 Then Exit Sub
    mPopTop = mPopTop + d * POP_MAX
    PopRender
End Sub

Private Sub ClosePopup()
    On Error Resume Next
    mPopFor = ""
    mPopN = 0
    mPopTop = 0
    mPopHover = -1
    If mFrm Is Nothing Then Exit Sub
    mFrm.Controls("zPop").Visible = False
End Sub

' Hover red panela -- pozadinski label se boji centralno (v6-ui-188),
' jer lokalno bojenje klase vise ne pokriva ceo red (tekst je odvojen).
Private Sub PopHoverRed(ByVal i As Long)
    Dim z As Object
    On Error Resume Next
    If Len(mPopFor) = 0 Then Exit Sub
    If i = mPopHover Then Exit Sub
    Set z = mFrm.Controls("zPop")
    If mPopHover >= 0 Then z.Controls("pop" & mPopHover).BackColor = C_WHITE
    z.Controls("pop" & i).BackColor = RGB(246, 244, 238)
    mPopHover = i
End Sub

Private Sub SelectPopup(ByVal idx As Long)
    Dim cmb As Object, nm As String, src As Long
    On Error Resume Next
    nm = mPopFor
    If Len(nm) = 0 Then Exit Sub
    If idx < 0 Or mPopTop + idx > mPopN - 1 Then
        ClosePopup
        Exit Sub
    End If
    src = mPopSrc(mPopTop + idx)
    ClosePopup
    Set cmb = FindCombo(nm)
    If cmb Is Nothing Then Exit Sub
    mPopMute = True                  ' upis okida Change - ne otvaraj panel ponovo
    cmb.ListIndex = src               ' povlaci i skrivenu ID kolonu
    mPopMute = False
End Sub

'=====================================================================
' PANEL "FILTERI"
' Cipovi pokrivaju najcesce preseke; panel dodaje ono sto u red cipova ne
' staje - mesec kao period i suzavanje na vrstu / partnera iz Osnovnih
' podataka. Uslovi se ne racunaju posebno: idu kroz isti "hay" string koji
' FillGrid vec gradi za pretragu, pa nema drugog puta kroz podatke.
'=====================================================================
Private Sub BuildFilterPanel(frm As Object)
    Dim z As Object
    Set z = NewFrame(frm, "zFlt", 0, 0, FLT_W, FLT_H, C_WHITE)
    z.Visible = False
    NewLbl z, "fltBd", "", 0, 0, FLT_W, FLT_H, 8, False, 0, C_BORDER
    NewLbl z, "fltBg", "", 1, 1, FLT_W - 2, FLT_H - 2, 8, False, 0, C_WHITE
    NewLbl z, "fltH1", Poruka("OTKUI_FLT_PERIOD"), 12, 12, 200, TxtH(TS_MICRO), _
           TS_MICRO, True, C_MUTED, -1
    BoxText z, "fltP0", Poruka("OTKUI_CHIP_DANAS"), 12, 30, 102, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    BoxText z, "fltP1", Poruka("OTKUI_CHIP_NEDELJA"), 122, 30, 102, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    BoxText z, "fltP2", Poruka("OTKUI_FLT_MESEC"), 12, 60, 102, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    BoxText z, "fltP3", Poruka("OTKUI_CHIP_SVE"), 122, 60, 102, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    NewLbl z, "fltH2", Poruka("OTKUI_FLT_SUZI"), 12, 98, 200, TxtH(TS_MICRO), _
           TS_MICRO, True, C_MUTED, -1
    BoxText z, "fltVrsta", Poruka("OTKUI_FLT_SAMO_VRSTA"), 12, 116, 212, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    BoxText z, "fltPart", Poruka("OTKUI_FLT_SAMO_PART"), 12, 146, 212, 24, _
            TS_META, False, C_FOREST, C_WHITE, C_BORDER_LT, "btn"
    BoxText z, "fltReset", Poruka("OTKUI_FLT_PONISTI"), 12, 184, 212, 26, _
            TS_META, False, C_MUTED, C_SAND_LT, C_BORDER_LT, "btn"
End Sub

Private Sub ToggleFilterPanel()
    If mFltOpen Then
        CloseFilterPanel
    Else
        OpenFilterPanel
    End If
End Sub

Private Sub OpenFilterPanel()
    Dim z As Object, b As Object, ax As Single, ay As Single
    On Error Resume Next
    ClosePopup
    Set z = mFrm.Controls("zFlt")
    Set b = mFrm.Controls("zGrid").Controls("btnFilteriB")
    AbsPos b, ax, ay
    z.Left = ax + b.width - FLT_W
    If z.Left < 0 Then z.Left = 0
    z.top = ay + b.Height + 4
    mFltOpen = True
    RenderFilterPanel
    z.Visible = True
    z.ZOrder 0
End Sub

Private Sub CloseFilterPanel()
    On Error Resume Next
    mFltOpen = False
    If mFrm Is Nothing Then Exit Sub
    mFrm.Controls("zFlt").Visible = False
End Sub

Private Sub RenderFilterPanel()
    Dim z As Object
    On Error Resume Next
    Set z = mFrm.Controls("zFlt")
    FltBox z, "fltP0", (mFilter = "danas")
    FltBox z, "fltP1", (mFilter = "nedelja")
    FltBox z, "fltP2", (mFilter = "mesec")
    FltBox z, "fltP3", (mFilter = "sve")
    FltBox z, "fltVrsta", (Len(mFltVrsta) > 0)
    FltBox z, "fltPart", (Len(mFltPart) > 0)
End Sub

' Polje partnera prikazuje "Ime Prezime - Otkupno mesto", a u koloni PARTNER
' stoji samo ime. Filter zato mora da odbaci odrednicu; sa njom u uslovu nije
' pogadjao nijedan red i lista je ispadala prazna.
Private Function PartnerBezOM(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, ChrW(183))
    If p > 0 Then s = Left$(s, p - 1)
    PartnerBezOM = Trim$(s)
End Function

Private Sub FltBox(z As Object, ByVal nm As String, ByVal on_ As Boolean)
    BoxState z, nm, IIf(on_, C_FOREST, C_WHITE), IIf(on_, C_CREAM, C_FOREST), on_
End Sub

' Klik u panelu. Vrsta i partner se citaju iz Osnovnih podataka - filter je
' uvek "ono sto je gore izabrano", pa nema treceg mesta gde se bira isto.
Private Sub FilterPanelClick(ByVal tag As String)
    On Error Resume Next
    Select Case tag
        Case "fltP0": mFilter = "danas"
        Case "fltP1": mFilter = "nedelja"
        Case "fltP2": mFilter = "mesec"
        Case "fltP3": mFilter = "sve"
        Case "fltVrsta"
            If Len(mFltVrsta) > 0 Then
                mFltVrsta = ""
            Else
                mFltVrsta = Trim$(CStr(mFrm.Controls("zCtx").Controls("cbVrsta").text))
                If Len(mFltVrsta) = 0 Then ShowToast Poruka("OTKUI_FLT_NEMA_VRSTE"), True
            End If
        Case "fltPart"
            If Len(mFltPart) > 0 Then
                mFltPart = ""
            Else
                mFltPart = PartnerBezOM(CStr(mFrm.Controls("zCtx").Controls("cbKupac").text))
                If Len(mFltPart) = 0 Then ShowToast Poruka("OTKUI_FLT_NEMA_PART"), True
            End If
        Case "fltReset"
            mFilter = "danas"
            mFltVrsta = ""
            mFltPart = ""
    End Select
    mSelRow = 0
    RenderFilterPanel
    RefreshFilterBadge
    ApplyChipVisual mFrm.Controls("zGrid")
    ReloadGrid
End Sub

' Dugme nosi broj aktivnih dodatnih uslova - inace se ne vidi da je filter
' ukljucen kad je panel zatvoren.
Private Sub RefreshFilterBadge()
    Dim n As Long
    On Error Resume Next
    If Len(mFltVrsta) > 0 Then n = n + 1
    If Len(mFltPart) > 0 Then n = n + 1
    If mFilter = "mesec" Then n = n + 1
    mFrm.Controls("zGrid").Controls("btnFilteriC").caption = _
        Poruka("OTKUI_BTN_FILTERI") & IIf(n > 0, " " & n, "")
    BoxState mFrm.Controls("zGrid"), "btnFilteri", _
             IIf(n > 0, C_FOREST, C_WHITE), IIf(n > 0, C_CREAM, C_MUTED), (n > 0)
End Sub

' Kombo po imenu. Trazi se u obe zone unosnog ekrana i u zoni AKTIVNOG
' ugovornog ekrana -- bez tog treceg mesta panel izbora ne bi mogao da se otvori
' nad poljem koje ekran sam nosi: PopKeyFor bi vratio prazno i klik na strelicu
' ne bi radio nista. Operater je to prijavio kao 'dropdowns ne rade'.
'
' Ljuska i dalje ne zna nijedan ekran po imenu -- zona se trazi po mScreen.
Private Function FindCombo(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCombo = mFrm.Controls("zCtx").Controls(nm)
    If FindCombo Is Nothing Then _
        Set FindCombo = mFrm.Controls("zForm").Controls(nm).Controls(nm & "T")
    If FindCombo Is Nothing And Len(mScreen) > 0 Then _
        Set FindCombo = mFrm.Controls("zScr_" & mScreen).Controls(nm).Controls(nm & "T")
End Function

' apsolutna pozicija kontrole na formi (sabira Left/Top svih Frame roditelja)
Private Sub AbsPos(ctl As Object, ByRef X As Single, ByRef Y As Single)
    Dim p As Object
    On Error Resume Next
    X = ctl.Left: Y = ctl.top
    Set p = ctl.parent
    Do While TypeName(p) = "Frame"
        X = X + p.Left: Y = Y + p.top
        Set p = p.parent
    Loop
End Sub

'=====================================================================
' LAYOUT
'=====================================================================
Public Sub LayoutOtkup(frm As Object)
    Dim wTot As Single, hTot As Single, navW As Single, rgtW As Single
    Dim mainX As Single, mainW As Single, ctrW As Single
    Dim cols As Long, stackRight As Boolean
    Dim formH As Single, gridY As Single, gridH As Single, i As Long, tw As Single
    On Error Resume Next
    If frm Is Nothing Then Exit Sub
    If mBuilding Then Exit Sub
    ClosePopup                          ' paneli su pozicionirani apsolutno
    CloseFilterPanel

    wTot = frm.InsideWidth: hTot = frm.InsideHeight
    If wTot < 400 Or hTot < 340 Then Exit Sub

    If wTot < BP_NARROW Then
        navW = SIDEBAR_MIN: cols = 2: stackRight = True
    Else
        navW = SIDEBAR_W: cols = 3: stackRight = False
    End If
    rgtW = IIf(stackRight, 0, RIGHT_W)
    mainX = navW
    mainW = wTot - navW
    ctrW = mainW - rgtW

    frm.Controls("zHdr").width = wTot
    With frm.Controls("zHdr")
        For i = 0 To 23
            .Controls("hgr" & i).Left = i * (wTot / 24)
            .Controls("hgr" & i).width = (wTot / 24) + 1
        Next i
        ' klaster desno, zdesna nalevo: X | Maticni | Operater | Snimi | Excel
        MoveBtn frm.Controls("zHdr"), "btnClose", wTot - 34, 10
        MoveBtn frm.Controls("zHdr"), "btnMatic", wTot - 164, 10
        MoveBtn frm.Controls("zHdr"), "btnSnimi", wTot - 304, 10
        MoveBtn frm.Controls("zHdr"), "btnExcel", wTot - 424, 10
        .Controls("hdrSaveTick").Left = wTot - 612
        .Controls("hdrSaved").Left = wTot - 596
    End With

    With frm.Controls("zNav")
        .Left = 0: .top = HEADER_H: .width = navW: .Height = hTot - HEADER_H - STATUS_H
        .Controls("navEdge").Left = navW - 1: .Controls("navEdge").Height = .Height
        .Controls("navFootLn").top = .Height - 24: .Controls("navFootLn").width = navW - 1
        .Controls("navSezona").top = .Height - 18
        .Controls("navVerzija").top = .Height - 18
        .Controls("navVerzija").Left = navW - 68
        ' profil sedi tacno iznad podnozja
        .Controls("profLn").top = .Height - PROFIL_H: .Controls("profLn").width = navW - 1
        .Controls("profBox").top = .Height - 54
        .Controls("profIco").top = CenterIco(.Height - 54, 26, TS_NAVICO)
        .Controls("profName").top = .Height - 52
        .Controls("profName").width = navW - 50
        .Controls("profSub").top = .Height - 39
        .Controls("profSub").width = navW - 50
        .Controls("profHit").top = .Height - 56: .Controls("profHit").width = navW - 1
    End With
    NavCollapse frm, (navW = SIDEBAR_MIN)

    ' Ugovorni ekran ne deli raspored sa ekranom dokumenata: dobija ceo
    ' prostor desno od sidebara i sam se rasporedi. Statusna traka ostaje
    ' zajednicka, pa se postavlja i ovde.
    If mScreen <> "DOKUMENTI" Then
        LayoutScreenZone frm, mainX, mainW, wTot, hTot
        With frm.Controls("zStatus")
            .Left = 0: .top = hTot - STATUS_H: .width = wTot: .Height = STATUS_H
            .Controls("stLnT").width = wTot
            .Controls("stHelp").Left = wTot - PAD - 70
            .Controls("stSys").Left = wTot - PAD - 250
            .Controls("stDot").Left = wTot - PAD - 260
        End With
        Exit Sub
    End If

    With frm.Controls("zKpi")
        .Left = mainX: .top = HEADER_H: .width = mainW: .Height = KPI_H
        .Controls("kpiLnB").width = mainW
        tw = mainW / 5
        For i = 0 To 4
            .Controls("kpiBg" & i).Left = i * tw: .Controls("kpiBg" & i).width = tw
            .Controls("kpiAcc" & i).Left = i * tw + PAD - 8
            .Controls("kpiC" & i).Left = i * tw + PAD: .Controls("kpiC" & i).width = tw - PAD * 2
            .Controls("kpiV" & i).Left = i * tw + PAD: .Controls("kpiV" & i).width = tw - PAD * 2
            .Controls("kpiS" & i).Left = i * tw + PAD: .Controls("kpiS" & i).width = tw - PAD * 2
            If i < 4 Then .Controls("kpiSep" & i).Left = (i + 1) * tw
        Next i
    End With

    With frm.Controls("zTitle")
        .Left = mainX: .top = HEADER_H + KPI_H: .width = mainW: .Height = TITLE_H
        .Controls("titLnB").width = mainW
        .Controls("titDatum").Left = mainW - PAD - 190
        ' pun naziv dokumenta ("Prijem ambalaze / Gotovinske uplate") ne staje
        ' u fiksnih 230pt - naslov uzima sve do datuma desno
        .Controls("titName").width = mainW - PAD * 2 - 200 - TIT_ICO_W - 10
        .Controls("titSub").width = mainW - PAD * 2 - 200 - TIT_ICO_W - 10
        PlaceTitleBadge frm
    End With

    ' Traka otpremnice stoji izmedju naslova i "Osnovnih podataka": opisuje
    ' IZVOR robe, pa mora biti iznad polja koja se iz njega pretpopunjavaju.
    Dim otpH As Single, i2 As Long
    otpH = 0
    If modeKey(ActiveMode) = "OTKUP" And Not mGridMax Then otpH = OTP_H
    frm.Controls("zOtp").Visible = (otpH > 0)
    If otpH > 0 Then
        With frm.Controls("zOtp")
            .Left = mainX: .top = HEADER_H + KPI_H + TITLE_H
            .width = ctrW: .Height = OTP_H
            .Controls("otpLnB").width = ctrW
            For i2 = 0 To 3
                .Controls("otpML" & i2).Left = ctrW - PAD - (4 - i2) * 132
                .Controls("otpMV" & i2).Left = ctrW - PAD - (4 - i2) * 132
                .Controls("otpMA" & i2).Left = ctrW - PAD - (4 - i2) * 132
            Next i2
        End With
    End If

    With frm.Controls("zCtx")
        .Left = mainX: .top = HEADER_H + KPI_H + TITLE_H + otpH: .width = ctrW: .Height = CTX_H
        .Controls("ctxLnB").width = ctrW
        .Controls("ctxHint").Left = ctrW - PAD - 52
        .Controls("ctxLn").Left = PAD + 104
        .Controls("ctxLn").width = ctrW - PAD * 2 - 104 - 52 - GAP
    End With
    LayoutCtx frm.Controls("zCtx"), ctrW

    formH = LayoutFields(frm.Controls("zForm"), cols, ctrW)
    With frm.Controls("zForm")
        .Left = mainX: .top = HEADER_H + KPI_H + TITLE_H + CTX_H
        .width = ctrW: .Height = formH
    End With
    frm.Controls("zForm").top = HEADER_H + KPI_H + TITLE_H + otpH + CTX_H

    ' Razvucena mreza: kontekst i forma se sklanjaju, mreza pocinje odmah ispod
    ' naslova dokumenta. Zone se ne unistavaju - drugi klik ih vraca netaknute.
    frm.Controls("zCtx").Visible = Not mGridMax
    frm.Controls("zForm").Visible = Not mGridMax
    If mGridMax Then
        gridY = HEADER_H + KPI_H + TITLE_H
    Else
        gridY = HEADER_H + KPI_H + TITLE_H + otpH + CTX_H + formH
    End If
    ' Donja granica zone mreze MORA da racuna i traku iznad tela (GRID_TOP:
    ' naslov, pretraga, cipovi). Ranije je bila 231pt bez nje, pa je zona bila
    ' visa od raspolozivog prostora i podnozje (Prikazano / Ukupno / strane)
    ' je ispadalo ispod donje ivice forme. Sada je minimum tacno ono sto zona
    ' zaista zahteva za tri reda.
    ' Minimum se NE sme nametnuti preko raspolozivog prostora. Ranije je zona
    ' u tesnom prozoru dobijala visinu za tri reda i podnozje (Prikazano /
    ' Ukupno / strane) je ispadalo ispod donje ivice. Traka otpremnice je taj
    ' slucaj napravila cestim: uzima 60pt koje je mreza do sada imala.
    ' Pravilo je obrnuto: podnozje se uvek vidi, a broj redova pada koliko
    ' treba - i do jednog.
    gridH = hTot - gridY - STATUS_H
    Dim gridMin As Single
    gridMin = GRID_TOP + GRID_HEAD_H + GRID_FOOT_H + GRID_ROW_H + 6
    If gridH < gridMin Then gridH = gridMin
    With frm.Controls("zGrid")
        .Left = mainX: .top = gridY: .width = ctrW: .Height = gridH
    End With
    LayoutGrid frm.Controls("zGrid"), ctrW, gridH

    With frm.Controls("zRight")
        If stackRight Then
            .Visible = False
        Else
            .Visible = True
            .Left = mainX + ctrW: .top = HEADER_H + KPI_H + TITLE_H
            .width = rgtW: .Height = hTot - .top - STATUS_H
            LayoutRight frm.Controls("zRight"), .Height
        End If
    End With

    With frm.Controls("zStatus")
        .Left = 0: .top = hTot - STATUS_H: .width = wTot: .Height = STATUS_H
        .Controls("stLnT").width = wTot
        .Controls("stHelp").Left = wTot - PAD - 70
        .Controls("stSys").Left = wTot - PAD - 250
        .Controls("stDot").Left = wTot - PAD - 260
    End With

    If mPageSize <> mLastPageSize Then
        mLastPageSize = mPageSize
        RenderGrid
    End If
End Sub

Private Function LayoutFields(z As Object, cols As Long, zw As Single) As Single
    Dim c As Object, colW As Single, col As Long, Y As Single, span As Long
    Dim g As Long, gk As Variant, imaPolja As Boolean
    Dim valFr As Object, btnY As Single, btnX As Single
    colW = (zw - 2 * PAD - (cols - 1) * GAP) / cols
    Y = 6
    gk = GrpKeys()

    ' Polja se redjaju PO GRUPAMA, fiksnim redosledom. Grupa bez ijednog
    ' vidljivog polja u datom rezimu se ne prikazuje - ni naslov ni linija.
    For g = 0 To UBound(gk)
        imaPolja = False
        col = 0
        For Each c In z.Controls
            If TypeName(c) = "Frame" Then
                If c.Visible And GrpOf(c) = CStr(gk(g)) Then
                    If Not imaPolja Then
                        imaPolja = True
                        z.Controls("grpH" & g).Visible = True
                        z.Controls("grpH" & g).top = Y
                        z.Controls("grpLn" & g).Visible = True
                        z.Controls("grpLn" & g).top = Y + 5
                        z.Controls("grpLn" & g).Left = PAD + GRP_LBL_W
                        z.Controls("grpLn" & g).width = zw - PAD * 2 - GRP_LBL_W
                        Y = Y + GRP_H
                    End If
                    span = SpanOf(c)
                    If span > cols Then span = cols
                    If col + span > cols Then
                        col = 0: Y = Y + FIELD_GRP_H + ROW_GAP
                    End If
                    c.Left = PAD + col * (colW + GAP)
                    c.width = colW * span + (span - 1) * GAP
                    c.top = Y
                    c.Height = FIELD_GRP_H
                    LayoutFieldInner c
                    col = col + span
                    If col >= cols Then col = 0: Y = Y + FIELD_GRP_H + ROW_GAP
                End If
            End If
        Next c
        If imaPolja Then
            If col > 0 Then Y = Y + FIELD_GRP_H + ROW_GAP
            ' vazduh izmedju grupa vise ne ide ovde nego u GRP_H, ispod natpisa
        Else
            z.Controls("grpH" & g).Visible = False
            z.Controls("grpLn" & g).Visible = False
        End If
    Next g

    ' AKCIONI RED. Redosled sleva: poruka | racunica + IZNOS | dugmad.
    ' Iznos i potvrda su jedna radnja pa dele red - ploca ne trosi ceo red.
    ' Redosled dugmadi: Sacuvaj (najcesce) najblize poljima, Otkazi najdalje.
    btnY = Y + 4
    btnX = zw - PAD - (196 + GAP + 132 + GAP + 72)
    MoveBtn z, "btnSacuvaj", btnX, btnY
    MoveBtn z, "btnSacuvajPrint", btnX + 196 + GAP, btnY
    MoveBtn z, "btnOtkazi", btnX + 196 + GAP + 132 + GAP, btnY

    ' Ploca ide uz LEVU ivicu reda (kao u viziji), poruka popunjava sredinu,
    ' dugmad ostaju desno.
    Set valFr = VisibleValFrame(z)
    If Not valFr Is Nothing Then
        valFr.width = VAL_PLATE_W + IIf(HasCtl(valFr, valFr.name & "K1"), _
                                        GAP + 4 + VAL_CALC_W + GAP + VAL_KG_W, 0)
        ' okvir ploce ne sme da zadje pod dugmad - beo je i pokrio bi ih
        If PAD + valFr.width > btnX - GAP Then valFr.width = btnX - GAP - PAD
        valFr.Left = PAD
        valFr.top = btnY                 ' ploca (28pt) i dugme (28pt) u istoj liniji
        valFr.Height = FIELD_H
        LayoutFieldInner valFr
        ' Kilogrami vise ne dele prostor sa porukom: toast je preseljen u
        ' naslovnu traku, pa ostaju stalno vidljivi.
    End If
    LayoutFields = btnY + 28 + 8
End Function

' Redosled grupa polja. Isti niz koristi i BuildForm za naslove.
Private Function GrpKeys() As Variant
    GrpKeys = Array("DOK", "ROBA", "AMB", "NOV")
End Function

' Tag polja je "fld:<span>:<grupa>". Grupa je PRAZNA kod ploce iznosa - ona se
' ne redja sa poljima nego ide u red sa dugmadima.
Private Function GrpOf(c As Object) As String
    Dim p As Variant
    If Left$(c.tag, 4) <> "fld:" Then Exit Function
    p = Split(c.tag, ":")
    If UBound(p) < 2 Then Exit Function
    GrpOf = CStr(p(2))
End Function

Private Function SpanOf(c As Object) As Long
    Dim p As Variant
    p = Split(c.tag, ":")
    If UBound(p) < 1 Then
        SpanOf = 1
    Else
        SpanOf = CLng(val(p(1)))
    End If
    If SpanOf < 1 Then SpanOf = 1
End Function

' U svakom rezimu je vidljiva najvise jedna ploca iznosa (VREDNOST ili
' NEISPLACENO), pa akcioni red ima tacno jedno mesto za nju.
Private Function VisibleValFrame(z As Object) As Object
    On Error Resume Next
    If z.Controls("fgVrednost").Visible Then
        Set VisibleValFrame = z.Controls("fgVrednost")
    ElseIf z.Controls("fgOstatak").Visible Then
        Set VisibleValFrame = z.Controls("fgOstatak")
    End If
End Function

' Segmenti (I/II klasa, 4 smera reversa) se rastezu preko cele sirine polja.
' Bez ovoga su ostajali na sirini iz trenutka kreiranja, pa je desno od njih
' zjapio prazan prostor.
Private Sub LayoutSegs(fr As Object)
    Dim c As Object, nms As Collection, i As Long, w As Single, X As Single
    On Error Resume Next
    Set nms = New Collection
    For Each c In fr.Controls
        If Left$(c.name, 3) = "seg" Then
            If IsNumeric(Right$(c.name, 1)) Then nms.Add c.name
        End If
    Next c
    If nms.count = 0 Then Exit Sub
    w = (fr.width - 2) / nms.count
    X = 1
    For i = 1 To nms.count
        MoveBox fr, CStr(nms(i)), X, 17, w - 2
        X = X + w
    Next i
End Sub

' Unutrasnjost polja prema sirini okvira: labela, ivica, jedinica, strelica i
' sam unos. JAVNO iz istog razloga kao NewFieldG -- ekran koji polje NAPRAVI
' mora i da ga RASPOREDI, inace mu unutrasnje kontrole ostanu na merama iz
' gradnje (180pt). Operater je to video kao jedinicu 'kg' nasred polja i unos
' odsecen sa leve strane.
Public Sub LayoutFieldInner(fr As Object)
    Dim nm As String, uW As Single, inW As Single, hw As Single
    Dim plateW As Single, px As Single
    nm = fr.name
    On Error Resume Next
    If HasCtl(fr, nm & "L") Then fr.Controls(nm & "L").width = fr.width
    If HasCtl(fr, nm & "L2") Then fr.Controls(nm & "L2").width = fr.width
    If HasCtl(fr, nm & "B") Then fr.Controls(nm & "B").width = fr.width
    If HasCtl(fr, nm & "F") Then fr.Controls(nm & "F").width = fr.width - 2
    ' KLASA I CENA ima svoje segmente koji NE smeju preko cele sirine polja -
    ' desno od njih stoje dve kutije cene
    If Not HasCtl(fr, "fgCena1B") Then LayoutSegs fr

    uW = 0
    If HasCtl(fr, nm & "UB") Then
        uW = fr.Controls(nm & "UB").width
        fr.Controls(nm & "UB").Left = fr.width - uW - 2
        fr.Controls(nm & "U").Left = fr.width - uW - 2
    End If
    If HasCtl(fr, nm & "D") Then
        fr.Controls(nm & "D").Left = fr.width - 19
        uW = 18
    End If
    ' dugme "prepisi" sedi levo od jedinice i uzima svoj deo sirine polja
    If HasCtl(fr, nm & "EQB") Then
        fr.Controls(nm & "EQB").Left = fr.width - uW - 2 - 26
        fr.Controls(nm & "EQ").Left = fr.width - uW - 2 - 26
        uW = uW + 26
    End If
    inW = fr.width - INPUT_PAD - uW - 3
    If inW < 20 Then inW = 20
    If HasCtl(fr, nm & "T") Then fr.Controls(nm & "T").width = inW

    ' KLASA I CENA: prekidac klase levo, pa dve jednake kutije cene
    If HasCtl(fr, "fgCena1B") Then
        MoveBox fr, "segKlasa1", 0, 17, SEG_KL_W
        MoveBox fr, "segKlasa2", SEG_KL_W + 1, 17, SEG_KL_W
        px = 2 * SEG_KL_W + 1 + GAP
        hw = (fr.width - px - GAP) / 2
        If hw < 44 Then hw = 44
        MoveShell fr, "fgCena1", px, 16, hw
        fr.Controls("fgCena1P").Left = px + 5
        fr.Controls("fgCena1T").Left = px + 20
        fr.Controls("fgCena1T").width = hw - 24
        MoveShell fr, "fgCena2", px + hw + GAP, 16, hw
        fr.Controls("fgCena2P").Left = px + hw + GAP + 5
        fr.Controls("fgCena2T").Left = px + hw + GAP + 20
        fr.Controls("fgCena2T").width = hw - 24
    End If
    ' PLOCA IZNOSA (VREDNOST / NEISPLACENO) - kao u viziji: ploca je uz LEVU
    ' ivicu reda, sadrzaj joj je u jednom redu (natpis, broj, RSD), zlatna
    ' ivica je levo, a racunica po klasama stoji DESNO od ploce.
    If HasCtl(fr, nm & "V") Then
        plateW = VAL_PLATE_W
        If plateW > fr.width Then plateW = fr.width
        MoveShell fr, nm, 0, 0, plateW
        If HasCtl(fr, nm & "A") Then fr.Controls(nm & "A").Left = 1
        fr.Controls(nm & "L").Left = INPUT_PAD + 3
        fr.Controls(nm & "L").width = 62
        fr.Controls(nm & "V").Left = INPUT_PAD + 3 + 62
        fr.Controls(nm & "V").width = plateW - (INPUT_PAD + 3 + 62) - 36
        fr.Controls(nm & "U").Left = plateW - 32
        fr.Controls(nm & "U").width = 26
        If HasCtl(fr, nm & "K1") Then
            fr.Controls(nm & "K1").Left = plateW + GAP + 4
            fr.Controls(nm & "K2").Left = plateW + GAP + 4
            fr.Controls(nm & "K1").width = VAL_CALC_W
            fr.Controls(nm & "K2").width = VAL_CALC_W
            fr.Controls(nm & "K1").Visible = True
            fr.Controls(nm & "K2").Visible = True
        End If
        If HasCtl(fr, nm & "Kg") Then
            fr.Controls(nm & "Kg").Left = plateW + GAP + 4 + VAL_CALC_W + GAP
            fr.Controls(nm & "Kg").width = fr.width - _
                (plateW + GAP + 4 + VAL_CALC_W + GAP)
        End If
    End If
End Sub

' Prvi red ("Osnovni podaci") koristi PUNU sirinu zone. Polja su do sada
' ostajala na fiksnim x koordinatama iz izgradnje, pa je desno od poslednjeg
' combo-a zjapio prazan pojas koliko god prozor bio sirok.
' Naslovi dokumenata nisu iste duzine, pa se znacka postavlja po IZMERENOJ
' sirini teksta - fiksni odmak bi kod "Prijem ambalaze / Gotovinske uplate"
' zavrsio usred naslova, a kod "Storno" daleko od njega.
' Drugi red plocice. Prazan string sakriva red - bez teksta ne ostaje trag.
Private Sub KpiSub(z As Object, ByVal i As Long, ByVal t As String)
    On Error Resume Next
    z.Controls("kpiS" & i).caption = t
End Sub

' Vrednost koja NE POSTOJI (nema izvor ili nema konteksta): em-crta u sivom.
' Razlicito od nule, koja je stvaran podatak.
Private Sub KpiPrazno(z As Object, ByVal i As Long)
    On Error Resume Next
    z.Controls("kpiV" & i).caption = ChrW(8212)
    z.Controls("kpiV" & i).ForeColor = RGB(168, 175, 164)
    z.Controls("kpiS" & i).caption = ""
End Sub

Private Sub PlaceTitleBadge(frm As Object)
    Dim z As Object, w As Single
    On Error Resume Next
    Set z = frm.Controls("zTitle")
    z.Controls("titRul").caption = z.Controls("titName").caption
    w = z.Controls("titRul").width
    z.Controls("titBadgeB").Left = z.Controls("titName").Left + w + 8
    z.Controls("titBadgeC").Left = z.Controls("titName").Left + w + 8
End Sub

Private Sub LayoutCtx(z As Object, ByVal zw As Single)
    Dim nm As Variant, wt As Variant, i As Long
    Dim X As Single, cell As Single, cw As Single, avail As Single, sumW As Single
    On Error Resume Next
    nm = Array("cbVrsta", "cbSorta", "cbOM", "cbVozac", "cbKupac")
    ' Polja NISU jednaka: partner nosi najduzi tekst ("Prezime Ime - Otkupno
    ' mesto") pa dobija vecu tezinu, a vrsta i sorta su kratke. Sa pet jednakih
    ' celija je partner ispadao odsecen. Natpis je iznad polja, pa cela celija
    ' pripada polju - red je time i siri nego dok je natpis stajao sa strane.
    wt = Array(1#, 1.15, 1.15, 1#, 1.7)
    sumW = 6#
    X = PAD
    avail = zw - 2 * PAD - 4 * GAP
    If avail < 560 Then avail = 560     ' ispod toga combo ostaje bez sirine
    For i = 0 To 4
        cell = avail * CSng(wt(i)) / sumW
        cw = cell
        z.Controls("ctxL" & i).Left = X
        z.Controls("ctxL" & i).width = cell
        z.Controls("ctxL" & i).top = CTX_LBL_Y
        MoveCtxCombo z, CStr(nm(i)), X, cw
        X = X + cell + GAP
    Next i
    z.Controls("ctxPal").Left = zw - PAD - 52 - GAP - 420
    z.Controls("ctxPal").width = 420
End Sub

' shell + combo + strelica se pomeraju zajedno (parnjak NewCtxCombo-a)
Private Sub MoveCtxCombo(z As Object, ByVal nm As String, ByVal X As Single, ByVal w As Single)
    On Error Resume Next
    MoveShell z, nm, X, CTX_FLD_Y, w
    z.Controls(nm).Left = X + INPUT_PAD
    z.Controls(nm).top = CTX_FLD_Y + 4
    z.Controls(nm).width = w - INPUT_PAD - 22
    z.Controls(nm & "D").Left = X + w - 19
    z.Controls(nm & "D").top = CenterY(CTX_FLD_Y + 1, 24, TS_META)
End Sub

Private Sub LayoutGrid(z As Object, zw As Single, zh As Single)
    Dim i As Long, r As Long, c As Long, X As Single, base As Variant
    Dim sc As Single, hd As Object, body As Object, ft As Object, bodyH As Single
    z.Controls("grdLnT").width = zw
    ' Prekidac lista: natpisi, sirine i broj dugmadi dolaze OD EKRANA.
    Dim lsX As Single, seg As Variant, sp As Variant
    lsX = PAD + 190
    seg = SegDefs()
    For i = 0 To MAX_SEG - 1
        If IsArray(seg) Then
            If i <= UBound(seg) Then
                sp = Split(CStr(seg(i)), "|")
                z.Controls("lsSeg" & i & "C").caption = Poruka(CStr(sp(1)))
                BoxShow z, "lsSeg" & i, True
                MoveBox z, "lsSeg" & i, lsX, 9, CSng(val(sp(3)))
                lsX = lsX + CSng(val(sp(3)))
                GoTo SledeciSeg
            End If
        End If
        BoxShow z, "lsSeg" & i, False
SledeciSeg:
    Next i
    z.Controls("grdSrc").Left = PAD + 176
    ' Radnje nad redom: koje su i koliko ih ima kaze EKRAN, za svoju aktivnu
    ' listu. Redjaju se zdesna nalevo, redom kojim su prijavljene.
    Dim act As Variant, ap As Variant, ax As Single
    act = ActDefs()
    ax = zw - PAD
    For i = 0 To MAX_ACT - 1
        If IsArray(act) Then
            If i <= UBound(act) Then
                ap = Split(CStr(act(i)), ":")
                z.Controls("btnAct" & i & "C").caption = Poruka(CStr(ap(1)))
                BoxShow z, "btnAct" & i, True
                ax = ax - CSng(val(ap(2)))
                MoveBtn z, "btnAct" & i, ax, 36
                ax = ax - GAP
                GoTo SledecaAkcija
            End If
        End If
        BoxShow z, "btnAct" & i, False
SledecaAkcija:
    Next i

    ' Opseg datuma je jedini deo reda radnji koji nije dugme. Pripada listi
    ' otpremnica; kad levo od dugmadi nema 296pt, ceo trio se sklanja - isto
    ' pravilo po kome se sklanjaju kolone nizeg prioriteta.
    Dim dV As Boolean
    dV = SpecDatLista() And (ax - 112 - 190 > PAD + 190)
    ShowSpecDat z, dV
    If dV Then
        Dim dx As Single
        dx = ax - 112
        MoveBtn z, "btnRedSpecDat", dx, 36
        MoveShell z, "specDo", dx - 6 - 74, 36, 74
        z.Controls("specDoT").Left = dx - 6 - 74 + 6
        z.Controls("specDoL").Left = dx - 6 - 74 - 18
        MoveShell z, "specOd", dx - 6 - 74 - 18 - 6 - 74, 36, 74
        z.Controls("specOdT").Left = dx - 6 - 74 - 18 - 6 - 74 + 6
        z.Controls("specOdL").Left = dx - 6 - 74 - 18 - 6 - 74 - 18
    End If
    MoveBtn z, "btnMax", zw - PAD - 26, 8
    MoveBtn z, "btnFilteri", zw - PAD - 26 - GAP - 88, 8
    Dim sx As Single: sx = zw - PAD - 88 - GAP - 200
    MoveShell z, "srch", sx, 8, 200
    z.Controls("srchIco").Left = sx + 5
    z.Controls("txtSearch").Left = sx + 24
    z.Controls("txtSearch").width = 200 - 24 - INPUT_PAD
    z.Controls("lblSearchPh").Left = z.Controls("txtSearch").Left
    z.Controls("lblSearchPh").Visible = (Len(mSearch) = 0)

    sc = zw - 2 * PAD

    ' Sirine iz opisa kolona. Kolone prioriteta 2 (sorta, klasa, kol. ambalaze,
    ' tip ambalaze) se sklanjaju dok red ne stane - uzak ekran tako ne lomi
    ' mrezu nego samo prikazuje manje detalja.
    If Not IsArray(mCols) Then SetGridCols modeKey(ActiveMode)

    Dim vis(0 To MAX_COLS - 1) As Boolean
    Dim fixW As Single, flexI As Long, dropped As Boolean
    flexI = -1
    Dim pass As Long
    For pass = 3 To 1 Step -1
        fixW = 0: flexI = -1
        For i = 0 To mColN - 1
            vis(i) = (CLng(val(ColF(CStr(mCols(i)), 4))) <= pass)
            If vis(i) Then
                If CSng(val(ColF(CStr(mCols(i)), 3))) = 0 Then
                    flexI = i
                Else
                    fixW = fixW + CSng(val(ColF(CStr(mCols(i)), 3)))
                End If
            End If
        Next i
        If sc - fixW >= 120 Then Exit For        ' partner kolona ima gde da stane
    Next pass

    X = 0
    For i = 0 To MAX_COLS - 1
        mColX(i) = X
        If i < mColN Then
            If Not vis(i) Then
                mColW(i) = 0
            ElseIf i = flexI Then
                mColW(i) = sc - fixW
                If mColW(i) < 120 Then mColW(i) = 120
            Else
                mColW(i) = CSng(val(ColF(CStr(mCols(i)), 3)))
            End If
        Else
            mColW(i) = 0
        End If
        X = X + mColW(i)
    Next i

    Set hd = z.Controls("grdHead")
    hd.Left = PAD: hd.top = GRID_TOP: hd.width = sc: hd.Height = GRID_HEAD_H
    For i = 0 To MAX_COLS - 1
        With hd.Controls("hd" & i)
            .Visible = (mColW(i) > 0)
            .Left = mColX(i) + 8
            .width = IIf(mColW(i) > 16, mColW(i) - 16, 1)
            .TextAlign = IIf(ColIsNum(i), fmTextAlignRight, fmTextAlignLeft)
        End With
    Next i

    Set body = z.Controls("grdBody")
    body.Left = PAD: body.top = GRID_TOP + GRID_HEAD_H: body.width = sc
    ' Telo NE SME da naraste preko onoga sto je zoni ostalo - inace pokrije
    ' podnozje. Zato nema donje granice vece od raspolozivog: kad je usko,
    ' smanjuje se broj redova, ne visina zone.
    '
    ' TOAST_H se REZERVISE. Traka poruka stoji tacno iznad podnozja
    ' (PostaviToast: footTop - TOAST_H - 4), a telo je do sada racunato samo sa
    ' rezervom od 6pt -- pa je poslednji red ulazio 24pt U traku. Poruka se
    ' zato crtala PREKO reda i drzala se samo ZOrder-om (ShowToast), sto je
    ' resenje za redosled crtanja, a ne za prostor. Sada telo staje TU gde
    ' traka pocinje: body.Bottom <= toast.Top.
    bodyH = zh - (GRID_TOP + GRID_HEAD_H) - GRID_FOOT_H - TOAST_H - 4
    If bodyH < GRID_ROW_H * 3 Then bodyH = GRID_ROW_H * 3

    mPageSize = Int(bodyH / GRID_ROW_H)
    If mPageSize > MAX_ROWS Then mPageSize = MAX_ROWS
    If mPageSize < 3 Then mPageSize = 3

    ' telo se skracuje na CEO broj redova: ostatak je bio prazan pojas ispod
    ' poslednjeg reda, a uz to je razbijao korak od 3pt
    bodyH = mPageSize * GRID_ROW_H
    body.Height = bodyH

    For r = 0 To MAX_ROWS - 1
        body.Controls("rb" & r).width = sc
        body.Controls("rl" & r).width = sc
        For c = 0 To MAX_COLS - 1
            With body.Controls("c" & r & "_" & c)
                .Left = mColX(c) + 8
                ' status pilula sama racuna sirinu po tekstu (PaintPill)
                If ColKind(c) <> "pill" Then _
                    .width = IIf(mColW(c) > 16, mColW(c) - 16, 1)
                .TextAlign = IIf(ColIsNum(c), fmTextAlignRight, fmTextAlignLeft)
            End With
            StyleGridCell body.Controls("c" & r & "_" & c), ColKind(c), (c = 0), r * GRID_ROW_H
        Next c
    Next r

    z.Controls("grdEmpty").Left = PAD
    z.Controls("grdEmpty").width = sc
    z.Controls("grdEmpty").top = body.top + bodyH / 2 - 16
    z.Controls("grdEmptyA").Left = PAD
    z.Controls("grdEmptyA").width = sc
    z.Controls("grdEmptyA").top = body.top + bodyH / 2 + 4

    Set ft = z.Controls("grdFoot")
    ft.Left = PAD: ft.top = zh - GRID_FOOT_H: ft.width = sc: ft.Height = GRID_FOOT_H
    PostaviToast z, sc, ft.top
    ft.Controls("ftLnT").width = sc
    ft.Controls("ftKg").Left = sc / 2 - 190
    ' Desna ivica novcanog dela je fiksna; drugi slot se dodaje ULEVO, da se
    ' jedan slot ne pomera kad se pojavi drugi.
    ft.Controls("ftVal").Left = sc / 2 - 10
    ft.Controls("ftVal2").Left = sc / 2 - 10
    X = sc - 21
    MoveBtn ft, "pgNext", X, 2
    For i = 5 To 1 Step -1
        X = X - 23
        MoveBtn ft, "pg" & i, X, 2
    Next i
    MoveBtn ft, "pgPrev", X - 23, 2

    ' Naslovi kolona zavise od mColW (skrivena kolona nema naslov), a mColW
    ' postaje tacan tek ovde. Pri prvom otvaranju je RefreshSortGlyphs isao
    ' PRE rasporeda, pa je zaglavlje ostajalo prazno dok korisnik ne klikne.
    RefreshSortGlyphs

    ' Geometrija je upravo preracunata nad tekucim mCols.
    mGeomStara = False
End Sub

' Tipografija celije po vrsti kolone. Pravilo: sve sto je BROJ, SIFRA ili DATUM
' ide u Consolas i desno je poravnato - inace se cifre ne poklapaju po redovima
' i kolona kg izgleda razbacano. Prva kolona je uvek broj dokumenta (sifra), pa
' i ona ide u mono, podebljano. Consolas nema Semibold rez, zato pravi Bold.
Private Sub StyleGridCell(lbl As Object, ByVal kind As String, ByVal isBroj As Boolean, _
                          ByVal rowTop As Single)
    Dim fn As String, fs As Single, bd As Boolean, fc As Long
    fn = F_UI: fs = TS_CELL: bd = False: fc = C_MUTED
    If isBroj Then
        fn = F_NUM: bd = True: fc = C_FOREST
    Else
        Select Case kind
            Case "date":              fn = F_NUM: fs = TS_CELL_SM
            Case "part":              fs = TS_CELL_LG: fc = C_FOREST
            Case "kg", "num", "sum0": fn = F_NUM: fc = C_FOREST
            Case "rest":              fn = F_NUM
            Case "rsd", "mult":       fn = F_NUM: bd = True: fc = C_FOREST
            Case "pill":              fs = TS_MICRO: bd = True
            Case "paypill":           fs = TS_CELL_SM: bd = True
        End Select
    End If
    With lbl
        .Font.name = fn
        .Font.Size = fs
        .Font.bold = bd
        .ForeColor = fc
        .Height = TxtH(fs)
        .top = CenterY(rowTop, GRID_ROW_H, fs)
    End With
End Sub

Private Sub LayoutRight(z As Object, zh As Single)
    Dim i As Long, Y As Single
    z.Controls("rgtEdge").Height = zh
    For i = 0 To 6
        z.Controls("mcKey" & i).Left = RIGHT_W - 13 - 30
        z.Controls("mcHint" & i).width = RIGHT_W - 42
    Next i
    Y = 30 + 7 * 40 + 12
    z.Controls("rgtSep1").top = Y: Y = Y + 9
    z.Controls("rgtHdr2").top = Y: Y = Y + 17
    z.Controls("luEmpty").top = Y
    z.Controls("luEmpty").width = RIGHT_W - 26
    For i = 0 To 2
        z.Controls("luK" & i).top = Y
        z.Controls("luV" & i).top = Y
        z.Controls("luV" & i).Left = RIGHT_W - 13 - 90
        z.Controls("luLn" & i).top = Y + 16
        Y = Y + 18
    Next i
    ' sekcija sinhronizacije: linija, naslov, stanje, pa dugme uz samo dno
    z.Controls("rgtSep2").top = zh - 78
    z.Controls("rgtHdr3").top = zh - 70
    z.Controls("syncDot").top = zh - 51
    z.Controls("syncTxt").top = zh - 52
    MoveBtn z, "btnSync", 13, zh - 30
End Sub

' Suzen sidebar (uzak prozor): ostaju samo glifovi. Sirinu, natpise i glifove
' stavki postavlja LayoutNav -- ovde se pamti SAMO stanje, jer vidljivost stavke
' zavisi i od sekcije, a NavCollapse o sekcijama ne zna nista. Ranije je ova
' rutina palila natpise svim stavkama redom, pa bi suzavanje prozora prikazalo
' i stavke sekcije koja se ne crta.
Private Sub NavCollapse(frm As Object, collapsed As Boolean)
    Dim z As Object
    On Error Resume Next
    Set z = frm.Controls("zNav")
    If z Is Nothing Then Exit Sub
    mNavUzak = collapsed
    z.Controls("navSezona").Visible = Not collapsed
    z.Controls("navVerzija").Visible = Not collapsed
    ' u suzenom sidebaru ostaje samo kvadrat sa glifom - ime nema gde da stane
    z.Controls("profName").Visible = Not collapsed
    z.Controls("profSub").Visible = Not collapsed
    LayoutNav frm
End Sub

'=====================================================================
' RENDER MREZE - puni vec napravljene Label-ove, ne pravi kontrole
'=====================================================================
Public Sub RenderGrid()
    Dim z As Object, body As Object, ft As Object
    Dim i As Long, r As Long, k As Long, from_ As Long, to_ As Long, shown As Long
    Dim kanal As Boolean
    Dim n As Long, celijaOK As Boolean, txt As String
    On Error Resume Next
    ' Brojac vazi za OVO crtanje. Prijavljuje se jednom, na kraju.
    mKvarCelija = 0
    mKvarKind = ""
    If mFrm Is Nothing Then Exit Sub
    kanal = Not ModeHasKgCol()                ' rezim bez kilograma nema ni zbir kg
    Set z = mFrm.Controls("zGrid")
    ' Nikad ne crtaj sa sirinama koje ne odgovaraju opisu kolona -- v.
    ' SetGridColsArr. Preracunava se SAMO kad se opis stvarno promenio, pa
    ' promena cipa ili strane ne placa raspored.
    OsveziGeometriju z
    Set body = z.Controls("grdBody")
    Set ft = z.Controls("grdFoot")

    If mPageSize < 1 Then mPageSize = 10
    If mPage < 1 Then mPage = 1
    If (mPage - 1) * mPageSize >= mViewN Then mPage = 1

    z.Controls("grdEmpty").Visible = (mViewN = 0)
    ' ponuda "Prikazi sve" nema smisla kad je filter vec "sve"
    z.Controls("grdEmptyA").Visible = (mViewN = 0 And mFilter <> "sve")
    shown = 0
    If mViewN > 0 Then
        from_ = (mPage - 1) * mPageSize + 1
        to_ = from_ + mPageSize - 1
        If to_ > mViewN Then to_ = mViewN
        shown = to_ - from_ + 1
    End If

    For i = 0 To MAX_ROWS - 1
        If i < shown Then
            r = from_ + i
            body.Controls("rb" & i).Visible = True
            body.Controls("rl" & i).Visible = True
            For k = 0 To MAX_COLS - 1
                With body.Controls("c" & i & "_" & k)
                    If mColW(k) = 0 Then
                        .Visible = False
                    Else
                        .Visible = True
                        ' UPIS SE RADI UVEK -- v. CelijaTekst. Racun je izvucen u
                        ' funkciju koja ne moze da pukne, pa preskocen upis (i sa
                        ' njim natpis prethodnog ekrana) vise nije moguc.
                        Select Case ColKind(k)
                            Case "pill"
                                n = CelijaBroj(mView(r, k + 1), celijaOK)
                                If celijaOK Then
                                    PaintPill body.Controls("c" & i & "_" & k), n
                                Else
                                    OcistiPilulu body.Controls("c" & i & "_" & k), mColW(k)
                                End If
                            Case "paypill"
                                n = CelijaBroj(mView(r, k + 1), celijaOK)
                                If celijaOK Then
                                    PaintPayPill body.Controls("c" & i & "_" & k), n
                                Else
                                    ' SAMO NATPIS. Ovo NIJE ista celija kao "pill" --
                                    ' v. OcistiPilulu: pozadinu joj vraca PaintRow, a
                                    ' sirinu drzi LayoutGrid. Ko je ovde ocisti kao
                                    ' pravu pilulu, ostavi je 16pt siru zauvek.
                                    .caption = vbNullString
                                End If
                            Case Else
                                txt = CelijaTekst(ColKind(k), mView(r, k + 1), celijaOK)
                                .caption = txt
                        End Select
                        If Not celijaOK Then
                            mKvarCelija = mKvarCelija + 1
                            If Len(mKvarKind) = 0 Then mKvarKind = ColKind(k) & "/" & k & _
                                " tip=" & TypeName(mView(r, k + 1)) & _
                                " vred=[" & CStr(mView(r, k + 1)) & "]"
                        End If
                    End If
                End With
            Next k
            PaintRow body, i, (from_ + i = mSelRow), (i = mHoverRow), RowMarked(from_ + i)
        Else
            body.Controls("rb" & i).Visible = False
            body.Controls("rl" & i).Visible = False
            For k = 0 To MAX_COLS - 1
                body.Controls("c" & i & "_" & k).Visible = False
            Next k
        End If
    Next i

    ' JEDNA prijava po crtanju. Prazna celija je istina, ali TIHA prazna celija
    ' je bila pola problema: prvi nalaz ove vrste (datum) trazio je zasebnu
    ' dijagnostiku da bi se uopste video. Od sada ostaje trag u logu.
    If mKvarCelija > 0 Then
        LogWarn "modOtkupUI.RenderGrid", _
                "Celija bez prikaza: " & mKvarCelija & _
                " (ekran=" & mScreen & ", prva=" & mKvarKind & ")"
    End If

    If mViewN > 0 Then
        ft.Controls("ftRange").caption = Poruka("OTKUI_FT_PRIKAZANO") & " " & from_ & ChrW(8211) & to_ & _
                                         " " & Poruka("OTKUI_FT_OD") & " " & mViewN
    Else
        ft.Controls("ftRange").caption = Poruka("OTKUI_FT_PRIKAZANO") & " 0"
    End If
    ' gotovinski promet i reversi nemaju kilograme - zbir kg bi bio prazna nula
    ft.Controls("ftKg").Visible = Not kanal
    ft.Controls("ftKg").caption = Poruka("OTKUI_FT_UKUPNO") & " " & FmtKg(mSumKg) & " " & Poruka("OTKUI_UNIT_KG")
    ' rezim bez kolone vrednosti (Zbirna) nema ni zbir
    ft.Controls("ftVal").Visible = ModeHasValCol()
    ft.Controls("ftVal2").Visible = False
    If mFtSlotN > 0 Then
        ' Ekran je poslao svoje novcane slotove -- zbir vrednosti tada nema sta
        ' da kaze preko njih.
        CrtajSlotovePodnozja ft
    Else
        ' Reversi se broje u komadima, sve ostalo u dinarima (i bez decimala).
        '
        ' Pitanje ide EKRANU, a ne globalnom ActiveMode: taj rezim pripada
        ' Dokumentima i na ugovornom ekranu ostaje onakav kakav ga je Dokumenta
        ' ostavila. Ko je bio na F7 pa presao na Uvoz izvoda video je "Ukupno
        ' 8.950 kom" umesto "Vrednost 8.950,00 RSD" -- novac izbrojan kao komadi.
        ft.Controls("ftVal").caption = PodnozjeValTekst(mSumVal)
    End If
    RenderPager ft
    RenderChipCounts z
    RefreshRowActions
End Sub

' status-pilula u boji - ono sto MSForms ListBox nije mogao.
' Sirina se racuna PO TEKSTU: preko cele kolone tinta (#EDF0F5 i sl.) je na
' belom redu prakticno nevidljiva, pa se pilula cita kao oblik tek kad
' pripije uz tekst. Zato LayoutGrid namerno NE postavlja sirinu kolone 5.
' TEKST CELIJE, i pravilo sta se desava kad vrednost NIJE ono sto kolona tvrdi.
'
' RenderGrid radi pod "On Error Resume Next", i to je namerno: pad jedne celije
' ne sme da obori crtanje cele mreze. Ali dok se tekst racunao U SAMOM UPISU
' (".caption = FmtBroj(CDbl(v), 0)"), pad konverzije nije preskakao samo racun
' nego i UPIS -- pa je u celiji OSTAJAO natpis od ranijeg crtanja. Vrednost sa
' PRETHODNOG EKRANA, bez ijedne poruke i bez ijednog traga.
'
' Tako je "26062026" (ddmmyyyy kao broj) u koloni datuma dalo "OSIROCENE_PAL" --
' vrstu reda sa ekrana Oporavak. Datum je zatvoren posebno (DatumSerijskiValidan),
' ali to je bio JEDAN ULAZ, ne klasa: isti ishod daju CDbl nad tekstom u
' kg/num/rsd/mult i CLng u pill/paypill kolonama.
'
' Zato konverzija ide OVDE, u funkciju koja NE MOZE da pukne, a upis se radi
' UVEK. Neuspeh je PRAZNA celija -- prazno je istina, tudji tekst nije.
Public Function CelijaTekst(ByVal kind As String, ByVal v As Variant, _
                            ByRef ok As Boolean) As String
    Dim d As Double

    ok = True
    On Error GoTo EH

    Select Case kind
        Case "date"
            CelijaTekst = FmtDatumKratko(v)
            ' Datum koji se NE MOZE prikazati je kvar prikaza, ne prazna celija.
            ' FmtDatumKratko sam odbija vrednosti van opsega (26062026 kao
            ' ddmmyyyy), pa bi bez ovoga bas PRVI nalaz ove vrste ostao
            ' neprebrojan i nevidljiv u logu. Prazna celija koja je prazna zato
            ' sto podatka nema i dalje je uredna.
            If Len(CelijaTekst) = 0 Then
                If Trim$(CStr(v)) <> "" And Trim$(CStr(v)) <> "0" Then ok = False
            End If
        Case "kg":           CelijaTekst = FmtKg(CDbl(v))
        Case "num", "sum0":  CelijaTekst = FmtBroj(CDbl(v), 0)
        ' DEC je broj sa decimalama koji NIJE zbirna velicina: tezina gajbice,
        ' cena po jedinici. "num" ih je zaokruzivao na ceo broj -- 0,5 kg se
        ' crtalo kao 1 -- a "rsd" bi ih upisao u podnozje kao zbir, sto
        ' sifarnik nema sta da sabira (v. UI_MIGRACIJA_KATALOG 26.4).
        Case "dec":          CelijaTekst = FmtBroj(CDbl(v), 2)
        Case "rsd", "mult":  CelijaTekst = FmtBroj(CDbl(v), 2)
        Case "rest"
            ' nula znaci "nema duga" - prazna celija je citljivija od 0,00
            d = CDbl(v)
            If d <> 0 Then CelijaTekst = FmtBroj(d, 2)
        Case Else:           CelijaTekst = CStr(v)
    End Select
    Exit Function
EH:
    ok = False
    CelijaTekst = vbNullString
End Function

' Isto pravilo za kolone koje se crtaju kao pilula. Neuspeh NE sme da postane
' nula: nula je kod PaintPill "Sacuvana", a kod PaintPayPill "Neplaceno" --
' dakle odredjen status, sto je svoja vrsta lazi. Zato se vraca i zastavica, a
' pozivalac na neuspeh celiju PRAZNI umesto da naslika bilo koji status.
Public Function CelijaBroj(ByVal v As Variant, ByRef ok As Boolean) As Long
    ok = True
    On Error GoTo EH
    CelijaBroj = CLng(v)
    Exit Function
EH:
    ok = False
    CelijaBroj = 0
End Function

' PRAVA PILULA ("pill") KOJA SE NE MOZE NASLIKATI SE BRISE CELA, ne samo natpis.
'
' Vazi SAMO za "pill". Dve vrste imaju dva ugovora, i mesanje to dvoje je greska
' koja je jednom vec napravljena bas ovde:
'
'   "pill"     PaintPill sam racuna sirinu po tekstu (PillW) i sam postavlja
'              BackColor i BackStyle. LayoutGrid tu kolonu NAMERNO preskace pri
'              dodeli sirine, a PaintRow je preskace pri vracanju pozadine reda.
'              Sve sto naslika, mora i da se ocisti -- inace ostane PRAZNA
'              OBOJENA KUTIJA nad novim podatkom.
'
'   "paypill"  PaintPayPill menja samo natpis, boju teksta, poravnanje i bold.
'              Sirinu drzi LayoutGrid (mColW - 16), a pozadinu vraca PaintRow.
'              Ko je ocisti kao pravu pilulu, postavi joj sirinu na PUNU sirinu
'              kolone -- i ona takva ostane, jer PaintPayPill sirinu ne vraca a
'              LayoutGrid se ponovo pusta tek kad se promeni opis kolona. Zato
'              se toj koloni brise SAMO natpis.
'
' Providna pozadina bez natpisa pusta boju reda da se vidi. Napomena iz PaintRow
' o mesanju providnih i neprozirnih kontrola tice se ISCRTAVANJA TEKSTA, a ovde
' teksta nema.
'
' Stil ostalih celija se ovde NE dira. Kanonski stil kolone postavlja
' StyleGridCell iz LayoutGrid-a (font, velicina, bold, boja, visina), a
' poravnanje LayoutGrid sam -- i to se preracuna cim se opis kolona promeni
' (mGeomStara). Crtanje koje bi ga "resetovalo" na neutralno samo bi ga
' pokvarilo: brojevi bi presli levo, a rsd/mult i prva kolona bi izgubili bold.
Private Sub OcistiPilulu(lbl As Object, ByVal sirina As Single)
    On Error Resume Next
    lbl.caption = vbNullString
    lbl.BackStyle = fmBackStyleTransparent
    lbl.width = sirina
End Sub

Private Sub PaintPill(lbl As Object, ByVal code As Long)
    Dim cap As String
    On Error Resume Next
    Select Case code
        Case 2
            cap = Poruka("OTKUI_STAT_OTKAZANA")
            lbl.BackColor = C_PILL_ERR_BG: lbl.ForeColor = C_PILL_ERR_FG
        Case 1
            cap = Poruka("OTKUI_STAT_POSLATO")
            lbl.BackColor = C_PILL_SND_BG: lbl.ForeColor = C_PILL_SND_FG
        Case Else
            ' dokument bez zbirne je NACRT - neutralno stanje, ne uspeh;
            ' zelena je rezervisana za zavrsen tok (placeno / fakturisano)
            cap = Poruka("OTKUI_STAT_SACUVANA")
            lbl.BackColor = C_DISABLED: lbl.ForeColor = C_MUTED
    End Select
    lbl.caption = "  " & cap & "  "
    lbl.TextAlign = fmTextAlignLeft
    lbl.width = PillW(cap)
    lbl.BackStyle = fmBackStyleOpaque
    lbl.Font.bold = True
End Sub

' gruba procena sirine teksta u Segoe UI (MSForms nema merenje bez AutoSize)
Private Function PillW(ByVal cap As String) As Single
    PillW = Len(cap) * TS_BODY * 0.52 + 18
    If PillW < 46 Then PillW = 46
End Function

Private Sub PaintRow(body As Object, ByVal i As Long, ByVal isSel As Boolean, _
                     ByVal isHover As Boolean, Optional ByVal isMark As Boolean = False)
    Dim bg As Long, c As Long
    On Error Resume Next
    If isMark Then
        ' oznacen red ima svoju boju i kad je izabran - inace se posle klika ne
        ' vidi da li je oznaka ostala
        bg = C_SOFT_BG
    ElseIf isSel Then
        bg = C_SAND
    ElseIf isHover Then
        bg = C_ROW_HOVER
    ElseIf (i Mod 2) = 1 Then
        bg = C_ROW_ALT
    Else
        bg = C_WHITE
    End If
    body.Controls("rb" & i).BackColor = bg
    ' SVE celije neprozirne, u boji reda. Mesanje providnih i neprozirnih
    ' Label-a u istoj mrezi daje razlicito iscrtavanje teksta - MSForms na
    ' providnom Label-u gubi ClearType i prelazi na sivo zagladjivanje, pa
    ' parni i neparni redovi izgledaju kao da imaju razlicit font.
    For c = 0 To MAX_COLS - 1
        If ColKind(c) = "pill" Then GoTo NextCell
        body.Controls("c" & i & "_" & c).BackStyle = fmBackStyleOpaque
        body.Controls("c" & i & "_" & c).BackColor = bg
NextCell:
    Next c
End Sub

Private Sub RenderPager(ft As Object)
    Dim i As Long, pages As Long, base As Long
    pages = 1
    If mPageSize > 0 Then pages = -Int(-mViewN / mPageSize)
    If pages < 1 Then pages = 1
    base = mPage - 2
    If base < 1 Then base = 1
    If base + 4 > pages Then base = pages - 4
    If base < 1 Then base = 1
    For i = 1 To 5
        Dim pn As Long: pn = base + i - 1
        BoxShow ft, "pg" & i, (pn <= pages)
        ft.Controls("pg" & i & "C").caption = CStr(pn)
        ft.Controls("pg" & i).tag = "pg:" & pn
        BoxState ft, "pg" & i, IIf(pn = mPage, C_FOREST, C_WHITE), _
                 IIf(pn = mPage, C_CREAM, C_FOREST), (pn = mPage)
    Next i
    BoxShow ft, "pgPrev", (pages > 1)
    BoxShow ft, "pgNext", (pages > 1)
End Sub

' Radnja se salje ekranu, a greska iz njega se PRIKAZUJE. Omotac u
' modUiScreens gusi greske da ekran koji padne ne obori aplikaciju - ali
' gusenje bez traga znaci dugme koje "ne radi" i niko ne zna zasto.
Private Function ScrAct(ByVal tag As String) As Boolean
    ScrAct = modUiScreens.ScrEvent(mScreen, tag, "Click")
    If Len(modUiScreens.ScrLastErr) > 0 Then
        Debug.Print "modOtkupUI.ScrAct: " & modUiScreens.ScrLastErr
        ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & modUiScreens.ScrLastErr, True
    End If
End Function

' Kljuc liste iza dugmeta prekidaca.
' Redni broj cipa iz taga, ili -1 ako tag nije cip. Javna je da bi se ugovor
' ljuske ("svaki cip koji ekran prijavi ume da se izabere") mogao TVRDITI u
' testu -- klik kroz formu se u harnessu ne moze odigrati.
' PODRAZUMEVANI SORT PO LISTI -- cist ugovor, bez stanja: rang-liste
' (KOOPERANTI na Dokumentima, RANG na Izvestajima) se otvaraju po rangu
' rastuce, sve ostalo po drugoj koloni opadajuce (datum u dokumentima).
' Javno da bi se ugovor mogao TVRDITI u testu -- klik kroz formu se u
' harnessu ne moze odigrati (isti razlog kao SegIndeksIzTaga).
Public Sub SortZaListu(ByVal kljuc As String, ByRef col As Long, ByRef asc As Boolean)
    If kljuc = "KOOPERANTI" Or kljuc = "RANG" Then
        col = 1: asc = True
    Else
        col = 2: asc = False
    End If
End Sub

' Ekran prvo dobija priliku da kaze svoj sort; ako ga nema, vazi pravilo ljuske.
' Bez toga bi sifarnici dobili "druga kolona opadajuce" -- pravilo napisano za
' datum dokumenta, koje nad Stanicama znaci nazive unazad.
Private Sub PrimeniSortZaListu(ByVal kljuc As String)
    Dim spec As String, p() As String
    spec = modUiScreens.ScrSort(mScreen)
    If Len(spec) > 0 Then
        p = Split(spec, ":")
        If UBound(p) >= 1 Then
            If val(p(0)) >= 1 Then
                mSortCol = CLng(val(p(0)))
                mSortAsc = (LCase$(Trim$(p(1))) = "asc")
                mSortLista = kljuc
                Exit Sub
            End If
        End If
    End If
    SortZaListu kljuc, mSortCol, mSortAsc
    mSortLista = kljuc
End Sub

' TEST SEAM: postavi sort kao da ga je ostavio DRUGI ekran (aktivacioni
' lifecycle test) i izvrsi aktivacioni korak -- ista procedura koju zove
' ActivateScreen, ne kopija. Tvrdo gejtovani.
Public Sub GridSortSetTest(ByVal col As Long, ByVal asc As Boolean, _
                           ByVal lista As String)
    If Not IsTestMode() Then Exit Sub
    mSortCol = col
    mSortAsc = asc
    mSortLista = lista
End Sub

Public Sub GridSortAktivacijaTest()
    If Not IsTestMode() Then Exit Sub
    PrimeniSortZaListu ActiveLista()
End Sub

' TEST SEAM: aktivan ekran za lifecycle testove (ActiveLista cita
' mScreen; bez forme se ekran ne moze aktivirati). Tvrdo gejtovan.
Public Sub GridScreenSetTest(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mScreen = kljuc
End Sub

' TEST SEAM: tekuci sort mreze (kolona, smer) -- tvrdo gejtovan.
Public Function GridSortTest(ByRef col As Long, ByRef asc As Boolean) As Boolean
    If Not IsTestMode() Then Exit Function
    col = mSortCol
    asc = mSortAsc
    GridSortTest = True
End Function

Public Function SegIndeksIzTaga(ByVal tag As String) As Long
    Dim rep As String
    SegIndeksIzTaga = -1
    If Left$(tag, 5) <> "lsSeg" Then Exit Function
    rep = Mid$(tag, 6)
    If Len(rep) = 0 Then Exit Function
    If Not IsNumeric(rep) Then Exit Function
    SegIndeksIzTaga = CLng(rep)
End Function

Private Function SegKey(ByVal i As Long) As String
    Dim seg As Variant
    seg = SegDefs()
    If Not IsArray(seg) Then Exit Function
    If i < 0 Or i > UBound(seg) Then Exit Function
    SegKey = Split(CStr(seg(i)), "|")(0)
End Function

' Pokretanje radnje nad redom. Ljuska proverava samo ono sto je njeno: da je
' red izabran kad radnja to trazi. Sve ostalo je na ekranu.
Private Sub RunRowAction(ByVal i As Long)
    Dim k As String, trebaRed As Boolean
    k = ActField(i, 0)
    If Len(k) = 0 Then Exit Sub

    ' oznacavanje vise redova je stanje MREZE, pa ga drzi ljuska
    If k = "mark" Then
        mMarkOn = Not mMarkOn
        If Not mMarkOn Then Set mMark = Nothing
        RefreshRowActions
        RenderGrid
        Exit Sub
    End If

    trebaRed = (ActField(i, 4) = "1")
    If k = "spec" Then trebaRed = (MarkCount() = 0 And mSelRow <= 0)
    If trebaRed And mSelRow <= 0 Then
        ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Sub
    End If
    If k = "spec" And MarkCount() = 0 And mSelRow <= 0 Then
        ShowToast Poruka("OTKUI_ERR_NEMA_OTP"), True
        Exit Sub
    End If

    ' True znaci "podaci su promenjeni" - tada se sve preracunava
    If ScrAct("act:" & k & ":" & mSelRow) Then
        mSelRow = 0
        RefreshFromData
        RefreshOtpTraka mFrm
    End If
End Sub

' Prekidac lista aktivnog ekrana: niz redova "KLJUC|natpis|naslov|sirina".
' Prazno = ekran nema prekidac.
Private Function SegDefs() As Variant
    On Error Resume Next
    SegDefs = modUiScreens.ScrListe(mScreen)
End Function

' Radnje nad redom za aktivnu listu: niz redova
' "kljuc:natpis:sirina:stil:trebaRed".
Private Function ActDefs() As Variant
    Dim sRad As String
    On Error Resume Next
    sRad = modUiScreens.ScrRadnje(mScreen)
    If Len(sRad) = 0 Then Exit Function
    ActDefs = Split(sRad, "|")
End Function

Private Function ActField(ByVal i As Long, ByVal idx As Long) As String
    Dim act As Variant, p As Variant
    act = ActDefs()
    If Not IsArray(act) Then Exit Function
    If i > UBound(act) Then Exit Function
    p = Split(CStr(act(i)), ":")
    If idx > UBound(p) Then Exit Function
    ActField = CStr(p(idx))
End Function

' Opseg datuma pripada listi otpremnica u F1 - jedini deo reda radnji koji nije
' dugme, pa ga ekran ne moze prijaviti kroz Scr_Radnje.
Private Function SpecDatLista() As Boolean
    If mScreen <> "DOKUMENTI" Then Exit Function
    SpecDatLista = (ActiveLista() = "OTPREMNICE")
End Function

' Ugaseno dugme mora da IZGLEDA ugaseno - inace operater klikne i nista se ne
' desi. Ista dva stanja kao kod "Filteri": prigusen tekst dok je mrtvo, pun
' kad radi.
' Opseg datuma nije jedna kontrola nego sest (dva natpisa, dve kutije, dva
' polja) - BoxShow ume samo dugme, pa ovde ide rucno.
Private Sub ShowSpecDat(z As Object, ByVal vis As Boolean)
    Dim nmv As Variant, i As Long
    On Error Resume Next
    nmv = Array("specOdL", "specOdB", "specOdF", "specOdT", _
                "specDoL", "specDoB", "specDoF", "specDoT")
    For i = 0 To UBound(nmv)
        z.Controls(CStr(nmv(i))).Visible = vis
    Next i
    BoxShow z, "btnRedSpecDat", vis
End Sub

Private Sub RefreshRowActions()
    Dim z As Object, act As Variant, p As Variant, i As Long
    Dim on_ As Boolean, fc As Long, bg As Long
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set z = mFrm.Controls("zGrid")
    act = ActDefs()
    If Not IsArray(act) Then Exit Sub
    For i = 0 To BazenStaje(UBound(act) + 1, MAX_ACT, "radnje") - 1
        p = Split(CStr(act(i)), ":")
        ' peto polje: 1 = radnja trazi izabran red, 0 = radi bez njega
        on_ = True
        If UBound(p) >= 4 Then
            If CStr(p(4)) = "1" Then on_ = (mSelRow > 0)
        End If
        ' oznacavanje je jedina radnja koju obradjuje sama ljuska; ono je
        ' "upaljeno" dok traje, i nosi broj oznacenih redova
        If CStr(p(0)) = "mark" Then
            z.Controls("btnAct" & i & "C").caption = Poruka(CStr(p(1))) & _
                                                     IIf(mMarkOn, " " & MarkCount(), "")
            BoxState z, "btnAct" & i, IIf(mMarkOn, C_SOFT_BG, C_WHITE), _
                     IIf(mMarkOn, C_GREEN, C_MUTED), mMarkOn
            GoTo SledecaR
        End If
        ' specifikacija radi i nad oznacenim redovima, ne samo nad izabranim
        If CStr(p(0)) = "spec" Then on_ = (MarkCount() > 0 Or mSelRow > 0)
        bg = C_WHITE
        fc = C_FOREST
        Select Case CStr(p(3))
            Case "danger": fc = C_RUST
            Case "soft":   If on_ Then bg = C_SOFT_BG
                           fc = C_GREEN
        End Select
        If Not on_ Then fc = C_DISABLED_FG
        BoxState z, "btnAct" & i, bg, fc, (CStr(p(3)) = "soft" And on_)
SledecaR:
    Next i
End Sub

'------------------------------------------------------- OZNACENI REDOVI
' Kljuc reda je njegova PRVA KOLONA (broj otpremnice) - preziveo sortiranje,
' pretragu i promenu strane, sto redni broj ne bi.
Private Function RowKeyAt(ByVal r As Long) As String
    On Error Resume Next
    If r < 1 Or r > mViewN Then Exit Function
    RowKeyAt = Trim$(CStr(mView(r, 1)))
End Function

Private Function RowMarked(ByVal r As Long) As Boolean
    If mMark Is Nothing Then Exit Function
    RowMarked = mMark.Exists(RowKeyAt(r))
End Function

Private Function MarkCount() As Long
    If mMark Is Nothing Then Exit Function
    MarkCount = mMark.count
End Function

Private Sub ToggleMark(ByVal tag As String)
    Dim i As Long, r As Long, k As String
    i = RowIndexFromTag(tag)
    If i < 0 Then Exit Sub
    r = (mPage - 1) * mPageSize + i + 1
    k = RowKeyAt(r)
    If Len(k) = 0 Then Exit Sub
    If mMark Is Nothing Then Set mMark = CreateObject("Scripting.Dictionary")
    If mMark.Exists(k) Then mMark.Remove k Else mMark(k) = True
    ' Red na koji je kliknuto ostaje i TEKUCI red. Bez toga oznacavanje ostavi
    ' ljusku bez izabranog reda, pa radnja koja ume da radi "nad izabranim" (a
    ' specifikacija ume) nema na cemu da radi cim oznake zataje.
    mSelRow = r
End Sub

' Prazni i oznake i sam rezim - zove se pri svakoj promeni liste, rezima ili
' ekrana, da oznake iz jedne liste ne prezive u drugu.
Private Sub ClearMarks()
    mMarkOn = False
    Set mMark = Nothing
End Sub

' Kljucevi oznacenih redova, spojeni "|". Ekran ih tumaci - ljuska ne zna sta
' su brojevi u prvoj koloni.
' Opseg datuma iznad liste otpremnica. Nije filter ljuske nego PARAMETAR koji
' ekran cita i primenjuje na svoje redove - ljuska ne zna sta je datum
' otpremnice. Isti par cita i stampa specifikacije.
Public Function GridDatOd() As String
    On Error Resume Next
    GridDatOd = Trim$(CStr(mFrm.Controls("zGrid").Controls("specOdT").text))
End Function

Public Function GridDatDo() As String
    On Error Resume Next
    GridDatDo = Trim$(CStr(mFrm.Controls("zGrid").Controls("specDoT").text))
End Function

' Dijagnostika iz Immediate prozora: ?OtkupUI_DiagMark()
Public Function OtkupUI_DiagMark() As String
    OtkupUI_DiagMark = "rezim=" & mMarkOn & " oznaka=" & MarkCount() & _
                       " red=" & mSelRow & " lista=" & ActiveLista() & _
                       " kljucevi=[" & MarkedKeys() & "]"
End Function

Public Function MarkedKeys() As String
    Dim kk As Variant, res As String
    If mMark Is Nothing Then Exit Function
    For Each kk In mMark.keys
        res = res & IIf(Len(res) > 0, "|", "") & CStr(kk)
    Next kk
    MarkedKeys = res
End Function

' Brojaci pripadaju cipovima liste dokumenata. Kad je slot pozajmljen ekranu,
' upis bi pregazio natpis koji je ekran dao.
Private Sub RenderChipCounts(z As Object)
    On Error Resume Next
    If mScrChipN > 0 Then Exit Sub
    z.Controls("chipOtkazaneC").caption = Poruka("OTKUI_CHIP_OTKAZANE") & "  " & mCntOtkaz
    z.Controls("chipBezZbirneC").caption = Poruka("OTKUI_CHIP_BEZZBIRNE") & "  " & mCntBezZb
    z.Controls("chipNefaktC").caption = Poruka("OTKUI_CHIP_NEFAKT") & "  " & mCntNefakt
    z.Controls("chipOtvoreneC").caption = Poruka("OTKUI_CHIP_OTVORENE") & "  " & mCntOtvor
End Sub

' Otpremnica i prijemnica pri ulasku preuzimaju broj zbirne sa kojom se radilo.
' U frmDokumenta je to radila RefreshBrojZbirneSuggestion (isti broj u obe
' sekcije) plus btnUnosZbr posle snimanja; ovde je jedan sticky broj.
' Rucni unos ima prednost - vec popunjeno polje se NE gazi.
' Skup polja zavisi od dokumenta. Gotovinski rezimi nemaju ni kilograme ni
' cenu ni ambalazu - ambalaza zivi iskljucivo u rezimu reversa. LayoutFields
' preskace sakrivena polja, pa se preostala sama presloze u mrezu.
Private Sub ApplyFormFields(frm As Object, ByVal mode As String)
    Dim z As Object
    On Error Resume Next
    Set z = frm.Controls("zForm")
    Select Case mode
        Case "F5", "F6"                     ' cist gotovinski promet
            FldShow z, "fgBrZbir", False
            FldShow z, "fgNovac", True
            FldShow z, "fgKgI", False
            FldShow z, "fgKgII", False
            FldShow z, "fgCena", False
            FldShow z, "fgKolAmb", False
            FldShow z, "fgKolAmbII", False
            FldShow z, "fgTipAmb", False
            FldShow z, "fgAmbPr", False
            FldShow z, "fgParcela", False
            FldShow z, "fgSmerRev", False
            FldShow z, "fgVrednost", False
            ' otvoreni blok i neisplaceni ostatak imaju smisla samo kod isplata
            FldShow z, "fgBlok", (mode = "F5")
            FldShow z, "fgOstatak", (mode = "F5")
            ' odakle novac izlazi pita se samo kod isplata; kod uplata se pita
            ' sta se zatvara, pa je tamo faktura umesto prekidaca
            FldShow z, "fgAvans", (mode = "F5")
            FldShow z, "fgFaktura", (mode = "F6")
        Case "F7"                           ' ambalaza: samo tip i kolicina
            FldShow z, "fgBrZbir", False
            FldShow z, "fgNovac", False
            FldShow z, "fgKgI", False
            FldShow z, "fgKgII", False
            FldShow z, "fgCena", False
            FldShow z, "fgKolAmb", True
            FldShow z, "fgKolAmbII", False
            FldShow z, "fgTipAmb", True
            FldShow z, "fgAmbPr", False
            FldShow z, "fgParcela", False
            FldShow z, "fgSmerRev", True
            FldShow z, "fgVrednost", False
            FldShow z, "fgBlok", False
            FldShow z, "fgOstatak", False
            FldShow z, "fgAvans", False
            FldShow z, "fgFaktura", False
        Case Else                           ' robni dokumenti
            ' u rezimu Zbirne broj dokumenta JESTE broj zbirne - zasebno polje
            ' bi bilo isti podatak dvaput
            FldShow z, "fgBrZbir", ModeVezujeZbirnu(mode)
            ' NOVAC NIJE POLJE ROBNOG DOKUMENTA. Kes isplate idu iskljucivo kroz
            ' F5 (Isplate) i F6 (Kupci - uplate); otkupni list ih vise ne nosi,
            ' iako legacy frmOtkup ima txtNovac/txtPrimalac pod KES_ISPLATE.
            FldShow z, "fgNovac", False
            FldShow z, "fgKgI", True
            FldShow z, "fgKgII", True
            FldShow z, "fgCena", True
            FldShow z, "fgKolAmb", True
            FldShow z, "fgKolAmbII", (mKlasa = 2)
            FldShow z, "fgTipAmb", True
            ' prazna ambalaza se izdaje uz otkup i vraca uz prijemnicu;
            ' otpremnica i zbirna je ne dodiruju
            FldShow z, "fgAmbPr", (mode = "F1" Or mode = "F4")
            FldShow z, "fgParcela", (mode = "F1" And IsPracenjeParcela())
            FldShow z, "fgSmerRev", False
            FldShow z, "fgVrednost", True
            FldShow z, "fgBlok", False
            FldShow z, "fgOstatak", False
            FldShow z, "fgAvans", False
            FldShow z, "fgFaktura", False
    End Select
End Sub

' Cetiri smera reversa su medjusobno iskljuciva - isti obrazac kao SetKlasa.
' n = 0 gasi sve cetiri (stanje "nije izabrano", u koje se rezim i vraca).
Private Sub SetSmerRev(ByVal n As Long)
    Dim z As Object, i As Long, sel As Boolean
    On Error Resume Next
    Set z = mFrm.Controls("zForm").Controls("fgSmerRev")
    mSmerRev = n
    MarkDirty
    For i = 1 To 4
        sel = (i = n)
        BoxState z, "segRev" & i, IIf(sel, C_FOREST, C_WHITE), IIf(sel, C_CREAM, C_MUTED), sel
    Next i
End Sub

' Isplata iz OM avansa vs virman firme - dva segmenta, isti obrazac. Prekidac
' se pusta samo kad su kes isplate ukljucene (legacy: tglIzOMAvansa.enabled =
' IsKesIsplate() And primalac izabran); kad nisu, izbor ostaje na virmanu.
Private Sub SetAvans(ByVal n As Long)
    Dim z As Object
    On Error Resume Next
    If n = 2 And Not IsKesIsplate() Then
        ShowToast Poruka("OTKUI_ERR_KES_ISKLJUCEN"), True
        Exit Sub
    End If
    Set z = mFrm.Controls("zForm").Controls("fgAvans")
    mIzAvansa = (n = 2)
    MarkDirty
    BoxState z, "segAvans1", IIf(n = 1, C_FOREST, C_WHITE), IIf(n = 1, C_CREAM, C_MUTED), (n = 1)
    BoxState z, "segAvans2", IIf(n = 2, C_FOREST, C_WHITE), IIf(n = 2, C_CREAM, C_MUTED), (n = 2)
End Sub

Private Sub FldShow(z As Object, ByVal nm As String, ByVal vis As Boolean)
    On Error Resume Next
    z.Controls(nm).Visible = vis
End Sub

Private Sub ApplyAktivnaZbirna(frm As Object, ByVal mode As String)
    Dim fld As Object
    On Error Resume Next
    Set fld = frm.Controls("zForm").Controls("fgBrZbir")
    If Not ModeVezujeZbirnu(mode) Then Exit Sub
    If Not fld.Visible Then Exit Sub
    If Len(mAktivnaZbirna) = 0 Then Exit Sub
    If Len(Trim$(fld.Controls("fgBrZbirT").text)) > 0 Then Exit Sub
    fld.Controls("fgBrZbirT").text = mAktivnaZbirna
End Sub

' Pamti broj zbirne kao "aktivan". Zove se iz F3 (broj dokumenta JESTE broj
' zbirne) i sa polja BROJ ZBIRNE u ostalim rezimima.
Private Sub SetAktivnaZbirna(ByVal broj As String)
    broj = Trim$(broj)
    If Len(broj) = 0 Then
        ' prazno polje znaci "ovaj dokument NEMA zbirnu" - bez ovoga bi
        ' ApplyAktivnaZbirna vratila stari broj pri sledecem ulasku u rezim
        mAktivnaZbirna = ""
        Exit Sub
    End If
    If StrComp(broj, mAktivnaZbirna, vbTextCompare) = 0 Then Exit Sub
    mAktivnaZbirna = broj
    ShowToast Poruka("OTKUI_MSG_ZBIRNA_AKTIVNA") & " " & broj, False
End Sub

Private Sub LoadRowIntoForm()
    Dim r As Long
    On Error Resume Next
    If mSelRow < 1 Or mSelRow > mViewN Then Exit Sub
    r = mSelRow
    SetFld "fgBrOtpr", CStr(mView(r, 1))
    ' U rezimu Zbirne broj dokumenta JESTE broj zbirne - klik na red je isto
    ' sto i lstZbirne_Click u staroj formi.
    If ActiveMode = "F3" Then SetAktivnaZbirna CStr(mView(r, 1))
    SetFld "fgKgI", FmtBroj(CDbl(mView(r, 4)), 0)
    If CDbl(mView(r, 4)) > 0 Then
        mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena1T").text = _
            FmtBroj(CDbl(mView(r, 5)) / CDbl(mView(r, 4)), 2)
    End If
    RecalcVrednost
    ShowToast Poruka("OTKUI_MSG_UCITANO") & " " & CStr(mView(r, 1)), False
End Sub

'=====================================================================
' REZIMI I STANJA
'=====================================================================
Public Sub SelectMode(frm As Object, key As String)
    SelectModeCore frm, key, True
End Sub

' Koliko dugmadi prekidaca ljuska uopste crta. Ekran koji prijavi vise lista od
' ovoga dobija TIHO odsecen rep -- LayoutGrid nacrta prvih MAX_SEG i stane, bez
' greske i bez traga. Ekran Storno je tako ostao bez liste "Izvodi": bila je u
' Scr_Liste, ali se nije mogla izabrati ni na koji nacin. Javno, da ta granica
' moze da se izmeri testom umesto da se otkrije na ekranu.
Public Function MaxPrekidaca() As Long
    MaxPrekidaca = MAX_SEG
End Function

' TEST SEAM: je li mapa imena kesirana pod datim kljucem. Prazna mapa i mapa
' koja nije kesirana daju isti rezultat pozivaocu, pa se razlika meri ovde.
Public Function MapaImenaKesirana(ByVal ck As String) As Boolean
    If mPartMap Is Nothing Then Exit Function
    MapaImenaKesirana = mPartMap.Exists(ck)
End Function

' Prelazak na drugi REZIM bez forme u argumentu - za ekran, koji formu nema.
' Postoji zbog ispravke posle storna (Faza D/13): posle storna starog dokumenta
' operater mora da zavrsi u rezimu u kome se unosi zamenski. Ekran zna KOJI je
' to rezim; formu drzi ljuska.
Public Sub IdiNaRezim(ByVal key As String)
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If Len(key) = 0 Then Exit Sub
    If key = ActiveMode Then Exit Sub
    SelectMode mFrm, key
End Sub

' Prelazak na drugi EKRAN bez forme u argumentu - par gornjem IdiNaRezim.
'
' Postoji zbog predaje ispravke: storno je svoj ekran, a zamenski dokument se
' unosi na ekranu dokumenata. Ceo tok je zato prelazak EKRANA -> prelazak
' REZIMA -> prefill, i redosled je obavezan (SelectMode cisti formu, pa bi
' prefill pre njega bio obrisan -- v. modScrStorno.OtvoriIspravku).
'
' ActivateScreen je Private i ostaje Private: ovo je jedini javni ulaz, pa
' ekran ne moze da preskoci proveru prava ni lenju gradnju zone.
'
' Sidebar se preboji POSLE prelaska, iz mScreen a ne iz kljuca: kad
' ActivateScreen odbije prelazak (nema prava, modul nedostaje), mScreen je
' ostao stari i oznaka mora da pokazuje ekran na kome zaista jesmo.
Public Sub IdiNaEkran(ByVal kljuc As String)
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If Len(kljuc) = 0 Then Exit Sub
    ' "Isti ekran" vise NIJE razlog za rano izlazenje: dok je panel otvoren,
    ' povratak na ekran ispod njega je isti kljuc -- a mora da zatvori panel.
    ' Odluku sada donosi ActivateScreen, koji zna i za panel.
    ActivateScreen mFrm, kljuc
    PaintNav mFrm, NavTagFor(AktivnaStavka())
End Sub

' Kljuc ekrana na kome jesmo. Javno zato sto je "storno je EKRAN, a ne rezim"
' tvrdnja koja se drugacije ne moze izmeriti: u testu se forma ne prikazuje,
' pa se vidljivost zona ne cita.
Public Function AktivanEkran() As String
    AktivanEkran = mScreen
End Function

Private Sub SelectModeCore(frm As Object, ByVal key As String, ByVal doReload As Boolean)
    Dim k As String
    On Error Resume Next
    mPopMute = True                    ' punjenje lista pise u combo -> ne otvaraj panel
    mLoading = True                    ' i to NIJE izmena korisnika
    ClosePopup
    ActiveMode = key
    k = modeKey(key)
    ' Ovde je do v6-ui-143 stajao grid-max za STORNO: F8 je crtao unosnu formu
    ' koju ne koristi, pa mu je forma sklanjana. Od kada je storno SVOJ EKRAN,
    ' svi preostali rezimi (F1..F7) su unosni i forma im pripada -- pa izuzetka
    ' nema. Grid-max ostaje samo kao operaterov izbor kroz prekidac.
    SetGridCols k                      ' pre LayoutGrid-a i pre naslova kolona
    ApplyFormFields frm, key           ' pre LayoutFields-a

    With frm.Controls("zTitle")
        .Controls("titName").caption = Poruka("OTKUI_TITLE_" & k)
        .Controls("titSub").caption = Poruka("OTKUI_SUB_" & k)
        .Controls("titBadgeC").caption = key
        .Controls("titIco").caption = ChrW(ModeIco(key))
        .Controls("titDatum").caption = FmtDatumPun(Now) & " " & ChrW(183) & " " & Poruka("OTKUI_SMENA_1")
    End With
    PlaceTitleBadge frm
    frm.Controls("zGrid").Controls("grdTitle").caption = Poruka("OTKUI_GRID_TITLE_" & k)
    RefreshListSeg frm
    RefreshOtpTraka frm
    ' Ime tabele je dijagnostika, ne informacija za operatera - u produkciji se
    ' ne prikazuje (UI_DEBUG=DA u tblLocalConfig ili dev build).
    frm.Controls("zGrid").Controls("grdSrc").caption = "list: " & ModeTable(key)
    frm.Controls("zGrid").Controls("grdSrc").Visible = IsDebugUI()
    frm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprL").caption = Poruka("OTKUI_FLDBROJ_" & k)
    ' "Sacuvaj" ne kaze STA se cuva - natpis nosi ime dokumenta
    frm.Controls("zForm").Controls("btnSacuvajC").caption = Poruka("OTKUI_BTN_SACUVAJ_" & k)
    ' partner je poslednje polje "Osnovnih podataka"; natpis prati sadrzaj
    ' liste - "Kupac" kad su u njoj samo kupci, inace "Partner"/"Primalac"
    ' UCase$ je OBAVEZAN - natpis se ovde prepisuje po rezimu, a NewLbl ga je
    ' pri gradnji stavio u verzal. Bez ovoga je "Kooperant" bio jedini natpis
    ' u redu koji nije verzal (VRSTA / SORTA / OTKUPNO MESTO / VOZAC).
    frm.Controls("zCtx").Controls("ctxL4").caption = UCase$(Poruka("OTKUI_FLDPART_" & k))
    FillFormPartner frm, key
    FillOpenBlokovi frm
    FillOpenFakture frm
    FillParcele frm
    ' Smer reversa i "isplata iz" pripadaju JEDNOM dokumentu - prelazak u drugi
    ' rezim ih vraca na pocetno stanje, kao i sve ostale izbore u formi.
    SetSmerRev 0
    SetAvans 1
    RefreshSaldoOM
    ' prazna ambalaza: uz otkup se izdaje, uz prijemnicu se vraca
    If key = "F1" Or key = "F4" Then
        frm.Controls("zForm").Controls("fgAmbPr").Controls("fgAmbPrL").caption = _
            UCase$(Poruka("OTKUI_FLD_AMB_PR_" & k))
    End If

    ClearMarks                       ' oznake pripadaju JEDNOJ listi jednog rezima
    mFilter = "danas"
    ' Lista otpremnica ima svoje cipove; "danas" bi u njoj pokazao praznu
    ' listu, a posao za koji ta lista postoji su bas neraspodeljene otpremnice.
    If ActiveLista() = "OTPREMNICE" Then mFilter = "otvorene"
    mSelRow = 0
    ApplyChipVisual frm.Controls("zGrid")
    ' "Bez zbirne" i "Nefakturisane" nemaju smisla nad zbirnom ni nad
    ' gotovinskim prometom (tblNovac nema BrojZbirne)
    ' povratak sa ugovornog ekrana - Filteri se vracaju samo ovde
    BoxShow frm.Controls("zGrid"), "btnFilteri", True
    ShowChip frm, "chipBezZbirne", ModeHasZbirna(key)
    ShowChip frm, "chipNefakt", ModeHasFaktura(key)
    ' Cipovi su se u F8 skrivali dok je taj rezim umeo samo da PRIKAZE
    ' stornirane iz tblOtpremnica - tada je svaki osim "Otkazane" izbacivao
    ' operatera iz storna. Od v6-ui-119 F8 cita svih devet tipova i stornira,
    ' pa mu cipovi rade isto sto i svuda: suzavaju listu, a "Otkazane" je i
    ' dalje pregled vec storniranih.
    ShowChip frm, "chipDanas", True
    ShowChip frm, "chipNedelja", True
    ShowChip frm, "chipSve", True
    ShowChip frm, "chipOtkazane", True
    ' Naslov i cipovi po LISTI idu POSLE cipova po rezimu, ne pre njih: red
    ' iznad postavlja cipove liste dokumenata bez obzira na izabranu listu, pa
    ' bi ranije pozvan RefreshGridTitle bio pregazen (u listi otpremnica su se
    ' vracali "Danas" / "Bez zbirne", koji tu nemaju sta da suze).
    RefreshGridTitle frm
    ' sirine kolona zavise od rezima - bez ovoga bi mreza zadrzala raspored
    ' prethodnog dokumenta (za vreme izgradnje raspored radi UserForm_Activate)
    If Not mBuilding Then LayoutOtkup frm
    RefreshSortGlyphs
    ApplyAktivnaZbirna frm, key
    RefreshModeCards frm
    RefreshKpi frm
    ' dodatni uslovi vaze preko rezima; dugme i panel moraju to da pokazu
    RefreshFilterBadge
    RenderFilterPanel
    ResetFokusOkvira
    ' podrazumevana vrsta/sorta iz Podesavanja - samo ako roba nije izabrana
    ApplyDefaultRoba
    ' nov dokument -> nista nije neupisano
    MarkClean
    mPopMute = False
    mLoading = False
    ' Svaki rezim ima svoj brojevni niz (otkupni list / otpremnica / zbirna /
    ' revers), pa se predlog racuna i pri promeni rezima - inace bi u polju
    ' ostao broj iz prethodnog niza. Rezimi bez niza ga ne diraju.
    If Not mBuilding Then RefreshBrojPredlog False
    If doReload Then
        mView = Empty: mViewN = 0
        ReloadGrid
    End If
End Sub

Private Sub RefreshModeCards(frm As Object)
    Dim z As Object, i As Long, m As Variant
    Dim mk As String, md As String, act As Boolean
    Set z = frm.Controls("zRight")
    ' redosled kartica prati brojeve tastera: F1..F7, svi unosni.
    ' Osma kartica (Storno, "warn" bojama) je otisla sa rezimom u v6-ui-143.
    m = Array("F1", "F2", "F3", "F4", "F5", "F6", "F7")
    For i = 0 To 6
        md = CStr(m(i))
        mk = modeKey(md)
        act = (md = ActiveMode)

        ' na kartici stoji KRATKO ime (108pt); pun naziv nosi veliki naslov
        z.Controls("mcT" & i).caption = Poruka("OTKUI_MODE_" & mk)
        z.Controls("mcT" & i).tag = md
        z.Controls("mcKey" & i).caption = md
        z.Controls("mcHint" & i).caption = Poruka("OTKUI_HINT_" & mk)

        If act Then
            ' aktivni: puna tamna ispuna + zlatna traka - vidi se iz daljine
            z.Controls("mc" & i & "B").BackColor = C_FOREST
            z.Controls("mc" & i & "F").BackColor = C_FOREST
            z.Controls("mc" & i & "A").BackColor = C_GOLD
            z.Controls("mcT" & i).ForeColor = C_CREAM
            z.Controls("mcHint" & i).ForeColor = RGB(168, 180, 160)
            z.Controls("mcKey" & i).ForeColor = C_GOLD
        Else
            z.Controls("mc" & i & "B").BackColor = C_BORDER_LT
            z.Controls("mc" & i & "F").BackColor = C_WHITE
            z.Controls("mc" & i & "A").BackColor = C_GOLD
            z.Controls("mcT" & i).ForeColor = C_FOREST
            z.Controls("mcHint" & i).ForeColor = C_MUTED
            z.Controls("mcKey" & i).ForeColor = C_MUTED
        End If
    Next i
    ' jedan prolaz kroz Btns za svih 7 kartica (svaka ima 4 vezane kontrole)
    RebaseSink "mcT", True            ' nove boje su sada osnova, ne stare
End Sub

Private Sub ApplyChipVisual(z As Object)
    Dim c As Object, on_ As Boolean
    For Each c In z.Controls
        If Left$(c.name, 4) = "chip" And Right$(c.name, 1) <> "B" And Right$(c.name, 1) <> "C" Then
            on_ = (ChipFilter(c.name) = mFilter)
            ' Crveni cip je 'Otkazane' liste dokumenata. Isti slot pozajmljen
            ' ekranu nema veze sa stornom, pa nosi obicnu boju.
            If c.name = "chipOtkazane" And mScrChipN = 0 Then
                BoxState z, c.name, IIf(on_, C_RUST, C_PILL_ERR_BG), IIf(on_, C_CREAM, C_RUST), on_
            Else
                BoxState z, c.name, IIf(on_, C_FOREST, C_WHITE), IIf(on_, C_CREAM, C_FOREST), on_
            End If
        End If
    Next c
End Sub

' Neutralna strelica stoji SAMO pod pokazivacem. Trinaest istih strelica u
' zaglavlju je davalo utisak Excel tabele, a nijedna od njih nije nosila
' informaciju - sortiranje pokazuje jedino aktivna kolona. Umesto strelice
' se ispisuju dva razmaka priblizno iste sirine, da natpis ne poskakuje kad
' pokazivac udje i izadje iz celije.
Private Sub RefreshSortGlyphs()
    Dim hd As Object, i As Long
    On Error Resume Next
    Set hd = mFrm.Controls("zGrid").Controls("grdHead")
    For i = 0 To MAX_COLS - 1
        Dim g As String
        If mColW(i) = 0 Then
            hd.Controls("hd" & i).caption = ""
        Else
            If mSortCol = i + 1 Then
                g = IIf(mSortAsc, ChrW(9650), ChrW(9660))
            ElseIf mHoverHd = i Then
                g = ChrW(8597)
            Else
                g = "  "
            End If
            hd.Controls("hd" & i).caption = UCase$(GridHeadCaption(i)) & " " & g
            hd.Controls("hd" & i).ForeColor = IIf(mSortCol = i + 1, C_FOREST, C_MUTED)
        End If
    Next i
End Sub

' Ulaz/izlaz pokazivaca nad zaglavljem kolone. Crta se samo zaglavlje -
' RenderGrid ovde nema sta da radi.
Private Sub HoverHead(ByVal i As Long)
    If i = mHoverHd Then Exit Sub
    mHoverHd = i
    RefreshSortGlyphs
End Sub

Public Sub SelectNav(frm As Object, ByVal nm As String)
    PaintNav frm, nm
    Dim kljuc As String
    If Not mNavKey Is Nothing Then
        If mNavKey.Exists(nm) Then kljuc = CStr(mNavKey(nm))
    End If
    If Len(kljuc) > 0 Then ActivateScreen frm, kljuc
End Sub

' Stavka sidebara na kojoj operater STVARNO jeste. Dok je panel otvoren to je
' panel, inace ekran. Bez ovoga bi PaintNav posle svakog neuspelog prelaska
' vratio oznaku na mScreen -- a mScreen je ekran ISPOD panela, pa bi sidebar
' pokazivao Partnere dok se gledaju Podesavanja.
Private Function AktivnaStavka() As String
    Dim k As String
    k = modUiPanel.PanelAktivan()
    If Len(k) = 0 Then k = mScreen
    AktivnaStavka = k
End Function

' nav tag ekrana koji je aktivan - za vracanje oznake posle neuspelog prelaska
Private Function NavTagFor(ByVal kljuc As String) As String
    Dim k As Variant
    If mNavKey Is Nothing Then Exit Function
    For Each k In mNavKey.keys
        If CStr(mNavKey(k)) = kljuc Then
            NavTagFor = CStr(k)
            Exit Function
        End If
    Next k
End Function

' samo bojenje sidebara - bez prelaska
'
' PRIGUSENA STAVKA OSTAJE PRIGUSENA. Do sada je bojenje svakoj neizabranoj
' stavci vracalo punu boju, pa su Marza i Sledljivost izgledale prigusene samo
' do prvog klika bilo gde u meniju -- posle toga su izgledale kao da rade.
' U maticnoj sekciji bi to bilo tri od pet stavki, pa se ispravlja ovde.
'
' Stanje se cita iz mape napravljene pri gradnji (mNavOff), ne racuna se ovde:
' bojenje ide na svaki klik, a odgovor "postoji li modul i ima li korisnik
' pravo" je prolaz kroz registar. Mapa nastaje tacno tamo gde je taj odgovor
' vec izracunat, pa se ne moze razici sa onim sto BuildNav nacrta.
Private Sub PaintNav(frm As Object, ByVal nm As String)
    Dim c As Object, z As Object, on_ As Boolean, off_ As Boolean, idx As String
    On Error Resume Next                ' navIco moze da ne postoji (kod 0)
    Set z = frm.Controls("zNav")
    For Each c In z.Controls
        If Left$(c.name, 3) = "nav" And Len(c.name) = 5 Then
            on_ = (c.name = nm)
            off_ = NavPrigusena(c.name)
            idx = Mid$(c.name, 4, 2)
            c.BackColor = IIf(on_, C_FOREST, C_SAND)
            c.Font.bold = on_                       ' marker stanja za hover (clsFlatBtn)
            z.Controls(c.name & "X").ForeColor = _
                IIf(on_, C_CREAM, IIf(off_, C_DISABLED_FG, RGB(52, 68, 44)))
            z.Controls(c.name & "X").Font.bold = on_
            z.Controls("navBar" & idx).BackColor = IIf(on_, C_GOLD, C_SAND)
            z.Controls("navIco" & idx).ForeColor = _
                IIf(on_, C_GOLD, IIf(off_, C_DISABLED_FG, C_ICON_OFF))
        End If
    Next c
End Sub

' Da li je stavka sidebara prigusena (ekran nema modul ili nema prava).
' Javna zbog testa - boja kontrole se u harnessu moze procitati, ali odluka o
' njoj ne moze bez klika kroz formu.
Public Function NavPrigusena(ByVal nmStavke As String) As Boolean
    If mNavOff Is Nothing Then Exit Function
    If mNavOff.Exists(nmStavke) Then NavPrigusena = mNavOff(nmStavke)
End Function

Private Sub RefreshKpi(frm As Object)
    Dim z As Object, cOtp As Long, cPrj As Long, om As String, sal As Double
    Dim st As String
    If frm Is Nothing Then Exit Sub
    On Error GoTo EH
    st = "zKpi": Set z = frm.Controls("zKpi")

    st = "danas": z.Controls("kpiV0").caption = FmtBroj(SumKgForDate(TBL_OTKUP, COL_OTK_DATUM, COL_OTK_KOLICINA, Int(Now)), 0)
    ' 0 kg je TACAN podatak, ne odsustvo podatka - zato ostaje nula, a ne crta
    KpiSub z, 0, CountForDate(TBL_OTKUP, COL_OTK_DATUM, Int(Now)) & " " & Poruka("OTKUI_KPI_SUB_DOK")

    st = "OM saldo": om = SelectedStanicaID(frm)
    If Len(om) > 0 Then
        sal = SaldoOMUkupno(om)
        z.Controls("kpiV1").caption = FmtBroj(sal, 0)
        z.Controls("kpiV1").ForeColor = IIf(sal < 0, C_RUST, C_FOREST)
        KpiSub z, 1, Poruka(IIf(sal < 0, "OTKUI_KPI_SUB_DUG", "OTKUI_KPI_SUB_POTRAZ"))
    Else
        ' bez izabranog otkupnog mesta vrednost NE POSTOJI - em-crta, ne nula
        KpiPrazno z, 1
    End If

    ' otvoreno kg jos nema izvor u semi - dok ga nema, plocica je tiha
    KpiPrazno z, 2

    st = "dokumenata"
    cOtp = CountForDate(TBL_OTPREMNICA, COL_OTP_DATUM, Int(Now))
    cPrj = CountForDate(TBL_PRIJEMNICA, COL_PRJ_DATUM, Int(Now))
    z.Controls("kpiV3").caption = cOtp & " otpr " & ChrW(183) & " " & cPrj & " prij"

    z.Controls("kpiV4").caption = Poruka("OTKUI_KPI_SPREMNO")
    z.Controls("kpiV4").ForeColor = C_GREEN
    KpiSub z, 4, Poruka("OTKUI_KPI_SUB_OK")
    Exit Sub
EH:
    Debug.Print "modOtkupUI.RefreshKpi PAO na koraku [" & st & "]: " & Err.Number & " " & Err.description
End Sub

'=====================================================================
' DISPATCHER
'=====================================================================
Public Sub UiEvent(ByVal tag As String, ByVal ev As String, ByVal arg As Variant)
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Select Case ev
        Case "Click":    UiClick tag
        Case "DblClick": UiDblClick tag
        Case "Change":   UiChange tag
        Case "Hover":    UiHover tag
        Case "Focus":    CloseOverlays tag: PostaviFokus tag
        Case "Blur":     SkiniFokus tag
        Case "KeyDown":  HandleGlobalKey CLng(arg), 0
        Case "Menu":     UiMenu tag
        Case "Drop":     OpenPopupFor tag
        Case "Commit":   ' rezervisano: normalizacija vrednosti posle unosa
        Case "Scroll":   ' rezervisano: zone sa sopstvenim klizacem
    End Select
End Sub

' Desni klik. Danas radi jedno: bira red na koji je pokazano, da radnja nad
' izabranom stavkom (S4) nikad ne odradi posao nad pogresnim redom. Meni sa
' radnjama dolazi uz radni sto otpremnice.
Private Sub UiMenu(ByVal tag As String)
    If RowIndexFromTag(tag) < 0 Then Exit Sub
    RowFromTag tag
    RenderGrid
End Sub

' Jedini dogadjaj koji sme da ODBIJE unos: klasa salje kod znaka, modul vraca
' kod koji prolazi (moze izmenjen) ili 0 = ne propusti.
Public Function UiAsk(ByVal tag As String, ByVal ev As String, ByVal arg As Variant) As Long
    UiAsk = CLng(arg)                   ' podrazumevano: propusti nepromenjeno
    On Error Resume Next
    If mFrm Is Nothing Then Exit Function
    Select Case ev
        Case "KeyPress": UiAsk = modUiKit.FilterKeyPress(tag, CLng(arg))
    End Select
End Function

' Taster iz kontrole KOJA GA JE PRIMILA. Postoji zato sto isti taster znaci
' razlicite stvari zavisno od toga gde je fokus: Alt+dole u combo-u otvara
' listu, a bilo gde drugde ne radi nista. F-tasteri ostaju globalni - F4 je
' rezim "Prijemnice" i u polju i van njega.
Public Function HandleKeyFrom(ByVal tag As String, ByVal KeyCode As Long, _
                              ByVal Shift As Long) As Boolean
    On Error Resume Next
    If KeyCode = vbKeyDown And (Shift And 4) <> 0 Then
        ' Alt+dole je standardni potez za otvaranje liste. Bez presretanja bi
        ' MSForms otvorio SVOJU listu - sa ivicom, klizacem i tudjim fontom,
        ' tacno ono sto je nas panel zamenio.
        If Len(PopKeyFor(tag)) > 0 Then
            OpenPopupFor tag
            HandleKeyFrom = True
            Exit Function
        End If
    End If
    HandleKeyFrom = HandleGlobalKey(KeyCode, Shift)
End Function

' Otvori/zatvori nas panel za kontrolu ciji je tag dat.
Private Sub OpenPopupFor(ByVal tag As String)
    Dim nm As String
    nm = PopKeyFor(tag)
    If Len(nm) = 0 Then Exit Sub
    If mPopFor = nm Then ClosePopup Else OpenPopup nm
End Sub

' Kljuc panela za dati tag kontrole, ili "" ako to nije combo. Kontekstni combo
' se zove isto kao svoj kljuc ("cbVrsta"), a combo u formi je kontrola unutar
' istoimenog okvira ("fgVrsta" -> "fgVrstaT") pa mu tag ima "T" viska. Provera
' tipa je obavezna u OBA slucaja: i tekstualno polje se zove nm & "T".
Private Function PopKeyFor(ByVal tag As String) As String
    Dim c As Object, nm As String
    On Error Resume Next
    Set c = FindCombo(tag)
    nm = tag
    If c Is Nothing And Len(tag) > 1 And Right$(tag, 1) = "T" Then
        nm = Left$(tag, Len(tag) - 1)
        Set c = FindCombo(nm)
    End If
    If c Is Nothing Then Exit Function
    If TypeName(c) <> "ComboBox" Then Exit Function
    PopKeyFor = nm
End Function

' MSForms UserForm NEMA KeyPreview. F-tasteri stizu samo sa kontrole koja drzi
' fokus i ima sink (TextBox / ComboBox preko clsFlatBtn) ili sa same forme dok
' nijedna kontrola nije fokusirana. Label NE MOZE da drzi fokus, pa klik na
' zaglavlje kolone, cip, paginaciju ili karticu rezima ostavlja fokus na Frame-u
' koji nema sink - i F-tasteri "prestanu da rade" dok korisnik rukom ne klikne u
' neko polje. Zato posle svakog klika na Label fokus vracamo u polje broja
' dokumenta. Izuzeci su klikovi koji sami upravljaju fokusom.
' Otvoreni panel se zatvara i kad fokus predje u drugu kontrolu (klik u polje,
' Tab), ne samo klikom. Sopstveni combo je izuzet - njegov Enter dolazi upravo
' pri otvaranju.
Private Sub CloseOverlays(ByVal tag As String)
    If Len(mPopFor) > 0 Then
        If tag <> mPopFor And tag <> mPopFor & "T" Then ClosePopup
    End If
    If mFltOpen Then CloseFilterPanel
End Sub

Private Sub UiClick(ByVal tag As String)
    UiClickCore tag
    Select Case True
        Case Left$(tag, 3) = "pop"                    ' izbor iz naseg dropdown-a
        Case Right$(tag, 1) = "D" And Len(tag) > 2    ' strelica dropdown-a
        Case tag = "btnClose", tag = "btnExcel"       ' forma se sakriva
        Case tag = "btnOperater"                      ' frmLogin je bio modalan
        Case Else: KeepKeys
    End Select
End Sub

Private Sub KeepKeys()
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If Not mFrm.Visible Then Exit Sub
    If Len(mPopFor) > 0 Then Exit Sub             ' panel je otvoren, ne diraj fokus
    mFrm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").SetFocus
End Sub

Private Sub UiClickCore(ByVal tag As String)
    Dim z As Object, c As Long

    ' skrol naseg dropdown-a - pre provere "pop" + broj, jer "popUp" nije broj
    If tag = "popUp" Then
        PopScroll -1
        Exit Sub
    End If
    If tag = "popDn" Then
        PopScroll 1
        Exit Sub
    End If
    ' izbor iz naseg dropdown-a
    If Left$(tag, 3) = "pop" And Len(tag) > 3 Then
        SelectPopup CLng(Mid$(tag, 4))
        Exit Sub
    End If
    ' klik u panelu Filteri ne sme da ga zatvori
    If Left$(tag, 3) = "flt" Then
        FilterPanelClick tag
        Exit Sub
    End If
    ' bilo koji drugi klik zatvara otvorene panele (osim njihovih sopstvenih dugmadi)
    If Len(mPopFor) > 0 And tag <> mPopFor & "D" Then ClosePopup
    If mFltOpen And tag <> "btnFilteri" Then CloseFilterPanel

    If Left$(tag, 2) = "rb" Or (Left$(tag, 1) = "c" And InStr(tag, "_") > 0) Then
        ' U rezimu oznacavanja klik NE bira red (u listi otpremnica izbor vodi
        ' pravo na blokove) nego ukljucuje/iskljucuje oznaku.
        If mMarkOn Then
            ToggleMark tag
            RefreshRowActions
        Else
            RowFromTag tag
        End If
        RenderGrid
        Exit Sub
    End If
    If Left$(tag, 3) = "nav" And Len(tag) = 5 Then
        SelectNav mFrm, tag
        Exit Sub
    End If
    If Left$(tag, 3) = "mcT" And Len(tag) = 4 Then
        Dim md As String: md = CStr(mFrm.Controls("zRight").Controls(tag).tag)
        ' aktivna kartica je sada u listi - klik na nju ne treba da preucitava
        If md <> ActiveMode Then SelectMode mFrm, md
        Exit Sub
    End If
    If Left$(tag, 4) = "chip" Then
        mFilter = ChipFilter(tag)
        Set z = mFrm.Controls("zGrid")
        ApplyChipVisual z
        RenderFilterPanel                ' period se bira na dva mesta - drzi ih u paru
        RefreshFilterBadge
        mSelRow = 0
        ReloadGrid
        Exit Sub
    End If
    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        c = CLng(Mid$(tag, 3)) + 1
        If mSortCol = c Then
            mSortAsc = Not mSortAsc
        Else
            mSortCol = c
            mSortAsc = True
        End If
        RefreshSortGlyphs
        ReloadGrid
        Exit Sub
    End If
    If Left$(tag, 2) = "pg" Then
        PagerClick tag
        Exit Sub
    End If
    ' prekidac liste - ljuska zna samo REDNI BROJ dugmeta; kljuc liste je
    ' ekranov, pa se cita iz njegove definicije i vraca mu nazad.
    '
    ' Redni broj se cita kroz SegIndeksIzTaga, ne merenjem duzine taga. Uslov je
    ' ranije glasio Len(tag) = 6, sto pokriva SAMO lsSeg0..lsSeg9 -- pa je
    ' jedanaesti cip (lsSeg10) imao sedam znakova, propadao kroz ovu granu i
    ' NIJE RADIO NISTA. Bez greske i bez traga: dugme se crta, boji se na hover,
    ' a klik nema kome da stigne.
    '
    ' Isti oblik greske kao MAX_SEG = 9 (v6-ui-143), samo na drugoj kapiji:
    ' tamo se cip nije CRTAO, ovde se crta ali ne reaguje. Zato ih meri i test
    ' odvojeno -- crtanje jednom tvrdnjom, dispecovanje drugom.
    Dim segI As Long: segI = SegIndeksIzTaga(tag)
    If segI >= 0 Then
        Dim segK As String
        segK = SegKey(segI)
        If Len(segK) = 0 Then Exit Sub
        If modUiScreens.ScrEvent(mScreen, "ls" & segK, "Click") Then
            mSelRow = 0
            mSearch = ""
            mFrm.Controls("zGrid").Controls("txtSearch").text = ""
            ' Sortiranje ne prelazi iz liste u listu: druga kolona je datum u
            ' listama dokumenata, a ime kooperanta u rangu - ista strelica bi
            ' znacila drugu stvar. Izbor zivi u SortZaListu (deli ga i
            ' RefreshFromData za auto-prelaz liste; recenzija #245, krug 16).
            PrimeniSortZaListu ActiveLista()
            RefreshListSeg mFrm
            If mScreen = "DOKUMENTI" Then
                ' povratak na "Svi listovi" vraca cipove koje je RefreshGridTitle
                ' sakrio; postavlja ih SelectModeCore, pa ide preko njega
                SelectModeCore mFrm, ActiveMode, True
                RefreshOtpTraka mFrm
            Else
                ' ugovorni ekran nema rezime ni cipove - dovoljni su naslov,
                ' radnje nad redom i ponovno citanje
                RefreshGridTitle mFrm
                ReloadGrid
            End If
            LayoutOtkup mFrm
        End If
        Exit Sub
    End If
    ' KONTROLE IZ ZONE UGOVORNOG EKRANA. Ljuska ih ne poznaje po imenu - prefiks
    ' "scr" je ceo dogovor, pa ekran moze da doda svoje dugme bez ijedne izmene
    ' ovde. Do v6-ui-143 zona nije imala nijednu kliktacu kontrolu (Oporavak i
    ' Palete drze samo labele), pa ovaj prolaz nije ni postojao - i dugme u zoni
    ' bi tiho "ne radilo".
    '
    ' True i dalje znaci "podaci su promenjeni" -- isto kao kod radnje nad redom.
    '
    ' Ali SAMO ako smo jos na istom ekranu. Radnja sme da PREDA posao drugom
    ' ekranu (storno -> ispravka -> unos zamenskog dokumenta), i tada je forma
    ' vec popunjena prefillom. RefreshFromData nad njom ponovo puni combo-e
    ' (FillZbirneCombo, mPartnerFor = "") -- a to obrise izbor partnera koji je
    ' prefill upravo postavio. Operater je to video kao "nisu sve stavke
    ' popunjene": sve sto je obicno polje ostane, a sve sto je lista se isprazni.
    ' Strelica combo-a se resava PRE grane ekrana: tag pocinje sa 'scr' kao i sve
    ' ostalo sto ekran nosi, pa bi inace otisao ekranu -- a on o panelu izbora ne
    ' zna nista, niti treba. Panel je ljuskin i radi isto nad svim poljima.
    If Right$(tag, 1) = "D" And Len(tag) > 2 Then
        If Len(PopKeyFor(Left$(tag, Len(tag) - 1))) > 0 Then
            OpenPopupFor Left$(tag, Len(tag) - 1)
            Exit Sub
        End If
    End If
    If Left$(tag, 3) = "scr" Then
        Dim preScr As String
        preScr = mScreen
        If ScrAct(tag) Then
            If mScreen = preScr Then
                mSelRow = 0
                RefreshFromData
            End If
        End If
        Exit Sub
    End If
    ' klik na strelicu "combo"-a -> nas panel (toggle)
    ' (CloseFilterPanel je vec odradjen gore)
    If Right$(tag, 1) = "D" And Len(tag) > 2 Then
        Dim cn As String: cn = Left$(tag, Len(tag) - 1)
        If mPopFor = cn Then
            ClosePopup
        Else
            OpenPopup cn
        End If
        Exit Sub
    End If

    Select Case tag
        Case "segKlasa1": SetKlasa 1
        Case "segKlasa2": SetKlasa 2
        Case "segRev1": SetSmerRev 1
        Case "segRev2": SetSmerRev 2
        Case "segRev3": SetSmerRev 3
        Case "segRev4": SetSmerRev 4
        Case "segAvans1": SetAvans 1
        Case "segAvans2": SetAvans 2
        Case "btnSacuvaj", "btnSacuvajPrint": CommitDokument (tag = "btnSacuvajPrint")
        Case "btnOtkazi": ClearForm
        Case "btnAmbEq": CopyAmbNaPraznu
        Case "btnMax"
            ' mreza se razvlaci do ispod naslova dokumenta - kontekst i forma
            ' se sklanjaju, drugi klik ih vraca
            mGridMax = Not mGridMax
            mFrm.Controls("zGrid").Controls("btnMaxI").caption = _
                ChrW(IIf(mGridMax, IC_MIN, IC_MAX))
            LayoutOtkup mFrm
            RenderGrid
        Case "btnClose": OtkupUI_Sakrij
        Case "btnSync": DoSync
        Case "btnSnimi": DoSaveWorkbook
        Case "btnExcel": DoShowExcel
        Case "btnOperater": DoSwitchOperater
        Case "btnFilteri": ToggleFilterPanel
        ' Radnja nad izabranim redom. Ljuska proverava SAMO da je red izabran;
        ' sta radnja znaci zna ekran (upis i stampa ostaju u poslovnim modulima).
        Case "btnRedSpecDat"
            ' Ekran sam procita opseg (GridDatOd / GridDatDo) - isti izvor po
            ' kome je i filtrirao listu, pa se stampa i lista ne mogu razici.
            ScrAct "act:specdat:0"
        Case "btnAct0", "btnAct1", "btnAct2", "btnAct3", "btnAct4"
            RunRowAction CLng(Mid$(tag, 7))
        Case "grdEmptyA"
            mFilter = "sve"
            mSelRow = 0
            ApplyChipVisual mFrm.Controls("zGrid")
            RenderFilterPanel
            RefreshFilterBadge
            ReloadGrid
        Case "btnMatic": PrebaciSekciju mFrm
    End Select
End Sub

'=====================================================================
' SEKCIJA LJUSKE
' Zlatno dugme u zaglavlju nije alatka nego ODREDISTE - isto sto je u legacy
' ljusci otvaralo ceo popup meni "Maticni podaci". Zato menja SKUP stavki
' sidebara, a ne otvara panel: sidebar tako i dalje govori istinu o tome gde
' si. Popup ispod dugmeta bi ostavio oznacenu stavku RADNE sekcije dok gledas
' Kupce. V. docs/UI_MIGRACIJA_KATALOG.md, 24.2.
'=====================================================================
Private Sub PrebaciSekciju(frm As Object)
    PostaviSekciju frm, IIf(mSekcija = SEK_MATICNI, SEK_RAD, SEK_MATICNI)
End Sub

' Javna zbog testa: klik kroz formu se u harnessu ne moze odigrati, a bas ovo
' je tvrdnja koju M0 nosi (isti razlog kao SegIndeksIzTaga i GridSortSetTest).
Public Sub OtkupUI_SekcijaTest(ByVal sekcija As String)
    PostaviSekciju mFrm, sekcija
End Sub

' Ono sto se desi posle zamene operatera, bez same zamene. Javno zbog testa:
' modAuth.Login trazi formu za prijavu, koja se u harnessu ne moze odigrati, a
' tvrdnja ("sidebar prati NOVA prava") je bas ovde.
Public Sub OtkupUI_PrimeniNovaPravaTest()
    PrimeniNovaPrava
End Sub

' Prelazak na ekran, sa VERDIKTOM. Javno zbog testa: klik kroz formu se u
' harnessu ne moze odigrati, a bas verdikt je ono na cemu stoji vracanje
' sekcije kad prelazak ne uspe (PostaviSekciju).
Public Function OtkupUI_PrelazakTest(ByVal kljuc As String) As Boolean
    OtkupUI_PrelazakTest = ActivateScreen(mFrm, kljuc)
End Function

' Radna povrsina bez ijednog dozvoljenog ekrana. Javno zbog testa: do tog
' stanja se u pogonu stize samo zamenom operatera na nalog bez prava, sto
' harness ne moze da odigra -- a bas tu je bilo curenje (mreza je zadrzavala
' redove prethodnog operatera).
Public Sub OtkupUI_IsprazniPovrsinuTest()
    If Not IsTestMode() Then Exit Sub
    mScreen = ""
    IsprazniRadnuPovrsinu mFrm
End Sub

' Zona ugovornog ekrana je promenila VISINU (npr. otvoren editor), pa raspored
' mora ponovo.
'
' Ekran ovo zove sam: ljuska ne moze da pogodi kada se to desilo, a
' LayoutScreenZone se izvrsava iskljucivo iz LayoutOtkup -- ni RefreshFromData
' ni ReloadGrid ga ne diraju (oni preracunavaju MREZU). Bez ovoga bi zona ostala
' na staroj visini, a polja editora bi se crtala ispod mreze -- vidljivo kao
' "dugme ne radi".
Public Sub OsveziRasporedEkrana()
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    LayoutOtkup mFrm
End Sub

Public Function AktivnaSekcija() As String
    AktivnaSekcija = mSekcija
End Function

' Visina koju su stavke AKTIVNE sekcije stvarno zauzele u poslednjem rasporedu.
' Javna zbog testa: mera na kojoj stoji cela odluka o sekcijama mora da se moze
' tvrditi, inace bi dodavanje ekrana tiho gurnulo stavke ispod profila.
Public Function NavVisinaSekcije() As Single
    NavVisinaSekcije = mNavVisina
End Function

' Slobodna visina za stavke sidebara na NAJMANJEM dozvoljenom prozoru. Sidebar
' nema skrol, pa je ovo tvrda granica, a ne preporuka.
Public Function NavSlobodnaVisina() As Single
    NavSlobodnaVisina = MIN_H - HEADER_H - STATUS_H - PROFIL_H
End Function

' Ime kontrole sidebara koja nosi dati ekran. Javno zbog testa - klik kroz
' formu se u harnessu ne moze odigrati.
Public Function NavTagZaEkran(ByVal kljuc As String) As String
    NavTagZaEkran = NavTagFor(kljuc)
End Function

Private Sub PostaviSekciju(frm As Object, ByVal sekcija As String)
    Dim kljuc As String, staraSekcija As String
    On Error GoTo EH
    If frm Is Nothing Then Exit Sub
    If sekcija <> SEK_RAD And sekcija <> SEK_MATICNI Then Exit Sub
    If sekcija = mSekcija Then Exit Sub

    ' Ekran sa kog se odlazi se PAMTI, po sekciji. Bez toga bi "Nazad na rad"
    ' uvek vracao na Unos dokumenata, pa bi provera jednog sifarnika usred
    ' fakturisanja kostala i povratak kroz meni.
    ZapamtiEkranSekcije mSekcija, mScreen

    staraSekcija = mSekcija
    mSekcija = sekcija
    OsveziDugmeSekcije frm
    LayoutNav frm

    kljuc = EkranZaSekciju(mSekcija)
    If Len(kljuc) = 0 Then
        ' Sekcija bez ijednog izgradjenog ekrana: oznaka u sidebaru se sklanja
        ' (nijedna stavka nije aktivna) i kaze se ZASTO. Glavna zona ostaje na
        ' prethodnom ekranu - prazna zona bi izgledala kao pad.
        PaintNav frm, ""
        ShowToast Poruka("OTKUI_SEK_NEMA_EKRANA"), False
        Exit Sub
    End If
    ' ActivateScreen sam radi LayoutOtkup na kraju - ne ponavlja se ovde, i sam
    ' izlazi kad se nista ne menja. Uslov "kljuc <> mScreen" je uklonjen jer je
    ' propustao jedan slucaj: panel otvoren nad ekranom koji je i cilj sekcije.
    '
    ' PRELAZAK MOZE DA NE USPE: operater odustane od odbacivanja nesacuvanog,
    ' ciljni ekran nema prava ili modul, gradnja padne. Do v6-ui-204 se sekcija
    ' menjala PRE toga i ostajala promenjena -- zlatno dugme i sidebar su
    ' pokazivali drugu sekciju od one iz koje je ekran koji se i dalje vidi.
    If Not ActivateScreen(frm, kljuc) Then
        mSekcija = staraSekcija
        OsveziDugmeSekcije frm
        LayoutNav frm
        PaintNav frm, NavTagFor(AktivnaStavka())
        Exit Sub
    End If
    PaintNav frm, NavTagFor(AktivnaStavka())
    Exit Sub
EH:
    LogErr "modOtkupUI.PostaviSekciju"
    ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Sub

Private Sub ZapamtiEkranSekcije(ByVal sekcija As String, ByVal kljuc As String)
    If Len(kljuc) = 0 Then Exit Sub
    If sekcija = SEK_MATICNI Then
        mZadnjiMaticni = kljuc
    Else
        mZadnjiRad = kljuc
    End If
End Sub

' Ekran na koji se ulazi u sekciju: zapamcen ako je jos dostupan, inace prvi
' dostupan iz registra. Provera je obavezna - prava se menjaju zamenom
' operatera, pa zapamcen ekran moze u medjuvremenu da postane zabranjen.
Private Function EkranZaSekciju(ByVal sekcija As String) As String
    Dim zapamcen As String
    zapamcen = IIf(sekcija = SEK_MATICNI, mZadnjiMaticni, mZadnjiRad)
    If Len(zapamcen) > 0 Then
        If modUiScreens.ScrSekcija(modUiScreens.ScrRowByKey(zapamcen)) = sekcija Then
            If modUiScreens.ScrAktivan(zapamcen) Then
                EkranZaSekciju = zapamcen
                Exit Function
            End If
        End If
    End If
    ' Registar odgovara koji je prvi dostupan - ljuska ne zna nijedan kljuc
    ' unapred.
    EkranZaSekciju = modUiScreens.ScrPrviUSekciji(sekcija)
End Function

' Natpis i glif zlatnog dugmeta prate sekciju: iz radne se ide u Maticne, iz
' maticne nazad na rad. Dugme koje u obe sekcije pise isto ne bi reklo gde si.
Private Sub OsveziDugmeSekcije(frm As Object)
    Dim z As Object
    On Error Resume Next
    Set z = frm.Controls("zHdr")
    If z Is Nothing Then Exit Sub
    If mSekcija = SEK_MATICNI Then
        z.Controls("btnMaticC").caption = Poruka("OTKUI_BTN_NAZAD_RAD")
        z.Controls("btnMaticI").caption = ChrW(IC_NAZAD)
    Else
        z.Controls("btnMaticC").caption = Poruka("OTKUI_BTN_MATICNI")
        z.Controls("btnMaticI").caption = ChrW(IC_MATICNI)
    End If
End Sub

'=====================================================================
' ALATKE APLIKACIJE
' Ove cetiri radnje NISU dokumenti pa nemaju veze sa CommitDokument-om
' (koji je namerno prazan sav). Vezane su na postojece rutine, ne na nove.
'=====================================================================
Private Sub DoSync()
    Dim ok As Boolean
    On Error GoTo EH
    ShowToast Poruka("OTKUI_MSG_SYNC_START"), False
    DoEvents
    ' postojeca ulazna tacka, napravljena bas za dugme (bez MsgBox-ova)
    ok = modGoogleSyncOrchestrator.SyncPWAFullCycle_ForButton()
    RefreshFromData                      ' kes tabela je posle sync-a zastareo
    RefreshSyncBadge mFrm, ok
    ShowToast IIf(ok, Poruka("OTKUI_MSG_SYNC_OK"), Poruka("OTKUI_MSG_SYNC_PAO")), Not ok
    Exit Sub
EH:
    RefreshSyncBadge mFrm, False
    ShowToast Poruka("OTKUI_MSG_SYNC_PAO"), True
End Sub

Private Sub RefreshSyncBadge(frm As Object, ByVal ok As Boolean)
    On Error Resume Next
    With frm.Controls("zRight")
        .Controls("syncDot").ForeColor = IIf(ok, C_GREEN, C_RUST)
        .Controls("syncTxt").Font.bold = True
        .Controls("syncTxt").ForeColor = IIf(ok, C_GREEN, C_RUST)
        .Controls("syncTxt").caption = IIf(ok, Poruka("OTKUI_SYNC_OK"), Poruka("OTKUI_SYNC_PAO"))
    End With
End Sub

Private Sub DoSaveWorkbook()
    On Error GoTo EH
    ThisWorkbook.Save
    ShowToast Poruka("OTKUI_MSG_WB_SNIMLJENA"), False
    Exit Sub
EH:
    ShowToast Poruka("OTKUI_MSG_WB_PALO"), True
End Sub

' Forma je modeless i preko celog ekrana, pa je Excel iza nje. Hide je jedini
' pouzdan nacin da se dodje do radnog lista: Unload se NE sme zvati odavde jer
' se ova rutina izvrsava iz clsFlatBtn eventa (sink forme je na steku).
' Sledece otvaranje je trenutno - 205 kontrola se ne gradi ponovo.
Private Sub DoShowExcel()
    On Error Resume Next
    Application.Visible = True
    mFrm.Hide
End Sub

' modAuth.Login sam radi Logout pa prikazuje frmLogin modalno - to JE zamena
' operatera. Otkazivanje prijave ne rusi sesiju (QuitAfterFailedLogin se zove
' iz StartApp, ne odavde), pa se u tom slucaju samo nista ne menja.
Private Sub DoSwitchOperater()
    Dim staro As String
    On Error GoTo EH
    If Not modAuth.AuthEnabled() Then
        ShowToast Poruka("OTKUI_MSG_AUTH_OFF"), True
        Exit Sub
    End If
    staro = OperaterText()
    If modAuth.Login() Then
        PrimeniNovaPrava
        RefreshOperater mFrm
        ShowToast Poruka("OTKUI_MSG_OPERATER") & " " & OperaterText(), False
    Else
        ' Neuspela prijava normalno VRACA prethodnu sesiju (modAuth.Login). Ako
        ' nije vracena -- nije je ni bilo, ili je vracanje puklo -- prikaz ne
        ' sme da tvrdi da je stari operater tu: prava se primenjuju na PRAZNU
        ' sesiju, sto radnu povrsinu i isprazni.
        If Not modAuth.JePrijavljen() Then
            PrimeniNovaPrava
            RefreshOperater mFrm
            ShowToast Poruka("OTKUI_MSG_ODJAVLJEN"), True
            Exit Sub
        End If
        ShowToast Poruka("OTKUI_MSG_OPERATER_ISTI") & " " & staro, False
    End If
    Exit Sub
EH:
    ShowToast Poruka("OTKUI_MSG_OPERATER_PAO"), True
End Sub

' Nov operater -- nova prava. Do v6-ui-199 se menjalo samo ime u zaglavlju, pa
' je admin mogao da otvori Korisnike, preda tastaturu operateru bez prava, a
' ekran i njegove mutacione komande su ostajali otvoreni.
'
' Redosled nije proizvoljan:
'   1) kes brane se BRISE pre svega -- inace bi se nova prava merila starim
'      odgovorom;
'   2) panel i editor se zatvaraju: oba nose stanje prethodnog operatera;
'   3) sidebar se precrtava da prigusenje odgovara novim pravima;
'   4) tek onda se proverava da li trenutni ekran sme da ostane.
'
' Ovo je UI brana. Tvrda brana stoji u piscu (modMaticniUnos), jer se do upisa
' stize i mimo ovog ekrana.
Private Sub PrimeniNovaPrava()
    On Error Resume Next
    modUiScreens.ScrResetCache
    ' False: ekran ispod se ionako ponovo cita nize (ili se menja), pa bi
    ' vracanje ovde bilo citanje viska.
    If modUiPanel.PanelAktivan() <> "" Then modUiPanel.PanelZatvori False
    If Len(mScreen) > 0 Then modUiScreens.ScrDeaktiviraj mScreen
    If mFrm Is Nothing Then Exit Sub

    ObnoviNavPrava
    If Len(mScreen) = 0 Then
        ' Prethodni operater nije imao pravo NI NA JEDAN ekran, pa je radna
        ' povrsina ostala prazna. Novi operater koji prava IMA mora da dobije
        ' ekran -- do v6-ui-204 je dobijao samo osvezen sidebar i ostajao pred
        ' praznom povrsinom do prvog rucnog klika.
        NaPrviDozvoljenEkran
        Err.Clear
        Exit Sub
    End If
    If Not modUiScreens.ScrDozvoljen(mScreen) Then
        ' Ekran koji novi operater ne sme da vidi ne ostaje otvoren.
        ShowToast Poruka("OTKUI_SCR_ZABRANJEN"), True
        NaPrviDozvoljenEkran
        Err.Clear
        Exit Sub
    End If
    PaintNav mFrm, NavTagFor(AktivnaStavka())
    LayoutOtkup mFrm
    RefreshFromData
    Err.Clear
End Sub

' Prigusenja sidebara po NOVIM pravima -- bez gradnje kontrola.
'
' Do v6-ui-201 je ovde stajalo BuildNav mFrm. Ono zove NewZone, a NewZone radi
' Controls.Add("Forms.Frame.1", "zNav") -- drugi put nad istom formom to je
' runtime greska "name already exists". Greska se gutala (On Error Resume Next
' u pozivaocu), pa se BuildNav prekidao na PRVOJ liniji: mNavOff je ostajao od
' PRETHODNOG operatera i sidebar je pokazivao tudja prava. Tvrde brane su
' drzale (ActivateScreen pita ScrDozvoljen na svaki klik), ali je meni lagao.
'
' Menjaju se samo PRAVA, a prava su jedino sto mNavOff drzi -- kontrole i
' preslikavanje stavka -> ekran (mNavKey) ostaju isti. Mapa se zato obilazi
' kroz mNavKey, ne kroz ponovljenu petlju po registru: dva spiska imena stavki
' bi mogla da se razidju, jedan ne moze.
Private Sub ObnoviNavPrava()
    Dim k As Variant
    If mNavKey Is Nothing Or mNavOff Is Nothing Then Exit Sub
    modUiScreens.ScrResetCache
    For Each k In mNavKey.keys
        mNavOff(CStr(k)) = Not modUiScreens.ScrAktivan(CStr(mNavKey(k)))
    Next k
End Sub

' Prvi ekran na koji novi operater SME. Postojalo je bezuslovno
' ActivateScreen SCR_POCETNI -- ali ako novi operater nema pravo ni na
' Dokumenta, taj prelazak se odbija i PRETHODNI (zabranjeni) ekran ostaje na
' sceni, sa podacima prethodnog operatera. Zato se trazi redom, a ako nema
' nijednog, radna povrsina se PRAZNI.
Private Sub NaPrviDozvoljenEkran()
    Dim r As Variant, kljuc As String
    If modUiScreens.ScrDozvoljen(SCR_POCETNI) Then
        ActivateScreen mFrm, SCR_POCETNI
        Exit Sub
    End If
    ' Redom kroz registar, PRESKACUCI panele: panel bi pokrio radnu povrsinu, a
    ' mScreen bi ispod njega ostao zabranjen ekran -- pa bi se on vratio cim se
    ' panel zatvori. Trazi se pravi ekran ili nijedan.
    For Each r In modUiScreens.ScrRows()
        kljuc = modUiScreens.ScrField(CStr(r), SCR_KLJUC)
        If Not modUiPanel.PanelPostoji(kljuc) Then
            If modUiScreens.ScrAktivan(kljuc) Then
                ActivateScreen mFrm, kljuc
                Exit Sub
            End If
        End If
    Next r
    ' Nijedan ekran nije dozvoljen. Prazna povrsina je jedini tacan prikaz --
    ' zadrzan ekran bi pokazivao podatke koje novi operater ne sme da vidi.
    mScreen = ""
    IsprazniRadnuPovrsinu mFrm
    ShowToast Poruka("OTKUI_SCR_NIJEDAN"), True
End Sub

' Radna povrsina bez ijednog dozvoljenog ekrana.
'
' Ne skriva samo mrezu -- BRISE joj stanje. Skrivena mreza sa tudjim redovima
' je i dalje tudji podaci: mView prezivi, RenderGrid ga vrati cim se bilo sta
' ponovo nacrta, a GridCell ga cita i bez crtanja. Zato prvo stanje, pa prikaz.
Private Sub IsprazniRadnuPovrsinu(frm As Object)
    On Error Resume Next
    mView = Empty
    mViewN = 0
    mSumKg = 0: mSumVal = 0
    mSelRow = 0
    mHoverRow = -1
    mPage = 1
    mFilter = "sve"
    mSearch = ""
    ClosePopup
    CloseFilterPanel
    ClearMarks
    RenderGrid                  ' iscrtaj PRAZNO preko starih redova
    ShowZones frm
    PaintNav frm, ""
    LayoutOtkup frm
    Err.Clear
End Sub

Private Sub RefreshOperater(frm As Object)
    On Error Resume Next
    frm.Controls("zHdr").Controls("btnOperaterC").caption = OperaterText()
    frm.Controls("zHdr").Controls("hdrStat").caption = HeaderStatusText()
    RefreshStatusBar frm
End Sub

' Prijavljeni app-korisnik; ako AUTH nije ukljucen - Windows nalog.
Private Function OperaterText() As String
    Dim s As String
    On Error Resume Next
    s = Trim$(GetCurrentUserIme())
    If Len(s) = 0 Then s = Trim$(Environ$("USERNAME"))
    If Len(s) = 0 Then s = Poruka("OTKUI_OPERATER_NEPOZNAT")
    OperaterText = s
End Function

' Dvoklik na red. Na unosnom ekranu ucitava red u formu -- kao i do sada. Na
' ugovornom ekranu ide EKRANU kao 'dbl:<red>', istim putem kao 'row:' i 'act:'.
'
' Zasto ekran ne moze da otvara detalj OBICNIM klikom: klik BIRA red, a radnje
' nad redom (zatvori paletu, storniraj, stampaj) rade bas nad izabranim. Da klik
' prebacuje listu, do tih radnji se ne bi moglo doci.
'
' Ljuska i dalje ne zna nijedan ekran po imenu -- zna samo svoj POCETNI, koji
' jedini ima formu za unos ispod mreze.
Private Sub UiDblClick(ByVal tag As String)
    RowFromTag tag
    If mScreen = SCR_POCETNI Then
        LoadRowIntoForm
        Exit Sub
    End If
    If mSelRow <= 0 Then Exit Sub
    If ScrAct("dbl:" & mSelRow) Then
        mSelRow = 0
        RefreshListSeg mFrm
        RefreshGridTitle mFrm
        ReloadGrid
        RefreshOtpTraka mFrm
        LayoutOtkup mFrm
    End If
End Sub

Private Sub UiChange(ByVal tag As String)
    ' KONTROLE UGOVORNOG EKRANA IDU EKRANU. Sve ispod ovog reda poznaje polja
    ' UNOSNOG ekrana po imenu (fgKgIT, cbKupac, fgBrZbirT...) -- to je i tacno,
    ' jer su ta polja njegova. Ali ekran koji ima SVOJA polja tu nema sta da
    ' trazi: ljuska ne zna sta ona znace, a ne sme ni da sazna.
    '
    ' Klik je ovaj put vec resio (grana 'scr' u UiClick); promena teksta nije
    ' imala svoju, pa ekran sa poljem za unos nije mogao ni da postoji. To je
    ' bas ono sto katalog vodi kao 'prosledjivanje dogadjaja sopstvenih
    ' kontrola ekranu' (Faza C, stavka 10).
    If Left$(tag, 3) = "scr" Then
        ' Kucanje suzava panel i ovde -- panel je ljuskin, isto kao nad poljima
        ' unosnog ekrana. Bez ovog reda bi kombo ekrana imao strelicu koja radi,
        ' a kucanje koje ne otvara nista.
        If Not mPopMute And Not mBuilding Then PopFromTyping tag
        ScrAct "chg:" & tag
        Exit Sub
    End If
    ' Promene koje pravi FillZbirneCombo (Clear + vracanje teksta) nisu unos
    ' operatera - ne diraju ni zapamcenu zbirnu ni "dokument je menjan".
    If mZbirnaFill And tag = "fgBrZbirT" Then Exit Sub
    If mPartnerFill And tag = "cbKupac" Then Exit Sub
    ' Kucanje u combo otvara nas panel i suzava listu. Programski upisi
    ' (punjenje lista, ClearForm, izbor iz panela) su pod mPopMute - bez toga
    ' bi se panel otvarao sam od sebe pri svakoj promeni rezima.
    If Not mPopMute And Not mBuilding Then PopFromTyping tag
    ' pretraga nije izmena dokumenta
    ' Pretraga i opseg datuma za stampu nisu izmene dokumenta.
    If tag <> "txtSearch" And Left$(tag, 4) <> "spec" Then MarkDirty
    Select Case tag
        Case "txtSearch"
            mSearch = Trim$(CStr(mFrm.Controls("zGrid").Controls("txtSearch").text))
            mFrm.Controls("zGrid").Controls("lblSearchPh").Visible = (Len(mSearch) = 0)
            mSelRow = 0
            ReloadGrid
        Case "specOdT", "specDoT"
            ' Opseg datuma suzava listu odmah, kao i pretraga. Nepotpun datum
            ' ("2", "21.") nije greska nego "jos nema granice" - ekran ga
            ' preskace, pa se lista ne prazni dok operater kuca.
            mSelRow = 0
            ReloadGrid
        Case "fgKgIT", "fgKgIIT", "fgCena1T", "fgCena2T"
            RecalcVrednost
        ' u bruto rezimu neto zavisi i od ambalaze - tara se racuna iz nje
        Case "fgKolAmbT", "fgTipAmbT"
            RecalcVrednost
        Case "fgBrZbirT"
            SetAktivnaZbirna mFrm.Controls("zForm").Controls("fgBrZbir").Controls("fgBrZbirT").text
        Case "fgBrOtprT"
            If ActiveMode = "F3" Then _
                SetAktivnaZbirna mFrm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").text
        Case "cbKupac"
            FillOpenBlokovi mFrm
            FillOpenFakture mFrm
            FillParcele mFrm
            ' Broj prijemnice zavisi od KUPCA. Legacy cmbKupac_Change prvo
            ' OBRISE broj pa trazi predlog - bez brisanja bi kod ne-hladnjaca
            ' kupca (koji predlog ne dobija) ostao broj hladnjace.
            If modeKey(ActiveMode) = "PRIJEMNICA" Then
                SetFld "fgBrOtpr", ""
                RefreshBrojPredlog True
            End If
        Case "fgParcelaT"
            ApplyParcelaProizvod
        Case "fgBlokT"
            ApplyBlokIzbor
        Case "cbVrsta"
            RefillSorta mFrm
            AutoFillCena
        Case "cbSorta"
            AutoFillCena
        Case "cbOM"
            OnStanicaChanged
            RefreshSaldoOM              ' avans se vodi po otkupnom mestu
            ' Lista otvorenih blokova je suzena na OM, pa promena OM menja i nju.
            ' Ne oslanjamo se na to sto OnStanicaChanged usput brise partnera:
            ' zavisnost od tudjeg sporednog efekta je tacno ono sto tiho zataji.
            FillOpenBlokovi mFrm
            RefreshKpi mFrm
            RefreshStatusBar mFrm
        Case "cbVozac"
            ' Zbirna se broji po vozacu, pa promena vozaca menja i njen niz.
            ' Bez vozaca niz ne postoji - broj se brise (legacy cmbVozac_Change).
            If modeKey(ActiveMode) = "ZBIRNA" Then
                If Len(Trim$(CStr(mFrm.Controls("zCtx").Controls("cbVozac").value))) = 0 Then
                    SetFld "fgBrOtpr", ""
                Else
                    RefreshBrojPredlog True
                End If
            End If
        Case "fgDatumT"
            OnDatumChanged
    End Select
End Sub

Private Sub UiHover(ByVal tag As String)
    Dim i As Long
    ' Redovi naseg dropdown-a (v6-ui-188): tekst i pozadina reda su
    ' odvojene kontrole, pa se hover boji CENTRALNO -- kao red mreze.
    If Left$(tag, 3) = "pop" And IsNumeric(Mid$(tag, 4)) Then
        PopHoverRed CLng(Mid$(tag, 4))
        Exit Sub
    End If
    ' zaglavlje kolone ("hd7") nosi samo strelicu sortiranja
    If Left$(tag, 2) = "hd" And IsNumeric(Mid$(tag, 3)) Then
        HoverHead CLng(Mid$(tag, 3))
        Exit Sub
    End If
    i = RowIndexFromTag(tag)
    If i = mHoverRow Then Exit Sub
    HoverHead -1                        ' pokazivac je napustio zaglavlje
    mHoverRow = i
    RenderGrid
End Sub

' "rb7" ili "c7_3" -> mSelRow (1-bazirano nad filtriranim skupom)
Private Sub RowFromTag(ByVal tag As String)
    Dim i As Long
    i = RowIndexFromTag(tag)
    If i < 0 Then Exit Sub
    mSelRow = (mPage - 1) * mPageSize + i + 1
    If mSelRow > mViewN Then mSelRow = 0
    ' Izbor reda se javlja ekranu: u listi otpremnica klik BIRA otpremnicu, a
    ' ljuska ne zna sta je otpremnica. Ako je ekran nesto uradio, mreza se
    ' puni ponovo.
    If mSelRow > 0 Then
        If modUiScreens.ScrEvent(mScreen, "row:" & mSelRow, "Click") Then
            mSelRow = 0
            RefreshListSeg mFrm
            RefreshGridTitle mFrm
            ReloadGrid
            RefreshOtpTraka mFrm
            LayoutOtkup mFrm
        End If
    End If
End Sub

Private Function RowIndexFromTag(ByVal tag As String) As Long
    RowIndexFromTag = -1
    On Error Resume Next
    If Left$(tag, 2) = "rb" Then
        RowIndexFromTag = CLng(Mid$(tag, 3))
    ElseIf Left$(tag, 1) = "c" And InStr(tag, "_") > 2 Then
        RowIndexFromTag = CLng(Mid$(tag, 2, InStr(tag, "_") - 2))
    End If
End Function

' focus/error ivica: tag je npr. "fgKgIT" -> shell se zove "fgKgI"
' Zeleni okvir dobija polje koje je upravo dobilo kursor, a gubi ga ono koje
' ga je imalo - bez obzira da li je za njega stigao dogadjaj napustanja.
Private Sub PostaviFokus(ByVal tag As String)
    If Len(mFokusTag) > 0 And mFokusTag <> tag Then ShellFor mFokusTag, "normal"
    mFokusTag = tag
    ShellFor tag, "focus"
End Sub

Private Sub SkiniFokus(ByVal tag As String)
    ShellFor tag, "normal"
    If mFokusTag = tag Then mFokusTag = ""
End Sub

' Sva polja u "normalno" stanje - pri promeni rezima i pri praznjenju forme,
' da zeleni okvir ne ostane na polju koje je u medjuvremenu sakriveno.
Private Sub ResetFokusOkvira()
    Dim c As Object, z As Object, nm As String
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    mFokusTag = ""
    Set z = mFrm.Controls("zForm")
    For Each c In z.Controls
        If TypeName(c) = "Frame" Then
            nm = c.name
            If HasCtl(c, nm & "B") Then ShellState c, nm, "normal"
        End If
    Next c
    Set z = mFrm.Controls("zCtx")
    For Each c In z.Controls
        If Left$(c.name, 2) = "cb" Then ShellState z, c.name, "normal"
    Next c
End Sub

Private Sub ShellFor(ByVal tag As String, ByVal state As String)
    Dim nm As String, fr As Object
    On Error Resume Next
    If tag = "txtSearch" Then
        ShellState mFrm.Controls("zGrid"), "srch", state
        Exit Sub
    End If
    If Left$(tag, 2) = "cb" Then
        ShellState mFrm.Controls("zCtx"), tag, state
        Exit Sub
    End If
    If Right$(tag, 1) <> "T" Then Exit Sub
    nm = Left$(tag, Len(tag) - 1)
    Set fr = mFrm.Controls("zForm").Controls(nm)
    If fr Is Nothing Then Exit Sub
    If mKlasa <> 2 And (nm = "fgKgII" Or nm = "fgKolAmbII") Then Exit Sub
    ShellState fr, nm, state
End Sub

' zadrzano: nativna lista se vise NE koristi - OpenPopup pokriva i duge liste
Private Sub DropCombo(ByVal nm As String)
    Dim c As Object
    On Error Resume Next
    Set c = FindCombo(nm)
    If c Is Nothing Then Exit Sub
    c.SetFocus
    c.DropDown
End Sub

Private Sub SetKlasa(ByVal k As Long)
    Dim fr As Object, t As Object
    On Error Resume Next
    mKlasa = k
    MarkDirty
    Set fr = mFrm.Controls("zForm").Controls("fgCena")
    BoxState fr, "segKlasa1", IIf(k = 1, C_FOREST, C_WHITE), IIf(k = 1, C_CREAM, C_MUTED), (k = 1)
    BoxState fr, "segKlasa2", IIf(k = 2, C_FOREST, C_WHITE), IIf(k = 2, C_CREAM, C_MUTED), (k = 2)

    Set fr = mFrm.Controls("zForm").Controls("fgKgII")
    Set t = fr.Controls("fgKgIIT")
    t.enabled = (k = 2)
    t.BackColor = IIf(k = 2, C_WHITE, C_DISABLED)
    t.ForeColor = IIf(k = 2, C_FOREST, C_DISABLED_FG)
    ShellState fr, "fgKgII", IIf(k = 2, "normal", "off")
    fr.Controls("fgKgIIL").ForeColor = IIf(k = 2, C_MUTED, C_DISABLED_FG)
    fr.Controls("fgKgIIU").ForeColor = IIf(k = 2, C_MUTED, C_DISABLED_FG)
    fr.Controls("fgKgIIUB").BackColor = IIf(k = 2, C_SAND, C_DISABLED)
    ' povratak na I klasu brise unetu kolicinu II klase: ostala bi nevidljiva
    ' u onemogucenom polju, a RecalcVrednost i CommitDokument je ne racunaju
    If k <> 2 Then t.text = ""

    ' druga kutija cene deli sudbinu kilograma II klase
    Set fr = mFrm.Controls("zForm").Controls("fgCena")
    Set t = fr.Controls("fgCena2T")
    t.enabled = (k = 2)
    t.BackColor = IIf(k = 2, C_WHITE, C_DISABLED)
    t.ForeColor = IIf(k = 2, C_FOREST, C_DISABLED_FG)
    ShellState fr, "fgCena2", IIf(k = 2, "normal", "off")
    fr.Controls("fgCena2P").ForeColor = IIf(k = 2, C_MUTED, C_DISABLED_FG)
    If k <> 2 Then t.text = ""

    ' Ukljucivanje II klase povlaci i njenu cenu iz cenovnika - do sada je
    ' kutija ostajala prazna dok je operater ne otkuca, iako cena za tu robu i
    ' klasu vec postoji.
    If k = 2 Then AutoFillCena
    ' gajbe II klase postoje samo dok II klasa postoji
    FldShow mFrm.Controls("zForm"), "fgKolAmbII", _
            (k = 2 And mFrm.Controls("zForm").Controls("fgKolAmb").Visible)
    If k <> 2 Then mFrm.Controls("zForm").Controls("fgKolAmbII").Controls("fgKolAmbIIT").text = ""
    If Not mBuilding Then LayoutOtkup mFrm
    RefreshPaletaInfo

    If k = 2 Then mFrm.Controls("zForm").Controls("fgKgII").Controls("fgKgIIT").SetFocus
    RecalcVrednost
End Sub

' cip se sakriva zajedno sa svojom ivicom
' ime | kljuc natpisa | sirina
Private Function ChipRow() As Variant
    ' Poslednji cip pripada ISKLJUCIVO listi otpremnica - u listi dokumenata se
    ' ne prikazuje. Stoji na kraju niza da indeks 5 (Otkazane) zadrzi svoj
    ' poseban rez u BuildGrid.
    ChipRow = Array("chipDanas|OTKUI_CHIP_DANAS|54", "chipNedelja|OTKUI_CHIP_NEDELJA|78", _
                    "chipSve|OTKUI_CHIP_SVE|40", "chipNefakt|OTKUI_CHIP_NEFAKT|98", _
                    "chipBezZbirne|OTKUI_CHIP_BEZZBIRNE|86", "chipOtkazane|OTKUI_CHIP_OTKAZANE|82", _
                    "chipOtvorene|OTKUI_CHIP_OTVORENE|132")
End Function

Private Sub ShowChip(frm As Object, ByVal nm As String, ByVal vis As Boolean)
    BoxShow frm.Controls("zGrid"), nm, vis
End Sub

' Sakriven cip ostavlja rupu u redu - polozaji su postavljeni pri izgradnji.
' Posle svake promene vidljivosti red se prepakuje sleva.
Private Sub LayoutChips(frm As Object)
    Dim z As Object, chip As Variant, p As Variant, i As Long, X As Single
    On Error Resume Next
    Set z = frm.Controls("zGrid")
    chip = ChipRow()
    X = PAD
    For i = 0 To UBound(chip)
        p = Split(chip(i), "|")
        ' Sirina dolazi iz mChipW, ne iz opisa: slot koji je pozajmio ekran
        ' nosi sirinu svog natpisa, pa bi opis vratio pogresnu meru.
        If mChipW(i) <= 0 Then mChipW(i) = CSng(p(2))
        If z.Controls(CStr(p(0))).Visible Then
            MoveBox z, CStr(p(0)), X, 36, mChipW(i)
            X = X + mChipW(i) + 6
        End If
    Next i
End Sub

' polje: label + shell + kontrola uvucena za INPUT_PAD (+ jedinica / strelica)
' Polje unosa (naslov + okvir + TextBox ili ComboBox + jedinica). JAVNO je da bi
' ga ugovorni ekran mogao koristiti u svojoj zoni: bez toga bi svaki ekran sa
' unosom crtao svoju verziju istog polja, pa bi se razisli u izgledu i ponasanju.
'
' Kontrole se OZICAVAJU ovde (WireInput), pa promena stize kroz UiChange. Ekran
' im daje imena sa prefiksom 'scr' da bi ta promena dosla NJEMU, a ne ljusci.
Public Sub NewFieldG(parent As Object, nm As String, cap As String, kind As String, _
                      unit As String, span As Long, isNum As Boolean, isFocus As Boolean, _
                      ByVal grp As String)
    Dim fr As Object
    Set fr = NewFrame(parent, nm, 0, 0, 180, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:" & span & ":" & grp
    NewLbl fr, nm & "L", UCase$(cap), 0, 0, 170, TxtH(TS_LABEL), TS_LABEL, True, C_MUTED, -1
    NewShell fr, nm, 0, 16, 180, FIELD_H, C_INPUT_BORDER, C_WHITE
    If kind = "cmb" Then
        NewCombo fr, nm & "T", "", INPUT_PAD, 20, 140, FIELD_H - 8
        NewChevron fr, nm & "D", 161, 17, 18, FIELD_H - 2
    Else
        NewTxt fr, nm & "T", "", INPUT_PAD, 20, 160, FIELD_H - 8, isNum
    End If
    If Len(unit) > 0 Then
        NewLbl fr, nm & "UB", "", 0, 17, 38, FIELD_H - 2, 8, False, 0, C_SAND
        NewLbl fr, nm & "U", unit, 0, CenterY(17, FIELD_H - 2, TS_MICRO), 38, TxtH(TS_MICRO), _
               TS_MICRO, False, C_DISABLED_FG, -1, fmTextAlignCenter
        fr.Controls(nm & "U").ZOrder 0
    End If
    ' isFocus vise NE boji polje pri gradnji. Zeleni okvir je stanje kursora,
    ' pa ga postavlja dogadjaj fokusa; ranije je ostajao zauvek na prvom polju i
    ' pokazivao fokus tamo gde ga nema.
    If isFocus Then mPrvoPolje = nm
End Sub

' Polovina polja CENA: kutija sa oznakom klase levo i brojem desno.
' Sirinu preuzima LayoutFieldInner - ovo su samo pocetne mere.
Private Sub NewCenaBox(fr As Object, nm As String, ByVal klasa As String, _
                       ByVal X As Single, ByVal w As Single)
    NewShell fr, nm, X, 16, w, FIELD_H, C_INPUT_BORDER, C_WHITE
    NewLbl fr, nm & "P", klasa, X + 5, CenterY(17, FIELD_H - 2, TS_MICRO), 16, TxtH(TS_MICRO), _
           TS_MICRO, True, C_MUTED, -1, fmTextAlignLeft
    fr.Controls(nm & "P").ZOrder 0
    NewTxt fr, nm & "T", "", X + 22, 20, w - 26, FIELD_H - 8, True
End Sub

' kontekst combo: shell + combo + strelica koja maskira nativno dugme
Private Sub NewCtxCombo(parent As Object, nm As String, X As Single, Y As Single, w As Single)
    NewShell parent, nm, X, Y, w, 26, C_INPUT_BORDER, C_WHITE
    NewCombo parent, nm, "", X + INPUT_PAD, Y + 4, w - INPUT_PAD - 22, 18
    NewChevron parent, nm & "D", X + w - 19, Y + 1, 18, 24
End Sub

#If VBA7 Then
Private Function FormHwnd(frm As Object) As LongPtr
    Dim h As LongPtr
#Else
Private Function FormHwnd(frm As Object) As Long
    Dim h As Long
#End If
    Dim oldCap As String, tag As String
    On Error Resume Next
    oldCap = frm.caption
    tag = "AgriX#" & Format$(Timer * 1000, "0")
    frm.caption = tag
    DoEvents
    h = FindWindow("ThunderDFrame", tag)              ' modeless UserForm
    If h = 0 Then h = FindWindow("ThunderXFrame", tag) ' modal
    frm.caption = oldCap
    FormHwnd = h
End Function

Public Sub GoFullScreen(frm As Object)
    Const TASKBAR_H As Single = 30
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim style As Long, wPt As Single, hPt As Single
    On Error Resume Next

    wPt = ScreenWidthPoints()
    hPt = ScreenHeightPoints() - TASKBAR_H
    If wPt < MIN_W Then wPt = MIN_W
    If hPt < MIN_H Then hPt = MIN_H

    frm.StartUpPosition = 0
    frm.Left = 0
    frm.top = 0
    frm.width = wPt
    frm.Height = hPt

    If mChromeRemoved Then Exit Sub
    h = FormHwnd(frm)
    If h = 0 Then Exit Sub            ' nismo nasli SVOJ prozor -> ne diraj nijedan
    style = GetWindowLong(h, GWL_STYLE)
    SetWindowLong h, GWL_STYLE, (style And Not WS_CAPTION) Or WS_THICKFRAME
    DrawMenuBar h
    frm.caption = vbNullString
    mChromeRemoved = True
End Sub

Public Sub MakeResizable()
    ' namerno prazno: raniji GetActiveWindow pristup je skidao/menjao stil
    ' POGRESNOM prozoru. Sve ide kroz GoFullScreen (FormHwnd).
End Sub

Public Sub FixVbeCaption()
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim st As Long, i As Long, cls As Variant
    cls = Array("wndclass_desked_gsk", "XLMAIN")
    For i = 0 To UBound(cls)
        h = FindWindow(CStr(cls(i)), vbNullString)
        If h <> 0 Then
            st = GetWindowLong(h, GWL_STYLE)
            SetWindowLong h, GWL_STYLE, st Or WS_CAPTION
            DrawMenuBar h
        End If
    Next i
End Sub

Public Sub ShowOtkupUI()
    ' Javan makro se moze pozvati i direktno (Alt+F8), mimo navigacije - zato se
    ' pravo na POCETNI ekran proverava ovde. Bez ovoga bi se ekran sa pravom
    ' upisa otvarao i korisniku kome ActivateScreen taj ekran ne bi dao.
    If Not modUiScreens.ScrDozvoljen(SCR_POCETNI) Then
        MsgBox Poruka("OTKUI_SCR_ZABRANJEN"), vbExclamation, APP_NAME
        Exit Sub
    End If
    frmOtkupUI.show vbModeless
End Sub

' Zatvaranje ekrana (dugme X u zaglavlju ili sistemski X). Forma se SAKRIVA, ne
' unload-uje, pa UserForm_Terminate ne puca i lock stanice bi ostao zauzet na
' serveru do gasenja Excela. Legacy frmOtkup pusta stanicu i u btnPovratak_Click
' i u UserForm_QueryClose - isto radi i ovaj ekran.
Public Sub OtkupUI_Sakrij()
    On Error Resume Next
    If Len(GetActiveStanica()) > 0 Then ReleaseStanicaLock GetActiveStanica()
    If Not mFrm Is Nothing Then mFrm.Hide
End Sub

' Gradnja ekrana je pukla. Zastavica mBuilding mora da padne, inace bi svaki
' sledeci dogadjaj cutke izlazio (svi handleri je proveravaju), pa bi ekran
' izgledao "ziv" a ne bi reagovao ni na sta.
Public Sub OtkupUI_BuildFailed(ByVal errNum As Long, ByVal errDesc As String)
    On Error Resume Next
    mBuilding = False
    LogError "modOtkupUI.BuildOtkupScreen", CStr(errNum) & " " & errDesc
End Sub

Private Sub ShowZones(frm As Object)
    Dim nmv As Variant, i As Long, dok As Boolean, prazno As Boolean
    dok = (mScreen = "DOKUMENTI")
    ' PRAZAN mScreen znaci "nijedan ekran nije dozvoljen" (NaPrviDozvoljenEkran).
    ' Naslov i mreza se tada NE crtaju. Do v6-ui-203 su se crtali, sa sadrzajem
    ' PRETHODNOG operatera -- mreza je zadrzavala njegove redove jer je
    ' LoadGridFromScreen za nepoznat kljuc izlazio pre nego sto isprazni mView.
    ' To nije kozmetika nego curenje podataka izmedju naloga.
    prazno = (Len(mScreen) = 0)
    nmv = Array("zHdr", "zNav", "zStatus")
    For i = 0 To UBound(nmv)
        On Error Resume Next
        frm.Controls(CStr(nmv(i))).Visible = True
    Next i
    ' Naslovna traka i mreza su ZAJEDNICKE - nosi ih i ugovorni ekran. Da
    ' nisu, svaki ekran bi izmislio svoj naslov i svoju listu, a lista je bas
    ' ono sto se deli (strane, sortiranje, pretraga, izbor).
    ' Dok je radna povrsina USTUPLJENA panelu, naslov i mreza se ne crtaju:
    ' panel pokriva isti prostor, a dve stvari u njemu su ista klasa kvara kao
    ' editor i geo panel jedan preko drugog.
    nmv = Array("zTitle", "zGrid")
    For i = 0 To UBound(nmv)
        On Error Resume Next
        frm.Controls(CStr(nmv(i))).Visible = (Not mPanelRezim) And (Not prazno)
    Next i
    On Error Resume Next
    frm.Controls("zPanel").Visible = mPanelRezim
    ' Ovo je samo ekran dokumenata: KPI traka, kontekstni red, forma, kartice
    ' i traka aktivne otpremnice.
    '
    ' zOtp je iz ovog spiska bio ISPAO. Njegovu vidljivost postavlja samo
    ' LayoutAll (grana ekrana dokumenata), pa je na ugovornim ekranima ostajao
    ' onakav kakav ga je Dokumenta ostavila -- a vidi se ili ne vidi zavisno od
    ' toga da li ga zona tog ekrana slucajno pokriva. Na ekranu Uvoz izvoda
    ' (najniza zona) se video, sa porukom "Nema izabrane otpremnice", koja tamo
    ' ne znaci nista.
    nmv = Array("zKpi", "zCtx", "zForm", "zRight", "zOtp")
    For i = 0 To UBound(nmv)
        On Error Resume Next
        frm.Controls(CStr(nmv(i))).Visible = dok
    Next i
    For Each nmv In frm.Controls
        On Error Resume Next
        If Left$(nmv.name, 5) = "zScr_" Then
            nmv.Visible = (Not dok) And (Not mPanelRezim) And _
                          (nmv.name = "zScr_" & mScreen)
        End If
    Next nmv
End Sub

' Prelazak na drugi ekran. VRACA da li smo posle poziva TAMO GDE SMO HTELI.
'
' False znaci da prelazak nije izvrsen: nema prava, nema modula, gradnja je
' pala, ili je operater odustao od odbacivanja nesacuvanog. Pozivalac koji je
' UZ prelazak promenio jos nesto (sekcija, zlatno dugme, sidebar) mora to da
' vrati -- inace zaglavlje pokazuje jednu sekciju, a ekran je iz druge.
'
' Zona ekrana se gradi LENJO - tek pri prvom ulasku -
' i ostaje izgradjena. Merenje na ekranu dokumenata: 863 kontrole, od toga 623
' deljeni hrom i mreza, pa je ekran u proseku 240 kontrola i oko 180 ms; drzati
' ih izgradjene je jeftinije nego ruseti ih pri svakom izlasku.
Private Function ActivateScreen(frm As Object, ByVal kljuc As String) As Boolean
    Dim z As Object, nm As String, greska As String, vratiEkran As Boolean
    Dim jePanel As Boolean, zatvorenPanel As Boolean

    ' Ponovljen klik na OTVOREN panel ne radi nista. Bez ovoga bi se panel
    ' rusio i gradio ispocetka, pa bi Podesavanja izgubila nesacuvanu izmenu
    ' na klik u stavku na kojoj operater vec jeste.
    If Len(kljuc) > 0 Then
        If modUiPanel.PanelAktivan() = UCase$(kljuc) Then
            ActivateScreen = True       ' vec smo tu
            Exit Function
        End If
    End If
    jePanel = modUiPanel.PanelPostoji(kljuc)

    ' Klik na stavku ekrana na kome VEC jesmo, bez otvorenog panela: nista se
    ' ne menja, pa se nista i ne dira.
    If (Not jePanel) And kljuc = mScreen And modUiPanel.PanelAktivan() = "" Then
        ActivateScreen = True           ' vec smo tu
        Exit Function
    End If

    ' ---- CILJ SE PROVERAVA PRVI ----------------------------------------
    ' Pre nego sto se ista ZATVORI. Do v6-ui-203 se panel zatvarao na vrhu, pa
    ' je klik na stavku ciji modul nedostaje (nepotpun import, self-update u
    ' toku) ostavljao i panel zatvoren i ekran ispod njega deaktiviran -- za
    ' potez koji nije nikuda ni odveo. Sada odbijen prelazak ne dira nista.
    '
    ' Pravo se proverava za SVAKI ekran, i za pocetni: ranije je "DOKUMENTI"
    ' bio izuzet, pa se na njega moglo vratiti i bez prava.
    If Not modUiScreens.ScrDozvoljen(kljuc) Then
        ShowToast Poruka("OTKUI_SCR_ZABRANJEN"), True
        PaintNav frm, NavTagFor(AktivnaStavka())
        Exit Function
    End If
    If (Not jePanel) And kljuc <> SCR_POCETNI Then
        ' Neuspeh mora da VRATI oznaku u sidebaru na ekran na kome zaista
        ' jesmo - inace bi meni pokazivao Palete dok je na ekranu jos Unos.
        If Not modUiScreens.ScrPostoji(kljuc) Then
            ShowToast Poruka("OTKUI_SCR_NEMA"), False
            PaintNav frm, NavTagFor(AktivnaStavka())
            Exit Function
        End If
    End If

    ' ---- NESACUVANO SE NE ODBACUJE TIHO --------------------------------
    ' Pita se za ekran koji ODLAZI, i to samo kad panel NIJE otvoren: dok je
    ' panel otvoren, ekran ispod je vec deaktiviran (editor mu je zatvoren pri
    ' otvaranju panela), pa nema sta da izgubi.
    If modUiPanel.PanelAktivan() = "" And Len(mScreen) > 0 Then
        If modUiScreens.ScrImaNesacuvano(mScreen) Then
            If MsgBox(Poruka("MATU_ASK_ODBACI_UNOS"), _
                      vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then
                PaintNav frm, NavTagFor(AktivnaStavka())
                Exit Function
            End If
        End If
    End If

    ' ---- TEK SADA se zatvara ono sto je otvoreno -----------------------
    ' Panel drzi CELU radnu povrsinu. Klik u sidebaru bi inace promenio ekran
    ' ispod njega -- operater bi video isti panel, a ljuska bi mislila da je na
    ' drugom ekranu.
    '
    ' Ekran ispod se vraca SAMO ako ostajemo na njemu. Kad se prelazi dalje,
    ' vraca ga sam prelazak -- pa bi vracanje ovde bilo citanje viska.
    If modUiPanel.PanelAktivan() <> "" Then
        ' Nesacuvano u panelu se NE odbacuje tiho. Ako operater odustane,
        ' oznaka u sidebaru se vraca na panel -- inace bi meni pokazivao ekran
        ' na koji se nije preslo.
        If Not modUiPanel.PanelSmemoDaZatvorimo() Then
            PaintNav frm, NavTagFor(AktivnaStavka())
            Exit Function
        End If
        vratiEkran = (Not jePanel) And (kljuc = mScreen)
        modUiPanel.PanelZatvori vratiEkran
        zatvorenPanel = True
    End If

    ' Stavka sidebara koja je PANEL, ne ekran. Grana je odvojena jer panel nema
    ' nista od onoga sto sledi: ni zonu, ni naslovnu traku, ni mrezu, ni sort.
    ' mScreen zato ostaje ekran ISPOD panela -- zatvaranje panela ga vraca bez
    ' ponovne gradnje, a "Nazad na rad" i dalje zna odakle se doslo.
    '
    ' Do v6-ui-200 se ovde nije dolazilo: panel je otvarao ekran MAT_SISTEM
    ' preko svog dugmeta "Otvori alatku". Taj ekran je bio spisak od dve stavke
    ' -- tacno ono sto sidebar vec radi -- pa je uklonjen, a panel je postao
    ' stavka sidebara kao i Korisnici.
    If jePanel Then
        ' Ekran koji ostaje ispod panela svejedno sprema svoje: editor koji bi
        ' prezivo iza panela pise u sifarnik koji se vise ne vidi.
        If Len(mScreen) > 0 Then modUiScreens.ScrDeaktiviraj mScreen
        ClosePopup
        CloseFilterPanel
        greska = modUiPanel.PanelOtvori(kljuc)
        If Len(greska) > 0 Then
            ShowToast greska, True
            PanelVracenNaEkran
            PaintNav frm, NavTagFor(AktivnaStavka())
            Exit Function
        End If
        ActivateScreen = True
        Exit Function
    End If
    ' Panel je zatvoren i ekran je vec vracen kad ostajemo na njemu.
    If kljuc = mScreen Then
        ActivateScreen = True
        Exit Function
    End If
    If kljuc <> SCR_POCETNI Then
        nm = "zScr_" & kljuc
        On Error Resume Next
        Set z = frm.Controls(nm)
        On Error GoTo 0
        If z Is Nothing Then
            Set z = NewZone(frm, nm, SIDEBAR_W, HEADER_H, 400, 400, C_CREAM)
            WireZone z
            If Not modUiScreens.ScrBuild(kljuc, z) Then
                ' Ekran koji padne u GRADNJI ne sme da obori aplikaciju. Ovo je
                ' jedini bail posle zatvaranja panela, pa se ekran ispod ovde
                ' vraca rucno -- inace bi ostao deaktiviran.
                z.Visible = False
                ShowToast Poruka("OTKUI_SCR_NEMA"), True
                If zatvorenPanel Then PanelVracenNaEkran
                PaintNav frm, NavTagFor(AktivnaStavka())
                Exit Function
            End If
        End If
    End If
    ' Ekran koji odlazi sprema svoje: editor, panel, izbor. Ugovorno i kasno
    ' vezano -- ljuska i dalje ne zna nijedan ekran po imenu.
    '
    ' TEK SADA, kad su sve provere prosle. Do v6-ui-201 je stajalo na vrhu, pa
    ' je odbijen prelazak (nema prava, nema modula, gradnja pala) ostavljao
    ' TRENUTNI ekran deaktiviran a i dalje na sceni: mreza se vidi, a radnje
    ' tise ne rade jer im je zona obrisana.
    If Len(mScreen) > 0 Then modUiScreens.ScrDeaktiviraj mScreen
    mScreen = kljuc
    ClosePopup
    CloseFilterPanel
    ClearMarks
    mFilter = "sve"                  ' ugovorni ekran nema cipove
    mSearch = ""
    ' Kolona sortiranja se NE prenosi izmedju ekrana: 9. kolona na Paletama i
    ' 9. na otkupnom listu nisu isti podatak. Ali ni tvrdi reset na 2/desc
    ' nije tacan za svaku listu -- rang-lista se otvara po rangu rastuce
    ' (recenzija #245, krug 17: povratak na Izvestaje sa aktivnim Rangom je
    ' vracao sort po imenu). Aktivacija zato primenjuje PODRAZUMEVANI sort
    ' aktivne liste ekrana -- isti SortZaListu ugovor kao klik i auto-prelaz.
    PrimeniSortZaListu ActiveLista()
    mSelRow = 0
    mHoverRow = -1
    mPage = 1
    ShowZones frm
    If kljuc = "DOKUMENTI" Then
        SelectModeCore frm, ActiveMode, True
    Else
        RefreshTitleFor frm, kljuc
        LayoutOtkup frm
        ReloadGrid
    End If
    LayoutOtkup frm
    ActivateScreen = True
End Function

' Ceo prostor desno od sidebara pripada ugovornom ekranu. Ljuska mu daje
' pravougaonik i pita ga da se rasporedi - sta u njemu stoji ne zna.
' Raspored ugovornog ekrana. Redosled je isti kao na ekranu dokumenata i to
' je namerno: zona ekrana stoji GORE, tamo gde je KPI traka, pa naslov, pa
' mreza. Prvo je bilo obrnuto (naslov pa brojke) i odmah se videlo da dva
' ekrana iste aplikacije imaju razlicit ritam.
Private Sub LayoutScreenZone(frm As Object, ByVal X As Single, ByVal w As Single, _
                             ByVal wTot As Single, ByVal hTot As Single)
    Dim z As Object, h As Single, gy As Single
    On Error Resume Next
    ' Panel dobija CELU radnu povrsinu -- sidebar, zaglavlje i statusna traka
    ' ostaju, sve ostalo je njegovo. Zona ekrana se tada ne meri: sakrivena je,
    ' a Scr_Layout nad sakrivenom zonom trosi na raspored koji niko ne vidi.
    If mPanelRezim Then
        With frm.Controls("zPanel")
            .Left = X: .top = HEADER_H: .width = w
            .Height = hTot - HEADER_H - STATUS_H
        End With
        Exit Sub
    End If

    Set z = frm.Controls("zScr_" & mScreen)
    If z Is Nothing Then Exit Sub
    z.Left = X
    z.top = HEADER_H
    z.width = w
    z.Height = 10
    h = modUiScreens.ScrLayout(mScreen, z, w, hTot - HEADER_H - TITLE_H - STATUS_H)
    If h < 1 Then h = 1
    z.Height = h
    With frm.Controls("zTitle")
        .Left = X: .top = HEADER_H + h: .width = w: .Height = TITLE_H
        .Controls("titLnB").width = w
        .Controls("titDatum").Left = w - PAD - 190
        .Controls("titName").width = w - PAD * 2 - 200 - TIT_ICO_W - 10
        .Controls("titSub").width = w - PAD * 2 - 200 - TIT_ICO_W - 10
    End With
    gy = HEADER_H + h + TITLE_H
    With frm.Controls("zGrid")
        .Left = X: .top = gy: .width = w
        .Height = hTot - gy - STATUS_H
        If .Height < GRID_TOP + GRID_HEAD_H + GRID_FOOT_H + 3 * GRID_ROW_H + 6 Then
            .Height = GRID_TOP + GRID_HEAD_H + GRID_FOOT_H + 3 * GRID_ROW_H + 6
        End If
    End With
    LayoutGrid frm.Controls("zGrid"), w, frm.Controls("zGrid").Height
End Sub

' Zona ugovornog ekrana, da ekranski modul moze da napuni svoje kontrole a da
' ne pamti referencu koja preziveti rusenje forme ne bi.
' Zona ugovornog ekrana, ili Nothing ako je nema.
'
' Err se BRISE pred povratak. Zona koje nema je odgovor 'Nothing', ne greska --
' a On Error Resume Next je samo PRESKACE, ne cisti. Bez brisanja Err ostaje
' postavljen i posle povratka, pa ga ScrGridData procita kao pad ekrana: mreza
' se isprazni uz 'Lista se nije ucitala' iako su podaci procitani ispravno.
' Isti put pogadja svaki ekran koji u Scr_Rows dopunjava svoju zonu.
Public Function ScreenZone(ByVal kljuc As String) As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Function
    Set ScreenZone = mFrm.Controls("zScr_" & kljuc)
    If Err.Number <> 0 Then Err.Clear
End Function

' Naslovna traka za ugovorni ekran - iz Scr_Meta, ne iz rezima dokumenata.
Private Sub RefreshTitleFor(frm As Object, ByVal kljuc As String)
    Dim meta As String, z As Object, ico As Long
    On Error Resume Next
    Set z = frm.Controls("zTitle")
    meta = modUiScreens.ScrMeta(kljuc)
    z.Controls("titName").caption = Poruka(MetaVal(meta, "naslov"))
    z.Controls("titSub").caption = Poruka(MetaVal(meta, "sub"))
    ico = CLng(Val(modUiScreens.ScrField(modUiScreens.ScrRowByKey(kljuc), SCR_IKONICA)))
    If ico <> 0 Then z.Controls("titIco").caption = ChrW(ico)
    z.Controls("titBadgeB").Visible = False
    z.Controls("titBadgeC").Visible = False
    z.Controls("titDatum").caption = Format$(Now, "dd.mm.yyyy.")
    ' mreza: naslov liste iz istog opisa, dijagnostika i cipovi se sklanjaju
    frm.Controls("zGrid").Controls("grdTitle").caption = Poruka(MetaVal(meta, "lista"))
    frm.Controls("zGrid").Controls("grdSrc").Visible = False
    ' Panel "Filteri" nudi vrste i partnere iz tabela dokumenata - na drugom
    ' ekranu bi bio prazan ili pogresan. Dok ekran ne bude mogao da prijavi
    ' svoje filtere, dugme se sklanja. Brza pretraga radi svuda.
    BoxShow frm.Controls("zGrid"), "btnFilteri", False
    Dim ch As Variant
    For Each ch In ChipRow()
        ShowChip frm, Split(CStr(ch), "|")(0), False
    Next ch
    LayoutChips frm
    ' Ekran koji ima prekidac lista dobija naslov iz definicije AKTIVNE liste,
    ' a ne iz opisa ekrana - inace bi u listi stavki i dalje pisalo "Palete".
    RefreshGridTitle frm
    RefreshListSeg frm
End Sub

' Vrednost polja iz opisa oblika "kljuc=A|naslov=B|sub=C".
Private Function MetaVal(ByVal meta As String, ByVal kljuc As String) As String
    Dim p As Variant
    For Each p In Split(meta, "|")
        If Left$(CStr(p), Len(kljuc) + 1) = kljuc & "=" Then
            MetaVal = Mid$(CStr(p), Len(kljuc) + 2)
            Exit Function
        End If
    Next p
End Function

Public Sub EnsureGridLoaded()
    FillCombos mFrm
    RefreshKpi mFrm
    RefreshStatusBar mFrm
    ' I brojaci uz stavke menija. Bez ovoga bi znacke bile prazne do PRVE promene
    ' podataka -- a zaostatak koji se vidi tek posto nesto uradis ne resava nista;
    ' ceo smisao je da se vidi cim se aplikacija otvori.
    OsveziNavBrojace
    If mViewN > 0 Then Exit Sub
    ReloadGrid
End Sub

' Brojaci uz stavke menija. Pita se SVAKI ekran -- ljuska ne zna koji ih ima --
' ali samo na promenu podataka, jer je izvor prolaz kroz tabele.
Public Sub OsveziNavBrojace()
    Dim r As Variant, kljuc As String, nm As Variant, z As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If mNavKey Is Nothing Then Exit Sub
    Set z = mFrm.Controls("zNav")
    If z Is Nothing Then Exit Sub
    Set mNavBroj = CreateObject("Scripting.Dictionary")
    For Each r In modUiScreens.ScrRows()
        kljuc = modUiScreens.ScrField(CStr(r), SCR_KLJUC)
        mNavBroj(kljuc) = modUiScreens.ScrBrojac(kljuc)
    Next r
    For Each nm In mNavKey.keys
        z.Controls(CStr(nm) & "B").caption = _
            BrojacTekst(CLng(mNavBroj(CStr(mNavKey(CStr(nm))))))
    Next nm
End Sub

' Nula se ne crta. Znacka postoji da bi se video ZAOSTATAK; nula uz svaku stavku
' bi je pretvorila u ukras koji se prestane primecivati.
'
' NEGATIVNO ZNACI "NE ZNAM", i crta se. Prazna znacka u ovom UI-ju nije odsustvo
' poruke nego poruka "nema sta da ceka"; ekran koji nije uspeo da procita svoje
' brojke ne sme da je posalje. Brojac ne moze legitimno da bude negativan, pa je
' to slobodan kanal koji ne trazi izmenu ugovora Scr_Brojac() As Long.
Private Function BrojacTekst(ByVal n As Long) As String
    If n < 0 Then
        BrojacTekst = "!"
        Exit Function
    End If
    If n <= 0 Then Exit Function
    If n > 999 Then BrojacTekst = "999+" Else BrojacTekst = CStr(n)
End Function

Public Sub RefreshFromData()
    On Error Resume Next
    modUiData.ResetCache
    ' i izvedene mape aktivnog ekrana - kasno vezano, ekran ne mora da ima
    Application.Run modUiScreens.ScrField(modUiScreens.ScrRowByKey(mScreen), _
                                          SCR_MODUL) & ".Scr_ResetCache"
    Err.Clear
    Set mPartMap = Nothing
    Set mColSpec = CreateObject("Scripting.Dictionary")
    mColSpec.CompareMode = 1
    RefreshKpi mFrm
    OsveziNavBrojace
    FillZbirneCombo mFrm      ' nove zbirne u picker
    mPartnerFor = ""          ' partner lista se ponovo puni
    ' Skup lista (tabova) sme da zavisi od konteksta ekrana (Izvestaji: tip
    ' entiteta bira koje liste postoje) -- posle promene konteksta geometrija
    ' je zastarela (LayoutGrid crta tabove iz Scr_Liste), a natpis i boje
    ' prate AKTIVNU listu, koju je ekran mogao da prebaci. Za ekrane sa
    ' statickim listama sve tri linije su idempotentne.
    mGeomStara = True
    ' Ekran je mogao da PREBACI aktivnu listu (auto-prelaz tipa/rezima) --
    ' nova lista dobija svoj podrazumevani sort, kao i na klik. Ista lista
    ' zadrzava operaterov izbor sorta (upis ga ne sme gaziti).
    If ActiveLista() <> mSortLista Then PrimeniSortZaListu ActiveLista()
    RefreshListSeg mFrm
    RefreshGridTitle mFrm
    ReloadGrid
End Sub

' Kljuc kesa je REZIM, ne tabela: F5 i F6 citaju istu tblNovac ali razlicite
' kolone vrednosti (Isplata / Uplata), pa bi kes po imenu tabele dao jednom
' rezimu tudje indekse.
Private Function ColSpecIdx(ByVal mk As String, ByVal tblName As String) As Variant
    Dim spec As Variant, ix(0 To 7) As Long, i As Long
    If mColSpec Is Nothing Then Set mColSpec = CreateObject("Scripting.Dictionary")
    If mColSpec.Exists(mk) Then
        ColSpecIdx = mColSpec(mk)
        Exit Function
    End If
    spec = ColumnSpec(mk)
    For i = 0 To 5
        ix(i) = ColIdx(tblName, CStr(spec(i)))
    Next i
    ix(6) = ColIdx(tblName, COL_STORNIRANO)
    ix(7) = ColIdx(tblName, CStr(spec(6)))   ' direktna vrednost (umesto kg * cena)
    mColSpec(mk) = ix
    ColSpecIdx = ix
End Function

'=====================================================================
' OPIS KOLONA MREZE
' Mreza vise nema fiksnih 6 kolona - svaki rezim opisuje svoje. Zapis je
' "KLJUC_NASLOVA|IZVORNA_KOLONA|VRSTA|SIRINA|PRIO".
'
' VRSTA:  txt  tekst
'         part tekst koji se razresava kroz PartnerLookup / RevPartner
'         date datum (serijski broj; prikaz kratak)
'         kg   ceo broj, ULAZI u zbir "Ukupno kg"
'         num  ceo broj, BEZ zbira (kol. ambalaze)
'         sum0 ceo broj, ulazi u zbir vrednosti (komadi reversa)
'         rsd  2 decimale, ulazi u zbir vrednosti; cita se iz izvorne kolone
'         mult 2 decimale, ulazi u zbir vrednosti; = izvorna kolona (cena) x kg
'         pill status pilula (racuna se, nema izvornu kolonu)
'
' SIRINA: 0 = fleksibilna (uzima ostatak reda; tacno jedna po rezimu)
' PRIO:   1 = uvek, 2 = sklanja se prva kad red nije dovoljno sirok
'=====================================================================
' Opis kolona MORA biti postavljen pre nego sto LayoutGrid izracuna sirine i
' pre nego sto RefreshSortGlyphs ispise naslove. Ranije ga je postavljao tek
' FillGrid - a on se izvrsava POSLE oba - pa su zaglavlja i sirine kasnili jedan
' rezim unazad, a celije prethodnog rezima su ostajale vidljive.
Private Sub SetGridCols(ByVal mk As String)
    SetGridColsArr GridCols(mk)
End Sub

' Isto, ali iz niza koji je dao EKRAN. Mreza od S4a ne pita ekran dokumenata
' kako izgledaju kolone - pita AKTIVAN ekran.
Private Sub SetGridColsArr(ByVal a As Variant)
    Dim potpis As String
    If Not IsArray(a) Then Exit Sub

    ' PROMENA OPISA KOLONA CINI GEOMETRIJU ZASTARELOM.
    '
    ' mColX / mColW racuna LayoutGrid, a on se zove iz RASPOREDA ekrana.
    ' ReloadGrid (promena liste, cipa, pretrage) zove samo LoadGridFromScreen i
    ' RenderGrid. Bez ove zastavice je RenderGrid crtao sa sirinama PRETHODNE
    ' liste: kolona koja je tamo bila skrivena (prioritet 4 -> sirina 0) ostajala
    ' je nevidljiva i u novoj listi, ma koliko joj vrednost bila ispravna.
    '
    ' Zaglavlje je pri tom umelo da bude vidljivo, jer ga osvezava RefreshSortGlyphs
    ' na kraju sledeceg rasporeda -- pa je izgled bio najgori moguci: naslov
    ' kolone stoji, celije prazne. Nadjeno merenjem na listi izvoda ekrana Uvoz
    ' izvoda (v6-ui-177); pogadja svaki ekran cije se liste razlikuju po broju
    ' vidljivih kolona.
    potpis = ColsPotpis(a)
    If potpis <> mColsPotpis Then
        mColsPotpis = potpis
        mGeomStara = True
    End If

    mCols = a
    mColN = BazenStaje(UBound(mCols) + 1, MAX_COLS, "kolone")
End Sub

' Potpis opisa kolona. Poredi se SADRZAJ, ne referenca: ekran vraca nov niz pri
' svakom citanju, pa bi poredjenje objekata uvek javljalo promenu.
Private Function ColsPotpis(ByVal a As Variant) As String
    Dim i As Long, s As String
    If Not IsArray(a) Then Exit Function
    For i = LBound(a) To UBound(a)
        s = s & CStr(a(i)) & vbLf
    Next i
    ColsPotpis = s
End Function

' Preracunaj geometriju ako je opis kolona u medjuvremenu promenjen. Zove se sa
' JEDNOG mesta -- s pocetka RenderGrid-a -- da bi svaki put kad se crta vazilo
' da mColW odgovara mCols.
Private Sub OsveziGeometriju(z As Object)
    If Not mGeomStara Then Exit Sub
    If z Is Nothing Then Exit Sub
    LayoutGrid z, z.width, z.Height
End Sub

Private Function SortedView(ByRef a() As Variant, ByVal n As Long, ByVal nc As Long, _
                            ByVal col As Long, ByVal asc As Boolean) As Variant
    Dim idx() As Long, outA() As Variant, i As Long, c As Long
    If col < 1 Or col > nc Then col = 2
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    If n > 1 Then QSortIdx idx, a, col, asc, 1, n

    ReDim outA(1 To n, 1 To nc)
    For i = 1 To n
        For c = 1 To nc
            outA(i, c) = a(idx(i), c)
        Next c
    Next i
    SortedView = outA
End Function

Private Sub QSortIdx(ByRef idx() As Long, ByRef k() As Variant, ByVal col As Long, _
                     ByVal asc As Boolean, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, t As Long
    Dim pivot As Variant
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pivot = k(idx((lo + hi) \ 2), col)
    Do While i <= j
        Do While CompareKey(k(idx(i), col), pivot, asc) < 0
            i = i + 1
        Loop
        Do While CompareKey(k(idx(j), col), pivot, asc) > 0
            j = j - 1
        Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSortIdx idx, k, col, asc, lo, j
    If i < hi Then QSortIdx idx, k, col, asc, i, hi
End Sub

Private Function CompareKey(ByVal X As Variant, ByVal Y As Variant, ByVal asc As Boolean) As Long
    Dim r As Long
    If IsNumeric(X) And IsNumeric(Y) Then
        r = Sgn(CDbl(X) - CDbl(Y))
    Else
        r = StrComp(CStr(X), CStr(Y), vbTextCompare)
    End If
    CompareKey = IIf(asc, r, -r)
End Function

' Redovi koje je dao ekran. Ugovor vraca niz:
'   0 = kolone (isti oblik kao GridCols), 1 = redovi(1..n, 1..kolona),
'   2 = n, 3 = zbir kg, 4 = zbir vrednosti
' Ljuska odavde radi sve ostalo - sortiranje, strane, pretragu, izbor, crtanje.
' Suzavanje iz panela "Filteri" - ekran ga cita, ljuska ga drzi.
' Celija prikazanog reda - ekran je cita kad korisnik izabere red. Bez ovoga
' bi ekran morao da drzi svoju kopiju istih podataka.
' Koliko redova mreza trenutno drzi. Ekran to ne moze da izvede iz svojih
' podataka: ljuska filtrira, sortira i strani, pa je njen broj jedini tacan.
' Postoji zbog dijagnostike -- radnja nad redom koji mreza nema mora da moze
' da kaze KOLIKO ih ima, inace se pad ne razlikuje od praznog ekrana.
Public Function GridBrojRedova() As Long
    GridBrojRedova = mViewN
End Function

' TEST SEAM: pusti raspored mreze nad zonom koju je test vec izgradio.
' Tvrdo gejtovan.
'
' LayoutGrid je Private i zove se samo iz gradnje forme, pa se geometrija tela
' i trake poruka drugacije ne moze izmeriti: test formu ne prikazuje, a bez
' rasporeda su sve kontrole na nuli. Same kontrole test cita direktno --
' seam vraca prostor, ne brojke, da tvrdnja ostane u testu.
Public Sub GridLayoutTest(z As Object, ByVal zw As Single, ByVal zh As Single)
    If Not IsTestMode() Then Exit Sub
    LayoutGrid z, zw, zh
End Sub

' TEST SEAM: sagradi mrezu nad datom formom i NACRTAJ je. Tvrdo gejtovan.
'
' Bez ovoga se pravilo "upis se radi uvek" moze izmeriti samo na CelijaTekst --
' dakle nad pravilom, ne nad crtanjem. A ceo kvar je bio bas u crtanju: racun je
' bio U UPISU, pa je pad preskakao upis.
'
' Mrezu NE gradi -- forma je vec nosi (UserForm_Initialize). Seam samo veze
' ljusku za tu formu i pusti isti raspored i isto crtanje koje ide u pogonu.
Public Sub GridRenderTest(frm As Object, ByVal zw As Single, ByVal zh As Single)
    If Not IsTestMode() Then Exit Sub
    Set mFrm = frm
    LayoutGrid frm.Controls("zGrid"), zw, zh
    RenderGrid
End Sub

' TEST SEAM: upisi vrednost u ucitanu mrezu. Tvrdo gejtovan.
'
' Fixture nosi ISPRAVNE vrednosti -- takav je i treba da bude. Vrednost koja se
' ne moze prikazati zato mora da se ubaci ovde, posle ucitavanja: sejanje takvog
' reda u tabelu bi obaralo tudje testove (jednom vec sedam, sa Overflow), a
' izmisljanje sopstvenog niza bi merilo izmisljotinu umesto puta kojim ide
' operater.
Public Sub GridTestVrednost(ByVal red As Long, ByVal kol As Long, ByVal v As Variant)
    If Not IsTestMode() Then Exit Sub
    If red < 1 Or red > mViewN Then Exit Sub
    If kol < 1 Or kol > MAX_COLS Then Exit Sub
    mView(red, kol) = v
End Sub

' TEST SEAM: pusti ljusku sa test-forme. Tvrdo gejtovan.
'
' GridRenderTest veze mFrm za formu koju je test napravio. Bez ovoga bi posle
' suite ostala ziva referenca na formu koja je Unload-ovana, pa bi svaki sledeci
' RenderGrid radio nad njom.
Public Sub GridOtkaciFormuTest()
    If Not IsTestMode() Then Exit Sub
    Set mFrm = Nothing
End Sub

' NATPIS JEDNOG NOVCANOG SLOTA, na jednom mestu (v. PodnozjeValTekst).
'
' Slot uvek nosi novac: pojam postoji zato sto jedan zbir nije dovoljan, a ne
' zato sto se broji nesto drugo. Zato ovde nema pitanja o komadima -- ekran koji
' broji komade nema sta da trazi od dva novcana slota.
' Jedinica je opcioni TRECI clan slota (kljuc poruke); bez njega RSD --
' ambalazni zbirovi (gajbe) ne smeju da se potpisu kao novac (krug 15).
Private Function PodnozjeSlotTekst(ByVal kljuc As String, ByVal iznos As Double, _
                                   Optional ByVal unitKljuc As String = "OTKUI_UNIT_RSD") As String
    PodnozjeSlotTekst = Poruka(kljuc) & " " & FmtBroj(iznos, 2) & " " & _
                        Poruka(unitKljuc)
End Function

' Dva novcana slota u podnozju.
'
' Desna ivica je fiksna, drugi slot se dodaje ULEVO -- da se jedan slot ne pomeri
' kad se pojavi drugi. Taj prostor levo od sredine deli sa zbirom kilograma: ako
' su oba tu, crta se samo prvi, i to se PRIJAVLJUJE kroz BazenStaje. Tiho
' odsecanje je klasa greske koju ljuska vise ne pravi.
Private Sub CrtajSlotovePodnozja(ft As Object)
    Dim n As Long, w As Single
    n = mFtSlotN
    If n > 1 Then
        If ft.Controls("ftKg").Visible Then n = BazenStaje(n, 1, "podnozje-uz-kg")
    End If

    w = ft.width
    ft.Controls("ftVal").Left = w / 2 - 10 - IIf(n > 1, 200, 0)
    ft.Controls("ftVal2").Left = w / 2 - 10

    ft.Controls("ftVal").Visible = True
    ft.Controls("ftVal").caption = SlotTekstIz(0)
    ft.Controls("ftVal2").Visible = (n > 1)
    If n > 1 Then ft.Controls("ftVal2").caption = SlotTekstIz(1)
End Sub

' Natpis i-tog slota iz onoga sto je ekran vratio. Van opsega daje prazno.
Private Function SlotTekstIz(ByVal i As Long) As String
    On Error Resume Next
    If i < 0 Or i >= mFtSlotN Then Exit Function
    If UBound(mFtSlot(i)) >= 2 Then
        SlotTekstIz = PodnozjeSlotTekst(CStr(mFtSlot(i)(0)), CDbl(mFtSlot(i)(1)), _
                                        CStr(mFtSlot(i)(2)))
    Else
        SlotTekstIz = PodnozjeSlotTekst(CStr(mFtSlot(i)(0)), CDbl(mFtSlot(i)(1)))
    End If
End Function

' TEST SEAM: natpis i-tog novcanog slota i njihov broj. Tvrdo gejtovani.
'
' Crtanje trazi pravu formu, a bas se u natpisu vidi da li je slot dobio SVOJ
' iznos i svoju jedinicu -- razlika izmedju "dva broja" i "dva TACNA broja".
Public Function GridPodnozjeSlotTest(ByVal i As Long) As String
    If Not IsTestMode() Then Exit Function
    GridPodnozjeSlotTest = SlotTekstIz(i)
End Function

Public Function GridPodnozjeSlotBrojTest() As Long
    If Not IsTestMode() Then Exit Function
    GridPodnozjeSlotBrojTest = mFtSlotN
End Function

' NATPIS ZBIRA VREDNOSTI U PODNOZJU, na jednom mestu.
'
' Reversi se broje u komadima, sve ostalo u dinarima (i bez decimala). Pitanje
' ide EKRANU, a ne globalnom ActiveMode: taj rezim pripada Dokumentima i na
' ugovornom ekranu ostaje onakav kakav ga je Dokumenta ostavila. Ko je bio na F7
' pa presao na Uvoz izvoda video je "Ukupno 8.950 kom" umesto
' "Vrednost 8.950,00 RSD" -- novac izbrojan kao komadi, bez decimala.
'
' Izdvojeno iz RenderGrid da bi se moglo IZMERITI: crtanje trazi pravu formu, a
' jedinica i broj decimala vide se samo u gotovom natpisu. Seam ispod zove BAS
' ovu funkciju -- kopija bi se razisla sa njom i tvrdnja bi merila kopiju.
Private Function PodnozjeValTekst(ByVal iznos As Double) As String
    If modUiScreens.ScrBrojiKomade(mScreen) Then
        PodnozjeValTekst = Poruka("OTKUI_FT_UKUPNO") & " " & _
                           FmtBroj(iznos, 0) & " " & Poruka("OTKUI_UNIT_KOM")
    Else
        PodnozjeValTekst = Poruka("OTKUI_FT_VREDNOST") & " " & _
                           FmtBroj(iznos, 2) & " " & Poruka("OTKUI_UNIT_RSD")
    End If
End Function

' TEST SEAM: natpis koji bi podnozje napisalo za tekuci ekran. Tvrdo gejtovan.
'
' Bez ovoga se "podnozje ne sme da broji novac u komadima" ne moze izmeriti:
' RenderGrid trazi pravu formu, a jedinica i broj decimala vide se samo u
' gotovom natpisu.
Public Function GridPodnozjeValTest(ByVal iznos As Double) As String
    If Not IsTestMode() Then Exit Function
    GridPodnozjeValTest = PodnozjeValTekst(iznos)
End Function

' TEST SEAM: da li ljuska za tekucu listu crta zbir vrednosti u podnozju.
' Tvrdo gejtovan.
'
' Bez ovoga se moze tvrditi samo da EKRAN vraca ispravan zbir (Scr_Rows), a ne i
' da ga ljuska PRIKAZUJE -- a upravo je ta razlika jednom vec sakrila podnozje
' liste izvoda uz zelenu suite.
Public Function GridImaValKolonuTest() As Boolean
    If Not IsTestMode() Then Exit Function
    GridImaValKolonuTest = ModeHasValCol()
End Function

' TEST SEAM: vrsta i-te kolone (0-bazirano). Tvrdo gejtovan. Test bez ovoga ne
' moze da nadje kolonu nad kojom konverzija UOPSTE moze da pukne -- nad txt
' kolonom svaka vrednost prolazi, pa bi tvrdnja merila nista.
Public Function GridKindKoloneTest(ByVal i As Long) As String
    If Not IsTestMode() Then Exit Function
    If i < 0 Or i > MAX_COLS - 1 Then Exit Function
    GridKindKoloneTest = ColKind(i)
End Function

' TEST SEAM: koliko redova mreza drzi. Tvrdo gejtovan. Bez ovoga se "celija je
' prazna" ne razlikuje od "nema sta da se crta".
Public Function GridRedovaTest() As Long
    If Not IsTestMode() Then Exit Function
    GridRedovaTest = mViewN
End Function

' TEST SEAM: koliko celija poslednje crtanje nije umelo da prikaze. Tvrdo
' gejtovan. Bez ovoga se "upis se radi UVEK" ne moze razlikovati od "vrednost je
' slucajno bila ispravna".
Public Function GridKvarCelijaTest() As Long
    If Not IsTestMode() Then Exit Function
    GridKvarCelijaTest = mKvarCelija
End Function

' TEST SEAM: koja je kolona prva zatajila. Tvrdo gejtovan.
Public Function GridKvarKindTest() As String
    If Not IsTestMode() Then Exit Function
    GridKvarKindTest = mKvarKind
End Function

' TEST SEAM: da li geometrija kolona zaostaje za opisom. Tvrdo gejtovan.
' Bez ovoga se pravilo "promena opisa kolona cini geometriju zastarelom" ne moze
' izmeriti: mColW je Private, a RenderGrid trazi pravu formu.
Public Function GridGeomStaraTest() As Boolean
    If Not IsTestMode() Then Exit Function
    GridGeomStaraTest = mGeomStara
End Function

' TEST SEAM: sirina i-te kolone (0-bazirano). Tvrdo gejtovan.
Public Function GridSirinaKoloneTest(ByVal i As Long) As Single
    If Not IsTestMode() Then Exit Function
    If i < 0 Or i > MAX_COLS - 1 Then Exit Function
    GridSirinaKoloneTest = mColW(i)
End Function

' TEST SEAM: isto sto RenderGrid radi pre crtanja. Tvrdo gejtovan.
Public Sub GridOsveziGeomTest(z As Object)
    If Not IsTestMode() Then Exit Sub
    OsveziGeometriju z
End Sub

' Test seam: mreza se puni BEZ forme. Tvrdo gejtovan.
'
' Postoji da bi klik na red mogao da se izmeri onako kako se STVARNO desava --
' preko GridCell nad pravim podacima ekrana. Bez ovoga bi test morao da izmisli
' red, pa bi merio sopstvenu izmisljotinu umesto puta kojim ide operater.
' Prazan kljuc vraca ljusku na pocetno stanje.
Public Sub GridTestLoad(ByVal kljuc As String)
    If Not IsTestMode() Then Exit Sub
    mFilter = "sve"
    mSearch = ""
    mPage = 1
    If Len(kljuc) = 0 Then
        mScreen = SCR_POCETNI
        mViewN = 0
        mView = Empty
        Exit Sub
    End If
    mScreen = kljuc
    LoadGridFromScreen
End Sub

' KOLIKO SME DA SE UZME iz onoga sto je ekran zatrazio.
'
' Bazen ljuske je konacan: segmenata, radnji, cipova i kolona ima tacno onoliko
' koliko je kontrola napravljeno pri gradnji. Ekran koji zatrazi vise nije
' gresio -- ali visak se do sada gubio BEZ IJEDNE PORUKE. To se dogodilo vec
' dvaput: jedanaesti cip se crtao a nije reagovao (v6-ui-143), a sesta radnja
' nad redom je tiho izbacila 'Nepotpune palete' (v6-ui-162).
'
' Odsecanje ostaje -- visak nema gde da se nacrta. Ono sto se menja je da
' odsecanje dobija IME: koji ekran, koja lista, koja vrsta i koliko je trazeno.
'
' Prijavljuje se JEDNOM po (ekran, lista, vrsta): ovo se zove pri svakom
' crtanju, pa bi inace log rastao bez kraja. Toast ide samo u dev buildu --
' operater od te brojke nema sta da uradi, ali trag mora da ostane uvek.
Public Function BazenStaje(ByVal trazeno As Long, ByVal bazen As Long, _
                           ByVal vrsta As String) As Long
    Dim k As String
    BazenStaje = trazeno
    If trazeno <= bazen Then Exit Function
    BazenStaje = bazen
    On Error Resume Next
    k = mScreen & "/" & ActiveLista() & "/" & vrsta
    If mBazenPrijave Is Nothing Then _
        Set mBazenPrijave = CreateObject("Scripting.Dictionary")
    If mBazenPrijave.Exists(k) Then Exit Function
    mBazenPrijave(k) = True
    LogWarn "modOtkupUI.BazenStaje", k & ": trazeno " & trazeno & _
            ", bazen ima " & bazen & " -- visak se NE crta"
    If IsDebugUI() Then _
        ShowToast Poruka("OTKUI_ERR_BAZEN") & " " & k & " " & _
                  trazeno & "/" & bazen, True
    Err.Clear
End Function

' Koliko je razlicitih prekoracenja prijavljeno. Postoji da bi se moglo
' TVRDITI da se isto prekoracenje prijavljuje jednom, a ne pri svakom crtanju.
Public Function BazenPrijavaBroj() As Long
    On Error Resume Next
    If mBazenPrijave Is Nothing Then Exit Function
    BazenPrijavaBroj = mBazenPrijave.count
End Function

Public Function GridCell(ByVal r As Long, ByVal c As Long) As Variant
    On Error Resume Next
    If r < 1 Or r > mViewN Then Exit Function
    If Not IsArray(mView) Then Exit Function
    GridCell = mView(r, c)
End Function

Public Function FltVrsta() As String
    FltVrsta = mFltVrsta
End Function

Public Function FltPart() As String
    FltPart = mFltPart
End Function

Private Sub LoadGridFromScreen()
    Dim d As Variant
    mViewN = 0: mSumKg = 0: mSumVal = 0
    mFtSlot = Empty: mFtSlotN = 0
    mCntOtkaz = 0: mCntBezZb = 0: mCntNefakt = 0: mCntOtvor = 0
    d = modUiScreens.ScrGridData(mScreen, mFilter, mSearch)
    ' Ekran koji je PUKAO i ekran koji nema listu daju isti Empty. Razlika mora
    ' da se vidi: bez ovoga mreza tiho ostaje na prethodnoj listi i prethodnom
    ' naslovu, pa prekidac izgleda kao da ne radi (v. ScrGridData).
    If Len(modUiScreens.ScrLastErr) > 0 Then
        Err.Raise ERR_UI_BASE + 22, modUiScreens.ScrLastErr, _
                  Poruka("OTKUI_ERR_LISTA")
    End If
    If Not IsArray(d) Then Exit Sub
    If UBound(d) < 4 Then Exit Sub
    SetGridColsArr d(0)
    mViewN = CLng(d(2))
    ' brojaci cipova - ekran ih vraca kao sesti clan; ekran koji nema cipove
    ' ga ne salje, pa ostaju nule
    If UBound(d) >= 5 Then
        If IsArray(d(5)) Then
            mCntOtkaz = CLng(d(5)(0))
            mCntBezZb = CLng(d(5)(1))
            mCntNefakt = CLng(d(5)(2))
            ' cetvrti clan salje samo lista otpremnica - stariji ekran ga nema
            If UBound(d(5)) >= 3 Then mCntOtvor = CLng(d(5)(3))
        End If
    End If
    mSumKg = CDbl(d(3))
    mSumVal = CDbl(d(4))
    ' NOVCANI SLOTOVI PODNOZJA -- sedmi clan, neobavezan.
    '
    ' Zbir vrednosti je JEDAN broj, a lista izvoda ima dva koja operater trazi:
    ' uplate i isplate. Njihov zbir (promet) se ne moze rastaviti nazad, pa
    ' podatak mora da stigne od ekrana, uz iste filtere pod kojima su i redovi
    ' izbrojani -- racunanje sa strane bi se razislo sa prikazanom listom.
    '
    ' Ekran koji ih ne salje se ponasa kao pre: podnozje crta zbir vrednosti.
    If UBound(d) >= 6 Then
        If IsArray(d(6)) Then
            mFtSlot = d(6)
            mFtSlotN = BazenStaje(UBound(mFtSlot) + 1, MAX_FT_VAL, "podnozje")
        End If
    End If
    ' Ekran bez ijednog reda vraca Empty umesto niza - tako to radi lista
    ' blokova dok otpremnica nije izabrana. Dodela Empty u tipizirani niz
    ' obara "Type mismatch", pa je mreza javljala gresku za stanje koje je
    ' potpuno ispravno (i crveni toast je posle toga ostajao na ekranu).
    If Not IsArray(d(1)) Then
        mViewN = 0
        mView = Empty
        Exit Sub
    End If
    ' Sortiranje je do sada zivelo unutar FillGrid, pa je ovaj put isao mimo
    ' njega: klik na zaglavlje je menjao mSortCol i ponovo citao redove, ali
    ' ih niko nije prerasporedjivao. Zato ide ovde, nad onim sto ekran vrati.
    ' Kopija u tipizirani niz je obavezna - SortedView prima a() As Variant,
    ' a d(1) je Variant koji SADRZI niz.
    Dim tmp() As Variant
    tmp = d(1)
    If mViewN > 0 Then
        mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
    Else
        mView = tmp
    End If
End Sub

Private Sub ReloadGrid()
    Dim errNum As Long, errSrc As String, errDesc As String
    If mBusyGrid Then Exit Sub
    If mBuilding Then Exit Sub
    mBusyGrid = True
    On Error GoTo EH
    ' Od S4b SVI ekrani idu istim putem: mreza pita aktivan ekran za redove.
    ' Ranije je ekran dokumenata imao svoj put (FillGrid je pisao pravo u
    ' stanje ljuske), pa je mreza umela da sluzi samo njega.
    LoadGridFromScreen
    mPage = 1
    RenderGrid
    ' Uspelo citanje gasi PRETHODNU gresku mreze. Crveni toast nema tajmer -
    ' namerno, greska ne sme da nestane sama - pa je bez ovoga ostajao na
    ' ekranu i posle poteza koji je sve doveo u red.
    If mGridToast Then
        mGridToast = False
        HideToast
    End If
    mBusyGrid = False
    Exit Sub
EH:
    ' Ranije je ovde stajalo On Error GoTo Fin, pa je svaka greska u punjenju
    ' tiho preskakala RenderGrid - mreza bi ostala na starom rezimu i izgledalo
    ' bi da "klik ne radi". Sada se vidi gde je puklo.
    '
    ' Korak vise ne dolazi iz stanja ljuske: ekran ga upisuje u IZVOR greske
    ' (Err.Raise ..., "modScrDokumenti.Scr_Rows[petlja po redovima]", ...), pa
    ' se cita odatle. Tako svaki ekran prijavljuje svoje korake, a ljuska ne
    ' mora da zna nijedan.
    ' Err se cita PRVI, pre bilo kog drugog poziva: Poruka() usput obrise
    ' stanje greske (ima svoj On Error Resume Next), pa je toast prijavljivao
    ' prazan izvor i sifru 0 - poruku od koje se nije imalo sta zakljuciti.
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.description
    mBusyGrid = False
    mGridToast = True
    Debug.Print "modOtkupUI.ReloadGrid PAO [" & mScreen & " / " & ActiveMode & _
                " / " & errSrc & "]: " & errNum & " " & errDesc
    ShowToast Poruka("OTKUI_MSG_MREZA_PALA") & " " & errSrc & _
              " (" & errNum & ")", True
End Sub

' Rezimi kojima je 4. kolona TEKST, a ne broj: gotovinski promet nosi kanal,
' reversi nose tip ambalaze. CompareKey pada na StrComp kad vrednost nije broj,
' pa sortiranje po toj koloni radi i za tekst - nije potrebno sedmo polje.
' Ima li aktivni rezim kolonu u kilogramima? Od toga zavisi zbir u podnozju.
' Podnozje mreze pita OPIS KOLONA, a ne rezim dokumenata -- to je vec jednom
' placeno (v. 13. i 14. u UI_MIGRACIJA_KATALOG). Racun je izdvojen u ciste
' funkcije nad opisom da bi se isto pitanje moglo postaviti i listi koja jos
' nije nacrtana: bez toga se tvrdnja "ova lista nema zbirove" moze proveriti
' tek posle crtanja, a to znaci tek kroz formu.
Public Function OpisImaKgKolonu(ByVal cols As Variant) As Boolean
    Dim i As Long
    If Not IsArray(cols) Then Exit Function
    For i = LBound(cols) To UBound(cols)
        If ColF(CStr(cols(i)), 2) = "kg" Then OpisImaKgKolonu = True: Exit Function
    Next i
End Function

Public Function OpisImaValKolonu(ByVal cols As Variant) As Boolean
    Dim i As Long
    If Not IsArray(cols) Then Exit Function
    For i = LBound(cols) To UBound(cols)
        Select Case ColF(CStr(cols(i)), 2)
            Case "rsd", "mult", "sum0", "rest": OpisImaValKolonu = True: Exit Function
        End Select
    Next i
End Function

' Prioritet kolone iz opisa (peto polje): 1 = crta se uvek, 4 = nikad (identitet).
Public Function ColPrioritet(ByVal spec As String) As Long
    ColPrioritet = CLng(val(ColF(spec, 4)))
End Function

Private Function ModeHasKgCol() As Boolean
    ModeHasKgCol = OpisImaKgKolonu(mCols)
End Function

' Da li lista uopste ima novcanu kolonu -- od toga zavisi da li se u podnozju
' crta zbir vrednosti.
'
' "rest" JE novcana kolona. Nije bila na spisku, pa je lista cije su sve novcane
' kolone "rest" ostajala bez podnozja: zbir se uredno racunao (Scr_Rows ga vraca),
' a ftVal.Visible bi bio False -- operater ne vidi promet. Nadjeno na listi
' IZVODA cim su joj kolone presle na "rest".
'
' Provereno da je prosirenje bezbedno: jedina druga "rest" kolona u repou je
' OTKUI_HD_OSTATAK na Dokumentima, koja stoji UZ "mult" kolonu -- tamo je
' ModeHasValCol vec bio True, pa se nista ne menja.
Private Function ModeHasValCol() As Boolean
    ModeHasValCol = OpisImaValKolonu(mCols)
End Function

' PLACENO nosi SAMO boju teksta, bez pozadine. Pozadinu (pilulu) nosi jedino
' kolona STATUS - dve pilule u susednim kolonama se takmice za paznju i red
' pocinje da lici na semafor. Pozadinu celije postavlja PaintRow, u boji reda.
Private Sub PaintPayPill(lbl As Object, ByVal code As Long)
    Dim cap As String, fc As Long
    On Error Resume Next
    Select Case code
        Case PAY_PLACENO
            cap = Poruka("OTKUI_PAY_PLACENO"): fc = C_GREEN
        Case PAY_DELIM
            cap = Poruka("OTKUI_PAY_DELIMICNO"): fc = C_AMBER
        Case PAY_NEFAKT
            cap = Poruka("OTKUI_PAY_NEFAKT"): fc = C_DISABLED_FG
        Case Else
            cap = Poruka("OTKUI_PAY_NEPLACENO"): fc = C_RUST
    End Select
    lbl.caption = cap
    lbl.ForeColor = fc
    lbl.TextAlign = fmTextAlignLeft
    lbl.Font.bold = True
End Sub

' Najnovije zbirne u combo BROJ ZBIRNE. Zamena za lstZbirne iz frmDokumenta:
' tamo je klik na red upisivao broj u prijemnicu, ovde se bira iz panela nad
' samim poljem koje taj broj i cuva. Broj zbirne je oblika "x/ddmmyy" pa nosi
' datum u sebi - zato je sam broj dovoljan za prepoznavanje.
Public Sub FillZbirneCombo(frm As Object)
    Dim CB As MSForms.ComboBox, src As Variant, iBroj As Long, iDat As Long
    Dim r As Long, n As Long, arr() As Variant, key() As Double, i As Long, j As Long
    Dim cur As String
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgBrZbir").Controls("fgBrZbirT")
    ' Punjenje liste NE sme da pojede upisani broj. ComboBox.Clear brise stavke
    ' ALI I vrednost polja, a BROJ ZBIRNE je kontekst koji mora da prezivi
    ' RefreshFromData posle snimanja (bez ovoga je prvi otkupni list dobijao
    ' zbirnu, a svaki sledeci ne). Zato: zapamti tekst, napuni, vrati - a za to
    ' vreme se promene ne racunaju kao izbor operatera.
    cur = CStr(CB.text)
    mZbirnaFill = True
    CB.Clear
    src = CachedTable(TBL_ZBIRNA)
    If Not IsArray(src) Then GoTo XIT
    iBroj = ColIdx(TBL_ZBIRNA, COL_ZBR_BROJ)
    iDat = ColIdx(TBL_ZBIRNA, COL_ZBR_DATUM)
    If iBroj < 1 Then GoTo XIT

    n = UBound(src, 1)
    ReDim arr(1 To n): ReDim key(1 To n)
    j = 0
    For r = 1 To n
        If Len(CellS(src, r, iBroj)) > 0 Then
            j = j + 1
            arr(j) = CellS(src, r, iBroj)
            key(j) = CellDate(src, r, iDat)
        End If
    Next r
    If j = 0 Then GoTo XIT

    ' opadajuce po datumu; lista je kratka pa je prosto umetanje dosta
    For i = 1 To j - 1
        For r = i + 1 To j
            If key(r) > key(i) Then
                Dim tk As Double, tv As Variant
                tk = key(i): key(i) = key(r): key(r) = tk
                tv = arr(i): arr(i) = arr(r): arr(r) = tv
            End If
        Next r
    Next i

    ' ogranicenja na POP_MAX vise nema - panel skroluje i suzava kucanjem
    For i = 1 To j
        CB.AddItem CStr(arr(i))
    Next i
XIT:
    On Error Resume Next
    If Len(cur) > 0 Then CB.text = cur
    mZbirnaFill = False
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillZbirneCombo PAO: " & Err.Number & " " & Err.description
    Resume XIT
End Sub

Public Sub FillCombos(frm As Object)
    ' modHelpers.FillCmb prima ByRef MSForms.ComboBox - Object se ne sme
    ' proslediti direktno (ByRef argument type mismatch), zato tipizirani lokali.
    ' Procedure-level "As MSForms." je bezbedno: IsHardModuleBody skenira SAMO
    ' module-level deo, pa modul ostaje "mek" za self-update.
    Dim ctx As Object
    Dim cbVrsta As MSForms.ComboBox, cbAmb As MSForms.ComboBox
    Dim st As String
    If mCombosFilled Then Exit Sub
    If frm Is Nothing Then Debug.Print "FillCombos: mFrm je Nothing": Exit Sub
    On Error GoTo EH
    mPopMute = True
    mLoading = True

    st = "zCtx":            Set ctx = frm.Controls("zCtx")
    st = "cbVrsta":         Set cbVrsta = ctx.Controls("cbVrsta")
    st = "VrstaVoca":       FillCmb cbVrsta, GetLookupList(TBL_KULTURE, "VrstaVoca", , , True)
    st = "cbOM":            FillComboDisplayID ctx.Controls("cbOM"), TBL_STANICE, "Naziv", "StanicaID"
    st = "cbVozac":         FillComboDisplayID ctx.Controls("cbVozac"), TBL_VOZACI, "Ime", "VozacID"
    ' cbKupac se NE puni ovde - to je partner polje i zavisi od rezima
    ' (FillFormPartner iz SelectModeCore); staticno punjenje bi ga pregazilo.
    st = "sorta":           RefillSorta frm
    st = "fgTipAmbT":       Set cbAmb = frm.Controls("zForm").Controls("fgTipAmb").Controls("fgTipAmbT")
    st = "TipAmbalaze":     FillCmb cbAmb, GetTipAmbalazeOptions()

    FillZbirneCombo frm
    mCombosFilled = True
    mPopMute = False
    mLoading = False
    Exit Sub
EH:
    ' NE gutati tiho - prazan combo bez poruke je bio glavni razlog zasto je
    ' izgledalo da "nista nije povezano".
    mPopMute = False
    mLoading = False
    Debug.Print "modOtkupUI.FillCombos PAO na koraku [" & st & "]: " & _
                Err.Number & " " & Err.description
End Sub

Private Sub RefillSorta(frm As Object)
    Dim ctx As Object, vrsta As String
    Dim cbSorta As MSForms.ComboBox
    On Error Resume Next
    Set ctx = frm.Controls("zCtx")
    vrsta = CStr(ctx.Controls("cbVrsta").value)
    Set cbSorta = ctx.Controls("cbSorta")
    cbSorta.Clear
    If Len(vrsta) = 0 Then Exit Sub
    FillCmb cbSorta, GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", vrsta, True)
End Sub

' Ponovo napuni listu partnera. Zove je ekran kad je tokom upisa kreiran nov
' kooperant - inace bi ga operater video tek posle ponovnog otvaranja.
Public Sub RefreshPartnerLista()
    On Error Resume Next
    Set mPartMap = Nothing
    mPartnerFor = ""
    FillFormPartner mFrm, ActiveMode
End Sub

Public Function PartnerMap(ByVal tblName As String, ByVal idCol As String, _
                            ByVal nameCol As String, ByVal nameCol2 As String) As Object
    Dim d As Object, src As Variant, iId As Long, iNm As Long, iNm2 As Long
    Dim r As Long, k As String, v As String
    ' KLJUC KESA NOSI I KOLONE. Ranije je bio samo ime tabele, a mapa zavisi i
    ' od toga koje se kolone citaju -- prvi pozivalac je time odlucivao sta svi
    ' ostali dobijaju (npr. "Ime" bez "Prezime").
    Dim ck As String
    ck = tblName & "|" & idCol & "|" & nameCol & "|" & nameCol2
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists(ck) Then
        Set PartnerMap = mPartMap(ck)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    src = CachedTable(tblName)
    If IsArray(src) Then
        iId = ColIdx(tblName, idCol)
        iNm = ColIdx(tblName, nameCol)
        iNm2 = ColIdx(tblName, nameCol2)
        If iId > 0 And iNm > 0 Then
            For r = 1 To UBound(src, 1)
                k = CellS(src, r, iId)
                If Len(k) > 0 Then
                    v = CellS(src, r, iNm)
                    If iNm2 > 0 Then v = Trim$(v & " " & CellS(src, r, iNm2))
                    d(k) = v
                End If
            Next r
        End If
    End If
    ' PRAZNA MAPA SE NE KESIRA. Nastaje kad tabela jos nije procitana ili kolone
    ' nisu nadjene -- a kesirana bi znacila da svako ime do kraja sesije pada na
    ' goli ID (operater je video "KOOP-00022" umesto imena). Prazna je "ne znam
    ' jos", ne "nema imena".
    If d.count > 0 Then Set mPartMap(ck) = d
    Set PartnerMap = d
End Function

' Koje tipove partnera nudi polje. Tamo gde dokument zna ko je druga strana
' lista je jednorodna (samo kupci); tamo gde ne zna - isplata ide i kooperantu
' i otkupnom mestu, revers svima - lista je mesovita i natpis polja je
' "Partner". Redosled = najverovatniji tip prvi.
Private Function PartnerSrcOrder(ByVal mode As String) As Variant
    Select Case mode
        Case "F1": PartnerSrcOrder = Array("KOOP")
        Case "F5": PartnerSrcOrder = Array("KOOP", "OM")
        Case "F7": PartnerSrcOrder = Array("KOOP", "OM", "KUP")
        Case Else: PartnerSrcOrder = Array("KUP")
    End Select
End Function

' Tekst stavke: kooperant nosi i svoje otkupno mesto, ostali samo naziv.
' Kooperant bez upisane stanice ostaje na golom imenu - visak separatora bi
' izgledao kao greska u podacima.
Private Function PartnerPrikaz(ByVal src As String, ByVal iD As String, _
                               ByVal naziv As String) As String
    Dim d As Object
    PartnerPrikaz = naziv
    If src <> "KOOP" Then Exit Function
    Set d = KoopOMMap()
    If d Is Nothing Then Exit Function
    If Not d.Exists(iD) Then Exit Function
    If Len(Trim$(CStr(d(iD)))) = 0 Then Exit Function
    PartnerPrikaz = naziv & " " & ChrW(183) & " " & CStr(d(iD))
End Function

' ID-jevi oznaceni kao neaktivni. Kolona "Aktivan" ne postoji u svakoj tabeli
' (schema drift) - tada se ne filtrira nista. Filter NE sme u PartnerMap: taj
' recnik razresava ID -> ime i u mrezi, gde stari dokumenti neaktivnog
' partnera i dalje moraju da prikazu ime.
Private Function NeaktivniSet(ByVal tblName As String) As Object
    Dim d As Object, src As Variant, iId As Long, iAk As Long, r As Long
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists("#NEAKT_" & tblName) Then
        Set NeaktivniSet = mPartMap("#NEAKT_" & tblName)
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    iAk = ColIdx(tblName, "Aktivan")
    If iAk > 0 Then
        src = CachedTable(tblName)
        iId = ColIdx(tblName, PartnerIdCol(tblName))
        If IsArray(src) And iId > 0 Then
            For r = 1 To UBound(src, 1)
                If StrComp(CellS(src, r, iAk), STATUS_NEAKTIVAN, vbTextCompare) = 0 Then
                    d(CellS(src, r, iId)) = True
                End If
            Next r
        End If
    End If
    Set mPartMap("#NEAKT_" & tblName) = d
    Set NeaktivniSet = d
End Function

' KooperantID -> StanicaID (sirov ID, ne naziv). Zaseban kes ("#KOOPST") jer
' PartnerMap kesira po IMENU TABELE, pa bi drugi par kolona nad tblKooperanti
' pregazio mapu imena koju koristi i mreza.
Private Function KoopStanicaMap() As Object
    Dim d As Object, src As Variant, iId As Long, iSt As Long, r As Long
    Dim k As String, sid As String
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists("#KOOPST") Then
        Set KoopStanicaMap = mPartMap("#KOOPST")
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    src = CachedTable(TBL_KOOPERANTI)
    If IsArray(src) Then
        iId = ColIdx(TBL_KOOPERANTI, COL_KOOP_ID)
        iSt = ColIdx(TBL_KOOPERANTI, COL_KOOP_STANICA)
        If iId > 0 And iSt > 0 Then
            For r = 1 To UBound(src, 1)
                k = CellS(src, r, iId)
                sid = CellS(src, r, iSt)
                If Len(k) > 0 And Len(sid) > 0 Then d(k) = sid
            Next r
        End If
    End If
    Set mPartMap("#KOOPST") = d
    Set KoopStanicaMap = d
End Function

Private Function PartnerIdCol(ByVal tblName As String) As String
    Select Case tblName
        Case TBL_KOOPERANTI: PartnerIdCol = COL_KOOP_ID
        Case TBL_STANICE:    PartnerIdCol = "StanicaID"
        Case Else:           PartnerIdCol = COL_KUP_ID
    End Select
End Function

' KooperantID -> naziv otkupnog mesta. Stanica se cuva kao ID
' (tblKooperanti.StanicaID), pa se razresava kroz istu PartnerMap listu
' stanica koju koristi i sam combo.
Private Function KoopOMMap() As Object
    Dim d As Object, om As Object, src As Variant
    Dim iId As Long, iSt As Long, r As Long, k As String, sid As String
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists("#KOOPOM") Then
        Set KoopOMMap = mPartMap("#KOOPOM")
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set om = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")
    src = CachedTable(TBL_KOOPERANTI)
    If IsArray(src) Then
        iId = ColIdx(TBL_KOOPERANTI, COL_KOOP_ID)
        iSt = ColIdx(TBL_KOOPERANTI, COL_KOOP_STANICA)
        If iId > 0 And iSt > 0 Then
            For r = 1 To UBound(src, 1)
                k = CellS(src, r, iId)
                sid = CellS(src, r, iSt)
                If Len(k) > 0 And Len(sid) > 0 Then
                    If om.Exists(sid) Then d(k) = om(sid) Else d(k) = sid
                End If
            Next r
        End If
    End If
    Set mPartMap("#KOOPOM") = d
    Set KoopOMMap = d
End Function

' tabela | ID kolona | naziv | drugi deo naziva ("" ako ga nema) | kljuc tipa
Private Function PartnerSrcDef(ByVal src As String) As Variant
    Select Case src
        Case "KOOP": PartnerSrcDef = Array(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime", "OTKUI_PTIP_KOOP")
        Case "OM":   PartnerSrcDef = Array(TBL_STANICE, "StanicaID", "Naziv", "", "OTKUI_PTIP_OM")
        Case Else:   PartnerSrcDef = Array(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "", "OTKUI_PTIP_KUP")
    End Select
End Function

' Puni PARTNER combo ("Osnovni podaci", poslednje polje) iz istih recnika koje
' mreza koristi za razresavanje ID -> ime, pa se lista i kolona PARTNER ne mogu
' raziici. FillComboDisplayID se ne moze upotrebiti jer prima JEDNU kolonu
' naziva, a kooperant ima Ime i Prezime.
' Oblik combo-a je isti kao kod svih ostalih u aplikaciji
' (modComboBinding.FillComboDisplayID): kolona 1 = prikaz, kolona 2 = skriven
' ID, bez ListWidth-a i bez dodatne vidljive kolone. Ranije je odrednica bila
' zasebna vidljiva kolona - zbog nje je ovaj dropdown izgledao drugacije od
' ostalih, pa je uvucena u sam tekst stavke.
' Kooperantu se uz ime dodaje NJEGOVO OTKUPNO MESTO: tip se podrazumeva, a dva
' istoimena kooperanta se razlikuju bas po stanici. Otkupna mesta i kupci idu
' pod svojim nazivom - u mesovitoj listi ih razdvaja redosled grupa.
' Neaktivni se preskacu, po istom pravilu kao FillComboDisplayID.
Private Sub FillFormPartner(frm As Object, ByVal mode As String)
    Dim CB As MSForms.ComboBox, ord As Variant, sd As Variant, d As Object
    Dim neakt As Object, i As Long, k As Variant
    Dim omFilt As String, koopSt As Object, preskoci As Boolean
    On Error GoTo EH
    Set CB = frm.Controls("zCtx").Controls("cbKupac")
    ' KOOP_FILTER_BY_OM (Podesavanja, default ON): kooperanti se suzavaju na
    ' izabrano otkupno mesto - legacy frmOtkup.FillKooperantCombo /
    ' FillComboKooperantiByStanica. Stanica ulazi u kljuc kesa, inace bi lista
    ' ostala od prethodnog otkupnog mesta.
    '
    ' IZUZETAK: u rezimu ISPLATA (F5) suzavanje vazi UVEK, bez obzira na
    ' zastavicu. Legacy frmDokumenta.cmbOtkupnoMesto_Change je za taj combo
    ' (cmbPrimalacOMUlaz) zvao FillComboKooperantiByStanica BEZUSLOVNO - zastavica
    ' je gejtovala samo otkupni list. Bez ovog izuzetka je iskljucena zastavica
    ' otvarala isplatu kooperantu sa TUDJEG otkupnog mesta, a red novca se i dalje
    ' knjizi na aktivno OM (SaveNovac omID) - novac bi otisao na pogresnu stanicu.
    If KoopFilterByOM() Or mode = "F5" Then omFilt = SelectedStanicaID(frm)
    If mPartnerFor = mode & "|" & omFilt Then Exit Sub    ' isto stanje, ne puni ponovo
    If Len(omFilt) > 0 Then Set koopSt = KoopStanicaMap()
    mPartnerFill = True
    CB.Clear
    ' Treca kolona nosi TIP partnera ("KOOP" / "OM" / "KUP"). U mesovitoj listi
    ' se tip ne moze zakljuciti iz ID-ja, a od njega zavisi kako se dokument
    ' knjizi: isplata kooperantu i isplata otkupnom mestu su dva razlicita tipa
    ' novca, a revers "izdato koop." kupca uopste ne prima (modNovacUnos).
    CB.ColumnCount = 3
    ' BEZ ListWidth-a i BEZ ListRows - lista se otvara tacno u sirini kontrole,
    ' kao kod svih ostalih dropdown-a; svaka druga vrednost napravi zaseban,
    ' siri okvir koji se vidi kao "druga forma". Kolona mora da stane u
    ' kontrolu umanjenu za vertikalni skroler, inace se javi i horizontalni.
    CB.ColumnWidths = "146 pt;0 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    ord = PartnerSrcOrder(mode)
    For i = LBound(ord) To UBound(ord)
        sd = PartnerSrcDef(CStr(ord(i)))
        Set d = PartnerMap(CStr(sd(0)), CStr(sd(1)), CStr(sd(2)), CStr(sd(3)))
        Set neakt = NeaktivniSet(CStr(sd(0)))
        If Not d Is Nothing Then
            For Each k In d.keys
                preskoci = neakt.Exists(CStr(k))
                ' suzavanje po otkupnom mestu vazi samo za kooperante -
                ' otkupna mesta i kupci nisu vezani za stanicu
                If Not preskoci And Len(omFilt) > 0 And CStr(ord(i)) = "KOOP" Then
                    If koopSt Is Nothing Then
                        preskoci = False
                    ElseIf Not koopSt.Exists(CStr(k)) Then
                        preskoci = True          ' bez stanice = nije sa ovog OM
                    Else
                        preskoci = (StrComp(CStr(koopSt(CStr(k))), omFilt, _
                                            vbTextCompare) <> 0)
                    End If
                End If
                If Not preskoci Then
                    CB.AddItem PartnerPrikaz(CStr(ord(i)), CStr(k), CStr(d(k)))
                    CB.List(CB.ListCount - 1, 1) = CStr(k)
                    CB.List(CB.ListCount - 1, 2) = CStr(ord(i))
                End If
            Next k
        End If
    Next i
    mPartnerFor = mode & "|" & omFilt
XIT:
    mPartnerFill = False
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillFormPartner PAO [" & mode & "]: " & Err.Number & " " & Err.description
    Resume XIT
End Sub

' ID izabranog partnera - druga, skrivena kolona combo-a (isti raspored
' kolona kao svuda drugde: modComboBinding.GetComboID).
Private Function PartnerID(frm As Object) As String
    Dim CB As MSForms.ComboBox
    On Error Resume Next
    Set CB = frm.Controls("zCtx").Controls("cbKupac")
    If CB.ListIndex >= 0 Then PartnerID = Trim$(CStr(CB.List(CB.ListIndex, 1)))
End Function

' Tip izabranog partnera ("KOOP" / "OM" / "KUP") - treca, skrivena kolona.
' Prazno kad nista nije izabrano ili je ime otkucano rukom; ekran to tumaci
' (F5 tada knjizi kao isplatu otkupnom mestu, F7 odbija smer ka kooperantu).
Private Function PartnerTip(frm As Object) As String
    Dim CB As MSForms.ComboBox
    On Error Resume Next
    Set CB = frm.Controls("zCtx").Controls("cbKupac")
    If CB.ListIndex >= 0 Then PartnerTip = Trim$(CStr(CB.List(CB.ListIndex, 2)))
End Function

' "Preostali kes" iz frmDokumenta (cmbOtkupBlok): otkupni blokovi izabranog
' primaoca koji nisu isplaceni do kraja. Racunicu NE ponavljamo ovde - koristi
' se postojeci modNovac.GetOpenOtkupi (7 kolona: brDok | OtkupID | vrednost |
' isplaceno | ostatak | datum | StanicaID), isti izvor koji koristi i pregled
' otvorenih blokova, pa se dva ekrana ne mogu raziici.
Private Sub FillOpenBlokovi(frm As Object)
    Dim CB As MSForms.ComboBox, koopID As String, o As Variant, i As Long
    Dim omID As String, n As Long
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT")
    CB.Clear
    mBlokIDs = Empty
    mBlokRest = Empty
    mBlokListaOk = True
    mBlokListaErr = ""
    SetOstatak 0
    If ActiveMode <> "F5" Then Exit Sub
    koopID = PartnerID(frm)
    If Len(koopID) = 0 Then Exit Sub
    ' PRAZNA TABELA I NEPOSTOJECA TABELA NISU ISTI ISHOD: GetTableData vraca isti
    ' Empty za oba, bez greske. Bez ove kapije bi nedostajuca tblOtkup prosla kao
    ' "kooperant nema otvorenih blokova" i isplata bi tiho postala AVANS.
    RequireTable TBL_OTKUP, "modOtkupUI.FillOpenBlokovi"
    o = GetOpenOtkupi(koopID)
    ' Prazno je ISTINA -- ali samo zato sto smo iznad utvrdili da tabela POSTOJI.
    If IsEmpty(o) Then Exit Sub
    ' BLOKOVI SE SUZAVAJU NA AKTIVNO OTKUPNO MESTO. GetOpenOtkupi vraca StanicaID
    ' u sedmoj koloni, a red novca se knjizi na aktivno OM (SaveNovac omID) - bez
    ' ovog filtera bi lista nudila blok sa OM-B dok je u kontekstu OM-A, pa bi
    ' razduzenje tog bloka bilo proknjizeno na pogresnu stanicu. Kapija za isti
    ' slucaj postoji i u core-u (SaveOMUlaz_TX); ovo je da se ne ponudi.
    omID = SelectedStanicaID(frm)
    ReDim mBlokIDs(1 To UBound(o, 1))
    ReDim mBlokRest(1 To UBound(o, 1))
    n = 0
    For i = 1 To UBound(o, 1)
        If Len(omID) = 0 Or StrComp(CStr(o(i, 7)), omID, vbTextCompare) = 0 Then
            n = n + 1
            mBlokIDs(n) = CStr(o(i, 2))
            mBlokRest(n) = CDbl(o(i, 5))
            CB.AddItem CStr(o(i, 1)) & "   " & ChrW(183) & "   " & _
                       FmtBroj(CDbl(o(i, 5)), 0) & " / " & FmtBroj(CDbl(o(i, 3)), 0)
        End If
    Next i
    ' Bez ijednog bloka sa ovog OM niz mora ostati PRAZAN, a ne pun praznih
    ' mesta: BlokID indeksira po rednom broju stavke u combo-u.
    If n = 0 Then
        mBlokIDs = Empty
        mBlokRest = Empty
    ElseIf n < UBound(o, 1) Then
        ReDim Preserve mBlokIDs(1 To n)
        ReDim Preserve mBlokRest(1 To n)
    End If
    Exit Sub
EH:
    mBlokListaOk = False
    mBlokListaErr = Err.description
    ' LogErr, ne Debug.Print: Immediate prozor u pogonu niko ne gleda, a ovaj pad
    ' menja TIP KNJIZENJA. Ide PRE ciscenja -- svaka On Error naredba resetuje Err.
    LogErr "modOtkupUI.FillOpenBlokovi"
    ' Delimicno napunjena lista je gora od prazne: izgleda kao potpuna.
    On Error Resume Next
    CB.Clear
    mBlokIDs = Empty
    mBlokRest = Empty
    SetOstatak 0
End Sub

' Izbor bloka puni plocu NEISPLACENO - iznos koji jos moze da se isplati.
Private Sub ApplyBlokIzbor()
    On Error Resume Next
    SetOstatak BlokOstatak()
End Sub

' Redni broj izabrane stavke u combo-u polja (1-bazno, 0 = nista izabrano).
' Sakriveno polje se racuna kao "nista izabrano": rezim koji ga ne prikazuje
' ne sme da posalje zatecenu vrednost u dokument (isto pravilo kao ParcelaID).
Private Function IzborIdx(ByVal grp As String) As Long
    Dim CB As MSForms.ComboBox
    On Error Resume Next
    If Not mFrm.Controls("zForm").Controls(grp).Visible Then Exit Function
    Set CB = mFrm.Controls("zForm").Controls(grp).Controls(grp & "T")
    If CB Is Nothing Then Exit Function
    IzborIdx = CB.ListIndex + 1
End Function

' OtkupID izabranog otkupnog bloka i njegov neisplaceni ostatak. Ostatak NE
' racunamo ovde - vrednost je ona koju je dala modNovac.GetOpenOtkupi.
Private Function BlokID() As String
    Dim i As Long
    On Error Resume Next
    i = IzborIdx("fgBlok")
    If i < 1 Then Exit Function
    If Not IsArray(mBlokIDs) Then Exit Function
    If i > UBound(mBlokIDs) Then Exit Function
    BlokID = CStr(mBlokIDs(i))
End Function

Private Function BlokOstatak() As Double
    Dim i As Long
    On Error Resume Next
    i = IzborIdx("fgBlok")
    If i < 1 Then Exit Function
    If Not IsArray(mBlokRest) Then Exit Function
    If i > UBound(mBlokRest) Then Exit Function
    BlokOstatak = CDbl(mBlokRest(i))
End Function

' Otvorene fakture izabranog kupca (F6). Racunicu NE ponavljamo - koristi se
' postojeci read-model modNovac.GetOpenFakture (6 kolona: BrojFakture |
' FakturaID | Iznos | Uplaceno | Preostalo | Datum), isti izvor koji koristi i
' frmDokumenta.FillOpenFakture, pa se dva ekrana ne mogu raziici. Taj model vec
' izbacuje stornirane i one bez ostatka, pa ovde filtera nema.
Private Sub FillOpenFakture(frm As Object)
    Dim CB As MSForms.ComboBox, kupID As String, f As Variant, i As Long
    Dim prikaz As String, datTxt As String
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgFaktura").Controls("fgFakturaT")
    CB.Clear
    mFakIDs = Empty
    mFakRest = Empty
    mFakListaOk = True
    mFakListaErr = ""
    If ActiveMode <> "F6" Then Exit Sub
    kupID = PartnerID(frm)
    If Len(kupID) = 0 Then Exit Sub
    ' Isti razlog kao kod blokova: bez ovoga bi nedostajuca tblFakture prosla kao
    ' "kupac nema otvorenih faktura" i uplata bi tiho postala AVANS KUPCA.
    RequireTable TBL_FAKTURE, "modOtkupUI.FillOpenFakture"
    f = GetOpenFakture(kupID)
    ' Prazno je ISTINA -- ali samo zato sto smo iznad utvrdili da tabela POSTOJI.
    If Not IsArray(f) Then Exit Sub
    CB.ColumnCount = 2
    CB.ColumnWidths = "268 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    ReDim mFakIDs(1 To UBound(f, 1))
    ReDim mFakRest(1 To UBound(f, 1))
    ' Iznosi se citaju CDbl-om, kao u FillOpenBlokovi i u legacy FillOpenFakture:
    ' kolone read-modela su vec brojevi. ParseNum bi ovde bio pogresan alat -
    ' on ocekuje SRPSKI zapis, pa bi "1234.56" pretvorio u 123456.
    For i = 1 To UBound(f, 1)
        mFakIDs(i) = NzToText(f(i, 2))
        mFakRest(i) = CDbl(f(i, 5))
        datTxt = ""
        If IsDate(f(i, 6)) Then datTxt = "   " & ChrW(183) & "   " & Format$(CDate(f(i, 6)), "dd.mm.yyyy")
        ' isti oblik kao lista otvorenih blokova: preostalo / ukupno
        prikaz = NzToText(f(i, 1)) & datTxt & "   " & ChrW(183) & "   " & _
                 FmtBroj(CDbl(f(i, 5)), 0) & " / " & FmtBroj(CDbl(f(i, 3)), 0)
        CB.AddItem prikaz
        CB.List(CB.ListCount - 1, 1) = NzToText(f(i, 2))
    Next i
    Exit Sub
EH:
    mFakListaOk = False
    mFakListaErr = Err.description
    LogErr "modOtkupUI.FillOpenFakture"
    On Error Resume Next
    CB.Clear
    mFakIDs = Empty
    mFakRest = Empty
End Sub

' Sme li se gotovinski dokument uopste knjiziti sa OVAKVOM listom.
'
' Kapija je siroka TACNO koliko i kvar, i to je merljivi deo:
'   - bez novca nema odluke faktura/blok, pa unos same ambalaze ne sme da stane;
'   - isplata otkupnom mestu ne dodiruje blokove, pa ni nju lista ne zaustavlja;
'   - u rezimima koji ove liste nemaju (F1-F4, F7) kapija cuti.
'
' Recnik je isti onaj koji ide u modNovacUnos, pa se pravilo proverava bez forme.
'
' Rezim se cita iz p("rezim") zato sto kapija treba da odlucuje nad ISTIM
' payload-om koji ide Save putu (modScrDokumenti po toj vrednosti bira rutinu).
' To NIJE nezavisan izvor istine: SkupiPolja ga pravi kao p("rezim") =
' modeKey(ActiveMode), dakle snimak istog ActiveMode po kome FillOpenBlokovi /
' FillOpenFakture odlucuju da li ce listu uopste citati. Njihova medjusobna
' uskladjenost ostaje INTEGRACIONA PRETPOSTAVKA sinhronog toka, a ne nesto sto
' ova kapija dokazuje.
Private Function NovacListaSme(ByVal p As Object, ByRef outPoruka As String) As Boolean
    outPoruka = ""
    NovacListaSme = True
    If Not p.Exists("rezim") Then Exit Function
    If Not p.Exists("novac") Then Exit Function
    If CDbl(p("novac")) <= 0 Then Exit Function

    Select Case CStr(p("rezim"))
        Case "AMB_ISPLATE"
            ' Blok se bira samo kooperantu; isplata OM-u ide drugim tipom novca.
            If UCase$(Trim$(CStr(p("partnerTip")))) <> "KOOP" Then Exit Function
            If Len(Trim$(CStr(p("kooperantID")))) = 0 Then Exit Function
            NovacListaSme = LjuskaListaSme(mBlokListaOk, mBlokListaErr, _
                                           "otkupnih blokova", "da bloka nema", outPoruka)
        Case "AMB_UPLATE"
            If Len(Trim$(CStr(p("kooperantID")))) = 0 Then Exit Function
            NovacListaSme = LjuskaListaSme(mFakListaOk, mFakListaErr, _
                                           "otvorenih faktura", "da fakture nema", outPoruka)
    End Select
End Function

' Tekst je isti kao u frmDokumenta.ListaSme -- legacy forma i ljuska namerno drze
' svoje kopije pravila unosa (v. UI_MIGRACIJA_KATALOG 5), pa se ni ovo ne deli
' kroz modul; deli se OBLIK poruke, da operater u obe forme vidi istu recenicu.
Private Function LjuskaListaSme(ByVal ucitanoOk As Boolean, ByVal greska As String, _
                                ByVal imeListe As String, ByVal praznoNeZnaci As String, _
                                ByRef outPoruka As String) As Boolean
    outPoruka = ""

    If Not ucitanoOk Then
        outPoruka = "Lista " & imeListe & " NIJE u" & ChrW(269) & "itana (" & _
                    greska & ")." & vbCrLf & _
                    "Prazna lista NE zna" & ChrW(269) & "i " & praznoNeZnaci & "."
        Exit Function
    End If

    LjuskaListaSme = True
End Function

' TEST SEAM: stanje ucitanosti obe liste. Tvrdo gejtovan.
Public Sub UiTestSetListaUcitanost(ByVal blokOk As Boolean, ByVal blokErr As String, _
                                   ByVal fakOk As Boolean, ByVal fakErr As String)
    If Not IsTestMode() Then Exit Sub
    mBlokListaOk = blokOk
    mBlokListaErr = blokErr
    mFakListaOk = fakOk
    mFakListaErr = fakErr
End Sub

' TEST SEAM: sme li se recnik knjiziti sa zatecenim stanjem listi.
Public Function UiTestNovacListaSme(ByVal p As Object) As Boolean
    If Not IsTestMode() Then Exit Function
    Dim s As String
    UiTestNovacListaSme = NovacListaSme(p, s)
End Function

' TEST SEAM: poruka koju bi operater dobio (prazna kad kapija pusta).
Public Function UiTestNovacListaPoruka(ByVal p As Object) As String
    If Not IsTestMode() Then Exit Function
    Dim s As String
    NovacListaSme p, s
    UiTestNovacListaPoruka = s
End Function

Private Function FakturaID() As String
    Dim i As Long
    On Error Resume Next
    i = IzborIdx("fgFaktura")
    If i < 1 Then Exit Function
    If Not IsArray(mFakIDs) Then Exit Function
    If i > UBound(mFakIDs) Then Exit Function
    FakturaID = CStr(mFakIDs(i))
End Function

Private Function FakturaOstatak() As Double
    Dim i As Long
    On Error Resume Next
    i = IzborIdx("fgFaktura")
    If i < 1 Then Exit Function
    If Not IsArray(mFakRest) Then Exit Function
    If i > UBound(mFakRest) Then Exit Function
    FakturaOstatak = CDbl(mFakRest(i))
End Function

' Raspoloziv avans otkupnog mesta - legacy UpdateOMAvansSaldo. Saldo racuna
' modNovac.GetOMAvansSaldo; ovde se samo prikazuje, i to u natpisu polja
' "ISPLATA IZ" (vidi napomenu uz gradnju tog polja).
Private Sub RefreshSaldoOM()
    Dim v As Double, stID As String, cap As String
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If Not mFrm.Controls("zForm").Controls("fgAvans").Visible Then Exit Sub
    cap = Poruka("OTKUI_FLD_AVANS")
    stID = SelectedStanicaID(mFrm)
    If Len(stID) > 0 Then
        v = GetOMAvansSaldo(stID)
        cap = cap & "   " & ChrW(183) & "   " & Poruka("OTKUI_LBL_SALDO_OM") & " " & _
              FmtBroj(v, 0) & " " & Poruka("OTKUI_UNIT_RSD")
    End If
    mFrm.Controls("zForm").Controls("fgAvans").Controls("fgAvansL").caption = cap
End Sub

' Parcele izabranog kooperanta - isti prikaz i isti izvor kao
' frmOtkup.cmbKooperant_Change (KatBroj | Kultura | ha), ID u skrivenoj koloni
' umesto u zagradi, pa ExtractParcelaID vise nije potreban.
Private Sub FillParcele(frm As Object)
    Dim CB As MSForms.ComboBox, koopID As String, d As Variant
    Dim iId As Long, iKo As Long, iKat As Long, iKul As Long, iPov As Long
    Dim r As Long
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT")
    CB.Clear
    If ActiveMode <> "F1" Then Exit Sub
    If Not IsPracenjeParcela() Then Exit Sub
    koopID = PartnerID(frm)
    If Len(koopID) = 0 Then Exit Sub
    d = CachedTable(TBL_PARCELE)
    If Not IsArray(d) Then Exit Sub
    iId = ColIdx(TBL_PARCELE, COL_PAR_ID)
    iKo = ColIdx(TBL_PARCELE, COL_PAR_KOOP)
    iKat = ColIdx(TBL_PARCELE, COL_PAR_KAT_BROJ)
    iKul = ColIdx(TBL_PARCELE, COL_PAR_KULTURA)
    iPov = ColIdx(TBL_PARCELE, COL_PAR_POVRSINA)
    If iId < 1 Or iKo < 1 Then Exit Sub
    CB.ColumnCount = 2
    CB.ColumnWidths = "268 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    For r = 1 To UBound(d, 1)
        If StrComp(CellS(d, r, iKo), koopID, vbTextCompare) = 0 Then
            CB.AddItem CellS(d, r, iKat) & "   " & ChrW(183) & "   " & _
                       CellS(d, r, iKul) & "   " & ChrW(183) & "   " & _
                       FmtBroj(ParseNum(CellS(d, r, iPov)), 2) & " ha"
            CB.List(CB.ListCount - 1, 1) = CellS(d, r, iId)
        End If
    Next r
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillParcele PAO: " & Err.Number & " " & Err.description
End Sub

Private Sub SetOstatak(ByVal v As Double)
    On Error Resume Next
    mFrm.Controls("zForm").Controls("fgOstatak").Controls("fgOstatakV").caption = FmtBroj(v, 0)
End Sub

Private Function SumKgForDate(ByVal tbl As String, ByVal datCol As String, _
                              ByVal kgCol As String, ByVal dKey As Double) As Double
    Dim src As Variant, iD As Long, iK As Long, r As Long
    src = CachedTable(tbl)
    If Not IsArray(src) Then Exit Function
    iD = ColIdx(tbl, datCol): iK = ColIdx(tbl, kgCol)
    If iD < 1 Or iK < 1 Then Exit Function
    For r = 1 To UBound(src, 1)
        If CellDate(src, r, iD) = dKey Then SumKgForDate = SumKgForDate + CellD(src, r, iK)
    Next r
End Function

Private Function CountForDate(ByVal tbl As String, ByVal datCol As String, _
                              ByVal dKey As Double) As Long
    Dim src As Variant, iD As Long, r As Long
    src = CachedTable(tbl)
    If Not IsArray(src) Then Exit Function
    iD = ColIdx(tbl, datCol)
    If iD < 1 Then Exit Function
    For r = 1 To UBound(src, 1)
        If CellDate(src, r, iD) = dKey Then CountForDate = CountForDate + 1
    Next r
End Function

Private Function SelectedStanicaID(frm As Object) As String
    On Error Resume Next
    SelectedStanicaID = GetComboID(frm.Controls("zCtx").Controls("cbOM"))
End Function

Private Function SaldoOMUkupno(ByVal stanicaID As String) As Double
    Dim res As Variant
    On Error Resume Next
    res = ReportSaldoOM(stanicaID, DateSerial(Year(Date), 1, 1), Date)
    If Not IsArray(res) Then Exit Function
    ' poslednji red = UKUPNO; kolona 5 = Saldo
    SaldoOMUkupno = CDbl(res(UBound(res, 1), 5))
End Function

Private Function HeaderStatusText() As String
    Dim u As String
    On Error Resume Next
    u = GetCurrentUserIme()
    If Len(u) = 0 Then u = GetCurrentUser()
    ' bez poznatog operatera segment se izostavlja - crtica je izgledala kao
    ' da nesto nedostaje (isto pravilo kao u statusnoj traci)
    HeaderStatusText = "Online"
    If Len(u) > 0 Then HeaderStatusText = HeaderStatusText & "  " & ChrW(183) & "  " & u
    HeaderStatusText = HeaderStatusText & "  " & ChrW(183) & "  " & Poruka("OTKUI_FOOT_SEZONA")
End Function

Private Function StatusLeftText(ByVal omNaziv As String) As String
    Dim u As String
    On Error Resume Next
    u = GetCurrentUserIme()
    If Len(u) = 0 Then u = GetCurrentUser()
    ' bez izabranog OM-a segment se IZOSTAVLJA; ranije je stajala crtica, pa je
    ' traka izgledala kao da nesto nedostaje
    StatusLeftText = Poruka("OTKUI_SMENA_1")
    If Len(omNaziv) > 0 Then _
        StatusLeftText = omNaziv & "  " & ChrW(183) & "  " & StatusLeftText
    If Len(u) > 0 Then _
        StatusLeftText = StatusLeftText & "  " & ChrW(183) & "  " & u
End Function

Private Sub RefreshStatusBar(frm As Object)
    Dim om As String
    On Error Resume Next
    om = Trim$(CStr(frm.Controls("zCtx").Controls("cbOM").value))
    frm.Controls("zStatus").Controls("stLeft").caption = StatusLeftText(om)
    ' profil u sidebaru nosi isti kontekst kao statusna traka
    frm.Controls("zNav").Controls("profName").caption = OperaterText()
    frm.Controls("zNav").Controls("profSub").caption = ProfilPodnaslov(om)
End Sub

' "smena 1 - Zaovine"; bez izabranog otkupnog mesta ostaje samo smena, ne
' prazan segment sa separatorom.
Private Function ProfilPodnaslov(ByVal om As String) As String
    ProfilPodnaslov = Poruka("OTKUI_SMENA_1")
    If Len(om) > 0 Then _
        ProfilPodnaslov = ProfilPodnaslov & " " & ChrW(183) & " " & om
End Function

' KRATAK DATUM ZA CELIJU MREZE.
'
' "IsNumeric" NAD Date-om JE FALSE. To nije sitnica nego cela klasa kvara: ekran
' koji vrednost preda onakvu kakva u tabeli jeste -- dakle kao Date -- dobijao je
' PRAZNU celiju, bez greske i bez traga. Zbog toga je Banka uvoz svoj datum
' konvertovao u serijski broj (modUiData.CellDate), ali lista FAKTURA to nije
' radila: njena kolona DATUM je bila prazna u svakom redu, i to se nije videlo
' jer nijedan test nije citao NACRTAN datum.
'
' Nadjeno je tek kad je crtanje pocelo da broji celije koje ne ume da prikaze
' (v. CelijaTekst): "date/1 tip=Date vred=[15.3.2026.]".
'
' Zato se Date prima direktno. Serijski broj i dalje prolazi istim putem, sa
' istom gornjom granicom.
Private Function FmtDatumKratko(ByVal v As Variant) As String
    Dim d As Double

    If IsDate(v) Then
        On Error Resume Next
        d = CDbl(CDate(v))
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
    ElseIf IsNumeric(v) Then
        d = CDbl(v)
    Else
        Exit Function
    End If

    If d <= 0 Then Exit Function
    ' GORNJA GRANICA. CDate van opsega baca Overflow, a RenderGrid radi pod
    ' "On Error Resume Next" -- pa bi upis celije bio preskocen i u njoj bi ostao
    ' natpis od RANIJEG crtanja. Pravilo zivi u modUiData; ovde je stitnik na
    ' samom mestu crtanja, jer ovamo stize i ono sto nije proslo kroz CellDate.
    If Not modUiData.DatumSerijskiValidan(d) Then Exit Function
    FmtDatumKratko = Format$(CDate(d), "dd.mm.")
End Function

Private Function FmtDatumPun(ByVal v As Variant) As String
    On Error Resume Next
    FmtDatumPun = Format$(CDate(v), "dd.mm.yyyy")
End Function

' Kilogrami se u ovom poslu vode na decimalu (60,6 kg), pa zaokruzivanje na
' ceo broj nije stvar ukusa nego netacan podatak: blok od 60,6 kg se ispisivao
' kao 61 i mreza se onda nije slagala sa trakom otpremnice, koja isti broj
' pise sa dve decimale. Ceo broj se i dalje pise bez decimala - lista ostaje
' mirna dok stvarno nema sta da se pokaze.
Private Function FmtKg(ByVal v As Double) As String
    If v = Int(v) Then
        FmtKg = FmtBroj(v, 0)
    Else
        FmtKg = FmtBroj(v, 2)
    End If
End Function

Private Function FmtBroj(ByVal d As Double, ByVal dec As Long) As String
    If dec = 0 Then
        FmtBroj = Format$(d, "#,##0")
    Else
        FmtBroj = Format$(d, "#,##0.00")
    End If
End Function

Public Function ParseNum(ByVal s As String) As Double
    s = Replace$(s, " ", "")
    s = Replace$(s, Chr$(160), "")
    s = Replace$(s, ".", "")
    s = Replace$(s, ",", ".")
    ParseNum = val(s)
End Function

Private Function FldText(ByVal grp As String) As String
    On Error Resume Next
    FldText = mFrm.Controls("zForm").Controls(grp).Controls(grp & "T").text
End Function

Private Sub SetFld(ByVal grp As String, ByVal v As String)
    On Error Resume Next
    mFrm.Controls("zForm").Controls(grp).Controls(grp & "T").text = v
End Sub

Private Function GridHeadCaption(ByVal i As Long) As String
    If Not IsArray(mCols) Then Exit Function
    If i > UBound(mCols) Then Exit Function
    GridHeadCaption = Poruka(ColF(CStr(mCols(i)), 0))
End Function

' Da li kolona nosi broj (desno poravnanje + Consolas)?
Private Function ColIsNum(ByVal i As Long) As Boolean
    Select Case ColKind(i)
        Case "kg", "num", "sum0", "rsd", "mult", "rest": ColIsNum = True
    End Select
End Function

Private Function ColKind(ByVal i As Long) As String
    If Not IsArray(mCols) Then Exit Function
    If i > UBound(mCols) Then Exit Function
    ColKind = ColF(CStr(mCols(i)), 2)
End Function

' DIRTY STATE. Znacka u zaglavlju je jedina povratna informacija da unos jos
' nije upisan - bez nje forma puna podataka tvrdi "Sve izmene sacuvane".
' Programsko punjenje polja (promena rezima, ClearForm, punjenje lista) NIJE
' izmena korisnika, zato mLoading.
Private Sub MarkDirty()
    If mBuilding Or mLoading Then Exit Sub
    If mDirty Then Exit Sub
    mDirty = True
    RenderDirtyBadge
End Sub

Private Sub MarkClean()
    mDirty = False
    RenderDirtyBadge
End Sub

Private Sub RenderDirtyBadge()
    Dim z As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set z = mFrm.Controls("zHdr")
    If mDirty Then
        z.Controls("hdrSaveTick").caption = ChrW(9679)      ' puna tackica
        z.Controls("hdrSaveTick").ForeColor = C_GOLD
        z.Controls("hdrSaved").caption = Poruka("OTKUI_HDR_NESACUVANO")
        z.Controls("hdrSaved").ForeColor = C_GOLD
    Else
        z.Controls("hdrSaveTick").caption = ChrW(10003)     ' kvacica
        z.Controls("hdrSaveTick").ForeColor = C_GREEN_LT
        z.Controls("hdrSaved").caption = Poruka("OTKUI_HDR_SACUVANO")
        z.Controls("hdrSaved").ForeColor = RGB(188, 198, 182)
    End If
End Sub

' Dijagnostika u UI-ju (ime izvorne tabele) samo kad je izricito ukljucena ili
' kad je build razvojni - u produkciji operater ne treba da vidi imena tabela.
' Dijagnosticke oznake na ekranu (danas: "list: tblOtkup" iznad mreze).
'
' Ranije je ovo bilo ukljuceno i kad build tag sadrzi "dev". To je u praksi
' znacilo UVEK: svaka nereleasovana radna sveska nosi tag v0.0.0-dev, pa je
' oznaka stajala na ekranu i kod operatera. Sada je iskljucivo svesna odluka:
' UI_DEBUG=DA u tblLocalConfig i nista drugo.
Private Function IsDebugUI() As Boolean
    On Error Resume Next
    IsDebugUI = (StrComp(GetLocalConfigValue("UI_DEBUG", ""), "DA", vbTextCompare) = 0)
End Function

' U sidebaru stoji VERZIJA PROGRAMA. Build UI-ja se prikazuje samo uz
' UI_DEBUG=DA.
'
' Od v6-ui-154 je na nereleasovanoj svesci na to mesto isao OTKUI_BUILD, jer
' svaka takva sveska nosi isti v0.0.0-dev: sa ekrana se nije videlo da li je
' posle ImportAllVBA u svesci nov ili star UI kod (dva puta je merenje islo nad
' neuvezenim buildom, pa se 'nije pomoglo' nije razlikovalo od 'nije uvezeno').
' Bilo je izricito PRIVREMENO, do kraja rada na storno ekranu -- a taj je
' zavrsen jos u Fazi D.
'
' Dijagnostika se ne gubi nego se vezuje za postojeci prekidac: UI_DEBUG=DA u
' tblLocalConfig. Ko meri, upali ga; ko radi, vidi verziju programa.
'
' Oba ne staju zajedno: raspored drzi ovu labelu na 55pt uz desnu ivicu
' sidebara, pa bi spojen tekst bio odsecen -- sto je i bio prvi pokusaj.
Private Function BuildTagOrBlank() As String
    Dim v As String
    On Error Resume Next
    v = "v" & BUILD_VERSION                        ' modBuildInfo (stamp-build.sh)
    If IsDebugUI() Then v = OTKUI_BUILD
    If Len(v) <= 1 Then v = "-"
    BuildTagOrBlank = v
End Function

Public Sub DetectDisplayFont()
    Dim ctl As Object, i As Long, res As String
    res = F_DISPLAY_FB
    On Error Resume Next
    Set ctl = Application.CommandBars.FindControl(iD:=1728)
    If Not ctl Is Nothing Then
        For i = 1 To ctl.ListCount
            If StrComp(ctl.List(i), F_DISPLAY, vbTextCompare) = 0 Then
                res = F_DISPLAY
                Exit For
            End If
        Next i
    End If
    SetLocalConfigValue "UI_DISPLAY_FONT", res, "Font naslova ekrana Otkup i prodaja"
    MsgBox "UI_DISPLAY_FONT = " & res, vbInformation
End Sub

' Traka poruke ide preko cele sirine mreze, tacno iznad podnozja (Prikazano /
' Ukupno / strane). Puna sirina je namerna: razlog odbijanja je cesto duga
' recenica, a poruka koja se sece ne vredi vise od one koje nema.
Private Sub PostaviToast(z As Object, ByVal sc As Single, ByVal footTop As Single)
    Dim tf As Object
    On Error Resume Next
    Set tf = z.Controls("tstScr")
    If tf Is Nothing Then Exit Sub
    tf.Left = PAD
    tf.width = sc
    tf.top = footTop - TOAST_H - 4
    tf.Controls("tstScrBar").Height = TOAST_H
    tf.Controls("tstScrMsg").width = sc - 18
End Sub

Public Sub ShowToast(ByVal msg As String, ByVal isErr As Boolean)
    Dim fr As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set fr = mFrm.Controls("zGrid").Controls("tstScr")
    If fr Is Nothing Then Exit Sub
    fr.BackColor = IIf(isErr, RGB(250, 240, 235), RGB(236, 245, 240))
    fr.Controls("tstScrBar").BackColor = IIf(isErr, C_RUST, C_GREEN_LT)
    fr.Controls("tstScrMsg").caption = msg
    fr.Controls("tstScrMsg").ForeColor = IIf(isErr, C_RUST, C_GREEN)
    ' Telo mreze je napravljeno POSLE trake (grdBody dolazi iza tstScr u
    ' BuildGrid), pa bi bez ovoga poruka stajala ISPOD redova i opet se ne bi
    ' videla -- druga varijanta istog kvara od koga se krenulo.
    fr.ZOrder 0
    fr.Visible = True
    CancelToastTimer
    If Not isErr Then
        mToastPending = Format$(Now + TimeSerial(0, 0, 4), "yyyy-mm-dd hh:nn:ss")
        Application.OnTime CDate(mToastPending), "HideToast", , True
    End If
End Sub

Public Sub HideToast()
    On Error Resume Next
    mToastPending = ""
    If mFrm Is Nothing Then Exit Sub
    mFrm.Controls("zGrid").Controls("tstScr").Visible = False
End Sub

Private Sub CancelToastTimer()
    On Error Resume Next
    If Len(mToastPending) = 0 Then Exit Sub
    Application.OnTime CDate(mToastPending), "HideToast", , False
    mToastPending = ""
End Sub

Public Sub StopOtkupUITimers()
    CancelToastTimer
End Sub

Public Sub OtkupUI_FormClosed()
    On Error Resume Next
    mPopFor = ""
    StopOtkupUITimers
    Set mHovered = Nothing
    Set mFrm = Nothing
    mView = Empty
    mViewN = 0
    ' Sekcija je stanje ljuske, ne forme: bez ciscenja bi sledeca gradnja
    ' (i sledeci test) krenula u maticnoj sekciji koju je neko ostavio.
    mSekcija = SEK_RAD
    mNavUzak = False
    mZadnjiRad = ""
    mZadnjiMaticni = ""
End Sub

Public Sub OtkupUI_Release()
    On Error Resume Next
    StopOtkupUITimers
    ' Panel u radnoj povrsini drzi referencu na OKVIR unutar ove forme, a modul
    ' panela drzi njega. Dok te reference postoje, forma se ne moze osloboditi --
    ' a ImportAllVBA prvo UKLANJA istoimenu komponentu, pa bi uklanjanje palo i
    ' import bi napravio frmOtkupUI1. Zato panel odlazi PRE mFrm.
    ' False: ekrana vise nema kome bi se vratila radna povrsina -- forma se ovog
    ' trenutka rusi. Citanje ekrana ovde bi bilo citanje u prazno.
    modUiPanel.PanelZatvori False
    mPanelRezim = False
    Set mHovered = Nothing
    Set Btns = Nothing
    modUiKit.ResetNumFields             ' drzi reference na kontrole oborene forme
    Set mFrm = Nothing
    modUiData.ResetCache
    Set mColSpec = Nothing
    Set mPartMap = Nothing
    mCombosFilled = False
    mPartnerFor = ""
    mBlokIDs = Empty
    mBlokRest = Empty
    mPopFor = ""
    mPopN = 0
    mPopTop = 0
    mPopMute = False
    mFltOpen = False
    mFltVrsta = ""
    mFltPart = ""
    mDirty = False
    mLoading = False
    mChromeRemoved = False
    mView = Empty
    mViewN = 0
End Sub

Public Sub NoteHover(ByVal b As clsFlatBtn)
    On Error Resume Next
    If Not mHovered Is Nothing Then
        If Not mHovered Is b Then mHovered.ResetVisual
    End If
    Set mHovered = b
End Sub

' Javi sinku da su mu se boje promenile spolja (centralni render), da ResetVisual
' posle hover-a ne vrati zastarelu. Bez ovoga bi svaka kontrola cija se boja
' menja POSLE WireBtn-a imala isti problem kao kartice rezima.
Public Sub RebaseSink(ByVal tag As String, Optional ByVal byPrefix As Boolean = False)
    Dim b As clsFlatBtn
    On Error Resume Next
    If Btns Is Nothing Then Exit Sub
    For Each b In Btns
        If byPrefix Then
            If Left$(b.SinkTag, Len(tag)) = tag Then b.Rebase
        Else
            If b.SinkTag = tag Then b.Rebase
        End If
    Next b
End Sub

Public Sub ResetAllBtnVisuals()
    On Error Resume Next
    ' pokazivac je na praznoj povrsini zone - i strelica sortiranja se gasi
    If mHoverHd >= 0 Then HoverHead -1
    If mHovered Is Nothing Then Exit Sub
    mHovered.ResetVisual
    Set mHovered = Nothing
End Sub

Public Function HandleGlobalKey(ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
    On Error Resume Next
    If KeyCode = vbKeyEscape And Len(mPopFor) > 0 Then
        ClosePopup                      ' Esc prvo zatvara dropdown
        HandleGlobalKey = True
        Exit Function
    End If
    If KeyCode = vbKeyEscape And mFltOpen Then
        CloseFilterPanel
        HandleGlobalKey = True
        Exit Function
    End If
    HandleGlobalKey = True
    Select Case KeyCode
        Case vbKeyF1: SelectMode mFrm, "F1"
        Case vbKeyF2: SelectMode mFrm, "F2"
        Case vbKeyF3: SelectMode mFrm, "F3"
        Case vbKeyF4: SelectMode mFrm, "F4"
        Case vbKeyF5: SelectMode mFrm, "F5"
        Case vbKeyF6: SelectMode mFrm, "F6"
        Case vbKeyF7: SelectMode mFrm, "F7"
        ' F8 je jedini od osam koji vise ne bira REZIM nego EKRAN: storno je
        ' od v6-ui-143 svoj ekran. Taster je ostao isti da operateru ne promeni
        ' prst -- promenilo se samo gde vodi.
        Case vbKeyF8: IdiNaEkran "STORNO"
        ' pomoc je sa F1 presla na F9 - F1 je sada rezim "Otkupni list"
        Case vbKeyF9: ShowToast Poruka("OTKUI_MSG_POMOC"), False
        Case vbKeyEscape
            If Len(mSearch) > 0 Then
                mFrm.Controls("zGrid").Controls("txtSearch").text = ""
                mSearch = ""
                ReloadGrid
            Else
                ' Hide, NE Unload: ova rutina se izvrsava iz clsFlatBtn eventa,
                ' a Unload forme dok je metod njenog sinka na steku je poznat
                ' izvor rusenja. Formu otpusta korisnik preko X (QueryClose).
                mFrm.Hide
            End If
        Case Else
            HandleGlobalKey = False
    End Select
End Function

Public Sub FocusKontekst()
    On Error Resume Next
    mFrm.Controls("zCtx").Controls("cbVrsta").SetFocus
End Sub

Public Sub FocusPretraga()
    On Error Resume Next
    mFrm.Controls("zGrid").Controls("txtSearch").SetFocus
End Sub

Private Sub CommitDokument(ByVal alsoPrint As Boolean)
    Dim p As Object, greska As String, fokus As String, dat As Double
    Dim listaPoruka As String
    On Error GoTo EH

    ' Datum se proverava OVDE jer je jedini podatak koji ljuska ume da procita
    ' pogresno; sve ostale provere pripadaju dokumentu, pa ih radi ekran.
    dat = ParseDatum(FldText("fgDatum"))
    If dat = 0 Then
        ShowToast Poruka("OTKUI_ERR_DATUM"), True
        Exit Sub
    End If

    Set p = SkupiPolja(dat, alsoPrint)

    ' UCITANOST LISTE SE PROVERAVA PRE UPISA, i pre grananja u modNovacUnos.
    ' Tamo prazan otkupID / fakturaID vec znaci AVANS -- modul ne moze da razlikuje
    ' "nema otvorenih" od "lista nije ucitana", jer tu razliku zna samo onaj ko je
    ' listu punio.
    If Not NovacListaSme(p, listaPoruka) Then
        ShowToast listaPoruka, True
        ' I u MsgBox: toast pise u usko polje akcionog reda, a bas rep poruke nosi
        ' RAZLOG -- isti obrazac kao kod neuspelog upisa nize.
        MsgBox listaPoruka, vbExclamation, APP_NAME
        Exit Sub
    End If

    greska = modUiScreens.ScrSave(mScreen, p)
    If Len(modUiScreens.ScrLastErr) > 0 Then
        Debug.Print "modOtkupUI.CommitDokument: " & modUiScreens.ScrLastErr
        ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & modUiScreens.ScrLastErr, True
        Exit Sub
    End If

    If Len(Trim$(greska)) > 0 Then
        ' Prazna poruka sa razmakom znaci "operater je odustao" - nista se ne
        ' prikazuje, forma ostaje kako jeste.
        If greska <> " " Then
            ShowToast greska, True
            fokus = ""
            If p.Exists("fokus") Then fokus = CStr(p("fokus"))
            ' NEUSPEO UPIS ide I u MsgBox. Toast pise u usko polje akcionog
            ' reda, pa se duga poruka vizuelno sece -- a bas rep nosi RAZLOG:
            ' SaveOtpremnica lepi "poruke" iza opsteg prefiksa, pa je operater
            ' video samo "Greska pri cuvanju otpremnice. Promene su vracene" i
            ' nista o tome ZASTO. Isti obrazac zbog kog upozorenje zbirne vec
            ' ide u MsgBox (v. modScrStorno.ObicanStorno).
            '
            ' Razlikovanje ide po FOKUSU, ne po duzini: validacija polja ga
            ' postavlja i pokazuje na polje, pa njoj toast i pripada. Pad upisa
            ' nema fokus -- nema polja na koje bi se pokazalo.
            If Len(fokus) = 0 Then MsgBox greska, vbExclamation, APP_NAME
            FokusNaPolje fokus
        End If
        Exit Sub
    End If

    ' Uspeh: poruke koje je ekran vratio idu u toast, forma se prazni, podaci
    ' se citaju ponovo (traka otpremnice, KPI, liste).
    ClearForm
    RefreshFromData
    RefreshOtpTraka mFrm
    Dim dopuna As String
    If p.Exists("poruke") Then dopuna = Trim$(CStr(p("poruke")))
    ShowToast PorukaUpisano(CStr(p("rezim"))) & " " & CStr(p("rezultat")) & _
              IIf(Len(dopuna) > 0, "  " & ChrW(183) & "  " & dopuna, ""), False

    ' UPOZORENJE UZ USPEH ide i u MsgBox. Dokument JESTE snimljen, ali nesto uz
    ' njega nije: prevezivanje paleta, auto-zbirna, ili zavrsetak ispravke koji
    ' je stao na safe-stopu ("vise ispravki na cekanju"). Toast to dvostruko
    ' gubi -- pise u usko polje, pa se rep sece, a uspesan toast se jos i sam
    ' sakrije posle cetiri sekunde. Operater tako propusti da mu je ostao posao.
    '
    ' Razlikovanje ide po OZNACI koju katalog vec nosi: ChrW(10007) = upozorenje
    ' (DOKUNOS_MSG_VISE_ISPRAVKI, _PALETE_NISU, _ZBIRNA_NIJE, _ISPRAVKA_NIJE),
    ' ChrW(10003) = informacija o uspehu. Cistu informaciju ne guramo u dijalog.
    '
    ' Oznaka je SIGNAL ZA RUTIRANJE, ne deo poruke, pa se pred dijalog skida.
    ' MsgBox crta kroz ANSI kodnu stranu, u kojoj ChrW(10007) ne postoji -- pa ga
    ' je operater video kao vodece '?' ispred recenice. U traci poruka, koja je
    ' Unicode, ista oznaka se crta ispravno i tamo OSTAJE: tamo nosi znacenje
    ' (crveno = nesto nije proslo), a u dijalogu ga vec nosi sam vbExclamation.
    If InStr(1, dopuna, ChrW(10007)) > 0 Then
        MsgBox PorukaZaDijalog(dopuna), vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Sub

' Tekst poruke bez oznake za rutiranje. Oznaka (ChrW(10007)) kaze SLOJU IZNAD da
' poruku treba pokazati u dijalogu -- nije deo recenice. MsgBox crta kroz ANSI
' kodnu stranu u kojoj tog znaka nema, pa ga je operater video kao vodece '?'.
'
' Javna je da bi se moglo tvrditi u testu: skidanje oznake unutar Sub-a koji
' otvara dijalog ne bi se moglo izmeriti, a dijalog u headless runu visi.
Public Function PorukaZaDijalog(ByVal txt As String) As String
    PorukaZaDijalog = Trim$(Replace(txt, ChrW(10007), ""))
End Function

' Vrednosti forme pod LOGICKIM imenima. Ljuska zna gde koje polje stoji, ekran
' zna sta ono znaci - ista podela kao kod ApplyPrefill, samo u drugom smeru.
Private Function SkupiPolja(ByVal dat As Double, ByVal stampaj As Boolean) As Object
    Dim p As Object, ctx As Object
    Dim cbOM As MSForms.ComboBox, cbVoz As MSForms.ComboBox
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    On Error Resume Next
    Set ctx = mFrm.Controls("zCtx")
    Set cbOM = ctx.Controls("cbOM")
    Set cbVoz = ctx.Controls("cbVozac")

    p("rezim") = modeKey(ActiveMode)
    p("datum") = CDate(dat)
    p("stanicaID") = GetComboID(cbOM)
    p("stanicaTekst") = Trim$(CStr(cbOM.value))
    p("kooperantID") = PartnerID(mFrm)
    ' Tip partnera (KOOP / OM / KUP) - u mesovitoj listi se ne moze zakljuciti
    ' iz ID-ja, a gotovinski rezimi po njemu biraju kako se dokument knjizi.
    p("partnerTip") = PartnerTip(mFrm)
    p("partnerTekst") = Trim$(CStr(ctx.Controls("cbKupac").value))
    p("vrsta") = Trim$(CStr(ctx.Controls("cbVrsta").value))
    p("sorta") = Trim$(CStr(ctx.Controls("cbSorta").value))
    p("vozacID") = GetComboID(cbVoz)
    p("brDok") = Trim$(FldText("fgBrOtpr"))
    p("brojZbirne") = Trim$(FldText("fgBrZbir"))
    p("parcelaID") = ParcelaID()
    p("tipAmb") = Trim$(FldText("fgTipAmb"))
    p("kolicinaI") = ParseNum(FldText("fgKgI"))
    p("cenaI") = CenaText(1)
    p("kolAmb") = ParseNum(FldText("fgKolAmb"))
    p("kolAmbIzdata") = ParseNum(FldText("fgAmbPr"))
    p("dveKlase") = (mKlasa = 2)
    p("kolicinaII") = IIf(mKlasa = 2, ParseNum(FldText("fgKgII")), 0#)
    p("cenaII") = IIf(mKlasa = 2, CenaText(2), 0#)
    p("kolAmbII") = IIf(mKlasa = 2, ParseNum(FldText("fgKolAmbII")), 0#)
    p("novac") = ParseNum(FldText("fgNovac"))
    ' --- gotovinski rezimi i reversi (F5 / F6 / F7) ---
    ' Sve cetiri vrednosti dolaze iz polja koja postoje samo u svom rezimu;
    ' IzborIdx vraca 0 za sakriveno polje, pa se u ostalim rezimima salje
    ' prazno i bez ijedne provere ovde (provere pripadaju ekranu i modulu).
    p("otkupID") = BlokID()
    ' Vidljiv tekst ide UZ ID: combo dopusta kucanje, pa "tekst ima a ID nema"
    ' znaci da stavka nije razresena. Ekran to ne tumaci - tumaci modul.
    p("blokTekst") = Trim$(FldText("fgBlok"))
    p("otkupOstatak") = BlokOstatak()
    p("izAvansa") = mIzAvansa
    p("fakturaID") = FakturaID()
    p("fakturaTekst") = Trim$(FldText("fgFaktura"))
    p("fakturaOstatak") = FakturaOstatak()
    p("smerRev") = mSmerRev
    p("stampajUvek") = stampaj
    Set SkupiPolja = p
End Function

' ID parcele iz prikaza ("123 - Njiva 2" -> "123"); prazno kad polje nije
' vidljivo ili nista nije izabrano.
' ID izabrane parcele = SKRIVENA druga kolona combo-a, kao kod svih ostalih
' dropdown-a (PartnerID / modComboBinding.GetComboID). Ranije se ID vadio iz
' teksta trazenjem " - ", a FillParcele gradi prikaz sa " " & ChrW(183) & " "
' - separator se nikad nije nasao, pa je u ParcelaID (i u tblOtkup) odlazio ceo
' prikazni string.
'
' TEST SEAM: Public, ne Private -- modTest.T_ParcelaID_IzSkriveneKolone cuva bas
' taj regres (prikaz umesto ID-ja).
Public Function ParcelaID() As String
    Dim CB As MSForms.ComboBox
    On Error Resume Next
    If Not mFrm.Controls("zForm").Controls("fgParcela").Visible Then Exit Function
    Set CB = mFrm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT")
    If CB Is Nothing Then Exit Function
    If CB.ListIndex >= 0 Then ParcelaID = Trim$(CStr(CB.List(CB.ListIndex, 1)))
End Function

' Izabrana parcela odredjuje ROBU: kultura parcele je sorta, a iz nje se cita
' vrsta. Legacy par: frmOtkup.cmbParcela_Change. Postavljanje vrste okida
' RefillSorta i cenu iz cenovnika, pa se sorta upisuje posle nje.
Private Sub ApplyParcelaProizvod()
    Dim parID As String, kultura As String, vrsta As String, ctx As Object
    On Error Resume Next
    If mLoading Or mBuilding Then Exit Sub
    parID = ParcelaID()
    If Len(parID) = 0 Then Exit Sub
    kultura = NzToText(LookupValue(TBL_PARCELE, COL_PAR_ID, parID, COL_PAR_KULTURA))
    If Len(kultura) = 0 Then Exit Sub
    vrsta = NzToText(LookupValue(TBL_KULTURE, "SortaVoca", kultura, "VrstaVoca"))
    If Len(vrsta) = 0 Then Exit Sub
    Set ctx = mFrm.Controls("zCtx")
    ctx.Controls("cbVrsta").value = vrsta
    RefillSorta mFrm
    ctx.Controls("cbSorta").value = kultura
    AutoFillCena
End Sub

' Potvrda upisa po rezimu - svaki dokument se zove svojim imenom. Nepokriven
' rezim pada na opstu potvrdu umesto da tvrdi da je upisan otkupni list.
Private Function PorukaUpisano(ByVal rezim As String) As String
    Select Case rezim
        Case "OTKUP":       PorukaUpisano = Poruka("OTKUNOS_MSG_UPISAN")
        Case "OTPREMNICA":  PorukaUpisano = Poruka("DOKUNOS_MSG_UPISANA_OTP")
        Case "ZBIRNA":      PorukaUpisano = Poruka("DOKUNOS_MSG_UPISANA_ZBR")
        Case "PRIJEMNICA":  PorukaUpisano = Poruka("DOKUNOS_MSG_UPISANA_PRIJ")
        Case "AMB_ISPLATE": PorukaUpisano = Poruka("NOVUNOS_MSG_UPISANA_ISPL")
        Case "AMB_UPLATE":  PorukaUpisano = Poruka("NOVUNOS_MSG_UPISANA_UPL")
        Case "REVERSI":     PorukaUpisano = Poruka("NOVUNOS_MSG_UPISAN_REV")
        Case Else:          PorukaUpisano = Poruka("OTKUI_MSG_UPISANO")
    End Select
End Function

' Fokus na polje koje je ekran prijavio kao sporno. Imena su LOGICKA, ista ona
' koja ekran koristi u recniku.
Private Sub FokusNaPolje(ByVal kljuc As String)
    On Error Resume Next
    If Len(kljuc) = 0 Then Exit Sub
    Select Case kljuc
        Case "stanicaID":    mFrm.Controls("zCtx").Controls("cbOM").SetFocus
        ' isto polje, dva imena: robni dokumenti ga zovu kooperantID, gotovinski
        ' rezimi i reversi partnerID (tamo partner ne mora biti kooperant)
        Case "kooperantID", "partnerID": mFrm.Controls("zCtx").Controls("cbKupac").SetFocus
        Case "novac":        mFrm.Controls("zForm").Controls("fgNovac").Controls("fgNovacT").SetFocus
        Case "otkupID":      mFrm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT").SetFocus
        Case "fakturaID":    mFrm.Controls("zForm").Controls("fgFaktura").Controls("fgFakturaT").SetFocus
        ' "smerRev" ovde NEMA granu i to nije propust: segmenti su Label-i
        ' (modUiKit.BoxText), a Label ne prima fokus. Poruka o gresci sama kaze
        ' sta se bira, pa kursor ostaje gde jeste umesto da odleti na tudje polje.
        Case "vrsta":        mFrm.Controls("zCtx").Controls("cbVrsta").SetFocus
        Case "sorta":        mFrm.Controls("zCtx").Controls("cbSorta").SetFocus
        Case "brDok":        mFrm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").SetFocus
        Case "kolicinaI":    mFrm.Controls("zForm").Controls("fgKgI").Controls("fgKgIT").SetFocus
        Case "kolicinaII":   mFrm.Controls("zForm").Controls("fgKgII").Controls("fgKgIIT").SetFocus
        Case "cenaI":        mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena1T").SetFocus
        Case "cenaII":       mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena2T").SetFocus
        Case "kolAmb":       mFrm.Controls("zForm").Controls("fgKolAmb").Controls("fgKolAmbT").SetFocus
        Case "kolAmbII":     mFrm.Controls("zForm").Controls("fgKolAmbII").Controls("fgKolAmbIIT").SetFocus
        Case "tipAmb":       mFrm.Controls("zForm").Controls("fgTipAmb").Controls("fgTipAmbT").SetFocus
        Case "parcelaID":    mFrm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT").SetFocus
        Case "vozacID":      mFrm.Controls("zCtx").Controls("cbVozac").SetFocus
    End Select
End Sub

' KONTEKST OTKUPNOG MESTA I DATUMA (Z3 + Z14).
'
' Otkupno mesto i datum nisu samo dva polja: iz njih se racuna predlog broja
' dokumenta, i oni cine "aktivni kontekst" koji modStanicaLock cuva izmedju
' dokumenata (na desktop instalaciji lock je no-op, ali kontekst se i dalje
' pamti). Legacy to radi u cmbOtkupnoMesto_Change i txtDatum_AfterUpdate.
Private Sub OnStanicaChanged()
    Dim ctx As Object, CB As MSForms.ComboBox, stID As String, dat As Date
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If mLoading Or mBuilding Then Exit Sub      ' prefill sam postavlja kontekst
    Set ctx = mFrm.Controls("zCtx")
    Set CB = ctx.Controls("cbOM")
    stID = GetComboID(CB)
    If Len(stID) = 0 Then
        ' operater je obrisao izbor - pusti aktivnu stanicu i skloni predlog
        ' broja: bez stanice niz ne postoji, a zatecen broj bi bio tudji
        ' (legacy cmbOtkupnoMesto_Change: txtBrojDokumenta / txtBrojOtp = "")
        If Len(GetActiveStanica()) > 0 Then ReleaseStanicaLock GetActiveStanica()
        SetFld "fgBrOtpr", ""
        Exit Sub
    End If

    ' Kooperant i parcela pripadaju PRETHODNOM otkupnom mestu - legacy ih brise
    ' na samom pocetku cmbOtkupnoMesto_Change. Vazi za rezime u kojima je
    ' partner kooperant; kupac (otpremnica, prijemnica) sa stanicom ne stoji u
    ' toj vezi, pa se ne dira.
    If PartnerJeKooperant(ActiveMode) Then
        mFrm.Controls("zCtx").Controls("cbKupac").value = ""
        mFrm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT").Clear
    End If
    ' lista kooperanata prati stanicu kad je KOOP_FILTER_BY_OM ukljucen
    RefreshPartnerLista

    ' Promena otkupnog mesta VAN konteksta izabrane otpremnice: njen datum,
    ' zbirna i roba vise ne vaze. Isto sto legacy radi u cmbOtkupnoMesto_Change.
    NapustiOtpremnicu stID

    dat = DatumIzPolja()
    If Not AcquireStanicaLock(stID, dat) Then _
        ShowToast Poruka("OTKUI_ERR_STANICA") & " " & stID, True

    ' MALINA: vozac je par-vozac otkupnog mesta (VozacID = StanicaID), pa se
    ' bira sam - isto kao u legacy. Ako par-vozac ne postoji, ostaje prazno.
    If IsMalinaMode() Then
        Dim cbVoz As MSForms.ComboBox
        Set cbVoz = ctx.Controls("cbVozac")
        SelectComboByDisplayID cbVoz, stID
    End If

    ' BROJ DOKUMENTA UVEK PRATI STANICU. Predlog se racuna i kad zakljucavanje
    ' stanice padne (mreza, tudji lock): broj je stvar prikaza, a pravo na upis
    ' se ionako proverava pri snimanju. Ranije je pad locka radio Exit Sub, pa
    ' je u polju ostajao broj prethodnog otkupnog mesta.
    ' checkRemote=True kao u legacy: promena stanice je redak potez i sme da
    ' pita i Google (drugi uredjaj je mogao zauzeti broj). Bez upita se ovde
    ' dobija broj koji na serveru vec postoji.
    RefreshBrojPredlog True
    RefreshStatusBar mFrm
End Sub

' Da li je partner dokumenta kooperant (pa ga stanica filtrira i brise).
' Izvedeno iz iste liste izvora koju koristi i sam combo - bez drugog spiska
' rezima koji bi mogao da odluta.
Private Function PartnerJeKooperant(ByVal mode As String) As Boolean
    Dim ord As Variant
    On Error Resume Next
    ord = PartnerSrcOrder(mode)
    If Not IsArray(ord) Then Exit Function
    PartnerJeKooperant = (CStr(ord(LBound(ord))) = "KOOP")
End Function

' Izlazak iz konteksta otpremnice kad se promeni otkupno mesto. Polja se pisu
' pod mLoading da ugnjezdeni okidaci (datum, vrsta) ne bi ponovo ulazili u
' OnStanicaChanged - stanje se sredjuje jednom, na kraju.
Private Sub NapustiOtpremnicu(ByVal stID As String)
    Dim otpSt As String
    On Error Resume Next
    If mScreen <> "DOKUMENTI" Then Exit Sub
    otpSt = CStr(Application.Run("modScrDokumenti.Scr_OtpStanica"))
    Err.Clear
    If Len(otpSt) = 0 Then Exit Sub
    If StrComp(otpSt, stID, vbTextCompare) = 0 Then Exit Sub

    Application.Run "modScrDokumenti.Scr_OtpOtkazi"
    Err.Clear

    mLoading = True
    SetDatumDanas mFrm.Controls("zForm")
    SetFld "fgBrZbir", ""
    mAktivnaZbirna = ""
    ' roba je bila prepisana sa otpremnice - vrati podrazumevanu iz Podesavanja
    mFrm.Controls("zCtx").Controls("cbVrsta").value = ""
    mFrm.Controls("zCtx").Controls("cbSorta").value = ""
    mLoading = False
    ApplyDefaultRoba

    mSelRow = 0                       ' lista je druga - stari izbor ne vazi
    RefreshOtpTraka mFrm
    RefreshListSeg mFrm
    RefreshGridTitle mFrm
    ReloadGrid
    LayoutOtkup mFrm
End Sub

Private Sub OnDatumChanged()
    Dim dat As Date
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If mLoading Or mBuilding Then Exit Sub
    If Not TryParseDateValue(FldText("fgDatum"), dat) Then Exit Sub
    If Len(GetActiveStanica()) > 0 Then
        If GetActiveDatum() <> dat Then AcquireStanicaLock GetActiveStanica(), dat
    End If
    ' checkRemote=True kao legacy txtDatum_AfterUpdate: drugi dan = drugi niz,
    ' pa se pita i Google da broj ne bude vec zauzet
    RefreshBrojPredlog True
End Sub

' Datum iz polja; danas kad polje jos nije citljivo.
Private Function DatumIzPolja() As Date
    Dim d As Date
    DatumIzPolja = Date
    On Error Resume Next
    If TryParseDateValue(FldText("fgDatum"), d) Then DatumIzPolja = d
End Function

' Predlog broja dokumenta za tekuci kontekst. Kad je auto-broj iskljucen u
' Podesavanjima, generator vrati prazno i polje ostaje operateru.
Private Sub RefreshBrojPredlog(Optional ByVal checkRemote As Boolean = True)
    Dim entID As String, kind As String, sug As String, mk As String
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    mk = modeKey(ActiveMode)

    ' PRIJEMNICA ne ide kroz SuggestNextBroj - ima svoj generator (niz po kupcu,
    ' x-deo fiksno "1"), i to samo za hladnjaca-kupca.
    If mk = "PRIJEMNICA" Then PredlogPrijemnice: Exit Sub

    kind = KindZaRezim(mk)
    If Len(kind) = 0 Then Exit Sub
    entID = EntitetZaBroj(kind)
    If Len(entID) = 0 Then Exit Sub
    sug = SuggestNextBroj(kind, entID, DatumIzPolja(), checkRemote)
    If Len(sug) = 0 Then Exit Sub
    SetFld "fgBrOtpr", sug
End Sub

' Koji brojevni niz pripada rezimu. Rezimi koji nemaju svoj niz vracaju prazno i
' predlog se ne racuna - broj tada unosi operater:
'   AMB_ISPLATE / AMB_UPLATE (gotovinski promet)  -> slobodan unos
'   PRIJEMNICA vec je obradjena pre ovog poziva   -> vidi PredlogPrijemnice
'   STORNO nije nov dokument                      -> nema broj
' REVERSI su ovde od v6-ui-112: modeKey vraca "REVERSI", a stajalo je "REVERS",
' pa se uslov nikad nije poklopio i revers nije dobijao broj. Niz mu je isti kao
' otpremnici (stanica/ddmmyy-seq), samo se skenira tblAmbalaza.
Private Function KindZaRezim(ByVal key As String) As String
    Select Case key
        Case "OTKUP":      KindZaRezim = KIND_OTK
        Case "OTPREMNICA": KindZaRezim = KIND_OTP
        Case "ZBIRNA":     KindZaRezim = KIND_ZBR
        Case "REVERSI":    KindZaRezim = KIND_REV
    End Select
End Function

' CIJI je to niz. Otkupni list, otpremnica i revers se broje po OTKUPNOM MESTU,
' a zbirna po VOZACU (tblZbirna.VozacID) - kao u legacy
' frmDokumenta.RefreshBrojZbirneSuggestion.
'
' Ovo je i razlog zasto se "S" prefiks pojavljivao van zbirnih: ApplyMirrorPrefix
' gleda da li je entitet mirror-vozac (VozacID == StanicaID), pa je stanica
' podmetnuta kao vozac davala "S..." i tamo gde mu nije mesto. Prefiks je sada
' moguc samo u rezimu zbirne, jer samo on i salje vozaca.
Private Function EntitetZaBroj(ByVal kind As String) As String
    Dim ctx As Object, CB As MSForms.ComboBox
    On Error Resume Next
    Set ctx = mFrm.Controls("zCtx")
    If kind = KIND_ZBR Then
        Set CB = ctx.Controls("cbVozac")
    Else
        Set CB = ctx.Controls("cbOM")
    End If
    If CB Is Nothing Then Exit Function
    EntitetZaBroj = GetComboID(CB)
End Function

' PRIJEMNICA: auto-broj SAMO za hladnjaca-kupca (CFG_MALINA_DEFAULT_KUPAC).
' Ostali kupci nose svoj eksterni, nezavisni broj - polje se tada NE dira.
' x-deo je fiksno "1" (konvencija za hladnjacu), a dnevni niz se broji po kupcu;
' sve to zna postojeci GenerateBrojPrijemnice, pa se racun ovde ne ponavlja.
' Legacy par: frmDokumenta.RefreshBrojPrijSuggestion.
Private Sub PredlogPrijemnice()
    Dim CB As MSForms.ComboBox, kupacID As String, hlad As String, sug As String
    On Error Resume Next
    Set CB = mFrm.Controls("zCtx").Controls("cbKupac")
    If CB Is Nothing Then Exit Sub
    kupacID = GetComboID(CB)
    If Len(kupacID) = 0 Then Exit Sub
    hlad = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(hlad) = 0 Then Exit Sub
    If StrComp(kupacID, hlad, vbTextCompare) <> 0 Then Exit Sub
    sug = GenerateBrojPrijemnice(kupacID, DatumIzPolja())
    If Len(sug) > 0 Then SetFld "fgBrOtpr", sug
End Sub

' PODRAZUMEVANI PROIZVOD iz Podesavanja (DEFAULT_VRSTA_VOCA / DEFAULT_SORTA_VOCA).
'
' Postavlja se SAMO kad je vrsta prazna - isto pravilo koje legacy koristi u
' frmDokumenta ("If cmbVrstaVoca.value = "" Then ApplyDefaultProizvod"). Izbor
' operatera i ono sto je prepisano sa otpremnice se ne diraju.
'
' Postavljanje vrste okida cenu iz cenovnika i tip ambalaze iz kulture, isto
' kao da je operater sam izabrao robu.
Private Sub ApplyDefaultRoba()
    Dim ctx As Object, zf As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set ctx = mFrm.Controls("zCtx")
    Set zf = mFrm.Controls("zForm")
    If Not zf.Controls("fgCena").Visible Then Exit Sub        ' rezim bez robe
    If Len(Trim$(CStr(ctx.Controls("cbVrsta").value))) > 0 Then Exit Sub
    ApplyDefaultProizvod ctx.Controls("cbVrsta"), ctx.Controls("cbSorta")
End Sub

' CENA IZ CENOVNIKA I TIP AMBALAZE IZ KULTURE.
'
' Isti potez koji legacy radi u AutoFillCenaOtkup / AutoFillCenaDok, na svaku
' promenu vrste ili sorte. Cena NIJE podatak dokumenta nego vazeca cena za tu
' robu i klasu (tblCenovnik, append-only istorija), a tip ambalaze pripada
' kulturi - operater ih ne kuca nego ispravlja kad treba.
'
' Postavlja se SAMO ako postoji vrednost: prazan cenovnik ne sme da obrise ono
' sto je operater vec upisao, niti ono sto je prepisano sa otpremnice. Zato
' prefill sa otpremnice ide POSLE ovoga (ista relacija kao u legacy: cenovnik
' je podrazumevano, otpremnica ima poslednju rec).
Private Sub AutoFillCena()
    Dim ctx As Object, zf As Object, vrsta As String, sorta As String
    Dim cI As Double, cII As Double, ta As String
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If mBuilding Then Exit Sub
    Set ctx = mFrm.Controls("zCtx")
    Set zf = mFrm.Controls("zForm")
    ' rezimi bez robe (gotovinski promet) nemaju ni cenu ni ambalazu
    If Not zf.Controls("fgCena").Visible Then Exit Sub

    vrsta = Trim$(CStr(ctx.Controls("cbVrsta").value))
    sorta = Trim$(CStr(ctx.Controls("cbSorta").value))
    If Len(vrsta) = 0 Then Exit Sub

    cI = GetVazecaCena(vrsta, sorta, KLASA_I)
    If cI > 0 Then zf.Controls("fgCena").Controls("fgCena1T").text = Format$(cI, "0.######")
    If mKlasa = 2 Then
        cII = GetVazecaCena(vrsta, sorta, KLASA_II)
        If cII > 0 Then zf.Controls("fgCena").Controls("fgCena2T").text = Format$(cII, "0.######")
    End If

    ta = GetKulturaTipAmbalaze(vrsta, sorta)
    If Len(ta) > 0 Then zf.Controls("fgTipAmb").Controls("fgTipAmbT").text = ta

    RefreshPaletaInfo
    RecalcVrednost
End Sub

' "Paleta br. 186/2026: jos 23 gajbi do zatvaranja (217/240)" - za izabranu
' robu i klasu. Racun radi postojeci GajbeDoZatvaranjaPaleteInfo; kad je
' paletiranje iskljuceno, on vraca prazno pa reda nema.
Private Sub RefreshPaletaInfo()
    Dim ctx As Object, vrsta As String, sorta As String, info As String, i2 As String
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set ctx = mFrm.Controls("zCtx")
    vrsta = Trim$(CStr(ctx.Controls("cbVrsta").value))
    sorta = Trim$(CStr(ctx.Controls("cbSorta").value))
    If Len(vrsta) > 0 Then
        info = GajbeDoZatvaranjaPaleteInfo(vrsta, sorta, KLASA_I)
        If mKlasa = 2 Then
            i2 = GajbeDoZatvaranjaPaleteInfo(vrsta, sorta, KLASA_II)
            If Len(i2) > 0 Then info = info & "   " & ChrW(183) & "   II: " & i2
        End If
    End If
    ctx.Controls("ctxPal").caption = info
End Sub

' PREPISIVANJE POZNATOG (prefill). Ekran zna sta je otpremnica i sta se sa nje
' moze prepisati; ljuska zna gde ta polja stoje. Zato ekran salje LOGICKA imena
' ("vrsta", "omid", "cena"), a ne imena kontrola:
'
'   datum brdok brzbirne vrsta sorta omid vozacid cena tipamb
'
' Sve ide pod mLoading/mPopMute: prepisivanje nije izmena operatera, pa ne sme
' ni da upali "nesacuvane izmene" ni da otvori nas dropdown.
Public Sub ApplyPrefill(ByVal spec As String)
    Dim par As Variant, kv As Variant, k As String, v As String
    Dim ctx As Object, zf As Object
    ' Tipizirani lokali: SetComboByID prima MSForms.ComboBox, a Controls(...)
    ' vraca Object - direktno prosledjivanje ume da padne na neslaganju tipa.
    ' Procedure-level "As MSForms." je bezbedno (IsHardModuleBody gleda samo
    ' modul-level deo, pa modul ostaje "mek" za self-update).
    Dim cbOM As MSForms.ComboBox, cbVoz As MSForms.ComboBox
    Dim cbPart As MSForms.ComboBox, cbPar As MSForms.ComboBox
    Dim fokus As String, imaBroj As Boolean
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    If Len(spec) = 0 Then Exit Sub
    Set ctx = mFrm.Controls("zCtx")
    Set zf = mFrm.Controls("zForm")
    Set cbOM = ctx.Controls("cbOM")
    Set cbVoz = ctx.Controls("cbVozac")
    Set cbPart = ctx.Controls("cbKupac")
    Set cbPar = zf.Controls("fgParcela").Controls("fgParcelaT")

    mPopMute = True
    mLoading = True
    ' Izabrana otpremnica postavlja CEO svoj kontekst, pa se zatecena zbirna
    ' brise pre prepisa: otpremnica koja zbirnu nema ne sme da nasledi zbirnu
    ' prethodne. (Datum i broj dokumenta uvek stizu u spec-u, pa oni nemaju
    ' ovaj problem.)
    zf.Controls("fgBrZbir").Controls("fgBrZbirT").text = ""
    For Each par In Split(spec, "|")
        kv = Split(CStr(par), "=")
        If UBound(kv) >= 1 Then
            k = CStr(kv(0))
            v = CStr(kv(1))
            Select Case k
                Case "datum":    zf.Controls("fgDatum").Controls("fgDatumT").text = v
                Case "brdok"
                    If Len(v) > 0 Then
                        zf.Controls("fgBrOtpr").Controls("fgBrOtprT").text = v
                        imaBroj = True
                    End If
                Case "brzbirne": zf.Controls("fgBrZbir").Controls("fgBrZbirT").text = v
                Case "cena":     If Len(v) > 0 Then zf.Controls("fgCena").Controls("fgCena1T").text = v
                Case "tipamb":   If Len(v) > 0 Then zf.Controls("fgTipAmb").Controls("fgTipAmbT").text = v
                Case "vrsta"
                    ctx.Controls("cbVrsta").value = v
                    ' sorta zavisi od vrste - lista se puni tek posle nje
                    RefillSorta mFrm
                Case "sorta":    ctx.Controls("cbSorta").value = v
                Case "omid":     If Len(v) > 0 Then SetComboByID cbOM, v
                Case "vozacid":  If Len(v) > 0 Then SetComboByID cbVoz, v
                ' --- ispravka posle storna (Z10) salje i ostatak dokumenta ---
                ' Partner: kooperant na otkupu, kupac na zbirnoj i prijemnici.
                ' Parcele zavise od njega, pa se lista puni odmah - inace bi
                ' "parcela" ispod imala praznu listu u kojoj nema sta da nadje.
                Case "partnerid"
                    If Len(v) > 0 Then
                        SetComboByID cbPart, v
                        FillParcele mFrm
                    End If
                ' Druga klasa PRE svojih polja: SetKlasa ih tek otkriva
                ' (fgKgII / fgKolAmbII su sakriveni u jednoklasnom dokumentu).
                Case "dveklase": SetKlasa CLng(val(v))
                Case "kol1":     If Len(v) > 0 Then zf.Controls("fgKgI").Controls("fgKgIT").text = v
                Case "kol2":     If Len(v) > 0 Then zf.Controls("fgKgII").Controls("fgKgIIT").text = v
                Case "amb1":     If Len(v) > 0 Then zf.Controls("fgKolAmb").Controls("fgKolAmbT").text = v
                Case "amb2":     If Len(v) > 0 Then zf.Controls("fgKolAmbII").Controls("fgKolAmbIIT").text = v
                Case "cena2":    If Len(v) > 0 Then zf.Controls("fgCena").Controls("fgCena2T").text = v
                Case "ambpr":    If Len(v) > 0 Then zf.Controls("fgAmbPr").Controls("fgAmbPrT").text = v
                Case "parcela":  If Len(v) > 0 Then SetComboByID cbPar, v
                Case "fokus":    fokus = v
            End Select
        End If
    Next par

    RecalcVrednost
    mPopMute = False
    mLoading = False

    ' BROJ DOKUMENTA. Prefill ga ne donosi kad novi dokument mora da dobije SVOJ
    ' broj -- kod ispravke posle storna to je pravilo, ne propust (stari broj
    ' pripada storniranom dokumentu). Ali predlog se do sada nije ni racunao:
    ' RefreshBrojPredlog visi o promeni stanice ili datuma, a prefill oba
    ' postavlja pod "mLoading = True", pa se nijedan event ne okine. Operater je
    ' zato dobijao popunjenu formu sa PRAZNIM brojem otpremnice.
    '
    ' Racuna se tek OVDE, posle "mLoading = False": generator cita stanicu i
    ' datum iz POLJA, a ona su popunjena tek na kraju petlje. SelectModeCore ga
    ' racuna ranije (RefreshBrojPredlog False), ali tada stanice jos nema --
    ' EntitetZaBroj vrati prazno i predlog se preskoci.
    '
    ' Remote provera se PRESKACE u testu: SuggestNextBroj sa checkRemote pita
    ' Google, a suite ne sme da zavisi od mreze (isti gard kao kod SetFocus-a).
    ' U produkciji ostaje ukljucena -- broj dokumenta je poslovni podatak, i
    ' lokalni sken ga na drugoj masini moze predloziti dvaput.
    If Not imaBroj Then RefreshBrojPredlog (Not IsTestMode())

    MarkClean
    ' Kuda fokus: kod prefilla sa otpremnice (F1) partner NIJE popunjen, pa
    ' kursor ide na njega - to je jedino sto operateru ostaje. Kod ispravke
    ' posle storna popunjeno je i to, pa ide na kolicinu: greska koja se
    ' ispravlja je skoro uvek u njoj.
    ' IsTestMode gard je isti kao u ClearForm: forma koja nije .Show-ovana ne
    ' moze da primi fokus, a u nevidljivom Excelu SetFocus ne puca nego
    ' TRAJNO visi (modTestMode).
    If IsTestMode() Then Exit Sub
    If fokus = "kolicina" Then
        zf.Controls("fgKgI").Controls("fgKgIT").SetFocus
    Else
        ctx.Controls("cbKupac").SetFocus
    End If
End Sub

' Datum dokumenta: podrazumevano danas. Ista vrednost koju je ekran do sada
' cutke pretpostavljao - sada se vidi i moze da se promeni.
Private Sub SetDatumDanas(z As Object)
    On Error Resume Next
    z.Controls("fgDatum").Controls("fgDatumT").text = Format$(Date, "dd.mm.yyyy")
End Sub

' Datum iz polja kao serijski broj; 0 znaci "necitljivo". Trailing tacka
' ("11.08.2026.") je nacin na koji se datum kod nas pise, pa se skida pre
' provere umesto da obori unos.
'
' TEST SEAM: Public, ne Private -- modTest.T_ParseDatum_Ugovor je zove direktno
' (isti razlog kao frmOtkup.ClearOtkupFields; .claude/rules/otkup-i-dokumenta.md
' odeljak 2). Poziva je samo ovaj modul i test.
Public Function ParseDatum(ByVal s As String) As Double
    Dim t As String, d As Date
    On Error Resume Next
    t = Trim$(s)
    Do While Right$(t, 1) = "."
        t = Left$(t, Len(t) - 1)
    Loop
    If Len(t) = 0 Then Exit Function
    ' SAMO deterministicki parser. IsDate/CDate citaju po Windows locale-u, pa
    ' bi isti tekst "01.02.2026" na MDY masini bio 2. januar, a na DMY masini
    ' 1. februar - dok predlog broja i zakljucavanje stanice isti tekst vec
    ' citaju kroz TryParseDateValue. Upis i kontekst NE SMEJU da se raziidju
    ' (ista klasa greske kao AUD-007 / RF-09).
    If TryParseDateValue(t, d) Then ParseDatum = CDbl(d)
End Function

' UGOVOR, ne "ciscenje forme". Ista tri ponasanja i ista tri razloga kao
' frmOtkup.ClearOtkupFields (.claude/rules/otkup-i-dokumenta.md odeljak 1 i 5):
' datum i broj zbirne su KONTEKST otpremnice i ostaju, partner se brise.
'
' TEST SEAM: Public, ne Private -- modTest.T_ClearForm_Ugovor je zove direktno,
' bez voznje celog upisa (koji trazi stanica-lock, PDF izlaz i auto-lanac
' hladnjace). Pre izmene ove rutine procitaj sta test tvrdi.
Public Sub ClearForm()
    Dim nmv As Variant, i As Long, imaOtp As Boolean
    On Error Resume Next
    mPopMute = True
    mLoading = True
    ' BROJ ZBIRNE se NE prazni - on je kontekst, kao i datum: svi blokovi jedne
    ' otpremnice idu na istu zbirnu. Legacy ClearOtkupFields ga takodje ne dira
    ' (cisti ga tek ResetDatumKontekst, kad se otpremnica napusti). Uz to bi
    ' praznjenje polja okinulo SetAktivnaZbirna "" i obrisalo zapamcenu zbirnu.
    ' fgNovac je u istom spisku: iznos pripada JEDNOM dokumentu, a zatecen iznos
    ' u polju posle snimanja je najskuplja greska ovog rezima - sledeca potvrda
    ' bi isplatila isti novac drugi put.
    nmv = Array("fgBrOtpr", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr", "fgNovac")
    For i = 0 To UBound(nmv)
        mFrm.Controls("zForm").Controls(CStr(nmv(i))).Controls(CStr(nmv(i)) & "T").text = ""
    Next i
    mFrm.Controls("zForm").Controls("fgFaktura").Controls("fgFakturaT").text = ""
    ' cena ima dve kutije pa ne staje u sablon grp -> grp & "T"
    mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena1T").text = ""
    mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena2T").text = ""
    mFrm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT").text = ""
    mFrm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT").text = ""
    ' PARTNER se prazni kao i sve ostalo - sledeci dokument je nov kooperant
    ' (legacy ClearOtkupFields: cmbKooperant.value = "", cmbParcela.Clear). Bez
    ' ovoga zatecen kooperant ostaje u polju, pa bi unos samo kilograma tiho
    ' otisao na prethodnog. Prazna vrednost okida i FillOpenBlokovi/FillParcele,
    ' koji na prazan partner ciste svoje liste.
    mFrm.Controls("zCtx").Controls("cbKupac").value = ""
    ' DATUM: dok je otpremnica aktivna, SVI njeni blokovi nose NJEN datum -
    ' vracanje na danas bi drugi i svaki sledeci blok upisalo pod danasnjim
    ' datumom, a RefreshBrojPredlog bi mu dao i danasnji broj (otpremnica
    ' 8/220726 od 22.07 dobijala je blok 8/110826 od 11.08). Legacy
    ' ClearOtkupFields datum uopste ne dira. Bez otpremnice ostaje staro
    ' ponasanje: danas, jer prazno polje bi bilo greska koju operater mora da
    ' ispravi pri svakom novom dokumentu.
    imaOtp = ImaAktivnuOtpremnicu()
    If Not imaOtp Then SetDatumDanas mFrm.Controls("zForm")
    If ParseDatum(FldText("fgDatum")) = 0 Then SetDatumDanas mFrm.Controls("zForm")
    SetOstatak 0
    SetKlasa 1
    ' Smer reversa i "isplata iz" su izbori JEDNOG dokumenta, kao i klasa:
    ' zadrzan smer bi sledeci revers proknjizio u suprotnom pravcu.
    SetSmerRev 0
    SetAvans 1
    ' Posle isplate/uplate saldo avansa i lista otvorenih faktura vise nisu
    ' tacni - citaju se ponovo, da sledeci dokument ne racuna po starom stanju.
    RefreshSaldoOM
    FillOpenFakture mFrm
    ' Cena je obrisana zajedno sa poljima, a vrsta/sorta su ostale izabrane -
    ' pa se cena i tip ambalaze vracaju iz cenovnika i kulture. Ako robe nema
    ' (prvi unos), uzima se podrazumevana iz Podesavanja.
    ResetFokusOkvira
    ApplyDefaultRoba
    AutoFillCena
    ' sledeci dokument dobija sledeci broj - bez upita na Google (lokalno je
    ' vec upisan red koji je upravo snimljen)
    RefreshBrojPredlog False
    RecalcVrednost
    mPopMute = False
    mLoading = False
    MarkClean
    ' Dok je otpremnica aktivna, sledecem bloku operater unosi samo kooperanta
    ' i kolicine - zato kursor ide pravo na kooperanta, isto kao posle prefilla
    ' i kao u legacy ClearOtkupFields (cmbKooperant.SetFocus).
    '
    ' IsTestMode gard je isti kao u frmOtkup.ClearOtkupFields: forma koja nije
    ' .Show-ovana ne moze da primi fokus, a u nevidljivom Excelu SetFocus ne
    ' puca nego TRAJNO visi (modTestMode). U produkciji je IsTestMode() uvek
    ' False -- flag postavlja iskljucivo test modul -- pa je ponasanje isto.
    If imaOtp Then
        If Not IsTestMode() Then mFrm.Controls("zCtx").Controls("cbKupac").SetFocus
        Exit Sub
    End If
    If Len(mPrvoPolje) = 0 Then mPrvoPolje = "fgBrOtpr"
    If Not IsTestMode() Then _
        mFrm.Controls("zForm").Controls(mPrvoPolje).Controls(mPrvoPolje & "T").SetFocus
End Sub

' Da li je otpremnica jos izabrana. Ista pitalica koju koristi traka otpremnice
' (Scr_OtpInfo vraca prazno kad otpremnice nema), kasno vezano - ljuska ne sme
' da rano vezuje modul ekrana (zamka #19).
Private Function ImaAktivnuOtpremnicu() As Boolean
    Dim info As String
    On Error Resume Next
    If mScreen <> "DOKUMENTI" Then Exit Function
    info = CStr(Application.Run("modScrDokumenti.Scr_OtpInfo"))
    Err.Clear
    ImaAktivnuOtpremnicu = (Len(info) > 0)
End Function

' Vrednost je zbir po klasama - svaka klasa ima svoju cenu, pa se kilogrami
' II klase NE mogu sabrati sa kilogramima I klase pre mnozenja.
Private Sub RecalcVrednost()
    Dim v1 As Double, v2 As Double
    On Error Resume Next
    v1 = ParseNum(FldText("fgKgI")) * CenaText(1)
    If mKlasa = 2 Then v2 = ParseNum(FldText("fgKgII")) * CenaText(2)
    SetCalcLine 1, Poruka("OTKUI_SEG_KLASA_I"), ParseNum(FldText("fgKgI")), CenaText(1), (v1 > 0)
    SetCalcLine 2, Poruka("OTKUI_SEG_KLASA_II"), ParseNum(FldText("fgKgII")), CenaText(2), _
                (mKlasa = 2 And v2 > 0)
    SetKgLine
    PaintVrednost v1 + v2
End Sub

' ZIVI ZBIR KILOGRAMA. Dva slucaja, isti kao u legacy UpdateUkupnoKg:
'
'   bruto rezim (OTKUP_BRUTO_UNOS)  ->  neto posle tare, sa racunom kako je
'                                       dobijen; upisuje se NETO, pa operater
'                                       mora da ga vidi pre snimanja
'   dve klase                       ->  zbir I + II
'
' Van toga red je prazan: jedna klasa u neto rezimu nema sta da sabira.
Private Sub SetKgLine()
    Dim L As Object, kgI As Double, kgII As Double, amb As Double, tara As Double
    Dim tip As String, cap As String
    On Error Resume Next
    Set L = mFrm.Controls("zForm").Controls("fgVrednost").Controls("fgVrednostKg")
    If L Is Nothing Then Exit Sub
    kgI = ParseNum(FldText("fgKgI"))
    If mKlasa = 2 Then kgII = ParseNum(FldText("fgKgII"))

    If OtkupBrutoUnos() Then
        amb = ParseNum(FldText("fgKolAmb"))
        tip = Trim$(FldText("fgTipAmb"))
        If kgI > 0 And amb > 0 And Len(tip) > 0 Then
            tara = amb * GetTezinaGajbice(tip)
            If tara > 0 And tara < kgI Then
                cap = Poruka("OTKUI_KG_NETO") & " " & FmtBroj(kgI - tara, 2) & " " & _
                      Poruka("OTKUI_UNIT_KG") & "   (" & FmtBroj(kgI, 2) & " " & _
                      ChrW(8722) & " " & FmtBroj(tara, 2) & ")"
            End If
        End If
    End If

    If Len(cap) = 0 And mKlasa = 2 And (kgI > 0 Or kgII > 0) Then
        cap = Poruka("OTKUI_KG_UKUPNO") & " " & FmtBroj(kgI + kgII, 2) & " " & _
              Poruka("OTKUI_UNIT_KG") & "   (" & FmtBroj(kgI, 2) & " + " & _
              FmtBroj(kgII, 2) & ")"
    End If
    L.caption = cap
End Sub

' "I klasa  3.200 kg x 48,50" - jedan red po klasi, jer sa dve cene jedan
' proizvod vise ne opisuje iznos. Cena se UVEK unosi sa PDV-om; razlaganje na
' osnovicu i porez radi obracun, ne ovaj ekran - zato ovde nema napomene o PDV-u.
Private Sub SetCalcLine(ByVal idx As Long, ByVal klasa As String, ByVal kg As Double, _
                        ByVal cena As Double, ByVal show As Boolean)
    Dim L As Object
    On Error Resume Next
    Set L = mFrm.Controls("zForm").Controls("fgVrednost").Controls("fgVrednostK" & idx)
    If Not show Then
        L.caption = ""
        Exit Sub
    End If
    L.caption = klasa & "   " & FmtBroj(kg, 0) & " " & Poruka("OTKUI_UNIT_KG") & _
                "  " & ChrW(215) & "  " & FmtBroj(cena, 2)
End Sub

' Prazan iznos ne sme da bude najjaci element na ekranu: dok nema vrednosti
' ploca je samo tanak okvir sa crticom, a pun forest sa zlatnom ivicom dobija
' tek kad ima sta da pokaze.
Private Sub PaintVrednost(ByVal v As Double)
    Dim fr As Object, on_ As Boolean
    On Error Resume Next
    Set fr = mFrm.Controls("zForm").Controls("fgVrednost")
    on_ = (v > 0)
    fr.Controls("fgVrednostB").BackColor = IIf(on_, C_FOREST, C_BORDER_LT)
    fr.Controls("fgVrednostF").BackColor = IIf(on_, C_FOREST, C_WHITE)
    fr.Controls("fgVrednostA").BackColor = IIf(on_, C_GOLD, C_WHITE)
    fr.Controls("fgVrednostL").ForeColor = IIf(on_, RGB(166, 178, 160), C_MUTED)
    fr.Controls("fgVrednostV").ForeColor = IIf(on_, C_CREAM, C_MUTED)
    fr.Controls("fgVrednostU").ForeColor = IIf(on_, C_GOLD, C_MUTED)
    fr.Controls("fgVrednostV").caption = IIf(on_, FmtBroj(v, 0), ChrW(8212))
End Sub

' Cena klase iz svoje kutije (fgCena1T / fgCena2T). Ako II klasa nije uneta,
' vazi cena I klase - u praksi se cesto radi po jednoj ceni.
Private Function CenaText(ByVal klasa As Long) As Double
    Dim c As Double
    On Error Resume Next
    c = ParseNum(mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena" & klasa & "T").text)
    If klasa = 2 And c <= 0 Then
        c = ParseNum(mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena1T").text)
    End If
    CenaText = c
End Function

' Prepisuje kolicinu predate pune ambalaze u polje prazne - razmena je
' jedan-za-jedan pravilo, pa je rucno prekucavanje suvisan korak.
Private Sub CopyAmbNaPraznu()
    Dim z As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set z = mFrm.Controls("zForm")
    If Not z.Controls("fgAmbPr").Visible Then Exit Sub
    z.Controls("fgAmbPr").Controls("fgAmbPrT").text = _
        z.Controls("fgKolAmb").Controls("fgKolAmbT").text
End Sub


' Kljuc filtera iza cipa. Slot koji je pozajmio ekran vraca EKRANOV kljuc --
' ljuska ga ne tumaci, samo ga prosledi kroz Scr_Rows.
Private Function ChipFilter(ByVal tag As String) As String
    Dim ch As Variant, i As Long
    If mScrChipN > 0 Then
        ch = ChipRow()
        For i = 0 To mScrChipN - 1
            If Split(CStr(ch(i)), "|")(0) = tag Then
                ChipFilter = mScrChipKey(i)
                Exit Function
            End If
        Next i
    End If
    Select Case tag
        Case "chipDanas":     ChipFilter = "danas"
        Case "chipNedelja":   ChipFilter = "nedelja"
        Case "chipSve":       ChipFilter = "sve"
        Case "chipNefakt":    ChipFilter = "nefakt"
        Case "chipBezZbirne": ChipFilter = "bezzbirne"
        Case "chipOtkazane":  ChipFilter = "otkazane"
        Case "chipOtvorene":  ChipFilter = "otvorene"
        Case Else:            ChipFilter = "sve"
    End Select
End Function

Private Sub PagerClick(ByVal tag As String)
    Dim pages As Long
    pages = 1
    If mPageSize > 0 Then pages = -Int(-mViewN / mPageSize)
    If pages < 1 Then pages = 1
    Select Case tag
        Case "pgPrev": If mPage > 1 Then mPage = mPage - 1
        Case "pgNext": If mPage < pages Then mPage = mPage + 1
        Case Else
            Dim t As String
            t = CStr(mFrm.Controls("zGrid").Controls("grdFoot").Controls(tag).tag)
            If Left$(t, 3) = "pg:" Then mPage = CLng(Mid$(t, 4))
    End Select
    RenderGrid
End Sub

Private Sub EnsureBtns()
    If Btns Is Nothing Then Set Btns = New Collection
End Sub

Public Sub WireBtn(ctl As Object, ByVal tag As String, ByVal kind As String)
    Dim b As clsFlatBtn
    Set b = New clsFlatBtn
    b.Bind ctl, kind, tag
    EnsureBtns
    Btns.Add b
End Sub

Public Sub WireInput(ctl As Object, ByVal tag As String)
    Dim b As clsFlatBtn
    Set b = New clsFlatBtn
    b.Bind ctl, "input", tag
    EnsureBtns
    Btns.Add b
End Sub

Private Sub WireZone(fr As Object)
    Dim b As clsFlatBtn
    Set b = New clsFlatBtn
    b.Bind fr, "zone", fr.name
    EnsureBtns
    Btns.Add b
End Sub

' Alt+F8 -> DumpMdl2Sheet
' Ispise sve glifove fonta "Segoe MDL2 Assets" sa kodovima u NOVU radnu svesku
' (ne dira AgriX). Prazan kvadratic = glif ne postoji na ovoj verziji Windowsa.
' Izabrane kodove prepisi u IC_* konstante na vrhu ovog modula.
Public Sub DumpMdl2Sheet()
    ' Segoe MDL2 Assets ide do ~F8B3, ne do EA00 - ranija verzija je ispisivala
    ' samo prvih ~770 glifova od preko 4.500.
    Const FIRST_C As Long = &HE700&
    Const LAST_C As Long = &HF8B3&
    Const PER_ROW As Long = 10
    Dim wb As Workbook, ws As Worksheet
    Dim arr() As Variant, n As Long, rows_ As Long
    Dim r As Long, c As Long, code As Long

    n = LAST_C - FIRST_C + 1
    rows_ = -Int(-n / PER_ROW)
    ReDim arr(1 To rows_, 1 To PER_ROW * 2)

    r = 1: c = 1
    For code = FIRST_C To LAST_C
        arr(r, c) = ChrW(code)
        arr(r, c + 1) = "&H" & Hex$(code) & "&"
        c = c + 2
        If c > PER_ROW * 2 Then
            c = 1
            r = r + 1
        End If
    Next code

    Application.ScreenUpdating = False
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Range(ws.CellS(1, 1), ws.CellS(rows_, PER_ROW * 2)).value = arr
    For c = 1 To PER_ROW * 2 Step 2
        With ws.columns(c)
            .Font.name = "Segoe MDL2 Assets"
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
        End With
        With ws.columns(c + 1)
            .Font.name = "Consolas"
            .Font.Size = 7
            .ColumnWidth = 7
        End With
    Next c
    ws.rows.RowHeight = 24
    ws.CellS.VerticalAlignment = xlCenter
    ws.CellS(1, 1).Select
    Application.ScreenUpdating = True

    MsgBox "Ispisano " & n & " glifova u novu radnu svesku." & vbCrLf & _
           "Prazan kvadratic = glif ne postoji na ovoj verziji Windowsa." & vbCrLf & vbCrLf & _
           "Izabrane kodove prepisi u IC_* konstante u modOtkupUI.", vbInformation
End Sub

Public Sub DiagOtkupUI()
    Debug.Print String(58, "=")
    Debug.Print "modOtkupUI build : " & OTKUI_BUILD
    Debug.Print "korisnik         : [" & GetCurrentUserIme() & "] / [" & GetCurrentUser() & "]"
    Debug.Print String(58, "-")
    DiagList "GetLookupList(tblKulture,VrstaVoca)", GetLookupList(TBL_KULTURE, "VrstaVoca", , , True)
    DiagList "GetTipAmbalazeOptions()", GetTipAmbalazeOptions()
    Debug.Print String(58, "-")
    DiagTable TBL_STANICE
    DiagTable TBL_VOZACI
    DiagTable TBL_KUPCI
    DiagTable TBL_KULTURE
    DiagTable TBL_OTKUP
    DiagTable TBL_OTPREMNICA
    DiagTable TBL_PRIJEMNICA
    DiagTable TBL_NOVAC
    DiagTable TBL_AMBALAZA
    Debug.Print String(58, "-")
    DiagCols "F1"
    DiagCols "F2"
    DiagCols "F3"
    DiagCols "F4"
    DiagCols "F5"
    DiagCols "F6"
    DiagCols "F7"
    Debug.Print String(58, "=")
End Sub

Private Sub DiagList(ByVal what As String, ByVal v As Variant)
    Dim n As Long, sample As String
    If IsArray(v) Then
        n = UBound(v) - LBound(v) + 1
        If n > 0 Then sample = "  prvi=[" & CStr(v(LBound(v))) & "]"
    End If
    Debug.Print Left$(what & String(38, "."), 38) & " " & n & " stavki" & sample
End Sub

Private Sub DiagTable(ByVal tbl As String)
    Dim v As Variant, n As Long, cols As Long
    v = GetTableData(tbl)
    If IsArray(v) Then n = UBound(v, 1): cols = UBound(v, 2)
    Debug.Print Left$(tbl & String(38, "."), 38) & " " & n & " redova, " & cols & " kolona" & _
                IIf(n = 0, "   <-- PRAZNO ili tabela nije nadjena", "")
End Sub

' Po REZIMU: ispisuje SVE kolone koje rezim trazi i njihove indekse u tabeli.
' "(!)" znaci da kolona tog imena NE POSTOJI u tabeli - schema drift.
Private Sub DiagCols(ByVal mode As String)
    Dim cols As Variant, i As Long, s As String, srcCol As String, ix As Long
    Dim mk As String, tbl As String
    mk = modeKey(mode): tbl = ModeTable(mode)
    cols = GridCols(mk)
    For i = 0 To UBound(cols)
        srcCol = ColF(CStr(cols(i)), 1)
        If Len(srcCol) > 0 Then
            ix = ColIdx(tbl, srcCol)
            s = s & srcCol & "=" & ix & IIf(ix = 0, "(!)", "") & "  "
        End If
    Next i
    ix = ColIdx(tbl, COL_STORNIRANO)
    Debug.Print mode & " " & Left$(mk & String(12, " "), 12) & " " & _
                Left$(tbl & String(16, " "), 16) & s & _
                "Stornirano=" & ix & IIf(ix = 0, "(!)", "")
End Sub

'=====================================================================
' POSLEDNJA PROCEDURA U FAJLU - namerno.
' Ako je import/paste presecen, ova rutina NECE postojati, a "Sub or
' Function not defined" na BtnV/NewShell je tada simptom polovicnog
' uvoza, ne greske u kodu. Provera: Immediate -> ?OtkupUI_SelfCheck()
'=====================================================================
' Ispisuje SAMO ikonice koje ekran stvarno koristi: mesto upotrebe, kod,
' ZVANICNO ime glifa i iscrtan glif. Ime je tu da se promasaj vidi bez gledanja
' u sliku - "Palete = CalendarWeek" je ocigledno pogresno i na papiru.
' Zamena: nadji red, uzmi nov kod iz biraca ili DumpMdl2Sheet, javi "slot = kod".
' PAZI: VBA dozvoljava najvise 25 nastavaka linije (" _") po JEDNOJ naredbi.
' Ovaj spisak je bio jedan Array( ... ) sa 26 stavki i oborio je kompajler sa
' "Too many line continuations". Zato se sada gradi kroz Collection - dodavanje
' novog slota je nova naredba i limit se ne moze ponovo probiti.
Public Sub DumpMdl2Used()
    Dim wb As Workbook, ws As Worksheet, r As Long, i As Long
    Dim a As Collection
    Set a = New Collection
    IcoRow a, "IC_OTKUP", "Sidebar: Otkup i prodaja", "QuickNote", IC_OTKUP
    IcoRow a, "IC_BLOKOVI", "Sidebar: Otkupni blokovi", "Document", IC_BLOKOVI
    IcoRow a, "IC_PALETE", "Sidebar: Palete", "Library", IC_PALETE
    IcoRow a, "IC_AGRO", "Sidebar: Agrohemija", "Leaf", IC_AGRO
    IcoRow a, "IC_FAKT", "Sidebar: Fakturisanje", "PaymentCard", IC_FAKT
    IcoRow a, "IC_BANKA", "(nije vezano - Banka je razlozena)", "Bank", IC_BANKA
    IcoRow a, "IC_ISPLATA", "Naslov F5: Isplate", "Upload", IC_ISPLATA
    IcoRow a, "IC_UPLATA", "Naslov F6: Uplate", "Download", IC_UPLATA
    IcoRow a, "IC_UVOZ", "Sidebar: Uvoz izvoda", "PDF", IC_UVOZ
    IcoRow a, "IC_NALOZI", "Sidebar: Platni nalozi", "CheckList", IC_NALOZI
    IcoRow a, "IC_MARZA", "Sidebar: Marza", "AreaChart", IC_MARZA
    IcoRow a, "IC_IZVEST", "Sidebar: Izvestaji", "ReportDocument", IC_IZVEST
    IcoRow a, "IC_SLEDLJ", "Sidebar: Sledljivost", "Link", IC_SLEDLJ
    IcoRow a, "IC_SAVE", "Dugme: Sacuvaj", "Save", IC_SAVE
    IcoRow a, "IC_PRINT", "Dugme: Stampa", "Print", IC_PRINT
    IcoRow a, "IC_CLOSE", "Dugme: Zatvori", "Cancel", IC_CLOSE
    IcoRow a, "IC_FILTER", "Dugme: Filteri", "Filter", IC_FILTER
    IcoRow a, "IC_SEARCH", "Polje pretrage", "Zoom", IC_SEARCH
    IcoRow a, "IC_SYNC", "Desna kolona: stanje sinhronizacije", "Cloud", IC_SYNC
    IcoRow a, "IC_HELP", "Status traka: pomoc", "Help", IC_HELP
    IcoRow a, "IC_MATICNI", "Dugme: Maticni podaci", "Setting", IC_MATICNI
    IcoRow a, "IC_SYNCBTN", "Dugme: Sinhronizuj", "Sync", IC_SYNCBTN
    IcoRow a, "IC_SNIMI", "Dugme: Snimi", "SaveLocal", IC_SNIMI
    IcoRow a, "IC_EXCEL", "Dugme: Otvori Excel", "OpenInNewWindow", IC_EXCEL
    IcoRow a, "IC_OPERATER", "Dugme: Operater", "SwitchUser", IC_OPERATER
    IcoRow a, "IC_STORNO", "Naslov F8: Storno", "Undo", IC_STORNO
    IcoRow a, "IC_REVERS", "Naslov F7: Reversi (PRIVREMENO)", "Library", IC_REVERS
    IcoRow a, "IC_ENTER", "(nije vezano)", "ReturnKey", IC_ENTER
    IcoRow a, "IC_OPORAVAK", "Naslov: Oporavak", "Refresh", IC_OPORAVAK
    IcoRow a, "IC_MAT_PARTNERI", "Sidebar: Partneri", "People", IC_MAT_PARTNERI
    IcoRow a, "IC_MAT_ROBA", "Sidebar: Proizvodi i cene", "Tag", IC_MAT_ROBA
    IcoRow a, "IC_MAT_PAKOVANJE", "Sidebar: Ambalaza i pakovanje", "Package", IC_MAT_PAKOVANJE
    IcoRow a, "IC_MAT_KORISNICI", "Sidebar: Korisnici", "SwitchUser", IC_MAT_KORISNICI
    IcoRow a, "IC_MAT_SISTEM", "Sidebar: Podesavanja", "Setting", IC_MAT_SISTEM
    IcoRow a, "IC_MAT_ADMIN", "Sidebar: Administracija", "Admin", IC_MAT_ADMIN
    IcoRow a, "IC_NAZAD", "Dugme: Nazad na rad", "Back", IC_NAZAD

    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Range("A1:E1").value = Array("KONSTANTA", "GDE SE KORISTI", "MDL2 IME", "KOD", "GLIF")
    ws.Range("A1:E1").Font.bold = True
    For i = 1 To a.count                     ' Collection je 1-bazirana
        Dim p As Variant: p = Split(CStr(a(i)), "|")
        r = i + 1
        ws.CellS(r, 1).value = p(0)
        ws.CellS(r, 2).value = p(1)
        ws.CellS(r, 3).value = p(2)
        If CLng(p(3)) = 0 Then
            ws.CellS(r, 4).value = "0 (bez ikonice)"
        Else
            ws.CellS(r, 4).value = "&H" & Hex$(CLng(p(3))) & "&"
            ws.CellS(r, 5).value = ChrW(CLng(p(3)))
        End If
    Next i
    ws.columns("E").Font.name = F_ICON
    ws.columns("E").Font.Size = 20
    ws.columns("A:E").AutoFit
    ws.rows(1).AutoFilter
End Sub

Private Sub IcoRow(a As Collection, ByVal konst As String, ByVal gde As String, _
                   ByVal mdl2Ime As String, ByVal kod As Long)
    a.Add konst & "|" & gde & "|" & mdl2Ime & "|" & CStr(kod)
End Sub

Public Function OtkupUI_SelfCheck() As String
    ' Provera MORA biti kasno vezana (CallByName). Rano vezan poziv "b.Build"
    ' na STAROJ klasi ne pada u runtime-u nego obara COMPILE celog modula - a
    ' cela poenta ove funkcije je da radi bas kad je klasa stara. Isto pravilo
    ' kao zamka #19 u CLAUDE.md: nov simbol iza kasnog vezivanja.
    '
    ' modUiKit se cita direktno (UIKIT_BUILD): oba modula su "meka" i self-update
    ' ih isporucuje zajedno, a nepotpun uvoz modula je GLASAN (nedostaje simbol
    ' ili "Ambiguous name" zbog modUiKit1) - za razliku od klase, gde stara tiho
    ' nastavi da radi.
    Dim b As Object, cv As String, ver As String
    On Error Resume Next
    Set b = New clsFlatBtn
    cv = CStr(CallByName(b, "Build", VbGet))
    On Error GoTo 0
    ver = "modOtkupUI " & OTKUI_BUILD & " + modUiKit " & UIKIT_BUILD & _
          " + modScrDokumenti " & SCRDOK_BUILD & _
          " + modUiScreens " & UISCR_BUILD & _
          " + modStornoDok " & STORNODOK_BUILD & _
          " + clsFlatBtn " & IIf(Len(cv) = 0, "(stara)", cv)
    If Len(cv) = 0 Then
        OtkupUI_SelfCheck = ver & vbCrLf & _
            "PAZNJA: clsFlatBtn je STARA verzija (uvoz nije zamenio klasu; " & _
            "trazi komponentu clsFlatBtn1 u Project Exploreru)"
    ElseIf StaraKomponenta(cv) Or StaraKomponenta(UIKIT_BUILD) _
           Or StaraKomponenta(SCRDOK_BUILD) Or StaraKomponenta(UISCR_BUILD) _
           Or StaraKomponenta(STORNODOK_BUILD) Then
        OtkupUI_SelfCheck = ver & vbCrLf & _
            "PAZNJA: neki modul je stariji od " & OTKUI_MIN_BUILD & _
            " (uvoz nije prosao do kraja)"
    Else
        OtkupUI_SelfCheck = ver & vbCrLf & OtkupUI_Stats()
    End If
End Function

' Da li je komponenta starija od najstarije podrzane. Verzije se NE porede na
' jednakost: moduli se menjaju nezavisno (clsFlatBtn je zamrznut, ekrani se
' dodaju svojim tempom), pa je ugovor "ne stariji od OTKUI_MIN_BUILD".
' Neprepoznat oblik verzije se ne prijavljuje kao greska - provera je tu da
' uhvati nepotpun uvoz, ne da tumaci tudje sheme oznacavanja.
Private Function StaraKomponenta(ByVal v As String) As Boolean
    Dim a As Long, b As Long
    a = BuildSeq(v)
    b = BuildSeq(OTKUI_MIN_BUILD)
    If a = 0 Or b = 0 Then Exit Function
    StaraKomponenta = (a < b)
End Function

' "v6-ui-113" -> 113. Redni broj je jedini deo koji raste monotono.
Private Function BuildSeq(ByVal v As String) As Long
    Dim i As Long, c As String, d As String
    For i = Len(v) To 1 Step -1
        c = Mid$(v, i, 1)
        If c >= "0" And c <= "9" Then d = c & d Else Exit For
    Next i
    If Len(d) > 0 Then BuildSeq = CLng(d)
End Function

' Merilo rasta ekrana. Plan je da SVI ekrani predju na ovu jednu formu, pa je
' broj kontrola resurs koji se trosi - a granicu MSForms-a ne znamo unapred.
' Zato se meri od pocetka: koliko kontrola i koliko traje gradnja.
Public Function OtkupUI_Stats() As String
    Dim sinks As Long
    On Error Resume Next
    If mFrm Is Nothing Then
        OtkupUI_Stats = "ekran nije izgradjen"
        Exit Function
    End If
    If Not Btns Is Nothing Then sinks = Btns.count
    OtkupUI_Stats = "kontrola: " & mFrm.Controls.count & _
                    "  |  vezanih sinkova: " & sinks & _
                    "  |  gradnja: " & mBuildMs & " ms"
End Function

' Raspodela po zonama. Bez ovoga se ne moze odluciti sta u S3 ostaje deljeno
' (mreza, hrom) a sta svaki ekran nosi sam: prvo mora da se vidi ko trosi.
'
' PAZNJA: UserForm.Controls je RAVNA lista - vec sadrzi i sve sto je unutar
' okvira. Rekurzija kroz Frame.Controls zato broji iste kontrole po vise
' puta (prva verzija ove rutine prijavila je 2703 umesto stvarnog broja).
' Ovde se ide jednim prolazom kroz ravnu listu, a svaka kontrola se pripise
' svojoj zoni preko lanca .Parent.
Public Sub OtkupUI_StatsZone()
    Dim c As Object, d As Object, k As Variant, uk As Long
    On Error Resume Next
    If mFrm Is Nothing Then
        Debug.Print "ekran nije izgradjen"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In mFrm.Controls
        uk = uk + 1
        k = ZoneOf(c)
        d(k) = d(k) + 1
    Next c
    For Each k In d.keys
        Debug.Print Right$("      " & d(k), 6) & "  " & k
    Next k
    Debug.Print Right$("      " & uk, 6) & "  = UKUPNO (ravno, bez duplog brojanja)"
End Sub

' Najspoljasnji okvir u kome kontrola zivi = njena zona. Kontrola koja stoji
' pravo na formi vraca svoje ime (nema zone).
Private Function ZoneOf(ByVal c As Object) As String
    Dim p As Object
    On Error Resume Next
    Set p = c
    Do While TypeName(p.parent) = "Frame"
        Set p = p.parent
    Loop
    ZoneOf = p.name
End Function
'=== KRAJ FAJLA modOtkupUI ===
