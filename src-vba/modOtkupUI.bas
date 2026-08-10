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
Public Const OTKUI_BUILD   As String = "v6-ui-67"

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
Private Const KPI_H       As Single = 64      ' tri reda: naslov, vrednost, kontekst
Private Const TITLE_H     As Single = 46
Private Const CTX_H       As Single = 72      ' eyebrow red + natpis + polje
Private Const CTX_LBL_Y   As Single = 24      ' natpis IZNAD polja
Private Const CTX_FLD_Y   As Single = 38      ' vrh polja u zoni konteksta
Private Const RIGHT_W     As Single = 188
Private Const FIELD_H     As Single = 30
Private Const FIELD_GRP_H As Single = 50
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
Private Const GRID_TOP    As Single = 63     ' 63 + 21 = 84pt = 112 px tacno
Private Const GRID_HEAD_H As Single = 21
Private Const GRID_FOOT_H As Single = 24
Private Const MAX_ROWS    As Long = 22       ' redova mreze koji se PRAVE
Private Const MAX_COLS    As Long = 14       ' kolona mreze koje se PRAVE

' Sifre stanja placanja (odvojene od 0/1/2 koje nosi STATUS pilula).
' MORAJU biti u deklaracionom delu: VBA module-level Const iza prve procedure
' nije vidljiv procedurama iznad njega ("Variable not defined").
Public Const PAY_PLACENO As Long = 10
Public Const PAY_DELIM   As Long = 11
Public Const PAY_NEPLAC  As Long = 12
Private Const PAY_NEFAKT  As Long = 13
Private Const POP_MAX     As Long = 14       ' stavki u sopstvenom dropdown-u
Private Const POP_ITEM_H  As Single = 21
Private Const FLT_W       As Single = 236     ' panel "Filteri"
Private Const FLT_H       As Single = 222
Private Const MODE_SEP_H  As Single = 11      ' razmak iznad kartice Storno
Private Const TIT_ICO_W   As Single = 26      ' marker modula uz naslov
Private Const GRP_H       As Single = 15      ' eyebrow red grupe polja
Private Const GRP_LBL_W   As Single = 108     ' natpis grupe, pa linija do kraja
Private Const SEG_KL_W    As Single = 30      ' prekidac klase u polju KLASA I CENA
Private Const VAL_CALC_W  As Single = 150     ' racunica DESNO od ploce iznosa
Private Const VAL_PLATE_W As Single = 236     ' ploca iznosa (natpis + broj + RSD)
Private Const GAP         As Single = 10
Public Const ICO_INSET   As Single = 24     ' sirina zone ikonice u dugmetu
Private Const PAD         As Single = 16
Private Const STATUS_H    As Single = 24
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
Private Const IC_PALETE   As Long = &HE8F1&   ' Library
Private Const IC_AGRO     As Long = &HE8BE&   ' Leaf
Private Const IC_FAKT     As Long = &HE8C7&   ' PaymentCard
Private Const IC_BANKA    As Long = &HE825&   ' Bank
Private Const IC_UVOZ     As Long = &HEA90&   ' PDF - izvodi su PDF fajlovi
Public Const IC_NALOZI   As Long = &HE9D5&   ' CheckList - nalozi su spisak stavki
Private Const IC_MARZA    As Long = &HE9D2&   ' AreaChart
Public Const IC_IZVEST   As Long = &HE9F9&   ' ReportDocument
Private Const IC_SLEDLJ   As Long = &HE71B&   ' Link
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
Private Const IC_REFRESH  As Long = &HE72C&   ' Refresh - rezervisano
' Stanja
Private Const IC_SYNC     As Long = &HE753&   ' Cloud
Private Const IC_HELP     As Long = &HE897&   ' Help
Private Const IC_MATICNI  As Long = &HE713&   ' Setting
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
Private mCache As Object
Private mView As Variant
Private mViewN As Long
Private mSumKg As Double
Private mSumVal As Double
Private mPage As Long
Private mPageSize As Long
Private mSortCol As Long
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
Private mCntOtkaz As Long
Private mCntBezZb As Long
Private mCntNefakt As Long
Private mToday As Double
Private mChromeRemoved As Boolean
Private mCombosFilled As Boolean
Private mPartMap As Object
Private mSelRow As Long              ' izabran red mreze (1..mViewN), 0 = nijedan
Private mHoverRow As Long            ' red pod pokazivacem (0-bazirano), -1 = nijedan
Private mHoverHd As Long             ' zaglavlje kolone pod pokazivacem, -1 = nijedno
Private mBuildMs As Long             ' koliko je trajala gradnja ekrana (ms)
Private mPopFor As String            ' combo za koji je otvoren nas dropdown ("" = zatvoren)
Private mColX(0 To MAX_COLS - 1) As Single   ' x i sirina kolona mreze; postavlja
Private mColW(0 To MAX_COLS - 1) As Single   ' LayoutGrid, RenderGrid samo puni
Private mCols As Variant             ' opis kolona AKTIVNOG rezima (GridCols)
Private mColN As Long                ' koliko ih je stvarno vidljivo
' Poslednja zbirna sa kojom se radilo. Preuzeto iz frmDokumenta: tamo je
' RefreshBrojZbirneSuggestion upisivala isti broj u sekciju Zbirne I u sekciju
' Otpremnice, a btnUnosZbr ga je posle snimanja gurao u Prijemnicu. Ovde je to
' jedan sticky broj koji prati prelazak izmedju rezima.
Private mAktivnaZbirna As String
Private mFillStep As String          ' gde je FillGrid stigao (za prijavu greske)
Private mGridMax As Boolean          ' mreza razvucena do ispod naslova dokumenta
Private mSmerRev As Long             ' izabrani smer reversa (1..4)
Private mPartnerFor As String        ' za koji rezim je partner lista vec napunjena
Private mBlokIDs As Variant          ' OtkupID po indeksu combo-a otvorenih blokova
Private mBlokRest As Variant         ' neisplaceni ostatak po istom indeksu
Private mPopSrc() As Long            ' indeksi stavki combo-a koje panel prikazuje
Private mPopN As Long                ' koliko ih je posle suzavanja kucanjem
Private mPopTop As Long              ' prva prikazana (skrol po stranama)
Private mPopMute As Boolean          ' programski upis u combo - ne otvaraj panel
Private mFltOpen As Boolean          ' panel "Filteri" otvoren
Private mFltVrsta As String          ' dodatni uslov: vrsta iz Osnovnih podataka
Private mFltPart As String           ' dodatni uslov: partner iz Osnovnih podataka
Private mMonthStart As Double        ' prvi dan tekuceg meseca (filter "mesec")
Private mDirty As Boolean            ' ima neupisanih izmena u formi
Private mLoading As Boolean          ' program puni polja - to NIJE izmena korisnika
'=====================================================================
' BUILD
'=====================================================================
Public Sub BuildOtkupScreen(frm As Object)
    Dim t0 As Double: t0 = Timer
    mBuilding = True
    Set mFrm = frm
    Set Btns = New Collection
    Set mHovered = Nothing
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 1
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
    BuildCtx frm
    BuildForm frm
    BuildGrid frm
    BuildRight frm
    BuildStatus frm

    BuildPopup frm
    BuildFilterPanel frm
    SetKlasa 1                      ' II klasa starta onemogucena
    SelectModeCore frm, "F1", False
    ShowZones frm
    mBuilding = False
    mBuildMs = CLng((Timer - t0) * 1000)
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
Private Sub BuildNav(frm As Object)
    Dim z As Object, Y As Single, i As Long, j As Long
    Dim grp As Variant, itm As Variant, ico As Variant
    Set z = NewZone(frm, "zNav", 0, HEADER_H, SIDEBAR_W, 400, C_SAND)
    WireZone z
    NewLbl z, "navEdge", "", SIDEBAR_W - 1, 0, 1, 400, 8, False, 0, C_BORDER

    grp = Array(Poruka("OTKUI_NAVG_OPERACIJE"), Poruka("OTKUI_NAVG_FINANSIJE"), Poruka("OTKUI_NAVG_ANALITIKA"))
    Y = 12
    For i = 0 To 2
        NewLbl z, "navGrp" & i, UCase$(CStr(grp(i))), 13, Y, 130, TxtH(TS_NAVGRP), _
               TS_NAVGRP, True, C_MUTED, -1
        Y = Y + 20
        Select Case i
            Case 0
                ' Otkup i blokovi su spojeni: svi dokumenti su na ovom ekranu,
                ' pa dve stavke koje vode na isto mesto nemaju svrhu.
                itm = Array(Poruka("OTKUI_NAV_UNOS"), _
                            Poruka("OTKUI_NAV_PALETE"), Poruka("OTKUI_NAV_AGRO"))
                ico = Array(IC_OTKUP, IC_PALETE, IC_AGRO)
            Case 1
                ' Banka nije podnaslov nego se raspada na svoje dve radnje -
                ' operater bira sta radi, ne u koji modul ulazi.
                itm = Array(Poruka("OTKUI_NAV_FAKT"), Poruka("OTKUI_NAV_BANKA_UVOZ"), _
                            Poruka("OTKUI_NAV_BANKA_NALOZI"), Poruka("OTKUI_NAV_MARZA"))
                ico = Array(IC_FAKT, IC_UVOZ, IC_NALOZI, IC_MARZA)
            Case 2
                itm = Array(Poruka("OTKUI_NAV_IZVESTAJI"), Poruka("OTKUI_NAV_SLEDLJIVOST"))
                ico = Array(IC_IZVEST, IC_SLEDLJ)
        End Select
        For j = 0 To UBound(itm)
            Dim sel As Boolean: sel = (i = 0 And j = 0)
            NewLbl z, "navBar" & i & j, "", 0, Y, 3, NAV_H, 8, False, 0, IIf(sel, C_GOLD, C_SAND)
            If CLng(ico(j)) <> 0 Then
                NewLbl z, "navIco" & i & j, ChrW(CLng(ico(j))), 12, CenterIco(Y, NAV_H, TS_NAVICO), _
                       20, TxtH(TS_NAVICO), TS_NAVICO, False, IIf(sel, C_GOLD, C_ICON_OFF), -1, _
                       fmTextAlignLeft, F_ICON
            End If
            NewNavItem z, "nav" & i & j, CStr(itm(j)), 0, Y, SIDEBAR_W - 1, NAV_H, sel
            ' stavka menija je neprozirna i preko pune sirine - bez ZOrder-a
            ' bi prekrila i zlatnu traku selekcije i MDL2 ikonicu
            z.Controls("navBar" & i & j).ZOrder 0
            On Error Resume Next
            z.Controls("navIco" & i & j).ZOrder 0     ' moze da ne postoji (kod 0)
            On Error GoTo 0
            z.Controls("nav" & i & j & "X").ZOrder 0
            Y = Y + NAV_H
        Next j
        Y = Y + 12
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
End Sub

'------------------------------------------------------------ KPI ----
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
        NewSectionHdr z, "kpiC" & i, CStr(cap(i)), 0, 9, 120
        NewLbl z, "kpiV" & i, ChrW(8212), 0, 22, 120, TxtH(TS_KPI), IIf(i = 4, TS_H1, TS_KPI), True, C_FOREST, -1, _
               fmTextAlignLeft, IIf(i = 4, F_UI, F_NUM)
        ' treci red: kontekst uz vrednost; prazan kad nema sta da kaze
        NewLbl z, "kpiS" & i, "", 0, 44, 120, TxtH(TS_MICRO), TS_MICRO, False, RGB(146, 158, 140), -1
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

    NewLbl z, "titName", "-", PAD + TIT_ICO_W + 10, 3, 230, TxtH(TS_DISPLAY), TS_DISPLAY, True, C_FOREST, -1, fmTextAlignLeft, mDisplayFont
    NewLbl z, "titSub", "-", PAD + TIT_ICO_W + 10, 29, 380, TxtH(TS_META), TS_META, False, C_MUTED, -1
    ' znacka rezima (F1..F8) - stoji uz naslov, ne u liniji sa podnaslovom
    NewLbl z, "titBadgeB", "", 0, 6, 26, 15, 8, False, 0, C_SOFT_BG
    NewLbl z, "titBadgeC", "", 0, 7, 26, TxtH(TS_MICRO), TS_MICRO, True, C_GREEN, -1, fmTextAlignCenter, F_NUM
    z.Controls("titBadgeC").ZOrder 0
    ' lenjir za merenje naslova - znacka mora da stane tacno iza teksta
    NewLbl z, "titRul", "", 0, -400, 10, 14, TS_DISPLAY, True, C_WHITE, -1, fmTextAlignLeft, mDisplayFont
    z.Controls("titRul").AutoSize = True
    NewLbl z, "titDatum", "-", 0, CenterY(0, TITLE_H, TS_META), 190, TxtH(TS_META), _
           TS_META, False, C_MUTED, -1, fmTextAlignRight
End Sub

'------------------------------------------------------------ CTX ----
Private Sub BuildCtx(frm As Object)
    Dim z As Object, i As Long, nm As Variant, w As Variant, X As Single
    Set z = NewZone(frm, "zCtx", SIDEBAR_W, HEADER_H + KPI_H + TITLE_H, 620, CTX_H, C_SAND_LT)
    WireZone z
    NewLbl z, "ctxLnB", "", 0, CTX_H - 1, 620, 1, 8, False, 0, C_BORDER

    ' Naslov sekcije dobija isti tretman kao grupe polja: eyebrow levo, tanka
    ' linija preko sredine, hint desno - pa tek onda red polja. Ranije je stajao
    ' u istoj liniji sa poljima, pa je izgledao kao natpis prvog polja.
    NewSectionHdr z, "ctxHdr", Poruka("OTKUI_LBL_OSNOVNI"), PAD, 8, 96
    NewLbl z, "ctxLn", "", 0, 13, 100, 1, 8, False, 0, C_BORDER_LT
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
    NewLbl z, "ctxHint", "Ctrl+K", 0, 8, 52, TxtH(TS_MICRO), TS_MICRO, False, C_MUTED, -1, fmTextAlignRight
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
    NewSegBtn fr, "segRev1", Poruka("OTKUI_SEG_REV_IZD_KOOP"), 1, 17, 91, FIELD_H - 2, True
    NewSegBtn fr, "segRev2", Poruka("OTKUI_SEG_REV_PRI_KOOP"), 93, 17, 91, FIELD_H - 2, False
    NewSegBtn fr, "segRev3", Poruka("OTKUI_SEG_REV_IZD_OM"), 185, 17, 91, FIELD_H - 2, False
    NewSegBtn fr, "segRev4", Poruka("OTKUI_SEG_REV_PRI_OM"), 277, 17, 91, FIELD_H - 2, False


    ' VREDNOST - forest ploca (isti shell, druga ispuna). Ploca ne zauzima celo
    ' polje: levo od nje stoji racunica po klasi, pa se vidi ODAKLE iznos.
    ' Ploca iznosa NEMA grupu: ne redja se sa poljima nego deli red sa dugmadima
    ' (iznos i potvrda su jedna radnja), pa ne trosi ceo red za sebe.
    Set fr = NewFrame(z, "fgVrednost", 0, 0, 180, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:1:"
    NewLbl fr, "fgVrednostL2", "", 0, 0, 170, 12, TS_LABEL, True, C_MUTED, -1
    NewShell fr, "fgVrednost", 0, 16, 180, FIELD_H, C_FOREST, C_FOREST
    NewLbl fr, "fgVrednostA", "", 1, 17, 3, FIELD_H - 2, 8, False, 0, C_GOLD
    NewLbl fr, "fgVrednostL", Poruka("OTKUI_LBL_VREDNOST"), INPUT_PAD, CenterY(16, FIELD_H, TS_MICRO), _
           66, TxtH(TS_MICRO), TS_MICRO, False, RGB(166, 178, 160), -1
    NewLbl fr, "fgVrednostV", "0", 0, CenterY(16, FIELD_H, TS_VAL), 100, TxtH(TS_VAL), _
           TS_VAL, True, C_CREAM, -1, fmTextAlignRight, F_NUM
    NewLbl fr, "fgVrednostU", Poruka("OTKUI_UNIT_RSD"), 0, CenterY(16, FIELD_H, TS_MICRO), _
           28, TxtH(TS_MICRO), TS_MICRO, True, C_GOLD, -1, fmTextAlignRight
    ' dva reda racunice - po jedan za svaku klasu; stoje DESNO od ploce
    NewLbl fr, "fgVrednostK1", "", 0, 18, 120, TxtH(TS_MICRO), TS_MICRO, False, C_DISABLED_FG, -1
    NewLbl fr, "fgVrednostK2", "", 0, 31, 120, TxtH(TS_MICRO), TS_MICRO, False, C_DISABLED_FG, -1

    ' OTVORENI BLOK + NEISPLACENI OSTATAK (samo isplate). U frmDokumenta je to
    ' cmbOtkupBlok, poznat kao "preostali kes" - poslovno to nije kes nego
    ' neisplaceni ostatak po otkupnom bloku, pa polje tako i pise.
    NewFieldG z, "fgBlok", Poruka("OTKUI_FLD_BLOK"), "cmb", "", 1, False, False, "NOV"
    Set fr = NewFrame(z, "fgOstatak", 0, 0, 180, FIELD_GRP_H, C_WHITE)
    fr.tag = "fld:1:"
    NewLbl fr, "fgOstatakL2", "", 0, 0, 170, 12, TS_LABEL, True, C_MUTED, -1
    NewShell fr, "fgOstatak", 0, 16, 180, FIELD_H, C_FOREST, C_FOREST
    NewLbl fr, "fgOstatakL", Poruka("OTKUI_LBL_OSTATAK"), INPUT_PAD, CenterY(16, FIELD_H, TS_MICRO), _
           86, TxtH(TS_MICRO), TS_MICRO, False, RGB(166, 178, 160), -1
    NewLbl fr, "fgOstatakV", "0", 0, CenterY(16, FIELD_H, TS_VAL), 100, TxtH(TS_VAL), _
           TS_VAL, True, C_CREAM, -1, fmTextAlignRight, F_NUM
    NewLbl fr, "fgOstatakU", Poruka("OTKUI_UNIT_RSD"), 0, CenterY(16, FIELD_H, TS_MICRO), _
           28, TxtH(TS_MICRO), TS_MICRO, True, C_GOLD, -1, fmTextAlignRight

    ' akcije: sekundarno pa primarno, desno poravnato
    BtnV z, "btnOtkazi", Poruka("OTKUI_BTN_OTKAZI"), 0, 0, 72, 28, "plain"
    BtnV z, "btnSacuvajPrint", Poruka("OTKUI_BTN_SACUVAJ_PRINT"), 0, 0, 132, 28, "secondary", IC_PRINT
    ' Natpis je kontekstualan ("Sacuvaj otkupni list"...), postavlja ga
    ' SelectModeCore - zato je dugme sire nego generickim "Sacuvaj".
    BtnV z, "btnSacuvaj", Poruka("OTKUI_BTN_SACUVAJ_OTKUP"), 0, 0, 196, 28, "primary", IC_SAVE

    ' toast
    Set fr = NewFrame(z, "tstOk", 0, 0, 320, 26, RGB(236, 245, 240))
    fr.Visible = False
    NewLbl fr, "tstBar", "", 0, 0, 3, 26, 8, False, 0, C_GREEN_LT
    NewLbl fr, "tstMsg", "", 11, 6, 300, 15, TS_META, True, C_GREEN, -1
End Sub

'----------------------------------------------------------- GRID ----
Private Sub BuildGrid(frm As Object)
    Dim z As Object, i As Long, chip As Variant, X As Single, hd As Object, ft As Object
    Set z = NewZone(frm, "zGrid", SIDEBAR_W, 0, 620, 260, C_SAND_LT)
    WireZone z
    NewLbl z, "grdLnT", "", 0, 0, 620, 1, 8, False, 0, C_BORDER

    NewLbl z, "grdTitle", "-", PAD, 10, 180, 16, TS_H1, True, C_FOREST, -1
    NewLbl z, "grdSrc", "-", PAD + 176, 11, 120, 14, TS_MICRO, False, C_MUTED, C_SAND, fmTextAlignCenter, F_NUM
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
    NewLbl ft, "ftLnT", "", 0, 0, 620, 1, 8, False, 0, C_BORDER
    NewLbl ft, "ftRange", "-", PAD, 6, 170, 14, TS_META, False, C_MUTED, -1
    ' zbirovi su brojevi - Consolas, podebljano (Consolas nema Semibold rez)
    NewLbl ft, "ftKg", "-", 0, 6, 170, 14, TS_CELL_LG, True, C_FOREST, -1, fmTextAlignRight, F_NUM
    NewLbl ft, "ftVal", "-", 0, 6, 190, 14, TS_CELL_LG, True, C_FOREST, -1, fmTextAlignRight, F_NUM
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

    ' OSAM kartica - sada su prikazani SVI rezimi, ukljucujuci aktivni, a
    ' aktivni je obelezen (tamna ispuna + zlatna traka). Ranije se aktivni
    ' izostavljao, pa se spisak premetao pri svakoj promeni i nije se videlo
    ' gde si. Redosled je fiksan F1..F7, kartica nosi rezim u .Tag.
    ' Storno nije unosni rezim - ide POSLEDNJI i odvojen je linijom. Reversi su
    ' zato pomereni ispred njega (RefreshModeCards drzi isti redosled).
    For i = 0 To 7
        Dim cy As Single: cy = 30 + i * 40 + IIf(i = 7, MODE_SEP_H, 0)
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

    NewLbl z, "rgtSepM", "", 13, 30 + 7 * 40 + 5, RIGHT_W - 26, 1, 8, False, 0, C_BORDER
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
        ' jedan Label po redu: i pozadina i tekst, pa hover i klik rade na istoj
        ' kontroli (kutija je jedva visa od teksta, poravnanje ostaje uredno)
        NewLbl z, "pop" & i, "", 1, 1 + i * POP_ITEM_H, 178, POP_ITEM_H, _
               TS_BODY, False, C_FOREST, C_WHITE
        WireBtn z.Controls("pop" & i), "pop" & i, "btn"
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
            z.Controls("pop" & i).caption = "   " & CStr(cmb.List(mPopSrc(mPopTop + i), 0))
            z.Controls("pop" & i).BackColor = C_WHITE
            z.Controls("pop" & i).top = 1 + i * POP_ITEM_H
            z.Controls("pop" & i).width = z.width - 2
            z.Controls("pop" & i).Visible = True
        Else
            z.Controls("pop" & i).Visible = False
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
    If mFrm Is Nothing Then Exit Sub
    mFrm.Controls("zPop").Visible = False
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

Private Function FindCombo(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCombo = mFrm.Controls("zCtx").Controls(nm)
    If FindCombo Is Nothing Then _
        Set FindCombo = mFrm.Controls("zForm").Controls(nm).Controls(nm & "T")
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
        .Controls("profLn").top = .Height - 62: .Controls("profLn").width = navW - 1
        .Controls("profBox").top = .Height - 54
        .Controls("profIco").top = CenterIco(.Height - 54, 26, TS_NAVICO)
        .Controls("profName").top = .Height - 52
        .Controls("profName").width = navW - 50
        .Controls("profSub").top = .Height - 39
        .Controls("profSub").width = navW - 50
        .Controls("profHit").top = .Height - 56: .Controls("profHit").width = navW - 1
    End With
    NavCollapse frm, (navW = SIDEBAR_MIN)

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

    With frm.Controls("zCtx")
        .Left = mainX: .top = HEADER_H + KPI_H + TITLE_H: .width = ctrW: .Height = CTX_H
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

    ' Razvucena mreza: kontekst i forma se sklanjaju, mreza pocinje odmah ispod
    ' naslova dokumenta. Zone se ne unistavaju - drugi klik ih vraca netaknute.
    frm.Controls("zCtx").Visible = Not mGridMax
    frm.Controls("zForm").Visible = Not mGridMax
    If mGridMax Then
        gridY = HEADER_H + KPI_H + TITLE_H
    Else
        gridY = HEADER_H + KPI_H + TITLE_H + CTX_H + formH
    End If
    ' Donja granica zone mreze MORA da racuna i traku iznad tela (GRID_TOP:
    ' naslov, pretraga, cipovi). Ranije je bila 231pt bez nje, pa je zona bila
    ' visa od raspolozivog prostora i podnozje (Prikazano / Ukupno / strane)
    ' je ispadalo ispod donje ivice forme. Sada je minimum tacno ono sto zona
    ' zaista zahteva za tri reda.
    gridH = hTot - gridY - STATUS_H
    If gridH < GRID_TOP + GRID_HEAD_H + GRID_FOOT_H + 3 * GRID_ROW_H + 6 Then
        gridH = GRID_TOP + GRID_HEAD_H + GRID_FOOT_H + 3 * GRID_ROW_H + 6
    End If
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
    Dim valFr As Object, btnY As Single, btnX As Single, tw As Single
    colW = (zw - 2 * PAD - (cols - 1) * GAP) / cols
    Y = 10
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
                        col = 0: Y = Y + FIELD_GRP_H + GAP
                    End If
                    c.Left = PAD + col * (colW + GAP)
                    c.width = colW * span + (span - 1) * GAP
                    c.top = Y
                    c.Height = FIELD_GRP_H
                    LayoutFieldInner c
                    col = col + span
                    If col >= cols Then col = 0: Y = Y + FIELD_GRP_H + GAP
                End If
            End If
        Next c
        If imaPolja Then
            If col > 0 Then Y = Y + FIELD_GRP_H + GAP
            Y = Y + 6                       ' vazduh izmedju grupa
        Else
            z.Controls("grpH" & g).Visible = False
            z.Controls("grpLn" & g).Visible = False
        End If
    Next g

    ' AKCIONI RED. Redosled sleva: poruka | racunica + IZNOS | dugmad.
    ' Iznos i potvrda su jedna radnja pa dele red - ploca ne trosi ceo red.
    ' Redosled dugmadi: Sacuvaj (najcesce) najblize poljima, Otkazi najdalje.
    btnY = Y + 6
    btnX = zw - PAD - (196 + GAP + 132 + GAP + 72)
    MoveBtn z, "btnSacuvaj", btnX, btnY
    MoveBtn z, "btnSacuvajPrint", btnX + 196 + GAP, btnY
    MoveBtn z, "btnOtkazi", btnX + 196 + GAP + 132 + GAP, btnY

    ' Ploca ide uz LEVU ivicu reda (kao u viziji), poruka popunjava sredinu,
    ' dugmad ostaju desno.
    Set valFr = VisibleValFrame(z)
    tw = PAD
    If Not valFr Is Nothing Then
        valFr.width = VAL_PLATE_W + IIf(HasCtl(valFr, valFr.name & "K1"), GAP + 4 + VAL_CALC_W, 0)
        valFr.Left = PAD
        valFr.top = btnY - 17            ' ploca (30pt) centrirana na dugme (28pt)
        valFr.Height = FIELD_GRP_H
        LayoutFieldInner valFr
        tw = valFr.Left + valFr.width + GAP
    End If
    z.Controls("tstOk").top = btnY
    z.Controls("tstOk").Left = tw
    z.Controls("tstOk").width = btnX - tw - GAP
    If z.Controls("tstOk").width < 60 Then z.Controls("tstOk").width = 60
    z.Controls("tstOk").Controls("tstMsg").width = z.Controls("tstOk").width - 18

    LayoutFields = btnY + 28 + 14
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

Private Sub LayoutFieldInner(fr As Object)
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
        MoveShell fr, nm, 0, 16, plateW
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
    z.Controls("grdSrc").Left = PAD + 176
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
    bodyH = zh - (GRID_TOP + GRID_HEAD_H) - GRID_FOOT_H - 6
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
    ft.Controls("ftLnT").width = sc
    ft.Controls("ftKg").Left = sc / 2 - 190
    ft.Controls("ftVal").Left = sc / 2 - 10
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
    For i = 0 To 7
        z.Controls("mcKey" & i).Left = RIGHT_W - 13 - 30
        z.Controls("mcHint" & i).width = RIGHT_W - 42
    Next i
    Y = 30 + 8 * 40 + 12 + MODE_SEP_H
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

Private Sub NavCollapse(frm As Object, collapsed As Boolean)
    Dim c As Object, z As Object
    Set z = frm.Controls("zNav")
    For Each c In z.Controls
        If Left$(c.name, 6) = "navGrp" Then c.Visible = Not collapsed
        If Left$(c.name, 6) = "navIco" Then c.Left = IIf(collapsed, 11, 12)
        If Left$(c.name, 3) = "nav" And Len(c.name) = 6 Then
            If Right$(c.name, 1) = "X" Then c.Visible = Not collapsed
        End If
        If Left$(c.name, 3) = "nav" And Len(c.name) = 5 Then
            c.width = IIf(collapsed, SIDEBAR_MIN - 1, SIDEBAR_W - 1)
        End If
    Next c
    z.Controls("navSezona").Visible = Not collapsed
    z.Controls("navVerzija").Visible = Not collapsed
    ' u suzenom sidebaru ostaje samo kvadrat sa glifom - ime nema gde da stane
    z.Controls("profName").Visible = Not collapsed
    z.Controls("profSub").Visible = Not collapsed
End Sub

'=====================================================================
' RENDER MREZE - puni vec napravljene Label-ove, ne pravi kontrole
'=====================================================================
Public Sub RenderGrid()
    Dim z As Object, body As Object, ft As Object
    Dim i As Long, r As Long, k As Long, from_ As Long, to_ As Long, shown As Long
    Dim kanal As Boolean
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    kanal = Not ModeHasKgCol()                ' rezim bez kilograma nema ni zbir kg
    Set z = mFrm.Controls("zGrid")
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
                        Select Case ColKind(k)
                            Case "date":               .caption = FmtDatumKratko(mView(r, k + 1))
                            Case "kg", "num", "sum0":  .caption = FmtBroj(CDbl(mView(r, k + 1)), 0)
                            Case "rsd", "mult":        .caption = FmtBroj(CDbl(mView(r, k + 1)), 2)
                            Case "pill":               PaintPill body.Controls("c" & i & "_" & k), CLng(mView(r, k + 1))
                            Case "paypill":            PaintPayPill body.Controls("c" & i & "_" & k), CLng(mView(r, k + 1))
                            Case "rest"
                                ' nula znaci "nema duga" - prazna celija je citljivija od 0,00
                                If CDbl(mView(r, k + 1)) = 0 Then
                                    .caption = ""
                                Else
                                    .caption = FmtBroj(CDbl(mView(r, k + 1)), 2)
                                End If
                            Case Else:                 .caption = CStr(mView(r, k + 1))
                        End Select
                    End If
                End With
            Next k
            PaintRow body, i, (from_ + i = mSelRow), (i = mHoverRow)
        Else
            body.Controls("rb" & i).Visible = False
            body.Controls("rl" & i).Visible = False
            For k = 0 To MAX_COLS - 1
                body.Controls("c" & i & "_" & k).Visible = False
            Next k
        End If
    Next i

    If mViewN > 0 Then
        ft.Controls("ftRange").caption = Poruka("OTKUI_FT_PRIKAZANO") & " " & from_ & ChrW(8211) & to_ & _
                                         " " & Poruka("OTKUI_FT_OD") & " " & mViewN
    Else
        ft.Controls("ftRange").caption = Poruka("OTKUI_FT_PRIKAZANO") & " 0"
    End If
    ' gotovinski promet i reversi nemaju kilograme - zbir kg bi bio prazna nula
    ft.Controls("ftKg").Visible = Not kanal
    ft.Controls("ftKg").caption = Poruka("OTKUI_FT_UKUPNO") & " " & FmtBroj(mSumKg, 0) & " " & Poruka("OTKUI_UNIT_KG")
    ' rezim bez kolone vrednosti (Zbirna) nema ni zbir
    ft.Controls("ftVal").Visible = ModeHasValCol()
    ' reversi se broje u komadima, sve ostalo u dinarima (i bez decimala)
    If ActiveMode = "F7" Then
        ft.Controls("ftVal").caption = Poruka("OTKUI_FT_UKUPNO") & " " & _
                                       FmtBroj(mSumVal, 0) & " " & ModeValUnit(ActiveMode)
    Else
        ft.Controls("ftVal").caption = Poruka("OTKUI_FT_VREDNOST") & " " & _
                                       FmtBroj(mSumVal, 2) & " " & ModeValUnit(ActiveMode)
    End If
    RenderPager ft
    RenderChipCounts z
End Sub

' status-pilula u boji - ono sto MSForms ListBox nije mogao.
' Sirina se racuna PO TEKSTU: preko cele kolone tinta (#EDF0F5 i sl.) je na
' belom redu prakticno nevidljiva, pa se pilula cita kao oblik tek kad
' pripije uz tekst. Zato LayoutGrid namerno NE postavlja sirinu kolone 5.
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

Private Sub PaintRow(body As Object, ByVal i As Long, ByVal isSel As Boolean, ByVal isHover As Boolean)
    Dim bg As Long, c As Long
    On Error Resume Next
    If isSel Then
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

Private Sub RenderChipCounts(z As Object)
    On Error Resume Next
    z.Controls("chipOtkazaneC").caption = Poruka("OTKUI_CHIP_OTKAZANE") & "  " & mCntOtkaz
    z.Controls("chipBezZbirneC").caption = Poruka("OTKUI_CHIP_BEZZBIRNE") & "  " & mCntBezZb
    z.Controls("chipNefaktC").caption = Poruka("OTKUI_CHIP_NEFAKT") & "  " & mCntNefakt
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
            FldShow z, "fgTipAmb", False
            FldShow z, "fgAmbPr", False
            FldShow z, "fgParcela", False
            FldShow z, "fgSmerRev", False
            FldShow z, "fgVrednost", False
            ' otvoreni blok i neisplaceni ostatak imaju smisla samo kod isplata
            FldShow z, "fgBlok", (mode = "F5")
            FldShow z, "fgOstatak", (mode = "F5")
        Case "F7"                           ' ambalaza: samo tip i kolicina
            FldShow z, "fgBrZbir", False
            FldShow z, "fgNovac", False
            FldShow z, "fgKgI", False
            FldShow z, "fgKgII", False
            FldShow z, "fgCena", False
            FldShow z, "fgKolAmb", True
            FldShow z, "fgTipAmb", True
            FldShow z, "fgAmbPr", False
            FldShow z, "fgParcela", False
            FldShow z, "fgSmerRev", True
            FldShow z, "fgVrednost", False
            FldShow z, "fgBlok", False
            FldShow z, "fgOstatak", False
        Case Else                           ' robni dokumenti
            ' u rezimu Zbirne broj dokumenta JESTE broj zbirne - zasebno polje
            ' bi bilo isti podatak dvaput
            FldShow z, "fgBrZbir", ModeVezujeZbirnu(mode)
            FldShow z, "fgNovac", False
            FldShow z, "fgKgI", True
            FldShow z, "fgKgII", True
            FldShow z, "fgCena", True
            FldShow z, "fgKolAmb", True
            FldShow z, "fgTipAmb", True
            ' prazna ambalaza se izdaje uz otkup i vraca uz prijemnicu;
            ' otpremnica i zbirna je ne dodiruju
            FldShow z, "fgAmbPr", (mode = "F1" Or mode = "F4")
            FldShow z, "fgParcela", (mode = "F1" And IsPracenjeParcela())
            FldShow z, "fgSmerRev", False
            FldShow z, "fgVrednost", True
            FldShow z, "fgBlok", False
            FldShow z, "fgOstatak", False
    End Select
End Sub

' Cetiri smera reversa su medjusobno iskljuciva - isti obrazac kao SetKlasa.
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

Private Sub SelectModeCore(frm As Object, ByVal key As String, ByVal doReload As Boolean)
    Dim k As String
    On Error Resume Next
    mPopMute = True                    ' punjenje lista pise u combo -> ne otvaraj panel
    mLoading = True                    ' i to NIJE izmena korisnika
    ClosePopup
    ActiveMode = key
    k = modeKey(key)
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
    FillParcele frm
    ' prazna ambalaza: uz otkup se izdaje, uz prijemnicu se vraca
    If key = "F1" Or key = "F4" Then
        frm.Controls("zForm").Controls("fgAmbPr").Controls("fgAmbPrL").caption = _
            UCase$(Poruka("OTKUI_FLD_AMB_PR_" & k))
    End If

    If key = "F8" Then mFilter = "otkazane" Else mFilter = "danas"
    mSelRow = 0
    ApplyChipVisual frm.Controls("zGrid")
    ' "Bez zbirne" i "Nefakturisane" nemaju smisla nad zbirnom ni nad
    ' gotovinskim prometom (tblNovac nema BrojZbirne)
    ShowChip frm, "chipBezZbirne", ModeHasZbirna(key)
    ShowChip frm, "chipNefakt", ModeHasFaktura(key)
    LayoutChips frm
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
    ' nov dokument -> nista nije neupisano
    MarkClean
    mPopMute = False
    mLoading = False
    If doReload Then
        mView = Empty: mViewN = 0
        ReloadGrid
    End If
End Sub

Private Sub RefreshModeCards(frm As Object)
    Dim z As Object, i As Long, m As Variant
    Dim mk As String, md As String, act As Boolean, warn As Boolean
    Set z = frm.Controls("zRight")
    ' redosled kartica: unosni rezimi, pa reversi, pa STORNO kao poslednji
    ' brojevi tastera prate REDOSLED u koloni: reversi F7, storno F8 (poslednji)
    m = Array("F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8")
    For i = 0 To 7
        md = CStr(m(i))
        mk = modeKey(md)
        act = (md = ActiveMode)
        warn = (md = "F8")

        ' na kartici stoji KRATKO ime (108pt); pun naziv nosi veliki naslov
        z.Controls("mcT" & i).caption = Poruka("OTKUI_MODE_" & mk)
        z.Controls("mcT" & i).tag = md
        z.Controls("mcKey" & i).caption = md
        z.Controls("mcHint" & i).caption = Poruka("OTKUI_HINT_" & mk)

        If act Then
            ' aktivni: puna tamna ispuna + zlatna traka - vidi se iz daljine
            z.Controls("mc" & i & "B").BackColor = IIf(warn, C_RUST, C_FOREST)
            z.Controls("mc" & i & "F").BackColor = IIf(warn, C_RUST, C_FOREST)
            z.Controls("mc" & i & "A").BackColor = C_GOLD
            z.Controls("mcT" & i).ForeColor = C_CREAM
            z.Controls("mcHint" & i).ForeColor = RGB(168, 180, 160)
            z.Controls("mcKey" & i).ForeColor = C_GOLD
        Else
            z.Controls("mc" & i & "B").BackColor = IIf(warn, C_RUST, C_BORDER_LT)
            z.Controls("mc" & i & "F").BackColor = IIf(warn, C_PILL_ERR_BG, C_WHITE)
            z.Controls("mc" & i & "A").BackColor = IIf(warn, C_RUST, C_GOLD)
            z.Controls("mcT" & i).ForeColor = IIf(warn, C_RUST, C_FOREST)
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
            If c.name = "chipOtkazane" Then
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
    Dim c As Object, z As Object, on_ As Boolean, idx As String
    On Error Resume Next                ' navIco moze da ne postoji (kod 0)
    Set z = frm.Controls("zNav")
    For Each c In z.Controls
        If Left$(c.name, 3) = "nav" And Len(c.name) = 5 Then
            on_ = (c.name = nm)
            idx = Mid$(c.name, 4, 2)
            c.BackColor = IIf(on_, C_FOREST, C_SAND)
            c.Font.bold = on_                       ' marker stanja za hover (clsFlatBtn)
            z.Controls(c.name & "X").ForeColor = IIf(on_, C_CREAM, RGB(52, 68, 44))
            z.Controls(c.name & "X").Font.bold = on_
            z.Controls("navBar" & idx).BackColor = IIf(on_, C_GOLD, C_SAND)
            z.Controls("navIco" & idx).ForeColor = IIf(on_, C_GOLD, C_ICON_OFF)
        End If
    Next c
    If nm <> "nav00" Then ShowToast Poruka("OTKUI_TODO_NEVEZANO"), False
End Sub

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
        Case "DblClick": RowFromTag tag: LoadRowIntoForm
        Case "Change":   UiChange tag
        Case "Hover":    UiHover tag
        Case "Focus":    CloseOverlays tag: ShellFor tag, "focus"
        Case "Blur":     ShellFor tag, "normal"
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
        RowFromTag tag
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
        Case "btnClose": mFrm.Hide
        Case "btnSync": DoSync
        Case "btnSnimi": DoSaveWorkbook
        Case "btnExcel": DoShowExcel
        Case "btnOperater": DoSwitchOperater
        Case "btnFilteri": ToggleFilterPanel
        Case "grdEmptyA"
            mFilter = "sve"
            mSelRow = 0
            ApplyChipVisual mFrm.Controls("zGrid")
            RenderFilterPanel
            RefreshFilterBadge
            ReloadGrid
        Case "btnMatic": ShowToast Poruka("OTKUI_TODO_NEVEZANO"), False
    End Select
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
        RefreshOperater mFrm
        ShowToast Poruka("OTKUI_MSG_OPERATER") & " " & OperaterText(), False
    Else
        ShowToast Poruka("OTKUI_MSG_OPERATER_ISTI") & " " & staro, False
    End If
    Exit Sub
EH:
    ShowToast Poruka("OTKUI_MSG_OPERATER_PAO"), True
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

Private Sub UiChange(ByVal tag As String)
    ' Kucanje u combo otvara nas panel i suzava listu. Programski upisi
    ' (punjenje lista, ClearForm, izbor iz panela) su pod mPopMute - bez toga
    ' bi se panel otvarao sam od sebe pri svakoj promeni rezima.
    If Not mPopMute And Not mBuilding Then PopFromTyping tag
    ' pretraga nije izmena dokumenta
    If tag <> "txtSearch" Then MarkDirty
    Select Case tag
        Case "txtSearch"
            mSearch = Trim$(CStr(mFrm.Controls("zGrid").Controls("txtSearch").text))
            mFrm.Controls("zGrid").Controls("lblSearchPh").Visible = (Len(mSearch) = 0)
            mSelRow = 0
            ReloadGrid
        Case "fgKgIT", "fgKgIIT", "fgCena1T", "fgCena2T"
            RecalcVrednost
        Case "fgBrZbirT"
            SetAktivnaZbirna mFrm.Controls("zForm").Controls("fgBrZbir").Controls("fgBrZbirT").text
        Case "fgBrOtprT"
            If ActiveMode = "F3" Then _
                SetAktivnaZbirna mFrm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").text
        Case "cbKupac"
            FillOpenBlokovi mFrm
            FillParcele mFrm
        Case "fgBlokT"
            ApplyBlokIzbor
        Case "cbVrsta"
            RefillSorta mFrm
        Case "cbOM"
            RefreshKpi mFrm
            RefreshStatusBar mFrm
    End Select
End Sub

Private Sub UiHover(ByVal tag As String)
    Dim i As Long
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
    If nm = "fgKgII" And mKlasa <> 2 Then Exit Sub
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

    If k = 2 Then mFrm.Controls("zForm").Controls("fgKgII").Controls("fgKgIIT").SetFocus
    RecalcVrednost
End Sub

' cip se sakriva zajedno sa svojom ivicom
' ime | kljuc natpisa | sirina
Private Function ChipRow() As Variant
    ChipRow = Array("chipDanas|OTKUI_CHIP_DANAS|54", "chipNedelja|OTKUI_CHIP_NEDELJA|78", _
                    "chipSve|OTKUI_CHIP_SVE|40", "chipNefakt|OTKUI_CHIP_NEFAKT|98", _
                    "chipBezZbirne|OTKUI_CHIP_BEZZBIRNE|86", "chipOtkazane|OTKUI_CHIP_OTKAZANE|82")
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
        If z.Controls(CStr(p(0))).Visible Then
            MoveBox z, CStr(p(0)), X, 36, CSng(p(2))
            X = X + CSng(p(2)) + 6
        End If
    Next i
End Sub

' polje: label + shell + kontrola uvucena za INPUT_PAD (+ jedinica / strelica)
Private Sub NewFieldG(parent As Object, nm As String, cap As String, kind As String, _
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
    If isFocus Then ShellState fr, nm, "focus"
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
    frmOtkupUI.show vbModeless
End Sub

Private Sub ShowZones(frm As Object)
    Dim nmv As Variant, i As Long
    nmv = Array("zHdr", "zNav", "zKpi", "zTitle", "zCtx", "zForm", "zGrid", "zRight", "zStatus")
    For i = 0 To UBound(nmv)
        On Error Resume Next
        frm.Controls(CStr(nmv(i))).Visible = True
    Next i
End Sub

Public Sub EnsureGridLoaded()
    FillCombos mFrm
    RefreshKpi mFrm
    RefreshStatusBar mFrm
    If mViewN > 0 Then Exit Sub
    ReloadGrid
End Sub

Public Sub RefreshFromData()
    On Error Resume Next
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 1
    Set mPartMap = Nothing
    Set mColSpec = CreateObject("Scripting.Dictionary")
    mColSpec.CompareMode = 1
    RefreshKpi mFrm
    FillZbirneCombo mFrm      ' nove zbirne u picker
    mPartnerFor = ""          ' partner lista se ponovo puni
    ReloadGrid
End Sub

Public Sub FillGrid(ByVal tblName As String, ByVal filter As String, ByVal q As String)
    Dim src As Variant, r As Long, n As Long, keep As Boolean, nRows As Long
    Dim outA() As Variant, c As Long, mk As String
    Dim ix() As Long, kind() As String
    Dim iStorno As Long, iZbir As Long, iKg As Long, iDokTip As Long, iFakt As Long
    Dim iTip As Long, iNap As Long, iEntTip As Long, iBrojCol As Long
    Dim iKoopID As Long, iPartID As Long, iOtkID As Long
    Dim mKoop As Object, mStan As Object, mKup As Object
    Dim rev As Boolean, kanal As Boolean

    On Error GoTo EH
    mFillStep = "start"
    mViewN = 0
    mView = Empty
    mSumKg = 0
    mSumVal = 0
    mCntOtkaz = 0
    mCntBezZb = 0
    mCntNefakt = 0
    If Len(tblName) = 0 Then Exit Sub

    src = CachedTable(tblName)
    If Not IsArray(src) Then Exit Sub

    mFillStep = "GridCols"
    mk = modeKey(ActiveMode)
    If Not IsArray(mCols) Then SetGridCols mk     ' rezerva; postavlja SelectModeCore

    ' indeksi izvornih kolona - JEDNOM po pozivu, ne po redu
    ReDim ix(0 To mColN - 1)
    ReDim kind(0 To mColN - 1)
    iKg = -1
    For c = 0 To mColN - 1
        kind(c) = ColF(CStr(mCols(c)), 2)
        ix(c) = ColIdx(tblName, ColF(CStr(mCols(c)), 1))
        If kind(c) = "kg" Then iKg = c
    Next c

    mFillStep = "indeksi kolona"
    iStorno = ColIdx(tblName, COL_STORNIRANO)
    iZbir = ColIdx(tblName, ColBrojZbirne(mk))
    If Len(ColFakturaID(mk)) > 0 Then iFakt = ColIdx(tblName, ColFakturaID(mk))

    rev = (mk = "REVERSI")
    If rev Then
        iBrojCol = ColIdx(tblName, COL_AMB_DOK_ID)
        iDokTip = ColIdx(tblName, COL_AMB_DOK_TIP)
        iEntTip = ColIdx(tblName, COL_AMB_ENTITET_TIP)
    End If
    kanal = (mk = "AMB_ISPLATE" Or mk = "AMB_UPLATE")
    If kanal Then
        iTip = ColIdx(tblName, COL_NOV_TIP)
        iNap = ColIdx(tblName, COL_NOV_NAPOMENA)
        iEntTip = ColIdx(tblName, COL_NOV_ENTITET_TIP)
        iKoopID = ColIdx(tblName, COL_NOV_KOOP_ID)
        iPartID = ColIdx(tblName, COL_NOV_PARTNER_ID)
        iOtkID = ColIdx(tblName, COL_NOV_OTKUP_ID)
    End If

    mFillStep = "partner mape"
    If rev Or kanal Then
        Set mKoop = PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Set mStan = PartnerMap(TBL_STANICE, "StanicaID", "Naziv", "")
        Set mKup = PartnerMap(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
    End If

    ' Placanje: gotove bulk rutine iz modNovac, jedan prolaz po tabeli novca.
    ' Za otkup je vezivanje direktno (tblNovac.OtkupID); za prijemnicu ide preko
    ' fakture, pa se stanje cita NA NIVOU FAKTURE - vidi PayCode.
    Dim pay As Boolean, dPay As Object, dFakIzn As Object
    Dim iPayID As Long, iPayKol As Long, iPayCena As Long
    mFillStep = "placanje"
    pay = (mk = "OTKUP" Or mk = "PRIJEMNICA")
    If pay Then
        If mk = "OTKUP" Then
            Set dPay = modNovac.BuildIsplataDictByOtkup()
            iPayID = ColIdx(tblName, COL_OTK_ID)
            iPayKol = ColIdx(tblName, COL_OTK_KOLICINA)
            iPayCena = ColIdx(tblName, COL_OTK_CENA)
        Else
            Set dPay = modNovac.BuildUplataDictByFaktura()
            Set dFakIzn = FakturaIznosMap()
            iPayID = ColIdx(tblName, COL_PRJ_FAKTURA_ID)
        End If
        If dPay Is Nothing Then pay = False
    End If

    Dim pl As Variant, pmap As Object
    pl = PartnerLookup(ActiveMode)
    If Len(CStr(pl(0))) > 0 Then _
        Set pmap = PartnerMap(CStr(pl(0)), CStr(pl(1)), CStr(pl(2)), CStr(pl(3)))

    mFillStep = "petlja po redovima"
    nRows = UBound(src, 1)
    ReDim outA(1 To nRows, 1 To mColN)
    n = 0

    For r = 1 To nRows
        Dim vDatK As Double, vZbir As String, hay As String
        Dim isStorno As Boolean, bezZbirne As Boolean, bezFakture As Boolean
        Dim vKgRow As Double, cell As Variant

        ' reversi su podskup tblAmbalaza - ostalo iz te knjige ne ulazi
        If rev Then
            If Not RevRowVisible(CellS(src, r, iDokTip), CellS(src, r, iEntTip)) Then GoTo NextRow
        End If

        vDatK = 0
        vKgRow = 0
        hay = ""
        If iKg >= 0 Then vKgRow = CellD(src, r, ix(iKg))

        Dim pCode As Long, pRest As Double, duguje As Double, placeno As Double
        Dim payKey As String
        pCode = 0: pRest = 0
        If pay Then
            duguje = 0: placeno = 0
            payKey = CellS(src, r, iPayID)
            If mk = "OTKUP" Then
                duguje = CellD(src, r, iPayKol) * CellD(src, r, iPayCena)
                If dPay.Exists(payKey) Then placeno = CDbl(dPay(payKey))
                pCode = PayCode(duguje, placeno)
            ElseIf Len(payKey) = 0 Then
                pCode = PAY_NEFAKT              ' prijemnica jos nije na fakturi
            Else
                If Not dFakIzn Is Nothing Then
                    If dFakIzn.Exists(payKey) Then duguje = CDbl(dFakIzn(payKey))
                End If
                If dPay.Exists(payKey) Then placeno = CDbl(dPay(payKey))
                pCode = PayCode(duguje, placeno)
            End If
            pRest = duguje - placeno
            If pRest < 0 Then pRest = 0
        End If

        For c = 0 To mColN - 1
            Select Case kind(c)
                Case "txt"
                    cell = CellS(src, r, ix(c))
                    hay = hay & "|" & cell
                Case "part"
                    cell = CellS(src, r, ix(c))
                    If rev Then
                        cell = RevPartner(CellS(src, r, iEntTip), CStr(cell), mKoop, mStan, mKup)
                    ElseIf kanal Then
                        cell = NovacPartner(CellS(src, r, iEntTip), CellS(src, r, iKoopID), _
                                            CellS(src, r, iPartID), CellS(src, r, iOtkID), _
                                            CStr(cell), mKoop, mStan, mKup)
                    ElseIf Not pmap Is Nothing Then
                        If pmap.Exists(cell) Then cell = pmap(cell)
                    End If
                    hay = hay & "|" & cell
                Case "date"
                    vDatK = CellDate(src, r, ix(c))
                    cell = vDatK
                Case "kg", "num"
                    cell = CellD(src, r, ix(c))
                Case "sum0", "rsd"
                    cell = CellD(src, r, ix(c))
                    ' F5 i F6 dele tblNovac - red bez iznosa u SVOJOJ koloni
                    ' pripada drugom smeru i odbacuje se ovde, ne u filteru
                    If cell = 0 Then GoTo NextRow
                Case "mult"
                    cell = CellD(src, r, ix(c)) * vKgRow
                Case "paypill"
                    cell = pCode
                Case "rest"
                    ' ostatak ima smisla samo dok nije placeno do kraja
                    If pCode = PAY_DELIM Or pCode = PAY_NEPLAC Then
                        cell = pRest
                    Else
                        cell = 0
                    End If
                Case "osnov"
                    cell = OsnovNaziv(CellS(src, r, iDokTip), CellS(src, r, iBrojCol))
                    hay = hay & "|" & cell
                Case "kanal"
                    cell = KanalNaziv(KanalCode(CellS(src, r, iTip), CellS(src, r, iNap)))
                    hay = hay & "|" & cell
                Case Else
                    cell = ""
            End Select
            outA(n + 1, c + 1) = cell
        Next c

        vZbir = CellS(src, r, iZbir)
        isStorno = (iStorno > 0)
        If isStorno Then isStorno = (UCase$(CellS(src, r, iStorno)) = "DA")
        bezZbirne = (Len(vZbir) = 0)
        hay = hay & "|" & vZbir

        bezFakture = False
        If iFakt > 0 Then bezFakture = (Len(CellS(src, r, iFakt)) = 0)

        ' brojaci cipova - isti prolaz, bez zasebnog skena po cipu.
        ' "Bez zbirne" i "Nefakturisane" su RAZLICITI uslovi i broje se odvojeno.
        If isStorno Then
            mCntOtkaz = mCntOtkaz + 1
        Else
            If bezZbirne Then mCntBezZb = mCntBezZb + 1
            If bezFakture Then mCntNefakt = mCntNefakt + 1
        End If

        keep = MatchFilterFast(filter, vDatK, bezZbirne, isStorno, bezFakture)
        If keep And Len(q) > 0 Then keep = (InStr(1, hay, q, vbTextCompare) > 0)
        ' dodatni uslovi iz panela Filteri - isti "hay", bez drugog prolaza
        If keep And Len(mFltVrsta) > 0 Then keep = (InStr(1, hay, mFltVrsta, vbTextCompare) > 0)
        If keep And Len(mFltPart) > 0 Then keep = (InStr(1, hay, mFltPart, vbTextCompare) > 0)

        If keep Then
            n = n + 1
            For c = 0 To mColN - 1
                Select Case kind(c)
                    Case "kg":                 mSumKg = mSumKg + CDbl(outA(n, c + 1))
                    Case "rsd", "mult", "sum0": mSumVal = mSumVal + CDbl(outA(n, c + 1))
                    Case "pill":               outA(n, c + 1) = StatusCode(isStorno, bezZbirne)
                End Select
            Next c
        End If
NextRow:
    Next r

    mFillStep = "sortiranje"
    mViewN = n
    If n = 0 Then Exit Sub
    mView = SortedView(outA, n, mColN, mSortCol, mSortAsc)
    mFillStep = "OK"
    Exit Sub
EH:
    ' greska se NE guta - ReloadGrid je prijavljuje sa imenom koraka
    Err.Raise Err.Number, "modOtkupUI.FillGrid[" & mFillStep & "]", Err.description
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
    mCols = GridCols(mk)
    mColN = UBound(mCols) + 1
    If mColN > MAX_COLS Then mColN = MAX_COLS
End Sub

Private Function MatchFilterFast(ByVal filter As String, ByVal vDatK As Double, _
                                 ByVal bezZbirne As Boolean, ByVal isStorno As Boolean, _
                                 ByVal bezFakture As Boolean) As Boolean
    Select Case filter
        Case "otkazane":  MatchFilterFast = isStorno
        Case "bezzbirne": MatchFilterFast = (Not isStorno) And bezZbirne
        Case "nefakt":    MatchFilterFast = (Not isStorno) And bezFakture
        Case "danas":     MatchFilterFast = (Not isStorno) And (vDatK = mToday)
        Case "nedelja":   MatchFilterFast = (Not isStorno) And (vDatK >= mToday - 6)
        Case "mesec":     MatchFilterFast = (Not isStorno) And (vDatK >= mMonthStart)
        Case Else:        MatchFilterFast = Not isStorno
    End Select
End Function

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

Private Sub ReloadGrid()
    If mBusyGrid Then Exit Sub
    If mBuilding Then Exit Sub
    mBusyGrid = True
    On Error GoTo EH
    mToday = Int(Now)
    mMonthStart = CDbl(DateSerial(Year(Now), Month(Now), 1))
    FillGrid ModeTable(ActiveMode), mFilter, mSearch
    mPage = 1
    RenderGrid
    mBusyGrid = False
    Exit Sub
EH:
    ' Ranije je ovde stajalo On Error GoTo Fin, pa je svaka greska u FillGrid-u
    ' tiho preskakala RenderGrid - mreza bi ostala na starom rezimu i izgledalo
    ' bi da "klik ne radi". Sada se vidi gde je puklo.
    mBusyGrid = False
    Debug.Print "modOtkupUI.ReloadGrid PAO [" & ActiveMode & " / " & _
                ModeTable(ActiveMode) & " / korak " & mFillStep & "]: " & _
                Err.Number & " " & Err.description
    ShowToast Poruka("OTKUI_MSG_MREZA_PALA") & " " & mFillStep & " (" & Err.Number & ")", True
End Sub

' Rezimi kojima je 4. kolona TEKST, a ne broj: gotovinski promet nosi kanal,
' reversi nose tip ambalaze. CompareKey pada na StrComp kad vrednost nije broj,
' pa sortiranje po toj koloni radi i za tekst - nije potrebno sedmo polje.
' Ima li aktivni rezim kolonu u kilogramima? Od toga zavisi zbir u podnozju.
Private Function ModeHasKgCol() As Boolean
    Dim i As Long
    If Not IsArray(mCols) Then Exit Function
    For i = 0 To UBound(mCols)
        If ColF(CStr(mCols(i)), 2) = "kg" Then ModeHasKgCol = True: Exit Function
    Next i
End Function

Private Function ModeHasValCol() As Boolean
    Dim i As Long
    If Not IsArray(mCols) Then Exit Function
    For i = 0 To UBound(mCols)
        Select Case ColF(CStr(mCols(i)), 2)
            Case "rsd", "mult", "sum0": ModeHasValCol = True: Exit Function
        End Select
    Next i
End Function

' Citljiv osnov reda. DOK_TIP_OTKUP je tu jer kooperant i pri predaji PUNIH
' gajbi ima izlaz ambalaze - to nije revers, ali jeste njegovo kretanje.
Private Function OsnovNaziv(ByVal dokTip As String, ByVal dokID As String) As String
    Dim izOtkupa As Boolean
    On Error Resume Next
    izOtkupa = OtkupKoopMap().Exists(Trim$(dokID))
    Select Case Trim$(dokTip)
        Case DOK_TIP_OM_IZLAZ_KOOP
            ' Prazne gajbe uz otkup se knjize ISTIM tipom kao pravi revers
            ' (modOtkup.bas:611), samo im je DokumentID = OtkupID. Bez ove
            ' razlike bi dva reda istog otkupa nosila razlicit osnov: jedan
            ' "Uz otkup", drugi "Revers" - a revers dokument ne postoji.
            OsnovNaziv = IIf(izOtkupa, Poruka("OTKUI_OSN_OTKUP_PRAZNE"), _
                                       Poruka("OTKUI_OSN_REV_IZDATO"))
        Case DOK_TIP_OM_ULAZ_KOOP:   OsnovNaziv = Poruka("OTKUI_OSN_REV_POVRAT")
        Case DOK_TIP_OM_IZLAZ_FIRMA: OsnovNaziv = Poruka("OTKUI_OSN_REV_OM_IZDATO")
        Case DOK_TIP_OM_ULAZ_FIRMA:  OsnovNaziv = Poruka("OTKUI_OSN_REV_OM_PRIJEM")
        Case DOK_TIP_OTKUP:          OsnovNaziv = Poruka("OTKUI_OSN_OTKUP_PUNE")
        Case Else:                   OsnovNaziv = dokTip
    End Select
End Function

' FakturaID -> Iznos. Prijemnica ne nosi svoj dug nego ga nasledjuje od fakture,
' pa se ostatak racuna NA NIVOU FAKTURE. Ako jedna faktura pokriva vise
' prijemnica, isti ostatak stoji u svakom njenom redu - to je tacno, ali se
' odnosi na fakturu, ne na pojedinacnu prijemnicu.
Private Function FakturaIznosMap() As Object
    Dim d As Object, src As Variant, iId As Long, iIzn As Long, r As Long, k As String
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists("#FAKIZN") Then
        Set FakturaIznosMap = mPartMap("#FAKIZN")
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    src = CachedTable(TBL_FAKTURE)
    If IsArray(src) Then
        iId = ColIdx(TBL_FAKTURE, COL_FAK_ID)
        iIzn = ColIdx(TBL_FAKTURE, COL_FAK_IZNOS)
        If iId > 0 And iIzn > 0 Then
            For r = 1 To UBound(src, 1)
                k = CellS(src, r, iId)
                If Len(k) > 0 Then d(k) = CellD(src, r, iIzn)
            Next r
        End If
    End If
    Set mPartMap("#FAKIZN") = d
    Set FakturaIznosMap = d
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

' Kolona Partner u tblNovac NIJE primalac novca. SaveNovac se za isplatu
' kooperantu poziva sa partner:=naziv OTKUPNOG MESTA, entitetTip:="OM",
' omID:=stanicaID, a stvarni primalac ide u KooperantID (modDokumenta:3791,
' modBankaMapiranje.MapBankaImportAsKooperant). Zato se ime vuce iz KooperantID
' kad postoji; tek ako ga nema, red se odnosi na sam OM ili na kupca.
' OtkupID -> KooperantID. Isti pristup koji koristi pregled otkupnih blokova
' (modBankaExportPregled: BuildLookupDict(TBL_OTKUP, COL_OTK_ID, COL_OTK_KOOPERANT)),
' jednom umesto LookupValue po redu. Sluzi dvema stvarima: da se primalac isplate
' nadje i kad red novca nema KooperantID, i da se prepozna da li je red ambalaze
' nastao iz otkupa (DokumentID je tada OtkupID) ili iz zasebnog reversa.
Private Function OtkupKoopMap() As Object
    On Error Resume Next
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists("#OTKKOOP") Then
        Set OtkupKoopMap = mPartMap("#OTKKOOP")
        Exit Function
    End If
    Dim d As Object
    Set d = BuildLookupDict(TBL_OTKUP, COL_OTK_ID, COL_OTK_KOOPERANT)
    If d Is Nothing Then Set d = CreateObject("Scripting.Dictionary")
    Set mPartMap("#OTKKOOP") = d
    Set OtkupKoopMap = d
End Function

Private Function NovacPartner(ByVal entTip As String, ByVal koopID As String, _
                              ByVal partID As String, ByVal otkID As String, _
                              ByVal partTekst As String, _
                              mKoop As Object, mStan As Object, mKup As Object) As String
    If Len(Trim$(koopID)) > 0 Then
        If Not mKoop Is Nothing Then
            If mKoop.Exists(koopID) Then NovacPartner = mKoop(koopID): Exit Function
        End If
        NovacPartner = koopID
        Exit Function
    End If
    ' red vezan za otkup, a bez KooperantID -> primaoca daje sam otkup
    If Len(Trim$(otkID)) > 0 Then
        Dim k As String
        k = ""
        If OtkupKoopMap().Exists(Trim$(otkID)) Then k = CStr(OtkupKoopMap()(Trim$(otkID)))
        If Len(k) > 0 Then
            If Not mKoop Is Nothing Then
                If mKoop.Exists(k) Then NovacPartner = mKoop(k): Exit Function
            End If
            NovacPartner = k
            Exit Function
        End If
    End If
    Dim d As Object
    Select Case Trim$(entTip)
        Case "Kupac":     Set d = mKup
        Case "OM":        Set d = mStan
        Case "Kooperant": Set d = mKoop
    End Select
    If Not d Is Nothing Then
        If d.Exists(partID) Then NovacPartner = d(partID): Exit Function
    End If
    NovacPartner = partTekst
End Function

' Najnovije zbirne u combo BROJ ZBIRNE. Zamena za lstZbirne iz frmDokumenta:
' tamo je klik na red upisivao broj u prijemnicu, ovde se bira iz panela nad
' samim poljem koje taj broj i cuva. Broj zbirne je oblika "x/ddmmyy" pa nosi
' datum u sebi - zato je sam broj dovoljan za prepoznavanje.
Public Sub FillZbirneCombo(frm As Object)
    Dim CB As MSForms.ComboBox, src As Variant, iBroj As Long, iDat As Long
    Dim r As Long, n As Long, arr() As Variant, key() As Double, i As Long, j As Long
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgBrZbir").Controls("fgBrZbirT")
    CB.Clear
    src = CachedTable(TBL_ZBIRNA)
    If Not IsArray(src) Then Exit Sub
    iBroj = ColIdx(TBL_ZBIRNA, COL_ZBR_BROJ)
    iDat = ColIdx(TBL_ZBIRNA, COL_ZBR_DATUM)
    If iBroj < 1 Then Exit Sub

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
    If j = 0 Then Exit Sub

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
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillZbirneCombo PAO: " & Err.Number & " " & Err.description
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

Private Function PartnerMap(ByVal tblName As String, ByVal idCol As String, _
                            ByVal nameCol As String, ByVal nameCol2 As String) As Object
    Dim d As Object, src As Variant, iId As Long, iNm As Long, iNm2 As Long
    Dim r As Long, k As String, v As String
    If mPartMap Is Nothing Then Set mPartMap = CreateObject("Scripting.Dictionary")
    If mPartMap.Exists(tblName) Then
        Set PartnerMap = mPartMap(tblName)
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
    Set mPartMap(tblName) = d
    Set PartnerMap = d
End Function

' Koje tipove partnera nudi polje. Tamo gde dokument zna ko je druga strana
' lista je jednorodna (samo kupci); tamo gde ne zna - isplata ide i kooperantu
' i otkupnom mestu, revers svima, storno preko svih dokumenata - lista je
' mesovita i natpis polja je "Partner". Redosled = najverovatniji tip prvi.
Private Function PartnerSrcOrder(ByVal mode As String) As Variant
    Select Case mode
        Case "F1": PartnerSrcOrder = Array("KOOP")
        Case "F5": PartnerSrcOrder = Array("KOOP", "OM")
        Case "F7": PartnerSrcOrder = Array("KOOP", "OM", "KUP")
        Case "F8": PartnerSrcOrder = Array("OM", "KOOP", "KUP")
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
    On Error GoTo EH
    Set CB = frm.Controls("zCtx").Controls("cbKupac")
    If mPartnerFor = mode Then Exit Sub          ' isti redosled, ne puni ponovo
    CB.Clear
    CB.ColumnCount = 2
    ' BEZ ListWidth-a i BEZ ListRows - lista se otvara tacno u sirini kontrole,
    ' kao kod svih ostalih dropdown-a; svaka druga vrednost napravi zaseban,
    ' siri okvir koji se vidi kao "druga forma". Kolona mora da stane u
    ' kontrolu umanjenu za vertikalni skroler, inace se javi i horizontalni.
    CB.ColumnWidths = "146 pt;0 pt"
    CB.BoundColumn = 1
    CB.TextColumn = 1
    ord = PartnerSrcOrder(mode)
    For i = LBound(ord) To UBound(ord)
        sd = PartnerSrcDef(CStr(ord(i)))
        Set d = PartnerMap(CStr(sd(0)), CStr(sd(1)), CStr(sd(2)), CStr(sd(3)))
        Set neakt = NeaktivniSet(CStr(sd(0)))
        If Not d Is Nothing Then
            For Each k In d.keys
                If Not neakt.Exists(CStr(k)) Then
                    CB.AddItem PartnerPrikaz(CStr(ord(i)), CStr(k), CStr(d(k)))
                    CB.List(CB.ListCount - 1, 1) = CStr(k)
                End If
            Next k
        End If
    Next i
    mPartnerFor = mode
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillFormPartner PAO [" & mode & "]: " & Err.Number & " " & Err.description
End Sub

' ID izabranog partnera - druga, skrivena kolona combo-a (isti raspored
' kolona kao svuda drugde: modComboBinding.GetComboID).
Private Function PartnerID(frm As Object) As String
    Dim CB As MSForms.ComboBox
    On Error Resume Next
    Set CB = frm.Controls("zCtx").Controls("cbKupac")
    If CB.ListIndex >= 0 Then PartnerID = Trim$(CStr(CB.List(CB.ListIndex, 1)))
End Function

' "Preostali kes" iz frmDokumenta (cmbOtkupBlok): otkupni blokovi izabranog
' primaoca koji nisu isplaceni do kraja. Racunicu NE ponavljamo ovde - koristi
' se postojeci modNovac.GetOpenOtkupi (7 kolona: brDok | OtkupID | vrednost |
' isplaceno | ostatak | datum | StanicaID), isti izvor koji koristi i pregled
' otvorenih blokova, pa se dva ekrana ne mogu raziici.
Private Sub FillOpenBlokovi(frm As Object)
    Dim CB As MSForms.ComboBox, koopID As String, o As Variant, i As Long
    On Error GoTo EH
    Set CB = frm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT")
    CB.Clear
    mBlokIDs = Empty
    mBlokRest = Empty
    SetOstatak 0
    If ActiveMode <> "F5" Then Exit Sub
    koopID = PartnerID(frm)
    If Len(koopID) = 0 Then Exit Sub
    o = GetOpenOtkupi(koopID)
    If IsEmpty(o) Then Exit Sub
    ReDim mBlokIDs(1 To UBound(o, 1))
    ReDim mBlokRest(1 To UBound(o, 1))
    For i = 1 To UBound(o, 1)
        mBlokIDs(i) = CStr(o(i, 2))
        mBlokRest(i) = CDbl(o(i, 5))
        CB.AddItem CStr(o(i, 1)) & "   " & ChrW(183) & "   " & _
                   FmtBroj(CDbl(o(i, 5)), 0) & " / " & FmtBroj(CDbl(o(i, 3)), 0)
    Next i
    Exit Sub
EH:
    Debug.Print "modOtkupUI.FillOpenBlokovi PAO: " & Err.Number & " " & Err.description
End Sub

' Izbor bloka puni plocu NEISPLACENO - iznos koji jos moze da se isplati.
Private Sub ApplyBlokIzbor()
    Dim CB As MSForms.ComboBox, i As Long
    On Error Resume Next
    Set CB = mFrm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT")
    i = CB.ListIndex + 1
    If i < 1 Then SetOstatak 0: Exit Sub
    If Not IsArray(mBlokRest) Then SetOstatak 0: Exit Sub
    If i > UBound(mBlokRest) Then SetOstatak 0: Exit Sub
    SetOstatak CDbl(mBlokRest(i))
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

Private Function PartnerLookup(ByVal mode As String) As Variant
    Select Case mode
        Case "F2", "F8":  PartnerLookup = Array(TBL_STANICE, "StanicaID", "Naziv", "")
        Case "F3", "F4":  PartnerLookup = Array(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV, "")
        Case "F1":        PartnerLookup = Array(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
        Case Else:        PartnerLookup = Array("", "", "", "")
                          ' F5/F6: tblNovac vec ima tekstualnu kolonu Partner
    End Select
End Function

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

Private Function CachedTable(ByVal tblName As String) As Variant
    If mCache Is Nothing Then Set mCache = CreateObject("Scripting.Dictionary")
    If Not mCache.Exists(tblName) Then mCache(tblName) = GetTableData(tblName)
    CachedTable = mCache(tblName)
End Function

Private Function ColIdx(ByVal tblName As String, ByVal colName As String) As Long
    If Len(colName) = 0 Then Exit Function
    On Error Resume Next
    ColIdx = GetColumnIndex(tblName, colName)
End Function

Private Function CellS(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsEmpty(v) Then Exit Function
    CellS = Trim$(CStr(v))
End Function

Private Function CellD(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then CellD = CDbl(v)
End Function

Private Function CellDate(ByRef src As Variant, ByVal r As Long, ByVal c As Long) As Double
    If c < 1 Then Exit Function
    Dim v As Variant: v = src(r, c)
    If IsNumeric(v) Then
        CellDate = Int(CDbl(v))
    ElseIf IsDate(v) Then
        CellDate = Int(CDbl(CDate(v)))
    End If
End Function

Private Function FmtDatumKratko(ByVal v As Variant) As String
    If Not IsNumeric(v) Then Exit Function
    If CDbl(v) <= 0 Then Exit Function
    FmtDatumKratko = Format$(CDate(CDbl(v)), "dd.mm.")
End Function

Private Function FmtDatumPun(ByVal v As Variant) As String
    On Error Resume Next
    FmtDatumPun = Format$(CDate(v), "dd.mm.yyyy")
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
Private Function IsDebugUI() As Boolean
    On Error Resume Next
    If StrComp(GetLocalConfigValue("UI_DEBUG", ""), "DA", vbTextCompare) = 0 Then
        IsDebugUI = True
        Exit Function
    End If
    IsDebugUI = (InStr(1, BuildTagOrBlank(), "dev", vbTextCompare) > 0)
End Function

Private Function BuildTagOrBlank() As String
    On Error Resume Next
    BuildTagOrBlank = "v" & BUILD_VERSION          ' modBuildInfo (stamp-build.sh)
    If Len(BuildTagOrBlank) <= 1 Then BuildTagOrBlank = "-"
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

Public Sub ShowToast(ByVal msg As String, ByVal isErr As Boolean)
    Dim fr As Object
    On Error Resume Next
    If mFrm Is Nothing Then Exit Sub
    Set fr = mFrm.Controls("zForm").Controls("tstOk")
    fr.BackColor = IIf(isErr, RGB(250, 240, 235), RGB(236, 245, 240))
    fr.Controls("tstBar").BackColor = IIf(isErr, C_RUST, C_GREEN_LT)
    fr.Controls("tstMsg").caption = msg
    fr.Controls("tstMsg").ForeColor = IIf(isErr, C_RUST, C_GREEN)
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
    mFrm.Controls("zForm").Controls("tstOk").Visible = False
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
End Sub

Public Sub OtkupUI_Release()
    On Error Resume Next
    StopOtkupUITimers
    Set mHovered = Nothing
    Set Btns = Nothing
    modUiKit.ResetNumFields             ' drzi reference na kontrole oborene forme
    Set mFrm = Nothing
    Set mCache = Nothing
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
        Case vbKeyF8: SelectMode mFrm, "F8"
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
    Dim kg As Double, cena As Double
    kg = ParseNum(FldText("fgKgI"))
    If mKlasa = 2 Then kg = kg + ParseNum(FldText("fgKgII"))
    cena = CenaText(1)

    If Len(Trim$(FldText("fgBrOtpr"))) = 0 Then
        ShowToast Poruka("OTKUI_ERR_BROJ"), True: Exit Sub
    End If
    If kg <= 0 Then ShowToast Poruka("OTKUI_ERR_KOLICINA"), True: Exit Sub
    If cena <= 0 Then ShowToast Poruka("OTKUI_ERR_CENA"), True: Exit Sub

    ' SEAM: ovde ide postojeci upis (modOtkup / modDokumenta / clsTransaction).
    ' Namerno NIJE vezano - prototip ne sme da knjizi.
    MarkClean
    ShowToast Poruka("OTKUI_MSG_SNIMLJENO") & IIf(alsoPrint, " " & ChrW(183) & " " & Poruka("OTKUI_MSG_PRINT"), ""), False
End Sub

Private Sub ClearForm()
    Dim nmv As Variant, i As Long
    On Error Resume Next
    mPopMute = True
    mLoading = True
    nmv = Array("fgBrOtpr", "fgBrZbir", "fgKgI", "fgKgII", "fgKolAmb", "fgAmbPr")
    For i = 0 To UBound(nmv)
        mFrm.Controls("zForm").Controls(CStr(nmv(i))).Controls(CStr(nmv(i)) & "T").text = ""
    Next i
    ' cena ima dve kutije pa ne staje u sablon grp -> grp & "T"
    mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena1T").text = ""
    mFrm.Controls("zForm").Controls("fgCena").Controls("fgCena2T").text = ""
    mFrm.Controls("zForm").Controls("fgBlok").Controls("fgBlokT").text = ""
    mFrm.Controls("zForm").Controls("fgParcela").Controls("fgParcelaT").text = ""
    SetOstatak 0
    SetKlasa 1
    RecalcVrednost
    mPopMute = False
    mLoading = False
    MarkClean
    mFrm.Controls("zForm").Controls("fgBrOtpr").Controls("fgBrOtprT").SetFocus
End Sub

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
    PaintVrednost v1 + v2
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


Private Function ChipFilter(ByVal tag As String) As String
    Select Case tag
        Case "chipDanas":     ChipFilter = "danas"
        Case "chipNedelja":   ChipFilter = "nedelja"
        Case "chipSve":       ChipFilter = "sve"
        Case "chipNefakt":    ChipFilter = "nefakt"
        Case "chipBezZbirne": ChipFilter = "bezzbirne"
        Case "chipOtkazane":  ChipFilter = "otkazane"
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
    DiagCols "F8"
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
    IcoRow a, "IC_REFRESH", "(nije vezano)", "Refresh", IC_REFRESH

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
          " + clsFlatBtn " & IIf(Len(cv) = 0, "(stara)", cv)
    If Len(cv) = 0 Then
        OtkupUI_SelfCheck = ver & vbCrLf & _
            "PAZNJA: clsFlatBtn je STARA verzija (uvoz nije zamenio klasu; " & _
            "trazi komponentu clsFlatBtn1 u Project Exploreru)"
    ElseIf cv <> OTKUI_BUILD Or UIKIT_BUILD <> OTKUI_BUILD _
           Or SCRDOK_BUILD <> OTKUI_BUILD Then
        OtkupUI_SelfCheck = ver & vbCrLf & "PAZNJA: verzije se ne poklapaju"
    Else
        OtkupUI_SelfCheck = ver & vbCrLf & OtkupUI_Stats()
    End If
End Function

' Merilo rasta ekrana. Plan je da SVI ekrani predju na ovu jednu formu, pa je
' broj kontrola resurs koji se trosi - a granicu MSForms-a ne znamo unapred.
' Zato se meri od pocetka: koliko kontrola i koliko traje gradnja.
Public Function OtkupUI_Stats() As String
    Dim n As Long, sinks As Long
    On Error Resume Next
    If mFrm Is Nothing Then
        OtkupUI_Stats = "ekran nije izgradjen"
        Exit Function
    End If
    n = CountCtls(mFrm)
    If Not Btns Is Nothing Then sinks = Btns.count
    OtkupUI_Stats = "kontrola: " & n & "  |  vezanih sinkova: " & sinks & _
                    "  |  gradnja: " & mBuildMs & " ms"
End Function

Private Function CountCtls(ByVal parent As Object) As Long
    Dim c As Object, n As Long
    On Error Resume Next
    For Each c In parent.Controls
        n = n + 1
        If TypeName(c) = "Frame" Then n = n + CountCtls(c)
    Next c
    CountCtls = n
End Function
'=== KRAJ FAJLA modOtkupUI ===
