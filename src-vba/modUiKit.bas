Attribute VB_Name = "modUiKit"
'=====================================================================
' modUiKit - FABRIKA KONTROLA novog UI-ja (faza S2).
'
' Sve sto crta, meri i pomera: okviri i zone, natpisi, polja, combo,
' kutije sa ivicom, dugmad, cipovi, kartice, segmenti, zaglavlja kolona,
' celije mreze, strelice - plus merenje teksta (TxtH/LineH/CenterY) i
' interpolacija boja (Lerp).
'
' Ovde NEMA nijednog imena ekrana, tabele ili polja. Modul ne zna sta se
' crta, samo kako. Nov ekran koristi ove fabrike i ne dodaje svoje.
'
' Paleta (C_*), tipografija (TS_*, F_*) i vezivanje dogadjaja (WireBtn /
' WireInput) ostaju u modOtkupUI. Razlog nije estetski: clsFlatBtn je
' ZAMRZNUTA i imenom se veze za modOtkupUI.C_GREEN, .C_FOREST, .C_WHITE,
' .C_GOLD, .C_HDR_SURF i .C_BTN_HOVER. Premestanje tih konstanti oborilo
' bi compile klase, a klasa se po S0 vise ne dira. modOtkupUI je zato
' fasada: paleta + ruter dogadjaja + ekran; modUiKit je alat.
' VBA razresava nekvalifikovana javna imena kroz module, pa se telo
' fabrika nije menjalo - C_WHITE i TS_BODY se i dalje pisu bez prefiksa.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UIKIT_BUILD As String = "v6-ui-209"

' Brojcana polja (NewTxt isNum:=True) -> kontrola. Filter unosa mora da zna
' i KOJE je polje i sta u njemu vec stoji, pa se cuva sama kontrola, ne
' zastavica. Puni ga NewTxt, cita FilterKeyPress (preko modOtkupUI.UiAsk).
Private mNumTag As Object

' Formu su srusili - reference na njene kontrole vise ne vrede.
Public Sub ResetNumFields()
    Set mNumTag = Nothing
End Sub

' U brojcanom polju slovo nema sta da trazi. Filter radi SAMO nad poljima koja
' su napravljena kao brojcana (NewTxt isNum:=True) - pretraga, brojevi
' dokumenata i ostali tekst prolaze nedirnuti.
Public Function FilterKeyPress(ByVal tag As String, ByVal k As Long) As Long
    Dim ch As String, sep As String, t As Object
    FilterKeyPress = k
    If mNumTag Is Nothing Then Exit Function
    If Not mNumTag.Exists(tag) Then Exit Function
    If k < 32 Then Exit Function         ' Backspace / Enter / Tab ne idu ovuda
    ch = ChrW(k)
    If ch >= "0" And ch <= "9" Then Exit Function
    Set t = mNumTag(tag)
    sep = LocalDecSep()
    If ch = "." Or ch = "," Then
        ' numericka tastatura daje tacku, a Excel ocekuje lokalni separator;
        ' drugi separator se ne prima (osim ako zamenjuje oznaceni deo)
        If InStr(1, CStr(t.text), sep) > 0 And t.SelLength = 0 Then
            FilterKeyPress = 0
        Else
            FilterKeyPress = AscW(sep)
        End If
        Exit Function
    End If
    If ch = "-" Then
        If t.SelStart = 0 Then Exit Function
    End If
    FilterKeyPress = 0
End Function

' Decimalni znak ovog Windows-a, bez citanja registry-ja: CStr(1.5) ga vrati.
Public Function LocalDecSep() As String
    LocalDecSep = Mid$(CStr(1.5), 2, 1)
End Function

'=====================================================================
' PRIMITIVI - jedan izgled za sve komponente
'=====================================================================
' MSForms Label NEMA vertikalno poravnanje - tekst se uvek crta uz GORNJU
' ivicu kontrole. Sto je kutija visa od fonta, to je odstupanje vidljivije
' (npr. "kg" u 28pt kutiji visi na vrhu, ikonica sidebara je bila 6pt ispod
' teksta stavke). Zato:
'   - tekstualni Label dobija visinu po fontu (TxtH)
'   - Top mu se racuna kao centar kutije u kojoj stoji (CenterY)
'   - pozadinu, kad treba, crta ZASEBAN Label ispod njega
' Visina kutije za tekst. Bilo je fs*1.6+3 - to je bilo znatno vise od stvarne
' visine linije (Segoe UI ~1.33em), a MSForms Label tekst PORAVNAVA UZ GORNJU
' IVICU. Zato je CenterY centrirao kutiju, a tekst u njoj ostajao visoko:
' na dugmetu 22pt sa fontom 8.5 kutija je bila 16pt, glifovi ~11pt, pa je tekst
' visio oko 2.5pt iznad sredine. Isto se videlo i na ikonici u dugmetu.
' Sada je kutija tesno oko linije (1.4em + 1), pa centriranje kutije znaci i
' centriranje teksta. 1.4em je i dalje iznad 1.33em, tako da se donji delovi
' slova (p, g, j) ne odsecaju.
Public Function TxtH(ByVal fs As Single) As Single
    TxtH = Int(fs * 1.4) + 1
End Function

' Stvarna visina linije teksta (Segoe UI ~1.33em). Kutija je namerno malo visa
' od ovoga da se ne odsecaju p/g/j, pa se za CENTRIRANJE mora koristiti linija,
' a ne kutija - inace ostaje sistematska greska od pola razlike.
' MDL2 glif zauzima celu em kutiju (~fs), a ne visinu tekstualne linije
' (~1.33*fs). Centriranje po LineH ga zato dize par tacaka previsoko - vidi se
' na dugmadima sa ikonicom. Ikonice se centriraju ovim.
Public Function CenterIco(ByVal boxTop As Single, ByVal boxH As Single, ByVal fs As Single) As Single
    CenterIco = boxTop + Int((boxH - fs) / 2)
    If CenterIco < boxTop Then CenterIco = boxTop
End Function

Public Function LineH(ByVal fs As Single) As Single
    LineH = fs * 1.33
End Function

Public Function CenterY(ByVal boxTop As Single, ByVal boxH As Single, ByVal fs As Single) As Single
    ' Centrira se LINIJA teksta, ne kutija - MSForms tekst lepi uz gornju ivicu
    ' kutije, pa centriranje kutije ostavlja tekst visoko.
    ' Int, ne obicno /2: pola tacke pomeraja gura tekst na podpiksel granicu i
    ' GDI ga tada rasterizuje drugacije od suseda (isti uzrok kao GRID_ROW_H).
    CenterY = boxTop + Int((boxH - LineH(fs)) / 2)
    ' ako je kutija NIZA od teksta, centriranje bi ga izbacilo iznad ivice
    ' (tako su brojevi stranica visili van svojih dugmadi)
    If CenterY < boxTop Then CenterY = boxTop
End Function

' Kutija sa tekstom: ivica + ispuna + vertikalno centriran tekst.
' Vraca ISPUNU - ona je hit/hover meta i nosi BackColor stanja.
' Tekst je <nm>C i nosi ForeColor/Bold; ivica je <nm>B.
Public Function BoxText(parent As Object, nm As String, cap As String, _
                         X As Single, Y As Single, w As Single, h As Single, _
                         fs As Single, bold As Boolean, fc As Long, bg As Long, bd As Long, _
                         ByVal kind As String) As Object
    Dim fill As Object
    NewLbl parent, nm & "B", "", X, Y, w, h, 8, False, 0, bd
    Set fill = NewLbl(parent, nm, "", X + 1, Y + 1, w - 2, h - 2, 8, bold, fc, bg)
    NewLbl parent, nm & "C", cap, X + 1, CenterY(Y, h, fs), w - 2, TxtH(fs), _
           fs, bold, fc, -1, fmTextAlignCenter
    parent.Controls(nm & "C").ZOrder 0
    WireBtn fill, nm, kind
    ' Tekst stoji IZNAD ispune pa on hvata klikove u sredini kutije - mora i
    ' on da bude ozicen, sa ISTIM tagom. Vrsta "chev" znaci: klik prosledi,
    ' hover ne diraj (hover boji ispunu ispod, a tekst je providan).
    WireBtn parent.Controls(nm & "C"), nm, "chev"
    Set BoxText = fill
End Function

' REZ SE POTVRDJUJE, NE VERUJE MU SE.
'
' MSForms ume da izgubi upis Font.Bold: kontrola gradjena sa bold=False izlazi
' sa Font.Weight = 700. Mereno sondom u testu 89, sest krugova, i nijedno
' objasnjenje nije izdrzalo -- ni redosled (rez pre pozadine, rez posle
' pozadine, rez kroz ponovni lookup, rez dva puta), ni ozicenje, ni povratna
' vrednost, ni nasledjivanje fonta forme. Poslednje merenje je pokazalo i da
' PRVI upis nad izgradjenom kontrolom prolazi (400), a DRUGI isti takav upis
' vrati 700 -- dakle upis nije ni pouzdan ni idempotentan.
'
' Zato se ovde ne tvrdi ZASTO nego se tvrdi ISHOD: upisi, procitaj, i ako nije
' ono sto je trazeno -- upisi opet. Merilo je Font.Weight (400 normalan, 700
' bold), jer je Font.Bold iz nje izveden i sam po sebi ume da prevari.
'
' Petlja je ogranicena na tri pokusaja: kad ni to ne uhvati, kontrola ostaje
' kakva jeste. Ekran koji se zavrti u petlji je gori kvar od pogresnog reza.
Public Sub PostaviRez(ByVal ctl As Object, ByVal bold As Boolean)
    Dim i As Long
    On Error Resume Next
    For i = 1 To 3
        If (ctl.Font.Weight >= 700) = bold Then Exit Sub
        ctl.Font.bold = bold
    Next i
End Sub

' promena boja kutije: ispuna nosi pozadinu, <nm>C tekst
Public Sub BoxState(parent As Object, ByVal nm As String, ByVal bg As Long, _
                     ByVal fc As Long, ByVal bold As Boolean)
    On Error Resume Next
    parent.Controls(nm).BackColor = bg
    PostaviRez parent.Controls(nm), bold
    parent.Controls(nm & "C").ForeColor = fc
    PostaviRez parent.Controls(nm & "C"), bold
End Sub

Public Sub BoxShow(parent As Object, ByVal nm As String, ByVal vis As Boolean)
    On Error Resume Next
    parent.Controls(nm).Visible = vis
    parent.Controls(nm & "B").Visible = vis
    parent.Controls(nm & "C").Visible = vis
    ' i ikonica: BtnV je pravi kao ZASEBAN Label (jedan Label ne moze da drzi
    ' dva fonta). Bez ovog reda sakriveno dugme ostavlja glif da lebdi.
    parent.Controls(nm & "I").Visible = vis
End Sub

' SHELL: ivica + ispuna. Time se dobija padding koji MSForms nema i
' jedno mesto za focus/error stanje.
Public Sub NewShell(parent As Object, nm As String, X As Single, Y As Single, _
                     w As Single, h As Single, borderCol As Long, fillCol As Long)
    NewLbl parent, nm & "B", "", X, Y, w, h, 8, False, 0, borderCol
    NewLbl parent, nm & "F", "", X + 1, Y + 1, w - 2, h - 2, 8, False, 0, fillCol
End Sub

Public Sub MoveShell(parent As Object, nm As String, X As Single, Y As Single, w As Single)
    On Error Resume Next
    parent.Controls(nm & "B").Left = X: parent.Controls(nm & "B").top = Y
    parent.Controls(nm & "B").width = w
    parent.Controls(nm & "F").Left = X + 1: parent.Controls(nm & "F").top = Y + 1
    parent.Controls(nm & "F").width = w - 2
End Sub

Public Sub ShellState(parent As Object, ByVal nm As String, ByVal state As String)
    Dim b As Long, f As Long
    Select Case LCase$(state)
        Case "focus": b = C_GREEN: f = C_WHITE
        Case "error": b = C_RUST: f = RGB(253, 246, 243)
        Case "off":   b = C_BORDER_LT: f = C_DISABLED
        Case Else:    b = C_INPUT_BORDER: f = C_WHITE
    End Select
    On Error Resume Next
    parent.Controls(nm & "B").BackColor = b
    parent.Controls(nm & "F").BackColor = f
End Sub

' Jedno dugme za sve varijante - ujednacen izgled i stanja.
'
' icoChar: MDL2 kodna tacka (0 = bez ikonice). Ikonica MORA biti zaseban
' Label sa Font.Name = "Segoe MDL2 Assets" - jedan Label ne moze da drzi dva
' fonta, pa je ranije ChrW(&HE74E&) unutar caption-a crtan u Segoe UI i
' ispadao kao prazan kvadratic.
Public Function BtnV(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                      w As Single, h As Single, ByVal kindName As String, _
                      Optional ByVal icoChar As Long = 0) As Object
    Dim fc As Long, bg As Long, bd As Long, fs As Single, bold As Boolean
    Dim ix As Single, iw As Single, ial As Long
    Select Case kindName
        Case "primary":   fc = C_WHITE: bg = C_GREEN: bd = C_GREEN: fs = TS_ACTION: bold = True
        ' svetlo zelena ploca sa zelenim tekstom - ne takmici se sa primarnim
        Case "soft":      fc = C_GREEN: bg = C_SOFT_BG: bd = C_SOFT_BG: fs = TS_META: bold = True
        ' tercijarno: bez ivice, samo tekst (Otkazi)
        Case "plain":     fc = C_MUTED: bg = C_WHITE: bd = C_WHITE: fs = TS_META: bold = False
        Case "secondary": fc = C_FOREST: bg = C_WHITE: bd = C_BORDER: fs = TS_META: bold = False
        Case "ghost":     fc = C_MUTED: bg = C_WHITE: bd = C_BORDER_LT: fs = TS_META: bold = False
        Case "gold":      fc = C_FOREST: bg = C_GOLD: bd = C_GOLD: fs = TS_META: bold = True
        ' alatka na tamnom zaglavlju - vidljiva ploca, ne goli tekst
        Case "ghostdark": fc = C_CREAM: bg = C_HDR_SURF: bd = C_HDR_EDGE: fs = TS_META: bold = False
        ' isto, ali bez ploce (dugme za zatvaranje - ploca oko "X" je preteska)
        Case "ghostx":    fc = C_CREAM: bg = C_FOREST: bd = C_FOREST: fs = TS_META: bold = False
        Case "danger":    fc = C_WHITE: bg = C_RUST: bd = C_RUST: fs = TS_META: bold = True
        Case "page":      fc = C_FOREST: bg = C_WHITE: bd = C_BORDER_LT: fs = TS_META: bold = False
        Case Else:        fc = C_FOREST: bg = C_WHITE: bd = C_BORDER: fs = TS_META: bold = False
    End Select
    Dim L As Object
    Set L = BoxText(parent, nm, cap, X, Y, w, h, fs, bold, fc, bg, bd, "btn")
    If icoChar <> 0 Then
        If Len(cap) = 0 Then
            ix = X + 1: iw = w - 2: ial = fmTextAlignCenter    ' dugme samo sa ikonicom
        Else
            ix = X + 10: iw = 16: ial = fmTextAlignLeft
        End If
        NewLbl parent, nm & "I", ChrW(icoChar), ix, CenterIco(Y, h, fs), iw, TxtH(fs), _
               fs, False, fc, -1, ial, F_ICON
        parent.Controls(nm & "I").ZOrder 0
        ' Ikonica je iznad ispune (ZOrder 0) i kod dugmeta BEZ teksta pokriva
        ' celu njegovu povrsinu. Bez ovog vezivanja klik pada na mrtvu labelu i
        ' dugme "ne radi" - isti slucaj kao ranije kod kartica rezima.
        WireBtn parent.Controls(nm & "I"), nm, "chev"
        ' Tekst se centrira u delu DESNO od ikonice. Ranije se resavalo trima
        ' vodecim razmacima u caption-u, ali BoxText centrira ceo string - pa su
        ' i razmaci ulazili u centriranje i tekst je bezao u stranu.
        ' Uvlacenje se pamti u Tag-u da ga MoveBox postuje pri svakom rasporedu.
        If Len(cap) > 0 Then
            parent.Controls(nm & "C").tag = "ico:" & CStr(ICO_INSET)
            parent.Controls(nm & "C").Left = X + 1 + ICO_INSET
            parent.Controls(nm & "C").width = w - 2 - ICO_INSET
        End If
    End If
    Set BtnV = L
End Function

' pomeri kutiju (ivica + ispuna + centriran tekst)
Public Sub MoveBox(parent As Object, ByVal nm As String, X As Single, Y As Single, w As Single)
    Dim bh As Single
    On Error Resume Next
    bh = parent.Controls(nm & "B").Height
    parent.Controls(nm & "B").Left = X: parent.Controls(nm & "B").top = Y
    parent.Controls(nm & "B").width = w
    parent.Controls(nm).Left = X + 1: parent.Controls(nm).top = Y + 1
    parent.Controls(nm).width = w - 2
    ' dugme sa ikonicom: tekst se centrira desno od nje (Tag ga pamti)
    Dim ins As Single
    If Left$(parent.Controls(nm & "C").tag, 4) = "ico:" Then _
        ins = CSng(val(Mid$(parent.Controls(nm & "C").tag, 5)))
    parent.Controls(nm & "C").Left = X + 1 + ins
    parent.Controls(nm & "C").width = w - 2 - ins
    ' isto pravilo kao CenterY: centrira se linija teksta, ne kutija. Velicinu
    ' fonta uzimam sa same kontrole - MoveBox je ne dobija kao argument.
    parent.Controls(nm & "C").top = Y + Int((bh - LineH(parent.Controls(nm & "C").Font.Size)) / 2)
    If parent.Controls(nm & "C").top < Y Then parent.Controls(nm & "C").top = Y
End Sub

Public Sub MoveBtn(parent As Object, nm As String, X As Single, Y As Single)
    Dim bh As Single
    On Error Resume Next
    bh = parent.Controls(nm & "B").Height
    MoveBox parent, nm, X, Y, parent.Controls(nm & "B").width
    If HasCtl(parent, nm & "I") Then
        If parent.Controls(nm & "I").width > 20 Then
            parent.Controls(nm & "I").Left = X + 1        ' centrirana, dugme bez teksta
        Else
            parent.Controls(nm & "I").Left = X + 10
        End If
        parent.Controls(nm & "I").top = CenterIco(Y, bh, parent.Controls(nm & "I").Font.Size)
    End If
End Sub

Public Function ChipV(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                       w As Single, h As Single, sel As Boolean, warn As Boolean) As Object
    ' bez lokalnog l - BoxText vraca ispunu
    Set ChipV = BoxText(parent, nm, cap, X, Y, w, h, TS_META, sel, _
                        IIf(sel, C_CREAM, IIf(warn, C_RUST, C_FOREST)), _
                        IIf(sel, C_FOREST, IIf(warn, C_PILL_ERR_BG, C_WHITE)), _
                        IIf(warn, C_RUST, C_BORDER_LT), "chip")
End Function

Public Sub CardV(parent As Object, nm As String, X As Single, Y As Single, w As Single, h As Single)
    NewShell parent, nm, X, Y, w, h, C_BORDER_LT, C_WHITE
    NewLbl parent, nm & "A", "", X, Y, 3, h, 8, False, 0, C_BORDER
End Sub

' Eyebrow: 7.5pt, Semibold, verzal, prigusen. Isti rez za naslov sekcije,
' KPI natpis i naslov grupe polja - to je jedan tipografski nivo.
Public Function NewSectionHdr(parent As Object, nm As String, cap As String, _
                               X As Single, Y As Single, w As Single) As Object
    Set NewSectionHdr = NewLbl(parent, nm, UCase$(cap), X, Y, w, TxtH(TS_MICRO), TS_MICRO, True, C_MUTED, -1)
End Function

Public Function NewSegBtn(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                           w As Single, h As Single, sel As Boolean) As Object
    Set NewSegBtn = BoxText(parent, nm, cap, X, Y, w, h, TS_META, sel, _
                            IIf(sel, C_CREAM, C_MUTED), IIf(sel, C_FOREST, C_WHITE), _
                            IIf(sel, C_FOREST, C_BORDER_LT), "seg")
End Function

Public Function NewSortHdr(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                            w As Single, h As Single, rightAlign As Boolean) As Object
    Dim L As Object
    ' bez strelice u mirovanju - RefreshSortGlyphs je dodaje aktivnoj koloni
    ' i onoj pod pokazivacem
    Set L = NewLbl(parent, nm, UCase$(cap) & "   ", X, Y, w, h, TS_MICRO, True, C_MUTED, -1, _
                   IIf(rightAlign, fmTextAlignRight, fmTextAlignLeft))
    WireBtn L, nm, "hdr"
    Set NewSortHdr = L
End Function

Public Function NewNavItem(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                            w As Single, h As Single, sel As Boolean) As Object
    ' pozadina nosi boju i hit-zonu; tekst je zaseban i vertikalno centriran
    Dim bgL As Object, L As Object
    Set bgL = NewLbl(parent, nm, "", X, Y, w, h, 8, sel, 0, IIf(sel, C_FOREST, C_SAND))
    bgL.tag = cap
    Set L = NewLbl(parent, nm & "X", cap, X + 38, CenterY(Y, h, TS_NAV), w - 44, TxtH(TS_NAV), _
                   TS_NAV, sel, IIf(sel, C_CREAM, RGB(52, 68, 44)), -1)
    WireBtn bgL, nm, "nav"
    WireBtn L, nm, "nav"
    Set NewNavItem = bgL
End Function

' celija mreze; bg = -1 znaci providna
Public Function NewCell(parent As Object, nm As String, X As Single, Y As Single, _
                         w As Single, h As Single, bg As Long, fc As Long, _
                         fs As Single, fnt As String, align As Long) As Object
    Dim L As Object
    Set L = NewLbl(parent, nm, "", X, Y, w, h, fs, False, fc, bg, align, fnt)
    WireBtn L, nm, "row"
    Set NewCell = L
End Function

' Strelica dropdown-a. Nativno dugme ComboBox-a je iskljuceno
' (ShowDropButtonWhen = Never), pa ovaj Label nista ne maskira - samo crta
' strelicu. Zato ide u visini SVOG teksta i centrira se u kutiji; ranije je
' bio visok koliko i polje pa je strelica visila pri vrhu.
Public Function NewChevron(parent As Object, nm As String, X As Single, Y As Single, _
                            w As Single, h As Single) As Object
    Dim L As Object
    Set L = NewLbl(parent, nm, ChrW(9662), X, CenterY(Y, h, TS_META), w, TxtH(TS_META), _
                   TS_META, False, C_MUTED, -1, fmTextAlignCenter)
    L.ZOrder 0
    WireBtn L, nm, "chev"
    Set NewChevron = L
End Function

'=====================================================================
' OSNOVNE FABRIKE
'=====================================================================
Public Function NewLbl(parent As Object, nm As String, cap As String, X As Single, Y As Single, _
                       w As Single, h As Single, fs As Single, bold As Boolean, fc As Long, _
                       Optional bg As Long = -1, Optional align As Long = fmTextAlignLeft, _
                       Optional fnt As String = "") As Object
    Dim L As Object
    Set L = parent.Controls.Add("Forms.Label.1", nm, True)
    With L
        .Left = X: .top = Y: .width = w: .Height = h
        .caption = cap: .ForeColor = fc: .TextAlign = align
        With .Font
            .name = IIf(fnt = "", F_UI, fnt)
            .Size = fs
        End With
        .BorderStyle = fmBorderStyleNone
        If bg = -1 Then
            .BackStyle = fmBackStyleTransparent
        Else
            .BackStyle = fmBackStyleOpaque: .BackColor = bg
        End If
        .WordWrap = False
    End With
    ' Rez ide POSLE pozadine i kroz potvrdu -- vidi PostaviRez. Do v6-ui-174 je
    ' stajao unutar "With .Font" i tiho se gubio, pa je SVAKA runtime kontrola
    ' bila bold. Uniformno, pa se nije ni primetilo -- izaslo je tek kad je jedna
    ' tvrdnja trazila da NEIZABRAN segment NIJE bold. Uz to clsFlatBtn.IsSelected
    ' cita bas taj Font.Bold, pa je za "nav", "chip" i "seg" bio uvek True i
    ' hover nije prefarbavao nista.
    PostaviRez L, bold
    Set NewLbl = L
End Function

' TextBox BEZ svoje ivice - ivicu i padding daje shell oko njega
Public Function NewTxt(parent As Object, nm As String, val As String, X As Single, Y As Single, _
                       w As Single, h As Single, isNum As Boolean) As Object
    Dim t As Object
    Set t = parent.Controls.Add("Forms.TextBox.1", nm, True)
    With t
        .Left = X: .top = Y: .width = w: .Height = h
        .text = val
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleOpaque
        .BackColor = C_WHITE
        .ForeColor = C_FOREST
        With .Font
            .name = IIf(isNum, F_NUM, F_UI)
            .Size = TS_BODY
        End With
        .TextAlign = IIf(isNum, fmTextAlignRight, fmTextAlignLeft)
    End With
    ' isti razlog kao u NewLbl -- vidi PostaviRez
    PostaviRez t, isNum
    WireInput t, nm
    ' upis u spisak brojcanih polja - odavde zna FilterKeyPress koga da cuva
    If isNum Then
        If mNumTag Is Nothing Then Set mNumTag = CreateObject("Scripting.Dictionary")
        Set mNumTag(nm) = t
    End If
    Set NewTxt = t
End Function

' ComboBox BEZ svoje ivice; nativno dugme maskira NewChevron
Public Function NewCombo(parent As Object, nm As String, val As String, X As Single, Y As Single, _
                         w As Single, h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add("Forms.ComboBox.1", nm, True)
    With c
        .Left = X: .top = Y: .width = w: .Height = h
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = C_WHITE
        .ForeColor = C_FOREST
        With .Font
            .name = F_UI
            .Size = TS_BODY
        End With
        .style = fmStyleDropDownCombo
        ' Nativno dopunjavanje se tuklo sa suzavanjem u nasem panelu (PopIndex
        ' trazi PODNIZ, MatchEntry prepisuje tekst na prvi pogodak po pocetku).
        ' Izbor se ionako potvrdjuje iz panela, ne kucanjem.
        .MatchEntry = fmMatchEntryNone
        .ShowDropButtonWhen = fmShowDropButtonWhenNever
        .text = val
    End With
    WireInput c, nm
    Set NewCombo = c
End Function

' linearna interpolacija dve BGR boje - za fejk gradijent u zaglavlju
Public Function Lerp(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    r1 = c1 And &HFF&: g1 = (c1 \ &H100&) And &HFF&: b1 = (c1 \ &H10000) And &HFF&
    r2 = c2 And &HFF&: g2 = (c2 \ &H100&) And &HFF&: b2 = (c2 \ &H10000) And &HFF&
    Lerp = RGB(r1 + (r2 - r1) * t, g1 + (g2 - g1) * t, b1 + (b2 - b1) * t)
End Function

Public Function HasCtl(parent As Object, ByVal nm As String) As Boolean
    On Error Resume Next
    Dim o As Object: Set o = parent.Controls(nm)
    HasCtl = Not o Is Nothing
End Function

Public Function DisplayFont() As String
    On Error Resume Next
    DisplayFont = GetLocalConfigValue("UI_DISPLAY_FONT", F_DISPLAY_FB)
    If Len(DisplayFont) = 0 Then DisplayFont = F_DISPLAY_FB
End Function

' ============================================================
' STIL PANELA U RADNOJ POVRSINI (v6-ui-198)
'
' Paneli su do sada nosili legacy izgled (siva podloga, sistemski font, gusti
' redovi) zato sto je crtani legacy Style* helperima. Od migracije koriste ISTU
' paletu i tipografiju kao ostatak nove ljuske -- modUiKit konstante -- pa se ne
' vidi da su "stariji ekrani".
'
' Stoje OVDE, a ne u modulu panela: dva panela sa svojom kopijom istog
' pravila su dva mesta koja se prvom doradom razidju.
'
' Menja se SAMO prezentacija. Registar polja (ConfigEditorFields, 97 redova),
' citanje, upis (SaveConfigEditor) i ozicavanje (clsConfigBtn) ostaju netaknuti:
' kontrole su i dalje CommandButton/TextBox/ComboBox, jer WithEvents omotac hvata
' bas njih. Zamena tipa kontrole bi pokidala svaki klik u panelu.
' ============================================================
Public Sub PanelStilNaslov(ByVal l As MSForms.label)
    l.BackStyle = fmBackStyleTransparent
    l.BorderStyle = fmBorderStyleNone
    l.ForeColor = C_FOREST
    l.Font.name = F_UI
    l.Font.Size = TS_H1
    l.Font.Bold = True
End Sub

Public Sub PanelStilNapomena(ByVal l As MSForms.label)
    l.BackStyle = fmBackStyleTransparent
    l.BorderStyle = fmBorderStyleNone
    l.ForeColor = C_MUTED
    l.Font.name = F_UI
    l.Font.Size = TS_META
    l.Font.Bold = False
End Sub

' Natpis iznad polja -- verzalom i sitno, isto kao NewFieldG u ljusci.
Public Sub PanelStilNatpis(ByVal l As MSForms.label)
    l.caption = UCase$(l.caption)
    l.BackStyle = fmBackStyleTransparent
    l.BorderStyle = fmBorderStyleNone
    l.ForeColor = C_MUTED
    l.Font.name = F_UI
    l.Font.Size = TS_LABEL
    l.Font.Bold = True
End Sub

Public Sub PanelStilUnos(ByVal t As MSForms.Control)
    ' REDOSLED je bitan: u MSForms SpecialEffect i BorderStyle se iskljucuju --
    ' postavljanje SpecialEffect-a POSLE BorderStyle-a gasi ivicu. Zato prvo
    ' Flat, pa tek onda ivica i njena boja.
    t.SpecialEffect = fmSpecialEffectFlat
    t.BorderStyle = fmBorderStyleSingle
    t.BorderColor = C_INPUT_BORDER
    t.BackColor = C_WHITE
    t.ForeColor = C_FOREST
    t.Font.name = F_UI
    t.Font.Size = TS_BODY
End Sub

' Dugme u paleti ljuske. Vrste su iste kao modUiKit.BtnV, da se dva mesta ne
' razidju u tome sta "primarno" znaci.
Public Sub PanelStilDugme(ByVal b As MSForms.CommandButton, ByVal vrsta As String)
    Select Case vrsta
        Case "primary": b.BackColor = C_GREEN: b.ForeColor = C_WHITE: b.Font.Bold = True
        Case "soft":    b.BackColor = C_SOFT_BG: b.ForeColor = C_GREEN: b.Font.Bold = True
        Case "danger":  b.BackColor = C_RUST: b.ForeColor = C_WHITE: b.Font.Bold = True
        Case Else:      b.BackColor = C_WHITE: b.ForeColor = C_FOREST: b.Font.Bold = False
    End Select
    b.Font.name = F_UI
    b.Font.Size = TS_META
End Sub

' Hover za dugme iz dizajnera. Iste vrednosti koje clsFlatBtn daje istim
' podlogama, pa se runtime dugme i dugme iz .frx ne razidju u ponasanju.
' Panel ga ne koristi (tamo hover nije bio ni ranije); koriste ga forme koje
' su same sebi ekran -- prijava i plutajuca kartica nad Excelom.
Public Sub PanelStilDugmeHover(ByVal b As MSForms.CommandButton, ByVal vrsta As String, _
                               ByVal preko As Boolean)
    PanelStilDugme b, vrsta
    If Not preko Then Exit Sub
    Select Case b.BackColor
        Case C_GREEN: b.BackColor = C_BTN_HOVER
        Case C_WHITE: b.BackColor = RGB(246, 244, 238)
        Case Else:    b.BackColor = C_SAND
    End Select
End Sub

' Segment prekidaca grupa -- isti jezik kao prekidac lista na ekranima:
' izabran je pun (forest/cream), ostali su beli sa ivicom.
'
' Zamenio je pescanu traku sa strelicom: sklapajuce grupe su bile OBRAZAC STARE
' forme. Nova ljuska svuda pokazuje JEDNU listu i menja je prekidacem, pa i
' podesavanja sada rade tako.
Public Sub PanelStilSegment(ByVal b As MSForms.CommandButton, ByVal cap As String, _
                            ByVal izabran As Boolean)
    b.caption = cap
    If izabran Then
        b.BackColor = C_FOREST
        b.ForeColor = C_CREAM
    Else
        b.BackColor = C_WHITE
        b.ForeColor = C_MUTED
    End If
    b.Font.name = F_UI
    b.Font.Size = TS_META
    b.Font.Bold = izabran
End Sub

' Sirina segmenta po duzini natpisa. Racuna se, ne pogadja: grupe se zovu od
' "SEF" do "Otkup / dokumenta", pa fiksna sirina ili secka ili rasipa.
Public Function PanelSirinaSegmenta(ByVal cap As String) As Single
    PanelSirinaSegmenta = 22 + Len(cap) * 5.4
    If PanelSirinaSegmenta < 64 Then PanelSirinaSegmenta = 64
End Function

' Slika iz koda (modLogo). BackColor je vazan: PictureSizeMode = Zoom cuva odnos
' stranica, pa okvir koji mu ne odgovara ostavlja pojas oko slike -- a pozadina
' je u samu sliku vec pecena, pa taj pojas mora da bude iste boje.
'
' BackStyle je OPAQUE namerno: transparentan Image bi kroz taj pojas pustio ono
' sto je ispod (gradijent splash-a), pa bi se oko znaka video pravougaonik.
Public Function NewImg(parent As Object, nm As String, X As Single, Y As Single, _
                       w As Single, h As Single, ByVal bg As Long) As Object
    Dim im As Object
    Set im = parent.Controls.Add("Forms.Image.1", nm, True)
    With im
        .Left = X: .top = Y: .width = w: .Height = h
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleOpaque
        .BackColor = bg
        .PictureSizeMode = fmPictureSizeModeZoom
        .PictureAlignment = fmPictureAlignmentCenter
        .Visible = False
    End With
    Set NewImg = im
End Function

Public Function NewZone(parent As Object, nm As String, X As Single, Y As Single, _
                         w As Single, h As Single, bg As Long) As Object
    Dim f As Object
    Set f = NewFrame(parent, nm, X, Y, w, h, bg)
    f.Visible = False
    Set NewZone = f
End Function

Public Function NewFrame(parent As Object, nm As String, X As Single, Y As Single, _
                         w As Single, h As Single, bg As Long) As Object
    Dim f As Object
    Set f = parent.Controls.Add("Forms.Frame.1", nm, True)
    With f
        .Left = X: .top = Y: .width = w: .Height = h
        .caption = "": .BackColor = bg
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
    Set NewFrame = f
End Function
