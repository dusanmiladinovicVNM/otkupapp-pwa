# frmPalete — build guide (#44)

Kontrole MSForms forme žive u binarnom `.frx`, pa se forma ne može isporučiti kao
tekst. Zato: ti napraviš formu i kontrole u VBA designeru (imena su važna —
code-behind se vezuje po imenu), pa nalepiš kod ispod. Sva poslovna logika je u
`modPaletniList` (read-modeli + TX wrapperi); forma samo zove te funkcije.

## 1. Napravi formu

`Insert → UserForm`. U Properties:
- **(Name)** = `frmPalete`
- **Caption** = `Palete — pregled i obrada`

## 2. Kontrole — tačne pozicije

Vrednosti su u **tačkama (pt)** — iste jedinice koje Properties prozor pokazuje
(`Left`, `Top`, `Width`, `Height`). Za svaku kontrolu postavi **(Name)** + te
četiri vrednosti (+ Caption gde stoji). Postupak: prevuci kontrolu iz Toolbox-a,
pa u Properties upiši vrednosti iz reda.

Forma `frmPalete`: **Width = 730, Height = 452**.

| (Name) | Tip | Left | Top | Width | Height | Caption / osobina |
|---|---|---|---|---|---|---|
| `lblFilterGod` | Label | 8 | 11 | 42 | 14 | `Godina:` |
| `txtFilterGod` | TextBox | 52 | 8 | 46 | 18 | |
| `lblFilterVrsta` | Label | 108 | 11 | 36 | 14 | `Vrsta:` |
| `cmbFilterVrsta` | ComboBox | 146 | 8 | 96 | 18 | |
| `lblFilterStatus` | Label | 250 | 11 | 40 | 14 | `Status:` |
| `cmbFilterStatus` | ComboBox | 292 | 8 | 88 | 18 | |
| `lblFilterPre` | Label | 388 | 11 | 58 | 14 | `Prerađeno:` |
| `cmbFilterPre` | ComboBox | 448 | 8 | 68 | 18 | |
| `btnOsvezi` | CommandButton | 526 | 7 | 70 | 20 | `Osveži` |
| `lblPalete` | Label | 8 | 36 | 300 | 12 | `Palete (Ctrl/Shift za više)` |
| `lstPalete` | ListBox | 8 | 50 | 520 | 320 | **MultiSelect = `1 - fmMultiSelectMulti`** |
| `lblStavke` | Label | 536 | 36 | 186 | 12 | `Stavke izabrane palete` |
| `lstStavke` | ListBox | 536 | 50 | 186 | 150 | |
| `lblKutije` | Label | 536 | 214 | 56 | 14 | `Kutije:` |
| `txtKutije` | TextBox | 596 | 212 | 50 | 18 | |
| `lblKese` | Label | 536 | 238 | 56 | 14 | `Kese:` |
| `txtKese` | TextBox | 596 | 236 | 50 | 18 | |
| `lblNeto` | Label | 536 | 262 | 56 | 14 | `Neto izlaz:` |
| `txtNeto` | TextBox | 596 | 260 | 50 | 18 | |
| `lblNapomena` | Label | 536 | 286 | 56 | 14 | `Napomena:` |
| `txtNapomena` | TextBox | 596 | 284 | 126 | 18 | |
| `btnStampaj` | CommandButton | 8 | 384 | 92 | 24 | `Štampaj paletu` |
| `btnPDF` | CommandButton | 106 | 384 | 92 | 24 | `PDF palete` |
| `btnStampajNepotpune` | CommandButton | 204 | 384 | 110 | 24 | `Štampaj nepotpune` |
| `btnZatvori` | CommandButton | 320 | 384 | 92 | 24 | `Zatvori ručno` |
| `btnPreradi` | CommandButton | 418 | 384 | 110 | 24 | `Preradi izabrane` |
| `btnPovratak` | CommandButton | 630 | 384 | 92 | 24 | `Povratak` |

Raspored: filteri u vrhu (red ~8pt); ispod levo veliki `lstPalete` (520×320);
desno `lstStavke` pa polja za preradu (Kutije/Kese/Neto/Napomena); dole red
dugmadi.

Kolone (redom) u `lstPalete`: PaletaID(skriveno), Broj, Godina, Vrsta, Sorta,
Klasa, TipAmb, Gajbice, Kapacitet, Neto, Bruto, Status, Prerađeno.
Kolone u `lstStavke`: PrijemnicaID, BrojPrijemnice, BrojZbirne, Gajbice, NetoKg.

> `ColumnCount` i `ColumnWidths` za obe liste postavlja kod (`UserForm_Initialize`)
> — ne diraš ručno. Jedino `lstPalete.MultiSelect` postavi u Properties.

## 3. Nalepi code-behind

Desni klik na `frmPalete` → `View Code` → obriši sve → nalepi:

```vba
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Me.cmbFilterStatus.Clear
    Me.cmbFilterStatus.AddItem ""            ' Sve
    Me.cmbFilterStatus.AddItem "Otvorena"
    Me.cmbFilterStatus.AddItem "Zatvorena"

    Me.cmbFilterPre.Clear
    Me.cmbFilterPre.AddItem ""               ' Sve
    Me.cmbFilterPre.AddItem "Ne"
    Me.cmbFilterPre.AddItem "Da"

    Me.txtFilterGod.value = Year(Date)

    ' Vrsta voca iz sifarnika (isti izvor kao na prijemnici); "" na vrhu = Sve
    FillCmb Me.cmbFilterVrsta, GetLookupList(TBL_KULTURE, "VrstaVoca")
    Me.cmbFilterVrsta.AddItem "", 0

    Me.lstPalete.ColumnCount = 13
    Me.lstPalete.ColumnWidths = "0;30;32;50;50;30;40;40;48;48;50;50;50"
    Me.lstStavke.ColumnCount = 5
    Me.lstStavke.ColumnWidths = "72;38;54;32;36"   ' za lstStavke Width 250

    RefreshGrid
    Exit Sub
EH:
    MsgBox "Greska pri otvaranju: " & Err.Description, vbCritical, APP_NAME
End Sub

' Tema (kao ostale forme): krem pozadina + stilizovane kontrole/dugmad.
Private Sub UserForm_Activate()
    On Error Resume Next
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    StylePrimaryButton btnPreradi, "Preradi izabrane"
    StyleExitButton btnPovratak, "Povratak"
End Sub

Private Sub RefreshGrid()
    On Error GoTo EH
    Dim god As Long
    If IsNumeric(Me.txtFilterGod.value) Then god = CLng(Me.txtFilterGod.value)

    Dim data As Variant
    data = GetPaleteForGrid(god, Trim$(Me.cmbFilterVrsta.value), _
                            Trim$(Me.cmbFilterStatus.value), Trim$(Me.cmbFilterPre.value))

    Me.lstStavke.Clear
    If IsEmpty(data) Then
        Me.lstPalete.Clear
    Else
        Me.lstPalete.List = data
    End If
    Exit Sub
EH:
    MsgBox "Greska pri osvezavanju: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnOsvezi_Click()
    RefreshGrid
End Sub

' MultiSelect ListBox NE okida Click pouzdano -> Change + ListIndex (red sa
' fokusom = poslednji kliknut) za prikaz stavki desno.
Private Sub lstPalete_Change()
    Dim i As Long: i = Me.lstPalete.ListIndex
    If i < 0 Then
        Me.lstStavke.Clear
        Exit Sub
    End If
    Dim s As Variant: s = GetPaletaStavkeForGrid(CStr(Me.lstPalete.List(i, 0)))
    If IsEmpty(s) Then
        Me.lstStavke.Clear
    Else
        Me.lstStavke.List = s
    End If
End Sub

' Red sa fokusom (akcije nad jednom paletom).
Private Function CurrentPaletaID() As String
    Dim i As Long: i = Me.lstPalete.ListIndex
    If i >= 0 Then CurrentPaletaID = CStr(Me.lstPalete.List(i, 0))
End Function

Private Function SelectedPaletaIDs() As Collection
    Dim c As Collection: Set c = New Collection
    Dim i As Long
    For i = 0 To Me.lstPalete.ListCount - 1
        If Me.lstPalete.Selected(i) Then c.Add CStr(Me.lstPalete.List(i, 0))
    Next i
    Set SelectedPaletaIDs = c
End Function

Private Sub btnStampaj_Click()
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    PrintPaletniList pid
End Sub

Private Sub btnPDF_Click()
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    ExportPaletniListPDF pid, True
End Sub

Private Sub btnStampajNepotpune_Click()
    On Error GoTo EH
    Dim n As Long: n = PrintNepotpunePalete()
    MsgBox n & " nepotpunih paleta poslato na izlaz (po PALETA_PRINT_MODE).", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnZatvori_Click()
    On Error GoTo EH
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    ClosePaletaManual_TX pid
    RefreshGrid
    MsgBox "Paleta je zatvorena.", vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Paleta nije zatvorena: " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub btnPreradi_Click()
    On Error GoTo EH
    Dim ids As Collection: Set ids = SelectedPaletaIDs()
    If ids.count = 0 Then
        MsgBox "Izaberite bar jednu paletu (Ctrl/Shift za vise).", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim preID As String
    preID = SavePrerada_TX(ids, _
                CLng(Val(Me.txtKutije.value)), _
                CLng(Val(Me.txtKese.value)), _
                CDbl(Val(Replace(Me.txtNeto.value, ",", "."))), _
                Trim$(Me.txtNapomena.value))

    If preID <> "" Then ExportPreradaPDF preID, True

    Me.txtKutije.value = ""
    Me.txtKese.value = ""
    Me.txtNeto.value = ""
    Me.txtNapomena.value = ""
    RefreshGrid
    MsgBox "Prerada je sacuvana.", vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Prerada nije sacuvana: " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub btnPovratak_Click()
    Unload Me
End Sub
```

## 4. Pokretanje

Posle pravljenja forme, dodaj launcher u `modPaletniListUI` (ne ranije — referenca
na `frmPalete` ne kompajlira dok forma ne postoji):

```vba
Public Sub ShowPalete()
    frmPalete.Show
End Sub
```

Pa `Alt+F8 → ShowPalete` (ili dugme na traci). Za probu odmah: `frmPalete.Show` u
Immediate prozoru.

## 5. (opciono) Zaglavlja kolona

ListBox punjen preko `.List` ne prikazuje zaglavlja (ColumnHeads radi samo uz
RowSource). Najčistije: zaključana 1-red „header" lista sa ISTIM `ColumnWidths`
— savršeno poravnanje, bez skrol-rasinhronizacije (518 < 520).

**Pomeranje (da naslov ostane vidljiv).** Zaglavlje ide IZMEĐU naslova i liste:
naslovi ostaju na Top 36; zaglavlja na Top 48; liste se spuštaju na Top 62.

| Kontrola | Left | Top | Width | Height | Napomena |
|---|---|---|---|---|---|
| `lblPalete` | 8 | 36 | 300 | 12 | (ostaje) |
| `lstPaleteHdr` | 8 | 48 | 520 | 14 | novo; `Locked = True` |
| `lstPalete` | 8 | 62 | 520 | 308 | bilo Top 50 / H 320 |
| `lblStavke` | 536 | 36 | 186 | 12 | (ostaje) |
| `lstStavkeHdr` | 536 | 48 | 186 | 14 | novo; `Locked = True` |
| `lstStavke` | 536 | 62 | 186 | 138 | bilo Top 50 / H 150 |

**U `UserForm_Initialize`** (posle `lstPalete`/`lstStavke` kolona):

```vba
    ' --- zaglavlje paleta ---
    Me.lstPaleteHdr.ColumnCount = 13
    Me.lstPaleteHdr.ColumnWidths = "0;30;32;50;50;30;40;40;48;48;50;50;50"
    Dim hdr(0 To 0, 0 To 12) As Variant
    hdr(0, 1) = "Broj":   hdr(0, 2) = "God":     hdr(0, 3) = "Vrsta"
    hdr(0, 4) = "Sorta":  hdr(0, 5) = "Klasa":   hdr(0, 6) = "TipAmb"
    hdr(0, 7) = "Gajb":   hdr(0, 8) = "Kap":     hdr(0, 9) = "Neto"
    hdr(0, 10) = "Bruto": hdr(0, 11) = "Status": hdr(0, 12) = "Prer."
    Me.lstPaleteHdr.List = hdr
    Me.lstPaleteHdr.Locked = True

    ' --- zaglavlje stavki ---
    Me.lstStavkeHdr.ColumnCount = 5
    Me.lstStavkeHdr.ColumnWidths = "72;38;54;32;36"
    Dim hdrS(0 To 0, 0 To 4) As Variant
    hdrS(0, 0) = "PrijemID": hdrS(0, 1) = "BrPrij": hdrS(0, 2) = "Zbirna"
    hdrS(0, 3) = "Gajb":     hdrS(0, 4) = "Neto"
    Me.lstStavkeHdr.List = hdrS
    Me.lstStavkeHdr.Locked = True
```

> Header `ColumnWidths` su LITERALI (isti string kao glavne liste), namerno — da
> ne zavise od redosleda postavljanja. Ako menjaš širine, promeni na OBA mesta
> (glavna lista + njeno zaglavlje).

**U `UserForm_Activate`** (posle `ApplyThemeToControls`, da tema ne pregazi izgled):

```vba
    Me.lstPaleteHdr.Font.Bold = True
    Me.lstPaleteHdr.BackColor = BG_TOP()
    Me.lstStavkeHdr.Font.Bold = True
    Me.lstStavkeHdr.BackColor = BG_TOP()
```

**Alternativa — Label-i** iznad liste (približno poravnanje; `lstPalete` Left-base
≈10, Top 38): Broj 10 · God 40 · Vrsta 72 · Sorta 122 · Klasa 172 · TipAmb 202 ·
Gajb 242 · Kap 282 · Neto 330 · Bruto 378 · Status 428 · Prer 478. Header-lista
je urednija i tačnija.

## 6. Uvezivanje u glavni meni (frmOtkupAPP)

`frmOtkupAPP` je shell: sidebar nav dugmad (raspoređena u `SetupButtons`, korak
`topPos = topPos + 42`) + content-area; sekcije se otvaraju preko
`OpenContentForm frmX, btnX, "Naslov"` (embeduje i RAZVlači formu).

**Koraci (zajednički):**
1. U designeru `frmOtkupAPP` dodaj `CommandButton` u sidebar, **(Name) = `btnPalete`**
   (pozicija/stil nebitni — `SetupButtons` ih postavlja; možeš kopirati `btnTrace`).
2. U `SetupButtons`, na željeno mesto (npr. odmah posle „Otkup i prodaja"):

```vba
    StyleNavButton btnPalete, "Palete", topPos
    topPos = topPos + 42
```

   (Sidebar dobija 1 dugme više; ako zafali visine, smanji korak ili povećaj `fraSidebar`.)

**Otvaranje (ugrađena sekcija — kao ostale).** Handler u `frmOtkupAPP`:

```vba
Private Sub btnPalete_Click()
    OpenContentForm frmPalete, btnPalete, "Palete"
End Sub
```

Shell razvuče formu (`FitActiveContent` postavlja Width/Height na content-area),
pa frmPalete mora da se „prelije". Dodaj u frmPalete tri stvari:

**(a)** na vrh modula (ispod `Option Explicit`):

```vba
Private mChromeRemoved As Boolean
```

**(b)** u `UserForm_Activate` (kao ostale sekcije — ukloni naslovnu traku, sakrij Povratak):

```vba
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    Me.btnPovratak.Visible = False
```

**(c)** `UserForm_Resize` — pin dugmad dole, desni panel desno, liste popune ostatak:

```vba
Private Sub UserForm_Resize()
    On Error Resume Next
    Const PAD As Double = 8
    Const GAP As Double = 10
    Const PANELW As Double = 250      ' desni panel (stavke + prerada)
    Const TOPGRID As Double = 62
    Const BTNH As Double = 24

    Dim w As Double: w = Me.InsideWidth
    Dim h As Double: h = Me.InsideHeight
    If w < 420 Or h < 220 Then Exit Sub

    Dim panelX As Double: panelX = w - PAD - PANELW
    Dim btnTop As Double:  btnTop = h - PAD - BTNH
    Dim gridW As Double:   gridW = panelX - GAP - PAD

    ' leva strana: naslov / zaglavlje / lista paleta
    Me.lblPalete.Top = 36:      Me.lblPalete.Left = PAD
    Me.lstPaleteHdr.Top = 48:   Me.lstPaleteHdr.Left = PAD:  Me.lstPaleteHdr.width = gridW
    Me.lstPalete.Top = TOPGRID: Me.lstPalete.Left = PAD:     Me.lstPalete.width = gridW
    Me.lstPalete.Height = btnTop - GAP - TOPGRID

    ' desni panel: stavke
    Me.lblStavke.Top = 36:      Me.lblStavke.Left = panelX
    Me.lstStavkeHdr.Top = 48:   Me.lstStavkeHdr.Left = panelX: Me.lstStavkeHdr.width = PANELW
    Me.lstStavke.Top = TOPGRID: Me.lstStavke.Left = panelX:    Me.lstStavke.width = PANELW

    ' desni panel: polja prerade ispod stavki
    Dim fy As Double: fy = TOPGRID + Me.lstStavke.Height + GAP + 6
    Me.lblKutije.Top = fy + 2:    Me.lblKutije.Left = panelX
    Me.txtKutije.Top = fy:        Me.txtKutije.Left = panelX + 70
    Me.lblKese.Top = fy + 26:     Me.lblKese.Left = panelX
    Me.txtKese.Top = fy + 24:     Me.txtKese.Left = panelX + 70
    Me.lblNeto.Top = fy + 50:     Me.lblNeto.Left = panelX
    Me.txtNeto.Top = fy + 48:     Me.txtNeto.Left = panelX + 70
    Me.lblNapomena.Top = fy + 74: Me.lblNapomena.Left = panelX
    Me.txtNapomena.Top = fy + 72: Me.txtNapomena.Left = panelX + 70: Me.txtNapomena.width = PANELW - 70

    ' dugmad dole (btnPovratak je sakriven)
    Me.btnStampaj.Top = btnTop
    Me.btnPDF.Top = btnTop
    Me.btnStampajNepotpune.Top = btnTop
    Me.btnZatvori.Top = btnTop
    Me.btnPreradi.Top = btnTop
End Sub
```

## Napomene
- Širine kolona (`ColumnWidths`, u tačkama) su okvirne — doteraj po ekranu. Prva je
  `0` da sakrije `PaletaID` (forma ga čita interno za akcije).
- MSForms ListBox sa `.List` ne prikazuje zaglavlja kolona; po želji stavi Label-red
  iznad liste sa nazivima.
- Forma ne piše direktno u tabele — sve ide preko `*_TX` / read-model funkcija u
  `modPaletniList`.
