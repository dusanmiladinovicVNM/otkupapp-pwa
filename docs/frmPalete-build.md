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
    Me.lstStavke.ColumnWidths = "60;34;34;28;30"

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
    Me.lstPaleteHdr.ColumnWidths = Me.lstPalete.ColumnWidths
    Dim hdr(0 To 0, 0 To 12) As Variant
    hdr(0, 1) = "Broj":   hdr(0, 2) = "God":     hdr(0, 3) = "Vrsta"
    hdr(0, 4) = "Sorta":  hdr(0, 5) = "Klasa":   hdr(0, 6) = "TipAmb"
    hdr(0, 7) = "Gajb":   hdr(0, 8) = "Kap":     hdr(0, 9) = "Neto"
    hdr(0, 10) = "Bruto": hdr(0, 11) = "Status": hdr(0, 12) = "Prer."
    Me.lstPaleteHdr.List = hdr
    Me.lstPaleteHdr.Locked = True

    ' --- zaglavlje stavki ---
    Me.lstStavkeHdr.ColumnCount = 5
    Me.lstStavkeHdr.ColumnWidths = Me.lstStavke.ColumnWidths
    Dim hdrS(0 To 0, 0 To 4) As Variant
    hdrS(0, 0) = "PrijemID": hdrS(0, 1) = "BrPrij": hdrS(0, 2) = "Zbirna"
    hdrS(0, 3) = "Gajb":     hdrS(0, 4) = "Neto"
    Me.lstStavkeHdr.List = hdrS
    Me.lstStavkeHdr.Locked = True
```

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

## Napomene
- Širine kolona (`ColumnWidths`, u tačkama) su okvirne — doteraj po ekranu. Prva je
  `0` da sakrije `PaletaID` (forma ga čita interno za akcije).
- MSForms ListBox sa `.List` ne prikazuje zaglavlja kolona; po želji stavi Label-red
  iznad liste sa nazivima.
- Forma ne piše direktno u tabele — sve ide preko `*_TX` / read-model funkcija u
  `modPaletniList`.
