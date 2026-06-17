# frmPalete — build guide (#44)

Kontrole MSForms forme žive u binarnom `.frx`, pa se forma ne može isporučiti kao
tekst. Zato: ti napraviš formu i kontrole u VBA designeru (imena su važna —
code-behind se vezuje po imenu), pa nalepiš kod ispod. Sva poslovna logika je u
`modPaletniList` (read-modeli + TX wrapperi); forma samo zove te funkcije.

## 1. Napravi formu

`Insert → UserForm`. U Properties:
- **(Name)** = `frmPalete`
- **Caption** = `Palete — pregled i obrada`

## 2. Dodaj kontrole (prevuci iz Toolbox-a, postavi (Name) tačno)

Pozicije rasporedi po želji; bitni su **(Name)** i tip. Posebne osobine su u koloni „Napomena".

### Filteri (gore)
| (Name) | Tip | Napomena |
|---|---|---|
| `txtFilterGod` | TextBox | godina; prazno = sve |
| `cmbFilterVrsta` | ComboBox | vrsta voća; prazno = sve (slobodan unos) |
| `cmbFilterStatus` | ComboBox | puni se u kodu (Sve/Otvorena/Zatvorena) |
| `cmbFilterPre` | ComboBox | puni se u kodu (Sve/Ne/Da) |
| `btnOsvezi` | CommandButton | Caption `Osveži` |

(Po želji dodaj Label-e „Godina/Vrsta/Status/Prerađeno" — nisu vezani u kodu.)

### Leva lista — palete
| (Name) | Tip | Napomena |
|---|---|---|
| `lstPalete` | ListBox | **MultiSelect = 1 - fmMultiSelectMulti** (za „Preradi izabrane"). ColumnCount/širine se postavljaju u kodu |

Kolone (redom): PaletaID(skriveno), Broj, Godina, Vrsta, Sorta, Klasa, TipAmb,
Gajbice, Kapacitet, Neto, Bruto, Status, Prerađeno.

### Desna lista — stavke izabrane palete
| (Name) | Tip | Napomena |
|---|---|---|
| `lstStavke` | ListBox | ColumnCount se postavlja u kodu |

Kolone: PrijemnicaID, BrojPrijemnice, BrojZbirne, Gajbice, NetoKg.

### Polja za preradu (izlaz)
| (Name) | Tip |
|---|---|
| `txtKutije` | TextBox |
| `txtKese` | TextBox |
| `txtNeto` | TextBox |
| `txtNapomena` | TextBox |

### Dugmad (akcije)
| (Name) | Tip | Caption |
|---|---|---|
| `btnStampaj` | CommandButton | Štampaj paletu |
| `btnPDF` | CommandButton | PDF palete |
| `btnStampajNepotpune` | CommandButton | Štampaj nepotpune |
| `btnZatvori` | CommandButton | Zatvori ručno |
| `btnPreradi` | CommandButton | Preradi izabrane |
| `btnPovratak` | CommandButton | Povratak |

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

    Me.lstPalete.ColumnCount = 13
    Me.lstPalete.ColumnWidths = "0;35;30;55;55;30;45;45;55;55;60;55;55"
    Me.lstStavke.ColumnCount = 5
    Me.lstStavke.ColumnWidths = "75;60;60;45;55"

    RefreshGrid
    Exit Sub
EH:
    MsgBox "Greska pri otvaranju: " & Err.Description, vbCritical, APP_NAME
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

Private Sub lstPalete_Click()
    On Error Resume Next
    Dim pid As String: pid = FirstSelectedPaletaID()
    If pid = "" Then
        Me.lstStavke.Clear
        Exit Sub
    End If
    Dim s As Variant: s = GetPaletaStavkeForGrid(pid)
    If IsEmpty(s) Then
        Me.lstStavke.Clear
    Else
        Me.lstStavke.List = s
    End If
End Sub

Private Function FirstSelectedPaletaID() As String
    Dim i As Long
    For i = 0 To Me.lstPalete.ListCount - 1
        If Me.lstPalete.Selected(i) Then
            FirstSelectedPaletaID = CStr(Me.lstPalete.List(i, 0))
            Exit Function
        End If
    Next i
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
    Dim pid As String: pid = FirstSelectedPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    PrintPaletniList pid
End Sub

Private Sub btnPDF_Click()
    Dim pid As String: pid = FirstSelectedPaletaID()
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
    Dim pid As String: pid = FirstSelectedPaletaID()
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

## Napomene
- Širine kolona (`ColumnWidths`, u tačkama) su okvirne — doteraj po ekranu. Prva je
  `0` da sakrije `PaletaID` (forma ga čita interno za akcije).
- MSForms ListBox sa `.List` ne prikazuje zaglavlja kolona; po želji stavi Label-red
  iznad liste sa nazivima.
- Forma ne piše direktno u tabele — sve ide preko `*_TX` / read-model funkcija u
  `modPaletniList`.
