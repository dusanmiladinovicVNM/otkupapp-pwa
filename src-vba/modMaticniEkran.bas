Attribute VB_Name = "modMaticniEkran"
'=====================================================================
' modMaticniEkran - zajednicko telo tri maticna ekrana (korak M1).
'
' Partneri, Proizvodi i cene, Ambalaza i pakovanje se razlikuju SAMO po tome
' koje sekcije nose. Zona, dogadjaji i citanje redova su im isti, pa stoje
' ovde: tri kopije istog rasporeda bi se razisle prvom doradom.
'
' ZASTO OVO NIJE U modMaticniIzvor: taj modul opisuje PODATKE i ne sme da zna
' za kontrole ni za ljusku. Ovaj zna za zonu i za mrezu, a nista o tabelama.
'
' ZONA (M1) je pregled, ne editor: naziv sekcije, koliko zapisa ukupno, koliko
' aktivnih i koliko neaktivnih. Polja unosa i dugmad dolaze u M2 -- zona tada
' raste, po obrascu liste "Nova prerada" sa ekrana Palete.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATEKR_BUILD As String = "v6-ui-188"

' Visina zone je ista kao KPI traka, pa naslov ispod nje pada u isti red na
' svim ekranima -- isto pravilo koje vec postuju Palete i Oporavak.
Private Const MAT_ZONA_H As Single = KPI_H

'--------------------------------------------------------------- ZONA
Public Sub ZonaGradi(ByVal z As Object)
    Dim i As Long
    modUiKit.NewLbl z, "matCap", "", PAD, 6, 300, 11, TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "matBroj", ChrW(8212), PAD, 18, 300, 20, TS_KPI, True, C_FOREST, -1
    ' Napomena da se unos jos radi kroz staru formu. Stoji u zoni, ne u dijalogu:
    ' dijalog se zatvori i zaboravi, a operater koji ovde trazi dugme "Dodaj"
    ' mora da vidi gde ono jeste.
    modUiKit.NewLbl z, "matHint", Poruka("OTKUI_MAT_UNOS_LEGACY"), PAD, 40, 520, 13, _
                    TS_META, False, C_MUTED, -1

    For i = 0 To 1
        modUiKit.NewLbl z, "matKL" & i, "", 0, 6, 120, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "matKV" & i, ChrW(8212), 0, 18, 120, 20, TS_KPI, True, _
                        C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i
    modUiKit.NewLbl z, "matLnB", "", 0, MAT_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function ZonaRaspored(ByVal z As Object, ByVal w As Single) As Single
    Dim i As Long
    On Error Resume Next
    For i = 0 To 1
        z.Controls("matKL" & i).Left = w - PAD - (2 - i) * 150
        z.Controls("matKV" & i).Left = w - PAD - (2 - i) * 150
    Next i
    ' Napomena deli red sa dve plocice desno; na uskom prozoru se skloni umesto
    ' da dobije negativnu sirinu.
    z.Controls("matHint").Visible = (w - 2 * PAD - 320 > 120)
    If w - 2 * PAD - 320 > 120 Then z.Controls("matHint").width = w - 2 * PAD - 320
    z.Controls("matLnB").width = w
    ZonaRaspored = MAT_ZONA_H
End Function

' Brojke se pisu POSLE citanja redova, iz istog prolaza -- pa se broj u zoni i
' lista u mrezi ne mogu razici. Sekcija bez kolone statusa nema sta da razlozi
' na aktivne i neaktivne, pa te dve plocice ostaju prazne (em-crta), a ne nula:
' nula bi tvrdila da neaktivnih nema, a odgovor je da pojam ne postoji.
Private Sub OsveziZonu(ByVal ekran As String, ByVal lista As String)
    Dim z As Object, imaStatus As Boolean
    On Error Resume Next
    Set z = modOtkupUI.ScreenZone(ekran)
    If z Is Nothing Then Exit Sub
    imaStatus = (Len(modMaticniIzvor.MatStatusKolona(lista)) > 0)

    z.Controls("matCap").caption = UCase$(Poruka(NaslovListe(ekran, lista)))
    z.Controls("matBroj").caption = CStr(modMaticniIzvor.MatUkupno()) & " " & _
                                    Poruka("OTKUI_MAT_ZAPISA")
    z.Controls("matKL0").caption = UCase$(Poruka("OTKUI_MAT_AKTIVNIH"))
    z.Controls("matKL1").caption = UCase$(Poruka("OTKUI_MAT_NEAKTIVNIH"))
    If imaStatus Then
        z.Controls("matKV0").caption = CStr(modMaticniIzvor.MatAktivnih())
        z.Controls("matKV1").caption = CStr(modMaticniIzvor.MatNeaktivnih())
    Else
        z.Controls("matKV0").caption = ChrW(8212)
        z.Controls("matKV1").caption = ChrW(8212)
    End If
End Sub

' Naslov aktivne liste iz ISTOG spiska koji puni prekidac -- da se natpis u
' zoni i natpis na dugmetu ne mogu razici.
Private Function NaslovListe(ByVal ekran As String, ByVal lista As String) As String
    Dim a As Variant, r As Variant, p() As String
    a = modMaticniIzvor.MatSekcijeEkrana(ekran)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        p = Split(CStr(r), "|")
        If p(0) = lista Then
            NaslovListe = p(2)
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- REDOVI
Public Function Redovi(ByVal ekran As String, ByVal lista As String, _
                       ByVal filter As String, ByVal q As String) As Variant
    Redovi = modMaticniIzvor.MatRedovi(lista, filter, q)
    ' Zona se osvezava POSLE citanja, iz istog prolaza -- brojke i lista se tako
    ' ne mogu razici. Ekran se prosledjuje jer tri ekrana dele ovo telo, pa se
    ' mora znati CIJU zonu ljuska treba da vrati.
    OsveziZonu ekran, lista
End Function

'------------------------------------------------------------ DOGADJAJI
' M1 poznaje SAMO prekidac lista. Izbor reda vraca False (mreza se ne cita
' ponovo, operater ne gubi mesto u listi), a radnji nad redom jos nema -- one
' menjaju podatke i dolaze u M2, zajedno sa jednim piscem.
'
' Lista se prima ByRef zato sto je stanje EKRANA, ne ovog modula: tri ekrana
' dele telo, ali svaki pamti svoju aktivnu listu.
Public Function Dogadjaj(ByVal tag As String, ByRef lista As String) As Boolean
    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = lista Then Exit Function
        lista = Mid$(tag, 3)
        Dogadjaj = True
    End If
End Function
