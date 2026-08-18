Attribute VB_Name = "modScrOporavak"
'=====================================================================
' modScrOporavak - ekran "Oporavak" (Faza D, stavka 14).
'
' Ljuska ga ne poznaje po imenu: dobija ga preko Application.Run, da
' klijent kome ovaj modul nedostaje i dalje radi (zamka #19).
'
' ODAKLE DOLAZI: cetiri legacy panela iz frmDokumenta, koji su svi radili
' istu stvar - pokazivali sta je ostalo nedovrseno i nudili da se preveze:
'   "Nedovrseno"          m_btnNedovrseno_Click   -> GetNedovrseno
'   "Osiroceni dokumenti" m_btnRecovery_Click     -> GetOsirocenePrijemnice,
'                                                    GetPrijemniceSaOsirocenimPaletama
'   "Undo operacija"      ShowUndoOpsPanel        -> GetUndoableStornoOperations
'   upozorenje pri otvaranju forme (CheckVerwaisteDokumente)
'
' Ovde su to SEST LISTA istog prekidaca koji F1 i Palete vec koriste:
'
'   NEDOVRSENO  jedan pregled svega sto ceka (read-only)
'   PRIJEMNICE  osirotele prijemnice - zbirna im je stornirana ili je nema;
'               radnja "Prevezi" ide na AKTIVNU ZBIRNU
'   ZBIRNE      aktivne zbirne; klik na red bira CILJ za gornju radnju
'   PALETE      prijemnice sa osirotelim paletnim stavkama; radnja
'               "Prevezi" ide na AKTIVNU PRIJEMNICU
'   CILJPRIJ    aktivne prijemnice; klik na red bira CILJ za gornju radnju
'   UNDO        storno operacije koje se mogu vratiti; radnja "Vrati storno"
'
' ZASTO DVA CILJA, A NE DIJALOG SA LISTOM: prevezivanje uvek ima izvor i
' cilj, a novi UI za izbor cilja vec ima obrazac - aktivna otpremnica u F1
' i aktivna paleta na ekranu Palete. Isti obrazac: cilj se bira klikom na
' red u svojoj listi i stoji u zoni gore, gde se vidi sve vreme. Legacy je
' za to imao combo u panelu; ovde je lista, pa se cilj moze i pretraziti i
' sortirati.
'
' NISTA se ovde ne racuna niti upisuje. Podaci dolaze iz
' modStornoRecovery / modStornoZurnal / modDokumenta, a radnje iz
' modDokumenta (ReassignPrijemnicaToZbirna_TX, ReassignPaleteToPrijemnica_TX)
' i modStornoZurnal (UndoOperation_TX) - iste rutine koje zove i legacy.
'
' RAZLIKA U ODNOSU NA LEGACY, namerna: upozorenje na siroticice se vise NE
' pokazuje kao modalni dijalog pri otvaranju forme (CheckVerwaisteDokumente).
' Umesto toga stalno stoji lista "Nedovrseno" i brojka u zoni gore. Dijalog
' pri otvaranju se zatvarao i zaboravljao; lista ne moze da se zaboravi.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCROPO_BUILD As String = "v6-ui-144"

' Kolona liste Nedovrseno koja nosi CorrectionID. JEDAN broj, deljen izmedju
' opisa kolona, punjenja redova i radnje -- da se indeks ne moze razici. Javan
' je zato sto test tvrdi da bas TU stoji identitet: da je privatan, radnja bi
' mogla da procita drugu kolonu a da opis kolona i dalje izgleda ispravno.
Public Const NED_COL_CID As Long = 6

' Visina zone - ista kao na ekranu Palete, pa naslov ispod nje pada u isti
' red na oba ekrana.
Private Const OPO_ZONA_H As Single = KPI_H

Private mLista As String          ' kljuc aktivne liste
' Dva CILJA prevezivanja. Prazan cilj znaci "jos nije izabran" - radnja to
' javlja umesto da preveze na nesto nasumicno.
Private mCiljZbirna As String
Private mCiljPrijemnica As String
' Generacija izabranog CILJA. Drzi se uz broj, ne umesto njega: broj i dalje
' ide u prikaz i u labele na stavkama, generacija samo bira dokument.
Private mCiljZbirnaGen As String
Private mCiljPrijemnicaGen As String
Private mStep As String           ' korak za poruku o gresci
' Brojke u zoni; racuna ih Scr_Rows u istom prolazu kroz podatke.
Private mBrNedovrseno As Long
Private mBrOsirocene As Long
Private mBrUndo As Long

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=OPORAVAK|naslov=OTKUI_NAV_OPORAVAK|sub=OTKUI_SCROPO_SUB" & _
               "|lista=OTKUI_SCROPO_LISTA|oblik=lista|upis=ne"
End Function

Public Function Scr_Liste() As Variant
    Scr_Liste = Array( _
        "NEDOVRSENO|OTKUI_SEG_OPO_NED|OTKUI_GRID_TITLE_NEDOVRSENO|92", _
        "PRIJEMNICE|OTKUI_SEG_OPO_PRI|OTKUI_GRID_TITLE_OSIR_PRIJ|110", _
        "ZBIRNE|OTKUI_SEG_OPO_ZBR|OTKUI_GRID_TITLE_CILJ_ZBIRNA|92", _
        "PALETE|OTKUI_SEG_OPO_PAL|OTKUI_GRID_TITLE_OSIR_PAL|96", _
        "CILJPRIJ|OTKUI_SEG_OPO_CPR|OTKUI_GRID_TITLE_CILJ_PRIJ|116", _
        "UNDO|OTKUI_SEG_OPO_UND|OTKUI_GRID_TITLE_UNDO|100")
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = "NEDOVRSENO"
    Scr_Lista = mLista
End Function

' U listama koje prevezuju naslov nosi izabran CILJ - inace operater ne vidi
' kuda ce radnja da odvede red na koji klikne.
Public Function Scr_NaslovDopuna() As String
    Select Case Scr_Lista()
        Case "PRIJEMNICE": Scr_NaslovDopuna = mCiljZbirna
        Case "PALETE":     Scr_NaslovDopuna = mCiljPrijemnica
    End Select
End Function

Public Function Scr_Radnje() As String
    Select Case Scr_Lista()
        Case "NEDOVRSENO"
            ' Zaostao context ispravke blokira prevezivanje za nove dokumente
            ' tog tipa: safe-stop odbija da nagadja KOJU od vise ispravki novi
            ' dokument zamenjuje. Do sada se to razresavalo SAMO kroz legacy
            ' formu -- CancelCorrectionContext je postojao, ali mu nov UI nije
            ' imao ulaz, pa je lista bila cist pregled bez izlaza.
            '
            ' Radnja MENJA podatke i tesko se poziva nazad -- otud danger stil,
            ' isto kao "Vrati storno".
            Scr_Radnje = "odbaci:OTKUI_BTN_OPO_ODBACI:150:danger:1"
        Case "PRIJEMNICE"
            Scr_Radnje = "prevezipri:OTKUI_BTN_OPO_PREVEZI:96:soft:1"
        Case "PALETE"
            Scr_Radnje = "prevezipal:OTKUI_BTN_OPO_PREVEZI:96:soft:1"
        Case "UNDO"
            ' "Vrati storno" MENJA podatke i tesko se poziva nazad, pa nosi
            ' danger stil - isto kao samo dugme storna.
            Scr_Radnje = "vrati:OTKUI_BTN_OPO_VRATI:112:danger:1"
    End Select
End Function

'--------------------------------------------------------------- ZONA
' Levo dva cilja prevezivanja, desno tri brojke - isti raspored kao traka
' aktivne palete na ekranu Palete.
Public Sub Scr_Build(ByVal z As Object)
    Dim i As Long
    modUiKit.NewLbl z, "opoCap", UCase$(Poruka("OTKUI_OPO_CILJEVI")), PAD, 6, 200, 11, _
                    TS_MICRO, True, C_MUTED, -1
    modUiKit.NewLbl z, "opoZbr", ChrW(8212), PAD, 18, 300, 20, TS_KPI, True, C_FOREST, -1
    modUiKit.NewLbl z, "opoPrj", ChrW(8212), PAD, 40, 420, 13, TS_META, False, C_MUTED, -1

    For i = 0 To 2
        modUiKit.NewLbl z, "opoKL" & i, "", 0, 6, 130, 12, TS_MICRO, True, C_MUTED, -1
        modUiKit.NewLbl z, "opoKV" & i, ChrW(8212), 0, 18, 130, 20, TS_KPI, True, _
                        C_FOREST, -1, fmTextAlignLeft, F_NUM
    Next i
    modUiKit.NewLbl z, "opoLnB", "", 0, OPO_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Dim i As Long
    On Error Resume Next
    For i = 0 To 2
        z.Controls("opoKL" & i).Left = w - PAD - (3 - i) * 160
        z.Controls("opoKV" & i).Left = w - PAD - (3 - i) * 160
    Next i
    z.Controls("opoLnB").width = w
    Scr_Layout = OPO_ZONA_H
End Function

Private Sub OsveziZonu()
    Dim z As Object
    On Error Resume Next
    Set z = modOtkupUI.ScreenZone("OPORAVAK")
    If z Is Nothing Then Exit Sub
    z.Controls("opoZbr").caption = Poruka("OTKUI_OPO_CILJ_ZBIRNA") & " " & _
                                   IIf(Len(mCiljZbirna) = 0, ChrW(8212), mCiljZbirna)
    z.Controls("opoPrj").caption = Poruka("OTKUI_OPO_CILJ_PRIJ") & " " & _
                                   IIf(Len(mCiljPrijemnica) = 0, ChrW(8212), mCiljPrijemnica)
    z.Controls("opoKL0").caption = UCase$(Poruka("OTKUI_OPO_KPI_NED"))
    z.Controls("opoKV0").caption = CStr(mBrNedovrseno)
    z.Controls("opoKL1").caption = UCase$(Poruka("OTKUI_OPO_KPI_OSIR"))
    z.Controls("opoKV1").caption = CStr(mBrOsirocene)
    z.Controls("opoKL2").caption = UCase$(Poruka("OTKUI_OPO_KPI_UNDO"))
    z.Controls("opoKV2").caption = CStr(mBrUndo)
End Sub

'-------------------------------------------------------------- RADNJE
Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Dim broj As String
    On Error GoTo EH

    If Left$(tag, 2) = "ls" Then
        If Mid$(tag, 3) = Scr_Lista() Then Exit Function
        mLista = Mid$(tag, 3)
        Scr_Event = True
        Exit Function
    End If

    ' Klik na red u listi ciljeva BIRA cilj. Vraca False: izbor ne menja
    ' podatke, pa mreza ne sme da se procita ponovo - operater bi izgubio
    ' mesto u listi.
    If Left$(tag, 4) = "row:" Then
        Select Case Scr_Lista()
            ' Uz broj se pamti i GENERACIJA izabranog reda: broj je labela, a
            ' radnja mora da posalje bas ovaj dokument. Kolona je poslednja i
            ' prioriteta 3 -- na uskom ekranu se sakrije, u podacima ostaje.
            Case "ZBIRNE"
                broj = Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 1)))
                If Len(broj) > 0 Then
                    mCiljZbirna = broj
                    mCiljZbirnaGen = GenIzReda(CLng(Mid$(tag, 5)), 7)
                    OsveziZonu
                End If
            Case "CILJPRIJ"
                broj = Trim$(CStr(modOtkupUI.GridCell(CLng(Mid$(tag, 5)), 1)))
                If Len(broj) > 0 Then
                    mCiljPrijemnica = broj
                    mCiljPrijemnicaGen = GenIzReda(CLng(Mid$(tag, 5)), 7)
                    OsveziZonu
                End If
        End Select
        Exit Function
    End If

    If Left$(tag, 4) = "act:" Then
        Scr_Event = OpoAkcija(tag)
        Exit Function
    End If
    Exit Function
EH:
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Vraca True samo ako je radnja PROMENILA podatke (mreza se tada cita ponovo).
' Generacija iz reda mreze. Kolona je prioriteta 3, pa se na uskom ekranu ne
' CRTA - ali podatak ostaje u mView, a GridCell cita njega, ne sirinu.
Private Function GenIzReda(ByVal red As Long, ByVal kolona As Long) As String
    On Error Resume Next
    GenIzReda = Trim$(CStr(modOtkupUI.GridCell(red, kolona)))
End Function

Private Function OpoAkcija(ByVal tag As String) As Boolean
    Dim p() As String, red As Long, kljuc As String
    On Error GoTo EH
    p = Split(Mid$(tag, 5), ":")
    If UBound(p) < 1 Then Exit Function
    red = CLng(val(p(1)))
    kljuc = Trim$(CStr(modOtkupUI.GridCell(red, 1)))
    If Len(kljuc) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_NEMA_REDA"), True
        Exit Function
    End If

    ' GENERACIJA je poslednja kolona obe izvorne liste. Ona je identitet
    ' dokumenta - broj je labela. Bez nje bi prevezivanje pogodilo i dokument
    ' drugog kupca koji deli broj.
    Select Case p(0)
        Case "prevezipri": OpoAkcija = PreveziPrijemnicu(kljuc, GenIzReda(red, 8))
        Case "prevezipal": OpoAkcija = PreveziPalete(kljuc, GenIzReda(red, 7))
        Case "vrati":      OpoAkcija = VratiStorno(kljuc, red)
        Case "odbaci":     OpoAkcija = OdbaciIspravku(red)
        Case Else
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & p(0), True
    End Select
    Exit Function
EH:
    LogErr "modScrOporavak.OpoAkcija"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Osirotela prijemnica dobija novu zbirnu. Menja se SAMO BrojZbirne -
' PrijemnicaID ostaje, pa faktura i palete ostaju ispravne. Ceo posao radi
' postojeci ReassignPrijemnicaToZbirna_TX, ukljucujuci proveru da je cilj
' aktivna zbirna.
Private Function PreveziPrijemnicu(ByVal brojPrij As String, ByVal gen As String) As Boolean
    On Error GoTo EH
    If Len(mCiljZbirna) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_NEMA_ZBIRNE"), True
        Exit Function
    End If
    If MsgBox(Poruka("OTKUI_ASK_OPO_PRI") & " " & brojPrij & " " & _
              Poruka("OTKUI_ASK_OPO_NA") & " " & mCiljZbirna & "?", _
              vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function

    If ReassignPrijemnicaToZbirna_TX(brojPrij, mCiljZbirna, gen, mCiljZbirnaGen) Then
        Scr_ResetCache
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_OPO_PREVEZANO") & " " & brojPrij & _
                             " " & ChrW(8594) & " " & mCiljZbirna, False
        PreveziPrijemnicu = True
    Else
        ' Rutina vraca samo True/False; razlog (cilj ne postoji ili je
        ' storniran) se ne prosledjuje, pa poruka kaze sta da se proveri.
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_PRI") & " " & mCiljZbirna, True
    End If
    Exit Function
EH:
    LogErr "modScrOporavak.PreveziPrijemnicu"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Osirotele paletne stavke prelaze na aktivnu prijemnicu. force=True je
' namerno: ovaj ekran postoji bas da razresi slucajeve koje automatika nije
' umela, pa razlika u broju gajbica ne sme da bude kapija - prijavljuje se
' i koriguje u mestu (PaletaAdjustPrompt).
Private Function PreveziPalete(ByVal brojPrij As String, ByVal gen As String) As Boolean
    Dim upoz As String, gajbDiff As Boolean
    On Error GoTo EH
    If Len(mCiljPrijemnica) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_NEMA_PRIJ"), True
        Exit Function
    End If
    ' "Isti dokument" se meri GENERACIJOM kad je obe strane imaju. Broj je labela
    ' i nastaje PO KUPCU, pa ispravka koja menja kupca lako dobije isti poslovni
    ' broj kao original -- a to su dva dokumenta i prevezivanje je legitimno.
    ' Poredjenje po broju bi tu odbilo potpuno ispravnu operaciju.
    If Len(gen) > 0 And Len(mCiljPrijemnicaGen) > 0 Then
        If StrComp(gen, mCiljPrijemnicaGen, vbTextCompare) = 0 Then
            modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_ISTA"), True
            Exit Function
        End If
    ElseIf StrComp(brojPrij, mCiljPrijemnica, vbTextCompare) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_ISTA"), True
        Exit Function
    End If
    ' Kapija nad dvosmislenim CILJEM vise nije ovde: ista provera sada stoji u
    ' samom writeru (VlasniciPoBroju / RequireJedanVlasnikPoBroju), pa vazi za
    ' SVAKOG pozivaoca, i za legacy formu. Ekran samo prikaze razlog.
    If MsgBox(Poruka("OTKUI_ASK_OPO_PAL") & " " & brojPrij & " " & _
              Poruka("OTKUI_ASK_OPO_NA") & " " & mCiljPrijemnica & "?", _
              vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Function

    If ReassignPaleteToPrijemnica_TX(brojPrij, mCiljPrijemnica, upoz, True, gajbDiff, _
                                     gen, mCiljPrijemnicaGen) Then
        Scr_ResetCache
        MsgBox Poruka("OTKUI_MSG_OPO_PREVEZANO") & " " & brojPrij & " " & ChrW(8594) & _
               " " & mCiljPrijemnica & _
               IIf(Len(upoz) > 0, vbCrLf & vbCrLf & upoz, "") & _
               IIf(gajbDiff, vbCrLf & vbCrLf & PaletaAdjustPrompt(mCiljPrijemnica), ""), _
               vbInformation, APP_NAME
        PreveziPalete = True
    Else
        LogRelinkFailure brojPrij, mCiljPrijemnica, upoz
        MsgBox Poruka("OTKUI_ERR_OPO_PAL") & vbCrLf & vbCrLf & upoz, vbExclamation, APP_NAME
    End If
    Exit Function
EH:
    LogErr "modScrOporavak.PreveziPalete"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

' Vracanje storna. Cilja KONKRETNU operaciju po OperationID, ne poslednju po
' broju: isti broj dokumenta moze imati vise generacija, pa bi "poslednja po
' broju" vracala tudju. Zato je prva kolona liste bas OperationID.
'
' Kapija je UndoGuardReason (fail-closed) - ista koju dize i legacy dugme.
Private Function VratiStorno(ByVal opID As String, ByVal red As Long) As Boolean
    Dim tip As String, broj As String, razlog As String
    On Error GoTo EH
    tip = Trim$(CStr(modOtkupUI.GridCell(red, 3)))
    broj = Trim$(CStr(modOtkupUI.GridCell(red, 4)))

    razlog = UndoGuardReason(tip, broj)
    If Len(razlog) > 0 Then
        MsgBox razlog, vbExclamation, APP_NAME
        Exit Function
    End If

    If MsgBox(Poruka("OTKUI_ASK_OPO_UNDO") & " " & tip & " " & broj & "?" & vbCrLf & vbCrLf & _
              Poruka("OTKUI_ASK_OPO_UNDO2"), vbExclamation + vbYesNo + vbDefaultButton2, _
              APP_NAME) = vbNo Then Exit Function

    If UndoOperation_TX(opID) Then
        Scr_ResetCache
        modOtkupUI.ShowToast Poruka("OTKUI_MSG_OPO_VRACENO") & " " & broj, False
        VratiStorno = True
    Else
        modOtkupUI.ShowToast Poruka("OTKUI_ERR_OPO_UNDO") & " " & broj, True
    End If
    Exit Function
EH:
    LogErr "modScrOporavak.VratiStorno"
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & Err.description, True
End Function

Public Sub Scr_ResetCache()
    ' Ovaj ekran nema izvedenih mapa - svaka lista se cita sveza iz svog
    ' izvora pri svakom osvezavanju. Metod postoji zbog ugovora: ljuska ga
    ' zove posle svake promene podataka.
End Sub

'-------------------------------------------------------------- REDOVI
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Select Case Scr_Lista()
        Case "PRIJEMNICE": Scr_Rows = RowsOsiroceninPrijemnica(q): Exit Function
        Case "ZBIRNE":     Scr_Rows = RowsAktivneZbirne(q): Exit Function
        Case "PALETE":     Scr_Rows = RowsOsirocenePalete(q): Exit Function
        Case "CILJPRIJ":   Scr_Rows = RowsAktivnePrijemnice(q): Exit Function
        Case "UNDO":       Scr_Rows = RowsUndo(q): Exit Function
    End Select
    Scr_Rows = RowsNedovrseno(q)
End Function

'--- NEDOVRSENO: jedan pregled svega sto ceka -------------------------
' Poslednja kolona je NEVIDLJIVA (prioritet 4) i nosi CorrectionID. Isti obrazac
' koji ekran Storno koristi za GeneracijaID: broj je ono sto operater vidi, a
' radnja mora da pogodi TACAN zapis -- vise contexta moze da deli isti broj
' dokumenta (storno, pa opet storno istog broja).
Private Function NedGridCols() As Variant
    NedGridCols = Array( _
        "OTKUI_HDO_REF||txt|120|1", _
        "OTKUI_HDO_VRSTA_PROB||txt|150|2", _
        "OTKUI_HD_STATUS||txt|92|1", _
        "OTKUI_HDO_OPIS||part|0|1", _
        "OTKUI_HDO_AKCIJA||txt|210|2", _
        "OTKUI_HDO_CID||txt|0|4")
End Function

Private Function RowsNedovrseno(ByVal q As String) As Variant
    Dim src As Collection, i As Long, n As Long, outA() As Variant
    Dim d As Object, hay As String
    On Error GoTo EH
    mStep = "nedovrseno"
    Set src = GetNedovrseno()
    mBrNedovrseno = 0
    If Not src Is Nothing Then mBrNedovrseno = src.count
    If src Is Nothing Or mBrNedovrseno = 0 Then
        OsveziZonu
        RowsNedovrseno = Array(NedGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To mBrNedovrseno, 1 To NED_COL_CID)
    For i = 1 To src.count
        Set d = src(i)
        hay = CStr(d("ref")) & "|" & CStr(d("kind")) & "|" & CStr(d("status")) & "|" & CStr(d("opis"))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(d("ref"))
        outA(n, 2) = CStr(d("kind"))
        outA(n, 3) = CStr(d("status"))
        outA(n, 4) = CStr(d("opis"))
        outA(n, 5) = CStr(d("akcija"))
        outA(n, NED_COL_CID) = CStr(d("correctionID"))
Sledeci:
    Next i

    OsveziZonu
    mStep = "OK"
    RowsNedovrseno = Array(NedGridCols(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrOporavak.RowsNedovrseno[" & mStep & "]", Err.description
End Function

' Odbaci ZAOSTAO context ispravke. Dokumenti se NE diraju -- zatvara se samo
' zapis koji ceka zamenski dokument, pa safe-stop prestane da blokira
' prevezivanje za taj tip.
'
' Bira se po CorrectionID iz nevidljive kolone, ne po broju.
'
' Redovi koji NISU context (osirotele prijemnice, palete, izgubljeni blokovi)
' nemaju sta da odbace -- oni se resavaju prevezivanjem, pa se radnja nad njima
' ODBIJA umesto da tiho ne uradi nista.
' JEZGRO radnje: gasi TAJ context i nista drugo. Odvojeno od reda mreze i od
' MsgBox-a zato sto je bas ovo tvrdnja koju PR nosi -- da se gasi izabrani, a ne
' prvi, susedni ili bilo koji koji deli poslovni broj. MsgBox u headless runu
' visi, pa se kroz UI ne moze izmeriti.
Public Function OdbaciIspravkuCore(ByVal cid As String) As Boolean
    cid = Trim$(cid)
    If Len(cid) = 0 Then Exit Function
    OdbaciIspravkuCore = modStornoContext.CancelCorrectionContext(cid, _
            "Operater odbacio zaostalu ispravku sa ekrana Oporavak.")
End Function

Private Function OdbaciIspravku(ByVal red As Long) As Boolean
    Dim cid As String, opis As String, errDesc As String
    On Error GoTo EH
    cid = Trim$(CStr(modOtkupUI.GridCell(red, NED_COL_CID)))
    If Len(cid) = 0 Then
        modOtkupUI.ShowToast Poruka("OTKUI_OPO_ODBACI_NIJE_CTX"), True
        Exit Function
    End If

    opis = Trim$(CStr(modOtkupUI.GridCell(red, 1))) & "  " & ChrW(183) & "  " & _
           Trim$(CStr(modOtkupUI.GridCell(red, 2)))
    If MsgBox(Poruka("OTKUI_OPO_ODBACI_ASK") & vbCrLf & vbCrLf & opis, _
              vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then Exit Function

    If Not OdbaciIspravkuCore(cid) Then
        modOtkupUI.ShowToast Poruka("OTKUI_OPO_ODBACI_ERR"), True
        Exit Function
    End If
    modOtkupUI.ShowToast Poruka("OTKUI_OPO_ODBACI_OK"), False
    OdbaciIspravku = True
    Exit Function
EH:
    errDesc = Err.description
    LogErr "modScrOporavak.OdbaciIspravku"
    Err.Clear
    modOtkupUI.ShowToast Poruka("OTKUI_ERR_RADNJA") & " " & errDesc, True
End Function

'--- OSIROTELE PRIJEMNICE --------------------------------------------
' GetOsirocenePrijemnice, kolone 1..7:
'   Broj | Datum | Vrsta | Sorta | Kolicina | StaraZbirna | Status
Private Function OsirPrijGridCols() As Variant
    OsirPrijGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|3", _
        "OTKUI_HD_KG||txt|72|1", _
        "OTKUI_HD_BROJ_ZBIRNE||txt|110|1", _
        "OTKUI_HD_STATUS||txt|132|1", _
        "OTKUI_HDO_GENERACIJA||txt|110|3")
End Function

Private Function RowsOsiroceninPrijemnica(ByVal q As String) As Variant
    RowsOsiroceninPrijemnica = Rows2D(GetOsirocenePrijemnice(), OsirPrijGridCols(), 8, q, True)
End Function

'--- OSIROTELE PALETE -------------------------------------------------
' GetPrijemniceSaOsirocenimPaletama, kolone 1..6:
'   Broj | Datum | Vrsta | Sorta | Gajbica | Stavki
Private Function OsirPalGridCols() As Variant
    OsirPalGridCols = Array( _
        "OTKUI_HD_BROJ||txt|110|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|3", _
        "OTKUI_HDP_GAJBICA||num|72|1", _
        "OTKUI_HD_STAVKI||num|72|1", _
        "OTKUI_HDO_GENERACIJA||txt|110|3")
End Function

Private Function RowsOsirocenePalete(ByVal q As String) As Variant
    RowsOsirocenePalete = Rows2D(GetPrijemniceSaOsirocenimPaletama(), OsirPalGridCols(), 7, q, True)
End Function

' Zajednicko telo za dve liste koje dolaze kao 2D niz iz modDokumenta.
' brojiUZonu: ove dve liste zajedno cine brojku "osirotelo" u zoni.
Private Function Rows2D(ByVal src As Variant, ByVal cols As Variant, ByVal nCol As Long, _
                        ByVal q As String, ByVal brojiUZonu As Boolean) As Variant
    Dim r As Long, c As Long, n As Long, outA() As Variant, hay As String
    On Error GoTo EH
    mStep = "2D lista"
    If Not IsArray(src) Then
        If brojiUZonu Then mBrOsirocene = ZbirOsirocenih()
        OsveziZonu
        Rows2D = Array(cols, Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To UBound(src, 1), 1 To nCol)
    For r = 1 To UBound(src, 1)
        hay = ""
        For c = 1 To nCol
            hay = hay & "|" & NzToText(src(r, c))
        Next c
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        For c = 1 To nCol
            outA(n, c) = src(r, c)
        Next c
Sledeci:
    Next r

    If brojiUZonu Then mBrOsirocene = ZbirOsirocenih()
    OsveziZonu
    mStep = "OK"
    Rows2D = Array(cols, outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrOporavak.Rows2D[" & mStep & "]", Err.description
End Function

' Brojka u zoni je zbir OBE vrste siroticica, bez obzira koja je lista
' otvorena - inace bi prelazak izmedju lista menjao brojku koja treba da
' kaze koliko posla ukupno ceka.
Private Function ZbirOsirocenih() As Long
    Dim a As Variant, b As Variant
    On Error Resume Next
    a = GetOsirocenePrijemnice()
    b = GetPrijemniceSaOsirocenimPaletama()
    If IsArray(a) Then ZbirOsirocenih = ZbirOsirocenih + UBound(a, 1)
    If IsArray(b) Then ZbirOsirocenih = ZbirOsirocenih + UBound(b, 1)
End Function

'--- CILJEVI: aktivne zbirne i aktivne prijemnice ---------------------
Private Function CiljZbirnaGridCols() As Variant
    CiljZbirnaGridCols = Array( _
        "OTKUI_HD_BROJ||txt|120|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HDO_VLASNIK||txt|120|1", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|3", _
        "OTKUI_HD_KG||kg|76|1", _
        "OTKUI_HDO_GENERACIJA||txt|110|3")
End Function

Private Function RowsAktivneZbirne(ByVal q As String) As Variant
    ' Vlasnistvo zbirne je VOZAC + KUPAC. Broj zbirne generator drzi
    ' jedinstvenim, pa se dva reda istog broja mogu pojaviti samo iz rucnog
    ' unosa ili uvoza; tada sam kupac ne razlikuje dokumenta.
    RowsAktivneZbirne = RowsAktivni(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_DATUM, _
                                    COL_ZBR_VRSTA, COL_ZBR_SORTA, COL_ZBR_KOLICINA, _
                                    Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC), _
                                    CiljZbirnaGridCols(), q)
End Function

Private Function CiljPrijGridCols() As Variant
    CiljPrijGridCols = Array( _
        "OTKUI_HD_BROJ||txt|120|1", _
        "OTKUI_HD_DATUM||date|62|1", _
        "OTKUI_HDO_VLASNIK||txt|120|1", _
        "OTKUI_HD_VRSTA||txt|84|2", _
        "OTKUI_HD_SORTA||part|0|3", _
        "OTKUI_HD_KG||kg|76|1", _
        "OTKUI_HDO_GENERACIJA||txt|110|3")
End Function

Private Function RowsAktivnePrijemnice(ByVal q As String) As Variant
    RowsAktivnePrijemnice = RowsAktivni(TBL_PRIJEMNICA, COL_PRJ_BROJ, COL_PRJ_DATUM, _
                                        COL_PRJ_VRSTA, COL_PRJ_SORTA, COL_PRJ_KOLICINA, _
                                        COL_PRJ_KUPAC, CiljPrijGridCols(), q)
End Function

' Liste kandidata za cilj: nestornirani redovi, jedan red po DOKUMENTU.
'
' "Jedan red po dokumentu" NIJE isto sto i "jedan red po broju", i ta razlika
' je razlog zasto ova rutina izgleda ovako:
'
'   klase I i II dele broj I VLASNIKA -> to je jedan dokument, jedan red
'   dva kupca sa istim brojem         -> to su DVA dokumenta, dva reda
'
' Broj se racuna PO KUPCU (GenerateBrojPrijemnice), pa je drugi slucaj
' svakodnevan, ne teorijski. Dedup samo po broju sveo bi ta dva dokumenta na
' jedan red - operater bi izabrao "cilj" ne znajuci ciji je, a prevezivanje bi
' otislo na onaj koji zatekne poslednji.
'
' Zato je kljuc grupisanja broj + VLASNIK, a vlasnik se i PRIKAZUJE: kolizija
' mora da bude vidljiva, ne samo tacno obradjena.
' colVlasnik je NIZ kolona, ne jedna: vlasnistvo zbirne je vozac + kupac, isti
' kompozit koji koriste StornoZbirna, ApplyGeneracijaID i kapija u
' ReassignPrijemnicaToZbirna_TX. Sa samim kupcem bi dve zbirne istog broja i
' istog kupca a razlicitih vozaca -- u jezgru dva dokumenta -- pale u JEDAN red,
' pa operater ne bi mogao ni da izabere onaj koji mu treba.
Private Function RowsAktivni(ByVal tbl As String, ByVal colBroj As String, _
                             ByVal colDatum As String, ByVal colVrsta As String, _
                             ByVal colSorta As String, ByVal colKol As String, _
                             ByVal colVlasnik As Variant, _
                             ByVal cols As Variant, ByVal q As String) As Variant
    Dim src As Variant, r As Long, n As Long, outA() As Variant
    Dim iBr As Long, iDat As Long, iVr As Long, iSo As Long, iKol As Long
    Dim iSt As Long, iGen As Long
    Dim vlCols As Variant, iVl() As Long, v As Long
    Dim broj As String, vlasnik As String, kljuc As String, hay As String, sumKg As Double
    Dim gen As String
    Dim seen As Object
    On Error GoTo EH
    mStep = "cilj " & tbl

    src = modUiData.CachedTable(tbl)
    If Not IsArray(src) Then
        RowsAktivni = Array(cols, Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    iBr = modUiData.ColIdx(tbl, colBroj)
    iDat = modUiData.ColIdx(tbl, colDatum)
    iVr = modUiData.ColIdx(tbl, colVrsta)
    iSo = modUiData.ColIdx(tbl, colSorta)
    iKol = modUiData.ColIdx(tbl, colKol)
    If IsArray(colVlasnik) Then
        vlCols = colVlasnik
    Else
        vlCols = Array(colVlasnik)
    End If
    ReDim iVl(LBound(vlCols) To UBound(vlCols))
    For v = LBound(vlCols) To UBound(vlCols)
        iVl(v) = modUiData.ColIdx(tbl, CStr(vlCols(v)))
    Next v
    iSt = modUiData.ColIdx(tbl, COL_STORNIRANO)
    iGen = modUiData.ColIdx(tbl, COL_GENERACIJA_ID)
    If iBr = 0 Then
        RowsAktivni = Array(cols, Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ' Klase I i II su JEDAN dokument: dele broj i vlasnika. Prvi red pravi
    ' izlazni red, svaki sledeci samo DODAJE svoje kilograme - ranije se
    ' preskakao, pa je dvoklasni dokument prikazivao samo Klasu I, a i zbir
    ' u podnozju je bio manji od stvarnog.
    '
    ' Filter se primenjuje PRE nego sto se dokument upamti: da se prvo pamti,
    ' dokument koji ne pogadja pretragu prvim redom ne bi mogao da je pogodi
    ' drugim.
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    ReDim outA(1 To UBound(src, 1), 1 To 7)
    For r = 1 To UBound(src, 1)
        If iSt > 0 Then
            If UCase$(modUiData.CellS(src, r, iSt)) = "DA" Then GoTo Sledeci
        End If
        broj = modUiData.CellS(src, r, iBr)
        If Len(broj) = 0 Then GoTo Sledeci
        vlasnik = ""
        For v = LBound(iVl) To UBound(iVl)
            If iVl(v) > 0 Then
                If Len(vlasnik) > 0 Then vlasnik = vlasnik & " / "
                vlasnik = vlasnik & modUiData.CellS(src, r, iVl(v))
            End If
        Next v

        ' Kljuc grupisanja: GENERACIJA kad postoji, inace broj + pun vlasnik.
        ' Klase I i II istog dokumenta dele generaciju, pa se i dalje sabiraju u
        ' jedan red; dva razlicita dokumenta je vise ne dele, pa se ne spajaju.
        gen = ""
        If iGen > 0 Then gen = modUiData.CellS(src, r, iGen)
        If Len(gen) > 0 Then
            kljuc = "g" & Chr$(1) & gen
        Else
            kljuc = broj & Chr$(1) & vlasnik
        End If
        hay = broj & "|" & vlasnik & "|" & modUiData.CellS(src, r, iVr) & _
              "|" & modUiData.CellS(src, r, iSo)
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        If seen.Exists(kljuc) Then
            ' druga klasa istog dokumenta - samo dopuni kolicinu
            outA(CLng(seen(kljuc)), 6) = CDbl(outA(CLng(seen(kljuc)), 6)) + _
                                         modUiData.CellD(src, r, iKol)
            sumKg = sumKg + modUiData.CellD(src, r, iKol)
            GoTo Sledeci
        End If
        n = n + 1
        seen(kljuc) = n
        outA(n, 1) = broj
        outA(n, 2) = modUiData.CellDate(src, r, iDat)
        outA(n, 3) = vlasnik
        outA(n, 4) = modUiData.CellS(src, r, iVr)
        outA(n, 5) = modUiData.CellS(src, r, iSo)
        outA(n, 6) = modUiData.CellD(src, r, iKol)
        ' Identitet ciljnog dokumenta ide u red: radnja mora da posalje BAS
        ' ovaj dokument, ne "prvi sa tim brojem".
        outA(n, 7) = gen
        sumKg = sumKg + modUiData.CellD(src, r, iKol)
Sledeci:
    Next r

    mStep = "OK"
    RowsAktivni = Array(cols, outA, n, sumKg, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrOporavak.RowsAktivni[" & mStep & "]", Err.description
End Function

'--- UNDO: storno operacije koje se mogu vratiti ----------------------
' Prva kolona je OperationID i to nije previd: radnja cilja BAS tu
' operaciju, a isti broj dokumenta moze imati vise generacija (razni
' OperationID-evi). Broj i tip stoje uz njega, da operater vidi sta vraca.
Private Function UndoGridCols() As Variant
    UndoGridCols = Array( _
        "OTKUI_HDO_OPID||txt|150|1", _
        "OTKUI_HDO_VREME||txt|118|2", _
        "OTKUI_HDO_TIPDOK||txt|110|1", _
        "OTKUI_HD_BROJ||part|0|1", _
        "OTKUI_HD_STAVKI||num|66|2", _
        "OTKUI_HD_STATUS||txt|100|1")
End Function

Private Function RowsUndo(ByVal q As String) As Variant
    Dim src As Collection, i As Long, n As Long, outA() As Variant
    Dim d As Object, hay As String
    On Error GoTo EH
    mStep = "undo"
    Set src = GetUndoableStornoOperations()
    mBrUndo = 0
    If Not src Is Nothing Then mBrUndo = src.count
    If src Is Nothing Or mBrUndo = 0 Then
        OsveziZonu
        RowsUndo = Array(UndoGridCols(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    ReDim outA(1 To mBrUndo, 1 To 6)
    For i = 1 To src.count
        Set d = src(i)
        hay = CStr(d("opID")) & "|" & CStr(d("docType")) & "|" & CStr(d("broj")) & _
              "|" & CStr(d("status"))
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then GoTo Sledeci
        End If
        n = n + 1
        outA(n, 1) = CStr(d("opID"))
        outA(n, 2) = CStr(d("ts"))
        outA(n, 3) = CStr(d("docType"))
        outA(n, 4) = CStr(d("broj"))
        outA(n, 5) = CLng(d("count"))
        outA(n, 6) = CStr(d("status"))
Sledeci:
    Next i

    OsveziZonu
    mStep = "OK"
    RowsUndo = Array(UndoGridCols(), outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modScrOporavak.RowsUndo[" & mStep & "]", Err.description
End Function

'------------------------------------------------------------ TEST SEAM
' Test bira aktivnu listu i ciljeve bez klika po prekidacu. Tvrdo gejtovano,
' isto kao Scr_OtpTestSet - van test-rezima ne radi nista.
Public Sub Scr_OpoTestSet(ByVal lista As String, ByVal ciljZbirna As String, _
                          ByVal ciljPrijemnica As String, _
                          Optional ByVal ciljZbirnaGen As String = "", _
                          Optional ByVal ciljPrijemnicaGen As String = "")
    If Not IsTestMode() Then Exit Sub
    mLista = lista
    mCiljZbirna = ciljZbirna
    mCiljPrijemnica = ciljPrijemnica
    mCiljZbirnaGen = ciljZbirnaGen
    mCiljPrijemnicaGen = ciljPrijemnicaGen
End Sub
