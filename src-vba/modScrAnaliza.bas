Attribute VB_Name = "modScrAnaliza"
'=====================================================================
' modScrAnaliza - ekran "Analiza poslovanja" (grupa FINANSIJE).
'
' EKRAN JE U IZRADI i to i pise na njemu. U registru stoji zato sto je uzeo
' mesto reda MARZA, koji je od S3a pokazivao na modul koji nikad nije napisan
' (modScrMarza): stavka se crtala prigusena, a klik je govorio da ekrana nema.
' Legacy forma frmMarza je nudila tri pogleda na marzu (po kupcu, po otkupnom
' mestu, ukupno) i po korisnikovoj reci se nije koristila; obrisana je zajedno
' sa uvodjenjem ovog ekrana (docs/UI_MIGRACIJA_KATALOG.md par.27.15).
'
' Sta ovde treba da stane SIRE je od marze -- poslovno-finansijske analize nad
' podacima koje aplikacija vec ima. Racun koji je forma zvala i dalje stoji u
' modMarza, ali ga ovaj ekran NE zove: audit (docs/AUDIT_FM_TRIJAZA.md, FM-0106)
' je zabelezio da ta tri pogleda mesaju PROCENU sa OSTVARENOM marzom, pa se prvo
' bira sta se od toga uopste prikazuje -- a to je posao za sledeci korak, ne za
' ovaj.
'
' Do tada ekran radi tacno ono sto ugovor trazi i nista vise: zona kaze da je u
' izradi, mreza je prazna. Nista se ne racuna, pa se nista ni ne tvrdi.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRAN_BUILD As String = "v6-ui-213"

Private Const EKRAN As String = "ANALIZA"

' Zona je visine KPI trake -- ekran nema sta da racuna, pa ni da zauzme vise.
Private Const ANA_ZONA_H As Single = KPI_H

'--------------------------------------------------------- UGOVOR EKRANA
Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=" & EKRAN & "|naslov=OTKUI_NAV_ANALIZA|sub=OTKUI_SCRAN_SUB" & _
               "|lista=OTKUI_SCRAN_LISTA|oblik=zona|upis=ne"
End Function

'--------------------------------------------------------------- ZONA
Public Sub Scr_Build(ByVal z As Object)
    modUiKit.NewLbl z, "anaCap", UCase$(Poruka("OTKUI_SCRAN_CAP")), PAD, 6, 200, 11, _
                    TS_MICRO, True, C_GOLD, -1
    modUiKit.NewLbl z, "anaNaslov", Poruka("OTKUI_SCRAN_IZRADA"), PAD, 18, 460, 20, _
                    TS_KPI, True, C_FOREST, -1
    modUiKit.NewLbl z, "anaOpis", Poruka("OTKUI_SCRAN_OPIS"), PAD, 40, 620, 13, _
                    TS_META, False, C_MUTED, -1
    modUiKit.NewLbl z, "anaLn", "", 0, ANA_ZONA_H - 1, 100, 1, 8, False, 0, C_BORDER
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    On Error Resume Next
    z.Controls("anaOpis").width = w - 2 * PAD
    z.Controls("anaLn").width = w
    Scr_Layout = ANA_ZONA_H
End Function

' PRAZNA mreza, ali sa SVOJIM kolonama.
'
' Ekran koji vrati Empty ostavlja mrezu na kolonama PRETHODNOG ekrana (v.
' LoadGridFromScreen): zaglavlje tudje liste stoji, celije su prazne. Zato ide
' uredan prazan odgovor -- jedna kolona i nula redova.
Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Scr_Rows = Array(AnaGridCols(), Empty, 0, 0#, 0#)
End Function

Private Function AnaGridCols() As Variant
    AnaGridCols = Array("OTKUI_SCRAN_HD_POKAZ||txt|0|1")
End Function
