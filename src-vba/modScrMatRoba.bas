Attribute VB_Name = "modScrMatRoba"
'=====================================================================
' modScrMatRoba - ekran "Proizvodi i cene" (sekcija MATICNI, korak M1).
'
' Cetiri liste: Artikli, Kulture, Cenovnik, Vrsta gotovog proizvoda -- grupa
' "Proizvodi i cene" iz legacy menija Maticni podaci.
'
' Modul je TANAK NAMERNO: sve o podacima zna modMaticniIzvor, sve o zoni i
' dogadjajima modMaticniEkran. Ovde stoji samo koje sekcije ekran nosi.
'
' UNOS I IZMENA rade od M2 -- v. modScrMatPartneri i UI_MIGRACIJA_KATALOG 26.15.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRMR_BUILD As String = "v6-ui-193"

Private Const EKRAN As String = "MAT_ROBA"

Private mLista As String

Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=" & EKRAN & "|naslov=OTKUI_NAV_MAT_ROBA|sub=OTKUI_SCRMR_SUB" & _
               "|lista=OTKUI_GTM_ART|oblik=zona+mreza|upis=zona"
End Function

Public Function Scr_Liste() As Variant
    Scr_Liste = modMaticniIzvor.MatSekcijeEkrana(EKRAN)
End Function

Public Function Scr_Lista() As String
    If Len(mLista) = 0 Then mLista = modMaticniIzvor.MatPrvaSekcija(EKRAN)
    Scr_Lista = mLista
End Function

Public Function Scr_Cipovi() As String
    Scr_Cipovi = modMaticniIzvor.MatCipovi(Scr_Lista())
End Function

Public Function Scr_Radnje() As String
    Scr_Radnje = modMaticniEkran.Radnje(Scr_Lista())
End Function

Public Function Scr_Sort() As String
    Scr_Sort = modMaticniIzvor.MatSort(Scr_Lista())
End Function

Public Sub Scr_Build(ByVal z As Object)
    modMaticniEkran.ZonaGradi z
End Sub

Public Function Scr_Layout(ByVal z As Object, ByVal w As Single, ByVal h As Single) As Single
    Scr_Layout = modMaticniEkran.ZonaRaspored(z, w, Scr_Lista())
End Function

Public Function Scr_Rows(ByVal filter As String, ByVal q As String) As Variant
    Scr_Rows = modMaticniEkran.Redovi(EKRAN, Scr_Lista(), filter, q)
End Function

Public Function Scr_Event(ByVal tag As String, ByVal ev As String) As Boolean
    Scr_Event = modMaticniEkran.Dogadjaj(tag, mLista)
End Function

' Ljuska javlja da se ekran napusta -- editor, GEO panel i izbor odlaze s
' njim. Bez ovoga bi otvoren unos preziveo prelazak na drugi ekran (v.
' UI_MIGRACIJA_KATALOG 26.25).
Public Sub Scr_Deaktiviraj()
    modMaticniEkran.Deaktiviraj
End Sub

Public Sub Scr_ResetCache()
    modMaticniIzvor.MatResetCache
End Sub

'------------------------------------------------------------ TEST SEAM
Public Sub Scr_MrTestSet(ByVal lista As String)
    mLista = lista
End Sub
