Attribute VB_Name = "modScrMatPartneri"
'=====================================================================
' modScrMatPartneri - ekran "Partneri" (sekcija MATICNI, korak M1).
'
' Pet lista: Kooperanti, Stanice, Kupci, Vozaci, Parcele -- grupa "Sifarnici"
' iz legacy menija Maticni podaci.
'
' Modul je TANAK NAMERNO. Sve o podacima (tabela, PK, kolone, izvedene
' vrednosti, cipovi, brojke) zna modMaticniIzvor; ovde stoji samo koje sekcije
' ekran nosi i kako izgleda njegova zona. Ista podela vazi za modScrMatRoba i
' modScrMatPakovanje -- tri ekrana, jedan opis podataka.
'
' UNOS I IZMENA rade od M2: oba idu kroz modMaticniUnos, istog pisca koga zove
' i legacy forma -- v. docs/UI_MIGRACIJA_KATALOG.md 26.5 i 26.15.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRMP_BUILD As String = "v6-ui-193"

Private Const EKRAN As String = "MAT_PARTNERI"

Private mLista As String

Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=" & EKRAN & "|naslov=OTKUI_NAV_MAT_PARTNERI|sub=OTKUI_SCRMP_SUB" & _
               "|lista=OTKUI_GTM_KOOP|oblik=zona+mreza|upis=zona"
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

Public Sub Scr_ResetCache()
    modMaticniIzvor.MatResetCache
End Sub

'------------------------------------------------------------ TEST SEAM
Public Sub Scr_MpTestSet(ByVal lista As String)
    mLista = lista
End Sub
