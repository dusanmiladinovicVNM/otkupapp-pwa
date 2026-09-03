Attribute VB_Name = "modScrMatKorisnici"
'=====================================================================
' modScrMatKorisnici - ekran "Korisnici" (sekcija MATICNI, korak M4).
'
' Dve liste: Korisnici (nalozi) i Prava pristupa (matrica po oblasti). Druga
' ZAVISI od prve -- prava se citaju za korisnika izabranog u listi Korisnici --
' pa su na istom ekranu: prekidac je jedan potez, a promena ekrana bi izgubila
' izbor.
'
' Modul je TANAK, kao i ostala tri maticna ekrana: sve o podacima zna
' modMaticniIzvor (koji Korisnike i Prava prosledjuje u modMaticniKorisnici),
' a zonu, editor i dogadjaje deli modMaticniEkran. Cetvrti ekran sa cetvrtom
' kopijom istog rasporeda bio bi cetvrto mesto koje se prvom doradom razidje.
'
' BRANA: pored oblasti MaticniPodaci trazi i administraciju. Prava pristupa su
' jedina lista na kojoj se moze dodeliti pristup samom sebi, pa obicnom
' korisniku ne sme da bude otvorena. Odgovara se kroz Scr_Dozvoljen, isto kao
' na ekranu Podesavanja i alati -- stavka menija ostaje vidljiva ali prigusena.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const SCRMKOR_BUILD As String = "v6-ui-193"

Private Const EKRAN As String = "MAT_KORISNICI"

Private mLista As String

' Test seam za branu. Sme SAMO da je zatvori, nikad da je otvori -- iz istog
' razloga kao u modUiPanel: u headless runu je MozeAdministraciju
' anti-lockout (bez AUTH-a svi su admini), pa bi tvrdnja "ljuska postuje branu"
' inace merila dva puta True.
Private mBranaZatvorenaTest As Boolean

Public Function Scr_Meta() As String
    Scr_Meta = "kljuc=" & EKRAN & "|naslov=OTKUI_NAV_MAT_KORISNICI|sub=OTKUI_SCRMK_SUB" & _
               "|lista=OTKUI_GTM_KOR|oblik=zona+mreza|upis=zona"
End Function

Public Function Scr_Dozvoljen() As Boolean
    If mBranaZatvorenaTest And IsTestMode() Then Exit Function
    Scr_Dozvoljen = modAuth.MozeAdministraciju()
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
Public Function Scr_ImaNesacuvano() As Boolean
    Scr_ImaNesacuvano = modMaticniEkran.ImaNesacuvano(EKRAN)
End Function

Public Sub Scr_Deaktiviraj()
    modMaticniEkran.Deaktiviraj
End Sub

Public Sub Scr_ResetCache()
    modMaticniIzvor.MatResetCache
End Sub

'------------------------------------------------------------ TEST SEAM
Public Sub Scr_MkorTestSet(ByVal lista As String)
    mLista = lista
End Sub

' Zatvara branu za test. Otvaranje nije moguce -- v. komentar uz mBranaZatvorenaTest.
Public Sub Scr_MkorBranaZatvoriTest(ByVal zatvori As Boolean)
    If Not IsTestMode() Then Exit Sub
    mBranaZatvorenaTest = zatvori
End Sub
