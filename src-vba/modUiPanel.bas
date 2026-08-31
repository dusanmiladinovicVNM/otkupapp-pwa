Attribute VB_Name = "modUiPanel"
'=====================================================================
' modUiPanel - registar panela nove ljuske (korak M6).
'
' Panel je sadrzaj koji NIJE lista: zauzme celu radnu povrsinu, sam se
' rasporedi i sam zna svoje kontrole. Podesavanja (97 polja u 11 grupa) i Admin
' (12 komandi u 5 grupa) su takvi -- ni jedno ni drugo ne staje u ugovor ekrana
' (zona od 16 kontrola + mreza), a razlozi su izmereni u UI_MIGRACIJA_KATALOG
' 24.19.
'
' ZASTO OVAJ MODUL, a ne ljuska: modOtkupUI ne sme da zna nijedan panel po
' imenu, isto kao sto ne zna nijedan ekran. Ljuska daje samo PRAZAN OKVIR
' (PanelHost) i ustupanje radne povrsine (PanelRezim); ko taj okvir puni i cime,
' zna iskljucivo ovaj registar. Poziv graditelja je zato kasno vezan i
' kvalifikovan -- Application.Run "modPodesavanja.BuildConfigEditor".
'
' ZASTO NE U FORMU: frmOtkupUI je ljuska bez logike. Nista sto je do sada zivelo
' u frmStammdaten ne ide u nju -- ide u standardni modul, ovaj.
'
' BRANA JE TROSTRUKA i to nije visak: ekran Podesavanja i alati odgovara kroz
' Scr_Dozvoljen, ovaj registar proverava pre otvaranja, a sam graditelj jos
' jednom (AUD-033). Prava pristupa se menjaju zamenom operatera, pa nijedan sloj
' ne sme da veruje prethodnom.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UIPANEL_BUILD As String = "v6-ui-196"

Private Const SRC As String = "modUiPanel"

' Polja reda registra.
Private Const PAN_KLJUC As Long = 0
Private Const PAN_MODUL As Long = 1
Private Const PAN_GRADI As Long = 2
Private Const PAN_NASLOV As Long = 3

' Kljuc panela koji je trenutno u radnoj povrsini. Prazno = nijedan.
Private mAktivan As String

'-------------------------------------------------------------- REGISTAR
' Red: "KLJUC|modul|graditelj|naslov(katalog)".
'
' Graditelj prima JEDAN argument -- okvir domacina. To je isti potpis koji su te
' procedure vec imale (frm As Object), pa se telo panela nije menjalo: menja se
' samo KO je domacin.
Public Function PanelRedovi() As Variant
    PanelRedovi = Array( _
        "PODESAVANJA|modPodesavanja|BuildConfigEditor|OTKUI_MS_PODESAVANJA", _
        "ADMIN|modAdmin|BuildAdminPanel|OTKUI_MS_ADMIN")
End Function

Public Function PanelPolje(ByVal kljuc As String, ByVal idx As Long) As String
    Dim r As Variant, p() As String
    For Each r In PanelRedovi()
        p = Split(CStr(r), "|")
        If StrComp(p(PAN_KLJUC), kljuc, vbTextCompare) = 0 Then
            If idx <= UBound(p) Then PanelPolje = p(idx)
            Exit Function
        End If
    Next r
End Function

Public Function PanelPostoji(ByVal kljuc As String) As Boolean
    PanelPostoji = (Len(PanelPolje(kljuc, PAN_MODUL)) > 0)
End Function

' Kljuc panela koji je otvoren, ili "" ako nijedan.
Public Function PanelAktivan() As String
    PanelAktivan = mAktivan
End Function

'--------------------------------------------------------------- OTVARANJE
' Vraca "" kad je proslo, inace poruku za operatera.
Public Function PanelOtvori(ByVal kljuc As String) As String
    Dim host As Object, m As String, g As String
    On Error GoTo EH

    m = PanelPolje(kljuc, PAN_MODUL)
    g = PanelPolje(kljuc, PAN_GRADI)
    If Len(m) = 0 Or Len(g) = 0 Then
        PanelOtvori = Poruka("UIPAN_ERR_NEPOZNAT") & " " & kljuc
        Exit Function
    End If

    ' Brana registra. Graditelj ima svoju i ona ostaje -- ova samo sprecava da
    ' se radna povrsina ustupi panelu koji ce odmah odbiti da se izgradi, pa da
    ' operater ostane pred praznim okvirom.
    If Not modAuth.MozeAdministraciju() Then
        PanelOtvori = Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA")
        Exit Function
    End If

    Set host = modOtkupUI.PanelHost()
    If host Is Nothing Then
        PanelOtvori = Poruka("UIPAN_ERR_NEMA_MESTA")
        Exit Function
    End If

    ' Prethodni panel se sklanja PRE nego sto novi pocne da gradi -- inace bi
    ' dva panela delila okvir i kontrole bi im se preklopile.
    ZatvoriTiho
    IsprazniOkvir host

    ' Redosled je bitan: rezim PRE gradnje. Graditelj cita host.InsideWidth, a
    ' ona je tacna tek kad okvir dobije svoju meru.
    modOtkupUI.PanelRezim True
    mAktivan = UCase$(kljuc)

    Application.Run m & "." & g, host
    Exit Function
EH:
    PanelOtvori = Poruka("UIPAN_ERR_GRADNJA") & " " & Err.description
    LogError SRC & ".PanelOtvori(" & kljuc & ")", Err.description
    PanelZatvori
End Function

'--------------------------------------------------------------- ZATVARANJE
' Vraca radnu povrsinu ekranu. Bezbedno je zvati i kad nijedan panel nije
' otvoren -- panel se zatvara i iz svog dugmeta i iz ljuske, pa dvostruko
' zatvaranje mora da prodje bez traga.
Public Sub PanelZatvori()
    Dim host As Object
    On Error Resume Next
    ZatvoriTiho
    Set host = modOtkupUI.PanelHost()
    If Not host Is Nothing Then IsprazniOkvir host
    modOtkupUI.PanelRezim False
    Err.Clear
End Sub

' Pusta reference modula panela, ali NE dira okvir. Odvojeno zato sto redosled
' mora da bude: prvo omotaci (WithEvents), pa tek onda kontrole -- omotac koji
' prezivi svoju kontrolu je mrtva referenca koja puca pri sledecem dogadjaju.
Private Sub ZatvoriTiho()
    Dim m As String
    If Len(mAktivan) = 0 Then Exit Sub
    m = PanelPolje(mAktivan, PAN_MODUL)
    mAktivan = ""
    If Len(m) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run m & "." & OslobodiIme(m)
    Err.Clear
End Sub

' Ime procedure koja oslobadja reference datog modula. Izvedeno iz imena modula
' po dogovoru (modPodesavanja -> Podesavanja_Release), pa nov panel ne trazi red
' vise u registru -- samo da postuje isti dogovor.
Private Function OslobodiIme(ByVal m As String) As String
    OslobodiIme = Mid$(m, 4) & "_Release"
End Function

Private Sub IsprazniOkvir(ByVal host As Object)
    Dim i As Long
    On Error Resume Next
    For i = host.Controls.count - 1 To 0 Step -1
        host.Controls.Remove i
    Next i
    Err.Clear
End Sub
