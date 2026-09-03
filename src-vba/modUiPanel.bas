Attribute VB_Name = "modUiPanel"
'=====================================================================
' modUiPanel - registar panela nove ljuske (korak M6).
'
' Panel je sadrzaj koji NIJE lista: zauzme celu radnu povrsinu, sam se
' rasporedi i sam zna svoje kontrole. Podesavanja (97 polja u 11 grupa) i Admin
' (12 komandi u 5 grupa) su takvi -- ni jedno ni drugo ne staje u ugovor ekrana
' (zona od 16 kontrola + mreza), a razlozi su izmereni u UI_MIGRACIJA_KATALOG
' 26.19.
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
' BRANA JE TROSTRUKA i to nije visak: sidebar pita PanelDozvoljen pre nego sto
' stavku uopste nacrta punom bojom, PanelOtvori proverava jos jednom pre nego
' sto ustupi radnu povrsinu, a sam graditelj treci put (AUD-033). Prava pristupa
' se menjaju zamenom operatera, pa nijedan sloj ne sme da veruje prethodnom.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UIPANEL_BUILD As String = "v6-ui-203"

Private Const SRC As String = "modUiPanel"

' Polja reda registra.
Private Const PAN_KLJUC As Long = 0
Public Const PAN_MODUL As Long = 1
Private Const PAN_GRADI As Long = 2
Private Const PAN_NASLOV As Long = 3
Public Const PAN_LTAG As Long = 4

' Kljuc panela koji je trenutno u radnoj povrsini. Prazno = nijedan.
Private mAktivan As String

' Test seam za branu. Sme SAMO da je zatvori, nikad da je otvori -- brana koju
' test moze da otkljuca nije brana. Postoji zato sto se u headless runu
' administracija ne moze iskljuciti (MozeAdministraciju je anti-lockout: bez
' AUTH-a svi su admini), pa bi tvrdnja "sidebar postuje branu panela" inace
' merila dva puta True i prolazila i kad brane nema. Preseljen iz
' modScrMatSistem zajedno sa branom koju meri.
Private mBranaZatvorenaTest As Boolean

'-------------------------------------------------------------- REGISTAR
' Red: "KLJUC|modul|graditelj|naslov(katalog)|legacy Tag".
'
' Graditelj prima JEDAN argument -- okvir domacina. To je isti potpis koji su te
' procedure vec imale (frm As Object), pa se telo panela nije menjalo: menja se
' samo KO je domacin.
Public Function PanelRedovi() As Variant
    ' Kljuc je ISTI kao red u registru ekrana (modUiScreens), jer sidebar bira
    ' po njemu: MAT_PODESAVANJA / MAT_ADMIN. Jedan kljuc, dva registra -- ime
    ' koje se razidje znaci stavku koja se ne otvara.
    '
    ' PETO polje je LEGACY TAG iz starog menija (modMaticniLookups.MaticniSekcije).
    ' Nije prikaz nego SPONA: tvrdnja "novi UI dostize sve sto stari meni
    ' dostize" (test 174) za sifarnike ide kroz MatKljucIzLegacyTag, a za ova
    ' dva panela iskljucivo odavde. Do v6-ui-200 ju je drzao modScrMatSistem, u
    ' spisku alatki; kad je taj ekran uklonjen, spona je morala negde -- a
    ' registar panela je jedino mesto koje oba kraja vec zna.
    PanelRedovi = Array( _
        "MAT_PODESAVANJA|modPodesavanja|BuildConfigEditor|OTKUI_MS_PODESAVANJA|" & _
            "Pode" & ChrW(353) & "avanja", _
        "MAT_ADMIN|modAdmin|BuildAdminPanel|OTKUI_MS_ADMIN|Admin")
End Function

' Svi kljucevi registra, redom. Javno zbog testa: registar ekrana i registar
' panela moraju da nose ISTE kljuceve, a to se ne moze tvrditi bez spiska.
Public Function PanelKljucevi() As Variant
    Dim r As Variant, a() As Variant, i As Long
    r = PanelRedovi()
    ReDim a(LBound(r) To UBound(r))
    For i = LBound(r) To UBound(r)
        a(i) = Split(CStr(r(i)), "|")(PAN_KLJUC)
    Next i
    PanelKljucevi = a
End Function

' Kljuc panela za stavku starog menija, ili "" ako ta stavka nije panel.
' Ogledalo modMaticniIzvor.MatKljucIzLegacyTag, za druga dva Tag-a.
Public Function PanelKljucIzLegacyTag(ByVal tag As String) As String
    Dim r As Variant, p() As String
    For Each r In PanelRedovi()
        p = Split(CStr(r), "|")
        If UBound(p) >= PAN_LTAG Then
            If StrComp(p(PAN_LTAG), tag, vbTextCompare) = 0 Then
                PanelKljucIzLegacyTag = p(PAN_KLJUC)
                Exit Function
            End If
        End If
    Next r
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

' Sme li trenutni operater da otvori ovaj panel. Sidebar to pita PRE crtanja,
' pa prigusena stavka kaze istinu -- do v6-ui-201 je stajala puna, a otvaranje
' je odbijalo tek posle klika.
Public Function PanelDozvoljen(ByVal kljuc As String) As Boolean
    On Error Resume Next
    If mBranaZatvorenaTest Then Exit Function
    If Not PanelPostoji(kljuc) Then Exit Function
    PanelDozvoljen = modAuth.MozeAdministraciju()
    Err.Clear
End Function

' Zatvara branu za jednu tvrdnju. Otvaranje ide iskljucivo kroz False, koji
' vraca normalno ponasanje -- ne postoji vrednost koja branu zaobilazi.
Public Sub PanelBranaZatvoriTest(ByVal zatvori As Boolean)
    mBranaZatvorenaTest = zatvori
End Sub

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
'
' vratiEkran = True (podrazumevano) znaci: posle sklanjanja panela ekran ispod
' se PONOVO CITA. To nije kozmetika. Otvaranje panela je taj ekran deaktiviralo
' (modOtkupUI.ActivateScreen -> ScrDeaktiviraj), pa mu je obrisano stanje --
' kod maticnih i mZonaEkran, koju sve njegove radnje traze. Bez ponovnog
' citanja bi se mreza videla, a "Izmeni" bi tiho nista ne radila.
'
' False prosledjuje samo onaj ko SAM vraca ekran (prelazak na drugi ekran) ili
' onaj kome ekrana vise nema (rusenje ljuske) -- inace bi se ekran citao dvaput.
Public Sub PanelZatvori(Optional ByVal vratiEkran As Boolean = True)
    Dim host As Object, bioOtvoren As Boolean
    On Error Resume Next
    bioOtvoren = (Len(mAktivan) > 0)
    ZatvoriTiho
    Set host = modOtkupUI.PanelHost()
    If Not host Is Nothing Then IsprazniOkvir host
    modOtkupUI.PanelRezim False
    ' Samo ako je NESTO stvarno zatvoreno: PanelZatvori se zove i "za svaki
    ' slucaj", a citanje ekrana na prazno je cist trosak.
    If bioOtvoren And vratiEkran Then modOtkupUI.PanelVracenNaEkran
    Err.Clear
End Sub

' Zatvara panel SAMO ako je aktivan bas panel datog MODULA. Vraca True kad
' jeste zatvorio -- pozivalac po tome zna da li je bio u ljusci ili u legacy
' formi.
'
' ZASTO MODUL, A NE KLJUC: modul zna SVOJE ime i ono ne moze da se razidje sa
' registrom. Kljuc je STRANO ime -- kad su PODESAVANJA/ADMIN u v6-ui-201
' postali MAT_PODESAVANJA/MAT_ADMIN, poredjenja po kljucu u modPodesavanja i
' modAdmin su tiho prestala da vaze: "Nazad" je padao u legacy granu, radio
' Unload nad OKVIROM (ne formom), gutao gresku i ostavljao mrtav panel.
Public Function PanelZatvoriAko(ByVal modul As String) As Boolean
    If Len(mAktivan) = 0 Then Exit Function
    If StrComp(PanelPolje(mAktivan, PAN_MODUL), modul, vbTextCompare) <> 0 Then Exit Function
    PanelZatvori
    PanelZatvoriAko = True
End Function

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

' Ima li AKTIVAN panel nesacuvanih izmena. Neobavezno: panel koji tu proceduru
' ne implementira nema sta da izgubi, pa greska poziva znaci False, ne pad.
' Ime se izvodi iz imena modula, isti dogovor kao <Modul>_Release -- nov panel
' ne trazi red vise u registru, samo da postuje isti dogovor.
Public Function PanelImaNesacuvano() As Boolean
    Dim m As String, v As Variant
    If Len(mAktivan) = 0 Then Exit Function
    m = PanelPolje(mAktivan, PAN_MODUL)
    If Len(m) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    v = Application.Run(m & "." & Mid$(m, 4) & "_ImaNesacuvano")
    If Err.Number = 0 Then PanelImaNesacuvano = CBool(v)
    Err.Clear
End Function

' Smemo li da zatvorimo aktivan panel. Pita SAMO ako panel kaze da ima
' nesacuvano -- inace se nista ne prikazuje. Do v6-ui-201 se otkucano u
' Podesavanjima gubilo bez reci na klik u sidebar, isto kao sto se gubio unos
' maticnog editora pre nego sto je i on poceo da pita (MATU_ASK_ODBACI_UNOS).
Public Function PanelSmemoDaZatvorimo() As Boolean
    PanelSmemoDaZatvorimo = True
    If Not PanelImaNesacuvano() Then Exit Function
    PanelSmemoDaZatvorimo = (MsgBox(Poruka("UIPAN_ASK_ODBACI"), _
                                    vbExclamation + vbYesNo + vbDefaultButton2, _
                                    APP_NAME) = vbYes)
End Function

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
