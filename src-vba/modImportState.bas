Attribute VB_Name = "modImportState"
'Attribute VB_Name = "modImportState"
' ============================================================
' modImportState - stanje "prekinut VBA import", jedina kopija.
'
' PISE ga samo modVbaTools (faza 2 importa). CITA ga svaki put do Save-a, jer
' sveska ne sme da se snimi dok je VBA projekat mozda NEPOTPUN.
'
' Zasto zaseban modul, a ne u modVbaTools gde marker i nastaje:
'   run_vba.py (SELF_MODULE) i ImportAllVBA NE uvoze modVbaTools. Produkcioni
'   modul koji bi ga zvao kvalifikovano ne bi mogao da se kompajlira u test
'   svesci -- a modul koji se ne kompajlira obara CEO projekat, pa suite ne
'   padne nego VISI (.claude/rules/vba-izvor.md S3). Izmereno: prva verzija ove
'   kapije je zvala modVbaTools.ImportNijeDovrsen direktno i RunAllTests je
'   visila 389s umesto da prijavi ijedan rezultat.
'
' Formula sekcije i ime registra zive OVDE, a modVbaTools ih koristi -- dve
' kopije istog kljuca su tacno ona vrsta divergencije zbog koje postoji
' tools/vba_parity_check.py.
'
' NAMERNO ne pokriva self-update: njegov marker je u drugom registru
' ("AgriXSelfUpdate", modSelfUpdate), a njegov SaveWorkbookVerified je
' LEGITIMAN snimac novog projekta. Kapija koja bi njega zaustavila polomila bi
' bas mehanizam koji svesku popravlja.
' ============================================================
Option Explicit

' Registar u kome modVbaTools drzi stanje faze 2.
Public Const IMPORT_REG_APP As String = "AgriXVbaTools"

' Test seam (vidi ImportPendingTestSet) - samo test-rezim.
Private mTestPending As Boolean

' Sekcija je scope-ovana po IMENU SVESKE: dve otvorene kopije ne dele stanje
' faze 2, pa prekinut import u jednoj ne blokira snimanje druge.
'
' POZNATO OGRANICENJE, nasledjeno od modVbaTools.P2Section: kljuc je samo
' sanitizovano ime fajla. "AgriX-DEV.xlsm" i "AgriX_DEV.xlsm" daju ISTI kljuc, a
' dve sveske istog imena iz dva foldera ga svakako dele. Dok je marker bio samo
' recovery bookkeeping to je bila kozmetika; sada je granica bezbednosti
' snimanja, pa je greska u OBA smera moguca -- lazna blokada tudje sveske i
' propusteno upozorenje nad svojom. Stabilniji identitet (ime + hash pune
' putanje) je zaseban posao, ne siri se ovde: promena kljuca bi ostavila zive
' markere prekinutih importa nevidljivim.
Public Function ImportSekcija() As String
    Dim s As String, i As Long, ch As String, out As String
    s = ThisWorkbook.name
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Then out = out & ch
    Next i
    ImportSekcija = "import_" & out
End Function

' True dok stoji marker prekinutog importa.
'
' Marker znaci "raniji prolaz je mozda ostavio projekat nepotpun". Postavlja ga
' BeginImportTransaction, a brise se na tacno tri mesta (modVbaTools:
' RecoverImportState), i nijedno nije "backup je uspeo".
'
' 12.09.2026: zastita je do tada bio KOMENTAR u modVbaTools -- "sledeci Save bi
' ga zabetonirao bez ijedne reci". Import je pukao usred pisanja modLogo
' (odsecen B64_SPLASH2), dijalog je rekao "NE SNIMAJ svesku", a
' AutoSaveAfterCommit je 60s kasnije snimio. Steta je time prezivela zatvaranje
' i ponovno otvaranje fajla; jedini izlaz je bio backup.
' FAIL-OPEN JE NAMERNA POLITIKA, ne slucajan On Error Resume Next.
' Ako citanje registra pukne, funkcija vraca False i sveska se sme snimiti.
' Obrnut izbor (fail-closed) bi kvar registra pretvorio u svesku koja se NIKAD
' vise ne moze snimiti, na svakoj masini gde se to desi -- a bez ijednog nacina
' da operater to razresi iz aplikacije. Steta koju ova kapija sprecava nastaje
' samo u uskom prozoru posle prekinutog importa; steta od fail-closed kvara
' pogadja normalan rad. Zato: propusti, ali ostavi trag.
Public Function ImportNijeDovrsen() As Boolean
    Dim v As String
    If mTestPending Then
        ImportNijeDovrsen = True
        Exit Function
    End If
    On Error GoTo EH
    v = GetSetting(IMPORT_REG_APP, ImportSekcija(), "pending", "")
    ImportNijeDovrsen = (v = "1")
    Exit Function
EH:
    ImportNijeDovrsen = False
    LogErr "modImportState.ImportNijeDovrsen"
End Function

' Test seam, tvrdo gejtovan -- isti obrazac kao modScrDokumenti.Scr_OtpTestSet
' (.claude/rules/testovi.md S4). Van test-rezima ne radi nista, pa se registar
' prave masine ne moze zaprljati iz testa, niti se kapija moze ugasiti spolja.
Public Sub ImportPendingTestSet(ByVal ukljuci As Boolean)
    If Not IsTestMode() Then Exit Sub
    mTestPending = ukljuci
End Sub
