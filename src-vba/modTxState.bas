Attribute VB_Name = "modTxState"
'Attribute VB_Name = "modTxState"
' ============================================================
' modTxState - stanje "rollback nije bio potpun", jedina kopija.
'
' PISE ga samo clsTransaction.RollbackTx. CITA ga svaki put do UPISA i svaki
' put do SAVE-a.
'
' Zasto zaseban modul, a ne samo property na clsTransaction: lokalni tx objekat
' nestane kad pozivalac izadje iz procedure, pa je RollbackNepotpun na njemu
' signal ZA TOG POZIVAOCA -- ne granica bezbednosti. Bez ovog modula se sistem
' posle nepotpunog rollback-a vracao u PUNO operativan rezim: nova transakcija
' dozvoljena, AutoSaveAfterCommit zakazan, a ZatvoriAplikaciju radi
' Close SaveChanges:=True -- pa je operater mogao da zabetonira parcijalno
' vracen podatak samo zatvaranjem aplikacije, bez ijedne greske na ekranu.
' Prethodna implementacija je bila losa jer je ostavljala EnableEvents=False i
' transakciju aktivnom, ali je time sistem bio ocigledno polomljen; verzija bez
' ovog modula ga je vracala u ispravno stanje dok zna da podaci to nisu.
'
' Obrazac je isti kao modImportState, ali je POLITIKA obrnuta na dva mesta i to
' je namerno:
'
'   1. NEMA REGISTRA. Marker zivi samo u memoriji sesije. Recovery JE reload:
'      posto je Save zatvoren, na disku i dalje stoji stanje PRE transakcije.
'      Perzistiran marker bi svesku ucinio trajno nesnimljivom bez ijednog
'      nacina da operater to razresi iz aplikacije -- tacno steta koju
'      modImportState.ImportNijeDovrsen izricito odbija da napravi.
'   2. FAIL-CLOSED. Nema citanja koje moze da pukne (obicna Boolean
'      promenljiva), pa nema dileme fail-open/fail-closed kao kod registra:
'      dok marker stoji, i upis i snimanje su zatvoreni.
'
' Marker se NE brise iz produkcionog koda. Jedini izlaz je zatvaranje bez
' snimanja i ponovno otvaranje. Funkcija za ciscenje postoji samo u test-rezimu,
' jer bi javno "ocisti marker" bilo bas ona zaobilaznica koju kapija sprecava.
'
' Fajl mora ostati 100% ASCII.
' ============================================================
Option Explicit

' Postavlja SAMO clsTransaction.RollbackTx, i samo kad neka snapshot-ovana
' tabela nije vracena.
Private mKompromitovan As Boolean
Private mTabele As String
Private mRazlog As String

' Nepotpun rollback moze da se ponovi pre reload-a (drugi tok, druga
' transakcija), pa se imena tabela NADOVEZUJU -- prvi razlog se cuva jer je on
' bio prvi uzrok, a ne posledica rada koji je vec tekao preko lose osnove.
Public Sub OznaciNepotpunRollback(ByVal tabele As String, ByVal razlog As String)
    mKompromitovan = True
    If Len(mTabele) > 0 Then mTabele = mTabele & " | "
    mTabele = mTabele & tabele
    If Len(mRazlog) = 0 Then mRazlog = razlog
End Sub

' True dok ova sesija zna da podaci mogu biti nekonzistentni.
Public Function RollbackKompromitovan() As Boolean
    RollbackKompromitovan = mKompromitovan
End Function

Public Function NevraceneTabeleSesije() As String
    NevraceneTabeleSesije = mTabele
End Function

Public Function RazlogKompromisa() As String
    RazlogKompromisa = mRazlog
End Function

' Sme li ova sesija da pocne NOVU transakciju? Javno zbog testa: pogresan smer
' ove odluke ne pravi crven test nego upis PREKO podatka koji nije vracen.
' Isti obrazac kao modStanicaLock.BulkPushNaIzlasku.
Public Function UpisDozvoljen() As Boolean
    UpisDozvoljen = Not mKompromitovan
End Function

' Sme li sveska da se snimi? Danas isti odgovor, ali DRUGO pitanje -- citalac ne
' treba da zna da je to ista zastavica, da bi se jedno moglo promeniti bez
' tihog menjanja drugog.
Public Function SnimanjeDozvoljeno() As Boolean
    SnimanjeDozvoljeno = Not mKompromitovan
End Function

' Poruka o ishodu rollback-a, na JEDNOM mestu.
'
' EH putevi su tvrdili "promene vracene" bez obzira na ishod rollback-a. Posle
' nepotpunog rollback-a to je cinjenicno netacno: podaci su delimicno vraceni, a
' upis i snimanje su zakljucani -- operater bi iz te poruke zakljucio da moze da
' ponovi unos, pa bi tek sledeci upis ili Save saznao istinu.
'
' Cita se GLOBALNA brana, ne stanje jednog tx objekta, i to nije priblizno nego
' tacno: BeginTx je fail-closed dok brana stoji, pa nova transakcija ne moze ni
' da pocne. Marker zato moze biti postavljen samo transakcijom koja se upravo
' odmotava -- globalno i "ovaj tx" se ne mogu razici.
Public Function PorukaIshodaRollbacka(ByVal normalna As String) As String
    If Not mKompromitovan Then
        PorukaIshodaRollbacka = normalna
        Exit Function
    End If
    ' Rezerva se dodeljuje PRE citanja kataloga. Pozivna mesta su pod
    ' On Error Resume Next, pa bi greska u Poruka() ostavila PRAZAN string --
    ' a prazna poruka o gresci operateru izgleda kao da greske nema.
    PorukaIshodaRollbacka = "ROLLBACK NEPOTPUN -- upis i snimanje su zatvoreni. " & _
                            "Zatvorite aplikaciju BEZ snimanja i otvorite je " & _
                            "ponovo. Nevracene tabele: " & mTabele
    On Error Resume Next
    PorukaIshodaRollbacka = Poruka("APP_MSG_ROLLBACK_NEPOTPUN_NE_SNIMAM") & _
                            vbCrLf & mTabele
    On Error GoTo 0
End Function

' Test seam, tvrdo gejtovan -- isti obrazac kao
' modImportState.ImportPendingTestSet (.claude/rules/testovi.md S4). Van
' test-rezima ne radi nista, pa se kapija ne moze ugasiti spolja.
'
' SAMO reset, nema test-set-a: test koji meri ovu branu pravi PRAVI nepotpun
' rollback, pa bi postavljanje markera iz testa merilo zastavicu umesto puta.
Public Sub TxKompromisTestReset()
    If Not IsTestMode() Then Exit Sub
    mKompromitovan = False
    mTabele = ""
    mRazlog = ""
End Sub
