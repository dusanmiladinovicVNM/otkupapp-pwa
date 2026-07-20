Attribute VB_Name = "modStornoWarm"
Option Explicit

' ============================================================
' modStornoWarm - idle pre-warm keza za Storno centar (browse).
'
' Problem: prvo otvaranje browse-a ("Nadji za storno") gradi ceo skup aktivnih
' dokumenata (GetActiveDocumentsForStorno -> ~7 svezih citanja velikih tabela) ->
' 7-10s od klika do ekrana.
'
' Resenje: taj skup se gradi UNAPRED, dok je program neaktivan (Application.OnTime
' idle timer, isti obrazac kao modJournaling deferred AutoSave). Debounce: 60s posle
' poslednje izmene podataka. Rezultat zivi u modulu (g_stornoDocs), pa forma pri
' otvaranju uzima gotov kes (GetWarmStornoDocs) -> otvaranje je trenutno.
'
' Osvezavanje: svaka TX (clsTransaction.CommitTx) zove InvalidateStornoWarm ->
' postavi dirty + prezakazi warm za +60s. Warm se gradi SAMO kad je dirty (ili
' prazan) -> nema bespotrebnog trosenja dok se podaci ne menjaju.
'
' VBA je jednonitni: OnTime se okida samo kad je Excel neaktivan (nije u editu
' celije, nema pokrenutog makroa, nema modalne forme). Zato warm ne prekida rad;
' odradi se dok korisnik ne dira aplikaciju. Ako korisnik otvori browse pre nego
' sto se warm zavrsi -> GetWarmStornoDocs gradi sinhrono (fallback, kao pre).
' ============================================================

Private Const MOD_NAME As String = "modStornoWarm"
Private Const WARM_IDLE_SECONDS As Long = 60         ' debounce posle poslednje izmene
Private Const WARM_PROC As String = "modStornoWarm.StornoWarmTick"

Private g_stornoDocs As Collection                   ' warm kes (ceo skup, svi tipovi)
Private m_stornoDirty As Boolean                     ' podaci promenjeni -> warm zastareo
Private m_warmScheduled As Boolean                   ' ima zakazan OnTime
Private m_nextWarmTime As Date                       ' tacno vreme (za Cancel)
Private m_warming As Boolean                         ' reentrancy guard

' ------------------------------------------------------------
' JAVNI API
' ------------------------------------------------------------

' Podaci su promenjeni (poziva se iz clsTransaction.CommitTx). Oznaci warm zastarelim
' i prezakazi ga za +60s od SADA (svaka nova izmena gura rok -> pravi debounce).
Public Sub InvalidateStornoWarm()
    On Error Resume Next
    m_stornoDirty = True
    ScheduleStornoWarm
End Sub

' Zakazi (ili prezakazi) idle warm za +WARM_IDLE_SECONDS. Bootstrap iz StartApp.
Public Sub ScheduleStornoWarm()
    On Error Resume Next
    CancelStornoWarm
    m_nextWarmTime = Now + TimeSerial(0, 0, WARM_IDLE_SECONDS)
    Application.OnTime m_nextWarmTime, WARM_PROC
    m_warmScheduled = True
End Sub

' Otkazi pending OnTime. MORA se pozvati pre gasenja (modMain.ShutdownApp) -
' pending OnTime posle zatvaranja moze da REOTVORI workbook (poznata zamka).
Public Sub StopStornoWarm()
    CancelStornoWarm
End Sub

' OnTime callback. MORA biti Public u standardnom modulu. Gradi warm SAMO ako je
' dirty ili prazan; inace nista (idle bez izmena = bez trosenja).
Public Sub StornoWarmTick()
    On Error GoTo EH
    m_warmScheduled = False
    If m_warming Then Exit Sub
    If Not (m_stornoDirty Or g_stornoDocs Is Nothing) Then Exit Sub
    BuildWarm
    Exit Sub
EH:
    LogErr MOD_NAME & ".StornoWarmTick"
End Sub

' Forma uzima gotov skup ovde: svez warm -> trenutno; zastareo/prazan -> sinhroni
' fallback (isto kao ranije, ali sada retko - samo ako korisnik pretekne warm).
Public Function GetWarmStornoDocs() As Collection
    On Error GoTo EH
    ' Kesu se veruje SAMO ako je svez I NEPRAZAN. Poznata zamka: warm build jednom
    ' ispadne prazan (tranzientno -- npr. prvi idle tick po startu) pa se g_stornoDocs
    ' zaglavi prazan (m_stornoDirty=False, nema TX da invalidira) -> panel ostaje prazan
    ' iako podaci postoje. Prazan kes zato NE koristimo -> rebuild.
    If (Not m_stornoDirty) And (Not g_stornoDocs Is Nothing) Then
        If g_stornoDocs.count > 0 Then
            Set GetWarmStornoDocs = g_stornoDocs
            Exit Function
        End If
    End If
    BuildWarm
    If Not g_stornoDocs Is Nothing Then
        If g_stornoDocs.count > 0 Then
            Set GetWarmStornoDocs = g_stornoDocs
            Exit Function
        End If
    End If
    ' Warm build ispao prazan/nedostupan: direktan (nekeziran) build kao mreza -> forma
    ' nikad ne ostane prazna kad podaci postoje (ako su podaci stvarno prazni, i ovo je prazno).
    Set GetWarmStornoDocs = GetActiveDocumentsForStorno("Svi", "")
    Exit Function
EH:
    LogErr MOD_NAME & ".GetWarmStornoDocs"
    ' Ultimativni fallback: gradi direktno (bez keza) da forma nikad ne ostane prazna.
    On Error Resume Next
    Set GetWarmStornoDocs = GetActiveDocumentsForStorno("Svi", "")
End Function

' ------------------------------------------------------------
' PRIVATNO
' ------------------------------------------------------------

' Sagradi ceo skup u jednom prolazu. Batch-cache (BeginTableCache) skuplja ponovljena
' citanja (OTPREMNICA se cita 2x u GetActiveDocumentsForStorno) u jedno.
Private Sub BuildWarm()
    On Error GoTo EH
    m_warming = True
    BeginTableCache
    Dim res As Collection
    Set res = GetActiveDocumentsForStorno("Svi", "")
    EndTableCache
    Set g_stornoDocs = res
    m_stornoDirty = False
    m_warming = False
    Exit Sub
EH:
    On Error Resume Next
    EndTableCache
    m_warming = False
    LogErr MOD_NAME & ".BuildWarm"
End Sub

Private Sub CancelStornoWarm()
    If Not m_warmScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime m_nextWarmTime, WARM_PROC, , False
    On Error GoTo 0
    m_warmScheduled = False
End Sub
