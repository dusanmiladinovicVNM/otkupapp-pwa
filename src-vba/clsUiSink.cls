VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUiSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' clsUiSink -- genericki WithEvents omotac za RUNTIME kontrole formi.
'
' ZASTO POSTOJI: nove "Private WithEvents ..." deklaracije u SAMOJ FORMI lome
' self-update code-merge te forme (AddFromString nad formom kojoj se DODAJU
' event-sink deklaracije = klasa kvara koja je oborila update 2.16.1 ->
' 2.21.0; vidi docs/SELF_UPDATE.md, zamke #3/#7/#11). Zato WithEvents zivi
' OVDE: klasa je .cls pa sme kroz fazu 2 (Remove+Import) koja MSForms
' deklaracije podnosi, a deklaracioni blok forme ostaje stabilan (samo
' obicne reference bez event interfejsa).
'
' KORISCENJE (u formi-domacinu):
'   Private m_uiSinks As Object                    ' tag -> clsUiSink
'   ' posle Controls.Add:
'   WireSink m_btnX, "m_btnX"                      ' helper u formi
'   ' jedan Public dispatcher u formi:
'   Public Sub UiSinkEvent(tagName, ev, arg)       ' Select Case -> handleri
'
' Isti obrazac kao clsBlokUI / clsWheelList / clsAdminBtn (WithEvents u
' klasi, ne u formi). Podrzani eventi: btn.Click, lst.Change, lst.Click,
' lst.DblClick, txt.Change, cmb.Change, tgl.Change -- prosiri ovde (i u
' dispatcheru forme) po potrebi.
' NIKAD ne vracaj WithEvents u deklaracije formi.
' ============================================================

Public frm As Object               ' forma-domacin; mora imati Public Sub UiSinkEvent
Public SinkTag As String           ' logicko ime kontrole (npr. "m_btnScClose")

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents lst As MSForms.ListBox
Attribute lst.VB_VarHelpID = -1
Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents cmb As MSForms.ComboBox
Attribute cmb.VB_VarHelpID = -1
Public WithEvents tgl As MSForms.ToggleButton
Attribute tgl.VB_VarHelpID = -1

' Zakaci sink na kontrolu po tipu kontrole. Poziva forma (WireSink helper).
Public Sub Bind(ByVal hostForm As Object, ByVal ctl As Object, ByVal tagName As String)
    Set frm = hostForm
    SinkTag = tagName
    Select Case TypeName(ctl)
        Case "CommandButton": Set btn = ctl
        Case "ListBox":       Set lst = ctl
        Case "TextBox":       Set txt = ctl
        Case "ComboBox":      Set cmb = ctl
        Case "ToggleButton":  Set tgl = ctl
    End Select
End Sub

Private Sub btn_Click()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Click", Nothing
End Sub

' Hover: koordinate se NE prosledjuju (postojeci handleri ih ignorisu -> zovu
' ResetActionButtons/ButtonHover). Dispatcher u formi zove handler sa dummy 0,0,0,0.
Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "MouseMove", Nothing
End Sub

Private Sub lst_Change()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Change", Nothing
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "DblClick", Cancel
End Sub

Private Sub lst_Click()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Click", Nothing
End Sub

Private Sub txt_Change()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Change", Nothing
End Sub

Private Sub cmb_Change()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Change", Nothing
End Sub

Private Sub tgl_Change()
    If Not frm Is Nothing Then frm.UiSinkEvent SinkTag, "Change", Nothing
End Sub

' Prekini SVE reference (kontrola + forma) -> raskini krug forma<->sink pre nego
' sto self-update edituje VBProject. Forma zove ReleaseUiSinks (QueryClose/
' Terminate); Class_Terminate je mreza za slucaj da instanca ode van skopa.
' NB: ime NIJE "Release" - to je rezervisan clan IUnknown-a (svaki VBA class ga
' nasledjuje) pa bi dalo "Member already exists ... from which this object module
' derives". Zato ReleaseSink.
Public Sub ReleaseSink()
    Set btn = Nothing
    Set lst = Nothing
    Set txt = Nothing
    Set cmb = Nothing
    Set tgl = Nothing
    Set frm = Nothing
    SinkTag = vbNullString
End Sub

Private Sub Class_Terminate()
    ReleaseSink
End Sub
