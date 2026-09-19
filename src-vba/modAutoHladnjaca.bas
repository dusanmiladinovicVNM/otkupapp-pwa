Attribute VB_Name = "modAutoHladnjaca"
Option Explicit

' ============================================================
' modAutoHladnjaca -- oznake hladnjace (stanica, kupac) i relink stanje.
'
' S3b-1: AUTO-LANAC JE OBRISAN. AutoChainHladnjaca je od Otkup cutover-a bio
' pauziran (modOtkupUnos ga ne zove), a pravio je lanac PO KLASI kroz stari
' pisac otpremnice (SaveOtpremnica_TX) i vezu Otkup.OtpremnicaID -- oba su
' obrisana. Sa njim su otisli i backfill prijemnica hladnjace (F-090, u mapi
' "nije sposobnost") i test seam ArmHladnjacaTestFail.
'
' Lanac se NE prevodi ovde. Hladnjaca u novom modelu ide istim putem kao svaki
' dokument: otpremnica iz izvora (CreateOtpremnicaIzIzvora_TX), zbirna iz
' izdatih otpremnica (S4). Da li se auto-lanac uopste vraca odlucuje S3d/S4
' (plan S14.14); operater do tada dobija glasnu poruku OTKUNOS_MSG_LANAC_PAUZIRAN.
'
' Ostaje samo ono sto drugi moduli i dalje zovu: IsHladnjacaStanica (upis
' otkupa, storno), IsHladnjacaKupac (storno framework) i relink stanje
' (modOtkupUnos, modScrDokumenti).
' ============================================================

' Ispravka autohladnjace: posle storna otkupa-hladnjace ceo lanac
' (otpremnica+zbirna+prijemnica) je oboren, a palete su OSIROCENE. Kad operater
' izabere "Uneti ispravku", ekran zapamti broj te (stornirane) prijemnice ovde.
' Relink aparat je NEDOSTIZAN dok auto-lanca nema (modOtkupUnos).
Private mPendingRelinkOldPrij As String

Public Sub SetHladnjacaRelinkPending(ByVal oldPrijBroj As String)
    mPendingRelinkOldPrij = Trim$(oldPrijBroj)
End Sub

Public Function GetHladnjacaRelinkPending() As String
    GetHladnjacaRelinkPending = mPendingRelinkOldPrij
End Function

' Da li je stanica oznacena kao hladnjaca (tblStanice.JeHladnjaca = "Da").
Public Function IsHladnjacaStanica(ByVal stanicaID As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = Trim$(nz(LookupValue(TBL_STANICE, "StanicaID", stanicaID, COL_STA_JE_HLADNJACA), ""))
    IsHladnjacaStanica = (StrComp(v, "Da", vbTextCompare) = 0)
End Function

' Da li je KUPAC oznacen kao hladnjaca-kupac (interni cold-store tok). Isti signal
' kao frmDokumenta.RefreshBrojPrijSuggestion: kupac == CFG_MALINA_DEFAULT_KUPAC.
' Eksterni kupci -> False (za njih je zbirna poslednji interni dokument, a prijemnica
' eksterna -> storno framework ne kaskadira nizvodni tok). Prazan config / prazan
' kupac -> False.
Public Function IsHladnjacaKupac(ByVal kupacID As String) As Boolean
    On Error Resume Next
    kupacID = Trim$(kupacID)
    If Len(kupacID) = 0 Then Exit Function
    Dim h As String
    h = Trim$(GetConfigValue(CFG_MALINA_DEFAULT_KUPAC))
    If Len(h) = 0 Then Exit Function
    IsHladnjacaKupac = (StrComp(kupacID, h, vbTextCompare) = 0)
End Function
