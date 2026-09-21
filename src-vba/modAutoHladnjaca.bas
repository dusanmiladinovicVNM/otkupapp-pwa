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
' S3d: LANAC SE VRACA U KORACIMA, nad kanonom. Prvi korak je otpremnica
' (AutoLanacHladnjaca ispod). Zbirna (S4) i prijemnica (S6) dodaju se u ISTU
' funkciju kad ti dokumenti predju na kanon -- ne u nov orkestrator.
'
' Stari lanac se NIJE prevodio: delio je dokument po klasi kroz SaveOtpremnica_TX
' i vezu Otkup.OtpremnicaID (oba obrisana u S3b-1).
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

' ============================================================
' AUTO-LANAC HLADNJACE -- korak otpremnice (A-014, S3d).
'
' Roba na hladnjackoj stanici se MERI PRI PRIJEMU u hladnjacu, pa otpremnica
' nosi tacno ono sto i otkupni list: kilaza i ambalaza su 1:1, bez odstupanja.
' Zato otpremnica nema sta da ceka -- nastaje i IZDAJE se odmah, iz tog jednog
' bloka, u jednoj transakciji (CreateOtpremnicaIzIzvora_TX izvodi ocekivanje iz
' izvora, pa je "povezano = ocekivano" zadovoljeno samim nastankom).
'
' Vozac je MIRROR stanice (VozacID = StanicaID). Lanac ga ne pravi usput: upis
' otkupa ne sme da pise maticne podatke. Nema mirrora -> lanac staje i kaze zasto.
'
' COVEKOV IZBOR JE JACI OD AUTOMATIKE: blok koji je operater vezao za svoj nacrt
' na radnom stolu ne dobija jos jednu otpremnicu. Pripadnost se cita iz KANONA
' (tblOtpremnicaIzvori), ne iz ekrana -- ista cinjenica bez obzira ko zove.
'
' Otkup je vec upisan svojom transakcijom. Ako lanac padne, blok OSTAJE -- ne
' brise se i ne stornira. Operater dobija razlog i blok vidi na radnom stolu
' (lista "Bez otpremnice"), pa ga veze rucno. Tiho preskakanje bi znacilo da
' misli da je lanac odradjen.
'
' Vraca OtpremnicaID ("" = nije pravljena); `izvestaj` je ono sto operater cita.
' ============================================================
Public Function AutoLanacHladnjaca(ByVal otkupID As String, _
                                   ByRef izvestaj As String) As String
    Dim stanicaID As String, kulturaID As String, tipAmb As String, broj As String
    Dim datum As Date, greska As String, errDesc As String
    Dim h As Object, izvori As Collection

    izvestaj = ""
    On Error GoTo EH

    otkupID = Trim$(otkupID)
    If Len(otkupID) = 0 Then Exit Function

    If StrComp(Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_STORNIRANO), "")), _
               "Da", vbTextCompare) = 0 Then Exit Function

    stanicaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_STANICA), ""))
    If Not IsHladnjacaStanica(stanicaID) Then Exit Function

    If Len(modDokumenta.OtpremnicaZaOtkup(otkupID)) > 0 Then Exit Function

    If Not modMalina.IsManagedStationMirror(stanicaID) Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_MIRROR") & " " & stanicaID
        Exit Function
    End If

    datum = CDate(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM))
    kulturaID = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_KULTURA), ""))
    tipAmb = Trim$(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_TIP_AMB), ""))

    ' Broj ide iz niza STANICE, kao i svaka druga otpremnica; mirror prefiks je
    ' pravilo broja ZBIRNE (modBrojevi.ApplyMirrorPrefix) i ovde se ne primenjuje.
    broj = modBrojevi.GenerateBrojOtpremnice(stanicaID, datum)
    If Len(broj) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & Poruka("OTKUI_ERR_LANAC_BROJ") & " " & stanicaID
        Exit Function
    End If

    Set h = CreateObject("Scripting.Dictionary")
    h("Datum") = datum
    h("StanicaID") = stanicaID
    h("VozacID") = stanicaID
    h("KulturaID") = kulturaID
    h("TipAmbalaze") = tipAmb
    h("BrojOtpremnice") = broj

    Set izvori = New Collection
    izvori.Add otkupID

    AutoLanacHladnjaca = modDokumenta.CreateOtpremnicaIzIzvora_TX(h, izvori, greska)
    If Len(AutoLanacHladnjaca) = 0 Then
        izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & greska
        Exit Function
    End If

    ' Zbirna i prijemnica se ovde dodaju u S4 i S6; do tada operater zna sta jos
    ' ceka, umesto da misli da je lanac gotov.
    izvestaj = Poruka("OTKUI_MSG_LANAC_OTP") & " " & broj & " " & Poruka("OTKUI_MSG_LANAC_CEKA")
    Exit Function

EH:
    ' Opis se cita PRE LogErr-a (LogErr usput brise stanje greske).
    errDesc = Err.description
    LogErr "modAutoHladnjaca.AutoLanacHladnjaca"
    izvestaj = Poruka("OTKUI_ERR_LANAC") & " " & errDesc
    AutoLanacHladnjaca = ""
End Function
