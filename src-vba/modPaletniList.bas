Attribute VB_Name = "modPaletniList"
Option Explicit

' ============================================================
' modPaletniList — paletni list sveze robe + prerada (Phase 2)
'
' Inkrement 1: numeracija po godini (1..n, reset svake godine).
'   Pattern preuzet iz modFaktura.GenerateBrojFakture (per-year max+1),
'   NE iz modBrojevi (cija je kanon x/ddmmyy po stanici/danu).
'
' Inkrement 2 (sledece): OnPrijemnicaSaved (240-gajbica raspodela + rubna
'   paleta preko 2 prijemnice/zbirne), PrintPaletniList (reuse modPrint),
'   GetKooperantiZaPaletu (reuse modSledljivost.TraceByZbirna), SavePreradniList.
'
' Reuse: GetTableData / RequireColumnIndex / LogErr (postojeci helperi).
' Sema tabela: modSetup.EnsurePaletniListSchema (pokrenuti jednom).
' ============================================================

' Vraca sledeci redni broj palete za TEKUCU godinu (1 ako jos nema palete
' u ovoj godini). Prikaz na listu: BrojPalete & "/" & Godina.
Public Function GenerateBrojPalete() As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PALETA)

    Dim yr As Long
    yr = Year(Date)

    Dim maxN As Long
    maxN = 0

    If Not IsEmpty(data) Then
        Dim iBroj As Long, iGod As Long
        iBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, "GenerateBrojPalete")
        iGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, "GenerateBrojPalete")

        Dim r As Long, n As Long
        For r = 1 To UBound(data, 1)
            If CLng(Val(CStr(data(r, iGod)))) = yr Then
                n = CLng(Val(CStr(data(r, iBroj))))
                If n > maxN Then maxN = n
            End If
        Next r
    End If

    GenerateBrojPalete = maxN + 1
    Exit Function

EH:
    LogErr "modPaletniList.GenerateBrojPalete"
    GenerateBrojPalete = 0
End Function
