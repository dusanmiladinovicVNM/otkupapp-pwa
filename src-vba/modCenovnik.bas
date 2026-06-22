Attribute VB_Name = "modCenovnik"
'Attribute VB_Name = "modCenovnik"
Option Explicit

' ============================================================
' modCenovnik – cenovnik (cena po proizvodu), append-only istorija
'
' Model:
'   - tblCenovnik je append-only: svaka promena cene = NOVI red.
'   - Stari redovi se ne brisu (vidi se kretanje cena kroz vreme).
'   - VAZECA cena = poslednji (najnoviji po Datum-u) ne-stornirani red
'     za kljuc (VrstaVoca + SortaVoca + Klasa).
'
' Koristi se u frmOtkup / frmDokumenta za auto-popunjavanje cene.
'
' Odnos prema postojecem modelu cene (NE duplirati):
'   - tblArtikli.CenaPoJedinici = JEDNA tekuca cena po artiklu (pesticidi/
'     djubrivo). Prefill je inline LookupValue u frmAgrohemija/modAgrohemija
'     (single-current model). Za to NE koristiti modCenovnik.
'   - tblCenovnik = append-only ISTORIJA cena za otkup voca (vrsta+sorta+
'     klasa), gde inline LookupValue ne moze (vraca prvi, ne poslednji red).
'   Zato postoje dva modela; ovaj modul pokriva samo otkup voca.
'
' VAZNO (PWA): kada krene rad na PWA / mobilnom otkupu, tblCenovnik
' MORA da se sinhronizuje na Google Sheets (modStammdatenSync), da bi
' otkupci na stanicama dobijali istu vazecu cenu. Vidi instructions/.
' ============================================================

' Vraca vazecu (poslednju) cenu za dati proizvod i klasu.
' 0 ako nema cene u cenovniku (ili tabela ne postoji).
Public Function GetVazecaCena(ByVal vrstaVoca As String, _
                              ByVal sortaVoca As String, _
                              ByVal klasa As String) As Double
    On Error GoTo EH

    GetVazecaCena = 0

    Dim data As Variant
    data = GetTableData(TBL_CENOVNIK)
    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_CENOVNIK)
    If IsEmpty(data) Then Exit Function

    Dim cv As Long, cS As Long, ck As Long, cc As Long, cD As Long
    cv = GetColumnIndex(TBL_CENOVNIK, COL_CEN_VRSTA)
    cS = GetColumnIndex(TBL_CENOVNIK, COL_CEN_SORTA)
    ck = GetColumnIndex(TBL_CENOVNIK, COL_CEN_KLASA)
    cc = GetColumnIndex(TBL_CENOVNIK, COL_CEN_CENA)
    cD = GetColumnIndex(TBL_CENOVNIK, COL_CEN_DATUM)

    If cv = 0 Or cS = 0 Or ck = 0 Or cc = 0 Then Exit Function

    Dim i As Long
    Dim bestRow As Long: bestRow = 0
    Dim bestDate As Double: bestDate = -1#

    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(nz(data(i, cv))), Trim$(vrstaVoca), vbTextCompare) = 0 _
           And StrComp(Trim$(nz(data(i, cS))), Trim$(sortaVoca), vbTextCompare) = 0 _
           And StrComp(Trim$(nz(data(i, ck))), Trim$(klasa), vbTextCompare) = 0 Then

            Dim dv As Double
            dv = 0
            If cD > 0 Then
                If IsDate(data(i, cD)) Then dv = CDbl(CDate(data(i, cD)))
            End If

            ' Najnoviji datum pobedjuje; kod istog datuma pobedjuje kasniji red.
            If dv >= bestDate Then
                bestDate = dv
                bestRow = i
            End If
        End If
    Next i

    If bestRow > 0 Then
        If IsNumeric(data(bestRow, cc)) Then GetVazecaCena = CDbl(data(bestRow, cc))
    End If

    Exit Function

EH:
    LogErr "modCenovnik.GetVazecaCena"
    GetVazecaCena = 0
End Function

' Dodaje novi red cene (append-only). Vraca CenaID ili "" pri gresci.
Public Function AddCena(ByVal datum As Date, _
                        ByVal vrstaVoca As String, _
                        ByVal sortaVoca As String, _
                        ByVal klasa As String, _
                        ByVal cena As Double) As String
    On Error GoTo EH

    If Len(Trim$(vrstaVoca)) = 0 Then
        Err.Raise vbObjectError + 7701, "modCenovnik.AddCena", "VrstaVoca je obavezna."
    End If
    If Len(Trim$(klasa)) = 0 Then
        Err.Raise vbObjectError + 7702, "modCenovnik.AddCena", "Klasa je obavezna."
    End If
    If cena <= 0 Then
        Err.Raise vbObjectError + 7703, "modCenovnik.AddCena", "Cena mora biti veca od nule."
    End If

    Dim newID As String
    newID = GetNextID(TBL_CENOVNIK, COL_CEN_ID, "CEN-")

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 7704, "modCenovnik.AddCena", "GetNextID nije vratio CenaID."
    End If

    ' Redosled mora pratiti tblCenovnik:
    ' CenaID, Datum, VrstaVoca, SortaVoca, Klasa, Cena, CreatedAt, Stornirano
    Dim rowData As Variant
    rowData = Array( _
        newID, _
        datum, _
        Trim$(vrstaVoca), _
        Trim$(sortaVoca), _
        Trim$(klasa), _
        cena, _
        Now, _
        "")

    If AppendRow(TBL_CENOVNIK, rowData) <= 0 Then
        Err.Raise vbObjectError + 7705, "modCenovnik.AddCena", "AppendRow nije uspeo za tblCenovnik."
    End If

    AddCena = newID
    Exit Function

EH:
    LogErr "modCenovnik.AddCena"
    AddCena = ""
End Function

