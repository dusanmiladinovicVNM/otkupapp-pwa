"""Generise src-vba/modSchema.bas iz ispisa seme (tools/dump_schema.py --json).

Zasto generator a ne rucno kucanje: registar je 590 kolona preko 41 tabele.
Rucno odrzavan bi se razisao sa sveskom prvog dana, a to je bas bolest koju
modSchema leci. Ovako je regeneracija jedan poziv:

    python tools/dump_schema.py <sveska> --json <put.json>
    python tools/gen_schema_module.py --json <put.json>

Prepisuje CEO modSchema.bas. Rucne izmene idu u API sekciju ovog generatora
(SABLON_* konstante dole), ne u izlazni .bas -- inace ih sledeca regeneracija
pojede.

Izlaz je 100% ASCII sa CRLF prelomima (CLAUDE.md S3).

Imena tabela se preslikavaju na TBL_ konstante iz modConfig.bas. Tabela u
svesci koja nema konstantu je greska generatora, ne upozorenje: registar mora
biti 1:1 sa konstantama da bi vba_check pravilo SEMA_REGISTAR imalo smisla.
"""

import argparse
import io
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
IZLAZ = os.path.join(SRC_VBA, "modSchema.bas")
MODCONFIG = os.path.join(SRC_VBA, "modConfig.bas")

TBL_CONST = re.compile(r'^Public Const (TBL_\w+)\s+As String\s*=\s*"(\w+)"', re.M)


def tbl_konstante(path: str) -> dict:
    """tblIme -> TBL_KONSTANTA iz modConfig.bas."""
    src = io.open(path, encoding="ascii", errors="replace").read()
    return {m.group(2): m.group(1) for m in TBL_CONST.finditer(src)}


def spec_ime(tbl: str) -> str:
    """tblOtkupStavke -> SpecOtkupStavke (VBA ime procedure)."""
    osnova = tbl[3:] if tbl.lower().startswith("tbl") else tbl
    return "Spec" + osnova[0].upper() + osnova[1:]


ZAGLAVLJE = '''Attribute VB_Name = "modSchema"

Option Explicit

'=====================================================================
' modSchema -- REGISTAR SEME TABELA
'
' Od ovog modula je izvor istine za STRUKTURU tabela KOD, ne sveska.
' Do sada je vazilo obrnuto: spiskovi kolona osnovnih tabela ziveli su
' iskljucivo u .xlsm, pa se prazna sveska nije mogla rekonstruisati, a
' obrisana kolona se videla tek kao pad upisa satima kasnije.
'
' GENERISAN FAJL -- ne menjaj rukom:
'     python tools/dump_schema.py <sveska> --json <put.json>
'     python tools/gen_schema_module.py --json <put.json>
'
' Javni ulazi:
'   EnsureAllTables      kreira sto fali, dopunjava kolone. Idempotentno.
'   VerifySchema         prijavljuje odstupanja; NISTA ne menja.
'   SchemaReadyOrFail    tvrda kapija pred upis, nad zadatim tabelama.
'   SchemaRegistry       registar, za testove i alate.
'
' NE zvati unutar otvorene transakcije: clsTransaction.RestoreTable pada
' na neslaganje broja kolona, pa bi promena seme onemogucila rollback.
'
' EnsureAllTables se NE zove sa svakog starta -- prolaz kroz sve tabele i
' sve kolone je skup. Zove se iz setup-a i iz testova. Na startu ide samo
' VerifySchema (citanje), kroz health check.
'=====================================================================

' Vrste odstupanja koje VerifySchema vraca (prefiks stavke u koleciji).
Public Const SCHEMA_DRIFT_TABELA As String = "TABELA"
Public Const SCHEMA_DRIFT_KOLONA As String = "KOLONA"

' Kes registra. Registar je DEKLARACIJA, ne snimak sveske, pa se ne menja
' u toku rada -- kesiranje je bezbedno.
Private mReg As Object


'=====================================================================
' JAVNI API
'=====================================================================

Public Function SchemaRegistry() As Object
    If mReg Is Nothing Then Set mReg = BuildRegistry()
    Set SchemaRegistry = mReg
End Function

' Kolone jedne tabele, redosledom iz registra. Prazna kolekcija = tabela
' nije u registru (pozivalac razlikuje "nema je" od "nema kolona").
Public Function SchemaTableColumns(ByVal tblName As String) As Collection
    Dim reg As Object
    Set reg = SchemaRegistry()
    If reg.Exists(tblName) Then
        Set SchemaTableColumns = reg(tblName)("kolone")
    Else
        Set SchemaTableColumns = New Collection
    End If
End Function

Public Function SchemaTableSheet(ByVal tblName As String) As String
    Dim reg As Object
    Set reg = SchemaRegistry()
    If reg.Exists(tblName) Then SchemaTableSheet = CStr(reg(tblName)("sheet"))
End Function

' Kreira tabele kojih nema i dopunjava kolone koje fale. Idempotentno.
' Pad JEDNE tabele se zapise i NE zaustavlja ostale -- isti razlog kao
' EnsureKolonaSaTragom u modSetup: blanket "On Error Resume Next" je cutke
' preskakao ostatak posla, pa se posledica videla tek kao pad upisa.
Public Sub EnsureAllTables()
    Dim reg As Object
    Dim tblName As Variant

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        EnsureJednuTabelu CStr(tblName), reg(tblName)
    Next tblName
End Sub

' Odstupanja sveske od registra. NISTA ne menja -- dijagnostika ne sme da
' zameni gresku koju opisuje.
' Stavka je "VRSTA|tabela|kolona"; kolona je prazna kad fali cela tabela.
Public Function VerifySchema() As Collection
    Dim out As Collection
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim imena As Object
    Dim lc As ListColumn
    Dim kolone As Collection
    Dim i As Long

    Set out = New Collection
    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        Set lo = Nothing
        On Error Resume Next
        Set lo = modDataAccess.GetTable(CStr(tblName))
        On Error GoTo 0

        If lo Is Nothing Then
            out.Add SCHEMA_DRIFT_TABELA & "|" & CStr(tblName) & "|"
        Else
            Set imena = CreateObject("Scripting.Dictionary")
            imena.CompareMode = vbTextCompare
            For Each lc In lo.ListColumns
                If Not imena.Exists(lc.name) Then imena.Add lc.name, True
            Next lc

            Set kolone = reg(tblName)("kolone")
            For i = 1 To kolone.count
                If Not imena.Exists(CStr(kolone(i))) Then
                    out.Add SCHEMA_DRIFT_KOLONA & "|" & CStr(tblName) & "|" & _
                            CStr(kolone(i))
                End If
            Next i
        End If
    Next tblName

    Set VerifySchema = out
End Function

' Tvrda kapija pred upis. tblList je "tblA|tblB|tblC" -- proverava se SAMO
' to, jer pun prolaz po svakom upisu je preskup.
'
' Postoji zato sto nedostajuca tabela inace pukne tek na AppendRow-u, sa
' porukom koja o uzroku ne kaze nista. Sa novim modelom dokumenta to vise
' nije degradacija nego "model dokumenta ne postoji".
Public Sub SchemaReadyOrFail(ByVal sourceName As String, ByVal tblList As String)
    Dim reg As Object
    Dim delovi() As String
    Dim tblName As String
    Dim lo As ListObject
    Dim kolone As Collection
    Dim imena As Object
    Dim lc As ListColumn
    Dim i As Long
    Dim j As Long

    Set reg = SchemaRegistry()
    delovi = Split(tblList, "|")

    For i = LBound(delovi) To UBound(delovi)
        tblName = Trim$(delovi(i))
        If Len(tblName) > 0 Then
            If Not reg.Exists(tblName) Then
                Err.Raise vbObjectError + 9403, sourceName, _
                          "Tabela '" & tblName & "' nije u registru seme " & _
                          "(modSchema). Registar je izvor istine -- dopuni ga."
            End If

            Set lo = Nothing
            On Error Resume Next
            Set lo = modDataAccess.GetTable(tblName)
            On Error GoTo 0

            If lo Is Nothing Then
                Err.Raise vbObjectError + 9404, sourceName, _
                          "Tabela '" & tblName & "' ne postoji u svesci. " & _
                          "Pokreni modSchema.EnsureAllTables pa ponovi."
            End If

            Set imena = CreateObject("Scripting.Dictionary")
            imena.CompareMode = vbTextCompare
            For Each lc In lo.ListColumns
                If Not imena.Exists(lc.name) Then imena.Add lc.name, True
            Next lc

            Set kolone = reg(tblName)("kolone")
            For j = 1 To kolone.count
                If Not imena.Exists(CStr(kolone(j))) Then
                    Err.Raise vbObjectError + 9405, sourceName, _
                              "Tabeli '" & tblName & "' fali kolona '" & _
                              CStr(kolone(j)) & "'. Pokreni " & _
                              "modSchema.EnsureAllTables pa ponovi."
                End If
            Next j
        End If
    Next i
End Sub


'=====================================================================
' INTERNO
'=====================================================================

Private Sub EnsureJednuTabelu(ByVal tblName As String, ByVal spec As Object)
    Dim kolone As Collection
    Dim arr() As String
    Dim i As Long

    On Error GoTo EH

    Set kolone = spec("kolone")
    If kolone.count = 0 Then
        Err.Raise vbObjectError + 9406, "modSchema.EnsureJednuTabelu", _
                  "Registar nema nijednu kolonu za '" & tblName & "'."
    End If

    ReDim arr(1 To kolone.count)
    For i = 1 To kolone.count
        arr(i) = CStr(kolone(i))
    Next i

    modSetup.EnsureDataTable tblName, CStr(spec("sheet")), arr
    Exit Sub

EH:
    LogError "modSchema.EnsureAllTables", _
             "Tabela '" & tblName & "' nije obezbedjena: " & Err.description, _
             Err.Number
End Sub

Private Sub Reg(ByVal reg As Object, ByVal tblName As String, _
                ByVal sheetName As String, ByVal kolone As Collection)
    Dim d As Object

    If reg.Exists(tblName) Then
        Err.Raise vbObjectError + 9401, "modSchema.Reg", _
                  "Dupla tabela u registru: " & tblName
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("sheet") = sheetName
    Set d("kolone") = kolone
    reg.Add tblName, d
End Sub


'=====================================================================
' REGISTAR -- GENERISANO, ne menjaj rukom
'=====================================================================

'''


def gen_registar(tabele, mapa) -> str:
    red = []
    red.append("Private Function BuildRegistry() As Object")
    red.append("    Dim reg As Object")
    red.append("    Set reg = CreateObject(\"Scripting.Dictionary\")")
    red.append("    reg.CompareMode = vbTextCompare")
    red.append("")
    for t in tabele:
        red.append("    %s reg" % spec_ime(t["table"]))
    red.append("")
    red.append("    Set BuildRegistry = reg")
    red.append("End Function")
    red.append("")
    red.append("")

    for t in tabele:
        tbl, sheet, cols = t["table"], t["sheet"], t["headers"]
        red.append("Private Sub %s(ByVal reg As Object)" % spec_ime(tbl))
        red.append("    Dim k As Collection")
        red.append("    Set k = New Collection")
        for c in cols:
            red.append('    k.Add "%s"' % c)
        red.append('    Reg reg, %s, "%s", k' % (mapa[tbl], sheet))
        red.append("End Sub")
        red.append("")

    return "\n".join(red)


def main(argv) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--json", required=True, help="ispis tools/dump_schema.py --json")
    ap.add_argument("--out", default=IZLAZ)
    a = ap.parse_args(argv[1:])

    d = json.load(io.open(a.json, encoding="utf-8"))
    tabele = sorted(d["tables"], key=lambda x: x["table"].lower())
    mapa = tbl_konstante(MODCONFIG)

    bez_konstante = [t["table"] for t in tabele if t["table"] not in mapa]
    if bez_konstante:
        print("GRESKA: tabele u svesci bez TBL_ konstante u modConfig.bas:",
              file=sys.stderr)
        for t in bez_konstante:
            print("  " + t, file=sys.stderr)
        print("Dodaj konstante pa ponovi -- registar mora biti 1:1 sa njima.",
              file=sys.stderr)
        return 2

    prazne = [t["table"] for t in tabele if not t["headers"]]
    if prazne:
        print("GRESKA: tabele bez zaglavlja: " + ", ".join(prazne), file=sys.stderr)
        return 2

    telo = ZAGLAVLJE + gen_registar(tabele, mapa) + "\n"

    ne_ascii = [c for c in telo if ord(c) > 127]
    if ne_ascii:
        print("GRESKA: izlaz nije ASCII: %r" % sorted(set(ne_ascii)), file=sys.stderr)
        return 2

    # CRLF, binarno -- inace .bas postane LF i self-update ga vidi kao izmenjen
    with open(a.out, "wb") as fh:
        fh.write(telo.replace("\n", "\r\n").encode("ascii"))

    print("Upisano: %s" % a.out)
    print("  tabela: %d, kolona: %d"
          % (len(tabele), sum(len(t["headers"]) for t in tabele)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
