"""Generise src-vba/modSchema.bas iz KANONSKE seme schema/schema.json.

Izvor istine je schema/schema.json u gitu -- ne sveska. Sveska je posledica.

    python tools/gen_schema_module.py            # regenerisi modSchema.bas
    python tools/gen_schema_module.py --check    # exit 2 ako nisu u koraku (CI)

Menjanje seme = izmena schema/schema.json, pa regeneracija. Nova kolona ide
NA KRAJ liste: modDataAccess.AppendRow pise POZICIONO (pisci poput
modOtkup.SaveOtkup grade goli Array(...)), pa kolona ubacena u sredinu tiho
pomera sve iza sebe.

BOOTSTRAP (jednokratno, vec obavljeno): kanon je izvucen iz donor sveske preko
tools/dump_schema.py --json. Od tada dump sluzi samo za inspekciju i poredjenje
(tools/schema_diff.py). Ako se kanon ikad mora ponovo zasejati iz sveske, to je
svesna operacija: --iz-dumpa <put.json>.

Izlaz je 100% ASCII sa CRLF prelomima (CLAUDE.md S3).
"""

import argparse
import collections
import io
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
KANON = os.path.join(ROOT, "schema", "schema.json")
IZLAZ = os.path.join(SRC_VBA, "modSchema.bas")
MODCONFIG = os.path.join(SRC_VBA, "modConfig.bas")

TBL_CONST = re.compile(r'^Public Const (TBL_\w+)\s+As String\s*=\s*"(\w+)"', re.M)


def tbl_konstante(path: str) -> dict:
    src = io.open(path, encoding="ascii", errors="replace").read()
    return {m.group(2): m.group(1) for m in TBL_CONST.finditer(src)}


def spec_ime(tbl: str) -> str:
    osnova = tbl[3:] if tbl.lower().startswith("tbl") else tbl
    return "Spec" + osnova[0].upper() + osnova[1:]


def otisak(tabele) -> str:
    """FNV-1a 32-bit nad "tbl|kol|kol;..." REDOM.

    Nad UREDJENIM kolonama, ne nad skupom: preraspored kolona je promena seme
    isto koliko i nedostajuca kolona, jer je upis pozicion.
    """
    delovi = []
    for t in tabele:
        delovi.append(t["table"] + "|" + "|".join(t["columns"]))
    tekst = ";".join(delovi)

    h = 0x811C9DC5
    for ch in tekst:
        h ^= ord(ch) & 0xFF
        h = (h * 0x01000193) & 0xFFFFFFFF
    return "%08X" % h


ZAGLAVLJE = '''Attribute VB_Name = "modSchema"

Option Explicit

'=====================================================================
' modSchema -- REGISTAR SEME TABELA
'
' Izvor istine za STRUKTURU tabela je schema/schema.json u gitu. Ovaj modul
' je GENERISAN ARTEFAKT te datoteke, a sveska je posledica -- ne izvor.
' Do PR1 je vazilo obrnuto: spiskovi kolona ziveli su iskljucivo u .xlsm, pa
' se prazna sveska nije mogla rekonstruisati, a obrisana kolona se videla tek
' kao pad upisa satima kasnije.
'
' GENERISAN FAJL -- ne menjaj rukom:
'     (izmeni schema/schema.json, pa:)
'     python tools/gen_schema_module.py
'
' REDOSLED KOLONA JE DEO SEME. modDataAccess.AppendRow pise POZICIONO, a
' pisci poput modOtkup.SaveOtkup grade goli Array(...) -- kolona ubacena u
' sredinu tiho pomera sve iza sebe u pogresne kolone. Zato se nove kolone
' dodaju NA KRAJ, a otisak (SchemaFingerprint) racuna nad UREDJENIM kolonama.
'
' Javni ulazi:
'   EnsureAllTables      kreira sto fali, dopunjava kolone. Idempotentno.
'   VerifySchema         prijavljuje odstupanja; NISTA ne menja.
'   SchemaReadyOrFail    tvrda kapija pred upis, nad zadatim tabelama.
'   SchemaCheckOnStart   jeftina provera pri pokretanju (otisak).
'   SchemaRegistry       registar, za testove i alate.
'
' NE zvati EnsureAllTables unutar otvorene transakcije: clsTransaction.RestoreTable
' pada na neslaganje broja kolona, pa bi promena seme onemogucila rollback.
'=====================================================================

' Vrste odstupanja koje VerifySchema vraca (prefiks stavke u kolekciji).
Public Const SCHEMA_DRIFT_TABELA As String = "TABELA"
Public Const SCHEMA_DRIFT_KOLONA As String = "KOLONA"
Public Const SCHEMA_DRIFT_REDOSLED As String = "REDOSLED"

' Otisak kanonske seme (FNV-1a 32 nad "tbl|kol|kol;..." REDOM). Generisan
' zajedno sa registrom -- ne menjati rukom.
Public Const SCHEMA_FINGERPRINT As String = "@@OTISAK@@"

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

' Kolone jedne tabele, REDOM iz kanona. Prazna kolekcija = tabela nije u
' registru (pozivalac razlikuje "nema je" od "nema kolona").
Public Function SchemaTableColumns(ByVal tblName As String) As Collection
    Dim out As Collection
    Dim delovi() As String
    Dim i As Long

    Set out = New Collection
    delovi = Split(RegKolone(tblName), "|")
    For i = LBound(delovi) To UBound(delovi)
        If Len(delovi(i)) > 0 Then out.Add delovi(i)
    Next i

    Set SchemaTableColumns = out
End Function

Public Function SchemaTableSheet(ByVal tblName As String) As String
    Dim reg As Object
    Dim v As String
    Dim p As Long

    Set reg = SchemaRegistry()
    If Not reg.Exists(tblName) Then Exit Function

    v = CStr(reg(tblName))
    p = InStr(1, v, vbTab, vbBinaryCompare)
    If p > 0 Then SchemaTableSheet = Left$(v, p - 1)
End Function

' Kreira tabele kojih nema i dopunjava kolone koje fale. Idempotentno.
' Pad JEDNE tabele se zapise i NE zaustavlja ostale -- isti razlog kao
' EnsureKolonaSaTragom u modSetup: blanket "On Error Resume Next" je cutke
' preskakao ostatak posla, pa se posledica videla tek kao pad upisa.
'
' NE popravlja REDOSLED: premestanje kolone u postojecoj tabeli bi pomerilo
' podatke. Pogresan redosled je nalaz za coveka, ne nesto sto se leci u prolazu.
Public Sub EnsureAllTables()
    Dim reg As Object
    Dim tblName As Variant

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        EnsureJednuTabelu CStr(tblName)
    Next tblName
End Sub

' Odstupanja sveske od kanona. NISTA ne menja -- dijagnostika ne sme da
' zameni gresku koju opisuje.
' Stavka je "VRSTA|tabela|detalj".
Public Function VerifySchema() As Collection
    Dim out As Collection
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim imena As Object
    Dim lc As ListColumn
    Dim delovi() As String
    Dim i As Long
    Dim ocekivano As String
    Dim neslaganje As String

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

            ocekivano = RegKolone(CStr(tblName))
            delovi = Split(ocekivano, "|")
            For i = LBound(delovi) To UBound(delovi)
                If Len(delovi(i)) > 0 Then
                    If Not imena.Exists(delovi(i)) Then
                        out.Add SCHEMA_DRIFT_KOLONA & "|" & CStr(tblName) & "|" & _
                                delovi(i)
                    End If
                End If
            Next i

            ' Redosled: kanon mora biti PREFIKS stvarnog zaglavlja, po INDEKSU
            ' KOLONE. Visak na kraju je dozvoljen (kolona koju kanon jos ne zna),
            ' ali svako razilazenje PRE kraja znaci da je pozicion upis promasen.
            neslaganje = PrefiksNeslaganje(lo, SchemaTableColumns(CStr(tblName)))
            If Len(neslaganje) > 0 Then
                out.Add SCHEMA_DRIFT_REDOSLED & "|" & CStr(tblName) & _
                        "|" & neslaganje
            End If
        End If
    Next tblName

    Set VerifySchema = out
End Function

' Jeftina provera pri pokretanju: otisak sveske vs otisak kanona.
' Vraca "" kad je sve u redu, inace kratak opis (pozivalac odlucuje sta s tim).
'
' Otisak se racuna nad UREDJENIM kolonama, pa hvata i nedostatak i preraspored.
' Cena je ~590 citanja .Name -- reda velicine 10 ms, zanemarljivo po startu.
' (EnsureAllTables je skup jer PISE; ovo samo cita.)
Public Function SchemaCheckOnStart() As String
    Dim stvarni As String
    Dim odstupanja As Collection
    Dim i As Long
    Dim n As Long
    Dim prikaz As String

    On Error GoTo EH

    stvarni = SchemaFingerprintActual()
    If StrComp(stvarni, SCHEMA_FINGERPRINT, vbBinaryCompare) = 0 Then Exit Function

    ' Otisak se razlikuje -- tek sada se placa pun prolaz, da bi poruka rekla STA.
    Set odstupanja = VerifySchema()
    If odstupanja.count = 0 Then
        ' Otisak i VerifySchema mere isto (kanonski prefiks), pa ovo znaci da
        ' im se implementacije razisle -- kvar u kodu, ne u svesci.
        SchemaCheckOnStart = "Otisak (" & stvarni & " vs " & SCHEMA_FINGERPRINT & _
                             ") se ne slaze sa VerifySchema, koja ne vidi nijedno " & _
                             "odstupanje. Dve provere su se razisle -- prijavi kao bug."
        Exit Function
    End If

    n = odstupanja.count
    If n > 5 Then n = 5
    For i = 1 To n
        If Len(prikaz) > 0 Then prikaz = prikaz & "; "
        prikaz = prikaz & CStr(odstupanja(i))
    Next i
    If odstupanja.count > 5 Then
        prikaz = prikaz & "; ... (+" & CStr(odstupanja.count - 5) & ")"
    End If

    SchemaCheckOnStart = "Sema odstupa od kanona (" & CStr(odstupanja.count) & _
                         "): " & prikaz
    Exit Function

EH:
    SchemaCheckOnStart = "Provera seme nije izvrsena: " & Err.description
End Function

' FNV-1a 32 nad "tbl|kol|kol;..." REDOM, po registru. Isti algoritam kao
' tools/gen_schema_module.py.otisak -- kad se jedan menja, mora i drugi.
'
' Meri se KANONSKI PREFIKS zaglavlja, ne celo zaglavlje. Dva razloga:
'
'   1. modSetup.EnsureRuntimeSchema dodaje kolone na svakom startu i one idu
'      NA KRAJ. One su legitimne (samo ih kanon jos ne drzi), pa bi otisak nad
'      celim zaglavljem prijavljivao trajan lazan drift -- a provera koju
'      operater nauci da ignorise ne stiti nista.
'   2. Pozicion upis (AppendRow) zavisi TACNO od prefiksa: sve iza kanonske
'      poslednje kolone ne moze da pomeri nijednu vrednost koju pisci salju.
'
' Time se otisak i VerifySchema poklapaju po konstrukciji: oba mere isto.
' Kolona koja FALI ili je PREMESTENA unutar prefiksa menja otisak; rep ne.
Public Function SchemaFingerprintActual() As String
    Dim reg As Object
    Dim tblName As Variant
    Dim lo As ListObject
    Dim delovi As String
    Dim red As String
    Dim koliko As Long
    Dim i As Long

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        Set lo = Nothing
        On Error Resume Next
        Set lo = modDataAccess.GetTable(CStr(tblName))
        On Error GoTo 0

        red = CStr(tblName)
        If Not lo Is Nothing Then
            koliko = UBound(Split(RegKolone(CStr(tblName)), "|"))
            If koliko > lo.ListColumns.count Then koliko = lo.ListColumns.count
            For i = 1 To koliko
                red = red & "|" & lo.ListColumns(i).name
            Next i
        End If

        If Len(delovi) > 0 Then delovi = delovi & ";"
        delovi = delovi & red
    Next tblName

    SchemaFingerprintActual = Fnv1a32(delovi)
End Function

' Tvrda kapija pred upis. tblList je "tblA|tblB|tblC" -- proverava se SAMO
' to, jer pun prolaz po svakom upisu je preskup.
'
' Proverava se i REDOSLED, ne samo prisustvo: AppendRow pise poziciono, pa
' tabela sa svim kolonama u pogresnom rasporedu tiho upisuje u pogresna polja.
' To je gore od pada upisa.
Public Sub SchemaReadyOrFail(ByVal sourceName As String, ByVal tblList As String)
    Dim reg As Object
    Dim delovi() As String
    Dim tblName As String
    Dim lo As ListObject
    Dim i As Long
    Dim neslaganje As String

    Set reg = SchemaRegistry()
    delovi = Split(tblList, "|")

    For i = LBound(delovi) To UBound(delovi)
        tblName = Trim$(delovi(i))
        If Len(tblName) > 0 Then
            If Not reg.Exists(tblName) Then
                Err.Raise vbObjectError + 9403, sourceName, _
                          "Tabela '" & tblName & "' nije u kanonskoj semi " & _
                          "(schema/schema.json). Dopuni kanon pa regenerisi."
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

            neslaganje = PrefiksNeslaganje(lo, SchemaTableColumns(tblName))
            If Len(neslaganje) > 0 Then
                Err.Raise vbObjectError + 9405, sourceName, _
                          "Tabela '" & tblName & "' ne odgovara kanonskoj semi (" & _
                          neslaganje & "). Upis je POZICION, pa bi vrednosti " & _
                          "otisle u pogresne kolone. Pokreni " & _
                          "modSchema.EnsureAllTables; ako i posle toga odstupa, " & _
                          "redosled se mora popraviti rucno."
            End If
        End If
    Next i
End Sub


'=====================================================================
' INTERNO
'=====================================================================

' FNV-1a 32: h = (h Xor bajt) * 16777619 mod 2^32.
'
' Double, ne Long: VBA Long je SA ZNAKOM, pa bi i medjurezultat i konacna
' vrednost preticali. Tri zamke koje su ovde vec placene:
'
'   1. Xor nad celom 32-bitnom vrednoscu ne treba. Bajt je < 256, pa menja
'      SAMO donjih 8 bita -- radi se Xor nad tim bajtom, u Long-u gde je
'      bezbedan. (Prva verzija je vrtela petlju po 32 bita: tacno, ali
'      ~320.000 iteracija za ovaj string.)
'   2. Mul32 mora da svede visu polovinu po modulu PRE mnozenja sa 65536,
'      inace medjurezultat predje 2^53 i Double pocne da gubi cifre.
'   3. Hex$ prima Long (sa znakom), pa preti preko 2^31. Ispis ide iz dve
'      16-bitne polovine.
Public Function Fnv1a32(ByVal s As String) As String
    Dim h As Double
    Dim i As Long
    Dim b As Long
    Dim nizak As Double
    Dim visa As Long
    Dim niza As Long

    h = 2166136261#

    For i = 1 To Len(s)
        b = Asc(Mid$(s, i, 1)) And &HFF&
        nizak = h - Int(h / 256#) * 256#
        h = h - nizak + CDbl(CLng(nizak) Xor b)
        h = Mul32(h, 16777619#)
    Next i

    visa = CLng(Int(h / 65536#))
    niza = CLng(h - CDbl(visa) * 65536#)
    Fnv1a32 = Right$("0000" & Hex$(visa), 4) & Right$("0000" & Hex$(niza), 4)
End Function

Private Function Mul32(ByVal a As Double, ByVal b As Double) As Double
    ' (a * b) mod 2^32. a se deli na dve 16-bitne polovine da nijedan
    ' medjurezultat ne predje 2^53 (granica tacnosti Double-a).
    Dim lo As Double
    Dim hi As Double

    hi = Int(a / 65536#)
    lo = a - hi * 65536#

    lo = Modulo32(lo * b)
    hi = Modulo32(Modulo32(hi * b) * 65536#)

    Mul32 = Modulo32(lo + hi)
End Function

Private Function Modulo32(ByVal v As Double) As Double
    Modulo32 = v - Int(v / 4294967296#) * 4294967296#
End Function

' Da li zaglavlje tabele pocinje TACNO kanonskim kolonama, po INDEKSU.
'
' Poredjenje po stringu ("|A|B" u "|A|BExtra") daje LAZAN prolaz: InStr vrati 1,
' a kolona B ne postoji. Ranjiva je bas poslednja kanonska kolona, jer iza nje
' nema delimitera. Zato se poredi kolona po kolona.
'
' Vraca "" kad je sve u redu, inace opis PRVOG neslaganja -- pozivalac odlucuje
' da li ga prijavljuje ili dize gresku.
'
' Jedan helper za VerifySchema i SchemaReadyOrFail: dve kapije ne smeju da
' razviju razlicite definicije "ispravnog prefiksa".
Private Function PrefiksNeslaganje(ByVal lo As ListObject, _
                                   ByVal kolone As Collection) As String
    Dim i As Long
    Dim stvarno As String

    If lo.ListColumns.count < kolone.count Then
        PrefiksNeslaganje = "tabela ima " & CStr(lo.ListColumns.count) & _
                            " kolona, kanon trazi " & CStr(kolone.count)
        Exit Function
    End If

    For i = 1 To kolone.count
        stvarno = lo.ListColumns(i).name
        If StrComp(stvarno, CStr(kolone(i)), vbTextCompare) <> 0 Then
            PrefiksNeslaganje = "pozicija " & CStr(i) & ": ocekivano '" & _
                                CStr(kolone(i)) & "', stvarno '" & stvarno & "'"
            Exit Function
        End If
    Next i
End Function

Private Sub EnsureJednuTabelu(ByVal tblName As String)
    Dim kolone As Collection
    Dim arr() As String
    Dim i As Long

    On Error GoTo EH

    Set kolone = SchemaTableColumns(tblName)
    If kolone.count = 0 Then
        Err.Raise vbObjectError + 9406, "modSchema.EnsureJednuTabelu", _
                  "Kanon nema nijednu kolonu za '" & tblName & "'."
    End If

    ReDim arr(1 To kolone.count)
    For i = 1 To kolone.count
        arr(i) = CStr(kolone(i))
    Next i

    modSetup.EnsureDataTable tblName, SchemaTableSheet(tblName), arr
    Exit Sub

EH:
    LogError "modSchema.EnsureAllTables", _
             "Tabela '" & tblName & "' nije obezbedjena: " & Err.description, _
             Err.Number
End Sub

' Ime NIJE "Reg": parametar svake Spec* procedure se zove "reg", a VBA je
' case-insensitive -- parametar bi zaklonio proceduru, pa bi "Reg reg, ..."
' postalo pozivanje Dictionary-ja sa cetiri argumenta (runtime 438). Ista klasa
' greske koju hvata vba_check pravilo ZAKLONJENO, ali ono namerno preskace
' objekte, pa je ovo nasao tek test.
'
' Registar cuva STRING, ne ugnjezden objekat: "sheet" & vbTab & "|kol|kol".
'
' Prva verzija je drzala Dictionary po tabeli sa Collection kolona. Ne radi:
' "Set d(\"kolone\") = obj" nad late-bound Scripting.Dictionary trazi Property Set,
' koji Dictionary ne izlaze -- runtime 438 ("Object doesn't support this property
' or method"). vba_check to ne moze da vidi; uhvatila su ga tek dva testa.
'
' Uz to je brze: otisak i poredjenje redosleda ionako rade nad spojenim stringom,
' pa se nista ne sklapa dvaput.
Private Sub RegistrujTabelu(ByVal reg As Object, ByVal tblName As String, _
                ByVal sheetName As String, ByVal kolone As Collection)
    Dim spojeno As String
    Dim i As Long

    If reg.Exists(tblName) Then
        Err.Raise vbObjectError + 9401, "modSchema.RegistrujTabelu", _
                  "Dupla tabela u registru: " & tblName
    End If

    For i = 1 To kolone.count
        spojeno = spojeno & "|" & CStr(kolone(i))
    Next i

    reg.Add tblName, sheetName & vbTab & spojeno
End Sub

' "|kol|kol" iz registra -- oblik koji otisak i poredjenje redosleda traze.
Private Function RegKolone(ByVal tblName As String) As String
    Dim reg As Object
    Dim v As String
    Dim p As Long

    Set reg = SchemaRegistry()
    If Not reg.Exists(tblName) Then Exit Function

    v = CStr(reg(tblName))
    p = InStr(1, v, vbTab, vbBinaryCompare)
    If p > 0 Then RegKolone = Mid$(v, p + 1)
End Function


'=====================================================================
' REGISTAR -- GENERISANO iz schema/schema.json, ne menjaj rukom
'=====================================================================

'''


def gen_registar(tabele) -> str:
    red = []
    red.append("Private Function BuildRegistry() As Object")
    red.append("    Dim reg As Object")
    red.append('    Set reg = CreateObject("Scripting.Dictionary")')
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
        red.append("Private Sub %s(ByVal reg As Object)" % spec_ime(t["table"]))
        red.append("    Dim k As Collection")
        red.append("    Set k = New Collection")
        for c in t["columns"]:
            red.append('    k.Add "%s"' % c)
        red.append('    RegistrujTabelu reg, %s, "%s", k' % (t["const"], t["sheet"]))
        red.append("End Sub")
        red.append("")

    return "\n".join(red)


def izgradi(kanon_path: str) -> tuple:
    d = json.load(io.open(kanon_path, encoding="utf-8"))
    tabele = d["tables"]

    mapa = tbl_konstante(MODCONFIG)
    greske = []
    for t in tabele:
        if t["table"] not in mapa:
            greske.append("  %s: nema TBL_ konstantu u modConfig.bas" % t["table"])
        elif mapa[t["table"]] != t["const"]:
            greske.append("  %s: kanon kaze %s, modConfig kaze %s"
                          % (t["table"], t["const"], mapa[t["table"]]))
        if not t["columns"]:
            greske.append("  %s: nema nijednu kolonu" % t["table"])
    if greske:
        return None, greske

    telo = (ZAGLAVLJE.replace("@@OTISAK@@", otisak(tabele))
            + gen_registar(tabele) + "\n")

    ne_ascii = sorted({c for c in telo if ord(c) > 127})
    if ne_ascii:
        return None, ["  izlaz nije ASCII: %r" % ne_ascii]

    return telo, []


def main(argv) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--check", action="store_true",
                    help="exit 2 ako modSchema.bas nije u koraku sa kanonom")
    ap.add_argument("--kanon", default=KANON)
    ap.add_argument("--out", default=IZLAZ)
    ap.add_argument("--iz-dumpa", metavar="JSON",
                    help="SVESNO ponovo zasej kanon iz ispisa sveske "
                         "(tools/dump_schema.py --json). Nije normalan tok.")
    a = ap.parse_args(argv[1:])

    if a.iz_dumpa:
        src = json.load(io.open(a.iz_dumpa, encoding="utf-8"))
        mapa = tbl_konstante(MODCONFIG)
        doc = collections.OrderedDict()
        stari = json.load(io.open(a.kanon, encoding="utf-8"))
        doc["_o_fajlu"] = stari["_o_fajlu"]
        doc["schemaVersion"] = stari.get("schemaVersion", 1) + 1
        doc["tables"] = [
            collections.OrderedDict([
                ("table", t["table"]),
                ("const", mapa.get(t["table"], "TBL_?")),
                ("sheet", t["sheet"]),
                ("columns", t["headers"]),
            ])
            for t in sorted(src["tables"], key=lambda x: x["table"].lower())
        ]
        with io.open(a.kanon, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(doc, fh, indent=2, ensure_ascii=True)
            fh.write("\n")
        print("Kanon ponovo zasejan iz %s -> %s" % (a.iz_dumpa, a.kanon))

    telo, greske = izgradi(a.kanon)
    if telo is None:
        print("GRESKA u kanonu (%s):" % a.kanon, file=sys.stderr)
        for g in greske:
            print(g, file=sys.stderr)
        return 2

    ocekivano = telo.replace("\n", "\r\n").encode("ascii")

    if a.check:
        if not os.path.exists(a.out):
            print("Ne postoji: %s -- pokreni bez --check" % a.out, file=sys.stderr)
            return 2
        if open(a.out, "rb").read() != ocekivano:
            print("%s nije u koraku sa %s -- regenerisi ga "
                  "(python tools/gen_schema_module.py)"
                  % (os.path.basename(a.out), os.path.basename(a.kanon)),
                  file=sys.stderr)
            return 2
        print("%s: u koraku sa kanonom (otisak %s)"
              % (os.path.basename(a.out), otisak(
                  json.load(io.open(a.kanon, encoding="utf-8"))["tables"])))
        return 0

    with open(a.out, "wb") as fh:
        fh.write(ocekivano)

    d = json.load(io.open(a.kanon, encoding="utf-8"))
    print("Upisano: %s" % a.out)
    print("  tabela: %d, kolona: %d, otisak: %s"
          % (len(d["tables"]),
             sum(len(t["columns"]) for t in d["tables"]),
             otisak(d["tables"])))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
