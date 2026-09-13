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
import contextlib
import copy
import io
import json
import os
import re
import sys
import tempfile

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
'   SchemaReadyOrFail    tvrda kapija pred upis, nad zadatim tabelama (redosled
'                        kolona I ugovor o formatu celije).
'   SchemaCheckOnStart   jeftina provera pri pokretanju (otisak).
'   PrimeniFormateKanona primeni ugovor o formatu i vrati sta NIJE leglo.
'   FormatNeslaganje     procitaj stvarni format jedne tabele; nista ne menja.
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
' ============================================================
' FORMAT CELIJE -- deo ugovora, ne kozmetika.
'
' Kolona u General formatu Excel TIHO konvertuje pri upisu: "3/2026" postane
' datum, "0641234567" izgubi vodecu nulu, 18-cifreni racun kroz Double izgubi
' poslednje cifre. Vrednost se menja u TRENUTKU UPISA, pa naknadno postavljanje
' formata NE popravlja vec pokvarene celije -- ono ih samo prikaze kao broj.
' Zato format mora da stoji PRE prvog upisa i da se tera na svaki start:
' reinstall, self-update i import vracaju kolone na General.
'
' Spisak je generisan iz kljuca "formats" u schema/schema.json. Ne dopunjavati
' ovde -- ovaj modul je artefakt.
' ============================================================
' Primeni ugovor i PROVERI da je legao. Vraca opis neslaganja koja su prezivela
' primenu ("" = sve u redu). Ne dize gresku -- odluku nosi pozivalac, jer se ovo
' zove i sa starta (gde fail-soft mora da ostane) i iz tvrde kapije pred upisom.
'
' Ranije je cela procedura bila Sub pod jednim "On Error Resume Next" i nije
' vracala nista. Ugovor je time prakticno glasio "pokusaj da postavis @, pa
' nastavi rad" -- a kolona koja nije "@" TIHO menja vrednost pri prvom upisu,
' dakle bas ono zbog cega ugovor postoji. Sada se rezultat MERI.
Public Function PrimeniFormateKanona() As String
    Dim reg As Object, tblName As Variant
    Dim kolone As Object, kolName As Variant
    Dim lo As ListObject
    Dim lose As String

    Set reg = FormatRegistry()
    For Each tblName In reg.keys
        Set lo = Nothing
        On Error Resume Next
        Set lo = modDataAccess.GetTable(CStr(tblName))
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set kolone = reg(tblName)
            For Each kolName In kolone.keys
                If Not PostaviFormatKolone(lo, CStr(kolName), CStr(kolone(kolName))) Then
                    If Len(lose) > 0 Then lose = lose & ", "
                    lose = lose & CStr(tblName) & "." & CStr(kolName)
                End If
            Next kolName
        End If
    Next tblName

    If Len(lose) > 0 Then PrimeniFormateKanona = "format nije legao: " & lose
End Function

' Format se postavlja na CELU kolonu tabele (ListColumn.Range), ne na
' DataBodyRange. Prazna tabela nema DataBodyRange, pa bi izlazak na njemu ostavio
' sveze napravljenu svesku bez formata -- i PRVI upisan red bi bio pokvaren.
' Tacno to se desilo 12.09.2026. sa sveskom napravljenom iz kanona.
'
' Vraca True samo ako je format STVARNO na koloni. Upis koji "nije pukao" nije
' dokaz: zasticen list, spojene celije ili drugi COM klijent mogu da ga odbiju bez
' VBA greske, a mesan format nad opsegom vrati Null (pa CStr pukne pod Resume Next
' i ostane prazno -- sto se ovde racuna kao neslaganje, i tako treba).
Private Function PostaviFormatKolone(ByVal lo As ListObject, ByVal kolName As String, _
                                     ByVal semanticki As String) As Boolean
    Dim col As ListColumn
    Dim fmt As String
    Dim stvarni As String

    fmt = ExcelFormat(semanticki)
    ' Nepoznat semanticki naziv obara GENERISANJE (gen_schema_module), pa ovde ne
    ' moze da stigne; ako ipak stigne, nije neslaganje formata nego nema sta da se
    ' postavi.
    If Len(fmt) = 0 Then PostaviFormatKolone = True: Exit Function

    Set col = Nothing
    On Error Resume Next
    Set col = lo.ListColumns(kolName)
    On Error GoTo 0
    ' Kolone nema u zatecenoj svesci -- to prijavljuje sema, ne ugovor o formatu.
    ' Nepostojeca kolona ne moze da pokvari nijednu vrednost.
    If col Is Nothing Then PostaviFormatKolone = True: Exit Function

    On Error Resume Next
    col.Range.NumberFormat = fmt
    stvarni = CStr(col.Range.NumberFormat)
    On Error GoTo 0

    PostaviFormatKolone = (StrComp(stvarni, fmt, vbTextCompare) = 0)
End Function

' Neslaganje ugovora za JEDNU tabelu -- cita STVARNI NumberFormat sa kolone, ne
' pamti sta je primena pokusala. "" = tabela je u skladu.
'
' Postoji odvojeno od PrimeniFormateKanona jer tvrda kapija pred upisom mora da
' ume da PITA, bez sporednog efekta i bez prolaza kroz svih 30 tabela.
Public Function FormatNeslaganje(ByVal tblName As String) As String
    Dim reg As Object, kolone As Object, kolName As Variant
    Dim lo As ListObject, col As ListColumn
    Dim fmt As String, stvarni As String, lose As String

    Set reg = FormatRegistry()
    If Not reg.Exists(tblName) Then Exit Function
    Set kolone = reg(tblName)

    Set lo = Nothing
    On Error Resume Next
    Set lo = modDataAccess.GetTable(tblName)
    On Error GoTo 0
    ' Nepostojecu tabelu prijavljuje SchemaReadyOrFail svojom porukom -- ovde bi
    ' druga poruka o istom stanju samo zbunila.
    If lo Is Nothing Then Exit Function

    For Each kolName In kolone.keys
        fmt = ExcelFormat(CStr(kolone(kolName)))
        If Len(fmt) > 0 Then
            Set col = Nothing
            On Error Resume Next
            Set col = lo.ListColumns(CStr(kolName))
            On Error GoTo 0
            If Not col Is Nothing Then
                stvarni = ""
                On Error Resume Next
                stvarni = CStr(col.Range.NumberFormat)
                On Error GoTo 0
                If StrComp(stvarni, fmt, vbTextCompare) <> 0 Then
                    If Len(lose) > 0 Then lose = lose & ", "
                    lose = lose & tblName & "." & CStr(kolName) & " je '" & _
                           IIf(Len(stvarni) > 0, stvarni, "mesano") & "', ocekivano '" & fmt & "'"
                End If
            End If
        End If
    Next kolName

    FormatNeslaganje = lose
End Function

' Semanticko ime -> Excel format. Kanon nosi znacenje, ne sirov Excel string.
Private Function ExcelFormat(ByVal semanticki As String) As String
    Select Case LCase$(semanticki)
        Case "text":     ExcelFormat = "@"
        Case "decimal2": ExcelFormat = "0.00"
        Case Else:       ExcelFormat = ""
    End Select
End Function

Public Sub EnsureAllTables()
    Dim reg As Object
    Dim tblName As Variant

    Set reg = SchemaRegistry()

    For Each tblName In reg.keys
        EnsureJednuTabelu CStr(tblName)
    Next tblName

    ' Format ide ODMAH po pravljenju tabela, pre ijednog upisa.
    PrimeniFormateKanona
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

            ' FORMAT je deo ugovora koliko i redosled, i opasniji je od njega:
            ' pogresan redosled posalje vrednost u pogresnu kolonu i to se vidi,
            ' a kolona koja nije "@" TIHO promeni samu vrednost pri upisu
            ' ("3/2026" -> datum, vodeca nula otpadne) -- posle cega nijedna
            ' kasnija provera nema sa cim da uporedi.
            '
            ' Primena je fail-soft po dizajnu (start ne sme da se zakljuca, isti
            ' razlog kao za semu), pa DOKAZ mora da stoji ovde -- pred upisom,
            ' gde greska stvarno nastaje i gde je pozivalac ne guta nego prekida
            ' _TX. Prvo se POKUSAVA lecenje, kao i za kolone; tvrdo se staje tek
            ' ako format ni posle primene nije legao.
            neslaganje = FormatNeslaganje(tblName)
            If Len(neslaganje) > 0 Then
                PrimeniFormateKanona
                neslaganje = FormatNeslaganje(tblName)
            End If
            If Len(neslaganje) > 0 Then
                Err.Raise vbObjectError + 9406, sourceName, _
                          "Ugovor o formatu celije nije ispunjen (" & neslaganje & _
                          "). Excel bi pri upisu TIHO promenio vrednost, pa je " & _
                          "upis odbijen. Pokreni modSchema.EnsureAllTables; ako i " & _
                          "posle toga odstupa, list je verovatno zasticen ili je " & _
                          "kolona rucno preformatirana."
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


def gen_formati(formats) -> str:
    """Generisi FormatRegistry() iz kljuca "formats" u kanonu."""
    red = []
    red.append("' Ugovor o formatu celije, generisan iz schema/schema.json -> \"formats\".")
    red.append("Private Function FormatRegistry() As Object")
    red.append("    Dim reg As Object, k As Object")
    red.append('    Set reg = CreateObject("Scripting.Dictionary")')
    red.append("    reg.CompareMode = vbTextCompare")
    for t in sorted(formats):
        red.append("")
        red.append('    Set k = CreateObject("Scripting.Dictionary")')
        red.append("    k.CompareMode = vbTextCompare")
        for c in sorted(formats[t]):
            red.append('    k("%s") = "%s"' % (c, formats[t][c]))
        red.append('    Set reg("%s") = k' % t)
    red.append("")
    red.append("    Set FormatRegistry = reg")
    red.append("End Function")
    red.append("")
    red.append("")
    return "\n".join(red)


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
    formats = d.get("formats", {})

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
    # Ugovor o formatu se proverava naspram kolona: tipfeler mora da umre pri
    # generisanju, ne u runtime-u nad tudjom sveskom. Format nad nepostojecom
    # kolonom je tise od greske -- nikad se ne primeni, a spisak izgleda pokriven.
    #
    # OBLIK se proverava prvi. Bez toga "formats": "text" ili lista umesto mape
    # dize goli TypeError iz dubine generatora -- poruka koja ne imenuje ni kljuc
    # ni tabelu, pa covek trazi gresku u alatu umesto u kanonu.
    kol_po_tabeli = {t["table"]: set(t["columns"]) for t in tabele}
    if not isinstance(formats, dict):
        return None, ['  formats: mora biti mapa {tabela: {kolona: format}}, '
                      'a jeste %s' % type(formats).__name__]
    for t in sorted(formats):
        if not isinstance(formats[t], dict):
            greske.append("  formats: %s mora biti mapa {kolona: format}, a jeste %s"
                          % (t, type(formats[t]).__name__))
            continue
        if t not in kol_po_tabeli:
            greske.append("  formats: tabele %s nema u kanonu" % t)
            continue
        for c in sorted(formats[t]):
            if c not in kol_po_tabeli[t]:
                greske.append("  formats: %s nema kolonu %s" % (t, c))
            if formats[t][c] not in ("text", "decimal2"):
                greske.append("  formats: %s.%s ima nepoznat format %r"
                              % (t, c, formats[t][c]))

    # UJEDNACENOST IMENA. Join je jak koliko i njegova slabija strana: dok su obe
    # kolone bile General, obe su se kvarile isto i poredjenje se poklapalo.
    # Ugovor nad samo jednom stranom pravi ASIMETRIJU -- tj. REGRESIJU u odnosu na
    # stanje pre ugovora. Mereno 13.09.2026: kljuc sifarnika tblKese.TipKese je
    # ostao van ugovora dok je FK tblPrerada.TipKese usao.
    #
    # Namerno neujednaceno ime mora stajati u "formatsIzuzeci", i to sa razlogom:
    # izuzetak bez obrazlozenja je spisak imena bez znacenja, tj. sledeca rupa.
    #
    # STA OVA KAPIJA NE VIDI, i ne pretvara se da vidi: vezu izmedju kolona
    # RAZLICITOG imena koje nose ISTU vrednost. tblStornoVeze.ParentBroj drzi isti
    # broj kao tblZbirna.BrojZbirne, a ime mu je jedinstveno u kanonu -- kad izadje
    # iz ugovora, nijedna druga tabela nema kolonu tog imena, pa kapija nema sta da
    # poredi i cuti. Tu asimetriju je nasao covek (recenzija 13.09.2026), ne alat.
    # Da bi je alat video, kanon bi morao da nosi imenovane vrednosne domene -- to
    # je zaseban posao, ne uzgredna dopuna ovog pravila.
    izuzeci = d.get("formatsIzuzeci", {})
    if not isinstance(izuzeci, dict):
        return None, ['  formatsIzuzeci: mora biti mapa {ime kolone: razlog}']
    for ime in sorted(izuzeci):
        if not str(izuzeci[ime]).strip():
            greske.append("  formatsIzuzeci: %s nema obrazlozenje" % ime)
    pod_ugovorom = {c for t in formats if isinstance(formats[t], dict)
                    for c in formats[t]}
    for ime in sorted(pod_ugovorom - set(izuzeci)):
        rupe = sorted(t["table"] for t in tabele
                      if ime in t["columns"] and ime not in formats.get(t["table"], {}))
        if rupe:
            nosi = sorted(t for t in formats if ime in formats[t])
            greske.append("  formats: '%s' je pod ugovorom u %s, a NIJE u %s "
                          "-- join po toj koloni bi bio asimetrican "
                          "(ili dodaj kolone, ili upisi ime u formatsIzuzeci uz razlog)"
                          % (ime, ", ".join(nosi), ", ".join(rupe)))

    # Imena idu DOSLOVNO u VBA string literal ("%s"), pa navodnik u imenu pravi
    # nezatvoren literal i modul koji se ne kompajlira -- a modul koji se ne
    # kompajlira obara CEO projekat, pa greska stigne kao "Cannot run the macro"
    # na bilo kom makrou. Simptom ne pokazuje na krivca; ova kapija pokazuje.
    for t in sorted(formats):
        if not isinstance(formats[t], dict):
            continue
        for ime in [t] + sorted(formats[t]):
            if '"' in ime or "\n" in ime or "\r" in ime:
                greske.append("  formats: ime %r sadrzi navodnik ili prelom reda "
                              "-- generisani VBA literal bi ostao nezatvoren" % ime)
    if greske:
        return None, greske

    telo = (ZAGLAVLJE.replace("@@OTISAK@@", otisak(tabele))
            + gen_formati(formats)
            + gen_registar(tabele) + "\n")

    ne_ascii = sorted({c for c in telo if ord(c) > 127})
    if ne_ascii:
        return None, ["  izlaz nije ASCII: %r" % ne_ascii]

    return telo, []


# ============================================================
# SELF-TEST
# ============================================================
# Do 13.09.2026. ovaj alat je bio JEDINA CI kapija bez self-testa, i
# .claude/rules/testovi.md je to imenovao kao poznatu rupu: "--check koji nikad
# nije pokazan crven ne dokazuje da poredi otisak".
#
# Rupa je porasla kad je ugovor o formatu celije doneo PET novih validacija nad
# kljucem "formats". Scenario koji se hvata: neko refaktorise izgradi() i otkaci
# jednu granu -- --check ostane ZELEN (kanon i modSchema.bas su i dalje medjusobno
# u koraku) i niko ne primeti da kapija vise ne meri.
#
# Slucajevi idu KROZ izgradi(), istu funkciju koju zove CLI, a jedan ide kroz ceo
# main() nad pravim fajlom na disku. Isti razlog je zapisan za vba_check:
# self-test koji zove proveru direktno dokazuje da funkcija radi, ali ne i da je
# CLI zove -- otkacen jedan red tada ostavlja i repo-run i self-test zelene.
#
# Kanon je SINTETICKI (tri prave tabele, minimalne kolone), ne pravi
# schema/schema.json: slucajevi ne smeju da padnu zato sto je neko legitimno
# promenio pravi kanon.

def _sinteticki_kanon() -> dict:
    """Najmanji kanon koji prolazi izgradi(): prave tabele, minimalne kolone."""
    return collections.OrderedDict([
        ("_o_fajlu", ["sinteticki kanon za self-test"]),
        ("schemaVersion", 1),
        ("tables", [
            collections.OrderedDict([
                ("table", "tblKese"), ("const", "TBL_KESE"), ("sheet", "Kese"),
                ("columns", ["TipKese", "TezinaKg", "Aktivan"]),
            ]),
            collections.OrderedDict([
                ("table", "tblKutije"), ("const", "TBL_KUTIJE"), ("sheet", "Kutije"),
                ("columns", ["TipKutije", "TezinaKg", "Aktivan"]),
            ]),
            collections.OrderedDict([
                ("table", "tblPrerada"), ("const", "TBL_PRERADA"), ("sheet", "Prerada"),
                ("columns", ["BrojPrerade", "TipKese", "TipKutije", "CreatedAt"]),
            ]),
        ]),
        ("formats", collections.OrderedDict([
            ("tblKese", collections.OrderedDict([("TipKese", "text")])),
            ("tblKutije", collections.OrderedDict([("TipKutije", "text")])),
            ("tblPrerada", collections.OrderedDict([
                ("BrojPrerade", "text"), ("TipKese", "text"),
                ("TipKutije", "text"), ("CreatedAt", "text"),
            ])),
        ])),
        ("formatsIzuzeci", collections.OrderedDict([
            ("CreatedAt", "u tblPrerada je tekst radi slucaja, drugde je datum"),
        ])),
    ])


def _bez_ujednacenosti(d):
    # TipKese postoji u DVE tabele; skini jednu stranu -> kapija mora da vikne
    del d["formats"]["tblKese"]["TipKese"]


def _izuzetak_gasi_pravilo(d):
    _bez_ujednacenosti(d)
    d["formatsIzuzeci"]["TipKese"] = "namerno, radi dokaza"


def _jedinstveno_ime(d):
    # IZMERENA GRANICA: BrojPrerade postoji samo u tblPrerada, pa kad izadje iz
    # ugovora ime vise nije pod ugovorom NIGDE i kapija nema sta da poredi.
    # Vrednosna veza sa drugom kolonom (drugo ime, ista vrednost) joj je
    # nevidljiva. Slucaj stoji da niko kasnije ne pripise kapiji ono sto ne radi.
    del d["formats"]["tblPrerada"]["BrojPrerade"]


def _izuzetak_bez_razloga(d):
    d["formatsIzuzeci"]["CreatedAt"] = "   "


def _navodnik_u_imenu(d):
    d["tables"][0]["columns"].append('Tip"Kese')
    d["formats"]["tblKese"]['Tip"Kese'] = "text"


def _formats_nije_mapa(d):
    d["formats"] = "text"


def _tabela_nije_mapa(d):
    d["formats"]["tblKese"] = ["TipKese"]


def _nepostojeca_kolona(d):
    d["formats"]["tblKese"]["NemaOve"] = "text"


def _nepoznat_format(d):
    d["formats"]["tblKese"]["TipKese"] = "tekst"


def _tabela_van_kanona(d):
    d["formats"]["tblNemaMe"] = collections.OrderedDict([("X", "text")])


def _const_se_ne_poklapa(d):
    d["tables"][0]["const"] = "TBL_NETACNO"


# (naziv, izmena kanona, deo poruke koji MORA da se pojavi; None = mora PROCI)
SEMA_CASES = [
    ("cist kanon",                     None,                    None),
    ("ujednacenost: TipKese izbacen",  _bez_ujednacenosti,      "'TipKese' je pod ugovorom"),
    ("izuzetak GASI pravilo",          _izuzetak_gasi_pravilo,  None),
    ("GRANICA: jedinstveno ime cuti",  _jedinstveno_ime,        None),
    ("izuzetak bez obrazlozenja",      _izuzetak_bez_razloga,   "nema obrazlozenje"),
    ("navodnik u imenu kolone",        _navodnik_u_imenu,       "nezatvoren"),
    ("formats nije mapa",              _formats_nije_mapa,      "mora biti mapa"),
    ("formats[tabela] nije mapa",      _tabela_nije_mapa,       "mora biti mapa {kolona"),
    ("format nad nepostojecom kolonom", _nepostojeca_kolona,    "nema kolonu NemaOve"),
    ("nepoznat semanticki format",     _nepoznat_format,        "nepoznat format"),
    ("tabela iz formats nije u kanonu", _tabela_van_kanona,     "tabele tblNemaMe nema u kanonu"),
    ("const se ne poklapa sa modConfig", _const_se_ne_poklapa,  "modConfig kaze"),
]


def _pusti_kanon(doc) -> tuple:
    fd, put = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    try:
        with io.open(put, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(doc, fh, indent=2, ensure_ascii=True)
        return izgradi(put)
    finally:
        os.unlink(put)


def self_test() -> int:
    palo = []
    for naziv, izmeni, mora in SEMA_CASES:
        doc = copy.deepcopy(_sinteticki_kanon())
        if izmeni is not None:
            izmeni(doc)
        telo, greske = _pusti_kanon(doc)
        tekst = "\n".join(greske)
        if mora is None:
            if telo is None or greske:
                palo.append("  %s: ocekivano da PRODJE, palo sa: %s" % (naziv, tekst[:200]))
        else:
            if telo is not None:
                palo.append("  %s: ocekivan PAD (%r), a generisanje je proslo" % (naziv, mora))
            elif mora not in tekst:
                palo.append("  %s: poruka ne sadrzi %r; dobijeno: %s" % (naziv, mora, tekst[:200]))

    # Jedan slucaj kroz CEO main(), nad pravim fajlovima na disku. Dokazuje da CLI
    # zaista zove izgradi() i da --check poredi bajtove izlaza, a ne nesto drugo.
    fd, kput = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    fd, oput = tempfile.mkstemp(suffix=".bas")
    os.close(fd)
    try:
        doc = copy.deepcopy(_sinteticki_kanon())
        with io.open(kput, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(doc, fh, indent=2, ensure_ascii=True)
        # main() svoje poruke pise na stdout/stderr; ovde su OCEKIVANE
        # (jedan poziv MORA da padne), pa bi u CI logu izgledale kao kvar.
        def tiho(argv):
            with contextlib.redirect_stdout(io.StringIO()), \
                 contextlib.redirect_stderr(io.StringIO()):
                return main(argv)

        rc = tiho(["gen", "--kanon", kput, "--out", oput])
        if rc != 0:
            palo.append("  main() nad cistim kanonom: rc=%d, ocekivano 0" % rc)
        elif tiho(["gen", "--kanon", kput, "--out", oput, "--check"]) != 0:
            palo.append("  main() --check odmah posle generisanja: nije 0")
        else:
            # otkaci izlaz -> --check MORA da padne
            with open(oput, "ab") as fh:
                fh.write(b"' drift\r\n")
            if tiho(["gen", "--kanon", kput, "--out", oput, "--check"]) == 0:
                palo.append("  main() --check nad IZMENJENIM izlazom vratio 0 "
                            "-- kapija ne poredi bajtove")
    finally:
        for p in (kput, oput):
            try:
                os.unlink(p)
            except OSError:
                pass

    if palo:
        print("gen_schema_module --self-test: PALO", file=sys.stderr)
        for p in palo:
            print(p, file=sys.stderr)
        return 2
    print("gen_schema_module --self-test: %d slucajeva + main(), cisto"
          % len(SEMA_CASES))
    return 0


def main(argv) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--check", action="store_true",
                    help="exit 2 ako modSchema.bas nije u koraku sa kanonom")
    ap.add_argument("--kanon", default=KANON)
    ap.add_argument("--out", default=IZLAZ)
    ap.add_argument("--iz-dumpa", metavar="JSON",
                    help="SVESNO ponovo zasej kanon iz ispisa sveske "
                         "(tools/dump_schema.py --json). Nije normalan tok.")
    ap.add_argument("--self-test", action="store_true",
                    help="exit 2 ako validacije nad kanonom ne grizu")
    a = ap.parse_args(argv[1:])

    if a.self_test:
        return self_test()

    if a.iz_dumpa:
        src = json.load(io.open(a.iz_dumpa, encoding="utf-8"))
        mapa = tbl_konstante(MODCONFIG)
        # KRECE OD ZATECENOG KANONA, ne od praznog dokumenta. Ranije je ovde
        # stajao doc = OrderedDict() u koji su se rucno prenosila TRI kljuca
        # (_o_fajlu, schemaVersion, tables) -- pa je svaki kljuc uveden posle toga
        # tiho nestajao pri reseed-u. Konkretno "formats": ceo ugovor o formatu bi
        # se izgubio, sve kapije bi ostale ZELENE (--check poredi kanon sa
        # modSchema.bas, a oba bi bila prazna; otisak se racuna samo nad
        # "tables"), i sveska iz kanona bi opet bila u General formatu.
        # Prolaz kroz zatecen dokument je ista stvar za tri kljuca, a ne moze da
        # izgubi cetvrti.
        doc = json.load(io.open(a.kanon, encoding="utf-8"),
                        object_pairs_hook=collections.OrderedDict)
        stari = doc
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
