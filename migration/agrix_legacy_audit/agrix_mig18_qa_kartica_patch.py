#!/usr/bin/env python3
"""Build two TEST-ONLY VBA replacement modules from the exact QA workbook code snapshot.

Reads old Git revision; never writes the repository or workbook.  Do not import
these modules into a workbook compiled from any other revision.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import subprocess
from pathlib import Path

BASE_REF = '66f7897fa5231786f7d3d3dd3e566922c8e80ed0'
FILES = ('modIzvestaj.bas', 'modOtkupUI.bas')


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def only_once(s: str, old: str, new: str, label: str) -> str:
    count = s.count(old)
    if count != 1:
        raise ValueError(f'{label}: expected ONE exact anchor, got {count}; nothing written')
    return s.replace(old, new, 1)


def patch_report(source: str) -> str:
    """Mutate only the report card function and add its own private sort-key helper."""
    marker = 'Public Function ReportKarticaAmbalaze('
    if source.count(marker) != 1:
        raise ValueError('ReportKarticaAmbalaze anchor missing/duplicated')
    start = source.index(marker)
    end = source.index('\nEnd Function', start) + len('\nEnd Function')
    before, fn, after = source[:start], source[start:end], source[end:]

    fn = only_once(fn, 'ReDim tmpArr(1 To moves.count, 1 To 5)',
                   'ReDim tmpArr(1 To moves.count, 1 To 6)', 'hidden sort key dimension')
    fn = only_once(fn,
                   '            tmpArr(i, 5) = mv(4)\n',
                   '            tmpArr(i, 5) = mv(4)\n'
                   '            tmpArr(i, 6) = AmbKarticaSortKey(CDate(mv(0)), CStr(mv(1)), CDbl(mv(3)), i)\n',
                   'hidden sort key per move')
    fn = only_once(fn, 'arr = SortArray(tmpArr, 1, True, 2)',
                   'arr = SortArray(tmpArr, 6, True)', 'sort by business chronology')

    source = before + fn + after
    helper = '''
' MIG-18 QA-only: karta se uredjuje PRE racuna running salda.
' Za poznate oblike BrDok x/ddmmyy-n, sekvenca n ide BROJCANO, ne kao tekst.
' Nepoznati oblici: stabilan prikaz po tekstu, NE tvrditi da je istorijska satnica.
' Unutar jednog otkupa pozitivno kretanje pre negativnog (vizuelni sled).
' Sesti, skriveni atribut je SAMO kljuc sortiranja; izlazni ugovor ostaje 6 kolona.
Private Function AmbKarticaSortKey(ByVal d As Date, ByVal broj As String, _
                                   ByVal ulaz As Double, ByVal ordinal As Long) As String
    Dim t As String, seq As String, prefix As String
    Dim marker As Long
    broj = Trim$(broj)
    marker = InStrRev(broj, "-")
    If marker > 1 And marker < Len(broj) Then
        seq = Mid$(broj, marker + 1)
        prefix = Left$(broj, marker - 1)
        If Len(seq) > 0 And Len(seq) <= 10 Then
            If seq Like String$(Len(seq), "#") Then
                t = "0|" & UCase$(prefix) & "|" & Format$(CDbl(seq), "0000000000")
            End If
        End If
    End If
    If Len(t) = 0 Then t = "1|" & UCase$(broj)
    AmbKarticaSortKey = Format$(Year(d), "0000") & _
                       Format$(Month(d), "00") & Format$(Day(d), "00") & _
                       "|" & t & "|" & IIf(ulaz > 0, "0", "1") & _
                       "|" & Format$(ordinal, "0000000000")
End Function

'''
    return only_once(source, '\nPrivate Function BuildOtkupBrojDokDict() As Object',
                     '\n' + helper + 'Private Function BuildOtkupBrojDokDict() As Object',
                     'insert private helper')


def patch_ui(source: str) -> str:
    # A sorted running-card is NOT a generic sortable grid. The result set is
    # already arranged and reconciled in the report layer, before saldo calc.
    anchor = '''    If mViewN > 0 Then
        mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
    Else
        mView = tmp
    End If'''
    replacement = '''    If ActiveLista() = "AMBKARTICA" Then
        ' Running saldo is calculated in the report's business order.
        ' NEVER sort the finalized card again in the generic grid.
        mView = tmp
    ElseIf mViewN > 0 Then
        mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
    Else
        mView = tmp
    End If'''
    s = only_once(source, anchor, replacement, 'LoadGridFromScreen bypass double-sort')

    # Ignore header click for a running card: sorting it would lie about saldo.
    anchor = '''    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        c = CLng(Mid$(tag, 3)) + 1'''
    replacement = '''    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        If ActiveLista() = "AMBKARTICA" Then Exit Sub
        c = CLng(Mid$(tag, 3)) + 1'''
    s = only_once(s, anchor, replacement, 'block column click on AMBKARTICA')

    # Display a truthful glyph: predefined business order is date ascending.
    anchor = '''Private Sub PrimeniSortZaListu(ByVal kljuc As String)
    Dim spec As String, p() As String
    spec = modUiScreens.ScrSort(mScreen)'''
    replacement = '''Private Sub PrimeniSortZaListu(ByVal kljuc As String)
    Dim spec As String, p() As String
    If kljuc = "AMBKARTICA" Then
        mSortCol = 1: mSortAsc = True: mSortLista = kljuc
        Exit Sub
    End If
    spec = modUiScreens.ScrSort(mScreen)'''
    s = only_once(s, anchor, replacement, 'initial card sort header')
    return s


def read_git(repo: Path, ref: str, path: str) -> bytes:
    p = subprocess.run(['git', '-C', str(repo), 'show', f'{ref}:src-vba/{path}'],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
    if p.returncode:
        raise RuntimeError(f'git show failed for {path}: {p.stderr.decode("utf-8", "replace").strip()}')
    return p.stdout


def get_source(repo: Path, ref: str, fname: str, source_dir: Path | None) -> bytes:
    if source_dir is None:
        return read_git(repo, ref, fname)
    return (source_dir / fname).read_bytes()


def build(repo: Path, out: Path, ref: str = BASE_REF, source_dir: Path | None = None) -> dict:
    if out.exists():
        raise FileExistsError(f'Output already exists: {out}; choose a NEW --out-dir')
    if ref != BASE_REF and source_dir is None:
        raise ValueError('QA package is anchored to 66f7897f; do not use another --ref')
    originals: dict[str, bytes] = {}
    patched: dict[str, bytes] = {}
    for name in FILES:
        raw = get_source(repo, ref, name, source_dir)
        # Git sources are UTF-8; source from VBE exported modules can be CP1252.
        try:
            content = raw.decode('utf-8-sig')
            encoding = 'utf-8'
        except UnicodeDecodeError:
            content = raw.decode('cp1252')
            encoding = 'cp1252'
        lf = content.replace('\r\n', '\n')
        if f'Attribute VB_Name = "{name[:-4]}"' not in lf[:400]:
            raise ValueError(f'{name}: VB_Name identity did not match; nothing written')
        updated = patch_report(lf) if name == 'modIzvestaj.bas' else patch_ui(lf)
        if lf == updated:
            raise ValueError(f'{name}: patch did not change anything')
        # CRLF is correct for VBA exported modules on Windows.  Import only on
        # a BACKUP QA copy of the workbook.  UTF-8 input remains UTF-8.
        patched[name] = updated.replace('\n', '\r\n').encode(encoding)
        originals[name] = raw

    out.mkdir(parents=True)
    for name, data in patched.items():
        (out / name).write_bytes(data)
    info = {
        'format':'AGRIx-MIG18-QA-CARD-BAS-PATCH-v1',
        'status':'GENERATED_SOURCE_ONLY_NOT_IMPORTED',
        'repository_untouched':True,
        'xlsm_untouched':True,
        'source_commit':ref if source_dir is None else 'exported_from_workbook',
        'target_workbook_basis':'66f7897f plus QA-local patches',
        'modules':{n:{'original_sha256':sha(originals[n]),'patched_sha256':sha(patched[n])} for n in FILES},
        'scope':['sort report by date/parsed document sequence/sign before running saldo',
                 'bypass generic grid sort for AMBKARTICA',
                 'block grid header re-sort for AMBKARTICA'],
        'not_in_scope':['source code repository','financial entries','AMB-10 ledger','document prints','movement type label'],
    }
    (out / 'MIG18_QA_patch_manifest.json').write_text(json.dumps(info,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    return info


def cli() -> None:
    p=argparse.ArgumentParser(description='Generate 2 QA-only VBA module replacements, no XLSM edits')
    p.add_argument('--repo', default=str(Path.home()/'otkupapp-pwa'))
    p.add_argument('--out-dir', required=True)
    p.add_argument('--ref', default=BASE_REF)
    p.add_argument('--source-dir', help='optional: source exported FROM actual QA workbook, for strict parity')
    a=p.parse_args()
    info=build(Path(a.repo),Path(a.out_dir),a.ref,Path(a.source_dir) if a.source_dir else None)
    print('MIG-18 QA PATCH FILES READY — no Excel or Git modified')
    for n, v in info['modules'].items():print(' ',n,v['patched_sha256'])
    print('Output:',Path(a.out_dir).resolve())

if __name__ == '__main__':cli()
