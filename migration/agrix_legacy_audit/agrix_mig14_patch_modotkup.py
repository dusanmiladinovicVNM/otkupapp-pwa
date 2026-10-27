#!/usr/bin/env python3
"""Produce a separate, minimally patched modOtkup.bas for historical cancelled zero-kg records.

Read-only source, no workbook, no Git writes; fails closed when source shape differs.
"""
import argparse
import difflib
import hashlib
import json
from pathlib import Path

MARKER = 'Private Function StorniraniOtkupiZaBulkRead(ByVal sourceName As String) As Object'
HELPER = r'''Private Function StorniraniOtkupiZaBulkRead(ByVal sourceName As String) As Object
    ' MIG-14: istorijski stornirani dokumenti ne smeju rusiti zajednicki citac.
    ' Ne menja validaciju novog otkupa niti jedne aktivne stavke.
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set StorniraniOtkupiZaBulkRead = result

    Dim zagl As Variant
    zagl = GetTableData(TBL_OTKUP)
    If Not IsArray(zagl) Then Exit Function

    Dim cID As Long, cStorno As Long
    cID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, sourceName)
    cStorno = RequireColumnIndex(TBL_OTKUP, COL_STORNIRANO, sourceName)

    Dim i As Long, oid As String
    For i = 1 To UBound(zagl, 1)
        If StrComp(Trim$(NzToText(zagl(i, cStorno))), "Da", vbTextCompare) = 0 Then
            oid = Trim$(NzToText(zagl(i, cID)))
            If Len(oid) > 0 Then result(oid) = True
        End If
    Next i
End Function

'''

def sha(b):
    return hashlib.sha256(b).hexdigest()

def one(s, old, new, label):
    n = s.count(old)
    if n != 1:
        raise ValueError(f'{label}: expected exactly one exact anchor, found {n}; source untouched')
    return s.replace(old, new, 1)

def patch(s):
    if MARKER in s:
        raise ValueError('MIG-14 helper already present: refuse a second patch')
    if 'Attribute VB_Name = "modOtkup"' not in s:
        raise ValueError('Not a VBA modOtkup module')
    s = one(s,
        '    RequireJedinstvenZaglavljeOtkupa zagl, SRC\n',
        '    RequireJedinstvenZaglavljeOtkupa zagl, SRC\n\n'
        '    Dim istorijskiStornirani As Object\n'
        '    Set istorijskiStornirani = StorniraniOtkupiZaBulkRead(SRC)\n',
        'bulk read cancelled dictionary')
    s = one(s,
        '            RequireStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _\n'
        '                                d(i, cAmb), i, SRC\n',
        '            RequireStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _\n'
        '                                d(i, cAmb), i, SRC, istorijskiStornirani.Exists(oid)\n',
        'bulk call guard')
    s = one(s,
        'Private Sub RequireJedinstvenZaglavljeOtkupa(ByVal zagl As Object, _',
        HELPER + 'Private Sub RequireJedinstvenZaglavljeOtkupa(ByVal zagl As Object, _',
        'cancelled header helper')
    s = one(s,
        '                                ByVal red As Long, ByVal sourceName As String)\n'
        '    If Len(oid) = 0 Then',
        '                                ByVal red As Long, ByVal sourceName As String, _\n'
        '                                Optional ByVal dozvoliNuluZaStorno As Boolean = False)\n'
        '    If Len(oid) = 0 Then',
        'validator opt-in arg')
    # Scoped to only this sub, never change an unrelated validator.
    a = s.index('Private Sub RequireStavkaUgovor(')
    b = s.index('\nEnd Sub', a) + len('\nEnd Sub')
    sub = s[a:b]
    sub = one(sub,
        '    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then\n',
        '    If CDbl(kol) < 0 Or CDbl(cena) <= 0 Or _\n'
        '       (CDbl(kol) = 0 And Not dozvoliNuluZaStorno) Then\n',
        'quantity and price validation')
    s = s[:a] + sub + s[b:]
    return s

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--source', required=True, type=Path, help='Existing local src-vba/modOtkup.bas')
    ap.add_argument('--out-dir', required=True, type=Path, help='NEW isolated output directory')
    args = ap.parse_args()
    if not args.source.is_file():
        ap.error('Source .bas does not exist')
    raw = args.source.read_bytes()
    try:
        original = raw.decode('ascii')
    except UnicodeDecodeError as e:
        raise ValueError('Source is not 100% ASCII; refuse to rewrite unexpected encoding') from e
    newline = '\r\n' if b'\r\n' in raw else '\n'
    original_lf = original.replace('\r\n', '\n')
    updated_lf = patch(original_lf)
    updated = updated_lf.replace('\n', newline).encode('ascii')
    out = args.out_dir
    if out.exists() and any(out.iterdir()):
        raise ValueError(f'Output directory is not empty: {out}')
    out.mkdir(parents=True, exist_ok=True)
    dest = out/'modOtkup.bas'
    dest.write_bytes(updated)
    diff = ''.join(difflib.unified_diff(original_lf.splitlines(keepends=True), updated_lf.splitlines(keepends=True), fromfile='local-modOtkup.bas', tofile='MIG14-modOtkup.bas'))
    (out/'modOtkup_compat.diff').write_text(diff, encoding='utf-8')
    summary = {
      'format':'AGRIx-MIG14-STORNO-ZERO-BULK-READER-PATCH-v1',
      'status':'GENERATED_NOT_IMPORTED',
      'source_sha256':sha(raw),
      'patched_sha256':sha(updated),
      'modified_workbook':False,
      'modified_source_or_git':False,
      'code_scope':'modOtkup.StavkeOtkupaRedovi and RequireStavkaUgovor, historical Stornirano=Da only',
      'writer_unchanged':True,
      'VrednostOtkupa_unchanged':True,
      'runtime_qa_pending':True,
    }
    (out/'mig14_patch_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('MIG-14 PATCH GENERATED. Source/Git/XLSM untouched.')
    print('PATCHED MODULE:', dest)
    print('DIFF:', out/'modOtkup_compat.diff')
    print('SHA SOURCE:', summary['source_sha256'])
    print('SHA PATCHED:', summary['patched_sha256'])
    print('NEXT: import ONLY modOtkup.bas into QA XLSM copy, compile, save/reopen, run smoke tests.')

if __name__ == '__main__':
    main()
