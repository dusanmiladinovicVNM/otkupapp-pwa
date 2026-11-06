#!/usr/bin/env python3
"""Build complete modOtkupUI.bas from the pinned GitHub revision, read-only to repository.
No COM, no Excel access, no repository edits. Produces .bas and unified .diff.
"""
import argparse
import difflib
import hashlib
import pathlib
import subprocess
import sys

COMMIT = '66f7897fa5231786f7d3d3dd3e566922c8e80ed0'
PATH = 'src-vba/modOtkupUI.bas'
BLOB = 'dc81ce8aa69d84e517c4ba35867af1c9f2328ef7'

CLICK_OLD = '''    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        c = CLng(Mid$(tag, 3)) + 1
'''
CLICK_NEW = '''    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        ' MIG-ARCHIVE: ne dozvoliti naknadno sortiranje kartica sa
        ' kumulativnim saldom (i ostalih izvestaja sa datumom).
        ' Izvorni izvestaj vec vraca poslovni redosled dokumenata.
        If GridZadrziRedosledIzvestaja() Then
            mSortCol = 0
            mSortAsc = True
            RefreshSortGlyphs
            Exit Sub
        End If
        c = CLng(Mid$(tag, 3)) + 1
'''

HELP_OLD = '''Private Sub LoadGridFromScreen()
'''
HELP_NEW = '''' MIG-ARCHIVE: poslovni izvori izvestaja sortiraju detaljne redove po
' poslovnom broju dokumenta PRE racunanja running salda. UI sort po
' tekstu BROJ DOK ne sme da preslozi te redove nakon obracuna.
Private Function GridZadrziRedosledIzvestaja() As Boolean
    If StrComp(mScreen, "IZVESTAJI", vbTextCompare) <> 0 Then Exit Function
    If mColN < 1 Then Exit Function
    GridZadrziRedosledIzvestaja = _
        (StrComp(ColF(CStr(mCols(0)), 2), "date", vbTextCompare) = 0)
End Function

Private Sub LoadGridFromScreen()
'''

SORT_OLD = '''    If mViewN > 0 Then
        mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
    Else
        mView = tmp
    End If
'''
SORT_NEW = '''    If mViewN > 0 Then
        If GridZadrziRedosledIzvestaja() Then
            ' Ne sortirati vec obracunatu karticu: saldo svakog reda
            ' zavisi od izvornog dokumentnog redosleda.
            mView = tmp
            mSortCol = 0
            mSortAsc = True
        Else
            mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
        End If
    Else
        mView = tmp
    End If
'''


def make_patch(raw):
    original = raw.decode('utf-8-sig')
    if '\r\n' in original:
        newline = '\r\n'
    else:
        newline = '\n'
    text = original.replace('\r\n', '\n')
    for label, old, new in [('header click', CLICK_OLD, CLICK_NEW),
                            ('report order helper', HELP_OLD, HELP_NEW),
                            ('grid reload', SORT_OLD, SORT_NEW)]:
        count = text.count(old)
        if count != 1:
            raise RuntimeError(f'ABORT: {label} anchor count={count}, expected exactly 1; no .bas generated')
        text = text.replace(old, new, 1)
    if text == original.replace('\r\n','\n'):
        raise RuntimeError('No changes made')
    return original, text.replace('\n', newline)


def git_blob_sha(raw):
    return hashlib.sha1(b'blob ' + str(len(raw)).encode('ascii') + b'\0' + raw).hexdigest()


def main():
    ap = argparse.ArgumentParser(description='Generate complete, pinned .bas and Git diff; no repo mutation')
    ap.add_argument('--repo', required=True, help='Local Git repository, e.g. $HOME/otkupapp-pwa')
    ap.add_argument('--out-dir', required=True)
    args = ap.parse_args()
    repo = pathlib.Path(args.repo).expanduser().resolve()
    dest = pathlib.Path(args.out_dir).expanduser().resolve()
    proc = subprocess.run(['git','-C',str(repo),'show',f'{COMMIT}:{PATH}'],
        stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    if proc.returncode:
        sys.exit('ABORT: Cannot load pinned GitHub commit from local clone: ' + proc.stderr.decode(errors='replace'))
    raw = proc.stdout
    actual_sha = git_blob_sha(raw)
    if actual_sha != BLOB:
        sys.exit(f'ABORT: Git blob mismatch; got {actual_sha}, expected {BLOB}')
    before, after = make_patch(raw)
    delta = ''.join(difflib.unified_diff(
        before.splitlines(True), after.splitlines(True),
        fromfile=f'a/{PATH} ({COMMIT[:8]})', tofile='b/modOtkupUI.bas (MIG-ARCHIVE)', n=5))
    dest.mkdir(parents=True, exist_ok=True)
    fn = dest/'modOtkupUI.bas'
    diff = dest/'modOtkupUI.diff'
    note = dest/'SOURCE.txt'
    fn.write_bytes(after.encode('utf-8'))
    diff.write_text(delta, encoding='utf-8', newline='\n')
    note.write_text(f'Source: {COMMIT}\nPath: {PATH}\nOriginal Git blob SHA-1: {BLOB}\n'
      'Scope: UI date-report rows preserve the already-sorted business order; '+'\n'
      'no changes to any business tables or master repository.\n'
      'WARNING: Use only if modOtkupUI in archive matches pinned GitHub revision.\n',encoding='utf-8')
    assert fn.read_bytes() == after.encode('utf-8')
    print('COMPLETE_BAS_CREATED',fn)
    print('GIT_DIFF_CREATED',diff)
    print('SOURCE_MANIFEST',note)
    print('PINNED_BLOB',BLOB)
    print('NO_GIT_ORIGINAL_MODIFIED')

if __name__ == '__main__':
    main()
