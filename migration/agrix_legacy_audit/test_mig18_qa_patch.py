"""Offline, *synthetic* tests only.  Actual workbook compile remains a human VBA gate."""
from pathlib import Path
import tempfile
from agrix_mig18_qa_kartica_patch import patch_report,patch_ui,only_once,build

report = '''Attribute VB_Name = "modIzvestaj"
Option Explicit
Public Function ReportKarticaAmbalaze(ByVal kooperantID As String, _
                                      ByVal datumOd As Date, ByVal datumDo As Date) As Variant
    If moves.count > 0 Then
        Dim tmpArr() As Variant
        ReDim tmpArr(1 To moves.count, 1 To 5)
        For i = 1 To moves.count
            Dim mv As Variant
            mv = moves(i)
            tmpArr(i, 1) = mv(0)
            tmpArr(i, 2) = mv(1)
            tmpArr(i, 3) = mv(2)
            tmpArr(i, 4) = mv(3)
            tmpArr(i, 5) = mv(4)
        Next i
        arr = SortArray(tmpArr, 1, True, 2)
    End If
    ReportKarticaAmbalaze = KarticaAmbRezultatSaPocetnim(arr, pocetniSaldo)
End Function
Private Function BuildOtkupBrojDokDict() As Object
End Function
'''
ui = '''Attribute VB_Name = "modOtkupUI"
Option Explicit
Private Sub PrimeniSortZaListu(ByVal kljuc As String)
    Dim spec As String, p() As String
    spec = modUiScreens.ScrSort(mScreen)
End Sub
Private Sub FilterClick(ByVal tag As String)
    If Left$(tag, 2) = "hd" And Len(tag) = 3 Then
        c = CLng(Mid$(tag, 3)) + 1
        Exit Sub
    End If
End Sub
Private Sub LoadGridFromScreen()
    If mViewN > 0 Then
        mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)
    Else
        mView = tmp
    End If
End Sub
'''

def k(d,doc,positive,idx):
    # Mirror of VBA expression's total-order policy for behavioral QA fixture.
    import re
    from datetime import date
    m=re.fullmatch(r'(.*)-([0-9]{1,10})',doc.strip())
    suffix=('0|'+m.group(1).upper()+'|'+m.group(2).zfill(10)) if m else ('1|'+doc.upper())
    return d.strftime('%Y%m%d')+'|'+suffix+'|'+('0' if positive else '1')+'|'+str(idx).zfill(10)

def test():
    out=patch_report(report)
    assert 'SortArray(tmpArr, 6, True)' in out
    assert 'AmbKarticaSortKey' in out
    assert 'ReDim tmpArr(1 To moves.count, 1 To 6)' in out
    assert 'Like String$(Len(seq), "#")' in out
    u=patch_ui(ui)
    assert 'If ActiveLista() = "AMBKARTICA" Then' in u
    assert 'mSortCol = 1: mSortAsc = True: mSortLista = kljuc' in u
    assert 'If ActiveLista() = "AMBKARTICA" Then Exit Sub' in u
    from datetime import date
    d=date(2026,6,8)
    a=[('1/080626-10',True,40),('1/080626-2',False,13),('1/080626-9',True,30),('1/080626-2',True,30)]
    a=sorted(enumerate(a),key=lambda p:k(d,p[1][0],p[1][1],p[0]+1))
    assert [p[1][0] for p in a]==['1/080626-2','1/080626-2','1/080626-9','1/080626-10']
    assert [p[1][1] for p in a[:2]]==[True,False]
    assert [sum((1 if x[1][1] else -1)*x[1][2] for x in a[:n]) for n in [1,2]]==[30,17]
    assert 'arr = SortArray(tmpArr, 6, True)' not in report
    # Anchor drift must fail closed.
    try: patch_report(report.replace('tmpArr(i, 5) = mv(4)','tmpArr(i, 5) = mv(99)'))
    except ValueError: pass
    else:raise AssertionError('mutated report anchor unexpectedly accepted')
    try: patch_ui(ui.replace('mView = SortedView(tmp, mViewN, mColN, mSortCol, mSortAsc)', 'mView = SortedView(tmp, mViewN, mColN, 1, mSortAsc)'))
    except ValueError: pass
    else:raise AssertionError('mutated UI anchor unexpectedly accepted')
    with tempfile.TemporaryDirectory() as td:
        root=Path(td)
        exported=root/'source';exported.mkdir()
        (exported/'modIzvestaj.bas').write_text(report)
        (exported/'modOtkupUI.bas').write_text(ui)
        created=root/'created'
        info=build(root, created, source_dir=exported)
        assert info['status']=='GENERATED_SOURCE_ONLY_NOT_IMPORTED'
        assert (created/'modIzvestaj.bas').exists()
        assert (created/'modOtkupUI.bas').exists()
        assert (created/'MIG18_QA_patch_manifest.json').exists()
        try:build(root,created,source_dir=exported)
        except FileExistsError:pass
        else:raise AssertionError('existing out must not overwrite')
    print('PASS: source-only module patch; numeric doc ordering; same-doc + before -; double-sort guard; fail-closed anchors; no overwrites')
if __name__=='__main__':test()
