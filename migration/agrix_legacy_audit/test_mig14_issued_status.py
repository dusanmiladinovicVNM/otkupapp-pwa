import hashlib
import json
import tempfile
from pathlib import Path
import zipfile
from agrix_mig14_fix_issued_status import patch,sha,TABLES,examine_table_after,discover_tables,stringtable

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK='http://schemas.openxmlformats.org/package/2006/relationships'

def fixture(root,broken_status=False,source_mutation=False):
    root=Path(root);stage=root/'stage';stage.mkdir()
    spec=[]
    for name,idcol,source,staged,_,_ in TABLES:
        rows=[{idcol:('OTP' if name=='tblOtpremnica' else 'ZBR')+f'-{i:05d}','Stornirano': 'Da' if i==2 else ''} for i in range(1,4)]
        for filename in (source,staged):
            path=stage/filename
            path.write_text(''.join(json.dumps(x)+'\n' for x in rows),encoding='utf8')
        spec.append((name,idcol,source,staged,sha(stage/source),sha(stage/staged)))
    wb=f'''<workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="Otpremnica" sheetId="1" r:id="rId1"/><sheet name="Zbirna" sheetId="2" r:id="rId2"/></sheets></workbook>'''
    wrels=f'''<Relationships xmlns="{PK}"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="worksheets/sheet2.xml"/></Relationships>'''
    zname=root/'sample.xlsm'
    with zipfile.ZipFile(zname,'w') as z:
        z.writestr('xl/workbook.xml',wb)
        z.writestr('xl/_rels/workbook.xml.rels',wrels)
        for i,(name,idcol,*rest) in enumerate(spec,1):
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'''<Relationships xmlns="{PK}"><Relationship Id="rId9" Type="x" Target="../tables/table{i}.xml"/></Relationships>''')
            z.writestr(f'xl/tables/table{i}.xml',f'''<table xmlns="{NS}" name="{name}" displayName="{name}" ref="A1:C4" headerRowCount="1"><tableColumns count="3"><tableColumn id="1" name="{idcol}"/><tableColumn id="2" name="Stornirano"/><tableColumn id="3" name="IzdatoStatus"/></tableColumns></table>''')
            s=f'''<worksheet xmlns="{NS}"><dimension ref="A1:C4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{idcol}</t></is></c><c r="B1" t="inlineStr"><is><t>Stornirano</t></is></c><c r="C1" t="inlineStr"><is><t>IzdatoStatus</t></is></c></row>'''
            for j in range(2,5):
                key=('OTP' if name=='tblOtpremnica' else 'ZBR')+f'-{j-1:05d}'
                s+=f'''<row r="{j}"><c r="A{j}" t="inlineStr"><is><t>{key}</t></is></c><c r="B{j}" t="inlineStr"><is><t>{'Da' if j==3 else ''}</t></is></c>'''
                if j==2: s+=f'<c r="C{j}" s="10"/>'
                if j==3: s+=f'<c r="C{j}" s="10"></c>'
                if j==4 and (i==1 or broken_status):
                    s+=f'<c r="C{j}" t="inlineStr"><is><t>{"DRAFT" if broken_status else ""}</t></is></c>'
                s+='</row>'
            s+='</sheetData><tableParts count="1"/></worksheet>'
            z.writestr(f'xl/worksheets/sheet{i}.xml',s)
        z.writestr('xl/vbaProject.bin',b'fake-but-unchanged')
        z.writestr('docProps/core.xml',b'<untouched />')
    if source_mutation:(stage/TABLES[0][2]).open('a').write('{}\n')
    return zname,stage,tuple(spec)

def success():
    with tempfile.TemporaryDirectory() as td:
        x,s,defs=fixture(td)
        original=x.read_bytes();out=Path(td)/'out'
        patch(x,s,out,expected_sha=sha(x),defs=defs,n=3)
        assert x.read_bytes()==original
        final=out/'AgriX_MIG14_ISSUED_STATUS_PATCHED.xlsm'
        with zipfile.ZipFile(x) as a,zipfile.ZipFile(final) as b:
            assert a.namelist()==b.namelist()
            assert all(a.read(n)==b.read(n) for n in a.namelist() if not n.startswith('xl/worksheets/sheet'))
            found=discover_tables(b,defs);ss=stringtable(b)
            for d in defs:
                ids={('OTP' if d[0]=='tblOtpremnica' else 'ZBR')+f'-{i:05d}' for i in range(1,4)}
                assert examine_table_after(b,d,found,ss,ids,3)==3
            assert b'Da' in b.read('xl/worksheets/sheet1.xml')
            assert b'<c r="C4" t="inlineStr"><is><t>IZDATO</t></is></c>' in b.read('xl/worksheets/sheet2.xml')
        report=json.loads((out/'MIG14_issued_status_final_verify.json').read_text())
        assert report['pass'] and len(report['modified_parts'])==2

def blocked_draft():
    with tempfile.TemporaryDirectory() as td:
        x,s,defs=fixture(td,broken_status=True)
        try:patch(x,s,Path(td)/'out',expected_sha=sha(x),defs=defs,n=3)
        except ValueError:pass
        else:raise AssertionError('Must not overwrite existing draft status')
        assert not (Path(td)/'out').exists()

def blocked_source():
    with tempfile.TemporaryDirectory() as td:
        x,s,defs=fixture(td,source_mutation=True)
        try:patch(x,s,Path(td)/'out',expected_sha=sha(x),defs=defs,n=3)
        except ValueError:pass
        else:raise AssertionError('Mutated source must block')
        assert not (Path(td)/'out').exists()

def blocked_input():
    with tempfile.TemporaryDirectory() as td:
        x,s,defs=fixture(td)
        try:patch(x,s,Path(td)/'out',expected_sha='0'*64,defs=defs,n=3)
        except ValueError:pass
        else:raise AssertionError('Unknown input must block')
        assert not (Path(td)/'out').exists()

if __name__=='__main__':
    for t in (success,blocked_draft,blocked_source,blocked_input):
        t();print(t.__name__+': PASS')
    print('ALL FOUR SYNTHETIC TESTS PASS')
