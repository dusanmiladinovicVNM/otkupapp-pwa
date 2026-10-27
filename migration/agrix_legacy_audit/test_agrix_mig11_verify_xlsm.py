import importlib.util
import json
import tempfile
import zipfile
from pathlib import Path

p=Path(__file__).resolve().with_name('agrix_mig11_verify_xlsm.py')
spec=importlib.util.spec_from_file_location('check',p)
c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)

XML='<?xml version="1.0" encoding="utf-8"?>'
SS=c.SS;RD=c.RD;PK=c.PK

def cell(ref,value):
    return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'

def makexlsm(p,binary=b'VBA-SAME',bad=False):
    rows=[('A','OTP-1','OTK-1'),('B','ZB-1','OTP-1')]
    wb=f'{XML}<workbook xmlns="{SS}" xmlns:r="{RD}"><sheets><sheet name="Links1" sheetId="1" r:id="rId1"/><sheet name="Links2" sheetId="2" r:id="rId2"/></sheets></workbook>'
    wbr=f'{XML}<Relationships xmlns="{PK}"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>'
    with zipfile.ZipFile(p,'w') as z:
        z.writestr('xl/vbaProject.bin',binary)
        z.writestr('xl/workbook.xml',wb)
        z.writestr('xl/_rels/workbook.xml.rels',wbr)
        for i,(prefix,a,b) in enumerate(rows,1):
            r=[prefix,a,b]
            if bad and i==1:r[-1]='WRONG'
            header=c.REQUIRED['tblOtpremnicaIzvori' if i==1 else 'tblZbirnaIzvori']
            worksheet=f'{XML}<worksheet xmlns="{SS}"><sheetData><row r="1">'+''.join(cell(f'{col}1',val) for col,val in zip('ABC',header))+'</row><row r="2">'+''.join(cell(f'{col}2',val) for col,val in zip('ABC',r))+'</row></sheetData></worksheet>'
            z.writestr(f'xl/worksheets/sheet{i}.xml',worksheet)
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'{XML}<Relationships xmlns="{PK}"><Relationship Id="rId1" Target="../tables/table{i}.xml" /></Relationships>')
            cols=''.join(f'<tableColumn id="{n}" name="{v}" />' for n,v in enumerate(header,1))
            tbl=f'{XML}<table xmlns="{SS}" name="{list(c.REQUIRED)[i-1]}" displayName="{list(c.REQUIRED)[i-1]}" ref="A1:C2" headerRowCount="1"><tableColumns count="3">{cols}</tableColumns></table>'
            z.writestr(f'xl/tables/table{i}.xml',tbl)

with tempfile.TemporaryDirectory() as d:
    root=Path(d)
    stage=root/'stage';stage.mkdir()
    opi={'OtpremnicaIzvorID':'A','OtpremnicaID':'OTP-1','OtkupID':'OTK-1'}
    zbi={'ZbirnaIzvorID':'B','ZbirnaID':'ZB-1','OtpremnicaID':'OTP-1'}
    for table,data in [('tblOtpremnicaIzvori',opi),('tblZbirnaIzvori',zbi)]:
        (stage/(table+'.jsonl')).write_text(json.dumps(data)+'\n')
    master=root/'master.xlsm';output=root/'data.xlsm'
    makexlsm(master);makexlsm(output)
    result=c.check_workbook(output,master,stage)
    assert result['pass'] and result['vba_binary_preserved'] and all(x['exact_values_match'] for x in result['tables'].values()),result
    makexlsm(output,bad=True)
    result=c.check_workbook(output,master,stage)
    assert not result['pass'] and 'row values differ' in str(result['errors']),result
    makexlsm(output,binary=b'VBA-DIFFERENT')
    result=c.check_workbook(output,master,stage)
    assert not result['pass'] and not result['vba_binary_preserved'],result
print('MIG11 XLSM verifier PASS: exact links, bad content, VBA mismatch')
