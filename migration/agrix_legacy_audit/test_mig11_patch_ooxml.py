import tempfile,json,zipfile,subprocess,sys,hashlib
from pathlib import Path
from xml.etree import ElementTree as ET

HERE=Path(__file__).resolve().parent
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RPK='http://schemas.openxmlformats.org/package/2006/relationships'
OPICOLS=['OtpremnicaIzvorID','OtpremnicaID','OtkupID','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy']
ZBICOLS=['ZbirnaIzvorID','ZbirnaID','OtpremnicaID','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy']

def table_xml(name,cols,n):
    colxml=''.join(f'<tableColumn id="{i+1}" name="{c}"/>' for i,c in enumerate(cols))
    end=chr(65+len(cols)-1)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<table xmlns="{S}" id="{n}" name="{name}" displayName="{name}" ref="A1:{end}2" totalsRowShown="0">'
            f'<autoFilter ref="A1:{end}2"/>'
            f'<tableColumns count="{len(cols)}">{colxml}</tableColumns>'
            f'<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
            f'</table>').encode()

def sheet_xml(cols, extra='', stub=''):
    cxml=''.join(f'<c r="{chr(65+i)}1" t="inlineStr"><is><t>{c}</t></is></c>' for i,c in enumerate(cols))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         f'<worksheet xmlns="{S}" xmlns:r="{REL}"><dimension ref="A1:{chr(65+len(cols)-1)}2"/>'
         f'<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
         f'<sheetFormatPr defaultRowHeight="15"/><sheetData>'
         f'<row r="1">{cxml}</row><row r="2">{stub}</row>{extra}</sheetData>'
         f'<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>').encode()

def rels(path):
    return f'<Relationships xmlns="{RPK}"><Relationship Id="rId1" Type="{REL}/table" Target="../tables/{path}"/></Relationships>'.encode()

def create_book(path,stub='',extra=''):
    wb=f'<workbook xmlns="{S}" xmlns:r="{REL}"><sheets><sheet name="OtpremnicaIzvori" sheetId="1" r:id="rId1"/><sheet name="ZbirnaIzvori" sheetId="2" r:id="rId2"/></sheets></workbook>'
    wb_rels=(f'<Relationships xmlns="{RPK}">'
          f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
          f'<Relationship Id="rId2" Type="{REL}/worksheet" Target="worksheets/sheet2.xml"/>'
          f'</Relationships>')
    content=f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
    with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',content)
        z.writestr('xl/workbook.xml',wb)
        z.writestr('xl/_rels/workbook.xml.rels',wb_rels)
        for n,cols in [(1,OPICOLS),(2,ZBICOLS)]:
            z.writestr(f'xl/worksheets/sheet{n}.xml',sheet_xml(cols,extra if n==1 else '',stub if n==1 else ''))
            z.writestr(f'xl/worksheets/_rels/sheet{n}.xml.rels',rels(f'table{n}.xml'))
            z.writestr(f'xl/tables/table{n}.xml',table_xml('tblOtpremnicaIzvori' if n==1 else 'tblZbirnaIzvori',cols,n))
        z.writestr('xl/vbaProject.bin',b'DUMMY-UNCHANGED-VBA')
        z.writestr('xl/otherImportantPart.xml',b'<keep>EXACT BYTE EQUALITY</keep>')

def write_rows(path,rs):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(''.join(json.dumps(x)+'\n' for x in rs))

def case(opi=3,zbi=2,stub='',extra=''):
    root=Path(tempfile.mkdtemp(prefix='mig11_patch_test_'))
    indir=root/'input';indir.mkdir()
    book=indir/'already_migrated.xlsm';create_book(book,stub,extra)
    master=root/'master.xlsm';create_book(master,stub,extra)
    stage=root/'stage';stage.mkdir()
    write_rows(stage/'tblOtkup.jsonl',[{'OtkupID':f'OTK-{i}'} for i in range(opi)])
    write_rows(stage/'tblOtpremnica.jsonl',[{'OtpremnicaID':f'OTP-{i}'} for i in range(zbi)])
    write_rows(stage/'tblZbirna.jsonl',[{'ZbirnaID':f'ZBR-{i}'} for i in range(zbi)])
    write_rows(stage/'tblOtpremnicaIzvori.jsonl',[
        {'OtpremnicaIzvorID':f'OPI-M11-{i}','OtpremnicaID':f'OTP-{i % zbi}','OtkupID':f'OTK-{i}'} for i in range(opi)])
    write_rows(stage/'tblZbirnaIzvori.jsonl',[
        {'ZbirnaIzvorID':f'ZBI-M11-{i}','ZbirnaID':f'ZBR-{i}','OtpremnicaID':f'OTP-{i}'} for i in range(zbi)])
    return root,indir,master,stage

def invoke(root,indir,master,stage,opi,zbi):
    return subprocess.run([sys.executable,str(HERE/'agrix_mig11_write_links_ooxml.py'),
        '--input-dir',str(indir),'--master',str(master),'--stage',str(stage),
        '--out-dir',str(root/'output'),'--expected-opi',str(opi),'--expected-zbi',str(zbi)],
        text=True,capture_output=True)

r,i,m,s=case(1692,907)
p=invoke(r,i,m,s,1692,907)
print('REAL-SCALE TEST',p.returncode,p.stdout,p.stderr)
assert p.returncode==0
out=r/'output'/'AgriX_MIG11_LINKS_PATCHED.xlsm'
with zipfile.ZipFile(out) as z,zipfile.ZipFile(i/'already_migrated.xlsm') as old:
    assert z.read('xl/otherImportantPart.xml')==old.read('xl/otherImportantPart.xml')
    assert z.read('xl/vbaProject.bin')==old.read('xl/vbaProject.bin')
    assert b'ref="A1:G1693"' in z.read('xl/tables/table1.xml')
    assert b'ref="A1:G908"' in z.read('xl/tables/table2.xml')
assert json.loads((r/'output'/'mig11_final_verify.json').read_text())['pass']

r,i,m,s=case(3,2,stub='<c r="A2" t="inlineStr"><is><t>existing-data</t></is></c>')
p=invoke(r,i,m,s,3,2)
print('EXISTING DATA REFUSAL',p.returncode,p.stderr.strip())
assert p.returncode==2 and not (r/'output'/'AgriX_MIG11_LINKS_PATCHED.xlsm').exists()

r,i,m,s=case(3,2,extra='<row r="3"><c r="A3" t="n"><v>42</v></c></row>')
p=invoke(r,i,m,s,3,2)
print('EXPANSION COLLISION REFUSAL',p.returncode,p.stderr.strip())
assert p.returncode==2

r,i,m,s=case(3,2)
rows=(s/'tblOtpremnicaIzvori.jsonl').read_text().splitlines()
row=json.loads(rows[0]);row['OtpremnicaID']='NONEXISTENT';rows[0]=json.dumps(row)
(s/'tblOtpremnicaIzvori.jsonl').write_text('\n'.join(rows)+'\n')
p=invoke(r,i,m,s,3,2)
print('ORPHAN FK REFUSAL',p.returncode,p.stderr.strip())
assert p.returncode==2
print('ALL FOUR MIG11 PATCH GATES PASS')
