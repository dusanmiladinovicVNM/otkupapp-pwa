import hashlib
import json
from pathlib import Path
import tempfile
import zipfile

from agrix_mig14_fix_predlogcena import patch, sha, inspect_workbook_after, Q

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK='http://schemas.openxmlformats.org/package/2006/relationships'


def fixture(root, second_zero=False):
    root=Path(root);stage=root/'stage';stage.mkdir()
    records=[{'OtpremnicaID':'OTP-00440','PredlogCena':0},
             {'OtpremnicaID':'OTP-00001','PredlogCena':0 if second_zero else 120}]
    source=[{'OtpremnicaID':'OTP-00440','Cena':0}]
    checks={}
    for name,rows in [('tblOtpremnicaStavke.jsonl',records),('SOURCE__tblOtpremnica.jsonl',source)]:
        f=stage/name
        f.write_text(''.join(json.dumps(x)+'\n' for x in rows),encoding='utf8')
        checks[name]=sha(f)
    wb=f'''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{NS}" xmlns:r="{R}"><sheets><sheet name="OtpremnicaStavke" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    rels=f'''<Relationships xmlns="{PK}"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'''
    sheetrels=f'''<Relationships xmlns="{PK}"><Relationship Id="rId2" Type="x" Target="../tables/table8.xml"/></Relationships>'''
    table=f'''<table xmlns="{NS}" name="tblOtpremnicaStavke" displayName="tblOtpremnicaStavke" ref="A1:B3" headerRowCount="1"><tableColumns count="2"><tableColumn id="1" name="OtpremnicaID"/><tableColumn id="2" name="PredlogCena"/></tableColumns></table>'''
    sheet=f'''<worksheet xmlns="{NS}"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>OtpremnicaID</t></is></c><c r="B1" t="inlineStr"><is><t>PredlogCena</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>OTP-00440</t></is></c><c r="B2" s="21"><v>0</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>OTP-00001</t></is></c><c r="B3" s="21"><v>{0 if second_zero else 120}</v></c></row></sheetData><tableParts count="1"/></worksheet>'''
    xlsm=root/'input.xlsm'
    with zipfile.ZipFile(xlsm,'w') as z:
        for name,data in [
            ('xl/workbook.xml',wb),('xl/_rels/workbook.xml.rels',rels),
            ('xl/worksheets/sheet1.xml',sheet),
            ('xl/worksheets/_rels/sheet1.xml.rels',sheetrels),
            ('xl/tables/table8.xml',table),('xl/vbaProject.bin',b'fake-not-real')]:
            z.writestr(name,data)
    return xlsm,stage,checks


def test_success():
    with tempfile.TemporaryDirectory() as t:
        x,stage,checks=fixture(t)
        original=x.read_bytes()
        out=Path(t)/'out'
        patch(x,stage,out,expected_sha=sha(x),checks=checks,expected_rows=2)
        assert x.read_bytes()==original
        result=out/'AgriX_MIG14_PRICE_PATCHED.xlsm'
        with zipfile.ZipFile(x) as a,zipfile.ZipFile(result) as b:
            for fn in a.namelist():
                if fn !='xl/worksheets/sheet1.xml':assert a.read(fn)==b.read(fn)
            a1=a.read('xl/worksheets/sheet1.xml')
            b1=b.read('xl/worksheets/sheet1.xml')
            assert a1.replace(b'<c r="B2" s="21"><v>0</v></c>',b'<c r="B2" s="21"></c>')==b1
            p=inspect_workbook_after(b,2)
            assert p=={'target':'empty','positive':1,'count':2},p
        verify=json.loads((out/'MIG14_price_final_verify.json').read_text())
        assert verify['pass'] is True


def test_reject_double_zero():
    with tempfile.TemporaryDirectory() as t:
        x,stage,checks=fixture(t,second_zero=True)
        before=x.read_bytes();out=Path(t)/'out'
        try:patch(x,stage,out,expected_sha=sha(x),checks=checks,expected_rows=2)
        except ValueError:pass
        else:raise AssertionError('Second zero must be rejected')
        assert not out.exists() and x.read_bytes()==before


def test_reject_wrong_workbook():
    with tempfile.TemporaryDirectory() as t:
        x,stage,checks=fixture(t)
        out=Path(t)/'out'
        try:patch(x,stage,out,expected_sha='0'*64,checks=checks,expected_rows=2)
        except ValueError:pass
        else:raise AssertionError('Unexpected checksum must reject')
        assert not out.exists()


def test_reject_mutated_stage():
    with tempfile.TemporaryDirectory() as t:
        x,stage,checks=fixture(t)
        with (stage/'SOURCE__tblOtpremnica.jsonl').open('a') as f:f.write('{}\n')
        out=Path(t)/'out'
        try:patch(x,stage,out,expected_sha=sha(x),checks=checks,expected_rows=2)
        except ValueError:pass
        else:raise AssertionError('Mutated source must reject')
        assert not out.exists()

if __name__=='__main__':
    for t in [test_success,test_reject_double_zero,test_reject_wrong_workbook,test_reject_mutated_stage]:
        t();print(t.__name__+': PASS')
    print('ALL SYNTHETIC TESTS PASS')
