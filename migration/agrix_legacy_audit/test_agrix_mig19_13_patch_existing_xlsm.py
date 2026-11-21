#!/usr/bin/env python3
import tempfile
import json
from pathlib import Path
from zipfile import ZipFile
import unittest
import sys
from decimal import Decimal

sys.path.insert(0,str(Path(__file__).parent))
from agrix_mig19_13_patch_existing_xlsm import (
 make_xlsm,planned_rows,UNKNOWN,TRANSFERS,CANON_COLUMNS,DOC_COLUMNS,
 OMULAZ,parse,qn,M, XlsmReadOnly, verify_projection, ensure_source,
)


def escaped(v):
    from xml.sax.saxutils import escape
    return escape(str(v))


def make_sheet(headers,data,table_id):
    from agrix_mig19_13_patch_existing_xlsm import xl_col
    rows=[]
    for k,vals in enumerate([headers]+data,1):
        cells=''.join(f'<c r="{xl_col(j)}{k}" t="inlineStr"><is><t>{escaped(v)}</t></is></c>' for j,v in enumerate(vals,1) if v not in ('',None))
        rows.append(f'<row r="{k}">{cells}</row>')
    body=''.join(rows)
    return ('<?xml version="1.0" encoding="utf-8"?>'
     '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:xr4="http://schemas.microsoft.com/office/spreadsheetml/2017/revision4" mc:Ignorable="xr4">'
     f'<dimension ref="A1:{xl_col(len(headers))}{1+len(data)}"/>'
     f'<sheetData>{body}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')


def make_table(name,headers,n):
    from agrix_mig19_13_patch_existing_xlsm import xl_col
    cols=''.join(f'<tableColumn id="{i}" name="{escaped(v)}"/>' for i,v in enumerate(headers,1))
    return (f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{1 if name=="tblAmbalaza" else 2}" name="{name}" displayName="{name}" ref="A1:{xl_col(len(headers))}{n+1}"><autoFilter ref="A1:{xl_col(len(headers))}{n+1}"/><tableColumns count="{len(headers)}">{cols}</tableColumns></table>')

class Tests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        d=Path(self.temp.name)
        self.src=d/'old.xlsm'; self.out=d/'new.xlsm'
        legacy=[]
        for aid,(date,st,qty,orig) in OMULAZ.items():
            legacy.append(dict(AmbID=aid,Datum=date,TipAmbalaze='0,4kg',Kolicina=str(qty),DokumentID='OLD-'+aid,DokumentTIP='OMUlaz',CreatedAt='',Stornirano='Ne',EntitetTip='Stanica',EntitetID=st,Smer='Ulaz'))
        self.idx={r['AmbID']:r for r in legacy}
        self.new,self.docs=planned_rows(self.idx)
        origamb=dict.fromkeys(CANON_COLUMNS,'');origamb.update(AmbID='AMB-EXIST',Datum='2026-06-01',TipAmbalaze='0,4kg',Kolicina='12',DokumentID='OTK-01',DokumentTIP='Otkup',OdNalogTip='Kooperant',OdNalogID='KP-1',NaNalogTip='Stanica',NaNalogID='ST-00001',VrstaKretanja='AMBALAZA_UZ_ROBU')
        origdoc=dict.fromkeys(DOC_COLUMNS,'');origdoc.update(AmbDokID='ADK-EXIST',Vrsta='REVERS',BrojDokumenta='1/2026',Datum='2026-06-01',BrojOwnerTip='Stanica',BrojOwnerID='ST-00001')
        with ZipFile(self.src,'w') as z:
            z.writestr('xl/workbook.xml','''<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Ambalaza" sheetId="1" r:id="rId1"/><sheet name="AmbalazaDokument" sheetId="2" r:id="rId2"/></sheets></workbook>''')
            z.writestr('xl/_rels/workbook.xml.rels','''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/></Relationships>''')
            z.writestr('[Content_Types].xml','''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>''')
            z.writestr('xl/worksheets/sheet3.xml',make_sheet(CANON_COLUMNS,[[origamb[k] for k in CANON_COLUMNS]],1))
            z.writestr('xl/worksheets/sheet4.xml',make_sheet(DOC_COLUMNS,[[origdoc[k] for k in DOC_COLUMNS]],2))
            for num,sheet in [(1,3),(2,4)]:
                z.writestr(f'xl/worksheets/_rels/sheet{sheet}.xml.rels',f'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{num}.xml"/></Relationships>''')
            z.writestr('xl/tables/table1.xml',make_table('tblAmbalaza',CANON_COLUMNS,1))
            z.writestr('xl/tables/table2.xml',make_table('tblAmbalazaDokument',DOC_COLUMNS,1))
            z.writestr('xl/vbaProject.bin',b'DO_NOT_MODIFY_VBA')

    def test_source_rows(self):
        self.assertEqual(len(self.new),3)
        self.assertEqual(len(self.docs),3)
        self.assertEqual(sum(int(x['Kolicina']) for x in self.new),633)
        self.assertEqual({x['OdNalogID'] for x in self.new},{'ST-00001'})
        self.assertEqual({x['VrstaKretanja'] for x in self.new},{'PRENOS_INTERNO'})
        self.assertEqual({x['DokumentTIP'] for x in self.new},{'AmbalazaDokument'})
    def test_preserve_macro_and_readback(self):
        make_xlsm(self.src,self.out,self.new,self.docs,self.idx)
        a=XlsmReadOnly(self.out)
        try:
            self.assertEqual(len(a.read_table('tblAmbalaza',CANON_COLUMNS)),4)
            self.assertEqual(len(a.read_table('tblAmbalazaDokument',DOC_COLUMNS)),4)
            self.assertIn('MIG19_Nepoznati',a.sheets)
            vals=a.read_table('tblAmbalaza',CANON_COLUMNS)
            self.assertEqual([x['AmbID'] for x in vals[-3:]], [r['AmbID'] for r in self.new])
            self.assertEqual([int(x['Kolicina']) for x in vals[-3:]], [300,250,83])
        finally:a.close()
        with ZipFile(self.src) as old,ZipFile(self.out) as new:
            self.assertEqual(new.read('xl/vbaProject.bin'),old.read('xl/vbaProject.bin'))
            self.assertIsNone(new.testzip())
            self.assertIn('xmlns:xr4=',new.read('xl/worksheets/sheet3.xml').decode())
            self.assertIn('mc:Ignorable="xr4"',new.read('xl/worksheets/sheet3.xml').decode())
            changed={p for p in old.namelist() if old.read(p)!=new.read(p)}
            self.assertEqual(changed,{
                'xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml',
                'xl/worksheets/sheet3.xml','xl/worksheets/sheet4.xml',
                'xl/tables/table1.xml','xl/tables/table2.xml'})
            sheet=new.read('xl/worksheets/sheet5.xml').decode()
            for aid in UNKNOWN:self.assertIn(aid,sheet)
            self.assertNotIn('VrstaKretanja',sheet)
    def test_output_protected(self):
        with self.assertRaises(ValueError):make_xlsm(self.src,self.src,self.new,self.docs,self.idx)
    def test_repeat_abort(self):
        make_xlsm(self.src,self.out,self.new,self.docs,self.idx)
        with self.assertRaises(ValueError):make_xlsm(self.src,self.out,self.new,self.docs,self.idx)

if __name__=='__main__':unittest.main(verbosity=2)
