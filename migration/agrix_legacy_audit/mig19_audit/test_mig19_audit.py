import csv
import json
import tempfile
from pathlib import Path
import sys
import unittest
from zipfile import ZipFile
from xml.sax.saxutils import escape

sys.path.insert(0,str(Path(__file__).parent))
import agrix_mig19_malina_readonly_audit as audit

AMBCOLS=['AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd']
HEADCOLS=['AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Napomena','Stornirano','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','IspravkaOdID','ZamenjenSaID']

def mkrow(columns, values):return [values.get(x,'') for x in columns]

def colname(index):
    out=''
    while index:
        index,mod=divmod(index-1,26);out=chr(65+mod)+out
    return out

def sheet_xml(headers, rows):
    def cell(v,r,c):
        cr=colname(c)+str(r)
        if v=='':return ''
        return f'<c r="{cr}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'
    xml=['<?xml version="1.0" encoding="UTF-8"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>']
    for nr,row in enumerate([headers,*rows],1):
        xml.append(f'<row r="{nr}">')
        for nc,v in enumerate(row,1):xml.append(cell(v,nr,nc))
        xml.append('</row>')
    xml.append('</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
    return ''.join(xml)

def build_fixture(path,ledger,docs):
    m='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg='http://schemas.openxmlformats.org/package/2006/relationships'
    with ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml',f'<?xml version="1.0"?><workbook xmlns="{m}" xmlns:r="{r}"><sheets><sheet name="Ambalaza" sheetId="1" r:id="rId1"/><sheet name="AmbalazaDokument" sheetId="2" r:id="rId2"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{r}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>')
        for idx,(nm,cols,rows) in enumerate([('tblAmbalaza',AMBCOLS,ledger),('tblAmbalazaDokument',HEADCOLS,docs)],1):
            z.writestr(f'xl/worksheets/sheet{idx}.xml',sheet_xml(cols,[mkrow(cols,x) for x in rows]))
            z.writestr(f'xl/worksheets/_rels/sheet{idx}.xml.rels',f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{r}/table" Target="../tables/table{idx}.xml"/></Relationships>')
            z.writestr(f'xl/tables/table{idx}.xml',f'<table xmlns="{m}" id="{idx}" name="{nm}" ref="A1:{colname(len(cols))}{len(rows)+1}"><tableColumns count="{len(cols)}">'+''.join(f'<tableColumn id="{i}" name="{x}"/>' for i,x in enumerate(cols,1))+'</tableColumns></table>')


def amb(**kw):
    result={'AmbID':'AMB-A','Datum':'2026-06-24','TipAmbalaze':'0,4kg','Kolicina':'10','DokumentID':'ADK-1','DokumentTIP':'AmbalazaDokument','OdNalogTip':'Stanica','OdNalogID':'ST-00005','NaNalogTip':'Kooperant','NaNalogID':'KOP-1','VrstaKretanja':'IZDATA_PRAZNA','StornoOd':''}
    result.update(kw)
    return result

def doc(**kw):
    result={'AmbDokID':'ADK-1','Vrsta':'REVERS','BrojDokumenta':'5/240626','Datum':'2026-06-24','BrojOwnerTip':'Stanica','BrojOwnerID':'ST-00005','Stornirano':''}
    result.update(kw)
    return result

class Mig19Tests(unittest.TestCase):
    def run_fixture(self,ledger,headers):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            target=root/'qa.xlsm';out=root/'output'
            build_fixture(target,ledger,headers)
            before=audit.sha256(target)
            exit_code=audit.main(['--workbook',str(target),'--out-dir',str(out),'--expected-revers',str(len(headers)),'--expected-ledger',str(len(ledger))])
            results=json.loads((out/'summary.json').read_text(encoding='utf-8'))
            with (out/'findings.csv').open('r',encoding='utf-8-sig') as f:findings=list(csv.DictReader(f))
            self.assertEqual(before,audit.sha256(target))
            return exit_code,results,findings

    def test_valid_issue_has_no_p1(self):
        status,summary,findings=self.run_fixture([amb()],[doc()])
        self.assertEqual(summary['counts']['revers_headers'],1)
        self.assertTrue(summary['source_unchanged'])
        self.assertNotIn('P1',[r['severity'] for r in findings])
        self.assertEqual(status,1) # HOLD: unknown legacy provenance / negative without history

    def test_cold_storage_supply_valid_direction(self):
        row=amb(AmbID='AMB-B',OdNalogID='ST-00001',NaNalogTip='Stanica',NaNalogID='ST-00002',VrstaKretanja='PRENOS_INTERNO')
        status,summary,findings=self.run_fixture([row],[doc(BrojOwnerID='ST-00002')])
        self.assertNotIn('MALINA_SUPPLY_SOURCE',[x['code'] for x in findings])

    def test_other_station_supply_flagged(self):
        row=amb(AmbID='AMB-C',OdNalogID='ST-00005',NaNalogTip='Stanica',NaNalogID='ST-00002',VrstaKretanja='PRENOS_INTERNO')
        _,_,findings=self.run_fixture([row],[doc(BrojOwnerID='ST-00002')])
        self.assertIn('MALINA_SUPPLY_SOURCE',[x['code'] for x in findings])

    def test_reverse_storno_not_false_class_mismatch(self):
        orig=amb();storno=amb(AmbID='AMB-ST',OdNalogTip='Kooperant',OdNalogID='KOP-1',NaNalogTip='Stanica',NaNalogID='ST-00005',StornoOd='AMB-A')
        _,_,findings=self.run_fixture([orig,storno],[doc(Stornirano='Da')])
        self.assertNotIn('CLASS_MISMATCH',[x['code'] for x in findings])
        self.assertNotIn('STORNO_DIRECTION',[x['code'] for x in findings])

    def test_duplicate_business_effect_p1(self):
        _,summary,findings=self.run_fixture([amb(),amb(AmbID='AMB-B')],[doc()])
        self.assertEqual(summary['gate'],'FAIL')
        self.assertIn('AMB_INV_04_DUPLICATE',[x['code'] for x in findings])

if __name__=='__main__':unittest.main(verbosity=2)
