import contextlib, io, json, tempfile, unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile
from xml.sax.saxutils import escape

import agrix_mig19_11_istorijska_semantika as app

COLS=['AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP',
      'CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID',
      'NaNalogTip','NaNalogID','VrstaKretanja','StornoOd']

def create_xlsm(filename,rows):
    def colname(c):
        t=''
        while c:c,r=divmod(c-1,26);t=chr(65+r)+t
        return t
    def sheet():
        x=['<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>']
        for idx,row in enumerate([dict(zip(COLS,COLS))]+rows,1):
            x.append(f'<row r="{idx}">')
            for j,c in enumerate(COLS,1):
                v=row.get(c,'')
                if v!='':x.append(f'<c r="{colname(j)}{idx}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>')
            x.append('</row>')
        x.append('</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
        return ''.join(x)
    m='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg='http://schemas.openxmlformats.org/package/2006/relationships'
    with ZipFile(filename,'w') as z:
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{m}" xmlns:r="{r}"><sheets><sheet name="Ambalaza" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml',sheet())
        z.writestr('xl/worksheets/_rels/sheet1.xml.rels',f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{r}/table" Target="../tables/table1.xml"/></Relationships>')
        z.writestr('xl/tables/table1.xml',f'<table xmlns="{m}" id="1" name="tblAmbalaza" ref="A1:P{len(rows)+1}"><tableColumns count="16">'+''.join(f'<tableColumn id="{i}" name="{c}"/>' for i,c in enumerate(COLS,1))+'</tableColumns></table>')


class Integration(unittest.TestCase):
    def test_full_readonly_workflow(self):
        from test_agrix_mig19_11_istorijska_semantika import old, target
        oldrows=[old('Otkup','Kooperant','K-1','Izlaz',10,'AMB-00001'),
            old('Otkup','Stanica','ST-00002','Ulaz',10,'AMB-00002'),
            old('OM-Izlaz-Koop','Kooperant','K-1','Ulaz',4,'AMB-00003'),
            old('OM-Izlaz-Koop','Stanica','ST-00002','Izlaz',4,'AMB-00004')]
        news=[target('AMBALAZA_UZ_ROBU','Kooperant','K-1','Stanica','ST-00002',10,'M-0001'),
              target('IZDATA_PRAZNA','Stanica','ST-00002','Kooperant','K-1',4,'M-0002')]
        with tempfile.TemporaryDirectory() as td:
            root=Path(td);stage=root/'stage';stage.mkdir()
            src=stage/'UNCONVERTED__tblAmbalaza.jsonl'
            src.write_text(''.join(json.dumps(r)+'\n' for r in oldrows),encoding='utf-8')
            book=root/'QA.xlsm';create_xlsm(book,news)
            hashes=(app.sha256(src),app.sha256(book))
            out=root/'result'
            with patch.object(app,'STAGE_SHA',hashes[0]),contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(app.main(['--stage',str(stage),'--workbook',str(book),
                           '--out-dir',str(out),'--expected-target-sha',hashes[1]]),0)
            s=json.loads((out/'MIG1911_SUMMARY.json').read_text(encoding='utf-8'))
            self.assertEqual(s['legacy_otkup_slice_mismatches_before'],2)
            self.assertEqual(s['legacy_otkup_slice_mismatches_after'],0)
            self.assertEqual(s['unchanged_balance_mismatch_count'],0)
            self.assertEqual(s['canonical_source_grounded_relabels'],1)
            self.assertEqual((app.sha256(src),app.sha256(book)),hashes)

if __name__=='__main__':unittest.main()
