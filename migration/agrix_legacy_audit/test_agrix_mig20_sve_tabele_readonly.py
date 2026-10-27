import tempfile, unittest, json
from pathlib import Path
from zipfile import ZipFile
from xml.sax.saxutils import escape
from agrix_mig20_sve_tabele_readonly import run, save

def x(v):return escape(str(v), {'"':'&quot;'})
def col(n):
    s=''
    while n:
        n,i=divmod(n-1,26);s=chr(65+i)+s
    return s

def fixture(path,tabledata):
    # Strictly synthetic legacy-style workbook; no real user data.
    tables=list(tabledata)
    ws=[]; rel=[]
    for i,(name,(hdrs,records)) in enumerate(tabledata.items(),1):
        rows=[]
        for num,values in enumerate([hdrs]+records,1):
            cells=''.join(f'<c r="{col(k)}{num}" t="inlineStr"><is><t>{x(v)}</t></is></c>' for k,v in enumerate(values,1) if v is not None and str(v)!='')
            rows.append(f'<row r="{num}">{cells}</row>')
        xml=(f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             f'<sheetData>{"".join(rows)}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
        cols=''.join(f'<tableColumn id="{n}" name="{x(c)}"/>' for n,c in enumerate(hdrs,1))
        tbl=(f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="{name}" displayName="{name}" id="{i}" ref="A1:{col(len(hdrs))}{len(records)+1}">'
             f'<tableColumns count="{len(hdrs)}">{cols}</tableColumns></table>')
        ws.append((i,xml,tbl))
        rel.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    with ZipFile(path,'w') as z:
        names=''.join(f'<sheet name="{x(name)}" sheetId="{i}" r:id="rId{i}"/>' for i,name in enumerate(tables,1))
        z.writestr('xl/workbook.xml',f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{names}</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(rel)+'</Relationships>')
        for i,xml,tbl in ws:
            z.writestr(f'xl/worksheets/sheet{i}.xml',xml)
            z.writestr(f'xl/tables/table{i}.xml',tbl)
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{i}.xml"/></Relationships>')
        z.writestr('xl/vbaProject.bin',b'MACROS_DO_NOT_TOUCH')

def mkdata():
    return {
      'tblOtkup':(['OtkupID','BrojDokumenta','OtpremnicaID','Datum','StanicaID','Stornirano'],[
       ['OTK-01110','1/030726-14','OTP-00577','2026-07-03','ST-00001',''],
       ['OTK-01111','1/030726-15','OTP-00577','2026-07-03','ST-00001','']]),
      'tblOtkupStavke':(['OtkupStavkaID','OtkupID','Kolicina','Cena','KolAmbalaze'],[
       ['S-1','OTK-01110','28.5','420','15'],['S-2','OTK-01111','60','420','30']]),
      'tblOtpremnica':(['OtpremnicaID','BrojOtpremnice','Datum','StanicaID','Stornirano','IzdatoStatus','BrojZbirne'],[
       ['OTP-00577','1/030726-10','2026-07-03','ST-00001','Da','IZDATO',''],
       ['OTP-00581','1/030726-14','2026-07-03','ST-00001','','IZDATO','S1/030726-14'],
       ['OTP-00582','1/030726-15','2026-07-03','ST-00001','','IZDATO','S1/030726-15'],
       ['OTP-00666','2/050726','2026-07-05','ST-00002','','IZDATO','S2/050726']]),
      'tblOtpremnicaIzvori':(['OtpremnicaIzvorID','OtpremnicaID','OtkupID'],[
       ['OPI-X','OTP-00577','OTK-01110'],['OPI-Y','OTP-00577','OTK-01111']]),
      'tblOtpremnicaStavke':(['OtpremnicaStavkaID','OtpremnicaID','Kolicina','KolAmbalaze'],[
       ['OPTS-1','OTP-00581','28.5','15'],['OPTS-2','OTP-00582','60','30'],['OPTS-3','OTP-00666','82.1','41']]),
      'tblZbirna':(['ZbirnaID','BrojZbirne','Datum','Stornirano'],[
       ['ZBR-00581','S1/030726-14','2026-07-03',''],
       ['ZBR-00666','S2/050726','2026-07-05',''] ,
       ['ZBR-00319','S5/130626-3','2026-06-13','']]),
      'tblZbirnaIzvori':(['ZbirnaIzvorID','ZbirnaID','OtpremnicaID'],[
       ['ZBI-1','ZBR-00581','OTP-00581'],['ZBI-2','ZBR-00666','OTP-00666']]),
      'tblPrijemnica':(['PrijemnicaID','BrojZbirne','Datum','Stornirano'],[
       ['PRJ-00319','s5/130626-3','2026-06-13',''],['PRJ-00651','S2/050726','2026-07-05','Da']]),
      'tblAmbalaza':(['AmbID','DokumentTIP','DokumentID','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','StornoOd','Kolicina','TipAmbalaze','Datum'],[
       ['AMB-666','Otpremnica','OTP-00666','Stanica','ST-00002','Vozac','ST-00002','','41','0,4kg','2026-07-05']])
    }

class TestAudit(unittest.TestCase):
  def setUp(self):
    self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
    self.path=Path(self.tmp.name)/'test.xlsm'
    fixture(self.path,mkdata())
  def go(self):return run(self.path,Path(self.tmp.name)/'stage')
  def test_detects_all_three_patterns(self):
    rep,rows=self.go();codes={x['code'] for x in rows}
    self.assertIn('ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP',codes)
    self.assertIn('ACTIVE_OTKUP_ONLY_STORNO_OTP_LINKS',codes)
    self.assertIn('PRJ_ZBR_CASE_ONLY_OR_FORMAT',codes)
    self.assertIn('ACTIVE_CHAIN_NO_OTKUPS_AND_CANCELLED_PRJ',codes)
    self.assertGreaterEqual(rep['findings_summary']['by_priority']['P1'],4)
  def test_id_collision_is_detected(self):
    d=mkdata();d['tblZbirna'][1].append(['ZBR-00666','S2/050726','2026-07-05',''])
    fixture(self.path,d)
    _,rows=self.go();self.assertTrue(any(x['code']=='DUPLICATE_PRIMARY_ID' for x in rows))
  def test_after_581_manual_fix_581_no_longer_warns(self):
    d=mkdata();d['tblOtpremnicaIzvori'][1][0][1]='OTP-00581'
    d['tblOtkup'][1][0][2]='OTP-00581'
    fixture(self.path,d)
    _,rows=self.go();bad=[x for x in rows if x['code']=='ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP' and x['document_id']=='OTP-00581']
    self.assertEqual(bad,[])
  def test_stage_readonly_and_output(self):
    before=self.path.read_bytes()
    report,rows=self.go()
    out=save(Path(self.tmp.name)/'output',report,rows)
    self.assertEqual(self.path.read_bytes(),before)
    self.assertEqual(len(out),3)
    self.assertTrue(all(p.is_file() for p in out))
    self.assertEqual(json.loads(out[0].read_text(encoding='utf-8'))['format'],report['format'])

if __name__=='__main__':unittest.main(verbosity=2)
