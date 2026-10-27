import tempfile,json,hashlib,unittest,zipfile
from pathlib import Path
from xml.etree.ElementTree import Element,SubElement,tostring
from legacy_audit import AuditError
from test_legacy_amb10_target_adapter_v2 import TestMIG065
import legacy_amb10_target_adapter_v2 as adapter
import legacy_amb10_workbook_preflight_v2 as m

N='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR='http://schemas.openxmlformats.org/package/2006/relationships'

def xml(el):return tostring(el,encoding='utf-8',xml_declaration=True)

def test_template(path,overwrite=None,wrong_header=False,include_macros=True):
 tables=[('tblAmbalaza',list(adapter.linked.CANONICAL_COLUMNS)),
         ('tblAmbalazaDokument',sorted(adapter.REQUIRED_HEADERS)),
         ('tblOtkup',['OtkupID','Opis']),('tblOtpremnica',['OtpremnicaID']),('tblPrijemnica',['PrijemnicaID'])]
 if wrong_header:tables[0]=('tblAmbalaza',[x for x in tables[0][1] if x!='StornoOd'])
 with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED) as z:
  z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
  w=Element('{%s}workbook'%N);ss=SubElement(w,'{%s}sheets'%N)
  wrels=Element('{%s}Relationships'%PR)
  for i,(name,cols) in enumerate(tables,1):
   SubElement(ss,'{%s}sheet'%N,{'name':name,'sheetId':str(i),'{%s}id'%R:'rId'+str(i)})
   SubElement(wrels,'{%s}Relationship'%PR,{'Id':'rId'+str(i),'Type':R+'/worksheet','Target':'worksheets/sheet'+str(i)+'.xml'})
   t=Element('{%s}table'%N,{'name':name,'ref':'A1:'+col(len(cols))+'2','headerRowCount':'1'})
   tc=SubElement(t,'{%s}tableColumns'%N,{'count':str(len(cols))})
   for j,x in enumerate(cols,1):SubElement(tc,'{%s}tableColumn'%N,{'id':str(j),'name':x})
   z.writestr('xl/tables/table'+str(i)+'.xml',xml(t))
   ws=Element('{%s}worksheet'%N)
   dat=SubElement(ws,'{%s}sheetData'%N)
   r1=SubElement(dat,'{%s}row'%N,{'r':'1'})
   for j,_ in enumerate(cols,1):
    c=SubElement(r1,'{%s}c'%N,{'r':col(j)+'1','t':'inlineStr'})
    si=SubElement(c,'{%s}is'%N);SubElement(si,'{%s}t'%N).text=cols[j-1]
   if overwrite==name:
    row=SubElement(dat,'{%s}row'%N,{'r':'2'})
    c=SubElement(row,'{%s}c'%N,{'r':'A2'});SubElement(c,'{%s}v'%N).text='123'
   tps=SubElement(ws,'{%s}tableParts'%N,{'count':'1'})
   SubElement(tps,'{%s}tablePart'%N,{'{%s}id'%R:'x1'})
   z.writestr('xl/worksheets/sheet'+str(i)+'.xml',xml(ws))
   shrels=Element('{%s}Relationships'%PR)
   SubElement(shrels,'{%s}Relationship'%PR,{'Id':'x1','Type':R+'/table','Target':'../tables/table'+str(i)+'.xml'})
   z.writestr('xl/worksheets/_rels/sheet'+str(i)+'.xml.rels',xml(shrels))
  z.writestr('xl/workbook.xml',xml(w));z.writestr('xl/_rels/workbook.xml.rels',xml(wrels))
  if include_macros:z.writestr('xl/vbaProject.bin',b'FAKE TEST MACRO')

def col(n):
 s=''
 while n:n,k=divmod(n-1,26);s=chr(65+k)+s
 return s

class TestMIG066(unittest.TestCase):
 def prepare(self,base):
  linked,ready=TestMIG065().fixture(base)
  ad=base/'adapter'
  self.assertEqual(adapter.cli(['--linked',str(linked),'--readiness',str(ready),'--out',str(ad)]),0)
  _,ledger,_,_=m.verify_adapter(ad)
  parents={name:set() for name in m.PARENTS}
  for row in ledger:
   doc=row['target_column_values_WITHOUT_AMBID'];key=doc['DokumentTIP']
   if key in parents:parents[key].add(doc['DokumentID'])
  stage=base/'stage';stage.mkdir()
  counts={}
  for key,(table,pk) in m.PARENTS.items():
   items=[{pk:v} for v in sorted(parents[key])];counts[table]=len(items)
   (stage/(table+'.jsonl')).write_text(''.join(json.dumps(x)+'\n' for x in items))
  (stage/'agrix_stage_summary.json').write_text(json.dumps({'format':'MIG-02.0',
   'source_sha256':adapter.ready.SOURCE_SHA,'canonical_stage_counts':counts}))
  return ad,stage

 def test_realistic_end_to_end(self):
  with tempfile.TemporaryDirectory() as t:
   b=Path(t);ad,stage=self.prepare(b);xlsx=b/'blank.xlsm';test_template(xlsx)
   before=hashlib.sha256(xlsx.read_bytes()).hexdigest()
   out=b/'out'
   self.assertEqual(m.cli(['--adapter',str(ad),'--stage',str(stage),'--template',str(xlsx),'--out',str(out)]),0)
   s=json.loads((out/m.SUMMARY).read_text())
   self.assertEqual(s['candidate_ledger_rows'],4669)
   self.assertEqual(sum(s['staged_parent_preflight']['staged_parent_references_verified'].values()),4621)
   self.assertTrue(s['target_preflight']['all_required_tables_are_empty'])
   self.assertFalse(s['import_approved']);self.assertFalse(s['target_workbook_written'])
   self.assertEqual(hashlib.sha256(xlsx.read_bytes()).hexdigest(),before)
 def test_existing_target_data_refused(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'x.xlsm';test_template(p,overwrite='tblAmbalaza')
   with self.assertRaises(AuditError):m.inspect_target(p)
 def test_existing_parent_data_refused(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'x.xlsm';test_template(p,overwrite='tblOtkup')
   with self.assertRaises(AuditError):m.inspect_target(p)
 def test_wrong_schema_refused(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'x.xlsm';test_template(p,wrong_header=True)
   with self.assertRaises(AuditError):m.inspect_target(p)
 def test_missing_macro_refused(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'x.xlsm';test_template(p,include_macros=False)
   with self.assertRaises(AuditError):m.inspect_target(p)
 def test_missing_parent_stage_fails_closed(self):
  with tempfile.TemporaryDirectory() as t:
   b=Path(t);ad,stage=self.prepare(b)
   (stage/'tblOtkup.jsonl').unlink()
   _,ledger,_,_=m.verify_adapter(ad)
   with self.assertRaises(AuditError):m.verify_parent_stage(stage,ledger)
 def test_payload_tamper_fails_closed(self):
  with tempfile.TemporaryDirectory() as t:
   b=Path(t);ad,stage=self.prepare(b)
   with (ad/adapter.PRIVATE_ROWS).open('a') as f:f.write('{}\n')
   with self.assertRaises(AuditError):m.verify_adapter(ad)

if __name__=='__main__':unittest.main()
