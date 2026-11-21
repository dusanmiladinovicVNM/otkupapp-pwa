import tempfile,json,hashlib,unittest,zipfile
from pathlib import Path
from xml.etree.ElementTree import Element,SubElement,tostring
from legacy_audit import AuditError
import legacy_amb10_target_adapter_v2 as m
import legacy_amb10_import_readiness_v2 as r
from test_legacy_amb10_import_readiness_v2 import TestMIG064

class TestMIG065(unittest.TestCase):
 def fixture(self,root):
  linked=root/'linked';rows,headers,holds=TestMIG064().fixture(linked)
  inp=root/'ready'
  self.assertEqual(r.cli(['--linked',str(linked),'--out',str(inp)]),0)
  return linked,inp
 def template(self,path,cols=None,macros=True):
  def table(n,fields):
   root=Element('table',{'name':n})
   columns=SubElement(root,'tableColumns')
   for i,k in enumerate(fields):SubElement(columns,'tableColumn',{'id':str(i+1),'name':k})
   return tostring(root)
  with zipfile.ZipFile(path,'w') as z:
   z.writestr('xl/tables/table1.xml',table('tblAmbalaza', cols or m.linked.CANONICAL_COLUMNS))
   z.writestr('xl/tables/table2.xml',table('tblAmbalazaDokument',sorted(m.REQUIRED_HEADERS)))
   if macros:z.writestr('xl/vbaProject.bin',b'SYNTHETIC TEST ONLY')
 def test_end_to_end_full_fixture(self):
  with tempfile.TemporaryDirectory() as t:
   base=Path(t);linked,ready=self.fixture(base)
   xlsm=base/'template.xlsm';self.template(xlsm)
   out=base/'out'
   self.assertEqual(m.cli(['--linked',str(linked),'--readiness',str(ready),
                    '--template',str(xlsm),'--out',str(out)]),0)
   s=json.loads((out/m.SUMMARY).read_text())
   self.assertEqual(s['ledger_payload_rows_not_importable'],4669)
   self.assertEqual(s['revers_header_payload_rows'],48)
   self.assertEqual(s['unknown_document_holds'],6)
   self.assertEqual(s['unexplained_balance_residuals'],0)
   self.assertTrue(s['template_inspection']['target_macros_present'])
   self.assertTrue(s['template_inspection']['required_target_tables_present'])
   self.assertFalse(s['canonical_import_approved'])
   self.assertEqual(len((out/m.PRIVATE_ROWS).read_text().splitlines()),4669)
   self.assertEqual(len((out/m.PRIVATE_HEADERS).read_text().splitlines()),48)
   self.assertEqual(len((out/m.PRIVATE_EXCEPTIONS).read_text().splitlines()),8)
   self.assertEqual(hashlib.sha256(xlsm.read_bytes()).hexdigest(),s['template_inspection']['target_file_sha256'])
   with self.assertRaises(AuditError):m.cli(['--linked',str(linked),'--readiness',str(ready),'--out',str(out)])
 def test_upstream_tamper_detected(self):
  with tempfile.TemporaryDirectory() as t:
   base=Path(t);linked,ready=self.fixture(base)
   f=ready/r.PRIVATE_MAP
   with f.open('a') as o:o.write('{}\n')
   with self.assertRaises(AuditError):m.materialize(linked,ready)
 def test_wrong_target_columns_refused(self):
  with tempfile.TemporaryDirectory() as t:
   base=Path(t);p=base/'bad.xlsm'
   self.template(p,cols=list(m.linked.CANONICAL_COLUMNS)[1:])
   with self.assertRaises(AuditError):m.inspect_template(p)
 def test_suspect_macro_extension_refused(self):
  with tempfile.TemporaryDirectory() as t:
   base=Path(t);p=base/'bad.xlsm'
   self.template(p,macros=False)
   with self.assertRaises(AuditError):m.inspect_template(p)
 def test_provenance_crosswalk_drift_fails(self):
  with tempfile.TemporaryDirectory() as t:
   base=Path(t);linked,ready=self.fixture(base)
   s,rmap,holds=m.verify_readiness(ready)
   rmap[0]['document_id']='FAKE'
   with self.assertRaises(AuditError):
    _,headers,rows,unknown=r.verify_linked(linked)
    original,_,_,_,_=r.classify(rows,headers,unknown)
    if [m.canon(z) for z in original]!=[m.canon(z) for z in rmap]:raise AuditError('Crosswalk mismatch')

if __name__=='__main__': unittest.main()
