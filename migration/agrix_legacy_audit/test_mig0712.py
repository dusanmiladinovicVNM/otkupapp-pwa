import hashlib
import os
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
import agrix_mig0712_vba_isolation as a
class TestVbaIsolation(unittest.TestCase):
 def test_sha(self):
  self.assertEqual(a.sha(b'a'),hashlib.sha256(b'a').hexdigest())
 def test_invalid_zip(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'demo.xlsm';p.write_text('not a zip')
   with self.assertRaises(Exception):a.zip_vba(p)
 def test_missing_bin(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'demo.xlsm'
   with ZipFile(p,'w') as z:z.writestr('xl/workbook.xml','')
   with self.assertRaises(KeyError):a.zip_vba(p)
 def test_bad_ole_signature(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'demo.xlsm'
   with ZipFile(p,'w') as z:z.writestr('xl/vbaProject.bin',b'bad')
   with self.assertRaisesRegex(RuntimeError,'OLE'):a.zip_vba(p)
 def test_valid_ole_signature(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'demo.xlsm'
   data=a.CFB_SIGNATURE+b'0'*100
   with ZipFile(p,'w') as z:z.writestr('xl/vbaProject.bin',data)
   d=a.zip_vba(p)
   self.assertEqual(d['vba_binary_sha256'],a.sha(data))
   self.assertEqual(d['vba_binary_bytes'],len(data))
 def test_non_windows_fail_closed(self):
  if os.name=='nt':self.skipTest('Windows')
  with self.assertRaisesRegex(RuntimeError,'Windows'):
   a.main(['--template','a','--scratch','b','--out','c'])
if __name__=='__main__':unittest.main()
