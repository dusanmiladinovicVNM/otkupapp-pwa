import importlib.util
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from unittest.mock import patch

HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('builder',HERE/'agrix_build_master_manual_vbatools.py')
b=importlib.util.module_from_spec(spec)
spec.loader.exec_module(b)

class Tests(unittest.TestCase):
    def test_staging_excludes_only_tools(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp); (root/'src').mkdir()
            for f in ['modVbaTools.bas','modLogo.bas','modMain.bas','frmOtkupUI.frm','frmOtkupUI.frx','ThisWorkbook.doccls']:
                (root/'src'/f).write_text('placeholder',encoding='utf-8')
            got=b.stage_sources(root/'src',root)
            self.assertFalse((got/'modVbaTools.bas').exists())
            self.assertTrue((got/'modLogo.bas').exists())
            self.assertEqual(len(list(got.iterdir())),5)

    def test_stage_refuses_tools_missing(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp); (root/'src').mkdir()
            with self.assertRaisesRegex(RuntimeError,'Exactly one'):
                b.stage_sources(root/'src',root)

    def _xlsx_fixture(self, col2='B'):
        ns=b.TABLE_NS
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'test.xlsm'
            root=ET.Element(ns+'table',{'name':'tblX','displayName':'tblX','ref':'A1:B2'})
            cols=ET.SubElement(root,ns+'tableColumns',{'count':'2'})
            ET.SubElement(cols,ns+'tableColumn',{'id':'1','name':'A'})
            ET.SubElement(cols,ns+'tableColumn',{'id':'2','name':col2})
            with zipfile.ZipFile(p,'w') as z:
                z.writestr('xl/tables/table1.xml',ET.tostring(root))
                z.writestr('xl/vbaProject.bin',bytes.fromhex('D0CF11E0A1B11AE1')+b'hello')
            yield p

    def test_canonical_ooxml_pass(self):
        for p in self._xlsx_fixture():
            r=b.workbook_structure(p,{'tables':[{'table':'tblX','columns':['A','B']}]})
            self.assertEqual(r['tables'],1)
            self.assertEqual(r['columns'],2)

    def test_missing_last_column_fails(self):
        for p in self._xlsx_fixture('WRONG'):
            with self.assertRaisesRegex(RuntimeError,'schema mismatch tblX'):
                b.workbook_structure(p,{'tables':[{'table':'tblX','columns':['A','B']}]})

    def test_binary_missing_fails(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'bad.xlsm'
            with zipfile.ZipFile(p,'w') as z:
                z.writestr('xl/tables/table1.xml', '<table/>')
            with self.assertRaises(Exception):
                b.workbook_structure(p,{'tables':[]})

    def test_premanual_verification_accepts_exact_missing(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'test.xlsm'
            p.write_bytes(b'dummy')
            repo=Path(td)
            (repo/'schema').mkdir();(repo/'schema'/'schema.json').write_text('{}')
            with patch.object(b,'workbook_structure',return_value={'tables':47}),\
                 patch.object(b,'vbide_probe',return_value={'status':'EXPORT_PASS','names':{'modLogo','ThisWorkbook','modMain'},'components':3,'references_count':1}):
                result=b.verify(p,repo,{'modLogo','modMain','modVbaTools'},False,120)
                self.assertEqual(result['status'],'READY_FOR_MANUAL_VBATOOLS')

    def test_postmanual_verification_requires_tools(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'test.xlsm';p.write_bytes(b'dummy')
            repo=Path(td);(repo/'schema').mkdir();(repo/'schema'/'schema.json').write_text('{}')
            with patch.object(b,'workbook_structure',return_value={'tables':47}),\
                 patch.object(b,'vbide_probe',return_value={'status':'EXPORT_PASS','names':{'modMain','modLogo','ThisWorkbook'},'components':3,'references_count':1}):
                with self.assertRaisesRegex(RuntimeError,'missing VBA modules'):
                    b.verify(p,repo,{'modLogo','modMain','modVbaTools'},True,120)

    def test_postmanual_verification_pass_compile_pending(self):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'test.xlsm';p.write_bytes(b'dummy')
            repo=Path(td);(repo/'schema').mkdir();(repo/'schema'/'schema.json').write_text('{}')
            with patch.object(b,'workbook_structure',return_value={'tables':47}),\
                 patch.object(b,'vbide_probe',return_value={'status':'EXPORT_PASS','names':{'modMain','modVbaTools','modLogo','ThisWorkbook'},'components':4,'references_count':1}):
                result=b.verify(p,repo,{'modLogo','modMain','modVbaTools'},True,120)
                self.assertEqual(result['status'],'VBA_REOPEN_PASS_COMPILE_PENDING')

if __name__=='__main__':unittest.main()
