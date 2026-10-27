"""Portable tests of the safe bootstrap (Excel COM itself must run on Windows)."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import agrix_build_master_manual_vbatools as m

class BootstrapTests(unittest.TestCase):
    def test_bootstrap_contains_vba_tools(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td); src=root/'source'; src.mkdir(); root.joinpath('temp').mkdir()
            (src/'modLogo.bas').write_text('kanonski logo',encoding='ascii')
            (src/'modVbaTools.bas').write_text('full tools',encoding='ascii')
            (src/'modConfig.bas').write_text('full config',encoding='ascii')
            dest=m.stage_sources(src,root/'temp')
            self.assertEqual((dest/'modVbaTools.bas').read_text(),'full tools')
            self.assertEqual((dest/'modConfig.bas').read_text(),'full config')
            self.assertIn(m.STUB_TAG,(dest/'modLogo.bas').read_text())
            self.assertNotEqual(m.digest(dest/'modLogo.bas'),m.digest(src/'modLogo.bas'))
            self.assertEqual((src/'modLogo.bas').read_text(),'kanonski logo')
    def test_staged_vba_tool_must_exist(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td);src=root/'src';src.mkdir();(root/'temp').mkdir()
            (src/'modLogo.bas').write_text('source')
            with self.assertRaisesRegex(RuntimeError,'modVbaTools'):
                m.stage_sources(src,root/'temp')
    def test_stub_contains_full_public_api(self):
        code=m.LOGO_STUB
        for name in ('LogoSlika','LogoOtpusti','LogoOdnos','LogoKljuc',
                     'LOGO_SPLASH','LOGO_KARTICA','LOGO_MINI',
                     'LOGO_BG_SPLASH','LOGO_BG_KARTICA','LOGO_BG_MINI'):
            self.assertIn(name,code)
        self.assertNotIn('B64_MINI2',code)
    def test_canonical_logo_remains_large(self):
        self.assertLess(len(m.LOGO_STUB),2000)
    def test_pre_import_requires_stub(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td); (root/'master').mkdir()
            (root/'master'/'modLogo.bas').write_bytes(b'Attribute VB_Name = "modLogo"\r\n\''+m.STUB_TAG.encode())
            inv={'components':{'modLogo':{},'modVbaTools':{},'ThisWorkbook':{}},'references':[]}
            with patch('agrix_mig0712_vba_isolation.run_one',return_value=({'status':'EXPORT_PASS'},inv)) as mock:
                # To exercise full vbide_probe, its temp path is created inside;
                # mock run_one generates a module export in that directory.
                def fake_run_one(label,wb,output,tempdir,timeout):
                    p=Path(tempdir)/label;p.mkdir();(p/'modLogo.bas').write_text(m.STUB_TAG)
                    return {'status':'EXPORT_PASS'},inv
                mock.side_effect=fake_run_one
                self.assertEqual(m.vbide_probe('file.xlsm',120,False)['logo_state'],'BOOTSTRAP_STUB')
    def test_post_import_rejects_stub(self):
        inv={'components':{'modLogo':{},'modVbaTools':{},'ThisWorkbook':{}},'references':[]}
        with patch('agrix_mig0712_vba_isolation.run_one') as mock:
            def fake_run_one(label,wb,output,tempdir,timeout):
                p=Path(tempdir)/label;p.mkdir();(p/'modLogo.bas').write_text(m.STUB_TAG)
                return {'status':'EXPORT_PASS'},inv
            mock.side_effect=fake_run_one
            with self.assertRaisesRegex(RuntimeError,'did not restore'):
                m.vbide_probe('file.xlsm',120,True)
    def test_post_import_accepts_full_logo_shape(self):
        inv={'components':{'modLogo':{},'modVbaTools':{},'ThisWorkbook':{}},'references':[]}
        with patch('agrix_mig0712_vba_isolation.run_one') as mock:
            def fake_run_one(label,wb,output,tempdir,timeout):
                p=Path(tempdir)/label;p.mkdir();(p/'modLogo.bas').write_bytes(b'B64_MINI2\n'+b'X'*40000)
                return {'status':'EXPORT_PASS'},inv
            mock.side_effect=fake_run_one
            self.assertEqual(m.vbide_probe('file.xlsm',120,True)['logo_state'],'CANONICAL_SHAPE')
    def test_abort_if_vbide_export_fails(self):
        with patch('agrix_mig0712_vba_isolation.run_one',return_value=({'status':'FAIL'},None)):
            with self.assertRaisesRegex(RuntimeError,'export FAIL'):
                m.vbide_probe('file.xlsm',120,False)


    def test_canonical_prefix_allows_known_audit_runtime_columns(self):
        expected={'tblPrevoznici':['PrevoznikID','Naziv','Vozac','Registracija','Aktivan'],
                  'tblUtovar':['UtovarID','Godina'],
                  'tblUtovarStavke':['UtovarStavkaID','UtovarID']}
        actual={k:v[:] for k,v in expected.items()}
        actual['tblPrevoznici'] += list(m.RUNTIME_AUDIT_SUFFIX)
        result=m.canonical_table_structure(expected,actual)
        self.assertEqual(result,{'tblPrevoznici':list(m.RUNTIME_AUDIT_SUFFIX)})

    def test_all_three_runtime_audit_tables(self):
        expected={k:['ID','Name'] for k in m.ALLOWED_RUNTIME_SUFFIXES}
        actual={k:expected[k]+list(m.RUNTIME_AUDIT_SUFFIX) for k in expected}
        self.assertEqual(set(m.canonical_table_structure(expected,actual)),set(expected))

    def test_reject_canonical_column_missing_or_misordered(self):
        expected={'tblPrevoznici':['PrevoznikID','Naziv','Aktivan']}
        with self.assertRaisesRegex(RuntimeError,'canonical schema mismatch'):
            m.canonical_table_structure(expected,{'tblPrevoznici':['PrevoznikID','Aktivan','Naziv']})
        with self.assertRaisesRegex(RuntimeError,'canonical schema mismatch'):
            m.canonical_table_structure(expected,{'tblPrevoznici':['PrevoznikID','Naziv']})

    def test_reject_partial_unknown_or_duplicate_audit_suffix(self):
        expected={'tblPrevoznici':['ID','Naziv']}
        for suffix in (['CreatedAt'],list(reversed(m.RUNTIME_AUDIT_SUFFIX)),
                       list(m.RUNTIME_AUDIT_SUFFIX)+['Unknown'],
                       list(m.RUNTIME_AUDIT_SUFFIX)+['CreatedAt']):
            with self.subTest(suffix=suffix):
                with self.assertRaisesRegex(RuntimeError,'unsupported runtime suffix'):
                    m.canonical_table_structure(expected,{'tblPrevoznici':expected['tblPrevoznici']+suffix})

    def test_reject_audit_suffix_on_ordinary_table(self):
        expected={'tblOtkup':['OtkupID','Datum']}
        with self.assertRaisesRegex(RuntimeError,'unsupported runtime suffix'):
            m.canonical_table_structure(expected,{'tblOtkup':expected['tblOtkup']+list(m.RUNTIME_AUDIT_SUFFIX)})

    def test_full_workbook_structure_accepts_runtime_audit_in_saved_ooxml(self):
        import zipfile
        import xml.etree.ElementTree as ET
        with tempfile.TemporaryDirectory() as td:
            wb=Path(td)/'synthetic.xlsm'
            schema={'tables':[{'table':'tblPrevoznici','columns':['PrevoznikID','Naziv','Vozac','Registracija','Aktivan']},
                              {'table':'tblOtkup','columns':['OtkupID','Datum']}]}
            ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            def xml_for(name,columns):
                root=ET.Element('{%s}table'%ns, {'name':name})
                cs=ET.SubElement(root,'{%s}tableColumns'%ns)
                for i,c in enumerate(columns,1):
                    ET.SubElement(cs,'{%s}tableColumn'%ns, {'id':str(i),'name':c})
                return ET.tostring(root,encoding='utf-8')
            with zipfile.ZipFile(wb,'w') as z:
                z.writestr('xl/tables/table1.xml',xml_for('tblPrevoznici',
                  schema['tables'][0]['columns']+list(m.RUNTIME_AUDIT_SUFFIX)))
                z.writestr('xl/tables/table2.xml',xml_for('tblOtkup',schema['tables'][1]['columns']))
                z.writestr('xl/vbaProject.bin',bytes.fromhex('D0CF11E0A1B11AE1')+b'test')
            d=m.workbook_structure(wb,schema)
            self.assertEqual(d['canonical_columns'],7)
            self.assertEqual(d['runtime_columns'],4)
            self.assertEqual(d['runtime_extensions'],{'tblPrevoznici':list(m.RUNTIME_AUDIT_SUFFIX)})

    def test_reject_table_removed_or_new_unexpected_table(self):
        expected={'tblPrevoznici':['ID']}
        with self.assertRaisesRegex(RuntimeError,'table set mismatch'):
            m.canonical_table_structure(expected,{})
        with self.assertRaisesRegex(RuntimeError,'table set mismatch'):
            m.canonical_table_structure(expected,{'tblPrevoznici':['ID'],'tblOther':['ID']})

if __name__ == '__main__': unittest.main()
