"""MIG-07.11: local-only scratch recovery; no customer file is used."""
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
import agrix_mig07_goods_copy as m
from legacy_audit import AuditError


class TestMIG0711ScratchRecovery(unittest.TestCase):
    def _inputs(self, root):
        stage=root/'stage'
        stage.mkdir()
        template=root/'master.xlsm'
        template.write_bytes(b'dummy blank master fixture')
        schemafile=root/'schema.json'
        schemafile.write_text('{}')
        preflight=root/'preflight.json'
        preflight.write_text('{}')
        for table in m.TABLE_ORDER:
            (stage/(table+'.jsonl')).write_text('', encoding='utf-8')
        scratch_dir=root/'previous_attempt'
        scratch_dir.mkdir()
        scratch=scratch_dir/'_BUILD_IN_PROGRESS.xlsm'
        with zipfile.ZipFile(scratch,'w') as z:
            z.writestr('xl/vbaProject.bin', b'fixture-vba')
        return stage,template,schemafile,preflight,scratch

    def _run(self, stage, template, schemafile, preflight, scratch, out):
        schema={t:['ID'] for t in m.TABLE_ORDER}
        data={t:[] for t in m.TABLE_ORDER}
        kinds={t:['text'] for t in m.TABLE_ORDER}
        counts={t:0 for t in m.TABLE_ORDER}
        with (patch.object(m,'get_schema',return_value=schema),
              patch.object(m,'baseline_check',return_value={}),
              patch.object(m,'stage_load',return_value=data),
              patch.object(m,'prepared',return_value=(data,kinds,counts)),
              patch.object(m,'verify_data_roundtrip',return_value=counts) as audit):
            result=m.cli(['--stage',str(stage),'--template',str(template),
                '--preflight',str(preflight),'--schema',str(schemafile),
                '--out',str(out),'--check-scratch',str(scratch)])
        return result,audit

    def test_offline_check_does_not_modify_scratch_or_master_and_never_claims_vba(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)
            stage,template,schemafile,preflight,scratch=self._inputs(root)
            master_before=m.sha256(template)
            scratch_before=m.sha256(scratch)
            out=root/'reports'
            result,audit=self._run(stage,template,schemafile,preflight,scratch,out)
            self.assertEqual(result,0)
            audit.assert_called_once()
            self.assertEqual(master_before,m.sha256(template))
            self.assertEqual(scratch_before,m.sha256(scratch))
            report=json.loads((out/'agrix_scratch_data_audit.json').read_text())
            self.assertEqual(report['status'],'DATA_ROUNDTRIP_PASS_VBA_NOT_CHECKED')
            self.assertFalse(report['vba_preservation_checked'])
            self.assertFalse(report['production_import_approved'])
            self.assertFalse((out/'AgriX_MIG07_UNVALIDATED_COPY.xlsm').exists())

    def test_incomplete_scratch_is_rejected_before_any_output(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)
            stage,template,schemafile,preflight,scratch=self._inputs(root)
            scratch.write_bytes(b'not a ZIP file')
            out=root/'reports'
            with self.assertRaisesRegex(AuditError,'not a valid OOXML ZIP'):
                self._run(stage,template,schemafile,preflight,scratch,out)
            self.assertFalse(out.exists())

    def test_unexpected_scratch_basename_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)
            stage,template,schemafile,preflight,scratch=self._inputs(root)
            renamed=scratch.with_name('AgriX_MIG07_UNVALIDATED_COPY.xlsm')
            scratch.rename(renamed)
            with self.assertRaisesRegex(AuditError,'EXISTING _BUILD_IN_PROGRESS'):
                self._run(stage,template,schemafile,preflight,renamed,root/'report')

if __name__=='__main__':
    unittest.main()
