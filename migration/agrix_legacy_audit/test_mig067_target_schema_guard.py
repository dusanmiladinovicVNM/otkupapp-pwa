"""Regression: source workbook schema hash is NOT target Git schema SHA."""
import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import legacy_amb10_target_adapter_v2 as adapter
import legacy_amb10_workbook_preflight_v2 as pf
from test_legacy_amb10_workbook_preflight_v2 import test_template
from legacy_audit import AuditError


class TestTargetCanonicalSchema(unittest.TestCase):
    def test_target_schema_differs_from_legacy_source_pin_but_matches_actual_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td)
            file=b/'blank.xlsm'
            test_template(file)
            spec={'tables':[
                {'table':'tblAmbalaza','columns':list(adapter.linked.CANONICAL_COLUMNS)},
                {'table':'tblAmbalazaDokument','columns':sorted(adapter.REQUIRED_HEADERS)},
                {'table':'tblOtkup','columns':['OtkupID','Opis']},
                {'table':'tblOtpremnica','columns':['OtpremnicaID']},
                {'table':'tblPrijemnica','columns':['PrijemnicaID']},
            ]}
            schema=b/'schema.json'
            schema.write_text(json.dumps(spec),encoding='utf-8')
            target_sha=hashlib.sha256(schema.read_bytes()).hexdigest()
            self.assertNotEqual(target_sha,adapter.ready.SCHEMA_SHA)
            inspection=pf.inspect_target(file,schema_file=schema)
            self.assertEqual(inspection['full_schema_check']['status'],'FULL_TARGET_SCHEMA_MATCH')
            self.assertEqual(inspection['full_schema_check']['target_schema_sha256'],target_sha)
            self.assertEqual(inspection['full_schema_check']['source_workbook_schema_sha256'],adapter.ready.SCHEMA_SHA)

    def test_target_schema_incompatible_missing_column_stops(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td)
            file=b/'blank.xlsm';test_template(file)
            spec={'tables':[{'table':'tblAmbalaza','columns':['AmbID']}]}
            schema=b/'schema.json';schema.write_text(json.dumps(spec),encoding='utf-8')
            with self.assertRaisesRegex(AuditError,'FULL target schema'):
                pf.inspect_target(file,schema_file=schema)

if __name__ == '__main__':unittest.main()
