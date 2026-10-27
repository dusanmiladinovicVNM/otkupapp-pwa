"""Build fails before Excel opens for old AMB-10 workbook schema/commit."""
import sys
import tempfile
import json
import unittest
from pathlib import Path
from unittest.mock import patch
import subprocess

sys.path.insert(0,str(Path(__file__).resolve().parent))
import agrix_generate_and_preflight as m

NEW_AMB = ("AmbID", "Datum", "TipAmbalaze", "Kolicina", "DokumentID", "DokumentTIP",
                  "CreatedAt", "CreatedBy", "ModifiedAt", "ModifiedBy", "OdNalogTip",
                  "OdNalogID", "NaNalogTip", "NaNalogID", "VrstaKretanja", "StornoOd")
NEW_DOC = ("AmbDokID", "Vrsta", "BrojDokumenta", "Datum", "BrojOwnerTip",
                  "BrojOwnerID", "Napomena", "Stornirano", "CreatedAt", "CreatedBy",
                  "ModifiedAt", "ModifiedBy", "IspravkaOdID", "ZamenjenSaID")

class TestGuard(unittest.TestCase):
    def setup(self,base,legacy=False):
        (base/'schema').mkdir()
        amb=("AmbID", "Datum", "TipAmbalaze", "Kolicina","Smer","EntitetID","EntitetTip",
             "VozacID","DokumentID","DokumentTIP","Stornirano","CreatedAt","CreatedBy",
             "ModifiedAt","ModifiedBy","ReversID") if legacy else NEW_AMB
        s={'tables':[{'table':'tblAmbalaza','columns':amb},
                     {'table':'tblAmbalazaDokument','columns':NEW_DOC}]}
        (base/'schema'/'schema.json').write_text(json.dumps(s))
    def test_clean_stale_head_is_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td);self.setup(b)
            with patch.object(subprocess,'run',return_value=subprocess.CompletedProcess([],0,'newhead\n','')):
                with self.assertRaisesRegex(RuntimeError,'nije origin/main'):
                    m.assert_repo_migration_compatible(b,'oldhead',False)
    def test_same_head_legacy_schema_is_rejected(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td);self.setup(b,legacy=True)
            with patch.object(subprocess,'run',return_value=subprocess.CompletedProcess([],0,'samehead\n','')):
                with self.assertRaisesRegex(RuntimeError,'nije na podrzanoj'):
                    m.assert_repo_migration_compatible(b,'samehead',False)
    def test_new_head_new_schema_passes(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td);self.setup(b)
            with patch.object(subprocess,'run',return_value=subprocess.CompletedProcess([],0,'samehead\n','')):
                x=m.assert_repo_migration_compatible(b,'samehead',False)
            self.assertEqual(x['target_schema_column_count'],30)
    def test_dirty_checkout_fails_closed(self):
        with tempfile.TemporaryDirectory() as td:
            b=Path(td);self.setup(b)
            with self.assertRaisesRegex(RuntimeError,'izmenjene'):
                m.assert_repo_migration_compatible(b,'samehead',True)

if __name__=='__main__':unittest.main()
