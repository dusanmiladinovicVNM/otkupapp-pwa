import json
from pathlib import Path
import tempfile
import unittest
from decimal import Decimal
from legacy_audit import AuditError
from legacy_membership_v2 import process_stage
from legacy_membership_guard_v2 import inspect
from legacy_reconciliation_dossier_v2 import (make_report,csv_text,private_csv)
from test_legacy_membership_guard_v2 import otk,otp,zbr,receipt,create_stage

class DossierTests(unittest.TestCase):
    def _fixture(self,otks,otps,zbrs,prijs):
        tmp=tempfile.TemporaryDirectory();p=Path(tmp.name)
        stage,member=create_stage(p,otks,otps,zbrs,prijs)
        g=p/'guard';g.mkdir()
        candidates,rep=inspect(stage,member)
        (g/'agrix_membership_guard_summary.json').write_text(json.dumps(rep))
        for name,data in candidates.items():
            (g/(name+'.jsonl')).write_text(''.join(json.dumps(r)+'\n' for r in data))
        return tmp,stage,g
    def test_active_candidacy_rechecked(self):
        t,s,g=self._fixture([otk(),otk(id='O2',qty='9',storno='Da')],[otp()],[zbr()],[receipt()])
        try:
            report,_=make_report(s,g)
            self.assertEqual(report['verified_candidate_integrity']['otp_active_candidate_links'],1)
            self.assertEqual(report['verified_candidate_integrity']['zbr_headers_rechecked'],1)
        finally:t.cleanup()
    def test_unlinked_candidates_are_only_suggestions(self):
        t,s,g=self._fixture([otk(parent='',qty='100',BrojOtpremnice='T100')],[otp(id='T1',BrojOtpremnice='T100')],[zbr()],[receipt()])
        try:
            report,private=make_report(s,g)
            self.assertEqual(report['unlinked_otp_number_suggestions_not_approved']['HAS_NUMBER_STATION_CLASS_MATCH'],1)
            self.assertEqual(private['otpremnica_review_PRIVATE.csv'][0]['operator_decision'],'REVIEW_REQUIRED')
        finally:t.cleanup()
    def test_mismatch_has_signed_delta(self):
        t,s,g=self._fixture([otk(qty='98')],[otp(qty='100')],[zbr(qty='100')],[receipt()])
        try:
            report,pr=make_report(s,g)
            x=pr['otpremnica_review_PRIVATE.csv'][0]
            self.assertEqual(x['active_minus_otp_kg'],Decimal('-2'))
            self.assertEqual(report['quantity_mismatch_buckets']['LE_5_PERCENT'],1)
        finally:t.cleanup()
    def test_verified_zbr_link_can_have_unverified_upstream(self):
        t,s,g=self._fixture([otk(qty='95')],[otp(qty='100')],[zbr(qty='100')],[receipt()])
        try:
            report,pr=make_report(s,g)
            self.assertEqual(report['transitive_chain']['ZBR_LINK_WITH_UNVERIFIED_OTKUP_CHAIN'],1)
            self.assertEqual(pr['zbirna_upstream_review_PRIVATE.csv'][0]['reason'],'QTY_MISMATCH')
        finally:t.cleanup()
    def test_verified_chain_has_no_upstream_review(self):
        t,s,g=self._fixture([otk()],[otp()],[zbr()],[receipt()])
        try:
            report,pr=make_report(s,g)
            self.assertEqual(report['transitive_chain']['ZBR_LINK_WITH_VERIFIED_OTKUP_CHAIN'],1)
            self.assertFalse(pr['zbirna_upstream_review_PRIVATE.csv'])
        finally:t.cleanup()
    def test_casefold_receipt_has_explicit_parent(self):
        t,s,g=self._fixture([otk()],[otp()],[zbr()],[receipt(number='b')])
        try:
            report,pr=make_report(s,g)
            self.assertEqual(report['receipt_casefold_ordinal'],1)
            self.assertEqual(pr['prijemnica_review_PRIVATE.csv'][0]['exact_identity_rows'],'1')
        finally:t.cleanup()
    def test_private_csv_formula_escape(self):
        self.assertEqual(csv_text('=HYPERLINK("x")'),"'=HYPERLINK(\"x\")")
        self.assertEqual(csv_text('+cmd'),"'+cmd")
        self.assertEqual(csv_text('-12.50'),'-12.50')
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'private.csv';private_csv(path,[{'value':'=1+1'}],['value'])
            self.assertIn("'=1+1",path.read_text(encoding='utf-8-sig'))
    def test_guard_count_drift_rejected(self):
        t,s,g=self._fixture([otk()],[otp()],[zbr()],[receipt()])
        try:
            v=json.loads((g/'agrix_membership_guard_summary.json').read_text())
            v['candidate_safety']['CANDIDATE__tblOtpremnicaIzvori']['provisional_active_count']=999
            (g/'agrix_membership_guard_summary.json').write_text(json.dumps(v))
            with self.assertRaises(AuditError):make_report(s,g)
        finally:t.cleanup()
    def test_bad_candidate_quantity_rejected(self):
        t,s,g=self._fixture([otk()],[otp()],[zbr()],[receipt()])
        try:
            p=g/'PROVISIONAL_ACTIVE__tblOtpremnicaIzvori.jsonl'
            dat=json.loads(p.read_text().splitlines()[0]);dat['OtkupID']='UNKNOWN'
            p.write_text(json.dumps(dat)+'\n')
            with self.assertRaises(AuditError):make_report(s,g)
        finally:t.cleanup()

if __name__=='__main__':unittest.main()
