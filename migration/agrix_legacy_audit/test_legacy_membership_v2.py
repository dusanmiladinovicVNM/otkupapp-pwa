import json
from pathlib import Path
import tempfile
import unittest
from legacy_audit import AuditError
from legacy_membership_v2 import (link_otkup_to_otp, link_otp_to_zbirna,
                                  receipt_parent_report, process_stage, dec,
                                  quantity_compare)

HASH='a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'

def otk(**kw):
    r={'OtkupID':'O1','OtpremnicaID':'T1','StanicaID':'S','VrstaVoca':'Malina',
       'SortaVoca':'W','TipAmbalaze':'12/1','Klasa':'I','Kolicina':'100','Stornirano':''}
    r.update(kw);return r

def otp(**kw):
    r={'OtpremnicaID':'T1','StanicaID':'S','VozacID':'V','BrojZbirne':'B',
       'VrstaVoca':'Malina','SortaVoca':'W','TipAmbalaze':'12/1',
       'Klasa':'I','Kolicina':'100','Stornirano':''}
    r.update(kw);return r

def zbr(**kw):
    r={'ZbirnaID':'Z1','BrojZbirne':'B','VozacID':'V','KupacID':'K',
       'VrstaVoca':'Malina','SortaVoca':'W','TipAmbalaze':'12/1',
       'Klasa':'I','UkupnoKolicina':'100','Stornirano':''}
    r.update(kw);return r

class Tests(unittest.TestCase):
    def test_valid_exact_id(self):
        result,counts,rec,issues=link_otkup_to_otp(HASH,[otk()],[otp()])
        self.assertEqual(len(result),1)
        self.assertFalse(issues)
        self.assertEqual(rec['OTP_LINKS_QTY_VERIFIED'],1)
    def test_unknown_exact_id(self):
        result,counts,rec,issues=link_otkup_to_otp(HASH,[otk(OtpremnicaID='T404')],[otp()])
        self.assertFalse(result)
        self.assertEqual(counts['DANGLING_OTP_ID'],1)
    def test_station_mismatch(self):
        result,counts,rec,issues=link_otkup_to_otp(HASH,[otk(StanicaID='OTHER')],[otp()])
        self.assertFalse(result)
        self.assertEqual(counts['BUSINESS_CONTEXT_MISMATCH'],1)
        self.assertEqual(issues[0]['field_names'],['StanicaID'])
    def test_quantity_disagreement(self):
        result,counts,rec,issues=link_otkup_to_otp(HASH,[otk(Kolicina='99')],[otp()])
        self.assertFalse(result)
        self.assertEqual(rec['OTP_HEADERS_QTY_MISMATCH'],1)
    def test_one_zbr_by_driver_and_class(self):
        result,stats,rec,issues=link_otp_to_zbirna(HASH,[otp()],[zbr()])
        self.assertEqual(len(result),1)
        self.assertEqual(rec['ZBR_LINKS_QTY_VERIFIED'],1)
        self.assertFalse(issues)
    def test_dont_merge_wrong_driver(self):
        result,stats,rec,issues=link_otp_to_zbirna(HASH,[otp(VozacID='X')],[zbr()])
        self.assertFalse(result)
        self.assertEqual(stats['ZBR_SCOPE_NOT_FOUND'],1)
    def test_dont_merge_duplicate_same_scope(self):
        result,stats,rec,issues=link_otp_to_zbirna(HASH,[otp()],[zbr(),zbr(ZbirnaID='Z2')])
        self.assertFalse(result)
        self.assertEqual(stats['ZBR_SCOPE_AMBIGUOUS'],1)
    def test_storno_header_requires_review(self):
        result,stats,rec,issues=link_otp_to_zbirna(HASH,[otp()],[zbr(Stornirano='Da')])
        self.assertFalse(result)
        self.assertEqual(rec['ZBR_HEADERS_HISTORICAL_STORNO_REVIEW'],1)
    def test_receipt_buyer_ownership(self):
        rp,issues=receipt_parent_report([{'BrojZbirne':'B','VozacID':'V','KupacID':'OTHER','Klasa':'I'}],[zbr()])
        self.assertEqual(rp['SCOPE_MISMATCH'],1)
    def test_receipt_missing_business_no(self):
        rp,issues=receipt_parent_report([{'BrojZbirne':'NO','VozacID':'V','KupacID':'K','Klasa':'I'}],[zbr()])
        self.assertEqual(rp['ZBR_NUMBER_MISSING'],1)
        self.assertEqual(issues[0]['prijemnica_row'],1)
    def test_quantity_by_class_no_merge(self):
        yes,reason=quantity_compare([zbr()],[otp(Klasa='II')],'UkupnoKolicina','Kolicina',dec('0.01'))
        self.assertFalse(yes)
    def test_no_stage_pii_summary(self):
        with tempfile.TemporaryDirectory() as temp:
            stage=Path(temp)
            d={'format':'MIG-02.0','status':'STAGED_ONLY_NO_XLSM_CREATED','source_sha256':HASH,
               'schema_sha256':'abcd', 'source_row_counts':{'tblOtkup':1,'tblOtpremnica':1,
                'tblZbirna':1,'tblPrijemnica':1}}
            (stage/'agrix_stage_summary.json').write_text(json.dumps(d))
            for name,rows in {'SOURCE__tblOtkup':[otk()], 'SOURCE__tblOtpremnica':[otp()],
                              'SOURCE__tblZbirna':[zbr()],
                              'tblPrijemnica':[{'PrijemnicaID':'P1','BrojZbirne':'B','KupacID':'K','VozacID':'V','Klasa':'I'}]}.items():
                (stage/(name+'.jsonl')).write_text('\n'.join(json.dumps(r) for r in rows)+'\n')
            private,summary=process_stage(stage)
            self.assertEqual(len(private['CANDIDATE__tblOtpremnicaIzvori']),1)
            self.assertEqual(len(private['CANDIDATE__tblZbirnaIzvori']),1)
            public=json.dumps(summary)
            for needle in ('OTK-REAL-123','Malina','O1','T1','Z1'):
                self.assertNotIn('"'+needle+'"',public)
    def test_hash_mismatch_fails_closed(self):
        with tempfile.TemporaryDirectory() as temp:
            (Path(temp)/'agrix_stage_summary.json').write_text(json.dumps({'format':'MIG-02.0','status':'STAGED_ONLY_NO_XLSM_CREATED','source_sha256':'other'}))
            with self.assertRaises(AuditError):process_stage(Path(temp))

if __name__=='__main__': unittest.main()
