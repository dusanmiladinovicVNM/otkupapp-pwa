import json
from pathlib import Path
import tempfile
import unittest

from legacy_audit import AuditError
from legacy_membership_v2 import process_stage
from legacy_membership_guard_v2 import inspect, ratio_band
from legacy_stage_v2 import SOURCE_SHA


def otk(id='O1', parent='T1', qty='100', storno='', **kw):
    d={'OtkupID':id,'OtpremnicaID':parent,'StanicaID':'S','VrstaVoca':'Malina',
       'SortaVoca':'W','TipAmbalaze':'12/1','Klasa':'I','Kolicina':qty,'Stornirano':storno}
    d.update(kw);return d

def otp(id='T1', qty='100', storno='', **kw):
    d={'OtpremnicaID':id,'StanicaID':'S','VozacID':'V','BrojZbirne':'B',
       'VrstaVoca':'Malina','SortaVoca':'W','TipAmbalaze':'12/1',
       'Klasa':'I','Kolicina':qty,'Stornirano':storno}
    d.update(kw);return d

def zbr(id='Z1', qty='100', storno='', **kw):
    d={'ZbirnaID':id,'BrojZbirne':'B','VozacID':'V','KupacID':'K',
       'VrstaVoca':'Malina','SortaVoca':'W','TipAmbalaze':'12/1','Klasa':'I',
       'UkupnoKolicina':qty,'Stornirano':storno}
    d.update(kw);return d

def receipt(number='B',**kw):
    d={'PrijemnicaID':'P1','BrojZbirne':number,'VozacID':'V','KupacID':'K',
       'Klasa':'I','SortaVoca':'W','VrstaVoca':'Malina','TipAmbalaze':'12/1'}
    d.update(kw);return d


def create_stage(base,otks,otps,zbrs,prijs):
    stage=base/'stage';member=base/'member';stage.mkdir();member.mkdir()
    rows={'SOURCE__tblOtkup':otks,'SOURCE__tblOtpremnica':otps,
          'SOURCE__tblZbirna':zbrs,'tblPrijemnica':prijs}
    manifest={'format':'MIG-02.0','status':'STAGED_ONLY_NO_XLSM_CREATED',
              'source_sha256':SOURCE_SHA,'schema_sha256':'synthetic',
              'source_row_counts':{key.removeprefix('SOURCE__'):len(vals) for key,vals in rows.items()}}
    (stage/'agrix_stage_summary.json').write_text(json.dumps(manifest))
    for name,vals in rows.items():
        (stage/(name+'.jsonl')).write_text(''.join(json.dumps(r)+'\n' for r in vals))
    candidates,summary=process_stage(stage)
    (member/'agrix_membership_summary.json').write_text(json.dumps(summary))
    for name,vals in candidates.items():
        (member/(name+'.jsonl')).write_text(''.join(json.dumps(r)+'\n' for r in vals))
    return stage,member

class GuardTests(unittest.TestCase):
    def test_storno_otkup_candidate_is_unsafe(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk(),otk(id='O2',qty='20',storno='Da')],[otp()],[zbr()],[receipt()])
            clean,rep=inspect(s,m)
            self.assertEqual(len(clean['PROVISIONAL_ACTIVE__tblOtpremnicaIzvori']),1)
            self.assertEqual(rep['candidate_safety']['CANDIDATE__tblOtpremnicaIzvori']['historical_excluded_from_active_only'],1)
    def test_storno_otpremnica_candidate_is_unsafe(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk()],[otp(),otp(id='T2',qty='20',storno='Da')],[zbr()],[receipt()])
            clean,rep=inspect(s,m)
            self.assertEqual(len(clean['PROVISIONAL_ACTIVE__tblZbirnaIzvori']),1)
            self.assertEqual(rep['candidate_safety']['CANDIDATE__tblZbirnaIzvori']['historical_excluded_from_active_only'],1)
    def test_no_source_headers_audited_and_retained(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk()],[otp(),otp(id='T2')],[zbr()],[receipt()])
            _,rep=inspect(s,m)
            self.assertEqual(rep['otp_reconciliation']['counts']['NO_VERIFIED_OTKUP_SOURCE'],1)
            self.assertIn(2,[v['otp_row'] for v in rep['otp_reconciliation']['review']])
    def test_quantity_mismatch_categorized_not_overwritten(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk(qty='97')],[otp()],[zbr()],[receipt()])
            _,rep=inspect(s,m)
            self.assertEqual(rep['otp_reconciliation']['counts']['QTY_MISMATCH'],1)
            self.assertEqual(rep['otp_reconciliation']['review'][0]['direction'],'HEADER_GREATER')
    def test_casefold_number_explains_mig02_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk()],[otp()],[zbr(BrojZbirne='B')],[receipt(number='b')])
            _,rep=inspect(s,m)
            self.assertEqual(rep['receipt_reconciliation']['counts']['NUMBER_RESOLVED_BY_CASEFOLD'],1)
    def test_public_report_redacts_private_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk()],[otp()],[zbr()],[receipt()])
            _,rep=inspect(s,m)
            raw=json.dumps(rep)
            for value in ['Malina','"O1"','"T1"','"Z1"','100']:
                # Number 100 can incidentally occur in counts, so skip unsafe comparison
                if value=='100':continue
                self.assertNotIn(value,raw)
    def test_mismatched_input_hash_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)
            s,m=create_stage(p,[otk()],[otp()],[zbr()],[receipt()])
            mf=json.loads((m/'agrix_membership_summary.json').read_text());mf['source_sha256']='wrong'
            (m/'agrix_membership_summary.json').write_text(json.dumps(mf))
            with self.assertRaises(AuditError):inspect(s,m)
    def test_ratio_band(self):
        self.assertEqual(ratio_band('100','95'),('LE_5_PERCENT','HEADER_GREATER'))
        self.assertEqual(ratio_band('100','110'),('LE_20_PERCENT','SOURCES_GREATER'))

if __name__=='__main__':unittest.main()
