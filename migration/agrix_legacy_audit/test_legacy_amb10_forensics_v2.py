from collections import Counter
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from legacy_audit import AuditError
from legacy_membership_v2 import index_ids
import legacy_amb10_forensics_v2 as m


def leg(id='A1',typ='Otkup',doc='O1',owner='Kooperant',eid='K1',direction='Izlaz',
        qty='10',driver='',storno='',amb='12/1',date='45200'):
    return dict(AmbID=id,Datum=date,TipAmbalaze=amb,Kolicina=qty,Smer=direction,
                EntitetID=eid,EntitetTip=owner,VozacID=driver,DokumentID=doc,
                DokumentTip=typ,Stornirano=storno)


def source(legs):
    tables={
        'UNCONVERTED__tblAmbalaza':legs,
        'SOURCE__tblOtkup':[dict(OtkupID='O1',KooperantID='K1',StanicaID='S1',KolAmbalaze='10',KolAmbIzdata='10',Stornirano='')],
        'SOURCE__tblOtpremnica':[dict(OtpremnicaID='T1',StanicaID='S1',VozacID='D1',KolAmbalaze='10')],
        'tblPrijemnica':[dict(PrijemnicaID='P1',KupacID='C1',VozacID='D1')],
        'tblKooperanti':[dict(KooperantID='K1')],
        'tblStanice':[dict(StanicaID='S1')],
        'tblKupci':[dict(KupacID='C1')],
        'tblVozaci':[dict(VozacID='D1')],
        'tblTipAmbalaze':[dict(TipAmbalaze='12/1')],
    }
    idx={name:index_ids(tables[name],pk,name) for name,pk in m.DOC_PARENT.values()}
    idx.update({name:index_ids(tables[name],pk,name) for name,pk in m.MASTER.values()})
    mf={'schema_sha256':'dummy','canonical_stage_counts':{'UNCONVERTED__tblAmbalaza':len(legs)}}
    return mf,tables,idx


def analyze(legs):
    with patch.object(m,'inputs',return_value=source(legs)):
        return m.analyze(Path('/synthetic'))


class AMBForensics(unittest.TestCase):
    def test_otkup_double_legs(self):
        r,raw=analyze([leg(),leg(id='A2',owner='Stanica',eid='S1',direction='Ulaz')])
        self.assertEqual(r['pair_group_counts'],{'MIRRORED_PAIR_EVIDENCE_ONLY':1})
        self.assertEqual(r['rows_with_structural_review'],0)
        self.assertEqual(r['legacy_ambalaza_rows'],2)
        self.assertEqual(len(raw),2)
    def test_om_issue_pair(self):
        r,_=analyze([leg(typ='OM-Izlaz-Koop',owner='Kooperant',direction='Ulaz'),
                     leg(id='A2',typ='OM-Izlaz-Koop',owner='Stanica',eid='S1',direction='Izlaz')])
        self.assertEqual(r['pair_group_counts']['MIRRORED_PAIR_EVIDENCE_ONLY'],1)
    def test_pair_amount_mismatch_flagged(self):
        r,_=analyze([leg(),leg(id='A2',owner='Stanica',eid='S1',direction='Ulaz',qty='9')])
        self.assertEqual(r['pair_group_counts']['UNRESOLVED_DIFFERENT_DATE_QTY_OR_TYPE'],1)
        self.assertEqual(r['rows_with_structural_review'],2)
    def test_pair_missing_leg(self):
        r,_=analyze([leg()])
        self.assertEqual(r['pair_group_counts']['UNRESOLVED_LEG_COUNT'],1)
    def test_pair_owner_context(self):
        r,_=analyze([leg(),leg(id='A2',owner='Stanica',eid='S999',direction='Ulaz')])
        self.assertIn('UNRESOLVED_OWNER_CONTEXT',r['pair_group_counts'])
    def test_pair_mixed_storno(self):
        r,_=analyze([leg(),leg(id='A2',owner='Stanica',eid='S1',direction='Ulaz',storno='Da')])
        self.assertEqual(r['pair_group_counts']['UNRESOLVED_MIXED_STORNO_STATE'],1)
        self.assertEqual(r['legacy_state_counts'],{'ACTIVE':1,'STORNO':1})
    def test_driver_inverse_with_good_parent(self):
        r,_=analyze([leg(typ='Otpremnica',doc='T1',owner='Stanica',eid='S1',driver='D1')])
        self.assertEqual(r['single_leg_evidence_counts']['DRIVER_INVERSE_EVIDENCE_ONLY'],1)
        self.assertEqual(r['rows_with_derived_driver_balance'],1)
    def test_missing_doc_and_bad_driver(self):
        r,_=analyze([leg(typ='Prijemnica',doc='',owner='Kupac',eid='C1',direction='Ulaz',driver='D999')])
        self.assertEqual(r['missing_document_id_ordinals'],[1])
        self.assertEqual(r['single_leg_evidence_counts']['UNRESOLVED_PARENT_OR_OWNER'],1)
        self.assertEqual(r['review_reason_counts']['DRIVER_NOT_IN_MASTER'],1)
    def test_no_fake_opening_balance(self):
        r,_=analyze([leg()])
        self.assertEqual(r['negative_legacy_ending_balance_accounts']['account_type_counts'],{'Kooperant':1})
        self.assertNotIn('opening_entries',r)
    def test_unknown_type_pseudonym_no_public_values(self):
        r,_=analyze([leg(typ='HIDDEN_INTERNAL_CUSTOMER_NAME')])
        data=json.dumps(r)
        self.assertNotIn('HIDDEN_INTERNAL_CUSTOMER_NAME',data)
        self.assertIn('OTHER_',data)
    def test_csv_formula_escape(self):
        self.assertEqual(m.safe_value('=1+1'),"'=1+1")
        self.assertEqual(m.safe_value('@danger'),"'@danger")
        with tempfile.TemporaryDirectory() as p:
            f=Path(p)/'test.csv';m.private_csv(f,[{'id':'=IMPORTXML(1)'}],['id'])
            self.assertIn("'=IMPORTXML",f.read_text(encoding='utf-8-sig'))
    def test_real_legacy_dokumenttip_spelling_read_from_jsonl(self):
        # Realistic 7815-row stage with actual legacy DokumentTIP column.
        with tempfile.TemporaryDirectory() as p:
            stage = Path(p)
            manifest, tables, _ = source([
                {**leg(), 'DokumentTIP': 'Otkup'},
                {**leg(id='A2',owner='Stanica',eid='S1',direction='Ulaz'),
                 'DokumentTIP': 'Otkup'}
            ])
            for row in tables['UNCONVERTED__tblAmbalaza']:
                row.pop('DokumentTip', None)
            records=tables['UNCONVERTED__tblAmbalaza']
            tables['UNCONVERTED__tblAmbalaza']=records + [
                dict(records[1], AmbID='A'+str(i)) for i in range(3,7816)
            ]
            tables['tblOtkup']=[dict(OtkupID='O1')]
            tables['tblOtpremnica']=[dict(OtpremnicaID='T1')]
            manifest.update({'source_sha256':m.SOURCE_SHA,
                             'status':'STAGED_ONLY_NO_XLSM_CREATED',
                             'canonical_stage_counts':{k:len(v) for k,v in tables.items()}})
            (stage/'agrix_stage_summary.json').write_text(json.dumps(manifest),encoding='utf-8')
            for table, records in tables.items():
                (stage/(table+'.jsonl')).write_text(
                    ''.join(json.dumps(r)+'\n' for r in records),encoding='utf-8')
            mf,normalized,idx=m.inputs(stage)
            self.assertEqual(normalized['UNCONVERTED__tblAmbalaza'][0]['DokumentTip'],'Otkup')
            self.assertEqual(len(normalized['UNCONVERTED__tblAmbalaza']),7815)
            self.assertIn('KolAmbalaze',normalized['SOURCE__tblOtkup'][0])

    def test_required_storno_column_fails_closed(self):
        row=leg();row.pop('Stornirano')
        with self.assertRaisesRegex(AuditError,'Stornirano'):
            m.normalize_legacy_leg(row,1)

    def test_conflicting_case_variants_fails_closed(self):
        row=leg();row['DokumentTIP']='Otpremnica'
        with self.assertRaisesRegex(AuditError,'conflicting case variants'):
            m.normalize_legacy_leg(row,1)

    def test_forensic_parent_quantity_uses_unmodified_source(self):
        legs=[leg(qty='11'),leg(id='A2',owner='Stanica',eid='S1',direction='Ulaz',qty='11')]
        r,_=analyze(legs)
        self.assertEqual(r['review_reason_counts']['LEG_DOCUMENT_QUANTITY_MISMATCH_REVIEW'],2)

    def test_wrong_sha_refused(self):
        with tempfile.TemporaryDirectory() as p:
            stage=Path(p)
            (stage/'agrix_stage_summary.json').write_text(json.dumps({'source_sha256':'OTHER'}))
            with self.assertRaises(AuditError):m.inputs(stage)
    def test_stage_count_refused(self):
        with tempfile.TemporaryDirectory() as p:
            stage=Path(p)
            (stage/'agrix_stage_summary.json').write_text(json.dumps({'source_sha256':m.SOURCE_SHA,
                'status':'STAGED_ONLY_NO_XLSM_CREATED',
                'canonical_stage_counts':{'UNCONVERTED__tblAmbalaza':2}}))
            with self.assertRaises(AuditError):m.inputs(stage)

if __name__=='__main__': unittest.main()
