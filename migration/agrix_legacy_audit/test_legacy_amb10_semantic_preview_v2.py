import json
import hashlib
from pathlib import Path
import tempfile
from unittest import TestCase
from decimal import Decimal
from legacy_audit import AuditError
from legacy_membership_v2 import index_ids
from legacy_amb10_forensics_v2 import DOC_PARENT
from legacy_amb10_shadow_v2 import build_report
from legacy_amb10_forensics_v2 import analyze
import legacy_amb10_storno_lineage_v2 as storno
import legacy_amb10_semantic_preview_v2 as m


def e(dt,from_type,from_id,to_type,to_id,quantity='10',doc='T1',ordinals='1'):
    return {'document_type':dt,'document_id':doc,'from_type':from_type,'from_id':from_id,
            'to_type':to_type,'to_id':to_id,'tip_ambalaze':'T','quantity':quantity,
            'source_row_ordinals':ordinals}

def r(typ='Otpremnica',doc='T1',direction='Izlaz',owner='Stanica',eid='S1',driver='D1',q='10',n=1,storno=''):
    return {'AmbID':f'AMB-{n:05d}', 'DokumentTip':typ, 'DokumentID':doc,
            'Kolicina':q,'TipAmbalaze':'T','Smer':direction,'EntitetTip':owner,
            'EntitetID':eid,'VozacID':driver,'Datum':'45200','Stornirano':storno}

def parents():
    tab={
      'SOURCE__tblOtkup':[{'OtkupID':'O1','KooperantID':'K1','StanicaID':'S1','KolAmbalaze':'10','KolAmbIzdata':'10','TipAmbalaze':'T','Stornirano':''}],
      'SOURCE__tblOtpremnica':[{'OtpremnicaID':'T1','StanicaID':'S1','VozacID':'D1','KolAmbalaze':'10','TipAmbalaze':'T','Stornirano':''},
                               {'OtpremnicaID':'T2','StanicaID':'S1','VozacID':'D1','KolAmbalaze':'10','TipAmbalaze':'T','Stornirano':''}],
      'tblPrijemnica':[{'PrijemnicaID':'P1','KupacID':'C1','VozacID':'D1','KolAmbalaze':'10','KolAmbVracena':'5','TipAmbalaze':'T','Stornirano':''}],
      'tblStanice':[{'StanicaID':'S1'}], 'tblKooperanti':[{'KooperantID':'K1'}],
      'tblKupci':[{'KupacID':'C1'}], 'tblVozaci':[{'VozacID':'D1'}], 'tblTipAmbalaze':[{'TipAmbalaze':'T'}],
    }
    indices={name:index_ids(tab[name],pk,name) for name,pk in DOC_PARENT.values()}
    return tab,indices

class SemanticTests(TestCase):
    def test_otkup_can_propose_uz_robu(self):
        _,p=parents();originals=[r('Otkup','O1','Izlaz','Kooperant','K1','',n=1),r('Otkup','O1','Ulaz','Stanica','S1','',n=2)]
        x=m.classify_event(e('Otkup','Kooperant','K1','Stanica','S1',doc='O1',ordinals='1;2'),originals,p)
        self.assertEqual(x['VrstaKretanja_PROPOSAL'],'AMBALAZA_UZ_ROBU')
        self.assertEqual(x['evidence_status'],'DIRECTION_KIND_PROPOSED_ONLY')
        self.assertEqual(x['canonical_AMB10_import_authorized'],'NO')
    def test_issue_to_cooperator_always_hold_ownership(self):
        _,p=parents();originals=[r('OM-Izlaz-Koop','O1','Izlaz','Stanica','S1','',n=1),r('OM-Izlaz-Koop','O1','Ulaz','Kooperant','K1','',n=2)]
        x=m.classify_event(e('OM-Izlaz-Koop','Stanica','S1','Kooperant','K1',doc='O1',ordinals='1;2'),originals,p)
        self.assertEqual(x['evidence_status'],'HOLD_OWNERSHIP')
        self.assertEqual(x['VrstaKretanja_PROPOSAL'],'UNASSIGNED')
    def test_otp_station_driver_proposed(self):
        _,p=parents();src=[r()]
        x=m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1'),src,p)
        self.assertEqual(x['evidence_status'],'DIRECTION_KIND_PROPOSED_ONLY')
    def test_otp_wrong_driver_hold(self):
        _,p=parents();src=[r(driver='WRONG')]
        x=m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1'),src,p)
        self.assertEqual(x['evidence_status'],'HOLD_INTEGRITY')
        self.assertIn('DRIVER_NOT_EXACT',x['reason'])
    def test_otp_quantity_hold(self):
        _,p=parents();src=[r(q='11')]
        x=m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1'),src,p)
        self.assertEqual(x['evidence_status'],'HOLD_INTEGRITY')
    def test_receipt_ulaz_and_return(self):
        _,p=parents();src=[r('Prijemnica','P1','Ulaz','Kupac','C1',n=1),r('Prijemnica','P1','Izlaz','Kupac','C1',q='5',n=2)]
        first=m.classify_event(e('Prijemnica','Vozac','D1','Kupac','C1',doc='P1'),src,p)
        second=m.classify_event(e('Prijemnica','Kupac','C1','Vozac','D1',quantity='5',doc='P1',ordinals='2'),src,p)
        self.assertEqual(first['VrstaKretanja_PROPOSAL'],'AMBALAZA_UZ_ROBU')
        self.assertEqual(second['VrstaKretanja_PROPOSAL'],'POVRAT_PRAZNE')
    def test_storned_source_fail_closed(self):
        _,p=parents();src=[r(storno='Da')]
        with self.assertRaisesRegex(AuditError,'STORNO source'):
            m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1'),src,p)
    def test_mismatched_parent_status_hold(self):
        tab,p=parents();tab['SOURCE__tblOtpremnica'][0]['Stornirano']='Da'
        x=m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1'),[r()],p)
        self.assertEqual(x['evidence_status'],'HOLD_INTEGRITY')
        self.assertIn('PARENT_NOT_ACTIVE',x['reason'])
    def test_malformed_source_ordinal_fail(self):
        _,p=parents()
        with self.assertRaises(AuditError):m.classify_event(e('Otpremnica','Stanica','S1','Vozac','D1',ordinals='7000'),[r()],p)
    def test_e2e_synthetic_7815_and_baseline_guard(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);stage=root/'stage';stage.mkdir()
            tab,_=parents();tab['SOURCE__tblOtpremnica'][0]['Stornirano']='Da';tab['SOURCE__tblOtpremnica'][0]['KolAmbalaze']='500'
            rows=[r(doc='T2',n=n) for n in range(1,7816)]
            rows[699]=r(doc='T1',n=700,q='500')
            for x in rows:x['DokumentTIP']=x.pop('DokumentTip')
            tab['UNCONVERTED__tblAmbalaza']=rows
            tab['UNCONVERTED__tblStornoVeze']=[]
            manifest={'source_sha256':m.SOURCE_SHA,'status':'STAGED_ONLY_NO_XLSM_CREATED','schema_sha256':'SYNTHETIC_SCHEMA',
                'canonical_stage_counts':{k:len(v) for k,v in tab.items()}}
            (stage/'agrix_stage_summary.json').write_text(json.dumps(manifest),encoding='utf-8')
            for name,items in tab.items():
                (stage/(name+'.jsonl')).write_text(''.join(json.dumps(i)+'\n' for i in items),encoding='utf-8')
            baseline,raw=analyze(stage);shadow,events,review,_=build_report(raw,baseline)
            sh=root/'shadow';sh.mkdir();(sh/'agrix_amb10_shadow_summary.json').write_text(json.dumps(shadow),encoding='utf-8')
            storno_report,*_=storno.build(stage)
            st=root/'storno';st.mkdir();(st/'agrix_storno_lineage_summary.json').write_text(json.dumps(storno_report),encoding='utf-8')
            self.assertEqual(len(events),7814)
            digest=hashlib.sha256((stage/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest()
            out=root/'semantic'
            self.assertEqual(m.cli(['--stage',str(stage),'--shadow',str(sh),'--storno',str(st),'--out',str(out)]),0)
            rep=json.loads((out/'agrix_amb10_semantic_preview_summary.json').read_text())
            self.assertEqual(rep['physical_events_examined'],7814)
            self.assertEqual(rep['quarantine_active_source_rows'],1)
            self.assertTrue(rep['q00700_state_mismatch_isolated_from_draft'])
            self.assertEqual(rep['draft_evidence_status_counts'],{'DIRECTION_KIND_PROPOSED_ONLY':7814})
            self.assertNotIn('T2', (out/'agrix_amb10_semantic_preview_summary.json').read_text())
            self.assertTrue((out/'amb10_direction_and_kind_PROPOSALS_PRIVATE.csv').is_file())
            self.assertEqual(hashlib.sha256((stage/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest(),digest)
            with self.assertRaisesRegex(AuditError,'Output exists'):
                m.cli(['--stage',str(stage),'--shadow',str(sh),'--storno',str(st),'--out',str(out)])
            storno_report['focus']['parent_state']='ACTIVE'
            (st/'agrix_storno_lineage_summary.json').write_text(json.dumps(storno_report),encoding='utf-8')
            with self.assertRaisesRegex(AuditError,'Q-00700'):
                m.run(stage,sh,st)
