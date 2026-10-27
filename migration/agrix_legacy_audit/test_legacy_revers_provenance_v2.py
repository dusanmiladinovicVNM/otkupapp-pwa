import unittest
from legacy_audit import AuditError
from legacy_revers_provenance_v2 import classify, revers_id, revers_pairs

BASE={
 'DokumentTip':'OM-Izlaz-Koop','DokumentID':'OTK-12','ReversID':'',
 'EntitetTip':'Kooperant','EntitetID':'K-1','Kolicina':'20',
 'TipAmbalaze':'Gajba','Smer':'Ulaz','Stornirano':'',
}
OTK={'OtkupID':'OTK-12','KooperantID':'K-1','StanicaID':'S-1',
     'KolAmbIzdata':'20','TipAmbalaze':'Gajba','Stornirano':''}
class ReversTest(unittest.TestCase):
 def test_otkup_exact(self):
    self.assertEqual(classify(BASE,{'OTK-12':(1,OTK)})[0],'OTKUP_LINK_EXACT')
 def test_revers_with_pk(self):
    a=dict(BASE,DokumentID='REV-BROJ-10',ReversID='REV-991')
    self.assertEqual(classify(a,{})[0],'REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF')
 def test_label_only(self):
    a=dict(BASE,DokumentID='REV-BROJ-10')
    self.assertEqual(classify(a,{})[0],'BUSINESS_LABEL_ONLY_NO_OPAQUE_ID')
 def test_parent_present_wrong_ownership_no_guess(self):
    a=dict(BASE,EntitetID='WRONG',ReversID='REV-44')
    self.assertEqual(classify(a,{'OTK-12':(1,OTK)})[0],'OTKUP_ID_PRESENT_CONTEXT_REVIEW')
 def test_storno_mismatch(self):
    a=dict(BASE,Stornirano='Da')
    self.assertIn('OTKUP_STATE_MISMATCH',classify(a,{'OTK-12':(1,OTK)})[2])
 def test_quantity_mismatch(self):
    a=dict(BASE,Kolicina='21')
    self.assertIn('OTKUP_QTY_MISMATCH',classify(a,{'OTK-12':(1,OTK)})[2])
 def test_revers_case_alias(self):
    self.assertEqual(revers_id({'reversid':'R-1'}),'R-1')
 def test_revers_case_alias_conflict(self):
    with self.assertRaises(AuditError):revers_id({'ReversID':'R-1','reversid':'R-2'})
 def test_revers_legs(self):
    a=dict(BASE,DokumentID='REVNUM-1',ReversID='REV-1',
           Datum='2026-08-01',source_row_ordinal=1,
           provenance='REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF')
    b=dict(a,source_row_ordinal=2,EntitetTip='Stanica',EntitetID='S-1',Smer='Izlaz')
    s,p=revers_pairs([a,b]);self.assertEqual(s[1],'RECIPROCAL_WITHIN_REVERS_ID_EVIDENCE_ONLY');self.assertEqual(p[1],2)
 def test_different_revers_not_paired(self):
    a=dict(BASE,DokumentID='REVNUM-1',ReversID='REV-1',
           Datum='2026-08-01',source_row_ordinal=1,
           provenance='REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF')
    b=dict(a,source_row_ordinal=2,ReversID='REV-2',EntitetTip='Stanica',Smer='Izlaz')
    s,p=revers_pairs([a,b]);self.assertEqual(p,{})
 def test_lone_firma_has_no_fake_pair(self):
    a=dict(BASE,DokumentTip='OM-Ulaz-Firma',DokumentID='REVNUM-2',ReversID='REV-2',
           Datum='2026-08-01',source_row_ordinal=1,
           provenance='REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF')
    s,p=revers_pairs([a]);self.assertEqual(p,{})

if __name__=='__main__':unittest.main()

class FullCLIRegression(unittest.TestCase):
 def test_real_shape_7815_and_no_staged_mutations(self):
    import json, hashlib, tempfile
    from pathlib import Path
    import legacy_revers_provenance_v2 as m
    from legacy_stage_v2 import SOURCE_SHA
    with tempfile.TemporaryDirectory() as tmp:
        root=Path(tmp); st=root/'stage'; st.mkdir(); qt=root/'quarantine';qt.mkdir()
        master={
          'SOURCE__tblOtkup':[{'OtkupID':'OTK-1','KooperantID':'K1','StanicaID':'S1','KolAmbIzdata':'20','TipAmbalaze':'T','Stornirano':''}],
          'SOURCE__tblOtpremnica':[{'OtpremnicaID':'OTP-1','StanicaID':'S1','VozacID':'V1','KolAmbalaze':'500','Stornirano':'Da'},
                                     {'OtpremnicaID':'OTP-2','StanicaID':'S1','VozacID':'V1','KolAmbalaze':'10','Stornirano':''}],
          'tblPrijemnica':[], 'tblStanice':[{'StanicaID':'S1'}],
          'tblKooperanti':[{'KooperantID':'K1'}], 'tblKupci':[{'KupacID':'C1'}],
          'tblVozaci':[{'VozacID':'V1'}], 'tblTipAmbalaze':[{'TipAmbalaze':'T'}]
        }
        def row(n,typ='Otpremnica',doc='OTP-2',direction='Izlaz',et='Stanica',ei='S1',qty='10',rev=''):
            return dict(AmbID=f'AMB-{n:05d}',Datum='45500',DokumentTIP=typ,DokumentID=doc,ReversID=rev,
                        Kolicina=qty,TipAmbalaze='T',Smer=direction,EntitetTip=et,EntitetID=ei,
                        VozacID='V1',Stornirano='')
        rows=[row(i) for i in range(1,7816)]
        rows[699]=row(700,'Otpremnica','OTP-1',qty='500')
        pool=[i for i in range(1,7816) if i!=700][:97]
        ptr=0
        # 18 exact paired old revers events, labelled by number, not OtkupID.
        for g in range(18):
            for x in range(2):
                n=pool[ptr];ptr+=1
                rows[n-1]=row(n,'OM-Izlaz-Koop',f'BROJ-{g}',direction=('Ulaz' if x==0 else 'Izlaz'),
                    et=('Kooperant' if x==0 else 'Stanica'),ei=('K1' if x==0 else 'S1'),qty='20',rev=f'REV-{g}')
        # 25 old reverse receipts, also two legs, which original forensics did not model.
        for g in range(25):
            for x in range(2):
                n=pool[ptr];ptr+=1
                rows[n-1]=row(n,'OM-Ulaz-Koop',f'BROJ-U-{g}',direction=('Izlaz' if x==0 else 'Ulaz'),
                    et=('Kooperant' if x==0 else 'Stanica'),ei=('K1' if x==0 else 'S1'),qty='20',rev=f'REV-U-{g}')
        for x in range(5):
            n=pool[ptr];ptr+=1
            rows[n-1]=row(n,'OM-Ulaz-Firma' if x<3 else 'OM-Izlaz-Firma','BROJ-F',rev=f'REV-F{x}')
        for x in range(6):
            n=pool[ptr];ptr+=1
            rows[n-1]=row(n,'Kupci-Otpremnica','BROJ-UNKNOWN')
        self.assertEqual(ptr,97)
        master['UNCONVERTED__tblAmbalaza']=rows
        mf={'source_sha256':SOURCE_SHA,'schema_sha256':'SYNTHETIC_REVERSES',
            'status':'STAGED_ONLY_NO_XLSM_CREATED','canonical_stage_counts':{k:len(v) for k,v in master.items()}}
        (st/'agrix_stage_summary.json').write_text(json.dumps(mf),encoding='utf-8')
        for name,rs in master.items():
            (st/(name+'.jsonl')).write_text(''.join(json.dumps(x)+'\n' for x in rs),encoding='utf-8')
        prior={'format':'MIG-04.4','source_sha256':SOURCE_SHA,'schema_sha256':'SYNTHETIC_REVERSES',
            'quarantined_active_rows':98,'document_type_categories':{
                'OM-Izlaz-Koop':36,'OM-Ulaz-Koop':50,'OM-Ulaz-Firma':3,'OM-Izlaz-Firma':2,
                'OTHER_OR_UNKNOWN':6,'Otpremnica':1}}
        (qt/'agrix_amb10_quarantine_summary.json').write_text(json.dumps(prior),encoding='utf-8')
        before=hashlib.sha256((st/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest()
        out=root/'output'
        self.assertEqual(m.cli(['--stage',str(st),'--quarantine',str(qt),'--out',str(out)]),0)
        result=json.loads((out/'agrix_revers_provenance_summary.json').read_text(encoding='utf-8'))
        self.assertEqual(result['om_rows_examined'],91)
        self.assertEqual(result['om_quarantined_source_rows'],91)
        self.assertEqual(result['quarantined_om_legs_with_exact_revers_id_reciprocal'],86)
        self.assertEqual(result['quarantined_om_provenance'],{'REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF':91})
        self.assertFalse(result['ready_for_import'])
        self.assertEqual(before,hashlib.sha256((st/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest())
        with self.assertRaises(AuditError):m.cli(['--stage',str(st),'--quarantine',str(qt),'--out',str(out)])
