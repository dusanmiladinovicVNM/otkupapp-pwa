import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from legacy_audit import AuditError
import legacy_amb10_revers_document_plan_v2 as m

class TestMIG062(unittest.TestCase):
 def source(self,number='7/011026',day='2026-10-01',station='ST-1',kind='IZDATA_PRAZNA',doc='OM-Izlaz-Koop',fr='Stanica',to='Kooperant'):
  scope={'station':station,'day':day,'label':number}
  return dict(identity_status=m.bridge.CATEGORIES[2],approval_status='PROPOSAL_ONLY_NOT_CANONICAL',
    import_approved=False,legacy_scope_station_day_label=scope,old_doc_label=number,
    provisional_group_token='MIG-GROUP-1',old_doc_type=doc,from_account_type=fr,
    from_account_id=station if fr=='Stanica' else 'KOOP-1',to_account_type=to,
    to_account_id=station if to=='Stanica' else 'KOOP-1', proposed_kind=kind,
    source_row_ordinals='700;701',physical_case_id='PHY-1',quantity='25',type='GAJBA')
 def test_new_draft_id_stable_and_opaque(self):
  a=self.source();h,l=m.make_proposal(a,'source-hash')
  self.assertEqual(h['target_header_draft']['AmbDokID'],m.make_proposal(a,'source-hash')[0]['target_header_draft']['AmbDokID'])
  self.assertTrue(h['target_header_draft']['AmbDokID'].startswith('ADK-'))
  self.assertEqual(len(h['target_header_draft']['AmbDokID']),36)
  self.assertEqual(h['target_header_draft']['BrojDokumenta'],'7/011026')
  self.assertEqual(h['target_header_draft']['BrojOwnerTip'],'Stanica')
  self.assertFalse(h['import_approved']);self.assertFalse(l['import_approved'])
  self.assertEqual(l['draft_DokumentTIP'],'AmbalazaDokument')
 def test_distinct_scopes_or_source_create_distinct_ids(self):
  a=self.source();b=self.source(number='007/011026')
  self.assertNotEqual(m.make_proposal(a,'h')[0]['target_header_draft']['AmbDokID'],m.make_proposal(b,'h')[0]['target_header_draft']['AmbDokID'])
  self.assertNotEqual(m.make_proposal(a,'h')[0]['target_header_draft']['AmbDokID'],m.make_proposal(a,'other')[0]['target_header_draft']['AmbDokID'])
 def test_revers_return_and_firma(self):
  r=self.source(kind='POVRAT_PRAZNE',doc='OM-Ulaz-Koop',fr='Kooperant',to='Stanica')
  self.assertEqual(m.make_proposal(r,'h')[1]['draft_VrstaKretanja'],'POVRAT_PRAZNE')
  r=self.source(kind='PRENOS_INTERNO',doc='OM-Izlaz-Firma',fr='Stanica',to='Vozac');r['to_account_id']='VOZ-1'
  self.assertEqual(m.make_proposal(r,'h')[1]['draft_VrstaKretanja'],'PRENOS_INTERNO')
 def test_tampered_station_kind_and_label_rejected(self):
  bads=[]
  a=self.source();a['legacy_scope_station_day_label']['station']='ST-2';bads.append(a)
  a=self.source();a['proposed_kind']='VRACANJE_TUDJE_AMBALAZE';bads.append(a)
  a=self.source();a['old_doc_label']='DIFFERENT';bads.append(a)
  a=self.source();a['import_approved']=True;bads.append(a)
  a=self.source();a['legacy_scope_station_day_label']['label']='';bads.append(a)
  for row in bads:
   with self.subTest(row=row),self.assertRaises(AuditError):m.make_proposal(row,'h')
 def test_no_overwrite_or_source_write(self):
  with tempfile.TemporaryDirectory() as temp:
   args=[]
   for arg in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','out'):
    args.extend(['--'+arg,temp])
   with self.assertRaises(AuditError):m.cli(args)

 def test_full_synthetic_4675_integrated_no_unknown_fabrication(self):
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);paths={k:root/k for k in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent')}
   for p in paths.values():p.mkdir()
   ids=[];events=[]
   for j in range(4675):
    if j<3485:cat=m.bridge.CATEGORIES[0]
    elif j<4621:cat=m.bridge.CATEGORIES[1]
    elif j<4669:cat=m.bridge.CATEGORIES[2]
    else:cat=m.bridge.CATEGORIES[3]
    phys='PHY-%05d'%(j+1);src=str(j+1)
    i={'physical_case_id':phys,'source_row_ordinals':src,'identity_status':cat,
       'verified_existing_DokumentID':'OTK-%05d'%j if j<4621 else None,
       'provisional_group_token':None,'legacy_scope_station_day_label':None}
    ev={'physical_case_id':phys,'source_row_ordinals':src,'identity_status':cat,
       'proposed_kind':'AMBALAZA_UZ_ROBU' if j<3485 else 'IZDATA_PRAZNA',
       'approval_status':'PROPOSAL_ONLY_NOT_CANONICAL','import_approved':False,
       'provisional_group_token':None}
    if cat==m.bridge.CATEGORIES[2]:
     k=j-4621;label='REV-'+str(k)
     i['legacy_scope_station_day_label']={'station':'ST-1','day':'2026-10-01','label':label}
     i['provisional_group_token']='MIG-GROUP-'+str(k)
     ev.update(old_doc_type='OM-Izlaz-Koop',old_doc_label=label,
       proposed_kind='IZDATA_PRAZNA',provisional_group_token=i['provisional_group_token'],
       from_account_type='Stanica',from_account_id='ST-1',to_account_type='Kooperant',to_account_id='K-1',
       quantity='5',type='GAJBA')
    if cat==m.bridge.CATEGORIES[3]:ev['proposed_kind']=None;ev['old_doc_type']='UNKNOWN';ev['old_doc_label']=''
    ids.append(i);events.append(ev)
   base={'source_sha256':'source-hash','schema_sha256':'schema-hash'}
   biz=dict(base,active_physical_events=4675)
   for name,records in [(paths['identity']/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl',ids),
                         (paths['business-intent']/m.intent.PRIVATE,events)]:
    name.write_text(''.join(m.canon(z) for z in records),encoding='utf-8',newline='\n')
   (paths['identity']/'agrix_amb10_identity_bridge_summary.json').write_text(json.dumps(dict(base,identity_bridge_evidence_sha256=hashlib.sha256((paths['identity']/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl').read_bytes()).hexdigest())))
   (paths['business-intent']/'agrix_business_intent_summary.json').write_text(json.dumps(dict(biz,business_intent_evidence_sha256=hashlib.sha256((paths['business-intent']/m.intent.PRIVATE).read_bytes()).hexdigest())))
   argv=[]
   for k,p in paths.items():argv+=['--'+k,str(p)]
   argv+=['--out',str(root/'plan')]
   with patch.object(m.bridge,'build',return_value=(base,ids)),patch.object(m.intent,'build',return_value=(biz,events)):
    self.assertEqual(m.cli(argv),0)
    s=json.loads((root/'plan'/'agrix_revers_document_plan_summary.json').read_text())
    self.assertEqual(s['draft_new_historical_revers_headers'],48)
    self.assertEqual(s['unknown_document_holds_untouched'],6)
    self.assertEqual(s['events_linked_to_existing_goods_parent'],4621)
    self.assertEqual(s['draft_header_rows_import_approved'],0)
    headers=[json.loads(l) for l in (root/'plan'/m.PRIVATE_HEADERS).read_text().splitlines()]
    self.assertEqual(len({h['target_header_draft']['AmbDokID'] for h in headers}),48)
    self.assertEqual(len((root/'plan'/m.PRIVATE_UNKNOWNS).read_text().splitlines()),6)
    with self.assertRaises(AuditError):m.cli(argv)
    # A changed or extra private record invalidates the upstream chain.
    path=paths['business-intent']/m.intent.PRIVATE
    path.write_text(path.read_text(encoding='utf-8')+'\n',encoding='utf-8')
    argv[-1]=str(root/'rejected')
    with self.assertRaises(AuditError):m.cli(argv)
    self.assertFalse((root/'rejected').exists())

if __name__=='__main__':unittest.main()
