#!/usr/bin/env python3
"""MIG-06.1 — traceable, NON-AUTHORIZING historical writer-purpose proposals.

No legal-title inference, no AMB10 Import row, no synthetic ReversID, no XLSM.
Existing negative account/type balances are legitimate and never corrected.

Only purpose-intent is proposed, from OLD DOCUMENT WRITER and exact source
physical roles: Otkup.KolAmbIzdata issued -> nominal IZDATA_PRAZNA, Revers
OM-Izlaz-Koop/OM-Ulaz-Koop -> nominal IZDATA_PRAZNA/POVRAT_PRAZNE,
OM-Izlaz-Firma/OM-Ulaz-Firma -> PRENOS_INTERNO. None proves packaging ownership,
third-party obligations, historical cancellation timestamp or legal title.
"""
from __future__ import annotations
import argparse, hashlib, json, os, sys
from collections import Counter
from pathlib import Path
from legacy_audit import AuditError, trim
import legacy_amb10_identity_bridge_v2 as bridge
import legacy_amb10_chronology_preview_v2 as chrono
import legacy_amb10_forensics_v2 as f

VERSION='MIG-06.1'
PRIVATE='AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl'
PROVENANCE={
 'om-izlaz-koop':('Stanica','Kooperant','IZDATA_PRAZNA','LEGACY_REVERS_ISSUANCE_INTENT'),
 'om-ulaz-koop':('Kooperant','Stanica','POVRAT_PRAZNE','LEGACY_REVERS_RETURN_INTENT'),
 'om-izlaz-firma':('Stanica','Vozac','PRENOS_INTERNO','LEGACY_REVERS_STATION_DRIVER_TRANSFER'),
 'om-ulaz-firma':('Vozac','Stanica','PRENOS_INTERNO','LEGACY_REVERS_DRIVER_STATION_TRANSFER'),
}

def canon(obj):
 return json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'

def decision(event,identity,drycase):
 """Fail closed if scope/physical evidence diverges from prior pinned tools."""
 if (event['physical_case_id']!=identity['physical_case_id'] or
     identity['physical_case_id']!=drycase['physical_case_id'] or
     event['source_row_ordinals']!=identity['source_row_ordinals'] or
     event['source_row_ordinals']!=drycase['source_row_ordinals']):
  raise AuditError('Physical provenance / identity / dryrun mismatch')
 if identity['import_approved'] is not False or drycase['import_approved'] is not False or event['canonical_AMB10_import_authorized']!='NO':
  raise AuditError('Unexpected upstream import authorization')
 cat=identity['identity_status']; doc=f.key(event['legacy_DokumentTIP'])
 row={'physical_case_id':event['physical_case_id'],
      'source_row_ordinals':event['source_row_ordinals'],
      'identity_status':cat,
      'verified_existing_DokumentTIP':identity.get('verified_existing_DokumentTIP'),
      'verified_existing_DokumentID':identity.get('verified_existing_DokumentID'),
      'provisional_group_token':identity.get('provisional_group_token'),
      'old_doc_type':event['legacy_DokumentTIP'],
      'old_doc_label':event['legacy_DokumentID'],
      'quantity':event['Kolicina'],'type':event['TipAmbalaze'],
      'from_account_type':event['OdNalogTip'],'from_account_id':event['OdNalogID'],
      'to_account_type':event['NaNalogTip'],'to_account_id':event['NaNalogID'],
      'proposed_kind':None,'business_intent_evidence':None,
      'ownership_or_third_party_title':'UNPROVEN',
      'possible_foreign_return':'NOT_DETERMINED',
      'canonical_document_id':None,'storno_od':None,
      'approval_status':'PROPOSAL_ONLY_NOT_CANONICAL','import_approved':False}
 if cat==bridge.CATEGORIES[0]:
  prev=drycase.get('movement_preview')
  if not prev or drycase['blocking_reasons'] or prev['VrstaKretanja']!=event['VrstaKretanja_PROPOSAL']:
   raise AuditError('Already-evidenced goods kind was changed')
  if identity['verified_existing_DokumentID']!=prev['DokumentID'] or identity['verified_existing_DokumentTIP']!=prev['DokumentTIP']:
   raise AuditError('Goods parent documentary identity drift')
  row['proposed_kind']=prev['VrstaKretanja'];row['business_intent_evidence']='MIG05_6_EXACT_GOODS_PARENT_WRITER_PROPOSAL'
 elif cat==bridge.CATEGORIES[1]:
  if (doc!='om-izlaz-koop' or event['event_provenance']!='MIG05_ORIGINAL_PHYSICAL' or
      event['evidence_status']!='HOLD_OWNERSHIP' or
      event['OdNalogTip']!='Stanica' or event['NaNalogTip']!='Kooperant' or
      identity['verified_existing_DokumentTIP']!='Otkup' or
      identity['verified_existing_DokumentID']!=trim(event['legacy_DokumentID']) or
      drycase['movement_preview'] is not None):
   raise AuditError('Old Otkup issuance writer contract not proven')
  row['proposed_kind']='IZDATA_PRAZNA'
  row['business_intent_evidence']='OTKUP_KOL_AMB_IZDATA_OLD_WRITER_ISSUANCE_INTENT_ONLY'
  row['possible_foreign_return']='VRACANJE_TUDJE_AMBALAZE_REQUIRES_INDEPENDENT_LIABILITY_EVIDENCE'
 elif cat==bridge.CATEGORIES[2]:
  if (event['event_provenance']!='MIG05_2_RECOVERED_REVERS_PHYSICAL' or
      not identity.get('provisional_group_token') or identity.get('canonical_new_doc_id') or
      drycase['movement_preview'] is not None or
      doc not in PROVENANCE):
   raise AuditError('Scoped old Revers contract unproven')
  f_type,t_type,kind,why=PROVENANCE[doc]
  if event['OdNalogTip']!=f_type or event['NaNalogTip']!=t_type:
   raise AuditError('Revers old direction does not match document type')
  row['proposed_kind']=kind;row['business_intent_evidence']=why
  if kind=='IZDATA_PRAZNA':
   row['possible_foreign_return']='VRACANJE_TUDJE_AMBALAZE_REQUIRES_INDEPENDENT_LIABILITY_EVIDENCE'
 elif cat==bridge.CATEGORIES[3]:
  if event['event_provenance']!='MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY' or drycase['movement_preview'] is not None:
   raise AuditError('Unknown document shape unexpectedly changed')
  row['business_intent_evidence']='UNMODELLED_DOCUMENT_NO_KIND_PROVEN'
 else:raise AuditError('Unknown identity category: '+str(cat))
 return row

def build(stage,expanded,seven,dryrun,storno_archive,identity_dir):
 # Recalculate parent chain and exact PRIVATE evidence; refuse summary-only trust.
 summary,identities=bridge.build(stage,expanded,seven,dryrun,storno_archive)
 expected=json.loads((identity_dir/'agrix_amb10_identity_bridge_summary.json').read_text(encoding='utf-8'))
 h=bridge.assert_records(identity_dir/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl',identities)
 if expected.get('identity_bridge_evidence_sha256')!=h:
  raise AuditError('MIG-06.0 identity PRIVATE sha mismatch')
 for key,v in summary.items():
  if expected.get(key)!=v:raise AuditError('MIG-06.0 identity summary drift: '+key)
 _,events,_=chrono._verify_inputs(stage,expanded,seven)
 _,drys,parity,_=bridge.dry.build(stage,expanded,seven)
 if len(events)!=4675 or len(drys)!=4675 or len(identities)!=4675:
  raise AuditError('Physical event count changed')
 rows=[decision(e,i,d) for e,i,d in zip(events,identities,drys)]
 counts=Counter(r['identity_status'] for r in rows)
 kinds=Counter(r['proposed_kind'] or 'UNASSIGNED' for r in rows)
 origins=Counter(r['business_intent_evidence'] for r in rows)
 required={bridge.CATEGORIES[0]:3485,bridge.CATEGORIES[1]:1136,
          bridge.CATEGORIES[2]:48,bridge.CATEGORIES[3]:6}
 if dict(counts)!=required:raise AuditError('Historical document classification mismatch: '+str(counts))
 if kinds['UNASSIGNED']!=6 or sum(kinds.values())!=4675:
  raise AuditError('New proposals did not cover exactly 4669 scoped events')
 if not all(x['unexplained']=='0' for x in parity) or summary['legacy_source_negative_accounts_accepted']!=35:
  raise AuditError('Legacy balance parity or negative balance count changed')
 if any(r['import_approved'] or r['canonical_document_id'] or r['storno_od'] for r in rows):
  raise AuditError('No canonical import writing allowed')
 out={'format':VERSION,'status':'LEGACY_WRITER_BUSINESS_INTENT_EVIDENCE_ONLY_NO_AMB10_IMPORT',
      'source_sha256':summary['source_sha256'],'schema_sha256':summary['schema_sha256'],
      'identity_bridge_evidence_sha256':h,
      'active_physical_events':len(rows),
      'events_with_proposed_kind_not_approved':len(rows)-kinds['UNASSIGNED'],
      'events_without_kind_proposal':kinds['UNASSIGNED'],
      'existing_goods_parent_kind_proposals':counts[bridge.CATEGORIES[0]],
      'existing_otkup_issued_empty_intent_proposals':counts[bridge.CATEGORIES[1]],
      'scoped_old_revers_kind_proposals':counts[bridge.CATEGORIES[2]],
      'legacy_unmodelled_document_kind_held':counts[bridge.CATEGORIES[3]],
      'proposed_kind_counts':dict(sorted(kinds.items())),
      'business_intent_evidence_counts':dict(sorted(origins.items())),
      'remaining_events_without_existing_source_doc_pk':summary['remaining_events_without_existing_source_doc_pk'],
      'source_negative_accounts_accepted_unchanged':35,
      'unexplained_legacy_balance_residuals':0,
      'q00700_excluded_unchanged':True,
      'historical_storno_rows_in_separate_archive':335,
      'foreign_ownership_verified':False,'deficit_cover_generated':False,
      'canonical_amb_id_allocated':False,'canonical_amb_doc_allocated':False,
      'storno_od_allocated':False,'import_approved':False,
      'notes':['Existing Otkup.KolAmbIzdata and old OM-leg direction provide intended business action, NOT legal packaging title.',
               'Nominal IZDATA_PRAZNA remains conditional: real VRACANJE_TUDJE_AMBALAZE requires independent proven liability.',
               'Scoped Revers business numbers are NOT opaque existing canonical document IDs.',
               '6 unknown document events and Q-00700 remain held. Historic cancellation dates not invented.',
               'Negative source balances accepted exactly; no synthetic opening, obligation or monetary adjustments.',
               'Private source IDs, account balances and amounts remain only in local PRIVATE evidence.']}
 return out,rows

def render(s):
 fields=('active_physical_events','events_with_proposed_kind_not_approved',
 'events_without_kind_proposal','existing_otkup_issued_empty_intent_proposals',
 'scoped_old_revers_kind_proposals','proposed_kind_counts',
 'remaining_events_without_existing_source_doc_pk','source_negative_accounts_accepted_unchanged',
 'unexplained_legacy_balance_residuals','historical_storno_rows_in_separate_archive','import_approved')
 return 'AGRIX MIG-06.1 - HISTORICAL BUSINESS INTENT PROPOSALS / NO IMPORT\n'+''.join(f'{k}: {json.dumps(s[k],ensure_ascii=False,sort_keys=True)}\n' for k in fields)

def cli(argv=None):
 p=argparse.ArgumentParser(description='Old writer -> business purpose proposals; does not approve title or import')
 for arg in ('stage','expanded','seven','dryrun','storno-archive','identity','out'):
  p.add_argument('--'+arg,type=Path,required=True)
 a=p.parse_args(argv)
 inp=[getattr(a,x.replace('-','_')).resolve() for x in ('stage','expanded','seven','dryrun','storno-archive','identity')]
 out=a.out.resolve()
 if not all(x.is_dir() for x in inp) or out.exists() or any(out==x or out in x.parents or x in out.parents for x in inp):
  raise AuditError('Missing input or overlapping/existing destination')
 s,rows=build(*inp)
 out.mkdir(parents=True,exist_ok=False)
 try:
  dest=out/PRIVATE
  with dest.open('w',encoding='utf-8',newline='\n') as fh:
   for r in rows:fh.write(canon(r))
  os.chmod(dest,0o600)
  s['business_intent_evidence_sha256']=hashlib.sha256(dest.read_bytes()).hexdigest()
  (out/'agrix_business_intent_summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
  (out/'agrix_business_intent_summary.txt').write_text(render(s),encoding='utf-8')
 except BaseException:
  for path in out.iterdir():
   if path.is_file():path.unlink()
  out.rmdir();raise
 print(render(s));print('Share summaries ONLY; no target XLSM or canonical AMB-10 written.')
 return 0

if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
  print('MIG-06.1 STOP:',str(exc),file=sys.stderr);raise SystemExit(2)
