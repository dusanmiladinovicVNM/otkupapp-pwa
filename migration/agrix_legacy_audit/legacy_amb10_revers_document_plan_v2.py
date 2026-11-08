#!/usr/bin/env python3
"""MIG-06.2: deterministic migration-only AmbalazaDokument HEADER PLAN.

48 verified old station/day/label Revers scopes -> 48 NEW draft IDs. No live
workbook writes, no transaction replay, no ownership or foreign-title approval.
Original historical Revers had business labels, not canonical AmbDokID; this
script never claims otherwise. Unknown documents stay on HOLD. SHA-pinned inputs.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import os
import sys
import uuid
from collections import Counter
from pathlib import Path
from legacy_audit import AuditError, trim
import legacy_amb10_business_intent_v2 as intent
import legacy_amb10_identity_bridge_v2 as bridge
import legacy_amb10_forensics_v2 as f

VERSION='MIG-06.2'
PRIVATE_HEADERS='AMB10_REVERS_HEADER_DRAFT_PRIVATE.jsonl'
PRIVATE_LINKS='AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl'
PRIVATE_UNKNOWNS='AMB10_UNMODELLED_DOCUMENT_HOLD_PRIVATE.jsonl'
# Stable project-specific namespace; this is a proposed NEW migration identity,
# never a historical ReversID or an already-existing target AmbDokID.
NAMESPACE=uuid.UUID('c533ad66-bfad-5519-9efe-e36427427b8d')


def canon(obj):
 return json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'


def proposed_id(source_sha,scope):
 if set(scope)!={'station','day','label'} or not all(isinstance(v,str) and trim(v) for v in scope.values()):
  raise AuditError('Invalid or incomplete old Revers scope')
 payload=canon({'source':source_sha,'source_scope':scope,'format':VERSION})
 return 'ADK-'+uuid.uuid5(NAMESPACE,payload).hex.upper()


def make_proposal(r,src_sha):
 if r['identity_status']!=bridge.CATEGORIES[2] or r['approval_status']!='PROPOSAL_ONLY_NOT_CANONICAL' or r['import_approved'] is not False:
  raise AuditError('Not an approved historical Revers physical-evidence class')
 scope=r.get('legacy_scope_station_day_label')
 if not isinstance(scope,dict):raise AuditError('MIG-06.1 private Revers scope missing')
 # The previous file contains token but not the scope. MIG-06.0 has it;
 # caller must add only the identity-verified, byte-checked scope.
 scope=scope.copy()
 if trim(r.get('old_doc_label'))!=trim(scope['label']) or r['provisional_group_token'] is None:
  raise AuditError('Revers scope business label/token mismatch')
 typ=f.key(r['old_doc_type'])
 directions={
  'om-izlaz-koop':('Stanica','Kooperant','IZDATA_PRAZNA'),
  'om-ulaz-koop':('Kooperant','Stanica','POVRAT_PRAZNE'),
  'om-izlaz-firma':('Stanica','Vozac','PRENOS_INTERNO'),
  'om-ulaz-firma':('Vozac','Stanica','PRENOS_INTERNO'),
 }
 if typ not in directions:raise AuditError('Unexpected legacy Revers document type')
 src_t,dst_t,k=directions[typ]
 if (r['from_account_type'],r['to_account_type'],r['proposed_kind'])!=(src_t,dst_t,k):
  raise AuditError('Old Revers physical direction differs from planned business kind')
 station = r['from_account_id'] if src_t=='Stanica' else r['to_account_id']
 if trim(station)!=trim(scope['station']):raise AuditError('Station number-series owner differs from source physical event')
 docid=proposed_id(src_sha,scope)
 header={'AmbDokID':docid,'Vrsta':'REVERS','BrojDokumenta':scope['label'],
         'Datum':scope['day'],'BrojOwnerTip':'Stanica','BrojOwnerID':scope['station'],
         'Napomena':'Migriran istorijski Revers (izvorni broj zadrzan).','Stornirano':''}
 h={'target_header_draft':header,'migration_group_token':r['provisional_group_token'],
    'legacy_scope':scope,'source_row_ordinals':r['source_row_ordinals'],
    'source_legacy_doc_type':r['old_doc_type'],
    'creation_status':'DRAFT_NEW_ID_NOT_EXISTING_HISTORICAL_ID',
    'ownership_proven':False,'import_approved':False}
 link={'physical_case_id':r['physical_case_id'],'source_row_ordinals':r['source_row_ordinals'],
       'draft_DokumentTIP':'AmbalazaDokument','draft_DokumentID':docid,
       'draft_VrstaKretanja':r['proposed_kind'],
       'old_label':scope['label'],'old_document_type':r['old_doc_type'],
       'old_document_identity':'BUSINESS_LABEL_NO_EXISTING_AMBDOK_ID',
       'from_account_type':r['from_account_type'],'from_account_id':r['from_account_id'],
       'to_account_type':r['to_account_type'],'to_account_id':r['to_account_id'],
       'quantity':r['quantity'],'tip_ambalaze':r['type'],
       'title_or_ownership_verified':False,'import_approved':False}
 return h,link


def build(stage,expanded,seven,dryrun,storno_archive,identity_dir,intent_dir):
 # Reconstruct from pinned source and EVERY upstream private proof. Do not
 # trust a free-floating summary or allow edited private rows in this chain.
 old_summary,idents=bridge.build(stage,expanded,seven,dryrun,storno_archive)
 reported_id=json.loads((identity_dir/'agrix_amb10_identity_bridge_summary.json').read_text(encoding='utf-8'))
 id_digest=bridge.assert_records(identity_dir/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl',idents)
 if id_digest!=reported_id.get('identity_bridge_evidence_sha256'):
  raise AuditError('MIG-06.0 private identity checksum mismatch')
 if any(reported_id.get(k)!=v for k,v in old_summary.items()):
  raise AuditError('MIG-06.0 public identity summary changed')
 business_summary,events=intent.build(stage,expanded,seven,dryrun,storno_archive,identity_dir)
 reported=json.loads((intent_dir/'agrix_business_intent_summary.json').read_text(encoding='utf-8'))
 digest=bridge.assert_records(intent_dir/intent.PRIVATE,events)
 if reported.get('business_intent_evidence_sha256')!=digest:
  raise AuditError('MIG-06.1 private business evidence checksum mismatch')
 if any(reported.get(k)!=v for k,v in business_summary.items()):
  raise AuditError('MIG-06.1 public business summary changed')
 if len(idents)!=len(events) or len(events)!=4675:
  raise AuditError('Active physical events unexpectedly changed')
 headers=[];links=[];unknown=[]
 seen_ids=set();seen_scope=set();seen_group=set()
 for ident,ev in zip(idents,events):
  if ident['physical_case_id']!=ev['physical_case_id'] or ident['source_row_ordinals']!=ev['source_row_ordinals'] or ident['identity_status']!=ev['identity_status']:
   raise AuditError('MIG-06.0 and MIG-06.1 per-event identity drift')
  if ident['identity_status']==bridge.CATEGORIES[2]:
   rich=dict(ev,legacy_scope_station_day_label=ident.get('legacy_scope_station_day_label'))
   if ident['provisional_group_token']!=ev['provisional_group_token']:
    raise AuditError('Revers group token changed')
   h,l=make_proposal(rich,reported['source_sha256'])
   scope=h['legacy_scope'];scope_key=(f.key(scope['station']),scope['day'],f.key(scope['label']))
   if (scope_key in seen_scope or h['target_header_draft']['AmbDokID'] in seen_ids or
       ident['provisional_group_token'] in seen_group):
    raise AuditError('Duplicate/colliding Revers scope, planned ID or group token')
   seen_scope.add(scope_key);seen_ids.add(h['target_header_draft']['AmbDokID']);seen_group.add(ident['provisional_group_token'])
   headers.append(h);links.append(l)
  elif ident['identity_status']==bridge.CATEGORIES[3]:
   if ev['proposed_kind'] is not None or ident.get('verified_existing_DokumentID'):
    raise AuditError('Unknown document was silently assigned business meaning')
   unknown.append({'physical_case_id':ev['physical_case_id'],
                   'source_row_ordinals':ev['source_row_ordinals'],
                   'reason':'SOURCE_DOCUMENT_ID_AND_MOVEMENT_KIND_UNPROVEN',
                   'legacy_DokumentTIP':ev['old_doc_type'],
                   'legacy_DokumentID':ev['old_doc_label'],
                   'import_approved':False})
  else:
   if ident['identity_status'] not in bridge.CATEGORIES[:2]:raise AuditError('Unknown bridge class')
   if ident.get('verified_existing_DokumentID') is None:raise AuditError('Existing parent ID disappeared')
 if (len(headers),len(links),len(unknown))!=(48,48,6):
  raise AuditError('Expected precisely 48 historical Revers and six unknown-document holds')
 if any(h['import_approved'] or l['import_approved'] for h,l in zip(headers,links)):
  raise AuditError('Migration draft should never approve imports')
 public={'format':VERSION,'status':'DRAFT_NEW_AMBALAZA_DOCUMENT_IDENTITIES_NOT_IMPORT',
  'source_sha256':business_summary['source_sha256'],
  'schema_sha256':business_summary['schema_sha256'],
  'business_intent_evidence_sha256':digest,
  'identity_bridge_evidence_sha256':id_digest,
  'active_physical_events_unchanged':4675,
  'events_linked_to_existing_goods_parent':4621,
  'draft_new_historical_revers_headers':len(headers),
  'draft_new_revers_event_links':len(links),
  'draft_id_collisions':0,
  'new_document_identity_candidates_instead_of_historical_pk':len(headers),
  'unknown_document_holds_untouched':len(unknown),
  'events_with_legacy_kind_proposals_unchanged':4669,
  'negative_source_account_types_accepted':35,
  'unexplained_source_balance_residuals':0,
  'historical_cancelled_originals_archived':335,
  'q00700_held':True,
  'legal_title_verified':False,
  'draft_header_rows_import_approved':0,
  'canonical_amb10_ledger_entries_generated':0,
  'target_workbook_modified':False,
  'notes':[
   'New ADK- IDs are deterministic MIGRATION DRAFT identifiers, not IDs that existed in the old workbook.',
   'AmbDokument header plan uses Vrsta REVERS, BrojOwnerTip Stanica, and the original text business label and date.',
   'No implicit Revers for six unknown documents; source PK and business kind remain unproven.',
   'No physical or ownership changes, no deficit cover, no artificial opening, no canonical StornoOd.',
   'All customer identifiers and candidate headers remain exclusively in PRIVATE local JSONL.']}
 return public,headers,links,unknown


def report(s):
 fields=('active_physical_events_unchanged','events_linked_to_existing_goods_parent',
         'draft_new_historical_revers_headers','draft_new_revers_event_links','draft_id_collisions',
         'unknown_document_holds_untouched','events_with_legacy_kind_proposals_unchanged',
         'negative_source_account_types_accepted','unexplained_source_balance_residuals',
         'historical_cancelled_originals_archived','draft_header_rows_import_approved')
 return 'AGRIX MIG-06.2 — DETERMINISTIC HISTORICAL REVERS HEADER DRAFT / NO IMPORT\n'+''.join(
  k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in fields)


def cli(argv=None):
 p=argparse.ArgumentParser(description='Deterministic DRAFT new Revers document IDs; never updates source or target XLSM')
 for name in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','out'):
  p.add_argument('--'+name,required=True,type=Path)
 args=p.parse_args(argv)
 parents=[getattr(args,x.replace('-','_')).resolve() for x in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent')]
 out=args.out.resolve()
 if not all(x.is_dir() for x in parents) or out.exists() or any(out==x or out in x.parents or x in out.parents for x in parents):
  raise AuditError('Missing upstream directory, existing or overlapping output')
 summary,headers,links,unknown=build(*parents)
 out.mkdir(parents=True,exist_ok=False)
 try:
  files=[(PRIVATE_HEADERS,headers),(PRIVATE_LINKS,links),(PRIVATE_UNKNOWNS,unknown)]
  for name,records in files:
   dest=out/name
   with dest.open('w',encoding='utf-8',newline='\n') as fh:
    for item in records:fh.write(canon(item))
   os.chmod(dest,0o600)
  summary['private_draft_headers_sha256']=hashlib.sha256((out/PRIVATE_HEADERS).read_bytes()).hexdigest()
  summary['private_draft_links_sha256']=hashlib.sha256((out/PRIVATE_LINKS).read_bytes()).hexdigest()
  summary['private_unknown_holds_sha256']=hashlib.sha256((out/PRIVATE_UNKNOWNS).read_bytes()).hexdigest()
  (out/'agrix_revers_document_plan_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
  (out/'agrix_revers_document_plan_summary.txt').write_text(report(summary),encoding='utf-8')
 except BaseException:
  for p in out.iterdir():
   if p.is_file():p.unlink()
  out.rmdir();raise
 print(report(summary))
 print('Share summaries ONLY; deterministic headers and linked customer IDs remain PRIVATE locally. NO XLSM. NO IMPORT.')
 return 0


if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as err:
  print('MIG-06.2 STOP:',str(err),file=sys.stderr);raise SystemExit(2)
