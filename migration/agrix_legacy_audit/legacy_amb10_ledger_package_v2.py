#!/usr/bin/env python3
"""MIG-06.3 — linked private target-shaped AMB-10 migration review package.

NO XLSM write, NO canonical import, NO fake initial stock or deficit cover. Writes
incomplete records on purpose (AmbID, audit fields, StornoOd unset); records
must never be pasted into tblAmbalaza. Negative source balances are accepted.
"""
from __future__ import annotations
import argparse
from collections import Counter,defaultdict
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import sys

from legacy_audit import AuditError,trim
import legacy_amb10_revers_document_plan_v2 as plan
import legacy_amb10_identity_bridge_v2 as bridge
import legacy_amb10_business_intent_v2 as intent
import legacy_amb10_chronology_preview_v2 as chrono
import legacy_amb10_forensics_v2 as f

VERSION='MIG-06.3'
PRIVATE_HEADERS='AMB10_LINKED_DRAFT_HEADERS_PRIVATE.jsonl'
PRIVATE_ROWS='AMB10_LINKED_DRAFT_LEDGER_PRIVATE.jsonl'
PRIVATE_HOLDS='AMB10_LINKED_DRAFT_HOLDS_PRIVATE.jsonl'
CANONICAL_COLUMNS=('AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP',
  'CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID',
  'NaNalogTip','NaNalogID','VrstaKretanja','StornoOd')
OWN={'Stanica','Vozac','Firma'}
PARTNER={'Kupac','Kooperant'}
REAL=OWN|PARTNER
ALLOWED_KINDS={'AMBALAZA_UZ_ROBU','IZDATA_PRAZNA','POVRAT_PRAZNE','PRENOS_INTERNO'}


def canon(obj):
 return json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'


def verify_draft(rev_dir:Path, summary, headers,links,unknowns):
 public=json.loads((rev_dir/'agrix_revers_document_plan_summary.json').read_text(encoding='utf-8'))
 if any(public.get(k)!=v for k,v in summary.items()):
  raise AuditError('MIG-06.2 public Revers plan changed')
 for name,rows,key in ((plan.PRIVATE_HEADERS,headers,'private_draft_headers_sha256'),
                      (plan.PRIVATE_LINKS,links,'private_draft_links_sha256'),
                      (plan.PRIVATE_UNKNOWNS,unknowns,'private_unknown_holds_sha256')):
  digest=bridge.assert_records(rev_dir/name,rows)
  if digest!=public.get(key):raise AuditError('MIG-06.2 private Revers draft SHA mismatch: '+name)
 return public


def target_row(ev, biz, ident, link):
 """Target-shaped *proposal*, deliberately incomplete and unimportable."""
 if (ev['physical_case_id']!=biz['physical_case_id'] or
     ev['physical_case_id']!=ident['physical_case_id'] or
     ev['source_row_ordinals']!=biz['source_row_ordinals'] or
     ev['source_row_ordinals']!=ident['source_row_ordinals']):
  raise AuditError('Per-event physical provenance, business kind or identity drift')
 if biz['import_approved'] is not False or ident['import_approved'] is not False:
  raise AuditError('Upstream source unexpectedly authorized an import')
 if ev['canonical_AMB10_import_authorized']!='NO':
  raise AuditError('Upstream active physical event approval changed')
 c=ident['identity_status'];kind=biz['proposed_kind']
 if c==bridge.CATEGORIES[3]:
  if link is not None or kind is not None or ident.get('verified_existing_DokumentID'):
   raise AuditError('Unmodelled event acquired an invented document or business kind')
  return None,{'physical_case_id':ev['physical_case_id'],
    'source_row_ordinals':ev['source_row_ordinals'],
    'reason':'NO_SOURCE_DOCUMENT_ID_AND_NO_PROVEN_KIND',
    'legacy_DokumentTIP':ev['legacy_DokumentTIP'],
    'legacy_DokumentID':ev['legacy_DokumentID'],
    'import_approved':False}
 if kind not in ALLOWED_KINDS:
  raise AuditError('Business intent is not an allowed type for linked review')
 if c==bridge.CATEGORIES[2]:
  if link is None or link['physical_case_id']!=ev['physical_case_id'] or link['source_row_ordinals']!=ev['source_row_ordinals']:
   raise AuditError('Revers missing an exact MIG-06.2 event-to-document link')
  doc_type=link['draft_DokumentTIP'];doc_id=link['draft_DokumentID']
  if link['draft_VrstaKretanja']!=kind or link['title_or_ownership_verified'] is not False or link['import_approved'] is not False:
   raise AuditError('Revers planned business purpose or approval changed')
 else:
  if link is not None or c not in bridge.CATEGORIES[:2]:raise AuditError('Unexpected Revers link')
  doc_type=ident['verified_existing_DokumentTIP'];doc_id=ident['verified_existing_DokumentID']
  if not trim(doc_type) or not trim(doc_id):raise AuditError('Existing document identity vanished')
 day=chrono.date_key(ev['legacy_Datum'])
 qty=f.dec(ev['Kolicina'])
 if day is None or qty is None or not qty.is_finite() or qty<=0:
  raise AuditError('Missing/invalid date or quantity in target-shaped event')
 fr=(ev['OdNalogTip'],ev['OdNalogID']);to=(ev['NaNalogTip'],ev['NaNalogID'])
 if not all(fr) or not all(to) or fr==to:raise AuditError('Missing/same canonical account')
 if kind=='IZDATA_PRAZNA' and not (fr[0] in OWN and to[0] in PARTNER):
  raise AuditError('IZDATA_PRAZNA roles violate AMB-10 matrix')
 if kind=='POVRAT_PRAZNE' and not (fr[0] in PARTNER and to[0] in OWN):
  raise AuditError('POVRAT_PRAZNE roles violate AMB-10 matrix')
 if kind=='PRENOS_INTERNO' and not (fr[0] in OWN and to[0] in OWN):
  raise AuditError('PRENOS_INTERNO roles violate AMB-10 matrix')
 if kind=='AMBALAZA_UZ_ROBU' and not (fr[0] in REAL and to[0] in REAL):
  raise AuditError('AMBALAZA_UZ_ROBU roles violate AMB-10 matrix')
 # Prevent use as real import table: critical required fields intentionally unset.
 proposal={'AmbID':None,'Datum':day.isoformat(),'TipAmbalaze':trim(ev['TipAmbalaze']),
   'Kolicina':str(qty),'DokumentID':doc_id,'DokumentTIP':doc_type,
   'CreatedAt':None,'CreatedBy':None,'ModifiedAt':None,'ModifiedBy':None,
   'OdNalogTip':fr[0],'OdNalogID':fr[1],
   'NaNalogTip':to[0],'NaNalogID':to[1],
   'VrstaKretanja':kind,'StornoOd':None}
 if set(proposal)!=set(CANONICAL_COLUMNS) or not proposal['TipAmbalaze']:
  raise AuditError('Target schema candidate incomplete')
 blockers=['MIGRATION_ONLY_NOT_IMPORTABLE','LEGAL_TITLE_NOT_PROVEN',
           'HISTORICAL_CANCELLATIONS_NOT_REPLAYED','Q00700_PENDING']
 if kind=='IZDATA_PRAZNA':
  blockers.append('FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN')
 if c==bridge.CATEGORIES[2]:
  blockers.append('NEW_REVERS_HEADER_NOT_APPROVED')
 return {'physical_case_id':ev['physical_case_id'],
  'source_row_ordinals':ev['source_row_ordinals'],
  'identity_status':c,'business_intent_evidence':biz['business_intent_evidence'],
  'proposed_target_columns':proposal,'import_blockers':blockers,
  'canonical_AmbID_allocated':False,'import_approved':False},None


def collisions(rows):
 """Whole-document AMB-INV-04, AMB-INV-10 preflight for all known proposals."""
 keys=defaultdict(list);parties=defaultdict(set)
 for entry in rows:
  x=entry['proposed_target_columns']; doc=(x['DokumentTIP'],x['DokumentID'])
  identity=doc+(x['VrstaKretanja'],x['TipAmbalaze'])
  keys[identity].append(entry)
  parties[doc].add(frozenset(((x['OdNalogTip'],x['OdNalogID']),
                            (x['NaNalogTip'],x['NaNalogID']))))
 duplicate_keys={k for k,v in keys.items() if len(v)>1}
 multi_pairs={d for d,v in parties.items() if len(v)>1}
 for entry in rows:
  x=entry['proposed_target_columns'];doc=(x['DokumentTIP'],x['DokumentID'])
  identity=doc+(x['VrstaKretanja'],x['TipAmbalaze'])
  if identity in duplicate_keys:entry['import_blockers'].append('AMB_INV_04_COLLISION')
  if doc in multi_pairs:entry['import_blockers'].append('AMB_INV_10_PARTY_PAIR_COLLISION')
 return len(duplicate_keys),len(multi_pairs),sum('AMB_INV_04_COLLISION' in r['import_blockers'] or
   'AMB_INV_10_PARTY_PAIR_COLLISION' in r['import_blockers'] for r in rows)


def build(stage,expanded,seven,dryrun,storno_archive,identity_dir,intent_dir,revers_dir):
 old,headers,links,unknown=plan.build(stage,expanded,seven,dryrun,storno_archive,identity_dir,intent_dir)
 existing=verify_draft(revers_dir,old,headers,links,unknown)
 _,events,_=chrono._verify_inputs(stage,expanded,seven)
 _,identities=bridge.build(stage,expanded,seven,dryrun,storno_archive)
 _,intents=intent.build(stage,expanded,seven,dryrun,storno_archive,identity_dir)
 if not (len(events)==len(identities)==len(intents)==4675):raise AuditError('Input lineage incomplete')
 bycase={link['physical_case_id']:link for link in links}
 if len(bycase)!=48:raise AuditError('Duplicate Revers link physical event')
 known=set();rows=[];holds=[];used=set()
 for ev,ident,biz in zip(events,identities,intents):
  cid=ev['physical_case_id']
  if cid in used:raise AuditError('Duplicate physical case in ledger draft')
  used.add(cid)
  row,hold=target_row(ev,biz,ident,bycase.get(cid))
  if row:rows.append(row)
  else:holds.append(hold)
  source_ord=[int(z) for z in ev['source_row_ordinals'].split(';')]
  if any(i in known for i in source_ord):raise AuditError('Duplicate source leg in linked ledger')
  known.update(source_ord)
 if len(known)!=7479 or 700 in known or len(rows)!=4669 or len(holds)!=6:
  raise AuditError('Incorrect target-draft partition')
 # Check exact planned headers map and new Revers references (no ghosts).
 hdrids={h['target_header_draft']['AmbDokID'] for h in headers}
 reversrows=[r for r in rows if r['identity_status']==bridge.CATEGORIES[2]]
 if len(reversrows)!=48 or {r['proposed_target_columns']['DokumentID'] for r in reversrows}!=hdrids:
  raise AuditError('Dangling Revers draft document reference')
 if {x['physical_case_id'] for x in holds}!={x['physical_case_id'] for x in unknown}:
  raise AuditError('Unmodelled source documents not fully held')
 duplicate,party,collision_rows=collisions(rows)
 status_counts=Counter(x['identity_status'] for x in rows)
 kind_counts=Counter(x['proposed_target_columns']['VrstaKretanja'] for x in rows)
 blockers=Counter(b for x in rows for b in x['import_blockers'])
 # MIG-05.6 parity is still exact against 4,675 transfers + Q-00700 hold.
 if old['negative_source_account_types_accepted']!=35 or old['unexplained_source_balance_residuals']!=0:
  raise AuditError('Negative source balances or ending parity changed')
 if old['historical_cancelled_originals_archived']!=335 or not old['q00700_held']:
  raise AuditError('Historical exceptions disappeared')
 public={'format':VERSION,'status':'PRIVATE_LINKED_TARGET_SHAPE_PREVIEW_NO_IMPORT',
  'source_sha256':old['source_sha256'],'schema_sha256':old['schema_sha256'],
  'mig06_2_headers_sha256':existing['private_draft_headers_sha256'],
  'mig06_2_links_sha256':existing['private_draft_links_sha256'],
  'total_physical_events':len(events),'target_shaped_proposals_not_approved':len(rows),
  'unknown_document_event_holds':len(holds),
  'linked_new_revers_header_drafts':len(headers),
  'new_revers_rows_linked':len(reversrows),
  'rows_with_existing_parent':len(rows)-len(reversrows),
  'kind_counts':dict(sorted(kind_counts.items())),
  'identity_counts':dict(sorted(status_counts.items())),
  'inv04_colliding_keys':duplicate,'inv10_documents_multiple_party_pairs':party,
  'proposal_rows_with_inv04_or_inv10_hold':collision_rows,
  'blocker_counts':dict(sorted(blockers.items())),
  'source_active_legs_covered':len(known),'active_q00700_held':True,
  'historical_cancelled_legacy_rows_archived':335,
  'historical_cancelled_legacy_rows_automatically_replayed':0,
  'source_negative_account_types_allowed':35,
  'unexplained_account_balance_residuals':0,
  'approved_canonical_amb_ids':0,'synthetic_opening_balances':0,
  'foreign_liability_or_deficit_cover_generated':0,
  'target_xlsm_created_or_modified':False,'import_approved':False,
  'notes':[
   'Every proposed ledger row has AmbID/created audit columns unset: NEVER paste into Excel.',
   'Movement kinds are legacy writer proposals; no ownership or third-party liability is established.',
   'Six unknown-document physical cases, Q-00700 and 335 archived storno legs stay separate.',
   'Source negative balances are valid; no correction or artificial opening stock.',
   'INV-04/INV-10 preflight is advisory: any collision is an import blocker, not an auto-merge.',
   'New ADK- identifiers belong to migration header drafts, never masquerade as source IDs.'
  ]}
 return public,headers,rows,holds


def report(s):
 keys=('total_physical_events','target_shaped_proposals_not_approved','unknown_document_event_holds',
  'linked_new_revers_header_drafts','new_revers_rows_linked','rows_with_existing_parent',
  'kind_counts','inv04_colliding_keys','inv10_documents_multiple_party_pairs',
  'proposal_rows_with_inv04_or_inv10_hold','source_negative_account_types_allowed',
  'unexplained_account_balance_residuals','import_approved')
 return 'AGRIX MIG-06.3 — LINKED TARGET SHAPE DRAFT / NO IMPORT\n'+''.join(
   k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in keys)


def cli(argv=None):
 ap=argparse.ArgumentParser(description='Read-only linked AMB-10 target-shape preview; cannot import XLSM')
 for name in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','revers-plan','out'):
  ap.add_argument('--'+name,type=Path,required=True)
 a=ap.parse_args(argv)
 parents=[getattr(a,x.replace('-','_')).resolve() for x in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','revers-plan')]
 out=a.out.resolve()
 if not all(p.is_dir() for p in parents) or out.exists() or any(out==p or out in p.parents or p in out.parents for p in parents):
  raise AuditError('Inputs missing, output already exists or overlapping directory')
 s,headers,rows,holds=build(*parents)
 out.mkdir(parents=True,exist_ok=False)
 try:
  for name,recs in ((PRIVATE_HEADERS,headers),(PRIVATE_ROWS,rows),(PRIVATE_HOLDS,holds)):
   dest=out/name
   with dest.open('w',encoding='utf-8',newline='\n') as fh:
    for row in recs:fh.write(canon(row))
   os.chmod(dest,0o600)
  s['private_header_draft_sha256']=hashlib.sha256((out/PRIVATE_HEADERS).read_bytes()).hexdigest()
  s['private_target_draft_sha256']=hashlib.sha256((out/PRIVATE_ROWS).read_bytes()).hexdigest()
  s['private_unknown_document_holds_sha256']=hashlib.sha256((out/PRIVATE_HOLDS).read_bytes()).hexdigest()
  (out/'agrix_linked_target_draft_summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
  (out/'agrix_linked_target_draft_summary.txt').write_text(report(s),encoding='utf-8')
 except BaseException:
  for p in out.iterdir():
   if p.is_file():p.unlink()
  out.rmdir();raise
 print(report(s))
 print('Share summaries ONLY. Private draft JSONLs contain customer data and are NOT Excel import sheets.')
 return 0

if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
  print('MIG-06.3 STOP:',str(exc),file=sys.stderr);raise SystemExit(2)
