#!/usr/bin/env python3
"""MIG-06.4: read-only import readiness manifest for pinned MIG-06.3 data.

Produces a deterministic candidate AmbID mapping and *scoped* decision queue,
without granting import approval, generating Excel workbooks, or representing
unverified title and negative legacy balances as canonical approved state.

Only processes byte-checked MIG-06.3 PRIVATE JSONLs, never outputs customer data
in the shareable summary. One Q-00700 and 335 cancelled historical original
legs remain outside the migration candidate population.
"""
from __future__ import annotations
import argparse
from collections import Counter,defaultdict
from decimal import Decimal,InvalidOperation
import hashlib,json,os,sys,uuid
from pathlib import Path
from legacy_audit import AuditError
import legacy_amb10_ledger_package_v2 as linked

VERSION='MIG-06.4'
SOURCE_SHA='a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'
SCHEMA_SHA='b9916598d1f8d29b8e66b3fd0805d36f4c2e695df48fe6ef32155a7b6a765d73'
NAMESPACE=uuid.UUID('fe0e7562-6069-5b9f-af43-971809a4f192')
PRIVATE_MAP='AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl'
PRIVATE_UNKNOWN='AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl'
PRIVATE_GLOBAL='AMB10_GLOBAL_MIGRATION_GATES_PRIVATE.json'
SUMMARY='agrix_amb10_import_readiness_summary.json'
TEXT='agrix_amb10_import_readiness_summary.txt'
KINDS={ 'AMBALAZA_UZ_ROBU':3349,'IZDATA_PRAZNA':1154,'POVRAT_PRAZNE':161,'PRENOS_INTERNO':5 }
GROUPS={
 'EXISTING_GOODS_PARENT_AND_KIND_PROPOSED':3485,
 'EXISTING_OTKUP_PARENT_KIND_OWNERSHIP_HELD':1136,
 'SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED':48
}
GLOBAL=(
 ('EXPLICIT_MIGRATION_ROUTE_NOT_IMPLEMENTED', 'No approved import or XLSM write path currently exists'),
 ('LEGAL_TITLE_POLICY_NOT_EVIDENCED', 'Legacy direction and purpose do not prove ownership or foreign title'),
 ('NEGATIVE_SOURCE_BALANCE_COMPATIBILITY_NOT_PROVED','35 negative account/type ending balances are allowed by source policy; live AMB-INV-07 remains independent'),
 ('Q00700_DOCUMENT_STORNO_DISPOSITION_PENDING','One active AMB leg is attached to a storno Otpremnica; source parity depends on policy'),
 ('HISTORICAL_CANCELLED_ORIGINALS_ARCHIVE_ONLY','335 cancelled old source legs preserved outside canonical StornoOd timeline'),
 ('SIX_UNKNOWN_DOCUMENTS_NOT_CANONICAL','Six physical events have neither proven document identity nor business kind'),
)

def canon(v):
 return json.dumps(v,ensure_ascii=False,sort_keys=True)+'\n'

def load_verified_lines(path:Path,sha:str):
 if not path.is_file():raise AuditError('Missing private MIG-06.3 file: '+path.name)
 if hashlib.sha256(path.read_bytes()).hexdigest()!=sha:
  raise AuditError('MIG-06.3 private file byte checksum mismatch: '+path.name)
 rows=[]
 with path.open(encoding='utf-8',newline='') as f:
  for l in f:
   if not l.endswith('\n'):raise AuditError('Unterminated linked private JSONL')
   r=json.loads(l)
   if l!=canon(r):raise AuditError('Noncanonical upstream linked JSONL serialization')
   rows.append(r)
 return rows

def verify_linked(folder):
 sp=folder/'agrix_linked_target_draft_summary.json'
 if not sp.is_file():raise AuditError('Missing MIG-06.3 summary')
 s=json.loads(sp.read_text(encoding='utf-8'))
 if (s.get('format'),s.get('status'))!=('MIG-06.3','PRIVATE_LINKED_TARGET_SHAPE_PREVIEW_NO_IMPORT'):
  raise AuditError('Unrecognized MIG-06.3 baseline')
 for key,want in (('source_sha256',SOURCE_SHA),('schema_sha256',SCHEMA_SHA),
 ('total_physical_events',4675),('target_shaped_proposals_not_approved',4669),
 ('unknown_document_event_holds',6),('linked_new_revers_header_drafts',48),
 ('new_revers_rows_linked',48),('rows_with_existing_parent',4621),
 ('inv04_colliding_keys',0),('inv10_documents_multiple_party_pairs',0),
 ('proposal_rows_with_inv04_or_inv10_hold',0),('source_active_legs_covered',7479),
 ('historical_cancelled_legacy_rows_archived',335),
 ('historical_cancelled_legacy_rows_automatically_replayed',0),
 ('source_negative_account_types_allowed',35),('unexplained_account_balance_residuals',0),
 ('approved_canonical_amb_ids',0),('synthetic_opening_balances',0),
 ('foreign_liability_or_deficit_cover_generated',0),
 ('active_q00700_held',True),('import_approved',False),('target_xlsm_created_or_modified',False)):
  if s.get(key)!=want:raise AuditError('Changed MIG-06.3 baseline: '+key)
 if s.get('kind_counts')!=KINDS or s.get('identity_counts')!=GROUPS:
  raise AuditError('Kind/identity partition drift')
 for k,want in (('FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN',1154),
                ('NEW_REVERS_HEADER_NOT_APPROVED',48),
                ('LEGAL_TITLE_NOT_PROVEN',4669),
                ('Q00700_PENDING',4669),
                ('HISTORICAL_CANCELLATIONS_NOT_REPLAYED',4669),
                ('MIGRATION_ONLY_NOT_IMPORTABLE',4669)):
  if s.get('blocker_counts',{}).get(k)!=want:raise AuditError('Prior blocker baseline drift')
 h=load_verified_lines(folder/linked.PRIVATE_HEADERS,s['private_header_draft_sha256'])
 r=load_verified_lines(folder/linked.PRIVATE_ROWS,s['private_target_draft_sha256'])
 x=load_verified_lines(folder/linked.PRIVATE_HOLDS,s['private_unknown_document_holds_sha256'])
 if (len(h),len(r),len(x))!=(48,4669,6):raise AuditError('MIG-06.3 record counts drift')
 return s,h,r,x

def candidate_id(case,source_ord,source_sha=SOURCE_SHA):
 if not isinstance(case,str) or not case or not isinstance(source_ord,str) or not source_ord:
  raise AuditError('Candidate lacks provenance')
 return 'AMB-'+uuid.uuid5(NAMESPACE,canon({'source_sha256':source_sha,'case':case,
                    'source_row_ordinals':source_ord,'migration_version':VERSION})).hex.upper()

def classify(rows,headers,unknowns):
 header_ids=[]
 for h in headers:
  if h.get('import_approved') is not False or h.get('ownership_proven') is not False:
   raise AuditError('Revers header unexpectedly approved')
  hd=h['target_header_draft']
  if hd.get('Vrsta')!='REVERS' or hd.get('BrojOwnerTip')!='Stanica' or not all(
     hd.get(z) for z in ('AmbDokID','BrojDokumenta','Datum','BrojOwnerID')):
   raise AuditError('Invalid planned Revers header')
  header_ids.append(hd['AmbDokID'])
 if len(set(header_ids))!=48:raise AuditError('Revers header IDs duplicate')
 bydoc={k:0 for k in header_ids}
 mapping=[];seen_id=set();seen_case=set();seen_sources=set();class_count=Counter();kind_count=Counter();local=Counter();pair_roles=Counter()
 for r in rows:
  if r.get('import_approved') is not False or r.get('canonical_AmbID_allocated') is not False:
   raise AuditError('MIG-06.3 row unexpectedly marked approved')
  p=r['proposed_target_columns'];doc=(p['DokumentTIP'],p['DokumentID']);kind=p['VrstaKretanja']
  if set(p)!=set(linked.CANONICAL_COLUMNS) or p['AmbID'] is not None or p['StornoOd'] is not None or any(p[z] is not None for z in ('CreatedAt','CreatedBy','ModifiedAt','ModifiedBy')):
   raise AuditError('Canonical upstream record acquired import metadata unexpectedly')
  if kind not in KINDS:raise AuditError('Unknown proposed kind')
  try:qty=Decimal(str(p['Kolicina']))
  except (ValueError,InvalidOperation):raise AuditError('Invalid quantity')
  if not qty.is_finite() or qty<=0 or not p['Datum'] or not p['TipAmbalaze']:
   raise AuditError('Invalid date/quantity/type')
  case=r['physical_case_id'];ordtext=r['source_row_ordinals'];ordinals=[]
  for z in ordtext.split(';'):
   if not z.isdecimal():raise AuditError('Invalid physical source ordinal')
   ordinals.append(int(z))
  if len(ordinals) not in (1,2) or len(set(ordinals))!=len(ordinals) or any(z<1 or z>7815 or z==700 for z in ordinals):
   raise AuditError('Invalid source leg lineage')
  if case in seen_case or seen_sources.intersection(ordinals):raise AuditError('Duplicate case/source lineage')
  seen_case.add(case);seen_sources.update(ordinals)
  ci=candidate_id(case,ordtext)
  if ci in seen_id:raise AuditError('Candidate migration ID collision')
  seen_id.add(ci)
  c=r['identity_status'];class_count[c]+=1;kind_count[kind]+=1
  original=set(r.get('import_blockers',[]))
  expect={'MIGRATION_ONLY_NOT_IMPORTABLE','LEGAL_TITLE_NOT_PROVEN',
          'HISTORICAL_CANCELLATIONS_NOT_REPLAYED','Q00700_PENDING'}
  if kind=='IZDATA_PRAZNA':expect.add('FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN')
  if c=='SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED':expect.add('NEW_REVERS_HEADER_NOT_APPROVED')
  if original!=expect:raise AuditError('Upstream blocker drift for a physical event')
  factors=[]
  if kind=='IZDATA_PRAZNA':factors.append('ISSUANCE_FOREIGN_LIABILITY_POLICY_PENDING')
  if c=='SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED':
   if doc[0]!='AmbalazaDokument' or doc[1] not in bydoc:
    raise AuditError('Revers row does not reference planned new header')
   bydoc[doc[1]]+=1
   factors.append('NEW_REVERS_DRAFT_REQUIRES_HEADER_ACCEPTANCE')
  elif c in ('EXISTING_GOODS_PARENT_AND_KIND_PROPOSED','EXISTING_OTKUP_PARENT_KIND_OWNERSHIP_HELD'):
   if doc[0]=='AmbalazaDokument' or not all(doc):
    raise AuditError('Known goods event document is missing or revers-classified')
  else:raise AuditError('Unexpected identity class in linked known rows')
  for factor in factors:local[factor]+=1
  pair_roles[(p['OdNalogTip'],p['NaNalogTip'],kind)]+=1
  mapping.append({'physical_case_id':case,'source_row_ordinals':ordtext,
     'candidate_AmbID_NOT_ALLOCATED':ci,'document_type':doc[0],
     'document_id':doc[1], 'movement_kind_proposal':kind,
     'from_account_type':p['OdNalogTip'],'from_account_id':p['OdNalogID'],
     'to_account_type':p['NaNalogTip'],'to_account_id':p['NaNalogID'],
     'identity_class':c, 'local_decision_gates':factors,
     'global_migration_gates_apply':True,'import_approved':False})
 if set(bydoc.values())!={1}:raise AuditError('Not exactly one planned movement per new Revers header')
 if class_count!=GROUPS or kind_count!=KINDS or len(seen_sources)!=7473:
  raise AuditError('Candidate input population partition changed')
 hold_ords=set()
 if len({h['physical_case_id'] for h in unknowns})!=6:raise AuditError('Duplicate unknown document case')
 for h in unknowns:
  if h.get('import_approved') is not False or h.get('reason')!='NO_SOURCE_DOCUMENT_ID_AND_NO_PROVEN_KIND':
   raise AuditError('Unknown document was wrongly resolved')
  ids=[int(x) for x in h['source_row_ordinals'].split(';')]
  if len(ids)!=1 or ids[0] in seen_sources or ids[0] in hold_ords or ids[0]==700:
   raise AuditError('Unknown document source overlap')
  hold_ords.update(ids)
 if len(seen_sources|hold_ords)!=7479:raise AuditError('Known plus unknown source lineage changed')
 return mapping,local,kind_count,class_count,pair_roles

def build(folder):
 source,headers,rows,holds=verify_linked(folder)
 mapping,local,kinds,classes,roles=classify(rows,headers,holds)
 globals=[{'gate':k,'applies_to':'WHOLE_MIGRATION','reason':why,
           'status':'OPEN_NOT_APPROVED'} for k,why in GLOBAL]
 s={
 'format':VERSION,'status':'SCOPED_IMPORT_DECISION_MANIFEST_ONLY_NO_IMPORT',
 'source_sha256':SOURCE_SHA,'schema_sha256':SCHEMA_SHA,
 'parent_mig06_3_draft_sha256':source['private_target_draft_sha256'],
 'planned_new_revers_headers':len(headers),'candidate_id_reservations_not_allocations':len(mapping),
 'candidate_id_collisions':len(mapping)-len({x['candidate_AmbID_NOT_ALLOCATED'] for x in mapping}),
 'candidate_rows_with_existing_goods_parent':classes['EXISTING_GOODS_PARENT_AND_KIND_PROPOSED']+classes['EXISTING_OTKUP_PARENT_KIND_OWNERSHIP_HELD'],
 'candidate_rows_with_planned_revers_header':classes['SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED'],
 'unknown_document_held_events':len(holds),
 'q00700_parent_storno_conflict_held':1,
 'source_cancelled_legs_archived_not_replayed':335,
 'negative_source_balances_allowed':35,
 'unknown_balance_residuals':0,
 'kind_counts':dict(sorted(kinds.items())),
 'source_identity_counts':dict(sorted(classes.items())),
 'specific_local_gate_counts':dict(sorted(local.items())),
 'global_migration_gate_counts':len(globals),
 'global_gate_codes':[x['gate'] for x in globals],
 'standalone_candidate_rows_without_issuance_or_new_header_gate':sum(not m['local_decision_gates'] for m in mapping),
 'global_gates_satisfied':0,'row_approval_count':0,'import_approved':False,
 'target_xlsm_created_or_modified':False,
 'notes':[
  'This report distinguishes truly per-row gates from independent global migration release gates.',
  'Candidate AmbIDs are reserved in a PRIVATE crosswalk only; target rows still have no AmbID and are NEVER importable.',
  'No negative source balances were changed or synthetic opening/foreign liability created.',
  'A policy decision on source negative balances does NOT override normal AMB-INV-07 live writer rules.',
  '335 old cancelled originals remain an integrity-hashed archive; not retroactively fabricated StornoOd inverses.',
  'The six unknown documents and Q-00700 are held separately, without invented document identities.',
  'Zero local per-row gates means no identified local issue of these two types, NOT legal or global import approval.'
 ]}
 if s['candidate_id_collisions'] or len(mapping)!=4669 or s['standalone_candidate_rows_without_issuance_or_new_header_gate']<0:
  raise AuditError('Import readiness partition invalid')
 return s,mapping,holds,globals

def txt(s):
 keys=('candidate_id_reservations_not_allocations','candidate_id_collisions',
       'candidate_rows_with_existing_goods_parent','candidate_rows_with_planned_revers_header',
       'unknown_document_held_events','q00700_parent_storno_conflict_held',
       'source_cancelled_legs_archived_not_replayed','negative_source_balances_allowed',
       'kind_counts','specific_local_gate_counts',
       'standalone_candidate_rows_without_issuance_or_new_header_gate',
       'global_migration_gate_counts','global_gate_codes','import_approved')
 return 'AGRIX MIG-06.4 - SCOPED IMPORT READINESS / NO EXCEL WRITE\n'+''.join(
   k+': '+json.dumps(s[k],sort_keys=True,ensure_ascii=False)+'\n' for k in keys)

def cli(argv=None):
 p=argparse.ArgumentParser(description='Separate local and global AMB-10 release gates; NOT an import tool')
 p.add_argument('--linked',type=Path,required=True)
 p.add_argument('--out',type=Path,required=True)
 args=p.parse_args(argv);inp=args.linked.resolve();out=args.out.resolve()
 if not inp.is_dir() or out.exists() or out==inp or out in inp.parents or inp in out.parents:
  raise AuditError('Linked MIG-06.3 input missing, or output exists/overlaps input')
 s,mapping,holds,global_gates=build(inp)
 out.mkdir(parents=True,exist_ok=False)
 try:
  for name,records in ((PRIVATE_MAP,mapping),(PRIVATE_UNKNOWN,holds)):
   with (out/name).open('w',encoding='utf-8',newline='\n') as f:
    for row in records:f.write(canon(row))
   os.chmod(out/name,0o600)
  (out/PRIVATE_GLOBAL).write_text(json.dumps(global_gates,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  for name,key in ((PRIVATE_MAP,'candidate_crosswalk_sha256'),(PRIVATE_UNKNOWN,'unknown_hold_sha256'),(PRIVATE_GLOBAL,'global_gates_manifest_sha256')):
   s[key]=hashlib.sha256((out/name).read_bytes()).hexdigest()
  (out/SUMMARY).write_text(json.dumps(s,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  (out/TEXT).write_text(txt(s),encoding='utf-8')
 except BaseException:
  for item in out.iterdir():
   if item.is_file():item.unlink()
  out.rmdir();raise
 print(txt(s));print('Only share summaries. PRIVATE crosswalk includes customer IDs and is NOT an import table.')
 return 0

if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,KeyError,ValueError,TypeError,InvalidOperation,json.JSONDecodeError) as ex:
  print('MIG-06.4 STOP:',str(ex),file=sys.stderr);raise SystemExit(2)
