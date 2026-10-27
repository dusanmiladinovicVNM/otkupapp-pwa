#!/usr/bin/env python3
"""MIG-06.0 — resolve EXISTING document identity without inventing an AMB-10 document.

Strictly local, read-only reconciliation of 4,675 legacy physical events.
Corrects MIG-05.6's overbroad `CANONICAL_SOURCE_DOCUMENT_IDENTITY_UNRESOLVED`:
1,136 OM-Izlaz-Koop physical events are linked to real, ACTIVE OtkupID parents.
Ownership/purpose remains UNKNOWN, and 48 standalone historical Revers scopes
are *not* treated as historically existing opaque AmbDokIDs. No Excel writes.
"""
from __future__ import annotations
import argparse,hashlib,json,os,sys
from collections import Counter
from decimal import Decimal
from pathlib import Path
from legacy_audit import AuditError,is_active,trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_canonical_dryrun_v2 as dry
import legacy_amb10_chronology_preview_v2 as chronology
import legacy_amb10_storno_archive_v2 as archived

VERSION='MIG-06.0'
ARCHIVE_NAME='STORNO_ARCHIVE_ALL_CANCELLED_ORIGINALS_PRIVATE.jsonl'
DRY_NAME='AMB10_DRYRUN_ALL_CASES_PRIVATE.jsonl'
PARITY_NAME='AMB10_DRYRUN_BALANCE_PARITY_PRIVATE.jsonl'
CATEGORIES=(
 'EXISTING_GOODS_PARENT_AND_KIND_PROPOSED',
 'EXISTING_OTKUP_PARENT_KIND_OWNERSHIP_HELD',
 'SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED',
 'UNMODELLED_LEG_DOCUMENT_UNRESOLVED',
)

def canon(obj):
 return json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'

def assert_records(path,records):
 if not path.is_file():raise AuditError('Missing PRIVATE evidence: '+str(path))
 with path.open(encoding='utf-8',newline='') as inp:
  for line,expected in zip(inp,records):
   if line!=canon(expected):raise AuditError('Modified PRIVATE evidence: '+str(path))
  if next(inp,None) is not None:raise AuditError('Extra PRIVATE evidence: '+str(path))
 # zip() does not detect missing lines; compare size to expected serialized bytes.
 expected_digest=hashlib.sha256(''.join(canon(x) for x in records).encode('utf-8')).hexdigest()
 actual_digest=hashlib.sha256(path.read_bytes()).hexdigest()
 if expected_digest!=actual_digest:raise AuditError('Missing or modified PRIVATE evidence '+str(path))
 return actual_digest

def find_revers_station(rows):
 stations={trim(r['EntitetID']) for r in rows if f.key(r['EntitetTip'])=='stanica'}
 if len(stations)!=1 or not next(iter(stations)):
  raise AuditError('Legacy Revers has no unique station account')
 return next(iter(stations))

def prove_existing_otkup(ev,src,indices):
 if f.key(ev['legacy_DokumentTIP'])!='om-izlaz-koop':
  raise AuditError('Expected OM-Izlaz-Koop Otkup-linked event')
 if ev['event_provenance']!='MIG05_ORIGINAL_PHYSICAL' or ev['evidence_status']!='HOLD_OWNERSHIP':
  raise AuditError('Unproven legacy OM provenance passed as Otkup')
 ids=[int(i) for i in ev['source_row_ordinals'].split(';')]
 if len(ids)!=2 or ids[0]==ids[1]:raise AuditError('Otkup issued packaging must have two source legs')
 src_legs=[src[i-1] for i in ids]
 docid=trim(ev['legacy_DokumentID'])
 parent=indices['SOURCE__tblOtkup'].get(docid)
 if not parent or not is_active(parent[1]):raise AuditError('Otkup parent missing or storno')
 p=parent[1]
 if not docid or any(not is_active(z) or f.key(z['DokumentTip'])!='om-izlaz-koop' or
   trim(z['DokumentID'])!=docid or f.dec(z['Kolicina'])!=f.dec(ev['Kolicina']) or
   f.key(z['TipAmbalaze'])!=f.key(ev['TipAmbalaze']) for z in src_legs):
  raise AuditError('Otkup source legs cannot be tied to existing parent')
 if not (ev['OdNalogTip']=='Stanica' and trim(ev['OdNalogID'])==trim(p.get('StanicaID')) and
         ev['NaNalogTip']=='Kooperant' and trim(ev['NaNalogID'])==trim(p.get('KooperantID')) and
         f.dec(ev['Kolicina']) is not None and f.dec(ev['Kolicina'])>0 and
         f.dec(ev['Kolicina'])==f.dec(p.get('KolAmbIzdata')) and
         f.key(ev['TipAmbalaze'])==f.key(p.get('TipAmbalaze'))):
  raise AuditError('Otkup parent/leg identity or issued quantity mismatch')
 if {(z['EntitetTip'],f.key(z['Smer']),trim(z['EntitetID'])) for z in src_legs}!={
       ('Stanica','izlaz',trim(p['StanicaID'])),
       ('Kooperant','ulaz',trim(p['KooperantID']))}:
  raise AuditError('Otkup issued packaging mirror leg roles mismatch')
 return 'Otkup',docid

def scope_revers(ev,src,scopes):
 if ev['event_provenance']!='MIG05_2_RECOVERED_REVERS_PHYSICAL':
  raise AuditError('Expected independently proven Revers provenance')
 ids=[int(i) for i in ev['source_row_ordinals'].split(';')]
 if len(ids) not in (1,2) or len(set(ids))!=len(ids):raise AuditError('Bad Revers source count')
 legs=[src[i-1] for i in ids]
 typ=f.key(ev['legacy_DokumentTIP'])
 if typ not in ('om-izlaz-koop','om-ulaz-koop','om-izlaz-firma','om-ulaz-firma'):
  raise AuditError('Unknown old Revers type')
 if any(not is_active(z) or f.key(z['DokumentTip'])!=typ or
        trim(z['DokumentID'])!=trim(ev['legacy_DokumentID']) or
        f.dec(z['Kolicina'])!=f.dec(ev['Kolicina']) or
        f.key(z['TipAmbalaze'])!=f.key(ev['TipAmbalaze']) for z in legs):
  raise AuditError('Historical Revers source evidence mismatch')
 station=find_revers_station(legs)
 if len(legs)!= (1 if typ.endswith('firma') else 2):
  raise AuditError('Historical Revers event shape mismatch')
 raw_date=trim(legs[0]['Datum'])
 if any(f.key(z['Datum'])!=f.key(raw_date) for z in legs):
  raise AuditError('Historical Revers has inconsistent leg dates')
 parsed=chronology.date_key(raw_date)
 if parsed is None or not trim(ev['legacy_DokumentID']):
  raise AuditError('Historical Revers lacks scoped station/day/document label')
 scope=(station,parsed.isoformat(),trim(ev['legacy_DokumentID']))
 if scope in scopes:raise AuditError('Colliding legacy Revers business number within station/day')
 scopes.add(scope)
 token='MIG-GROUP-'+hashlib.sha256(canon(scope).encode('utf-8')).hexdigest()[:24]
 return scope,token

def identity_link(ev,old,source,indices,scopes):
 if old['physical_case_id']!=ev['physical_case_id'] or old['source_row_ordinals']!=ev['source_row_ordinals']:
  raise AuditError('Dry run and staging case order/identity diverged')
 if old['import_approved'] or ev.get('canonical_AMB10_import_authorized')!='NO':
  raise AuditError('Upstream evidence claims prohibited canonical authorization')
 rec={'physical_case_id':ev['physical_case_id'],
      'source_row_ordinals':ev['source_row_ordinals'],
      'legacy_doc_label':ev['legacy_DokumentID'],
      'legacy_doc_type':ev['legacy_DokumentTIP'],
      'legacy_event_provenance':ev['event_provenance'],
      'verified_existing_DokumentTIP':None,'verified_existing_DokumentID':None,
      'provisional_group_token':None,'provisional_document_kind':None,
      'identity_status':'','purpose_status':'UNPROVEN',
      'ownership_status':'UNPROVEN', 'canonical_new_doc_id':None,
      'storno_od':None,'import_approved':False}
 existing=old['movement_preview']
 if existing is not None:
  if ev['evidence_status']!='DIRECTION_KIND_PROPOSED_ONLY' or old['blocking_reasons']:
   raise AuditError('Proposed goods event unexpectedly not verified')
  if existing['DokumentID']!=trim(ev['legacy_DokumentID']) or existing['VrstaKretanja']!=ev['VrstaKretanja_PROPOSAL']:
   raise AuditError('Goods document identity or proposal mismatch')
  rec['identity_status']=CATEGORIES[0]
  rec['verified_existing_DokumentTIP']=existing['DokumentTIP']
  rec['verified_existing_DokumentID']=existing['DokumentID']
  rec['purpose_status']='KIND_PROPOSED_NOT_APPROVED'
 elif ev['event_provenance']=='MIG05_ORIGINAL_PHYSICAL':
  typ,docid=prove_existing_otkup(ev,source,indices)
  rec['identity_status']=CATEGORIES[1]
  rec['verified_existing_DokumentTIP']=typ
  rec['verified_existing_DokumentID']=docid
  rec['purpose_status']='ISSUED_OR_FOREIGN_RETURN_UNPROVEN'
 elif ev['event_provenance']=='MIG05_2_RECOVERED_REVERS_PHYSICAL':
  scope,token=scope_revers(ev,source,scopes)
  rec['identity_status']=CATEGORIES[2]
  rec['provisional_group_token']=token
  rec['provisional_document_kind']='REVERS_CANDIDATE_NOT_APPROVED'
  rec['legacy_scope_station_day_label']={'station':scope[0],'day':scope[1],'label':scope[2]}
  rec['purpose_status']='REVERS_KIND_AND_OWNERSHIP_UNPROVEN'
 elif ev['event_provenance']=='MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY':
  ids=[int(i) for i in ev['source_row_ordinals'].split(';')]
  if len(ids)!=1 or trim(source[ids[0]-1]['AmbID']) not in ev['source_amb_ids']:
   raise AuditError('Unknown document leg lineage mismatch')
  rec['identity_status']=CATEGORIES[3]
  rec['purpose_status']='UNKNOWN_DOCUMENT_AND_KIND'
 else:raise AuditError('Unexpected undocumented physical event provenance')
 return rec

def build(stage,expanded,seven,dryrun,storno_archive):
 report,old_cases,account_parity,history=dry.build(stage,expanded,seven)
 if report.get('format')!='MIG-05.6' or report.get('target_shaped_candidate_count')!=3485:
  raise AuditError('Pinned dry run evidence changed')
 stored=json.loads((dryrun/'agrix_amb10_dryrun_summary.json').read_text(encoding='utf-8'))
 if stored!=report:raise AuditError('Dry run public summary drift')
 assert_records(dryrun/DRY_NAME,old_cases)
 assert_records(dryrun/PARITY_NAME,account_parity)
 sr=json.loads((storno_archive/'agrix_storno_archive_summary.json').read_text(encoding='utf-8'))
 if sr.get('format')!='MIG-05.8' or sr.get('source_sha256')!=report['source_sha256'] or sr.get('schema_sha256')!=report['schema_sha256'] or sr.get('historical_storno_rows_archived')!=335 or sr.get('canonical_storno_od_links_generated')!=0 or sr.get('import_approved') is not False:
  raise AuditError('Missing or changed historical archive baseline')
 _,tables,indices=f.inputs(stage)
 all_raw=tables['UNCONVERTED__tblAmbalaza']
 original_archive=[dict({k:r[k] for k in f.FIELDS},source_row_ordinal=i,archive_class='CANCELLED_LEGACY_ORIGINAL_NOT_CANONICAL_EVENT') for i,r in enumerate(all_raw,1) if not is_active(r)]
 if len(original_archive)!=335:raise AuditError('Original cancelled rows drift')
 digest=assert_records(storno_archive/ARCHIVE_NAME,original_archive)
 if digest!=sr['legacy_archive_sha256']:
  raise AuditError('Historical archive checksum changed')
 if sr['historical_rows_still_physically_unresolved']!=8:
  raise AuditError('Archive historical unresolved population changed')
 # MIG-05.4 source evidence is independently and exactly re-derived by dry.build.
 _,events,_=chronology._verify_inputs(stage,expanded,seven)
 if len(events)!=len(old_cases) or len(events)!=4675:
  raise AuditError('Event count or dry-run lineage mismatch')
 scopes=set();rows=[]
 for ev,rec in zip(events,old_cases):
  rows.append(identity_link(ev,rec,all_raw,indices,scopes))
 counts=Counter(x['identity_status'] for x in rows)
 if (counts[CATEGORIES[0]],counts[CATEGORIES[1]],counts[CATEGORIES[2]],counts[CATEGORIES[3]])!=(3485,1136,48,6):
  raise AuditError('Document classification totals drift: '+str(counts))
 if len(scopes)!=48:raise AuditError('Historical Revers scope uniqueness drift')
 if report['source_negative_balance_accounts_preserved']!=35 or report['unexplained_account_balance_residuals']!=0:
  raise AuditError('Negative legacy balance preservation failed')
 if any(x['unexplained']!='0' for x in account_parity):
  raise AuditError('Legacy balance parity no longer exact including Q-00700')
 result={'format':VERSION,'status':'SOURCE_DOCUMENT_IDENTITY_BRIDGE_ONLY_NO_CANONICAL_IMPORT',
         'source_sha256':report['source_sha256'],'schema_sha256':report['schema_sha256'],
         'active_physical_events_reconciled':len(rows),'active_legacy_source_legs_covered':report['active_legacy_source_legs_covered'],
         'identity_classification_counts':dict(sorted(counts.items())),
         'previously_misclassified_otkup_linked_events':counts[CATEGORIES[1]],
         'events_linked_to_existing_source_documents':counts[CATEGORIES[0]]+counts[CATEGORIES[1]],
         'historical_revers_scoped_without_opaque_id':len(scopes),
         'unknown_document_identity_active_events':counts[CATEGORIES[3]],
         'remaining_events_without_existing_source_doc_pk':counts[CATEGORIES[2]]+counts[CATEGORIES[3]],
         'movement_kinds_still_unproven':counts[CATEGORIES[1]]+counts[CATEGORIES[2]]+counts[CATEGORIES[3]],
         'legacy_source_negative_accounts_accepted':report['source_negative_balance_accounts_preserved'],
         'legacy_unexplained_balance_residuals':report['unexplained_account_balance_residuals'],
         'historical_storno_rows_integrity_archived':len(original_archive),
         'historical_storno_rows_still_physically_unresolved':sr['historical_rows_still_physically_unresolved'],
         'q00700_state_conflict_still_on_hold':True,
         'no_ambalaza_dokument_allocated':True,'no_storno_od_allocated':True,
         'no_synthetic_openings':True,'import_approved':False,
         'notes':[
           'MIG-05.6 called 1136 true legacy OM-Izlaz-Koop Otkup-linked events missing source identity; this is corrected on existing OtkupID evidence.',
           '1136 Otkup-linked events still lack documented packaging ownership/purpose (IZDATA_PRAZNA vs VRACANJE_TUDJE_AMBALAZE).',
           '48 standalone historical Revers have a scoped business label, NOT an opaque historical AmbDokID; migration tokens only identify evidence groups.',
           '6 unmodelled documents and 1 Q-00700 remain held; 335 cancelled source rows stay private in the immutable archive.',
           'No valid canonical import table, no workbook modification, no foreign obligations, deficit cover or opening stock created.'
         ]}
 return result,rows

def render(s):
 keys=('active_physical_events_reconciled','identity_classification_counts',
       'previously_misclassified_otkup_linked_events',
       'events_linked_to_existing_source_documents','historical_revers_scoped_without_opaque_id',
       'unknown_document_identity_active_events','remaining_events_without_existing_source_doc_pk',
       'movement_kinds_still_unproven','legacy_source_negative_accounts_accepted',
       'legacy_unexplained_balance_residuals','historical_storno_rows_integrity_archived',
       'historical_storno_rows_still_physically_unresolved','import_approved')
 return 'AGRIX MIG-06.0 / IDENTITY BRIDGE ONLY — NO AMB10 IMPORT\n'+''.join(k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in keys)

def cli(argv=None):
 ap=argparse.ArgumentParser(description='Resolve legacy Otkup document identity; no canonical AMB-10 write')
 for name in ('stage','expanded','seven','dryrun','storno-archive','out'):
  ap.add_argument('--'+name,type=Path,required=True)
 ns=ap.parse_args(argv)
 paths=[getattr(ns,key.replace('-','_')).resolve() for key in ('stage','expanded','seven','dryrun','storno-archive')]
 out=ns.out.resolve()
 if out.exists() or any(out==p or out in p.parents or p in out.parents for p in paths):
  raise AuditError('Destination exists or overlaps source inputs')
 if not all(p.is_dir() for p in paths):raise AuditError('Missing required MIG stage directory')
 summary,records=build(*paths)
 out.mkdir(parents=True,exist_ok=False)
 try:
  private=out/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl'
  with private.open('w',encoding='utf-8',newline='\n') as fh:
   for r in records:fh.write(canon(r))
  os.chmod(private,0o600)
  summary['identity_bridge_evidence_sha256']=hashlib.sha256(private.read_bytes()).hexdigest()
  (out/'agrix_amb10_identity_bridge_summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  (out/'agrix_amb10_identity_bridge_summary.txt').write_text(render(summary),encoding='utf-8')
 except BaseException:
  for p in out.iterdir():
   if p.is_file():p.unlink()
  out.rmdir();raise
 print(render(summary));print('Share summaries ONLY; PRIVATE file stays local. NO XLSM.')
 return 0

if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
  print('MIG-06.0 STOP:',str(exc),file=sys.stderr);raise SystemExit(2)
