#!/usr/bin/env python3
"""MIG-06.5: bounded, immutable *offline* target payload and optional XLSM schema inspection.

Source for one day of controlled Excel migration, not an import-authorized ledger.
No writes to any workbook. Exact upstream private byte integrity required. Policy holds
and legacy cancelled originals are NOT bypassed. Local customer data stays private.
"""
from __future__ import annotations
import argparse,hashlib,json,os,sys,zipfile
from collections import Counter,defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET
from legacy_audit import AuditError
import legacy_amb10_import_readiness_v2 as ready
import legacy_amb10_ledger_package_v2 as linked

VERSION='MIG-06.5'
PRIVATE_ROWS='AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl'
PRIVATE_HEADERS='AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl'
PRIVATE_EXCEPTIONS='AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl'
SUMMARY='agrix_target_adapter_summary.json'
TEXT='agrix_target_adapter_summary.txt'
REQUIRED_HEADERS={'AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Napomena','Stornirano'}

def canon(value):
 return json.dumps(value,ensure_ascii=False,sort_keys=True)+'\n'

def verify_readiness(inp):
 p=inp/ready.SUMMARY
 if not p.is_file():raise AuditError('MIG-06.4 summary missing')
 s=json.loads(p.read_text(encoding='utf-8'))
 expect={'format':'MIG-06.4','status':'SCOPED_IMPORT_DECISION_MANIFEST_ONLY_NO_IMPORT',
 'source_sha256':ready.SOURCE_SHA,'schema_sha256':ready.SCHEMA_SHA,
 'candidate_id_reservations_not_allocations':4669,'candidate_id_collisions':0,
 'planned_new_revers_headers':48,'candidate_rows_with_existing_goods_parent':4621,
 'candidate_rows_with_planned_revers_header':48,'unknown_document_held_events':6,
 'q00700_parent_storno_conflict_held':1,'source_cancelled_legs_archived_not_replayed':335,
 'negative_source_balances_allowed':35,'unknown_balance_residuals':0,
 'standalone_candidate_rows_without_issuance_or_new_header_gate':3485,
 'global_migration_gate_counts':6,'global_gates_satisfied':0,'row_approval_count':0,
 'import_approved':False,'target_xlsm_created_or_modified':False}
 for k,v in expect.items():
  if s.get(k)!=v:raise AuditError('Changed MIG-06.4 required invariant: '+k)
 if s.get('kind_counts')!=ready.KINDS or s.get('source_identity_counts')!=ready.GROUPS:
  raise AuditError('MIG-06.4 kind/category drift')
 if s.get('specific_local_gate_counts')!={'ISSUANCE_FOREIGN_LIABILITY_POLICY_PENDING':1154,
 'NEW_REVERS_DRAFT_REQUIRES_HEADER_ACCEPTANCE':48}:raise AuditError('MIG-06.4 per-row gate drift')
 if s.get('global_gate_codes')!=[x[0] for x in ready.GLOBAL]:
  raise AuditError('MIG-06.4 global gate drift')
 checks=((ready.PRIVATE_MAP,'candidate_crosswalk_sha256'),
         (ready.PRIVATE_UNKNOWN,'unknown_hold_sha256'),
         (ready.PRIVATE_GLOBAL,'global_gates_manifest_sha256'))
 rows={}
 for filename,k in checks:
  path=inp/filename
  if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest()!=s[k]:
   raise AuditError('MIG-06.4 checksum mismatch: '+filename)
  if filename==ready.PRIVATE_GLOBAL:
   data=json.loads(path.read_text(encoding='utf-8'))
   if len(data)!=6 or any(x.get('status')!='OPEN_NOT_APPROVED' for x in data):
    raise AuditError('MIG-06.4 global policy gates unexpectedly changed')
  else:
   data=ready.load_verified_lines(path,s[k])
  rows[filename]=data
 if len(rows[ready.PRIVATE_MAP])!=4669 or len(rows[ready.PRIVATE_UNKNOWN])!=6:
  raise AuditError('MIG-06.4 private partition changed')
 return s,rows[ready.PRIVATE_MAP],rows[ready.PRIVATE_UNKNOWN]

def inspect_template(path):
 """Read Excel OOXML table metadata WITHOUT opening/saving the workbook."""
 if not path.is_file() or path.suffix.lower() not in ('.xlsx','.xlsm'):
  raise AuditError('Provide a local blank TARGET .xlsx or .xlsm workbook')
 if not zipfile.is_zipfile(path):raise AuditError('Target is not Excel OOXML ZIP')
 names={}; macro=False
 with zipfile.ZipFile(path) as z:
  listing=set(z.namelist());macro='xl/vbaProject.bin' in listing
  for entry in listing:
   if not (entry.startswith('xl/tables/') and entry.endswith('.xml')):continue
   root=ET.fromstring(z.read(entry))
   if root.tag.rsplit('}',1)[-1]!='table':continue
   name=root.attrib.get('name') or root.attrib.get('displayName')
   if not name or name in names:raise AuditError('Target ListObject name missing or duplicate')
   columns=[]
   for elem in root.iter():
    if elem.tag.rsplit('}',1)[-1]=='tableColumn':
     columns.append(elem.attrib.get('name',''))
   if not columns or len(set(columns))!=len(columns):
    raise AuditError('Invalid duplicate or absent Excel table columns: '+name)
   names[name]=columns
 if path.suffix.lower()=='.xlsm' and not macro:
  raise AuditError('Target extension .xlsm but xl/vbaProject.bin is absent')
 if path.suffix.lower()=='.xlsx' and macro:
  raise AuditError('Target extension .xlsx but workbook contains VBA binary')
 require={'tblAmbalaza':set(linked.CANONICAL_COLUMNS),
          'tblAmbalazaDokument':REQUIRED_HEADERS}
 missing={table:sorted(set(cols)-set(names.get(table,[]))) for table,cols in require.items()}
 if any(missing.values()):raise AuditError('Missing target required table/columns: '+json.dumps(missing,ensure_ascii=False))
 if set(names['tblAmbalaza'])!=set(linked.CANONICAL_COLUMNS):
  raise AuditError('Target tblAmbalaza columns differ from pinned new canonical schema')
 return {'target_extension':path.suffix.lower(), 'target_macros_present':macro,
         'target_file_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
         'table_count':len(names),'tblAmbalaza_column_count':len(names['tblAmbalaza']),
         'tblAmbalazaDokument_column_count':len(names['tblAmbalazaDokument']),
         'required_target_tables_present':True,
         'required_ambalaza_columns_exact':True,
         'required_document_columns_present':True}

def materialize(linked_dir,readiness_dir,template=None):
 ls,headers,proposals,holds=ready.verify_linked(linked_dir)
 rs,mapping,unknown=verify_readiness(readiness_dir)
 if ls['private_target_draft_sha256']!=rs['parent_mig06_3_draft_sha256']:
  raise AuditError('Cross-stage MIG-06.3 and MIG-06.4 baseline mismatch')
 # Re-derive every MIG-06.4 mapping from original byte-verified proposals; never
 # trust a supplied crosswalk with altered account, document or kind fields.
 recovered,local,kinds,classes,roles=ready.classify(proposals,headers,holds)
 if [canon(r) for r in recovered]!=[canon(r) for r in mapping]:
  raise AuditError('MIG-06.4 crosswalk differs from recalculated source draft')
 if [canon(r) for r in holds]!=[canon(r) for r in unknown]:
  raise AuditError('MIG-06.4 unknown hold records changed')
 header_ids=set();out_heads=[]
 for h in headers:
  p=h['target_header_draft'].copy()
  if h.get('import_approved') is not False or h.get('ownership_proven') is not False:
   raise AuditError('Historical Revers draft approval changed')
  if not REQUIRED_HEADERS.issuperset(p.keys()):
   raise AuditError('Unexpected new Revers document header fields')
  if not all(p.get(k) for k in ('AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID')):
   raise AuditError('Incomplete Revers target header fields')
  if p.get('Stornirano') not in (None,'') or p['AmbDokID'] in header_ids:
   raise AuditError('Bad target Revers header storno/identity')
  header_ids.add(p['AmbDokID'])
  out_heads.append({'proposed_target_header':p,'status':'NOT_IMPORT_APPROVED',
                    'migration_provenance':h.get('migration_group_token')})
 result=[];seen=set();revers_refs=Counter()
 for p,m in zip(proposals,mapping):
  if (p['physical_case_id'],p['source_row_ordinals'])!=(m['physical_case_id'],m['source_row_ordinals']):
   raise AuditError('Candidate and reservation ordinal mismatch')
  candidate=p['proposed_target_columns'].copy()
  if candidate.get('AmbID') is not None or m['candidate_AmbID_NOT_ALLOCATED'] in seen:
   raise AuditError('Allocated/duplicate source candidate AmbID')
  seen.add(m['candidate_AmbID_NOT_ALLOCATED'])
  if (candidate['DokumentTIP'],candidate['DokumentID'],candidate['VrstaKretanja'])!=(
   m['document_type'],m['document_id'],m['movement_kind_proposal']):
   raise AuditError('Crosswalk document or kind mismatch')
  if (candidate['OdNalogTip'],candidate['OdNalogID'],candidate['NaNalogTip'],candidate['NaNalogID'])!=(
   m['from_account_type'],m['from_account_id'],m['to_account_type'],m['to_account_id']):
   raise AuditError('Crosswalk physical accounts mismatch')
  if candidate['DokumentTIP']=='AmbalazaDokument':
   if candidate['DokumentID'] not in header_ids:raise AuditError('Dangling planned Revers header')
   revers_refs[candidate['DokumentID']]+=1
  # This is *not* an import row: reserved ID stays outside target column data.
  result.append({'reserved_AmbID_NOT_ALLOCATED':m['candidate_AmbID_NOT_ALLOCATED'],
      'target_column_values_WITHOUT_AMBID':candidate,
      'source_row_ordinals':p['source_row_ordinals'],'physical_case_id':p['physical_case_id'],
      'local_gates':m['local_decision_gates'],'global_gates_required':True,
      'status':'NOT_IMPORT_APPROVED'})
 if len(result)!=4669 or len(seen)!=4669 or set(revers_refs)!=header_ids or set(revers_refs.values())!={1}:
  raise AuditError('Target adapter referential integrity failed')
 exception=[{'status':'UNKNOWN_DOCUMENT_HELD','evidence':x} for x in holds]
 exception.append({'status':'PARENT_STORNO_CONFLICT_HELD','source_row_ordinal':700})
 exception.append({'status':'HISTORICAL_STORNO_SOURCE_LEGS_ARCHIVED','archived_legs':335,
   'canonical_inverse_replay':'NOT_PROVEN'})
 schema=None if template is None else inspect_template(template)
 summary={'format':VERSION,'status':'TARGET_ADAPTER_PAYLOAD_NOT_APPROVED_NO_XLSM_WRITE',
  'source_sha256':ready.SOURCE_SHA,'schema_sha256':ready.SCHEMA_SHA,
  'mig06_3_private_draft_sha256':ls['private_target_draft_sha256'],
  'mig06_4_private_crosswalk_sha256':rs['candidate_crosswalk_sha256'],
  'revers_header_payload_rows':len(out_heads),'ledger_payload_rows_not_importable':len(result),
  'parent_goods_document_ledger_rows':4621,'new_revers_document_ledger_rows':48,
  'unknown_document_holds':6,'q00700_held':True,
  'historical_cancelled_legs_archived_no_replay':335,
  'source_negative_balances_accepted_without_modification':35,
  'unexplained_balance_residuals':0,'global_release_gates_open':6,
  'local_issuance_title_gates_open':1154,'local_revers_header_gates_open':48,
  'kind_counts':dict(sorted(kinds.items())),
  'candidate_ambid_collisions':0,'referential_errors':0,
  'template_inspection':schema if schema else {'status':'NOT_PROVIDED'},
  'audit_fields_or_canonical_storno_populated':False,
  'canonical_import_approved':False,'xlsx_or_xlsm_modified':False,
  'note':'Adapter is an import shape contract and integrity-checked payload. Not safe to load into live workbook.'}
 return summary,out_heads,result,exception

def text_summary(s):
 fields=['revers_header_payload_rows','ledger_payload_rows_not_importable',
  'parent_goods_document_ledger_rows','new_revers_document_ledger_rows',
  'unknown_document_holds','q00700_held',
  'historical_cancelled_legs_archived_no_replay',
  'source_negative_balances_accepted_without_modification','unexplained_balance_residuals',
  'global_release_gates_open','local_issuance_title_gates_open','local_revers_header_gates_open',
  'referential_errors','template_inspection','canonical_import_approved']
 return 'AGRIX MIG-06.5 / TARGET ADAPTER PAYLOAD -- NOT EXCEL IMPORT\n'+''.join(
  k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in fields)

def cli(argv=None):
 ap=argparse.ArgumentParser(description='Build bounded local AMB-10 target payload; no Excel write')
 ap.add_argument('--linked',type=Path,required=True)
 ap.add_argument('--readiness',type=Path,required=True)
 ap.add_argument('--out',type=Path,required=True)
 ap.add_argument('--template',type=Path,help='Optional blank refactored target XLSM/XLSX; INSPECT ONLY, no modification')
 a=ap.parse_args(argv);l=a.linked.resolve();r=a.readiness.resolve();o=a.out.resolve()
 if not l.is_dir() or not r.is_dir() or o.exists():raise AuditError('Required staging missing or output already exists')
 for x in (l,r):
  if x==o or x in o.parents or o in x.parents:raise AuditError('Input/output overlap rejected')
 if a.template:
  t=a.template.resolve()
  if t==o or t in o.parents:raise AuditError('Template/output overlap rejected')
 summary,heads,rows,exception=materialize(l,r,a.template.resolve() if a.template else None)
 o.mkdir(parents=True,exist_ok=False)
 try:
  for filename,data,key in ((PRIVATE_HEADERS,heads,'private_header_payload_sha256'),
                (PRIVATE_ROWS,rows,'private_ledger_payload_sha256'),
                (PRIVATE_EXCEPTIONS,exception,'private_hold_manifest_sha256')):
   f=o/filename
   with f.open('w',encoding='utf-8',newline='\n') as stream:
    for row in data:stream.write(canon(row))
   os.chmod(f,0o600)
   summary[key]=hashlib.sha256(f.read_bytes()).hexdigest()
  (o/SUMMARY).write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  (o/TEXT).write_text(text_summary(summary),encoding='utf-8')
 except BaseException:
  for f in o.iterdir():
   if f.is_file():f.unlink()
  o.rmdir();raise
 print(text_summary(summary))
 print('Never upload PRIVATE JSONL. No Excel workbook was written or modified.')
 return 0

if __name__=='__main__':
 try:raise SystemExit(cli())
 except (AuditError,OSError,KeyError,ValueError,TypeError,json.JSONDecodeError,zipfile.BadZipFile,ET.ParseError) as ex:
  print('MIG-06.5 STOP:',str(ex),file=sys.stderr);raise SystemExit(2)
