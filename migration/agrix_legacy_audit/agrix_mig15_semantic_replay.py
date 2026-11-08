#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-15: signed-account historical packaging replay (READ ONLY).

Uses local private MIG-06 artifacts; never prints individual client data,
never writes to an xlsm, source/stage, or repo. Creates a PRIVATE per-case
crossing register and an aggregate summary. This is NOT import authorization.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import sys

EXPECTED = {
 'MigrationAmb10BusinessIntent_v2/AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl': ('9eaedfc62a880e27ff985bf0741715c87370d6633a470b33efa1ae4d6a2c4982',4675),
 'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl': ('e2f354b3af25d0c0098be2ff838e7f9b528f89999f38c1877ad8f956bb8917aa',4669),
 'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl': ('0269ead5b5c9e0d6bad62b0bfa65c2e929d47fbf7ad0f86130c2411cb460dd82',48),
 'MigrationAmb10ImportReadiness_v2/AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl': ('2155b5ad6c192a62f7b7f0b57d9f5a7b0d8ffcaca3ebba7273aad678f89d9215',4669),
 'MigrationAmb10ImportReadiness_v2/AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl': ('c3f9707efdca5022eec5d18115a4c3ed52e4f73ddc1d60ea8e4afef481632562',6),
 'MigrationStage_v2/UNCONVERTED__tblAmbalaza.jsonl': ('8a3d1c370b95f08e602216d124af12d67febb7ae3260360789bad7abca138bb5',7815),
}
SCHEMA_FIELDS={'Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd'}
KINDS={'AMBALAZA_UZ_ROBU','IZDATA_PRAZNA','POVRAT_PRAZNE','PRENOS_INTERNO','ULAZ_TUDJE_AMBALAZE','VRACANJE_TUDJE_AMBALAZE','NABAVKA','OTPIS'}


def digest(path):
 h=hashlib.sha256()
 with path.open('rb') as f:
  for part in iter(lambda:f.read(1<<20),b''):h.update(part)
 return h.hexdigest()


def read_jsonl(path):
 out=[]
 with path.open('r',encoding='utf-8-sig') as f:
  for i,line in enumerate(f,1):
   if not line.strip():continue
   obj=json.loads(line)
   if not isinstance(obj,dict):raise ValueError('Non-object JSON line')
   out.append(obj)
 return out


def txt(x):return '' if x is None else str(x).strip()

def amount(x):
 if isinstance(x,bool) or x is None or not txt(x):raise ValueError('Missing numeric quantity')
 try:d=Decimal(str(x))
 except InvalidOperation as e:raise ValueError('Non-numeric quantity') from e
 if not d.is_finite() or d<=0:raise ValueError('Nonpositive/nonfinite quantity')
 return d


def ordinal(obj):
 arr=obj.get('source_row_ordinals',[])
 if not isinstance(arr,list) or not arr:return 10**15
 try:return min(int(x) for x in arr)
 except (TypeError,ValueError):return 10**15


def date_sort_key(val):
 if isinstance(val,(int,float)) and not isinstance(val,bool):
  return datetime(1899,12,30)+timedelta(days=float(val))
 v=txt(val)
 if not v:raise ValueError('Empty date')
 if len(v)<=18 and v.replace('.','',1).isdigit() and v[:1].isdigit() and '.' not in v[4:]:
  # Numeric Excel date stored as a string.
  d=Decimal(v)
  if 20000<d<80000:return datetime(1899,12,30)+timedelta(days=float(d))
 for fmt in ('%Y-%m-%d','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S','%d.%m.%Y','%d.%m.%Y.','%d.%m.%Y %H:%M:%S','%Y-%m-%dT%H:%M:%S.%f','%Y-%m-%d %H:%M:%S.%f','%Y-%m-%d %H:%M','%Y-%m-%dT%H:%M','%d.%m.%Y %H:%M'):
  try:return datetime.strptime(v,fmt)
  except ValueError:continue
 raise ValueError('Unknown historical date format')


def getrow(obj):
 d=obj.get('target_column_values_WITHOUT_AMBID')
 if not isinstance(d,dict):raise ValueError('Target row mapping missing/object expected')
 return d


def account_pair(row, prefix):
 return (txt(row.get(prefix+'NalogTip')),txt(row.get(prefix+'NalogID')))


def acc_key(t,i,typ):return (t.casefold(),i.casefold(),typ.casefold())


def balancedelta(delta,a,b,typ,q):
 delta[acc_key(a[0],a[1],typ)]-=q
 delta[acc_key(b[0],b[1],typ)]+=q


def run(root,out, inventory_path=None):
 root=root.resolve();out=out.resolve()
 if out==root or root in out.parents:
  # We allow an output in the parent project, never inside the scanned raw input tree.
  if any((out==root/rel or root/rel in out.parents) for rel in EXPECTED):
   raise ValueError('Output cannot replace input')
 if inventory_path:
  inventory=json.loads(inventory_path.read_text(encoding='utf-8'))
  if inventory.get('issues'):raise ValueError('The previous inventory reported issues')
  if inventory.get('matched_payload_count',0)<8:raise ValueError('Unexpected payload inventory')
 data={}
 for rel,(expected_hash,expected_n) in EXPECTED.items():
  path=root/rel
  if not path.is_file():raise ValueError('Required local payload missing: '+rel)
  if digest(path)!=expected_hash:raise ValueError('Changed payload SHA256: '+rel)
  rows=read_jsonl(path)
  if len(rows)!=expected_n:raise ValueError('Changed payload record count: '+rel)
  data[rel]=rows
 def obj(rel):return data[next(k for k in data if k.endswith('/'+rel))]
 base=obj('AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl')
 led=obj('AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl')
 headers=obj('AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl')
 reserves=obj('AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl')
 holds=obj('AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl')
 bybase={};byled={};byres={};byhold={}
 errors=Counter();move_counts=Counter();approval_status=Counter();local_gates=Counter()
 for row in base:
  case=txt(row.get('physical_case_id'))
  if not case or case in bybase:errors['DUPLICATE_OR_EMPTY_BASE_CASE']+=1
  bybase[case]=row
 for row in led:
  case=txt(row.get('physical_case_id'))
  if not case or case in byled:errors['DUPLICATE_OR_EMPTY_LEDGER_CASE']+=1
  byled[case]=row
 for row in reserves:
  case=txt(row.get('physical_case_id'))
  if not case or case in byres:errors['DUPLICATE_OR_EMPTY_RESERVATION_CASE']+=1
  byres[case]=row
 for row in holds:
  case=txt(row.get('physical_case_id'))
  if not case or case in byhold:errors['DUPLICATE_OR_EMPTY_HELD_CASE']+=1
  byhold[case]=row
 if (set(byled)&set(byhold)):errors['LEDGER_AND_HELD_OVERLAP']+=len(set(byled)&set(byhold))
 if (set(bybase)-set(byled)-set(byhold)):errors['UNCLASSIFIED_BASE_CASES']+=len(set(bybase)-set(byled)-set(byhold))
 if (set(byled)-set(bybase)):errors['LEDGER_CASE_NOT_IN_BASE']+=len(set(byled)-set(bybase))
 if (set(byhold)-set(bybase)):errors['HELD_CASE_NOT_IN_BASE']+=len(set(byhold)-set(bybase))
 if set(byled)!=set(byres):errors['RESERVATION_MEMBERSHIP_DIFFERENCE']+=len(set(byled)^set(byres))
 candidate_bal=defaultdict(Decimal);base_bal=defaultdict(Decimal);held_bal=defaultdict(Decimal)
 event_rows=[];reserved_seen=set();docs=set();header_fields=Counter()
 for row in headers:
  h=row.get('proposed_target_header')
  if not isinstance(h,dict):errors['INVALID_HEADER_PAYLOAD']+=1;continue
  header_fields.update(h.keys())
  if not txt(h.get('AmbDokID')):errors['MISSING_DRAFT_AMBDOKID']+=1
  if txt(h.get('AmbDokID')) in docs:errors['DUPLICATE_DRAFT_AMBDOKID']+=1
  docs.add(txt(h.get('AmbDokID')))
 for case,row in bybase.items():
  typ=txt(row.get('type'))
  frm=(txt(row.get('from_account_type')),txt(row.get('from_account_id')))
  to=(txt(row.get('to_account_type')),txt(row.get('to_account_id')))
  try:q=amount(row.get('quantity'))
  except ValueError:errors['INVALID_BASE_QUANTITY']+=1;continue
  balancedelta(base_bal,frm,to,typ,q)
  if case in byhold:balancedelta(held_bal,frm,to,typ,q)
 for case,outer in byled.items():
  base_row=bybase.get(case)
  if base_row is None:continue
  try:
   row=getrow(outer)
   missing=SCHEMA_FIELDS-set(row)
   if missing:errors['MISSING_TARGET_COLUMNS']+=1;continue
   q=amount(row.get('Kolicina'))
   typ=txt(row.get('TipAmbalaze'))
   frm=account_pair(row,'Od');to=account_pair(row,'Na')
   kind=txt(row.get('VrstaKretanja'))
   dt=date_sort_key(row.get('Datum'))
   if kind not in KINDS:errors['INVALID_MOVEMENT_KIND']+=1
   if frm==to:errors['SELF_TRANSFER']+=1
   if not all((typ,frm[0],to[0])):errors['EMPTY_ACCOUNT_OR_PACKAGING_TYPE']+=1
   if frm[0].casefold() in ('firma','spoljnissvet','spoljn isvet') and frm[1]:errors['SYSTEM_ACCOUNT_WITH_ID']+=1
   if to[0].casefold()=='firma' and to[1]:errors['SYSTEM_ACCOUNT_WITH_ID']+=1
   if not txt(row.get('DokumentID')) or not txt(row.get('DokumentTIP')):errors['MISSING_SOURCE_DOCUMENT']+=1
   if txt(row.get('StornoOd')):errors['SOURCE_STORNO_NOT_EMPTY']+=1
   eq_fields=((q,amount(base_row.get('quantity'))),
              (typ,txt(base_row.get('type'))),
              (frm[0],txt(base_row.get('from_account_type'))),
              (frm[1],txt(base_row.get('from_account_id'))),
              (to[0],txt(base_row.get('to_account_type'))),
              (to[1],txt(base_row.get('to_account_id'))))
   if any(a!=b for a,b in eq_fields):errors['SOURCE_TARGET_PHYSICAL_MISMATCH']+=1
   res=byres.get(case)
   if not res:errors['RESERVATION_NOT_FOUND']+=1
   else:
    cid=txt(outer.get('reserved_AmbID_NOT_ALLOCATED'))
    rid=txt(res.get('candidate_AmbID_NOT_ALLOCATED'))
    if not cid or cid!=rid:errors['RESERVATION_ID_MISMATCH']+=1
    if cid in reserved_seen:errors['DUPLICATE_RESERVED_AMBID']+=1
    reserved_seen.add(cid)
    if txt(res.get('movement_kind_proposal'))!=kind:errors['KIND_DIFFERENCE_BETWEEN_PLANS']+=1
   approval_status[txt(outer.get('status'))]+=1
   for gate in outer.get('local_gates',[]) or []:local_gates[txt(gate)]+=1
   move_counts[kind]+=1
   balancedelta(candidate_bal,frm,to,typ,q)
   event_rows.append({'case':case,'date':dt,'ordinal':ordinal(outer),'kind':kind,'type':typ,
                      'from':frm,'to':to,'quantity':q})
  except (ValueError,TypeError,KeyError) as e:
   errors['BAD_TARGET_EVENT']+=1
 combined=defaultdict(Decimal,candidate_bal)
 for k,v in held_bal.items():combined[k]+=v
 mismatch=[k for k in (set(base_bal)|set(combined)) if base_bal[k]!=combined[k]]
 if mismatch:errors['BALANCE_CONSERVATION_WITH_HELD_MISMATCH_ACCOUNTS']=len(mismatch)
 crossing=Counter();crossing_register=[];balances=defaultdict(Decimal)
 # Date then original legacy row ordinal is a stable historical replay hypothesis.
 # Same-day true chronology is NOT proven when only day precision exists.
 for ev in sorted(event_rows,key=lambda x:(x['date'],x['ordinal'],x['case'])):
  for typ,aid,sgn in ((ev['from'][0],ev['from'][1],-1),(ev['to'][0],ev['to'][1],1)):
   if typ.casefold()!='kooperant':continue
   k=acc_key(typ,aid,ev['type']);b=balances[k];q=ev['quantity'];after=b+sgn*q
   if ev['kind']=='IZDATA_PRAZNA' and sgn==1 and b<0:
    old_due=min(-b,q);remainder=q-old_due
    crossing['ISSUANCE_STARTS_WITH_NEGATIVE_BALANCE']+=1
    if remainder>0:crossing['ISSUANCE_CROSSES_ZERO']+=1
    crossing_register.append({'physical_case_id':ev['case'],'classification':'ISSUANCE_BEFORE_NEGATIVE_TO_ZERO',
                               'pre_balance':str(b),'quantity':str(q),'potential_legacy_due_part':str(old_due),
                               'potential_new_debit_part':str(remainder),'post_balance':str(after),
                               'date':ev['date'].isoformat(),'source_ordinal':ev['ordinal']})
   if ev['kind']=='POVRAT_PRAZNE' and sgn==-1 and after<0:
    crossing['RETURN_ENDS_WITH_NEGATIVE_BALANCE']+=1
    if b>0:crossing['RETURN_CROSSES_ZERO']+=1
    crossing_register.append({'physical_case_id':ev['case'],'classification':'RETURN_TO_NEGATIVE_BALANCE',
                               'pre_balance':str(b),'quantity':str(q),
                               'potential_existing_debit_repaid':str(min(max(b,Decimal(0)),q)),
                               'potential_new_legacy_due':str(-min(after,Decimal(0))),
                               'post_balance':str(after),'date':ev['date'].isoformat(),
                               'source_ordinal':ev['ordinal']})
   balances[k]=after
 negative_end=sum(1 for (typ,_,_),v in balances.items() if typ=='kooperant' and v<0)
 summary={
  'format':'AGRIx-MIG15-SEMANTIC-REPLAY-v1','status':'READ_ONLY_REPLAY_COMPLETED',
  'source_and_target_workbooks_modified':False,
  'source_counts':{'legacy_rows':7815,'physical_events':len(base),'adapter_candidates':len(led),
                   'new_revers_header_drafts':len(headers),'unmodelled_held_events':len(holds)},
  'fingerprints_verified':len(EXPECTED),'candidate_physical_integrity_errors':dict(sorted(errors.items())),
  'candidate_kind_counts':dict(sorted(move_counts.items())),
  'balance_conservation_reconstructed_plus_holds_matches_all_physical':not bool(mismatch),
  'historical_crossings':dict(sorted(crossing.items())),
  'historical_crossing_cases_private':len(crossing_register),
  'negative_kooperant_account_type_ending_replay_count_excluding_held':negative_end,
  'same_day_order':'DATE_THEN_LEGACY_SOURCE_ORDINAL__NOT_PROVEN_INTRADAY',
  'policy_interpretation':'Old negative kooperant physical balance = historical company obligation; not automatic legal title documentation',
  'existing_adapter_rows_status_counts':dict(sorted(approval_status.items())),
  'header_field_names':sorted(header_fields),
  'readiness':'NOT_APPROVED_FOR_CANONICAL_XLSM_IMPORT',
  'timing_precision_warning':'Same-day original row order is a deterministic hypothesis, not proven business ordering',
  'remaining_gates':['CROSSING_CLASSIFICATION_NOT_A_CANONICAL_EVENT','HISTORICAL_FOREIGN_LIABILITY_REQUIRES_PROVEN_LEDGER_POLICY',
                     'SIX_UNMODELLED_DOCUMENTS_HELD','Q00700_STORNO_CONFLICT_HELD',
                     '335_CANCELLED_OLD_LEGS_ARCHIVE_ONLY','48_REVERS_HEADERS_ARE_NOT_ALLOCATED_CANONICAL_IDS'],
  'customer_data_in_summary':False
 }
 out.mkdir(parents=True,exist_ok=True)
 (out/'MIG15_semantic_replay_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 with (out/'MIG15_crossings_PRIVATE.jsonl').open('w',encoding='utf-8') as f:
  for row in crossing_register:f.write(json.dumps(row,ensure_ascii=False)+'\n')
 print('MIG-15 REPLAY COMPLETE — candidate/held source balance conserved:',not bool(mismatch))
 print('Physical integrity errors:',dict(sorted(errors.items())))
 print('Crossing candidates (private details stay local):',len(crossing_register))
 print('Summary:',out/'MIG15_semantic_replay_summary.json')
 print('IMPORT NOT EXECUTED: historical policy is not yet an AMB-10 canonical ledger.')
 return 0 if not errors else 2


def main():
 ap=argparse.ArgumentParser(description=__doc__)
 ap.add_argument('--project-dir',type=Path,required=True)
 ap.add_argument('--out',type=Path,required=True)
 ap.add_argument('--inventory',type=Path)
 a=ap.parse_args()
 try:return run(a.project_dir,a.out,a.inventory)
 except Exception as e:
  print(f'MIG-15 FAIL-CLOSED: {type(e).__name__}: {str(e)}',file=sys.stderr)
  return 1

if __name__=='__main__':sys.exit(main())
