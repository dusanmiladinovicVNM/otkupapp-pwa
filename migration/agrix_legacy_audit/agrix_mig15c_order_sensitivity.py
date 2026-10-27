#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-15C: source-order sensitivity + six held case operator brief (read-only).

No XLSM, source or Git modifications. Private row details remain in local output.
This is an evidence tool, not a canonical AMB-10 import.
"""
from __future__ import annotations
import argparse,collections,csv,datetime,hashlib,json,sys
from decimal import Decimal
from pathlib import Path
from agrix_mig15_semantic_replay import EXPECTED,read_jsonl,digest,txt,amount,ordinal,date_sort_key

META_FILES={
 'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl':('804b97f62797fcbd439779b80ea1f4e550237cc9e45ac00d06d1c8c6d330df1d',8),
 'MigrationAmb10ReversDocumentPlan_v2/AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl':('66d6ac47e56b8f881bc1957aa5964d0a5736f747721447bcd1ed9f5846de37b1',48),
}

def verified(root,rel,spec):
 path=root/rel
 if not path.is_file():raise ValueError('Required evidence file missing: '+rel)
 if digest(path)!=spec[0]:raise ValueError('Evidence SHA256 changed: '+rel)
 rows=read_jsonl(path)
 if len(rows)!=spec[1]:raise ValueError('Evidence row count changed: '+rel)
 return rows

def prepare(root):
 data={rel:verified(root,rel,spec) for rel,spec in {**EXPECTED,**META_FILES}.items()}
 def get(suffix):return next(v for k,v in data.items() if k.endswith(suffix))
 base=get('AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl')
 led=get('AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl')
 holds=get('AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl')
 bybase={txt(x.get('physical_case_id')):x for x in base}
 rows=[]
 for x in led:
  case=txt(x.get('physical_case_id'))
  if case not in bybase:raise ValueError('Ledger case missing from base')
  d=x.get('target_column_values_WITHOUT_AMBID')
  if not isinstance(d,dict):raise ValueError('Missing target column values')
  q=amount(d.get('Kolicina'))
  o=bybase[case]
  if q!=amount(o.get('quantity')):raise ValueError('Quantity mismatch')
  dt=date_sort_key(d.get('Datum'))
  rows.append({'case':case,'date':dt.date(),'moment':dt,'ord':ordinal(x),'q':q,
   'kind':txt(d.get('VrstaKretanja')),'typ':txt(d.get('TipAmbalaze')),
   'from':(txt(d.get('OdNalogTip')),txt(d.get('OdNalogID'))),
   'to':(txt(d.get('NaNalogTip')),txt(d.get('NaNalogID')))})
 return rows,holds,bybase,get('AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl'),get('AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl')

def replay(rows,reverse=False):
 # A conservative same-DAY test; timestamps in old ledger may lack meaningful ordering.
 byday=collections.defaultdict(list)
 for ev in rows:byday[ev['date']].append(ev)
 bal=collections.defaultdict(Decimal);parts={};days=0
 for day in sorted(byday):
  evs=sorted(byday[day],key=lambda e:(e['moment'],e['ord'],e['case']),reverse=reverse)
  if len(evs)>1:days+=1
  for e in evs:
   q=e['q'];typ=e['typ'].casefold();case=e['case']
   foreign_in=Decimal(0);foreign_return=Decimal(0)
   for acc,direction in ((e['from'],-1),(e['to'],1)):
    if acc[0].casefold()!='kooperant':continue
    key=(acc[0].casefold(),acc[1].casefold(),typ)
    before=bal[key];after=before+q*direction
    if direction==-1:
     foreign_in=max(Decimal(0),-after)-max(Decimal(0),-before)
    else:
     foreign_return=min(max(Decimal(0),-before),q)
    bal[key]=after
   parts[case]=(foreign_in,foreign_return)
 return parts,bal,days

def run(root,out):
 root=root.resolve();out=out.resolve()
 if out==root:raise ValueError('Output must not equal project root')
 for rel in {**EXPECTED,**META_FILES}:
  source=root/rel
  if out==source or out in source.parents or source in out.parents:
   raise ValueError('Output cannot overwrite or be nested in a source evidence path')
 rows,holds,bybase,exceptions,revers=prepare(root)
 base,balances,days=replay(rows)
 reverse,revbal,_=replay(rows,True)
 if balances!=revbal:raise ValueError('End balances changed by intraday order')
 changed=[e for e in rows if base[e['case']]!=reverse[e['case']]]
 if len({e['case'] for e in rows})!=len(rows):raise ValueError('Duplicate physical cases')
 cases_by_date=collections.Counter(e['date'].isoformat() for e in changed)
 psrc=out/'MIG15C_order_sensitive_PRIVATE.csv'
 out.mkdir(parents=True,exist_ok=True)
 with psrc.open('w',newline='',encoding='utf-8') as f:
  w=csv.writer(f);w.writerow(['physical_case_id','date','source_ordinal','kind','type','base_new_liability','base_return_liability','reverse_new_liability','reverse_return_liability'])
  for ev in changed:
   x,y=base[ev['case']],reverse[ev['case']]
   w.writerow([ev['case'],ev['date'],ev['ord'],ev['kind'],ev['typ'],str(x[0]),str(x[1]),str(y[0]),str(y[1])])
 # Local-only operator register: private case IDs + former documentary evidence. No fabrication.
 with (out/'MIG15C_exception_register_PRIVATE.csv').open('w',newline='',encoding='utf-8') as f:
  w=csv.writer(f);w.writerow(['category','physical_case_id','source_ordinals','old_DokumentTIP','old_DokumentID_or_label','reason','operator_note'])
  for h in holds:
   c=txt(h.get('physical_case_id'));b=bybase.get(c,{})
   w.writerow(['UNMODELLED',c,txt(h.get('source_row_ordinals')),txt(h.get('legacy_DokumentTIP')) or txt(b.get('old_doc_type')),
     txt(h.get('legacy_DokumentID')) or txt(b.get('old_doc_label')),txt(h.get('reason')),''])
  for e in exceptions:
   w.writerow(['STORNO_EXCEPTION','',txt(e.get('source_row_ordinal')),'','',txt(e.get('status')),''])
  # The full 48 revers docs already have proven scoped old identity, but new AmbDokID is migration generated.
  for x in revers:
   w.writerow(['REVERS_DRAFT',txt(x.get('physical_case_id')),txt(x.get('source_row_ordinals')),txt(x.get('old_document_type')),
     txt(x.get('old_label')), 'NEW_CANONICAL_HEADER_ID_REQUIRED',''])
 count=collections.Counter({'events_total':len(rows),'held_documents':len(holds),'revers_drafts':len(revers),'storno_exceptions':len(exceptions)})
 for c,(new,ret) in base.items():
  if new>0:count['events_with_derived_foreign_entry']+=1
  if ret>0:count['events_with_derived_foreign_return']+=1
  if new>0 and ret>0:count['dual_exposure_event_cases']+=1
 for x in changed:
  if base[x['case']][0]!=reverse[x['case']][0]:count['source_order_sensitive_new_liability_events']+=1
  if base[x['case']][1]!=reverse[x['case']][1]:count['source_order_sensitive_liability_return_events']+=1
 summary={'format':'AGRIx-MIG15C-ORDER-SENSITIVITY-v1','status':'READ_ONLY_POLICY_DIAGNOSTIC',
  'candidate_events_checked':len(rows),'held_events':len(holds),'revers_drafts':len(revers),
  'storno_exception_register_rows':len(exceptions),'same_day_groups_with_multiple_events':days,
  'derived_event_counts_date_then_source':dict(sorted(count.items())),
  'order_sensitive_event_cases':len(changed),
  'order_sensitive_distinct_dates':len(cases_by_date),
  'historical_ending_balances_independent_of_same_day_order':True,
  'strict_exact_input_sha256_verified':len(EXPECTED)+len(META_FILES),
  'customer_data_in_summary':False,
  'cannot_equate_legacy_signed_balance_to_new_physical_balance':True,
  'interpretation':'Counterfactual split amounts, not canonical AMB-10 ledger entries or established legal ownership',
  'import_approved':False,'xlsm_modified':False,
  'next_action':'Review private six undocumented physical events, original Q00700 disposition and any order-sensitive dates. Define exact canonical foreign-liability migration semantics before writing XLSM.'}
 (out/'MIG15C_summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
 print('MIG15C READ-ONLY COMPLETE; candidate events:',len(rows),'day-order-sensitive:',len(changed),'held:',len(holds))
 print('Send only:',out/'MIG15C_summary.json')
 print('PRIVATE local operator register:',out/'MIG15C_exception_register_PRIVATE.csv')
 return summary

def main():
 ap=argparse.ArgumentParser(description=__doc__)
 ap.add_argument('--project-dir',required=True,type=Path);ap.add_argument('--out',required=True,type=Path)
 a=ap.parse_args()
 try:run(a.project_dir,a.out)
 except Exception as e:
  print('MIG15C BLOCKED:',type(e).__name__,str(e),file=sys.stderr);sys.exit(2)
if __name__=='__main__':main()
