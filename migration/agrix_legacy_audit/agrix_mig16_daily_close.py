#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG16: closed-day legacy packaging policy, local-only release diagnostics.

This program produces evidence for the agreed historical daily-closing policy.
It does NOT manufacture AMB-10 book events or alter an XLSM. It is deliberately
not a canonical import writer: an inferred liability is not a source transfer.
"""
from __future__ import annotations
import argparse
from collections import Counter,defaultdict
import csv,hashlib,json,sys,zipfile
from pathlib import Path
from decimal import Decimal
from datetime import date
from xml.etree import ElementTree as ET
from agrix_mig15_semantic_replay import EXPECTED,digest,read_jsonl,txt,amount,date_sort_key

EXTRA={
 'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl':('804b97f62797fcbd439779b80ea1f4e550237cc9e45ac00d06d1c8c6d330df1d',8),
 'MigrationAmb10ReversDocumentPlan_v2/AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl':('66d6ac47e56b8f881bc1957aa5964d0a5736f747721447bcd1ed9f5846de37b1',48),
}

NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def probe_workbook(path:Path):
    from agrix_mig11_write_links_ooxml import table_locations,parse_ref,q
    if not path.is_file():raise ValueError('Input workbook missing')
    with zipfile.ZipFile(path) as z:
        if z.testzip() or 'xl/vbaProject.bin' not in z.namelist():raise ValueError('XLSM ZIP/VBA failed')
        loc,_=table_locations(z)
        result={}
        cols_ref={
          'tblAmbalaza':['AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd'],
          'tblAmbalazaDokument':['AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Napomena','Stornirano','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','IspravkaOdID','ZamenjenSaID']}
        for tab in ('tblAmbalaza','tblAmbalazaDokument'):
            if tab not in loc:raise ValueError(f'Missing target {tab}')
            tp,sp,root=loc[tab]
            cols=[x.attrib['name'] for x in root.find(q('tableColumns')).findall(q('tableColumn'))]
            if cols!=cols_ref[tab]:raise ValueError(f'{tab} canonical column order mismatch')
            c1,r1,c2,r2=parse_ref(root.attrib['ref'])
            if r2 not in (r1,r1+1):raise ValueError(f'{tab} not empty; overwrite forbidden')
            sheet=ET.fromstring(z.read(sp));occupied=0
            for row in sheet.findall(q('sheetData')+'/'+q('row')):
                idx=int(row.attrib.get('r','0'))
                if idx<=r1:continue
                for cell in row.findall(q('c')):
                    pos=cell.attrib.get('r','');col=''.join(x for x in pos if x.isalpha())
                    n=0
                    for x in col:n=n*26+ord(x.upper())-64
                    if c1<=n<=c2:
                        if cell.find(q('v')) is not None and cell.find(q('v')).text not in (None,''):occupied+=1
                        if cell.find(q('is')) is not None and ''.join(t.text or '' for t in cell.findall('.//'+q('t'))).strip():occupied+=1
            if occupied:raise ValueError(f'{tab} has occupied target cells; overwrite forbidden')
            result[tab]={'schema_exact':True,'empty':True,'width':len(cols)}
        result['vba_present']=True;result['xlsm_sha256']=digest(path)
        return result

def make_rows(project:Path):
    datasets={}
    for rel,(h,n) in {**EXPECTED,**EXTRA}.items():
        p=project/rel
        if not p.is_file():raise ValueError('Missing source payload: '+rel)
        if digest(p)!=h:raise ValueError('Known source evidence changed: '+rel)
        d=read_jsonl(p)
        if len(d)!=n:raise ValueError('Source evidence count changed: '+rel)
        datasets[rel]=d
    def get(name):return datasets[next(k for k in datasets if k.endswith('/'+name))]
    base=get('AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl')
    led=get('AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl')
    reserves=get('AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl')
    holds=get('AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl')
    revers=get('AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl')
    exc=get('AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl')
    bybase={txt(r.get('physical_case_id')):r for r in base}
    byres={txt(r.get('physical_case_id')):r for r in reserves}
    byhold={txt(r.get('physical_case_id')):r for r in holds}
    byled={txt(r.get('physical_case_id')):r for r in led}
    if len(bybase)!=4675 or len(byled)!=4669 or len(byhold)!=6 or len(byres)!=4669:raise ValueError('Duplicate physical IDs detected')
    if set(bybase)!=set(byled)|set(byhold) or set(byled)&set(byhold) or set(byres)!=set(byled):raise ValueError('Incomplete physical case partition')
    rows=[];dates=defaultdict(set)
    for case,x in byled.items():
        d=x.get('target_column_values_WITHOUT_AMBID')
        if not isinstance(d,dict):raise ValueError('Ledger row not an object')
        b=bybase[case];r=byres[case]
        q=amount(d.get('Kolicina'))
        if q!=amount(b.get('quantity')):raise ValueError('Source quantity differs')
        fm=(txt(d.get('OdNalogTip')),txt(d.get('OdNalogID')))
        to=(txt(d.get('NaNalogTip')),txt(d.get('NaNalogID')))
        bf=(txt(b.get('from_account_type')),txt(b.get('from_account_id')))
        bt=(txt(b.get('to_account_type')),txt(b.get('to_account_id')))
        typ=txt(d.get('TipAmbalaze'))
        if fm!=bf or to!=bt or typ!=txt(b.get('type')):raise ValueError('Base physical transfer mismatch')
        if txt(x.get('reserved_AmbID_NOT_ALLOCATED'))!=txt(r.get('candidate_AmbID_NOT_ALLOCATED')):raise ValueError('Reserved ID mismatch')
        if txt(d.get('VrstaKretanja'))!=txt(r.get('movement_kind_proposal')):raise ValueError('Movement kind changed')
        day=date_sort_key(d.get('Datum')).date()
        rows.append({'case':case,'day':day,'q':q,'typ':typ,'kind':txt(d.get('VrstaKretanja')),'from':fm,'to':to})
        dates[day].add(case)
    return rows,holds,revers,exc,len(datasets),bybase

def compute_daily_policy(rows):
    """Do not inspect intraday ordering. Return candidate-only policy diagnostics.

    A historic signed account balance is NOT AMB-10 physical stock.
    The required new obligation is a target, not a manufactured movement.
    """
    per=defaultdict(lambda: {'delta':Decimal(0),'issued':Decimal(0),'outflow':Decimal(0),'n':0})
    for r in rows:
        for acc,sgn in ((r['from'],-1),(r['to'],1)):
            if acc[0].casefold()!='kooperant':continue
            k=(acc[1].casefold(),r['typ'].casefold(),r['day'])
            z=per[k];z['delta']+=r['q']*sgn;z['n']+=1
            if sgn==1 and r['kind']=='IZDATA_PRAZNA':z['issued']+=r['q']
            if sgn==-1:z['outflow']+=r['q']
    for k,z in per.items():
        if z['n']==0:raise ValueError('Unexpected empty date account group')
    last=defaultdict(Decimal);daily=[];summary=Counter();hold_qty=Decimal(0)
    for account,typ,day in sorted(per,key=lambda k:(k[2],k[0],k[1])):
        z=per[(account,typ,day)];key=(account,typ)
        before=last[key];after=before+z['delta'];last[key]=after
        old_debt=max(Decimal(0),-before);new_debt=max(Decimal(0),-after)
        raise_debt=max(Decimal(0),new_debt-old_debt)
        clear_debt=max(Decimal(0),old_debt-new_debt)
        issue_possible=z['issued'];source_out=z['outflow']
        uncovered_return=max(Decimal(0),clear_debt-issue_possible)
        unsupported_cover=max(Decimal(0),raise_debt-source_out)
        if raise_debt>0:summary['days_with_increased_historical_obligation']+=1
        if clear_debt>0:summary['days_with_decreased_historical_obligation']+=1
        if uncovered_return>0:summary['days_obligation_reduction_without_matching_issuance']+=1
        if unsupported_cover>0:summary['days_obligation_increase_without_matching_outflow']+=1
        if after<0:summary['negative_account_day_closes']+=1
        if after==0 and before<0:summary['debt_closed_at_day_end']+=1
        if before<0 and after>0:summary['crossed_negative_to_positive_at_day_end']+=1
        if before>0 and after<0:summary['crossed_positive_to_negative_at_day_end']+=1
        summary['account_type_day_closes']+=1
        daily.append({'day':day.isoformat(),'kooperant_id':account,'tip_ambalaze':typ,
                      'starting_signed_balance':str(before),'ending_signed_balance':str(after),
                      'historical_obligation_at_start':str(old_debt),'historical_obligation_at_close':str(new_debt),
                      'obligation_increase':str(raise_debt),'obligation_decrease':str(clear_debt),
                      'matching_candidate_issue_capacity':str(issue_possible),
                      'matching_candidate_outflow_capacity':str(source_out),
                      'not_assignable_decrease':str(uncovered_return),
                      'not_assignable_increase':str(unsupported_cover)})
    summary['distinct_kooperant_type_pairs']=len(last)
    summary['negative_account_type_endings_candidate_only']=sum(v<0 for v in last.values())
    return daily,dict(sorted(summary.items()))

def run(project:Path,out:Path,workbook:Path|None=None):
    project=project.resolve();out=out.resolve()
    if out.exists():raise ValueError('Output directory already exists — refusing overwrite')
    for rel in {**EXPECTED,**EXTRA}:
        source=project/rel
        if out==source or source in out.parents or out in source.parents:
            raise ValueError('Output cannot be nested inside or overwrite source evidence')
    rows,holds,revers,exceptions,verified,bybase=make_rows(project)
    daily,calc=compute_daily_policy(rows)
    out_of_policy_date_count=sum(r['day']>date(2026,8,31) for r in rows)
    workbook_report=probe_workbook(workbook.resolve()) if workbook else None
    hold_touches=sum(any(txt(bybase[txt(h.get('physical_case_id'))].get(k)).casefold()=='kooperant' for k in ('from_account_type','to_account_type')) for h in holds)
    # Strict release gates deliberately remain. This stage does not invent an
    # event for a daily obligation nor auto-approve unresolved documents.
    gates={
       'HELD_UNMODELLED_PHYSICAL_CASES':len(holds),
       'LEGACY_STORNO_EXCEPTIONS':len(exceptions),
       'HISTORICAL_REVERSES_REQUIRE_CANONICAL_HEADERS':len(revers),
       'Q00700_STORNO_DISPOSITION':1,
       'EVENTS_AFTER_HISTORICAL_POLICY_CUTOFF':out_of_policy_date_count,
       'SOURCE_CANCELLED_ROWS_ARCHIVED_ONLY':335,
       'OBLIGATION_CHANGES_NOT_YET_ENCODED_AS_PROVEN_AMB10_EVENTS':(calc.get('days_with_increased_historical_obligation',0)+calc.get('days_with_decreased_historical_obligation',0)),
       'REDUCTION_DAYS_WITHOUT_SAME_DAY_ELIGIBLE_ISSUANCE':calc.get('days_obligation_reduction_without_matching_issuance',0),
       'INCREASE_DAYS_WITHOUT_SAME_DAY_DOCUMENTED_OUTFLOW':calc.get('days_obligation_increase_without_matching_outflow',0),
    }
    out.mkdir(parents=True)
    with (out/'MIG16_daily_close_PRIVATE.csv').open('w',encoding='utf-8',newline='') as f:
        wr=csv.DictWriter(f,fieldnames=list(daily[0]) if daily else [])
        if daily:wr.writeheader();wr.writerows(daily)
    summary={'format':'AGRIx-MIG16-DAILY-CLOSE-POLICY-v1','status':'POLICY_ACCEPTED_DAILY_CLOSE_RECONCILED_CANDIDATE_ONLY',
      'operator_policy':'2026-08-31 and before: closing signed packaging balance per kooperant and packaging type is authoritative; intraday source order is not authoritative',
      'source_case_integrity_checked':4675,'candidate_physical_events_checked':len(rows),
      'physical_cases_not_yet_in_canonical_candidate':len(holds),'held_cases_touching_cooperant':hold_touches,
      'historical_revers_header_drafts':len(revers),'source_evidence_sha256_verified':verified,
      'historical_policy_cutoff_inclusive':'2026-08-31',
      'day_close_diagnostics_candidate_only':calc,
      'global_release_gates':gates,'canonical_import_approved':False,
      'reason':'Daily balances prove historic exposure but do not themselves instantiate AMB-10 foreign-ownership events. Six physical cases are missing canonical parents; archived storno and Q00700 remain held.',
      'xlsm_written':False,'workbook_preflight':workbook_report,
      'private_detail_file':'MIG16_daily_close_PRIVATE.csv','customer_values_excluded_from_summary':True}
    (out/'MIG16_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('MIG16 DAILY CLOSING POLICY: LOCAL DIAGNOSTIC COMPLETE')
    print('Cases:',len(rows),'held:',len(holds),'candidate-only closing groups:',calc.get('account_type_day_closes',0))
    print('No workbook modifications; do not import MIG-06.5 JSONL as-is.')
    print('SHARE ONLY:',out/'MIG16_summary.json')
    return summary

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project-dir',required=True,type=Path)
    parser.add_argument('--out',required=True,type=Path)
    parser.add_argument('--input-xlsm',type=Path)
    args=parser.parse_args()
    try:run(args.project_dir,args.out,args.input_xlsm)
    except Exception as e:
        print('MIG16 BLOCKED:',type(e).__name__,'—',str(e).split(':')[0],file=sys.stderr)
        return 2
    return 0
if __name__=='__main__':sys.exit(main())
