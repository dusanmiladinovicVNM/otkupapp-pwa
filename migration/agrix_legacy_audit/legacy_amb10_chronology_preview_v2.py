#!/usr/bin/env python3
"""MIG-05.5: read-only chronological stress test of proven LEGACY physical vectors.

NOT a chronological canonical AMB-10 ledger. Historical canceled movements,
legal title, opening stock and intraday ordering cannot be recovered from a
net-active-only ledger. All amounts/accounts remain in PRIVATE files.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import re
import sys

from legacy_audit import AuditError, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as shadow
import legacy_seven_hold_rescue_v2 as seven

VERSION='MIG-05.5'
EXPECTED_STATUS='SEVEN_HOLD_PHYSICAL_EVIDENCE_ONLY_NOT_AMB10_APPROVED'
ACCOUNT_COLUMNS=('account_type','account_id','tip_ambalaze','net_active_ending',
 'day_best_zero_opening_floor','day_worst_zero_opening_floor',
 'minimum_opening_best_case','minimum_opening_worst_case',
 'uncertain_same_day_event_groups','date_unknown_events','ownership','opening_approved')
CASE_COLUMNS=('physical_case_id','source_row_ordinals','legacy_Datum',
 'parsed_date','date_status','row_ordinal_min','row_ordinal_max',
 'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','TipAmbalaze','Kolicina',
 'VrstaKretanja_PROPOSAL','evidence_status','same_source_leg_dates',
 'source_date_decreasing_vs_previous','chronology_approved')


def date_key(raw):
    """Convert a date to a day, with NO inferred time-of-day.

    Excel 1900 serial uses 1899-12-30 for days >=61. Reject invalid or
    ambiguous date text; callers must keep DATE_UNRESOLVED on hold.
    """
    t=trim(raw)
    if not t:return None
    try:
        val=Decimal(t)
    except (InvalidOperation,ValueError):
        val=None
    if val is not None:
        if not val.is_finite() or val < 61 or val > 90000:return None
        return date(1899,12,30)+timedelta(days=int(val))
    try:
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}',t):return date.fromisoformat(t)
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?',t):
            return datetime.fromisoformat(t).date()
        if re.fullmatch(r'\d{1,2}\.\d{1,2}\.\d{4}\.?',t):
            return datetime.strptime(t.rstrip('.'),'%d.%m.%Y').date()
        if re.fullmatch(r'\d{1,2}\.\d{1,2}\.\d{4}\.?(?:\s+)\d{1,2}:\d{2}(?::\d{2})?',t):
            return date.fromisoformat(date(*[int(z) for z in t.split(' ')[0].rstrip('.').split('.')[::-1]]).isoformat())
    except (ValueError,OverflowError):return None
    return None


def _verify_inputs(stage:Path,expanded_dir:Path,seven_dir:Path):
    if not all(x.is_dir() for x in (stage,expanded_dir,seven_dir)):
        raise AuditError('Missing MIG-02 / MIG-05.3 / MIG-05.4 input directory')
    expected_report,_,expected_events,_=seven.build(stage,expanded_dir)
    recorded_path=seven_dir/'agrix_seven_hold_summary.json'
    if not recorded_path.is_file():raise AuditError('MIG-05.4 summary missing')
    recorded=json.loads(recorded_path.read_text(encoding='utf-8'))
    if recorded!=expected_report:raise AuditError('MIG-05.4 summary is inconsistent with pinned source')
    if recorded.get('format')!='MIG-05.4' or recorded.get('status')!=EXPECTED_STATUS:
        raise AuditError('Wrong upstream report format')
    if (recorded.get('new_physical_events_total')!=4675 or
        recorded.get('new_physical_source_legs')!=7479 or
        recorded.get('remaining_active_hold_source_ordinals')!=[700] or
        recorded.get('historical_storno_rows_unconverted')!=335 or
        recorded.get('ending_balance_mismatch_account_types_after')!=2 or
        recorded.get('unexplained_account_residual_count')!=0):
        raise AuditError('Pinned source reconciliation changed')
    private_path=seven_dir/'expanded_plus_held_physical_cases_PRIVATE.jsonl'
    if not private_path.is_file():raise AuditError('MIG-05.4 private evidence file missing')
    signature=hashlib.sha256()
    event_count=0
    with private_path.open(encoding='utf-8',newline='') as inp:
        for line in inp:
            if not line.endswith('\n'):raise AuditError('MIG-05.4 private JSONL not newline terminated')
            if event_count>=len(expected_events):raise AuditError('Extra evidence event in MIG-05.4')
            expected=json.dumps(expected_events[event_count],ensure_ascii=False,sort_keys=True)+'\n'
            if line!=expected:raise AuditError('MIG-05.4 evidence record content/ordering changed')
            signature.update(line.encode('utf-8'))
            event_count+=1
    if event_count!=len(expected_events):raise AuditError('Missing MIG-05.4 private event')
    return recorded,expected_events,signature.hexdigest()


def chronological_stress(events, source_dates):
    """Date-scoped best/worst hypothetical prefix floors, opening always UNKNOWN.

    Sort by documented DATE, never assert a time-of-day. For each account/day,
    credit-first is the optimistic min-prefix, debit-first conservative. These
    bounds are risk diagnostics only, NOT an opening stock instruction.
    """
    detail=[];used=[];bad_date=0;date_disagreements=0
    status_counts=Counter();kind_counts=Counter();source_type=Counter()
    account_day=defaultdict(lambda:defaultdict(lambda:[Decimal(0),Decimal(0),0]))
    account_date_unknown=Counter();source_inversions=0
    for e in events:
        if e.get('canonical_AMB10_import_authorized')!='NO':
            raise AuditError('Prior physical evidence wrongly declares import approval')
        status=e['evidence_status'];status_counts[status]+=1
        kind_counts[e['VrstaKretanja_PROPOSAL']]+=1
        source_type[e['event_provenance']]+=1
        ids=[int(v) for v in e['source_row_ordinals'].split(';')]
        if not ids or len(ids)!=len(set(ids)) or len(ids)>2 or any(n<1 or n>7815 for n in ids):
            raise AuditError('Invalid physical lineage')
        used+=ids
        if (e.get('ownership')!='UNPROVEN' or e.get('opening_balance')!='UNPROVEN' or
           e.get('historical_storno')!='NOT_REPLAYED'):
            raise AuditError('Source evidence has unexpected authority for ownership/opening/storno')
        d=date_key(e['legacy_Datum'])
        parsed=[date_key(source_dates[n]) for n in ids]
        same=bool(d and all(z==d for z in parsed))
        if not same:date_disagreements+=1
        if d is None:bad_date+=1
        ord_min=min(ids);ord_max=max(ids)
        qty=f.dec(e['Kolicina'])
        if qty is None or qty<=0:raise AuditError('Invalid physical event amount')
        typ=trim(e['TipAmbalaze']);key_from=(e['OdNalogTip'],e['OdNalogID'],typ)
        key_to=(e['NaNalogTip'],e['NaNalogID'],typ)
        if not all(key_from) or not all(key_to) or key_from==key_to:
            raise AuditError('Invalid physical transfer account identity')
        for acc,delta in ((key_from,-qty),(key_to,qty)):
            if d is None:account_date_unknown[acc]+=1
            else:
                slots=account_day[acc][d]
                if delta<0:slots[1]+=delta
                else:slots[0]+=delta
                slots[2]+=1
        detail.append(dict(physical_case_id=e['physical_case_id'],
            source_row_ordinals=e['source_row_ordinals'],legacy_Datum=e['legacy_Datum'],
            parsed_date=d.isoformat() if d else '',date_status='DATE_EXACT_DAY_ONLY' if same else 'DATE_CONFLICT_OR_UNKNOWN',
            row_ordinal_min=ord_min,row_ordinal_max=ord_max,
            OdNalogTip=key_from[0],OdNalogID=key_from[1],NaNalogTip=key_to[0],
            NaNalogID=key_to[1],TipAmbalaze=typ,Kolicina=e['Kolicina'],
            VrstaKretanja_PROPOSAL=e['VrstaKretanja_PROPOSAL'],
            evidence_status=status,same_source_leg_dates='YES' if same else 'NO',
            source_date_decreasing_vs_previous='',chronology_approved='NO'))
    if len(used)!=7479 or len(events)!=4675 or len(used)!=len(set(used)):
        raise AuditError('Physical event lineage partition invalid')
    seq=sorted(range(len(detail)),key=lambda i:detail[i]['row_ordinal_min'])
    prior_date=None
    for i in seq:
        d=date_key(detail[i]['legacy_Datum'])
        if prior_date is not None and d is not None and d<prior_date:
            source_inversions+=1;detail[i]['source_date_decreasing_vs_previous']='YES'
        elif d is None:detail[i]['source_date_decreasing_vs_previous']='UNKNOWN'
        else:detail[i]['source_date_decreasing_vs_previous']='NO'
        if d is not None:prior_date=d
    accounts=[];bestcount=worstcount=0;samedayaccounts=0;endneg=0
    for acc in sorted(set(account_day)|set(account_date_unknown)):
        net=Decimal(0);min_best=Decimal(0);min_worst=Decimal(0);same_day=0
        for d, (credit,debit,count) in sorted(account_day[acc].items()):
            if count>1:same_day+=1
            min_best=min(min_best,net,net+credit+debit)
            min_worst=min(min_worst,net,net+debit)
            net+=credit+debit
        min_open_best=max(Decimal(0),-min_best)
        min_open_worst=max(Decimal(0),-min_worst)
        if min_open_best:bestcount+=1
        if min_open_worst:worstcount+=1
        if same_day:samedayaccounts+=1
        if net<0:endneg+=1
        accounts.append(dict(account_type=acc[0],account_id=acc[1],tip_ambalaze=acc[2],
            net_active_ending=str(net),day_best_zero_opening_floor=str(min_best),
            day_worst_zero_opening_floor=str(min_worst),
            minimum_opening_best_case=str(min_open_best),
            minimum_opening_worst_case=str(min_open_worst),
            uncertain_same_day_event_groups=same_day,date_unknown_events=account_date_unknown[acc],
            ownership='UNPROVEN',opening_approved='NO'))
    report={
        'physical_case_count':len(events),
        'physical_legacy_source_legs':len(used),
        'private_evidence_lineage_unique':len(set(used))==len(used),
        'date_unknown_event_count':bad_date,
        'source_legs_date_conflict_or_unknown_events':date_disagreements,
        'legacy_row_order_date_descents':source_inversions,
        'hypothetical_best_day_order_accounts_with_zero_opening_deficit':bestcount,
        'hypothetical_worst_day_order_accounts_with_zero_opening_deficit':worstcount,
        'account_type_balances_with_multiple_events_same_day':samedayaccounts,
        'account_type_balances_with_negative_net_without_opening':endneg,
        'active_evidence_status_counts':dict(sorted(status_counts.items())),
        'physical_business_kind_counts':dict(sorted(kind_counts.items())),
        'event_provenance_counts':dict(sorted(source_type.items()))
    }
    return report,detail,accounts


def build(stage:Path,expanded:Path,seven_dir:Path):
    recorded,events,digest=_verify_inputs(stage,expanded,seven_dir)
    forensic,raw=f.analyze(stage)
    if forensic.get('legacy_state_counts')!={'ACTIVE':7480,'STORNO':335}:
        raise AuditError('Legacy source status population changed')
    sources={r['source_row_ordinal']:r for r in raw}
    proof,details,accounts=chronological_stress(events,{n:r['Datum'] for n,r in sources.items()})
    used={int(t) for e in events for t in e['source_row_ordinals'].split(';')}
    if used!={n for n,r in sources.items() if f.is_active(r)}-{700}:
        raise AuditError('Current source active partition does not match physical projection')
    if proof['physical_legacy_source_legs']!=recorded['new_physical_source_legs']:
        raise AuditError('Source physical leg counts drifted')
    public={
        'format':VERSION,'status':'NET_ACTIVE_CHRONOLOGY_STRESS_ONLY_NOT_IMPORT_APPROVED',
        'source_sha256':recorded['source_sha256'],
        'schema_sha256':recorded['schema_sha256'],
        'parent_expanded_physical_evidence_sha256':recorded['parent_expanded_physical_evidence_sha256'],
        'input_mig05_4_evidence_sha256':digest,
        **proof,
        'still_held_active_ordinal':700,
        'still_held_active_legs':1,
        'historical_storno_rows_not_replayed':335,
        'legacy_negative_account_type_balances':recorded['legacy_negative_account_types_unchanged'],
        'unexplained_ending_balance_residual_accounts':0,
        'business_ownership_verified':False,
        'canonical_amb10_import_approved':False,
        'notes':[
            'BEST/WORST are hypothetical starting-from-zero intra-day order stress tests, never historical balances.',
            'Legacy stornirano rows (335) are excluded from net-active sequence; actual lifecycle chronological reconstruction is NOT proven.',
            'The apparent date order may differ from ledger row insertion order; no within-day time can be assumed.',
            'If any date is unresolved, reported zero-opening balance-floor diagnostics are incomplete and not a pass.',
            'Calculated minimum hypothetical opening is NOT evidence of existing physical stock or permission to manufacture opening balances.',
            'Only sums/dates/counts appear publicly; PRIVATE files contain sensitive source identifiers and quantities.',
            'One active Q-00700 and six reconstructed-but-document-unknown cases remain blocked for canonical import.',
        ]}
    return public,details,accounts


def render(s):
    keys=['physical_case_count','physical_legacy_source_legs','date_unknown_event_count',
        'source_legs_date_conflict_or_unknown_events','legacy_row_order_date_descents',
        'hypothetical_best_day_order_accounts_with_zero_opening_deficit',
        'hypothetical_worst_day_order_accounts_with_zero_opening_deficit',
        'account_type_balances_with_multiple_events_same_day',
        'account_type_balances_with_negative_net_without_opening',
        'active_evidence_status_counts','physical_business_kind_counts','still_held_active_legs',
        'historical_storno_rows_not_replayed','legacy_negative_account_type_balances',
        'canonical_amb10_import_approved']
    return ('AGRIX MIG-05.5 / READ ONLY CHRONOLOGY STRESS, NO AMB-10\n'+
        ''.join(k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in keys)+
        'Do not publish PRIVATE files or use hypothetical openings for imports.\n')


def cli(argv=None):
    p=argparse.ArgumentParser(description='MIG-05.5 hypothetical date/ordering stress test')
    for x in ('stage','expanded','seven','out'):p.add_argument('--'+x,required=True,type=Path)
    a=p.parse_args(argv)
    stage=a.stage.resolve();expanded=a.expanded.resolve();seven_dir=a.seven.resolve();out=a.out.resolve()
    if out.exists() or any(out==x or out in x.parents or x in out.parents for x in (stage,expanded,seven_dir)):
        raise AuditError('Output folder already exists/overlaps input folder; use fresh output')
    report,details,accounts=build(stage,expanded,seven_dir)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_amb10_chronology_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (out/'agrix_amb10_chronology_summary.txt').write_text(render(report),encoding='utf-8')
        f.private_csv(out/'chronology_cases_PRIVATE.csv',details,CASE_COLUMNS)
        f.private_csv(out/'chrono_balance_stress_PRIVATE.csv',accounts,ACCOUNT_COLUMNS)
    except BaseException:
        for x in out.iterdir():
            if x.is_file():x.unlink()
        out.rmdir()
        raise
    print(render(report));print('Upload ONLY agrix_amb10_chronology_summary.json/txt.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,TypeError,KeyError,OverflowError,InvalidOperation) as exc:
        print('MIG-05.5 STOP:',exc,file=sys.stderr);raise SystemExit(2)
