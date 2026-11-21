#!/usr/bin/env python3
"""MIG-04.3: account-level exact residual attribution for MIG-04.2 physical shadow.

Read-only. No AMB-10 canonical transfer, ownership, opening stock, or XLSM.
Public output contains counts only; private files have exact accounts/quantities/IDs.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as s

VERSION = 'MIG-04.3'


def _amounts(rows, include_ordinals=None):
    """Legacy net effects; preserve decimals and account identities exactly."""
    amounts=defaultdict(Decimal)
    per_account_rows=defaultdict(set)
    for r in rows:
        ordinal=r['source_row_ordinal']
        if include_ordinals is not None and ordinal not in include_ordinals:
            continue
        for account,delta in s.raw_contributions(r):
            amounts[account]+=delta
            per_account_rows[account].add(ordinal)
    return amounts,per_account_rows


def _events(events):
    amounts=defaultdict(Decimal)
    for ev in events:
        s.add_event(amounts,ev)
    return amounts


def attribute(rows,events,review):
    """Prove legacy = shadow + excluded-active-rows + unexplained-model-error.

    Ending balance equality is weaker than physical-history identity. An exact
    zero residual proves only the arithmetic of this physical projection.
    """
    row_idx={r['source_row_ordinal']:r for r in rows}
    quarantined_ordinals={r['row'] for r in review}
    if len(row_idx)!=len(rows) or len(quarantined_ordinals)!=len(review):
        raise AuditError('Duplicate legacy/review row ordinal')
    if not quarantined_ordinals.issubset(row_idx):
        raise AuditError('Review refers to source row that does not exist')
    event_ordinal_list=[]
    for e in events:
        for token in str(e['source_row_ordinals']).split(';'):
            try:ordinal=int(token)
            except ValueError:raise AuditError('Invalid source row ordinal in projected event')
            event_ordinal_list.append(ordinal)
    if len(event_ordinal_list)!=len(set(event_ordinal_list)):
        raise AuditError('Projected physical event uses a legacy source row twice')
    if set(event_ordinal_list)&quarantined_ordinals:
        raise AuditError('Projected event contains a quarantined source row')
    if not set(event_ordinal_list).issubset(row_idx):
        raise AuditError('Projected event source row missing from input')
    legacy,_ = _amounts(rows)
    quarantined,touch_rows = _amounts(rows,quarantined_ordinals)
    shadow = _events(events)
    groups=sorted(set(legacy)|set(shadow)|set(quarantined))
    breakdown=[]
    stats=Counter({k:0 for k in (
        'ending_balances_equal','ending_balances_mismatched',
        'mismatches_fully_attributed_to_quarantine','mismatches_not_attributed_to_quarantine',
        'model_identity_failures_all_accounts','accounts_touched_by_active_quarantined_rows',
        'touched_accounts_with_zero_quarantine_net','negative_legacy_accounts',
        'negative_mismatch_accounts','active_review_rows','storno_review_rows',
        'active_review_rows_with_balance_effect','review_rows_without_balance_effect')})
    account_type_mismatch=Counter()
    for k in groups:
        l=legacy[k]; h=shadow[k]; q=quarantined[k]
        d=l-h
        model_gap=d-q
        if d:
            stats['ending_balances_mismatched']+=1
            account_type_mismatch[k[0]]+=1
            if model_gap == 0:stats['mismatches_fully_attributed_to_quarantine']+=1
            else:stats['mismatches_not_attributed_to_quarantine']+=1
        else: stats['ending_balances_equal']+=1
        if model_gap:stats['model_identity_failures_all_accounts']+=1
        if l<0:
            stats['negative_legacy_accounts']+=1
            if d:stats['negative_mismatch_accounts']+=1
        touched=bool(touch_rows.get(k))
        if touched:
            stats['accounts_touched_by_active_quarantined_rows']+=1
            if not q:stats['touched_accounts_with_zero_quarantine_net']+=1
        breakdown.append({
            'account_type':k[0], 'account_id':k[1], 'tip_ambalaze':k[2],
            'legacy_ending_balance':str(l),'shadow_ending_balance':str(h),
            'excluded_review_net':str(q),'model_unexplained_residual':str(model_gap),
            'legacy_minus_shadow':str(d),
            'review_source_rows':';'.join(str(x) for x in sorted(touch_rows.get(k,set()))),
            'review_source_row_count':len(touch_rows.get(k,set())),
            'ending_balance_matches':'YES' if d==0 else 'NO',
            'model_identity_holds':'YES' if model_gap==0 else 'NO',
            'legacy_negative':'YES' if l<0 else 'NO',
        })
    review_details=[]; reasons=Counter(); kinds=Counter()
    for r0 in review:
        r=row_idx[r0['row']]
        legs=s.raw_contributions(r)
        active=is_active(r)
        if active:stats['active_review_rows']+=1
        else:stats['storno_review_rows']+=1
        if legs:stats['active_review_rows_with_balance_effect']+=1
        else:stats['review_rows_without_balance_effect']+=1
        rawtype=f.key(r.get('DokumentTip'))
        docclass=next((named for named in sorted(f.NAMED_TYPES) if f.key(named)==rawtype),'OTHER_OR_UNKNOWN')
        kinds[docclass]+=1
        flags=trim(r.get('forensic_flags'))
        for flag in flags.split(';') if flags else [r0['reason']]:
            reasons[flag]+=1
        review_details.append({
            'source_row_ordinal':r0['row'],'account_type':r.get('EntitetTip',''),
            'account_id':r.get('EntitetID',''), 'tip_ambalaze':r.get('TipAmbalaze',''),
            'smer':r.get('Smer',''), 'vozac_id':r.get('VozacID',''),
            'dokument_tip':r.get('DokumentTip',''), 'dokument_id':r.get('DokumentID',''),
            'amb_id':r.get('AmbID',''), 'kolicina':r.get('Kolicina',''),
            'stornirano':r.get('Stornirano',''),
            'forensic_flags':flags,'quarantine_reason':r0['reason'],
            'legacy_account_effect':';'.join(
                '{}:{}:{}:{}'.format(*acc,val) for acc,val in legs),
        })
    stats['account_type_combinations']=len(groups)
    stats['quarantined_review_rows']=len(review)
    stats['projected_event_count']=len(events)
    stats['projected_source_row_count']=len(event_ordinal_list)
    stats['historical_source_rows_not_projected']=(len(rows)-len(event_ordinal_list)-len(review))
    return breakdown,review_details,dict(stats),dict(reasons),dict(kinds),dict(account_type_mismatch)


def build_summary(stage,shadow_dir):
    baseline,rows=f.analyze(stage)
    shadow_report,events,review,old_accounts=s.build_report(rows,baseline)
    prev_file=shadow_dir/'agrix_amb10_shadow_summary.json'
    if not prev_file.is_file():raise AuditError('MIG-04.2 summary missing')
    prev=json.loads(prev_file.read_text(encoding='utf-8'))
    if prev.get('format')!='MIG-04.2' or prev.get('source_sha256')!=f.SOURCE_SHA:
        raise AuditError('Wrong MIG-04.2 version or source SHA')
    for field in ('schema_sha256','legacy_rows','legacy_state_counts','mig04_1_evidence',
                  'shadow_classification','shadow_active_events_total','account_balance_comparison',
                  'negative_legacy_by_account_type'):
        if prev.get(field)!=shadow_report.get(field):
            raise AuditError('MIG-04.2 baseline mismatch in '+field)
    accounts,review_rows,counts,reasons,kinds,acc_type=attribute(rows,events,review)
    prior=prev['account_balance_comparison']
    if counts['ending_balances_equal']!=prior['account_type_balances_equal'] or \
       counts['ending_balances_mismatched']!=prior['account_type_balance_mismatches'] or \
       counts['account_type_combinations']!=prior['account_type_combinations']:
        raise AuditError('Residual counts disagree with MIG-04.2 baseline')
    if counts['negative_legacy_accounts']!=prior['negative_legacy_account_types']:
        raise AuditError('Negative account count disagrees with MIG-04.2')
    if counts['quarantined_review_rows']!=prev['shadow_classification']['review_rows']:
        raise AuditError('Quarantined row count disagrees with MIG-04.2')
    if counts['active_review_rows'] + counts['storno_review_rows'] != len(review):
        raise AuditError('Review state partition failed')
    if counts['projected_source_row_count']+counts['quarantined_review_rows']+counts['historical_source_rows_not_projected']!=len(rows):
        raise AuditError('Source row partition failed')
    if counts['historical_source_rows_not_projected'] != (
          2*prev['shadow_classification']['historical_storno_pair_groups'] +
          prev['shadow_classification']['historical_storno_single_legs']):
        raise AuditError('Historical classification partition failed')
    status=('PHYSICAL_RESIDUAL_EXPLAINED_BY_QUARANTINE_NOT_AMB10_APPROVED'
            if counts['model_identity_failures_all_accounts']==0
            else 'PHYSICAL_PROJECTION_MODEL_INCONSISTENCY_REVIEW_REQUIRED')
    report={
        'format':VERSION,'status':status,'source_sha256':f.SOURCE_SHA,
        'schema_sha256':baseline['schema_sha256'],
        'source_legacy_rows':len(rows),
        'model_checks':counts,
        'quarantine_forensic_flag_counts_overlapping':dict(sorted(reasons.items())),
        'quarantine_document_type_counts':dict(sorted(kinds.items())),
        'mismatching_account_type_counts':dict(sorted(acc_type.items())),
        'notes':[
            'EXACT_IDENTITY: legacy ending balance = projected shadow + excluded active legacy rows + residual model error.',
            'If model error is zero, ending balance discrepancies are explained ARITHMETICALLY by quarantined legacy legs, not resolved for AMB-10.',
            'No ownership, movement kind, opening inventory, nonnegative event ordering, or third-party packaging obligation proven.',
            'Historical stornirano rows remain in old source; NOT converted to AMB-10 StornoOd inverse.',
            'No legacy balances altered, no stock created, no target XLSM written.',
            'Private account IDs, quantities, documents and per-row details remain on operator PC.',
        ]}
    return report,accounts,review_rows


def txt_summary(report):
    out=['AGRIX MIG-04.3 -- SHADOW BALANCE RESIDUAL ATTRIBUTION',
         'READ ONLY / NOT CANONICAL AMB-10 OR IMPORT AUTHORIZATION',
         'Source SHA256: '+report['source_sha256'],
         'Status: '+report['status'],
         '', 'MODEL CHECKS:']
    out.extend('  {}: {}'.format(k,v) for k,v in sorted(report['model_checks'].items()))
    out.append('REVIEW FORENSIC FLAGS (OVERLAP):')
    out.extend('  {}: {}'.format(k,v) for k,v in report['quarantine_forensic_flag_counts_overlapping'].items())
    out.append('REVIEW DOCUMENT TYPES (COUNTS ONLY):')
    out.extend('  {}: {}'.format(k,v) for k,v in report['quarantine_document_type_counts'].items())
    out.append('MODEL NOTE: exact arithmetic accounting is NOT AMB-10 conversion approval.')
    return '\n'.join(out)+'\n'


def cli(argv=None):
    ap=argparse.ArgumentParser(description='Read-only account residual explanation for physical shadow')
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--shadow',type=Path,required=True)
    ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args(argv)
    stage=args.stage.resolve();shadow=args.shadow.resolve();dest=args.out.resolve()
    if not stage.is_dir() or not shadow.is_dir():
        raise AuditError('Stage or MIG-04.2 shadow input directory missing')
    if stage==shadow or stage in shadow.parents or shadow in stage.parents:
        raise AuditError('Input directories overlap')
    for p in (stage,shadow):
        if dest==p or p in dest.parents or dest in p.parents:
            raise AuditError('Input/output directory overlap')
    if dest.exists():raise AuditError('Output folder exists; choose a fresh path')
    summary,accounts,review=build_summary(stage,shadow)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        (dest/'agrix_amb10_residual_summary.json').write_text(
            json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (dest/'agrix_amb10_residual_summary.txt').write_text(txt_summary(summary),encoding='utf-8')
        f.private_csv(dest/'amb10_residual_accounts_PRIVATE.csv',accounts,[
            'account_type','account_id','tip_ambalaze','legacy_ending_balance',
            'shadow_ending_balance','legacy_minus_shadow','excluded_review_net',
            'model_unexplained_residual','review_source_rows','review_source_row_count',
            'ending_balance_matches','model_identity_holds','legacy_negative'])
        f.private_csv(dest/'amb10_residual_review_rows_PRIVATE.csv',review,[
            'source_row_ordinal','account_type','account_id','tip_ambalaze','smer',
            'vozac_id','dokument_tip','dokument_id','amb_id','kolicina',
            'stornirano','forensic_flags','quarantine_reason','legacy_account_effect'])
    except BaseException:
        for p in dest.iterdir():
            if p.is_file():p.unlink()
        dest.rmdir()
        raise
    print(txt_summary(summary))
    print('Share only the two summary files; both PRIVATE.csv files stay local.')
    return 0


if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as e:
        print('MIG-04.3 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
