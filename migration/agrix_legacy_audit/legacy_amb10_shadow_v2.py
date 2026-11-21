#!/usr/bin/env python3
"""MIG-04.2: PRIVATE, read-only shadow balance reconciliation of legacy AMB-10 evidence.

Creates NO canonical AMB-10 entries, no opening stock, no ownership interpretation,
no storno inverse, no workbook. The shadow is only a directional *physical* model
of the high-confidence old posting shapes, NOT an import candidate.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_amb10_forensics_v2 import (
    analyze as forensic_analyze, inputs, dec, key, private_csv, SOURCE_SHA,
)

VERSION='MIG-04.2'


def raw_contributions(r):
    """Exactly the legacy *ending balance* rule, never interpret ownership."""
    q=dec(r.get('Kolicina'))
    if not is_active(r) or q is None or q<=0 or key(r.get('Smer')) not in ('ulaz','izlaz'):
        return []
    t=trim(r.get('EntitetTip')); eid=trim(r.get('EntitetID'))
    p=trim(r.get('TipAmbalaze')); driver=trim(r.get('VozacID'))
    sign=1 if key(r['Smer'])=='ulaz' else -1
    out=[]
    if t and eid:out.append(((t,eid,p), q*sign))
    if t in ('Stanica','Kupac') and driver:
        out.append((('Vozac',driver,p),-q*sign))
    return out


def add_event(delta, event):
    q=dec(event['quantity'])
    if q is None or q<=0:
        raise AuditError('Invalid shadow event quantity')
    src=(event['from_type'],event['from_id'],event['tip_ambalaze'])
    dst=(event['to_type'],event['to_id'],event['tip_ambalaze'])
    delta[src]-=q
    delta[dst]+=q


def project_rows(rows):
    """Consume ONLY classification from the established MIG-04.1 forensic implementation.

    Structural proof is conservative: any per-row flag or inconsistent document
    state sends the row/group for review; it is never guessed into a transfer.
    """
    groups=defaultdict(list)
    singles=[]; review=[]; used=set()
    for r in rows:
        typ=key(r['DokumentTip'])
        if typ in ('otkup','om-izlaz-koop'):
            groups[(typ,trim(r['DokumentID']))].append(r)
        else:singles.append(r)
    events=[]; summary=Counter({k:0 for k in ('active_physical_pair_events','active_physical_single_events','historical_storno_pair_groups','historical_storno_single_legs','review_rows')})
    def excluded(r,reason):
        used.add(r['source_row_ordinal'])
        review.append({'row':r['source_row_ordinal'], 'reason':reason,
                       'forensic_flags':r.get('forensic_flags',''),
                       'evidence_class':r.get('evidence_class','')})
        summary['review_rows']+=1
    for (typ,docid),legs in sorted(groups.items()):
        if len(legs)!=2 or any(r.get('evidence_class')!='MIRRORED_PAIR_EVIDENCE_ONLY' or r.get('forensic_flags') for r in legs):
            for r in legs:excluded(r,'PAIR_NOT_CLEAN_OR_NOT_EXACT')
            continue
        a,b=legs
        if is_active(a)!=is_active(b):
            for r in legs:excluded(r,'PAIR_MIXED_STORNO')
            continue
        if not is_active(a):
            summary['historical_storno_pair_groups']+=1
            for r in legs:used.add(r['source_row_ordinal'])
            continue
        if typ=='otkup':
            fst=next(r for r in legs if r['EntitetTip']=='Kooperant')
            snd=next(r for r in legs if r['EntitetTip']=='Stanica')
            kind='OTKUP_PHYSICAL_ROBUST_DIRECTION'
        else:
            fst=next(r for r in legs if r['EntitetTip']=='Stanica')
            snd=next(r for r in legs if r['EntitetTip']=='Kooperant')
            kind='OM_OUT_PHYSICAL_OWNERSHIP_UNKNOWN'
        event=dict(source_row_ordinals=';'.join(str(r['source_row_ordinal']) for r in legs),
                   document_type=a['DokumentTip'],document_id=docid,
                   from_type=fst['EntitetTip'],from_id=fst['EntitetID'],
                   to_type=snd['EntitetTip'],to_id=snd['EntitetID'],
                   tip_ambalaze=a['TipAmbalaze'],quantity=a['Kolicina'],
                   directional_evidence=kind,ownership='UNPROVEN',
                   movement_type='NOT_ASSIGNED',storno_od='NOT_ASSIGNED')
        events.append(event)
        summary['active_physical_pair_events']+=1
        for r in legs:used.add(r['source_row_ordinal'])
    for r in singles:
        klass=r.get('evidence_class')
        if klass!='DRIVER_INVERSE_EVIDENCE_ONLY' or r.get('forensic_flags'):
            excluded(r,'SINGLE_NOT_CLEAN_OR_NOT_PROVEN')
            continue
        if not is_active(r):
            summary['historical_storno_single_legs']+=1
            used.add(r['source_row_ordinal'])
            continue
        t=r['EntitetTip'];d=trim(r['VozacID'])
        if t=='Stanica' and key(r['Smer'])=='izlaz':
            ft,fi,tt,ti='Stanica',r['EntitetID'],'Vozac',d
        elif t=='Kupac' and key(r['Smer'])=='ulaz':
            ft,fi,tt,ti='Vozac',d,'Kupac',r['EntitetID']
        elif t=='Kupac' and key(r['Smer'])=='izlaz':
            ft,fi,tt,ti='Kupac',r['EntitetID'],'Vozac',d
        else:
            excluded(r,'SINGLE_UNEXPECTED_DIRECTION')
            continue
        events.append(dict(source_row_ordinals=str(r['source_row_ordinal']),
             document_type=r['DokumentTip'],document_id=r['DokumentID'],
             from_type=ft,from_id=fi,to_type=tt,to_id=ti,
             tip_ambalaze=r['TipAmbalaze'],quantity=r['Kolicina'],
             directional_evidence='DRIVER_INVERSE_PHYSICAL_ONLY',ownership='UNPROVEN',
             movement_type='NOT_ASSIGNED',storno_od='NOT_ASSIGNED'))
        used.add(r['source_row_ordinal'])
        summary['active_physical_single_events']+=1
    ordinals={r['source_row_ordinal'] for r in rows}
    if used!=ordinals or len(rows)!=len(ordinals):
        raise AuditError('Internal shadow invariant failed: row coverage not exact')
    if len(events)!=summary['active_physical_pair_events']+summary['active_physical_single_events']:
        raise AuditError('Internal shadow invariant failed: event count')
    return events,review,summary


def reconcile(rows,events):
    legacy=defaultdict(Decimal);shadow=defaultdict(Decimal)
    for r in rows:
        for account,delta in raw_contributions(r):legacy[account]+=delta
    for ev in events:add_event(shadow,ev)
    keys=sorted(set(legacy)|set(shadow))
    compare=[];counts=Counter({k:0 for k in ('account_type_balance_mismatches','account_type_balances_equal','negative_legacy_account_types','negative_legacy_with_shadow_mismatch','negative_legacy_with_shadow_match')});negative={'Kooperant':0,'Stanica':0,'Vozac':0,'Kupac':0,'Firma':0,'Other':0}
    for account in keys:
        l=legacy[account];s=shadow[account]
        mismatch=l!=s
        if mismatch:counts['account_type_balance_mismatches']+=1
        else:counts['account_type_balances_equal']+=1
        if l<0:
            counts['negative_legacy_account_types']+=1
            negative[account[0] if account[0] in negative else 'Other']+=1
            if mismatch:counts['negative_legacy_with_shadow_mismatch']+=1
            else:counts['negative_legacy_with_shadow_match']+=1
        compare.append(dict(account_type=account[0],account_id=account[1],
                            tip_ambalaze=account[2],legacy_ending_balance=str(l),
                            shadow_ending_balance=str(s),unexplained_delta=str(l-s),
                            legacy_negative='YES' if l<0 else 'NO',
                            exact_match='YES' if not mismatch else 'NO',
                            ownership='NOT_PROVEN',opening_stock='NOT_PROVEN'))
    counts['account_type_combinations']=len(keys)
    return compare,counts,{k:v for k,v in negative.items() if v}


def build_report(rows,baseline):
    events,review,stats=project_rows(rows)
    compare,account_stats,neg=reconcile(rows,events)
    if len(rows)!=baseline['legacy_ambalaza_rows'] or len(rows)!=7815:
        raise AuditError('Legacy row count does not match pinned MIG-04.1 source')
    expected=baseline['negative_legacy_ending_balance_accounts']['account_type_counts']
    if {k:v for k,v in neg.items() if v}!=expected:
        raise AuditError('Shadow disagrees with MIG-04.1 legacy negative balance baseline')
    if len(review)!=stats['review_rows']:
        raise AuditError('Shadow review row count mismatch')
    if baseline['pair_group_counts'].get('MIRRORED_PAIR_EVIDENCE_ONLY',0)*2 + baseline['pair_group_counts'].get('UNRESOLVED_DOCUMENT_ID',0)*2 + sum(baseline['single_leg_evidence_counts'].values())!=7815:
        raise AuditError('Source evidence group partition does not add to legacy rows')
    summary={
        'format':VERSION,
        'status':'READ_ONLY_PHYSICAL_SHADOW_NO_CANONICAL_AMB10_AND_NO_IMPORT',
        'source_sha256':SOURCE_SHA,
        'schema_sha256':baseline['schema_sha256'],
        'legacy_rows':len(rows),
        'legacy_state_counts':baseline['legacy_state_counts'],
        'mig04_1_evidence':{'mirrored_pair_groups':baseline['pair_group_counts'].get('MIRRORED_PAIR_EVIDENCE_ONLY',0),
                            'driver_inverse_legs':baseline['single_leg_evidence_counts'].get('DRIVER_INVERSE_EVIDENCE_ONLY',0),
                            'structural_review_rows':baseline['rows_with_structural_review']},
        'shadow_classification':dict(stats),
        'shadow_active_events_total':len(events),
        'account_balance_comparison':dict(account_stats),
        'negative_legacy_by_account_type':neg,
        'review_row_ordinals_first_120':[r['row'] for r in review][:120],
        'notes':[
            'Shadow rows are PHYSICAL DIRECTION evidence, NOT accepted AMB-10 events.',
            'Even matching shadow ending balances DO NOT prove ownership, movement kind, or opening stock.',
            'Historical STORNO is preserved in old source archive; no AMB-10 inverse entries were made.',
            'No invented SpoljniSvet, ULAZ_TUDJE, NABAVKA, deficits, or initial balances.',
            'All private amounts, account IDs, document IDs, document dates and numbers remain local.',
            'DO NOT import private event CSV to canonical tblAmbalaza.',
        ]}
    return summary,events,review,compare


def report_txt(r):
    out=['AGRIX MIG-04.2 -- PHYSICAL SHADOW RECONCILIATION / NOT AMB-10 CONVERSION',
         'READ ONLY. Source SHA256: '+r['source_sha256'],
         'Legacy rows: '+str(r['legacy_rows']),
         'Evidence: '+json.dumps(r['mig04_1_evidence'],ensure_ascii=False),
         'Shadow: '+json.dumps(r['shadow_classification'],ensure_ascii=False),
         'Shadow active physical events: '+str(r['shadow_active_events_total']),
         'Legacy vs shadow account/type: '+json.dumps(r['account_balance_comparison'],ensure_ascii=False),
         'Negative legacy by account: '+json.dumps(r['negative_legacy_by_account_type'],ensure_ascii=False),
         'Review ordinals (first 120): '+str(r['review_row_ordinals_first_120']),
         'STOP: No AMB-10 movement, storno inverse, opening stock or XLSM generated.']
    return '\n'.join(out)+'\n'


def cli(argv=None):
    p=argparse.ArgumentParser(description='Read-only physical transfer shadow with legacy balance comparison')
    p.add_argument('--stage',type=Path,required=True)
    p.add_argument('--out',type=Path,required=True)
    args=p.parse_args(argv)
    stage=args.stage.resolve();out=args.out.resolve()
    if not stage.is_dir():raise AuditError('MIG-02 staging directory not found')
    if out==stage or stage in out.parents or out in stage.parents:
        raise AuditError('Input/output overlap')
    if out.exists():raise AuditError('Output folder exists; choose fresh folder, never overwrite')
    # Reuse tested MIG-04.1 reader/classifier, including original-source checksum.
    baseline,raw=forensic_analyze(stage)
    summary,events,review,compare=build_report(raw,baseline)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_amb10_shadow_summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_amb10_shadow_summary.txt').write_text(report_txt(summary),encoding='utf-8')
        event_fields=['source_row_ordinals','document_type','document_id','from_type','from_id','to_type','to_id',
                     'tip_ambalaze','quantity','directional_evidence','ownership','movement_type','storno_od']
        private_csv(out/'amb10_shadow_events_PRIVATE.csv',events,event_fields)
        private_csv(out/'amb10_shadow_accounts_PRIVATE.csv',compare,[
              'account_type','account_id','tip_ambalaze','legacy_ending_balance','shadow_ending_balance',
              'unexplained_delta','legacy_negative','exact_match','ownership','opening_stock'])
        private_csv(out/'amb10_shadow_review_PRIVATE.csv',review,['row','reason','forensic_flags','evidence_class'])
    except BaseException:
        for f in out.iterdir():
            if f.is_file():f.unlink()
        out.rmdir()
        raise
    print(report_txt(summary))
    print('Only share the two summary files; private CSVs must remain local.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as e:
        print('MIG-04.2 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
