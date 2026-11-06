#!/usr/bin/env python3
"""MIG-05.2: read-only legacy Revers physical rescue by strict label/scope proof.

This is deliberately NOT an opaque ReversID reconstruction. Staged records only
have business-number labels; uniqueness is proven only over observed old records.
The result never authorizes legal ownership, AMB-10 events, stock or storno.
"""
from __future__ import annotations
import argparse
from collections import defaultdict, Counter
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as s
from legacy_revers_provenance_v2 import classify, VERSION as PROVENANCE_VERSION

VERSION='MIG-05.2'
OM_KOOP={
  'om-izlaz-koop':frozenset({('kooperant','ulaz'),('stanica','izlaz')}),
  'om-ulaz-koop':frozenset({('kooperant','izlaz'),('stanica','ulaz')}),
}
FIRMA={'om-izlaz-firma':'izlaz','om-ulaz-firma':'ulaz'}
ALLOWED_FLAGS={
  'om-izlaz-koop':{'ORPHAN_DOCUMENT_ID','UNRESOLVED_DOCUMENT_ID'},
  'om-ulaz-koop':{'DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING'},
  'om-izlaz-firma':{'DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING'},
  'om-ulaz-firma':{'DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING'},
}
COLUMNS=['source_row_ordinals','document_type','document_label','event_shape',
         'from_type','from_id','to_type','to_id','tip_ambalaze','quantity',
         'stanica_id','koop_id','vozac_id','status','reasons','ownership','movement_kind',
         'opaque_revers_id','import_authorized']


def _flags(r):
    return set(filter(None,(r.get('forensic_flags') or '').split(';')))


def _row_clean(r, masters):
    typ=f.key(r.get('DokumentTip'))
    reasons=[]
    q=f.dec(r.get('Kolicina'))
    if q is None or q<=0: reasons.append('INVALID_QUANTITY')
    if not is_active(r): reasons.append('NONACTIVE_SOURCE')
    if not trim(r.get('DokumentID')): reasons.append('MISSING_DOCUMENT_LABEL')
    if not trim(r.get('Datum')): reasons.append('MISSING_DATE')
    if f.key(r.get('Smer')) not in ('izlaz','ulaz'): reasons.append('INVALID_DIRECTION')
    if not trim(r.get('TipAmbalaze')): reasons.append('MISSING_AMB_TYPE')
    t=f.key(r.get('EntitetTip'))
    tab={'kooperant':'tblKooperanti','stanica':'tblStanice'}.get(t)
    if not tab or trim(r.get('EntitetID')) not in masters[tab]:
        reasons.append('UNKNOWN_ACCOUNT')
    if typ in OM_KOOP and trim(r.get('VozacID')):
        reasons.append('UNEXPECTED_DRIVER_ON_KOOP_LEG')
    if typ in FIRMA:
        driver=trim(r.get('VozacID'))
        if not driver or driver not in masters['tblVozaci']:
            reasons.append('DRIVER_NOT_PROVEN')
    if _flags(r)-ALLOWED_FLAGS.get(typ,set()):
        reasons.append('ADDITIONAL_FORENSIC_FLAG')
    return reasons


def _event_dict(rr, from_row, to_row,shape):
    a=rr[0]
    return dict(source_row_ordinals=';'.join(str(z['source_row_ordinal']) for z in sorted(rr,key=lambda x:x['source_row_ordinal'])),
         document_type=a['DokumentTip'], document_id=a['DokumentID'],
         from_type=from_row[0],from_id=from_row[1],to_type=to_row[0],to_id=to_row[1],
         tip_ambalaze=a['TipAmbalaze'],quantity=a['Kolicina'],
         directional_evidence=shape,ownership='UNPROVEN',movement_type='UNASSIGNED',storno_od='NOT_ASSIGNED')


def _effects_of_source(rr):
    x=defaultdict(Decimal)
    for r in rr:
        for k,v in s.raw_contributions(r):x[k]+=v
    return {k:v for k,v in x.items() if v}


def _effects_of_event(ev):
    x=defaultdict(Decimal);s.add_event(x,ev)
    return {k:v for k,v in x.items() if v}


def physical_rescue(om,masters):
    """Conservative: complete pair or station/driver leg, never bare label as PK."""
    by_sig=defaultdict(list)
    for r in om:
        sig=(f.key(r.get('DokumentTip')),f.key(r.get('DokumentID')),
             f.key(r.get('Datum')),f.key(r.get('TipAmbalaze')),
             str(f.dec(r.get('Kolicina'))))
        by_sig[sig].append(r)
    # Label is only unique *scoped* to station + date. Multi-event group may be
    # legitimate, but without historical Revers PK cannot prove identity.
    all_results=[]
    for sig,legs in sorted(by_sig.items()):
        typ=sig[0];reasons=[]; event=None
        for r in legs: reasons+=_row_clean(r,masters)
        if typ in OM_KOOP:
            if len(legs)!=2:reasons.append('REVERSE_PAIR_CARDINALITY')
            else:
                roles=frozenset((f.key(r.get('EntitetTip')),f.key(r.get('Smer'))) for r in legs)
                if roles!=OM_KOOP[typ]:reasons.append('REVERSE_PAIR_ROLES')
                if not reasons:
                    koop=next(r for r in legs if f.key(r['EntitetTip'])=='kooperant')
                    sta=next(r for r in legs if f.key(r['EntitetTip'])=='stanica')
                    if typ=='om-izlaz-koop':fr=('Stanica',sta['EntitetID']);to=('Kooperant',koop['EntitetID'])
                    else:fr=('Kooperant',koop['EntitetID']);to=('Stanica',sta['EntitetID'])
                    event=_event_dict(legs,fr,to,'EXACT_OLD_REVERS_KOOP_LEGS_NO_OPAQUE_ID')
        elif typ in FIRMA:
            if len(legs)!=1:reasons.append('FIRMA_SINGLE_CARDINALITY')
            else:
                a=legs[0]
                if f.key(a.get('EntitetTip'))!='stanica' or f.key(a.get('Smer'))!=FIRMA[typ]:
                    reasons.append('FIRMA_LEG_ROLE')
                if not reasons:
                    station=('Stanica',a['EntitetID']);driver=('Vozac',a['VozacID'])
                    fr,to=(station,driver) if FIRMA[typ]=='izlaz' else (driver,station)
                    event=_event_dict(legs,fr,to,'EXACT_OLD_REVERS_FIRMA_DRIVER_INVERSE_NO_OPAQUE_ID')
        else:reasons.append('NOT_A_REVERS_CATEGORY')
        if event is not None and _effects_of_source(legs)!=_effects_of_event(event):
            reasons.append('ACCOUNTING_EFFECT_DIFFERS');event=None
        station=next((trim(r.get('EntitetID')) for r in legs if f.key(r.get('EntitetTip'))=='stanica'),'')
        all_results.append((legs,event,sorted(set(reasons)),station))
    by_scope=defaultdict(list)
    for pos,(legs,event,reasons,station) in enumerate(all_results):
        a=legs[0];scope=(f.key(station),f.key(a.get('Datum')),f.key(a.get('DokumentID')))
        if all(scope): by_scope[scope].append(pos)
    collisions=set()
    for indices in by_scope.values():
        if len(indices)>1:collisions.update(indices)
    out=[];candidate=[]
    for i,(legs,event,reasons,station) in enumerate(all_results):
        if i in collisions:reasons=sorted(set(reasons)|{'MULTIPLE_TRANSFERS_SAME_LABEL_STATION_DAY'});event=None
        a=legs[0];ok=(event is not None and not reasons)
        r0=next((r for r in legs if f.key(r.get('EntitetTip'))=='kooperant'),{})
        result=dict(source_row_ordinals=';'.join(str(r['source_row_ordinal']) for r in sorted(legs,key=lambda z:z['source_row_ordinal'])),
           document_type=a.get('DokumentTip',''), document_label=a.get('DokumentID',''),
           event_shape='PAIR_KOOP' if f.key(a.get('DokumentTip')) in OM_KOOP else 'SINGLE_FIRMA',
           from_type=event['from_type'] if event else '',from_id=event['from_id'] if event else '',
           to_type=event['to_type'] if event else '',to_id=event['to_id'] if event else '',
           tip_ambalaze=a.get('TipAmbalaze',''),quantity=a.get('Kolicina',''),
           stanica_id=station,koop_id=r0.get('EntitetID',''),
           vozac_id=next((r.get('VozacID','') for r in legs if trim(r.get('VozacID'))),''),
           status='PHYSICAL_EVIDENCE_ONLY' if ok else 'HOLD_STRUCTURAL',
           reasons=';'.join(reasons),ownership='UNPROVEN',movement_kind='UNASSIGNED',
           opaque_revers_id='NONE_IN_LEGACY_SOURCE',import_authorized='NO')
        out.append(result)
        if ok:candidate.append(event)
    consumed=sum(len(z[0]) for z in all_results)
    if consumed!=len(om):raise AuditError('Revers row conservation failed')
    return out,candidate


def build(stage,quarantine,provenance):
    manifest,tables,idx=f.inputs(stage)
    orig=json.loads((quarantine/'agrix_amb10_quarantine_summary.json').read_text(encoding='utf-8'))
    prov=json.loads((provenance/'agrix_revers_provenance_summary.json').read_text(encoding='utf-8'))
    for report,version in [(orig,'MIG-04.4'),(prov,PROVENANCE_VERSION)]:
        if report.get('format')!=version or report.get('source_sha256')!=f.SOURCE_SHA or report.get('schema_sha256')!=manifest['schema_sha256']:
            raise AuditError('Input summary version, schema, or source hash mismatch')
    if orig.get('quarantined_active_rows')!=98 or prov.get('om_quarantined_source_rows')!=91 or prov.get('om_rows_with_revers_id')!=0 or prov.get('all_om_provenance')!={'BUSINESS_LABEL_ONLY_NO_OPAQUE_ID':91,'OTKUP_LINK_EXACT':2374}:
        raise AuditError('MIG-05.1 provenance baseline drift')
    prior,all_rows=f.analyze(stage)
    om=[]
    exact_otkup_count=0
    for r in all_rows:
        typ=f.key(r.get('DokumentTip'))
        if typ not in (set(OM_KOOP)|set(FIRMA)):continue
        state,rev,warnings=classify(r,idx['SOURCE__tblOtkup'])
        if state=='BUSINESS_LABEL_ONLY_NO_OPAQUE_ID':
            if not r.get('forensic_flags'):raise AuditError('Revers business-label row not in quarantine')
            om.append(r)
        elif state=='OTKUP_LINK_EXACT':exact_otkup_count+=1
        else:raise AuditError('Unanticipated OM source provenance')
    if len(om)!=91 or exact_otkup_count!=2374:
        raise AuditError('OM source provenance count changed: expected 91 Revers and 2374 Otkup legs')
    records,events=physical_rescue(om,idx)
    original_report,original_events,review,accounts=s.build_report(all_rows,prior)
    if len(original_events)!=4621 or len(review)!=98:
        raise AuditError('MIG-04.2 baseline projection has drifted')
    rescued_ord={int(v) for e in events for v in e['source_row_ordinals'].split(';')}
    if any(v not in {r['source_row_ordinal'] for r in om} for v in rescued_ord):
        raise AuditError('Physical rescue included non-Revers row')
    if len(rescued_ord)!=sum(len(e['source_row_ordinals'].split(';')) for e in events):
        raise AuditError('Duplicate row in physical rescue events')
    original_counts=original_report['account_balance_comparison']
    new_accounts,new_counts,neg=s.reconcile(all_rows,original_events+events)
    reason_counts=Counter(code for z in records for code in filter(None,z['reasons'].split(';')))
    shape_counts=Counter(z['event_shape'] for z in records if z['status']=='PHYSICAL_EVIDENCE_ONLY')
    hold_rows=sum(len(z['source_row_ordinals'].split(';')) for z in records if z['status']!='PHYSICAL_EVIDENCE_ONLY')
    safe={
       'format':VERSION,'status':'REVERS_PHYSICAL_SCOPE_EVIDENCE_ONLY_NO_AMB10_AUTHORIZATION',
       'source_sha256':f.SOURCE_SHA,'schema_sha256':manifest['schema_sha256'],
       'legacy_revers_rows_examined':len(om),
       'legacy_revers_physical_cases':len(records),
       'physical_candidate_events_by_shape':dict(sorted(shape_counts.items())),
       'physical_candidate_events_total':len(events),
       'physical_candidate_source_rows':len(rescued_ord),
       'revers_rows_still_on_hold':hold_rows,
       'revers_hold_reason_counts':dict(sorted(reason_counts.items())),
       'old_shadow_event_count':len(original_events),
       'expanded_shadow_event_count':len(original_events)+len(events),
       'old_shadow_balance_mismatches':original_counts['account_type_balance_mismatches'],
       'expanded_shadow_balance_mismatches':new_counts['account_type_balance_mismatches'],
       'legacy_negative_account_types_unchanged':new_counts['negative_legacy_account_types'],
       'still_quarantined_rows_total':98-len(rescued_ord),
       'non_revers_quarantine_rows':7,
       'amb_00700_stays_on_hold':700 not in rescued_ord,
       'historical_storno_unconverted':335,
       'ready_for_import':False,
       'notes':['Business label + date + station is only scoped physical evidence, not an opaque historical ReversID.',
                'Kooperant two-leg postings and Firma station/driver postings must reconcile exactly by account.',
                'Multiple transfer signatures sharing the same station/day/business-number are held for identity review.',
                'No business kind or packaging ownership decided; no AMB-10 document, StornoOd or XLSM produced.',
                'Private IDs, quantities and documents stay in local CSV only.']}
    return safe,records


def text_report(x):
    show=['source_sha256','legacy_revers_rows_examined','legacy_revers_physical_cases',
          'physical_candidate_events_by_shape','physical_candidate_source_rows',
          'revers_rows_still_on_hold','revers_hold_reason_counts',
          'old_shadow_event_count','expanded_shadow_event_count',
          'old_shadow_balance_mismatches','expanded_shadow_balance_mismatches',
          'legacy_negative_account_types_unchanged','still_quarantined_rows_total',
          'amb_00700_stays_on_hold']
    return 'AGRIX MIG-05.2 - LEGACY REVERS LABEL/SCOPE PHYSICAL RESCUE / READ ONLY\n'+''.join(
         f'{k}: {json.dumps(x[k],ensure_ascii=False,sort_keys=True)}\n' for k in show)+\
         'STOP: physical evidence ONLY, no ReversID recovered, no AMB-10 or XLSM written.\n'


def cli(argv=None):
    ap=argparse.ArgumentParser(description='Verify old Revers two-leg and Firma single-leg physical evidence')
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--quarantine',type=Path,required=True)
    ap.add_argument('--provenance',type=Path,required=True)
    ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args(argv)
    inp=[args.stage.resolve(),args.quarantine.resolve(),args.provenance.resolve()]
    dest=args.out.resolve()
    if not all(p.is_dir() for p in inp):raise AuditError('Input directory missing')
    if dest.exists() or any(dest==d or dest in d.parents or d in dest.parents for d in inp):
        raise AuditError('Output already exists or overlaps input; choose NEW output dir')
    report,private=build(*inp)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        (dest/'agrix_revers_scope_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (dest/'agrix_revers_scope_summary.txt').write_text(text_report(report),encoding='utf-8')
        f.private_csv(dest/'revers_scope_evidence_PRIVATE.csv',private,COLUMNS)
    except BaseException:
        for file in dest.iterdir():
            if file.is_file():file.unlink()
        dest.rmdir()
        raise
    print(text_report(report));return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as ex:
        print('MIG-05.2 STOP:',ex,file=sys.stderr);raise SystemExit(2)
