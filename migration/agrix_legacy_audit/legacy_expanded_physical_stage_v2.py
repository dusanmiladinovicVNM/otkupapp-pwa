#!/usr/bin/env python3
"""MIG-05.3 / local extended physical evidence staging; NEVER AMB-10 canonical.

Reconstruct original MIG-05 shadow from the pinned source. Recalculate all old
Revers candidates (do not trust CSV inputs), verify MIG-05.2 and MIG-05.0 reports,
prove the full 7,815-row partition and the account-by-account exact residual.
Local private JSONL/CSV outputs contain customer information and MUST NOT be shared.
"""
from __future__ import annotations

import argparse
import hashlib
from collections import Counter
from decimal import Decimal
import json
import os
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as shadow
import legacy_amb10_residual_v2 as residual
import legacy_amb10_semantic_preview_v2 as semantic
import legacy_revers_label_scope_v2 as revers
from legacy_revers_provenance_v2 import classify

VERSION='MIG-05.3'
EVENT_FIELDS=(
 'physical_case_id','source_row_ordinals','source_amb_ids','legacy_DokumentTIP',
 'legacy_DokumentID','legacy_Datum','TipAmbalaze','Kolicina',
 'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','physical_evidence',
 'event_provenance','VrstaKretanja_PROPOSAL','evidence_status','reason',
 'ownership','opening_balance','historical_storno','canonical_AMB10_import_authorized',
)
HOLD_FIELDS=(
 'source_row_ordinal','amb_id','date','legacy_DokumentTIP','legacy_DokumentID',
 'EntitetTip','EntitetID','VozacID','Smer','TipAmbalaze','Kolicina',
 'forensic_flags','legacy_stornirano','reconciliation_state',
)
ACCOUNT_FIELDS=(
 'account_type','account_id','tip_ambalaze','legacy_ending_balance',
 'expanded_physical_balance','excluded_7_source_net','unexplained_residual',
 'excluded_source_row_ordinals','legacy_negative','ending_balances_equal',
)


def _summary(path:Path, filename:str, expected_version:str, schema:str):
    target=path/filename
    if not target.is_file():raise AuditError('Missing summary '+str(target))
    obj=json.loads(target.read_text(encoding='utf-8'))
    if (obj.get('format')!=expected_version or obj.get('source_sha256')!=f.SOURCE_SHA or
            obj.get('schema_sha256')!=schema):
        raise AuditError('Input summary source/schema/format mismatch: '+str(target))
    return obj


def _validate_report(actual:dict, recorded:dict, name:str):
    """Count/identity facts must be rederived from stage rather than trusted."""
    for k,value in actual.items():
        if k in ('notes',):continue
        if recorded.get(k)!=value:raise AuditError(name+' summary drift: '+k)


def _only_source_events(om_rows, indexes):
    results,events=revers.physical_rescue(om_rows,indexes)
    if any(r['status']!='PHYSICAL_EVIDENCE_ONLY' for r in results):
        raise AuditError('Old Revers physical proof has new HOLD; stop')
    return events


def _required_baselines(baseline, old_report, new_report, semantic_report):
    if baseline['legacy_state_counts']!={'ACTIVE':7480,'STORNO':335}:
        raise AuditError('Unexpected legacy source states; migration must stop')
    if (old_report['shadow_active_events_total']!=4621 or
            old_report['account_balance_comparison']['account_type_balance_mismatches']<0):
        raise AuditError('Baseline shadow drift')
    if (new_report['legacy_revers_rows_examined']!=91 or
            new_report['physical_candidate_events_total']!=48 or
            new_report['physical_candidate_source_rows']!=91 or
            new_report['revers_rows_still_on_hold']!=0 or
            new_report['expanded_shadow_event_count']!=4669):
        raise AuditError('MIG-05.2 reported baseline drift')
    if (semantic_report['physical_events_examined']!=4621 or
            semantic_report.get('integrity_hold_reasons') or
            semantic_report['draft_evidence_status_counts'].get('HOLD_INTEGRITY',0)!=0 or
            sum(semantic_report['draft_evidence_status_counts'].values())!=4621):
        raise AuditError('MIG-05.0 reported semantic baseline drift')


def _mtype(r):
    typ=f.key(r.get('DokumentTip'))
    if typ=='otpremnica':return 'Otpremnica'
    if typ=='kupci-otpremnica':return 'Kupci-Otpremnica'
    # Never expose unknown document identifiers in public summaries.
    return 'OTHER_OR_UNKNOWN'


def build(stage:Path,quarantine:Path,provenance:Path,revers_scope:Path,semantics:Path):
    mf,_,idx=f.inputs(stage)
    forensic_report,source=f.analyze(stage)
    source_index={r['source_row_ordinal']:r for r in source}
    if len(source_index)!=len(source):raise AuditError('Duplicate source row ordinal')
    original_report,base_events,review,_=shadow.build_report(source,forensic_report)

    _summary(quarantine,'agrix_amb10_quarantine_summary.json','MIG-04.4',mf['schema_sha256'])
    prov=_summary(provenance,'agrix_revers_provenance_summary.json','MIG-05.1',mf['schema_sha256'])
    scoped=_summary(revers_scope,'agrix_revers_scope_summary.json','MIG-05.2',mf['schema_sha256'])
    sem_ref=_summary(semantics,'agrix_amb10_semantic_preview_summary.json','MIG-05.0',mf['schema_sha256'])
    _required_baselines(forensic_report,original_report,scoped,sem_ref)

    om=[];exact=0
    for r in source:
        if f.key(r.get('DokumentTip')) not in (set(revers.OM_KOOP)|set(revers.FIRMA)):continue
        state,*_=classify(r,idx['SOURCE__tblOtkup'])
        if state=='BUSINESS_LABEL_ONLY_NO_OPAQUE_ID':
            if not r.get('forensic_flags'):raise AuditError('Unflagged Revers row unexpectedly rescanned')
            om.append(r)
        elif state=='OTKUP_LINK_EXACT':exact+=1
        else:raise AuditError('Unknown old OM provenance')
    if (len(om),exact)!=(91,2374) or prov['om_quarantined_source_rows']!=91:
        raise AuditError('OM classification changed')
    events_new=_only_source_events(om,idx)
    if len(events_new)!=48:raise AuditError('Revers case count drift')
    new_report,_=revers.build(stage,quarantine,provenance)
    _validate_report(new_report,scoped,'MIG-05.2')

    events=base_events+events_new
    old_ids=set();mapped=[]
    source_rows=source
    existing_drafts=[semantic.classify_event(ev,source_rows,idx) for ev in base_events]
    status_counts=Counter();kind_counts=Counter();origin_counts=Counter()
    for i,(event,proposal) in enumerate(
            [(ev,existing_drafts[i]) for i,ev in enumerate(base_events)] +
            [(ev,None) for ev in events_new],1):
        src=[int(n) for n in event['source_row_ordinals'].split(';')]
        if len(src) not in (1,2) or len(set(src))!=len(src) or any(n in old_ids for n in src):
            raise AuditError('An active source leg occurs in two physical cases')
        old_ids.update(src)
        legs=[source_index[n] for n in src]
        if any(not is_active(r) for r in legs):raise AuditError('Storno row projected as active')
        if proposal is not None:
            # Retain previously independently verified semantic proposals.
            if proposal['evidence_status'] not in ('DIRECTION_KIND_PROPOSED_ONLY','HOLD_OWNERSHIP'):
                raise AuditError('Unexpected older semantic status')
            status=proposal['evidence_status']
            kind=proposal['VrstaKretanja_PROPOSAL']
            reason=proposal['reason']
            provenance_type='MIG05_ORIGINAL_PHYSICAL'
        else:
            status='HOLD_OWNERSHIP';kind='UNASSIGNED'
            reason='LEGACY_REVERS_BUSINESS_KIND_AND_LEGAL_TITLE_UNPROVEN'
            provenance_type='MIG05_2_RECOVERED_REVERS_PHYSICAL'
        status_counts[status]+=1;kind_counts[kind]+=1;origin_counts[provenance_type]+=1
        mapped.append({
            'physical_case_id':f'PHY-{i:05d}',
            'source_row_ordinals':event['source_row_ordinals'],
            'source_amb_ids':';'.join(trim(r['AmbID']) for r in legs),
            'legacy_DokumentTIP':event['document_type'],
            'legacy_DokumentID':event['document_id'],
            'legacy_Datum':trim(legs[0]['Datum']),
            'TipAmbalaze':event['tip_ambalaze'],'Kolicina':event['quantity'],
            'OdNalogTip':event['from_type'],'OdNalogID':event['from_id'],
            'NaNalogTip':event['to_type'],'NaNalogID':event['to_id'],
            'physical_evidence':event['directional_evidence'],
            'event_provenance':provenance_type,
            'VrstaKretanja_PROPOSAL':kind,'evidence_status':status,
            'reason':reason,'ownership':'UNPROVEN','opening_balance':'UNPROVEN',
            'historical_storno':'NOT_REPLAYED','canonical_AMB10_import_authorized':'NO'
        })

    active={r['source_row_ordinal'] for r in source if is_active(r)}
    storno={r['source_row_ordinal'] for r in source if not is_active(r)}
    remaining=active-old_ids
    if (len(old_ids),len(remaining),len(storno))!=(7473,7,335):
        raise AuditError('Full row partition incorrect')
    if old_ids&storno or old_ids&remaining or storno&remaining or len(old_ids|remaining|storno)!=len(source):
        raise AuditError('Source partition contains duplicates or omissions')
    if 700 not in remaining:raise AuditError('Q-00700 must remain isolated')
    if remaining!={r['row'] for r in review}-{n for e in events_new for n in (int(v) for v in e['source_row_ordinals'].split(';'))}:
        raise AuditError('Remaining rows not explained by prior quarantine')

    # Exact independent account equation for every account/type:
    # legacy balance = projected physical balance + excluded active source impact.
    accounts,old_review,counts,_,_,_=residual.attribute(source,events,
        [{'row':n,'reason':'REMAINS_UNMAPPED'} for n in sorted(remaining)])
    if (counts['model_identity_failures_all_accounts']!=0 or
            counts['ending_balances_mismatched']!=scoped['expanded_shadow_balance_mismatches']):
        raise AuditError('Expanded shadow physical account reconciliation failure')
    if (counts['negative_legacy_accounts']!=scoped['legacy_negative_account_types_unchanged'] or
            counts['historical_source_rows_not_projected']!=335):
        raise AuditError('Negative-account or historical-storno baseline drift')
    if counts['projected_source_row_count']!=7473 or counts['quarantined_review_rows']!=7:
        raise AuditError('Residual reconciliation source partition drift')

    by_public_type=Counter(_mtype(source_index[n]) for n in remaining)
    raw_account_rows=[]
    for x in accounts:
        raw_account_rows.append(dict(
            account_type=x['account_type'],account_id=x['account_id'],
            tip_ambalaze=x['tip_ambalaze'],legacy_ending_balance=x['legacy_ending_balance'],
            expanded_physical_balance=x['shadow_ending_balance'],
            excluded_7_source_net=x['excluded_review_net'],
            unexplained_residual=x['model_unexplained_residual'],
            excluded_source_row_ordinals=x['review_source_rows'],
            legacy_negative=x['legacy_negative'],
            ending_balances_equal=x['ending_balance_matches'],
        ))
    private_holds=[]
    for n in sorted(remaining):
        r=source_index[n]
        private_holds.append(dict(source_row_ordinal=n,amb_id=r.get('AmbID',''),
            date=r.get('Datum',''),legacy_DokumentTIP=r.get('DokumentTip',''),
            legacy_DokumentID=r.get('DokumentID',''),
            EntitetTip=r.get('EntitetTip',''),EntitetID=r.get('EntitetID',''),
            VozacID=r.get('VozacID',''),Smer=r.get('Smer',''),
            TipAmbalaze=r.get('TipAmbalaze',''),Kolicina=r.get('Kolicina',''),
            forensic_flags=r.get('forensic_flags',''),legacy_stornirano=r.get('Stornirano',''),
            reconciliation_state='EXCLUDED_ACTIVE_SOURCE_NOT_YET_PROJECTED'))

    expected_status=Counter(sem_ref['draft_evidence_status_counts'])
    expected_status['HOLD_OWNERSHIP']+=48
    if status_counts!=expected_status:
        raise AuditError('Expanded semantic classifications changed')
    expected_kinds=Counter(sem_ref['proposed_business_movement_kind_counts'])
    expected_kinds['UNASSIGNED']+=48
    if kind_counts!=expected_kinds:
        raise AuditError('Old semantics proposals unexpectedly changed')
    if len(mapped)!=4669 or len(events)!=4669:raise AuditError('Event count mismatch')

    public={
        'format':VERSION,'status':'EXPANDED_PHYSICAL_STAGING_ONLY_NO_AMB10_OR_XLSM',
        'source_sha256':f.SOURCE_SHA,'schema_sha256':mf['schema_sha256'],
        'source_amb_rows':len(source),'active_source_rows':len(active),
        'historical_storno_rows_unconverted':len(storno),
        'base_physical_events':len(base_events),
        'rescued_legacy_revers_physical_events':len(events_new),
        'expanded_physical_events':len(mapped),
        'expanded_physical_evidence_sha256':hashlib.sha256(
            ''.join(json.dumps(v,ensure_ascii=False,sort_keys=True)+'\n' for v in mapped).encode('utf-8')).hexdigest(),
        'physical_source_legs':len(old_ids),'still_excluded_active_source_legs':len(remaining),
        'remaining_hold_document_types':dict(sorted(by_public_type.items())),
        'remaining_hold_row_ordinals':sorted(remaining),
        'prior_balance_mismatch_accounts':original_report['account_balance_comparison']['account_type_balance_mismatches'],
        'expanded_balance_mismatch_accounts':counts['ending_balances_mismatched'],
        'balance_mismatches_explained_by_remaining_holds':counts['mismatches_fully_attributed_to_quarantine'],
        'unexplained_account_model_residual_count':counts['model_identity_failures_all_accounts'],
        'negative_legacy_account_types':counts['negative_legacy_accounts'],
        'evidence_status_counts':dict(sorted(status_counts.items())),
        'business_kind_proposal_counts':dict(sorted(kind_counts.items())),
        'source_classification_counts':dict(sorted(origin_counts.items())),
        'q00700_remains_excluded':700 in remaining,
        'ready_for_import':False,
        'notes':[
            'Only extended PHYSICAL transfers; no canonical AMB-10 AmbID, title, opening, StornoOd or documents.',
            '8 account balance gaps are accounted for ARITHMETICALLY by 7 excluded active legacy rows.',
            '138? no: see local test command for actual test totals.',
            'Proposed business kinds do not prove legal title/ownership or point-in-time stock feasibility.',
            'All private records contain customer data. Never upload or paste into tblAmbalaza.'
        ]}
    return public,mapped,private_holds,raw_account_rows


def render(x):
    keys=['source_sha256','source_amb_rows','active_source_rows','physical_source_legs',
          'historical_storno_rows_unconverted','base_physical_events',
          'rescued_legacy_revers_physical_events','expanded_physical_events',
          'expanded_physical_evidence_sha256',
          'still_excluded_active_source_legs','remaining_hold_document_types',
          'prior_balance_mismatch_accounts','expanded_balance_mismatch_accounts',
          'balance_mismatches_explained_by_remaining_holds',
          'unexplained_account_model_residual_count','negative_legacy_account_types',
          'evidence_status_counts','business_kind_proposal_counts',
          'q00700_remains_excluded','ready_for_import']
    return ('AGRIX MIG-05.3 / LOCAL PHYSICAL STAGE, NOT AMB-10 IMPORT\n'+
            ''.join(k+': '+json.dumps(x[k],ensure_ascii=False,sort_keys=True)+'\n' for k in keys)+
            'NO canonical AMB-10 movements or XLSM files created.\n')


def cli(argv=None):
    parser=argparse.ArgumentParser(description='MIG-05.3 / expanded physical evidence stage; read-only source')
    for name in ('stage','quarantine','provenance','revers-scope','semantic','out'):
        parser.add_argument('--'+name,type=Path,required=True)
    a=parser.parse_args(argv)
    all_ins=[a.stage.resolve(),a.quarantine.resolve(),a.provenance.resolve(),
             a.revers_scope.resolve(),a.semantic.resolve()]
    out=a.out.resolve()
    if not all(x.is_dir() for x in all_ins):raise AuditError('Required input folder missing')
    if out.exists() or any(out==x or out in x.parents or x in out.parents for x in all_ins):
        raise AuditError('Output folder exists or overlaps input; use fresh folder')
    summary,events,holds,accounts=build(*all_ins)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_expanded_physical_summary.json').write_text(
            json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_expanded_physical_summary.txt').write_text(render(summary),encoding='utf-8')
        f.private_csv(out/'expanded_physical_cases_PRIVATE.csv',events,EVENT_FIELDS)
        f.private_csv(out/'seven_remaining_legacy_rows_PRIVATE.csv',holds,HOLD_FIELDS)
        f.private_csv(out/'expanded_balance_reconciliation_PRIVATE.csv',accounts,ACCOUNT_FIELDS)
        # JSONL is machine-readable source-evidence, not canonical AMB-10 JSONL.
        evidence=out/'expanded_physical_cases_PRIVATE.jsonl'
        with evidence.open('w',encoding='utf-8',newline='\n') as fh:
            for ev in events:fh.write(json.dumps(ev,ensure_ascii=False,sort_keys=True)+'\n')
        os.chmod(evidence,0o600)
    except BaseException:
        for child in out.iterdir():
            if child.is_file():child.unlink()
        out.rmdir();raise
    print(render(summary));print('Share ONLY aggregate JSON/TXT; PRIVATE files remain on your PC.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as ex:
        print('MIG-05.3 STOP:',ex,file=sys.stderr);raise SystemExit(2)
