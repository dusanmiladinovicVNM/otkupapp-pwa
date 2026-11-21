#!/usr/bin/env python3
"""MIG-05.0: Strict, READ ONLY draft of AMB-10 direction/purpose from legacy v2.

Every physical event is either marked PROPOSED (NOT approved) or HOLD. Do not
import any file produced by this tool. Ownership, opening stock, obligations,
chronological nonnegative guarantees and STORNO inverses remain unproven.
"""
from __future__ import annotations
import argparse
from collections import Counter
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_amb10_forensics_v2 import analyze as forensics, inputs, dec, key, private_csv, SOURCE_SHA
from legacy_amb10_shadow_v2 import build_report

VERSION='MIG-05.0'
KIND_WITH_GOODS='AMBALAZA_UZ_ROBU'
KIND_RETURN='POVRAT_PRAZNE'
OWNERSHIP_HOLD='OWNERSHIP_UNKNOWN_ISSUE_VS_THIRD_PARTY_RETURN'
PRIVATE_COLUMNS=[
    'source_row_ordinals','source_amb_ids','DokumentTIP','DokumentID','Datum','TipAmbalaze','Kolicina',
    'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja_PROPOSAL',
    'evidence_status','reason','title_or_ownership','opening_balance','storno_replay','canonical_AMB10_import_authorized'
]


def same(a,b):
    return bool(trim(a)) and key(a)==key(b)


def classify_event(event, originals, indices):
    """Classify a proven *physical* shadow event, never create a canonical ledger row."""
    ords=[int(t) for t in event['source_row_ordinals'].split(';')]
    if len(ords) not in (1,2) or len(ords)!=len(set(ords)):
        raise AuditError('Invalid event lineage ordinals')
    if any(n<=0 or n>len(originals) for n in ords):
        raise AuditError('Invalid event lineage ordinal range')
    rows=[originals[n-1] for n in ords]
    if not all(is_active(r) for r in rows):raise AuditError('STORNO source accidentally entered active draft')
    dt=key(event['document_type']); docid=trim(event['document_id'])
    category={'otkup':'SOURCE__tblOtkup','om-izlaz-koop':'SOURCE__tblOtkup',
              'otpremnica':'SOURCE__tblOtpremnica','prijemnica':'tblPrijemnica'}
    if dt not in category:raise AuditError('Unknown physical shadow document type')
    parent=indices[category[dt]].get(docid)
    reason=[];proposal='UNASSIGNED'
    if not parent or not is_active(parent[1]):reason.append('PARENT_NOT_ACTIVE_OR_MISSING')
    else:
        p=parent[1]
        expected_type=trim(p.get('TipAmbalaze'))
        if not expected_type or not same(expected_type,event['tip_ambalaze']):
            reason.append('PARENT_AMBALAZA_TYPE_NOT_EXACT')
        if dt=='otkup':
            expected_kind=KIND_WITH_GOODS
            expected_qty='KolAmbalaze'
            fields=(('Kooperant','KooperantID','from_type','from_id'),
                    ('Stanica','StanicaID','to_type','to_id'))
            if len(rows)!=2:reason.append('EXPECTED_TWO_LEGS')
        elif dt=='om-izlaz-koop':
            expected_kind='UNRESOLVED'
            expected_qty='KolAmbIzdata'
            fields=(('Stanica','StanicaID','from_type','from_id'),
                    ('Kooperant','KooperantID','to_type','to_id'))
            if len(rows)!=2:reason.append('EXPECTED_TWO_LEGS')
        elif dt=='otpremnica':
            expected_kind=KIND_WITH_GOODS
            expected_qty='KolAmbalaze'
            fields=(('Stanica','StanicaID','from_type','from_id'),
                    ('Vozac','VozacID','to_type','to_id'))
            if len(rows)!=1:reason.append('EXPECTED_ONE_LEG')
            if not all(same(r.get('VozacID'),p.get('VozacID')) for r in rows):
                reason.append('DRIVER_NOT_EXACT')
        else:
            if event['from_type']=='Vozac' and event['to_type']=='Kupac':
                expected_kind=KIND_WITH_GOODS;expected_qty='KolAmbalaze'
                fields=(('Vozac','VozacID','from_type','from_id'),
                        ('Kupac','KupacID','to_type','to_id'))
            elif event['from_type']=='Kupac' and event['to_type']=='Vozac':
                expected_kind=KIND_RETURN;expected_qty='KolAmbVracena'
                fields=(('Kupac','KupacID','from_type','from_id'),
                        ('Vozac','VozacID','to_type','to_id'))
            else:
                expected_kind='UNRESOLVED';expected_qty='';fields=()
                reason.append('UNKNOWN_RECEIPT_DIRECTION')
            if len(rows)!=1:reason.append('EXPECTED_ONE_LEG')
            if not all(same(r.get('VozacID'),p.get('VozacID')) for r in rows):
                reason.append('DRIVER_NOT_EXACT')
        for role,field,ev_role,ev_id in fields:
            if event[ev_role]!=role or not same(p.get(field),event[ev_id]):
                reason.append('PARENT_'+role.upper()+'_IDENTITY_MISMATCH')
        amount=dec(event['quantity']);expected=dec(p.get(expected_qty)) if expected_qty else None
        if amount is None or amount<=0 or expected is None or expected!=amount:
            reason.append('PARENT_QUANTITY_NOT_EXACT')
        if any(not same(r.get('DokumentID'),docid) or key(r.get('DokumentTip'))!=dt or
               dec(r.get('Kolicina'))!=amount or not same(r.get('TipAmbalaze'),event['tip_ambalaze'])
               for r in rows):
            reason.append('SOURCE_LEG_VALUE_ID_MISMATCH')
        # An exact old leg pair proves physical movement, not ownership/title.
        proposal=expected_kind
    if reason:
        evidence='HOLD_INTEGRITY';proposal='UNASSIGNED'
    elif dt=='om-izlaz-koop':
        evidence='HOLD_OWNERSHIP';proposal='UNASSIGNED'
    else:
        evidence='DIRECTION_KIND_PROPOSED_ONLY'
    return dict(
        source_row_ordinals=event['source_row_ordinals'],
        source_amb_ids=';'.join(trim(r['AmbID']) for r in rows),
        DokumentTIP=event['document_type'],DokumentID=docid,
        Datum=trim(rows[0]['Datum']), TipAmbalaze=event['tip_ambalaze'],
        Kolicina=event['quantity'],
        OdNalogTip=event['from_type'],OdNalogID=event['from_id'],
        NaNalogTip=event['to_type'],NaNalogID=event['to_id'],
        VrstaKretanja_PROPOSAL=proposal,evidence_status=evidence,
        reason=';'.join(sorted(set(reason))) if reason else (
            OWNERSHIP_HOLD if evidence=='HOLD_OWNERSHIP' else 'REQUIRES_OPENING_OWNERSHIP_AND_CHRONOLOGY_GATE'),
        title_or_ownership='NOT_VERIFIED', opening_balance='NOT_VERIFIED',
        storno_replay='NOT_DONE', canonical_AMB10_import_authorized='NO')


def run(stage,shadow_dir,storno_dir):
    base,rows=forensics(stage)
    shadow,events,review,_=build_report(rows,base)
    shadow_ref=json.loads((shadow_dir/'agrix_amb10_shadow_summary.json').read_text(encoding='utf-8'))
    storno_ref=json.loads((storno_dir/'agrix_storno_lineage_summary.json').read_text(encoding='utf-8'))
    for name,ref in [('shadow',shadow_ref),('storno',storno_ref)]:
        if ref.get('source_sha256')!=SOURCE_SHA or ref.get('schema_sha256')!=base['schema_sha256']:
            raise AuditError('Prior '+name+' source/schema mismatch')
    if shadow_ref.get('format')!='MIG-04.2' or storno_ref.get('format')!='MIG-04.6':
        raise AuditError('Wrong prior report version')
    if shadow_ref.get('shadow_active_events_total')!=len(events):
        raise AuditError('Shadow event count changed since validated snapshot')
    if shadow_ref.get('shadow_classification')!=shadow['shadow_classification']:
        raise AuditError('Shadow classification drift')
    if shadow_ref.get('account_balance_comparison')!=shadow['account_balance_comparison']:
        raise AuditError('Shadow balance drift')
    q=storno_ref.get('focus',{})
    if not (q.get('id_and_ordinal_same_source_record') and q.get('requested_source_data_ordinal')==700
            and q.get('amb_state')=='ACTIVE' and q.get('parent_state')=='STORNO'
            and storno_ref.get('linked_storno_document_counts',{}).get('DOCUMENT_LEG_STATE_MISMATCH')==1):
        raise AuditError('Required Q-00700 storno status baseline has drifted')
    _,tables,idx=inputs(stage)
    source=tables['UNCONVERTED__tblAmbalaza']
    drafts=[classify_event(e,source,idx) for e in events]
    flat=[int(n) for ev in events for n in ev['source_row_ordinals'].split(';')]
    if len(flat)!=len(set(flat)) or 700 in flat:
        raise AuditError('Duplicate physical lineage or Q-00700 entered draft')
    if len(flat)!=shadow['shadow_classification']['active_physical_pair_events']*2+shadow['shadow_classification']['active_physical_single_events']:
        raise AuditError('Physical lineage row count invariant failed')
    if len(flat)+len(review)+base['legacy_state_counts'].get('STORNO',0)!=len(source):
        raise AuditError('Whole source coverage invariant failed')
    count=Counter(r['evidence_status'] for r in drafts)
    kind=Counter(r['VrstaKretanja_PROPOSAL'] for r in drafts)
    holds=Counter(reason for d in drafts if d['evidence_status']=='HOLD_INTEGRITY' for reason in d['reason'].split(';'))
    public={
       'format':VERSION, 'status':'READ_ONLY_SEMANTIC_PROPOSALS_NO_CANONICAL_AMB10_AND_NO_IMPORT',
       'source_sha256':SOURCE_SHA,'schema_sha256':base['schema_sha256'],
       'source_amb_rows':len(source), 'physical_events_examined':len(events),
       'source_rows_in_physical_events':len(flat),
       'quarantine_active_source_rows':len(review),
       'historical_storno_source_rows':base['legacy_state_counts'].get('STORNO',0),
       'draft_evidence_status_counts':dict(sorted(count.items())),
       'proposed_business_movement_kind_counts':dict(sorted(kind.items())),
       'integrity_hold_reasons':dict(sorted(holds.items())),
       'q00700_state_mismatch_isolated_from_draft':700 not in flat,
       'negative_legacy_account_type_combinations':shadow['account_balance_comparison']['negative_legacy_account_types'],
       'ready_for_import':False,
       'notes':['Purpose proposed solely from legacy document writer evidence, not authoritative business/legal ownership.',
                'OM-Izlaz-Koop has ambiguous title: IZDATA_PRAZNA vs VRACANJE_TUDJE; always HOLD.',
                'Even DIRECTION_KIND_PROPOSED_ONLY needs opening stock, ownership, chronology and storno gate.',
                'Q-00700 stays quarantined; no automatic storno/cancellation.',
                'No AMB-10 AmbID, StornoOd, opening balances or foreign obligations generated.',
                'All private IDs, amounts, dates and account names remain local. Do not import CSV into tblAmbalaza.']}
    return public,drafts


def render(s):
    lines=['AGRIX MIG-05 / AMB-10 SEMANTIC PREVIEW - READ ONLY',
           'Source SHA256: '+s['source_sha256'],
           'Physical events examined: '+str(s['physical_events_examined']),
           'Source legs in physical events: '+str(s['source_rows_in_physical_events']),
           'Candidate evidence statuses: '+json.dumps(s['draft_evidence_status_counts'],sort_keys=True),
           'Proposed business movement kinds: '+json.dumps(s['proposed_business_movement_kind_counts'],sort_keys=True),
           'Integrity holds: '+json.dumps(s['integrity_hold_reasons'],sort_keys=True),
           'Quarantined rows: '+str(s['quarantine_active_source_rows']),
           'Historical STORNO rows not replayed: '+str(s['historical_storno_source_rows']),
           'Q-00700 still quarantined: '+str(s['q00700_state_mismatch_isolated_from_draft']),
           'Negative legacy accounts: '+str(s['negative_legacy_account_type_combinations']),
           'STOP: drafts are NOT import-approved canonical AMB-10 ledger entries.']
    return '\n'.join(lines)+'\n'


def cli(argv=None):
    p=argparse.ArgumentParser(description='Fail-closed AMB-10 movement semantics DRY RUN; never import results')
    p.add_argument('--stage',type=Path,required=True)
    p.add_argument('--shadow',type=Path,required=True)
    p.add_argument('--storno',type=Path,required=True)
    p.add_argument('--out',type=Path,required=True)
    a=p.parse_args(argv)
    stage=a.stage.resolve();shadow=a.shadow.resolve();storno=a.storno.resolve();out=a.out.resolve()
    for d in (stage,shadow,storno):
        if not d.is_dir():raise AuditError('Required input directory missing')
        if out==d or out in d.parents or d in out.parents:raise AuditError('Input/output overlap')
    if out.exists():raise AuditError('Output exists; choose a new directory')
    summary,drafts=run(stage,shadow,storno)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_amb10_semantic_preview_summary.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_amb10_semantic_preview_summary.txt').write_text(render(summary),encoding='utf-8')
        private_csv(out/'amb10_direction_and_kind_PROPOSALS_PRIVATE.csv',drafts,PRIVATE_COLUMNS)
    except BaseException:
        for file in out.iterdir():
            if file.is_file():file.unlink()
        out.rmdir()
        raise
    print(render(summary));print('Only share summary JSON and TXT. Private CSV must remain local.');return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as e:
        print('MIG-05 STOP:',e,file=sys.stderr);raise SystemExit(2)
