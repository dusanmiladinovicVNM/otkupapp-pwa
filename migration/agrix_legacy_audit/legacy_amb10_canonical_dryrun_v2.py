#!/usr/bin/env python3
"""MIG-05.6 — strictly local, read-only target-shaped AMB-10 PREVIEW.

Does not infer opening stock, legal title, purchases, covers, document identities,
storno inverse entries or import approval. Negative legacy balances ARE VALID and
must be preserved EXACTLY (including sign), not repaired.

No XLSM reading/writing here; this accepts only the source-pinned MIG-02 / 05.3 /
05.4 local evidence. Private outputs are NOT valid CSVs for importing into Excel.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
import os
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as shadow
import legacy_amb10_chronology_preview_v2 as chronology

VERSION='MIG-05.6'
DOCS={'otkup':('SOURCE__tblOtkup','OtkupID','Otkup'),
      'otpremnica':('SOURCE__tblOtpremnica','OtpremnicaID','Otpremnica'),
      'prijemnica':('tblPrijemnica','PrijemnicaID','Prijemnica')}
PARTNERS={'Kooperant','Kupac'}
OWN={'Stanica','Vozac','Firma'}
REAL=PARTNERS|OWN
KNOWN_KINDS={'AMBALAZA_UZ_ROBU','POVRAT_PRAZNE'}


def canonical_preview(event:dict,indices:dict):
    """Target column *proposal* only, not a valid target ledger row.

    Failed guard returns explicit HOLDS. Even pass is BLOCKED for ownership,
    replay of storno, opening title and migration-only negative-balances gate.
    """
    status=event['evidence_status']
    kind=event['VrstaKretanja_PROPOSAL']
    holds=[]
    if event.get('canonical_AMB10_import_authorized')!='NO':
        raise AuditError('Upstream event unexpectedly claims import authorization')
    if status!='DIRECTION_KIND_PROPOSED_ONLY' or kind not in KNOWN_KINDS:
        holds.append('PURPOSE_OR_OWNERSHIP_NOT_PROVEN')
    doc=DOCS.get(f.key(event['legacy_DokumentTIP']))
    doc_id=trim(event.get('legacy_DokumentID'))
    if not doc or not doc_id:holds.append('CANONICAL_SOURCE_DOCUMENT_IDENTITY_UNRESOLVED')
    elif doc_id not in indices[doc[0]]:holds.append('CANONICAL_SOURCE_PARENT_NOT_FOUND')
    elif not is_active(indices[doc[0]][doc_id][1]):holds.append('CANONICAL_SOURCE_PARENT_STORNIRANO')
    if kind in KNOWN_KINDS:
        fr=event['OdNalogTip'];to=event['NaNalogTip']
        if fr not in REAL or to not in REAL or fr==to and event['OdNalogID']==event['NaNalogID']:
            holds.append('INVALID_CANONICAL_ACCOUNT_ROLES')
        if kind=='POVRAT_PRAZNE' and not (fr in PARTNERS and to in OWN):
            holds.append('KIND_ACCOUNT_MATRIX_CONFLICT')
    quantity=f.dec(event['Kolicina'])
    if quantity is None or quantity<=0:raise AuditError('Invalid physical event quantity')
    d=chronology.date_key(event['legacy_Datum'])
    if d is None:holds.append('DATE_UNPARSEABLE')
    if status=='DIRECTION_KIND_PROPOSED_ONLY' and not holds:
        proposal={
          'Datum':d.isoformat(),'TipAmbalaze':event['TipAmbalaze'],
          'Kolicina':str(quantity),'DokumentID':doc_id,'DokumentTIP':doc[2],
          'OdNalogTip':event['OdNalogTip'],'OdNalogID':event['OdNalogID'],
          'NaNalogTip':event['NaNalogTip'],'NaNalogID':event['NaNalogID'],
          'VrstaKretanja':kind,
        }
    else:proposal=None
    return proposal,sorted(set(holds))


def _ledger_delta(e,net):
    q=f.dec(e['Kolicina'])
    if q is None or q<=0:raise AuditError('Invalid converted physical quantity')
    fr=(e['OdNalogTip'],e['OdNalogID'],e['TipAmbalaze'])
    to=(e['NaNalogTip'],e['NaNalogID'],e['TipAmbalaze'])
    if not all(fr) or not all(to) or fr==to:
        raise AuditError('Bad physical transfer accounts')
    net[fr]-=q;net[to]+=q


def _balances(raw,events,expected_negative_accounts):
    """Independently reconcile every source account and tip, negatives unchanged."""
    legacy=defaultdict(Decimal);shadow_net=defaultdict(Decimal)
    for r in raw:
        for acc,delta in shadow.raw_contributions(r):legacy[acc]+=delta
    for e in events:_ledger_delta(e,shadow_net)
    q=raw[699]
    if q['source_row_ordinal']!=700 or not is_active(q):
        raise AuditError('Q-00700 no longer the active legacy hold')
    held=defaultdict(Decimal)
    for acc,delta in shadow.raw_contributions(q):held[acc]+=delta
    account_rows=[];different=negatives=unexplained=0
    for acc in sorted(set(legacy)|set(shadow_net)|set(held)):
        actual=legacy[acc];proj=shadow_net[acc];hold=held[acc]
        diff=actual-proj;residual=diff-hold
        if diff:different+=1
        if actual<0:negatives+=1
        if residual:unexplained+=1
        account_rows.append({'account_type':acc[0],'account_id':acc[1],
            'TipAmbalaze':acc[2],'source_net_active_ending':str(actual),
            'all_4675_physical_ending':str(proj),'q00700_unprojected_effect':str(hold),
            'legacy_minus_4675':str(diff),'unexplained':str(residual),
            'legacy_negative_allowed':'YES' if actual<0 else 'NO'})
    if different!=2 or negatives!=expected_negative_accounts or unexplained:
        raise AuditError('Legacy balance parity drift — no auto-correction allowed (different=%s negatives=%s unexplained=%s)' %(different,negatives,unexplained))
    return account_rows,{'source_negative_balance_accounts_preserved':negatives,
        'accounts_different_until_q00700_resolved':different,
        'unexplained_account_balance_residuals':unexplained,
        'synthetic_opening_balances_created':0,
        'migration_balance_policy':'PRESERVE_SOURCE_SIGN_AND_EXACT_AMOUNT'}


def build(stage:Path,expanded:Path,seven_dir:Path):
    manifest,tables,indices=f.inputs(stage)
    seven_report,events,upstream_digest=chronology._verify_inputs(stage,expanded,seven_dir)
    if len(events)!=4675:raise AuditError('Unexpected MIG-05.4 physical population')
    base,raw=f.analyze(stage)
    q=raw[699]
    if f.key(q.get('DokumentTip'))!='otpremnica' or not trim(q.get('DokumentID')):
        raise AuditError('Q-00700 no longer refers to an Otpremnica')
    parent=indices['SOURCE__tblOtpremnica'].get(trim(q['DokumentID']))
    if (not parent or is_active(parent[1]) or not is_active(q) or
        f.dec(q['Kolicina'])!=f.dec(parent[1].get('KolAmbalaze')) or
        trim(q['EntitetID'])!=trim(parent[1].get('StanicaID')) or
        trim(q['VozacID'])!=trim(parent[1].get('VozacID'))):
        raise AuditError('Q-00700 storno-parent evidence changed')
    if len(raw)!=7815 or base['legacy_state_counts']!={'ACTIVE':7480,'STORNO':335}:
        raise AuditError('Pinned historical source population changed')
    origins=[];held=[];sources=[];key_groups=defaultdict(list);pair_by_doc=defaultdict(set)
    ordinals=[]
    for ev in events:
        case=ev['physical_case_id']
        ids=[int(n) for n in ev['source_row_ordinals'].split(';')]
        ordinals+=ids
        if any(not is_active(raw[n-1]) for n in ids):
            raise AuditError('Historical storno passed into active event')
        candidate,reasons=canonical_preview(ev,indices)
        record={'physical_case_id':case,'source_row_ordinals':ev['source_row_ordinals'],
           'legacy_source_amb_ids':ev['source_amb_ids'],
           'legacy_doc_label':ev['legacy_DokumentID'],
           'source_evidence_status':ev['evidence_status'],
           'movement_preview':candidate,'blocking_reasons':reasons,
           'opening_stock':'SOURCE_EXACT_NEGATIVES_ALLOWED__NO_OPENING_CREATED',
           'ownership':'UNPROVEN','historical_storno':'NOT_REPLAYED',
           'new_amb_id':None,'storno_od':None,
           'import_approved':False}
        if candidate is not None:
            key=(candidate['DokumentTIP'],candidate['DokumentID'],
                 candidate['VrstaKretanja'],candidate['TipAmbalaze'])
            key_groups[key].append(record)
            parties=frozenset(((candidate['OdNalogTip'],candidate['OdNalogID']),
                               (candidate['NaNalogTip'],candidate['NaNalogID'])))
            pair_by_doc[(candidate['DokumentTIP'],candidate['DokumentID'])].add(parties)
        origins.append(record)
    if len(ordinals)!=7479 or len(set(ordinals))!=7479 or 700 in ordinals:
        raise AuditError('Preview uses repeated or held source leg')
    if set(ordinals)!={r['source_row_ordinal'] for r in raw if is_active(r)}-{700}:
        raise AuditError('Active evidence partition missing or added source rows')
    # AMB-INV-04 original identity and AMB-INV-10 single business party pair.
    for group in key_groups.values():
        if len(group)>1:
            for rec in group:
                rec['blocking_reasons'].append('AMB_INV_04_ORIGINAL_IDENTITY_COLLISION')
                rec['movement_preview']=None
    for rec in origins:
        preview=rec['movement_preview']
        if preview and len(pair_by_doc[(preview['DokumentTIP'],preview['DokumentID'])])>1:
            rec['blocking_reasons'].append('AMB_INV_10_DOCUMENT_PARTY_PAIR_COLLISION')
            rec['movement_preview']=None
        if rec['blocking_reasons']:held.append(rec)
    reason_counts=Counter(reason for v in origins for reason in v['blocking_reasons'])
    preview=[r for r in origins if r['movement_preview'] is not None]
    # These STILL ARE NOT IMPORT READY: title, opening, lifecycle, and not-yet
    # reconciled Q-00700 prevent a valid canonical ledger.
    if len(preview)+len(held)!=len(origins):raise AuditError('Preview partition drift')
    balances,checks=_balances(raw,events,seven_report['legacy_negative_account_types_unchanged'])
    historical=[{'source_row_ordinal':r['source_row_ordinal'],'AmbID':r['AmbID'],
                 'legacy_DokumentTIP':r['DokumentTip'],'legacy_DokumentID':r['DokumentID'],
                 'legacy_Datum':r['Datum'],'legacy_Kolicina':r['Kolicina'],
                 'Stornirano':r['Stornirano'],
                 'status':'STORNO_HISTORY_UNCONVERTED','StornoOd':None,
                 'import_approved':False} for r in raw if not is_active(r)]
    if len(historical)!=335:raise AuditError('Storno population drift')
    s={
       'format':VERSION,'status':'TARGET_SHAPED_DRYRUN_NOT_IMPORT_APPROVED',
       'source_sha256':f.SOURCE_SHA,'schema_sha256':manifest['schema_sha256'],
       'verified_mig05_4_private_evidence_sha256':upstream_digest,
       'physical_events_total':len(events),'active_legacy_source_legs_covered':len(ordinals),
       'active_source_leg_still_held':700,
       'target_shaped_candidate_count':len(preview),'target_shaped_hold_count':len(held),
       'hold_reason_counts':dict(sorted(reason_counts.items())),
       'no_canonical_amb_id_allocated':True,
       'no_storno_inverse_written':True,'historical_storno_source_rows_unreplayed':len(historical),
       'ownership_verified':False,'import_approved':False,
       **checks,
       'notes':[
           'TARGET-SHAPED candidate is evidence-only; no row is legal to paste into tblAmbalaza.',
           'Source negative balances (including 35 negative account/type pairs) are ALLOWED; do not correct or create artificial opening stock.',
           'Originals only are represented; 335 historical STORNO source legs and Q-00700 remain separately unconverted.',
           'No invented AmbalazaDokument/Revers PK, no legal title, no deficit cover, no automatic third-party liabilities.',
           'The target workbook, existing source files and customer data were not modified.',
       ]}
    return s,origins,balances,historical


def render(s):
    fields=['physical_events_total','active_legacy_source_legs_covered',
            'target_shaped_candidate_count','target_shaped_hold_count',
            'hold_reason_counts','source_negative_balance_accounts_preserved',
            'accounts_different_until_q00700_resolved','unexplained_account_balance_residuals',
            'historical_storno_source_rows_unreplayed','import_approved']
    return ('AGRIX MIG-05.6 / TARGET-SHAPED DRY RUN, NO IMPORT\n'+''.join(
        k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in fields)+
        'Private files are not import tables. Negative source balances are intentionally preserved.\n')


def cli(argv=None):
    p=argparse.ArgumentParser(description='AgriX MIG-05.6 guarded canonical AMB-10 preview')
    for name in ('stage','expanded','seven','out'):p.add_argument('--'+name,type=Path,required=True)
    a=p.parse_args(argv)
    stage=a.stage.resolve();expanded=a.expanded.resolve();seven_dir=a.seven.resolve();out=a.out.resolve()
    if out.exists() or any(out==x or out in x.parents or x in out.parents for x in (stage,expanded,seven_dir)):
        raise AuditError('Output exists or overlaps source directories')
    if not all(x.is_dir() for x in (stage,expanded,seven_dir)):
        raise AuditError('Missing required input stage folder')
    s,origins,balances,historical=build(stage,expanded,seven_dir)
    out.mkdir(parents=True,exist_ok=False)
    try:
        for name,seq in (
            ('AMB10_DRYRUN_ALL_CASES_PRIVATE.jsonl',origins),
            ('AMB10_DRYRUN_BALANCE_PARITY_PRIVATE.jsonl',balances),
            ('AMB10_HISTORICAL_STORNO_HOLD_PRIVATE.jsonl',historical)):
            with (out/name).open('w',encoding='utf-8',newline='\n') as fh:
                for rec in seq:fh.write(json.dumps(rec,ensure_ascii=False,sort_keys=True)+'\n')
            os.chmod(out/name,0o600)
        (out/'agrix_amb10_dryrun_summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (out/'agrix_amb10_dryrun_summary.txt').write_text(render(s),encoding='utf-8')
    except BaseException:
        for child in out.iterdir():
            if child.is_file():child.unlink()
        out.rmdir();raise
    print(render(s))
    print('Upload ONLY agrix_amb10_dryrun_summary.json/txt; KEEP PRIVATE JSONL LOCAL')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
        print('MIG-05.6 STOP:',exc,file=sys.stderr)
        raise SystemExit(2)
