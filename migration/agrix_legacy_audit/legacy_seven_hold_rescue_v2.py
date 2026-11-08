#!/usr/bin/env python3
"""MIG-05.4: evidence-only examination of seven outstanding active legacy AMB legs.

NEVER converts AMB-10 / writes XLSM. Driver-inverse physical vectors are only
accounting evidence; old document class, title, opening stock, storno and legal
ownership stay UNPROVEN. Source customer identifiers appear ONLY in local files.
"""
from __future__ import annotations
import argparse
from collections import Counter
import csv
import hashlib
import json
import os
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as shadow
import legacy_amb10_residual_v2 as residual
import legacy_expanded_physical_stage_v2 as expanded

VERSION='MIG-05.4'
EXPECTED_HOLDS=frozenset({700,1377,2317,3849,3939,7115,7815})
ROWS=('source_row_ordinal','AmbID','legacy_DokumentTIP','legacy_DokumentID','Datum',
      'EntitetTip','EntitetID','VozacID','Smer','TipAmbalaze','Kolicina','Stornirano',
      'forensic_flags','physical_proof_status','physical_hold_reasons',
      'provisional_from_type','provisional_from_id','provisional_to_type',
      'provisional_to_id','missing_document_label','business_kind','ownership',
      'canonical_document_identity','storno_status','import_approved')


def checked_summary(expanded_dir:Path,manifest:dict):
    p=expanded_dir/'agrix_expanded_physical_summary.json'
    if not p.is_file():raise AuditError('MIG-05.3 public summary missing')
    s=json.loads(p.read_text(encoding='utf-8'))
    checks={
        'format':'MIG-05.3',
        'source_sha256':f.SOURCE_SHA,
        'schema_sha256':manifest['schema_sha256'],
        'active_source_rows':7480,
        'physical_source_legs':7473,
        'expanded_physical_events':4669,
        'historical_storno_rows_unconverted':335,
        'still_excluded_active_source_legs':7,
        'unexplained_account_model_residual_count':0,
        'q00700_remains_excluded':True,
        'ready_for_import':False,
    }
    for k,v in checks.items():
        if s.get(k)!=v:raise AuditError('MIG-05.3 summary contract changed: '+k)
    holds=s.get('remaining_hold_row_ordinals',[])
    if (not isinstance(holds,list) or len(holds)!=7 or len(set(holds))!=7 or
            holds!=sorted(holds) or 700 not in holds):
        raise AuditError('MIG-05.3 seven hold ordinals invalid')
    if manifest['schema_sha256']=='b9916598d1f8d29b8e66b3fd0805d36f4c2e695df48fe6ef32155a7b6a765d73' and set(holds)!=EXPECTED_HOLDS:
        raise AuditError('Actual pinned seven hold ordinals drifted')
    return s


def load_prior_evidence(expanded_dir:Path,s:dict):
    p=expanded_dir/'expanded_physical_cases_PRIVATE.jsonl'
    if not p.is_file():raise AuditError('MIG-05.3 private event JSONL missing')
    hasher=hashlib.sha256()
    rows=[];source_used=[]
    with p.open(encoding='utf-8') as src:
        for line in src:
            if not line.endswith('\n'):raise AuditError('Physical JSONL record lacks newline')
            obj=json.loads(line)
            canonical=json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'
            # refuse any reordered/altered/corrupt private event file
            if line!=canonical:raise AuditError('MIG-05.3 physical evidence JSONL canonical format mismatch')
            hasher.update(canonical.encode('utf-8'))
            if obj.get('canonical_AMB10_import_authorized')!='NO':
                raise AuditError('Prior physical event import authorization unexpectedly changed')
            if obj.get('evidence_status') not in ('DIRECTION_KIND_PROPOSED_ONLY','HOLD_OWNERSHIP'):
                raise AuditError('Unknown prior physical evidence state')
            source_used.extend(int(z) for z in obj['source_row_ordinals'].split(';'))
            rows.append(obj)
    if hasher.hexdigest()!=s['expanded_physical_evidence_sha256']:
        raise AuditError('MIG-05.3 private evidence checksum mismatch')
    if len(rows)!=4669 or len(source_used)!=7473 or len(source_used)!=len(set(source_used)):
        raise AuditError('Prior physical event partition mismatch')
    if set(source_used)&set(s['remaining_hold_row_ordinals']):raise AuditError('Prior projection contains held row')
    return rows,set(source_used)


def physical_vector_only(r:dict, masters:dict):
    """Driver-account inverse exactly matching old balance; never assign purpose."""
    why=[];e=None
    typ=f.key(r.get('EntitetTip'));direction=f.key(r.get('Smer'))
    if not is_active(r):why.append('NOT_ACTIVE')
    if f.dec(r.get('Kolicina')) is None or f.dec(r.get('Kolicina'))<=0:
        why.append('INVALID_QUANTITY')
    if not trim(r.get('Datum')):why.append('MISSING_DATE')
    pack=trim(r.get('TipAmbalaze'))
    if not pack or f.key(pack) not in {f.key(v) for v in masters['tblTipAmbalaze']}:
        why.append('UNKNOWN_PACKAGING_TYPE')
    mastername={'stanica':'tblStanice','kupac':'tblKupci'}.get(typ)
    if mastername is None:why.append('NOT_DRIVER_INVERSE_ACCOUNT_CLASS')
    elif trim(r.get('EntitetID')) not in masters[mastername]:
        why.append('UNKNOWN_SOURCE_ACCOUNT')
    driver=trim(r.get('VozacID'))
    if not driver or driver not in masters['tblVozaci']:
        why.append('UNKNOWN_OR_MISSING_DRIVER')
    if direction not in ('ulaz','izlaz'):why.append('UNKNOWN_DIRECTION')
    if why:return None,sorted(set(why))
    source_role='Stanica' if typ=='stanica' else 'Kupac'
    account=(source_role,trim(r['EntitetID']));driver_role=('Vozac',driver)
    fr,to=(driver_role,account) if direction=='ulaz' else (account,driver_role)
    e={
        'source_row_ordinals':str(r['source_row_ordinal']),
        'document_type':r['DokumentTip'],'document_id':r['DokumentID'],
        'from_type':fr[0],'from_id':fr[1],'to_type':to[0],'to_id':to[1],
        'tip_ambalaze':pack,'quantity':r['Kolicina'],
        'directional_evidence':'EXACT_LEGACY_DRIVER_INVERSE_VECTOR_ONLY',
    }
    from collections import defaultdict
    from decimal import Decimal
    result=defaultdict(Decimal)
    shadow.add_event(result,e)
    source=defaultdict(Decimal)
    for key,amount in shadow.raw_contributions(r):source[key]+=amount
    if {k:v for k,v in result.items() if v}!={k:v for k,v in source.items() if v}:
        return None,['LEGACY_VS_PROPOSED_VECTOR_MISMATCH']
    return e,[]


def build(stage:Path,expanded_dir:Path):
    manifest,tables,idx=f.inputs(stage)
    prior=checked_summary(expanded_dir,manifest)
    existing,used=load_prior_evidence(expanded_dir,prior)
    baseline,raw=f.analyze(stage)
    if len(raw)!=7815 or baseline['legacy_state_counts']!={'ACTIVE':7480,'STORNO':335}:
        raise AuditError('Pinned source AMB population changed')
    sources={r['source_row_ordinal']:r for r in raw}
    remaining={r['source_row_ordinal'] for r in raw if is_active(r)}-used
    if remaining!=set(prior['remaining_hold_row_ordinals']):
        raise AuditError('Seven residual source rows differ from MIG-05.3')
    if any(not is_active(sources[n]) for n in remaining):raise AuditError('Held row not active')
    masters={name:idx[name] for name in ('tblStanice','tblKupci','tblVozaci')}
    masters['tblTipAmbalaze']={f.key(r['TipAmbalaze']):r for r in tables['tblTipAmbalaze']}

    private=[];candidate_events=[];state=Counter();role=Counter();missing_doc=0
    for n in sorted(remaining):
        r=sources[n];parent_mismatch=(n==700)
        if parent_mismatch:
            status='HOLD_STORNO_PARENT_CONFLICT';reasons=['ACTIVE_LEG_STORNO_PARENT'];new=None
        else:
            new,reasons=physical_vector_only(r,masters)
            status='PHYSICAL_VECTOR_RECONSTRUCTED_SEMANTICS_HELD' if new else 'HOLD_PHYSICAL_UNPROVEN'
        state[status]+=1
        role[f.key(r.get('EntitetTip')) or 'blank']+=1
        if not trim(r.get('DokumentID')):missing_doc+=1
        private.append({
            'source_row_ordinal':n,'AmbID':r['AmbID'],
            'legacy_DokumentTIP':r['DokumentTip'],'legacy_DokumentID':r['DokumentID'],
            'Datum':r['Datum'],'EntitetTip':r['EntitetTip'],'EntitetID':r['EntitetID'],
            'VozacID':r['VozacID'],'Smer':r['Smer'],'TipAmbalaze':r['TipAmbalaze'],
            'Kolicina':r['Kolicina'],'Stornirano':r['Stornirano'],
            'forensic_flags':r.get('forensic_flags',''),
            'physical_proof_status':status,'physical_hold_reasons':';'.join(reasons),
            'provisional_from_type':new['from_type'] if new else '',
            'provisional_from_id':new['from_id'] if new else '',
            'provisional_to_type':new['to_type'] if new else '',
            'provisional_to_id':new['to_id'] if new else '',
            'missing_document_label':'YES' if not trim(r.get('DokumentID')) else 'NO',
            'business_kind':'UNASSIGNED','ownership':'UNPROVEN',
            'canonical_document_identity':'NOT_RESOLVED',
            'storno_status':'NOT_REPLAYED', 'import_approved':'NO'})
        if new:
            # Nothing synthetic is allowed to look like an approved AMB document.
            case={
                'physical_case_id':f'PHY-{4669+len(candidate_events)+1:05d}',
                'source_row_ordinals':str(n),'source_amb_ids':trim(r['AmbID']),
                'legacy_DokumentTIP':r['DokumentTip'],'legacy_DokumentID':r['DokumentID'],
                'legacy_Datum':trim(r['Datum']),'TipAmbalaze':trim(r['TipAmbalaze']),
                'Kolicina':r['Kolicina'],'OdNalogTip':new['from_type'],
                'OdNalogID':new['from_id'],'NaNalogTip':new['to_type'],
                'NaNalogID':new['to_id'],'physical_evidence':new['directional_evidence'],
                'event_provenance':'MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY',
                'VrstaKretanja_PROPOSAL':'UNASSIGNED',
                'evidence_status':'HOLD_DOCUMENT_TYPE_AND_OWNERSHIP',
                'reason':'UNKNOWN_DOCUMENT_SEMANTICS_AND_IDENTITY_NOT_CANONICAL',
                'ownership':'UNPROVEN','opening_balance':'UNPROVEN',
                'historical_storno':'NOT_REPLAYED','canonical_AMB10_import_authorized':'NO'}
            candidate_events.append(case)

    still=remaining-{int(x['source_row_ordinals']) for x in candidate_events}
    if 700 not in still:raise AuditError('Q-00700 must remain isolated')
    if len(still)+len(candidate_events)!=7:raise AuditError('Seven leg partition incorrect')
    after=existing+candidate_events
    shadow_events=[dict(source_row_ordinals=e['source_row_ordinals'],
        from_type=e['OdNalogTip'],from_id=e['OdNalogID'],
        to_type=e['NaNalogTip'],to_id=e['NaNalogID'],
        tip_ambalaze=e['TipAmbalaze'],quantity=e['Kolicina']) for e in after]
    account_rows,_,counts,_,_,_=residual.attribute(
        raw,shadow_events,[{'row':n,'reason':'NOT_RECONSTRUCTED'} for n in sorted(still)])
    if (counts['model_identity_failures_all_accounts']!=0 or
        counts['quarantined_review_rows']!=len(still) or
        counts['projected_source_row_count']!=7473+len(candidate_events) or
        counts['historical_source_rows_not_projected']!=335 or
        counts['negative_legacy_accounts']!=prior['negative_legacy_account_types']):
        raise AuditError('After-rescue physical/account residual invariant failed')
    if len(after)!=4669+len(candidate_events):raise AuditError('Physical event total inconsistent')
    if len(set(x['physical_case_id'] for x in after))!=len(after):
        raise AuditError('Physical case IDs duplicated')
    report={
        'format':VERSION,'status':'SEVEN_HOLD_PHYSICAL_EVIDENCE_ONLY_NOT_AMB10_APPROVED',
        'source_sha256':f.SOURCE_SHA,'schema_sha256':manifest['schema_sha256'],
        'parent_expanded_physical_evidence_sha256':prior['expanded_physical_evidence_sha256'],
        'source_rows_total':7815,'active_source_rows':7480,'historical_storno_rows_unconverted':335,
        'previous_physical_events':4669,'previous_physical_source_legs':7473,
        'previous_excluded_active_legs':7,
        'seven_case_status_counts':dict(sorted(state.items())),
        'legacy_roles_for_seven':dict(sorted(role.items())),
        'seven_missing_document_labels':missing_doc,
        'recovered_extra_physical_vector_cases':len(candidate_events),
        'new_physical_events_total':len(after),
        'new_physical_source_legs':7473+len(candidate_events),
        'remaining_active_hold_source_ordinals':sorted(still),
        'remaining_active_hold_source_leg_count':len(still),
        'ending_balance_mismatch_account_types_after':counts['ending_balances_mismatched'],
        'unexplained_account_residual_count':counts['model_identity_failures_all_accounts'],
        'legacy_negative_account_types_unchanged':counts['negative_legacy_accounts'],
        'new_unmodelled_legs_have_canonical_document_identity':False,
        'any_import_approved':False,
        'notes':[
          'Unknown legacy DokumentTIP remains private; no assumption of Kupci-Otpremnica.',
          'Driver inverse vector matching is accounting evidence, not source document/ownership proof.',
          'Q-00700 remains a parent-storno versus active-amb status conflict.',
          'Missing document label may still have a reconstructed physical vector, NEVER identity approval.',
          'No target workbook changes, no new AMB-10 documents/rows, no StornoOd generated.',
        ],
    }
    return report,private,after,account_rows


def render(s:dict):
    keys=['source_sha256','previous_physical_events','previous_excluded_active_legs',
        'seven_case_status_counts','legacy_roles_for_seven','seven_missing_document_labels',
        'recovered_extra_physical_vector_cases','new_physical_events_total',
        'remaining_active_hold_source_ordinals','ending_balance_mismatch_account_types_after',
        'unexplained_account_residual_count','legacy_negative_account_types_unchanged',
        'any_import_approved']
    return ('AGRIX MIG-05.4 / SEVEN LEG PHYSICAL VECTOR ONLY / NO AMB-10 IMPORT\n'+
        ''.join(k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in keys)+
        'Do not share PRIVATE evidence or write to the target XLSM.\n')


def cli(argv=None):
    p=argparse.ArgumentParser(description='MIG-05.4 read-only seven held source legs')
    for name in ('stage','expanded','out'):p.add_argument('--'+name,type=Path,required=True)
    a=p.parse_args(argv)
    stage=a.stage.resolve();old=a.expanded.resolve();out=a.out.resolve()
    if not stage.is_dir() or not old.is_dir():raise AuditError('Missing input directory')
    if out.exists() or any(out==x or out in x.parents or x in out.parents for x in (stage,old)):
        raise AuditError('Output folder exists/overlaps input; use fresh output folder')
    report,private,events,accounts=build(stage,old)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_seven_hold_summary.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_seven_hold_summary.txt').write_text(render(report),encoding='utf-8')
        f.private_csv(out/'seven_case_reconstruction_PRIVATE.csv',private,ROWS)
        fields=('account_type','account_id','tip_ambalaze','legacy_ending_balance',
                'shadow_ending_balance','excluded_review_net','model_unexplained_residual',
                'review_source_rows','legacy_negative','ending_balance_matches')
        f.private_csv(out/'post_rescue_account_balances_PRIVATE.csv',accounts,fields)
        dest=out/'expanded_plus_held_physical_cases_PRIVATE.jsonl'
        with dest.open('w',encoding='utf-8',newline='\n') as fh:
            for ev in events:fh.write(json.dumps(ev,ensure_ascii=False,sort_keys=True)+'\n')
        os.chmod(dest,0o600)
    except BaseException:
        for child in out.iterdir():
            if child.is_file():child.unlink()
        out.rmdir();raise
    print(render(report));print('Upload ONLY agrix_seven_hold_summary.json/txt, NEVER private files.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
        print('MIG-05.4 STOP:',exc,file=sys.stderr);raise SystemExit(2)
