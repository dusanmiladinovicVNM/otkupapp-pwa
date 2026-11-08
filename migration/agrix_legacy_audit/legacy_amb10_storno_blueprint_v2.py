#!/usr/bin/env python3
"""MIG-05.7 read-only legacy STORNO history reconstruction, never canonical rows.

335 old Stornirano rows are cancelled ORIGINAL legs, not new AMB-10 contra-entries.
A historical cancellation would require a linked original + reverse and a known
business date/identity; this tool proves physical grouping ONLY and creates neither.
Never repairs negatives, invents legal ownership or mutates .xlsm/staging.
"""
from __future__ import annotations
import argparse
from collections import Counter,defaultdict
from decimal import Decimal
import hashlib,json,os,sys
from pathlib import Path
from legacy_audit import AuditError,is_active,trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as sh
import legacy_amb10_canonical_dryrun_v2 as dry
import legacy_seven_hold_rescue_v2 as seven

VERSION='MIG-05.7'

def _read_verified_private(path:Path,expected:list):
    if not path.is_file():raise AuditError('Missing private upstream file '+path.name)
    digest=hashlib.sha256();count=0
    with path.open('r',encoding='utf-8',newline='') as inp:
        for line in inp:
            if count>=len(expected):raise AuditError('Extra private upstream records')
            good=json.dumps(expected[count],ensure_ascii=False,sort_keys=True)+'\n'
            if line!=good:raise AuditError('Private upstream record mutated: '+path.name)
            digest.update(line.encode('utf-8'));count+=1
    if count!=len(expected):raise AuditError('Missing private upstream records')
    return digest.hexdigest()

def _balances(rows):
    out=defaultdict(Decimal)
    for r in rows:
        for k,v in sh.raw_contributions(r):out[k]+=v
    return {k:v for k,v in out.items() if v}

def _event_delta(event):
    amount=f.dec(event['Kolicina'])
    if amount is None or amount<=0:raise AuditError('Nonpositive event amount')
    d=defaultdict(Decimal)
    od=(event['OdNalogTip'],event['OdNalogID'],event['TipAmbalaze'])
    na=(event['NaNalogTip'],event['NaNalogID'],event['TipAmbalaze'])
    if not all(od+na) or od==na:raise AuditError('Invalid endpoints')
    d[od]-=amount;d[na]+=amount
    return {k:v for k,v in d.items() if v}

def _from_two(a,b,master_ids):
    """Old book legs are one transfer ONLY if every account delta matches."""
    if {f.key(a['Smer']),f.key(b['Smer'])}!={'ulaz','izlaz'}:return None,'PAIR_DIRECTIONS_INVALID'
    if {a['EntitetTip'],b['EntitetTip']}!={'Kooperant','Stanica'}:
        return None,'PAIR_NOT_KOOPERANT_STANICA'
    if any(f.key(a[z])!=f.key(b[z]) for z in ('Datum','TipAmbalaze','Kolicina','DokumentTip','DokumentID')):
        return None,'PAIR_DIFFERENT_SOURCE_FIELDS'
    if f.dec(a['Kolicina']) is None or f.dec(a['Kolicina'])<=0:return None,'INVALID_QUANTITY'
    if any(trim(z['VozacID']) for z in (a,b)):return None,'PAIRED_WITH_DRIVER_NOT_SUPPORTED'
    if any(trim(z['EntitetID']) not in master_ids[{'Kooperant':'tblKooperanti','Stanica':'tblStanice'}[z['EntitetTip']]] for z in (a,b)):
        return None,'ACCOUNT_MASTER_NOT_FOUND'
    directions={f.key(x['Smer']):x for x in (a,b)}
    from_leg=directions['izlaz'];to_leg=directions['ulaz']
    ev={'OdNalogTip':from_leg['EntitetTip'],'OdNalogID':from_leg['EntitetID'],
        'NaNalogTip':to_leg['EntitetTip'],'NaNalogID':to_leg['EntitetID'],
        'TipAmbalaze':trim(a['TipAmbalaze']),'Kolicina':trim(a['Kolicina'])}
    # Hypothetical ACTIVE original from the two old cancelled records. Storno
    # flag gets cleared IN MEMORY ONLY for vector comparison, not in source.
    expected=_balances([dict(z,Stornirano='') for z in (a,b)])
    if _event_delta(ev)!=expected:return None,'PAIR_BALANCE_VECTOR_MISMATCH'
    return ev,''

def _from_one(a,masters):
    # Use established MIG-05.4 physical classifier, but only hypothetically
    # treat old cancelled leg as pre-cancellation to check its vector.
    rr=dict(a,Stornirano='')
    ev,why=seven.physical_vector_only(rr,masters)
    if why:return None,'SINGLE_'+','.join(why)
    x={'OdNalogTip':ev['from_type'],'OdNalogID':ev['from_id'],
       'NaNalogTip':ev['to_type'],'NaNalogID':ev['to_id'],
       'TipAmbalaze':ev['tip_ambalaze'],'Kolicina':ev['quantity']}
    if _event_delta(x)!=_balances([rr]):return None,'SINGLE_VECTOR_MISMATCH'
    return x,''

def reconstruct(storno:list,masters:dict):
    """Group only unique evidenced cancelled entries; never pair across state."""
    by_document=defaultdict(list)
    for r in storno:
        if is_active(r):raise AuditError('ACTIVE row in storno history')
        by_document[(f.key(r['DokumentTip']),trim(r['DokumentID']))].append(r)
    events=[];holds=[];shape=Counter();reasons=Counter();used=[]
    for (_typ,_doc),records in sorted(by_document.items()):
        # Distinct original events with same business label are NOT merged.
        # Partition by exact date, packaging type, quantity and state.
        by_signature=defaultdict(list)
        for r in records:
            sig=(f.key(r['Datum']),f.key(r['TipAmbalaze']),trim(r['Kolicina']))
            by_signature[sig].append(r)
        # Ambiguous repetition of a document label may denote multiple events;
        # all group entries remain held rather than creating an invented PK.
        unique_scope=len(by_signature)==1
        for sig,legs in sorted(by_signature.items()):
            legs=sorted(legs,key=lambda r:r['source_row_ordinal'])
            ids=[r['source_row_ordinal'] for r in legs];used.extend(ids)
            kind='MIRROR_PAIR' if len(legs)==2 else 'DRIVER_SINGLE' if len(legs)==1 else 'AMBIGUOUS_GROUP'
            ev=None;why=''
            if not unique_scope:why='MULTIPLE_SIGNATURES_SAME_LEGACY_DOCUMENT_LABEL'
            elif len(legs)==2:ev,why=_from_two(legs[0],legs[1],masters)
            elif len(legs)==1:ev,why=_from_one(legs[0],masters)
            else:why='UNIQUE_STORNO_ORIGINAL_NOT_PROVABLE'
            if not all(trim(r.get('AmbID')) for r in legs):why=why or 'MISSING_OLD_AMB_ID';ev=None
            if not all(trim(r.get('Datum')) for r in legs):why=why or 'MISSING_ORIGINAL_DATE';ev=None
            if not all(f.dec(r['Kolicina']) is not None and f.dec(r['Kolicina'])>0 for r in legs):
                why=why or 'INVALID_QUANTITY';ev=None
            # No source document identity, ownership, storno date or original-to-inverse
            # canonical reference is established by any of these physical checks.
            common={'source_row_ordinals':';'.join(map(str,ids)),
              'legacy_amb_ids':';'.join(trim(r['AmbID']) for r in legs),
              'legacy_DokumentTip':legs[0]['DokumentTip'],
              'legacy_DokumentID':legs[0]['DokumentID'],
              'legacy_Datum':legs[0]['Datum'],
              'TipAmbalaze':legs[0]['TipAmbalaze'],
              'Kolicina':legs[0]['Kolicina'],
              'physical_group_shape':kind,
              'cancelled_legacy_state':'STORNO',
              'historical_cancellation_timestamp':'UNPROVEN',
              'new_AMB_ID':None,'new_StornoOd':None,
              'canonical_document_id':'UNRESOLVED',
              'business_kind':'UNPROVEN','ownership':'UNPROVEN',
              'import_approved':False}
            if ev is not None and not why:
                common.update(ev)
                common['proven_physical_vector_only']=True
                common['replay_policy']='ORIGINAL_AND_INVERSE_REQUIRED_BUT_NOT_GENERATED'
                events.append(common);shape[kind]+=1
                # Mathematical cancellation check, not actual lifecycle replay.
                delta=_event_delta(common)
                if any(v+(-v)!=0 for v in delta.values()):raise AuditError('Inverse vector arithmetic failed')
            else:
                common['proven_physical_vector_only']=False
                common['hold_reason']=why or 'NO_PHYSICAL_EVIDENCE'
                common['replay_policy']='HOLD_NO_REPLAY'
                holds.append(common);reasons[common['hold_reason']]+=len(legs)
    if len(used)!=len(set(used)) or set(used)!={r['source_row_ordinal'] for r in storno}:
        raise AuditError('Historical storno source partition is incomplete or duplicate')
    return events,holds,shape,reasons

def build(stage,expanded,seven_dir,dry_dir):
    manifest,tables,indices=f.inputs(stage)
    expected,all_cases,balances,historical=dry.build(stage,expanded,seven_dir)
    report=dry_dir/'agrix_amb10_dryrun_summary.json'
    if not report.is_file() or json.loads(report.read_text(encoding='utf-8'))!=expected:
        raise AuditError('Stored MIG-05.6 report differs from recalculated source')
    source_digests={}
    for fname,source in [('AMB10_DRYRUN_ALL_CASES_PRIVATE.jsonl',all_cases),
                         ('AMB10_DRYRUN_BALANCE_PARITY_PRIVATE.jsonl',balances),
                         ('AMB10_HISTORICAL_STORNO_HOLD_PRIVATE.jsonl',historical)]:
        source_digests[fname]=_read_verified_private(dry_dir/fname,source)
    baseline,raw=f.analyze(stage)
    if len(raw)!=7815 or baseline['legacy_state_counts']!={'ACTIVE':7480,'STORNO':335}:
        raise AuditError('Pinned source population drift')
    storno=[r for r in raw if not is_active(r)]
    if len(storno)!=len(historical) or any(r['source_row_ordinal']!=h['source_row_ordinal'] or
            r['AmbID']!=h['AmbID'] for r,h in zip(storno,historical)):
        raise AuditError('Historical storno lineage differs from MIG-05.6')
    masters={k:indices[k] for k in ('tblStanice','tblKooperanti','tblKupci','tblVozaci')}
    masters['tblTipAmbalaze']={f.key(r['TipAmbalaze']):r for r in tables['tblTipAmbalaze']}
    ev,held,shapes,reasons=reconstruct(storno,masters)
    account_impact={}
    for r in storno:
        if sh.raw_contributions(r):raise AuditError('Stornirano source row changes active balance')
    n_rows=sum(len(x['source_row_ordinals'].split(';')) for x in ev+held)
    if n_rows!=335:raise AuditError('Historical STORNO originals missing from record set')
    summary={
      'format':VERSION,'status':'CANCELLED_LEGACY_ORIGINALS_PHYSICAL_BLUEPRINT_ONLY_NO_AMB10_REPLAY',
      'source_sha256':f.SOURCE_SHA,'schema_sha256':manifest['schema_sha256'],
      'dryrun_record_checksum_sha256':source_digests['AMB10_DRYRUN_ALL_CASES_PRIVATE.jsonl'],
      'historical_source_storno_legs':len(storno),
      'cancelled_physical_event_candidates':len(ev),
      'cancelled_physical_source_legs_covered':sum(len(x['source_row_ordinals'].split(';')) for x in ev),
      'historical_source_legs_held':sum(len(x['source_row_ordinals'].split(';')) for x in held),
      'cancelled_physical_shape_counts':dict(sorted(shapes.items())),
      'held_storno_source_leg_reason_counts':dict(sorted(reasons.items())),
      'active_physical_events_unchanged':expected['physical_events_total'],
      'active_physical_source_legs_unchanged':expected['active_legacy_source_legs_covered'],
      'source_negative_balance_accounts_preserved':expected['source_negative_balance_accounts_preserved'],
      'source_balance_effect_of_cancelled_originals':0,
      'hypothetical_original_plus_inverse_net_effect':0,
      'historical_cancellation_dates_proven':0,
      'canonical_storno_od_links_generated':0,
      'canonical_import_approved':False,
      'notes':['These 335 old STORNO legs are cancelled ORIGINALS, not 335 AMB-10 inverses.',
         'Even physical pairs do not prove business kind, ownership, unique target documents or cancellation time.',
         'Source negative accounts and Q-00700 remain unchanged, with no artificial opening.',
         'No canonical rows or target workbook written. All customer data remains local in PRIVATE JSONL.']
    }
    return summary,ev,held

def render(s):
    ks=['historical_source_storno_legs','cancelled_physical_event_candidates',
      'cancelled_physical_source_legs_covered','historical_source_legs_held',
      'cancelled_physical_shape_counts','held_storno_source_leg_reason_counts',
      'source_negative_balance_accounts_preserved','source_balance_effect_of_cancelled_originals',
      'canonical_storno_od_links_generated','canonical_import_approved']
    return 'AGRIX MIG-05.7 - CANCELLED HISTORICAL EVENT BLUEPRINT / READ ONLY\n'+''.join(
        k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in ks)+\
        'NO XLSM. No migration of STORNO without historical cancellation evidence.\n'

def cli(argv=None):
    ap=argparse.ArgumentParser(description='MIG-05.7 read-only historic storno blueprint')
    for n in ('stage','expanded','seven','dryrun','out'):ap.add_argument('--'+n,type=Path,required=True)
    args=ap.parse_args(argv)
    paths=[getattr(args,n).resolve() for n in ('stage','expanded','seven','dryrun')]
    out=args.out.resolve()
    if out.exists() or any(out==p or out in p.parents or p in out.parents for p in paths):
        raise AuditError('Output exists or overlaps input; never overwrite')
    if not all(p.is_dir() for p in paths):raise AuditError('Missing required source folder')
    s,events,holds=build(*paths)
    out.mkdir(parents=True,exist_ok=False)
    try:
        for name,data in [('STORNO_PHYSICAL_CANCELLED_EVENTS_PRIVATE.jsonl',events),
                          ('STORNO_HISTORY_HELD_PRIVATE.jsonl',holds)]:
            p=out/name
            with p.open('w',encoding='utf-8',newline='\n') as dst:
                for r in data:dst.write(json.dumps(r,ensure_ascii=False,sort_keys=True)+'\n')
            os.chmod(p,0o600)
        (out/'agrix_storno_blueprint_summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (out/'agrix_storno_blueprint_summary.txt').write_text(render(s),encoding='utf-8')
    except BaseException:
        for child in out.iterdir():
            if child.is_file():child.unlink()
        out.rmdir();raise
    print(render(s));print('Upload ONLY agrix_storno_blueprint_summary.json/txt; private JSONL stays LOCAL')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
        print('MIG-05.7 STOP:',exc,file=sys.stderr);raise SystemExit(2)
