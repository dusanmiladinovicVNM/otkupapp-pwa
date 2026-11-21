#!/usr/bin/env python3
"""MIG-05.8: immutable, read-only archive of the legacy cancelled AMB rows.

A legacy Stornirano=Da is a cancelled ORIGINAL (not a canonical inverse).
Preserves all rows privately, analyses previously unsupported driver pairs with
exact legacy account effects, and NEVER books or invents a cancellation time.
"""
from __future__ import annotations
import argparse,hashlib,json,os,sys
from collections import Counter
from decimal import Decimal
from pathlib import Path
from legacy_audit import AuditError,is_active,trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_storno_blueprint_v2 as prior

VERSION='MIG-05.8'


def _canon(obj):
    return json.dumps(obj,ensure_ascii=False,sort_keys=True)+'\n'


def _driver_pair_analysis(hold,by_row):
    """Propose a *physical* 2-account vector only when EXACTLY conserving.

    Never fill a third leg from SpoljniSvet or manufacture a new opening balance.
    Some historical mirror pairs with a derived driver have three nonzero account
    postings; such a pair cannot be one two-sided canonical transfer.
    """
    ids=[int(v) for v in hold['source_row_ordinals'].split(';')]
    if len(ids)!=2 or len(ids)!=len(set(ids)):
        raise AuditError('Invalid driver pair partition')
    legs=[by_row[i] for i in ids]
    if any(is_active(r) for r in legs):raise AuditError('An active row entered STORNO pair review')
    if any(not trim(r.get('VozacID')) for r in legs):
        # Previous stage holds if ANY leg has a driver: don't accidentally
        # assert both legs had a driver.
        if not any(trim(r.get('VozacID')) for r in legs):
            raise AuditError('Prior driver hold has no driver evidence')
    contribution=prior._balances([dict(r,Stornirano='') for r in legs])
    nonzero={k:v for k,v in contribution.items() if v}
    total=sum(nonzero.values(),Decimal(0))
    qty=f.dec(legs[0]['Kolicina'])
    if qty is None or qty<=0 or any(f.dec(r['Kolicina'])!=qty for r in legs):
        return {'class':'INVALID_QUANTITY','raw_account_count':len(nonzero), 'raw_net_conserved':total==0},None
    negatives=[(k,v) for k,v in nonzero.items() if v<0]
    positives=[(k,v) for k,v in nonzero.items() if v>0]
    two_party=(len(nonzero)==2 and len(negatives)==1 and len(positives)==1
               and -negatives[0][1]==qty and positives[0][1]==qty and total==0)
    if not two_party:
        reason='NONCONSERVING_DERIVED_DRIVER_VECTOR' if total!=0 else 'MULTI_ACCOUNT_OR_NONUNIT_VECTOR'
        return {'class':reason,'raw_account_count':len(nonzero),'raw_net_conserved':total==0},None
    neg=negatives[0][0];pos=positives[0][0]
    ev={'source_row_ordinals':hold['source_row_ordinals'],
        'OdNalogTip':neg[0],'OdNalogID':neg[1],
        'NaNalogTip':pos[0],'NaNalogID':pos[1],
        'TipAmbalaze':neg[2], 'Kolicina':trim(legs[0]['Kolicina']),
        'evidence':'EXACT_CANCELLED_ORIGINAL_LEGACY_VECTOR_ONLY',
        'document_identity':'UNPROVEN','movement_kind':'UNPROVEN',
        'cancellation_date':'UNPROVEN','StornoOd':None,
        'import_approved':False}
    if prior._event_delta(ev)!=nonzero:raise AuditError('Driver pair vector equality failure')
    return {'class':'TWO_ACCOUNT_EXACT_VECTOR_ONLY','raw_account_count':2,'raw_net_conserved':True},ev


def build(stage,expanded,seven,dryrun,blueprint):
    previous,old_events,old_holds=prior.build(stage,expanded,seven,dryrun)
    stored=blueprint/'agrix_storno_blueprint_summary.json'
    if not stored.is_file() or json.loads(stored.read_text(encoding='utf-8'))!=previous:
        raise AuditError('Stored MIG-05.7 blueprint summary is missing or changed')
    for name,expected in (('STORNO_PHYSICAL_CANCELLED_EVENTS_PRIVATE.jsonl',old_events),
                          ('STORNO_HISTORY_HELD_PRIVATE.jsonl',old_holds)):
        prior._read_verified_private(blueprint/name,expected)
    manifest,tables,indices=f.inputs(stage)
    raw=tables['UNCONVERTED__tblAmbalaza']
    if len(raw)!=7815:raise AuditError('Source population drift')
    cancelled=[]
    for n,r in enumerate(raw,1):
        if not is_active(r):
            record={k:r[k] for k in f.FIELDS}
            record['source_row_ordinal']=n
            record['archive_class']='CANCELLED_LEGACY_ORIGINAL_NOT_CANONICAL_EVENT'
            cancelled.append(record)
    if len(cancelled)!=previous['historical_source_storno_legs']:
        raise AuditError('Cancelled original coverage changed')
    by_row={r['source_row_ordinal']:r for r in cancelled}
    existing_ids={int(i) for e in old_events for i in e['source_row_ordinals'].split(';')}
    held_ids={int(i) for h in old_holds for i in h['source_row_ordinals'].split(';')}
    if existing_ids&held_ids or existing_ids|held_ids!=set(by_row):
        raise AuditError('Previous STORNO history partition not exhaustive')
    recovered=[];remaining=[];classified=[];counts=Counter();driver_groups=0
    for h in old_holds:
        why=h['hold_reason']
        if why=='PAIRED_WITH_DRIVER_NOT_SUPPORTED':
            driver_groups+=1
            verdict,event=_driver_pair_analysis(h,by_row)
            classification=verdict['class']
            case={'source_row_ordinals':h['source_row_ordinals'],
                  'legacy_amb_ids':h['legacy_amb_ids'],'legacy_document_type':h['legacy_DokumentTip'],
                  'legacy_document_label':h['legacy_DokumentID'],
                  'classification':classification,
                  'raw_account_count':verdict['raw_account_count'],
                  'raw_net_conserved':verdict['raw_net_conserved'],
                  'import_approved':False}
            classified.append(case)
            if event:
                recovered.append(event)
            else:
                remaining.append(dict(h,new_review_class=classification))
            counts[classification]+=len(h['source_row_ordinals'].split(';'))
        else:
            remaining.append(dict(h,new_review_class='UNRESOLVED_LABEL_COLLISION'))
            counts['UNRESOLVED_LABEL_COLLISION']+=len(h['source_row_ordinals'].split(';'))
    recovered_ids={int(i) for e in recovered for i in e['source_row_ordinals'].split(';')}
    remaining_ids={int(i) for h in remaining for i in h['source_row_ordinals'].split(';')}
    if existing_ids&recovered_ids or (existing_ids|recovered_ids)&remaining_ids or existing_ids|recovered_ids|remaining_ids!=set(by_row):
        raise AuditError('STORNO history partition after reclassification failed')
    if sum(counts.values())!=len(held_ids):raise AuditError('New STORNO classification incomplete')
    archive_digest=hashlib.sha256(''.join(_canon(r) for r in cancelled).encode('utf-8')).hexdigest()
    result={
      'format':VERSION,
      'status':'HISTORICAL_CANCELLED_ORIGINAL_ARCHIVE_AND_VECTOR_REVIEW_ONLY',
      'source_sha256':f.SOURCE_SHA,
      'schema_sha256':manifest['schema_sha256'],
      'upstream_blueprint_format':previous['format'],
      'historical_storno_rows_archived':len(cancelled),
      'legacy_archive_sha256':archive_digest,
      'previously_proven_cancelled_physical_events':len(old_events),
      'previously_proven_source_rows':len(existing_ids),
      'driver_pair_groups_reexamined':driver_groups,
      'driver_pair_rows_reexamined':sum(len(x['source_row_ordinals'].split(';')) for x in classified),
      'additional_exact_two_account_vectors':len(recovered),
      'additional_source_rows_with_exact_two_account_vector':len(recovered_ids),
      'historical_rows_still_physically_unresolved':len(remaining_ids),
      'historical_unresolved_class_counts':dict(sorted(counts.items())),
      'historical_originals_accounting_balance_effect':0,
      'source_negative_balance_accounts_preserved':previous['source_negative_balance_accounts_preserved'],
      'known_active_hold_q00700_unchanged':True,
      'canonical_storno_od_links_generated':0,
      'historical_cancellation_dates_proven':0,
      'import_approved':False,
      'notes':[
        'All historical cancelled originals are preserved in a private integrity-hashed archive.',
        'A derived driver account can make an old pair a non-conserving vector; never invent offset or opening.',
        'Archiving STORNO entries does not replay them in canonical AMB-10; it preserves historical evidence only.',
        'Source negative balances are accepted as-is and not repaired.',
        'No customer IDs, quantities or business numbers in shareable summary; no source or target workbook changed.'
      ]
    }
    return result,cancelled,classified,recovered,remaining


def render(s):
    ks=('historical_storno_rows_archived','previously_proven_cancelled_physical_events',
      'previously_proven_source_rows','driver_pair_groups_reexamined','driver_pair_rows_reexamined',
      'additional_exact_two_account_vectors','additional_source_rows_with_exact_two_account_vector',
      'historical_rows_still_physically_unresolved','historical_unresolved_class_counts',
      'historical_originals_accounting_balance_effect','source_negative_balance_accounts_preserved',
      'canonical_storno_od_links_generated','import_approved','legacy_archive_sha256')
    return 'AGRIX MIG-05.8 - HISTORICAL CANCELLED SOURCE ARCHIVE / NO REPLAY\n'+''.join(
        k+': '+json.dumps(s[k],ensure_ascii=False,sort_keys=True)+'\n' for k in ks)+\
        'NO AMB-10, NO TARGET XLSM. PRIVATE archive and evidence remain LOCAL.\n'


def cli(argv=None):
    p=argparse.ArgumentParser(description='MIG-05.8 cancelled original archive, read-only')
    for a in ('stage','expanded','seven','dryrun','blueprint','out'):
        p.add_argument('--'+a,required=True,type=Path)
    ns=p.parse_args(argv)
    inputs=[getattr(ns,a).resolve() for a in ('stage','expanded','seven','dryrun','blueprint')]
    out=ns.out.resolve()
    if out.exists() or any(out==q or out in q.parents or q in out.parents for q in inputs):
        raise AuditError('Output exists or overlaps input; never overwrite')
    if not all(q.is_dir() for q in inputs):raise AuditError('Source directory missing')
    result,archive,driver_cases,recovered,remaining=build(*inputs)
    out.mkdir(parents=True,exist_ok=False)
    try:
        files={
            'STORNO_ARCHIVE_ALL_CANCELLED_ORIGINALS_PRIVATE.jsonl':archive,
            'STORNO_DRIVER_PAIR_VECTOR_REVIEW_PRIVATE.jsonl':driver_cases,
            'STORNO_ADDITIONAL_EXACT_VECTORS_PRIVATE.jsonl':recovered,
            'STORNO_HISTORY_STILL_HELD_PRIVATE.jsonl':remaining,
        }
        for name,records in files.items():
            path=out/name
            with path.open('w',encoding='utf-8',newline='\n') as f_out:
                for row in records:f_out.write(_canon(row))
            os.chmod(path,0o600)
        if hashlib.sha256((out/'STORNO_ARCHIVE_ALL_CANCELLED_ORIGINALS_PRIVATE.jsonl').read_bytes()).hexdigest()!=result['legacy_archive_sha256']:
            raise AuditError('Archive checksum mismatch after write')
        (out/'agrix_storno_archive_summary.json').write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_storno_archive_summary.txt').write_text(render(result),encoding='utf-8')
    except BaseException:
        for child in out.iterdir():
            if child.is_file():child.unlink()
        out.rmdir();raise
    print(render(result));print('Send only agrix_storno_archive_summary.json and .txt; PRIVATE files stay local')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as exc:
        print('MIG-05.8 STOP:',exc,file=sys.stderr);raise SystemExit(2)
