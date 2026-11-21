#!/usr/bin/env python3
"""AGRIx MIG-19.3: legacy -> AMB-10 historical saldo parity; READ ONLY.

Goal: EXACT agreement with recorded legacy balances, not judgement about stock plausibility.
Frozen domain references:
 Legacy: otkupapp-pwa@4d3e2caaa6cec2994b506d19e090bb263afcb5cf,
   modAmbalaza.GetAmbalazeStanje: active Smer Ulaz (+), Izlaz (-).
 Canonical: otkupapp-pwa@66f7897fa5231786f7d3d3dd3e566922c8e80ed0,
   modAmbalaza.AmbSaldoNaloga: Na (+), Od (-), including storno inverse.
 Migration-only ULAZ_TUDJE_AMBALAZE provenance is not legacy physical movement;
   keep canonical physical saldo intact, but remove its signed contribution in
   legacy-equivalent projection.

Source: MigrationStage_v2/UNCONVERTED__tblAmbalaza.jsonl (7815 rows).
Target: MIG18 QA .xlsm (4811 rows). No source or target writes.

Outputs: reconciliation CSV files and shareable summary.json (counts only).
Does NOT establish event-level completeness just because balances match.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import csv
import hashlib
import json
from pathlib import Path
import re
import sys

# Reuses the audited and already installed OOXML read-only reader from MIG-19.1.
# The ZIP package includes it as a local dependency; no third-party libraries.
from agrix_mig19_malina_readonly_audit import XlsmReadOnly, sha256, ddate, dnum, norm, key, write_csv

STAGE_SHA = '8a3d1c370b95f08e602216d124af12d67febb7ae3260360789bad7abca138bb5'
SOURCE_XLSM_SHA = 'a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'
PARTIES = {'KOOPERANT', 'STANICA'}
STABLE_DOCS = {'OTKUP', 'OTPREMNICA', 'PRIJEMNICA'}
DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')


def norm_type(v):
    t = re.sub(r'\s+', '', key(v))
    return re.sub(r'(?<=\d)\.(?=\d)', ',', t)


def number(v, where, positive=False):
    x = dnum(v)
    if x is None or not x.is_finite() or (positive and x <= 0):
        raise ValueError(f'Invalid numeric quantity at {where}: {v!r}')
    return x


def date(v, where):
    x = ddate(v)
    if not DATE_RE.fullmatch(x):
        raise ValueError(f'Invalid date at {where}: {v!r} -> {x!r}')
    return x


def void(v):
    # This is the legacy VBA ExcludeStornirano convention: Stornirano = 'Da'.
    return key(v) == 'DA'


def dec(v):
    return format(v.normalize(), 'f')


def entry(acct, amb, day, qty, document, docid):
    return ((acct[0], acct[1], amb), day, qty,
            (acct[0], acct[1], amb, day, document, docid))


def ledger_views(old_rows, new_rows, cutoff='2026-08-31'):
    """Pure math: old signed, canonical physical, canonical legacy-equivalent.

    Returns each ledger daily deltas and stable source-document deltas;
    no account type collapsing, no minus banning, no made-up stock.
    """
    old, physical, projected, cover = (defaultdict(Decimal) for _ in range(4))
    olddoc, newdoc = defaultdict(Decimal), defaultdict(Decimal)
    diag = Counter()
    all_orig_ids = Counter()
    rows_after_cutoff = Counter()

    for i, r in enumerate(old_rows, 1):
        where = f'legacy row {i} AmbID={r.get("AmbID", "")}'
        all_orig_ids[key(r.get('AmbID'))] += 1
        if void(r.get('Stornirano')):
            diag['legacy_stornirano'] += 1
            continue
        q = number(r.get('Kolicina'), where)
        day = date(r.get('Datum'), where)
        if day > cutoff:
            rows_after_cutoff['legacy'] += 1
            continue
        diag['legacy_active'] += 1
        direct = key(r.get('Smer'))
        if direct not in {'ULAZ', 'IZLAZ'}:
            raise ValueError(f'Unknown old Smer {direct!r} at {where}')
        acct = (key(r.get('EntitetTip')), key(r.get('EntitetID')))
        if not acct[0] or (acct[0] in {'KOOPERANT', 'STANICA', 'VOZAC', 'KUPAC'} and not acct[1]):
            raise ValueError(f'Empty legacy account at {where}')
        typ = norm_type(r.get('TipAmbalaze'))
        if not typ: raise ValueError('Missing type '+where)
        signed = q if direct == 'ULAZ' else -q
        dkey=(acct[0],acct[1],typ,day)
        old[dkey]+=signed
        doctype=key(r.get('DokumentTIP'))
        docid=key(r.get('DokumentID'))
        if doctype in STABLE_DOCS and docid and acct[0] in PARTIES:
            olddoc[(acct[0],acct[1],typ,day,doctype,docid)] += signed

    old_dupes=sum(n-1 for k,n in all_orig_ids.items() if n>1 and k)
    diag['legacy_duplicate_AmbID'] = old_dupes
    diag['legacy_empty_AmbID'] = all_orig_ids.get('',0)

    all_target_ids=Counter()
    for i, r in enumerate(new_rows, 1):
        where = f'canonical row {i} AmbID={r.get("AmbID", "")}'
        all_target_ids[key(r.get('AmbID'))]+=1
        q=number(r.get('Kolicina'),where,positive=True)
        day=date(r.get('Datum'),where)
        if day>cutoff:
            rows_after_cutoff['canonical']+=1
            continue
        typ=norm_type(r.get('TipAmbalaze'))
        if not typ:raise ValueError('Missing type '+where)
        from_acct=(key(r.get('OdNalogTip')),key(r.get('OdNalogID')))
        to_acct=(key(r.get('NaNalogTip')),key(r.get('NaNalogID')))
        if not from_acct[0] or not to_acct[0]:
            raise ValueError('Empty source/destination account type '+where)
        if from_acct==to_acct:
            raise ValueError('Same source/destination account '+where)
        vk=key(r.get('VrstaKretanja'))
        for acct,amt in ((from_acct,-q),(to_acct,q)):
            physical[(acct[0],acct[1],typ,day)] += amt
            if vk != 'ULAZ_TUDJE_AMBALAZE':
                projected[(acct[0],acct[1],typ,day)] += amt
            else:
                cover[(acct[0],acct[1],typ,day)] += amt
            dt=key(r.get('DokumentTIP')); di=key(r.get('DokumentID'))
            if vk != 'ULAZ_TUDJE_AMBALAZE' and dt in STABLE_DOCS and di and acct[0] in PARTIES:
                newdoc[(acct[0],acct[1],typ,day,dt,di)] += amt
        diag['canonical_rows']+=1
        if vk=='ULAZ_TUDJE_AMBALAZE':diag['cover_rows']+=1
        if key(r.get('StornoOd')):diag['storno_inverse_rows']+=1

    diag['canonical_duplicate_AmbID']=sum(n-1 for k,n in all_target_ids.items() if n>1 and k)
    diag['canonical_empty_AmbID']=all_target_ids.get('',0)
    return old,physical,projected,cover,olddoc,newdoc,diag,rows_after_cutoff


def aggregate(buckets, include=PARTIES):
    t=defaultdict(Decimal)
    for (ac,ident,typ,day),quantity in buckets.items():
        if ac in include:
            t[(ac,ident,typ)]+=quantity
    return t


def reconcile(old,physical,projected,cover,olddoc,newdoc):
    totals=[]
    oldt=aggregate(old); newt=aggregate(projected); physt=aggregate(physical); covert=aggregate(cover)
    for k in sorted(set(oldt)|set(newt)|set(physt)):
        legacy=oldt[k]; migrated=newt[k]; difference=migrated-legacy
        totals.append(dict(NalogTip=k[0],NalogID=k[1],TipAmbalaze=k[2],
                           StariSaldo=dec(legacy),NoviIstorijskiSaldo=dec(migrated),
                           NoviFizickiSaldo=dec(physt[k]),MigracionoPokrice=dec(covert[k]),
                           Razlika=dec(difference),Status='PASS' if difference==0 else 'MISMATCH'))

    # All distinct observation dates per account/type: compare closing totals,
    # never assume intraday order is meaningful in the legacy workbook.
    byacct=defaultdict(set)
    for ac,i,typ,day in set(old)|set(projected):
        if ac in PARTIES:byacct[(ac,i,typ)].add(day)
    days=[]
    for (ac,i,typ), dates in sorted(byacct.items()):
        a=b=Decimal(0)
        for day in sorted(dates):
            a+=old[(ac,i,typ,day)];b+=projected[(ac,i,typ,day)]
            days.append(dict(NalogTip=ac,NalogID=i,TipAmbalaze=typ,Datum=day,
                             StariZakljucniSaldo=dec(a),NoviZakljucniSaldo=dec(b),
                             Razlika=dec(b-a),Status='PASS' if a==b else 'MISMATCH'))

    documents=[]
    for k in sorted(set(olddoc)|set(newdoc)):
        a,b=olddoc[k],newdoc[k]
        documents.append(dict(NalogTip=k[0],NalogID=k[1],TipAmbalaze=k[2],Datum=k[3],
                              DokumentTIP=k[4],DokumentID=k[5],
                              StariPotpisaniEfekat=dec(a),NoviPotpisaniEfekat=dec(b),
                              Razlika=dec(b-a),Status='PASS' if a==b else 'MISMATCH'))
    return totals,days,documents


def stage_rows(path):
    rows=[]
    with path.open('r',encoding='utf-8-sig') as f:
        for n,line in enumerate(f,1):
            if not line.strip():continue
            r=json.loads(line)
            if not isinstance(r,dict):raise ValueError(f'Stage JSONL item {n} is not an object')
            rows.append(r)
    required={'AmbID','Datum','TipAmbalaze','Kolicina','Smer','EntitetTip','EntitetID','Stornirano'}
    if not rows:raise ValueError('Empty staging file')
    missing=required-set(rows[0])
    if missing:raise ValueError('Missing stage fields: '+', '.join(sorted(missing)))
    return rows


def source_semantic(r):
    # Stage and XLSM may store dates/numbers in different but equivalent forms.
    cols=['AmbID','Smer','EntitetTip','EntitetID','VozacID','DokumentTIP','DokumentID','Stornirano']
    base=tuple(key(r.get(c)) for c in cols)
    return base+(date(r.get('Datum'),'original/stage date'),norm_type(r.get('TipAmbalaze')),
                 dec(number(r.get('Kolicina'),'original/stage quantity')))


def verify_source_against_stage(source_path, stage):
    z=XlsmReadOnly(source_path)
    try:
        original=z.read_table('tblAmbalaza',{'AmbID','Datum','TipAmbalaze','Kolicina','Smer','EntitetTip','EntitetID','Stornirano'})
    finally:z.close()
    a=Counter(map(source_semantic,original)); b=Counter(map(source_semantic,stage))
    return dict(original_rows=len(original),stage_rows=len(stage),
                different_source_entries=sum((a-b).values()),
                different_stage_entries=sum((b-a).values()),
                rows_identical_in_legacy_fields=a==b)


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage',type=Path,required=True,help='MigrationStage_v2 directory')
    p.add_argument('--workbook',type=Path,required=True,help='Migrated QA .xlsm')
    p.add_argument('--out-dir',type=Path,required=True,help='Output directory, never a workbook directory')
    p.add_argument('--source-workbook',type=Path,help='Optional immutable original AgriX_Avgust_2026_v2.xlsm')
    p.add_argument('--cutoff',default='2026-08-31',help='Inclusive historical end date, ISO YYYY-MM-DD')
    args=p.parse_args(argv)
    if not DATE_RE.fullmatch(args.cutoff):p.error('Cutoff must be YYYY-MM-DD')
    src=args.stage/'UNCONVERTED__tblAmbalaza.jsonl'
    for f in (src,args.workbook):
        if not f.is_file():p.error('Input missing: '+str(f))
    if args.source_workbook and not args.source_workbook.is_file():p.error('Source workbook missing: '+str(args.source_workbook))
    out=args.out_dir.resolve()
    protected=[src.resolve(),args.workbook.resolve()]
    if args.source_workbook:protected.append(args.source_workbook.resolve())
    if out in protected or any(out==x.parent for x in protected):p.error('Output directory must be separate from inputs')
    hashes_before={str(x):sha256(x) for x in protected}
    stage_sha=sha256(src)
    source_sha=sha256(args.source_workbook) if args.source_workbook else None
    old=stage_rows(src)
    reader=XlsmReadOnly(args.workbook)
    try:
        new=reader.read_table('tblAmbalaza',{'AmbID','Datum','TipAmbalaze','Kolicina',
                        'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd','DokumentTIP','DokumentID'})
    finally:reader.close()
    source_reconciliation=verify_source_against_stage(args.source_workbook,old) if args.source_workbook else None
    x=ledger_views(old,new,args.cutoff)
    old_d,phys_d,proj_d,cover_d,old_docs,new_docs,diag,after=x
    totals,days,docs=reconcile(old_d,phys_d,proj_d,cover_d,old_docs,new_docs)
    tally={
        'account_type_pairs_compared':len(totals),
        'account_type_pairs_mismatch':sum(r['Status']!='PASS' for r in totals),
        'station_type_pairs_mismatch':sum(r['Status']!='PASS' and r['NalogTip']=='STANICA' for r in totals),
        'cooperant_type_pairs_mismatch':sum(r['Status']!='PASS' and r['NalogTip']=='KOOPERANT' for r in totals),
        'daily_closes_compared':len(days),
        'daily_closes_mismatch':sum(r['Status']!='PASS' for r in days),
        'stable_document_account_slices_compared':len(docs),
        'stable_document_account_slices_mismatch':sum(r['Status']!='PASS' for r in docs),
        'legacy_rows':len(old),'canonical_rows':len(new),
        **dict(diag),
        'after_cutoff_legacy':after['legacy'],
        'after_cutoff_canonical':after['canonical'],
    }
    input_integrity = (stage_sha==STAGE_SHA and (not args.source_workbook or source_sha==SOURCE_XLSM_SHA)
                       and (not source_reconciliation or source_reconciliation['rows_identical_in_legacy_fields']))
    # A mismatch in any parity proof is HOLD for remediation, never interpreted as
    # evidence a stock amount is economically impossible.
    parity_ok=(tally['account_type_pairs_mismatch']==0 and tally['daily_closes_mismatch']==0
               and tally['stable_document_account_slices_mismatch']==0)
    data_integrity_ok=(tally['legacy_duplicate_AmbID']==0 and tally['legacy_empty_AmbID']==0
                      and tally['canonical_duplicate_AmbID']==0 and tally['canonical_empty_AmbID']==0
                      and tally['legacy_rows']==7815 and tally['legacy_stornirano']==335 and
                      tally['canonical_rows']==4811 and
                      tally['after_cutoff_legacy']==0 and tally['after_cutoff_canonical']==0)
    hashes_after={str(x):sha256(x) for x in protected}
    unchanged=hashes_before==hashes_after
    if not unchanged:raise RuntimeError('Input file hash changed during READ ONLY inspection')
    out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG193_SALDO_SVI.csv',totals,
              ['NalogTip','NalogID','TipAmbalaze','StariSaldo','NoviIstorijskiSaldo',
               'NoviFizickiSaldo','MigracionoPokrice','Razlika','Status'])
    write_csv(out/'MIG193_SALDO_RAZLIKE.csv',[r for r in totals if r['Status']!='PASS'],
              ['NalogTip','NalogID','TipAmbalaze','StariSaldo','NoviIstorijskiSaldo',
               'NoviFizickiSaldo','MigracionoPokrice','Razlika','Status'])
    write_csv(out/'MIG193_DNEVNA_ZATVARANJA_RAZLIKE.csv',[r for r in days if r['Status']!='PASS'],
              ['NalogTip','NalogID','TipAmbalaze','Datum','StariZakljucniSaldo',
               'NoviZakljucniSaldo','Razlika','Status'])
    write_csv(out/'MIG193_DOKUMENT_RAZLIKE.csv',[r for r in docs if r['Status']!='PASS'],
              ['NalogTip','NalogID','TipAmbalaze','Datum','DokumentTIP','DokumentID',
               'StariPotpisaniEfekat','NoviPotpisaniEfekat','Razlika','Status'])
    result={
        'format':'AGRIx-MIG19.3-FAITHFUL-HISTORIC-PARITY-v1',
        'status':'SALDO_AND_HISTORICAL_PARITY_PASS' if parity_ok and input_integrity and data_integrity_ok
                   else 'HOLD_ONE_OR_MORE_GATES_NOT_MET',
        'complete_migration_100_percent_proven':False,
        'reason_completeness_not_proven':'Balances + dated + stable document account parity cannot by themselves prove every source row maps to a specific canonical event; event-level provenance and archived storno correspondence are separate requirements.',
        'correctness_target':'Exact historical signed balance, not nonnegative physical stock',
        'source_stage_sha256':stage_sha,
        'source_stage_matches_pinned_inventory':stage_sha==STAGE_SHA,
        'source_original_sha256_if_supplied':source_sha,
        'source_original_matches_known_sha_if_supplied':None if source_sha is None else source_sha==SOURCE_XLSM_SHA,
        'source_workbook_vs_staging_if_supplied':source_reconciliation,
        'target_sha256':hashes_before[str(args.workbook.resolve())],
        'input_hashes_unchanged':unchanged,
        'cutoff_inclusive':args.cutoff,
        'counts':tally,
        'gate_account_and_date_parity':parity_ok,
        'gate_frozen_source_integrity':input_integrity,
        'gate_preflight_counts_and_ids':data_integrity_ok,
        'calculation':'Legacy Smer=Ulaz(+)/Izlaz(-), Stornirano=Da excluded. Canonical Na(+)/Od(-) with original+contra; legacy projection omits ULAZ_TUDJE_AMBALAZE. Physical canonical balance is reported separately.',
        'remaining_event_evidence':'Historical old-row-to-new-event mapping, six OMuUlaz cases and archived 335 cancelled source rows are not proved by this saldo script.'
    }
    (out/'MIG193_SUMMARY.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'status':result['status'],'counts':tally,
                      'stage_sha_matches_inventory':stage_sha==STAGE_SHA,
                      'unchanged':unchanged,'out_dir':str(out),
                      'event_level_completeness_proven':False},ensure_ascii=False,indent=2))
    return 0 if parity_ok and input_integrity and data_integrity_ok else 2


if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:
        print(f'MIG19_3_ERROR: {type(e).__name__}: {e}',file=sys.stderr)
        sys.exit(3)
