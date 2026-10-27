#!/usr/bin/env python3
"""MIG-04.4: local-only operator dossier for 98 quarantined active AMB-10 legs.

Never approves source rows, movement purpose, opening stock, or writes Excel.
No raw IDs, quantities, dates or unknown document strings in shared summaries.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, trim
import legacy_amb10_forensics_v2 as f
import legacy_amb10_residual_v2 as r
import legacy_amb10_shadow_v2 as s

VERSION = 'MIG-04.4'


def _case_class(flags):
    fset=set(filter(None,flags.split(';')))
    if 'MISSING_DOCUMENT_ID' in fset:
        return 'MISSING_DOCUMENT_ID'
    if 'DOCUMENT_LEG_STATE_MISMATCH_REVIEW' in fset:
        return 'DOCUMENT_STATE_CONFLICT'
    if 'ORPHAN_DOCUMENT_ID' in fset or 'UNRESOLVED_DOCUMENT_ID' in fset:
        return 'DOCUMENT_ID_NOT_FOUND'
    if 'DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING' in fset:
        return 'DOCUMENT_TYPE_NOT_MODELLED'
    return 'OTHER_STRUCTURAL_REVIEW'


def _doc_category(raw):
    token=f.key(raw)
    return next((t for t in sorted(f.NAMED_TYPES) if f.key(t)==token),'OTHER_OR_UNKNOWN')


def _candidate_peer(a,b):
    """Exact reciprocal OLD physical legs only; never AMB-10 business meaning."""
    if a['source_row_ordinal']==b['source_row_ordinal']:
        return False
    for col in ('DokumentTip','DokumentID','Datum','TipAmbalaze','Kolicina'):
        if f.key(a.get(col))!=f.key(b.get(col)):
            return False
    if not trim(a.get('DokumentID')):return False
    if f.key(a.get('Smer'))==f.key(b.get('Smer')):return False
    if f.key(a.get('Smer')) not in ('ulaz','izlaz') or f.key(b.get('Smer')) not in ('ulaz','izlaz'):
        return False
    # Do not call two completely unrelated accounts a reciprocal pair.
    if (f.key(a['EntitetTip']),f.key(b['EntitetTip'])) not in {
        ('kooperant','stanica'),('stanica','kooperant'),
        ('stanica','firma'),('firma','stanica'),
        ('stanica','vozac'),('vozac','stanica'),
        ('kupac','vozac'),('vozac','kupac')}:
        return False
    return True


def triage(review,accounts):
    """Return local operator tables and *aggregate-only* report metadata."""
    if len({p['source_row_ordinal'] for p in review})!=len(review):
        raise AuditError('Duplicate quarantine ordinal')
    neg={ (a['account_type'],a['account_id'],a['tip_ambalaze'])
          for a in accounts if a['legacy_negative']=='YES'}
    by_type=Counter();by_case=Counter();by_flag=Counter()
    for z in review:
        by_type[_doc_category(z['dokument_tip'])]+=1
        by_case[_case_class(z.get('forensic_flags',''))]+=1
        for flag in filter(None,z.get('forensic_flags','').split(';')):
            by_flag[flag]+=1
    return by_type,by_case,by_flag,neg


def exact_pair_candidates(review,source_index):
    """Conservative unique reciprocal candidates in quarantine (not approvals)."""
    refs={z['source_row_ordinal'] for z in review}
    if not refs.issubset(source_index):raise AuditError('Unknown source quarantine ordinal')
    groups=defaultdict(list)
    for ix in refs:
        a=source_index[ix]
        if not trim(a.get('DokumentID')):continue
        sig=tuple(f.key(a.get(t)) for t in ('DokumentTip','DokumentID','Datum','TipAmbalaze','Kolicina'))
        groups[sig].append(a)
    pairs={}; ambiguous=0
    for rr in groups.values():
        if len(rr)==2 and _candidate_peer(*rr):
            i,j=sorted(t['source_row_ordinal'] for t in rr)
            pairs[i]=j;pairs[j]=i
        elif len(rr)>1:
            ambiguous+=len(rr)
    return pairs,ambiguous


def build(stage,shadow,residual):
    res_report,accounts,review=r.build_summary(stage,shadow)
    prior_path=residual/'agrix_amb10_residual_summary.json'
    if not prior_path.is_file():raise AuditError('MIG-04.3 summary missing')
    prior=json.loads(prior_path.read_text(encoding='utf-8'))
    for k in ('format','source_sha256','schema_sha256','model_checks',
              'quarantine_forensic_flag_counts_overlapping','quarantine_document_type_counts',
              'mismatching_account_type_counts'):
        expected=VERSION if k=='format' else res_report.get(k)
        if k=='format':expected='MIG-04.3'
        if prior.get(k)!=expected:raise AuditError('MIG-04.3 baseline mismatch: '+k)
    checks=res_report['model_checks']
    if checks['model_identity_failures_all_accounts']!=0 or checks['storno_review_rows']!=0:
        raise AuditError('Unresolved model residual or historical storno quarantine')
    if len(review)!=checks['quarantined_review_rows']:
        raise AuditError('Quarantine count drift')
    baseline,src=f.analyze(stage)
    by_row={x['source_row_ordinal']:x for x in src}
    if set(x['source_row_ordinal'] for x in review)!=set(x for x,z in by_row.items() if z.get('forensic_flags')):
        raise AuditError('Quarantine no longer matches structural evidence')
    types,classes,flags,negative=triage(review,accounts)
    pairs,ambiguous=exact_pair_candidates(review,by_row)
    private=[]
    for item in sorted(review,key=lambda z:z['source_row_ordinal']):
        ix=item['source_row_ordinal'];row=by_row[ix]
        effects=s.raw_contributions(row)
        touched_negative=any(key in negative for key,value in effects)
        private.append({
          'source_row_ordinal':ix,
          'quarantine_case':_case_class(item.get('forensic_flags','')),
          'physical_peer_row_candidate':pairs.get(ix,''),
          'peer_link_status':('RECIPROCAL_ONLY_NO_BUSINESS_PROOF' if ix in pairs
                             else 'NO_UNIQUE_RECIPROCAL_PROOF'),
          'document_tip':row.get('DokumentTip',''),
          'document_id':row.get('DokumentID',''),
          'amb_id':row.get('AmbID',''),
          'datum':row.get('Datum',''),
          'tip_ambalaze':row.get('TipAmbalaze',''),
          'kolicina':row.get('Kolicina',''),
          'smer':row.get('Smer',''),
          'entitet_tip':row.get('EntitetTip',''),
          'entitet_id':row.get('EntitetID',''),
          'vozac_id':row.get('VozacID',''),
          'forensic_flags':item.get('forensic_flags',''),
          'negative_ending_account_touched':'YES' if touched_negative else 'NO',
          'operator_resolution':'UNREVIEWED',
          'independent_evidence_reference':'',
          'business_movement_purpose':'UNPROVEN',
          'ownership_of_packaging':'UNPROVEN',
          'verified_source_document':'UNPROVEN',
          'operator_note':'',
        })
    neg_details=[]
    for item in accounts:
        if item['legacy_negative']!='YES':continue
        neg_details.append({
           'account_type':item['account_type'],'account_id':item['account_id'],
           'tip_ambalaze':item['tip_ambalaze'],
           'legacy_ending_balance':item['legacy_ending_balance'],
           'shadow_ending_balance':item['shadow_ending_balance'],
           'quarantine_net':item['excluded_review_net'],
           'quarantine_source_rows':item['review_source_rows'],
           'operator_resolution':'UNREVIEWED',
           'verified_starting_stock':'UNPROVEN',
           'verified_third_party_ownership':'UNPROVEN',
           'documentary_evidence_reference':'',
           'operator_note':'',
        })
    neg_details.sort(key=lambda z:(z['account_type'],z['account_id'],z['tip_ambalaze']))
    result={
      'format':VERSION,
      'status':'PRIVATE_TRIAGE_ONLY_NO_AMB10_APPROVAL_OR_IMPORT',
      'source_sha256':f.SOURCE_SHA,
      'schema_sha256':baseline['schema_sha256'],
      'quarantined_active_rows':len(review),
      'reason_counts_overlapping':dict(sorted(flags.items())),
      'primary_triage_cases':dict(sorted(classes.items())),
      'document_type_categories':dict(sorted(types.items())),
      'exact_reciprocal_old_legs_candidate_pairs':len(pairs)//2,
      'candidate_paired_old_legs':len(pairs),
      'old_legs_without_unique_reciprocal_evidence':len(review)-len(pairs),
      'ambiguous_signature_legs':ambiguous,
      'negative_legacy_account_type_combinations':len(neg_details),
      'negative_legacy_accounts_affected_by_quarantine':sum(bool(a['quarantine_source_rows']) for a in neg_details),
      'source_storno_rows_preserved_not_replayed':checks['historical_source_rows_not_projected'],
      'model_identity_failures_all_accounts':checks['model_identity_failures_all_accounts'],
      'notes':[
        'Each quarantine row and every negative account is in PRIVATE local operator CSV; none is approved.',
        'Reciprocal old physical legs are not sufficient to prove AMB-10 kind, title or legal ownership.',
        'No business-number heuristics, automatic replacements, opening stock, foreign ownership, or AMB-10 events.',
        'Do not upload PRIVATE.csv, import them to Excel as live business tables, or paste into tblAmbalaza.',
      ],
    }
    if sum(types.values())!=len(review) or sum(classes.values())!=len(review):
        raise AuditError('Triage classification count mismatch')
    if len(neg_details)!=checks['negative_legacy_accounts']:
        raise AuditError('Negative account classification mismatch')
    return result,private,neg_details


def report_txt(x):
    lines=['AGrIX MIG-04.4 — AMB-10 QUARANTINE OPERATOR DOSSIER',
           'READ ONLY / NOT AMB-10 CONVERSION OR IMPORT',
           'Source SHA256: '+x['source_sha256'],
           str(x['quarantined_active_rows'])+' source rows are excluded from physical shadow, NOT discarded.',
           '','QUARANTINE: '+str(x['quarantined_active_rows']),
           'Reciprocal leg candidates (NOT APPROVAL): '+str(x['exact_reciprocal_old_legs_candidate_pairs']),
           'Unpaired/quarantined legs: '+str(x['old_legs_without_unique_reciprocal_evidence']),
           'Negative legacy accounts: '+str(x['negative_legacy_account_type_combinations']),
           'Negative accounts touched by quarantine: '+str(x['negative_legacy_accounts_affected_by_quarantine']),
           '','PRIMARY TRIAGE:']
    lines.extend('  '+a+': '+str(b) for a,b in x['primary_triage_cases'].items())
    lines.append('DOCUMENT TYPES:')
    lines.extend('  '+a+': '+str(b) for a,b in x['document_type_categories'].items())
    lines.append('STOP: inspect local PRIVATE.csv; no AMB-10 writes, no target XLSM.');return '\n'.join(lines)+'\n'


def cli(argv=None):
    p=argparse.ArgumentParser(description='Private local operator review for MIG-04.3 quarantine')
    for opt in ('stage','shadow','residual','out'):
        p.add_argument('--'+opt,required=True,type=Path)
    ns=p.parse_args(argv)
    inputs=[getattr(ns,k).resolve() for k in ('stage','shadow','residual')]
    dest=ns.out.resolve()
    if len(set(inputs))!=3 or any(not x.is_dir() for x in inputs):
        raise AuditError('Inputs missing or overlapping')
    if any(a in b.parents for a in inputs for b in inputs if a!=b):
        raise AuditError('Input folders nested')
    if any(dest==u or dest in u.parents or u in dest.parents for u in inputs):
        raise AuditError('Input/output folders overlap')
    if dest.exists():raise AuditError('Output folder exists; choose fresh path')
    summary,review,negative=build(*inputs)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        (dest/'agrix_amb10_quarantine_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (dest/'agrix_amb10_quarantine_summary.txt').write_text(report_txt(summary),encoding='utf-8')
        f.private_csv(dest/'amb10_quarantine_operator_PRIVATE.csv',review,[
            'source_row_ordinal','quarantine_case','physical_peer_row_candidate','peer_link_status',
            'document_tip','document_id','amb_id','datum','tip_ambalaze','kolicina','smer',
            'entitet_tip','entitet_id','vozac_id','forensic_flags',
            'negative_ending_account_touched','operator_resolution',
            'independent_evidence_reference','business_movement_purpose','ownership_of_packaging',
            'verified_source_document','operator_note'])
        f.private_csv(dest/'amb10_negative_accounts_operator_PRIVATE.csv',negative,[
            'account_type','account_id','tip_ambalaze','legacy_ending_balance',
            'shadow_ending_balance','quarantine_net','quarantine_source_rows',
            'operator_resolution','verified_starting_stock','verified_third_party_ownership',
            'documentary_evidence_reference','operator_note'])
    except BaseException:
        for file in dest.iterdir():
            if file.is_file():file.unlink()
        dest.rmdir()
        raise
    print(report_txt(summary))
    print('Send ONLY summary JSON/TXT, not PRIVATE CSVs.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,ValueError,OSError,KeyError,TypeError,json.JSONDecodeError) as err:
        print('MIG-04.4 STOP:',err,file=sys.stderr)
        raise SystemExit(2)
