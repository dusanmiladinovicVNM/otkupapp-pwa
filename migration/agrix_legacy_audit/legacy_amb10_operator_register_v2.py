#!/usr/bin/env python3
"""MIG-04.5: audited operator evidence register, NEVER AMB-10 import authorization.

Operates on PRIVATE MIG-04.4 CSV files on operator's PC; only aggregate reports
are appropriate for sharing.  Group exactly reciprocal physical legs to avoid
reviewing both halves independently.  Default: no approvals or business inferences.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from pathlib import Path
import os
import sys
from collections import Counter

from legacy_audit import AuditError

VERSION = 'MIG-04.5'
Q_FIELDS = ('case_id','source_row_ordinals','physical_case','primary_causes',
            'document_categories','operator_resolution','direction_verified',
            'movement_kind','packaging_ownership','source_document_basis',
            'independent_evidence_reference','operator_initials','operator_note')
N_FIELDS = ('case_id','account_type','account_id','tip_ambalaze',
            'operator_resolution','opening_state_evidence','packaging_ownership',
            'independent_evidence_reference','operator_initials','operator_note')
REVIEW_FIELDS = ('operator_resolution','direction_verified','movement_kind',
                 'packaging_ownership','source_document_basis',
                 'independent_evidence_reference','operator_initials','operator_note')
N_REVIEW_FIELDS = ('operator_resolution','opening_state_evidence','packaging_ownership',
                   'independent_evidence_reference','operator_initials','operator_note')
KINDS = {'AMBALAZA_UZ_ROBU','IZDATA_PRAZNA','POVRAT_PRAZNE','PRENOS_INTERNO',
         'VRACANJE_TUDJE_AMBALAZE','NABAVKA','OTPIS'}
OWNERSHIP = {'OWN_FIRM','OWN_PARTNER','MIXED_WITH_EVIDENCE'}
SOURCE_BASIS = {'EXACT_ID_VERIFIED','INDEPENDENT_SOURCE_DOCUMENT',
                'DOCUMENTLESS_EVENT_CONFIRMED_WITH_EVIDENCE'}

def read_csv(path:Path, expected):
    if not path.is_file():raise AuditError('Required local PRIVATE CSV missing: '+path.name)
    with path.open('r',newline='',encoding='utf-8-sig') as fd:
        reader=csv.DictReader(fd)
        if not reader.fieldnames or len(reader.fieldnames)!=len(set(reader.fieldnames)):
            raise AuditError('Missing or duplicate headers in '+path.name)
        if not set(expected).issubset(set(reader.fieldnames)):
            raise AuditError('Missing required PRIVATE CSV columns: '+path.name)
        out=list(reader)
    if any(None in x for x in out):raise AuditError('Ragged CSV in '+path.name)
    return out

def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()

def q_case_sort(item):return int(item['source_row_ordinal'])

def dossier_load(folder):
    summary_path=folder/'agrix_amb10_quarantine_summary.json'
    if not summary_path.is_file():raise AuditError('Missing MIG-04.4 JSON summary')
    summary=json.loads(summary_path.read_text(encoding='utf-8'))
    if summary.get('format')!='MIG-04.4' or summary.get('status')!='PRIVATE_TRIAGE_ONLY_NO_AMB10_APPROVAL_OR_IMPORT':
        raise AuditError('Not a supported MIG-04.4 dossier')
    p1=folder/'amb10_quarantine_operator_PRIVATE.csv'
    p2=folder/'amb10_negative_accounts_operator_PRIVATE.csv'
    q=read_csv(p1,('source_row_ordinal','physical_peer_row_candidate','quarantine_case','document_tip'))
    n=read_csv(p2,('account_type','account_id','tip_ambalaze'))
    if len(q)!=summary['quarantined_active_rows'] or len(n)!=summary['negative_legacy_account_type_combinations']:
        raise AuditError('Dossier row counts disagree with baseline')
    rows={}
    for z in q:
        try:ix=int(z['source_row_ordinal'])
        except (ValueError,TypeError):raise AuditError('Bad source ordinal')
        if ix<=0 or ix in rows:raise AuditError('Duplicate/nonpositive source ordinal')
        rows[ix]=z
    pairs=set(); visited=set();cases=[]
    for ix in sorted(rows):
        if ix in visited:continue
        peer=rows[ix].get('physical_peer_row_candidate','').strip()
        group=(ix,)
        if peer:
            try: j=int(peer)
            except ValueError:raise AuditError('Invalid reciprocal peer ordinal')
            if j==ix or j not in rows or rows[j].get('physical_peer_row_candidate','').strip()!=str(ix):
                raise AuditError('Nonreciprocal physical peer evidence')
            if j in visited:raise AuditError('Peer reused')
            group=tuple(sorted((ix,j)))
            pairs.add(group)
        visited.update(group)
        categories=sorted({z.get('document_tip','').strip() for z in (rows[x] for x in group)})
        # Only known public categories; free-form client content is never exported to summaries.
        causes=sorted({z.get('quarantine_case','').strip() for z in (rows[x] for x in group)})
        cases.append({
          'case_id':f'Q-{group[0]:05d}',
          'source_row_ordinals':'+'.join(map(str,group)),
          'physical_case':'RECIPROCAL_TWO_LEGS' if len(group)==2 else 'SINGLE_UNPAIRED_LEG',
          'primary_causes':';'.join(causes),
          'document_categories':';'.join(categories),
          'operator_resolution':'PENDING',
          'direction_verified':'UNPROVEN',
          'movement_kind':'UNPROVEN',
          'packaging_ownership':'UNPROVEN',
          'source_document_basis':'UNPROVEN',
          'independent_evidence_reference':'',
          'operator_initials':'',
          'operator_note':'',
        })
    if len(pairs)!=summary['exact_reciprocal_old_legs_candidate_pairs']:
        raise AuditError('Reciprocal pair counts disagree with MIG-04.4')
    if len(cases)!=len(q)-len(pairs):raise AuditError('Physical case partition broken')
    if len(q)-2*len(pairs)!=summary['old_legs_without_unique_reciprocal_evidence']:
        raise AuditError('Singleton counts disagree with MIG-04.4')
    neg=[]
    for j,z in enumerate(n,1):
        neg.append({
          'case_id':f'N-{j:03d}',
          'account_type':z['account_type'],
          'account_id':z['account_id'],
          'tip_ambalaze':z['tip_ambalaze'],
          'operator_resolution':'PENDING',
          'opening_state_evidence':'UNPROVEN',
          'packaging_ownership':'UNPROVEN',
          'independent_evidence_reference':'',
          'operator_initials':'',
          'operator_note':'',
        })
    if len({(r['account_type'],r['account_id'],r['tip_ambalaze']) for r in neg})!=len(neg):
        raise AuditError('Duplicate negative account/type reference')
    return summary,cases,neg,{'q_csv_sha256':sha(p1),'n_csv_sha256':sha(p2)}

def dump_csv(path,rows,fields):
    with path.open('w',encoding='utf-8-sig',newline='') as fd:
        wr=csv.DictWriter(fd,fieldnames=fields,extrasaction='raise');wr.writeheader();wr.writerows(rows)
    try:os.chmod(path,0o600)
    except OSError:pass

def safe_summary(summary,cases,neg,counts,neg_counts,mode,issues):
    return {
      'format':VERSION,
      'status':'OPERATOR_EVIDENCE_REGISTER_ONLY_NOT_AMB10_APPROVAL',
      'mode':mode,
      'source_sha256':summary['source_sha256'],
      'schema_sha256':summary['schema_sha256'],
      'quarantine_source_rows':summary['quarantined_active_rows'],
      'physical_reciprocal_pairs':summary['exact_reciprocal_old_legs_candidate_pairs'],
      'physical_singleton_legs':summary['old_legs_without_unique_reciprocal_evidence'],
      'operator_physical_cases':len(cases),
      'negative_balance_account_cases':len(neg),
      'movement_review_status_counts':dict(sorted(counts.items())),
      'negative_review_status_counts':dict(sorted(neg_counts.items())),
      'validation_issue_counts':dict(sorted(Counter(issues).items())),
      'ready_for_import':False,
      'note':'Decisions evidence does not create AMB-10 events or prove interim stocks, ownership, or historical storno.'
    }

def report_txt(x):
    lines=['AGRIX MIG-04.5 - LOCAL AMB-10 OPERATOR DECISION REGISTER',
           'Mode: '+x['mode']+' / NO CANONICAL AMB-10 AND NO XLSM',
           'Source SHA256: '+x['source_sha256'],
           'Physical cases: '+str(x['operator_physical_cases']),
           'Negative account cases: '+str(x['negative_balance_account_cases']),
           'Movement statuses: '+json.dumps(x['movement_review_status_counts'],sort_keys=True),
           'Negative statuses: '+json.dumps(x['negative_review_status_counts'],sort_keys=True),
           'Validation issues: '+json.dumps(x['validation_issue_counts'],sort_keys=True),
           'NO IMPORT APPROVAL. Never upload PRIVATE CSVs.']
    return '\n'.join(lines)+'\n'

def build_results(folder,mode,decisions=None):
    baseline,cases,neg,src_hashes=dossier_load(folder)
    issues=[]
    if mode=='prepare':
        q_dec=cases;n_dec=neg
    elif mode=='check':
        if not decisions or not decisions.is_dir():raise AuditError('Missing decision folder')
        manifest_path=decisions/'agrix_amb10_operator_manifest_PRIVATE.json'
        if not manifest_path.is_file():raise AuditError('Missing operator manifest')
        manifest=json.loads(manifest_path.read_text(encoding='utf-8'))
        for field in ('source_sha256','schema_sha256'):
            if manifest.get(field)!=baseline[field]:raise AuditError('Operator manifest source/schema drift')
        if manifest.get('source_private_csv_hashes')!=src_hashes:
            raise AuditError('Operator dossier PRIVATE CSV changed since preparing decisions')
        if manifest.get('case_set_sha256')!=case_fingerprint(cases,neg):
            raise AuditError('Operator case set drift')
        q_dec=read_csv(decisions/'amb10_movement_decisions_PRIVATE.csv',Q_FIELDS)
        n_dec=read_csv(decisions/'amb10_negative_decisions_PRIVATE.csv',N_FIELDS)
        for data,expected,fields,editable in (
            (q_dec,cases,Q_FIELDS,REVIEW_FIELDS),(n_dec,neg,N_FIELDS,N_REVIEW_FIELDS)):
            if len(data)!=len(expected):raise AuditError('Modified decision case count')
            lookup={r['case_id']:r for r in expected}
            seen=set()
            for row in data:
                ident=row['case_id'];
                if ident not in lookup or ident in seen:raise AuditError('Unknown/duplicate decision case ID')
                seen.add(ident)
                if set(row)!=set(fields):raise AuditError('Decision CSV schema drift')
                if any(row[k]!=lookup[ident][k] for k in fields if k not in editable):
                    raise AuditError('Read-only decision evidence columns modified: '+ident)
    else:raise AuditError('Unknown mode')
    for row in q_dec:
        decision=row['operator_resolution'].strip()
        if decision not in ('PENDING','INVESTIGATING','DOCUMENTED_FOR_REVIEW','HISTORICAL_UNRESOLVED'):
            raise AuditError('Invalid movement operator_resolution: '+row['case_id'])
        if decision=='DOCUMENTED_FOR_REVIEW':
            if row['direction_verified'].strip()!='YES' or row['movement_kind'].strip() not in KINDS:
                issues.append('MOVEMENT_DIRECTION_OR_KIND_NOT_DOCUMENTED')
            if row['packaging_ownership'].strip() not in OWNERSHIP:
                issues.append('MOVEMENT_OWNERSHIP_NOT_DOCUMENTED')
            if row['source_document_basis'].strip() not in SOURCE_BASIS:
                issues.append('MOVEMENT_SOURCE_BASIS_NOT_DOCUMENTED')
            if not row['independent_evidence_reference'].strip() or not row['operator_initials'].strip():
                issues.append('MOVEMENT_INDEPENDENT_PROOF_MISSING')
            if row['movement_kind'].strip()=='PRENOS_INTERNO' and row['packaging_ownership'].strip()!='OWN_FIRM':
                issues.append('INTERNAL_TRANSFER_REQUIRES_FIRM_OWNERSHIP')
        elif decision=='HISTORICAL_UNRESOLVED':
            if not row['operator_initials'].strip() or not row['operator_note'].strip():
                issues.append('UNRESOLVED_MOVEMENT_LACKS_NOTE')
        elif decision=='PENDING' and any(row[k].strip()!='UNPROVEN' for k in ('direction_verified','movement_kind','packaging_ownership','source_document_basis')):
            issues.append('PENDING_MOVEMENT_HAS_UNREVIEWED_ANNOTATIONS')
    for row in n_dec:
        decision=row['operator_resolution'].strip()
        if decision not in ('PENDING','INVESTIGATING','EVIDENCE_PROVIDED','HISTORICAL_UNRESOLVED'):
            raise AuditError('Invalid negative operator_resolution: '+row['case_id'])
        if decision=='EVIDENCE_PROVIDED':
            if (row['opening_state_evidence'].strip() not in ('DOCUMENTED','UNKNOWN_BUT_DOCUMENTED') or
                row['packaging_ownership'].strip() not in OWNERSHIP or
                not row['independent_evidence_reference'].strip() or not row['operator_initials'].strip()):
                issues.append('NEGATIVE_ACCOUNT_INDEPENDENT_PROOF_MISSING')
        elif decision=='HISTORICAL_UNRESOLVED':
            if not row['operator_initials'].strip() or not row['operator_note'].strip():
                issues.append('UNRESOLVED_NEGATIVE_LACKS_NOTE')
        elif decision=='PENDING' and (row['opening_state_evidence'].strip()!='UNPROVEN' or row['packaging_ownership'].strip()!='UNPROVEN'):
            issues.append('PENDING_NEGATIVE_HAS_UNREVIEWED_ANNOTATIONS')
    counts=Counter(x['operator_resolution'].strip() for x in q_dec)
    n_counts=Counter(x['operator_resolution'].strip() for x in n_dec)
    report=safe_summary(baseline,cases,neg,counts,n_counts,mode,issues)
    return report,q_dec,n_dec,cases,neg,src_hashes

def case_fingerprint(q,n):
    data=json.dumps({'q':[ {k:v for k,v in row.items() if k not in REVIEW_FIELDS} for row in q],
                     'n':[ {k:v for k,v in row.items() if k not in N_REVIEW_FIELDS} for row in n]},
                    ensure_ascii=False,sort_keys=True,separators=(',',':'))
    return hashlib.sha256(data.encode('utf-8')).hexdigest()

def cli(argv=None):
    ap=argparse.ArgumentParser(description='MIG-04.5 local operator evidence register, no AMB10 import')
    ap.add_argument('--mode',choices=['prepare','check'],required=True)
    ap.add_argument('--dossier',type=Path,required=True)
    ap.add_argument('--decisions',type=Path,help='Existing MIG-04.5 operator decisions directory (check only)')
    ap.add_argument('--out',type=Path,required=True)
    ns=ap.parse_args(argv);folder=ns.dossier.resolve();dest=ns.out.resolve()
    decision_dir=ns.decisions.resolve() if ns.decisions else None
    if not folder.is_dir():raise AuditError('Dossier directory not found')
    if ns.mode=='check' and decision_dir is None:raise AuditError('check needs --decisions')
    if ns.mode=='prepare' and decision_dir is not None:raise AuditError('prepare cannot accept --decisions')
    sources=[folder]+([decision_dir] if decision_dir else [])
    if any(dest==p or dest in p.parents or p in dest.parents for p in sources):
        raise AuditError('Output directory overlaps an input')
    if dest.exists():raise AuditError('Output exists; choose a fresh path')
    report,q,n,cases,neg,hashes=build_results(folder,ns.mode,decision_dir)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        (dest/'agrix_amb10_operator_summary.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (dest/'agrix_amb10_operator_summary.txt').write_text(report_txt(report),encoding='utf-8')
        if ns.mode=='prepare':
            dump_csv(dest/'amb10_movement_decisions_PRIVATE.csv',q,Q_FIELDS)
            dump_csv(dest/'amb10_negative_decisions_PRIVATE.csv',n,N_FIELDS)
            manifest={'format':VERSION,'source_sha256':report['source_sha256'],
                      'schema_sha256':report['schema_sha256'],
                      'case_set_sha256':case_fingerprint(cases,neg),
                      'source_private_csv_hashes':hashes,
                      'note':'Must not be manually modified or uploaded.'}
            path=dest/'agrix_amb10_operator_manifest_PRIVATE.json'
            path.write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
            try:os.chmod(path,0o600)
            except OSError:pass
    except BaseException:
        for p in dest.iterdir():
            if p.is_file():p.unlink()
        dest.rmdir()
        raise
    print(report_txt(report))
    if report['validation_issue_counts']:print('WARNING: review annotations incomplete or inconsistent; no import approved.')
    print('Share summary JSON/TXT ONLY. All PRIVATE files stay local.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as ex:
        print('MIG-04.5 STOP:',ex,file=sys.stderr);raise SystemExit(2)
