#!/usr/bin/env python3
"""MIG-10.1: evidence-only triage of the ONE known historical Otkup exception.

Private MIG-02 staging never leaves the operator PC. Only category flags and
source ordinals are written in the shareable summary. No XLSM/COM access.
"""
from __future__ import annotations
import argparse
from decimal import Decimal, InvalidOperation
from pathlib import Path
import json
import sys

# Reuse MIG-10 exact source identity, 1702-row invariant, and row-by-row contract.
from agrix_mig10_otkup_contract_audit import audit, load_jsonl, AuditError, EXPECTED

SOURCE_LINK_FIELDS = ('OtpremnicaID', 'BrojOtpremnice', 'BrojZbirne')
CORRECTION_FIELDS = ('IspravkaOd', 'IspravkaOdID', 'ZamenjenSa', 'ZamenjenSaID',
                     'KorekcijaOd', 'IspravkaZa', 'IspravkaZaID')

def classified_storno(val):
    s = '' if val is None else str(val).strip().casefold()
    if s in ('da', 'true', '1', 'yes'):
        return 'CANCELLED'
    if s in ('', 'ne', 'false', '0', 'no'):
        return 'NOT_CANCELLED'
    return 'UNKNOWN_STATUS'


def classified_numeric(val):
    if val is None or isinstance(val, bool):
        return 'MISSING_OR_INVALID'
    s = str(val).strip()
    if not s:
        return 'MISSING_OR_INVALID'
    try:
        d = Decimal(s)
    except InvalidOperation:
        return 'MISSING_OR_INVALID'
    if not d.is_finite():
        return 'MISSING_OR_INVALID'
    return 'NEGATIVE' if d < 0 else 'ZERO' if d == 0 else 'POSITIVE'


def present_fields(row, names):
    return {k: bool(str(row.get(k, '') or '').strip()) for k in names if k in row}


def triage(stage, row_number):
    earlier = audit(stage)  # checks manifest, source SHA, schema SHA, stage totals
    if earlier['issue_counts'] != {'QUANTITY_NONPOSITIVE': 1} or earlier['affected_source_row_count'] != 1:
        raise AuditError('MIG-10 finding changed; run/review MIG-10 before this narrow triage')
    issues = earlier['issues_by_source_row_ordinal']
    if len(issues) != 1 or issues[0]['source_row_ordinal'] != row_number:
        raise AuditError('Historical exception ordinal differs; refuse to infer identity')
    files = ['SOURCE__tblOtkup','tblOtkup','tblOtkupStavke']
    rows = [load_jsonl(stage/(f+'.jsonl'),EXPECTED)[row_number-1] for f in files]
    src, head, line = rows
    id_source = str(src.get('OtkupID','')).strip()
    if not id_source or id_source != str(head.get('OtkupID','')).strip() or id_source != str(line.get('OtkupID','')).strip():
        raise AuditError('Otkup identity no longer matches across source/header/item')
    st_src = classified_storno(src['Stornirano']) if 'Stornirano' in src else 'UNKNOWN_STATUS'
    st_head = classified_storno(head['Stornirano']) if 'Stornirano' in head else 'UNKNOWN_STATUS'
    if st_src != st_head:
        raise AuditError('Storno status changed source-to-stage')
    q1, q2 = classified_numeric(src.get('Kolicina')), classified_numeric(line.get('Kolicina'))
    p1, p2 = classified_numeric(src.get('Cena')), classified_numeric(line.get('Cena'))
    a1, a2 = classified_numeric(src.get('KolAmbalaze')), classified_numeric(line.get('KolAmbalaze'))
    if (q1,p1,a1) != (q2,p2,a2):
        raise AuditError('Business numeric category changed source-to-stage')
    if q1 not in ('ZERO','NEGATIVE'):
        raise AuditError('Expected one nonpositive quantity, found another value')
    if st_src == 'CANCELLED':
        decision = 'CANCELLED_HISTORICAL_EXCEPTION_NEEDS_EXPLICIT_READ_POLICY'
    elif st_src == 'NOT_CANCELLED':
        decision = 'ACTIVE_HISTORICAL_EXCEPTION_NEEDS_OPERATOR_DECISION'
    else:
        decision = 'UNRECOGNIZED_STORNO_STATUS_NEEDS_OPERATOR_DECISION'
    return {
       'format':'MIG-10.1-HISTORICAL-OTKUP-EXCEPTION-TRIAGE',
       'status':'EVIDENCE_ONLY_NO_MUTATION',
       'source_row_ordinal':row_number,
       'evidence':{'original_and_staged_identity_match':True,
                   'original_and_staged_numeric_categories_match':True,
                   'original_and_staged_cancellation_status_match':True,
                   'historical_cancellation_status':st_src,
                   'quantity_category':q1,
                   'price_category':p1,
                   'packaging_quantity_category':a1,
                   'legacy_document_links_present':present_fields(src,SOURCE_LINK_FIELDS),
                   'legacy_correction_links_present':present_fields(src,CORRECTION_FIELDS)},
       'decision_gate':decision,
       'historical_row_retained':True,
       'source_workbook_modified':False,
       'stage_modified':False,
       'master_or_data_copy_modified':False,
       'source_or_stage_private_ids_quantities_prices_not_exported':True,
       'data_migration_production_approved':False,
       'note':'A cancelled old document with zero quantity is not by itself proof that its goods, finance or ambalaza effects may be discarded.'
    }


def cli():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage', required=True)
    p.add_argument('--out',required=True,help='New output folder separate from private stage')
    p.add_argument('--ordinal',type=int,default=123)
    a=p.parse_args()
    st=Path(a.stage).expanduser().resolve()
    out=Path(a.out).expanduser().resolve()
    if not 1<=a.ordinal<=EXPECTED:
        raise AuditError('Invalid source ordinal')
    if out.exists():
        raise AuditError('Output folder already exists: use a fresh directory')
    if st==out or st in out.parents or out in st.parents:
        raise AuditError('Output folder must not nest inside private staging')
    result=triage(st,a.ordinal)
    out.mkdir(parents=True,exist_ok=False)
    (out/'agrix_mig101_exception_summary.json').write_text(json.dumps(result,indent=2,ensure_ascii=True)+'\n',encoding='utf-8')
    e=result['evidence']
    lines=[result['format'], 'STATUS: '+result['status'],
           'Source ordinal: '+str(result['source_row_ordinal']),
           'Cancellation: '+e['historical_cancellation_status'],
           'Quantity category: '+e['quantity_category'],
           'Price category: '+e['price_category'],
           'Packaging quantity category: '+e['packaging_quantity_category'],
           'Legacy links present: '+json.dumps(e['legacy_document_links_present'],sort_keys=True),
           'Correction links present: '+json.dumps(e['legacy_correction_links_present'],sort_keys=True),
           'DECISION: '+result['decision_gate'],
           'READ-ONLY. NO XLSM WRITES. NO PRIVATE NUMERIC VALUES, NAMES OR IDS EXPORTED.',
           'NO PRODUCTION GO.']
    (out/'agrix_mig101_exception_summary.txt').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print('\n'.join(lines),flush=True)
    print('Share ONLY agrix_mig101_exception_summary.json and .txt',flush=True)
    return 0

if __name__=='__main__':
    try:sys.exit(cli())
    except (AuditError,ValueError,OSError,KeyError,json.JSONDecodeError) as exc:
        print('MIG-10.1 STOP: '+str(exc),file=sys.stderr)
        sys.exit(2)
