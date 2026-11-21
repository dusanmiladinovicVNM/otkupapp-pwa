#!/usr/bin/env python3
"""MIG-10.0 read-only local forensic audit of canonical OtkupStavke against source.

No XLSM or Excel required; never uploads, edits, or prints business values.
Compares exact stage and original source columns from MIG-02; reports row ordinals only.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import hashlib
import json
import os
import sys

SOURCE_SHA = 'a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'
SCHEMA_SHA = 'b9916598d1f8d29b8e66b3fd0805d36f4c2e695df48fe6ef32155a7b6a765d73'
EXPECTED = 1702
COLS = ('Klasa','Kolicina','Cena','KolAmbalaze','BrutoKg')

class AuditError(Exception): pass


def load_jsonl(path, expected):
    if not path.is_file(): raise AuditError('Missing private staging file: ' + path.name)
    with path.open(encoding='utf-8') as f:
        rows = [json.loads(x) for x in f if x.strip()]
    if len(rows)!=expected or not all(isinstance(r,dict) for r in rows):
        raise AuditError('Unexpected record count/shape for '+path.name)
    return rows


def num(raw):
    if raw is None or isinstance(raw,bool): return None
    s = str(raw).strip()
    if not s: return None
    try:
        d = Decimal(s)
        return d if d.is_finite() else None
    except InvalidOperation:
        return None


def is_same(src, stage):
    # Source JSONL cell is copied without numerical conversion by MIG-02.
    # Missing field (rather than explicit empty field) is equivalent to empty.
    a = '' if src is None else str(src)
    b = '' if stage is None else str(stage)
    return a == b


def scan(heads, lines, srcrows, lookup=None):
    if not (len(heads)==len(lines)==len(srcrows)==EXPECTED):
        raise AuditError('Refuse changing MIG-02 1702-row ordinal contract')
    issues = defaultdict(list)
    origin = Counter()
    lookup_ord = []
    out_issues = []
    seen_h, seen_l = set(), set()
    by_class = defaultdict(set)
    for k, (head, line, src) in enumerate(zip(heads, lines, srcrows), 1):
        hid = str(head.get('OtkupID','')).strip()
        lid = str(line.get('OtkupID','')).strip()
        sid = str(src.get('OtkupID','')).strip()
        if lookup and (lookup == hid or lookup == sid or lookup == lid):
            lookup_ord.append(k)
        bad=[]
        def add(code):
            bad.append(code)
            issues[code].append(k)
        if not hid or not lid or not sid or hid!=lid or hid!=sid: add('PARENT_SOURCE_LINE_ID_MISMATCH')
        if hid in seen_h or (lid in seen_l): add('DUPLICATE_DOCUMENT_ID')
        seen_h.add(hid);seen_l.add(lid)
        klas = str(line.get('Klasa','')).strip()
        if klas not in ('I','II'): add('INVALID_CLASS')
        if hid and klas and klas in by_class[hid]: add('DUPLICATE_CLASS_PER_OTKUP')
        by_class[hid].add(klas)
        q, p = num(line.get('Kolicina')), num(line.get('Cena'))
        if q is None: add('QUANTITY_NOT_NUMERIC_OR_MISSING')
        elif q <= 0: add('QUANTITY_NONPOSITIVE')
        if p is None: add('PRICE_NOT_NUMERIC_OR_MISSING')
        elif p <= 0: add('PRICE_NONPOSITIVE')
        a = line.get('KolAmbalaze','')
        if a is not None and str(a).strip():
            ad = num(a)
            if ad is None: add('AMBALAZA_NOT_NUMERIC')
            elif ad<0: add('AMBALAZA_NEGATIVE')
            elif ad != ad.to_integral_value(): add('AMBALAZA_NOT_INTEGER')
        bruto = line.get('BrutoKg','')
        if bruto is not None and str(bruto).strip():
            bd=num(bruto)
            if bd is None: add('BRUTO_NOT_NUMERIC')
            elif bd<0: add('BRUTO_NEGATIVE')
            elif q is not None and q>0 and bd>0 and bd<q: add('BRUTO_BELOW_NET')
        for c in COLS:
            if not is_same(src.get(c,''),line.get(c,'')):
                add('SOURCE_TO_STAGE_MISMATCH_'+c.upper())
        # Attribution does not reveal actual customer amounts/IDs.
        if bad:
            differs=any(x.startswith('SOURCE_TO_STAGE_MISMATCH') or x=='PARENT_SOURCE_LINE_ID_MISMATCH' for x in bad)
            origin['stage_differs_from_source' if differs else 'unchanged_historical_source']+=1
            out_issues.append({'source_row_ordinal':k,'findings':bad})
    # Tests only own stage; validation of actual XLSM was separately proven by MIG-09.
    result={
      'format':'MIG-10.0-OTKUP-STAVKE-LOCAL-CONTRACT-AUDIT',
      'status':'BLOCKED_BY_OTKUP_CONTRACT' if issues else 'OTKUP_CONTRACT_PASS',
      'stage_rows':len(lines), 'source_rows':len(srcrows),
      'reader_contract':'modOtkup.StavkeOtkupaRedovi (git 66f7897f)',
      'issue_counts':dict(sorted((k,len(v)) for k,v in issues.items())),
      'affected_source_row_count':len(out_issues),
      'origin_of_flagged_rows':dict(origin),
      'lookup_source_row_ordinals':lookup_ord,
      'issues_by_source_row_ordinal':out_issues,
      'customer_ids_names_prices_quantities_dates_not_exported':True,
      'read_only':True, 'changes_to_customer_data':0,
      'migration_approved':False,
      'note':'A zero/negative source value is evidence of a historical exception, not permission to fabricate a corrected value.'
    }
    return result


def audit(stage, lookup=None):
    manifest=stage/'agrix_stage_summary.json'
    if not manifest.is_file(): raise AuditError('MIG-02 stage manifest missing')
    m=json.loads(manifest.read_text(encoding='utf-8'))
    if m.get('format')!='MIG-02.0' or m.get('source_sha256')!=SOURCE_SHA or m.get('schema_sha256')!=SCHEMA_SHA:
        raise AuditError('Wrong stage/source/schema identity, refuse read')
    c=m.get('canonical_stage_counts',{})
    if any(c.get(n)!=EXPECTED for n in ('tblOtkup','tblOtkupStavke','SOURCE__tblOtkup')):
        raise AuditError('Unexpected stage manifest count')
    return scan(*(load_jsonl(stage/(n+'.jsonl'),EXPECTED) for n in
                  ('tblOtkup','tblOtkupStavke','SOURCE__tblOtkup')),lookup=lookup)


def cli():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--stage',required=True)
    ap.add_argument('--out',required=True,help='fresh folder for shareable aggregate reports')
    ap.add_argument('--lookup-id',help='optional LOCAL lookup; only source row ordinal is reported, never ID')
    args=ap.parse_args()
    stage=Path(args.stage).expanduser().resolve()
    out=Path(args.out).expanduser().resolve()
    if out.exists(): raise AuditError('Output folder already exists; use a fresh folder')
    if out==stage or stage in out.parents or out in stage.parents:
        raise AuditError('Audit output must be disjoint from private stage')
    report=audit(stage,args.lookup_id)
    out.mkdir(parents=True,exist_ok=False)
    js=out/'agrix_mig10_otkup_contract_summary.json'
    tx=out/'agrix_mig10_otkup_contract_summary.txt'
    js.write_text(json.dumps(report,indent=2,ensure_ascii=True)+'\n',encoding='utf-8')
    lines=['MIG-10.0 read-only Otkup stage/source contract audit',
           'Status: '+report['status'],
           f"Rows scanned: {report['stage_rows']}; affected source ordinals: {report['affected_source_row_count']}",
           'Attributed: '+json.dumps(report['origin_of_flagged_rows'],sort_keys=True),
           'Issue count by type:']
    lines += [f'  {k}: {v}' for k,v in report['issue_counts'].items()]
    lines += ['Lookup source ordinal(s): '+str(report['lookup_source_row_ordinals']),
              'Private source values, IDs, money and names: NOT EXPORTED.',
              'Read only: source, stage, master and migrated XLSM not modified.',
              'NO automatic correction, NO production GO.']
    tx.write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print('\n'.join(lines),flush=True)
    print('Share ONLY '+js.name+' and '+tx.name,flush=True)
    return 0

if __name__=='__main__':
    try:sys.exit(cli())
    except (AuditError,ValueError,OSError,KeyError,json.JSONDecodeError) as exc:
        print('MIG-10 STOP:',exc,file=sys.stderr)
        sys.exit(2)
