#!/usr/bin/env python3
"""MIG-03: verify legacy document memberships against PRIVATE MIG-02 stage.

Strictly read-only against source/stage. Produces NEW, PRIVATE candidate JSONL
files, NOT canonical import files. No heuristics that fabricate IDs or links.
Public summary contains counts and stage row ordinals only.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import sys
import uuid

from legacy_audit import AuditError, trim, is_active
from legacy_stage_v2 import SOURCE_SHA, write_private_jsonl

VERSION = "MIG-03.0"
# This is a numeric comparison threshold, not an authorization to alter data.
KG_TOLERANCE = Decimal('0.01')
AMBIENT_TOLERANCE = Decimal('0')


def dec(value):
    value = trim(value)
    if not value: return None
    try:
        n = Decimal(value)
        if n.is_finite(): return n
    except (InvalidOperation, ValueError):
        pass
    return None


def eqkey(a, b):
    return trim(a).casefold() == trim(b).casefold()


def get_rows(stage: Path, table: str) -> list[dict]:
    path = stage / (table + '.jsonl')
    if not path.is_file():
        raise AuditError('Missing staged table: '+table)
    out=[]
    with path.open('r', encoding='utf-8') as f:
        for ordinal,line in enumerate(f,1):
            try: r=json.loads(line)
            except json.JSONDecodeError as e: raise AuditError(f'{table}: invalid JSONL at {ordinal}') from e
            if not isinstance(r, dict):
                raise AuditError(f'{table}: non-dict JSONL at {ordinal}')
            out.append(r)
    return out


def index_ids(rows, pk, name):
    idx={}
    for n,r in enumerate(rows,1):
        key=trim(r.get(pk,''))
        if not key or key in idx:
            raise AuditError(f'{name}: empty or duplicate PK near ordinal {n}')
        idx[key]=(n,r)
    return idx


def new_id(source_hash, table, a, b):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f'agrix-mig03:{source_hash}:{table}:{a}:{b}'))


def diff_fields(child, parent, fields):
    return [k for k in fields if trim(child.get(k,'')) and trim(parent.get(k,'')) and not eqkey(child[k],parent[k])]


def amount_ok(a,b,tolerance):
    if a is None or b is None: return False
    return abs(a-b)<=tolerance


def quantity_compare(parents, children, parent_amount, child_amount, tolerance):
    """Reconcile independently by class. No missing fields treated as zero."""
    class_keys=set(trim(r.get('Klasa','')) for r in parents+children)
    if '' in class_keys:
        return False,'MISSING_CLASS'
    for cls in class_keys:
        pa=[r for r in parents if trim(r.get('Klasa',''))==cls]
        ch=[r for r in children if trim(r.get('Klasa',''))==cls]
        try:
            p=[dec(r.get(parent_amount,'')) for r in pa]
            c=[dec(r.get(child_amount,'')) for r in ch]
            if not p or not c or any(v is None for v in p+c): return False,'MISSING_QTY'
            if not amount_ok(sum(p),sum(c),tolerance): return False,'QTY_MISMATCH'
        except (TypeError, InvalidOperation): return False,'BAD_QTY'
    return True,'OK'


def classified(rows, key):
    return [r for r in rows if is_active(r)==key]


def quantity_by_active_state(parent, children, parent_amt, child_amt):
    """Compare aggregate active membership for active header; if header storno,
    historic book values are not necessarily current membership balances.
    Storno is always reported, never dropped.
    """
    act=is_active(parent)
    if not act:
        return 'HISTORICAL_STORNO_REVIEW'
    relevant=[r for r in children if is_active(r)]
    if not relevant:
        return 'NO_ACTIVE_SOURCES'
    ok,why=quantity_compare([parent], relevant,parent_amt,child_amt,KG_TOLERANCE)
    return 'OK' if ok else why


def link_otkup_to_otp(source_hash, otk, otp):
    otp_index=index_ids(otp,'OtpremnicaID','SOURCE__tblOtpremnica')
    by_parent=defaultdict(list)
    stats=Counter(); problems=[]
    for ordinal,child in enumerate(otk,1):
        pid=trim(child.get('OtpremnicaID',''))
        if not pid:
            stats['NO_EXACT_OTP_ID']+=1
            continue
        item=otp_index.get(pid)
        if not item:
            stats['DANGLING_OTP_ID']+=1
            problems.append({'otkup_row':ordinal,'otp_row':None,'reason':'DANGLING_OTP_ID'})
            continue
        o_ordinal,parent=item
        reasons=diff_fields(child,parent,('StanicaID','VrstaVoca','SortaVoca','TipAmbalaze','Klasa'))
        # Old oto could have optional empty VozacID; don't compare it to OTP driver.
        if reasons:
            stats['BUSINESS_CONTEXT_MISMATCH']+=1
            problems.append({'otkup_row':ordinal,'otp_row':o_ordinal,'reason':'BUSINESS_CONTEXT_MISMATCH',
                             'field_names':reasons})
            continue
        stats['EXACT_ID_CONTEXT_OK']+=1
        by_parent[pid].append((ordinal, child))
    candidates=[]; rec=Counter()
    for o_ordinal,parent in enumerate(otp,1):
        pid=trim(parent.get('OtpremnicaID',''))
        src_pairs=by_parent.get(pid,[])
        if not src_pairs:
            rec['NO_VERIFIED_OTKUP_SOURCE']+=1
            continue
        src=[row for _,row in src_pairs]
        verdict=quantity_by_active_state(parent,src,'Kolicina','Kolicina')
        if verdict=='OK':
            # Cannot infer whether stored gajbe sum is valid: examine separately,
            # but do not block membership on differences in book movement semantics.
            for _,child in src_pairs:
                cid=trim(child['OtkupID'])
                candidates.append({'OtpremnicaIzvorID':new_id(source_hash,'tblOtpremnicaIzvori',pid,cid),
                                   'OtpremnicaID':pid,'OtkupID':cid})
            rec['OTP_HEADERS_QTY_VERIFIED']+=1
            rec['OTP_LINKS_QTY_VERIFIED']+=len(src_pairs)
        else:
            rec['OTP_HEADERS_'+verdict]+=1
            problems.append({'otp_row':o_ordinal,'otkup_rows':[n for n,_ in src_pairs],
                             'reason':verdict})
    return candidates,stats,rec,problems


def link_otp_to_zbirna(source_hash, otp,zbr):
    by_number=defaultdict(list)
    for ordinal,parent in enumerate(zbr,1):
        n=trim(parent.get('BrojZbirne',''))
        if n: by_number[n.casefold()].append((ordinal,parent))
    verified=defaultdict(list)
    stats=Counter();problems=[]
    for ordinal,child in enumerate(otp,1):
        n=trim(child.get('BrojZbirne',''))
        if not n:
            stats['NO_ZBR_NUMBER']+=1
            continue
        group=by_number.get(n.casefold(),[])
        if not group:
            stats['ZBR_NUMBER_MISSING']+=1
            problems.append({'otp_row':ordinal,'reason':'ZBR_NUMBER_MISSING'})
            continue
        # Composite identity: business number+driver+class, then compare fruits.
        scoped=[(j,z) for j,z in group if eqkey(child.get('VozacID',''),z.get('VozacID',''))
                and eqkey(child.get('Klasa',''),z.get('Klasa',''))]
        if len(scoped)!=1:
            why='ZBR_SCOPE_NOT_FOUND' if not scoped else 'ZBR_SCOPE_AMBIGUOUS'
            stats[why]+=1
            problems.append({'otp_row':ordinal,'zbirna_rows':[j for j,_ in scoped], 'reason':why})
            continue
        z_ordinal,parent=scoped[0]
        differences=diff_fields(child,parent,('VrstaVoca','SortaVoca','TipAmbalaze'))
        if differences:
            stats['ZBR_CONTEXT_MISMATCH']+=1
            problems.append({'otp_row':ordinal,'zbirna_row':z_ordinal,'reason':'ZBR_CONTEXT_MISMATCH',
                             'field_names':differences})
            continue
        stats['ZBR_SCOPE_OK']+=1
        verified[trim(parent['ZbirnaID'])].append((ordinal,child))
    rec=Counter();candidates=[]
    for z_ordinal,parent in enumerate(zbr,1):
        zid=trim(parent.get('ZbirnaID',''))
        pairs=verified.get(zid,[])
        if not pairs:
            rec['NO_VERIFIED_OTPREMNICA_SOURCE']+=1
            continue
        verdict=quantity_by_active_state(parent,[r for _,r in pairs],'UkupnoKolicina','Kolicina')
        if verdict=='OK':
            for _,child in pairs:
                cid=trim(child['OtpremnicaID'])
                candidates.append({'ZbirnaIzvorID':new_id(source_hash,'tblZbirnaIzvori',zid,cid),
                                   'ZbirnaID':zid,'OtpremnicaID':cid})
            rec['ZBR_HEADERS_QTY_VERIFIED']+=1
            rec['ZBR_LINKS_QTY_VERIFIED']+=len(pairs)
        else:
            rec['ZBR_HEADERS_'+verdict]+=1
            problems.append({'zbirna_row':z_ordinal,'otp_rows':[n for n,_ in pairs],
                             'reason':verdict})
    return candidates,stats,rec,problems


def receipt_parent_report(prij,zbr):
    bynumber=defaultdict(list)
    for ordinal,row in enumerate(zbr,1):
        n=trim(row.get('BrojZbirne',''))
        if n: bynumber[n.casefold()].append((ordinal,row))
    report=Counter();problems=[]
    for ordinal,row in enumerate(prij,1):
        n=trim(row.get('BrojZbirne',''))
        if not n:
            report['EMPTY_REFERENCE']+=1;continue
        options=bynumber.get(n.casefold(),[])
        if not options:
            report['ZBR_NUMBER_MISSING']+=1
            problems.append({'prijemnica_row':ordinal,'reason':'ZBR_NUMBER_MISSING'})
            continue
        candidates=[(j,z) for j,z in options if eqkey(row.get('VozacID',''),z.get('VozacID',''))
                    and eqkey(row.get('KupacID',''),z.get('KupacID',''))
                    and eqkey(row.get('Klasa',''),z.get('Klasa',''))]
        if len(candidates)!=1:
            reason='SCOPE_MISMATCH' if not candidates else 'SCOPE_AMBIGUOUS'
            report[reason]+=1
            problems.append({'prijemnica_row':ordinal,'zbirna_rows':[j for j,_ in candidates],
                             'reason':reason})
            continue
        j,z=candidates[0]
        differences=diff_fields(row,z,('VrstaVoca','SortaVoca','TipAmbalaze'))
        if differences:
            report['BUSINESS_CONTEXT_MISMATCH']+=1
            problems.append({'prijemnica_row':ordinal,'zbirna_row':j,
                             'reason':'BUSINESS_CONTEXT_MISMATCH','field_names':differences})
        else:
            report['SCOPE_VERIFIED']+=1
    return report,problems


def process_stage(stage:Path, source_hash=SOURCE_SHA):
    mf=stage/'agrix_stage_summary.json'
    if not mf.is_file():raise AuditError('Stage manifest missing')
    summary=json.loads(mf.read_text(encoding='utf-8'))
    if summary.get('format')!='MIG-02.0' or summary.get('status')!='STAGED_ONLY_NO_XLSM_CREATED':
        raise AuditError('Input is not MIG-02 stage')
    if summary.get('source_sha256','').lower()!=source_hash.lower():
        raise AuditError('Source SHA mismatch; DO NOT combine two workbooks')
    otk=get_rows(stage,'SOURCE__tblOtkup')
    otp=get_rows(stage,'SOURCE__tblOtpremnica')
    zbr=get_rows(stage,'SOURCE__tblZbirna')
    prij=get_rows(stage,'tblPrijemnica')
    for table,rows in [('tblOtkup',otk),('tblOtpremnica',otp),('tblZbirna',zbr),('tblPrijemnica',prij)]:
        expected=summary.get('source_row_counts',{}).get(table)
        if len(rows)!=expected:raise AuditError(f'{table}: source count drift from MIG-02 manifest')
    index_ids(otk,'OtkupID','tblOtkup')
    index_ids(otp,'OtpremnicaID','tblOtpremnica')
    index_ids(zbr,'ZbirnaID','tblZbirna')
    otpc,ots,otr,otp_issues=link_otkup_to_otp(source_hash,otk,otp)
    zbc,zs,zr,zbr_issues=link_otp_to_zbirna(source_hash,otp,zbr)
    rc,rcissues=receipt_parent_report(prij,zbr)
    counts={
        'otkup_to_otp_match':dict(ots),'otkup_to_otp_reconciliation':dict(otr),
        'otp_to_zbr_match':dict(zs),'otp_to_zbr_reconciliation':dict(zr),
        'prijemnica_to_zbr_match':dict(rc),
        'candidate_otp_sources':len(otpc),'candidate_zbr_sources':len(zbc),
        'remaining_issues':len(otp_issues)+len(zbr_issues)+len(rcissues),
    }
    report={'format':VERSION,'status':'CANDIDATES_ONLY_NOT_APPROVED',
      'source_sha256':source_hash,'input_stage_schema_sha256':summary.get('schema_sha256'),
      'source_counts':{t:len(v) for t,v in [('tblOtkup',otk),('tblOtpremnica',otp),
          ('tblZbirna',zbr),('tblPrijemnica',prij)]},
      'statistics':counts,
      'issues_by_row_ordinal':{
        'otkup_to_otp':otp_issues,'otp_to_zbr':zbr_issues,
        'prijemnica_to_zbr':rcissues},
      'conditions':['Candidate links are NOT import authorization',
        'All IDs/amounts and linkage content remain in local private JSONL only',
        'Historical storno headers need independent treatment, not automatic active membership',
        'AMB-10 conversion NOT done, no XLSM created or changed',
        'Aggregate discrepancies must not be silently overwritten']}
    return {'CANDIDATE__tblOtpremnicaIzvori':otpc,
            'CANDIDATE__tblZbirnaIzvori':zbc}, report


def cli(argv=None):
    ap=argparse.ArgumentParser(description='MIG-03: prove exact header->source membership from private MIG-02 stage')
    ap.add_argument('--stage',required=True,help='private directory MigrationStage_v2 (NE upload)')
    ap.add_argument('--out',required=True,help='fresh local destination directory')
    args=ap.parse_args(argv)
    stage=Path(args.stage).expanduser().resolve()
    dest=Path(args.out).expanduser().resolve()
    if not stage.is_dir():raise AuditError('MIG-02 stage directory not found')
    if dest.exists():raise AuditError('Output already exists; use a fresh directory')
    if dest==stage or dest.is_relative_to(stage) or stage.is_relative_to(dest):
        raise AuditError('Output directory overlaps stage')
    payload,report=process_stage(stage)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        for name,data in payload.items():
            write_private_jsonl(dest/(name+'.jsonl'), data)
        (dest/'agrix_membership_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        lines=['AGRIX MIG-03 -- DOCUMENT MEMBERSHIP / RECONCILIATION',
          'Status: CANDIDATES ONLY, NO XLSM WRITTEN',
          'Source SHA256: '+report['source_sha256']]
        for heading,stats in report['statistics'].items():
            if isinstance(stats,dict):
                lines.append(heading+':')
                for k,v in sorted(stats.items()):lines.append('  %-37s %6s'%(k,v))
            else: lines.append('%s: %s'%(heading,stats))
        lines+=['','ISSUES BY SOURCE ROW ORDINAL (no IDs or amounts):']
        for section,items in report['issues_by_row_ordinal'].items():
            lines.append('  %s: %d'%(section,len(items)))
            for entry in items[:35]:lines.append('    '+json.dumps(entry,ensure_ascii=False))
            if len(items)>35:lines.append('    ... and %d more in JSON'%(len(items)-35))
        lines+=['','STOP: candidate JSONL is PRIVATE. Do not import it or upload it.',
               'Safe to share: agrix_membership_summary.json/.txt only.']
        txt='\n'.join(lines)+'\n'
        (dest/'agrix_membership_summary.txt').write_text(txt,encoding='utf-8')
        print(txt)
        print('Output: '+str(dest))
    except BaseException:
        for f in dest.glob('*'):f.unlink(missing_ok=True)
        try:dest.rmdir()
        except OSError:pass
        raise
    return 0


if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError, OSError, ValueError, json.JSONDecodeError) as e:
        print('MIG-03 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
