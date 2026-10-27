#!/usr/bin/env python3
"""MIG-03.1: read-only membership guard and discrepancy triage for legacy _v2.

NO Excel file written. Reads local private MIG-02 + MIG-03 directories, creates
privacy-safe summary and NEW local provisional *active-only* candidate JSONL.
Historical memberships are never silently fabricated or deleted: they remain
explicitly unresolved for subsequent migration of storno lineage.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_membership_v2 import (get_rows, index_ids, diff_fields,
                                 quantity_by_active_state, dec, eqkey, KG_TOLERANCE)
from legacy_stage_v2 import SOURCE_SHA, write_private_jsonl

VERSION = 'MIG-03.1'
DOCS = ('SOURCE__tblOtkup', 'SOURCE__tblOtpremnica',
        'SOURCE__tblZbirna', 'tblPrijemnica')

def read_json(path):
    if not path.is_file():
        raise AuditError('Required manifest missing: '+path.name)
    d=json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(d,dict):
        raise AuditError('Invalid manifest: '+path.name)
    return d


def ratio_band(parent, active):
    p=dec(parent);a=dec(active)
    if p is None or a is None or p <= 0:
        return 'INVALID_NUMERIC_OR_ZERO_PARENT', 'UNDETERMINED'
    delta=a-p
    proportion=abs(delta)/p
    bucket=('LE_0_1_PERCENT' if proportion<=Decimal('.001') else
            'LE_1_PERCENT' if proportion<=Decimal('.01') else
            'LE_5_PERCENT' if proportion<=Decimal('.05') else
            'LE_20_PERCENT' if proportion<=Decimal('.2') else
            'GT_20_PERCENT')
    return bucket, ('SOURCES_GREATER' if delta>0 else 'HEADER_GREATER' if delta<0 else 'EQUAL')


def sum_qty(rows, field):
    vals=[dec(r.get(field)) for r in rows]
    return None if not vals or any(x is None for x in vals) else sum(vals,Decimal(0))


def eqsum(x,y):
    return x is not None and y is not None and abs(x-y) <= KG_TOLERANCE


def zbr_scope(child, z):
    return (eqkey(child.get('VozacID',''),z.get('VozacID','')) and
            eqkey(child.get('Klasa',''),z.get('Klasa','')) and
            not diff_fields(child,z,('VrstaVoca','SortaVoca','TipAmbalaze')))


def inspect(stage:Path,membership:Path):
    manifest=read_json(stage/'agrix_stage_summary.json')
    baseline=read_json(membership/'agrix_membership_summary.json')
    if manifest.get('format')!='MIG-02.0' or manifest.get('status')!='STAGED_ONLY_NO_XLSM_CREATED':
        raise AuditError('Requires MIG-02.0 source-pinned stage')
    if manifest.get('source_sha256')!=SOURCE_SHA or baseline.get('source_sha256')!=SOURCE_SHA:
        raise AuditError('Source SHA differs from locked _v2 snapshot')
    if baseline.get('format')!='MIG-03.0' or baseline.get('status')!='CANDIDATES_ONLY_NOT_APPROVED':
        raise AuditError('Requires MIG-03.0 membership candidates')
    if baseline.get('input_stage_schema_sha256') != manifest.get('schema_sha256'):
        raise AuditError('Stage schema checksum mismatch')
    tables={key:get_rows(stage,key) for key in DOCS}
    for key, rows in tables.items():
        orig=key.removeprefix('SOURCE__')
        if len(rows)!=manifest.get('source_row_counts',{}).get(orig):
            raise AuditError('Stage row count mismatch: '+orig)
    otk,otp,zbr,prij=(tables[x] for x in DOCS)
    otki=index_ids(otk,'OtkupID','source otkup')
    otpi=index_ids(otp,'OtpremnicaID','source otpremnica')
    zbri=index_ids(zbr,'ZbirnaID','source zbirna')
    stats=Counter(); details={}
    candidate_specs=(('CANDIDATE__tblOtpremnicaIzvori','OtpremnicaID','OtkupID',otpi,otki),
                     ('CANDIDATE__tblZbirnaIzvori','ZbirnaID','OtpremnicaID',zbri,otpi))
    clean={}
    for name,pk,fk,parents,children in candidate_specs:
        rows=get_rows(membership,name)
        expected=baseline.get('statistics',{}).get('candidate_otp_sources' if pk=='OtpremnicaID' else 'candidate_zbr_sources')
        if len(rows)!=expected:
            raise AuditError('Candidate count drift: '+name)
        seen=set();out=[];flagged=[]
        for r in rows:
            pid=trim(r.get(pk));cid=trim(r.get(fk))
            if not pid or not cid or (pid,cid) in seen or pid not in parents or cid not in children:
                raise AuditError('Duplicate/dangling candidate in '+name)
            seen.add((pid,cid))
            pi,pr=parents[pid];ci,cr=children[cid]
            if not is_active(pr) or not is_active(cr):
                reason=('STORNO_PARENT' if not is_active(pr) else 'STORNO_SOURCE')
                flagged.append({'parent_row':pi,'source_row':ci,'reason':reason})
                stats[name+'_REVIEW_'+reason]+=1
            else:
                out.append(r)
        clean['PROVISIONAL_ACTIVE__'+name.removeprefix('CANDIDATE__')]=out
        details[name]={'original_count':len(rows),'provisional_active_count':len(out),
                       'historical_excluded_from_active_only':len(flagged),
                       'historical_review':flagged}

    byotp=defaultdict(list)
    for k,r in enumerate(otk,1):
        pid=trim(r.get('OtpremnicaID'))
        if pid:
            byotp[pid].append((k,r))
    otp_counts=Counter();otp_review=[]
    for idx,parent in enumerate(otp,1):
        pid=trim(parent.get('OtpremnicaID'))
        raw=byotp.get(pid,[])
        good=[(i,r) for i,r in raw if not diff_fields(r,parent,('StanicaID','VrstaVoca','SortaVoca','TipAmbalaze','Klasa'))]
        rejected=[(i,r) for i,r in raw if (i,r) not in good]
        verdict=('NO_VERIFIED_OTKUP_SOURCE' if not good else
                 quantity_by_active_state(parent,[r for _,r in good],'Kolicina','Kolicina'))
        if verdict=='OK':
            otp_counts['VERIFIED']+=1
            continue
        otp_counts[verdict]+=1
        active=[r for _,r in good if is_active(r)]
        hist=[r for _,r in good if not is_active(r)]
        parent_amount=dec(parent.get('Kolicina'))
        active_amount=sum_qty(active,'Kolicina')
        all_amount=sum_qty([r for _,r in good],'Kolicina')
        band,direction=ratio_band(parent_amount,active_amount)
        reason_detail={'otp_row':idx,'reason':verdict,
                       'raw_source_count':len(raw),'valid_source_count':len(good),
                       'active_valid_count':len(active),'storno_valid_count':len(hist),
                       'context_rejected_count':len(rejected)}
        if verdict=='NO_VERIFIED_OTKUP_SOURCE':
            reason_detail['cause']='NO_DIRECT_ID_LINKS' if not raw else 'ALL_DIRECT_LINKS_FAILED_CONTEXT'
        if verdict=='QTY_MISMATCH':
            reason_detail.update({'direction':direction,'relative_difference_bucket':band,
                                  'all_status_sources_would_match':eqsum(parent_amount,all_amount),
                                  'excluded_context_source_rows':[i for i,_ in rejected]})
            otp_counts['QTY_'+direction]+=1
            otp_counts['QTY_BAND_'+band]+=1
            if eqsum(parent_amount,all_amount):otp_counts['QTY_ALL_STATUSES_WOULD_MATCH']+=1
        otp_review.append(reason_detail)
    details['otp_reconciliation']={'counts':dict(otp_counts),'review':otp_review}

    byzbrnum=defaultdict(list)
    for i,row in enumerate(otp,1):
        n=trim(row.get('BrojZbirne'))
        if n:byzbrnum[n.casefold()].append((i,row))
    zb_counts=Counter();zb_review=[]
    for i,z in enumerate(zbr,1):
        number=trim(z.get('BrojZbirne')).casefold()
        raws=byzbrnum.get(number,[])
        pairs=[(j,o) for j,o in raws if zbr_scope(o,z)]
        verdict=('NO_VERIFIED_OTPREMNICA_SOURCE' if not pairs else
                 quantity_by_active_state(z,[row for _,row in pairs],
                                          'UkupnoKolicina','Kolicina'))
        if verdict=='OK':
            zb_counts['VERIFIED']+=1;continue
        zb_counts[verdict]+=1
        zb_review.append({'zbirna_row':i,'reason':verdict,'raw_number_matches':len(raws),
                          'scope_matched_sources':len(pairs),
                          'active_sources':sum(is_active(o) for _,o in pairs),
                          'storno_sources':sum(not is_active(o) for _,o in pairs)})
    details['zbr_reconciliation']={'counts':dict(zb_counts),'review':zb_review}

    # Exact/casefold discrepancy check: MIG-02's old number match is case sensitive.
    znums={trim(z.get('BrojZbirne')) for z in zbr if trim(z.get('BrojZbirne'))}
    bynum=defaultdict(list)
    for i,z in enumerate(zbr,1):
        if trim(z.get('BrojZbirne')):
            bynum[trim(z['BrojZbirne']).casefold()].append((i,z))
    rc_counts=Counter();rc_review=[]
    for i,p in enumerate(prij,1):
        raw=trim(p.get('BrojZbirne'))
        if not raw:
            rc_counts['EMPTY_NUMBER']+=1;continue
        if raw not in znums:
            rc_counts['MIG02_EXACT_NUMBER_NOT_FOUND']+=1
            if raw.casefold() in bynum:
                rc_counts['NUMBER_RESOLVED_BY_CASEFOLD']+=1
                rc_review.append({'prijemnica_row':i,'reason':'CASEFOLD_MATCH_NOT_EXACT'})
            else:
                rc_counts['NUMBER_ABSENT_EVEN_AFTER_CASEFOLD']+=1
                rc_review.append({'prijemnica_row':i,'reason':'NUMBER_NOT_FOUND'})
        group=bynum.get(raw.casefold(),[])
        valid=[(j,z) for j,z in group if eqkey(p.get('VozacID',''),z.get('VozacID','')) and
               eqkey(p.get('KupacID',''),z.get('KupacID','')) and
               eqkey(p.get('Klasa',''),z.get('Klasa',''))]
        if not valid:
            if group:
                rc_counts['SCOPE_MISMATCH']+=1
                rc_review.append({'prijemnica_row':i,'reason':'SCOPE_MISMATCH',
                    'same_number_candidates':len(group),
                    'fields_with_some_match':[key for key in ('VozacID','KupacID','Klasa')
                                               if any(eqkey(p.get(key,''),z.get(key,'')) for _,z in group)]})
            continue
        if len(valid)>1:
            rc_counts['AMBIGUOUS']+=1
            rc_review.append({'prijemnica_row':i,'reason':'AMBIGUOUS','candidates':len(valid)})
            continue
        j,z=valid[0]
        diffs=diff_fields(p,z,('VrstaVoca','SortaVoca','TipAmbalaze'))
        if diffs:
            rc_counts['BUSINESS_CONTEXT_MISMATCH']+=1
            rc_review.append({'prijemnica_row':i,'zbirna_row':j,
                              'reason':'BUSINESS_CONTEXT_MISMATCH','field_names':diffs})
    details['receipt_reconciliation']={'counts':dict(rc_counts),'review':rc_review}

    report={'format':VERSION,'status':'READ_ONLY_ACTIVE_CANDIDATES_PROVISIONAL_NOT_IMPORT_AUTHORIZATION',
            'source_sha256':SOURCE_SHA,'stage_schema_sha256':manifest.get('schema_sha256'),
            'original_mig03_issue_count':baseline.get('statistics',{}).get('remaining_issues'),
            'candidate_safety':{k:v for k,v in details.items() if k.startswith('CANDIDATE__')},
            'otp_reconciliation':details['otp_reconciliation'],
            'zbr_reconciliation':details['zbr_reconciliation'],
            'receipt_reconciliation':details['receipt_reconciliation'],
            'notes':['Historical membership must be handled separately, not as active candidate.',
                     'A valid ID link is not proof of correct business composition.',
                     'QTY_MISMATCH is a discrepancy, not automatically a business-data error.',
                     'Preserve legacy raw source and full storno history.',
                     'No business values or names are present in this public report.',
                     'No XLSM written; no AMB-10 conversion performed.']}
    return clean,report


def cli(argv=None):
    parser=argparse.ArgumentParser(description='MIG-03.1 read-only membership guard, _v2')
    parser.add_argument('--stage',required=True)
    parser.add_argument('--membership',required=True)
    parser.add_argument('--out',required=True)
    args=parser.parse_args(argv)
    stage=Path(args.stage).expanduser().resolve()
    member=Path(args.membership).expanduser().resolve()
    dest=Path(args.out).expanduser().resolve()
    if not stage.is_dir() or not member.is_dir():
        raise AuditError('Stage and membership directories must exist')
    if dest.exists() or any(a==dest or dest.is_relative_to(a) or a.is_relative_to(dest) for a in (stage,member)):
        raise AuditError('Output must be fresh and independent of inputs')
    provisional,report=inspect(stage,member)
    dest.mkdir(parents=True,exist_ok=False)
    try:
        for name,rows in provisional.items():
            write_private_jsonl(dest/(name+'.jsonl'),rows)
        (dest/'agrix_membership_guard_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        lines=['AGRIX MIG-03.1 - ACTIVE MEMBERSHIP GUARD / LEGACY _v2','READ ONLY - NO XLSM CREATED',
               'Source SHA256: '+report['source_sha256'],'MIG-03 original issues: '+str(report['original_mig03_issue_count']),
               '', 'CANDIDATE STORNO SAFETY:']
        for name,info in report['candidate_safety'].items():
            lines.extend(['  '+name+': original=%d provisional_active=%d historical_review=%d'%
                         (info['original_count'],info['provisional_active_count'],info['historical_excluded_from_active_only'])])
            for itm in info['historical_review'][:30]: lines.append('    '+json.dumps(itm,ensure_ascii=False))
        for k,label in [('otp_reconciliation','OTPREMNICA TRIAGE'),('zbr_reconciliation','ZBIRNA TRIAGE'),
                        ('receipt_reconciliation','PRIJEMNICA TRIAGE')]:
            val=report[k]
            lines.extend(['',label])
            for reason,n in sorted(val['counts'].items()):lines.append('  %-42s %5d'%(reason,n))
            lines.append('  Review row ordinals:')
            for itm in val['review'][:60]:lines.append('    '+json.dumps(itm,ensure_ascii=False))
            if len(val['review'])>60:lines.append('    ... %d more in JSON'%(len(val['review'])-60))
        lines+=['','STOP: PROVISIONAL_ACTIVE__ files are PRIVATE and NOT approved for XLSM import.',
                'Share ONLY agrix_membership_guard_summary.json and .txt.']
        (dest/'agrix_membership_guard_summary.txt').write_text('\n'.join(lines)+'\n',encoding='utf-8')
        print('\n'.join(lines[:31]));print('Reports saved:',dest)
    except BaseException:
        for f in dest.glob('*'):f.unlink(missing_ok=True)
        dest.rmdir()
        raise
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (OSError,AuditError,ValueError,KeyError,json.JSONDecodeError) as e:
        print('MIG-03.1 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
