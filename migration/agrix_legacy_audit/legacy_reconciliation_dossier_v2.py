#!/usr/bin/env python3
"""MIG-03.2 forensic review dossier, PINNED AgriX legacy _v2.

Reads local private MIG-02 stage and MIG-03.1 active guard. Emits PRIVATE CSV
for manual comparison and a PUBLIC row-ordinal-only summary. Neither import
nor automatic relinking nor Excel execution. Never upload CSV to chat/Git.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
import csv
from decimal import Decimal
import json
import os
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_stage_v2 import SOURCE_SHA
from legacy_membership_v2 import get_rows, index_ids, dec, diff_fields, eqkey, quantity_by_active_state
from legacy_membership_guard_v2 import read_json, ratio_band, sum_qty

VERSION = 'MIG-03.2'
DATASETS = ('SOURCE__tblOtkup','SOURCE__tblOtpremnica','SOURCE__tblZbirna','tblPrijemnica')
MEMBERSHIPS = (('PROVISIONAL_ACTIVE__tblOtpremnicaIzvori','OtpremnicaID','OtkupID','SOURCE__tblOtpremnica','SOURCE__tblOtkup'),
               ('PROVISIONAL_ACTIVE__tblZbirnaIzvori','ZbirnaID','OtpremnicaID','SOURCE__tblZbirna','SOURCE__tblOtpremnica'))


def sum_nullable(items, field):
    return sum_qty(items,field)


def csv_text(v):
    """Excel will not execute source-controlled strings as spreadsheet formulas."""
    s=str(v) if v is not None else ''
    if s and (s[0] in '=+@\t\r\n' or (s[0]=='-' and not _numeric(s))):
        return "'"+s
    return s


def _numeric(s):
    try:return Decimal(s).is_finite()
    except Exception:return False


def private_csv(path:Path, rows:list[dict], fields:list[str]):
    fd=os.open(str(path),os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
    with os.fdopen(fd,'w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=fields,delimiter=',',extrasaction='ignore')
        writer.writeheader()
        for r in rows:writer.writerow({c:csv_text(r.get(c,'')) for c in fields})


def make_report(stage:Path,guard:Path):
    mf=read_json(stage/'agrix_stage_summary.json')
    gd=read_json(guard/'agrix_membership_guard_summary.json')
    if mf.get('format')!='MIG-02.0' or mf.get('status')!='STAGED_ONLY_NO_XLSM_CREATED':
        raise AuditError('Expected MIG-02 stage')
    if gd.get('format')!='MIG-03.1' or gd.get('status')!='READ_ONLY_ACTIVE_CANDIDATES_PROVISIONAL_NOT_IMPORT_AUTHORIZATION':
        raise AuditError('Expected MIG-03.1 provisional guard')
    if mf.get('source_sha256')!=SOURCE_SHA or gd.get('source_sha256')!=SOURCE_SHA:
        raise AuditError('Source SHA mismatch -- no mixed workbooks')
    if mf.get('schema_sha256')!=gd.get('stage_schema_sha256'):
        raise AuditError('Schema SHA mismatch')
    rows={table:get_rows(stage,table) for table in DATASETS}
    for t,rs in rows.items():
        if len(rs)!=mf['source_row_counts'][t.removeprefix('SOURCE__')]:
            raise AuditError('Stage table count differs: '+t)
    otk,otp,zbr,prij=(rows[t] for t in DATASETS)
    otki=index_ids(otk,'OtkupID','otkup');otpi=index_ids(otp,'OtpremnicaID','otp');zbri=index_ids(zbr,'ZbirnaID','zbr')
    indexes={'SOURCE__tblOtkup':otki,'SOURCE__tblOtpremnica':otpi,'SOURCE__tblZbirna':zbri}
    # Independent reconstruction of the ACTIVE candidate totals. Cannot accept
    # a forged summary count or a CSV produced from a different staging graph.
    verified_parent_counts={}
    verified_parent_ids={}
    for fname,pkey,ckey,parent_table,child_table in MEMBERSHIPS:
        cand=get_rows(guard,fname)
        expected=gd['candidate_safety']['CANDIDATE__'+fname.removeprefix('PROVISIONAL_ACTIVE__')]['provisional_active_count']
        if len(cand)!=expected:raise AuditError('Guard candidate count mismatch: '+fname)
        byparent=defaultdict(list);seen=set();seen_child=set()
        for link in cand:
            parent_id=trim(link.get(pkey,''));child_id=trim(link.get(ckey,''))
            if not parent_id or not child_id or (parent_id,child_id) in seen:
                raise AuditError('Bad/duplicated provisional link')
            seen.add((parent_id,child_id))
            if child_id in seen_child:raise AuditError('A child has >1 candidate parent')
            seen_child.add(child_id)
            if parent_id not in indexes[parent_table] or child_id not in indexes[child_table]:
                raise AuditError('Dangling provisional parent/child')
            p=indexes[parent_table][parent_id][1];c=indexes[child_table][child_id][1]
            if not is_active(p) or not is_active(c):raise AuditError('Provisional candidate has storno row')
            if parent_table=='SOURCE__tblOtpremnica':
                if diff_fields(c,p,('StanicaID','VrstaVoca','SortaVoca','TipAmbalaze','Klasa')):
                    raise AuditError('Provisional OTP link context mismatch')
            else:
                if not (eqkey(c.get('BrojZbirne'),p.get('BrojZbirne')) and
                        eqkey(c.get('VozacID'),p.get('VozacID')) and
                        eqkey(c.get('Klasa'),p.get('Klasa')) and
                        not diff_fields(c,p,('VrstaVoca','SortaVoca','TipAmbalaze'))):
                    raise AuditError('Provisional ZBR link context mismatch')
            byparent[parent_id].append(c)
        pa_field='Kolicina' if parent_table=='SOURCE__tblOtpremnica' else 'UkupnoKolicina'
        for pid,children in byparent.items():
            p=indexes[parent_table][pid][1]
            if quantity_by_active_state(p,children,pa_field,'Kolicina')!='OK':
                raise AuditError('Provisional candidate amount reconciliation failed')
        verified_parent_counts[parent_table]=len(byparent)
        verified_parent_ids[parent_table]=set(byparent)
    if verified_parent_counts['SOURCE__tblOtpremnica']!=gd['otp_reconciliation']['counts'].get('VERIFIED',0):
        raise AuditError('OTP verified-header count mismatch')
    if verified_parent_counts['SOURCE__tblZbirna']!=gd['zbr_reconciliation']['counts'].get('VERIFIED',0):
        raise AuditError('ZBR verified-header count mismatch')

    # End-to-end chain: a verified Zbirna->Otpremnica edge is NOT automatically
    # a verified Zbirna->Otpremnica->Otkup chain. Enumerate unresolved upstream.
    otp_reason_by_ordinal={v['otp_row']:v['reason'] for v in gd['otp_reconciliation']['review']}
    transitive=Counter(); upstream_dossier=[]
    for link in get_rows(guard,'PROVISIONAL_ACTIVE__tblZbirnaIzvori'):
        oid=trim(link['OtpremnicaID']); zid=trim(link['ZbirnaID'])
        if oid in verified_parent_ids['SOURCE__tblOtpremnica']:
            transitive['ZBR_LINK_WITH_VERIFIED_OTKUP_CHAIN']+=1
        else:
            transitive['ZBR_LINK_WITH_UNVERIFIED_OTKUP_CHAIN']+=1
            ono,orow=otpi[oid];zno,zrow=zbri[zid]
            reason=otp_reason_by_ordinal.get(ono,'NOT_CLASSIFIED')
            transitive['UPSTREAM_'+reason]+=1
            upstream_dossier.append({'zbirna_row':zno,'otp_row':ono,
                'reason':reason,'ZbirnaID':zid,'OtpremnicaID':oid,
                'BrojZbirne':trim(zrow.get('BrojZbirne')),
                'BrojOtpremnice':trim(orow.get('BrojOtpremnice')),
                'OtpremnicaKg':trim(orow.get('Kolicina')),
                'operator_decision':'REVIEW_REQUIRED','operator_note':''})
    if (transitive['ZBR_LINK_WITH_VERIFIED_OTKUP_CHAIN']+
        transitive['ZBR_LINK_WITH_UNVERIFIED_OTKUP_CHAIN'] !=
            gd['candidate_safety']['CANDIDATE__tblZbirnaIzvori']['provisional_active_count']):
        raise AuditError('ZBR chain coverage incomplete')

    byotp=defaultdict(list)
    bynumber=defaultdict(list)
    for ordinal,c in enumerate(otk,1):
        if trim(c.get('OtpremnicaID')):byotp[trim(c['OtpremnicaID'])].append((ordinal,c))
        if trim(c.get('BrojOtpremnice')):bynumber[trim(c['BrojOtpremnice']).casefold()].append((ordinal,c))
    otp_dossier=[]
    missing_alt=Counter();qband=Counter();mismatch_rows=[];none_rows=[]
    for issue in gd['otp_reconciliation']['review']:
        ordinal=issue['otp_row'];p=otp[ordinal-1];pid=trim(p.get('OtpremnicaID'))
        raw=byotp.get(pid,[])
        good=[(i,r) for i,r in raw if not diff_fields(r,p,('StanicaID','VrstaVoca','SortaVoca','TipAmbalaze','Klasa'))]
        active=[r for _,r in good if is_active(r)];storno=[r for _,r in good if not is_active(r)]
        pqty=dec(p.get('Kolicina'));aqty=sum_nullable(active,'Kolicina');histqty=sum_nullable(storno,'Kolicina')
        diff=(aqty-pqty) if aqty is not None and pqty is not None else None
        pct=(abs(diff)/pqty*100) if diff is not None and pqty and pqty>0 else None
        other=[]
        if issue['reason']=='NO_VERIFIED_OTKUP_SOURCE':
            for i,c in bynumber.get(trim(p.get('BrojOtpremnice')).casefold(),[]):
                if (eqkey(c.get('StanicaID'),p.get('StanicaID')) and
                    eqkey(c.get('Klasa'),p.get('Klasa'))):
                    other.append((i,c))
            missing_alt['HAS_NUMBER_STATION_CLASS_MATCH' if other else 'NO_NUMBER_STATION_CLASS_MATCH']+=1
            if other:
                missing_alt['ACTIVE_MATCHES']+=sum(is_active(c) for _,c in other)
                missing_alt['STORNO_MATCHES']+=sum(not is_active(c) for _,c in other)
            none_rows.append(ordinal)
        if issue['reason']=='QTY_MISMATCH':
            qband[issue['relative_difference_bucket']]+=1
            mismatch_rows.append(ordinal)
        otp_dossier.append({
            'otp_row':ordinal,'reason':issue['reason'],'OtpremnicaID':pid,
            'BrojOtpremnice':trim(p.get('BrojOtpremnice')),
            'BrojZbirne':trim(p.get('BrojZbirne')),
            'StanicaID':trim(p.get('StanicaID')),
            'Klasa':trim(p.get('Klasa')),
            'SortaVoca':trim(p.get('SortaVoca')),
            'otp_stornirana':not is_active(p),
            'otp_kolicina_kg':pqty,
            'active_source_kolicina_kg':aqty,
            'storno_source_kolicina_kg':histqty,
            'active_minus_otp_kg':diff,
            'abs_difference_pct':round(pct,6) if pct is not None else '',
            'direct_otkup_rows':';'.join(str(i) for i,_ in raw),
            'valid_active_rows':';'.join(str(i) for i,r in good if is_active(r)),
            'valid_storno_rows':';'.join(str(i) for i,r in good if not is_active(r)),
            'rejected_context_rows':';'.join(str(i) for i,c in raw if (i,c) not in good),
            'number_scope_alt_row_count':len(other),
            'number_scope_alt_rows':';'.join(str(i) for i,_ in other),
            'number_scope_alt_active':sum(is_active(r) for _,r in other),
            'number_scope_alt_kg':sum_nullable([r for _,r in other if is_active(r)],'Kolicina'),
            'operator_decision':'REVIEW_REQUIRED',
            'operator_note':''
        })
    if len(otp_dossier)!=len(gd['otp_reconciliation']['review']):
        raise AuditError('OTP review count drift')
    zbr_dossier=[]
    for item in gd['zbr_reconciliation']['review']:
        ordinal=item['zbirna_row'];z=zbr[ordinal-1]
        zbr_dossier.append({'zbirna_row':ordinal,'reason':item['reason'],
                            'ZbirnaID':trim(z.get('ZbirnaID')),'BrojZbirne':trim(z.get('BrojZbirne')),
                            'VozacID':trim(z.get('VozacID')),'KupacID':trim(z.get('KupacID')),
                            'Stornirano':trim(z.get('Stornirano')),'UkupnoKolicina':trim(z.get('UkupnoKolicina')),
                            'operator_decision':'REVIEW_REQUIRED','operator_note':''})
    zb_index_by_num=defaultdict(list)
    for zidx,z in enumerate(zbr,1):
        n=trim(z.get('BrojZbirne'))
        if n:zb_index_by_num[n.casefold()].append((zidx,z))
    prij_dossier=[]
    for item in gd['receipt_reconciliation']['review']:
        ordinal=item['prijemnica_row'];p=prij[ordinal-1]
        raw=trim(p.get('BrojZbirne'));zopts=zb_index_by_num.get(raw.casefold(),[])
        matched=[(i,z) for i,z in zopts if eqkey(z.get('VozacID'),p.get('VozacID'))
                 and eqkey(z.get('KupacID'),p.get('KupacID')) and eqkey(z.get('Klasa'),p.get('Klasa'))]
        prij_dossier.append({'prijemnica_row':ordinal,'reason':item['reason'],
                             'PrijemnicaID':trim(p.get('PrijemnicaID')),
                             'BrojPrijemnice':trim(p.get('BrojPrijemnice')),
                             'BrojZbirne':raw,'KupacID':trim(p.get('KupacID')),
                             'VozacID':trim(p.get('VozacID')),
                             'Klasa':trim(p.get('Klasa')),
                             'prijemnica_stornirana':not is_active(p),
                             'casefold_candidates_rows':';'.join(str(i) for i,_ in zopts),
                             'exact_identity_rows':';'.join(str(i) for i,_ in matched),
                             'operator_decision':'REVIEW_REQUIRED','operator_note':''})
    context_dossier=[]
    for ordinal,c in enumerate(otk,1):
        pid=trim(c.get('OtpremnicaID'))
        if not pid or pid not in otpi:continue
        oidx,p=otpi[pid];diff=diff_fields(c,p,('StanicaID','VrstaVoca','SortaVoca','TipAmbalaze','Klasa'))
        if not diff:continue
        context_dossier.append({'otkup_row':ordinal,'otp_row':oidx,
            'OtkupID':trim(c.get('OtkupID')),'OtpremnicaID':pid,
            'mismatch_fields':';'.join(diff),
            'otkup_sorta':trim(c.get('SortaVoca')),'otp_sorta':trim(p.get('SortaVoca')),
            'otkup_vrsta':trim(c.get('VrstaVoca')),'otp_vrsta':trim(p.get('VrstaVoca')),
            'otkup_klasa':trim(c.get('Klasa')),'otp_klasa':trim(p.get('Klasa')),
            'operator_decision':'REVIEW_REQUIRED','operator_note':''})
    public={
        'format':VERSION,
        'status':'PRIVATE_OPERATOR_DOSSIER_NOT_IMPORT_AUTHORIZATION',
        'source_sha256':SOURCE_SHA,
        'schema_sha256':mf['schema_sha256'],
        'verified_candidate_integrity':{
            'otp_active_candidate_links':gd['candidate_safety']['CANDIDATE__tblOtpremnicaIzvori']['provisional_active_count'],
            'zbr_active_candidate_links':gd['candidate_safety']['CANDIDATE__tblZbirnaIzvori']['provisional_active_count'],
            'otp_headers_rechecked':verified_parent_counts['SOURCE__tblOtpremnica'],
            'zbr_headers_rechecked':verified_parent_counts['SOURCE__tblZbirna'],
        },
        'review_case_counts':{
            'otpremnica':len(otp_dossier),
            'zbirna':len(zbr_dossier),
            'prijemnica':len(prij_dossier),
            'otkup_context_mismatches':len(context_dossier),
            'zbr_upstream_review':len(upstream_dossier)
        },
        'unlinked_otp_number_suggestions_not_approved':dict(missing_alt),
        'transitive_chain':dict(transitive),
        'quantity_mismatch_buckets':dict(qband),
        'quantity_mismatch_otp_ordinals':mismatch_rows,
        'no_verified_source_otp_ordinals':none_rows,
        'receipt_casefold_ordinal':next((r['prijemnica_row'] for r in prij_dossier if r['reason']=='CASEFOLD_MATCH_NOT_EXACT'),None),
        'notes':['Sensitive review CSVs stay only on local Windows machine and must not be uploaded.',
                 'Number-based alternatives are suggestions to investigate, NOT evidence of membership.',
                 'Do not correct quantities by applying a tolerance or overwrite historical records.',
                 'No data copied to XLSM or changed, no membership approved.']
    }
    return public, {'otpremnica_review_PRIVATE.csv':otp_dossier,
                    'zbirna_storno_review_PRIVATE.csv':zbr_dossier,
                    'prijemnica_review_PRIVATE.csv':prij_dossier,
                    'otkup_context_review_PRIVATE.csv':context_dossier,
                    'zbirna_upstream_review_PRIVATE.csv':upstream_dossier}


def cli(argv=None):
    p=argparse.ArgumentParser(description='MIG-03.2 private operator review dossier, read-only')
    p.add_argument('--stage',required=True)
    p.add_argument('--guard',required=True)
    p.add_argument('--out',required=True)
    a=p.parse_args(argv)
    stage=Path(a.stage).expanduser().resolve();guard=Path(a.guard).expanduser().resolve();out=Path(a.out).expanduser().resolve()
    if not stage.is_dir() or not guard.is_dir():raise AuditError('Missing stage or guard directory')
    if out.exists() or any(out==x or out.is_relative_to(x) or x.is_relative_to(out) for x in (stage,guard)):
        raise AuditError('Output must be fresh independent directory')
    public,private=make_report(stage,guard)
    out.mkdir(parents=True,exist_ok=False)
    try:
        for fname,rows in private.items():
            fields=list(rows[0]) if rows else ['source_row','operator_decision','operator_note']
            private_csv(out/fname,rows,fields)
        (out/'agrix_reconciliation_dossier_summary.json').write_text(json.dumps(public,ensure_ascii=False,indent=2),encoding='utf-8')
        lines=['AGRIX MIG-03.2 -- HISTORICAL MEMBERSHIP FORENSIC DOSSIER','READ ONLY, NOT IMPORT AUTHORIZATION',
               'Source SHA256: '+SOURCE_SHA,'', 'VERIFIED ACTIVE CANDIDATE INTEGRITY']
        for k,v in public['verified_candidate_integrity'].items():lines.append(f'  {k}: {v}')
        lines.append('');lines.append('PRIVATE REVIEW CSV ROW COUNTS')
        for k,v in public['review_case_counts'].items():lines.append(f'  {k}: {v}')
        lines.append('');lines.append('NO EXACT OTP SOURCE: BUSINESS-NUMBER ALTERNATIVES (UNPROVEN)')
        for k,v in sorted(public['unlinked_otp_number_suggestions_not_approved'].items()):lines.append(f'  {k}: {v}')
        lines.append('');lines.append('ZBIRNA TRANSITIVE CHAIN')
        for k,v in sorted(public['transitive_chain'].items()):lines.append(f'  {k}: {v}')
        lines.append('');lines.append('QTY MISMATCH BUCKETS')
        for k,v in sorted(public['quantity_mismatch_buckets'].items()):lines.append(f'  {k}: {v}')
        lines+=['','NOTE: Private CSVs contain identifiers, numbers and amounts. Keep local, do NOT upload.',
                'STOP: NO XLSM modified or created; no link inferred/approved.']
        (out/'agrix_reconciliation_dossier_summary.txt').write_text('\n'.join(lines)+'\n',encoding='utf-8')
        print('\n'.join(lines));print('Destination:',out)
    except BaseException:
        for f in out.glob('*'):f.unlink(missing_ok=True)
        out.rmdir()
        raise
    return 0


if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,json.JSONDecodeError) as ex:
        print('MIG-03.2 STOP:',ex,file=sys.stderr)
        raise SystemExit(2)
