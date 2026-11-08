#!/usr/bin/env python3
"""MIG21: source-grounded document linkage repair for MALINA_MODE.

Run from ~/Desktop/AgriX DEV/agrix_legacy_audit on a CLEAN MIG19.13 xlsm.
Default: read-only diagnostic and precise repair-plan. --apply only applies
strict 1:1 verified links to a NEW macro-enabled copy using local Microsoft Excel.
NO statuses, quantities, palettes, packaging, or original workbook ever changed.
"""
from __future__ import annotations
import argparse, collections, csv, hashlib, json, os, shutil, subprocess, sys, zipfile
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from agrix_mig19_malina_readonly_audit import XlsmReadOnly, dnum, ddate, key, norm, sha256

VERSION='AGRIx-MIG21-DOC-LINK-REPAIR-v1'
ROOT=Path('..')
DEFAULT=ROOT/'MigrationMIG19_13_PatchedXLSM'/'AgriX_MIG19_13_AMBALAZA_QA_PARTIAL.xlsm'
TABS=['tblOtkup','tblOtkupStavke','tblOtpremnica','tblOtpremnicaStavke','tblOtpremnicaIzvori','tblZbirna','tblZbirnaIzvori','tblPrijemnica']
Q0=Decimal('0');CENT=Decimal('0.01'); VOID={'DA','TRUE','1'}

def void(r):return key(r.get('Stornirano')) in VOID

def rounded(x):
    v=dnum(x)
    return v.quantize(CENT) if v is not None else None

def amounts(rows,owner_col):
    by=collections.defaultdict(list)
    for r in rows:by[key(r.get(owner_col))].append(r)
    return by

def sum_col(rows,col):
    nums=[rounded(r.get(col)) for r in rows]
    return sum((v for v in nums if v is not None),Q0)

def signature_match(otk,otp,stage_otk,stage_otp,otk_lines,otp_lines):
    """Strict 1:1: all business identity and each class's precise quantity and packaging."""
    checks={}
    for label,a,b in [('number',otk.get('BrojDokumenta'),otp.get('BrojOtpremnice')),
                      ('date',ddate(otk.get('Datum')),ddate(otp.get('Datum'))),
                      ('station',otk.get('StanicaID'),otp.get('StanicaID')),
                      ('fruit',otk.get('VrstaVoca'),otp.get('VrstaVoca')),
                      ('variety',otk.get('SortaVoca'),otp.get('SortaVoca')),
                      ('packaging',otk.get('TipAmbalaze'),otp.get('TipAmbalaze'))]:
        checks[label]=bool(key(a)) and key(a)==key(b)
    # Raw SOURCE facts must agree with canonical headers (not the old corrupted backlink).
    for label,col in [('number','BrojDokumenta'),('date','Datum'),('station','StanicaID'),
                      ('fruit','VrstaVoca'),('variety','SortaVoca'),('packaging','TipAmbalaze')]:
        old=(ddate(stage_otk.get(col)) if label=='date' else key(stage_otk.get(col)))
        now=(ddate(otk.get(col)) if label=='date' else key(otk.get(col)))
        checks['raw_otk_'+label]=bool(old) and old==now
    for label,col in [('number','BrojOtpremnice'),('date','Datum'),('station','StanicaID'),
                      ('fruit','VrstaVoca'),('variety','SortaVoca'),('packaging','TipAmbalaze')]:
        old=(ddate(stage_otp.get(col)) if label=='date' else key(stage_otp.get(col)))
        now=(ddate(otp.get(col)) if label=='date' else key(otp.get(col)))
        checks['raw_otp_'+label]=bool(old) and old==now
    def clas(rows):
        d=collections.defaultdict(lambda:[Q0,Q0]); ok=True
        for r in rows:
            c=key(r.get('Klasa'))
            v=rounded(r.get('Kolicina')); a=rounded(r.get('KolAmbalaze'))
            if not c or v is None or a is None: ok=False; continue
            d[c][0]+=v;d[c][1]+=a
        return {c:(x[0],x[1]) for c,x in d.items()},ok
    a,avalid=clas(otk_lines);b,bvalid=clas(otp_lines)
    checks['line_items']=bool(otk_lines and otp_lines and avalid and bvalid and a==b)
    return checks

def read_stage(stage,name):
    stage=Path(stage)
    possibilities=list(stage.rglob(name+'.jsonl'))
    # Prefer an exact root file, otherwise demand a unique nested canonical stage file.
    direct=stage/(name+'.jsonl')
    if direct.is_file():path=direct
    elif len(possibilities)==1:path=possibilities[0]
    elif not possibilities:raise FileNotFoundError(f'Missing original stage file: {name}.jsonl inside {stage}')
    else:raise ValueError(f'Ambiguous stage files for {name}: {possibilities}')
    pk={'tblOtkup':'OtkupID','tblOtpremnica':'OtpremnicaID'}[name]
    obj={}
    for lineno,line in enumerate(path.open(encoding='utf-8-sig'),1):
        if not line.strip():continue
        r=json.loads(line)
        if not isinstance(r,dict) or not key(r.get(pk)) or key(r[pk]) in obj:
            raise ValueError(f'Stage corrupt/duplicate {path} line {lineno}')
        obj[key(r[pk])]=r
    return obj

def read_xlsm(path):
    book=XlsmReadOnly(path)
    try:return {t:book.read_table(t,set()) for t in TABS}
    finally:book.close()

def pkmap(rows,k):
    counts=collections.Counter(key(r.get(k)) for r in rows)
    bad={name:n for name,n in counts.items() if n>1 or not name}
    return {key(r[k]):r for r in rows if key(r.get(k))},bad

def value_pairs(T):
    otks,_=pkmap(T['tblOtkup'],'OtkupID')
    otps,_=pkmap(T['tblOtpremnica'],'OtpremnicaID')
    by_otp=amounts(T['tblOtpremnicaIzvori'],'OtpremnicaID')
    by_otk=amounts(T['tblOtpremnicaIzvori'],'OtkupID')
    otkli=amounts(T['tblOtkupStavke'],'OtkupID')
    otpli=amounts(T['tblOtpremnicaStavke'],'OtpremnicaID')
    return otks,otps,by_otp,by_otk,otkli,otpli

def build(path,stage):
    T=read_xlsm(path)
    otks,otps,by_otp,by_otk,otkli,otpli=value_pairs(T)
    stage_otk=read_stage(stage,'tblOtkup');stage_otp=read_stage(stage,'tblOtpremnica')
    blockers=[];incons=[];cases=[];repairs=[]
    for name,k in [('tblOtkup','OtkupID'),('tblOtpremnica','OtpremnicaID'),
                   ('tblOtpremnicaIzvori','OtpremnicaIzvorID'),('tblZbirna','ZbirnaID'),
                   ('tblZbirnaIzvori','ZbirnaIzvorID')]:
        _,bad=pkmap(T[name],k)
        if bad:blockers.append(f'DUPLICATE_OR_BLANK_KEY {name}: {bad}')
    for id in ('ZBR-00666','ZBR-00667','ZBR-00668'):
        # Source workbook was accidentally edited once; demand exact stage coverage here.
        if id in {key(x.get('ZbirnaID')) for x in T['tblZbirna']}:
            continue
        blockers.append('MISSING_STAGE_ZBR: '+id)
    all_otp_numbers=collections.defaultdict(list)
    for pid,o in otps.items():
        if not void(o):all_otp_numbers[key(o.get('BrojOtpremnice'))].append(pid)
    all_otk_numbers=collections.defaultdict(list)
    for oid,o in otks.items():
        if not void(o):all_otk_numbers[key(o.get('BrojDokumenta'))].append(oid)
    for pid,otp in sorted(otps.items()):
        if void(otp):continue
        source=by_otp.get(pid,[])
        active=[key(z.get('OtkupID')) for z in source if key(z.get('OtkupID')) in otks and not void(otks[key(z.get('OtkupID'))])]
        p_kg=sum_col(otpli.get(pid,[]),'Kolicina');p_amb=sum_col(otpli.get(pid,[]),'KolAmbalaze')
        q_kg=sum((sum_col(otkli.get(oid,[]),'Kolicina') for oid in active),Q0)
        q_amb=sum((sum_col(otkli.get(oid,[]),'KolAmbalaze') for oid in active),Q0)
        dk=p_kg-q_kg;da=p_amb-q_amb
        class_ = ('NO_ACTIVE_OTKUP' if not active else 'KG_AMB_DIFF' if dk or da else 'OK')
        if class_=='OK':continue
        case=dict(category=class_,OTP=pid,broj=norm(otp.get('BrojOtpremnice')),
                  date=ddate(otp.get('Datum')),station=norm(otp.get('StanicaID')),
                  variety=norm(otp.get('SortaVoca')),status=norm(otp.get('IzdatoStatus')),
                  existing_links=len(source),active_sources=','.join(active),
                  otp_kg=str(p_kg),source_kg=str(q_kg),difference_kg=str(dk),
                  otp_gajbe=str(p_amb),source_gajbe=str(q_amb),difference_gajbe=str(da),
                  result='REVIEW',reason='No strict unique 1:1 proof')
        candidates=[]
        business=key(otp.get('BrojOtpremnice'))
        if not active and pid not in {'OTP-00666'} and business and len(all_otp_numbers.get(business,[]))==1:
            # ONLY SAME exact historical business number; no approximate or sequence matching.
            for oid in all_otk_numbers.get(business,[]):
                otk=otks[oid]
                if oid not in stage_otk or pid not in stage_otp:continue
                checks=signature_match(otk,otp,stage_otk[oid],stage_otp[pid],otkli.get(oid,[]),otpli.get(pid,[]))
                if not all(checks.values()):continue
                links_for_otk=by_otk.get(oid,[])
                active_existing=[x for x in links_for_otk if key(x.get('OtpremnicaID')) in otps and not void(otps[key(x.get('OtpremnicaID'))])]
                if active_existing:continue
                # If OPI exists, it must be a UNIQUE wrong link to a VOID OTP.
                if len(links_for_otk)>1:continue
                link=(links_for_otk[0] if links_for_otk else None)
                if link and (key(link.get('OtpremnicaID')) not in otps or not void(otps[key(link.get('OtpremnicaID'))])):continue
                old_backlink=key(otk.get('OtpremnicaID'))
                if old_backlink and old_backlink!=pid and (old_backlink not in otps or not void(otps[old_backlink])):continue
                operation='MOVE' if link else 'ADD'
                candidates.append((oid,link,operation,old_backlink,checks))
        if len(candidates)==1:
            oid,link,operation,old_backlink,checks=candidates[0]
            # Do not silently change historic link if the source was already active elsewhere.
            used_src=key(link.get('OtpremnicaID')) if link else ''
            linkid=key(link.get('OtpremnicaIzvorID')) if link else 'OPI-M21-'+hashlib.sha256(f'{pid}|{oid}'.encode()).hexdigest()[:22].upper()
            proposed=dict(action=operation,OtpremnicaIzvorID=linkid,OtkupID=oid,
                          old_OtpremnicaID=used_src,new_OtpremnicaID=pid,
                          old_Otkup_Backlink=old_backlink,expected_broj=norm(otp.get('BrojOtpremnice')),
                          source_evidence='number+date+station+fruit+variety+packaging+all_classes_kg_gajbe+source_stage',
                          number=norm(otp.get('BrojOtpremnice')))
            repairs.append(proposed);case.update(result='AUTO_ELIGIBLE',reason=operation+' exact unique source '+oid)
        elif len(candidates)>1:case.update(reason=f'{len(candidates)} equally matching candidates: unsafe')
        elif pid=='OTP-00666':case.update(result='HISTORICAL_STORNO_CASE',reason='Known manually investigated cancellation; never auto-add an OTK')
        cases.append(case)
    # Pairwise residual repair: only an exact unique OTK row whose removal from one
    # OTP and attachment to another makes BOTH kg AND packaging totals exact.
    # This requires ORIGINAL-staged business identity, and an independent target
    # document affinity (BrojZbirne or BrojOtpremnice), not arithmetic alone.
    case_idx={x['OTP']:x for x in cases}
    moved_already={r['OtkupID'] for r in repairs}
    target_claims=collections.defaultdict(list)
    pair_candidates=[]
    zbr_by_otp=collections.defaultdict(list)
    zbrs={key(z.get('ZbirnaID')):z for z in T['tblZbirna']}
    for zlink in T['tblZbirnaIzvori']:
        zid=key(zlink.get('ZbirnaID'));pid=key(zlink.get('OtpremnicaID'))
        if zid in zbrs and not void(zbrs[zid]):zbr_by_otp[pid].append(zbrs[zid])
    for sid,src_case in case_idx.items():
        if src_case['category']!='KG_AMB_DIFF':continue
        src_kg=Decimal(src_case['difference_kg']);src_amb=Decimal(src_case['difference_gajbe'])
        if not(src_kg<0 and src_amb<0):continue
        for tid,tgt_case in case_idx.items():
            if tid==sid or tgt_case['category']!='KG_AMB_DIFF':continue
            tgt_kg=Decimal(tgt_case['difference_kg']);tgt_amb=Decimal(tgt_case['difference_gajbe'])
            if (src_kg+tgt_kg)!=0 or (src_amb+tgt_amb)!=0:continue
            src=otps[sid];tgt=otps[tid]
            if void(src) or void(tgt):continue
            if tid not in stage_otp or sid not in stage_otp:continue
            if len(zbr_by_otp.get(tid,[]))!=1:continue
            # Raw-document headers unchanged, and station/fruit/variety/packaging
            # equal on BOTH sides; cross-station or cross-sort pairing forbidden.
            facts=['StanicaID','VrstaVoca','SortaVoca','TipAmbalaze']
            if any(key(src.get(x))!=key(tgt.get(x)) or not key(src.get(x)) for x in facts):continue
            if any(key(src.get(x))!=key(stage_otp[sid].get(x)) or
                   key(tgt.get(x))!=key(stage_otp[tid].get(x)) for x in facts):continue
            for link in by_otp.get(sid,[]):
                oid=key(link.get('OtkupID'))
                if oid in moved_already or oid not in otks or oid not in stage_otk:continue
                o=otks[oid]
                if void(o) or len(by_otk.get(oid,[]))!=1:continue
                # Distinct source attribution: membership was not altered by
                # this migration; the raw OTK business numbers must agree.
                raw=stage_otk[oid]
                if any(key(o.get(x))!=key(raw.get(x)) for x in facts+['BrojDokumenta']):continue
                if ddate(o.get('Datum'))!=ddate(tgt.get('Datum')):continue
                if any(key(o.get(x))!=key(tgt.get(x)) for x in facts):continue
                if not raw.get('BrojZbirne') and not raw.get('BrojOtpremnice'):continue
                matched_affinity=(key(raw.get('BrojZbirne'))==key(zbr_by_otp[tid][0].get('BrojZbirne')) or
                                  key(raw.get('BrojOtpremnice'))==key(tgt.get('BrojOtpremnice')))
                if not matched_affinity:continue
                if sum_col(otkli.get(oid,[]),'Kolicina')!=tgt_kg or sum_col(otkli.get(oid,[]),'KolAmbalaze')!=tgt_amb:continue
                backlink=key(o.get('OtpremnicaID'))
                if backlink and backlink not in (sid,tid):continue
                pair_candidates.append((sid,tid,oid,link,backlink))
                target_claims[(sid,tid)].append(oid)
    # A pair is automatically fixable only if exactly one OTK can explain it,
    # and that OTK does not also participate in another pair candidate.
    candidate_claims=collections.Counter(oid for sid,tid,oid,link,b in pair_candidates)
    reserved=set()
    for sid,tid,oid,link,back in pair_candidates:
        if (sid,tid) in reserved:continue
        if len(target_claims[(sid,tid)])!=1 or candidate_claims[oid]!=1:continue
        rid=key(link.get('OtpremnicaIzvorID'))
        repairs.append(dict(action='MOVE',OtpremnicaIzvorID=rid,OtkupID=oid,
            old_OtpremnicaID=sid,new_OtpremnicaID=tid,old_Otkup_Backlink=back,
            expected_broj=norm(otps[tid].get('BrojOtpremnice')),
            source_evidence='exact_pair_residual_kg_gajbe+raw_stage+date_station_fruit_variety+unique_broj_zbirne_or_otp',
            number=norm(otps[tid].get('BrojOtpremnice'))))
        case_idx[sid].update(result='AUTO_ELIGIBLE_PAIR',reason=f'MOVE {oid} -> {tid} via exact documented pair')
        case_idx[tid].update(result='AUTO_ELIGIBLE_PAIR',reason=f'MOVE {oid} <- {sid} via exact documented pair')
        reserved.add((sid,tid))
    # No more than one new OTP per old Otkup; apply will refuse if ambiguity.
    used=collections.Counter(r['OtkupID'] for r in repairs)
    if any(n>1 for n in used.values()):blockers.append('SAME_OTK_TARGETED_MULTIPLE_TIMES')
    # Structured aggregate without fabricated corrections.
    zbrs,_=pkmap(T['tblZbirna'],'ZbirnaID');zbiby=amounts(T['tblZbirnaIzvori'],'ZbirnaID')
    for zid,z in zbrs.items():
        if void(z):continue
        linked=[key(v.get('OtpremnicaID')) for v in zbiby.get(zid,[])]
        if linked and not any(pid in otps and not void(otps[pid]) for pid in linked):
            incons.append(dict(kind='ACTIVE_ZBR_NO_ACTIVE_OTP',id=zid,others=','.join(linked)))
    prj=T['tblPrijemnica']
    by_zbr=collections.defaultdict(list)
    for zid,z in zbrs.items():by_zbr[key(z.get('BrojZbirne'))].append(z)
    for p in prj:
        if void(p):continue
        number=key(p.get('BrojZbirne'))
        found=by_zbr.get(number,[])
        if number and found and not any(not void(z) for z in found):
            incons.append(dict(kind='ACTIVE_PRJ_STORNO_ZBR',id=key(p.get('PrijemnicaID')),others=','.join(key(z['ZbirnaID']) for z in found)))
        elif number and not found:
            incons.append(dict(kind='ACTIVE_PRJ_ZBR_MISSING',id=key(p.get('PrijemnicaID')),others=number))
    # Existing stage stable statuses must not change except explicitly user-approved manual cases.
    result=dict(format=VERSION,status='BLOCKED_DIRTY_SOURCE' if blockers else 'PLAN_READY',
                input={'workbook':str(path.resolve()),'sha256':sha256(path),'stage':str(stage)},
                blocker_count=len(blockers),blockers=blockers,
                found_cases=len(cases),auto_eligible=len(repairs),manual_cases=len(cases)-len(repairs),
                missing_source=[c['OTP'] for c in cases if c['category']=='NO_ACTIVE_OTKUP'],
                qty_differences=[c['OTP'] for c in cases if c['category']=='KG_AMB_DIFF'],
                document_lifecycle=incons,
                note='No palette analysis, no automatic storno. No quantity edits.')
    return result,cases,repairs,T

def table_snap(T):
    return {name:[{k:v for k,v in row.items() if k!='_excel_row'} for row in rows] for name,rows in T.items()}

def check_readback(before,after,repairs):
    approved={(r['OtpremnicaIzvorID'],r['OtkupID'],r['new_OtpremnicaID']) for r in repairs}
    old_links={key(x['OtpremnicaIzvorID']):x for x in before['tblOtpremnicaIzvori']}
    new_links={key(x['OtpremnicaIzvorID']):x for x in after['tblOtpremnicaIzvori']}
    for r in repairs:
        rid=r['OtpremnicaIzvorID'];cur=new_links.get(rid)
        if not cur or key(cur.get('OtpremnicaID'))!=r['new_OtpremnicaID'] or key(cur.get('OtkupID'))!=r['OtkupID']:
            raise AssertionError('Repair readback failed: '+rid)
    for id,old in old_links.items():
        new=new_links.get(id)
        if not new:raise AssertionError('Existing OPI lost '+id)
        changed={c for c in old if c not in ('_excel_row',) and norm(old[c])!=norm(new.get(c))}
        if changed and not (id in {r['OtpremnicaIzvorID'] for r in repairs if r['action']=='MOVE'} and changed=={'OtpremnicaID'}):
            raise AssertionError(f'Unapproved OPI difference {id} {changed}')
    if len(new_links)!=len(old_links)+sum(r['action']=='ADD' for r in repairs):
        raise AssertionError('OPI count mismatch')
    old_otk={key(x['OtkupID']):x for x in before['tblOtkup']}
    new_otk={key(x['OtkupID']):x for x in after['tblOtkup']}
    for id,o in old_otk.items():
        n=new_otk.get(id)
        if not n:raise AssertionError('Lost otkup '+id)
        changed={c for c in o if c!='_excel_row' and norm(o[c])!=norm(n.get(c))}
        if changed:
            allowed={'OtpremnicaID','BrojOtpremnice'} if id in {r['OtkupID'] for r in repairs} else set()
            if not changed<=allowed:raise AssertionError(f'Changed forbidden OTK fields {id} {changed}')
    for name in TABS:
        if name in ('tblOtpremnicaIzvori','tblOtkup'):continue
        a=table_snap(before)[name];b=table_snap(after)[name]
        if a!=b:raise AssertionError(f'Other table data changed: {name}')

def all_other_tables_digest(path):
    book=XlsmReadOnly(path)
    try:
        names=sorted(n for n in book.tables if n not in {'tblOtpremnicaIzvori','tblOtkup'})
        digest=hashlib.sha256()
        for name in names:
            rows=book.read_table(name,set())
            for r in rows:
                r.pop('_excel_row',None)
            digest.update(name.encode('utf-8'))
            digest.update(json.dumps(rows,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf-8'))
        return names,digest.hexdigest()
    finally:book.close()

def save(out,report,cases,repair):
    out.mkdir(parents=True,exist_ok=True)
    (out/'MIG21_SUMMARY.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    (out/'MIG21_PLAN.json').write_text(json.dumps({'source_sha256':report['input']['sha256'],'operations':repair},ensure_ascii=False,indent=2),encoding='utf-8')
    cols=['category','OTP','broj','date','station','variety','status','existing_links','active_sources','otp_kg','source_kg','difference_kg','otp_gajbe','source_gajbe','difference_gajbe','result','reason']
    with (out/'MIG21_PREDMETI.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=cols,delimiter=';');w.writeheader();w.writerows(cases)
    with (out/'MIG21_KOREKCIJE.csv').open('w',encoding='utf-8-sig',newline='') as f:
        cols2=['action','OtpremnicaIzvorID','OtkupID','old_OtpremnicaID','new_OtpremnicaID','old_Otkup_Backlink','number','source_evidence']
        w=csv.DictWriter(f,fieldnames=cols2,delimiter=';');w.writeheader();w.writerows([{k:r.get(k,'') for k in cols2} for r in repair])

def vbahash(path):
    with zipfile.ZipFile(path) as z:
        return hashlib.sha256(z.read('xl/vbaProject.bin')).hexdigest()

def apply(path,out,report,repair,before,ps1):
    if report['blockers']:raise RuntimeError('NO WRITE: contaminated/invalid QA source')
    if not repair:print('NO WRITE: no fully-proven candidates; see MIG21_PREDMETI.csv');return None
    dest=out/'AgriX_MIG21_DOKUMENTNE_VEZE_QA.xlsm'
    if dest.exists():raise FileExistsError(f'Refusing to overwrite output: {dest}')
    prior=sha256(path);vb=vbahash(path);names,stable_hash=all_other_tables_digest(path)
    shutil.copy2(path,dest)
    try:
        ps='powershell.exe'
        result=subprocess.run([ps,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(ps1.resolve()),
                               '-Workbook',str(dest.resolve()),'-Plan',str((out/'MIG21_PLAN.json').resolve())],
                              capture_output=True,text=True,timeout=180)
        if result.returncode:raise RuntimeError('Excel COM failure: '+result.stderr+' '+result.stdout)
        print(result.stdout.strip())
        after=read_xlsm(dest)
        check_readback(before,after,repair)
        if vb!=vbahash(dest):raise AssertionError('VBA binary hash changed; do not use copy')
        names2,stable_hash2=all_other_tables_digest(dest)
        if names!=names2 or stable_hash!=stable_hash2:raise AssertionError('Unrelated Excel ListObject data changed')
        if prior!=sha256(path):raise AssertionError('Original source changed during apply')
        return dest
    except Exception:
        if dest.exists(): dest.rename(dest.with_suffix('.REJECTED.xlsm'))
        raise

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--workbook',default=str(DEFAULT))
    p.add_argument('--stage',default='../MigrationStage_v2')
    p.add_argument('--out-dir',default='../MigrationMIG21_DocumentLinks')
    p.add_argument('--apply',action='store_true',help='Produce NEW XLSM only for exact 1:1 source-grounded repairs.')
    a=p.parse_args(argv)
    inp=Path(a.workbook);out=Path(a.out_dir);stage=Path(a.stage)
    if not inp.is_file():p.error(f'XLSM not found: {inp}. Use --workbook exact path of CLEAN backup.')
    report,cases,repairs,original=build(inp,stage)
    save(out,report,cases,repairs)
    if a.apply and not report['blockers']:
        built=apply(inp,out,report,repairs,original,Path(__file__).with_name('MIG21_APPLY_EXCEL_COM.ps1'))
        if built:report['status']='APPLIED_AND_READBACK_VALIDATED';report['new_workbook']=str(built.resolve());report['new_sha256']=sha256(built)
        else:report['status']='NO_AUTO_ELIGIBLE_CORRECTIONS'
        save(out,report,cases,repairs)
    elif a.apply:report['status']='BLOCKED_DIRTY_SOURCE';save(out,report,cases,repairs)
    print(json.dumps({k:v for k,v in report.items() if k not in ('document_lifecycle','qty_differences','missing_source')},ensure_ascii=False,indent=2))
    print('REPORTS:',out/'MIG21_PREDMETI.csv',out/'MIG21_KOREKCIJE.csv',out/'MIG21_SUMMARY.json',sep='\n  ')
    if report['blockers']:sys.exit(2)

if __name__=='__main__':main()
