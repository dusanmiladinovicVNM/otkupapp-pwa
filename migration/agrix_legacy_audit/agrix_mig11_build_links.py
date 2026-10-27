#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIx MIG-11 v1.1 — restore historical document membership in canonical stage.

READ ONLY: original stage, legacy workbook, compiled VBA master and repo.
Writes to NEW output: two canonical JSONL link tables, redacted summary and
PRIVATE local issues CSV. Optionally makes a NEW COMPLETE stage snapshot with
those two link tables for existing MIG-09 OOXML importer.
No inference from row order, coincident counts or guessed date/number.
Python 3.10+, stdlib only.
"""
import argparse
import csv
import hashlib
import json
import shutil
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

CANON = ('tblOtkup', 'tblOtpremnica', 'tblZbirna')
LINK = ('tblOtpremnicaIzvori', 'tblZbirnaIzvori')
PRIMARY = {'tblOtkup':'OtkupID', 'tblOtpremnica':'OtpremnicaID', 'tblZbirna':'ZbirnaID'}
OPTIONAL_MATCH = ('StanicaID','BrojDokumenta','BrojOtpremnice','BrojZbirne',
                  'Datum','Stornirano')

def norm(v):
    return '' if v is None else str(v).strip().casefold()

def text(v):
    return '' if v is None else str(v).strip()

def cancelled(v):
    return norm(v) in ('da', 'yes', 'true', '1')

def hashfile(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1<<20), b''):
            h.update(b)
    return h.hexdigest()

def file_named(stage,name):
    matches=[p for p in stage.rglob(name) if p.is_file() and
             not any(x.casefold() in ('reports','audit','diagnostics','logs') for x in p.relative_to(stage).parts[:-1])]
    if len(matches)!=1:
        raise ValueError(f'Ocekivao jedan {name} u stage, pronasao {len(matches)}')
    return matches[0]

def stage_rows(path):
    if path.suffix.lower()!='.jsonl':
        raise ValueError('Samo provereni JSONL stage format je podrzan')
    rows=[]
    with path.open('r',encoding='utf-8-sig') as f:
        for line_no,line in enumerate(f,1):
            if not line.strip(): continue
            try: r=json.loads(line)
            except Exception as e: raise ValueError(f'Pogresan JSON u {path.name}:{line_no}') from e
            if not isinstance(r,dict): raise ValueError(f'Nije dictionary {path.name}:{line_no}')
            rows.append(r)
    return rows

def unique_index(rows,field,label):
    i={}
    for n,r in enumerate(rows,1):
        k=norm(r.get(field))
        if not k: raise ValueError(f'{label}: prazno {field} na redu {n}')
        if k in i: raise ValueError(f'{label}: duplikat {field} (bez izlaganja ID-a)')
        i[k]=r
    return i

def stable_id(prefix,from_id,to_id):
    digest=hashlib.sha256((prefix+'\0'+norm(from_id)+'\0'+norm(to_id)).encode('utf-8')).hexdigest().upper()[:24]
    return f'{prefix}M11-{digest}'

def load(stage):
    data={}; files={}
    for t in CANON:
        f=file_named(stage,t+'.jsonl'); files[t]=f; data[t]=stage_rows(f)
        f=file_named(stage,'SOURCE__'+t+'.jsonl'); files['SOURCE__'+t]=f
        data['SOURCE__'+t]=stage_rows(f)
    for t in LINK:
        ms=[p for p in stage.rglob(t+'.jsonl') if p.is_file()]
        if len(ms)>1: raise ValueError(f'Dva kanonska {t} ulaza')
        files[t]=ms[0] if ms else None
        data[t]=stage_rows(ms[0]) if ms else []
    return data,files

def align_source(data):
    indices={}; sources={}
    for table,key in PRIMARY.items():
        indices[table]=unique_index(data[table],key,table)
        sources[table]=unique_index(data['SOURCE__'+table],key,'SOURCE__'+table)
        if indices[table].keys()!=sources[table].keys():
            raise ValueError(f'{table}: SOURCE ID skup nije jednak kanonskom ID skupu')
        for iid, cr in indices[table].items():
            sr=sources[table][iid]
            for col in OPTIONAL_MATCH:
                if col in cr and col in sr and norm(cr.get(col))!=norm(sr.get(col)):
                    raise ValueError(f'{table}: SOURCE/kanonski konflikt kolone {col}; prekinut import')
    return indices,sources

def lookup_zbr(label,otp,zbrs):
    """Resolve old BrojZbirne by *unique* match; if reused, by uniquely agreeing evidence.
    On contradictory driver/date evidence do not guess.
    """
    label=norm(label)
    if not label: return (None,'NO_BROJ_ZBIRNE')
    candidates=[r for r in zbrs if norm(r.get('BrojZbirne'))==label]
    if not candidates: return (None,'ZBR_NUMBER_NOT_FOUND')
    if len(candidates)==1: return (candidates[0],'UNIQUE_NUMBER')
    driver=norm(otp.get('VozacID'))
    date=norm(otp.get('Datum'))
    by_driver=[r for r in candidates if driver and norm(r.get('VozacID'))==driver]
    by_date=[r for r in candidates if date and norm(r.get('Datum'))==date]
    by_both=[r for r in candidates if driver and date and norm(r.get('VozacID'))==driver and norm(r.get('Datum'))==date]
    if len(by_both)==1: return (by_both[0],'UNIQUE_NUMBER_DRIVER_DATE')
    if len(by_both)>1: return (None,'AMBIGUOUS_NUMBER_DRIVER_DATE')
    if len(by_driver)==1 and len(by_date)==1:
        if norm(by_driver[0].get('ZbirnaID'))!=norm(by_date[0].get('ZbirnaID')):
            return (None,'CONFLICT_DRIVER_VS_DATE')
        return (by_driver[0],'UNIQUE_NUMBER_DRIVER_AND_DATE')
    if len(by_driver)==1 and not by_date:
        return (by_driver[0],'UNIQUE_NUMBER_DRIVER')
    if len(by_date)==1 and not by_driver:
        # A mismatching known driver is contradictory evidence, not a usable date-only match.
        if driver and norm(by_date[0].get('VozacID')) and norm(by_date[0].get('VozacID'))!=driver:
            return (None,'DRIVER_CONTRADICTS_DATE_MATCH')
        return (by_date[0],'UNIQUE_NUMBER_DATE')
    return (None,'AMBIGUOUS_NUMBER')

def build(data):
    canon, src=align_source(data)
    otk=canon['tblOtkup']; otp=canon['tblOtpremnica']; zbr=canon['tblZbirna']
    used_ids=set(); issues=[]; legacy_mismatches=[]; stats=Counter(); risk=Counter(); unlinked=Counter()
    def issue(t,why,a='',b=''):
        issues.append({'table':t,'reason':why,'source_id':a,'target_reference':b})
        stats[t+':'+why]+=1
    def existing_check(t,rows,source_col,target_col,keys):
        pairs={}; by_id={}
        for row in rows:
            a=norm(row.get(source_col)); b=norm(row.get(target_col)); ix=norm(row.get(keys))
            if not a or not b or not ix: raise ValueError(f'{t}: postojeci link ima prazan FK/ID')
            if ix in used_ids: raise ValueError(f'{t}: duplikat link ID')
            used_ids.add(ix)
            pair=(a,b)
            if pair in pairs: raise ValueError(f'{t}: duplirani par clanova')
            pairs[pair]=row
            by_id.setdefault(a,set()).add(b)
        return pairs,by_id
    old_opi, by_otk = existing_check('tblOtpremnicaIzvori',data['tblOtpremnicaIzvori'],
                                    'OtkupID','OtpremnicaID','OtpremnicaIzvorID')
    old_zbi, by_otp = existing_check('tblZbirnaIzvori',data['tblZbirnaIzvori'],
                                    'OtpremnicaID','ZbirnaID','ZbirnaIzvorID')
    if any(a not in otk or any(b not in otp for b in bs) for a,bs in by_otk.items()):
        raise ValueError('Postojeci OPI link ima dangling foreign key')
    if any(a not in otp or any(b not in zbr for b in bs) for a,bs in by_otp.items()):
        raise ValueError('Postojeci ZBI link ima dangling foreign key')
    opi=list(data['tblOtpremnicaIzvori']); zbi=list(data['tblZbirnaIzvori'])
    for k, r in sorted(otk.items()):
        raw=src['tblOtkup'][k]; target=norm(raw.get('OtpremnicaID'))
        if not target:
            unlinked['OPI_NO_HISTORICAL_OTP_REFERENCE']+=1
            continue
        if target not in otp:
            issue('OPI','DANGLING_OTP_REFERENCE',text(r.get('OtkupID')),text(raw.get('OtpremnicaID')))
            continue
        # OtpremnicaID is an EXPLICIT historical reference to an opaque ID.
        # BrojZbirne is a human-facing historical number; a disagreement is
        # kept as an auditable historical anomaly, NOT used to override that ID.
        # A conflicting prior canonical membership remains a hard error.
        if k in by_otk and (by_otk[k]!={target}):
            issue('OPI','CONFLICT_WITH_EXISTING_LINK',text(r.get('OtkupID')),text(raw.get('OtpremnicaID')))
            continue
        old_broj=norm(raw.get('BrojZbirne'))
        otp_broj=norm(src['tblOtpremnica'][target].get('BrojZbirne'))
        if old_broj and otp_broj and old_broj!=otp_broj:
            risk['OPI_LEGACY_BROJ_ZBIRNE_MISMATCH']+=1
            legacy_mismatches.append({
                'OtkupID':text(r['OtkupID']),
                'OtpremnicaID':text(otp[target]['OtpremnicaID']),
                'Otkup_BrojZbirne':text(raw.get('BrojZbirne')),
                'Otpremnica_BrojZbirne':text(src['tblOtpremnica'][target].get('BrojZbirne')),
                'LinkEvidence':'SOURCE__tblOtkup.OtpremnicaID',
                'Disposition':'PRESERVE_EXPLICIT_ID; REVIEW_HISTORICAL_NUMBER'
            })
        if (k,target) in old_opi:
            stats['OPI:EXISTING']+=1
            continue
        # Only historical fact with direct opaque ID, never inferred from similar values.
        cr=otp[target]
        if cancelled(r.get('Stornirano')):
            risk['OPI_CANCELLED_SOURCE']+=1
        if cancelled(cr.get('Stornirano')):
            risk['OPI_CANCELLED_TARGET']+=1
        if cancelled(r.get('Stornirano')) and not cancelled(cr.get('Stornirano')):
            risk['OPI_CANCELLED_OTK_ACTIVE_OTP']+=1
        if norm(r.get('StanicaID')) and norm(cr.get('StanicaID')) and norm(r.get('StanicaID'))!=norm(cr.get('StanicaID')):
            risk['OPI_STANICA_DISAGREES']+=1
        izvor=text(r['OtkupID']); dest=text(cr['OtpremnicaID'])
        rid=stable_id('OPI-',izvor,dest)
        if norm(rid) in used_ids: raise ValueError('Generated OPI ID collision')
        used_ids.add(norm(rid))
        opi.append({'OtpremnicaIzvorID':rid,'OtpremnicaID':dest,'OtkupID':izvor})
        old_opi[k,target]=opi[-1]
        by_otk.setdefault(k,set()).add(target)
        stats['OPI:BUILT_DIRECT_ID']+=1
    for k, r in sorted(otp.items()):
        raw=src['tblOtpremnica'][k]; label=raw.get('BrojZbirne')
        target,how=lookup_zbr(label,raw,data['tblZbirna'])
        if target is None:
            if how=='NO_BROJ_ZBIRNE':
                unlinked['ZBI_NO_HISTORICAL_ZBR_REFERENCE']+=1
            else:
                issue('ZBI',how,text(r.get('OtpremnicaID')),text(label))
            continue
        dest=norm(target['ZbirnaID'])
        if k in by_otp and by_otp[k]!={dest}:
            issue('ZBI','CONFLICT_WITH_EXISTING_LINK',text(r.get('OtpremnicaID')),text(label))
            continue
        if (k,dest) in old_zbi:
            stats['ZBI:EXISTING']+=1
            continue
        if cancelled(r.get('Stornirano')): risk['ZBI_CANCELLED_SOURCE']+=1
        if cancelled(target.get('Stornirano')): risk['ZBI_CANCELLED_TARGET']+=1
        if cancelled(r.get('Stornirano')) and not cancelled(target.get('Stornirano')): risk['ZBI_CANCELLED_OTP_ACTIVE_ZBR']+=1
        if norm(r.get('VozacID')) and norm(target.get('VozacID')) and norm(r.get('VozacID'))!=norm(target.get('VozacID')):
            risk['ZBI_VOZAC_DISAGREES']+=1
        a=text(r['OtpremnicaID']); b=text(target['ZbirnaID']); rid=stable_id('ZBI-',a,b)
        if norm(rid) in used_ids: raise ValueError('Generated ZBI ID collision')
        used_ids.add(norm(rid))
        zbi.append({'ZbirnaIzvorID':rid,'ZbirnaID':b,'OtpremnicaID':a})
        old_zbi[k,dest]=zbi[-1]
        by_otp.setdefault(k,set()).add(dest)
        stats['ZBI:BUILT_'+how]+=1
    # Validate generated FK consistency and active membership (one active parent per source).
    for row in opi:
        if norm(row['OtkupID']) not in otk or norm(row['OtpremnicaID']) not in otp:
            raise ValueError('OPI outbound FK invariant failed')
    for row in zbi:
        if norm(row['OtpremnicaID']) not in otp or norm(row['ZbirnaID']) not in zbr:
            raise ValueError('ZBI outbound FK invariant failed')
    for a, bs in by_otk.items():
        if sum(not cancelled(otp[b].get('Stornirano')) for b in bs)>1:
            raise ValueError('OPI: jedan otkup ima vise aktivnih otpremnica')
    for a, bs in by_otp.items():
        if sum(not cancelled(zbr[b].get('Stornirano')) for b in bs)>1:
            raise ValueError('ZBI: jedna otpremnica ima vise aktivnih zbirnih')
    return opi,zbi,issues,stats,risk,unlinked,legacy_mismatches

def write_jsonl(path,rows):
    with path.open('x',encoding='utf-8',newline='\n') as f:
        for row in rows:
            f.write(json.dumps(row,ensure_ascii=False,separators=(',',':'))+'\n')

def make_snapshot(stage,out,opi,zbi,canonical_rel_dir):
    if out.exists(): raise ValueError('Snapshot stage vec postoji; bez prepisivanja')
    if stage in out.parents or out==stage: raise ValueError('Snapshot ne moze biti unutar originalnog stage-a')
    shutil.copytree(stage,out)
    for t, rows in zip(LINK,(opi,zbi)):
        matches=[p for p in out.rglob(t+'.jsonl') if p.is_file()]
        if len(matches)>1: raise ValueError('Stage snapshot duplira canonical '+t)
        # Place a NEW link alongside the actual canonical tblOtkup.jsonl;
        # original stages may have a nested canonical/ folder, not their root.
        dest=matches[0] if matches else out/canonical_rel_dir/(t+'.jsonl')
        dest.parent.mkdir(parents=True,exist_ok=True)
        if dest.exists(): dest.unlink() # only isolated copy, never source
        write_jsonl(dest,rows)
    return out

def main(argv=None):
    p=argparse.ArgumentParser(description='MIG-11 izvorni istorijski links -> kanonske link tabele')
    p.add_argument('--stage',type=Path,required=True)
    p.add_argument('--out',type=Path,required=True)
    p.add_argument('--stage-copy',action='store_true',help='kreira punu kopiju stage-a + dve kanonske veze')
    a=p.parse_args(argv)
    stage=a.stage.resolve(); out=a.out.resolve()
    if not stage.is_dir(): p.error('Stage ne postoji')
    if out==stage or stage in out.parents: p.error('Izlaz mora biti van originalnog stage-a')
    if out.exists(): p.error('Izlaz vec postoji; izaberi novi direktorijum')
    try:
        data,files=load(stage)
        opi,zbi,issues,stats,risk,unlinked,legacy_mismatches=build(data)
        if len(opi)>len(data['tblOtkup'])+len(data['tblOtpremnicaIzvori']):
            raise ValueError('Broj OPI veza prelazi moguc broj izvora')
        if len(zbi)>len(data['tblOtpremnica'])+len(data['tblZbirnaIzvori']):
            raise ValueError('Broj ZBI veza prelazi moguc broj izvora')
        # Check whole result first; no output if input-level error.
        out.mkdir(parents=True,exist_ok=False)
        overlay=out/'overlay'; overlay.mkdir()
        write_jsonl(overlay/'tblOtpremnicaIzvori.jsonl',opi)
        write_jsonl(overlay/'tblZbirnaIzvori.jsonl',zbi)
        with (out/'mig11_issues_PRIVATE_LOCAL.csv').open('x',encoding='utf-8-sig',newline='') as f:
            writer=csv.DictWriter(f,fieldnames=['table','reason','source_id','target_reference'])
            writer.writeheader(); writer.writerows(issues)
        # PRIVATE: full unaltered evidence for the operator, never upload to chat.
        with (out/'mig11_number_mismatches_PRIVATE_LOCAL.csv').open('x',encoding='utf-8-sig',newline='') as f:
            writer=csv.DictWriter(f,fieldnames=[
                'OtkupID','OtpremnicaID','Otkup_BrojZbirne',
                'Otpremnica_BrojZbirne','LinkEvidence','Disposition'])
            writer.writeheader(); writer.writerows(legacy_mismatches)
        snapshot='NOT_REQUESTED'
        if a.stage_copy and not issues:
            canonical_parent=files['tblOtkup'].parent.relative_to(stage)
            snap=make_snapshot(stage,out/'stage_MIG11_ready',opi,zbi,canonical_parent)
            snapshot=snap.name
        elif a.stage_copy and issues:
            snapshot='BLOCKED_BY_RELATION_ERRORS'
        status=('PARTIAL_LINKS_NEED_REVIEW' if issues else
                'READY_WITH_HISTORICAL_NOTES' if risk or unlinked else 'READY_COMPLETE')
        report={
            'format':'AGRIx-MIG-11-LINKS-BUILD-v1.1',
            'created_utc':datetime.now(timezone.utc).isoformat(timespec='seconds'),
            'status':status,
            'safety':'no source/stage/xlsm/VBA modified; old ids linked only with explicit evidence',
            'source_table_rows':{t:len(data[t]) for t in CANON},
            'links':{'tblOtpremnicaIzvori':len(opi),'tblZbirnaIzvori':len(zbi)},
            'new_link_rows':{'tblOtpremnicaIzvori':len(opi)-len(data['tblOtpremnicaIzvori']),
                             'tblZbirnaIzvori':len(zbi)-len(data['tblZbirnaIzvori'])},
            'outcomes':dict(sorted(stats.items())),
            'unresolved':dict(sorted(Counter(x['table']+':'+x['reason'] for x in issues).items())),
            'historically_unlinked':dict(sorted(unlinked.items())),
            'historical_warnings':dict(sorted(risk.items())),
            'input_sha256':{k:hashfile(v) if v else None for k,v in files.items()},
            'snapshot':snapshot,
            'historical_number_mismatches_preserved_by_explicit_id':len(legacy_mismatches),
            'manual_historical_review_required':bool(legacy_mismatches),
            'note':'Output only; MIG-09 OOXML NOT run. Direct historical OtpremnicaID preserved even when legacy BrojZbirne differs; discrepancy in PRIVATE_LOCAL.csv. Review before production. Storno runtime issue remains separate.',
        }
        (out/'mig11_build_summary.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        print('MIG-11',status)
        print('LINKS:',report['new_link_rows'])
        print('UNRESOLVED:',report['unresolved'])
        print('HISTORICALLY UNLINKED:',report['historically_unlinked'])
        print('WARNINGS:',report['historical_warnings'])
        print('HISTORICAL NUMBER MISMATCHES (direct ID preserved):',len(legacy_mismatches))
        print('STAGE COPY:',snapshot)
        print('No XLSM or VBA modified. Share ONLY mig11_build_summary.json')
        return 0  # Missing original references are reported, not execution failures.
    except (ValueError,OSError) as e:
        print('MIG-11 INPUT ERROR:',str(e),file=sys.stderr)
        return 2

if __name__=='__main__':
    sys.exit(main())
