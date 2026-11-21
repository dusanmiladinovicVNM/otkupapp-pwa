#!/usr/bin/env python3
"""AgriX MIG-12 legacy packaging archive + MIG-13 read-only financial preflight.

Local-only operation. DOES NOT write .xlsm, source, staging, code or Git.
No private row IDs, prices, amounts or partner names in public summary.
Python 3.10+, stdlib only.
"""
from __future__ import annotations
import argparse
import csv
import hashlib
import json
import os
import shutil
import sys
import tempfile
import zipfile
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path

EXPECTED = {
    'UNCONVERTED__tblAmbalaza.jsonl': ('8a3d1c370b95f08e602216d124af12d67febb7ae3260360789bad7abca138bb5', 7815),
    'SECONDARY__tblNovac.jsonl': ('7731dca7d0167ca9dad69374d4f35ae54e05fb1cacf3800fb67d4191736077b0', 39),
    'UNCONVERTED__tblStornoVeze.jsonl': ('630d0e0e387cd4fce594bd98ff90a86ebbaa22c034fb0a69089cd8d17ac11d86', 2),
}
AMB_REPORTS = [
    'MigrationAmb10DecisionCheck_v2/agrix_amb10_operator_summary.json',
    'MigrationAmb10Decisions_v2/agrix_amb10_operator_summary.json',
    'MigrationAmb10DryRun_v2/agrix_amb10_dryrun_summary.json',
    'MigrationAmb10Forensics_v2/agrix_amb10_forensics_summary.json',
    'MigrationAmb10IdentityBridge_v2/agrix_amb10_identity_bridge_summary.json',
    'MigrationAmb10ImportReadiness_v2/agrix_amb10_import_readiness_summary.json',
    'MigrationAmb10Quarantine_v2/agrix_amb10_quarantine_summary.json',
    'MigrationAmb10Residual_v2/agrix_amb10_residual_summary.json',
    'MigrationAmb10Semantic_v2/agrix_amb10_semantic_preview_summary.json',
    'MigrationAmb10Shadow_v2/agrix_amb10_shadow_summary.json',
]
EXPECTED_AMB_SOURCE = 'a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'
EXPECTED_MIG11_XLSM = '1974d2a8b4c6d6dbd4a8810c2c8d86068f27a92eb4c1cc2829539b51a0f4a5d7'
FIN_CANON = (
    'NovacID','BrojDokumenta','Datum','Partner','PartnerID','EntitetTip','OMID',
    'KooperantID','FakturaID','VrstaVoca','Tip','Uplata','Isplata','Napomena',
    'Stornirano','OtkupID','OsirocenoOd','CreatedAt','CreatedBy','ModifiedAt',
    'ModifiedBy','IspravkaOd','ZamenjenSa','CorrectionID','IzdatoStatus','GeneracijaID'
)
FK = {'OtkupID':('tblOtkup','OtkupID'),
      'KooperantID':('tblKooperanti','KooperantID'),
      'FakturaID':('tblFakture','FakturaID'),
      'OMID':('tblStanice','StanicaID')}

def sha(path:Path)->str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1<<20), b''): h.update(b)
    return h.hexdigest()

def load_jsonl(path:Path):
    rows=[]
    with path.open('r',encoding='utf-8-sig') as f:
        for i,line in enumerate(f,1):
            if not line.strip(): continue
            try: obj=json.loads(line)
            except json.JSONDecodeError as e: raise ValueError(f'Invalid JSONL in {path.name}, line {i}') from e
            if not isinstance(obj,dict): raise ValueError(f'Non-object JSONL in {path.name}, line {i}')
            rows.append(obj)
    return rows

def locate(stage:Path,filename:str,required=True):
    files=[p for p in stage.rglob(filename) if p.is_file() and not any(
        x.casefold() in ('logs','reports','diagnostics','audit') for x in p.relative_to(stage).parts[:-1])]
    if len(files)!=1:
        if not required and not files: return None
        raise ValueError(f'Expected exactly one {filename}, found {len(files)}')
    return files[0]

def norm(x): return '' if x is None else str(x).strip().casefold()

def missing(x): return x is None or (isinstance(x,str) and not x.strip())

def decimal_value(x):
    if x is None or x=='':return Decimal('0')
    if isinstance(x,bool):return None
    raw=str(x).strip().replace(' ','').replace('\u00a0','')
    if ',' in raw and '.' in raw:
        if raw.rfind(',') > raw.rfind('.'):
            raw=raw.replace('.','').replace(',','.')
        else: raw=raw.replace(',','')
    elif ',' in raw: raw=raw.replace(',','.')
    try:
        d=Decimal(raw)
        return d if d.is_finite() else None
    except (InvalidOperation,ValueError): return None

def preflight(finance, stage:Path):
    canonical_set=set(FIN_CANON)
    unknown=Counter(); issue_counts=Counter(); detail=[]; ids=set(); refs={}
    missing_ref_tables=[]
    for field,(tab,idcol) in FK.items():
        p=locate(stage,tab+'.jsonl',required=False)
        if p is None:
            missing_ref_tables.append(tab)
            refs[field]=None
        else:
            refrows=load_jsonl(p)
            known={norm(r.get(idcol)) for r in refrows if not missing(r.get(idcol))}
            refs[field]=known
    existing=locate(stage,'tblNovac.jsonl',required=False)
    existing_ids=set()
    if existing:
        existing_ids={norm(row.get('NovacID')) for row in load_jsonl(existing) if not missing(row.get('NovacID'))}
    columns=Counter(); nonempty=Counter()
    for i,row in enumerate(finance,1):
        errors=[]
        for k,v in row.items():
            columns[k]+=1
            if not missing(v): nonempty[k]+=1
            if k not in canonical_set and not missing(v):
                unknown[k]+=1
                errors.append('UNRECOGNIZED_NONEMPTY_COLUMN')
        rid=norm(row.get('NovacID'))
        if not rid: errors.append('MISSING_NOVAC_ID')
        elif rid in ids or rid in existing_ids: errors.append('DUPLICATE_OR_EXISTING_NOVAC_ID')
        else: ids.add(rid)
        if missing(row.get('Datum')): errors.append('MISSING_DATE')
        if missing(row.get('Tip')): errors.append('MISSING_TYPE')
        if missing(row.get('EntitetTip')): errors.append('MISSING_ENTITY_TYPE')
        if missing(row.get('Partner')) and missing(row.get('PartnerID')):
            errors.append('MISSING_PARTNER')
        pay=decimal_value(row.get('Uplata'))
        out=decimal_value(row.get('Isplata'))
        if pay is None or out is None: errors.append('NON_NUMERIC_AMOUNT')
        elif pay<0 or out<0: errors.append('NEGATIVE_AMOUNT')
        elif (pay>0 and out>0):errors.append('DOUBLE_DIRECTION')
        elif pay==0 and out==0:errors.append('ZERO_AMOUNT')
        for field in FK:
            k=norm(row.get(field))
            if k and refs[field] is None: errors.append('FK_TABLE_NOT_AVAILABLE_'+field.upper())
            elif k and k not in refs[field]: errors.append('ORPHAN_'+field.upper())
        for err in set(errors): issue_counts[err]+=1
        if errors: detail.append({'row_ordinal':i,'issue_codes':';'.join(sorted(set(errors)))})
    if missing_ref_tables: issue_counts['OPTIONAL_REFERENCE_TABLES_NOT_FOUND']=len(missing_ref_tables)
    return {
        'rows':len(finance),
        'columns_present':sorted(columns),
        'nonempty_counts_by_canonical_column':{k:nonempty[k] for k in FIN_CANON if nonempty[k]},
        'unknown_nonempty_column_names':sorted(unknown),
        'existing_canonical_novac_table':bool(existing),
        'missing_optional_reference_tables':sorted(set(missing_ref_tables)),
        'issue_counts':dict(sorted(issue_counts.items())),
        'blocked_rows':len(detail),
        'ready_for_candidate':len(detail)==0 and not unknown,
        'not_approved_to_import':True,
        'note':'Read-only check of 39 secondary ledger rows. Cross-transaction reconciliation and historical payment behavior not yet verified.'
    },detail

def write_json(path,data):
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

def execute(stage:Path,out:Path,root:Path,workbook:Path|None):
    if not stage.is_dir(): raise ValueError('Stage directory does not exist')
    if out==stage or stage in out.parents or out in stage.parents:
        raise ValueError('Output directory may not overlap staging')
    if out.exists(): raise ValueError('Output already exists; nothing overwritten')
    loaded={}; infos={}
    for file_name,(expected_hash,count) in EXPECTED.items():
        p=locate(stage,file_name)
        digest=sha(p)
        if digest!=expected_hash:raise ValueError(file_name+': SHA mismatch, input dataset changed; no output')
        rows=load_jsonl(p)
        if len(rows)!=count:raise ValueError(file_name+': count mismatch; no output')
        loaded[file_name]=(p,rows)
        infos[file_name]={'rows':len(rows),'sha256':digest,'sha_expected_match':True}
    copies=[]; summaries=[]
    for rel in AMB_REPORTS:
        p=root/rel
        if p.is_file():
            v=json.loads(p.read_text(encoding='utf-8-sig'))
            if v.get('source_sha256')!=EXPECTED_AMB_SOURCE:
                raise ValueError(rel+': source fingerprint mismatch')
            copies.append((p,rel))
            summaries.append({'file':rel,'sha256':sha(p),'status':v.get('status'),
                              'import_approved':v.get('import_approved')})
    import_ready=root/'MigrationAmb10ImportReadiness_v2/agrix_amb10_import_readiness_summary.json'
    if not import_ready.is_file():raise ValueError('Required existing AMB-10 import-readiness report is missing')
    readiness=json.loads(import_ready.read_text(encoding='utf-8-sig'))
    if readiness.get('import_approved') is not False or readiness.get('row_approval_count')!=0:
        raise ValueError('AMB readiness is changed: this archive-only tool refuses to infer import permissions')
    workbook_info=None
    if workbook is not None:
        if not workbook.is_file():raise ValueError('MIG-11 workbook missing')
        workbook_info={'exists':True,'sha256':sha(workbook),
                       'matches_last_verified_mig11_sha256':sha(workbook)==EXPECTED_MIG11_XLSM,
                       'edited':False}
    finance,details=preflight(loaded['SECONDARY__tblNovac.jsonl'][1],stage)
    # Source-fingerprinted records archived byte-for-byte; never normalize historical values.
    archive_manifest={'format':'AGRIx-MIG12-LEGACY-ARCHIVE-v1',
                     'source_workbook_original_sha256':EXPECTED_AMB_SOURCE,
                     'source_files':{k:v for k,v in infos.items() if k!='SECONDARY__tblNovac.jsonl'},
                     'embedded_reports':summaries,
                     'no_live_amb10_rows_written':True,
                     'archived_original_content_byte_identical':True,
                     'amb10_import_approved':False,
                     'historical_negative_balances_must_remain_unchanged':True}
    parent=out.parent
    parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='agrix_mig12_finance_',dir=parent) as temp:
        tmp=Path(temp)
        archive=tmp/'MIG12_AMB10_HISTORICAL_ARCHIVE.zip'
        with zipfile.ZipFile(archive,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
            for k in ('UNCONVERTED__tblAmbalaza.jsonl','UNCONVERTED__tblStornoVeze.jsonl'):
                z.write(loaded[k][0],arcname='SOURCE_JSONL/'+k)
            for p,rel in copies:z.write(p,arcname='REPORTS/'+rel)
            z.writestr('README.txt',
                'AGRIX ORIGINAL AMB HISTORY — ARCHIVE ONLY, NOT LIVE AMB-10.\n'
                'Preserve original chronology, ownership uncertainty, cancellation and negative balances.\n'
                'Do not load unconverted records into tblAmbalaza.\n')
            z.writestr('archive_manifest.json',json.dumps(archive_manifest,ensure_ascii=False,indent=2)+'\n')
        # Verify archived bytes, not only ZIP CRC.
        with zipfile.ZipFile(archive) as z:
            for k in ('UNCONVERTED__tblAmbalaza.jsonl','UNCONVERTED__tblStornoVeze.jsonl'):
                digest=hashlib.sha256(z.read('SOURCE_JSONL/'+k)).hexdigest()
                if digest!=infos[k]['sha256']:raise ValueError('ZIP byte preservation failed: '+k)
        fin_raw=tmp/'MIG13_SECONDARY_tblNovac_SOURCE.jsonl'
        shutil.copyfile(loaded['SECONDARY__tblNovac.jsonl'][0],fin_raw)
        if sha(fin_raw)!=infos['SECONDARY__tblNovac.jsonl']['sha256']:
            raise ValueError('Financial source copy hash mismatch')
        with (tmp/'MIG13_finance_issues_PRIVATE_LOCAL.csv').open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=['row_ordinal','issue_codes']);w.writeheader();w.writerows(details)
        if finance['ready_for_candidate']:
            candidate=tmp/'MIG13_tblNovac_CANDIDATE_NOT_IMPORTED.jsonl'
            with candidate.open('w',encoding='utf-8') as f:
                for row in loaded['SECONDARY__tblNovac.jsonl'][1]:
                    f.write(json.dumps(row,ensure_ascii=False,separators=(',',':'))+'\n')
            finance['candidate_created_not_imported']=True
        else:finance['candidate_created_not_imported']=False
        summary={'format':'AGRIx-MIG12-ARCHIVE-MIG13-FINANCE-PREFLIGHT-v1',
                 'archive':{'status':'PASS','legacy_ambalaza_rows':7815,
                            'historical_storno_rows':335,
                            'unconverted_storno_veze_rows':2,
                            'archive_zip_sha256':sha(archive),
                            'archive_verified_byte_identical':True,
                            'historical_reports_preserved':len(copies)},
                 'input_integrity':infos,
                 'mig11_workbook':workbook_info,
                 'finance':finance,
                 'workbook_updated':False,
                 'vba_or_git_modified':False,
                 'next_action':'Finance candidate needs exact mapping/reconciliation before separate OOXML import; AMB archive remains non-live.'}
        write_json(tmp/'MIG12_MIG13_summary.json',summary)
        os.rename(tmp,out)
    return summary

def main():
    p=argparse.ArgumentParser(description='Archive historical ambalaza + audit 39 finance records, read-only')
    p.add_argument('--stage',type=Path,required=True)
    p.add_argument('--out',type=Path,required=True)
    p.add_argument('--project-dir',type=Path,required=True,help='parent AgriX DEV folder with MigrationAmb10* reports')
    p.add_argument('--workbook',type=Path,help='optional: SHA-only check of successful MIG-11 XLSM')
    a=p.parse_args()
    try:
        r=execute(a.stage.resolve(),a.out.resolve(),a.project_dir.resolve(),
                  a.workbook.resolve() if a.workbook else None)
    except (ValueError,OSError,KeyError,json.JSONDecodeError) as e:
        print('BLOCKED:',str(e),file=sys.stderr)
        return 2
    print('MIG-12 ARCHIVE:',r['archive']['status'],'| rows:',r['archive']['legacy_ambalaza_rows'])
    print('MIG-13 FINANCE:',r['finance']['rows'],'rows | blocked:',r['finance']['blocked_rows'])
    print('FINANCE ISSUES:',r['finance']['issue_counts'])
    print('XLSM / VBA / GIT: NOT MODIFIED')
    print('SUMMARY:',a.out/'MIG12_MIG13_summary.json')
    return 0

if __name__=='__main__':sys.exit(main())
