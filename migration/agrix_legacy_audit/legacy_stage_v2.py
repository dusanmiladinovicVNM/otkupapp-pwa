#!/usr/bin/env python3
"""AgriX MIG-02: source-pinned, read-only conversion to LOCAL canonical JSONL staging.

DOES NOT CREATE OR MODIFY XLSM. Staged JSONL contains real confidential values:
NEVER upload the stage directory or commit it to git. The summary contains only counts
and source row ordinals (no real IDs, names, numbers, amounts, dates).
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import os
from pathlib import Path
import sys
import uuid

from legacy_audit import AuditError, Table, parse_xlsm, is_active, sha256_file, trim

SOURCE_SHA = 'a3987ab1f1a1f45bfc508f502a8c80926c6f321387a50a11652cfef422650a6c'
PLAN_VERSION = 'MIG-02.0'
AUDIT_TABLES = ('tblOtkup','tblOtpremnica','tblZbirna','tblPrijemnica','tblAmbalaza',
                'tblMagacin','tblPaleta','tblPaletaStavka','tblPrerada','tblPreradaStavka')
HEADER_LINE_SPEC = {
 'tblOtkup': ('tblOtkupStavke','OtkupID','OtkupStavkaID',
              {'Klasa':'Klasa','Kolicina':'Kolicina','Cena':'Cena',
               'KolAmbalaze':'KolAmbalaze','BrutoKg':'BrutoKg'}),
 'tblOtpremnica': ('tblOtpremnicaStavke','OtpremnicaID','OtpremnicaStavkaID',
                   {'Klasa':'Klasa','Kolicina':'Kolicina','KolAmbalaze':'KolAmbalaze',
                    'BrutoKg':'BrutoKg','Cena':'PredlogCena'}),
 'tblZbirna': ('tblZbirnaStavke','ZbirnaID','ZbirnaStavkaID',
               {'Klasa':'Klasa','UkupnoKolicina':'Kolicina',
                'UkupnoAmbalaze':'KolAmbalaze'}),
}
PASSTHROUGH = ('tblKooperanti','tblStanice','tblKupci','tblVozaci',
               'tblKulture','tblParcele','tblTipAmbalaze','tblTipPalete',
               'tblKutije','tblKese','tblVrstaGotovihProizvoda',
               'tblArtikli','tblMagacin','tblPrijemnica','tblPaleta',
               'tblPaletaStavka','tblPrerada','tblPreradaStavka',
               'tblCenovnik')
SECONDARY = ('tblNovac','tblBankaImport','tblPartnerMap')
QUARANTINE = ('tblAmbalaza','tblStornoVeze')
# Config, users and UI text intentionally come from CURRENT fresh workbook,
# never from the old customer installation.
NOT_IMPORTED = ('tblSEFConfig','tblLocalConfig','tblConfig','tblKorisnici','tblPoruke')


def canon_schema(path: Path) -> dict[str,list[str]]:
    blob=json.loads(path.read_text(encoding='utf-8'))
    schema={v['table']:v['columns'] for v in blob['tables']}
    required=set(HEADER_LINE_SPEC)|{x[0] for x in HEADER_LINE_SPEC.values()}|set(PASSTHROUGH)
    required.add('tblAmbalaza')
    missing=required-set(schema)
    if missing:
        raise AuditError('Current schema missing expected tables: '+', '.join(sorted(missing)))
    return schema


def pk_exists_and_unique(table: Table, name:str) -> None:
    if name not in table.columns:
        raise AuditError(f'{table.name}: missing PK column {name}')
    ids=[trim(r.get(name,'')) for r in table.rows]
    if '' in ids or len(set(ids))!=len(ids):
        raise AuditError(f'{table.name}: empty/duplicated PKs - stage blocked')


def ensure_inputs(tables:dict[str,Table]) -> None:
    for required in AUDIT_TABLES:
        if required.lower() not in tables:
            raise AuditError(f'Source missing table {required}')
    for old, (line_table, pk, line_pk, spec) in HEADER_LINE_SPEC.items():
        t=tables[old.lower()]
        for c in [pk,*spec.keys()]:
            if c not in t.columns:
                raise AuditError(f'{old}: missing required legacy column {c}')
        pk_exists_and_unique(t,pk)
    for table, pk in [('tblPrijemnica','PrijemnicaID'),('tblPaleta','PaletaID'),
                       ('tblPaletaStavka','StavkaID'),('tblPrerada','PreradaID'),
                       ('tblPreradaStavka','StavkaID'),('tblMagacin','MagacinID')]:
        pk_exists_and_unique(tables[table.lower()],pk)
    for i,r in enumerate(tables['tblMagacin'.lower()].rows,1):
        if trim(r.get('ArtikalID',''))!='ART-POC-DUG' or trim(r.get('Tip',''))!='Izlaz':
            raise AuditError(f'tblMagacin row {i}: not a legacy opening debt, do not stage')
    for t in tables.values():
        if t.name in (*HEADER_LINE_SPEC,*PASSTHROUGH) and t.formulas_without_cache:
            raise AuditError(f'{t.name}: {t.formulas_without_cache} un-cached formulas; refuse to migrate')


def canonical_copy(src:dict[str,str], target_cols:list[str], skipped:set[str]|None=None) -> dict[str,str]:
    skipped=skipped or set()
    return {key:src[key] for key in target_cols if key in src and key not in skipped}


def deterministic_id(source_hash:str, table:str, pk:str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL,'agrix-legacy-v2:'+source_hash+':'+table+':'+pk))


def transform_one(table: Table, schema:dict[str,list[str]], source_hash:str) -> tuple[list[dict],list[dict]]:
    line_table,pk,line_pk,mapping=HEADER_LINE_SPEC[table.name]
    header_cols=schema[table.name]
    line_cols=schema[line_table]
    headers,lines=[],[]
    for ordinal,src in enumerate(table.rows,1):
        header=canonical_copy(src,header_cols, set(mapping.keys()))
        # Header must have no legacy cached business quantities, even if schema retains them.
        for k in ('Klasa','Kolicina','Cena','KolAmbalaze','BrutoKg',
                  'UkupnoKolicina','UkupnoAmbalaze','PredlogCena'):
            if k in header and k in mapping:
                del header[k]
        # Old generation metadata belongs to old row-pair hack, never the new header.
        header.pop('GeneracijaID',None)
        # Business-number membership caches are NOT the canonical relation.
        if table.name=='tblOtkup':
            for k in ('OtpremnicaID','BrojOtpremnice','BrojZbirne','ZbirnaRoditeljID'):
                header.pop(k,None)
        elif table.name=='tblOtpremnica':
            for k in ('BrojZbirne','ZbirnaRoditeljID'):
                header.pop(k,None)
        elif table.name=='tblZbirna':
            for k in ('VrstaVoca','SortaVoca','TipAmbalaze'):
                header.pop(k,None)
        headers.append(header)
        line={line_pk:deterministic_id(source_hash,line_table,trim(src[pk])),
              pk:src[pk], 'RedniBroj':'1'}
        for old_key,new_key in mapping.items():
            if new_key in line_cols and trim(src.get(old_key,'')):
                line[new_key]=src[old_key]
        if 'Klasa' not in line:
            raise AuditError(f'{table.name} row {ordinal}: empty class, cannot assume I')
        for k in ('CreatedAt','CreatedBy','ModifiedAt','ModifiedBy'):
            if k in line_cols and trim(src.get(k,'')):
                line[k]=src[k]
        if not trim(line.get('Kolicina','')):
            raise AuditError(f'{table.name} row {ordinal}: empty quantity')
        lines.append(line)
    return headers,lines


def simple_id_ref(tables:dict[str,Table], src_name:str, col:str, target_name:str, target_pk:str):
    src=tables[src_name.lower()]
    tgt=tables[target_name.lower()]
    idx={trim(r.get(target_pk,'')) for r in tgt.rows}
    missing=[];present=0
    for i,r in enumerate(src.rows,1):
        val=trim(r.get(col,''))
        if not val: continue
        if val in idx:present+=1
        else:missing.append(i)
    return {'present':present,'missing_rows':missing,'missing_count':len(missing)}


def make_plan(tables:dict[str,Table],schema:dict[str,list[str]], source_hash:str) -> tuple[dict[str,list[dict]],dict]:
    ensure_inputs(tables)
    out:dict[str,list[dict]]={}
    notes=[]
    for old,(line_table,*_) in HEADER_LINE_SPEC.items():
        h,l=transform_one(tables[old.lower()],schema,source_hash)
        out[old]=h;out[line_table]=l
        # Retain entire legacy source rows to support exact cross-document proof.
        out['SOURCE__'+old]=list(tables[old.lower()].rows)
    for table in PASSTHROUGH:
        src=tables.get(table.lower())
        out[table]=[canonical_copy(r,schema[table]) for r in src.rows] if src else []
    for table in SECONDARY:
        t=tables.get(table.lower())
        if t:
            out['SECONDARY__'+table]=[canonical_copy(r,schema[table]) for r in t.rows] if table in schema else list(t.rows)
    for table in QUARANTINE:
        src=tables.get(table.lower())
        out['UNCONVERTED__'+table]=list(src.rows) if src else []
    for t in ('tblOtkup','tblOtpremnica','tblZbirna'):
        line=HEADER_LINE_SPEC[t][0]
        assert len(out[t])==len(out[line])==len(tables[t.lower()].rows)
    refs={
     'paleta_item_to_receipt':simple_id_ref(tables,'tblPaletaStavka','PrijemnicaID','tblPrijemnica','PrijemnicaID'),
     'paleta_item_to_paleta':simple_id_ref(tables,'tblPaletaStavka','PaletaID','tblPaleta','PaletaID'),
     'prerada_item_to_paleta':simple_id_ref(tables,'tblPreradaStavka','PaletaID','tblPaleta','PaletaID'),
     'prerada_item_to_prerada':simple_id_ref(tables,'tblPreradaStavka','PreradaID','tblPrerada','PreradaID'),
     'old_otkup_to_otp_id':simple_id_ref(tables,'tblOtkup','OtpremnicaID','tblOtpremnica','OtpremnicaID'),
    }
    # Business-number links are not enough: do NOT auto-create Izvori rows.
    # Human/identity verification remains a necessary follow-up.
    numbers={trim(r.get('BrojZbirne','')) for r in tables['tblZbirna'.lower()].rows}
    missing_zbr={}
    for table in ('tblOtkup','tblOtpremnica','tblPrijemnica'):
        missing=[i for i,r in enumerate(tables[table.lower()].rows,1)
                 if trim(r.get('BrojZbirne','')) and trim(r['BrojZbirne']) not in numbers]
        missing_zbr[table]=missing
    unknown_types=collections.Counter(trim(r.get('DokumentTip','')) for r in tables['tblAmbalaza'.lower()].rows)
    unlinked_amb=[i for i,r in enumerate(tables['tblAmbalaza'.lower()].rows,1)
                  if not trim(r.get('DokumentID',''))]
    report={
      'format':PLAN_VERSION,'status':'STAGED_ONLY_NO_XLSM_CREATED','source_sha256':source_hash,
      'source_row_counts':{t.name:len(t.rows) for t in tables.values()},
      'canonical_stage_counts':{key:len(value) for key,value in out.items()},
      'source_to_target_id_policy':'preserve legacy parent PK, deterministic new item UUID',
      'exact_id_links':refs,'missing_zbirna_by_business_number_ordinals':missing_zbr,
      'unresolved_ambalaza':{'rows':len(out['UNCONVERTED__tblAmbalaza']),
                             'document_type_categories':len(unknown_types),
                             'rows_missing_document_id':unlinked_amb},
      'blocked_next_steps':[
       'AMB-10 source direction/ownership/storno conversion for every legacy row',
       'Exact identity ownership validation and reconstruction of OtpremnicaIzvori/ZbirnaIzvori',
       'Resolve all missing Zbirna business-number references without guessing',
       'Confirm old-to-new paleta/prerada state lineage and no dangling ID links',
       'Windows Excel COM target-write and VBA runtime/roundtrip verification'
      ],
      'not_automatically_copied':list(NOT_IMPORTED),
      'notes':['SECONDARY__ data is saved locally, not automatically booked',
               'UNCONVERTED__ tables are retained raw locally; NEVER inject into target canonical table',
               'Canonical stage JSONL contains PRIVATE customer data; do not upload it',
               'NO assumed PWA sync history or synthetic financial transactions']}
    return out,report


def write_private_jsonl(path:Path, records:list[dict]) -> None:
    fd=os.open(str(path),os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
    with os.fdopen(fd,'w',encoding='utf-8',newline='\n') as f:
        for r in records:
            f.write(json.dumps(r,ensure_ascii=False,separators=(',',':'))+'\n')


def cli(argv=None)->int:
    ap=argparse.ArgumentParser(description='MIG-02 convert legacy source into LOCAL canonical staging, NOT a working XLSM')
    ap.add_argument('--source',required=True)
    ap.add_argument('--schema',required=True,help='repo schema/schema.json at pinned target revision')
    ap.add_argument('--out',required=True)
    ap.add_argument('--source-sha256',default=SOURCE_SHA)
    args=ap.parse_args(argv)
    source=Path(args.source).expanduser().resolve()
    schema_path=Path(args.schema).expanduser().resolve()
    dest=Path(args.out).expanduser().resolve()
    if dest==source.parent or source.is_relative_to(dest) or schema_path.is_relative_to(dest):
        raise AuditError('Output directory would overlap with input; refused')
    if dest.exists():raise AuditError('Output exists. For idempotency use a FRESH new directory, no overwrite')
    actual=sha256_file(source)
    if actual.lower()!=args.source_sha256.lower():
        raise AuditError(f'Source SHA mismatch: expected {args.source_sha256}, actual {actual}')
    schema=canon_schema(schema_path)
    raw=parse_xlsm(source)
    converted,report=make_plan(raw,schema,actual)
    report['schema_sha256']=hashlib.sha256(schema_path.read_bytes()).hexdigest()
    dest.mkdir(parents=True,exist_ok=False)
    try:
        for name,rows in converted.items():
            write_private_jsonl(dest/(name+'.jsonl'),rows)
        (dest/'agrix_stage_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        text=['AGRIX MIG-02 -- LOCAL CANONICAL STAGING, NOT TARGET XLSM',
              f'SOURCE SHA256: {actual}', 'NO legacy workbook changed.',
              'STAGED (row counts only):']
        text+=['  %-37s %6d'%(k,len(v)) for k,v in converted.items()]
        text+=['','UNRESOLVED CORE ZBIRNA NUMBER LINKS (row ordinals):']
        text+=['  %s: %s'%(k,v) for k,v in report['missing_zbirna_by_business_number_ordinals'].items()]
        text+=['','EXACT ID LINK CHECKS:']
        text+=['  %s: present=%s missing_count=%s missing_rows=%s'%
              (k,v['present'],v['missing_count'],v['missing_rows'][:35]) for k,v in refs_sorted(report['exact_id_links'])]
        text+=['','AMBALAZA NOT CONVERTED: %d rows; missing_document_id=%s'%
               (report['unresolved_ambalaza']['rows'],report['unresolved_ambalaza']['rows_missing_document_id']),
               'STOP: no output is an executable XLSM. Do not import UNCONVERTED__ files.',
               'PRIVATE JSONL FILES MUST STAY LOCAL. Share only agrix_stage_summary.json.']
        (dest/'agrix_stage_summary.txt').write_text('\n'.join(text)+'\n',encoding='utf-8')
    except BaseException:
        # No partial apparent-success directory; preserve original inputs always.
        for f in dest.glob('*.jsonl'):f.unlink(missing_ok=True)
        for f in dest.glob('agrix_stage_summary.*'):f.unlink(missing_ok=True)
        try:dest.rmdir()
        except OSError:pass
        raise
    print('\n'.join(text))
    print('\nOutput: '+str(dest))
    return 0


def refs_sorted(refs):
    return sorted(refs.items())


if __name__=='__main__':
    try:
        raise SystemExit(cli())
    except (AuditError, OSError, ValueError, json.JSONDecodeError) as e:
        print('MIG-02 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
