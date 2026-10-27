#!/usr/bin/env python3
"""MIG-06.6 -- real target workbook (OOXML) and migration-order preflight.

Strictly read-only. Does not copy, save, populate, mutate or authorize any XLSM.
The linked AMB-10 draft remains intentionally non-importable until business policy
and historical exceptions have been resolved. Does not repair negative balances.
"""
from __future__ import annotations
import argparse, hashlib, json, posixpath, re, sys, zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from legacy_audit import AuditError
import legacy_amb10_target_adapter_v2 as adapter

VERSION='MIG-06.6'
SUMMARY='agrix_workbook_preflight_summary.json'
TEXT='agrix_workbook_preflight_summary.txt'
MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PACKAGE_REL='http://schemas.openxmlformats.org/package/2006/relationships'
COLUMNS={'tblAmbalaza':tuple(adapter.linked.CANONICAL_COLUMNS),
         'tblAmbalazaDokument':tuple(sorted(adapter.REQUIRED_HEADERS)),
         'tblOtkup':('OtkupID',),'tblOtpremnica':('OtpremnicaID',),
         'tblPrijemnica':('PrijemnicaID',)}
PARENTS={'Otkup':('tblOtkup','OtkupID'),
         'Otpremnica':('tblOtpremnica','OtpremnicaID'),
         'Prijemnica':('tblPrijemnica','PrijemnicaID')}


def read_xml(z, entry):
    if entry not in z.namelist():raise AuditError('Missing workbook part: '+entry)
    info=z.getinfo(entry)
    if info.file_size>40_000_000:raise AuditError('Oversized XML workbook part: '+entry)
    try:return ET.fromstring(z.read(entry))
    except ET.ParseError as e:raise AuditError('Invalid OOXML XML: '+entry) from e


def rels(z,base,relpath):
    root=read_xml(z,relpath)
    result={}
    for e in root:
        if e.tag.rsplit('}',1)[-1]!='Relationship':continue
        if e.get('TargetMode')=='External':continue
        id=e.get('Id');target=e.get('Target','')
        if not id or not target or id in result:raise AuditError('Invalid OOXML relationship '+relpath)
        if target.startswith('/'):
            full=target.lstrip('/')
        else:
            full=posixpath.normpath(posixpath.join(posixpath.dirname(base),target))
        if full.startswith('../') or full not in z.namelist():
            raise AuditError('Broken/unsafe OOXML internal relationship in '+relpath)
        result[id]=full
    return result


def col_num(name):
    n=0
    for c in name.upper():
        if c not in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':raise AuditError('Bad column coordinate: '+name)
        n=n*26+(ord(c)-64)
    return n


def parse_ref(ref):
    m=re.fullmatch(r'\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?',ref or '')
    if not m:raise AuditError('Invalid table ref: '+str(ref))
    c1,r1=col_num(m[1]),int(m[2]);c2,r2=col_num(m[3] or m[1]),int(m[4] or m[2])
    if not (1<=c1<=c2<=16384 and 1<=r1<=r2<=1048576):raise AuditError('Table ref out of Excel bounds')
    return c1,r1,c2,r2


def cell_data(c):
    # Empty styled cells are acceptable; values, shared-strings, formulae and inline text aren't.
    for el in c:
        kind=el.tag.rsplit('}',1)[-1]
        if kind in ('f','is'):return True
        if kind=='v' and el.text not in ('',None):return True
    return False


def inspect_target(path, extra_blank_tables=(), schema_file=None):
    if path.suffix.lower()!='.xlsm' or not path.is_file():
        raise AuditError('Only an existing, newly generated blank .xlsm target is accepted')
    sha=hashlib.sha256(path.read_bytes()).hexdigest()
    extra_blank_tables=set(extra_blank_tables)
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:raise AuditError('Corrupt target ZIP member')
        if 'xl/vbaProject.bin' not in z.namelist():raise AuditError('XLSM contains no VBA binary')
        for name in ('[Content_Types].xml','xl/workbook.xml','xl/_rels/workbook.xml.rels'):
            read_xml(z,name)
        root=read_xml(z,'xl/workbook.xml')
        wr=rels(z,'xl/workbook.xml','xl/_rels/workbook.xml.rels')
        sheets={}
        for el in root.iter():
            if el.tag.rsplit('}',1)[-1]!='sheet':continue
            name=el.get('name');id=el.get('{'+REL+'}id');target=wr.get(id)
            if not name or not target or target in sheets.values():raise AuditError('Missing/duplicated sheet relationships')
            sheets[name]=target
        if not sheets:raise AuditError('Target contains no worksheets')
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ssroot=read_xml(z,'xl/sharedStrings.xml')
            for si in ssroot:
                if si.tag.rsplit('}',1)[-1]=='si':
                    shared.append(''.join((x.text or '') for x in si.iter() if x.tag.rsplit('}',1)[-1]=='t'))
        tables={}
        for sheet_name, sheet_path in sheets.items():
            sheet_root=read_xml(z,sheet_path)
            sheet_rels_path=posixpath.join(posixpath.dirname(sheet_path),'_rels',posixpath.basename(sheet_path)+'.rels')
            tparts=[e for e in sheet_root.iter() if e.tag.rsplit('}',1)[-1]=='tablePart']
            if not tparts:continue
            sr=rels(z,sheet_path,sheet_rels_path)
            for part in tparts:
                dest=sr.get(part.get('{'+REL+'}id'))
                if not dest:raise AuditError('TablePart relationship unresolved: '+sheet_name)
                table=read_xml(z,dest)
                if table.tag.rsplit('}',1)[-1]!='table':raise AuditError('Table relationship points to non-table')
                name=table.get('name') or table.get('displayName')
                if not name or name in tables:raise AuditError('Missing/duplicated ListObject name')
                start_c,start_r,end_c,end_r=parse_ref(table.get('ref'))
                if name in COLUMNS and (table.get('headerRowCount','1')!='1' or table.get('totalsRowCount','0') not in ('0','')):
                    raise AuditError('Required import table has no standard header or has totals: '+name)
                columns=[el.get('name','') for el in table.iter() if el.tag.rsplit('}',1)[-1]=='tableColumn']
                if len(columns)!=len(set(columns)) or not columns or (end_c-start_c+1)!=len(columns):
                    raise AuditError('Invalid table column count/identity: '+name)
                meaningful=[]
                header={}
                for c in sheet_root.iter():
                    if c.tag.rsplit('}',1)[-1]!='c':continue
                    addr=c.get('r','');match=re.fullmatch('([A-Z]{1,3})([0-9]+)',addr)
                    if not match:continue
                    cn,rn=col_num(match[1]),int(match[2])
                    if start_c<=cn<=end_c and start_r<=rn<=end_r:
                        if rn==start_r:header[cn]=c
                        elif cell_data(c):meaningful.append(addr)
                # Check actual header labels, not merely tableColumn metadata.
                if name in COLUMNS:
                    for ix,expected in enumerate(columns):
                        head=header.get(start_c+ix)
                        if head is None:raise AuditError('Required table missing physical header cell: '+name)
                        mode=head.get('t','')
                        if mode=='s':
                            vs=[e.text for e in head if e.tag.rsplit('}',1)[-1]=='v']
                            try:label=shared[int(vs[0])]
                            except (IndexError,ValueError,TypeError):raise AuditError('Invalid shared header string in '+name)
                        elif mode=='inlineStr':
                            label=''.join((e.text or '') for e in head.iter() if e.tag.rsplit('}',1)[-1]=='t')
                        else:
                            label=''.join((e.text or '') for e in head if e.tag.rsplit('}',1)[-1]=='v')
                        if label!=expected:raise AuditError('Physical Excel header differs from ListObject definition: '+name)
                # For the migration tables, real pre-existing records are unsafe.
                if (name in COLUMNS or name in extra_blank_tables) and meaningful:
                    raise AuditError('Target table NOT BLANK: '+name+' (nonempty cell in '+meaningful[0]+')')
                tables[name]={'sheet':sheet_name,'ref':table.get('ref'),
                              'columns':columns,'existing_meaningful_data_cells':len(meaningful),
                              'table_data_row_slots':max(0,end_r-start_r)}
        for name,required in COLUMNS.items():
            if name not in tables:raise AuditError('Missing target ListObject: '+name)
            actual=tables[name]['columns']
            if name=='tblAmbalaza':
                if actual!=list(required):
                    # Table row data import is by NAME, so don't require order unless order is verified by target rows.
                    if set(actual)!=set(required) or len(actual)!=len(required):
                        raise AuditError('Canonical AMB-10 columns missing/extra: '+name)
            else:
                missing=set(required)-set(actual)
                if missing:raise AuditError('Target ListObject missing '+','.join(sorted(missing))+': '+name)
        missing_staged=extra_blank_tables-set(tables)
        if missing_staged:raise AuditError('Missing destination table(s) required by MIG-02 stage: '+','.join(sorted(missing_staged)))
        schema_check={'status':'NOT_PROVIDED_FULL_SCHEMA_NOT_VERIFIED'}
        if schema_file is not None:
            raw=schema_file.read_bytes()
            schema_sha=hashlib.sha256(raw).hexdigest()
            # MIG-02 SCHEMA_SHA fingerprints the OLD CUSTOMER workbook, not
            # the new repository's schema/schema.json.  Comparing them was
            # incorrect: the migration changes precisely this schema.
            # Check the target against every ordered column in the provided
            # canonical target schema instead, including the AMB-10 columns.
            spec=json.loads(raw)
            source_tables={x['table']:x['columns'] for x in spec['tables']}
            mismatches=[]
            for t,fields in source_tables.items():
                if t not in tables or tables[t]['columns']!=list(fields):mismatches.append(t)
            if mismatches:
                raise AuditError('Template does not match pinned FULL target schema: '+','.join(sorted(mismatches)))
            schema_check={'status':'FULL_TARGET_SCHEMA_MATCH','target_schema_sha256':schema_sha,
                          'source_workbook_schema_sha256':adapter.ready.SCHEMA_SHA,
                          'schema_tables_verified':len(source_tables)}
        return {'status':'SCHEMA_AND_BLANKNESS_PASS_NOT_IMPORT_APPROVAL',
                'target_file_sha256':sha,'vba_project_sha256':hashlib.sha256(z.read('xl/vbaProject.bin')).hexdigest(),
                'target_workbook_name_suppressed':True,'worksheet_count':len(sheets),
                'listobject_count':len(tables),
                'required_tables':{name:tables[name] for name in COLUMNS},
                'all_required_tables_are_empty':True,
                'all_staged_target_tables_empty':True,
                'target_stage_table_count':len(extra_blank_tables),
                'full_schema_check':schema_check}


def verify_adapter(folder):
    summaryfile=folder/adapter.SUMMARY
    if not summaryfile.is_file():raise AuditError('Missing MIG-06.5 summary')
    s=json.loads(summaryfile.read_text(encoding='utf-8'))
    fixed={'format':'MIG-06.5','status':'TARGET_ADAPTER_PAYLOAD_NOT_APPROVED_NO_XLSM_WRITE',
          'source_sha256':adapter.ready.SOURCE_SHA,'schema_sha256':adapter.ready.SCHEMA_SHA,
          'revers_header_payload_rows':48,'ledger_payload_rows_not_importable':4669,
          'parent_goods_document_ledger_rows':4621,'new_revers_document_ledger_rows':48,
          'unknown_document_holds':6,'q00700_held':True,
          'historical_cancelled_legs_archived_no_replay':335,
          'source_negative_balances_accepted_without_modification':35,
          'unexplained_balance_residuals':0,'global_release_gates_open':6,
          'candidate_ambid_collisions':0,'referential_errors':0,
          'canonical_import_approved':False,'xlsx_or_xlsm_modified':False,
          'audit_fields_or_canonical_storno_populated':False}
    for k,v in fixed.items():
        if s.get(k)!=v:raise AuditError('Changed MIG-06.5 invariant: '+k)
    res={}
    for fn,key in ((adapter.PRIVATE_ROWS,'private_ledger_payload_sha256'),
                   (adapter.PRIVATE_HEADERS,'private_header_payload_sha256'),
                   (adapter.PRIVATE_EXCEPTIONS,'private_hold_manifest_sha256')):
        f=folder/fn
        if not f.is_file() or hashlib.sha256(f.read_bytes()).hexdigest()!=s.get(key):
            raise AuditError('MIG-06.5 private byte checksum mismatch: '+fn)
        objs=[]
        with f.open('r',encoding='utf-8') as stream:
            for line in stream:
                x=json.loads(line)
                if line!=adapter.canon(x):raise AuditError('Noncanonical MIG-06.5 private JSONL')
                objs.append(x)
        res[fn]=objs
    if (len(res[adapter.PRIVATE_ROWS]),len(res[adapter.PRIVATE_HEADERS]),len(res[adapter.PRIVATE_EXCEPTIONS]))!=(4669,48,8):
        raise AuditError('MIG-06.5 private cardinalities changed')
    return s,res[adapter.PRIVATE_ROWS],res[adapter.PRIVATE_HEADERS],res[adapter.PRIVATE_EXCEPTIONS]


def verify_parent_stage(stage,ledger):
    sp=stage/'agrix_stage_summary.json'
    if not sp.is_file():raise AuditError('Required MIG-02 staging summary absent')
    meta=json.loads(sp.read_text(encoding='utf-8'))
    if meta.get('format')!='MIG-02.0' or meta.get('source_sha256')!=adapter.ready.SOURCE_SHA:
        raise AuditError('MIG-02 source pin mismatch')
    registry={}
    for table,pk in PARENTS.values():
        file=stage/(table+'.jsonl')
        if not file.is_file():raise AuditError('Missing parent stage: '+table)
        ids=set();count=0
        with file.open('r',encoding='utf-8') as f:
            for line in f:
                row=json.loads(line);count+=1;v=str(row.get(pk,'')).strip()
                if not v or v in ids:raise AuditError('Stage parent empty/duplicate ID '+table)
                ids.add(v)
        if meta.get('canonical_stage_counts',{}).get(table)!=count:
            raise AuditError('Stage header count mismatch '+table)
        registry[table]=ids
    matched=Counter()
    for r in ledger:
        doc=r['target_column_values_WITHOUT_AMBID']
        dtype=doc['DokumentTIP']
        if dtype in PARENTS:
            table,pk=PARENTS[dtype]
            if doc['DokumentID'] not in registry[table]:
                raise AuditError('Ambalaža parent absent from staged '+table)
            matched[table]+=1
        elif dtype!='AmbalazaDokument':
            raise AuditError('Unrecognized target parent type in candidate payload')
    if sum(matched.values())!=4621:raise AuditError('Unmatched staged goods-parent candidate total')
    return {'stage_source_sha256_verified':True,'staged_parent_references_verified':dict(sorted(matched.items())),
            'all_goods_parents_must_be_imported_before_ambalaza':True}


def run(adapterdir,stage,template,schema_file=None):
    asm,ledger,headers,excepts=verify_adapter(adapterdir)
    stage_info=json.loads((stage/'agrix_stage_summary.json').read_text(encoding='utf-8'))
    required_stage_tables={k for k in stage_info.get('canonical_stage_counts',{}) if k.startswith('tbl')}
    meta=inspect_target(template,required_stage_tables,schema_file)
    stage_meta=verify_parent_stage(stage,ledger)
    out={'format':VERSION,'status':'TARGET_PREFLIGHT_PASS_WITH_UNRESOLVED_BUSINESS_GATES',
         'source_sha256':asm['source_sha256'],'schema_sha256':asm['schema_sha256'],
         'adapter_ledger_sha256':asm['private_ledger_payload_sha256'],
         'adapter_header_sha256':asm['private_header_payload_sha256'],
         'candidate_ledger_rows':len(ledger),'draft_revers_headers':len(headers),
         'unknown_doc_holds':6,'amb00700_held':True,'archived_old_storno_legs':335,
         'negative_source_balances_accepted':35,'open_global_gates':6,
         'import_approved':False,'target_workbook_written':False,
         'target_preflight':meta,'staged_parent_preflight':stage_meta,
         'next_action':'Import verified masters + goods documents into a COPY first; no AMB10 write until global gates resolved',
         'warning':'This preflight does NOT grant canonical AMB-10 import approval.'}
    return out


def text_summary(s):
    m=s['target_preflight'];p=s['staged_parent_preflight']
    return ('AGRIX MIG-06.6 / TARGET WORKBOOK PREFLIGHT -- NO WRITE\n'
      'Target schema and blank tables: '+m['status']+'\n'
      'Required AMB/parent tables: '+', '.join(m['required_tables'])+'\n'
      'All staged target tables are empty: '+str(m['all_staged_target_tables_empty'])+' ('+str(m['target_stage_table_count'])+' tables)\n'
      'Full pinned schema validation: '+m['full_schema_check']['status']+'\n'
      'Target workbook SHA256: '+m['target_file_sha256']+'\n'
      'Source stage goods document references verified: '+str(sum(p['staged_parent_references_verified'].values()))+'\n'
      'AMB-10 candidate rows (NOT APPROVED): '+str(s['candidate_ledger_rows'])+'\n'
      'Open global business gates: '+str(s['open_global_gates'])+'\n'
      'Import approved: false. Workbook modified: false.\n')


def cli(argv=None):
    ap=argparse.ArgumentParser(description='Inspect blank macro-enabled new AgriX template and parent staging; no write')
    ap.add_argument('--adapter',type=Path,required=True)
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--template',type=Path,required=True)
    ap.add_argument('--out',type=Path,required=True)
    ap.add_argument('--schema',type=Path,help='Optional pinned schema/schema.json for full table/column verification; strongly recommended')
    a=ap.parse_args(argv)
    inp=a.adapter.resolve();stage=a.stage.resolve();template=a.template.resolve();out=a.out.resolve()
    if not inp.is_dir() or not stage.is_dir() or out.exists():raise AuditError('Input missing or output already exists')
    if not template.is_file() or template.suffix.lower()!='.xlsm':raise AuditError('Provide the fresh blank .xlsm template')
    for x in (inp,stage):
        if x==out or x in out.parents or out in x.parents:raise AuditError('Input/output overlap')
    if out==template or out in template.parents:raise AuditError('Output overlaps target workbook location')
    before=hashlib.sha256(template.read_bytes()).hexdigest()
    s=run(inp,stage,template,a.schema.resolve() if a.schema else None)
    if hashlib.sha256(template.read_bytes()).hexdigest()!=before:raise AuditError('Template changed during audit')
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/SUMMARY).write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (out/TEXT).write_text(text_summary(s),encoding='utf-8')
    except BaseException:
        for child in out.iterdir():child.unlink()
        out.rmdir();raise
    print(text_summary(s))
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,TypeError,zipfile.BadZipFile,ET.ParseError) as e:
        print('MIG-06.6 STOP:',e,file=sys.stderr);raise SystemExit(2)
