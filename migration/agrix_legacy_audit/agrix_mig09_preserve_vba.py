#!/usr/bin/env python3
"""MIG-09: deterministic, OOXML-only migration from private MIG-02 stage.

Writes a NEW .xlsm copy of the manually compiled 66f7897f master; DOES NOT
launch Excel, VBIDE, ImportAllVBA, SelfUpdate, or rewrite vbaProject.bin.
Only 25 canonical goods/master tables receive rows. All source workbooks and
stage remain local/unmodified. Output is NOT approved for operational use.

Prerequisites (same folder): legacy_audit.py, agrix_mig07_goods_copy.py.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET

import agrix_mig07_goods_copy as mig
from legacy_audit import AuditError, parse_xlsm, relationships, parse_range, N, NS_MAIN, NS_REL

VERSION = 'MIG-09.0'
NEW_MASTER_SHA = '0d5a4936df5481289d88ecfed5df22be16597806b5f86d9481d6be2df0676de6'
EXPECTED_TABLES = 47
STAGE_ROWS = 9610
AUDIT_SUFFIX = ['CreatedAt','CreatedBy','ModifiedAt','ModifiedBy']
ALLOWED_SUFFIX = {'tblPrevoznici': AUDIT_SUFFIX, 'tblUtovar': AUDIT_SUFFIX,
                  'tblUtovarStavke': AUDIT_SUFFIX}
XML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XML_SP = 'http://www.w3.org/XML/1998/namespace'
NS = {'x': XML_NS}
RELID = '{'+NS_REL+'}id'
NO_USER_DATA_NAMES = ('agrix_mig09_summary.json','agrix_mig09_summary.txt')


def hash_bytes(b): return hashlib.sha256(b).hexdigest()
def sha(path): return mig.sha256(Path(path))


def ensure_pinned_preflight(preflight):
    p = json.loads(Path(preflight).read_text(encoding='utf-8'))
    if (p.get('format') != 'MIG-06.6' or
        p.get('status') != 'TARGET_PREFLIGHT_PASS_WITH_UNRESOLVED_BUSINESS_GATES' or
        p.get('source_sha256') != mig.SOURCE_SHA or
        p.get('schema_sha256') != mig.SCHEMA_SHA or
        p.get('open_global_gates') != 6 or
        p.get('import_approved') is not False):
        raise AuditError('Historical source/schema preflight identity or six AMB-10 gates changed')
    # Old master hash in this historical report is deliberately obsolete.
    # It must NOT be treated as approval of the new manually compiled master.
    return p


def full_canonical_schema(path):
    if sha(path) != mig.SCHEMA_SHA:
        raise AuditError('Canonical schema changed')
    spec = json.loads(Path(path).read_text(encoding='utf-8'))
    tables = {t['table']:t for t in spec['tables']}
    if len(tables) != EXPECTED_TABLES or len({t['sheet'] for t in spec['tables']}) != EXPECTED_TABLES:
        raise AuditError('Expected 47 tables, each on its own worksheet')
    for name in mig.TABLE_ORDER:
        if name not in tables: raise AuditError('Staged table missing: '+name)
    return tables


def index_parts(z):
    names = set(z.namelist())
    root = ET.fromstring(z.read('xl/workbook.xml'))
    bookrels = relationships(z,'xl/workbook.xml')
    sheet_names = {}
    for sheet in root.findall('x:sheets/x:sheet',N):
        p = bookrels.get(sheet.get(RELID,''))
        if not p: raise AuditError('Unresolved workbook sheet relationship')
        if p in sheet_names: raise AuditError('Duplicate worksheet relationship')
        sheet_names[p] = sheet.get('name')
    by_table = {}
    for part,sheet_name in sheet_names.items():
        rels = relationships(z,part)
        root = ET.fromstring(z.read(part))
        for tp in root.findall('x:tableParts/x:tablePart',N):
            table_path=rels.get(tp.get(RELID,''))
            if not table_path or table_path not in names:
                raise AuditError('Broken worksheet->table relation')
            table_xml=ET.fromstring(z.read(table_path))
            tname=table_xml.get('name') or table_xml.get('displayName')
            if not tname or tname in by_table: raise AuditError('Duplicate/missing table name')
            by_table[tname] = (part,table_path,sheet_name)
    return by_table


def check_full_structure(z, schema):
    by_table=index_parts(z)
    if set(by_table)!=set(schema):
        raise AuditError('Canonical 47-table set mismatch')
    extensions={}
    for name in sorted(schema):
        part,path,sheet=by_table[name]
        expected=schema[name]
        if sheet != expected['sheet']:
            raise AuditError('Worksheet mapping mismatch '+name)
        root=ET.fromstring(z.read(path))
        cols=[c.get('name') for c in root.findall('x:tableColumns/x:tableColumn',NS)]
        canonical=expected['columns']
        if cols[:len(canonical)]!=canonical:
            raise AuditError('Canonical table column prefix/order mismatch '+name)
        tail=cols[len(canonical):]
        if tail and tail!=ALLOWED_SUFFIX.get(name):
            raise AuditError('Unapproved runtime table tail '+name)
        if tail: extensions[name]=tail
        if root.get('headerRowCount','1') != '1' or root.get('totalsRowCount','0') != '0':
            raise AuditError('Unapproved header/total row metadata '+name)
        c1,r1,c2,r2 = parse_range(root.get('ref',''))
        if (c1,r1,c2)!=(1,1,len(cols)) or r2<2:
            raise AuditError('Unexpected table range '+name)
    return by_table,extensions


def assert_blank_stage_tables(template, schema):
    parsed=parse_xlsm(template)
    for name in mig.TABLE_ORDER:
        t=parsed.get(name.lower())
        if t is None or t.columns!=schema[name]['columns']:
            raise AuditError('Blank template staged schema mismatch '+name)
        if t.rows:
            raise AuditError('Master already has business rows in '+name)
    # Ledger/financial & transition tables deliberately NOT imported.
    for name in mig.FORBIDDEN:
        if name in ('tblSEFConfig','tblLocalConfig','tblPoruke','tblKorisnici','tblConfig'):
            continue
        t=parsed.get(name.lower())
        if t is not None and t.rows:
            raise AuditError('Master has pre-existing forbidden business data '+name)


def excel_letters(colnum):
    if not 1 <= colnum <= 16384: raise AuditError('Column number out of Excel limits')
    v=''
    while colnum:
        colnum,r=divmod(colnum-1,26)
        v=chr(65+r)+v
    return v


def legal_xml_text(t, table, col, row):
    # XML 1.0 characters only. Do NOT silently remove or normalize values.
    for ch in t:
        cp=ord(ch)
        if cp in (9,10,13) or 0x20 <= cp <= 0xD7FF or 0xE000 <= cp <= 0xFFFD or 0x10000 <= cp <= 0x10FFFF:
            continue
        raise AuditError('Unsafe XML character in '+table+'/'+col+' row '+str(row))
    return t


def make_rows(rows,cols,kinds,styles,tablename):
    # XML is constructed directly; template VBA binary and package metadata remain untouched.
    out=[]
    for i,row in enumerate(rows,2):
        pieces=[f'<row r="{i}">']
        for j,(value,kind) in enumerate(zip(row,kinds),1):
            if value is None: continue
            col=cols[j-1]
            addr=excel_letters(j)+str(i)
            if kind=='text':
                value=legal_xml_text(str(value),tablename,col,i-1)
                value=escape(value).replace('\r','&#13;')
                pieces.append(f'<c r="{addr}" s="{styles["text"]}" t="inlineStr"><is><t xml:space="preserve">{value}</t></is></c>')
            else:
                import math
                if not isinstance(value,(int,float)) or not math.isfinite(value):
                    raise AuditError('Nonfinite or invalid numeric in '+tablename+'/'+col)
                sty=styles['datetime'] if kind=='date' and (col in mig.DATETIME_FIELDS or col.endswith('Vreme')) else styles['date'] if kind=='date' else None
                attr=f' s="{sty}"' if sty is not None else ''
                pieces.append(f'<c r="{addr}"{attr} t="n"><v>{repr(value)}</v></c>')
        pieces.append('</row>')
        out.append(''.join(pieces))
    return ''.join(out).encode('utf-8')


def inject_styles(stylesxml):
    """Append three NumberFormat XFs without touching any pre-existing styles."""
    root=ET.fromstring(stylesxml)
    xfs=root.find('x:cellXfs',NS)
    if xfs is None: raise AuditError('Master lacks xl/styles.xml cellXfs')
    if root.find('x:cellStyleXfs',NS) is None:
        raise AuditError('Master lacks base cellStyleXfs styles')
    count=len(xfs.findall('x:xf',NS))
    if int(xfs.get('count',str(count)))!=count:
        raise AuditError('Style cellXfs count mismatch')
    existing_fmt=[int(x.get('numFmtId')) for x in root.findall('x:numFmts/x:numFmt',NS)]
    fmt0=max([163]+existing_fmt)+1
    # Preserve original namespace and mc:Ignorable prefix bindings by modifying
    # raw XML bytes, rather than reserializing whole workbook style structures.
    def inject_numfmts(b):
        block=(f'<numFmt numFmtId="{fmt0}" formatCode="dd.mm.yyyy"/>'
               f'<numFmt numFmtId="{fmt0+1}" formatCode="dd.mm.yyyy hh:mm:ss"/>').encode()
        full=re.search(rb'<numFmts\b[^>]*>.*?</numFmts>',b,re.S)
        if full:
            token=full.group(0)
            m=re.search(rb'\bcount="(\d+)"',token)
            if not m: raise AuditError('Missing numFmts count')
            actual=len(root.findall('x:numFmts/x:numFmt',NS))
            if int(m.group(1))!=actual:raise AuditError('numFmts count mismatch')
            updated=re.sub(rb'\bcount="\d+"',f'count="{actual+2}"'.encode(),token,count=1)
            updated=updated.replace(b'</numFmts>',block+b'</numFmts>')
            return b[:full.start()]+updated+b[full.end():]
        selfclosed=re.search(rb'<numFmts\b[^>]*/>',b)
        if selfclosed:
            repl=b'<numFmts count="2">'+block+b'</numFmts>'
            return b[:selfclosed.start()]+repl+b[selfclosed.end():]
        root_open=re.search(rb'<styleSheet\b[^>]*>',b)
        if not root_open: raise AuditError('No styleSheet root XML')
        pos=root_open.end()
        return b[:pos]+b'<numFmts count="2">'+block+b'</numFmts>'+b[pos:]
    # Text (49 builtin), Date & DateTime (custom ids).
    b=inject_numfmts(stylesxml)
    xfs_block=re.search(rb'<cellXfs\b[^>]*>.*?</cellXfs>',b,re.S)
    if not xfs_block:raise AuditError('Cannot locate style cellXfs XML')
    original=xfs_block.group(0)
    xf=(b'<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'+
        f'<xf numFmtId="{fmt0}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.encode()+
        f'<xf numFmtId="{fmt0+1}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'.encode())
    replacement=re.sub(rb'\bcount="\d+"',f'count="{count+3}"'.encode(),original,count=1)
    replacement=replacement.replace(b'</cellXfs>',xf+b'</cellXfs>')
    b=b[:xfs_block.start()]+replacement+b[xfs_block.end():]
    ET.fromstring(b) # fail on broken XML *before* touching output
    return b, {'text':count,'date':count+1,'datetime':count+2}


def new_table_xml(raw,expected_columns,rowcount,tablename):
    t=ET.fromstring(raw)
    if t.get('name')!=tablename:raise AuditError('Wrong OOXML table part '+tablename)
    cols=t.findall('x:tableColumns/x:tableColumn',NS)
    if [c.get('name') for c in cols]!=expected_columns:
        raise AuditError('Table changed from precheck '+tablename)
    old=t.get('ref')
    if not old:raise AuditError('No existing table ref')
    range_ = f'A1:{excel_letters(len(expected_columns))}{rowcount+1}'
    start=re.search(rb'<table\b[^>]*>',raw)
    if not start:raise AuditError('Cannot patch table root')
    original=start.group(0)
    modified,n = re.subn(rb'\bref="[^"]+"',f'ref="{range_}"'.encode(),original,count=1)
    if n!=1:raise AuditError('Cannot patch table ref')
    data=raw[:start.start()]+modified+raw[start.end():]
    # A real ListObject has its own autoFilter; keep its range equal to the table.
    af=re.search(rb'<autoFilter\b[^>]*>',data)
    if af:
        orig=af.group(0)
        upd,n=re.subn(rb'\bref="[^"]+"',f'ref="{range_}"'.encode(),orig,count=1)
        if n!=1:raise AuditError('Cannot patch table autoFilter')
        data=data[:af.start()]+upd+data[af.end():]
    ET.fromstring(data)
    return data


def new_sheet_xml(raw,rows,cols,kinds,styles,tablename):
    tree=ET.fromstring(raw)
    rows_before=tree.findall('x:sheetData/x:row',NS)
    # For this pinned blank master, each data worksheet has only a header and
    # optionally an empty physical row 2. Other layout is NOT blindly cleared.
    if any(int(x.get('r','0'))>2 for x in rows_before):
        raise AuditError('Master has unexpected non-header rows in '+tablename)
    sheetdata=re.search(rb'<sheetData\b[^>]*>.*?</sheetData>',raw,re.S)
    if not sheetdata:raise AuditError('Unsupported worksheet sheetData XML '+tablename)
    sd=sheetdata.group(0)
    opening=re.match(rb'<sheetData\b[^>]*>',sd)
    header=re.search(rb'<row\b[^>]*\br="1"[^>]*>.*?</row>',sd,re.S)
    if header is None:raise AuditError('Worksheet header missing '+tablename)
    # Refuse multirow/hidden content rather than discard customer/supplemental data.
    if len(rows_before) not in (1,2): raise AuditError('Unexpected existing rows '+tablename)
    b=opening.group(0)+header.group(0)+make_rows(rows,cols,kinds,styles,tablename)+b'</sheetData>'
    new=raw[:sheetdata.start()]+b+raw[sheetdata.end():]
    ref=f'A1:{excel_letters(len(cols))}{len(rows)+1}'
    dim=re.search(rb'<dimension\b[^>]*>',new)
    if dim:
        upd,n=re.subn(rb'\bref="[^"]+"',f'ref="{ref}"'.encode(),dim.group(0),count=1)
        if n!=1:raise AuditError('Dimension missing reference '+tablename)
        new=new[:dim.start()]+upd+new[dim.end():]
    ET.fromstring(new)
    return new


def build_ooxml_copy(template,target,schema,stage_matrices,kinds):
    """Never opens Excel; do not touch any original file or ZIP entry except styles,
    25 *data worksheets* and their 25 table definition XML parts."""
    with zipfile.ZipFile(template) as zin:
        parts,suffixes=check_full_structure(zin,schema)
        if 'xl/styles.xml' not in zin.namelist():raise AuditError('Missing styles.xml')
        if any(x.startswith('_xmlsignatures/') for x in zin.namelist()):
            raise AuditError('Digitally signed OOXML package needs special handling')
        newstyle,styles=inject_styles(zin.read('xl/styles.xml'))
        edits={'xl/styles.xml':newstyle}
        for name in mig.TABLE_ORDER:
            sheetpath,tablepath,sheetname=parts[name]
            if sheetpath in edits or tablepath in edits:raise AuditError('Two targets share same OOXML part')
            cols=schema[name]['columns']
            rows=stage_matrices[name]
            edits[sheetpath]=new_sheet_xml(zin.read(sheetpath),rows,cols,kinds[name],styles,name)
            edits[tablepath]=new_table_xml(zin.read(tablepath),cols,len(rows),name)
        if len(edits) != 1+2*len(mig.TABLE_ORDER):
            raise AuditError('Expected 25 sheets + 25 table XML + styles; some parts overlap')
        vba_before=hash_bytes(zin.read('xl/vbaProject.bin'))
        with zipfile.ZipFile(target,'x') as zout:
            for info in zin.infolist():
                # Retain ZIP members names and metadata, byte-identical uncompressed
                # content for all files other than the 51 controlled XML parts.
                zout.writestr(info,edits.get(info.filename,zin.read(info.filename)))
    with zipfile.ZipFile(target) as z:
        if z.testzip() is not None:raise AuditError('New workbook ZIP corruption')
        vba_after=hash_bytes(z.read('xl/vbaProject.bin'))
        if vba_before!=vba_after:
            raise AuditError('VBA binary changed; NEVER approve')
        for name in mig.TABLE_ORDER:
            if name not in parts: raise AuditError('Missing table unexpectedly')
    return {'vba_binary_sha256':vba_before,
            'vba_binary_preserved_exactly':True,
            'vba_automatic_compile_performed':False,
            'runtime_extensions':suffixes,
            'xml_parts_changed':len(edits)}


def cli(argv=None):
    a=argparse.ArgumentParser(description='MIG-09 OOXML-only: protect manually compiled VBA master')
    a.add_argument('--stage',type=Path,required=True)
    a.add_argument('--master',type=Path,required=True)
    a.add_argument('--preflight',type=Path,required=True)
    a.add_argument('--schema',type=Path,required=True)
    a.add_argument('--out',type=Path,required=True)
    mode=a.add_mutually_exclusive_group(required=True)
    mode.add_argument('--precheck-only',action='store_true')
    mode.add_argument('--write-copy',action='store_true')
    args=a.parse_args(argv)
    master=args.master.resolve();stage=args.stage.resolve(); schemafile=args.schema.resolve();out=args.out.resolve()
    preflight=args.preflight/'agrix_workbook_preflight_summary.json' if args.preflight.is_dir() else args.preflight
    preflight=preflight.resolve()
    if not master.is_file() or not stage.is_dir() or not schemafile.is_file() or not preflight.is_file():
        raise AuditError('Missing source input')
    if out.exists(): raise AuditError('Output folder already exists (no overwrite allowed)')
    if any(out==p or out in p.parents or p in out.parents for p in (master,stage,schemafile,preflight)):
        raise AuditError('Unsafe output path')
    if sha(master)!=NEW_MASTER_SHA:
        raise AuditError('New manual master SHA mismatch; MUST be exactly user-confirmed SHA')
    ensure_pinned_preflight(preflight)
    schema=full_canonical_schema(schemafile)
    input_schemas={n:entry['columns'] for n,entry in schema.items()}
    staged=mig.stage_load(stage,input_schemas)
    mats,kinds,counts=mig.prepared(staged,input_schemas)
    if sum(counts.values())!=STAGE_ROWS:raise AuditError('Expected exactly 9610 rows')
    protected={p:sha(p) for p in [master,schemafile,preflight]+[stage/(name+'.jsonl') for name in mig.TABLE_ORDER]}
    with zipfile.ZipFile(master) as z:
        if z.testzip():raise AuditError('Damaged source master OOXML')
        parts,extensions=check_full_structure(z,schema)
        if 'xl/vbaProject.bin' not in z.namelist():raise AuditError('Master has no VBA project')
        if len(z.read('xl/vbaProject.bin'))<1000000:
            raise AuditError('Unexpectedly small VBA project in manual master')
    assert_blank_stage_tables(master,schema)
    print(f'[{VERSION}] PRECHECK PASS 47 canonical tables, runtime={sorted(extensions)}, {sum(counts.values())} private stage rows',flush=True)
    if args.precheck_only:
        print(f'[{VERSION}] No output created; no workbook modified',flush=True)
        return 0
    out.mkdir(parents=True,exist_ok=False)
    tmp=out/'_WRITE_IN_PROGRESS.xlsm'
    dest=out/'AgriX_MIG09_DATA_COPY_NEEDS_MANUAL_VBA_CHECK.xlsm'
    try:
        print(f'[{VERSION}] Direct OOXML-only write starting; Excel/VBIDE will NOT open',flush=True)
        audit=build_ooxml_copy(master,tmp,schema,mats,kinds)
        print(f'[{VERSION}] Saved OOXML; verifying 9610 full data rows',flush=True)
        verified=mig.verify_data_roundtrip(tmp,mats,input_schemas,kinds)
        if sum(verified.values())!=STAGE_ROWS:
            raise AuditError('Full readback count mismatch')
        for p,oldsha in protected.items():
            if sha(p)!=oldsha:raise AuditError('Protected input changed during operation')
        os.replace(tmp,dest)
        summary={'format':VERSION,'status':'DATA_ROUNDTRIP_PASS_VBA_BINARY_PRESERVED_MANUAL_COMPILE_REQUIRED',
                 'source_sha256':mig.SOURCE_SHA,'schema_sha256':mig.SCHEMA_SHA,
                 'manual_master_sha256':NEW_MASTER_SHA,'master_unchanged':sha(master)==NEW_MASTER_SHA,
                 'trial_copy_sha256':sha(dest),'copy_filename':dest.name,
                 'stage_rows_verified':STAGE_ROWS,'tables_verified':len(mig.TABLE_ORDER),
                 'rows_per_table':verified,'vba':audit,
                 'manual_compile_and_reopen_confirmed_after_data_migration':False,
                 'app_business_writers_tested':False,
                 'legacy_storno_replayed':False,'amb10_imported':False,
                 'secondary_financial_postings_imported':False,'six_amb10_business_gates_open':True,
                 'production_approved':False,
                 'no_customer_values_in_report':True}
        (out/NO_USER_DATA_NAMES[0]).write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/NO_USER_DATA_NAMES[1]).write_text(
            f'{VERSION} DATA ROUNDTRIP PASS: {STAGE_ROWS} rows / {len(mig.TABLE_ORDER)} tables\n'
            f'VBA BINARY BYTE-IDENTICAL TO MANUAL MASTER; no Excel COM or VBIDE used\n'
            'REQUIRES MANUAL Excel open / Debug > Compile / close / reopen, without saving original master\n'
            'AMB-10 / financial / business relations NOT MIGRATED OR APPROVED\n',encoding='utf-8')
        print(f'[{VERSION}] DATA ROUNDTRIP PASS, VBA BINARY PRESERVED; MANUAL VBA COMPILE REQUIRED',flush=True)
        print(f'[{VERSION}] trial file: {dest}',flush=True)
        print('Share ONLY agrix_mig09_summary.json and .txt; the .xlsm with private data stays local.',flush=True)
        return 0
    except BaseException:
        for p in [tmp,dest,out/NO_USER_DATA_NAMES[0],out/NO_USER_DATA_NAMES[1]]:
            p.unlink(missing_ok=True)
        try:out.rmdir()
        except OSError:pass
        raise


if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,ValueError,OSError,KeyError,TypeError,zipfile.BadZipFile,ET.ParseError) as exc:
        print(f'{VERSION} STOP: {exc}',file=sys.stderr,flush=True)
        raise SystemExit(2)
