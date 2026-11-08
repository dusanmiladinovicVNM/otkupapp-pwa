#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-11: append validated canonical membership records into two EMPTY Excel tables.

Works on a NEW XLSM copy; never alters VBA, original master, original source,
legacy stage, or input MIG-09/11 workbook. Python 3.10+, standard library only.
Refuses to overwrite existing cells, tables, or files.
"""
import argparse
from collections import Counter
from hashlib import sha256
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import json, os, re, sys, tempfile, zipfile, copy

from agrix_mig11_verify_xlsm import (REQUIRED, SS, q, relationships, resolve_path,
                                    RANGE, index_column, check_workbook)

TABLE_SET=set(REQUIRED)
ROW_RE=re.compile(rb'<row\b[^>]*(?:/>|>.*?</row>)', re.DOTALL)
SD_RE=re.compile(rb'(<sheetData\b[^>]*>)(.*?)(</sheetData>)', re.DOTALL)
CELLREF_RE=re.compile(r'^([A-Z]+)([0-9]+)$')
SHEET_DIM_RE=re.compile(rb'(<dimension\b[^>]*\bref\s*=\s*)(["\'])([^"\']+)\2')
ATTR_REF_RE=re.compile(rb'(\bref\s*=\s*)(["\'])([^"\']+)\2')


def col_name(n):
    out=''
    while n:
        n,rem=divmod(n-1,26)
        out=chr(65+rem)+out
    return out


def pos(coord):
    m=CELLREF_RE.fullmatch(coord.replace('$','').upper())
    if not m:raise ValueError('Unsupported A1 cell address')
    return index_column(m.group(1)),int(m.group(2))


def parse_ref(ref):
    m=RANGE.fullmatch(ref)
    if not m:raise ValueError('Expected rectangular table range, not '+str(ref))
    return (index_column(m.group(1)),int(m.group(2)),
            index_column(m.group(3)),int(m.group(4)))


def replace_attr_ref(tag,expected,new_value,reason):
    m=ATTR_REF_RE.search(tag)
    if not m or m.group(3).decode('utf-8')!=expected:
        raise ValueError(reason+': reference does not match table')
    return tag[:m.start(3)]+new_value.encode('ascii')+tag[m.end(3):]


def patch_table_xml(payload,old_ref,new_ref):
    root=ET.fromstring(payload)
    if root.attrib.get('ref')!=old_ref:
        raise ValueError('Table ref changed between parse and patch')
    start=re.search(rb'<table\b[^>]*>',payload)
    if start is None:raise ValueError('Unknown table XML format')
    new_tag=replace_attr_ref(start.group(0),old_ref,new_ref,'Table')
    modified=payload[:start.start()]+new_tag+payload[start.end():]
    filter_el=root.find(q('autoFilter'))
    if filter_el is not None:
        match=re.search(rb'<autoFilter\b[^>]*>',modified)
        if match is None:
            # Optional self-closing filter tag has closing >, so normally matched.
            raise ValueError('Missing table autoFilter XML')
        filt_ref=filter_el.attrib.get('ref')
        if filt_ref not in (old_ref,None):
            raise ValueError('Unexpected autoFilter ref; manual review required')
        if filt_ref is not None:
            filt_tag=replace_attr_ref(match.group(0),old_ref,new_ref,'AutoFilter')
            modified=modified[:match.start()]+filt_tag+modified[match.end():]
    if root.find(q('sortState')) is not None:
        raise ValueError('Nonempty or sorted table template: refusing edit')
    ET.fromstring(modified)
    return modified


def is_cell_blank(cell):
    if cell.find(q('f')) is not None:return False
    value=cell.find(q('v'))
    if value is not None and (value.text or '').strip():return False
    for item in cell.iter(q('t')):
        if (item.text or '').strip():return False
    return True


def make_rows(stage_rows,fields,cols,start_col,start_row,old_cells):
    byfield={name:start_col+i for i,name in enumerate(cols)}
    styles={}
    for cell in old_cells:
        p=cell.attrib.get('r','')
        if p:
            col,_=pos(p)
            s=cell.attrib.get('s')
            if s is not None:
                if not s.isdigit():raise ValueError('Unexpected cell style ID')
                styles[col]=s
    segments=[]
    for seq, obj in enumerate(stage_rows):
        rowno=start_row+seq
        cells=[]
        for field in cols:
            value=obj.get(field,'')
            if value is None or value=='':continue
            if not isinstance(value,str):raise ValueError('Link values must be strings')
            cnum=byfield[field]
            addr=f'{col_name(cnum)}{rowno}'
            style=f' s="{styles[cnum]}"' if cnum in styles else ''
            # Always inlineStr: no changes to sharedStrings.xml, no formula injection.
            xval=escape(value,{'"':'&quot;',"'":'&apos;'})
            cells.append(f'<c r="{addr}"{style} t="inlineStr"><is><t>{xval}</t></is></c>')
        segments.append(f'<row r="{rowno}" spans="{start_col}:{start_col+len(cols)-1}">{"".join(cells)}</row>')
    return ''.join(segments).encode('utf-8')


def dimension_patch(payload,first_col,last_col,first_row,last_row):
    found=SHEET_DIM_RE.search(payload)
    if found is None:return payload  # dimension is optional
    old=found.group(3).decode('utf-8')
    if ':' in old:
        a,b=old.split(':',1)
    else:
        a=b=old
    try:
        c1,r1=pos(a)
        c2,r2=pos(b)
    except ValueError as err:
        raise ValueError('Unsupported worksheet dimension') from err
    new=f'{col_name(min(c1,first_col))}{min(r1,first_row)}:{col_name(max(c2,last_col))}{max(r2,last_row)}'
    return payload[:found.start(3)]+new.encode('ascii')+payload[found.end(3):]


def patch_sheet_xml(payload,table,cols,records):
    root=ET.fromstring(payload)
    old_ref=table.attrib['ref']
    c1,r1,c2,r2=parse_ref(old_ref)
    if int(table.attrib.get('headerRowCount','1'))!=1 or int(table.attrib.get('totalsRowCount','0'))!=0:
        raise ValueError('Table needs a one-row header and no totals row')
    if not records:raise ValueError('Cannot replace table with empty stage')
    if c2-c1+1 != len(cols):raise ValueError('Table XML columns disagree with ref')
    first_data=r1+1
    last_data=first_data+len(records)-1
    sheetdata=root.find(q('sheetData'))
    if sheetdata is None:raise ValueError('Worksheet sheetData missing')
    stub_cells=[]
    for row in sheetdata.findall(q('row')):
        ri=int(row.attrib['r'])
        # Tables containing existing actual user data must never be modified.
        if first_data <= ri <= max(last_data,r2):
            for c in row.findall(q('c')):
                if not is_cell_blank(c):
                    raise ValueError('Occupied data cell in expansion range: refuse overwrite')
                if ri==first_data:
                    ci,_=pos(c.attrib.get('r',''))
                    if c1<=ci<=c2:stub_cells.append(c)
    new_rows=make_rows(records,REQUIRED[table.attrib['name']],cols,c1,first_data,stub_cells)
    match=SD_RE.search(payload)
    if not match:raise ValueError('Unsupported worksheet sheetData serialization')
    body=match.group(2)
    row_matches=list(ROW_RE.finditer(body))
    if len(row_matches)!=len(sheetdata.findall(q('row'))):
        raise ValueError('Unsupported row serialization; no edits performed')
    # Never change worksheet rows outside the newly filled table range.
    data_fragments=[]
    inserted=False
    prior=-1
    for found in row_matches:
        raw=found.group(0)
        rm=re.search(rb'<row\b[^>]*\br\s*=\s*["\']([0-9]+)["\']',raw)
        if not rm:raise ValueError('Row lacking r number')
        ri=int(rm.group(1))
        if ri<=prior:raise ValueError('Worksheet rows not strictly ascending')
        prior=ri
        if ri>=first_data and not inserted:
            data_fragments.append(new_rows)
            inserted=True
        elif ri>last_data and not inserted:
            data_fragments.append(new_rows)
            inserted=True
        if ri < first_data or ri > max(last_data,r2):
            data_fragments.append(raw)
    if not inserted:data_fragments.append(new_rows)
    # Header and unrelated rows remain bytewise identical.
    modified=payload[:match.start(2)]+b''.join(data_fragments)+payload[match.end(2):]
    modified=dimension_patch(modified,c1,c2,r1,last_data)
    parsed=ET.fromstring(modified)
    sd=parsed.find(q('sheetData'))
    allrows=[int(x.attrib['r']) for x in sd.findall(q('row'))]
    if len(allrows)!=len(set(allrows)) or allrows!=sorted(allrows):
        raise ValueError('Duplicate/nonmonotonic worksheet rows after patch')
    if len([n for n in allrows if first_data<=n<=last_data])!=len(records):
        raise ValueError('Row count mismatch in patched sheet')
    return modified


def stage_table(stage,name):
    files=list(stage.rglob(name+'.jsonl'))
    if len(files)!=1:raise ValueError(name+': expected exactly one canonical stage file')
    rows=[]
    for i,line in enumerate(files[0].read_text(encoding='utf-8-sig').splitlines(),1):
        if not line.strip():continue
        obj=json.loads(line)
        if not isinstance(obj,dict):raise ValueError(name+': JSON object expected')
        for k in REQUIRED[name]:
            if not isinstance(obj.get(k),str) or not obj[k].strip():
                raise ValueError(name+f': empty/non-string {k} on row {i}')
            if obj[k]!=obj[k].strip():raise ValueError(name+': untrimmed ID')
        rows.append(obj)
    ids=[r[REQUIRED[name][0]].casefold() for r in rows]
    if len(ids)!=len(set(ids)):raise ValueError(name+': duplicate link IDs')
    pairs=[tuple(r[x].casefold() for x in REQUIRED[name][1:]) for r in rows]
    if len(pairs)!=len(set(pairs)):raise ValueError(name+': duplicate relationship')
    return rows


def index_ids(stage,table,key):
    files=list(stage.rglob(table+'.jsonl'))
    if len(files)!=1:raise ValueError('Missing canonical '+table)
    out=set()
    for line in files[0].read_text(encoding='utf-8-sig').splitlines():
        if not line.strip():continue
        obj=json.loads(line)
        val=obj.get(key)
        if not isinstance(val,str) or not val.strip():raise ValueError(table+': bad ID')
        v=val.casefold()
        if v in out:raise ValueError(table+': duplicate ID')
        out.add(v)
    return out


def check_links(stage,records):
    keys={
       'tblOtkup':index_ids(stage,'tblOtkup','OtkupID'),
       'tblOtpremnica':index_ids(stage,'tblOtpremnica','OtpremnicaID'),
       'tblZbirna':index_ids(stage,'tblZbirna','ZbirnaID'),
    }
    seen_opi=set(); seen_zbi=set()
    for r in records['tblOtpremnicaIzvori']:
        if r['OtkupID'].casefold() not in keys['tblOtkup'] or r['OtpremnicaID'].casefold() not in keys['tblOtpremnica']:
            raise ValueError('OPI: orphan link')
        if r['OtkupID'].casefold() in seen_opi:raise ValueError('OPI: same otkup linked more than once')
        seen_opi.add(r['OtkupID'].casefold())
    for r in records['tblZbirnaIzvori']:
        if r['OtpremnicaID'].casefold() not in keys['tblOtpremnica'] or r['ZbirnaID'].casefold() not in keys['tblZbirna']:
            raise ValueError('ZBI: orphan link')
        if r['OtpremnicaID'].casefold() in seen_zbi:raise ValueError('ZBI: same otpremnica linked more than once')
        seen_zbi.add(r['OtpremnicaID'].casefold())


def table_locations(z):
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    wb_rel=relationships(z,'xl/workbook.xml')
    result={}
    sheets={}
    ridns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for sh in wb.findall('.//'+q('sheets')+'/'+q('sheet')):
        rid=sh.attrib.get(ridns)
        if rid not in wb_rel:raise ValueError('Broken workbook sheet relationship')
        sheetpath=wb_rel[rid]
        sheets[sheetpath]=sh.attrib.get('name','?')
        for target in relationships(z,sheetpath).values():
            if not target.startswith('xl/tables/') or not target.endswith('.xml'):continue
            tr=ET.fromstring(z.read(target))
            table=tr.attrib.get('name') or tr.attrib.get('displayName')
            if table in result:raise ValueError('Duplicate table '+table)
            result[table]=(target,sheetpath,tr)
    return result,sheets


def prepare_patches(z,records):
    locate,sheets=table_locations(z)
    patches={}
    for name in REQUIRED:
        if name not in locate:raise ValueError('Missing canonical XLSM table '+name)
        tablepath,sheetpath,tab=locate[name]
        if sum(s==sheetpath for _,s,_ in locate.values())>1:
            raise ValueError('Two tables share link-table sheet; dedicated sheet required')
        cols=tab.find(q('tableColumns'))
        if cols is None:raise ValueError('Table columns missing')
        columns=[c.attrib['name'] for c in cols.findall(q('tableColumn'))]
        if len(columns)!=len(set(columns)) or any(key not in columns for key in REQUIRED[name]):
            raise ValueError(name+': unexpected table column definitions')
        for p in (sheetpath,tablepath):
            if p in patches:raise ValueError('Two link tables share an OOXML part')
        c1,r1,c2,r2=parse_ref(tab.attrib['ref'])
        if r2-r1!=1:
            raise ValueError(name+': expected exactly one empty template data row, found different table size')
        new_ref=f'{col_name(c1)}{r1}:{col_name(c2)}{r1+len(records[name])}'
        patches[tablepath]=patch_table_xml(z.read(tablepath),tab.attrib['ref'],new_ref)
        patches[sheetpath]=patch_sheet_xml(z.read(sheetpath),tab,columns,records[name])
    return patches


def main(argv=None):
    p=argparse.ArgumentParser(description='Write only MIG-11 relation tables to a new isolated XLSM copy')
    p.add_argument('--input-dir',type=Path,required=True,help='Directory with exactly one previously imported XLSM (MIG-11 failed verification)')
    p.add_argument('--master',type=Path,required=True)
    p.add_argument('--stage',type=Path,required=True)
    p.add_argument('--out-dir',type=Path,required=True)
    p.add_argument('--expected-opi',type=int,default=1692)
    p.add_argument('--expected-zbi',type=int,default=907)
    a=p.parse_args(argv)
    try:
        inputs=list(a.input_dir.glob('*.xlsm'))
        if len(inputs)!=1:raise ValueError('Input directory must contain exactly one XLSM; found '+str(len(inputs)))
        source=inputs[0].resolve()
        master=a.master.resolve();stage=a.stage.resolve();outdir=a.out_dir.resolve()
        if not master.is_file() or not stage.is_dir():raise ValueError('Missing master or stage')
        if outdir.exists():raise ValueError('Output directory exists; refusing overwrite')
        if source==master:raise ValueError('Refusing to use master as input XLSM')
        if outdir==a.input_dir.resolve() or outdir==stage or stage in outdir.parents:
            raise ValueError('Output directory overlaps input or staging')
        records={name:stage_table(stage,name) for name in REQUIRED}
        if len(records['tblOtpremnicaIzvori'])!=a.expected_opi or len(records['tblZbirnaIzvori'])!=a.expected_zbi:
            raise ValueError('Source stage counts changed; expected %d + %d'%(a.expected_opi,a.expected_zbi))
        check_links(stage,records)
        with zipfile.ZipFile(source) as zin,zipfile.ZipFile(master) as m:
            if zin.testzip() or m.testzip():raise ValueError('ZIP CRC check failed')
            if len(zin.namelist())!=len(set(zin.namelist())):raise ValueError('Duplicate XLSM ZIP part')
            vbafile='xl/vbaProject.bin'
            if vbafile not in zin.namelist() or vbafile not in m.namelist():raise ValueError('Missing VBA project')
            vhash=sha256(zin.read(vbafile)).hexdigest()
            if vhash!=sha256(m.read(vbafile)).hexdigest():
                raise ValueError('Input workbook has different VBA than compiled master')
            patches=prepare_patches(zin,records)
            # Atomic publishing only after validation and full verification.
            outdir.mkdir(parents=True,exist_ok=False)
            final=outdir/'AgriX_MIG11_LINKS_PATCHED.xlsm'
            partial=outdir/'AgriX_MIG11_LINKS_PATCHED.xlsm.part'
            try:
                with zipfile.ZipFile(partial,'w') as zout:
                    for item in zin.infolist():
                        payload=patches.get(item.filename)
                        if payload is None:payload=zin.read(item.filename)
                        zout.writestr(copy.copy(item),payload)
                with zipfile.ZipFile(partial) as z:
                    if z.testzip():raise ValueError('Output XLSM failed ZIP CRC')
                    for name in zin.namelist():
                        if name not in patches and zin.read(name)!=z.read(name):
                            raise ValueError('Unintended change to XLSM OOXML part')
                outcome=check_workbook(partial,master,stage)
                if not outcome['pass']:
                    raise ValueError('Post-write XLSM verification failed: '+str(outcome['errors']))
                os.replace(partial,final)
                outcome.update({
                    'format':'AGRIx-MIG-11-XLSM-PATCH-VERIFY-v2',
                    'workbook':final.name,
                    'source_workbook_sha256':sha256(source.read_bytes()).hexdigest(),
                    'result_workbook_sha256':sha256(final.read_bytes()).hexdigest(),
                    'modified_parts':sorted(patches),
                    'all_other_xlsm_parts_byte_identical':True,
                    'original_input_master_and_stage_untouched':True,
                    'note':'Technical data import PASS; Excel compile/reopen and cancelled-history runtime validation still pending.'
                })
                report=outdir/'mig11_final_verify.json'
                report.write_text(json.dumps(outcome,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
                print('MIG-11 XLSM PATCH AND VERIFY: PASS')
                print('LINKS:',{k:v['actual_rows'] for k,v in outcome['tables'].items()})
                print('UNCHANGED OOXML parts:',len(zin.namelist())-len(patches))
                print('VBA IDENTICAL: YES')
                print('OUTPUT:',final)
                print('SHARE ONLY:',report)
                return 0
            except Exception:
                if partial.exists():partial.unlink()
                if final.exists():final.unlink()
                if outdir.exists() and not any(outdir.iterdir()):outdir.rmdir()
                raise
    except (ValueError,OSError,KeyError,ET.ParseError,zipfile.BadZipFile,json.JSONDecodeError) as exc:
        print('MIG-11 PATCH BLOCKED:',str(exc),file=sys.stderr)
        return 2

if __name__=='__main__':sys.exit(main())
