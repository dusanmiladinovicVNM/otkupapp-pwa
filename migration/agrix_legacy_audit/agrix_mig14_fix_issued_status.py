#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-14 historic issued status: exact source-bound OOXML-only transformation.

Only historical Otpremnica and Zbirna headers, each sourced from pre-cutover
legacy program without a draft lifecycle, receive IzdatoStatus=IZDATO.
The Stornirano column is NEVER changed; an issued-and-later-cancelled document
is still a document that was originally issued.

Creates a NEW XLSM, never saves Excel via COM, never changes VBA or Git.
Requires a pinned verified input SHA and source JSONL checksums.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import os
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RD='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK='http://schemas.openxmlformats.org/package/2006/relationships'
Q=lambda s:'{'+NS+'}'+s
EXPECTED_INPUT_SHA='d22182bbee9555134015ee39e6e04668932d6f6ad973a265c0be6b0d90ed550c'
EXPECTED_ROWS=907
STATUS='IZDATO'
TABLES=(
    ('tblOtpremnica','OtpremnicaID','SOURCE__tblOtpremnica.jsonl','tblOtpremnica.jsonl',
     'd02436aa32a0e14ae6ca19172227f83b358b3ab0e329a7f040f97d29748e611a',
     'fec7aadeb5ad36d2e46058baa89d2c836e219a29d79051a3866727e1017e8531'),
    ('tblZbirna','ZbirnaID','SOURCE__tblZbirna.jsonl','tblZbirna.jsonl',
     'c4b66c1ca6b32771fe239b56af4917cd7a37ddefc6eba1efde5381f9715638e7',
     'b82259cc09e4f8168b5b38a77c45cf32b5a76524718e1f7a3a113fa524baf084')
)
RANGE=re.compile(r'^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$',re.I)
COORD=re.compile(r'^\$?([A-Z]+)\$?(\d+)$',re.I)
SD_RE=re.compile(rb'(<sheetData\b[^>]*>)(.*?)(</sheetData>)',re.S)
ROW_RE=re.compile(rb'<row\b[^>]*(?:/>|>.*?</row>)',re.S)
CELL_RE=re.compile(rb'<c\b[^>]*(?:/>|>.*?</c>)',re.S)
R_ATTR_RE=re.compile(rb'\br\s*=\s*["\']([^"\']+)["\']')
T_ATTR_RE=re.compile(rb'\bt\s*=\s*(["\'])[^"\']*\1')


def sha(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for block in iter(lambda:f.read(1<<20),b''):h.update(block)
    return h.hexdigest()

def colidx(s):
    n=0
    for c in s.upper():n=n*26+ord(c)-64
    return n

def colname(n):
    s=''
    while n:n,k=divmod(n-1,26);s=chr(65+k)+s
    return s

def refxy(addr):
    m=COORD.fullmatch(addr)
    if not m:raise ValueError('Invalid cell reference')
    return colidx(m.group(1)),int(m.group(2))

def refbounds(value):
    m=RANGE.fullmatch(value)
    if not m:raise ValueError('Unexpected table range format')
    return colidx(m.group(1)),int(m.group(2)),colidx(m.group(3)),int(m.group(4))

def rels(z,owner):
    rp=posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')
    if rp not in z.namelist():return {}
    result={}
    for relation in ET.fromstring(z.read(rp)).findall('{'+PK+'}Relationship'):
        if relation.attrib.get('TargetMode')=='External':continue
        target=relation.attrib.get('Target','')
        result[relation.attrib['Id']]=(target.lstrip('/') if target.startswith('/') else
              posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)))
    return result

def discover_tables(z,tabledefs=TABLES):
    names={t[0] for t in tabledefs}
    found={}
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    wrels=rels(z,'xl/workbook.xml')
    for sheet in wb.findall('.//'+Q('sheets')+'/'+Q('sheet')):
        sp=wrels.get(sheet.attrib.get('{'+RD+'}id'))
        if not sp:raise ValueError('Missing workbook sheet relationship')
        for part in rels(z,sp).values():
            if not (part.startswith('xl/tables/') and part.endswith('.xml')):continue
            t=ET.fromstring(z.read(part))
            name=t.attrib.get('name') or t.attrib.get('displayName')
            if name in names:
                if name in found:raise ValueError('Duplicate workbook canonical table: '+name)
                found[name]=(sp,part,t)
    if set(found)!=names:raise ValueError('Missing canonical status tables: '+str(sorted(names-set(found))))
    return found

def stringtable(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    return [''.join(t.text or '' for t in si.iter(Q('t')))
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(Q('si'))]

def get_cell_value(c,shared):
    if c is None:return ''
    if c.find(Q('f')) is not None:raise ValueError('Formula in migrated status/identity cell')
    t=c.attrib.get('t','')
    if t=='inlineStr':return ''.join(x.text or '' for x in c.iter(Q('t')))
    v=c.find(Q('v'))
    if v is None or v.text is None:return ''
    if t=='s':
        try:return shared[int(v.text)]
        except (IndexError,ValueError) as ex:raise ValueError('Invalid shared string reference') from ex
    if t=='e':raise ValueError('Error cell in migration table')
    return v.text

def jsonl(path):
    with path.open(encoding='utf-8-sig') as f:
        for line in f:
            if line.strip():yield json.loads(line)

def stage_ids(stage,tables=TABLES,n=EXPECTED_ROWS):
    out={}
    for name,id_col,source_file,stage_file,source_sha,stage_sha in tables:
        for fn,digest in ((source_file,source_sha),(stage_file,stage_sha)):
            path=stage/fn
            if not path.is_file() or sha(path)!=digest:
                raise ValueError('Historical source checksum mismatch: '+fn)
        ids=[]
        for fn in (source_file,stage_file):
            rows=list(jsonl(stage/fn))
            if len(rows)!=n:raise ValueError('Unexpected historical count in '+fn)
            keys=[]
            for row in rows:
                key=str(row.get(id_col,'')).strip()
                if not key:raise ValueError('Empty historic header key in '+fn)
                if str(row.get('IzdatoStatus','')).strip():
                    raise ValueError('Historical source includes explicit issued status in '+fn)
                keys.append(key)
            if len(keys)!=len(set(keys)):raise ValueError('Duplicate historical header ID in '+fn)
            ids.append(set(keys))
        if ids[0]!=ids[1]:raise ValueError('Historical stage diverged from original header IDs for '+name)
        out[name]=ids[0]
    return out

def examine_table(z,tdef,found,shared,expected_ids,n=EXPECTED_ROWS):
    table,id_col,*_=tdef
    sp,tp,tab=found[table]
    cols=[c.attrib['name'] for c in tab.find(Q('tableColumns')).findall(Q('tableColumn'))]
    if len(cols)!=len(set(cols)) or id_col not in cols or 'IzdatoStatus' not in cols or 'Stornirano' not in cols:
        raise ValueError('Incomplete/duplicate canonical columns: '+table)
    l,h,r,last=refbounds(tab.attrib['ref'])
    if r-l+1!=len(cols) or last-h!=n:
        raise ValueError('Historical target table dimensions mismatch: '+table)
    if tab.attrib.get('headerRowCount','1')!='1' or tab.attrib.get('totalsRowCount','0')!='0':
        raise ValueError('Unexpected table header or total row: '+table)
    sd=ET.fromstring(z.read(sp)).find(Q('sheetData'))
    if sd is None:raise ValueError('No worksheet rows: '+table)
    id_ix=l+cols.index(id_col);status_ix=l+cols.index('IzdatoStatus')
    observed={}; targets={}; nonempty=0
    for row in sd.findall(Q('row')):
        rn=int(row.attrib['r'])
        if rn<=h or rn>last:continue
        cells={}
        for c in row.findall(Q('c')):
            ci,rr=refxy(c.attrib['r'])
            if rr!=rn or ci in cells:raise ValueError('Invalid/duplicate worksheet cell')
            cells[ci]=c
        key=get_cell_value(cells.get(id_ix),shared).strip()
        if key in observed or not key:raise ValueError('Duplicate/empty target header identity: '+table)
        observed[key]=rn
        raw=get_cell_value(cells.get(status_ix),shared).strip()
        if raw:nonempty+=1
        if raw not in ('',):raise ValueError('Pre-existing issued status in '+table+': refusing overwrite')
        targets[rn]=colname(status_ix)+str(rn)
    if len(observed)!=n or set(observed)!=expected_ids:
        raise ValueError('Migrated target IDs differ from original/stage: '+table)
    if nonempty:raise ValueError('Target statuses not blank: '+table)
    return {'table':table,'sheet':sp,'table_part':tp,'targets':targets,'before_empty':len(targets)}

def change_cell_raw(raw):
    m=re.match(rb'<c\b[^>]*?(?:/>|>)',raw,re.S)
    if not m:raise ValueError('Unable to read target cell opening tag')
    opening=m.group(0)
    if raw.endswith(b'/>'):
        body=b''
    else:
        if not raw.endswith(b'</c>'):raise ValueError('Unexpected cell closure')
        body=raw[len(opening):-4]
    if body.strip():
        parsed=ET.fromstring(raw)
        # Cell already proved semantically EMPTY in examine_table. Preserve
        # fail-closed behavior for metadata, formulas and unsupported children.
        if any(ch.tag.split('}')[-1] not in ('v','is') for ch in parsed):
            raise ValueError('Refusing to remove unsupported cell metadata')
        if any(ch.tag.split('}')[-1] != 't' for x in parsed if x.tag.split('}')[-1]=='is' for ch in x):
            raise ValueError('Refusing to remove unexpected inline XML')
        if ''.join(parsed.itertext()).strip():
            raise ValueError('Refusing to overwrite nonempty status cell')
    attrs=opening[2:-2] if opening.endswith(b'/>') else opening[2:-1]
    if T_ATTR_RE.search(attrs):attrs=T_ATTR_RE.sub(b't="inlineStr"',attrs,count=1)
    else:attrs+=b' t="inlineStr"'
    return b'<c'+attrs+b'><is><t>IZDATO</t></is></c>'

def patch_sheet(raw,targets):
    sd=SD_RE.search(raw)
    if not sd:raise ValueError('Unexpected sheetData encoding')
    body=sd.group(2)
    rows=list(ROW_RE.finditer(body))
    seen=set();edits=[]
    for rm in rows:
        rowraw=rm.group(0)
        start=rowraw.split(b'>',1)[0]
        m=R_ATTR_RE.search(start)
        if not m:continue
        rn=int(m.group(1))
        if rn not in targets:continue
        addr=targets[rn].encode('ascii')
        if rn in seen:raise ValueError('Duplicate target row')
        seen.add(rn)
        cells=list(CELL_RE.finditer(rowraw))
        target=[];next_cell=None
        target_ix=refxy(addr.decode())[0]
        for cm in cells:
            a=R_ATTR_RE.search(cm.group(0).split(b'>',1)[0])
            if not a:raise ValueError('Cell without an address')
            raw_addr=a.group(1)
            ci,rr=refxy(raw_addr.decode())
            if rr!=rn:raise ValueError('Row/cell address mismatch')
            if raw_addr.upper()==addr.upper():target.append(cm)
            if ci>target_ix and next_cell is None:next_cell=cm
        if len(target)>1:raise ValueError('Duplicate historical status cell')
        if target:
            cm=target[0]
            newcell=change_cell_raw(cm.group(0))
            newrow=rowraw[:cm.start()]+newcell+rowraw[cm.end():]
        else:
            newcell=b'<c r="'+addr+b'" t="inlineStr"><is><t>IZDATO</t></is></c>'
            if next_cell is None:
                cut=rowraw.rfind(b'</row>')
                if cut<0:raise ValueError('Unsupported self closing data row')
            else:cut=next_cell.start()
            newrow=rowraw[:cut]+newcell+rowraw[cut:]
        edits.append((rm.start(),rm.end(),newrow))
    if seen!=set(targets):raise ValueError('Incomplete targeted rows; refuse partial status import')
    newbody=body
    for s,e,x in reversed(edits):newbody=newbody[:s]+x+newbody[e:]
    changed=raw[:sd.start(2)]+newbody+raw[sd.end(2):]
    ET.fromstring(changed)
    if changed==raw:raise ValueError('Patch made no changes')
    return changed

def patch(input_xlsm,stage,outdir,expected_sha=EXPECTED_INPUT_SHA,defs=TABLES,n=EXPECTED_ROWS):
    src=input_xlsm.resolve();stage=stage.resolve();outdir=outdir.resolve()
    if not src.is_file() or not stage.is_dir():raise ValueError('Missing input XLSM or stage')
    if outdir.exists():raise ValueError('Output directory already exists; no overwrite permitted')
    if outdir==src.parent or outdir==stage or stage in outdir.parents:
        raise ValueError('Output must be separate from original XLSM and stage')
    if sha(src)!=expected_sha:
        raise ValueError('Input XLSM SHA differs from verified MIG-14 PRICE baseline: do not patch an unknown revision')
    ids=stage_ids(stage,defs,n)
    with zipfile.ZipFile(src) as zin:
        if zin.testzip():raise ValueError('Damaged XLSM input')
        names=zin.namelist()
        if len(names)!=len(set(names)) or 'xl/vbaProject.bin' not in names:
            raise ValueError('Duplicate ZIP members or missing compiled VBA')
        found=discover_tables(zin,defs)
        ss=stringtable(zin)
        info=[examine_table(zin,d,found,ss,ids[d[0]],n) for d in defs]
        if len({x['sheet'] for x in info})!=len(info):raise ValueError('Unexpected shared worksheet; migration blocked')
        patches={t['sheet']:patch_sheet(zin.read(t['sheet']),t['targets']) for t in info}
        outdir.mkdir(parents=True,exist_ok=False)
        result=outdir/'AgriX_MIG14_ISSUED_STATUS_PATCHED.xlsm'
        tmp=outdir/(result.name+'.part')
        try:
            with zipfile.ZipFile(tmp,'w') as zout:
                zout.comment=zin.comment
                for member in zin.infolist():
                    zout.writestr(copy.copy(member),patches.get(member.filename,zin.read(member.filename)))
            with zipfile.ZipFile(tmp) as zout:
                if zout.testzip():raise ValueError('Output CRC failure')
                if set(zout.namelist())!=set(names) or len(zout.namelist())!=len(names):
                    raise ValueError('XLSM ZIP structure changed')
                for member in names:
                    before=zin.read(member);after=zout.read(member)
                    if member in patches:
                        if after!=patches[member]:raise ValueError('Wrong target sheet XML output')
                    elif before!=after:raise ValueError('Unrelated workbook part changed: '+member)
                if zin.read('xl/vbaProject.bin')!=zout.read('xl/vbaProject.bin'):
                    raise ValueError('Compiled VBA changed')
                newinfo=[examine_table_after(zout,d,discover_tables(zout,defs),stringtable(zout),ids[d[0]],n) for d in defs]
            os.replace(tmp,result)
            report={
                'format':'AGRIx-MIG14-LEGACY-ISSUED-STATUS-PATCH-v1',
                'status':'TECHNICAL_DATA_PATCH_PASS',
                'pass':True,
                'historical_issuance_policy':'Old pre-2026-09 program had no draft lifecycle; legacy document headers that match source are issued, even when later stornirano',
                'input_workbook_sha256':expected_sha,
                'result_workbook_sha256':sha(result),
                'per_table':{t['table']:{'historical_rows_checked':n,'status_izdato_set':n,'status_empty_after':0,'other_columns_untouched':True} for t in info},
                'modified_parts':sorted(patches),
                'vba_binary_byte_identical':True,
                'all_other_ooxml_parts_byte_identical':True,
                'all_table_parts_byte_identical':True,
                'stornirano_and_other_business_values_unchanged':True,
                'source_stage_and_input_untouched':True,
                'customer_values_excluded':True,
                'excel_runtime_qa_pending':True
            }
            (outdir/'MIG14_issued_status_final_verify.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
            print('MIG-14 ISSUED STATUS: TECHNICAL DATA PATCH PASS')
            for t in info:print(t['table']+': '+str(n)+' verified historic statuses -> IZDATO')
            print('Only '+str(len(patches))+' worksheet XML parts changed; VBA, data, tables and other parts preserved')
            print('SHARE ONLY:',outdir/'MIG14_issued_status_final_verify.json')
        except BaseException:
            tmp.unlink(missing_ok=True)
            result.unlink(missing_ok=True)
            (outdir/'MIG14_issued_status_final_verify.json').unlink(missing_ok=True)
            if outdir.exists() and not any(outdir.iterdir()):outdir.rmdir()
            raise

def examine_table_after(z,tdef,found,ss,expected_ids,n):
    name,id_col,*_=tdef
    sp,tp,table=found[name]
    cols=[c.attrib['name'] for c in table.find(Q('tableColumns')).findall(Q('tableColumn'))]
    left,h,right,last=refbounds(table.attrib['ref'])
    id_ix=left+cols.index(id_col);status_ix=left+cols.index('IzdatoStatus')
    sd=ET.fromstring(z.read(sp)).find(Q('sheetData'))
    rows=[]
    for r in sd.findall(Q('row')):
        ri=int(r.attrib['r'])
        if not h<ri<=last:continue
        cells={refxy(c.attrib['r'])[0]:c for c in r.findall(Q('c'))}
        ident=get_cell_value(cells.get(id_ix),ss).strip()
        status=get_cell_value(cells.get(status_ix),ss).strip()
        if status!=STATUS:raise ValueError('Roundtrip issued status not correct')
        rows.append(ident)
    if len(rows)!=n or len(set(rows))!=n or set(rows)!=expected_ids:
        raise ValueError('Roundtrip migrated IDs changed')
    return len(rows)

def main():
    ap=argparse.ArgumentParser(description='Set historical issued lifecycle status in isolated, source-verified XLSM')
    ap.add_argument('--stage',required=True,type=Path)
    ap.add_argument('--input-xlsm',required=True,type=Path)
    ap.add_argument('--out-dir',required=True,type=Path)
    a=ap.parse_args()
    try:patch(a.input_xlsm,a.stage,a.out_dir)
    except Exception as e:
        print('MIG-14 ISSUED STATUS PATCH BLOCKED:',type(e).__name__,'—',str(e),file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':sys.exit(main())
