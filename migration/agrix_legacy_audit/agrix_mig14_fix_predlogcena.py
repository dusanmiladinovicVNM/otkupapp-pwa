#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-14/OTP-00440: remove ONLY one legacy zero from optional PredlogCena.

Python 3.10+, stdlib only. Reads verified staging and saved MIG-14 QA workbook.
Creates a NEW XLSM copy with ONLY ONE worksheet XML element changed. It never
changes source, existing workbooks, VBA binary, table definitions, or Git.
"""
from __future__ import annotations

import argparse
import copy
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK = 'http://schemas.openxmlformats.org/package/2006/relationships'
Q = lambda s: '{'+NS+'}'+s
TABLE = 'tblOtpremnicaStavke'
DOC_ID = 'OTP-00440'  # historical exception, no customer data in log output
EXPECTED_ROWS = 907
EXPECTED_INPUT_SHA = '8dedb685aae4d69948f526c6983302cb4f23e77afd7d57eb39d88946d749bf33'
STAGE_CHECKS = {
    'tblOtpremnicaStavke.jsonl':'c6164d6dc730b491c10f9bcce08c25079bbb2db148856e1c3c699903387e40db',
    'SOURCE__tblOtpremnica.jsonl':'d02436aa32a0e14ae6ca19172227f83b358b3ab0e329a7f040f97d29748e611a',
}
RANGE = re.compile(r'^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$',re.I)
COORD = re.compile(r'^\$?([A-Z]+)\$?(\d+)$',re.I)
SD_RE = re.compile(rb'(<sheetData\b[^>]*>)(.*?)(</sheetData>)',re.DOTALL)
ROW_RE = re.compile(rb'<row\b[^>]*(?:/>|>.*?</row>)',re.DOTALL)
CELL_RE = re.compile(rb'<c\b[^>]*(?:/>|>.*?</c>)',re.DOTALL)
R_ATTR_RE = re.compile(rb'\br\s*=\s*["\']([^"\']+)["\']')
V_RE = re.compile(rb'<v(?:\s[^>]*)?>([^<]*)</v>')


def sha(p):
    h=hashlib.sha256()
    with open(p,'rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()


def colidx(s):
    n=0
    for c in s.upper():n=n*26+ord(c)-64
    return n


def colname(n):
    s=''
    while n:n,k=divmod(n-1,26);s=chr(65+k)+s
    return s


def cellxy(addr):
    m=COORD.fullmatch(addr)
    if not m:raise ValueError('Invalid cell reference')
    return colidx(m.group(1)),int(m.group(2))


def decode_ref(s):
    m=RANGE.fullmatch(s)
    if not m:raise ValueError('Unsupported table range')
    return colidx(m.group(1)),int(m.group(2)),colidx(m.group(3)),int(m.group(4))


def rels(z,owner):
    rp=posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')
    if rp not in z.namelist():return {}
    root=ET.fromstring(z.read(rp)); out={}
    for r in root.findall('{'+PK+'}Relationship'):
        if r.attrib.get('TargetMode')=='External':continue
        target=r.attrib.get('Target','')
        out[r.attrib['Id']]=(target.lstrip('/') if target.startswith('/') else
            posixpath.normpath(posixpath.join(posixpath.dirname(owner),target)))
    return out


def find_table(z):
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    wbrel=rels(z,'xl/workbook.xml')
    found=[]
    for s in wb.findall('.//'+Q('sheets')+'/'+Q('sheet')):
        sp=wbrel.get(s.attrib.get('{'+RD+'}id'))
        if sp is None:raise ValueError('Broken workbook relationship')
        for tp in rels(z,sp).values():
            if not tp.startswith('xl/tables/') or not tp.endswith('.xml'):continue
            tbl=ET.fromstring(z.read(tp))
            if tbl.attrib.get('name')==TABLE or tbl.attrib.get('displayName')==TABLE:
                found.append((sp,tp,tbl))
    if len(found)!=1:raise ValueError('Exactly one linked target table is required')
    return found[0]


def stringtable(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter(Q('t'))) for si in root.findall(Q('si'))]


def read_cell(c,shared):
    if c is None:return ''
    if c.find(Q('f')) is not None:raise ValueError('Formula in canonical imported table')
    typ=c.attrib.get('t','')
    if typ=='inlineStr':return ''.join(t.text or '' for t in c.iter(Q('t')))
    v=c.find(Q('v'))
    if v is None:return ''
    if typ=='s':return shared[int(v.text or '-1')]
    if typ=='e':raise ValueError('Excel cell error in canonical table')
    return v.text or ''


def classify(v):
    if v is None or str(v).strip()=='':return 'empty'
    try:
        a=Decimal(str(v).strip().replace(',','.'))
        if not a.is_finite():return 'invalid'
    except InvalidOperation:return 'invalid'
    return 'zero' if a==0 else ('positive' if a>0 else 'negative')


def load_stage(stage, checks=STAGE_CHECKS, expected_rows=EXPECTED_ROWS):
    rows={}
    for name,h in checks.items():
        path=stage/name
        if not path.is_file() or sha(path)!=h:raise ValueError('Historical stage checksum mismatch: '+name)
        with path.open(encoding='utf-8-sig') as f:
            rows[name]=[json.loads(line) for line in f if line.strip()]
    items=rows['tblOtpremnicaStavke.jsonl']
    if len(items)!=expected_rows:raise ValueError('Historical stage row count changed')
    counts={k:sum(classify(row.get('PredlogCena'))==k for row in items) for k in ('positive','zero','empty','negative','invalid')}
    if counts!={'positive':expected_rows-1,'zero':1,'empty':0,'negative':0,'invalid':0}:
        raise ValueError('Stage does not contain exactly one historic zero proposed price')
    zero=[row for row in items if classify(row.get('PredlogCena'))=='zero']
    if zero[0].get('OtpremnicaID')!=DOC_ID:raise ValueError('Historic zero price belongs to unexpected document')
    source=[row for row in rows['SOURCE__tblOtpremnica.jsonl'] if row.get('OtpremnicaID')==DOC_ID]
    if len(source)!=1 or classify(source[0].get('Cena'))!='zero':
        raise ValueError('Source does not prove a matching historic price of zero')
    return items


def inspect_workbook(z,expected_rows=EXPECTED_ROWS):
    if z.testzip():raise ValueError('Damaged XLSM ZIP')
    names=z.namelist()
    if len(names)!=len(set(names)):raise ValueError('Duplicate ZIP members')
    if 'xl/vbaProject.bin' not in names:raise ValueError('Missing compiled VBA binary')
    sp,tp,tab=find_table(z)
    columns=[c.attrib['name'] for c in tab.find(Q('tableColumns')).findall(Q('tableColumn'))]
    if not {'OtpremnicaID','PredlogCena'}<=set(columns):raise ValueError('Target table columns absent')
    if len(columns)!=len(set(columns)):raise ValueError('Duplicate table header')
    left,head,right,last=decode_ref(tab.attrib['ref'])
    if len(columns)!=right-left+1 or last-head!=expected_rows:raise ValueError('Unexpected table dimensions')
    if tab.attrib.get('totalsRowCount','0')!='0':raise ValueError('Table totals unsupported')
    if tab.attrib.get('headerRowCount','1')!='1':raise ValueError('Unexpected table header')
    root=ET.fromstring(z.read(sp))
    data=root.find(Q('sheetData'))
    if data is None:raise ValueError('Missing worksheet cells')
    mapping={}
    wanted_ids={}
    shared=stringtable(z)
    for row in data.findall(Q('row')):
        rn=int(row.attrib['r'])
        if not head<rn<=last:continue
        cells={}
        for cell in row.findall(Q('c')):
            ci,rr=cellxy(cell.attrib.get('r',''))
            if rr!=rn:raise ValueError('Cell outside its row')
            if ci in cells:raise ValueError('Duplicate worksheet cell')
            cells[ci]=cell
        oid=read_cell(cells.get(left+columns.index('OtpremnicaID')),shared).strip()
        price=read_cell(cells.get(left+columns.index('PredlogCena')),shared)
        mapping[rn]=(oid,classify(price),price)
        if oid==DOC_ID:wanted_ids[rn]=cells.get(left+columns.index('PredlogCena'))
    if len(mapping)!=expected_rows or len(wanted_ids)!=1:
        raise ValueError('Unexpected migrated row count or missing historical exception')
    counts={k:sum(v[1]==k for v in mapping.values()) for k in ('positive','zero','empty','negative','invalid')}
    if counts!={'positive':expected_rows-1,'zero':1,'empty':0,'negative':0,'invalid':0}:
        raise ValueError('Workbook candidate price pattern differs from historical staging')
    rowno=next(iter(wanted_ids))
    target=wanted_ids[rowno]
    if target is None or target.attrib.get('t','') not in ('','n') or target.find(Q('f')) is not None:
        raise ValueError('Target must be a literal numeric zero cell')
    if classify(read_cell(target,shared))!='zero':raise ValueError('Target is not numeric zero')
    if [k for k in target if k.tag!=Q('v')]:raise ValueError('Target cell contains unsupported metadata')
    addr=colname(left+columns.index('PredlogCena'))+str(rowno)
    return {'sheetpart':sp,'tablepart':tp,'cell':addr,'row':rowno,'counts':counts}


def patch_sheet(payload,addr,rowno):
    sd=SD_RE.search(payload)
    if not sd:raise ValueError('Unsupported sheetData markup')
    rows=list(ROW_RE.finditer(sd.group(2)))
    if not rows:raise ValueError('No worksheet rows')
    target_row=[]
    for r in rows:
        ma=R_ATTR_RE.search(r.group(0).split(b'>',1)[0])
        if ma and int(ma.group(1))==rowno:target_row.append(r)
    if len(target_row)!=1:raise ValueError('Worksheet target row not unique')
    m=target_row[0]
    rowraw=m.group(0)
    cells=[]
    for cm in CELL_RE.finditer(rowraw):
        ma=R_ATTR_RE.search(cm.group(0).split(b'>',1)[0])
        if ma and ma.group(1).decode('ascii').upper()==addr.upper():cells.append(cm)
    if len(cells)!=1:raise ValueError('Worksheet target cell not unique')
    cm=cells[0]
    raw=cm.group(0)
    val=V_RE.findall(raw)
    if len(val)!=1 or classify(val[0].decode('utf-8'))!='zero':
        raise ValueError('Zero value not found in expected OOXML shape')
    if b'<f' in raw or b'<is' in raw:raise ValueError('Refusing to change formula/string')
    changed=V_RE.sub(b'',raw,count=1)
    if changed==raw:raise ValueError('No surgical change')
    editedrow=rowraw[:cm.start()]+changed+rowraw[cm.end():]
    body=sd.group(2)
    editedbody=body[:m.start()]+editedrow+body[m.end():]
    result=payload[:sd.start(2)]+editedbody+payload[sd.end(2):]
    ET.fromstring(result)
    return result


def patch(input_xlsm, stage, outdir, expected_sha=EXPECTED_INPUT_SHA,
          checks=STAGE_CHECKS, expected_rows=EXPECTED_ROWS):
    src=input_xlsm.resolve();outdir=outdir.resolve();stage=stage.resolve()
    if not src.is_file() or not stage.is_dir():raise ValueError('Input workbook or stage absent')
    if outdir.exists():raise ValueError('Output directory exists; refusing overwrite')
    if outdir in (src.parent,stage) or outdir==src.parent.parent or stage in outdir.parents:
        raise ValueError('Output directory must be isolated from staging and input')
    if sha(src)!=expected_sha:raise ValueError('MIG-14 QA SHA differs from verified 15:14 input; do not patch an unknown revision')
    stage_rows=load_stage(stage,checks,expected_rows)
    with zipfile.ZipFile(src) as zin:
        info=inspect_workbook(zin,expected_rows)
        original=zin.read(info['sheetpart'])
        changed=patch_sheet(original,info['cell'],info['row'])
        if changed==original:raise ValueError('No change in output sheet')
        outdir.mkdir(parents=True,exist_ok=False)
        final=outdir/'AgriX_MIG14_PRICE_PATCHED.xlsm'
        partial=outdir/(final.name+'.part')
        try:
            with zipfile.ZipFile(partial,'w') as zout:
                zout.comment=zin.comment
                for member in zin.infolist():
                    zout.writestr(copy.copy(member),changed if member.filename==info['sheetpart'] else zin.read(member.filename))
            with zipfile.ZipFile(partial) as zout:
                after=ET.fromstring(zout.read(info['sheetpart']))
                original_root=ET.fromstring(original)
                if zout.testzip():raise ValueError('Written XLSM ZIP CRC failure')
                if len(zout.namelist())!=len(zin.namelist()) or set(zout.namelist())!=set(zin.namelist()):
                    raise ValueError('ZIP member inventory changed')
                for name in zin.namelist():
                    if name!=info['sheetpart'] and zin.read(name)!=zout.read(name):
                        raise ValueError('Unrelated OOXML part unexpectedly changed')
                if zin.read('xl/vbaProject.bin')!=zout.read('xl/vbaProject.bin'):
                    raise ValueError('VBA binary changed')
                if zin.read(info['tablepart'])!=zout.read(info['tablepart']):
                    raise ValueError('Excel table definition changed')
                after_info=inspect_workbook_after(zout,expected_rows)
                if after_info['target']!='empty' or after_info['positive']!=expected_rows-1:
                    raise ValueError('Price edit did not roundtrip as expected')
            os.replace(partial,final)
            report={
                'format':'AGRIx-MIG14-OTP-PREDLOGCENA-PATCH-v1',
                'status':'TECHNICAL_DATA_PATCH_PASS', 'pass':True,
                'input_workbook_sha256':expected_sha,
                'result_workbook_sha256':sha(final),
                'source_stage_rows_checked':len(stage_rows),
                'target_table':TABLE,
                'rows_in_table':expected_rows,
                'historic_zero_price_cells_cleared':1,
                'remaining_nonpositive_numeric_proposed_prices':0,
                'source_historic_zero_cena_preserved_in_stage':True,
                'modified_parts':[info['sheetpart']],
                'vba_binary_byte_identical_to_input':True,
                'all_other_ooxml_parts_byte_identical':True,
                'all_table_parts_byte_identical':True,
                'source_workbook_and_stage_untouched':True,
                'migrated_goods_and_financial_rows_unchanged':True,
                'historic_ambalaza_not_imported':True,
                'next_action':'Excel QA then use this new XLSM as baseline for packaging import',
                'customer_values_in_report':False
            }
            (outdir/'MIG14_price_final_verify.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
            print('MIG-14 PRICE PATCH: TECHNICAL PASS')
            print('One historical zero optional proposed price converted to blank; remaining positives:',expected_rows-1)
            print('Only one worksheet XML part modified; VBA preserved; all other parts byte-identical')
            print('SHARE ONLY:',outdir/'MIG14_price_final_verify.json')
        except BaseException:
            partial.unlink(missing_ok=True)
            final.unlink(missing_ok=True)
            (outdir/'MIG14_price_final_verify.json').unlink(missing_ok=True)
            if outdir.exists() and not any(outdir.iterdir()):outdir.rmdir()
            raise


def inspect_workbook_after(z,expected_rows):
    sp,tp,tab=find_table(z)
    cols=[c.attrib['name'] for c in tab.find(Q('tableColumns')).findall(Q('tableColumn'))]
    left,head,right,last=decode_ref(tab.attrib['ref'])
    data=ET.fromstring(z.read(sp)).find(Q('sheetData'))
    ss=stringtable(z)
    count=0;positives=0;target=None
    for row in data.findall(Q('row')):
        n=int(row.attrib['r'])
        if not head<n<=last:continue
        cells={cellxy(c.attrib['r'])[0]:c for c in row.findall(Q('c'))}
        oid=read_cell(cells.get(left+cols.index('OtpremnicaID')),ss).strip()
        price=classify(read_cell(cells.get(left+cols.index('PredlogCena')),ss))
        count+=1
        if oid==DOC_ID:
            if target is not None:raise ValueError('Duplicate target document after patch')
            target=price
        elif price=='positive':positives+=1
        else:raise ValueError('Unexpected proposed price outside target document')
    if count!=expected_rows:return {'target':target,'positive':positives,'count':count}
    return {'target':target,'positive':positives,'count':count}


def main(argv=None):
    ap=argparse.ArgumentParser(description='Clear one proven historical optional PredlogCena numeric zero in a new XLSM copy')
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--input-xlsm',type=Path,required=True)
    ap.add_argument('--out-dir',type=Path,required=True)
    args=ap.parse_args(argv)
    try:patch(args.input_xlsm,args.stage,args.out_dir)
    except Exception as e:
        print('MIG-14 PRICE PATCH BLOCKED:',type(e).__name__,'—',str(e),file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':sys.exit(main())
