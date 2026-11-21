#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIx MIG-11: read-only XLSM link-table and VBA-binary verification.

No Excel COM and no workbook modification. Python 3.10+, stdlib only.
"""
import argparse
import csv
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

SS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
PK='http://schemas.openxmlformats.org/package/2006/relationships'
RD='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REQUIRED={
 'tblOtpremnicaIzvori':('OtpremnicaIzvorID','OtpremnicaID','OtkupID'),
 'tblZbirnaIzvori':('ZbirnaIzvorID','ZbirnaID','OtpremnicaID'),
}
RANGE=re.compile(r'^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$',re.I)
CELL=re.compile(r'^\$?([A-Z]+)\$?(\d+)$',re.I)

def q(s):return '{'+SS+'}'+s

def sha(data):return hashlib.sha256(data).hexdigest()

def index_column(label):
    n=0
    for c in label.upper():n=n*26+ord(c)-ord('A')+1
    return n

def cell_position(pos):
    m=CELL.fullmatch(pos)
    if not m:raise ValueError('Invalid worksheet cell reference')
    return index_column(m.group(1)),int(m.group(2))

def resolve_path(from_file,target):
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(from_file),target))

def relationships(z,owner):
    relpath=posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')
    if relpath not in z.namelist():return {}
    root=ET.fromstring(z.read(relpath))
    result={}
    for rel in root.findall('{'+PK+'}Relationship'):
        if rel.attrib.get('TargetMode')=='External':continue
        result[rel.attrib['Id']]=resolve_path(owner,rel.attrib['Target'])
    return result

def shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter(q('t'))) for si in root.findall(q('si'))]

def string_value(cell,shared):
    typ=cell.attrib.get('t','')
    if typ=='inlineStr':
        return ''.join(t.text or '' for t in cell.iter(q('t')))
    v=cell.find(q('v'))
    if v is None:return ''
    value=v.text or ''
    if typ=='s':return shared[int(value)]
    if typ=='e':raise ValueError('Excel error cell in link table')
    return value

def parse_sheet(z, sheet_path, first_row, last_row, mincol, maxcol, shared):
    root=ET.fromstring(z.read(sheet_path))
    cells={}
    for row in root.findall('.//'+q('sheetData')+'/'+q('row')):
        r=int(row.attrib['r'])
        if r<=first_row or r>last_row:continue
        for cell in row.findall(q('c')):
            col, rr=cell_position(cell.attrib['r'])
            if rr==r and mincol<=col<=maxcol:
                cells[(r,col)]=string_value(cell,shared)
    return cells

def load_stage(stage,table):
    matches=[p for p in stage.rglob(table+'.jsonl') if p.is_file()]
    if len(matches)!=1:raise ValueError(f'{table}: expected one stage JSONL, found {len(matches)}')
    rows=[]
    with matches[0].open(encoding='utf-8-sig') as f:
        for line in f:
            if line.strip():
                row=json.loads(line)
                rows.append(tuple(str(row.get(k,'') or '').strip() for k in REQUIRED[table]))
    return rows

def check_workbook(xlsm,master,stage):
    outcome={"format":"AGRIx-MIG-11-XLSM-VERIFY-v1", "workbook":xlsm.name,
             "vba_binary_preserved":False, "tables":{}, "pass":False, "errors":[]}
    with zipfile.ZipFile(master) as m, zipfile.ZipFile(xlsm) as z:
        for check, archive in (("master",m),("output",z)):
            bad=archive.testzip()
            if bad:raise ValueError(check+': corrupt OOXML member')
        vba='xl/vbaProject.bin'
        if vba not in m.namelist() or vba not in z.namelist():
            raise ValueError('VBA binary missing from master or output')
        outcome['vba_binary_preserved']=(sha(m.read(vba))==sha(z.read(vba)))
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        wb_rel=relationships(z,'xl/workbook.xml')
        sheetpaths=[]
        for sh in wb.findall('.//'+q('sheets')+'/'+q('sheet')):
            rid=sh.attrib.get('{'+RD+'}id')
            if rid not in wb_rel:raise ValueError('Workbook sheet relationship missing')
            sheetpaths.append(wb_rel[rid])
        table_location={}
        for sp in sheetpaths:
            for path in relationships(z,sp).values():
                if path.startswith('xl/tables/') and path.endswith('.xml'):
                    table_location[path]=sp
        shared=shared_strings(z)
        tables={}
        for path in z.namelist():
            if not re.fullmatch(r'xl/tables/[^/]+\.xml',path):continue
            root=ET.fromstring(z.read(path))
            table=root.attrib.get('name') or root.attrib.get('displayName')
            if table in REQUIRED:
                if table in tables:raise ValueError('Duplicate XLSM table '+table)
                tables[table]=(root,path)
        for name, fields in REQUIRED.items():
            expected=load_stage(stage,name)
            row_result={"expected_rows":len(expected),"actual_rows":None,"exact_values_match":False}
            outcome['tables'][name]=row_result
            if name not in tables:
                outcome['errors'].append(name+': table not found in workbook')
                continue
            table,path=tables[name]
            if path not in table_location:
                outcome['errors'].append(name+': no worksheet relation to XML table')
                continue
            m_range=RANGE.fullmatch(table.attrib.get('ref',''))
            if not m_range:raise ValueError(name+': unsupported table ref')
            mincol,firstrow=index_column(m_range.group(1)),int(m_range.group(2))
            maxcol,lastrow=index_column(m_range.group(3)),int(m_range.group(4))
            headerrows=int(table.attrib.get('headerRowCount','1'))
            totals=int(table.attrib.get('totalsRowCount','0'))
            datafirst=firstrow+headerrows
            datalast=lastrow-totals
            header=table.find(q('tableColumns'))
            if header is None:raise ValueError(name+': missing table columns')
            columns=[c.attrib['name'] for c in header.findall(q('tableColumn'))]
            if len(set(columns))!=len(columns):raise ValueError(name+': duplicate table column names')
            for col in fields:
                if col not in columns:raise ValueError(name+': missing required column '+col)
            row_result['actual_rows']=max(0,datalast-datafirst+1)
            if row_result['actual_rows']!=len(expected):
                outcome['errors'].append(name+': row count differs')
                continue
            cells=parse_sheet(z,table_location[path],datafirst-1,datalast,mincol,maxcol,shared)
            found=[]
            for row in range(datafirst,datalast+1):
                found.append(tuple(cells.get((row,mincol+columns.index(key)),'').strip() for key in fields))
            row_result['exact_values_match']=Counter(found)==Counter(expected)
            if not row_result['exact_values_match']:
                outcome['errors'].append(name+': row values differ from canonical stage')
        if not outcome['vba_binary_preserved']:
            outcome['errors'].append('vbaProject.bin differs from compiled master')
    outcome['pass']=(not outcome['errors'])
    return outcome

def main(argv=None):
    ap=argparse.ArgumentParser(description='Verify MIG-11 imported XLSM link tables and VBA binary')
    ap.add_argument('--xlsm-dir',type=Path,required=True)
    ap.add_argument('--master',type=Path,required=True)
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args(argv)
    try:
        books=list(args.xlsm_dir.rglob('*.xlsm'))
        if len(books)!=1:raise ValueError(f'Expected exactly 1 output XLSM, found {len(books)}')
        if args.out.exists():raise ValueError('Report output already exists; refuse overwrite')
        result=check_workbook(books[0],args.master,args.stage)
        args.out.parent.mkdir(parents=True,exist_ok=True)
        args.out.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        print('MIG-11 XLSM VERIFY:', 'PASS' if result['pass'] else 'FAIL')
        print('TABLES:',{k:{'expected_rows':v['expected_rows'],'actual_rows':v['actual_rows'],'exact_match':v['exact_values_match']} for k,v in result['tables'].items()})
        print('VBA BINARY PRESERVED:',result['vba_binary_preserved'])
        print('SHARE ONLY:',args.out.name)
        return 0 if result['pass'] else 2
    except (ValueError,OSError,KeyError,ET.ParseError,zipfile.BadZipFile) as exc:
        print('MIG-11 VERIFY ERROR:',str(exc),file=sys.stderr)
        return 2

if __name__=='__main__':sys.exit(main())
