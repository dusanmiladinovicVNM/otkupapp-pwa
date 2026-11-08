#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-17.2: Finance date-format OOXML patch, NO business-value changes.

Only the 39 numeric Datum cells of tblNovac receive an already-present
Excel date style copied from tblOtkup.Datum. VBA and all other ZIP parts
remain byte-identical to the input. Creates a NEW .xlsm only. No Excel COM.
"""
from __future__ import annotations
import argparse, copy, hashlib, json, os, posixpath, re, sys, zipfile
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
Q=lambda s:'{'+MAIN+'}'+s
INPUT_SHA='ca1bf03c2654020a1bfe5f1175024d45898cf427d42e82c7a97c85da9315bc02'
EPOCH=date(1899,12,30)
COORD=re.compile(r'^([A-Z]+)(\d+)$')
RANGE=re.compile(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$')
SD_RE=re.compile(rb'(<sheetData\b[^>]*>)(.*?)(</sheetData>)',re.S)
ROW_RE=re.compile(rb'<row\b[^>]*(?:/>|>.*?</row>)',re.S)
CELL_RE=re.compile(rb'<c\b[^>]*(?:/>|>.*?</c>)',re.S)
RATTR=re.compile(rb'\br\s*=\s*["\']([^"\']+)["\']')
SATTR=re.compile(rb'\bs\s*=\s*["\'][^"\']+["\']')

def sha(path):
    h=hashlib.sha256()
    with open(path,'rb') as f:
        for c in iter(lambda:f.read(1<<20),b''):h.update(c)
    return h.hexdigest()

def idx(x):
    z=0
    for ch in x.upper():z=z*26+ord(ch)-64
    return z

def coord(s):
    m=COORD.fullmatch(s)
    if not m:raise ValueError('Unsupported coordinate')
    return idx(m[1]),int(m[2])

def bounds(s):
    m=RANGE.fullmatch(s.replace('$',''))
    if not m:raise ValueError('Unsupported table reference')
    return idx(m[1]),int(m[2]),idx(m[3]),int(m[4])

def rels(z,owner):
    rp=posixpath.join(posixpath.dirname(owner),'_rels',posixpath.basename(owner)+'.rels')
    if rp not in z.namelist():return {}
    out={}
    for r in ET.fromstring(z.read(rp)).findall('{'+PKG+'}Relationship'):
        if r.get('TargetMode')=='External':continue
        dest=r.get('Target','')
        dest=(dest.lstrip('/') if dest.startswith('/') else
              posixpath.normpath(posixpath.join(posixpath.dirname(owner),dest)))
        out[r.attrib['Id']]=dest
    return out

def find(z):
    wanted={'tblOtkup','tblNovac','tblAmbalaza','tblAmbalazaDokument'}
    found={};wb=ET.fromstring(z.read('xl/workbook.xml'));wrels=rels(z,'xl/workbook.xml')
    for sheet in wb.findall('.//'+Q('sheets')+'/'+Q('sheet')):
        sheetpart=wrels.get(sheet.get('{'+OFFICE+'}id'))
        if not sheetpart:raise ValueError('Missing sheet linkage')
        for part in rels(z,sheetpart).values():
            if not part.startswith('xl/tables/') or not part.endswith('.xml'):continue
            tbl=ET.fromstring(z.read(part));name=tbl.get('name') or tbl.get('displayName')
            if name in wanted:
                if name in found:raise ValueError('Duplicate canonical table')
                cols=[x.get('name') for x in tbl.find(Q('tableColumns')).findall(Q('tableColumn'))]
                found[name]=(sheetpart,part,bounds(tbl.get('ref')),cols)
    if set(found)!=wanted:raise ValueError('Missing MIG17 canonical table(s)')
    return found

def isdate_style_list(z):
    styles=ET.fromstring(z.read('xl/styles.xml'))
    fmts={int(x.get('numFmtId')):x.get('formatCode','') for x in styles.findall(Q('numFmts')+'/'+Q('numFmt'))}
    xfs=styles.find(Q('cellXfs'))
    if xfs is None:raise ValueError('No Excel cellXfs')
    out=[]
    for xf in xfs.findall(Q('xf')):
        n=int(xf.get('numFmtId','0'))
        if n in set(range(14,23))|set(range(27,37))|set(range(45,48))|set(range(50,59)):
            out.append(True);continue
        f=fmts.get(n,'')
        f=re.sub(r'"[^"]*"|\\.|\[[^]]*\]','',f)
        out.append(bool(re.search(r'[dmyhs]',f,re.I)))
    return out

def table_date_cells(z,found,name):
    sp,tp,(c1,r1,c2,r2),cols=found[name]
    if len(cols)!=c2-c1+1 or 'Datum' not in cols:
        raise ValueError('Unexpected table columns: '+name)
    needed=c1+cols.index('Datum')
    sd=ET.fromstring(z.read(sp)).find(Q('sheetData'))
    if sd is None:raise ValueError('No sheetData: '+name)
    rows=[]
    for row in sd.findall(Q('row')):
        rn=int(row.get('r','0'))
        if rn<=r1 or rn>r2:continue
        cs=[]
        for c in row.findall(Q('c')):
            ci,ri=coord(c.get('r',''))
            if ri==rn and ci==needed:cs.append(c)
        if len(cs)!=1:raise ValueError('Expected one numeric Datum cell per row: '+name)
        c=cs[0]
        if c.find(Q('f')) is not None or c.get('t') in ('s','inlineStr','e','str'):
            raise ValueError('Non-numeric/formula date cell: '+name)
        v=c.find(Q('v'))
        if v is None or not(v.text or '').strip():raise ValueError('Blank date cell: '+name)
        try:dt=Decimal(v.text)
        except InvalidOperation as exc:raise ValueError('Invalid Excel date serial: '+name) from exc
        if dt!=dt.to_integral_value():raise ValueError('Unexpected non-midnight historical date')
        d=EPOCH+timedelta(days=int(dt))
        if not date(2026,1,1)<=d<=date(2026,8,31):
            raise ValueError('Historical date outside expected 2026 scope')
        rows.append((c.get('r'),int(c.get('s','0')),v.text))
    if len(rows)!=(r2-r1):raise ValueError('Date row count diverged in '+name)
    return sp,rows

def rewrite(raw,targets,style):
    sd=SD_RE.search(raw)
    if not sd:raise ValueError('Unsupported sheetData serialization')
    body=sd.group(2);seen=set()
    def cell_sub(m):
        cell=m.group(0)
        attr=cell.split(b'>',1)[0]
        found=RATTR.search(attr)
        if not found:return cell
        key=found.group(1).decode('ascii')
        if key not in targets:return cell
        if key in seen:raise ValueError('Duplicate edited cell reference')
        seen.add(key)
        head=re.match(rb'<c\b[^>]*?(?:/>|>)',cell,re.S)
        if head is None:raise ValueError('Bad cell start')
        opening=head.group(0)
        if opening.endswith(b'/>'):raise ValueError('Unexpected empty financial date cell')
        first=opening[:-1]
        if SATTR.search(first):first=SATTR.sub(b's="'+str(style).encode()+b'"',first,count=1)
        else:first+=b' s="'+str(style).encode()+b'"'
        return first+b'>'+cell[len(opening):]
    def row_sub(m):return CELL_RE.sub(cell_sub,m.group(0))
    after=ROW_RE.sub(row_sub,body)
    if seen!=targets:raise ValueError('Not all financial date cells were patched')
    return raw[:sd.start(2)]+after+raw[sd.end(2):]

def run(input_path,output_dir,expected_sha=INPUT_SHA):
    input_path=input_path.resolve();output_dir=output_dir.resolve()
    if not input_path.is_file():raise ValueError('Input workbook does not exist')
    if not input_path.name.endswith('.xlsm'):raise ValueError('Only XLSM source supported')
    if expected_sha and sha(input_path)!=expected_sha:
        raise ValueError('Input SHA differs from MIG17.1 verified workbook: do not patch an unverified version')
    if output_dir.exists():raise ValueError('Output directory exists; choose new one, do not overwrite')
    if output_dir in input_path.parents or output_dir==input_path.parent:
        raise ValueError('Invalid output directory')
    with zipfile.ZipFile(input_path,'r') as z:
        if z.testzip():raise ValueError('Invalid ZIP archive')
        if 'xl/vbaProject.bin' not in z.namelist():raise ValueError('VBA binary missing')
        if len(z.namelist())!=len(set(z.namelist())):raise ValueError('Duplicate ZIP member name')
        found=find(z)
        if found['tblNovac'][2][3]-found['tblNovac'][2][1]!=39:raise ValueError('Expect 39 finance entries')
        if found['tblOtkup'][2][3]-found['tblOtkup'][2][1]!=1702:raise ValueError('Expect 1702 purchase headers')
        if found['tblAmbalaza'][2][3]-found['tblAmbalaza'][2][1]!=4811:raise ValueError('Expect 4811 packaging rows')
        if found['tblAmbalazaDokument'][2][3]-found['tblAmbalazaDokument'][2][1]!=48:raise ValueError('Expect 48 revers headers')
        date_styles=isdate_style_list(z)
        if not date_styles:raise ValueError('No styles')
        orig_sp,orig=table_date_cells(z,found,'tblOtkup')
        cand=[style for _,style,_ in orig]
        if not all(style<len(date_styles) and date_styles[style] for style in cand):
            raise ValueError('Original purchase dates not all formatted as dates')
        style=Counter(cand).most_common(1)[0][0]
        nov_sp,finance=table_date_cells(z,found,'tblNovac')
        if any(styleid<len(date_styles) and date_styles[styleid] for _,styleid,_ in finance):
            raise ValueError('Some tblNovac dates already date-formatted; refusing partial/duplicate patch')
        targets={coord for coord,_,_ in finance}
        raw_before=z.read(nov_sp);raw_after=rewrite(raw_before,targets,style)
        if raw_after==raw_before:raise ValueError('No patch written')
        # validate lexical patch changed no numeric values
        v1=[v for _,_,v in finance]
        original_members={name:hashlib.sha256(z.read(name)).hexdigest() for name in z.namelist()}
        output_dir.mkdir(parents=True)
        out=output_dir/'AgriX_MIG17_2_NOVAC_DATES_QA.xlsm'
        tmp=output_dir/(out.name+'.tmp')
        try:
            with zipfile.ZipFile(tmp,'w') as wr:
                for item in z.infolist():
                    wr.writestr(copy.copy(item),raw_after if item.filename==nov_sp else z.read(item.filename))
            with zipfile.ZipFile(tmp) as f:
                if f.testzip():raise ValueError('Patched ZIP test failed')
                newer=find(f)
                new_sp,updated=table_date_cells(f,newer,'tblNovac')
                if new_sp!=nov_sp or len(updated)!=39:raise ValueError('Postpatch finance mismatch')
                if [x[2] for x in updated]!=v1:raise ValueError('Financial date values changed')
                if not all(x[1]==style for x in updated):raise ValueError('Finance date style not applied')
                if not date_styles[style]:raise ValueError('New style not recognized as Excel date')
                altered=[name for name in original_members if original_members[name]!=hashlib.sha256(f.read(name)).hexdigest()]
                if altered!=[nov_sp]:raise ValueError('Unexpected modified OOXML/ZIP parts')
                if f.read('xl/vbaProject.bin')!=z.read('xl/vbaProject.bin'):
                    raise ValueError('VBA changed')
            os.replace(tmp,out)
        finally:
            if tmp.exists():tmp.unlink()
        summary={
            'format':'AGRIx-MIG172-FINANCE-DATE-STYLE-OOXML-v1',
            'status':'TECHNICAL_DATA_STYLE_PATCH_PASS','pass':True,
            'input_workbook_sha256':sha(input_path),
            'output_workbook_sha256':sha(out),
            'date_style_taken_from':'tblOtkup.Datum (existing verified Excel date style)',
            'finance_rows_checked':39,'finance_date_cells_corrected':39,
            'numeric_financial_date_values_preserved':True,
            'financial_amounts_and_other_business_cells_preserved':True,
            'modified_parts':[nov_sp],
            'vba_binary_identical':True,'all_other_ooxml_parts_identical':True,
            'finance_report_runtime_test_pending':True,
            'old_report_526812_50_not_yet_reconciled':True,
            'ambalaza_unmodified':True,'private_customer_values_in_summary':False}
        (output_dir/'MIG172_final_verify.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    print('MIG-17.2 DATA STYLE PATCH PASS; finance date cells corrected:',len(finance))
    print('XLSM:',out)
    print('SHARE ONLY:',output_dir/'MIG172_final_verify.json')
    return summary

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input-xlsm',type=Path,required=True)
    parser.add_argument('--out-dir',type=Path,required=True)
    a=parser.parse_args()
    try:run(a.input_xlsm,a.out_dir)
    except Exception as e:
        print('MIG17.2 BLOCKED:',type(e).__name__,'—',str(e),file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':sys.exit(main())
