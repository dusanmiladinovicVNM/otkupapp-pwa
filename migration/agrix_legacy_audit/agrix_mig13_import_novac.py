#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIX MIG-13: verified, isolated import of 39 historical tblNovac records.

No Excel COM, no VBIDE, no edits to source, staging, original workbook or Git.
MIG-12 packaging history remains an archive, not a live ledger.
Python >=3.10, standard library only. Run from local agrix_legacy_audit directory.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import os
import re
import sys
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_EVEN
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from agrix_mig12_archive_mig13_finance import (
    EXPECTED, EXPECTED_MIG11_XLSM, FIN_CANON, preflight, locate, load_jsonl, sha, decimal_value
)
from agrix_mig11_write_links_ooxml import (
    q, table_locations, parse_ref, col_name, pos, patch_table_xml, dimension_patch,
    is_cell_blank, ROW_RE, SD_RE
)
from agrix_mig11_verify_xlsm import shared_strings, string_value, relationships, index_column, cell_position

EXPECTED_ROWS=39
AMOUNT_FIELDS={'Uplata','Isplata'}
DATE_FIELDS={'Datum','CreatedAt','ModifiedAt'}
EXPECTED_MASTER_SHA='0d5a4936df5481289d88ecfed5df22be16597806b5f86d9481d6be2df0676de6'
XML_CONTROL=re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def nonempty(x):
    return x is not None and not (isinstance(x,str) and not x.strip())


def valid_text(x):
    if isinstance(x,(dict,list,bool)):raise ValueError('Unexpected structured/bool cell value')
    value=str(x)
    if len(value)>32767 or XML_CONTROL.search(value):raise ValueError('Invalid Excel text value')
    return value


def canonical_amount(x):
    """Return (Excel-safe amount, treatment), preserving the source JSONL separately.

    Excel keeps at most 15 significant decimal digits. Zero padding is not
    precision; 1e-15 binary-float tails near an exact <=4dp amount are not
    meaningful currency. Anything beyond the strict tolerance is REJECTED.
    No unrestricted rounding and no text-cell fallback for money.
    """
    original=decimal_value(x)
    if original is None or not original.is_finite():
        raise ValueError('Nonfinite/invalid amount')
    if original<0:raise ValueError('Negative amount')

    # Strip exact decimal zeros without Decimal.normalize() context rounding.
    if original.is_zero():
        if len(original.as_tuple().digits)<=15: return original,'EXACT'
        return Decimal('0'),'EXACT_ZERO'
    sign,digits,exponent=original.as_tuple()
    if len(digits)<=15:return original,'EXACT'
    trimmed=list(digits)
    while trimmed and trimmed[-1]==0:
        trimmed.pop();exponent+=1
    exact=Decimal((sign,tuple(trimmed),exponent))
    if len(trimmed)<=15:
        return exact, ('EXACT_TRAILING_ZERO' if exact.as_tuple()!=original.as_tuple() else 'EXACT')

    # Exact binary/serialization tail only. Use the *finest* supported scale
    # (4 decimals), so no meaningful amount digits are lost if 3/4dp existed.
    # Cent-only rounding would be less conservative for such legacy values.
    nearest=exact.quantize(Decimal('0.0001'), rounding=ROUND_HALF_EVEN)
    ns,nd,ne=nearest.as_tuple()
    strip=list(nd)
    while strip and strip[-1]==0:
        strip.pop();ne+=1
    candidate=Decimal((ns,tuple(strip),ne)) if strip else Decimal('0')
    if abs(exact-candidate)<=Decimal('0.000000001') and len(candidate.as_tuple().digits)<=15:
        return candidate, 'FLOAT_TAIL_MAX_4DP'
    raise ValueError('Amount has unsupported real precision (not a harmless floating-point tail)')


def valid_decimal(x):
    return canonical_amount(x)[0]


def precision_summary(records):
    """Aggregate-only evidence; never put business values, IDs or partners in logs."""
    modes=Counter()
    by_field=Counter()
    rejected=Counter()
    for row in records:
        for field in ('Uplata','Isplata'):
            try:
                _,mode=canonical_amount(row.get(field,''))
            except ValueError:
                rejected[field]+=1
            else:
                modes[mode]+=1
                if mode=='FLOAT_TAIL_MAX_4DP':by_field[field]+=1
    return {'total_numeric_cells_checked':2*len(records),
            'float_tail_normalized':sum(by_field.values()),
            'float_tail_normalized_by_field':dict(sorted(by_field.items())),
            'exact_trailing_zero_or_zero_normalized':sum(v for k,v in modes.items() if k in ('EXACT_TRAILING_ZERO','EXACT_ZERO')),
            'unrepresentable_numeric_cells_by_field':dict(sorted(rejected.items()))}


def parse_date_serial(value):
    """Keep real sortable Excel dates. Reject ambiguous locale/epoch/offset inputs."""
    if isinstance(value,bool):raise ValueError('Boolean date')
    if isinstance(value,(int,float)):
        n=Decimal(str(value))
        if not n.is_finite() or not (Decimal(20000) < n < Decimal(80000)):
            raise ValueError('Unexpected numeric date serial')
        return n
    if not isinstance(value,str):raise ValueError('Unsupported date type')
    s=value.strip()
    if not s:raise ValueError('Empty required date')
    if re.fullmatch(r'\d{5}(?:\.\d+)?',s):
        n=Decimal(s)
        if not Decimal(20000)<n<Decimal(80000):raise ValueError('Invalid serial date range')
        return n
    formats=['%Y-%m-%d','%Y-%m-%d %H:%M:%S.%f','%Y-%m-%d %H:%M:%S',
             '%Y-%m-%d %H:%M','%Y-%m-%dT%H:%M:%S.%f','%Y-%m-%dT%H:%M:%S',
             '%Y-%m-%dT%H:%M','%d.%m.%Y.','%d.%m.%Y','%d.%m.%Y %H:%M:%S','%d.%m.%Y %H:%M']
    dt=None
    for fmt in formats:
        try:dt=datetime.strptime(s,fmt);break
        except ValueError:pass
    if dt is None:
        raise ValueError('Ambiguous date string (timezone, unsupported format or locale); no import')
    elapsed=dt-datetime(1899,12,30)
    if not 20000<elapsed.days<80000:raise ValueError('Date outside expected period')
    out=Decimal(elapsed.days)+Decimal(elapsed.seconds)/Decimal(86400)+Decimal(elapsed.microseconds)/Decimal(86400000000)
    return out


def cell_conversion(field,val):
    if not nonempty(val):return ('', '')
    if field in AMOUNT_FIELDS:
        return ('n',format(valid_decimal(val),'f'))
    if field in DATE_FIELDS:
        return ('n',format(parse_date_serial(val),'f'))
    return ('s',valid_text(val))


def render_rows(records,columns,c1,r1,stub_cells):
    styles={}
    for c in stub_cells:
        coord=c.attrib.get('r','')
        if coord:
            x,_=pos(coord)
            if c.attrib.get('s') is not None:styles[x]=c.attrib['s']
    chunks=[]
    for i,obj in enumerate(records):
        rn=r1+i;cells=[]
        for off,key in enumerate(columns):
            typ,value=cell_conversion(key,obj.get(key,''))
            if not typ:continue
            col=c1+off;ref=f'{col_name(col)}{rn}'
            style=(f' s="{styles[col]}"' if col in styles else '')
            if typ=='n':
                cells.append(f'<c r="{ref}"{style}><v>{value}</v></c>')
            else:
                preserve=' xml:space="preserve"' if value[:1].isspace() or value[-1:].isspace() else ''
                cells.append(f'<c r="{ref}"{style} t="inlineStr"><is><t{preserve}>{escape(value)}</t></is></c>')
        chunks.append(f'<row r="{rn}" spans="{c1}:{c1+len(columns)-1}">{"".join(cells)}</row>')
    return ''.join(chunks).encode('utf-8')


def patch_sheet(payload,table,cols,rows):
    root=ET.fromstring(payload)
    c1,r1,c2,r2=parse_ref(table.attrib['ref']);first=r1+1;last=r1+len(rows)
    if r2 not in (r1,r1+1):raise ValueError('tblNovac is not an empty table template')
    if table.attrib.get('headerRowCount','1')!='1' or table.attrib.get('totalsRowCount','0')!='0':
        raise ValueError('Unsupported table header/totals')
    if c2-c1+1!=len(cols):raise ValueError('tblNovac column width mismatch')
    if root.find(q('mergeCells')) is not None:raise ValueError('Merged cells in finance sheet')
    data=root.find(q('sheetData'))
    if data is None:raise ValueError('Missing sheetData')
    stubs=[]
    for row in data.findall(q('row')):
        r=int(row.attrib['r'])
        if first<=r<=last:
            for c in row.findall(q('c')):
                if not is_cell_blank(c):raise ValueError('Finance target has occupied cells — refusing overwrite')
                if r==first and c1<=pos(c.attrib['r'])[0]<=c2:stubs.append(c)
    rendered=render_rows(rows,cols,c1,first,stubs)
    match=SD_RE.search(payload)
    if not match:raise ValueError('Unsupported XML sheetData encoding')
    rowmatches=list(ROW_RE.finditer(match.group(2)))
    if len(rowmatches)!=len(data.findall(q('row'))):raise ValueError('Unsupported sheet row serialization')
    blocks=[];inserted=False;prior=0
    for m in rowmatches:
        raw=m.group(0);mat=re.search(rb'<row\b[^>]*\br\s*=\s*["\'](\d+)["\']',raw)
        if mat is None:raise ValueError('Row without r')
        rn=int(mat.group(1))
        if rn<=prior:raise ValueError('Nonmonotonic row indexes')
        prior=rn
        if rn>=first and not inserted:blocks.append(rendered);inserted=True
        if rn<first or rn>last:blocks.append(raw)
    if not inserted:blocks.append(rendered)
    edited=payload[:match.start(2)]+b''.join(blocks)+payload[match.end(2):]
    edited=dimension_patch(edited,c1,c2,r1,last)
    parse=ET.fromstring(edited)
    seen=[int(r.attrib['r']) for r in parse.find(q('sheetData')).findall(q('row'))]
    if seen!=sorted(set(seen)) or sum(first<=x<=last for x in seen)!=len(rows):raise ValueError('Bad worksheet row layout')
    return edited


def read_finance(z,tabpath,sheetpath,expected_cols,expected_records):
    table=ET.fromstring(z.read(tabpath));c1,r1,c2,r2=parse_ref(table.attrib['ref'])
    if r2-r1!=len(expected_records):raise ValueError('Finance table size mismatch')
    sheet=ET.fromstring(z.read(sheetpath))
    strings=shared_strings(z)
    present={}
    for row in sheet.findall('.//'+q('sheetData')+'/'+q('row')):
        n=int(row.attrib['r'])
        if not r1<n<=r2:continue
        for c in row.findall(q('c')):
            col,rr=cell_position(c.attrib['r'])
            if rr==n and c1<=col<=c2:
                present[(n,col)]=string_value(c,strings)
    for i,obj in enumerate(expected_records):
        for k,col in enumerate(expected_cols):
            typ,expected=cell_conversion(col,obj.get(col,''));got=present.get((r1+1+i,c1+k),'')
            if not typ:
                if got!='':raise ValueError('Unexpected finance cell value')
            elif typ=='s':
                if expected!=got:raise ValueError('Finance string roundtrip mismatch')
            else:
                try:
                    actual=Decimal(got)
                    original=Decimal(expected)
                except InvalidOperation as e:raise ValueError('Finance numeric cell not numeric') from e
                if col in DATE_FIELDS:
                    # Sub-second datetime rounding is permitted under 1 millisecond.
                    if abs(actual-original)>Decimal('0.00000001157'):
                        raise ValueError('Excel date roundtrip mismatch')
                elif actual!=original:
                    raise ValueError('Amount roundtrip mismatch')
    return True


def inspect_input(stage,source,finance_raw,workbook,master):
    expected_name='SECONDARY__tblNovac.jsonl'
    origin=locate(stage,expected_name)
    expected_hash=EXPECTED[expected_name][0]
    if sha(origin)!=expected_hash or sha(finance_raw)!=expected_hash:
        raise ValueError('Candidate and original stage fingerprints must match known finance SHA')
    records=load_jsonl(finance_raw)
    if len(records)!=EXPECTED_ROWS:raise ValueError('Expected 39 finance records')
    stats,private=preflight(records,stage)
    if stats['blocked_rows'] or not stats['ready_for_candidate']:
        raise ValueError('Financial records are not ready for import')
    canonical=locate(stage,'tblNovac.jsonl',required=False)
    if canonical:raise ValueError('Stage already contains canonical tblNovac; refuse duplicate import')
    # Important: other (purchase) columns such as SOURCE tblOtkup.Novac must NOT
    # be converted to additional ledger entries. Only this exact source is imported.
    baseline=sha(workbook)
    if baseline!=EXPECTED_MIG11_XLSM:
        raise ValueError('Input MIG-11 workbook fingerprint differs; stop before any edits')
    if sha(master)!=EXPECTED_MASTER_SHA:raise ValueError('Compiled master fingerprint differs')
    precision=precision_summary(records)
    if precision['unrepresentable_numeric_cells_by_field']:
        groups=', '.join(f'{k}={v}' for k,v in precision['unrepresentable_numeric_cells_by_field'].items())
        raise ValueError('Real Excel precision loss would occur in '+groups+'; no workbook written')
    return records,stats,baseline,precision


def run(stage,source,master,finance_raw,outdir):
    for p in (stage,source,master,finance_raw):
        if not p.exists():raise ValueError('Required local input missing: '+p.name)
    if outdir.exists():raise ValueError('Output directory already exists; refuse overwrite')
    if outdir in (stage,source.parent,master.parent,finance_raw.parent):
        raise ValueError('Output must be independent')
    rec,stats,baseline,precision=inspect_input(stage,source,finance_raw,source,master)
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(master) as zm:
        if zin.testzip() or zm.testzip():raise ValueError('ZIP CRC fail')
        if len(zin.namelist())!=len(set(zin.namelist())):raise ValueError('Duplicate OOXML member')
        vba='xl/vbaProject.bin'
        if vba not in zin.namelist() or vba not in zm.namelist():raise ValueError('Missing VBA binary')
        if zin.read(vba)!=zm.read(vba):raise ValueError('VBA differs from compiled master')
        loc,_=table_locations(zin)
        if 'tblNovac' not in loc:raise ValueError('Target XLSM missing tblNovac')
        tabpath,sheetpath,tab=loc['tblNovac']
        if sum(s==sheetpath for _,s,_ in loc.values())!=1:raise ValueError('Finance sheet is shared with another Excel table')
        cols=[x.attrib['name'] for x in tab.find(q('tableColumns')).findall(q('tableColumn'))]
        if cols!=list(FIN_CANON):raise ValueError('XLSM tblNovac schema order differs from expected reference')
        _,r1,_,r2=parse_ref(tab.attrib['ref'])
        if r2 not in (r1,r1+1):raise ValueError('Existing finance data not empty')
        c1,_,c2,_=parse_ref(tab.attrib['ref'])
        newref=f'{col_name(c1)}{r1}:{col_name(c2)}{r1+len(rec)}'
        patches={tabpath:patch_table_xml(zin.read(tabpath),tab.attrib['ref'],newref),
                 sheetpath:patch_sheet(zin.read(sheetpath),tab,cols,rec)}
        outdir.mkdir(parents=True,exist_ok=False)
        part=outdir/'AgriX_MIG13_FINANCE_PATCHED.xlsm.part'
        target=outdir/'AgriX_MIG13_FINANCE_PATCHED.xlsm'
        try:
            with zipfile.ZipFile(part,'w') as zout:
                for zi in zin.infolist():
                    zout.writestr(copy.copy(zi),patches.get(zi.filename,zin.read(zi.filename)))
            with zipfile.ZipFile(part) as z:
                if z.testzip():raise ValueError('Output zip bad')
                if z.read(vba)!=zin.read(vba):raise ValueError('VBA binary differs')
                for filename in zin.namelist():
                    if filename not in patches and zin.read(filename)!=z.read(filename):
                        raise ValueError('Other OOXML part unexpectedly altered')
                read_finance(z,tabpath,sheetpath,cols,rec)
            os.replace(part,target)
            result={
                'format':'AGRIx-MIG-13-FINANCE-OOXML-VERIFY-v1.1',
                'status':'TECHNICAL_DATA_IMPORT_PASS', 'pass':True,
                'workbook':target.name,'table':'tblNovac',
                'expected_rows':len(rec),'actual_rows':len(rec),
                'roundtrip_matches_effective_canonical_values':True,
                'original_numeric_amounts_preserved_exactly':precision['float_tail_normalized']==0,
                'numeric_precision':precision,
                'original_finance_jsonl_byte_unchanged':True,
                'vba_binary_preserved':True,'all_other_xlsm_parts_byte_identical':True,
                'modified_parts':sorted(patches),
                'input_workbook_sha256':baseline,'result_workbook_sha256':sha(target),
                'historical_source_sha256':sha(finance_raw),
                'mig11_relations_preserved':True,
                'customer_rows_in_report':False,
                'source_stage_master_untouched':True,
                'historical_ambalaza_archive_only':True,
                'excel_runtime_qa_pending':True,
                'note':'Imported only 39 existing secondary ledger entries; no ledger synthesized from old purchase summary fields. Functional finance reconciliation and Excel runtime remain pending.'
            }
            (outdir/'mig13_final_verify.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
            print('MIG-13 XLSM FINANCE IMPORT: TECHNICAL PASS')
            print('tblNovac:',len(rec),' /',len(rec),'rows; all cells verified; VBA byte-identical')
            print('Harmless float tails normalized:',precision['float_tail_normalized'],
                  '(raw finance JSONL unchanged; details in final verify JSON)')
            print('SHARE ONLY:',outdir/'mig13_final_verify.json')
        except BaseException:
            if part.exists():part.unlink()
            if target.exists():target.unlink()
            report=outdir/'mig13_final_verify.json'
            if report.exists():report.unlink()
            if outdir.is_dir() and not any(outdir.iterdir()):outdir.rmdir()
            raise


def main(argv=None):
    ap=argparse.ArgumentParser(description='Patch only canonical tblNovac with existing 39 legacy ledger records')
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--input-xlsm',type=Path,required=True)
    ap.add_argument('--master',type=Path,required=True)
    ap.add_argument('--financial-jsonl',type=Path,required=True)
    ap.add_argument('--out-dir',type=Path,required=True)
    args=ap.parse_args(argv)
    try:run(*[p.resolve() for p in (args.stage,args.input_xlsm,args.master,args.financial_jsonl,args.out_dir)])
    except Exception as e:
        # Never emit record IDs, amounts, partner names or personal details to chat.
        print('MIG-13 IMPORT BLOCKED:',type(e).__name__,'—',str(e).split(':')[0],file=sys.stderr)
        return 2
    return 0
if __name__=='__main__':sys.exit(main())
