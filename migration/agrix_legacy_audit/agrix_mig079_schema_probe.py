#!/usr/bin/env python3
"""MIG-07.9 synthetic table/Excel persistence probe; NO CUSTOMER DATA.

Runs the SAME excel_write() method on a disposable copy of the exactly pinned
blank XLSM, filling 25 tables with synthetic primary-column labels and empty
other columns. Does not read or require MIG-02, the legacy XLSM, or staging.
No diagnostic ZIP or XLSM is retained. Output JSON includes only public schema
column names, expected/actual row counts, and booleans (never cell contents).
"""
from __future__ import annotations
import argparse
from contextlib import contextmanager
import json
import os
from pathlib import Path
import shutil
import tempfile
import sys
from unittest.mock import patch

import agrix_mig07_goods_copy as mig
from legacy_audit import AuditError, parse_xlsm

# Proven counts from the previous user-visible MIG-07.7 run; the actual test
# contents are SYNTHETIC and do not derive any field from the customer workbook.
COUNTS=(8,103,1,8,10,2,3,3,2,1,3,2,10,1702,1702,907,907,907,907,893,188,1063,93,140,45)
assert len(COUNTS)==len(mig.TABLE_ORDER) and sum(COUNTS)==9610


def synthetic_matrices(schema, names):
    # The first column of every staged table is a textual identifier/type.
    matrices={}; kinds={}
    counts=dict(zip(mig.TABLE_ORDER,COUNTS))
    for table in names:
        columns=schema[table]
        if mig.column_type(columns[0]) != 'text':
            raise AuditError('Synthetic first-column key is not text: '+table)
        kinds[table]=[mig.column_type(x) for x in columns]
        matrices[table]=[tuple([f'SYNTHETIC-MIG079-{table}-{i:06d}']+
                              ['']*(len(columns)-1)) for i in range(1,counts[table]+1)]
    return matrices,kinds


def com_metadata(path, names):
    """Read Excel ListColumns and visible table headers; never touch values."""
    import pythoncom
    import win32com.client as w32
    pythoncom.CoInitialize()
    xl=None; wb=None
    try:
        xl=w32.DispatchEx('Excel.Application')
        xl.Visible=False; xl.DisplayAlerts=False; xl.EnableEvents=False
        xl.AskToUpdateLinks=False; xl.AutomationSecurity=3
        wb=xl.Workbooks.Open(str(path),UpdateLinks=0,ReadOnly=True,
                             IgnoreReadOnlyRecommended=True,AddToMru=False)
        all_objs={str(o.Name):o for ws in wb.Worksheets for o in ws.ListObjects}
        snapshot={}
        for name in names:
            lo=all_objs.get(name)
            if lo is None:
                snapshot[name]={'missing':True}; continue
            headers=[str(lo.ListColumns.Item(i).Name) for i in range(1,lo.ListColumns.Count+1)]
            h=lo.HeaderRowRange.Value2
            if isinstance(h,tuple):
                if len(h)==1 and isinstance(h[0],tuple): h=h[0]
                h=[str(x or '') for x in h]
            else:h=[str(h or '')]
            snapshot[name]={'ListColumns':headers,'HeaderRowRange':h,
                           'ListRows':int(lo.ListRows.Count),
                           'RangeRows':int(lo.Range.Rows.Count)}
        return snapshot
    finally:
        if wb is not None:
            try:wb.Close(SaveChanges=False)
            except Exception: pass
        if xl is not None:
            try:xl.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()


def column_report(names, schema, parsed, com=None):
    report={}
    for name in names:
        table=parsed.get(name.casefold())
        if table is None:
            report[name]={'status':'MISSING_TABLE'};continue
        expected=schema[name]; actual=table.columns
        item={'status':'MATCH' if expected==actual else 'MISMATCH',
              'expected_columns':len(expected),'saved_columns':len(actual),
              'saved_meaningful_rows':len(table.rows)}
        if expected!=actual:
            item['column_difference']=mig.schema_column_difference(expected,actual)
        if com and name in com:
            values=com[name]
            item['excel_list_columns_match']=values.get('ListColumns')==expected
            item['excel_header_cells_match']=values.get('HeaderRowRange')==expected
            if values.get('ListColumns')!=expected or values.get('HeaderRowRange')!=expected:
                # Public schema identifiers ONLY; no business values.
                item['excel_list_columns']=values.get('ListColumns')
                item['excel_header_cells']=values.get('HeaderRowRange')
            item['excel_ListRows']=values.get('ListRows')
            item['excel_RangeRows']=values.get('RangeRows')
        report[name]=item
    return report


def main(argv=None):
    p=argparse.ArgumentParser(description='Reproduce Excel table schema drift with synthetic data only')
    p.add_argument('--template',required=True,type=Path)
    p.add_argument('--schema',required=True,type=Path)
    p.add_argument('--out',required=True,type=Path,help='new output directory, only redacted JSON/TXT')
    p.add_argument('--first-three',action='store_true',help='test 8 + 103 + 1 rows instead of 9610')
    args=p.parse_args(argv)
    master=args.template.resolve();schema_file=args.schema.resolve();out=args.out.resolve()
    if out.exists():raise AuditError('Output already exists; choose new directory')
    if not master.is_file() or not schema_file.is_file():raise AuditError('Missing template/schema')
    if out in master.parents or out in schema_file.parents or out==master or out==schema_file:
        raise AuditError('Unsafe output location')
    if os.name!='nt':raise AuditError('Windows Excel is required to reproduce Windows COM behavior')
    if mig.sha256(master)!=mig.TEMPLATE_SHA:
        raise AuditError('Blank master SHA is different from preflight; do not run')
    schema=mig.get_schema(schema_file)
    names=mig.TABLE_ORDER[:3] if args.first_three else mig.TABLE_ORDER
    matrices,kinds=synthetic_matrices(schema,names)
    out.mkdir(parents=True)
    result={'format':'MIG-07.9-SYNTHETIC-SCHEMA-PROBE','status':'STARTED',
            'no_customer_data_read':True,'template_unchanged_expected':mig.TEMPLATE_SHA,
            'tables_tested':len(names),'synthetic_rows':sum(len(x) for x in matrices.values())}
    try:
        with tempfile.TemporaryDirectory(prefix='AGRIX_SYNTHETIC_SCHEMA_PROBE_') as d:
            copy=Path(d)/'SYNTHETIC_ONLY.xlsm'
            shutil.copy2(master,copy)
            before=parse_xlsm(copy)
            result['before']=column_report(names,schema,before)
            try:
                with patch.object(mig,'TABLE_ORDER',tuple(names)):
                    mig.excel_write(copy,matrices,schema,kinds)
            except Exception as exc:
                result['status']='EXCEL_WRITE_FAILED'
                result['write_error_type']=type(exc).__name__
                result['write_error_message']=str(exc)[:800]
            else:
                after=parse_xlsm(copy)
                result['after']=column_report(names,schema,after)
                try:
                    c=com_metadata(copy,names)
                    result['after']=column_report(names,schema,after,c)
                except Exception as e:
                    result['com_inspection_error_type']=type(e).__name__
                    result['com_inspection_error']=str(e)[:300]
                result['status']='SYNTHETIC_SCHEMA_MATCH' if all(
                    v['status']=='MATCH' and v['saved_meaningful_rows']==len(matrices[n])
                    and v.get('excel_list_columns_match',True)
                    and v.get('excel_header_cells_match',True)
                    for n,v in result['after'].items()) else 'SYNTHETIC_SCHEMA_DRIFT'
            result['template_unchanged']=mig.sha256(master)==mig.TEMPLATE_SHA
        (out/'agrix_schema_probe_summary.json').write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_schema_probe_summary.txt').write_text(
            'AgriX MIG-07.9 synthetic workbook schema probe\n'+
            'Status: '+result['status']+'\n'+
            'Synthetic rows: '+str(result['synthetic_rows'])+'\n'+
            'Source/customer data: NOT READ\n'+
            'Blank master hash unchanged: '+str(result['template_unchanged'])+'\n'+
            'Mismatches: '+', '.join(n for n,v in result.get('after',{}).items() if v['status']!='MATCH')+'\n',
            encoding='utf-8')
        print('MIG-07.9 SCHEMA PROBE:',result['status'])
        print('No customer data was accessed. Share only agrix_schema_probe_summary.json and .txt')
        return 0 if result['status']=='SYNTHETIC_SCHEMA_MATCH' else 2
    except BaseException:
        # no synthetic workbook survives temp context, no partial public results
        raise

if __name__=='__main__':
    try:raise SystemExit(main())
    except (AuditError,OSError,ValueError,KeyError,TypeError) as e:
        print('MIG-07.9 PROBE STOP:',e,file=sys.stderr)
        raise SystemExit(2)
