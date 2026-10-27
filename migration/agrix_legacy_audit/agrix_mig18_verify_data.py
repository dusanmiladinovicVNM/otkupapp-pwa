#!/usr/bin/env python3
"""Read-only semantic OOXML cell comparison across the QA pre/post VBA import.

Does NOT compare VBA: its binary MUST change after imported code is saved.
Flags any sheet cell/formula content drift; ignores styles/caches/worksheet XML layout.
"""
from __future__ import annotations
import argparse
import json
import posixpath
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def ss(z:zipfile.ZipFile) -> list[str]:
    path='xl/sharedStrings.xml'
    if path not in z.namelist():return []
    root=ET.fromstring(z.read(path))
    return [''.join(n.text or '' for n in si.iter(NS+'t')) for si in root.findall(NS+'si')]


def cv(t:str|None,raw:str|None,shared:list[str],inline:ET.Element|None,formula:str|None):
    if formula is not None:return ['formula',formula]
    if t=='s':
        assert raw is not None
        return ['text',shared[int(raw)]]
    if t=='inlineStr':
        return ['text',''.join(n.text or '' for n in inline.iter(NS+'t')) if inline is not None else '']
    if raw is None or raw=='':return None
    if t in ('str','e'):
        return [t,raw]
    if t=='b':return ['bool',raw]
    try:
        n=Decimal(raw)
        return ['num',str(n.normalize())]
    except InvalidOperation:
        return ['other',raw]


def read_cells(path:Path)->tuple[dict[str,list[str]],bool]:
    cells={}
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None:raise ValueError(f'ZIP CRC FAIL: {path}')
        has_vba = 'xl/vbaProject.bin' in z.namelist()
        if not has_vba:raise ValueError('Missing xl/vbaProject.bin')
        shared=ss(z)
        for sh in sorted(p for p in z.namelist() if p.startswith('xl/worksheets/sheet') and p.endswith('.xml') and '/' not in p[len('xl/worksheets/'):]):
            root=ET.fromstring(z.read(sh))
            data=root.find(NS+'sheetData')
            if data is None:continue
            for row in data.findall(NS+'row'):
                for c in row.findall(NS+'c'):
                    loc=c.get('r')
                    if not loc:continue
                    value=c.find(NS+'v')
                    formula=c.find(NS+'f')
                    v=cv(c.get('t'),value.text if value is not None else None,shared,c.find(NS+'is'), formula.text if formula is not None else None)
                    if v is not None:cells[f'{sh}:{loc}']=v
    return cells,has_vba


def run(before:Path,after:Path,out:Path|None):
    if before.resolve()==after.resolve():raise ValueError('Input paths must differ')
    a,va=read_cells(before)
    b,vb=read_cells(after)
    delta=[k for k in a.keys()|b.keys() if a.get(k)!=b.get(k)]
    result={'format':'AGRIx-MIG18-QA-SEMANTIC-CELL-CHECK-v1','status':'PASS' if not delta else 'REVIEW_REQUIRED',
            'cells_before':len(a),'cells_after':len(b),'mismatch_count':len(delta),
            'mismatch_addresses_first_10':sorted(delta)[:10],
            'both_have_vba_binary':va and vb,'data_write_performed':False,
            'note':'All worksheet raw cell/formula values compared; styles and formula cached results excluded. This is not a VBA compile test.'}
    if out is not None:
        if out.exists():raise FileExistsError(f'Report already exists {out}')
        out.parent.mkdir(parents=True,exist_ok=True)
        out.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    print(json.dumps(result,indent=2,ensure_ascii=False))
    if delta:raise SystemExit(2)

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('--before',required=True)
    p.add_argument('--after',required=True)
    p.add_argument('--out')
    a=p.parse_args();run(Path(a.before),Path(a.after),Path(a.out) if a.out else None)
