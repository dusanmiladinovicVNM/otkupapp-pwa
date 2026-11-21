#!/usr/bin/env python3
"""Read-only audit of live VBA exports. Never edits either workbook or VBA files.

Run with --dir <snapshot_dir>. Prints important declarations and variable usage,
then writes ACTUAL_VBA_ANALYSIS.txt to the snapshot directory.
"""
import argparse
import json
from pathlib import Path
import re

PATTERNS = {
 'modAmbalaza': [
   r'(?i)^\s*(?:Public|Private)?\s*Function\s+AmbKretanjaNaloga\b',
   r'(?i)AMB_VK_ULAZ_TUDJE|ULAZ_TUDJE_AMBALAZE',
   r'(?i)COL_AMB_VRSTA_KRETANJA|VrstaKretanja',
   r'(?i)AmbKretanjaNaloga\s*=|res\([^\n]*,\s*[689]\)',
   r'(?i)\bFunction\s+AmbSaldoNaloga\b',
 ],
 'modIzvestaj': [
   r'(?i)^\s*(?:Public|Private)?\s*Function\s+(?:ReportKarticaAmbalaze|ReportKarticaKooperanta)\b',
   r'(?i)AmbKretanjaNaloga\s*\(',
   r'(?i)ULAZ_TUDJE_AMBALAZE|AMB_VK_ULAZ_TUDJE',
   r'(?i)\bKarticaAmbRezultatSaPocetnim\b',
 ],
 'modScrIzvestaji': [
   r'(?i)ReportKarticaAmbalaze|ReportKarticaKooperanta',
   r'(?i)mZonaKartSaldo\b|mZonaKartSaldoAmb\b',
 ],
 'modAmbalazaUgovor': [
   r'(?i)AMB_VK_ULAZ_TUDJE|ULAZ_TUDJE_AMBALAZE',
 ],
}

def nearby(lines, where, a=3, b=6):
    lo = max(0,where-a); hi=min(len(lines),where+b+1)
    return '\n'.join(f'{j+1:5d} | {lines[j]}' for j in range(lo,hi))

def audit(folder):
    out=[]
    manifest = folder/'SOURCE_MANIFEST.json'
    if manifest.exists():
        raw=json.loads(manifest.read_text(encoding='utf-8-sig'))
        out.append(f"SOURCE: {raw.get('workbook')}\nSHA256: {raw.get('workbook_sha256')}\n")
    else:
        out.append('MANIFEST MISSING\n')
    for module, pats in PATTERNS.items():
        path=folder/f'{module}.txt'
        out.append(f'\n====== {module} ======')
        if not path.exists():
            out.append('MISSING MODULE EXPORT');continue
        lines=path.read_text(encoding='utf-8-sig',errors='replace').splitlines()
        out.append(f'LINES: {len(lines)}')
        for pat in pats:
            indices=[i for i,line in enumerate(lines) if re.search(pat,line)]
            out.append(f'PATTERN: {pat} | HITS={len(indices)}')
            for i in indices[:6]: out.append(nearby(lines,i))
        if module=='modAmbalaza':
            signature='\n'.join(lines)
            has_func=bool(re.search(r'(?im)^\s*(Public|Private)?\s*Function\s+AmbKretanjaNaloga\b',signature))
            returns_9=bool(re.search(r'\bres\([^\n]*,\s*9\)',signature,re.I))
            out.append(f'CLASSIFICATION: has_amb_kretanja={has_func} mentions_result_9th_column={returns_9}')
    report=folder/'ACTUAL_VBA_ANALYSIS.txt'
    report.write_text('\n'.join(out)+'\n',encoding='utf-8-sig')
    return report

if __name__=='__main__':
    p=argparse.ArgumentParser()
    p.add_argument('--dir',required=True,type=Path)
    args=p.parse_args()
    report=audit(args.dir)
    print('ANALYSIS_READ_ONLY_OK')
    print('REPORT='+str(report))
