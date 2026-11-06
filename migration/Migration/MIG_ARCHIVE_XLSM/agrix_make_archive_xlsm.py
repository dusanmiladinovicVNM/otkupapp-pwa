#!/usr/bin/env python3
"""AgriX 2026: copy migrated XLSM + add read-only archival worksheets.

Windows Excel 2021 COM performs the XLSM modifications. The source file
is only read with zip/xml and copied to a new file. Original VBA modules
and historical business tables remain unchanged.
"""
from __future__ import annotations
import argparse
import csv
import hashlib
import json
import os
import shutil
import subprocess
import sys
import zipfile
from collections import defaultdict, Counter
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from agrix_mig_archive import read_workbook, export_otkups, extract_issues
from agrix_mig19_malina_readonly_audit import sha256, norm

ARCHIVE_VERSION = 'AGRIx-MIG-ARCHIVE-XLSM-v1'
DATA_COLUMNS = ('tblOtkup', 'tblOtkupStavke', 'tblOtpremnica',
                'tblOtpremnicaIzvori', 'tblOtpremnicaStavke', 'tblZbirna',
                'tblZbirnaIzvori', 'tblZbirnaStavke', 'tblPrijemnica',
                'tblAmbalaza', 'tblNovac')

def safe_csv_value(v):
    value = '' if v is None else str(v)
    # Prevent formula execution from free-text in imported historical data.
    try:
        Decimal(value)  # signed decimal data is safe numeric input
        return value
    except Exception:
        return "'" + value if value.startswith(('=', '+', '-', '@', '\t', '\r')) else value

def write_csv(path: Path, header: list[str], rows: list[list]):
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        w.writerow(header)
        for line in rows:
            w.writerow([safe_csv_value(v) for v in line])

def as_dec(v):
    if v is None or v == '':
        return None
    return Decimal(str(v))

def as_text(v):
    return '' if v is None else str(v)

def build_csvs(output: Path, records: list[dict], issues: list[dict], tables: dict):
    files = {}
    def emit(name, headers, body):
        name = name + '.csv'
        write_csv(output/name, headers, body)
        files[name] = len(body)

    active = [r for r in records if not r['storno']]
    summary = defaultdict(list)
    coop = defaultdict(list)
    for r in active:
        summary[(r['station_id'], r['station'])].append(r)
        coop[(r['coop_id'], r['coop'], r['station_id'], r['station'])].append(r)

    def total(rows, field):
        return str(sum((as_dec(r[field]) for r in rows if r[field] is not None), Decimal(0)))
    def incomplete(rows):
        return sum(r['kg'] is None or r['value'] is None for r in rows)
    om = [[sid, name, len(rows), total(rows,'kg'), total(rows,'value'), incomplete(rows)]
          for (sid,name), rows in sorted(summary.items())]
    emit('ARHIVA_OM', ['OM ID','Naziv OM','Aktivnih otkupa','Kilograma (poznato)',
           'Vrednost RSD (poznato)','Nepotpunih otkupa'], om)
    koops = [[kid, name, sid, station, len(rows), total(rows,'kg'),
              total(rows,'value'), incomplete(rows)] for (kid,name,sid,station),rows in sorted(coop.items())]
    emit('ARHIVA_KOOPERANTI', ['KooperantID','Ime i prezime','OM ID','OM',
        'Aktivnih otkupa','Kilograma (poznato)','Vrednost RSD (poznato)',
        'Nepotpunih otkupa'], koops)
    otk = [[r['id'],r['excel_row'],r['date'],r['business'],r['station_id'],r['station'],
           r['coop_id'],r['coop'],r['fruit'],r['variety'],r['classes'],
           as_text(r['kg']),as_text(r['value']), 'DA' if r['storno'] else 'NE', ', '.join(r['flags'])]
           for r in sorted(records,key=lambda x:(x['date'],x['id']))]
    emit('ARHIVA_OTKUPI', ['OtkupID','Excel red','Datum','Broj dokumenta',
        'OM ID','OM','KooperantID','Kooperant','Vrsta voća','Sorta',
        'Klasa','Kg','Vrednost RSD','Stornirano','Nepotpuno'],otk)
    ix = [[x.get('priority'),x.get('category'),x.get('code'),x.get('table'),
            x.get('excel_row'),x.get('document_id'),x.get('related_id'),
            x.get('date'),x.get('problem'),x.get('impact'),x.get('blocks'),
            'ZA PREGLED'] for x in issues]
    emit('ARHIVA_IZUZECI',['Prioritet','Grupa','Šifra problema','Tabela',
            'Excel red','DokumentID','Povezani ID','Datum','Opis',
            'Uticaj na pregled','Blokira izveštaj','Pregled'],ix)
    by_table_row = {(name, str(r.get('_excel_row'))):r
                    for name, rr in tables.items() for r in rr}
    raw=[]
    for x in issues:
        row=by_table_row.get((x['table'],str(x['excel_row'])), {})
        source_number=next((norm(row.get(y)) for y in
            ('BrojDokumenta','BrojOtpremnice','BrojZbirne','BrojPrijemnice')
            if norm(row.get(y))), '')
        raw.append([x['document_id'],x['code'],x['table'],x['excel_row'],
                    source_number, norm(row.get('Datum')),norm(row.get('StanicaID')),
                    norm(row.get('KooperantID')),norm(row.get('VozacID')),
                    norm(row.get('Stornirano')),norm(row.get('IzdatoStatus')),
                    json.dumps(row,ensure_ascii=False,default=str)[:30000] if row else ''])
    emit('ARHIVA_SIROVI',['DokumentID','Problem','Izvorna tabela','Excel red',
            'Broj','Datum','StanicaID','KooperantID','VozacID',
            'Stornirano','IzdatoStatus','Originalni red (JSON)'],raw)
    return files, len(active), sum(r['kg'] is None or r['value'] is None for r in active)

def project_bytes(path: Path):
    with zipfile.ZipFile(path) as z:
        matches = [n for n in z.namelist() if n.lower().endswith('vbaproject.bin')]
        if len(matches)!=1:
            raise RuntimeError(f'Izvorni XLSM nema jedinstveni VBA projekat: {matches}')
        return hashlib.sha256(z.read(matches[0])).hexdigest()

def workbook_signature(tables):
    # Check all imported business table values, including worksheet row identity.
    # This is independent of archive worksheets and formatting.
    digest=hashlib.sha256()
    for name in sorted(tables):
        if not name.startswith('tbl'):
            continue
        digest.update(name.encode())
        for row in tables[name]:
            digest.update(json.dumps(row,ensure_ascii=False,sort_keys=True,
                 default=str,separators=(',',':')).encode('utf-8'))
    return digest.hexdigest()

def find_powershell():
    for c in ('powershell.exe','powershell'):
        p=shutil.which(c)
        if p: return p
    raise RuntimeError('PowerShell nije pronadjen; pokretanje je namenjeno za Windows sa Excel 2021.')

def check_output(source: Path, candidate: Path, before_tables, vba_digest: str):
    with zipfile.ZipFile(candidate) as z:
        invalid = z.testzip()
        if invalid: raise RuntimeError('Izlazni XLSM ima neispravan ZIP entry '+invalid)
        required = [s for s in z.namelist() if s.lower().endswith('vbaproject.bin')]
        if len(required)!=1: raise RuntimeError('Novi XLSM je izgubio VBA projekat')
    # Excel may refresh compiled VBA caches when saving a macro-enabled book.
    # Preserve code by never editing VBProject; report byte differences explicitly.
    vba_same=(project_bytes(candidate)==vba_digest)
    if not vba_same:
        print('UPOZORENJE: Excel je ponovo snimio VBA binarni projekat; '+
              'makroi su sacuvani, ali byte-for-byte hash nije isti. '+
              'Pre aktiviranja originalnog UI proveri Compile.',flush=True)
    out_tables=read_workbook(candidate)
    if workbook_signature(before_tables)!=workbook_signature(out_tables):
        raise RuntimeError('Postojeći poslovni redovi se razlikuju od izvora; STOP!')
    # All worksheet references must have been saved; Archive sheets are not ListObjects.
    with zipfile.ZipFile(candidate) as z:
        wb_xml=z.read('xl/workbook.xml').decode('utf-8')
        if 'ARHIVA_OM' not in wb_xml or 'ARHIVA_IZUZECI' not in wb_xml:
            raise RuntimeError('Nisu dodati obavezni arhivski listovi')

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--workbook',required=True,type=Path,help='originalni postojeći XLSM')
    ap.add_argument('--out-dir',required=True,type=Path,help='NOV i prazan izlazni folder')
    ap.add_argument('--prepare-only',action='store_true',help='bez COM-a, samo CSV/manifest preflight')
    args=ap.parse_args()
    src=args.workbook.resolve(); out=args.out_dir.resolve()
    if not src.is_file() or src.suffix.lower()!='.xlsm':ap.error('Nije nadjen XLSM: '+str(src))
    if out==src.parent or src in out.parents:ap.error('Izlazni folder mora biti nezavisan od ulazne sveske')
    if out.exists() and any(out.iterdir()):ap.error('Izlazni folder vec ima fajlove; izaberi nov/prázan folder.')
    out.mkdir(parents=True,exist_ok=True)
    sha_orig=sha256(src)
    vba_orig=project_bytes(src)
    print('MIG-ARCHIVE: procitavam originalni XLSM bez upisa...',flush=True)
    T=read_workbook(src)
    records,extra=export_otkups(T)
    issues=extract_issues(T,extra)
    code_counts=Counter(x['code'] for x in issues)
    duplicates=sum(1 for x in issues if x['code']=='DUPLICATE_PRIMARY_ID')
    files,active,partial=build_csvs(out,records,issues,T)
    meta={'version':ARCHIVE_VERSION,'source':str(src),'source_sha256':sha_orig,
          'vba_sha256':vba_orig,'source_table_count':len(T),
          'total_otkup_records':len(records),'active_otkups':active,
          'active_otkups_with_missing_values':partial,'exceptions':len(issues),
          'duplicate_document_primary_id_findings':duplicates,
          'codes':dict(code_counts),'sheet_rows':files,
          'source_unchanged':None,'workbook_generated':not args.prepare_only,
          'generated':datetime.now().astimezone().isoformat(timespec='seconds')}
    meta_path=out/'ARHIVA_2026_MAKE_SUMMARY.json'
    if not args.prepare_only:
        dst=out/'AgriX_2026_ARHIVA.xlsm'
        candidate=out/'__ARHIVA_BUILD__.xlsm'
        if dst.exists() or candidate.exists():raise RuntimeError('Izlazni fajl vec postoji')
        shutil.copy2(src,candidate)
        powershell=find_powershell()
        ps=Path(__file__).with_name('agrix_archive_excel_com.ps1')
        print('Kreiram arhivske Excel listove u NOVOJ kopiji (Excel COM)...',flush=True)
        res=subprocess.run([powershell,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(ps),
             '-Workbook',str(candidate),'-CsvDir',str(out)],text=True,
              stdout=subprocess.PIPE,stderr=subprocess.STDOUT,errors='replace')
        print(res.stdout,flush=True)
        if res.returncode!=0:
            candidate.unlink(missing_ok=True)
            raise RuntimeError(f'Excel COM je zaustavljen, original nije menjan (exit {res.returncode}).')
        print('Proveravam nepromenjen VBA i sve poslovne tabele...',flush=True)
        try:
            check_output(src,candidate,T,vba_orig)
            meta['vba_project_byte_identical'] = (project_bytes(candidate)==vba_orig)
            if sha256(src)!=sha_orig: raise RuntimeError('IZVORNI XLSM JE PROMENJEN -- STOP')
            candidate.rename(dst)
            meta['output_xlsm']=str(dst)
            meta['output_sha256']=sha256(dst)
        except Exception:
            candidate.unlink(missing_ok=True)
            raise
    if sha256(src)!=sha_orig:raise RuntimeError('Izvorni XLSM je promenjen! STOP.')
    meta['source_unchanged']=True
    meta_path.write_text(json.dumps(meta,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'status':'PREPARED_ONLY' if args.prepare_only else 'ARCHIVE_XLSM_CREATED',
                      **meta},ensure_ascii=False,indent=2))
    return 0
if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:
        print('MIG-ARCHIVE XLSM STOP:',e,file=sys.stderr)
        sys.exit(2)
