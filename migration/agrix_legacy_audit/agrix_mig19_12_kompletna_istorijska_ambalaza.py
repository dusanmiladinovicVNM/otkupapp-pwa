#!/usr/bin/env python3
"""MIG19.12: assemble a COMPLETE historical packaging evidence workbook from local inputs.

Source: 7,815 frozen original legacy rows, 4,811 frozen AMB10 QA rows.
The six OMUlaz source rows are represented as three internal transfers and three
explicit UNKNOWN-SOURCE historical receipts, WITHOUT mislabelling procurement.
AMB-00700 is excluded from active historical saldo only AFTER verifying the
parent Otpremnica is storno in the original staged parent table.

This makes a verified HISTORICAL .xlsx archive, NOT a runnable AMB10 .xlsm: the
three unknown-source entries have no permitted AMB10 movement kind as of 66f7897f.
The output explicitly declares its runtime-compatibility HOLD rather than inventing
NABAVKA or a real counterparty. Input XLSM is NEVER edited.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED, is_zipfile
from xml.sax.saxutils import escape, quoteattr

from agrix_mig19_3_parity import (
    XlsmReadOnly, sha256, stage_rows, ledger_views, reconcile,
    key, norm_type, number, date, void, dec, STAGE_SHA,
)
from agrix_mig19_11_istorijska_semantika import (
    HISTORIC_DESCRIPTIONS, EXPECTED_TARGET_SHA, OMULAZ, READ_COLUMNS,
    semantic_target,
)

PARENT_SHA = 'd02436aa32a0e14ae6ca19172227f83b358b3ab0e329a7f040f97d29748e611a'
PARENT_FILE = 'SOURCE__tblOtpremnica.jsonl'
EXCLUDED = 'AMB-00700'
SOURCE_COUNT = 7815
CANON_COUNT = 4811
MIG_VERSION = 'MIG19.12'
COLUMNS = [
    'AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP',
    'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja',
    'StornoOd','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy'
]


def stage_jsonl(path):
    out=[]
    with Path(path).open(encoding='utf-8-sig') as f:
        for n,line in enumerate(f,1):
            if not line.strip():continue
            x=json.loads(line)
            if not isinstance(x,dict): raise ValueError(f'Bad JSONL #{n}: {path}')
            out.append(x)
    return out


def parent_storno_proof(old, parent):
    hits=[r for r in old if key(r.get('AmbID'))==EXCLUDED]
    if len(hits)!=1: raise ValueError('AMB-00700 not unique in frozen original')
    a=hits[0]
    if any((key(a.get('DokumentTIP'))!='OTPREMNICA',key(a.get('EntitetTip'))!='STANICA',
            key(a.get('EntitetID'))!='ST-00004',key(a.get('Smer'))!='IZLAZ',
            norm_type(a.get('TipAmbalaze'))!='0,4KG',
            date(a.get('Datum'),'AMB-00700')!='2026-06-22',
            number(a.get('Kolicina'),'AMB-00700')!=Decimal(500))):
        raise ValueError('AMB-00700 original fields do not match recorded exception')
    docid=key(a.get('DokumentID'))
    p=[r for r in parent if key(r.get('OtpremnicaID'))==docid]
    if len(p)!=1:
        raise ValueError(f'Parent OtpremnicaID {docid!r} not unique/found: {len(p)}')
    if key(p[0].get('Stornirano'))!='DA':
        raise ValueError(f'Parent OtpremnicaID {docid} is not marked Stornirano=Da')
    if key(p[0].get('StanicaID'))!='ST-00004':
        raise ValueError('Parent station is not ST-00004')
    if number(p[0].get('KolAmbalaze'),'parent 500')!=Decimal(500):
        raise ValueError('Parent packaging amount is not 500')
    return dict(LegacyAmbID=EXCLUDED,OtpremnicaID=docid,
                BrojOtpremnice=str(p[0].get('BrojOtpremnice','')),
                Datum=date(a.get('Datum'),'parent'),
                Kolicina=500,StornoNaAmbRedu='NE',StornoNaRoditelju='DA',
                Odluka='IZUZET_IZ_AKTIVNOG_OBRACUNA_SAMO_U_DERIVATU')


def build_six(old,new):
    byid={}
    for r in old:
        aid=key(r.get('AmbID'))
        if aid in OMULAZ:
            if aid in byid: raise ValueError(f'Duplicate {aid}')
            byid[aid]=r
    if len(byid)!=6:raise ValueError(f'Expected six OMUlaz records, got {len(byid)}')
    existing={key(r.get('AmbID')) for r in new}
    added=[];evidence=[]
    for aid,(day,dest,qty,origin) in OMULAZ.items():
        r=byid[aid]
        if (void(r.get('Stornirano')) or key(r.get('DokumentTIP'))!='OMULAZ'
            or key(r.get('EntitetTip'))!='STANICA'
            or key(r.get('EntitetID'))!=dest or key(r.get('Smer'))!='ULAZ'
            or number(r.get('Kolicina'),aid,True)!=qty
            or date(r.get('Datum'),aid)!=day or norm_type(r.get('TipAmbalaze'))!='0,4KG'):
            raise ValueError(f'OMUlaz source mismatch: {aid}')
        # Prefix ID is deterministic, distinct and transparent about source lineage.
        nid='AMB-MIG19-'+aid.removeprefix('AMB-')
        if nid in existing:raise ValueError(f'Canonical identity collision: {nid}')
        existing.add(nid)
        isinternal=origin is not None
        c=dict(AmbID=nid, Datum=day,TipAmbalaze=str(r.get('TipAmbalaze')),
               Kolicina=dec(qty),DokumentID=str(r.get('DokumentID') or aid),
               DokumentTIP='OMUlaz',OdNalogTip='STANICA' if isinternal else 'NEPOZNAT',
               OdNalogID=origin or '',NaNalogTip='STANICA',NaNalogID=dest,
               VrstaKretanja='PRENOS_INTERNO' if isinternal else 'ISTORIJSKI_ULAZ_IZVOR_NEPOZNAT',
               StornoOd='',CreatedAt=str(r.get('CreatedAt','')),CreatedBy=str(r.get('CreatedBy','')),
               ModifiedAt=str(r.get('ModifiedAt','')),ModifiedBy=str(r.get('ModifiedBy','')))
        added.append(c)
        evidence.append(dict(IzvorniAmbID=aid,NoviAmbID=nid,Datum=day,Kolicina=dec(qty),
                  TipAmbalaze=norm_type(r.get('TipAmbalaze')),
                  OdNalogTip=c['OdNalogTip'],OdNalogID=origin or 'NEPOZNAT',
                  NaNalogTip='STANICA',NaNalogID=dest,
                  Oznaka='AMB10_VALIDAN_PRENOS' if isinternal else 'ARHIVA_SAMO_AMB10_HOLD',
                  IzvorniDokumentTip=str(r.get('DokumentTIP')),IzvorniDokumentID=str(r.get('DokumentID') or '')))
    return added,evidence


def historical_old_projection(original, parentproof):
    # Every historical source row is retained in the archive. Only the derived
    # ACTIVE calculation excludes the row whose parent is confirmed cancelled.
    derived=[dict(r,Stornirano='Da') if key(r.get('AmbID'))==EXCLUDED else dict(r)
             for r in original]
    # User-confirmed source of 3 physical transfers is ST-00001. The original
    # single-leg receipts were incomplete: historical reconstructed ST1 also
    # receives the missing sender legs. These derived legs are NOT original rows.
    for aid,(day,dest,qty,src) in OMULAZ.items():
        if src:
            r=next(x for x in original if key(x.get('AmbID'))==aid)
            derived.append(dict(AmbID='DERIVED-COUNTERLEG-'+aid,Datum=day,
              TipAmbalaze=r.get('TipAmbalaze'),Kolicina=dec(qty),Smer='Izlaz',
              EntitetTip='Stanica',EntitetID=src,VozacID='',
              DokumentTIP='OMUlaz',DokumentID=r.get('DokumentID') or aid,
              Stornirano='Ne'))
    return derived


def matrix_summary(original,new,added,parentproof):
    transformed=historical_old_projection(original,parentproof)
    new_complete=[*new,*added]
    olddaily,physical,newdaily,cover,od,nd,diags,after=ledger_views(transformed,new_complete)
    balances,dates,docs=reconcile(olddaily,physical,newdaily,cover,od,nd)
    bad=[x for x in balances if x['Status']!='PASS']
    bad_dates=[x for x in dates if x['Status']!='PASS']
    source_nolog=[r for r in balances if r['NalogTip'] not in ('STANICA','KOOPERANT')]
    # The double-sided reconstructed semantics must produce exact balances
    # for EVERY cooperant and station, not just the four known accounts.
    if bad or bad_dates:
        raise ValueError(f'ACCEPTANCE FAIL: mismatched accounts={len(bad)} dated={len(bad_dates)}; '
             + str(bad[:8]))
    vals={(x['NalogTip'],x['NalogID'],x['TipAmbalaze']):x for x in balances}
    expected={('STANICA','ST-00001','0,4KG'):Decimal(-16094),
              ('STANICA','ST-00002','0,4KG'):Decimal(-3577),
              ('STANICA','ST-00004','0,4KG'):Decimal(-4173),
              ('STANICA','ST-00005','0,4KG'):Decimal(-8093)}
    for acc,amount in expected.items():
        if acc not in vals or Decimal(vals[acc]['NoviIstorijskiSaldo'])!=amount:
            raise ValueError(f'Expected corrected station saldo {acc}={amount}, '
                            f'found {vals.get(acc)}')
    return balances,dates,dict(account_pairs=len(balances),account_mismatches=0,
         dated_closes=len(dates),dated_mismatches=0,
         added_internal_transfers=3,added_unknown_receipts=3,
         canonical_original_rows=len(new),canonical_archive_rows=len(new_complete),
         original_legacy_rows=len(original),source_rows_active_by_flag=sum(not void(x.get('Stornirano')) for x in original),
         historical_new_kind_unknown_source_rows=3,
         stop_for_AMB10_runtime=True)


def xlcol(n):
    s=''
    while n:
        n,a=divmod(n-1,26);s=chr(65+a)+s
    return s


def safe_text(value):
    s=str(value if value is not None else '')
    # XLSX inline strings are never evaluated as formulas.
    return ''.join(ch for ch in s if ch=='\t' or ch=='\n' or ord(ch)>=32)


def worksheet_xml(headers,rows,number_cols=(),filter_=True):
    # Generate streaming OOXML to avoid retaining huge ElementTrees (16k+ rows).
    cols=len(headers);lastrow=len(rows)+1
    yield '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    yield '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    yield f'<dimension ref="A1:{xlcol(cols)}{max(1,lastrow)}"/>'
    yield '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    yield '<sheetFormatPr defaultRowHeight="16"/><cols>'
    for j,h in enumerate(headers,1):
        wid=15 if h in ('Datum','Kolicina') else min(34,max(12,len(h)+2))
        yield f'<col min="{j}" max="{j}" width="{wid}" customWidth="1"/>'
    yield '</cols><sheetData>'
    def cell(j,rn,val,heading=False):
        ref=f'{xlcol(j)}{rn}';v=safe_text(val)
        if not heading and j in number_cols and v and re.fullmatch(r'-?(?:\d+\.?\d*|\.\d+)',v):
            return f'<c r="{ref}" s="2"><v>{v}</v></c>'
        return (f'<c r="{ref}" s="{1 if heading else 0}" t="inlineStr">'
                f'<is><t>{escape(v)}</t></is></c>')
    yield '<row r="1" ht="25" customHeight="1">'
    for j,h in enumerate(headers,1):yield cell(j,1,h,True)
    yield '</row>'
    for rn,row in enumerate(rows,2):
        yield f'<row r="{rn}">'
        for j,x in enumerate(row,1):
            if x not in (None,''):yield cell(j,rn,x)
        yield '</row>'
    yield '</sheetData>'
    if filter_: yield f'<autoFilter ref="A1:{xlcol(cols)}{max(1,lastrow)}"/>'
    yield '</worksheet>'


def build_archive(filename,sheets):
    wbns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    relns='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with ZipFile(filename,'w',compression=ZIP_DEFLATED,compresslevel=6) as z:
        x=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           f'<workbook xmlns="{wbns}" xmlns:r="{relns}"><sheets>']
        for i,(name,_,_,_) in enumerate(sheets,1):
            x.append(f'<sheet name={quoteattr(name)} sheetId="{i}" r:id="rId{i}"/>')
        x.append('</sheets></workbook>')
        z.writestr('xl/workbook.xml',''.join(x))
        z.writestr('xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?>'+
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
           ''.join(f'<Relationship Id="rId{i}" Type="{relns}/worksheet" Target="worksheets/sheet{i}.xml"/>'
                   for i in range(1,len(sheets)+1))+'</Relationships>')
        z.writestr('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?>'+
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
            '<Default Extension="xml" ContentType="application/xml"/>'+
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+
            ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    for i in range(1,len(sheets)+1))+'</Types>')
        z.writestr('_rels/.rels','<?xml version="1.0" encoding="UTF-8"?>'+
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
           f'<Relationship Id="rId1" Type="{relns}/officeDocument" Target="xl/workbook.xml"/>'+
           '</Relationships>')
        z.writestr('xl/styles.xml','<?xml version="1.0" encoding="UTF-8"?>'+
           '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
           '<fonts count="2"><font><sz val="10"/><name val="Aptos"/></font>'+
           '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font></fonts>'+
           '<fills count="3"><fill><patternFill patternType="none"/></fill>'+
           '<fill><patternFill patternType="gray125"/></fill>'+
           '<fill><patternFill patternType="solid"><fgColor rgb="FF12324A"/><bgColor indexed="64"/></patternFill></fill></fills>'+
           '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'+
           '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'+
           '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'+
           '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1"/>'+
           '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>')
        for i,(_,headers,rows,nums) in enumerate(sheets,1):
            # Passing an iterator makes writing memory-safe even for large ledgers.
            with z.open(f'xl/worksheets/sheet{i}.xml','w') as o:
                for part in worksheet_xml(headers,rows,nums):o.write(part.encode('utf-8'))
    with ZipFile(filename) as z:
        if z.testzip():raise ValueError('Output XLSX archive CRC invalid')
        if not {'xl/workbook.xml','xl/worksheets/sheet1.xml','xl/styles.xml'}.issubset(z.namelist()):
            raise ValueError('Output XLSX is incomplete')


def generate(stage_dir, qa, outdir):
    src=stage_dir/'UNCONVERTED__tblAmbalaza.jsonl'
    par=stage_dir/PARENT_FILE
    for p in (src,par,qa):
        if not p.is_file(): raise FileNotFoundError(f'Missing local file: {p}')
    orig_sha=(sha256(src),sha256(par),sha256(qa))
    if orig_sha!=(STAGE_SHA,PARENT_SHA,EXPECTED_TARGET_SHA):
        raise ValueError(f'Checksum mismatch. Got original/parent/QA {orig_sha}; refusing stale/unverified data')
    old=stage_rows(src); parents=stage_jsonl(par)
    if len(old)!=SOURCE_COUNT:raise ValueError(f'Unexpected legacy count {len(old)}')
    parentproof=parent_storno_proof(old,parents)
    z=XlsmReadOnly(qa)
    try:can=z.read_table('tblAmbalaza',set(COLUMNS))
    finally:z.close()
    if len(can)!=CANON_COUNT:raise ValueError(f'Unexpected canonical count {len(can)}')
    added,linked=build_six(old,can)
    totals,dailies,checks=matrix_summary(old,can,added,parentproof)
    # Confirm that the 1,179 reverse movements are semantic relabels only,
    # not write operations; these remain unchanged physically in archive.
    _,retagged,ungrounded=semantic_target(old,can)
    if ungrounded: raise ValueError('Historic issuance semantic matching not exhaustive')
    checks['source_grounded_issuance_labels']=len(retagged)
    checks['storno_parent_verified']=True
    checks['canonical_AMB10_compatible']=False
    checks['ready_for_runtime_import']=False
    checks['status']='HISTORICAL_ARCHIVE_PASS__AMB10_UNKNOWN_SOURCE_HOLD'
    checks['baseline_staging_sha256']=orig_sha[0]
    checks['baseline_parent_sha256']=orig_sha[1]
    checks['baseline_qa_sha256']=orig_sha[2]
    checks['parent_storno_detail']=parentproof
    checks['important_note']='Three UNKNOWN source receipts are preserved as historical archive rows with noncanonical movement kind; do not import into active AMB10 before explicit domain mapping.'

    oldkeys=['AmbID','Datum','TipAmbalaze','Kolicina','Smer','EntitetTip','EntitetID',
             'VozacID','DokumentTIP','DokumentID','Stornirano','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy']
    main_headers=[*COLUMNS,'IstorijskiIzvorniAmbID','StatusMigracije']
    main_rows=[]
    for r in can:
        main_rows.append([r.get(c,'') for c in COLUMNS]+['','POSTOJECI_AMB10_QA'])
    for r,e in zip(added,linked):
        main_rows.append([r.get(c,'') for c in COLUMNS]+[e['IzvorniAmbID'],e['Oznaka']])
    overview=[
      ['Rezultat','HISTORIJSKA_SALDA_PASS / AMB10_RUNTIME_HOLD'],
      ['Istorijski cutoff','2026-08-31'],['Broj izvornih redova',len(old)],
      ['Broj postojećih kanonskih redova',len(can)],['Dodati istorijski događaji',len(added)],
      ['Ukupno u SVA_AMBALAZA',len(main_rows)],
      ['Tri prenosa iz ST-00001','633 komada, pune dve strane'],
      ['Tri istorijska prijema na ST-00001','18 komada, NEPOZNAT izvor, ne NABAVKA'],
      ['AMB-00700','Izbačen iz aktivnog istorijskog obračuna jer je OTP roditelj storno'],
      ['Usklađeni parovi nalog/tip',f"{checks['account_pairs']} / {checks['account_pairs']}"],
      ['Usklađena dnevna zatvaranja',f"{checks['dated_closes']} / {checks['dated_closes']}"],
      ['Storno izvor',parentproof['OtpremnicaID']],
      ['Izvori','Original JSONL + roditeljske otpremnice + MIG18 QA XLSM'],
      ['Status AMB10 integracije','HOLD: tip ISTORIJSKI_ULAZ_IZVOR_NEPOZNAT nije u postojećoj AMB10 matrici'],
      ['VAŽNO','Ovo je kompletna istorijska XLSX sveska, NE izvršna VBA .xlsm sveska'],
      ['Uputstvo','Ne uvoziti tri UNKNOWN reda u aktivni AMB10 bez podrške za nepoznat izvor'],
    ]
    sheet_defs=[
      ('PREGLED',['Kontrola','Rezultat'],overview,()),
      ('SALDA',list(totals[0].keys()),[[x[k] for k in totals[0]] for x in totals],()),
      ('DNEVNA_PROVERA',list(dailies[0].keys()),[[x[k] for k in dailies[0]] for x in dailies],()),
      ('SVA_AMBALAZA',main_headers,main_rows,(4,)),
      ('ISTORIJA_7815',oldkeys+['DokumentStorniran'],
       [[r.get(k,'') for k in oldkeys]+['DA' if key(r.get('AmbID'))==EXCLUDED else ''] for r in old],(4,)),
      ('DODATO_6',list(linked[0].keys()),[[x[k] for k in linked[0]] for x in linked],()),
      ('PRAVILA',['DokumentTIP','Istorijsko značenje'],list(HISTORIC_DESCRIPTIONS.items()),()),
      ('DOKAZ_STORNA',list(parentproof.keys()),[[parentproof[k] for k in parentproof]],()),
    ]
    if (sha256(src),sha256(par),sha256(qa))!=orig_sha:
        raise RuntimeError('A source file changed during build')
    outdir.mkdir(parents=True,exist_ok=True)
    dest=outdir/'AgriX_MIG19_12_KOMPLETNA_ISTORIJSKA_AMBALAZA.xlsx'
    if dest.exists():raise FileExistsError(f'Refusing overwrite: {dest}')
    temp=outdir/'AgriX_MIG19_12_KOMPLETNA_ISTORIJSKA_AMBALAZA.tmp.xlsx'
    try:
        build_archive(temp,sheet_defs)
        os.replace(temp,dest)
    finally:
        if temp.exists():temp.unlink()
    checks['output_file']=str(dest)
    checks['output_sha256']=sha256(dest)
    (outdir/'MIG1912_SUMMARY.json').write_text(json.dumps(checks,indent=2,ensure_ascii=False),encoding='utf-8')
    if (sha256(src),sha256(par),sha256(qa))!=orig_sha:
        raise RuntimeError('Source was modified during export')
    return checks


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage',required=True,type=Path)
    p.add_argument('--workbook',required=True,type=Path)
    p.add_argument('--out-dir',required=True,type=Path)
    a=p.parse_args(argv)
    report=generate(a.stage.resolve(),a.workbook.resolve(),a.out_dir.resolve())
    print(json.dumps(report,indent=2,ensure_ascii=False))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:
        print(f'MIG1912_BLOCKED {type(e).__name__}: {e}',file=sys.stderr)
        sys.exit(3)
