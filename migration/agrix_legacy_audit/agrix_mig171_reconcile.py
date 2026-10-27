#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIx MIG-17.1: local READ-ONLY historic AMB / payment reconciliation.

No Excel automation, no workbook writing, no Git and no network. All private
per-person figures are written ONLY into *_PRIVATE.csv. Send ONLY summary.json.
"""
from __future__ import annotations
import argparse, csv, hashlib, json, re, sys, zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

from agrix_mig11_write_links_ooxml import table_locations, parse_ref, q, pos
from agrix_mig16_daily_close import make_rows, compute_daily_policy

D=Decimal
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ZERO=D(0)
EPOCH=date(1899,12,30)
CUT=date(2026,8,31)


def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for v in iter(lambda:f.read(1024*1024),b''):h.update(v)
    return h.hexdigest()

def dec(v,field='numeric'):
    if isinstance(v,Decimal):return v
    if v is None or str(v).strip()=='':return ZERO
    try:
        a=D(str(v).strip().replace(',','.'))
        if not a.is_finite():raise ValueError
        return a
    except (InvalidOperation,ValueError) as e:raise ValueError(f'Unusable {field} value') from e

def norm(x):return str(x or '').strip().casefold()

def day(x):
    if x is None or str(x).strip()=='':return None
    if isinstance(x, date):return x
    s=str(x).strip()
    try:
        d=D(s)
        if d >= 1 and d < 200000:
            return EPOCH+timedelta(days=int(d))
    except (ValueError,InvalidOperation):pass
    for fmt in ('%Y-%m-%d','%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S', '%d.%m.%Y', '%d/%m/%Y'):
        try:return datetime.strptime(s[:len(datetime.now().strftime(fmt))],fmt).date()
        except ValueError:pass
    raise ValueError('Unsupported stored date format (private value omitted)')

def styles(z):
    if 'xl/styles.xml' not in z.namelist():return []
    root=ET.fromstring(z.read('xl/styles.xml'))
    customs={}
    for node in root.findall(N+'numFmts/'+N+'numFmt'):
        customs[int(node.attrib['numFmtId'])]=node.attrib.get('formatCode','')
    xfs=root.find(N+'cellXfs')
    if xfs is None:return []
    out=[]
    for x in xfs.findall(N+'xf'):
        n=int(x.attrib.get('numFmtId','0'))
        if n in set(range(14,23))|set(range(27,37))|set(range(45,48))|set(range(50,59)):
            out.append(True);continue
        f=customs.get(n,'')
        # Remove quoted / escaped literals, then check date/time pattern.
        f=re.sub(r'"[^"]*"|\\.|\[[^]]*\]', '', f)
        out.append(bool(re.search(r'[dmyhs]',f,re.I)))
    return out

def shared(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(x.text or '' for x in si.iter(N+'t')) for si in root.findall(N+'si')]

def value(c,string_table):
    t=c.attrib.get('t','')
    if t=='inlineStr':return ''.join(x.text or '' for x in c.iter(N+'t'))
    v=c.find(N+'v')
    if v is None or v.text is None:return ''
    if t=='s':
        i=int(v.text)
        if i>=len(string_table):raise ValueError('Invalid shared string reference')
        return string_table[i]
    if t in ('e','b'):return v.text
    return v.text

def read_book(path,tables):
    """Return canonical rows plus date-cell Excel formatting classifications."""
    result={};formats={}
    with zipfile.ZipFile(path) as z:
        if z.testzip():raise ValueError('Damaged XLSM archive')
        if 'xl/vbaProject.bin' not in z.namelist():raise ValueError('XLSM lacks VBA binary')
        loc,_=table_locations(z)
        sst=shared(z);xs=styles(z)
        for name in tables:
            if name not in loc:raise ValueError('Missing canonical table '+name)
            tp,sp,root=loc[name]
            cols=[c.attrib['name'] for c in root.find(q('tableColumns')).findall(q('tableColumn'))]
            c1,r1,c2,r2=parse_ref(root.attrib['ref'])
            if len(cols)!=(c2-c1+1):raise ValueError('Unexpected table width for '+name)
            raw=ET.fromstring(z.read(sp));sd=raw.find(N+'sheetData')
            if sd is None:raise ValueError('sheetData not found')
            table=[];styles_found=Counter(); datecol='Datum'
            for row in sd.findall(N+'row'):
                ri=int(row.attrib.get('r',0))
                if ri<=r1 or ri>r2:continue
                obj={k:'' for k in cols}
                for c in row.findall(N+'c'):
                    coord=c.attrib.get('r','')
                    if not coord:continue
                    col,rr=pos(coord)
                    if rr!=ri or not(c1<=col<=c2):continue
                    key=cols[col-c1]; obj[key]=value(c,sst)
                    if key==datecol:
                        st=int(c.attrib.get('s',0))
                        styles_found['date_format' if st<len(xs) and xs[st] else 'not_date_format']+=1
                        styles_found['date_cell_numeric' if c.attrib.get('t','') not in ('s','inlineStr','str') else 'date_cell_text']+=1
                table.append(obj)
            if len(table)!=(r2-r1):raise ValueError(f'Expected {r2-r1} rows in {name}, parsed {len(table)}')
            result[name]=table;formats[name]=dict(styles_found)
    return result,formats

def csv_out(path,records,cols):
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.DictWriter(f,fieldnames=cols,extrasaction='ignore');w.writeheader();w.writerows(records)

def amb_diagnostics(book,project,station,expected_covers=None):
    cases,_,_,_,_,_=make_rows(project)
    daily,_=compute_daily_policy(cases)
    expected={}
    for r in daily:
        k=(norm(r['kooperant_id']),norm(r['tip_ambalaze']))
        # Last daily record for each key by date, not lexical/source order
        daykey=day(r['day'])
        if k not in expected or daykey>expected[k][0]: expected[k]=(daykey,dec(r['ending_signed_balance']))
    signed={k:v[1] for k,v in expected.items()}
    physical=defaultdict(lambda:ZERO);covers=defaultdict(lambda:ZERO)
    names=Counter();cover_count=0
    for row in book['tblAmbalaza']:
        ktype=norm(row.get('TipAmbalaze')); val=dec(row.get('Kolicina'),'Amb quantity')
        names[norm(row.get('VrstaKretanja'))]+=1
        for side,sgn in [('Na',1),('Od',-1)]:
            if norm(row.get(side+'NalogTip'))=='kooperant':
                k=(norm(row.get(side+'NalogID')),ktype)
                physical[k]+=sgn*val
        if norm(row.get('VrstaKretanja'))=='ulaz_tudje_ambalaze' and norm(row.get('CreatedBy'))=='mig-17':
            if norm(row.get('OdNalogTip'))!='spoljnisvet':
                raise ValueError('MIG-17 foreign cover has unexpected from account')
            if norm(row.get('NaNalogTip'))!='kooperant':raise ValueError('MIG-17 foreign cover has wrong partner')
            covers[(norm(row.get('NaNalogID')),ktype)]+=val
            cover_count+=1
    if expected_covers is not None and cover_count!=expected_covers:
        raise ValueError('MIG17 foreign cover count diverged')
    keys=set(signed)|set(physical)|set(covers)
    records=[];bad=0
    for k in sorted(keys):
        p=physical[k];c=covers[k];e=signed.get(k,ZERO);adj=p-c
        diff=adj-e
        if diff!=0:bad+=1
        records.append({'kooperant_id':k[0],'tip_ambalaze':k[1],'old_signed_close':str(e),
            'new_physical':str(p),'mig17_cover':str(c),'derived_legacy_signed':str(adj),'difference':str(diff)})
    coop_at_station=set()
    for r in book['tblOtkup']:
        if norm(r.get('StanicaID'))==norm(station) and norm(r.get('Stornirano'))!='da':
            coop_at_station.add(norm(r.get('KooperantID')))
    def agg(field):return sum((dec(r[field]) for r in records if r['kooperant_id'] in coop_at_station),ZERO)
    return records,{
        'canonical_amb_rows':len(book['tblAmbalaza']),
        'legacy_account_type_pairs':len(signed),
        'compared_pairs':len(records),'mismatch_pairs':bad,
        'inferred_foreign_cover_rows':cover_count,
        'station_kooperants_from_otkup':len(coop_at_station),
        'station_adjustment_nonzero':agg('mig17_cover')!=0,
        'station_legacy_signed_matches_expected_237':agg('old_signed_close')==D(237),
        'station_legacy_reconstructed_matches_expected_237':agg('derived_legacy_signed')==D(237),
        'station_original_and_reconstructed_equal':agg('old_signed_close')==agg('derived_legacy_signed'),
        'all_global_signed_parities':bad==0
    }

def source_finance_audit(project, book):
    stage=project/'MigrationStage_v2'
    ps=list(stage.rglob('SECONDARY__tblNovac.jsonl'))
    if len(ps)!=1:raise ValueError('Missing/ambiguous original secondary finance JSONL')
    source=ps[0]
    if sha(source)!='7731dca7d0167ca9dad69374d4f35ae54e05fb1cacf3800fb67d4191736077b0':
        raise ValueError('Source finance checksum changed')
    original=[]
    with source.open(encoding='utf-8-sig') as f:
        for line in f:
            if line.strip():original.append(json.loads(line))
    if len(original)!=39:raise ValueError('Unexpected original finance row count')
    old={norm(x.get('NovacID')):x for x in original}
    newer={norm(x.get('NovacID')):x for x in book['tblNovac']}
    if len(old)!=39 or len(newer)!=len(book['tblNovac']) or set(old)!=set(newer):
        raise ValueError('Source and workbook financial identities differ')
    discrepancies=Counter()
    diff_q=ZERO
    for key,src in old.items():
        dst=newer[key]
        for col in ['KooperantID','OMID','Stornirano','Tip','PartnerID','OtkupID']:
            if norm(src.get(col))!=norm(dst.get(col)):
                discrepancies[col]+=1
        for col in ['Isplata','Uplata']:
            delta=abs(dec(src.get(col))-dec(dst.get(col)))
            if delta>D('0.00000001'):discrepancies[col]+=1
            diff_q+=delta
        if day(src.get('Datum'))!=day(dst.get('Datum')):discrepancies['Datum']+=1
    src_total=sum((dec(r.get('Isplata')) for r in old.values()),ZERO)
    return {'source_rows':len(original),'workbook_finance_rows':len(newer),
            'per_field_mismatch_counts':dict(discrepancies),
            'source_workbook_data_reconciled':not discrepancies,
            'source_all_isplata_matches_old_526812_50':abs(src_total-D('526812.50'))<D('0.00000001'),
            'source_checksum_verified':True}


def legacy_finance_source(project):
    stage=project/'MigrationStage_v2'
    p=stage/'SOURCE__tblOtkup.jsonl'
    if not p.exists():
        xs=list(stage.rglob('SOURCE__tblOtkup.jsonl'))
        if len(xs)!=1:raise ValueError('Cannot uniquely identify SOURCE__tblOtkup.jsonl')
        p=xs[0]
    # Only candidate columns - discovering does NOT certify their bookkeeping meaning.
    count=Counter();summed=defaultdict(lambda:ZERO)
    candidate=set()
    with p.open(encoding='utf-8-sig') as f:
        for line in f:
            if not line.strip():continue
            row=json.loads(line)
            for key,v in row.items():
                if re.search(r'ispla|plac|novac|uplata|gotovina|paid',key,re.I):
                    candidate.add(key)
                    if str(v or '').strip():
                        count[key]+=1
                        try:summed[key]+=dec(v)
                        except ValueError:pass
    return [{'field':k,'nonempty_source_rows':count[k], 'sum_if_numeric_NOT_A_CONFIRMED_PAYMENT':str(summed[k])} for k in sorted(candidate)]

def finance_diagnostics(book,station):
    known={norm(r.get('KooperantID')):norm(r.get('StanicaID')) for r in book['tblKooperanti']}
    summary=Counter();total=defaultdict(lambda:ZERO);detail=[]
    for row in book['tblNovac']:
        summary['rows']+=1
        amount=dec(row.get('Isplata'),'Novac.Isplata')
        total['all_isplata']+=amount
        payment_day=day(row.get('Datum'))
        koop=norm(row.get('KooperantID'));om=norm(row.get('OMID'))
        in_window=payment_day is not None and date(2026,1,1)<=payment_day<=date(2026,10,21)
        if not koop:summary['missing_koop']+=1
        if not in_window:summary['not_in_report_date_window']+=1
        if amount<=0:summary['no_positive_isplata']+=1
        effective=om or known.get(koop,'')
        station_match=effective==norm(station)
        if not station_match:summary['other_om_or_unknown_om']+=1
        if norm(row.get('Stornirano'))=='da':summary['stornirano_rows']+=1
        passes=bool(koop and in_window and station_match and norm(row.get('Stornirano'))!='da')
        if passes:summary['in_report_scope_rows']+=1;total['in_report_scope_isplata']+=amount
        if passes and amount>0:summary['in_report_scope_positive_isplata_rows']+=1
        detail.append({'NovacID':row.get('NovacID'),'KooperantID':row.get('KooperantID'),
            'OMID':row.get('OMID'),'Datum':row.get('Datum'),'Isplata':str(amount),
            'Stornirano':row.get('Stornirano'),'in_report_date':in_window,'in_report_station':station_match,
            'passes_ReportSaldoOM_document_filters':passes})
    totals={
        'rows_checked':summary['rows'],
        'in_report_scope_rows':summary['in_report_scope_rows'],
        'in_report_scope_positive_payment_rows':summary['in_report_scope_positive_isplata_rows'],
        'row_filters':dict(summary),
        'nonzero_paid_expected_from_book':total['in_report_scope_isplata']!=0,
        'book_paid_matches_old_report_526812_50':total['in_report_scope_isplata']==D('526812.50'),
    }
    return detail,totals, {k:str(v) for k,v in total.items()}

def run(project,input_xlsm,outdir,station):
    project=project.resolve();input_xlsm=input_xlsm.resolve();outdir=outdir.resolve()
    if not input_xlsm.is_file():raise ValueError('XLSM missing')
    if outdir.exists():raise ValueError('Output folder exists; use a new name to avoid overwriting forensic results')
    if project==outdir or outdir in input_xlsm.parents:raise ValueError('Bad output path')
    book,formats=read_book(input_xlsm,['tblAmbalaza','tblNovac','tblOtkup','tblKooperanti'])
    if len(book['tblAmbalaza'])!=4811 or len(book['tblNovac'])!=39:
        raise ValueError('Not the known MIG17 QA workbook: expected 4811 packaging and 39 finance rows')
    amb_records,amb=amb_diagnostics(book,project,station,expected_covers=95)
    finance_records,fin,fin_totals=finance_diagnostics(book,station)
    source_verification=source_finance_audit(project,book)
    source_fin=legacy_finance_source(project)
    outdir.mkdir(parents=True)
    csv_out(outdir/'MIG171_ambalaza_PRIVATE.csv',amb_records,
            ['kooperant_id','tip_ambalaze','old_signed_close','new_physical','mig17_cover','derived_legacy_signed','difference'])
    csv_out(outdir/'MIG171_finance_PRIVATE.csv',finance_records,
            ['NovacID','KooperantID','OMID','Datum','Isplata','Stornirano','in_report_date','in_report_station','passes_ReportSaldoOM_document_filters'])
    csv_out(outdir/'MIG171_old_otkup_possible_payment_fields_PRIVATE.csv',source_fin,
            ['field','nonempty_source_rows','sum_if_numeric_NOT_A_CONFIRMED_PAYMENT'])
    (outdir/'MIG171_private_finance_aggregate.json').write_text(json.dumps(fin_totals,indent=2)+'\n',encoding='utf-8')
    status='AMB_PARITY_PASS_FINANCE_REQUIRES_REVIEW' if amb['all_global_signed_parities'] else 'AMB_RECONCILIATION_MISMATCH'
    summary={'format':'AGRIx-MIG171-AMBUSINESS-FINANCE-READONLY-v1',
       'status':status,'workbook_modified':False,'git_modified':False,
       'workbook_basename':input_xlsm.name,'workbook_sha256':sha(input_xlsm),
       'historical_cutoff':'2026-08-31',
       'station_check':station,
       'ambalaza':amb,
       'finance':fin,
       'finance_source_integrity':source_verification,
       'excel_date_cell_styles':formats,
       'source_payment_candidate_fields':[{'field':x['field'],'nonempty_source_rows':x['nonempty_source_rows']} for x in source_fin],
       'source_finance_field_interpretation':'Not proven: amounts in SOURCE__tblOtkup are not automatically payable transactions',
       'private_reports_created':3,
       'next_action':'Use aggregate result to decide narrowly scoped data/UI correction; never append finance blindly',
       'customer_values_in_shareable_summary':False}
    (outdir/'MIG171_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('MIG-17.1 READ ONLY:',status)
    print('AMB matching pairs:',len(amb_records)-amb['mismatch_pairs'],'/',len(amb_records))
    print('FINANCE scoped rows:',fin['in_report_scope_rows'],'positive payments:',fin['in_report_scope_positive_payment_rows'])
    print('DO NOT SHARE *_PRIVATE.csv or private finance aggregate')
    print('SHARE ONLY:',outdir/'MIG171_summary.json')
    return summary

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--project-dir',required=True,type=Path)
    ap.add_argument('--input-xlsm',required=True,type=Path)
    ap.add_argument('--out-dir',required=True,type=Path)
    ap.add_argument('--station-id',default='ST-00001')
    a=ap.parse_args()
    try:run(a.project_dir,a.input_xlsm,a.out_dir,a.station_id)
    except Exception as e:
        print('MIG17.1 BLOCKED:',type(e).__name__,'—',str(e),file=sys.stderr)
        return 2
    return 0

if __name__=='__main__':sys.exit(main())
