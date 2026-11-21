#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MIG-17: isolated historical AMB-10 QA importer. STANDARD LIBRARY ONLY.

The old system had a signed packaging balance; negative cooperative balances
are historically a liability of the company. Within-day ordering has NO
business authority. This migration explicitly records compensating physical
entries ULAZ_TUDJE (from external boundary to partner) for net increases in
historical debt and classifies the debt-clearing portion of documented issued
packaging as VRACANJE_TUDJE. It NEVER equates the new physical stock with the
old signed balance, and never backdates audit creation timestamps.

The six unmatched physical cases and canceled old records remain quarantined.
The result is a QA workbook, NOT a claim of full canonical cutover. Writes to
NEW XLSM only, replacing two formerly empty table parts and worksheets, VBA
byte-identical. No COM, Excel save, Git change or network access.
"""
from __future__ import annotations
import argparse, copy, hashlib, json, os, re, sys, zipfile
from pathlib import Path
from collections import defaultdict,Counter
from datetime import date,datetime,timezone,timedelta
from decimal import Decimal
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET
from agrix_mig16_daily_close import make_rows,compute_daily_policy,probe_workbook
from agrix_mig15_semantic_replay import read_jsonl,txt,date_sort_key,amount,digest,ordinal
from agrix_mig11_write_links_ooxml import table_locations,patch_table_xml,parse_ref,col_name,dimension_patch,ROW_RE,SD_RE,replace_attr_ref,q,pos

INPUT_SHA='622ea7e68467a2f77b372463f73f9819cd43eb7a7f22973a19d692f713c068dc'
BASE_LEDGER='MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl'
HEADERS='MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl'
COOP='kooperant'
COLS={
 'tblAmbalaza':['AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd'],
 'tblAmbalazaDokument':['AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Napomena','Stornirano','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','IspravkaOdID','ZamenjenSaID']
}
KINDS={'AMBALAZA_UZ_ROBU','IZDATA_PRAZNA','POVRAT_PRAZNE','PRENOS_INTERNO','ULAZ_TUDJE_AMBALAZE','VRACANJE_TUDJE_AMBALAZE','NABAVKA','OTPIS'}
ACC={'kooperant','stanica','kupac','vozac','firma','spoljn isvet','spoljn isvet'}
SYS={'firma','spoljn isvet'}
SOURCE_LIMIT=date(2026,8,31)
NUMERIC={'Kolicina'}
DATES={'Datum','CreatedAt','ModifiedAt'}


def norm_num(d):
    d=Decimal(str(d))
    if not d.is_finite() or d<=0:raise ValueError('Amount must be finite and > 0')
    significant=len(d.normalize().as_tuple().digits)
    if significant>15:raise ValueError('Excel cannot faithfully hold a quantity with >15 significant digits')
    return format(d,'f')


def stable_id(kind,key):
    return 'AMB-M17-'+hashlib.sha256((kind+'|'+key).encode('utf-8')).hexdigest()[:18].upper()


def case_to_row(k,case,now):
    raw=dict(k['target_column_values_WITHOUT_AMBID'])
    raw.setdefault('StornoOd','')
    raw['AmbID']=txt(k['reserved_AmbID_NOT_ALLOCATED'])
    if not raw['AmbID']:raise ValueError('Empty reserved candidate ID')
    raw['CreatedAt']=now
    raw['CreatedBy']='MIG-17'
    raw['ModifiedAt']='';raw['ModifiedBy']=''
    for name in COLS['tblAmbalaza']:
        raw.setdefault(name,'')
    if any(txt(raw.get(x)) for x in ('StornoOd','ModifiedAt','ModifiedBy')):raise ValueError('Unexpected preexisting canonical storno/audit')
    if txt(raw.get('VrstaKretanja')) not in KINDS:raise ValueError('Unknown movement kind')
    if txt(raw.get('DokumentID'))=='' or txt(raw.get('DokumentTIP'))=='':raise ValueError('Missing historical documentary identity')
    day=date_sort_key(raw['Datum']).date()
    if day>SOURCE_LIMIT:raise ValueError('Postcutoff event in legacy adapter')
    raw['Datum']=day.isoformat()
    raw['Kolicina']=norm_num(raw['Kolicina'])
    return raw


def account_of(row,prefix):
    return (txt(row.get(prefix+'NalogTip')).casefold(),txt(row.get(prefix+'NalogID')).casefold())


def group_for(row,prefix):
    return (account_of(row,prefix)[1],txt(row['TipAmbalaze']).casefold(),date.fromisoformat(row['Datum']))


def convert_physical(rows,daily,now):
    """Daily CLOSE drives debt policy, independent of source intraday order.

    For decrease: a documented station->cooperator issuance can be relabeled
    VRACANJE_TUDJE (identical physical vector). For increase: add explicitly
    marked MIGRATION deficit-cover SpoljniSvet->Kooperant to the documented
    partner outflows. Extra rows are not disguised legacy transactions.
    """
    original={x['case']:case_to_row(x['raw'],x['case'],now) for x in rows}
    indexed=defaultdict(list)
    for x in rows:
        k=x['case'];r=original[k]
        if account_of(r,'Na')[0]==COOP:indexed.setdefault(group_for(r,'Na'),[]).append(k)
        if account_of(r,'Od')[0]==COOP:indexed.setdefault(group_for(r,'Od'),[]).append(k)
    entries=defaultdict(list);foreign_in=defaultdict(Decimal);foreign_return=defaultdict(Decimal)
    work={case:dict(r) for case,r in original.items()}
    extra=[]; provenance=[]; derived=Counter()
    for x in daily:
        day=date.fromisoformat(x['day']);key=(x['kooperant_id'],x['tip_ambalaze'],day)
        debt_plus=Decimal(x['obligation_increase']);debt_minus=Decimal(x['obligation_decrease'])
        candidates=sorted(set(indexed.get(key,[])),key=lambda case:(date.fromisoformat(original[case]['Datum']),case))
        take=debt_minus
        for case in candidates:
            row=work[case]
            if take<=0:break
            if row['VrstaKretanja']!='IZDATA_PRAZNA' or account_of(row,'Na')!=(COOP,key[0]):continue
            v=Decimal(row['Kolicina']);part=min(v,take)
            if part<=0:continue
            if part==v:
                row['VrstaKretanja']='VRACANJE_TUDJE_AMBALAZE'
                extra_ref=row['AmbID']
            else:
                # Split physical transfer; preserve existing reserved ID on original
                row['Kolicina']=norm_num(v-part)
                replacement=dict(row)
                replacement['AmbID']=stable_id('RETURN',case)
                replacement['Kolicina']=norm_num(part)
                replacement['VrstaKretanja']='VRACANJE_TUDJE_AMBALAZE'
                extra.append(replacement);extra_ref=replacement['AmbID'];derived['split_issue']+=1
            foreign_return[key]+=part
            provenance.append({'role':'FOREIGN_RETURN_DAILY_CLOSE','case':case,'AmbID':extra_ref,'qty':str(part),'date':day.isoformat()})
            take-=part
            derived['relabel_return_events']+=1
        if take!=0:raise ValueError('Decreased foreign liability exceeds documented issuance capacity')
        take=debt_plus
        for case in candidates:
            r=original[case]
            if take<=0:break
            if account_of(r,'Od')!=(COOP,key[0]):continue
            part=min(Decimal(r['Kolicina']),take)
            if part<=0:continue
            created=dict(r)
            created.update({'AmbID':stable_id('COVER',case),'Kolicina':norm_num(part),
                            'OdNalogTip':'SpoljniSvet','OdNalogID':'',
                            'NaNalogTip':r['OdNalogTip'],'NaNalogID':r['OdNalogID'],
                            'VrstaKretanja':'ULAZ_TUDJE_AMBALAZE', 'CreatedBy':'MIG-17'})
            extra.append(created)
            foreign_in[key]+=part
            provenance.append({'role':'FOREIGN_COVER_DAILY_CLOSE','case':case,'AmbID':created['AmbID'],'qty':str(part),'date':day.isoformat()})
            take-=part
            derived['foreign_cover_events']+=1
        if take!=0:raise ValueError('Increased foreign liability exceeds documented partner outflow capacity')
    # Check each original movement's physical transfer survived exactly even if split
    resulting=list(work.values())+extra
    ids=[r['AmbID'].casefold() for r in resulting]
    if len(ids)!=len(set(ids)):raise ValueError('Generated AmbID collision')
    physical=defaultdict(Decimal)
    for r in resulting:
        if r['VrstaKretanja']=='ULAZ_TUDJE_AMBALAZE':continue
        physical[(r['OdNalogTip'].casefold(),r['OdNalogID'].casefold(),r['NaNalogTip'].casefold(),r['NaNalogID'].casefold(),r['TipAmbalaze'].casefold(),r['DokumentID'].casefold(),r['DokumentTIP'].casefold(),r['Datum'])]+=Decimal(r['Kolicina'])
    expected=defaultdict(Decimal)
    for r in original.values():
        expected[(r['OdNalogTip'].casefold(),r['OdNalogID'].casefold(),r['NaNalogTip'].casefold(),r['NaNalogID'].casefold(),r['TipAmbalaze'].casefold(),r['DokumentID'].casefold(),r['DokumentTIP'].casefold(),r['Datum'])]+=Decimal(r['Kolicina'])
    if physical!=expected:raise ValueError('Physical provenance changed during classification/splitting')
    # AMB-INV-04: one canonical movement identity per document/kind/account pair/type.
    unique_keys=set()
    for row in resulting:
        identity=tuple(txt(row[k]).casefold() for k in ('DokumentTIP','DokumentID','TipAmbalaze','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja'))
        if identity in unique_keys:
            raise ValueError('AMB-INV-04 duplicate document movement identity after split/coverage')
        unique_keys.add(identity)
    # Verify daily obligation independently of allocation source order
    byday=defaultdict(lambda:[Decimal(0),Decimal(0)])
    for k,val in foreign_in.items():byday[k][0]+=val
    for k,val in foreign_return.items():byday[k][1]+=val
    debt=defaultdict(Decimal)
    daily_groups={(x['kooperant_id'],x['tip_ambalaze'],date.fromisoformat(x['day'])):x for x in daily}
    for key in sorted(daily_groups,key=lambda z:(z[2],z[0],z[1])):
        x=daily_groups[key];debt_key=key[:2]
        if debt[debt_key]!=Decimal(x['historical_obligation_at_start']):raise ValueError('Initial daily debt mismatch')
        u,v=byday[key];debt[debt_key]+=u-v
        if debt[debt_key]!=Decimal(x['historical_obligation_at_close']):raise ValueError('Daily ending debt mismatch')
    return resulting,provenance,dict(derived),len(original)


def validated_headers(project,now):
    raw=read_jsonl(project/HEADERS)
    if len(raw)!=48:raise ValueError('Unexpected canonical revers header count')
    result=[];ids=set()
    for obj in raw:
        h=obj.get('proposed_target_header')
        if not isinstance(h,dict):raise ValueError('Invalid revers header payload')
        row={k:txt(h.get(k)) for k in COLS['tblAmbalazaDokument']}
        day=date_sort_key(h.get('Datum')).date()  # original may be numeric Excel serial
        ident=row['AmbDokID']
        if not ident or ident.casefold() in ids:raise ValueError('Duplicate or empty revers identity')
        ids.add(ident.casefold())
        if row['Vrsta']!='REVERS' or not row['BrojDokumenta'] or not row['BrojOwnerTip']:
            raise ValueError('Unexpected historical revers header shape')
        if row['Stornirano'] not in ('','Ne','NE','ne'):
            raise ValueError('Unexpected revers storno status')
        row.update({'Datum':day.isoformat(),
                    'CreatedAt':now,'CreatedBy':'MIG-17', 'ModifiedAt':'', 'ModifiedBy':'',
                    'IspravkaOdID':'', 'ZamenjenSaID':''})
        result.append(row)
    return result


def excel_serial(text_value):
    if text_value=='':return ''
    dt=date_sort_key(text_value)
    serial=Decimal((dt.date()-date(1899,12,30)).days)+Decimal(dt.hour*3600+dt.minute*60+dt.second)/Decimal(86400)
    return str(serial)


def emitted_rows(records,cols,first_col,first_row,old_cells):
    styles={}
    for c in old_cells:
        ref=c.attrib.get('r','')
        if ref:
            col,_=pos(ref)
            if c.attrib.get('s'):styles[col]=c.attrib['s']
    fragments=[]
    for j,obj in enumerate(records):
        rn=first_row+j;xml=[]
        for i,key in enumerate(cols):
            val=obj.get(key,'')
            if val is None or val=='':continue
            col=first_col+i;ref=f'{col_name(col)}{rn}';style=f' s="{styles[col]}"' if col in styles else ''
            if key in NUMERIC or key in DATES:
                v=norm_num(val) if key in NUMERIC else excel_serial(val)
                xml.append(f'<c r="{ref}"{style}><v>{v}</v></c>')
            else:
                s=escape(txt(val))
                if s:xml.append(f'<c r="{ref}"{style} t="inlineStr"><is><t>{s}</t></is></c>')
        fragments.append(f'<row r="{rn}" spans="{first_col}:{first_col+len(cols)-1}">{"".join(xml)}</row>')
    return ''.join(fragments).encode('utf-8')


def patch_amb_sheet(raw,tab,columns,records):
    c1,r1,c2,r2=parse_ref(tab.attrib['ref']);first=r1+1;last=r1+len(records)
    if not records or (r2-r1)!=1 or (c2-c1+1)!=len(columns):raise ValueError('Unexpected empty template dimensions')
    parsed=ET.fromstring(raw);sd=parsed.find(q('sheetData'))
    if sd is None:raise ValueError('sheetData missing')
    stub=[]
    for row in sd.findall(q('row')):
        ri=int(row.attrib.get('r','0'))
        if first<=ri<=max(last,r2):
            for c in row.findall(q('c')):
                if c.find(q('f')) is not None or (c.find(q('v')) is not None and (c.find(q('v')).text or '').strip()) or ''.join(x.text or '' for x in c.iter(q('t'))).strip():
                    raise ValueError('Occupied cell in destination table expansion range')
                if ri==first:stub.append(c)
    new_rows=emitted_rows(records,columns,c1,first,stub)
    m=SD_RE.search(raw)
    if not m:raise ValueError('Unsupported sheetData encoding')
    old=m.group(2);matches=list(ROW_RE.finditer(old))
    if len(matches)!=len(sd.findall(q('row'))):raise ValueError('Unexpected worksheet row encoding')
    out=[];inserted=False;previous=0
    for found in matches:
        item=found.group(0)
        match=re.search(rb'<row\b[^>]*\br\s*=\s*["\'](\d+)["\']',item)
        if not match:raise ValueError('Worksheet row missing index')
        rn=int(match.group(1))
        if rn<=previous:raise ValueError('Out of order source worksheet rows')
        previous=rn
        if rn>=first and not inserted:
            out.append(new_rows);inserted=True
        if rn<first or rn>max(last,r2):out.append(item)
    if not inserted:out.append(new_rows)
    changed=raw[:m.start(2)]+b''.join(out)+raw[m.end(2):]
    changed=dimension_patch(changed,c1,c2,r1,last)
    root=ET.fromstring(changed)
    rows=root.find(q('sheetData')).findall(q('row'))
    if len([1 for r in rows if first<=int(r.attrib['r'])<=last])!=len(records):raise ValueError('New sheet row count mismatch')
    return changed


def build_patches(z,tabledata):
    located,_=table_locations(z);patches={}
    for name,records in tabledata.items():
        if name not in located:raise ValueError('Missing table '+name)
        tp,sp,tab=located[name]
        cols=[c.attrib['name'] for c in tab.find(q('tableColumns')).findall(q('tableColumn'))]
        if cols!=COLS[name]:raise ValueError('Unexpected canonical schema for '+name)
        if sum(loc[1]==sp for loc in located.values())!=1:raise ValueError('Target sheet shares another table')
        c1,r1,c2,r2=parse_ref(tab.attrib['ref']);new_ref=f'{col_name(c1)}{r1}:{col_name(c2)}{r1+len(records)}'
        if sp in patches or tp in patches:raise ValueError('Collision between modified OOXML parts')
        patches[sp]=patch_amb_sheet(z.read(sp),tab,cols,records)
        patches[tp]=patch_table_xml(z.read(tp),tab.attrib['ref'],new_ref)
    return patches


def check_rows(z,counts):
    located,_=table_locations(z)
    for t,n in counts.items():
        tp,sp,tab=located[t];c1,r1,c2,r2=parse_ref(tab.attrib['ref'])
        if r2-r1!=n:raise ValueError('Postwrite Excel table size mismatch')
        sd=ET.fromstring(z.read(sp)).find(q('sheetData'))
        actual=sum(r1<int(r.attrib['r'])<=r2 for r in sd.findall(q('row')))
        if actual!=n:raise ValueError('Postwrite worksheet row count mismatch')


def preflight(project,source):
    rows,holds,headers,exceptions,checks,base=make_rows(project)
    if len(holds)!=6 or len(headers)!=48 or len(exceptions)!=8:raise ValueError('Known historical quarantine count diverged')
    if any(x['day']>SOURCE_LIMIT for x in rows):raise ValueError('Source date outside historical cutoff')
    if any(txt(base[txt(h['physical_case_id'])].get(side)).casefold()==COOP for h in holds for side in ('from_account_type','to_account_type')):
        raise ValueError('Held physical case touches cooperant; daily parity not complete')
    daily,counts=compute_daily_policy(rows)
    if counts.get('account_type_day_closes')!=1458 or counts.get('days_with_increased_historical_obligation')!=95 or counts.get('days_with_decreased_historical_obligation')!=48:
        raise ValueError('Daily closing policy diverges from verified MIG16')
    if counts.get('days_obligation_reduction_without_matching_issuance',0) or counts.get('days_obligation_increase_without_matching_outflow',0):
        raise ValueError('Daily close lacks original documentary physical transfers')
    if digest(source)!=INPUT_SHA:raise ValueError('Wrong input workbook SHA. Use exact untouched MIG14_STATUS baseline')
    workbook=probe_workbook(source)
    return rows,holds,headers,exceptions,daily,counts,workbook


def run(project,source,outdir):
    project=project.resolve();source=source.resolve();outdir=outdir.resolve()
    if outdir.exists():raise ValueError('Output folder already exists; refusing overwrite')
    if not source.is_file():raise ValueError('Input XLSM not found')
    if project in outdir.parents and any((project/x) in outdir.parents for x in ('MigrationStage_v2','MigrationAmb10TargetAdapter_v2','MigrationAmb10ImportReadiness_v2')):
        raise ValueError('Output folder overlaps historical source')
    rows,holds,head,exceptions,daily,close,probe=preflight(project,source)
    raw_by_case={txt(x['physical_case_id']):x for x in read_jsonl(project/BASE_LEDGER)}
    if len(raw_by_case)!=4669:raise ValueError('Invalid/duplicate physical cases')
    for item in rows:item['raw']=raw_by_case[item['case']]
    now=datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S')
    books,crosswalk,derivation,original_count=convert_physical(rows,daily,now)
    headers=validated_headers(project,now)
    header_ids={x['AmbDokID'].casefold() for x in headers}
    for b in books:
        if txt(b['DokumentTIP']).casefold()=='ambalazadokument' and b['DokumentID'].casefold() not in header_ids:
            raise ValueError('Unresolved AmbalazaDokument FK')
    # Provenance local ONLY, intentionally not embedded in generated office workbook.
    tabledata={'tblAmbalaza':books,'tblAmbalazaDokument':headers}
    with zipfile.ZipFile(source) as zin:
        if zin.testzip():raise ValueError('Bad XLSM ZIP')
        names=zin.namelist()
        if len(names)!=len(set(names)) or 'xl/vbaProject.bin' not in names:raise ValueError('Missing VBA or duplicate XLSM ZIP part')
        patches=build_patches(zin,tabledata)
        outdir.mkdir(parents=True)
        temp=outdir/'AgriX_MIG17_AMBALAZA_QA.xlsm.part';final=outdir/'AgriX_MIG17_AMBALAZA_QA.xlsm'
        try:
            with zipfile.ZipFile(temp,'w') as zout:
                for item in zin.infolist():zout.writestr(copy.copy(item),patches.get(item.filename,zin.read(item.filename)))
            with zipfile.ZipFile(temp) as result:
                if result.testzip():raise ValueError('Bad output XLSM ZIP')
                if set(result.namelist())!=set(names):raise ValueError('ZIP member set changed')
                for name in names:
                    if name not in patches and result.read(name)!=zin.read(name):raise ValueError('Unintended OOXML part mutation')
                if result.read('xl/vbaProject.bin')!=zin.read('xl/vbaProject.bin'):raise ValueError('VBA mutated')
                check_rows(result,{k:len(v) for k,v in tabledata.items()})
            os.replace(temp,final)
        except Exception:
            temp.unlink(missing_ok=True)
            final.unlink(missing_ok=True)
            raise
    outcross=outdir/'MIG17_provenance_PRIVATE.jsonl'
    with outcross.open('w',encoding='utf-8') as f:
        for p in crosswalk:f.write(json.dumps(p,ensure_ascii=False)+'\n')
    manifest={'format':'AGRIx-MIG17-HISTORICAL-AMBALAZA-QA-v1',
      'status':'TECHNICAL_QA_IMPORT_WITH_KNOWN_HISTORICAL_HOLDS',
      'canonical_final_approved':False,'output_xlsm_created':True,
      'input_xlsm_sha256':INPUT_SHA,'output_xlsm_sha256':digest(final),
      'candidate_physical_events':original_count,'canonical_ledger_rows_written':len(books),
      'revers_headers_written':len(headers),
      'derived_migration_events':derivation,
      'daily_closings_checked':len(daily),'daily_debt_reconciled':True,
      'physical_ledger_vectors_preserved_except_explicit_external_liability_covers':True,
      'source_quarantine':{'unmodelled_physical_events':len(holds),'storno_exception_register':len(exceptions),'cancelled_legacy_rows_archived':335,'q00700_held':True},
      'modified_parts':sorted(patches),'vba_binary_byte_identical':True,
      'all_other_ooxml_parts_byte_identical':True,
      'business_policy':'Historical signed daily close authoritative; foreign obligations derived only by documented same-day capacities; intra-day order not asserted',
      'warning':'This is QA, not a complete operational cutover. Six physical events and Q00700 remain quarantined. Synthetic foreign-coverage rows are migration inference, not proof of legal title.',
      'customer_private_values_in_report':False,'excel_runtime_qa_pending':True}
    (outdir/'MIG17_final_verify.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('MIG-17 QA XLSM CREATED; KNOWN HISTORICAL HOLDS REMAIN')
    print('LEDGER:',len(books),'REVERS HEADERS:',len(headers),'DAILY CLOSINGS:',len(daily))
    print('VBA: UNCHANGED; INPUT: UNCHANGED')
    print('OUTPUT:',final)
    print('SHARE ONLY:',outdir/'MIG17_final_verify.json')
    return manifest


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--project-dir',required=True,type=Path)
    ap.add_argument('--input-xlsm',required=True,type=Path)
    ap.add_argument('--out-dir',required=True,type=Path)
    a=ap.parse_args()
    try:run(a.project_dir,a.input_xlsm,a.out_dir)
    except Exception as e:
        print('MIG17 BLOCKED:',type(e).__name__,'—',str(e),file=sys.stderr)
        return 2
    return 0
if __name__=='__main__':sys.exit(main())
