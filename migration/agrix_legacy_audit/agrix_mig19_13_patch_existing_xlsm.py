#!/usr/bin/env python3
"""MIG19.13 safe patch existing AgriX QA XLSM on operator's Windows computer.

Transfers: 3 source-proven OMuUlaz rows -> AMB-10 PRENOS_INTERNO, with
  3 source-linked REVERS headers. The other 3 legacy single-sided receipts
  with UNKNOWN origin go into *a sheet in the XLSM*, not into an invalid
  AMB10 ledger row, until the code/model supports unknown origin.
AMB-00700: parent-storno proof -> no active ledger entry; leave source intact.

Fails closed on inputs, versions, rows, account parity, collisions or ambiguity.
Outputs new XLSM; never modifies original XLSM or staging.
"""
from __future__ import annotations
import argparse
import copy
from collections import Counter
from datetime import datetime, date as dcls
from decimal import Decimal
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from agrix_mig19_3_parity import XlsmReadOnly, sha256, stage_rows, ledger_views, reconcile, date, number, norm_type, key, void, STAGE_SHA, dec
from agrix_mig19_11_istorijska_semantika import OMULAZ, EXPECTED_TARGET_SHA
from agrix_mig19_12_kompletna_istorijska_ambalaza import stage_jsonl, parent_storno_proof, PARENT_SHA

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('',M)
ET.register_namespace('r',R)
ET.register_namespace('x14ac','http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
ET.register_namespace('mc','http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x15','http://schemas.microsoft.com/office/spreadsheetml/2010/11/main')
ET.register_namespace('xr','http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
ET.register_namespace('xr2','http://schemas.microsoft.com/office/spreadsheetml/2015/revision2')
ET.register_namespace('xr3','http://schemas.microsoft.com/office/spreadsheetml/2016/revision3')
ET.register_namespace('vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')

CANON_COLUMNS=['AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd']
DOC_COLUMNS=['AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Napomena','Stornirano','CreatedAt','CreatedBy','ModifiedAt','ModifiedBy','IspravkaOdID','ZamenjenSaID']
UNKNOWN=['AMB-01377','AMB-07115','AMB-07815']
TRANSFERS=['AMB-02317','AMB-03849','AMB-03939']
EXCLUDE='AMB-00700'
EXPECTED_BEFORE={('STANICA','ST-00001','0,4KG'):(-15461,-15479),('STANICA','ST-00002','0,4KG'):(-3577,-3660),('STANICA','ST-00004','0,4KG'):(-4673,-4173),('STANICA','ST-00005','0,4KG'):(-8093,-8643)}


def parse(s):return ET.fromstring(s)
def qn(uri,name):return '{'+uri+'}'+name
def xml(root, original=None):
    result=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if original:
        # ElementTree may drop unused namespace declarations that Excel still
        # references in the literal mc:Ignorable attribute. Preserve originals.
        src=original.decode('utf-8-sig') if isinstance(original,bytes) else str(original)
        src_start=re.search(r'<(?:[A-Za-z_][\w.]*:)?[A-Za-z_][\w.]*\s[^>]*>',src)
        out=result.decode('utf-8')
        dst_start=re.search(r'<(?:[A-Za-z_][\w.]*:)?[A-Za-z_][\w.]*\s[^>]*>',out)
        if src_start and dst_start:
            decls=re.findall(r'xmlns(?::[A-Za-z_][\w.-]*)?=["\'][^"\']+["\']',src_start.group(0))
            prepend=''
            for decl in decls:
                attr=decl.split('=',1)[0]
                if not re.search(r'(?<![\w:])'+re.escape(attr)+r'=',dst_start.group(0)):
                    prepend+=' '+decl
            if prepend:
                out=out[:dst_start.end()-1]+prepend+out[dst_start.end()-1:]
                result=out.encode('utf-8')
    return result
def xl_col(n):
    out=''
    while n:n,k=divmod(n-1,26);out=chr(65+k)+out
    return out

def excel_date(value):return (datetime.strptime(value,'%Y-%m-%d')-datetime(1899,12,30)).days

def ensure_source(old,parents):
    if len(old)!=7815:raise ValueError('Frozen source rows must be 7815')
    idx={}
    for r in old:
        aid=key(r.get('AmbID'))
        if aid in idx:raise ValueError('Duplicate old AmbID '+aid)
        idx[aid]=r
    if len(idx)!=7815:raise ValueError('Old ID count changed')
    proof=parent_storno_proof(old,parents)
    for aid,(day,dest,qty,src) in OMULAZ.items():
        if aid not in idx:raise ValueError('Missing original '+aid)
        r=idx[aid]
        if (void(r.get('Stornirano')) or date(r.get('Datum'),aid)!=day
            or norm_type(r.get('TipAmbalaze'))!='0,4KG'
            or number(r.get('Kolicina'),aid)!=qty
            or key(r.get('DokumentTIP'))!='OMULAZ'
            or key(r.get('Smer'))!='ULAZ'
            or key(r.get('EntitetTip'))!='STANICA'
            or key(r.get('EntitetID'))!=dest):
            raise ValueError('Source signature changed '+aid)
    return idx,proof


def target_preflight(old,new,docs):
    if len(new)!=4811 or len(docs)!=48:raise ValueError('Frozen QA row counts differ: amb=%s docs=%s'%(len(new),len(docs)))
    for expected in ['AmbID','Datum','TipAmbalaze','Kolicina','DokumentTIP','DokumentID','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd']:
        if expected not in new[0]:raise ValueError('Missing AMB10 column: '+expected)
    ids={key(r['AmbID']) for r in new}
    if len(ids)!=len(new):raise ValueError('Canonical duplicate AmbID')
    for aid in OMULAZ:
        if 'AMB-MIG19-'+aid[4:] in ids:raise ValueError('Already patched '+aid)
    od,phy,nd,co,odoc,ndoc,_,_=ledger_views(old,new)
    balances,_,_=reconcile(od,phy,nd,co,odoc,ndoc)
    wrong=[r for r in balances if r['Status']!='PASS']
    if len(wrong)!=4:raise ValueError('Expected 4 mismatches before: '+str(wrong[:8]))
    exp={(r['NalogTip'],r['NalogID'],r['TipAmbalaze']):(int(Decimal(r['StariSaldo'])),int(Decimal(r['NoviIstorijskiSaldo']))) for r in wrong}
    if exp!=EXPECTED_BEFORE:raise ValueError('Before saldos changed: '+str(exp))


def planned_rows(idx):
    a=[];docs=[]
    for aid in TRANSFERS:
        original=idx[aid];day,dest,qty,src=OMULAZ[aid]
        suffix=aid.split('-')[1]
        header='ADK-MIG19-'+suffix
        newid='AMB-MIG19-'+suffix
        doc=dict(AmbDokID=header,Vrsta='REVERS',BrojDokumenta='MIG19/'+suffix,
          Datum=day,BrojOwnerTip='Stanica',BrojOwnerID='ST-00001',
          Napomena='Istorijski OMuUlaz '+aid+'; izvorni dokument '+str(original.get('DokumentID') or '')+'; ST-00001 -> '+dest,
          Stornirano='',CreatedAt=original.get('CreatedAt') or '',CreatedBy='MIG19',
          ModifiedAt='',ModifiedBy='',IspravkaOdID='',ZamenjenSaID='')
        row=dict(AmbID=newid,Datum=day,TipAmbalaze=original.get('TipAmbalaze') or '0,4kg',
          Kolicina=qty,DokumentID=header,DokumentTIP='AmbalazaDokument',
          CreatedAt=original.get('CreatedAt') or '',CreatedBy='MIG19',ModifiedAt='',ModifiedBy='',
          OdNalogTip='Stanica',OdNalogID='ST-00001',NaNalogTip='Stanica',NaNalogID=dest,
          VrstaKretanja='PRENOS_INTERNO',StornoOd='')
        a.append(row);docs.append(doc)
    return a,docs


def verify_projection(old,existing,append):
    # Frozen old is unchanged and counts 500 active by flag; parent is cancelled.
    # Reconstruct old business effect without mutating old source rows.
    calc_old=[dict(r, Stornirano='Da') if key(r.get('AmbID'))==EXCLUDE else dict(r) for r in old]
    for aid in TRANSFERS:
        day,dest,qty,src=OMULAZ[aid]
        orig=next(r for r in old if key(r.get('AmbID'))==aid)
        calc_old.append(dict(AmbID='ONLY-PROJECTION-'+aid,Datum=day,TipAmbalaze=orig['TipAmbalaze'],
               Kolicina=qty,Smer='Izlaz',EntitetTip='Stanica',EntitetID=src,
               DokumentTIP='OMUlaz',DokumentID=orig.get('DokumentID') or aid,Stornirano='Ne'))
    o,p,n,cover,do,dn,diag,_=ledger_views(calc_old,existing+append)
    all_balances,all_dates,_=reconcile(o,p,n,cover,do,dn)
    diffs=[x for x in all_balances if x['Status']!='PASS']
    # Exactly the three UNKNOWN sourced receipts, not modelled in active AMB-10.
    if len(diffs)!=1 or (diffs[0]['NalogTip'],diffs[0]['NalogID'],diffs[0]['TipAmbalaze'],diffs[0]['Razlika'])!=('STANICA','ST-00001','0,4KG','-18'):
        raise ValueError('Unexpected residual mismatch: '+str(diffs[:8]))
    needed={(r['NalogTip'],r['NalogID'],r['TipAmbalaze']):r for r in all_balances}
    exp={'ST-00001':-16112,'ST-00002':-3577,'ST-00004':-4173,'ST-00005':-8093}
    for station,score in exp.items():
        val=Decimal(needed[('STANICA',station,'0,4KG')]['NoviIstorijskiSaldo'])
        if val!=score:raise ValueError(f'Unexpected station {station}: {val} != {score}')
    other=[d for d in all_dates if d['Status']!='PASS' and not (d['NalogTip']=='STANICA' and d['NalogID']=='ST-00001' and d['TipAmbalaze']=='0,4KG')]
    if other:raise ValueError('Unexpected daily mismatch outside unknown source ST1: '+str(other[:5]))
    return dict(saldo_pairs_compared=len(all_balances),saldo_mismatches=len(diffs),
         mismatch=diffs[0],dated_mismatch_count=sum(d['Status']!='PASS' for d in all_dates),
         internal_transfer_qty=633,unknown_receipt_qty=18,station_new_saldo=exp)


def get_table_location(z,ro,name):
    sheet,path,ref=ro.tables[name]
    # table relationship mapping, resolvable via workbook reader's tables list
    for entry in z.namelist():
        if not re.fullmatch(r'xl/tables/table\d+\.xml',entry):continue
        root=parse(z.read(entry))
        if root.get('name')==name or root.get('displayName')==name:
            return sheet,path,ref,entry
    raise ValueError('Cannot locate table OOXML for '+name)


def table_append(z,ro,name,records):
    sheet,sheetpath,ref,tablepath=get_table_location(z,ro,name)
    root=parse(z.read(tablepath)); sr=parse(z.read(sheetpath))
    assert re.fullmatch(r'[A-Z]+\d+:[A-Z]+\d+',ref),ref
    match=re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)',ref)
    colstart,rstart,colend,rend=match.groups();rstart=int(rstart);rend=int(rend)
    if colstart!='A' or rstart!=1:raise ValueError('Safe append requires table to start A1: '+name)
    names=[c.get('name') for c in root.findall('.//'+qn(M,'tableColumn'))]
    required=CANON_COLUMNS if name=='tblAmbalaza' else DOC_COLUMNS
    if names!=required:raise ValueError('Unexpected schema or column order for '+name+' '+repr(names))
    data=sr.find(qn(M,'sheetData'))
    if data is None:raise ValueError('Missing sheetData '+sheet)
    others=[r for r in data.findall(qn(M,'row')) if int(r.get('r','0'))>rend and r.findall(qn(M,'c'))]
    if others:raise ValueError('Unsafe: nonempty rows below table '+name+' '+str([r.get('r') for r in others[:5]]))
    rowmap={int(r.get('r')):r for r in data.findall(qn(M,'row'))}
    proto=rowmap.get(rend)
    if proto is None:raise ValueError('Last table row not found: '+name)
    styles={}
    for c in proto.findall(qn(M,'c')):
        col=re.match(r'([A-Z]+)',c.get('r','')).group(1)
        if col in [xl_col(i+1) for i in range(len(names))]:styles[col]=c.get('s')
    for j,values in enumerate(records,1):
        n=rend+j
        row=ET.Element(qn(M,'row'),{'r':str(n)})
        for i,namecol in enumerate(names,1):
            v=values.get(namecol,'')
            if v is None or str(v)=='':continue
            col=xl_col(i)
            attrs={'r':col+str(n)}
            if styles.get(col):attrs['s']=styles[col]
            is_num=(namecol=='Kolicina')
            if namecol=='Datum':
                v=excel_date(str(v));is_num=True
            if is_num:
                cell=ET.SubElement(row,qn(M,'c'),attrs)
                ET.SubElement(cell,qn(M,'v')).text=str(v)
            else:
                cell=ET.SubElement(row,qn(M,'c'),dict(attrs,t='inlineStr'))
                si=ET.SubElement(cell,qn(M,'is'))
                ET.SubElement(si,qn(M,'t')).text=str(v)
        data.append(row)
    newend=rend+len(records);end=colend+str(newend)
    root.set('ref',f'A1:{end}')
    auto=root.find(qn(M,'autoFilter'))
    if auto is not None:auto.set('ref','A1:'+end)
    for prop in ('dimension','autoFilter'):
        if prop=='dimension':node=sr.find(qn(M,'dimension'))
        else:node=sr.find(qn(M,'autoFilter'))
        if node is not None:
            prev=node.get('ref','')
            if prop=='dimension':
                mm=re.match(r'([A-Z]+)\d+:([A-Z]+)(\d+)',prev)
                if mm and int(mm.group(3))<newend:node.set('ref',mm.group(1)+'1:'+mm.group(2)+str(newend))
            elif prev==ref:node.set('ref','A1:'+end)
    return {sheetpath:xml(sr,z.read(sheetpath)),tablepath:xml(root,z.read(tablepath))}


def new_unknown_sheet(z,idx,existing_paths):
    # Adds visible archive worksheet INSIDE XLSM so six OMu historical cases
    # are together in workbook, without inserting noncanonical ledger kinds.
    wbpath='xl/workbook.xml';rels='xl/_rels/workbook.xml.rels';ct='[Content_Types].xml'
    wb=parse(z.read(wbpath));rootrel=parse(z.read(rels));types=parse(z.read(ct))
    sheets=wb.find(qn(M,'sheets'))
    names={x.get('name') for x in sheets}
    sheetname='MIG19_Nepoznati'
    if sheetname in names:raise ValueError('Already contains '+sheetname)
    nums=[int(re.search(r'sheet(\d+)\.xml',x).group(1)) for x in existing_paths if re.search(r'^xl/worksheets/sheet\d+\.xml$',x)]
    number=max(nums+[0])+1
    sheetpath=f'xl/worksheets/sheet{number}.xml'
    maxsheet=max([int(x.get('sheetId')) for x in sheets]+[0])+1
    ridset={x.get('Id') for x in rootrel}
    k=1
    while 'rId'+str(k) in ridset:k+=1
    rid='rId'+str(k)
    ET.SubElement(sheets,qn(M,'sheet'),{'name':sheetname,'sheetId':str(maxsheet),qn(R,'id'):rid})
    ET.SubElement(rootrel,qn(P,'Relationship'),{'Id':rid,'Type':R+'/worksheet','Target':f'worksheets/sheet{number}.xml'})
    ET.SubElement(types,qn(C,'Override'),{'PartName':'/'+sheetpath,'ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
    hdr=['IzvorniAmbID','Datum','TipAmbalaze','Kolicina','EntitetTip','EntitetID','DokumentTip','DokumentID','Smer','OdNalog','Status','Napomena']
    src=[]
    for aid in UNKNOWN:
        r=idx[aid]
        src.append([aid,date(r.get('Datum'),aid),r.get('TipAmbalaze'),str(number_orig(r)),
              'Stanica','ST-00001',r.get('DokumentTIP'),r.get('DokumentID'),
              r.get('Smer'),'NEPOZNAT','ISTORIJSKI_UNOS_VIDELJIV__AMB10_HOLD',
              'Bez izmisljanja NABAVKA / Firma / Vozac. Aktivni AMB10 ne podrzava izvor nepoznat.'])
    ns='<worksheet xmlns="'+M+'"><dimension ref="A1:L4"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="17"/><sheetData>'
    for rn,values in enumerate([hdr]+src,1):
        ns+=f'<row r="{rn}">'
        for col,v in enumerate(values,1):
            from xml.sax.saxutils import escape
            ns+=f'<c r="{xl_col(col)}{rn}" t="inlineStr"><is><t>{escape(str(v or ""))}</t></is></c>'
        ns+='</row>'
    ns+='</sheetData><autoFilter ref="A1:L4"/></worksheet>'
    return {wbpath:xml(wb,z.read(wbpath)),rels:xml(rootrel,z.read(rels)),ct:xml(types,z.read(ct)),sheetpath:ns.encode('utf-8')}

def number_orig(r):return number(r.get('Kolicina'),'original')


def make_xlsm(original,outpath,amb,docs,idx):
    if outpath.resolve()==original.resolve():raise ValueError('Output cannot overwrite input')
    with ZipFile(original) as z:
        if z.testzip() is not None:raise ValueError('Input XLSM zip corruption')
        if 'xl/vbaProject.bin' not in z.namelist():raise ValueError('Input is not a VBA-containing XLSM')
        vbabefore=z.read('xl/vbaProject.bin')
        ro=XlsmReadOnly(original)
        try:
            parts={}
            parts.update(table_append(z,ro,'tblAmbalaza',amb))
            parts.update(table_append(z,ro,'tblAmbalazaDokument',docs))
            parts.update(new_unknown_sheet(z,idx,set(z.namelist())))
        finally:ro.close()
        # atomic creation, never produce partial workbook as final path
        temp=outpath.with_suffix('.xlsm.incomplete')
        if temp.exists():temp.unlink()
        try:
            with ZipFile(temp,'w') as dst:
                for info in z.infolist():
                    # ZipFile.writestr mutates ZipInfo offsets; copy to preserve source reader.
                    dst.writestr(copy.copy(info),parts.pop(info.filename,z.read(info.filename)))
                for p,b in parts.items():dst.writestr(p,b)
            with ZipFile(temp) as chk:
                if chk.testzip() is not None:raise ValueError('Output zip corrupted')
                if chk.read('xl/vbaProject.bin')!=vbabefore:
                    raise ValueError('VBA bytes were changed')
            if outpath.exists():raise ValueError('Refusing overwrite of existing output '+str(outpath))
            temp.replace(outpath)
        finally:
            if temp.exists():temp.unlink()


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage',type=Path,default=Path('../MigrationStage_v2'))
    p.add_argument('--workbook',type=Path,default=Path('../MigrationGoods_MIG18_KARTICA_QA/AgriX_MIG18_KARTICA_QA.xlsm'))
    p.add_argument('--out-dir',type=Path,default=Path('../MigrationMIG19_13_PatchedXLSM'))
    p.add_argument('--apply',action='store_true',help='Write NEW XLSM copy, never original; otherwise preview only')
    args=p.parse_args(argv)
    if not args.workbook.is_file():
        raise FileNotFoundError('Existing XLSM not found: '+str(args.workbook)+'; inspect ../MigrationGoods_MIG*QA/ folders')
    source=args.stage/'UNCONVERTED__tblAmbalaza.jsonl'
    parent=args.stage/'SOURCE__tblOtpremnica.jsonl'
    if sha256(source)!=STAGE_SHA:raise ValueError('Legacy staging checksum mismatch')
    if sha256(parent)!=PARENT_SHA:raise ValueError('Original Otpremnica checksum mismatch')
    initial=sha256(args.workbook)
    if initial!=EXPECTED_TARGET_SHA:raise ValueError('QA XLSM checksum mismatch; expected '+EXPECTED_TARGET_SHA+' got '+initial)
    old=stage_rows(source); parents=stage_jsonl(parent)
    idx,parentproof=ensure_source(old,parents)
    book=XlsmReadOnly(args.workbook)
    try:
        new=book.read_table('tblAmbalaza',CANON_COLUMNS)
        docs=book.read_table('tblAmbalazaDokument',DOC_COLUMNS)
    finally:book.close()
    target_preflight(old,new,docs)
    amb,doc=planned_rows(idx)
    if any(key(x['AmbID']) in {key(r['AmbID']) for r in new} for x in amb):raise ValueError('Duplicate AmbID')
    if any(key(x['AmbDokID']) in {key(r['AmbDokID']) for r in docs} for x in doc):raise ValueError('Duplicate AmbDokID')
    verification=verify_projection(old,new,amb)
    args.out_dir.mkdir(parents=True,exist_ok=True)
    final_path=args.out_dir/'AgriX_MIG19_13_AMBALAZA_QA_PARTIAL.xlsm'
    result={'format':'AGRIx-MIG19.13-SURGICAL-EXISTING-XLSM-v1',
            'status':'DRY_RUN_OK' if not args.apply else 'WROTE_PARTIAL_AMB10_QA_WITH_3_UNKNOWN_IN_ARCHIVE',
            'original_qa_sha256':initial,'source_verified':True,'parent_storno_verified':True,
            'parent_storno':parentproof,'existing_canonical_amb_rows':len(new),
            'added_valid_internal_transfers':3,'preserved_unknown_receipts_in_workbook_sheet':3,
            'added_valid_revers_headers':3,'source_amended':False,
            'unknown_source_operational_ledger_status':'HOLD_REQUIRES_SUPPORTED_AMB10_MOVEMENT',
            'projection':verification,'original_sha_unchanged':True,'new_workbook':str(final_path) if args.apply else None}
    if args.apply:
        make_xlsm(args.workbook,final_path,amb,doc,idx)
        out=XlsmReadOnly(final_path)
        try:
            reread=out.read_table('tblAmbalaza',CANON_COLUMNS)
            rdocs=out.read_table('tblAmbalazaDokument',DOC_COLUMNS)
            if len(reread)!=len(new)+3 or len(rdocs)!=len(docs)+3:
                raise ValueError('Readback row counts mismatch')
            if any(key(x['AmbID']) not in {key(y['AmbID']) for y in reread} for x in amb):
                raise ValueError('Added IDs not found after readback')
            reread_projection=verify_projection(old,new,reread[len(new):])
            result['readback_validated']=True
            result['workbook_sha256']=sha256(final_path)
        finally:out.close()
    if sha256(args.workbook)!=initial:raise ValueError('Original workbook changed unexpectedly')
    fname='MIG1913_SUMMARY.json'
    (args.out_dir/fname).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:
        print('MIG19_13_ERROR:',type(e).__name__,str(e),file=sys.stderr)
        sys.exit(3)
