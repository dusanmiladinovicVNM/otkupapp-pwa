#!/usr/bin/env python3
"""AGRIx MIG-19 Malina: independent, strictly read-only QA audit.

Frozen reference: otkupapp-pwa 66f7897fa5231786f7d3d3dd3e566922c8e80ed0
Tables: tblAmbalaza, tblAmbalazaDokument (schema/schema.json at frozen ref).
Reads .xlsm OOXML with Python standard library; never opens/saves workbook for writing.
Does NOT reconstruct undocumented legacy links. Provenance needs --lineage-csv.
"""
from __future__ import annotations
import argparse
import csv
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {
    'm': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
QN = lambda p, name: '{' + NS[p] + '}' + name
CENTER = 'ST-00001'
REVERSES = {'REVERS', 'REVERS_PARTNERA'}
MATRIX = {
    'AMBALAZA_UZ_ROBU': ('REALAN','REALAN'),
    'IZDATA_PRAZNA': ('SOPSTVENI','PARTNER'),
    'POVRAT_PRAZNE': ('PARTNER','SOPSTVENI'),
    'PRENOS_INTERNO': ('SOPSTVENI','SOPSTVENI'),
    'ULAZ_TUDJE_AMBALAZE': ('GRANICA','PARTNER'),
    'VRACANJE_TUDJE_AMBALAZE': ('SOPSTVENI','PARTNER'),
    'NABAVKA': ('GRANICA','SOPSTVENI'),
    'OTPIS': ('REALAN','GRANICA'),
}
OWN={'STANICA','VOZAC','FIRMA'}
PARTNER={'KOOPERANT','KUPAC'}
ALLOWED={
    'REVERS':{'IZDATA_PRAZNA','POVRAT_PRAZNE','PRENOS_INTERNO','VRACANJE_TUDJE_AMBALAZE','ULAZ_TUDJE_AMBALAZE'},
    'REVERS_PARTNERA':{'POVRAT_PRAZNE','ULAZ_TUDJE_AMBALAZE'},
}
OMULAZ = [
    ('AMB-01377','2026-06-23',5,CENTER,None),
    ('AMB-02317','2026-06-27',300,'ST-00005',CENTER),
    ('AMB-03849','2026-06-28',250,'ST-00005',CENTER),
    ('AMB-03939','2026-06-29',83,'ST-00002',CENTER),
    ('AMB-07115','2026-07-14',10,CENTER,None),
    ('AMB-07815','2026-07-23',3,CENTER,None),
]


def norm(x):
    if x is None: return ''
    return str(x).strip()


def key(x): return norm(x).upper()


def dnum(value):
    s=norm(value).replace(' ','').replace('\u00a0','')
    if ',' in s and '.' not in s: s=s.replace(',','.')
    try: return Decimal(s)
    except (InvalidOperation, ValueError): return None


def ddate(value):
    s=norm(value)
    if not s: return ''
    # Numeric cells that contain Excel date serials; date part only.
    try:
        n=Decimal(s)
        if n>=Decimal('20000') and n<Decimal('100000'):
            return (datetime(1899,12,30)+timedelta(days=int(n))).date().isoformat()
    except InvalidOperation: pass
    for fmt in ('%Y-%m-%d','%d.%m.%Y','%d.%m.%Y.','%d/%m/%Y','%Y-%m-%dT%H:%M:%S'):
        try: return datetime.strptime(s,fmt).date().isoformat()
        except ValueError: pass
    return s


def sha256(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()


def zip_path(base,target):
    # OOXML relationship targets are relative to their parent-part folder.
    target=target.lstrip('/')
    if target.startswith('xl/'):return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def cell_col(ref):
    letters=re.match(r'^([A-Z]+)', ref)
    if not letters:raise ValueError('Unexpected cell address '+ref)
    col=0
    for ch in letters.group(1):col=col*26+ord(ch)-64
    return col


def address_row(ref):
    mt=re.match(r'^[A-Z]+([0-9]+)$',ref)
    if not mt:raise ValueError('Unexpected row address '+ref)
    return int(mt.group(1))


class XlsmReadOnly:
    def __init__(self,path):
        self.path=Path(path)
        self.z=ZipFile(self.path,'r')
        self.sst=[]
        if 'xl/sharedStrings.xml' in self.z.namelist():
            root=ET.fromstring(self.z.read('xl/sharedStrings.xml'))
            for si in root.findall('m:si',NS):
                self.sst.append(''.join(t.text or '' for t in si.iter(QN('m','t'))))
        wb=ET.fromstring(self.z.read('xl/workbook.xml'))
        relroot=ET.fromstring(self.z.read('xl/_rels/workbook.xml.rels'))
        relmap={x.get('Id'):zip_path('xl/workbook.xml',x.get('Target'))
                for x in relroot.findall('p:Relationship',NS)
                if not x.get('TargetMode')=='External'}
        self.sheets={x.get('name'):relmap[x.get(QN('r','id'))]
                for x in wb.findall('m:sheets/m:sheet',NS)}
        self.tables={}
        for sheet,path in self.sheets.items():
            relfile=posixpath.join(posixpath.dirname(path),'_rels',posixpath.basename(path)+'.rels')
            if relfile not in self.z.namelist():continue
            rels=ET.fromstring(self.z.read(relfile))
            for x in rels.findall('p:Relationship',NS):
                if not x.get('Type','').endswith('/table'):continue
                tablepath=zip_path(path,x.get('Target'))
                meta=ET.fromstring(self.z.read(tablepath))
                name=meta.get('name') or meta.get('displayName')
                self.tables[name]=(sheet,path,meta.get('ref'))

    def read_table(self,name,required):
        if name not in self.tables:
            raise ValueError(f'Missing Excel ListObject {name} (available: {sorted(self.tables)})')
        sheet,path,ref=self.tables[name]
        mt=re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$',ref or '')
        if not mt:raise ValueError(f'Unexpected table range for {name}: {ref}')
        c1,c2=cell_col(mt.group(1)),cell_col(mt.group(3))
        r1,r2=int(mt.group(2)),int(mt.group(4))
        tree=ET.fromstring(self.z.read(path))
        out=[]; headers=None; formula_no_cache=[]
        for row in tree.findall('m:sheetData/m:row',NS):
            rn=int(row.get('r','0'))
            if rn<r1 or rn>r2:continue
            fields=['']*(c2-c1+1)
            for c in row.findall('m:c',NS):
                cidx=cell_col(c.get('r',''))
                if not c1<=cidx<=c2:continue
                typ=c.get('t','')
                v=c.find('m:v',NS)
                if typ=='inlineStr':
                    ins=c.find('m:is',NS)
                    val=''.join(x.text or '' for x in ins.iter(QN('m','t'))) if ins is not None else ''
                elif v is None:
                    val=''
                    if c.find('m:f',NS) is not None:formula_no_cache.append(c.get('r',''))
                elif typ=='s':
                    idx=int(v.text or '0')
                    if idx<0 or idx>=len(self.sst):raise ValueError(f'Bad sharedString at {c.get("r")}')
                    val=self.sst[idx]
                elif typ=='b':val='True' if v.text=='1' else 'False'
                else:val=v.text or ''
                fields[cidx-c1]=val
            if rn==r1: headers=[norm(x) for x in fields];continue
            if any(norm(x) for x in fields):
                if headers is None:raise ValueError(f'Header missing in {name}')
                out.append({'_excel_row':rn, **{headers[i]:fields[i] for i in range(len(headers))}})
        if headers is None:raise ValueError(f'No table header in {name}')
        missing=sorted(set(required)-set(headers))
        if missing:raise ValueError(f'Wrong schema in {name}: missing {missing}; found {headers}')
        if formula_no_cache:raise ValueError(f'Uncached formula cells in {name}: {formula_no_cache[:8]}')
        return out

    def close(self):self.z.close()


def report(row, severity, code, why, findings, amb=''):
    findings.append({'severity':severity,'code':code,
        'AmbDokID':norm(row.get('AmbDokID') or row.get('DokumentID')),
        'BrojDokumenta':norm(row.get('BrojDokumenta')),
        'AmbID':norm(row.get('AmbID')) or amb,
        'Datum':ddate(row.get('Datum')),
        'TipAmbalaze':norm(row.get('TipAmbalaze')),
        'Kolicina':norm(row.get('Kolicina')),
        'OdNalogTip':norm(row.get('OdNalogTip')),
        'OdNalogID':norm(row.get('OdNalogID')),
        'NaNalogTip':norm(row.get('NaNalogTip')),
        'NaNalogID':norm(row.get('NaNalogID')),
        'reason':why})


def account_class(t):
    if t=='SPOLJNISVET':return {'GRANICA'}
    if t in OWN:return {'SOPSTVENI','REALAN'}
    if t in PARTNER:return {'PARTNER','REALAN'}
    return set()


def book_row_check(row,findings):
    a,b=key(row['OdNalogTip']),key(row['NaNalogTip'])
    aid,bid=key(row['OdNalogID']),key(row['NaNalogID'])
    v=key(row['VrstaKretanja'])
    q=dnum(row['Kolicina'])
    if q is None or q<=0:report(row,'P1','QUANTITY_INVALID','Kolicina mora biti pozitivan broj.',findings)
    if not key(row['TipAmbalaze']):report(row,'P1','TYPE_EMPTY','Nedostaje tip ambalaze.',findings)
    if not norm(row['DokumentID']) or not norm(row['DokumentTIP']):
        report(row,'P1','SOURCE_DOCUMENT_EMPTY','Nedostaje identitet izvornog dokumenta.',findings)
    if a not in OWN|PARTNER|{'SPOLJNISVET'} or b not in OWN|PARTNER|{'SPOLJNISVET'}:
        report(row,'P1','UNKNOWN_ACCOUNT','Nepoznata vrsta naloga.',findings)
    for t,x,side in ((a,aid,'Od'),(b,bid,'Na')):
        if t in {'FIRMA','SPOLJNISVET'} and x:
            report(row,'P1','SYSTEM_ID','Sistemski nalog mora imati prazan ID: '+side,findings)
        if t not in {'FIRMA','SPOLJNISVET'} and not x:
            report(row,'P1','ACCOUNT_ID_EMPTY','Nesistemski nalog mora imati ID: '+side,findings)
    if a==b and aid==bid:
        report(row,'P1','SELF_TRANSFER','Kanonski AMB-INV-02 zabranjuje isti nalog na obe strane.',findings)
    if {a,b}=={'STANICA','VOZAC'} and aid and aid==bid:
        report(row,'P2','MIRROR_PSEUDO_TRANSFER','Malina mirror Stanica/Vozac istog ID: proveriti da ovo nije lazni fizicki transfer.',findings)
    if v not in MATRIX:
        report(row,'P1','UNKNOWN_MOVEMENT','VrstaKretanja nije u kanonskoj AMB-10 matrici.',findings)
    else:
        ca,cb=MATRIX[v]
        # A storno keeps original kind but swaps physical endpoints.
        # The frozen VBA writer explicitly validates storno in reversed order.
        if norm(row.get('StornoOd')): ca,cb=cb,ca
        if ca not in account_class(a) or cb not in account_class(b):
            report(row,'P1','CLASS_MISMATCH',f'{v} zahteva {ca} -> {cb} za ovaj red; trenutno {a}->{b}.',findings)


def lineage(path):
    if path is None:return {}
    out={}
    with Path(path).open('r',encoding='utf-8-sig',newline='') as f:
        for r in csv.DictReader(f):
            legacy=key(r.get('LegacyAmbID'))
            current=key(r.get('NewAmbID'))
            if not legacy or not current:raise ValueError('Lineage CSV requires non-empty LegacyAmbID,NewAmbID')
            if legacy in out:raise ValueError('Duplicate legacy lineage ID '+legacy)
            out[legacy]=current
    # One canonical physical event cannot be presented as two independent OMu legacy movements.
    omu_links=[out[x] for x in out if x in {v[0] for v in OMULAZ}]
    if len(omu_links)!=len(set(omu_links)):
        raise ValueError('Two distinct OMuUlaz source events point to the same canonical AmbID')
    return out


def write_csv(path,rows,cols):
    with Path(path).open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=cols,extrasaction='ignore')
        w.writeheader()
        for row in rows:
            # protect CSV opening in Excel from formula injection
            safe={k:("'"+str(row.get(k,'')) if isinstance(row.get(k,''),str) and row.get(k,'').startswith(('=','+','-','@')) else row.get(k,'')) for k in cols}
            w.writerow(safe)


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--workbook',required=True,type=Path,help='Existing MIG-18.x/19 QA .xlsm; READ ONLY')
    p.add_argument('--stage',type=Path,help='Legacy staging path, referenced only in metadata (not parsed)')
    p.add_argument('--out-dir',required=True,type=Path)
    p.add_argument('--lineage-csv',type=Path,help='Optional verified source->canonical linkage: LegacyAmbID,NewAmbID')
    p.add_argument('--expected-revers',type=int,default=48)
    p.add_argument('--expected-ledger',type=int,default=4811)
    args=p.parse_args(argv)
    if not args.workbook.is_file():p.error('Workbook does not exist: '+str(args.workbook))
    if args.workbook.suffix.lower() not in ('.xlsm','.xlsx'):
        p.error('Workbook must be .xlsm or .xlsx')
    if args.out_dir.resolve()==args.workbook.resolve().parent:
        p.error('Use a separate output directory to keep the workbook location untouched.')
    checksum_before=sha256(args.workbook)
    needamb={'AmbID','Datum','TipAmbalaze','Kolicina','DokumentID','DokumentTIP','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd'}
    needhead={'AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Stornirano'}
    reader=XlsmReadOnly(args.workbook)
    try:
        book=reader.read_table('tblAmbalaza',needamb)
        docs=reader.read_table('tblAmbalazaDokument',needhead)
    finally:reader.close()
    checks=[]; findings=[]
    if len(book)!=args.expected_ledger:
        checks.append(f'expected {args.expected_ledger} tblAmbalaza rows, found {len(book)}')
    docids=defaultdict(list)
    numbers=defaultdict(list)
    for d in docs:
        docids[key(d['AmbDokID'])].append(d)
        nums=(key(d['BrojOwnerTip']),key(d['BrojOwnerID']),ddate(d['Datum']),key(d['BrojDokumenta']))
        if all(nums):numbers[nums].append(d)
    for k,ds in docids.items():
        if not k:report(ds[0],'P1','AMBDOK_ID_EMPTY','AmbDokID cannot be empty.',findings)
        elif len(ds)>1:report(ds[0],'P1','AMB_DOC_ID_DUPLICATE',f'{k} appears {len(ds)} times.',findings)
    for k,ds in numbers.items():
        if len(ds)>1:report(ds[0],'P1','AMB_DOC_NUMBER_DUPLICATE',f'{len(ds)} document headers share owner/date/number.',findings)
    rev=[x for x in docs if key(x['Vrsta']) in REVERSES]
    if len(rev)!=args.expected_revers:
        checks.append(f'expected {args.expected_revers} revers headers, found {len(rev)}')
    idmap=defaultdict(list)
    bydocument=defaultdict(list)
    normalized_rows={}
    ledger_saldo=defaultdict(Decimal)
    for r in book:
        aid=key(r['AmbID'])
        if not aid:report(r,'P1','AMB_ID_MISSING','AmbID is empty.',findings)
        else:idmap[aid].append(r)
        doc_key=(key(r['DokumentTIP']),key(r['DokumentID']))
        bydocument[doc_key].append(r)
        book_row_check(r,findings)
        q=dnum(r['Kolicina'])
        if q and q>0:
            tp=key(r['TipAmbalaze'])
            ledger_saldo[(key(r['OdNalogTip']),key(r['OdNalogID']),tp)]-=q
            ledger_saldo[(key(r['NaNalogTip']),key(r['NaNalogID']),tp)]+=q
        if aid:normalized_rows[aid]=r
    for aid,rows in idmap.items():
        if len(rows)>1:report(rows[0],'P1','DUPLICATE_AMBID',f'AmbID {aid}: {len(rows)} rows.',findings)
    for (dtype,did),entries in bydocument.items():
        if dtype=='AMBALAZADOKUMENT' and (not did or did not in docids):
            report(entries[0],'P1','ORPHAN_AMB_DOC_LEDGER',f'Ledger references missing tblAmbalazaDokument heading {did}.',findings)
    for r in book:
        st=key(r['StornoOd'])
        if st:
            if st not in idmap:
                report(r,'P1','ORPHAN_STORNO',f'StornoOd {st} missing.',findings)
            elif len(idmap[st])==1:
                o=idmap[st][0]
                if (key(r['OdNalogTip']),key(r['OdNalogID']),key(r['NaNalogTip']),key(r['NaNalogID'])) != (key(o['NaNalogTip']),key(o['NaNalogID']),key(o['OdNalogTip']),key(o['OdNalogID'])):
                    report(r,'P1','STORNO_DIRECTION','Storno must reverse original source and destination.',findings)
                if (key(r['TipAmbalaze']),key(r['VrstaKretanja']),dnum(r['Kolicina']))!=(key(o['TipAmbalaze']),key(o['VrstaKretanja']),dnum(o['Kolicina'])):
                    report(r,'P1','STORNO_CONTENT','Storno must match original type, movement and quantity.',findings)
    storno_by_original=defaultdict(list)
    for r in book:
        if key(r['StornoOd']):storno_by_original[key(r['StornoOd'])].append(r)
    for original,rows in storno_by_original.items():
        if len(rows)>1:report(rows[0],'P1','STORNO_DUPLICATE',f'{original} has {len(rows)} reversals.',findings)
    revers_info=[]
    revers_ledger=[]
    for h in rev:
        docid=key(h['AmbDokID'])
        lines=bydocument.get(('AMBALAZADOKUMENT',docid),[])
        original=[r for r in lines if not norm(r['StornoOd'])]
        business=[r for r in original if key(r['VrstaKretanja'])!='ULAZ_TUDJE_AMBALAZE']
        coverage=[r for r in original if key(r['VrstaKretanja'])=='ULAZ_TUDJE_AMBALAZE']
        storned=key(h['Stornirano']) in ('DA','TRUE','1')
        for r in lines:
            revers_ledger.append({
                'AmbDokID':h['AmbDokID'],'BrojDokumenta':h['BrojDokumenta'],'VrstaDokumenta':h['Vrsta'],
                'AmbID':r['AmbID'],'Datum':ddate(r['Datum']),'TipAmbalaze':r['TipAmbalaze'],
                'Kolicina':r['Kolicina'],'OdNalogTip':r['OdNalogTip'],'OdNalogID':r['OdNalogID'],
                'NaNalogTip':r['NaNalogTip'],'NaNalogID':r['NaNalogID'],
                'VrstaKretanja':r['VrstaKretanja'],'StornoOd':r['StornoOd'],
                'role':('STORNO' if key(r['StornoOd']) else
                        'COVERAGE' if key(r['VrstaKretanja'])=='ULAZ_TUDJE_AMBALAZE' else 'BUSINESS')})
        if not lines:
            report(h,'P1' if not storned else 'P2','REVERS_NO_LEDGER','Revers header has no ledger rows.',findings)
        if not business:
            report(h,'P1' if not storned else 'P2','REVERS_NO_BUSINESS_ROW','Revers header has no original physical movement.',findings)
        if len(business)>0:
            pairs={frozenset(((key(x['OdNalogTip']),key(x['OdNalogID'])),(key(x['NaNalogTip']),key(x['NaNalogID'])))) for x in business}
            if len(pairs)!=1:
                report(h,'P1','REVERS_MULTIPLE_PAIRS',f'AMB-INV-10: {len(pairs)} different partner pairs.',findings)
        owner_type=key(h['BrojOwnerTip']);owner_id=key(h['BrojOwnerID'])
        if not owner_id:report(h,'P1','REVERS_OWNER_MISSING','BrojOwnerID missing.',findings)
        if key(h['Vrsta'])=='REVERS' and owner_type not in OWN:
            report(h,'P1','REVERS_OWNER_WRONG','REVERS requires SOPSTVENI BrojOwnerTip.',findings)
        if key(h['Vrsta'])=='REVERS_PARTNERA' and owner_type not in PARTNER:
            report(h,'P1','PARTNER_OWNER_WRONG','REVERS_PARTNERA requires PARTNER owner.',findings)
        if len({key(x['AmbID']) for x in lines})!=len(lines):
            report(h,'P1','REVERS_DUPLICATE_ROW','Duplicate AmbID within document.',findings)
        effects=defaultdict(list)
        for r in original:
            if ddate(r['Datum'])!=ddate(h['Datum']):
                report(r,'P2','REVERS_DATE_MISMATCH',f'Ledger date differs from header date {ddate(h["Datum"])}.',findings)
            v=key(r['VrstaKretanja']);fromt=key(r['OdNalogTip']);tot=key(r['NaNalogTip'])
            fromid=key(r['OdNalogID']);toid=key(r['NaNalogID']);typ=key(r['TipAmbalaze'])
            effects[(v,typ)].append(r)
            if v not in ALLOWED[key(h['Vrsta'])]:
                report(r,'P1','REVERS_MOVEMENT_FORBIDDEN',f'{h["Vrsta"]} does not allow {v}.',findings)
            if v=='IZDATA_PRAZNA':
                if (fromt,tot)!=('STANICA','KOOPERANT'):
                    report(r,'P2','MALINA_ISSUE_PARTIES','Expected Stanica -> Kooperant for Malina historical issue.',findings)
                elif owner_type=='STANICA' and owner_id!=fromid:
                    report(r,'P2','ISSUE_WRONG_STATION','Station source differs from revers number owner; needs source verification.',findings)
            elif v=='POVRAT_PRAZNE':
                if key(h['Vrsta'])=='REVERS_PARTNERA':
                    if not (fromt=='KUPAC' and tot=='VOZAC' and fromid==owner_id):
                        report(r,'P1','PARTNER_REVERS_PROBLEM','Partner revers must be Kupac -> Vozac with matching Kupac owner (snapshot ODL-9/10).',findings)
                elif (fromt,tot)!=('KOOPERANT','STANICA'):
                    report(r,'P2','MALINA_RETURN_PARTIES','Expected Kooperant -> Stanica for historical return.',findings)
                elif owner_type=='STANICA' and owner_id!=toid:
                    report(r,'P2','RETURN_WRONG_STATION','Return destination differs from revers number owner; needs verification.',findings)
            elif v=='PRENOS_INTERNO':
                if tot in ('STANICA','VOZAC') and toid!=CENTER:
                    if (fromt,fromid,tot)!=('STANICA',CENTER,'STANICA'):
                        report(r,'P2','MALINA_SUPPLY_SOURCE',f'Supplying {tot} {toid}: expected Stanica {CENTER} -> Stanica {toid}.',findings)
                elif toid==CENTER:
                    report(r,'P3','TRANSFER_TO_COLD_STORAGE','Revers arrives at Hladnjaca; verify original source, do not auto-reverse.',findings)
        for (v,typ),rows in effects.items():
            if len(rows)>1:
                report(rows[0],'P1','AMB_INV_04_DUPLICATE',f'{len(rows)} original effects share (document, kind={v}, type={typ}).',findings)
        members={((key(x['OdNalogTip']),key(x['OdNalogID']))) for x in business} | {((key(x['NaNalogTip']),key(x['NaNalogID']))) for x in business}
        for x in coverage:
            if (key(x['NaNalogTip']),key(x['NaNalogID'])) not in members:
                report(x,'P1','COVERAGE_FOREIGN_PARTNER','Coverage targets account outside its document partner pair.',findings)
        revers_info.append({'AmbDokID':h['AmbDokID'],'Vrsta':h['Vrsta'], 'BrojDokumenta':h['BrojDokumenta'],
            'Datum':ddate(h['Datum']),'BrojOwnerTip':h['BrojOwnerTip'],'BrojOwnerID':h['BrojOwnerID'],
            'Stornirano':h['Stornirano'],'all_rows':len(lines),'business_rows':len(business),
            'coverage_rows':len(coverage),'storno_rows':len(lines)-len(original),
            'kinds':'|'.join(sorted({key(x['VrstaKretanja']) for x in business})),
            'movement_ids':'|'.join(key(x['AmbID']) for x in business),
            'source_provenance':'INTERNAL_ONLY_UNTIL_LEGACY_LINKED'})
    # Critical: explicit lineage is the only way to claim an exact source->canonical match.
    link=lineage(args.lineage_csv)
    omulaz=[]
    for legacy,datum,qty,dest,src in OMULAZ:
        candidates=[]
        for r in book:
            if (ddate(r['Datum'])==datum and key(r['NaNalogTip'])=='STANICA' and key(r['NaNalogID'])==dest
                and dnum(r['Kolicina'])==Decimal(qty) and key(r['TipAmbalaze']).replace(' ','').replace('.',',') in ('0,4KG','0,4KG.')
                and key(r['VrstaKretanja']) not in ('ULAZ_TUDJE_AMBALAZE',)):
                candidates.append(r)
        actual=normalized_rows.get(link.get(legacy,'')) if legacy in link else None
        status='UNLINKED_REVIEW'
        if src is None:
            status='HISTORICAL_INTAKE_UNRESOLVED'
            if actual:status='LINEAGE_LINKED_CHECK_ORIGIN'
        elif actual:
            if (key(actual['OdNalogTip']),key(actual['OdNalogID']),key(actual['NaNalogTip']),key(actual['NaNalogID']),key(actual['VrstaKretanja']),dnum(actual['Kolicina']),ddate(actual['Datum'])) == ('STANICA',src,'STANICA',dest,'PRENOS_INTERNO',Decimal(qty),datum) and key(actual['TipAmbalaze']).replace(' ','').replace('.',',')=='0,4KG' and not key(actual['StornoOd']):
                status='LINEAGE_VERIFIED'
            else:
                status='LINEAGE_CONTRADICTS_POLICY'
                report(actual,'P1','OMULAZ_LINEAGE_MISMATCH',f'{legacy} expected {src}->{dest} {qty} 0,4kg on {datum}.',findings)
        elif legacy in link:
            status='LINEAGE_NEW_ID_NOT_FOUND'
            checks.append(f'OMUlaz lineage points to missing AmbID: {legacy} -> {link[legacy]}')
        omulaz.append({'LegacyAmbID':legacy,'Datum':datum,'TipAmbalaze':'0,4kg','Kolicina':qty,'ExpectedOd':src or 'NEUTVRDJENO','ExpectedNa':dest,
            'lineage_AmbID':link.get(legacy,''),'possible_matching_AmbIDs':'|'.join(key(x['AmbID']) for x in candidates),
            'candidate_count':len(candidates),'status':status})
    negatives=[]
    for (t,i,tp),bal in sorted(ledger_saldo.items()):
        if bal<0 and t!='SPOLJNISVET':
            allowed=(t=='STANICA' and i==CENTER)
            negatives.append({'NalogTip':t,'NalogID':i,'TipAmbalaze':tp,'Saldo':str(bal),
                'allowed_by_mig19_policy':'YES' if allowed else 'NO_REVIEW'})
            if not allowed:
                checks.append(f'negative saldo outside historical Hladnjaca exception: {t}/{i}/{tp}={bal}')
    args.out_dir.mkdir(parents=True,exist_ok=True)
    write_csv(args.out_dir/'all_saldos.csv',
              [{'NalogTip':t,'NalogID':i,'TipAmbalaze':tp,'Saldo':str(v),
                'historical_negative_exception':'YES' if t=='STANICA' and i==CENTER else 'NO'}
               for (t,i,tp),v in sorted(ledger_saldo.items()) if t!='SPOLJNISVET'],
              ['NalogTip','NalogID','TipAmbalaze','Saldo','historical_negative_exception'])
    write_csv(args.out_dir/'revers_ledger_rows.csv',revers_ledger,
              ['AmbDokID','BrojDokumenta','VrstaDokumenta','AmbID','Datum','TipAmbalaze','Kolicina',
               'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','StornoOd','role'])
    write_csv(args.out_dir/'revers_headers.csv',revers_info,['AmbDokID','Vrsta','BrojDokumenta','Datum','BrojOwnerTip','BrojOwnerID','Stornirano','all_rows','business_rows','coverage_rows','storno_rows','kinds','movement_ids','source_provenance'])
    cols=['severity','code','AmbDokID','BrojDokumenta','AmbID','Datum','TipAmbalaze','Kolicina','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','reason']
    write_csv(args.out_dir/'findings.csv',findings,cols)
    write_csv(args.out_dir/'omulaz6_review.csv',omulaz,['LegacyAmbID','Datum','TipAmbalaze','Kolicina','ExpectedOd','ExpectedNa','lineage_AmbID','possible_matching_AmbIDs','candidate_count','status'])
    write_csv(args.out_dir/'negative_saldos.csv',negatives,['NalogTip','NalogID','TipAmbalaze','Saldo','allowed_by_mig19_policy'])
    hash_after=sha256(args.workbook)
    if checksum_before!=hash_after:checks.append('CRITICAL: source workbook checksum changed during audit')
    severity=Counter(x['severity'] for x in findings)
    gate='HOLD'
    if findings or checks: gate='FAIL' if any(x['severity']=='P1' for x in findings) else 'HOLD'
    # This script cannot conclude absolute PASS without original evidence lineage.
    result={
        'gate':gate,'gate_scope':'local-read-only-policy-audit; NOT Excel-runtime PASS or legacy-source reconciled',
        'workbook':str(args.workbook),'sha256_before':checksum_before,'sha256_after':hash_after,
        'source_unchanged':checksum_before==hash_after,
        'reference_sha':'66f7897fa5231786f7d3d3dd3e566922c8e80ed0',
        'counts':{'ledger_rows':len(book),'amb_dok_headers':len(docs),'revers_headers':len(rev),
                  'revers_ledger_rows':sum(x['all_rows'] for x in revers_info),
                  'revers_business_rows':sum(x['business_rows'] for x in revers_info),
                  'revers_coverage_rows':sum(x['coverage_rows'] for x in revers_info),
                  'revers_storno_rows':sum(x['storno_rows'] for x in revers_info),
                  'findings':len(findings),'finding_severity':dict(severity),
                  'negative_saldo_accounts':len(negatives),
                  'omulaz_verified_with_lineage':sum(x['status']=='LINEAGE_VERIFIED' for x in omulaz)},
        'reconciliation_checks':checks,
        'unproven':['Historical reversal provenance for all 48 headers without external legacy mapping',
                    'OMUlaz origin for 3 intake rows at ST-00001','signed versus physical source balances',
                    'Excel runtime report/UI behavior'],
        'stage_path_if_given':str(args.stage) if args.stage else None,
        'note':'--stage is informational: opaque staging formats are NOT guessed. Supply verified --lineage-csv for OMu identity mapping.'
    }
    (args.out_dir/'summary.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'gate':gate,'counts':result['counts'],'checks':checks[:12],'out_dir':str(args.out_dir)},ensure_ascii=False,indent=2))
    return 2 if gate=='FAIL' else 1  # HOLD is nonzero for CI / command chains.

if __name__=='__main__':
    try: sys.exit(main())
    except Exception as exc:
        print(f'MIG19_AUDIT_ERROR: {type(exc).__name__}: {exc}',file=sys.stderr)
        sys.exit(3)
