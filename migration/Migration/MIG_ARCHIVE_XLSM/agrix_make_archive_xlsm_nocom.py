#!/usr/bin/env python3
"""AGRIx MIG-ARCHIVE v1.3: macro-preserving read-only archival sheets, NO Excel COM.

Only changes three package metadata entries and adds six new worksheet XML parts
in a NEW XLSM. Every original worksheet, table, VBA part, external link and image
part is copied UNCHANGED. This is a Windows-independent packaging operation.
The result should be opened with macros DISABLED / read-only for archive viewing.
"""
from __future__ import annotations
import argparse,csv,hashlib,json,os,re,sys,zipfile
from datetime import datetime
from pathlib import Path
from collections import Counter
from xml.sax.saxutils import escape, quoteattr
from xml.etree import ElementTree as ET

from agrix_mig_archive import read_workbook, export_otkups, extract_issues
from agrix_mig19_malina_readonly_audit import sha256
from archive_tables import build_csvs,project_bytes,workbook_signature

V='AGRIx-MIG-ARCHIVE-XLSM-v1.3-NOCOM'
MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RREL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKGREL='http://schemas.openxmlformats.org/package/2006/relationships'
CTNS='http://schemas.openxmlformats.org/package/2006/content-types'
ARCHIVE=['ARHIVA_START','ARHIVA_OM','ARHIVA_OTKUPI','ARHIVA_KOOPERANTI','ARHIVA_IZUZECI','ARHIVA_SIROVI']
NUMERIC={'Aktivnih otkupa','Kilograma (poznato)','Vrednost RSD (poznato)',
         'Nepotpunih otkupa','Excel red','Kg','Vrednost RSD'}
BAD_SHA={'c5bc94125ef706b6a04a0bfae06a4a03a299897dbdaa465e8e305eab08fbfe5d'}

# Preserve all original XML namespaces, tags and other workbook metadata by
# replacing exactly one closing XML tag, rather than serializing whole workbooks.
def insert_before_close(src:bytes, tag:str, payload:str)->bytes:
    s=src.decode('utf-8-sig')
    rx=re.compile(r'</(?:[A-Za-z_][\w.-]*:)?'+re.escape(tag)+r'\s*>')
    matches=list(rx.finditer(s))
    if len(matches)!=1:
        raise RuntimeError('OOXML metadata missing/ambiguous closing '+tag)
    at=matches[0].start()
    return (s[:at]+payload+s[at:]).encode('utf-8-sig' if src.startswith(b'\xef\xbb\xbf') else 'utf-8')

def excel_col(n:int)->str:
    x=''
    while n:
        n,k=divmod(n-1,26)
        x=chr(65+k)+x
    return x

def textcell(address,text):
    # Inline strings have NO formula semantics. Escaped XML, preserve whitespace.
    if text is None: text=''
    t=escape(str(text).replace('\x00',''))
    # XML 1.0 disallows control chars even when escaped.
    t=''.join(c for c in t if c in '\t\n\r' or ord(c)>=32)
    return f'<c r="{address}" t="inlineStr"><is><t xml:space="preserve">{t}</t></is></c>'

def numcell(address,text):
    from decimal import Decimal, InvalidOperation
    try:
        v=Decimal(str(text).strip())
        if not v.is_finite(): return textcell(address,text)
        return f'<c r="{address}"><v>{str(v)}</v></c>'
    except (ValueError,InvalidOperation): return textcell(address,text)

def sheet_xml(header:list[str],rows:list[list[str]],start:bool=False)->bytes:
    nc=max(len(header),max((len(x) for x in rows),default=0))
    nr=1+len(rows)
    # Cap to legal Excel size and preserve original CSV values exactly as text.
    if nr>1048576 or nc>16384: raise RuntimeError('Archive sheet exceeds Excel worksheet limits')
    chunks=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<worksheet xmlns="{MAIN}">',
            '<sheetViews><sheetView workbookViewId="0" showGridLines="0">',
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>',
            '</sheetView></sheetViews>',
            '<sheetFormatPr defaultRowHeight="17"/>', '<cols>']
    for col in range(1,nc+1):
        w=60 if (header[col-1] if col<=len(header) else '') in ('Opis','Originalni red (JSON)','Uticaj na pregled') else (25 if col==1 else 21)
        if start: w=96 if col==2 else 14
        chunks.append(f'<col min="{col}" max="{col}" width="{w}" customWidth="1"/>')
    chunks.append('</cols><sheetData>')
    for index, values in enumerate([header]+rows,1):
        chunks.append(f'<row r="{index}">')
        for col,value in enumerate(values,1):
            addr=f'{excel_col(col)}{index}'
            label=header[col-1] if col<=len(header) else ''
            # Keep all identifier columns and textual dates as strings.
            if not start and index>1 and label in NUMERIC:
                chunks.append(numcell(addr,value))
            else:
                chunks.append(textcell(addr,value))
        chunks.append('</row>')
    chunks.append('</sheetData>')
    if not start and rows:
        chunks.append(f'<autoFilter ref="A1:{excel_col(nc)}{nr}"/>')
    # Basic protection is accidental edit prevention, NOT security. Only the
    # separately provided safe launcher disables macros and enforces ReadOnly.
    chunks.append('<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0" sort="0"/>')
    chunks.append('</worksheet>')
    xml=''.join(chunks).encode('utf-8')
    ET.fromstring(xml)
    return xml

def csv_rows(p:Path):
    with p.open(encoding='utf-8-sig',newline='') as f:
        r=csv.reader(f,delimiter=';')
        h=next(r)
        return h,list(r)

def xml_start():
    h=['','ARHIVA 2026 / READ-ONLY PREGLED']
    entries=[
      ['', 'AgriX - istorijski poslovni podaci iz migrirane XLSM sveske'],
      ['', 'ARHIVA_OM: zbirni pregled po otkupnim mestima'],
      ['', 'ARHIVA_OTKUPI: pregled svih otkupa; filtiranje po OM/kooperantu'],
      ['', 'ARHIVA_KOOPERANTI: zbirni podaci po kooperantu i OM'],
      ['', 'ARHIVA_IZUZECI: problemi dokumentnih veza (bez korekcija)'],
      ['', 'ARHIVA_SIROVI: izvorni problematični redovi za proveru'],
      ['', 'Brojke iz OTK i OTK stavki. Izuzeci nisu uklonjeni iz istorije.'],
      ['', 'Ovo nije finansijska kartica; isplate nisu uključene u ove zbirove.'],
      ['', 'Za otvaranje koristi OPEN_ARCHIVE_SAFE.ps1: ReadOnly + macros OFF.'],
      ['', 'Ne uključivati makroe ni upis u arhivsku kopiju.'],
    ]
    return h,entries

def assert_clean(path:Path,T,accept_duplicates:bool=False)->list:
    required={'tblOtkup':1702,'tblOtkupStavke':1702,'tblOtpremnica':907,
              'tblZbirna':907,'tblPrijemnica':893}
    problems=[]
    for name,n in required.items():
        actual=len(T.get(name,[]))
        if actual!=n:problems.append(f'{name}: {actual} redova; ocekivano {n}')
    zbr=[str(r.get('ZbirnaID','')).strip() for r in T['tblZbirna']]
    dups=[(x,n) for x,n in Counter(zbr).items() if n>1 or not x]
    if dups: problems.append(f'dupli/prazni ZbirnaID: {dups[:15]}')
    if sha256(path) in BAD_SHA: problems.append('poznata rucno ostecena MIG19.13 kopija (SHA)')
    if problems and not accept_duplicates:
        raise RuntimeError('Sveska nije cista arhivska baza: '+'; '.join(problems))
    return problems

def build_without_excel(source:Path,target:Path,output:Path)->list[str]:
    replacement={}
    with zipfile.ZipFile(source) as zin:
        srcnames=zin.namelist()
        if len(set(srcnames))!=len(srcnames):raise RuntimeError('Dupli ZIP partovi u ulazu')
        if any('/_xmlsignatures/' in s or s.startswith('_xmlsignatures/') for s in srcnames):
            raise RuntimeError('Potpisan OOXML paket: dodavanje sheet-ova bi invalidiralo potpis')
        for name in ('xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'):
            if name not in srcnames:raise RuntimeError('Nedostaje OOXML deo '+name)
        wbraw=zin.read('xl/workbook.xml')
        relraw=zin.read('xl/_rels/workbook.xml.rels')
        ctraw=zin.read('[Content_Types].xml')
        wbr=ET.fromstring(wbraw)
        sheetgroup=wbr.find(f'{{{MAIN}}}sheets')
        if sheetgroup is None:raise RuntimeError('XLSM nema Sheets cvor')
        oldnames={x.attrib.get('name','') for x in sheetgroup}
        overlap=oldnames.intersection(ARCHIVE)
        if overlap:raise RuntimeError('Vec ima arhivske listove: '+str(overlap))
        maxid=max((int(x.attrib['sheetId']) for x in sheetgroup),default=0)
        relations=ET.fromstring(relraw)
        oldrids={x.attrib.get('Id','') for x in relations}
        oldfiles=set(srcnames)
        paths={}
        for i,name in enumerate(ARCHIVE,1):
            sid=maxid+i
            idx=1
            while f'xl/worksheets/sheet{idx}.xml' in oldfiles or f'xl/worksheets/sheet{idx}.xml' in paths.values():idx+=1
            path=f'xl/worksheets/sheet{idx}.xml'
            rid=f'rIdMIGARCHIVE{sid}'
            if rid in oldrids:raise RuntimeError('ID collision '+rid)
            paths[name]=(sid,rid,path)
            oldfiles.add(path)
        sheetns_match=re.search(r'<([A-Za-z_][\w.-]*:)?sheets(?:\s|>)',wbraw.decode('utf-8-sig'))
        pfx=sheetns_match.group(1) if sheetns_match else ''
        if not sheetns_match:raise RuntimeError('Nije pronadjen tag sheets')
        newsheets=[];newrels=[];newct=[]
        for name,(sid,rid,path) in paths.items():
            newsheets.append(f'<{pfx}sheet name={quoteattr(name)} sheetId="{sid}" r:id="{rid}" xmlns:r="{RREL}"/>')
            newrels.append(f'<Relationship Type="{RREL}/worksheet" Target="/{path}" Id="{rid}"/>')
            newct.append(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        workbook_xml=insert_before_close(wbraw,'sheets',''.join(newsheets))
        rel_xml=insert_before_close(relraw,'Relationships',''.join(newrels))
        ct_xml=insert_before_close(ctraw,'Types',''.join(newct))
        for content in (workbook_xml,rel_xml,ct_xml):ET.fromstring(content)
        replacement={'xl/workbook.xml':workbook_xml,
                     'xl/_rels/workbook.xml.rels':rel_xml,
                     '[Content_Types].xml':ct_xml}
        newparts={}
        for name,(_,_,path) in paths.items():
            h,rows=xml_start() if name=='ARHIVA_START' else csv_rows(output/f'{name}.csv')
            newparts[path]=sheet_xml(h,rows,name=='ARHIVA_START')
        with zipfile.ZipFile(target,'w',allowZip64=True) as zout:
            for info in zin.infolist():
                zout.writestr(info,replacement.get(info.filename,zin.read(info.filename)))
            for path,data in newparts.items():
                zout.writestr(path,data,compress_type=zipfile.ZIP_DEFLATED,compresslevel=6)
    return list(newparts)

def validate_package(src:Path,dst:Path,old_tables,old_vba:str,newparts):
    with zipfile.ZipFile(src) as a,zipfile.ZipFile(dst) as b:
        if b.testzip() is not None:raise RuntimeError('Izlaz ima neispravan ZIP element')
        diff_allowed={'xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'}
        for name in a.namelist():
            if name not in diff_allowed and a.read(name)!=b.read(name):
                raise RuntimeError('Originalni ZIP deo PROMENJEN: '+name)
        for name in newparts: ET.fromstring(b.read(name))
        wb=ET.fromstring(b.read('xl/workbook.xml'))
        found={x.attrib['name'] for x in wb.find(f'{{{MAIN}}}sheets')}
        if not set(ARCHIVE).issubset(found):raise RuntimeError('Nedostaje arhivski list')
        rel=ET.fromstring(b.read('xl/_rels/workbook.xml.rels'))
        relids={x.attrib.get('Id') for x in rel}
        for x in wb.find(f'{{{MAIN}}}sheets'):
            if x.attrib.get(f'{{{RREL}}}id') not in relids:
                raise RuntimeError('Nepostojeca workbook relacija za '+x.attrib.get('name',''))
        typ=ET.fromstring(b.read('[Content_Types].xml'))
        cts={x.attrib.get('PartName') for x in typ if x.tag.endswith('Override')}
        for x in newparts:
            if '/'+x not in cts:raise RuntimeError('Nedostaje ContentType za '+x)
    if project_bytes(dst)!=old_vba:raise RuntimeError('VBA PROJECT BINARY JE PROMENJEN')
    newtables=read_workbook(dst)
    if workbook_signature(old_tables)!=workbook_signature(newtables):
        raise RuntimeError('Poslovne tabele NISU identicne sa izvorom')

def matches_mig(path:Path)->bool:
    s=str(path).lower().replace('-','_').replace('.','_')
    return bool(re.search(r'mig19_1[34](?:\D|$)',s)) and path.suffix.lower()=='.xlsm' and 'arhiva' not in s

def find_source(root:Path):
    cands=[]
    for dir,dirs,files in os.walk(root):
        dirs[:]=[x for x in dirs if not x.startswith('.') and not x.lower().startswith(('migrationarchivexlsm','node_modules','.git'))]
        for f in files:
            p=Path(dir)/f
            if matches_mig(p):cands.append(p)
    if not cands:raise RuntimeError('Nije pronadjena MIG19.13 / MIG19.14 XLSM ispod: '+str(root))
    cands.sort(key=lambda p:(not bool(re.search(r'mig19_14(?:\D|$)',str(p).lower().replace('-','_').replace('.','_'))), -p.stat().st_mtime))
    print(f'Pronadjeno MIG19.13/14 kandidata: {len(cands)}',flush=True)
    failures=[]
    for p in cands:
        try:
            T=read_workbook(p);assert_clean(p,T);vba=project_bytes(p)
            print('IZABRANA CISTA SVESKA: '+str(p),flush=True)
            return p,T,vba
        except Exception as e:
            print('ODBACENA: '+str(p)+' | '+str(e)[:320],flush=True)
            failures.append(str(p)+' | '+str(e)[:220])
    raise RuntimeError('Nema ciste 13/14 sveske. Ne pravimo arhivu iz ostecene kopije. '+ ' // '.join(failures[:6]))

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--root',type=Path,help='Trazi cistu MIG19.13/14 ispod AgriX DEV')
    p.add_argument('--workbook',type=Path,help='Eksplicitna cista MIG19.13/14 XLSM')
    p.add_argument('--out-dir',type=Path,required=True,help='PRAZAN NOV folder')
    args=p.parse_args()
    if (args.root is None)==(args.workbook is None):p.error('Navedi TACNO JEDNO: --root ILI --workbook')
    out=args.out_dir.resolve()
    if out.exists() and any(out.iterdir()):p.error('Izlazni folder nije prazan: '+str(out))
    if args.workbook:
        source=args.workbook.resolve();
        if not source.is_file():p.error('Sveska ne postoji: '+str(source))
        T=read_workbook(source); assert_clean(source,T);vba=project_bytes(source)
    else:
        source,T,vba=find_source(args.root.resolve())
    if out==source.parent or source in out.parents:p.error('Izlazni folder ne sme biti izvorni folder')
    before_hash=sha256(source)
    print(f'IZVOR: {source}\nSHA: {before_hash}',flush=True)
    records,extra=export_otkups(T)
    issues=extract_issues(T,extra)
    out.mkdir(parents=True,exist_ok=True)
    filecounts,active,missing=build_csvs(out,records,issues,T)
    candidate=out/'__ARHIVA_BUILD__.xlsm'
    final=out/'AgriX_2026_ARHIVA.xlsm'
    if candidate.exists() or final.exists():raise RuntimeError('Izlaz vec postoji')
    try:
        print('Dodajem arhivske listove DIREKTNO u novu ZIP/XLSM kopiju (bez Excel COM-a)...',flush=True)
        newparts=build_without_excel(source,candidate,out)
        print('Proveravam VBA, svih 47 tabela i originalne ZIP delove...',flush=True)
        validate_package(source,candidate,T,vba,newparts)
        if sha256(source)!=before_hash:raise RuntimeError('Izvor je promenjen tokom rada')
        candidate.replace(final)
    except Exception:
        candidate.unlink(missing_ok=True)
        raise
    data={'version':V,'status':'ARCHIVE_XLSM_CREATED','source':str(source),'source_sha256':before_hash,
          'output':str(final),'output_sha256':sha256(final),'rows_otkup':len(records),
          'tables_verified':len(T),'exceptions':len(issues),'active_otkup':active,
          'missing_otkup_values':missing,'csv_rows':filecounts,'business_tables_identical':True,
          'vba_binary_identical':True,'original_package_parts_unchanged_except_metadata':True,
          'excel_com_used':False,'source_unchanged':True,'created':datetime.now().isoformat(timespec='seconds')}
    (out/'ARHIVA_2026_MAKE_SUMMARY.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(data,ensure_ascii=False,indent=2),flush=True)

if __name__=='__main__':
    try:main()
    except Exception as e:
        print('MIG-ARCHIVE v1.3 STOP: '+str(e),file=sys.stderr)
        sys.exit(2)
