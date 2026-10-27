#!/usr/bin/env python3
"""MIG-ARCHIVE / istorija malina 2026. READ ONLY.

Ulaz: postojeca migrirana XLSM (Excel ListObjects), nikada je ne otvara u Excel-u
niti pise u nju. Izlaz: nezavisan offline HTML pregled + CSV izuzeci + CSV otkupi.
Upozorenja ne skrivaju istorijske dokumente. Izvestaji po OM/kooperantu koriste
OTKUP + OTKUP_STAVKE, ne rekonstruisu vrednost preko OTP/ZBR.
"""
from __future__ import annotations
import argparse, csv, hashlib, html, json, re, sys
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from datetime import datetime
from agrix_mig19_malina_readonly_audit import XlsmReadOnly, dnum, ddate, norm, key, sha256
import agrix_mig20_sve_tabele_readonly as mig20

VERSION = 'AGRIx-MIG-ARCHIVE-v1'
REQUIRED = ['tblOtkup','tblOtkupStavke','tblStanice','tblKooperanti',
            'tblOtpremnica','tblOtpremnicaIzvori','tblOtpremnicaStavke',
            'tblZbirna','tblZbirnaIzvori','tblPrijemnica']
QZERO = Decimal('0')
STOP = {'DA','TRUE','1'}
ARCHIVE_CODES = {
 'ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP': ('NEPOTPUN IZVOR OTP','OTP vrednost nije dokaziva'),
 'ACTIVE_ZBR_CONTAINS_STORNO_OTP': ('NEPOTPUN LANAC ZBR','ZBR ima storniranu OTP'),
 'ACTIVE_ZBR_NO_ACTIVE_OTP_SOURCE': ('NEPOTPUN LANAC ZBR','ZBR nema aktivnu OTP'),
 'DUPLICATE_ZBR_OTP_MEMBERSHIP': ('SPORNA VEZA','ZBR ima ponovljenu vezu OTP'),
 'OTP_IN_MULTIPLE_ACTIVE_ZBR': ('SPORNA VEZA','OTP je u vise aktivnih ZBR'),
 'ACTIVE_PRJ_ONLY_STORNO_ZBR': ('NEPOTPUN LANAC PRJ','PRJ vezana za storniranu ZBR'),
 'ACTIVE_PRJ_NO_ZBR_NUMBER': ('NEPOTPUN LANAC PRJ','Broj zbirne nije pronadjen'),
 'ACTIVE_CHAIN_NO_OTKUPS_AND_CANCELLED_PRJ': ('NEPOTPUN LANAC ZBR','ZBR sa OTP bez izvora i storniranom PRJ'),
 'ONE_OTKUP_MULTIPLE_OTP': ('SPORNA VEZA','Otkup povezan sa vise OTP'),
 'DUPLICATE_OTKUP_OTP_MEMBERSHIP': ('SPORNA VEZA','Duplirana veza otkupa i OTP'),
 'INCOMPLETE_OTPREMNICA_SOURCE_LINK': ('SPORNA VEZA','Nepotpuna veza izvora OTP'),
 'ACTIVE_OTKUP_ONLY_STORNO_OTP_LINKS': ('SPORNA VEZA','Otkup vezan samo za stornirane OTP'),
 'ACTIVE_OTKUP_LEGACY_BACKLINK_TO_STORNO_OTP': ('SPORNA VEZA','Zaglavlje otkupa upucuje na storniranu OTP'),
 'ORPHAN_FOREIGN_KEY': ('NEPOSTOJECI ID','Veza pokazuje na nepostojeci red'),
 'DUPLICATE_PRIMARY_ID': ('DUPLI ID','Identifikator nije jedinstven'),
 'EMPTY_PRIMARY_ID': ('PRAZAN ID','Dokument nema identitet'),
 'ACTIVE_OTKUP_NO_LINES': ('NEDOSTAJU STAVKE','Otkup nema stavke'),
 'PRJ_ZBR_NUMBER_AMBIGUOUS': ('DVOSMISLENA VEZA','Vise zbirnih pod istim brojem'),
 'PRJ_ZBR_CASE_ONLY_OR_FORMAT': ('FORMAT BROJA','Razlika u formatu broja zbirne'),
}


def is_storno(r): return key(r.get('Stornirano')) in STOP

def q(r, field):
    x = dnum(r.get(field))
    if x is None or not x.is_finite(): return None
    return x

def clean_str(s): return norm(s)

def csv_safe(v):
    s='' if v is None else str(v)
    # Excel formula injection in exported CSV from historical free text.
    return "'"+s if s.startswith(('=','+','-','@','\t','\r')) else s

def csv_save(path, rows, fields):
    with path.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=fields,delimiter=';',extrasaction='ignore')
        w.writeheader()
        for row in rows:w.writerow({col:csv_safe(row.get(col,'')) for col in fields})

def read_workbook(path):
    ro=XlsmReadOnly(path)
    try:
        missing=sorted(set(REQUIRED)-set(ro.tables))
        if missing:raise RuntimeError('Nedostaju obavezne ListObject tabele: '+', '.join(missing))
        # Audit reads all available tables; avoid losing orphan checks.
        tables={name:ro.read_table(name,set()) for name in sorted(ro.tables)}
    finally:ro.close()
    for table,required_fields in {
      'tblOtkup':['OtkupID','Datum','StanicaID','KooperantID','BrojDokumenta','Stornirano'],
      'tblOtkupStavke':['OtkupID','Kolicina','Cena'],
      'tblKooperanti':['KooperantID','Ime','Prezime'],
      'tblStanice':['StanicaID','Naziv'],
      'tblOtpremnica':['OtpremnicaID','IzdatoStatus','Stornirano'],
    }.items():
        if tables[table]:
            absent=[x for x in required_fields if x not in tables[table][0]]
            if absent:raise RuntimeError(f'{table} nedostaju kolone {absent}')
    return tables


def export_otkups(T):
    stations={key(r.get('StanicaID')):clean_str(r.get('Naziv')) for r in T['tblStanice']}
    coop={key(r.get('KooperantID')):r for r in T['tblKooperanti']}
    by_otk=defaultdict(list)
    for line in T['tblOtkupStavke']:
        by_otk[key(line.get('OtkupID'))].append(line)
    records=[]; extra=[]
    ids=[key(r.get('OtkupID')) for r in T['tblOtkup']]
    if not all(ids) or len(set(ids))!=len(ids):
        raise RuntimeError('tblOtkup ima prazan ili dupliran OtkupID; zbirni prikaz nije pouzdan. STOP.')
    lineids=[key(r.get('OtkupStavkaID')) for r in T['tblOtkupStavke'] if key(r.get('OtkupStavkaID'))]
    if len(set(lineids))!=len(lineids):
        raise RuntimeError('tblOtkupStavke ima duplirani OtkupStavkaID; zbirni prikaz nije pouzdan. STOP.')
    records=[]; extra=[]
    for r in T['tblOtkup']:
        oid=key(r.get('OtkupID'))
        lines=by_otk.get(oid,[])
        kg=QZERO;value=QZERO;missing_price=False;missing_qty=False
        classes=set()
        for item in lines:
            k=q(item,'Kolicina');p=q(item,'Cena')
            if k is None:missing_qty=True
            else:kg+=k
            if k is None or p is None:missing_price=True
            else:value+=k*p
            if key(item.get('Klasa')):classes.add(clean_str(item.get('Klasa')))
        koop_id=key(r.get('KooperantID'))
        koop_name=' '.join(clean_str(coop.get(koop_id,{}).get(k)) for k in ['Ime','Prezime']).strip()
        if not koop_name:koop_name=koop_id or '(bez ID)'
        sid=key(r.get('StanicaID')); sn=stations.get(sid,sid or '(bez OM)')
        flags=[]
        if not lines:flags.append('NEMA_STAVKI')
        if missing_qty:flags.append('NEMA_KOLICINE')
        if missing_price:flags.append('NEMA_CENE')
        if koop_id not in coop:flags.append('NEPOSTOJECI_KOOPERANT')
        if sid not in stations:flags.append('NEPOSTOJECE_OM')
        date=ddate(r.get('Datum'))
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',date):flags.append('NEVALIDAN_DATUM')
        out={
           'id':oid,'excel_row':r.get('_excel_row'),'date':date,'business':clean_str(r.get('BrojDokumenta')),
           'station_id':sid,'station':sn,'coop_id':koop_id,'coop':koop_name,
           'fruit':clean_str(r.get('VrstaVoca')),'variety':clean_str(r.get('SortaVoca')),
           'classes':', '.join(sorted(classes)),'storno':is_storno(r),
           'kg':str(kg) if lines and not missing_qty else None,
           'value':str(value) if lines and not missing_price else None,
           'flags':flags
        }
        records.append(out)
        for fl in flags:
            extra.append({'priority':'P1' if fl in ('NEMA_STAVKI','NEMA_KOLICINE') else 'P2',
                  'code':'ARCHIVE_'+fl,'table':'tblOtkup','excel_row':r.get('_excel_row'),
                  'document_id':oid,'related_id':'','date':date,
                  'problem':f'Otkup {oid}: {fl}. BrojDokumenta={out["business"]}',
                  'action':'Rucni pregled; arhiva ne menja podatke.'})
    return records, extra


def extract_issues(T, extra):
    # Reuse independently audited, already checked relationships. No amb/palette/stage
    # mandatory checks: archive scope is OTK -> OTP -> ZBR -> PRJ only.
    a=mig20.Audit(T)
    mig20.scan_global(a)
    mig20.scan_otk_otp(a)
    mig20.scan_zbr(a)
    mig20.scan_prj(a)
    mig20.detect_cross_link(a)
    allowed_tables={'tblOtkup','tblOtkupStavke','tblOtpremnica',
                    'tblOtpremnicaStavke','tblOtpremnicaIzvori','tblZbirna',
                    'tblZbirnaStavke','tblZbirnaIzvori','tblPrijemnica','tblPrijemnicaStavke'}
    issues=[];seen=set()
    for f in list(a.findings)+extra:
        code=f['code'];table=f['table']
        if code not in ARCHIVE_CODES and not code.startswith('ARCHIVE_'):continue
        if code=='ACTIVE_OTKUP_NO_LINES' and any(x['document_id']==f['document_id'] and x['code']=='ARCHIVE_NEMA_STAVKI' for x in extra):continue
        if code in ('ORPHAN_FOREIGN_KEY','DUPLICATE_PRIMARY_ID','EMPTY_PRIMARY_ID') and table not in allowed_tables:continue
        if table.startswith('tblPaleta'):continue
        kt=(code,table,key(f.get('document_id')),key(f.get('related_id')))
        if kt in seen:continue
        seen.add(kt)
        category,impact=ARCHIVE_CODES.get(code,('NEPOTPUN OTKUP','Nepotpuni podaci za otkup'))
        # Only these reports have evidence of a hard runtime guard in this source.
        blocks='VOZACI_ROBA_VREDNOST' if code=='ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP' else ''
        issues.append({'priority':f['priority'],'category':category,'code':code,
           'table':table,'excel_row':f.get('excel_row',''),
           'document_id':f['document_id'],'related_id':f.get('related_id',''),
           'date':f.get('date',''),'problem':f['problem'],'impact':impact,
           'blocks':blocks,'manual_decision':'','manual_note':''})
    issues.sort(key=lambda x:({'P1':0,'P2':1,'P3':2,'INFO':3}.get(x['priority'],9),x['category'],x['document_id']))
    return issues


def render_html(records, issues, meta):
    data=json.dumps({'records':records,'issues':issues,'meta':meta},ensure_ascii=False,separators=(',',':')).replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026')
    page=r'''<!DOCTYPE html><html lang="sr"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/><title>AgriX | Arhiva 2026</title>
<style>
:root{color-scheme:light;font-family:Segoe UI,Arial,sans-serif;--green:#194d45;--teal:#0b7162;--line:#dbe4e1;--muted:#61716c;--paper:#fff;--bg:#f4f7f5}*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:#172823;font-size:14px}header{background:#173f38;color:#fff;padding:27px max(22px,calc((100vw - 1300px)/2)) 22px}header h1{font-size:27px;margin:0 0 6px}header p{opacity:.85;margin:0}main{max-width:1344px;padding:22px;margin:0 auto}nav{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:18px}button.tab{background:#fff;border:1px solid var(--line);padding:12px 18px;border-radius:8px;color:#244c43;font-weight:600;cursor:pointer}button.tab.active{background:var(--green);color:white}.notice{border-left:4px solid #e1a23d;background:#fff9ea;padding:12px 15px;margin:0 0 18px;line-height:1.5}.filters{display:flex;gap:12px;align-items:end;flex-wrap:wrap;background:#fff;padding:17px;border-radius:11px;border:1px solid var(--line);margin-bottom:16px}label{display:flex;flex-direction:column;gap:5px;font-weight:600;color:#456158;font-size:12px}select,input{border:1px solid #cadbd3;border-radius:7px;background:#fff;color:#1c342b;padding:9px 11px;font-size:14px;min-width:145px}input[type=search]{min-width:235px}.check{flex-direction:row;align-items:center;gap:7px;padding:10px}.check input{min-width:0}.metricrow{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:12px 0 20px}.metric{background:#fff;border:1px solid var(--line);padding:17px;border-radius:10px}.metric b{display:block;font-size:25px;margin-top:7px}.metric span{color:var(--muted);font-size:12px}.surface{background:#fff;border:1px solid var(--line);border-radius:11px;padding:15px}.summary{color:var(--muted);padding:0 4px 14px}.overflow{overflow:auto;max-height:68vh;border:1px solid #e4ece9;border-radius:8px}table{border-collapse:collapse;width:100%;white-space:nowrap}th{position:sticky;top:0;z-index:2;background:#edf4f0;text-align:left;color:#35554b;font-size:12px;border-bottom:1px solid #cfded7;padding:11px 12px}td{padding:9px 12px;border-bottom:1px solid #edf0ef}tbody tr:hover{background:#f4f8f5}.num{text-align:right;font-variant-numeric:tabular-nums}.warn{color:#915c05;font-weight:700}.muted{color:var(--muted)}.bad{background:#fff2d5;color:#8b4e00;border-radius:5px;padding:3px 7px;font-size:12px}.detail{cursor:pointer;color:#176d61;text-decoration:underline}.footer{margin-top:14px;color:#6a7a74;font-size:12px;line-height:1.5}.activeFilter{background:#eaf4ed;padding:8px 10px;margin-bottom:12px}.none{padding:35px;text-align:center;color:#788881}@media(max-width:700px){.metricrow{grid-template-columns:repeat(2,1fr)}header h1{font-size:22px}main{padding:12px}input[type=search]{min-width:160px}}
</style></head><body>
<header><h1>AgriX · Arhiva 2026</h1><p>Istorijska dokumentacija · read-only · bez povezivanja na internet · podaci ostaju neizmenjeni</p></header>
<main><div class="notice" id="notice">Izveštaji po otkupnim mestima i kooperantima računaju se direktno iz otkupnih listova i stavki. Neproverene veze OTP/ZBR su u zasebnoj evidenciji. Nisu predstavljene kao nula.</div>
<nav id="tabs"><button class="tab active" data-view="om">OM · zbirno</button><button class="tab" data-view="otk">OM · pojedinačno</button><button class="tab" data-view="coop">Kooperanti</button><button class="tab" data-view="issues">Izuzeci</button></nav>
<div class="filters"><label>Otkupno mesto<select id="station"></select></label><label>Kooperant<select id="cooperator"></select></label><label>Datum od<input type="date" id="dateFrom" value="2026-01-01"/></label><label>Datum do<input type="date" id="dateTo" value="2026-12-31"/></label><label>Pretraga<input type="search" id="search" placeholder="ID, broj, naziv, tekst..."/></label><label class="check"><input type="checkbox" id="storno"/> Uključi stornirane</label></div>
<div id="kpis" class="metricrow"></div><section class="surface"><div class="summary" id="summary"></div><div class="overflow" id="display"></div></section><div class="footer" id="footer"></div></main>
<script id="archive-data" type="application/json">__DATA__</script>
<script>
'use strict';
const data=JSON.parse(document.getElementById('archive-data').textContent);
if(data.meta.duplicate_primary_id_findings){document.getElementById('notice').textContent+=' UPOZORENJE: '+data.meta.duplicate_primary_id_findings+' dupliranih identifikatora u izvornom fajlu; proveri registar izuzetaka i po potrebi pređi na čistu rezervnu kopiju.';}
const $=id=>document.getElementById(id);
let view='om';
const esc=x=>String(x??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const currency=x=>x===null||x===undefined||x===''?'—':new Intl.NumberFormat('sr-RS',{minimumFractionDigits:2,maximumFractionDigits:2}).format(Number(x));
const kg=x=>x===null||x===undefined||x===''?'—':new Intl.NumberFormat('sr-RS',{maximumFractionDigits:3}).format(Number(x));
const n=x=>new Intl.NumberFormat('sr-RS').format(x);
const calc=(rows,k)=>rows.some(r=>r[k]===null)?null:rows.reduce((v,r)=>v+Number(r[k]),0);
function listSelect(id,items,allLabel){let x=$(id),parts=['<option value="">'+esc(allLabel)+'</option>'];for(let [k,name] of items)parts.push('<option value="'+esc(k)+'">'+esc(name)+'</option>');x.innerHTML=parts.join('')}
const sitems=[...new Map(data.records.map(r=>[r.station_id,r.station])).entries()].filter(x=>x[0]).sort((a,b)=>a[1].localeCompare(b[1]));
const citems=[...new Map(data.records.map(r=>[r.coop_id,r.coop+' ['+r.coop_id+']'])).entries()].filter(x=>x[0]).sort((a,b)=>a[1].localeCompare(b[1]));
listSelect('station',sitems,'Sva otkupna mesta');listSelect('cooperator',citems,'Svi kooperanti');
const getFilters=()=>({s:$('station').value,c:$('cooperator').value,f:$('dateFrom').value,t:$('dateTo').value,q:$('search').value.toLowerCase().trim(),v:$('storno').checked});
function filteredRecords(f){return data.records.filter(r=>(!f.s||r.station_id===f.s)&&(!f.c||r.coop_id===f.c)&&(!f.f||r.date>=f.f)&&(!f.t||r.date<=f.t)&&(f.v||!r.storno)&&(!f.q||[r.id,r.business,r.coop,r.station,r.fruit,r.variety,r.classes].join(' ').toLowerCase().includes(f.q)))}
function showTable(headers,rows,count){const limit=350;let html='<table><thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>';
if(!rows.length){$('display').innerHTML='<div class="none">Nema stavki za ovaj filter.</div>';return}
for(let row of rows.slice(0,limit))html+='<tr>'+row.map((x,i)=>'<td '+(typeof x==='object'&&x&&x.n?'class="num"':'')+'>'+esc(typeof x==='object'&&x&&x.n?x.v:x)+'</td>').join('')+'</tr>';
$('display').innerHTML=html+'</tbody></table>';if(rows.length>limit)$('summary').textContent+=' · Prvih '+n(limit)+' od '+n(rows.length)+' redova; CSV sadrži sve.'}
const num=x=>({v:x,n:true});
function render(){let f=getFilters(),rows=filteredRecords(f),ids=new Set(rows.map(r=>r.id)),totalKg=calc(rows,'kg'),totalValue=calc(rows,'value'),bad=rows.filter(r=>r.kg===null||r.value===null).length;
$('kpis').innerHTML=[['Otkupa u pregledu',n(rows.length)],['Kilograma (dokazano)',kg(rows.filter(r=>r.kg!==null).reduce((s,r)=>s+Number(r.kg),0))],['Vrednost (dokazano)',currency(rows.filter(r=>r.value!==null).reduce((s,r)=>s+Number(r.value),0))],['Nepotpuni otkupi',n(bad)]].map(([k,v])=>'<div class="metric"><span>'+esc(k)+'</span><b>'+esc(v)+'</b></div>').join('');
$('summary').textContent='';
if(view==='om'){let groups=new Map();for(let r of rows){let id=r.station_id;if(!groups.has(id))groups.set(id,[]);groups.get(id).push(r)}let out=[...groups].map(([id,v])=>[id,v[0].station,num(v.length),num(v.filter(z=>z.kg!==null).reduce((s,z)=>s+Number(z.kg),0).toLocaleString('sr-RS',{maximumFractionDigits:3})),num(v.filter(z=>z.value!==null).reduce((s,z)=>s+Number(z.value),0).toLocaleString('sr-RS',{minimumFractionDigits:2,maximumFractionDigits:2})),v.filter(z=>z.value===null||z.kg===null).length]);
$('summary').textContent='Ukupno '+n(out.length)+' otkupnih mesta. Iznosi su zbir poznatih vrednosti, nepotpuni dokumenti navedeni odvojeno.';showTable(['OM ID','Naziv','Broj otkupa','Kg','Vrednost RSD','Nepotpuno'],out);}
else if(view==='otk'){$('summary').textContent='Lista istorijskih otkupa. Klik na broj izvornog dokumenta nije potreban: svi originalni ID-jevi su prikazani.';
showTable(['OtkupID','Datum','Broj dokumenta','OM','Kooperant','Voće / sorta','Klasa','Kg','Vrednost RSD','Status'],rows.sort((a,b)=>a.date.localeCompare(b.date)||a.id.localeCompare(b.id)).map(r=>[r.id,r.date,r.business,r.station,r.coop,r.fruit+' '+r.variety,r.classes,num(kg(r.kg)),num(currency(r.value)),r.storno?'STORNO':r.flags.length?'NEPOTPUNO':'AKTIVAN']));}
else if(view==='coop'){let c=f.c;if(c){$('summary').textContent='Pojedinačna istorija izabranog kooperanta; nema OTP/ZBR prevezivanja.';showTable(['Datum','Broj otkupa','OtkupID','OM','Voće / sorta','Kg','Vrednost RSD','Status'],rows.sort((a,b)=>a.date.localeCompare(b.date)||a.id.localeCompare(b.id)).map(r=>[r.date,r.business,r.id,r.station,r.fruit+' '+r.variety,num(kg(r.kg)),num(currency(r.value)),r.storno?'STORNO':r.flags.length?'NEPOTPUNO':'AKTIVAN']));}
else{let group=new Map();for(let r of rows){if(!group.has(r.coop_id))group.set(r.coop_id,[]);group.get(r.coop_id).push(r)}$('summary').textContent='Izaberi kooperanta iz gornje liste za njegovu pojedinačnu istoriju.';
showTable(['KooperantID','Kooperant','OM','Broj otkupa','Kg','Vrednost RSD','Nepotpuno'],[...group].map(([id,v])=>[id,v[0].coop,v[0].station,num(v.length),num(kg(v.filter(r=>r.kg!==null).reduce((a,r)=>a+Number(r.kg),0))),num(currency(v.filter(r=>r.value!==null).reduce((a,r)=>a+Number(r.value),0))),v.filter(r=>r.kg===null||r.value===null).length]));}}
else{let ix=data.issues.filter(x=>(!f.f||!x.date||x.date>=f.f)&&(!f.t||!x.date||x.date<=f.t)&&(!f.q||[x.code,x.category,x.document_id,x.related_id,x.problem,x.table].join(' ').toLowerCase().includes(f.q)));
$('summary').textContent=n(ix.length)+' zapis(a) za ručnu proveru. Neusklađeni dokumenti nisu automatski izbrisani ili stornirani. CSV sadrži celokupan skup.';
showTable(['Prioritet','Grupa','Dokument','Veza','Tabela / Excel red','Problem','Uticaj'],ix.map(x=>[x.priority,x.category,x.document_id,x.related_id,x.table+' / '+x.excel_row,x.problem,x.impact]));}
$('footer').textContent='Izvorna XLSM: '+data.meta.source+' · SHA256 '+data.meta.sha+' · Generisano '+data.meta.generated+' · Ovaj HTML nema API, mrežne pozive niti mogućnost upisa u XLSM. Za pun finansijski saldo koristi validiran originalni izveštaj.';
}
for(let btn of document.querySelectorAll('.tab'))btn.addEventListener('click',()=>{view=btn.dataset.view;document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===btn));render()});
for(let el of document.querySelectorAll('.filters input,.filters select'))el.addEventListener('input',render);
render();
</script></body></html>'''
    return page.replace('__DATA__',data)


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook',type=Path,required=True,help='postojeca QA XLSM (SAMO CITANJE)')
    parser.add_argument('--out-dir',type=Path,required=True,help='NOVI folder za izvestaj, ne XLSM')
    args=parser.parse_args(argv)
    path=args.workbook.resolve()
    if not path.is_file():parser.error(f'Ne postoji XLSM: {path}')
    if path.suffix.lower()!='.xlsm':parser.error('--workbook mora biti XLSM')
    dest=args.out_dir.resolve()
    # Only create new files outside source workbook. Do not silently overwrite user notes.
    if dest.exists() and any(dest.iterdir()):parser.error('Out-dir mora biti NOV ili prazan (bez prepisivanja rucnih beleški).')
    digest=sha256(path)
    tables=read_workbook(path)
    records,extra=export_otkups(tables)
    issues=extract_issues(tables,extra)
    # Exact table row evidence, for users who want to inspect historical records
    # without modifying the workbook or reconstructing missing parents.
    original_by_row={(table,str(r.get('_excel_row'))):r
                     for table,rows in tables.items() for r in rows}
    raw_issues=[]
    for issue in issues:
        r=original_by_row.get((issue['table'],str(issue['excel_row'])),{})
        name=next((norm(r.get(f)) for f in ('BrojDokumenta','BrojOtpremnice',
                         'BrojZbirne','BrojPrijemnice') if norm(r.get(f))), '')
        raw_issues.append({**issue,'source_number':name,
             'source_date':ddate(r.get('Datum')) if r else '',
             'source_station':norm(r.get('StanicaID')),
             'source_cooperator':norm(r.get('KooperantID')),
             'source_driver':norm(r.get('VozacID')),
             'source_storno':norm(r.get('Stornirano')),
             'source_issue_status':norm(r.get('IzdatoStatus')),
             'source_row_json':json.dumps(r,ensure_ascii=False,sort_keys=True) if r else ''})
    if sha256(path)!=digest:raise RuntimeError('XLSM se promenila tokom citanja; rezultat nije validan.')
    if len(tables.get('tblOtkup',[]))!=len(records):raise RuntimeError('Broj otkupa nije identican ulazu')
    meta={'version':VERSION,'source':str(path),'sha':digest,
          'duplicate_primary_id_findings':sum(i['code']=='DUPLICATE_PRIMARY_ID' for i in issues),
          'generated':datetime.now().astimezone().isoformat(timespec='seconds'),
          'rows':len(records),'tables_read':len(tables),'issues':len(issues),
          'scope':'OM, kooperanti, OTK/OTP/ZBR/PRJ; palete iskljucene',
          'no_changes_to_workbook':True}
    dest.mkdir(parents=True,exist_ok=True)
    (dest/'ARHIVA_2026_PREGLED.html').write_text(render_html(records,issues,meta),encoding='utf-8')
    fields=['id','excel_row','date','business','station_id','station','coop_id','coop',
            'fruit','variety','classes','storno','kg','value','flags']
    csv_save(dest/'ARHIVA_2026_OTKUPI.csv',
       [{**r,'flags':','.join(r['flags'])} for r in records],fields)
    fx=['priority','category','code','table','excel_row','document_id','related_id','date',
        'problem','impact','blocks']
    csv_save(dest/'ARHIVA_2026_IZUZECI.csv',issues,fx)
    csv_save(dest/'ARHIVA_2026_IZUZECI_SIROVI_REDOVI.csv',raw_issues,fx+
       ['source_number','source_date','source_station','source_cooperator','source_driver',
        'source_storno','source_issue_status','source_row_json'])
    review_path=dest/'ARHIVA_2026_RUCNI_PREGLED.csv'
    csv_save(review_path,issues,fx+['manual_decision','manual_note'])
    (dest/'ARHIVA_2026_SUMMARY.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2),encoding='utf-8')
    if sha256(path)!=digest:raise RuntimeError('Nakon izvoza XLSM hash vise nije isti! STOP.')
    print(json.dumps({'status':'READ_ONLY_COMPLETE',**meta,
         'codes':dict(Counter(i['code'] for i in issues)),
         'output_files':[str(x) for x in sorted(dest.iterdir())]},ensure_ascii=False,indent=2))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as e:print(f'MIG-ARCHIVE STOP: {e}',file=sys.stderr);sys.exit(2)
