#!/usr/bin/env python3
"""AGRIx MIG20 — independent full-workbook relationship/lifecycle audit (READ ONLY).

Reads all Excel ListObjects in an existing XLSM, optionally checks stage JSONL
row-id coverage. No Excel automation, no VBA execution, no workbook modifications.
Stdlib only; needs agrix_mig19_malina_readonly_audit.py beside this file.

Usage (Git Bash, inside agrix_legacy_audit):
  python -u agrix_mig20_sve_tabele_readonly.py \
      --workbook '../MigrationMIG19_13_PatchedXLSM/AgriX_MIG19_13_AMBALAZA_QA_PARTIAL.xlsm' \
      --stage '../MigrationStage_v2' \
      --out-dir '../MigrationMIG20_FullAudit'

Known legacy: MALINA_MODE ST-00001 may go negative, station mirrors driver.
WARNING: Source data can be imperfect; findings are CANDIDATES, never auto-fixes.
"""
from __future__ import annotations
import argparse, csv, json, re, sys
from pathlib import Path
from collections import Counter, defaultdict
from decimal import Decimal
from datetime import datetime
from agrix_mig19_malina_readonly_audit import XlsmReadOnly, dnum, ddate, norm, key, sha256

VERSION='AGRIx-MIG20-FULL-WORKBOOK-RELATIONSHIP-AUDIT-v1'
ID_COL={
 'tblOtkup':'OtkupID','tblOtkupStavke':'OtkupStavkaID',
 'tblOtpremnica':'OtpremnicaID','tblOtpremnicaStavke':'OtpremnicaStavkaID',
 'tblOtpremnicaIzvori':'OtpremnicaIzvorID',
 'tblZbirna':'ZbirnaID','tblZbirnaStavke':'ZbirnaStavkaID',
 'tblZbirnaIzvori':'ZbirnaIzvorID',
 'tblPrijemnica':'PrijemnicaID','tblPrijemnicaStavke':'PrijemnicaStavkaID',
 'tblAmbalaza':'AmbID','tblAmbalazaDokument':'AmbDokID',
 'tblPaleta':'PaletaID','tblPaletaStavka':'StavkaID',
 'tblPrerada':'PreradaID','tblPreradaStavka':'StavkaID',
 'tblFakture':'FakturaID','tblFakturaStavke':'StavkaID',
 'tblUtovar':'UtovarID','tblUtovarStavke':'UtovarStavkaID',
 'tblKooperanti':'KooperantID','tblStanice':'StanicaID','tblKupci':'KupacID',
 'tblVozaci':'VozacID','tblNovac':'NovacID','tblKulture':'KulturaID',
 'tblBankaImport':'BankaImportID','tblStornoVeze':'CorrectionID',
}
REFS=[('tblOtkup','KooperantID','tblKooperanti'),
 ('tblOtkup','StanicaID','tblStanice'),
 ('tblOtpremnica','StanicaID','tblStanice'),
 ('tblPrijemnica','KupacID','tblKupci'),
 ('tblPrijemnica','VozacID','tblVozaci'),
 ('tblNovac','OtkupID','tblOtkup'),('tblNovac','FakturaID','tblFakture'),
 ('tblUtovarStavke','PreradaID','tblPrerada'),
 ('tblOtkupStavke','OtkupID','tblOtkup'),
 ('tblOtpremnicaStavke','OtpremnicaID','tblOtpremnica'),
 ('tblOtpremnicaIzvori','OtpremnicaID','tblOtpremnica'),
 ('tblOtpremnicaIzvori','OtkupID','tblOtkup'),
 ('tblZbirnaStavke','ZbirnaID','tblZbirna'),
 ('tblZbirnaIzvori','ZbirnaID','tblZbirna'),
 ('tblZbirnaIzvori','OtpremnicaID','tblOtpremnica'),
 ('tblPrijemnicaStavke','PrijemnicaID','tblPrijemnica'),
 ('tblPaletaStavka','PaletaID','tblPaleta'),
 ('tblPaletaStavka','PrijemnicaID','tblPrijemnica'),
 ('tblPreradaStavka','PreradaID','tblPrerada'),
 ('tblPreradaStavka','PaletaID','tblPaleta'),
 ('tblFakturaStavke','FakturaID','tblFakture'),
 ('tblFakturaStavke','PrijemnicaID','tblPrijemnica'),
 ('tblUtovarStavke','UtovarID','tblUtovar')]
STORNO={'DA','TRUE','1'}
ISSUED={'IZDATO','PROSLEDJENO','POSLATO'}
SUSPECT_NUMS={'OTP-00581','OTP-00582','OTP-00666','OTK-01110','OTK-01111','ZBR-00666','PRJ-00651'}


def is_void(r):return key(r.get('Stornirano')) in STORNO

def is_active(r):return not is_void(r)

def q(row,col):
    x=dnum(row.get(col))
    return x if x is not None and x.is_finite() else None

def numstr(value):return str(value.normalize()) if isinstance(value,Decimal) else str(value)

def docref(r):
    for c in ('OtkupID','OtpremnicaID','ZbirnaID','PrijemnicaID','AmbID',
              'FakturaID','PaletaID','PreradaID','OtkupStavkaID','OtpremnicaIzvorID',
              'ZbirnaIzvorID','PrijemnicaStavkaID','AmbDokID','PaletaStavkaID','NovacID'):
        if key(r.get(c)):return key(r.get(c))
    return f"red={r.get('_excel_row','?')}"

class Audit:
    def __init__(self,tables):
        self.T=tables;self.findings=[];self.seen=set();self.info={}
    def add(self,level,code,table,row,description,other='',recommendation=''):
        rid=docref(row) if isinstance(row,dict) else norm(row)
        k=(code,table,rid,norm(other))
        if k in self.seen:return
        self.seen.add(k)
        self.findings.append(dict(priority=level,code=code,table=table,
           excel_row=row.get('_excel_row','') if isinstance(row,dict) else '',
           document_id=rid,related_id=norm(other),
           date=ddate(row.get('Datum')) if isinstance(row,dict) else '',
           problem=description,action=recommendation))
    def t(self,name):return self.T.get(name,[])
    def index(self,table,column=None):
        column=column or ID_COL.get(table,'')
        return {key(r.get(column)):r for r in self.t(table) if key(r.get(column))}
    def many(self,table,col):
        out=defaultdict(list)
        for r in self.t(table):
            if key(r.get(col)):out[key(r[col])].append(r)
        return out


def scan_global(A):
    for name,rows in A.T.items():
        col=ID_COL.get(name)
        if not col and rows:
            base=name[3:]
            candidate=base+'ID'
            if candidate in rows[0]:col=candidate
        if not col or not rows or col not in rows[0]:continue
        groups=A.many(name,col)
        for id_,dupes in groups.items():
            if len(dupes)>1:
                A.add('P1','DUPLICATE_PRIMARY_ID',name,dupes[0],
                    f"Identifikator {id_} postoji {len(dupes)} puta (Excel redovi: "+
                    ', '.join(str(x.get('_excel_row')) for x in dupes[:12])+')',id_,
                    'Ne brisati naslepo; dokazati poreklo duplikata i zadrzati originalne ID veze.')
        for r in rows:
            if not key(r.get(col)):
                A.add('P1','EMPTY_PRIMARY_ID',name,r,'Prazan primarni identifikator.',
                      recommendation='Izolovati red i proveriti mapiranje sa originalnim ID-jem.')
    for child,foreign,parent in REFS:
        if not A.t(child) or not A.t(parent):continue
        if foreign not in A.t(child)[0] or ID_COL.get(parent,'') not in A.t(parent)[0]:continue
        parents=A.index(parent)
        for r in A.t(child):
            v=key(r.get(foreign))
            if v and v not in parents:
                A.add('P1','ORPHAN_FOREIGN_KEY',child,r,
                      f'{foreign}={v} ne postoji u {parent}.',v,
                      'Proveriti izvorni roditeljski ID i originalnu vezu, ne izmisljati roditelja.')


def scan_otk_otp(A):
    otks=A.index('tblOtkup');otps=A.index('tblOtpremnica')
    links=A.t('tblOtpremnicaIzvori')
    items=A.many('tblOtkupStavke','OtkupID')
    otp_lines=A.many('tblOtpremnicaStavke','OtpremnicaID')
    by_otp=A.many('tblOtpremnicaIzvori','OtpremnicaID')
    by_otk=A.many('tblOtpremnicaIzvori','OtkupID')
    zbi_otp=A.many('tblZbirnaIzvori','OtpremnicaID');zbr=A.index('tblZbirna')
    for l in links:
        oid=key(l.get('OtkupID'));pid=key(l.get('OtpremnicaID'))
        if not oid or not pid:
            A.add('P1','INCOMPLETE_OTPREMNICA_SOURCE_LINK','tblOtpremnicaIzvori',l,
                  f'Nepotpuna veza: OtkupID={oid}, OtpremnicaID={pid}.')
    sig=defaultdict(list)
    for l in links:
        if key(l.get('OtkupID')) and key(l.get('OtpremnicaID')):
            sig[(key(l.get('OtpremnicaID')),key(l.get('OtkupID')))].append(l)
    for (pid,oid),ll in sig.items():
        if len(ll)>1:A.add('P1','DUPLICATE_OTKUP_OTP_MEMBERSHIP',
            'tblOtpremnicaIzvori',ll[0],f'Veza {pid} -> {oid} se ponavlja {len(ll)} puta.',pid)
    for oid,lnk in by_otk.items():
        if len({key(x.get('OtpremnicaID')) for x in lnk})>1:
            if oid in otks and is_active(otks[oid]):
                pids=sorted({key(x.get('OtpremnicaID')) for x in lnk})
                A.add('P1','ONE_OTKUP_MULTIPLE_OTP','tblOtkup',otks[oid],
                      f'Jedan AKTIVAN otkup je clan vise otpremnica: {pids}.',','.join(pids))
    for oid,otk in otks.items():
        if is_void(otk):continue
        if not items.get(oid):A.add('P1','ACTIVE_OTKUP_NO_LINES','tblOtkup',otk,
             'Aktivan otkup nema tblOtkupStavke; vrednost i kolicina nisu dokazive.')
        oldpid=key(otk.get('OtpremnicaID'))
        member_pids={key(x.get('OtpremnicaID')) for x in by_otk.get(oid,[])}
        if oldpid and oldpid in otps and is_void(otps[oldpid]):
            A.add('P1','ACTIVE_OTKUP_LEGACY_BACKLINK_TO_STORNO_OTP','tblOtkup',otk,
                  'Aktivan otkup u zaglavlju pokazuje na STORNIRANU otpremnicu.',oldpid,
                  'Proveriti stvarni BrojDokumenta/BrojOtpremnice, ne oslanjati se na stari OtpremnicaID.')
        if oldpid and member_pids and oldpid not in member_pids:
            A.add('P2','OTK_BACKLINK_DIFFERS_FROM_CANONICAL_MEMBERSHIP','tblOtkup',otk,
                  f'Legacy OtpremnicaID={oldpid}, kanonska veza={sorted(member_pids)}.',oldpid,
                  'Uporediti oba izvorna dokumenta; kanonsku vezu ne menjati bez dokaza.')
        if member_pids and all(pid in otps and is_void(otps[pid]) for pid in member_pids):
            A.add('P1','ACTIVE_OTKUP_ONLY_STORNO_OTP_LINKS','tblOtkup',otk,
                  'Aktivan otkup ima samo veze ka storniranim otpremnicama.',','.join(sorted(member_pids)))
    by_business=defaultdict(list)
    for o in otks.values():
        business=key(o.get('BrojDokumenta'))
        if business and is_active(o):by_business[business].append(o)
    for pid,otp in otps.items():
        if is_void(otp):continue
        source_links=by_otp.get(pid,[])
        active_sources=[otks[key(l.get('OtkupID'))] for l in source_links
                        if key(l.get('OtkupID')) in otks and is_active(otks[key(l.get('OtkupID'))])]
        child_zbr=[zbr[key(l.get('ZbirnaID'))] for l in zbi_otp.get(pid,[])
                   if key(l.get('ZbirnaID')) in zbr and is_active(zbr[key(l.get('ZbirnaID'))])]
        status=key(otp.get('IzdatoStatus'))
        physically_issued=(bool(zbi_otp.get(pid)) or status in ISSUED or
                           (bool(key(otp.get('BrojZbirne'))) and bool(otp_lines.get(pid))))
        if not source_links or not active_sources:
            kind=('NONE' if not source_links else 'ONLY_STORNED_OR_BROKEN')
            business=key(otp.get('BrojOtpremnice'))
            candidates=[]
            for o in by_business.get(business,[]):
                if (ddate(o.get('Datum'))==ddate(otp.get('Datum')) and
                    key(o.get('StanicaID'))==key(otp.get('StanicaID'))):
                    candidates.append(key(o.get('OtkupID')))
            msg=(f'Aktivna OTP bez AKTIVNOG izvornog otkupa ({kind}); '
                 f'OPI veza={len(source_links)}, vezana za aktivnu ZBR={bool(child_zbr)}, '
                 f'status={status or "PRAZNO"}.')
            if candidates:msg+=f' Kanditati po broju/datumu/stanici: {candidates} (NISU dokazani).'
            A.add('P1' if physically_issued else 'P2','ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP',
                  'tblOtpremnica',otp,msg,','.join(candidates[:8]),
                  'Proveriti OTK zaglavlje/stavke i izvorni BrojDokumenta. Ako je istorijski ponisteno, dokumentovati storno.')
        kg=sum((q(x,'Kolicina') or Decimal(0)) for x in otp_lines.get(pid,[]))
        amb=sum((q(x,'KolAmbalaze') or Decimal(0)) for x in otp_lines.get(pid,[]))
        kg_otk=Decimal(0); amb_otk=Decimal(0)
        for source in active_sources:
            oid=key(source['OtkupID'])
            kg_otk+=sum((q(x,'Kolicina') or Decimal(0)) for x in items.get(oid,[]))
            amb_otk+=sum((q(x,'KolAmbalaze') or Decimal(0)) for x in items.get(oid,[]))
        if active_sources and otp_lines.get(pid) and kg_otk!=kg:
            A.add('P2','OTP_KG_VS_SOURCE_OTKUP_KG','tblOtpremnica',otp,
                  f'OTP stavke={numstr(kg)} kg; povezani aktivni otkupi={numstr(kg_otk)} kg.',
                  ','.join(key(x.get('OtkupID')) for x in active_sources),
                  'Nije automatski kvar ako su podeljena otprema/korekcija; proveriti clanstvo i verzije.')
        if active_sources and otp_lines.get(pid) and amb_otk!=amb:
            A.add('P3','OTP_AMB_VS_SOURCE_OTKUP_AMB','tblOtpremnica',otp,
                  f'OTP stavke={numstr(amb)} gajbica; povezani aktivni otkupi={numstr(amb_otk)}.',
                  recommendation='U MALINA_MODE testirati stvarni broj gajbica i istorijske korekcije.')


def scan_zbr(A):
    zbr=A.index('tblZbirna');otp=A.index('tblOtpremnica')
    zbi=A.t('tblZbirnaIzvori'); by_zbr=A.many('tblZbirnaIzvori','ZbirnaID')
    by_otp=A.many('tblZbirnaIzvori','OtpremnicaID')
    lines_z=A.many('tblZbirnaStavke','ZbirnaID')
    lines_o=A.many('tblOtpremnicaStavke','OtpremnicaID')
    sig=defaultdict(list)
    for r in zbi:
        x=(key(r.get('ZbirnaID')),key(r.get('OtpremnicaID')))
        if x[0] and x[1]:sig[x].append(r)
    for x,ls in sig.items():
        if len(ls)>1:A.add('P1','DUPLICATE_ZBR_OTP_MEMBERSHIP','tblZbirnaIzvori',ls[0],
              f'Ponovljeno clanstvo ZBR={x[0]} OTP={x[1]} x{len(ls)}.',x[1])
    for oid,members in by_otp.items():
        active=[x for x in members if key(x.get('ZbirnaID')) in zbr and is_active(zbr[key(x.get('ZbirnaID'))])]
        if len(active)>1:A.add('P1','OTP_IN_MULTIPLE_ACTIVE_ZBR','tblZbirnaIzvori',active[0],
                             f'Otpremnica je u {len(active)} aktivnih zbirnih.',oid)
    for zid,z in zbr.items():
        if is_void(z):continue
        membership=by_zbr.get(zid,[])
        active=[];cancelled=[]
        for x in membership:
            pid=key(x.get('OtpremnicaID'))
            if pid not in otp:continue
            if is_void(otp[pid]):cancelled.append(pid)
            else:active.append(otp[pid])
        if cancelled:A.add('P1','ACTIVE_ZBR_CONTAINS_STORNO_OTP','tblZbirna',z,
                           f'Aktivna ZBR ima stornirane OTP izvore: {cancelled}',','.join(cancelled))
        if not active:
            A.add('P1' if lines_z.get(zid) else 'P2','ACTIVE_ZBR_NO_ACTIVE_OTP_SOURCE',
                  'tblZbirna',z,f'Aktivna zbirna bez AKTIVNE otpremnice; linkova ukupno={len(membership)}.',
                  recommendation='Proveriti storno/korekciju i tblZbirnaIzvori.')
        kg_z=sum((q(x,'Kolicina') or Decimal(0)) for x in lines_z.get(zid,[]))
        # Canonical cols in ZBRStavke may use UkupnoKolicina, never assume.
        if lines_z.get(zid) and 'Kolicina' not in lines_z[zid][0]:
            kg_z=sum((q(x,'UkupnoKolicina') or Decimal(0)) for x in lines_z[zid])
        kg_o=sum((q(x,'Kolicina') or Decimal(0)) for source in active
                 for x in lines_o.get(key(source.get('OtpremnicaID')),[]))
        if lines_z.get(zid) and active and kg_z!=kg_o:
            A.add('P2','ZBR_KG_DIFF_VS_MEMBER_OTPS','tblZbirna',z,
                  f'ZBR stavke={numstr(kg_z)} kg; aktivne OTP stavke={numstr(kg_o)} kg.',
                  recommendation='Proveriti stavke i clanstvo, ne menjati iznose automatski.')
        if lines_z.get(zid) and q(z,'UkupnoKolicina') is not None and q(z,'UkupnoKolicina')!=kg_z:
            A.add('P3','ZBR_HEADER_VS_LINE_KG','tblZbirna',z,
                  f'Header={numstr(q(z,"UkupnoKolicina"))} kg, stavke={numstr(kg_z)} kg.',
                  recommendation='Ako je kanonska istina u stavkama, ne prepisivati header napamet.')


def scan_prj(A):
    zbr=A.index('tblZbirna'); prj=A.t('tblPrijemnica')
    zbr_no=defaultdict(list)
    for r in zbr.values():
        n=key(r.get('BrojZbirne'))
        if n:zbr_no[n].append(r)
    prj_items=A.many('tblPrijemnicaStavke','PrijemnicaID')
    for p in prj:
        znum=norm(p.get('BrojZbirne')); k=key(znum)
        if not k:continue
        cand=zbr_no.get(k,[])
        active=[c for c in cand if is_active(c)]
        if not cand and is_active(p):
            A.add('P1','ACTIVE_PRJ_NO_ZBR_NUMBER','tblPrijemnica',p,
                  f'Ne postoji zbirna za BrojZbirne={znum!r}.',znum,
                  'Proveriti case, whitespace, ID i istorijski poslovni broj.')
        if len(cand)>1 and is_active(p):
            A.add('P2','PRJ_ZBR_NUMBER_AMBIGUOUS','tblPrijemnica',p,
                  f'Poslovni broj {znum} ima {len(cand)} ZBR zaglavlja; nije jedinstven.',
                  ','.join(key(x.get('ZbirnaID')) for x in cand),
                  'Odrediti cilj preko dokumentnog identiteta, ne samo poslovnog broja.')
        if cand and is_active(p) and not active:
            A.add('P1','ACTIVE_PRJ_ONLY_STORNO_ZBR','tblPrijemnica',p,
                  f'Aktivna prijemnica ukazuje samo na stornirane zbirne broj={znum}.',
                  ','.join(key(x.get('ZbirnaID')) for x in cand))
        exact=[c for c in cand if norm(c.get('BrojZbirne'))==znum]
        if cand and not exact and is_active(p):
            A.add('P2','PRJ_ZBR_CASE_ONLY_OR_FORMAT','tblPrijemnica',p,
                  f'Prijemnica BrojZbirne={znum!r}, ZBR vrednosti={[norm(x.get("BrojZbirne")) for x in cand]}.',
                  recommendation='Normalizovati poredjenje (UPPER/TRIM); originalne brojeve sacuvati.')
        if len(active)==1 and is_active(p):
            parent=active[0]
            if ddate(parent.get('Datum')) and ddate(p.get('Datum')) and ddate(parent.get('Datum'))!=ddate(p.get('Datum')):
                A.add('P3','PRJ_ZBR_DATE_DIFF','tblPrijemnica',p,
                      f'PRJ datum={ddate(p.get("Datum"))}, ZBR datum={ddate(parent.get("Datum"))}.',key(parent.get('ZbirnaID')))
        if is_void(p) and active:
            A.add('INFO','STORNO_PRJ_WITH_ACTIVE_ZBR','tblPrijemnica',p,
                  'Stornirana prijemnica uz aktivnu zbirnu; stari VBA je dozvoljavao pojedinacni storno.',
                  key(active[0].get('ZbirnaID')),
                  'Samo informacija: nije automatski kvar niti automatski nalog za storno ZBR.')
        pid=key(p.get('PrijemnicaID'))
        if is_active(p) and prj_items.get(pid):
            qty=sum((q(x,'Kolicina') or Decimal(0)) for x in prj_items[pid])
            header=q(p,'Kolicina')
            if header is not None and qty!=header:
                A.add('P3','PRJ_HEADER_VS_LINES_KG','tblPrijemnica',p,
                      f'Header {numstr(header)} kg != kanonske stavke {numstr(qty)} kg.')
    # Archived storned PRJ can be physically re-used by paleta rows; do not auto detach.
    pal=A.t('tblPaletaStavka');pidx=A.index('tblPrijemnica')
    for x in pal:
        pid=key(x.get('PrijemnicaID'))
        if pid in pidx and is_void(pidx[pid]) and is_active(x):
            A.add('P2','ACTIVE_PALETA_ITEM_FROM_STORNO_PRJ','tblPaletaStavka',x,
                  f'Aktivna paletna stavka upucuje na storniranu prijemnicu {pid}.',pid,
                  'Proveriti da li je paleta stvarno razduzena ili rezervisana za rucno prevezivanje.')


def scan_amb(A):
    ledger=A.t('tblAmbalaza')
    idx=A.index('tblAmbalaza');otk=A.index('tblOtkup');otp=A.index('tblOtpremnica');prj=A.index('tblPrijemnica')
    ambit=A.index('tblAmbalazaDokument')
    inverse=A.many('tblAmbalaza','StornoOd')
    for oid,inv in inverse.items():
        if len(inv)>1:
            A.add('P1','MULTIPLE_AMB_INVERSES','tblAmbalaza',inv[0],
                  f'Jedan original AmbID={oid} ima {len(inv)} kontra-stavova.',oid)
        parent=idx.get(oid)
        for row in inv:
            if parent is None:continue # reported by foreign check below
            cols=[('OdNalogTip','NaNalogTip'),('OdNalogID','NaNalogID'),('NaNalogTip','OdNalogTip'),('NaNalogID','OdNalogID')]
            reversed_ok=all(key(row.get(a))==key(parent.get(b)) for a,b in cols)
            sameqty=q(row,'Kolicina')==q(parent,'Kolicina')
            sametype=key(row.get('TipAmbalaze'))==key(parent.get('TipAmbalaze'))
            if not (reversed_ok and sameqty and sametype):
                A.add('P1','BAD_AMB_INVERSE','tblAmbalaza',row,
                      f'Kontra-stav nije exact inverse originala {oid}.',oid)
    for r in ledger:
        ambid=key(r.get('AmbID'))
        old=key(r.get('StornoOd'))
        if old and old not in idx:
            A.add('P1','AMB_STORNO_TARGET_NOT_FOUND','tblAmbalaza',r,
                  f'StornoOd={old} ne pokazuje na postojeci AmbID.',old)
        if not old:
            if (key(r.get('OdNalogTip')),key(r.get('OdNalogID')))==(key(r.get('NaNalogTip')),key(r.get('NaNalogID'))):
                A.add('P1','AMB_SAME_ACCOUNT_BOTH_SIDES','tblAmbalaza',r,
                      'Prenos sa istog naloga na isti nalog nema saldo efekat.')
            dt=key(r.get('DokumentTIP'));di=key(r.get('DokumentID'))
            parents={'OTKUP':otk,'OTPREMNICA':otp,'PRIJEMNICA':prj,'AMBALAZADOKUMENT':ambit}
            if dt in parents:
                parent=parents[dt].get(di)
                if parent is None:
                    A.add('P1','AMB_ORPHAN_BUSINESS_DOCUMENT','tblAmbalaza',r,
                          f'Ledger pokazuje na nepostojeci dokument {dt} {di}.',di)
                elif is_void(parent) and not inverse.get(ambid):
                    A.add('P2','STORNO_DOCUMENT_WITH_UNREVERSED_AMB','tblAmbalaza',r,
                          f'Roditelj {dt} {di} je storniran, a AMB red nema kontra-stav.',di,
                          'Proveriti migriranu istoriju pre korekcije; ne dodavati proizvoljan storno.')
        if q(r,'Kolicina') is None or q(r,'Kolicina')<=0:
            A.add('P1','AMB_NONPOSITIVE_OR_INVALID_KOLICINA','tblAmbalaza',r,
                  f'Neispravna kolicina={r.get("Kolicina")!r}.')
    for rid in ('AMB-01377','AMB-07115','AMB-07815'):
        # Their source is unknown and archived in MIG19_Nepoznati, intentionally not in live ledger.
        if rid in idx:
            A.add('P2','UNKNOWN_SOURCE_DIRECTLY_IMPORTED','tblAmbalaza',idx[rid],
                  'Jednostrani istorijski prijem nepoznatog porekla nalazi se u kanonskoj knjizi.',rid)
    # ST-00001 has accepted historic negative; do not emit negative-balance P1.
    A.info['malina_station1_negative_allowed']=True
    transfers=[('02317','2026-06-27','ST-00005',Decimal('300')),
               ('03849','2026-06-28','ST-00005',Decimal('250')),
               ('03939','2026-06-29','ST-00002',Decimal('83'))]
    for suffix,day,destination,qty in transfers:
        expected_id='AMB-MIG19-'+suffix
        matches=[x for x in ledger if key(x.get('AmbID'))==expected_id]
        if len(matches)!=1:
            A.add('P1','MIG19_TRANSFER_ABSENT_OR_DUPLICATE','tblAmbalaza',expected_id,
                  f'Potvrdeni PRENOS_INTERNO {expected_id} nije tacno jedan red (ima {len(matches)}).',
                  destination,'Proveriti da je MIG19.13 izvrsen i da su append redovi sacuvani.')
            continue
        r=matches[0]
        valid=(key(r.get('VrstaKretanja'))=='PRENOS_INTERNO' and
               key(r.get('OdNalogTip'))=='STANICA' and
               key(r.get('OdNalogID'))=='ST-00001' and
               key(r.get('NaNalogTip'))=='STANICA' and
               key(r.get('NaNalogID'))==destination and
               q(r,'Kolicina')==qty and ddate(r.get('Datum'))==day)
        if not valid:
            A.add('P1','MIG19_TRANSFER_SIGNATURE_CHANGED','tblAmbalaza',r,
                  f'Ocekivano ST-00001 -> {destination}, {numstr(qty)} na {day}.',destination)
    A.info['unknown_ST1_receipts_expected_archived']=['AMB-01377','AMB-07115','AMB-07815']


# Ova polja su bila poslovne cinjenice u direktnom staging-u.
# Ne poredimo izvedene canonical kolone, runtime/created audit timestampove, niti
# zaglavlja koja su prebacena u stavke i zamenjena canonical source-of-truth.
STABLE_DIRECT_FIELDS={
 'tblOtkup':['Datum','KooperantID','StanicaID','BrojDokumenta','VrstaVoca','SortaVoca','Stornirano'],
 'tblOtpremnica':['Datum','StanicaID','BrojOtpremnice','VrstaVoca','SortaVoca','Stornirano'],
 'tblZbirna':['Datum','VozacID','BrojZbirne','VrstaVoca','SortaVoca','Stornirano'],
 'tblPrijemnica':['Datum','KupacID','BrojPrijemnice','BrojZbirne','VrstaVoca','SortaVoca','Stornirano'],
 'tblPaleta':['PrijemnicaID','Datum','TipPalete','Stornirano'],
 'tblPrerada':['Datum','Stornirano'],
}

def field_compare(col,a,b):
    # Legacy datetime serial -> canonical date: compare day only.
    if col.lower().startswith('datum'):
        return ddate(a)==ddate(b)
    if col in {'Kolicina','Cena','KolAmbalaze','BrutoKg'}:
        da,db=dnum(a),dnum(b)
        return da==db if da is not None and db is not None else norm(a)==norm(b)
    # For source authenticity, case differences are separately visible as exact field changes.
    return norm(a)==norm(b)

def scan_stage_counts(A,stage):
    """Source rows and IDs coverage when corresponding <table>.jsonl exists.
    Deliberately conservative: no blanket assertion that all legacy config tables
    must be copied to a new runtime workbook.
    """
    s=Path(stage)
    if not s.exists():
        A.info['stage_status']='MISSING (audit still ran on XLSM; source-vs-target not verified)'
        return
    A.info['stage_status']='PRESENT'
    expected={}
    for fp in s.rglob('*.jsonl'):
        name=fp.stem
        if name.startswith(('SOURCE__','UNCONVERTED__','SECONDARY__')):continue
        if name not in A.T:continue
        pk=ID_COL.get(name)
        count=0; ids=set(); empties=0; source_rows={}
        try:
            with fp.open('r',encoding='utf-8-sig') as fh:
                for lineno,line in enumerate(fh,1):
                    if not line.strip():continue
                    val=json.loads(line)
                    count+=1
                    if pk and isinstance(val,dict):
                        if key(val.get(pk)):
                            ids.add(key(val.get(pk)))
                            source_rows[key(val.get(pk))]=val
                        else:empties+=1
        except (ValueError,UnicodeError,OSError) as exc:
            A.add('P2','STAGE_SOURCE_UNREADABLE',name,fp.name,
                  f'Staging JSONL ne moze da se ucita: {exc!s}')
            continue
        target=A.t(name)
        expected[name]=(count,len(target))
        if count!=len(target):
            A.add('P2','STAGE_TARGET_ROWCOUNT_DIFFERENCE',name,fp.name,
                  f'Stage {count} redova, QA {len(target)} redova; proveriti namerne filtracije.',
                  recommendation='Ovo je broj redova, ne dokaz tacnosti polja ili kompletne relacije.')
        if pk and ids and target and pk in target[0]:
            got={key(r.get(pk)) for r in target if key(r.get(pk))}
            miss=sorted(ids-got)
            added=sorted(got-ids)
            if miss:A.add('P2','STAGE_SOURCE_IDS_MISSING_IN_QA',name,fp.name,
                          f'Nedostaje {len(miss)} stage ID-jeva u QA; primeri: {miss[:12]}',
                          recommendation='Razdvojiti legitimne transformacije od izgubljenih redova.')
            if added:A.add('P3','QA_IDS_NOT_PRESENT_IN_DIRECT_STAGE',name,fp.name,
                           f'QA ima {len(added)} dodatnih ID-jeva; primeri: {added[:12]}')
            if name in STABLE_DIRECT_FIELDS and source_rows:
                target_by_id={key(r.get(pk)):r for r in target}
                changes=Counter(); samples=defaultdict(list)
                for source_id,srow in source_rows.items():
                    trow=target_by_id.get(source_id)
                    if not trow:continue
                    for col in STABLE_DIRECT_FIELDS[name]:
                        if col not in srow or col not in trow:continue
                        if not field_compare(col,srow[col],trow[col]):
                            changes[col]+=1
                            if len(samples[col])<7:
                                samples[col].append((source_id,norm(srow[col]),norm(trow[col])))
                for col,n in changes.items():
                    A.add('P2' if col not in {'Stornirano','BrojZbirne'} else 'P3',
                          'STAGE_STABLE_FIELD_CHANGED',name,fp.name,
                          f'{col}: {n} razlicitih vrednosti stage / QA. Primeri: {samples[col]}',
                          col,
                          'Proveriti da li je izmena namerna QA korekcija; identicni ID-jevi nisu dokaz identicnog polja.')
    A.info['stage_direct_table_counts']=expected
    # Sekundarni staging nije dokaz da je uvezen; eksplicitno ga popisi.
    secondary=[]
    for prefix in ('SECONDARY__','UNCONVERTED__'):
        for fp in s.rglob(prefix+'*.jsonl'):
            name=fp.stem[len(prefix):]
            if name not in {'tblBankaImport','tblPartnerMap','tblStornoVeze','tblNovac'}:continue
            with fp.open('r',encoding='utf-8-sig') as fh:
                source_count=sum(1 for line in fh if line.strip())
            target_count=len(A.t(name)) if name in A.T else None
            secondary.append({'stage_file':fp.name,'source_rows':source_count,
                              'target_table':name,'target_rows':target_count})
            if target_count is None:
                A.add('P2','SECONDARY_STAGE_TARGET_TABLE_ABSENT',name,fp.name,
                      f'Sekundarna istorijska tabela {name} nije pronadena u QA XLSM; stage ima {source_count} redova.')
            elif source_count!=target_count:
                A.add('P2','SECONDARY_STAGE_NOT_FULLY_IMPORTED',name,fp.name,
                      f'Sekundarni stage {source_count} redova, QA {target_count}. '
                      'Ovo moze biti namerno, ali znaci da nismo dokazali pun import.')
    A.info['secondary_stage_inventory']=secondary
    A.info['stage_coverage_note']='Samo postojeci direktni JSONL fajlovi. Transformacije, sekundarne tabele i svaki pojedinacni atribut zahtevaju zasebnu semanticku proveru.'


def detect_cross_link(A):
    """Identify suspicious chain patterns, including historical PRJ-storno/OTP-no-source."""
    zbr=A.index('tblZbirna');otp=A.index('tblOtpremnica');prj=A.t('tblPrijemnica')
    by_otp=A.many('tblOtpremnicaIzvori','OtpremnicaID')
    zbi=A.many('tblZbirnaIzvori','ZbirnaID')
    p_by_no=defaultdict(list)
    for p in prj:
        if key(p.get('BrojZbirne')):p_by_no[key(p.get('BrojZbirne'))].append(p)
    for zid,z in zbr.items():
        if is_void(z):continue
        members=[otp.get(key(x.get('OtpremnicaID'))) for x in zbi.get(zid,[])]
        members=[m for m in members if m is not None and is_active(m)]
        if not members:continue
        if all(not by_otp.get(key(o.get('OtpremnicaID'))) for o in members):
            storned_prj=[p for p in p_by_no.get(key(z.get('BrojZbirne')),[]) if is_void(p)]
            if storned_prj:
                A.add('P1','ACTIVE_CHAIN_NO_OTKUPS_AND_CANCELLED_PRJ','tblZbirna',z,
                      f'Aktivna zbirna sa {len(members)} OTP bez izvora; {len(storned_prj)} povezane stornirane PRJ.',
                      ','.join(key(x.get('OtpremnicaID')) for x in members),
                      'Tip OTP-00666: proveriti opravdanost zatvaranja lanca, bez izmisljanja otkupa.')


def choose_workbook(explicit):
    if explicit:return Path(explicit)
    paths=[Path('../MigrationMIG19_13_PatchedXLSM/AgriX_MIG19_13_AMBALAZA_QA_PARTIAL.xlsm')]
    possible=list(Path('../MigrationMIG19_13_PatchedXLSM').glob('AgriX_MIG19_13*.xlsm')) if Path('../MigrationMIG19_13_PatchedXLSM').exists() else []
    possible+=list(Path('../').glob('AgriX_MIG19_13*.xlsm'))
    available=sorted({x.resolve() for x in paths+possible if x.is_file()})
    if not available:
        raise FileNotFoundError('Nije nadjena QA XLSM. Prosledi --workbook "../putanja/do/backup_sveske.xlsm"')
    if len(available)>1:
        print('UPOZORENJE: nadjeno vise QA sveski; koristim imenovanu polaznu ako postoji. '
              'Za izmene koje si rucno uradio, OBAVEZNO prosledi --workbook:',file=sys.stderr)
        for p in available:print('  '+str(p),file=sys.stderr)
    if paths[0].is_file():return paths[0]
    return max(available,key=lambda p:p.stat().st_mtime)


def run(workbook,stage):
    workbook=Path(workbook)
    if not workbook.is_file():raise FileNotFoundError(f'XLSM ne postoji: {workbook}')
    ro=XlsmReadOnly(workbook)
    try:
        tables={}
        for name in sorted(ro.tables):
            try:tables[name]=ro.read_table(name,set())
            except Exception as exc:
                # Incomplete schema/read must be visible as a BLOCKER, not hidden.
                tables[name]=[]
                print('READ ERROR',name,repr(exc),file=sys.stderr)
                tables.setdefault('__READ_ERRORS__',[]).append({'table':name,'error':str(exc)})
        A=Audit({k:v for k,v in tables.items() if not k.startswith('__')})
        A.info['tables_in_xlsm']=len(ro.tables)
        A.info['table_rows_total']=sum(len(v) for k,v in A.T.items())
        A.info['tables_rows']={k:len(v) for k,v in sorted(A.T.items())}
        A.info['source_workbook_sha256']=sha256(workbook)
        A.info['source_workbook_path']=str(workbook.resolve())
        A.info['read_errors']=tables.get('__READ_ERRORS__',[])
        for e in A.info['read_errors']:
            A.add('P1','TABLE_UNREADABLE',e['table'],e['table'],e['error'])
    finally:ro.close()
    scan_global(A)
    scan_otk_otp(A)
    scan_zbr(A)
    scan_prj(A)
    scan_amb(A)
    detect_cross_link(A)
    scan_stage_counts(A,stage)
    A.findings.sort(key=lambda x:({'P1':0,'P2':1,'P3':2,'INFO':3}.get(x['priority'],9),x['code'],x['document_id']))
    counts=Counter(x['priority'] for x in A.findings)
    codes=Counter(x['code'] for x in A.findings)
    hot=[x for x in A.findings if x['document_id'] in SUSPECT_NUMS or
         any(id_ in x['related_id'] for id_ in SUSPECT_NUMS)]
    result={
      'format':VERSION,'status':'READ_ONLY_FINDINGS_NOT_AUTO_VERDICTS',
      'input':{'workbook':A.info['source_workbook_path'],'sha256':A.info['source_workbook_sha256'],
               'stage':str(stage)},
      'coverage':{'tables_read':A.info['tables_in_xlsm'],'rows_read':A.info['table_rows_total'],
                  'table_row_counts':A.info['tables_rows'],'stage_status':A.info.get('stage_status'),
                  'stage_direct_table_counts':A.info.get('stage_direct_table_counts',{}),
                  'secondary_stage_inventory':A.info.get('secondary_stage_inventory',[]),
                  'read_errors':A.info['read_errors']},
      'findings_summary':{'total':len(A.findings),'by_priority':dict(counts),
                          'by_code':dict(sorted(codes.items()))},
      'hot_document_findings':hot[:100],
      'sample_top_priority':A.findings[:35],
      'policy':{'ST_00001_historical_negative_allowed':True,
               'station_equals_driver_mirror_in_MALINA_MODE':True,
               'unknown_ST1_receipts_not_fabricated':[5,10,3],
               'negative_balance_alone_is_not_migration_corruption':True,
               'no_auto_repair':True},
      'warning':'Provera ID-jeva, clansta i salda ne dokazuje potpuno paritet SVAKOG polja. Uzorci po dokumentima zahtevaju proveru kada ih audit izdvoji.'}
    return result,A.findings


def save(out,summary,findings):
    out=Path(out);out.mkdir(parents=True,exist_ok=True)
    with (out/'MIG20_SUMMARY.json').open('w',encoding='utf-8') as f:
        json.dump(summary,f,ensure_ascii=False,indent=2)
    cols=['priority','code','table','excel_row','document_id','related_id','date','problem','action']
    with (out/'MIG20_ANOMALIJE.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=cols,delimiter=';');w.writeheader();w.writerows(findings)
    with (out/'MIG20_PRIORITETI.txt').open('w',encoding='utf-8') as f:
        f.write('MIG20 — read-only, puna sveska; NIJEDNA izmena XLSM\n')
        f.write(f"XLSM: {summary['input']['workbook']}\nSHA256: {summary['input']['sha256']}\n")
        f.write(f"TABELE: {summary['coverage']['tables_read']} | REDOVI: {summary['coverage']['rows_read']}\n")
        f.write(f"NALAZI: {summary['findings_summary']['by_priority']}\n\n")
        for x in findings:
            f.write(f"[{x['priority']}] {x['code']} {x['table']} {x['document_id']} row={x['excel_row']} other={x['related_id']}\n")
            f.write('   '+x['problem']+'\n')
            if x['action']:f.write('   AKCIJA: '+x['action']+'\n')
    return [out/n for n in ['MIG20_SUMMARY.json','MIG20_ANOMALIJE.csv','MIG20_PRIORITETI.txt']]


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--workbook',help='Exact .xlsm file to audit. Read-only.')
    ap.add_argument('--stage',default='../MigrationStage_v2')
    ap.add_argument('--out-dir',default='../MigrationMIG20_FullAudit')
    args=ap.parse_args()
    book=choose_workbook(args.workbook)
    report,findings=run(book,args.stage)
    files=save(args.out_dir,report,findings)
    print(json.dumps({k:v for k,v in report.items() if k not in ('sample_top_priority','coverage','hot_document_findings')},ensure_ascii=False,indent=2))
    print('TABELE:',report['coverage']['tables_read'],'REDOVI:',report['coverage']['rows_read'])
    print('TOP:',[(x['priority'],x['code'],x['document_id'],x['related_id']) for x in report['sample_top_priority'][:20]])
    print('IZVESTAJI:',*[str(x) for x in files],sep='\n  ')
    if report['coverage']['read_errors']:sys.exit(2)

if __name__=='__main__':main()
