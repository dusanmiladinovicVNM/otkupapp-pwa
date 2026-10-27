#!/usr/bin/env python3
"""Read-only provenance check for KOOP-00019; no Excel COM / no workbook edits.

Reuses the XLSM parser already shipped in MIG_ARCHIVE_XLSM_v1.3.
"""
import argparse
import sys
from pathlib import Path
from collections import defaultdict


def txt(x):
    return '' if x is None else str(x).strip()


def run(workbook: Path, parser: Path):
    sys.path.insert(0, str(parser.resolve()))
    from agrix_mig_archive import read_workbook
    t=read_workbook(str(workbook))
    coop='KOOP-00019'
    r=t['tblAmbalaza']
    good=[]
    for x in r:
        odT,odI,naT,naI=[txt(x.get(k)) for k in ('OdNalogTip','OdNalogID','NaNalogTip','NaNalogID')]
        if not ((odT.casefold()=='kooperant' and odI==coop) or (naT.casefold()=='kooperant' and naI==coop)):
            continue
        dat=txt(x.get('Datum'))
        if '2026' not in dat: continue
        if '07-08' not in dat and '07-11' not in dat and '8.7.2026' not in dat and '11.7.2026' not in dat and '08.07.2026' not in dat and '11.07.2026' not in dat:
            continue
        qty=float(x.get('Kolicina') or 0)
        s=qty if naT.casefold()=='kooperant' and naI==coop else -qty
        tech=(txt(x.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE')
        good.append((dat,txt(x.get('DokumentID')),txt(x.get('AmbID')),
                     txt(x.get('VrstaKretanja')),s,0 if tech else s,txt(x.get('StornoOd'))))
    print('Read-only ledger rows related to KOOP-00019 on Jul-08 and Jul-11:')
    print('Datum | Document | AmbID | VrstaKretanja | physical | business | StornoOd')
    for x in sorted(good): print(' | '.join(map(str,x)))
    if not good:
        print('NO MATCHES - inspect date encoding and workbook path; no business claims made.')
    vals=defaultdict(float)
    for dat,dok,aid,typ,phy,biz,sto in good: vals[(dat,'physical')]+=phy; vals[(dat,'business')]+=biz
    print('Daily totals (physical vs business, for source ledger row subset only):')
    for k,v in sorted(vals.items()): print(k,v)
    match=[x for x in good if x[1]=='OTK-01472' and x[3]=='AMBALAZA_UZ_ROBU' and x[5]<0]
    if match:
        print('CHECK OTK-01472 PHYSICAL OUTGOING:',sum(x[5] for x in match))
    else:
        print('CHECK OTK-01472: MISSING physical outgoing (-3 expected from screenshot)')


if __name__=='__main__':
    ap=argparse.ArgumentParser()
    ap.add_argument('--workbook',required=True,type=Path)
    ap.add_argument('--parser-dir',required=True,type=Path)
    a=ap.parse_args()
    run(a.workbook,a.parser_dir)
