#!/usr/bin/env python3
"""AGRIx MIG-19.2: read-only diagnosis of MIG-19.1 output and Malina mirror balances.

Requires agrix_mig19_malina_readonly_audit.py in same directory (MIG-19.1 parser).
Does not write to a workbook, does not change business ledgers, never fabricates entries.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
import csv
from decimal import Decimal
import hashlib
import json
from pathlib import Path
from agrix_mig19_malina_readonly_audit import XlsmReadOnly, dnum, key, ddate


def readcsv(path):
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def sha256(path):
    h = hashlib.sha256()
    with path.open('rb') as f:
        for buf in iter(lambda: f.read(1048576), b''):
            h.update(buf)
    return h.hexdigest()


def jdump(obj):
    return json.dumps(obj, indent=2, ensure_ascii=False)


def prepare(args):
    output = args.audit_out
    for name in ('summary.json', 'findings.csv', 'all_saldos.csv', 'revers_ledger_rows.csv'):
        if not (output / name).is_file():
            raise ValueError(f'Cannot find {output / name}; use MIG-19.1 --out-dir')
    previous = json.loads((output / 'summary.json').read_text(encoding='utf-8-sig'))
    findings = readcsv(output / 'findings.csv')
    saldos = readcsv(output / 'all_saldos.csv')
    reverses = readcsv(output / 'revers_ledger_rows.csv')
    workbook = args.workbook or Path(previous['workbook'])
    if not workbook.is_file():
        raise ValueError(f'Missing workbook: {workbook}. Supply --workbook path.')
    before = sha256(workbook)
    if previous.get('sha256_after') and before != previous['sha256_after']:
        raise ValueError('Workbook SHA-256 differs from MIG-19.1 output; refuse mixed snapshots')
    reader = XlsmReadOnly(workbook)
    try:
        stations = reader.read_table('tblStanice', {'StanicaID'})
        ledger = reader.read_table('tblAmbalaza', {
            'AmbID', 'DokumentID', 'DokumentTIP', 'Datum', 'TipAmbalaze', 'Kolicina',
            'OdNalogTip', 'OdNalogID', 'NaNalogTip', 'NaNalogID', 'VrstaKretanja', 'StornoOd'})
        docs = reader.read_table('tblAmbalazaDokument', {'AmbDokID', 'Vrsta'})
    finally:
        reader.close()
    if sha256(workbook) != before:
        raise ValueError('Workbook changed while reading')
    if len(ledger) != previous['counts']['ledger_rows']:
        raise ValueError('Workbook ledger count differs from MIG-19.1 results')
    station_ids = {key(r['StanicaID']) for r in stations if key(r['StanicaID'])}
    if 'ST-00001' not in station_ids:
        raise ValueError('ST-00001 does not exist in tblStanice; cannot apply Malina mirror analysis')
    revers_ids = {key(x['AmbDokID']) for x in docs if key(x['Vrsta']) in ('REVERS','REVERS_PARTNERA')}
    # P2 counts, 3 examples each; do not mistake one P2 finding for one business document.
    p2 = [f for f in findings if key(f['severity']) == 'P2']
    bycode = Counter(key(f['code']) for f in findings)
    codeinfo = []
    for code, n in sorted(bycode.items(), key=lambda v:(-v[1],v[0])):
        matches = [f for f in findings if key(f['code']) == code]
        codeinfo.append({
            'code':code, 'count':n,
            'sample_AmbIDs':[x['AmbID'] for x in matches[:3]],
            'sample_DocumentIDs':[x['AmbDokID'] for x in matches[:3]]})
    # Raw station balances + mirror driver's balances, by exact mirrored StationID and packaging type.
    vals = defaultdict(Decimal)
    for s in saldos:
        # MIG-19.1 Excel-safe CSV export prefixes a literal apostrophe before negatives.
        q = dnum(s['Saldo'][1:] if s['Saldo'].startswith("'-") else s['Saldo'])
        if q is not None:
            vals[(key(s['NalogTip']),key(s['NalogID']),key(s['TipAmbalaze']))] += q
    account_types = sorted({tp for (t,id,tp) in vals if id in station_ids and t in ('STANICA','VOZAC')})
    combined = []
    for station_id in sorted(station_ids):
        for typ in account_types:
            station = vals[('STANICA',station_id,typ)]
            driver = vals[('VOZAC',station_id,typ)]
            if not station and not driver: continue
            combined.append({'StationID':station_id, 'TipAmbalaze':typ,
                'StanicaSaldo':str(station),'MirrorVozacSaldo':str(driver),
                'CombinedPhysicalCandidate':str(station+driver),
                'CombinedNegative':'YES' if station+driver < 0 else 'NO',
                'CentralNegativeException':'YES' if station_id=='ST-00001' else 'NO'})
    # Classify mirror-pair ledger rows by source document, direction, and movement.
    mirror_rows=[]; by_doc=Counter(); by_dir=Counter(); by_kind=Counter(); by_station=Counter(); amounts=defaultdict(Decimal)
    for r in ledger:
        a,b=key(r['OdNalogTip']),key(r['NaNalogTip'])
        sa,sb=key(r['OdNalogID']),key(r['NaNalogID'])
        if {a,b}!={'STANICA','VOZAC'} or sa != sb or sa not in station_ids:
            continue
        direction = f'{a}->{b}'
        dt=key(r['DokumentTIP']); kind=key(r['VrstaKretanja'])
        mirror_rows.append(r)
        by_doc[dt]+=1; by_dir[direction]+=1;by_kind[kind]+=1;by_station[sa]+=1
        q=dnum(r['Kolicina']) or Decimal(0)
        amounts[(sa, key(r['TipAmbalaze']),direction)] += q
    # Reverse findings can be identified by exact referenced AmbID or AmbDokID.
    rev_ambids = {key(r['AmbID']) for r in reverses}
    revers_findings = [f for f in findings if key(f['AmbID']) in rev_ambids or
         key(f['AmbDokID']) in revers_ids]
    diag = {
        'gate':'TRIAGE_ONLY_NOT_FINAL_PASS',
        'source_workbook_sha256':before,
        'mig19_1_gate':previous['gate'],
        'audit_findings_total':len(findings),
        'audit_findings_p2':len(p2),
        'findings_by_code':codeinfo,
        'revers_findings_count':len(revers_findings),
        'revers_findings_by_code':dict(Counter(key(f['code']) for f in revers_findings)),
        'mirror_ledger_row_count':len(mirror_rows),
        'mirror_rows_by_document_type':dict(by_doc),
        'mirror_rows_by_direction':dict(by_dir),
        'mirror_rows_by_kind':dict(by_kind),
        'mirror_rows_by_station':dict(by_station),
        'station_types_negative_raw':sum(vals[('STANICA',sid,tp)]<0 for sid in station_ids for tp in account_types),
        'station_types_negative_combined':sum(r['CombinedNegative']=='YES' for r in combined),
        'station_types_negative_combined_outside_central':sum(
            r['CombinedNegative']=='YES' and r['CentralNegativeException']!='YES' for r in combined),
        'combined_balances':combined,
        'scope_note':('CombinedPhysicalCandidate is a diagnostic view, NOT a new canonical ledger balance. '
                       'It treats station and driver with identical ID as one physical holder only for MALINA_MODE. '
                       'It does not prove historical correctness or authorize rewriting ledger rows.'),
        'source_verification_note':'No legacy-to-canonical lineage is reconstructed from approximate date/quantity matches.'
    }
    # Print summary without all rows, revealing only technical IDs.
    print('MIG19_2_TRIAGE')
    print('Source SHA-256 unchanged:',before == sha256(workbook))
    print('Findings by code:',dict(bycode))
    print('Revers-specific findings:',len(revers_findings), dict(Counter(key(x['code']) for x in revers_findings)))
    print('Mirror same-ID rows:',len(mirror_rows))
    print('Mirror by document type:',dict(by_doc))
    print('Mirror by direction:',dict(by_dir))
    print('Mirror by movement:',dict(by_kind))
    print('Mirror by station:',dict(by_station))
    print('Negative station balances before combining:',diag['station_types_negative_raw'])
    print('Negative combined balances (all stations):',diag['station_types_negative_combined'])
    print('Negative combined balances (excluding ST-00001):',diag['station_types_negative_combined_outside_central'])
    print('StationID, Packaging, StanicaSaldo, MirrorVozacSaldo, CombinedPhysicalCandidate')
    for v in combined:
        print(', '.join(v[k] for k in ('StationID','TipAmbalaze','StanicaSaldo','MirrorVozacSaldo','CombinedPhysicalCandidate')))
    print('This is diagnosis only, not a migration patch or historical PASS.')
    if args.save:
        # Files are written only when explicitly requested and never to the workbook.
        target = output / 'mig19_2_triage_summary.json'
        target.write_text(jdump(diag),encoding='utf-8')
        print('Wrote',target)
    return diag


def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--audit-out',required=True,type=Path,help='MIG-19.1 --out-dir')
    p.add_argument('--workbook',type=Path,help='Override workbook path embedded in MIG19 summary.json')
    p.add_argument('--save',action='store_true',help='Write JSON summary into audit output directory (not workbook)')
    args=p.parse_args()
    prepare(args)

if __name__=='__main__':
    main()
