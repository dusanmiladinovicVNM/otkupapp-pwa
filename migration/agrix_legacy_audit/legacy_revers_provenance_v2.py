#!/usr/bin/env python3
"""MIG-05.1: read-only audit of mixed legacy OM/Revers document identities.

The old workbook used OM-Izlaz-Koop for both otkup's KolAmbIzdata and a
standalone Revers; the latter stored business number in DokumentID and may
have ReversID separately. NEVER assume a document number equals an opaque PK.
No canonical AMB-10 entry, storno inverse, balance fix, or XLSM is produced.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_amb10_forensics_v2 import inputs, analyze as forensics, key, dec, private_csv, SOURCE_SHA

VERSION='MIG-05.1'
OM_TYPES={'om-izlaz-koop','om-ulaz-koop','om-izlaz-firma','om-ulaz-firma'}
PAIR_TYPES={'om-izlaz-koop','om-ulaz-koop'}
ROLE_MAP={
  'om-izlaz-koop':frozenset({('kooperant','ulaz'),('stanica','izlaz')}),
  'om-ulaz-koop':frozenset({('kooperant','izlaz'),('stanica','ulaz')}),
}
PRIVATE_FIELDS=['source_row_ordinal','AmbID','DokumentTip','DokumentID','ReversID','Datum',
                'TipAmbalaze','Kolicina','Smer','EntitetTip','EntitetID','VozacID','Stornirano',
                'forensic_flags','provenance','revers_pair_status','possible_peer_ordinal',
                'identity_warning','canonical_AMB10_import_authorized']


def revers_id(row):
    """Excel ListObject columns compare case-insensitively, Python dicts do not."""
    matches=[trim(v) for k,v in row.items() if k.casefold()=='reversid']
    if len(set(matches))>1:
        raise AuditError('Conflicting ReversID case-variant columns')
    return matches[0] if matches else ''


def classify(row, otkups):
    """One row's identity provenance. Never guess by business number."""
    typ=key(row.get('DokumentTip'))
    if typ not in OM_TYPES:raise AuditError('Not an OM row')
    doc=trim(row.get('DokumentID'))
    rev=revers_id(row)
    parent=otkups.get(doc) if doc and typ=='om-izlaz-koop' else None
    warnings=[]
    if parent:
        if rev:warnings.append('REVERS_ID_PRESENT_WITH_OTKUP_DOCUMENT_ID')
        checks=[
           ('KooperantID', 'Kooperant'),
           ('StanicaID', 'Stanica'),
        ]
        owner=key(row.get('EntitetTip'))
        expect=next((name for col,name in checks if key(name)==owner),None)
        if not expect or key(row.get('EntitetID'))!=key(parent[1].get(expect+'ID')):
            warnings.append('OTKUP_ACCOUNT_MISMATCH')
        if dec(row.get('Kolicina'))!=dec(parent[1].get('KolAmbIzdata')):
            warnings.append('OTKUP_QTY_MISMATCH')
        if key(row.get('TipAmbalaze'))!=key(parent[1].get('TipAmbalaze')):
            warnings.append('OTKUP_AMB_TYPE_MISMATCH')
        if is_active(row)!=is_active(parent[1]):
            warnings.append('OTKUP_STATE_MISMATCH')
        if key(row.get('Smer')) != ('ulaz' if owner=='kooperant' else 'izlaz'):
            warnings.append('OTKUP_DIRECTION_MISMATCH')
        if not warnings:return 'OTKUP_LINK_EXACT',rev,''
        return 'OTKUP_ID_PRESENT_CONTEXT_REVIEW',rev,';'.join(warnings)
    if rev:
        return 'REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF',rev,''
    if doc:
        return 'BUSINESS_LABEL_ONLY_NO_OPAQUE_ID',rev,''
    return 'MISSING_ID_AND_LABEL',rev,''


def revers_pairs(rows):
    """Two source legs exactly reciprocal under same ReversID & same event values.

    This is *physical evidence* only, never proof of business number ownership,
    legal title, opening inventory, or the new AMB-10 movement kind.
    """
    groups=defaultdict(list)
    for r in rows:
        if r['provenance']!='REVERS_ID_EVIDENCE_NOT_DOCUMENT_PROOF':continue
        typ=key(r.get('DokumentTip'))
        if typ not in PAIR_TYPES:continue
        sig=(key(r['ReversID']),typ,key(r.get('DokumentID')),key(r.get('Datum')),
             key(r.get('TipAmbalaze')),str(dec(r.get('Kolicina'))),is_active(r))
        groups[sig].append(r)
    statuses={};candidate_peers={}
    for items in groups.values():
        typ=key(items[0]['DokumentTip'])
        roles=frozenset((key(r.get('EntitetTip')),key(r.get('Smer'))) for r in items)
        if len(items)==2 and roles==ROLE_MAP[typ] and all(dec(r.get('Kolicina')) is not None and dec(r.get('Kolicina'))>0 for r in items):
            a,b=items
            statuses[a['source_row_ordinal']]='RECIPROCAL_WITHIN_REVERS_ID_EVIDENCE_ONLY'
            statuses[b['source_row_ordinal']]='RECIPROCAL_WITHIN_REVERS_ID_EVIDENCE_ONLY'
            candidate_peers[a['source_row_ordinal']]=b['source_row_ordinal']
            candidate_peers[b['source_row_ordinal']]=a['source_row_ordinal']
        else:
            st='AMBIGUOUS_REVERS_EVENT_GROUP' if len(items)>2 else 'NO_EXACT_REVERS_RECIPROCAL'
            for r in items: statuses[r['source_row_ordinal']]=st
    return statuses,candidate_peers


def run(stage,quarantine):
    manifest,tables,idx=inputs(stage)
    prior=json.loads((quarantine/'agrix_amb10_quarantine_summary.json').read_text(encoding='utf-8'))
    if prior.get('format')!='MIG-04.4' or prior.get('source_sha256')!=SOURCE_SHA or prior.get('schema_sha256')!=manifest.get('schema_sha256'):
        raise AuditError('MIG-04.4 baseline source/schema mismatch')
    if prior.get('quarantined_active_rows')!=98:
        raise AuditError('MIG-04.4 quarantine count drift')
    _,evidence=forensics(stage)
    raw=tables['UNCONVERTED__tblAmbalaza']
    if len(evidence)!=len(raw):raise AuditError('Forensics source count drift')
    otkups=idx['SOURCE__tblOtkup']
    om_rows=[];count_all=Counter();counts_q=Counter();quarantine_rows=Counter();label_count=0
    id_present=0
    for pos,(row,ev) in enumerate(zip(raw,evidence),1):
        if key(row.get('DokumentTip')) not in OM_TYPES:continue
        prov,rev,warning=classify(row,otkups)
        obj=dict(row)
        obj.update({'source_row_ordinal':pos,'ReversID':rev,'forensic_flags':ev.get('forensic_flags',''),
                    'provenance':prov,'revers_pair_status':'NOT_APPLICABLE',
                    'possible_peer_ordinal':'','identity_warning':warning,
                    'canonical_AMB10_import_authorized':'NO'})
        om_rows.append(obj)
        count_all[prov]+=1
        if rev:id_present+=1
        if trim(obj['DokumentID']):label_count+=1
        if trim(ev.get('forensic_flags')):
            counts_q[prov]+=1
            quarantine_rows[key(row.get('DokumentTip'))]+=1
    expected_types={key(k):v for k,v in prior['document_type_categories'].items() if key(k) in OM_TYPES}
    if sum(expected_types.values())!=sum(quarantine_rows.values()) or any(quarantine_rows.get(k,0)!=v for k,v in expected_types.items()):
        raise AuditError('OM quarantine classification drift')
    status,peers=revers_pairs(om_rows)
    for row in om_rows:
        i=row['source_row_ordinal']
        if i in status:row['revers_pair_status']=status[i]
        if i in peers:row['possible_peer_ordinal']=peers[i]
    rev_pair_q=sum(bool(r.get('forensic_flags')) and r.get('revers_pair_status')=='RECIPROCAL_WITHIN_REVERS_ID_EVIDENCE_ONLY' for r in om_rows)
    c=Counter(r['revers_pair_status'] for r in om_rows)
    report={
        'format':VERSION,
        'status':'MIXED_LEGACY_REVERS_IDENTITY_EVIDENCE_ONLY_NO_AMB10_IMPORT',
        'source_sha256':SOURCE_SHA,
        'schema_sha256':manifest.get('schema_sha256'),
        'source_amb_rows':len(raw),
        'om_rows_examined':len(om_rows),
        'om_quarantined_source_rows':sum(counts_q.values()),
        'quarantine_reference':prior['quarantined_active_rows'],
        'non_om_quarantine_rows':prior['quarantined_active_rows']-sum(counts_q.values()),
        'all_om_provenance':dict(sorted(count_all.items())),
        'quarantined_om_provenance':dict(sorted(counts_q.items())),
        'om_quarantine_type_counts':dict(sorted(quarantine_rows.items())),
        'om_rows_with_revers_id':id_present,
        'om_rows_without_revers_id':len(om_rows)-id_present,
        'revers_id_group_pair_status_all':dict(sorted(c.items())),
        'quarantined_om_legs_with_exact_revers_id_reciprocal':rev_pair_q,
        'om_rows_with_nonempty_document_id':label_count,
        'ready_for_import':False,
        'notes':[
            'Old OM-Izlaz-Koop served BOTH otkup and standalone Revers. DokumentID could be OtkupID or business label.',
            'Only exact OtkupID AND context agreement establishes an otkup-origin posting.',
            'Standalone Revers evidence uses ReversID where available; business number alone never treated as unique PK.',
            'An exact reciprocal ReversID pair proves legacy physical legs, NOT the business movement kind or packaging title.',
            'NO source or staging mutated; historical storno replay, title, opening state, AMB-10 entry and target XLSM untouched.',
        ],
    }
    return report,om_rows


def txt(report):
    keys=['source_sha256','source_amb_rows','om_rows_examined','om_quarantined_source_rows',
          'non_om_quarantine_rows','all_om_provenance','quarantined_om_provenance',
          'om_rows_with_revers_id','om_rows_without_revers_id','revers_id_group_pair_status_all',
          'quarantined_om_legs_with_exact_revers_id_reciprocal']
    return 'AGRIX MIG-05.1 - MIXED OLD REVERSE PROVENANCE / READ ONLY\n'+''.join(
        f'{k}: {json.dumps(report[k],sort_keys=True,ensure_ascii=False)}\n' for k in keys)+\
        'STOP: no source, AMB-10 ledger or XLSM changed. Share ONLY summaries.\n'


def cli(argv=None):
    ap=argparse.ArgumentParser(description='Read-only legacy OM/Revers ID provenance diagnostic')
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--quarantine',type=Path,required=True)
    ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args(argv)
    st,qa,out=(getattr(args,p).resolve() for p in ('stage','quarantine','out'))
    if not st.is_dir() or not qa.is_dir():raise AuditError('Stage/quarantine directories missing')
    if out.exists():raise AuditError('Output already exists: choose a NEW output folder')
    for d in (st,qa):
        if out==d or out in d.parents or d in out.parents:raise AuditError('Output overlaps input')
    report,rows=run(st,qa)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_revers_provenance_summary.json').write_text(json.dumps(report,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_revers_provenance_summary.txt').write_text(txt(report),encoding='utf-8')
        private_csv(out/'revers_provenance_PRIVATE.csv',rows,PRIVATE_FIELDS)
    except BaseException:
        for f in out.iterdir():
            if f.is_file():f.unlink()
        out.rmdir()
        raise
    print(txt(report));return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print('MIG-05.1 STOP:',exc,file=sys.stderr);raise SystemExit(2)
