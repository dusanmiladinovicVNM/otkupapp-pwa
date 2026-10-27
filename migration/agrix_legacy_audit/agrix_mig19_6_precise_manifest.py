#!/usr/bin/env python3
"""MIG-19.6: exact SOURCE-vs-TARGET document forensic + 7-event migration manifest.

READ ONLY, no recommendations to invent counterparties or silently change AMB10.
Requires locally installed MIG19.3/4/5 scripts and MIG19.1 OOXML reader.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from collections import defaultdict, Counter
from decimal import Decimal
from pathlib import Path
import sys

from agrix_mig19_3_parity import (
    XlsmReadOnly, stage_rows, ledger_views, reconcile, number, key, norm_type,
    date, void, dec, sha256, write_csv, STAGE_SHA,
)
from agrix_mig19_4_difference_isolation import gate_sum, day_isolation, PRIOR_TARGET_SHA
from agrix_mig19_5_evidence_and_parity import PACK, EXPECTED_STATION_GAPS, old_signed, endpoint

SEVEN = {
    'AMB-00700': ('ST-00004', '2026-06-22', Decimal('-500')),
    'AMB-01377': ('ST-00001', '2026-06-23', Decimal('5')),
    'AMB-02317': ('ST-00005', '2026-06-27', Decimal('300')),
    'AMB-03849': ('ST-00005', '2026-06-28', Decimal('250')),
    'AMB-03939': ('ST-00002', '2026-06-29', Decimal('83')),
    'AMB-07115': ('ST-00001', '2026-07-14', Decimal('10')),
    'AMB-07815': ('ST-00001', '2026-07-23', Decimal('3')),
}
READ_COLUMNS = {'AmbID','Datum','TipAmbalaze','Kolicina','OdNalogTip','OdNalogID',
                'NaNalogTip','NaNalogID','VrstaKretanja','StornoOd','DokumentTIP','DokumentID'}


def doc_group(row, i, where):
    return (date(row.get('Datum'),where+str(i)), norm_type(row.get('TipAmbalaze')),
            key(row.get('DokumentTIP')), key(row.get('DokumentID')))


def grouped_vectors(old_rows, new_rows, cutoff='2026-08-31'):
    """Legacy and canonical comparable per-document account contributions.

    Both representations retain transport provenance by grouping on original
    document key (day, packaging, doc type, doc ID). Cover rows are tracked
    separately, never confused with canonical historical projection.
    """
    leg=defaultdict(lambda: defaultdict(Decimal))
    new=defaultdict(lambda: defaultdict(Decimal))
    cnt_old=Counter(); cnt_new=Counter(); cnt_cover=Counter()
    old_detail=defaultdict(list); new_detail=defaultdict(list)
    for i,r in enumerate(old_rows,1):
        if void(r.get('Stornirano')):continue
        group=doc_group(r,i,'legacy#')
        if group[0]>cutoff:continue
        acct=(key(r.get('EntitetTip')),key(r.get('EntitetID')))
        qty=old_signed(r)
        leg[group][acct]+=qty
        cnt_old[group]+=1
        if key(r.get('AmbID')) in SEVEN:
            old_detail[group].append({'AmbID':key(r.get('AmbID')),'Account':f'{acct[0]}/{acct[1]}',
                'Signed':dec(qty),'LegacyOrdinal':i,'VozacID':key(r.get('VozacID'))})
    for i,r in enumerate(new_rows,1):
        group=doc_group(r,i,'canonical#')
        if group[0]>cutoff:continue
        q=number(r.get('Kolicina'),f'canonical#{i}',positive=True)
        if key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE':
            cnt_cover[group]+=1
            continue
        a,b=endpoint(r,'Od'),endpoint(r,'Na')
        new[group][a]-=q;new[group][b]+=q
        cnt_new[group]+=1
        new_detail[group].append({'AmbID':key(r.get('AmbID')),
            'Od':f'{a[0]}/{a[1]}','Na':f'{b[0]}/{b[1]}','Qty':dec(q),
            'VrstaKretanja':key(r.get('VrstaKretanja')),
            'CanonicalOrdinal':i,'StornoOd':key(r.get('StornoOd'))})
    return leg,new,cnt_old,cnt_new,cnt_cover,old_detail,new_detail


def classify_doc(old_vec,new_vec):
    touched={k for k in set(old_vec)|set(new_vec) if k[0] in ('KOOPERANT','STANICA')}
    if all(old_vec.get(k,0)==new_vec.get(k,0) for k in touched):return 'EXACT'
    old_present={k for k in touched if old_vec.get(k,0)!=0}
    new_present={k for k in touched if new_vec.get(k,0)!=0}
    if old_present!=new_present:return 'PARTICIPANTS_DIFFER'
    if old_present and all(new_vec.get(k,0)==-old_vec.get(k,0) for k in old_present):return 'EXACT_DIRECTION_REVERSED'
    if old_present and all(old_vec.get(k,0)*new_vec.get(k,0)>0 for k in old_present):return 'SAME_SIGN_DIFFERENT_MAGNITUDE'
    return 'MIXED_OR_COMPLEX'


def docs_audit(leg,new,cl,cn,cc):
    hist=Counter(); by_kind=Counter(); mismatch=[]; netzero=0
    allk=set(leg)|set(new)
    for g in sorted(allk):
        if g[2] not in ('OTKUP','OTPREMNICA','PRIJEMNICA'):continue
        a,b=leg.get(g,{}),new.get(g,{})
        typ=classify_doc(a,b)
        hist[(g[2],typ)]+=1
        if typ=='EXACT':continue
        touched=sorted({k for k in set(a)|set(b) if k[0] in ('KOOPERANT','STANICA')})
        delta=sum((b.get(k,Decimal(0))-a.get(k,Decimal(0)) for k in touched),Decimal(0))
        if delta==0:netzero+=1
        parts=' ; '.join(f'{k[0]}/{k[1]}:old={dec(a.get(k,Decimal(0)))},new={dec(b.get(k,Decimal(0)))}'
                         for k in touched)
        mismatch.append(dict(Datum=g[0],TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
           Classification=typ,OldRows=cl[g],NewNonCoverRows=cn[g],CoverRows=cc[g],
           ParticipantCounts=len(touched),NetDelta=dec(delta),ParticipantsWithEffects=parts))
        by_kind[(g[2],typ)]+=1
    return mismatch,hist,netzero


def seven_manifest(oldrows,newrows,legs,news,old_details,new_details,change_pts):
    byid=defaultdict(list)
    for ordinal,r in enumerate(oldrows,1):byid[key(r.get('AmbID'))].append((ordinal,r))
    change_index={(r['NalogID'],r['Datum'],r['TipAmbalaze']):Decimal(r['PromenaRazlike']) for r in change_pts}
    output=[]
    for aid,(st,day,expected) in SEVEN.items():
        entries=byid[aid]
        if len(entries)!=1:
            raise ValueError(f'Legacy ID {aid} appears {len(entries)} times; refuse manifest')
        ord,r=entries[0];v=old_signed(r)
        group=doc_group(r,ord,'legacy#')
        if (group[0],group[1],v,key(r.get('EntitetTip')),key(r.get('EntitetID'))) != (
              day,PACK,expected,'STANICA',st):
            raise ValueError('Pinned source signature mismatch for '+aid)
        if void(r.get('Stornirano')):raise ValueError('Source record is cancelled: '+aid)
        # Isolate exact canonical rows carrying this SOURCE document identity,
        # never match by date/quantity alone.
        same_doc=new_details.get(group,[])
        projected_effect=sum((news.get(group,{}).get(('STANICA',st),Decimal(0)),),Decimal(0))
        legacy_effect=legs.get(group,{}).get(('STANICA',st),Decimal(0))
        matched_daily_delta=change_index.get((st,day,PACK),Decimal(0))
        # Indispensable condition: this date's aggregate discrepancy is exactly
        # the effect of the pinned old row, but document co-occurrence is still needed.
        isolated = matched_daily_delta == -expected
        existing_exact=[x for x in same_doc if
          (x['Na']==f'STANICA/{st}' if expected>0 else x['Od']==f'STANICA/{st}') and
          Decimal(x['Qty'])==abs(expected)]
        output.append(dict(SourceAmbID=aid,LegacyOrdinal=ord,Datum=day,TipAmbalaze=PACK,
           Stanica=st,LegacySignedEffect=dec(expected),SourceDokumentTIP=group[2],
           SourceDokumentID=group[3],SourceVozacID=key(r.get('VozacID')),
           EntireSourceDocStationEffect=dec(legacy_effect),
           EntireCanonicalDocStationEffect=dec(projected_effect),
           SourceDocCanonicalNonCoverRows=len(same_doc),
           ExactSameDocumentEndpointQtyCandidates=len(existing_exact),
           CandidateCanonicalIDs=' | '.join(x['AmbID'] for x in existing_exact),
           AggregateDailyGapEqualsSourceEffect=isolated,
           CorrectionSign=dec(expected),
           Action='RECONSTRUCT_FROM_SOURCE_WITH_PROVEN_CANONICAL_COUNTERPARTY_AND_DOC_MAPPING',
           ProofLevel='SOURCE_AND_DAILY_VECTOR_ONLY_NOT_CANONICAL_EVENT_LINEAGE'))
    return output


def main(argv=None):
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--stage',type=Path,required=True)
    ap.add_argument('--workbook',type=Path,required=True)
    ap.add_argument('--out-dir',type=Path,required=True)
    ap.add_argument('--expected-target-sha',default=PRIOR_TARGET_SHA)
    args=ap.parse_args(argv)
    source=args.stage/'UNCONVERTED__tblAmbalaza.jsonl'
    if not source.is_file() or not args.workbook.is_file():ap.error('Missing source staging or QA workbook')
    protected=(source.resolve(),args.workbook.resolve())
    if args.out_dir.resolve() in set(protected)|{p.parent for p in protected}:ap.error('Invalid output location')
    before={str(p):sha256(p) for p in protected}
    if before[str(source.resolve())]!=STAGE_SHA:raise RuntimeError('Staging SHA mismatch')
    if before[str(args.workbook.resolve())]!=args.expected_target_sha:raise RuntimeError('Target workbook SHA mismatch')
    olds=stage_rows(source);wb=XlsmReadOnly(args.workbook)
    try:news=wb.read_table('tblAmbalaza',READ_COLUMNS)
    finally:wb.close()
    old,phy,projected,cover,olddoc,newdoc,diag,af=ledger_views(olds,news)
    acct,dates,docs=reconcile(old,phy,projected,cover,olddoc,newdoc)
    gaps,_=gate_sum(acct)
    gm={(r['NalogID'],r['TipAmbalaze']):Decimal(r['Razlika']) for r in gaps}
    if gm!=EXPECTED_STATION_GAPS:raise RuntimeError('Unexpected account gaps. Refuse to issue frozen manifest.')
    changes,_,_=day_isolation(old,projected,gaps)
    if len(changes)!=7:raise RuntimeError('Unexpected number of change dates')
    sourcevec,targetvec,cs,ct,cc,old_d,new_d=grouped_vectors(olds,news)
    seven=seven_manifest(olds,news,sourcevec,targetvec,old_d,new_d,changes)
    mismatch,hist,nzero=docs_audit(sourcevec,targetvec,cs,ct,cc)
    if before!={str(p):sha256(p) for p in protected}:raise RuntimeError('Inputs were altered')
    out=args.out_dir.resolve();out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG196_SEVEN_SOURCE_RECONSTRUCTION_PRIVATE.csv',seven,
        ['SourceAmbID','LegacyOrdinal','Datum','TipAmbalaze','Stanica','LegacySignedEffect',
         'SourceDokumentTIP','SourceDokumentID','SourceVozacID','EntireSourceDocStationEffect',
         'EntireCanonicalDocStationEffect','SourceDocCanonicalNonCoverRows',
         'ExactSameDocumentEndpointQtyCandidates','CandidateCanonicalIDs',
         'AggregateDailyGapEqualsSourceEffect','CorrectionSign','Action','ProofLevel'])
    write_csv(out/'MIG196_DOCUMENT_VECTOR_DIFFERENCES_PRIVATE.csv',mismatch,
        ['Datum','TipAmbalaze','DokumentTIP','DokumentID','Classification',
         'OldRows','NewNonCoverRows','CoverRows','ParticipantCounts','NetDelta',
         'ParticipantsWithEffects'])
    summary={
      'format':'AGRIx-MIG19.6-CORRECTION-MANIFEST-READONLY-v1',
      'status':'READ_ONLY_NO_CORRECTIONS_APPLIED',
      'stage_sha_verified':True,'target_sha256':before[str(args.workbook.resolve())],
      'inputs_unchanged':True,
      'final_balance_mismatches':len(gaps),
      'daily_change_points':len(changes),
      'seven_source_records_verified':len(seven),
      'seven_each_daily_gap_exactly_matches_source':all(r['AggregateDailyGapEqualsSourceEffect'] for r in seven),
      'seven_same_document_exact_endpoint_qty_candidate_counts':{
          r['SourceAmbID']:r['ExactSameDocumentEndpointQtyCandidates'] for r in seven},
      'seven_same_document_noncover_row_counts':{
          r['SourceAmbID']:r['SourceDocCanonicalNonCoverRows'] for r in seven},
      'document_vector_classes':[dict(DokumentTIP=d,Class=k,Count=n)
         for (d,k),n in sorted(hist.items())],
      'document_vectors_with_nonzero_mismatch':len(mismatch),
      'document_mismatches_with_zero_net_effect':nzero,
      'event_lineage_proven':False,
      'ready_to_patch_xlsm':False,
      'no_write_applied':True,
      'next_gate':'Inspect replacement not append for 7; classify OTKUP document reversals and their original provenance',
    }
    (out/'MIG196_SUMMARY.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print(f'MIG196_ERROR: {type(exc).__name__}: {exc}',file=sys.stderr)
        sys.exit(3)
