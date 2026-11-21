#!/usr/bin/env python3
"""MIG-19.8 — multi-event OTKUP reconciliation, frozen inputs, no workbook writes.

What MIG-19.7 could not do: compare whole legacy source-document transfer against
ALL canonical non-cover movements (one to many), without assuming one physical row.
Never infer missing partner, never auto-correct a quantity, never modify XLSM.
"""
from __future__ import annotations
import argparse, json, sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path
from agrix_mig19_3_parity import (
    STAGE_SHA, XlsmReadOnly, stage_rows, sha256, key, dec,
    number, norm_type, date, void, write_csv, ledger_views, reconcile,
)
from agrix_mig19_4_difference_isolation import PRIOR_TARGET_SHA
from agrix_mig19_6_precise_manifest import SEVEN, READ_COLUMNS, doc_group, classify_doc
from agrix_mig19_7_verified_replacement_plan import orig_pair

ZERO=Decimal(0)
DOC_COLS=['Datum','TipAmbalaze','DokumentTIP','DokumentID','SourceLegRows','SourceAmbIDs',
    'LegacyFrom','LegacyTo','LegacyQuantity','TargetNonCoverRows','TargetCoverRows',
    'TargetKinds','TargetEndpoints','TargetAmbIDs','SourcePairProven','Status',
    'OldFromSigned','OldToSigned','TargetFromSigned','TargetToSigned',
    'FromDelta','ToDelta','ExactSourceToTargetRows','ExactSourceToTargetQtyTotal',
    'ReversedTargetRows','ReversedTargetQtyTotal','TargetSamePairOnly',
    'SameDirectionSplitEqualsOriginal','VectorCompletelyEqual','CanonicalTouchesOtherAccounts']
LEDGER_COLS=['Datum','TipAmbalaze','DokumentTIP','DokumentID','AmbID','Ordinal',
    'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','Kolicina','VrstaKretanja','StornoOd','IsCover']


def group_rows(old,new,cutoff='2026-08-31'):
    olds=defaultdict(list); news=defaultdict(list); covers=defaultdict(list)
    for i,r in enumerate(old,1):
        if void(r.get('Stornirano')):continue
        g=doc_group(r,i,'old')
        if g[0]<=cutoff and g[2]=='OTKUP':olds[g].append((i,r))
    for i,r in enumerate(new,1):
        g=doc_group(r,i,'new')
        if g[0]>cutoff or g[2]!='OTKUP':continue
        if key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE':covers[g].append((i,r))
        else:news[g].append((i,r))
    return olds,news,covers


def signed_source(row):
    amt=number(row.get('Kolicina'),'source OTKUP')
    sm=key(row.get('Smer'))
    if sm not in ('ULAZ','IZLAZ'):raise ValueError('Unknown legacy OTKUP direction '+sm)
    return amt if sm=='ULAZ' else -amt


def evaluate(g, olditems,newitems,coveritems):
    """All canonical rows in one document accounted for; no lossy 1:1 heuristic."""
    originals=[r for _,r in olditems]
    pair,reason=orig_pair(originals,g)
    if not pair:
        # Keep diagnostically visible rather than assuming source evidence.
        return {'Datum':g[0], 'TipAmbalaze':g[1],'DokumentTIP':g[2], 'DokumentID':g[3],
            'SourceLegRows':len(olditems),'SourceAmbIDs':'|'.join(key(r.get('AmbID')) for _,r in olditems),
            'LegacyFrom':'','LegacyTo':'','LegacyQuantity':'','TargetNonCoverRows':len(newitems),
            'TargetCoverRows':len(coveritems),'TargetKinds':'','TargetEndpoints':'',
            'TargetAmbIDs':'|'.join(key(r.get('AmbID')) for _,r in newitems),
            'SourcePairProven':False,'Status':'SOURCE_PAIR_NOT_PROVEN:'+reason,
            **{c:'' for c in DOC_COLS[16:]}}
    src,dst,q=pair
    srcs=f'{src[0]}/{src[1]}';dsts=f'{dst[0]}/{dst[1]}'
    oldvec=defaultdict(Decimal); newvec=defaultdict(Decimal)
    target_edges=Counter(); kinds=Counter(); same_qty=ZERO; inverse_qty=ZERO
    same_count=inv_count=0
    canonical_other=set()
    for _,r in olditems:oldvec[(key(r.get('EntitetTip')),key(r.get('EntitetID')))]+=signed_source(r)
    for _,r in newitems:
        a=(key(r.get('OdNalogTip')),key(r.get('OdNalogID')))
        b=(key(r.get('NaNalogTip')),key(r.get('NaNalogID')))
        n=number(r.get('Kolicina'),'target OTKUP',positive=True)
        newvec[a]-=n;newvec[b]+=n
        kinds[key(r.get('VrstaKretanja'))]+=1
        target_edges[f'{a[0]}->{b[0]}']+=1
        if a==src and b==dst: same_qty+=n;same_count+=1
        if a==dst and b==src: inverse_qty+=n;inv_count+=1
        for party in (a,b):
            if party not in (src,dst):canonical_other.add(party)
    all_equal=all(oldvec[x]==newvec[x] for x in set(oldvec)|set(newvec))
    source_effect=newvec[src];dest_effect=newvec[dst]
    onlypair=not canonical_other
    splitok=onlypair and inv_count==0 and same_count==len(newitems) and same_qty==q
    if not newitems:status='TARGET_NO_ROWS'
    elif all_equal:status='FULL_VECTOR_EXACT'
    elif onlypair and inv_count>0:status='SAME_PAIR_WITH_REVERSE_COMPONENTS'
    elif onlypair:status='SAME_PAIR_QUANTITY_OR_DIRECTION_MISMATCH'
    else:status='EXTRA_CANONICAL_PARTICIPANTS'
    return {'Datum':g[0],'TipAmbalaze':g[1],'DokumentTIP':g[2],'DokumentID':g[3],
        'SourceLegRows':len(olditems),'SourceAmbIDs':'|'.join(key(r.get('AmbID')) for _,r in olditems),
        'LegacyFrom':srcs,'LegacyTo':dsts,'LegacyQuantity':dec(q),
        'TargetNonCoverRows':len(newitems),'TargetCoverRows':len(coveritems),
        'TargetKinds':'|'.join(f'{k}:{v}' for k,v in sorted(kinds.items())),
        'TargetEndpoints':'|'.join(f'{k}:{v}' for k,v in sorted(target_edges.items())),
        'TargetAmbIDs':'|'.join(key(r.get('AmbID')) for _,r in newitems),
        'SourcePairProven':True,'Status':status,
        'OldFromSigned':dec(oldvec[src]),'OldToSigned':dec(oldvec[dst]),
        'TargetFromSigned':dec(source_effect),'TargetToSigned':dec(dest_effect),
        'FromDelta':dec(source_effect-oldvec[src]),'ToDelta':dec(dest_effect-oldvec[dst]),
        'ExactSourceToTargetRows':same_count,'ExactSourceToTargetQtyTotal':dec(same_qty),
        'ReversedTargetRows':inv_count,'ReversedTargetQtyTotal':dec(inverse_qty),
        'TargetSamePairOnly':onlypair,'SameDirectionSplitEqualsOriginal':splitok,
        'VectorCompletelyEqual':all_equal,'CanonicalTouchesOtherAccounts':len(canonical_other)}


def analysis(old,new):
    oldg,newg,covers=group_rows(old,new)
    rows=[]; ledgers=[];shapes=Counter();kinds=Counter(); statuses=Counter(); pair_n=Counter()
    oldtot,newtot=Counter(),Counter()
    mismatch=0; exact=0; source_pair_good=0
    for g in sorted(set(oldg)|set(newg)):
        res=evaluate(g,oldg.get(g,[]),newg.get(g,[]),covers.get(g,[]))
        rows.append(res)
        oldtot[len(oldg.get(g,[]))]+=1
        newtot[len(newg.get(g,[]))]+=1
        if res.get('SourcePairProven'):source_pair_good+=1
        if res['Status']=='FULL_VECTOR_EXACT':exact+=1
        else:
            mismatch+=1
            shapes[(res['TargetNonCoverRows'],res['TargetCoverRows'],res['Status'])]+=1
            kinds[res['TargetKinds']]+=1
            pair_n[(res['SourceLegRows'],res['TargetNonCoverRows'])]+=1
        for idx,r in newg.get(g,[])+covers.get(g,[]):
            ledgers.append(dict(Datum=g[0],TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
                AmbID=key(r.get('AmbID')),Ordinal=idx,OdNalogTip=key(r.get('OdNalogTip')),
                OdNalogID=key(r.get('OdNalogID')),NaNalogTip=key(r.get('NaNalogTip')),
                NaNalogID=key(r.get('NaNalogID')),
                Kolicina=dec(number(r.get('Kolicina'),'new quantity',positive=True)),
                VrstaKretanja=key(r.get('VrstaKretanja')),
                StornoOd=key(r.get('StornoOd')),
                IsCover='YES' if key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE' else 'NO'))
    summary={'otkup_document_groups':len(rows),'whole_doc_vector_exact':exact,
       'whole_doc_vector_not_exact':mismatch,'source_two_leg_pair_proven':source_pair_good,
       'target_noncover_row_count_distribution':dict(sorted(newtot.items())),
       'source_leg_row_count_distribution':dict(sorted(oldtot.items())),
       'mismatched_source_target_row_count_distribution':[{'OldRows':a,'NewRows':b,'Docs':c}
           for (a,b),c in sorted(pair_n.items())],
       'mismatched_target_shapes':[{'NewRows':a,'CoverRows':b,'Status':c,'Docs':n}
           for (a,b,c),n in sorted(shapes.items())],
       'mismatched_movement_kind_signatures':[{'Signature':k,'Docs':v}
           for k,v in kinds.most_common(20)],
       'all_1136_legacy_pair_condition':sum(r['SourcePairProven'] for r in rows if r['Status']!='FULL_VECTOR_EXACT')==1136,
       'all_1136_more_than_one_noncover_target':sum(r['TargetNonCoverRows']!=1 for r in rows if r['Status']!='FULL_VECTOR_EXACT')==1136,
       'target_row_details':len(ledgers)}
    return rows,ledgers,summary


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage',required=True,type=Path)
    p.add_argument('--workbook',required=True,type=Path)
    p.add_argument('--out-dir',required=True,type=Path)
    args=p.parse_args(argv)
    source=(args.stage/'UNCONVERTED__tblAmbalaza.jsonl').resolve()
    target=args.workbook.resolve()
    for f in (source,target):
        if not f.is_file():p.error(f'Missing input {f}')
    if args.out_dir.resolve() in (source,target,source.parent,target.parent):p.error('Output location must be separate')
    pre={str(f):sha256(f) for f in (source,target)}
    if pre[str(source)]!=STAGE_SHA:raise RuntimeError('Frozen source staging SHA changed')
    if pre[str(target)]!=PRIOR_TARGET_SHA:raise RuntimeError('Frozen QA workbook SHA changed')
    old=stage_rows(source)
    r=XlsmReadOnly(target)
    try:new=r.read_table('tblAmbalaza',READ_COLUMNS)
    finally:r.close()
    # Exact parity checks re-run independently before showing statistics.
    x=ledger_views(old,new)
    totals,dates,docs=reconcile(*x[:6])
    if (sum(z['Status']=='MISMATCH' for z in totals),sum(z['Status']=='MISMATCH' for z in dates),
            sum(z['Status']=='MISMATCH' for z in docs)) != (4,98,2273):
        raise RuntimeError('Unexpected parity baseline; do not report on changed inputs')
    rows,details,summary=analysis(old,new)
    if summary['whole_doc_vector_not_exact']!=1136 or not summary['all_1136_legacy_pair_condition']:
        raise RuntimeError('Unexpected OTKUP source-vector state; refuse frozen-case conclusion')
    if pre!={str(f):sha256(f) for f in (source,target)}:raise RuntimeError('Read-only hash changed')
    out=args.out_dir.resolve();out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG198_OTKUP_FULL_DOCUMENTS_PRIVATE.csv',rows,DOC_COLS)
    write_csv(out/'MIG198_OTKUP_CANONICAL_ROW_EVIDENCE_PRIVATE.csv',details,LEDGER_COLS)
    summary.update(format='AGRIx-MIG19.8-OTKUP-MULTIEVENT-READONLY-v1',
       status='SOURCE_GROUNDED_DOCUMENT_DECOMPOSITION_NO_EDIT',
       stage_sha_verified=True,target_sha_verified=True,inputs_unchanged=True,
       canonical_rows=len(new),legacy_rows=len(old),
       current_total_saldo_mismatches=4,current_daily_close_mismatches=98,
       current_document_account_slice_mismatches=2273,
       unresolved_seven_source_ambids=sorted(SEVEN),
       ready_to_modify_xlsm=False,
       note='Document-level matching requires ALL source legs and ALL non-cover canonical event rows. No silent rounding, balances adjusted, or inferred physical movement.')
    (out/'MIG198_SUMMARY.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print(f'MIG198_ERROR: {type(exc).__name__}: {exc}',file=sys.stderr)
        sys.exit(3)
