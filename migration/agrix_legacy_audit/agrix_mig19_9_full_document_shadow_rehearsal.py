#!/usr/bin/env python3
"""MIG-19.9: original-source-faithful OTKUP whole-document shadow rehearsal (read only).

Never modifies source JSONL, XLSM or Git. No deployable AMB-10 ledger writes.
Replaces complete canonical non-cover document groups IN MEMORY with abstract
source-signed 2-leg vectors. Simulated rows carry an intentionally INVALID
movement kind so they CANNOT be mistaken for an AMB-10 writer manifest.

Scope is narrow: 1136 divergent OTKUP document/date/pack groups, not six OMu
exceptions, not archived historical cancellations, not source->new event lineage.
"""
from __future__ import annotations
import argparse
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path
from agrix_mig19_3_parity import (
    STAGE_SHA, XlsmReadOnly, stage_rows, sha256, key, dec, number,
    ledger_views, reconcile, write_csv
)
from agrix_mig19_4_difference_isolation import PRIOR_TARGET_SHA
from agrix_mig19_8_multi_event_otkup import group_rows, evaluate, DOC_COLS, ZERO
from agrix_mig19_7_verified_replacement_plan import orig_pair, kind_valid
from agrix_mig19_6_precise_manifest import READ_COLUMNS, SEVEN, doc_group

GROUP_FIELDS=[
    'Datum','TipAmbalaze','DokumentTIP','DokumentID','Status','SourceAmbIDs',
    'SourceFrom','SourceTo','SourceQuantity', 'CanonicalNonCoverRows',
    'CanonicalCoverRows','CanonicalAmbIDs','ForwardRows','ForwardQuantity',
    'ReverseRows','ReverseQuantity','CanonicalSignedNetSourceToDest',
    'CanonicalMinusLegacyOnSource','CanonicalMinusLegacyOnDest',
    'ExactForwardAnchorIDs','SameDirectionCompatibleAnchorIDs',
    'AnyStornoLinkage','AnyCover','OtherCanonicalParties',
    'SourceVectorRehearsalEligible', 'Reason',
]
ACTION_FIELDS=['Datum','TipAmbalaze','DokumentTIP','DokumentID',
    'CanonicalAmbID','CanonicalVrstaKretanja','CanonicalOd','CanonicalNa',
    'CanonicalQuantity','CanonicalStornoOd','ExistingDirectionRelativeToSource',
    'PreviewOnly','Message']


def doc_reconstruction_plan(old,new):
    """Return descriptive manifest and protected group keys, with NO writer instructions."""
    oldg,newg,covers=group_rows(old,new)
    refs={key(r.get('StornoOd')) for r in new if key(r.get('StornoOd'))}
    groups=[]; actions=[]; replay_keys=set(); statuses=Counter();shapes=Counter();reasons=Counter()
    has_exact_anchor=has_same_kind_anchor=has_external=has_storno=has_cover=0
    total_forward=total_reverse=ZERO
    for g in sorted(set(oldg)|set(newg)):
        srcitems=oldg.get(g,[]); tgtitems=newg.get(g,[]);cvritems=covers.get(g,[])
        res=evaluate(g,srcitems,tgtitems,cvritems)
        if res.get('Status')=='FULL_VECTOR_EXACT':continue
        p,why=orig_pair([r for _,r in srcitems],g)
        if not p: raise ValueError(f'Legacy pair not proven {g}: {why}')
        src,dst,q=p
        same=opposite=ZERO;fcount=ocount=0
        exact_anchor=[];same_kind=[];others=set()
        st_link=False
        ids=[]
        for idx,r in tgtitems:
            oid=key(r.get('AmbID'));ids.append(oid)
            a=(key(r.get('OdNalogTip')),key(r.get('OdNalogID')))
            b=(key(r.get('NaNalogTip')),key(r.get('NaNalogID')))
            qty=number(r.get('Kolicina'),f'canonical OTKUP row {idx}',positive=True)
            if a==src and b==dst:
                direction='SOURCE_DIRECTION';same+=qty;fcount+=1
                if kind_valid(key(r.get('VrstaKretanja')),src,dst):
                    same_kind.append(oid)
                    if qty==q:exact_anchor.append(oid)
            elif a==dst and b==src:
                direction='REVERSE_DIRECTION';opposite+=qty;ocount+=1
            else:direction='OTHER_ENDPOINTS';others.update({a,b}-{src,dst})
            if key(r.get('StornoOd')) or oid in refs:st_link=True
            actions.append(dict(Datum=g[0],TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
                CanonicalAmbID=oid,CanonicalVrstaKretanja=key(r.get('VrstaKretanja')),
                CanonicalOd=f'{a[0]}/{a[1]}',CanonicalNa=f'{b[0]}/{b[1]}',
                CanonicalQuantity=dec(qty),CanonicalStornoOd=key(r.get('StornoOd')),
                ExistingDirectionRelativeToSource=direction,PreviewOnly=True,
                Message='DOCUMENT_GROUP_RECONSTRUCTION_EVIDENCE_ONLY_NO_DELETE_OR_WRITE'))
        for _,r in cvritems:
            aid=key(r.get('AmbID'))
            if key(r.get('StornoOd')) or aid in refs:st_link=True
        if not tgtitems: raise ValueError(f'No canonical OTKUP rows for {g}')
        if same+opposite==0:raise ValueError(f'No source-pair target match at {g}')
        if any(key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE' for _,r in tgtitems):
            raise ValueError(f'Derived cover improperly in noncover target group {g}')
        # We can always compare OLD 2-leg signed vector against a shadow transfer;
        # that does NOT imply writing the shadow event is legal under AMB-10.
        eligible=not others and len(srcitems)==2 and len(tgtitems) in (2,3)
        if not eligible:raise ValueError(f'Unexpected group shape; cannot rehearse {g}')
        replay_keys.add(g)
        statuses[res['Status']]+=1;shapes[(fcount,ocount,len(cvritems))]+=1
        total_forward+=same;total_reverse+=opposite
        if exact_anchor:has_exact_anchor+=1
        if same_kind:has_same_kind_anchor+=1
        if others:has_external+=1
        if st_link:has_storno+=1
        if cvritems:has_cover+=1
        if fcount==0:reason='NO_FORWARD_COMPONENT'
        elif same==q and opposite:reason='SOURCE_EXACT_FORWARD_PLUS_REVERSE'
        elif same>q and opposite:reason='FORWARD_EXCEEDS_SOURCE_WITH_REVERSE'
        elif same<q and opposite:reason='FORWARD_BELOW_SOURCE_WITH_REVERSE'
        else:reason='OTHER_MULTI_EVENT_NET_MISMATCH'
        reasons[reason]+=1
        groups.append(dict(Datum=g[0],TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
            Status=res['Status'],SourceAmbIDs='|'.join(key(r.get('AmbID')) for _,r in srcitems),
            SourceFrom=f'{src[0]}/{src[1]}',SourceTo=f'{dst[0]}/{dst[1]}',
            SourceQuantity=dec(q),CanonicalNonCoverRows=len(tgtitems),CanonicalCoverRows=len(cvritems),
            CanonicalAmbIDs='|'.join(ids),ForwardRows=fcount,ForwardQuantity=dec(same),
            ReverseRows=ocount,ReverseQuantity=dec(opposite),
            CanonicalSignedNetSourceToDest=dec(same-opposite),
            CanonicalMinusLegacyOnSource=dec(q-(same-opposite)),
            CanonicalMinusLegacyOnDest=dec((same-opposite)-q),
            ExactForwardAnchorIDs='|'.join(exact_anchor),
            SameDirectionCompatibleAnchorIDs='|'.join(same_kind),
            AnyStornoLinkage=st_link,AnyCover=bool(cvritems),OtherCanonicalParties=len(others),
            SourceVectorRehearsalEligible=eligible,Reason=reason))
    return groups,actions,replay_keys,dict(
        divergent_docs=len(groups),counts_by_reason=dict(sorted(reasons.items())),
        counts_by_component_shape=[dict(ForwardRows=k[0],ReverseRows=k[1],CoverRows=k[2],Docs=v)
            for k,v in sorted(shapes.items())],
        exact_same_direction_anchor_count=has_exact_anchor,
        any_same_direction_compatible_anchor_count=has_same_kind_anchor,
        docs_with_storno_dependency=has_storno,docs_with_derived_cover=has_cover,
        docs_with_unexpected_external_parties=has_external,
        sum_forward_quantity=dec(total_forward),sum_reverse_quantity=dec(total_reverse))


def full_source_projection_shadow(old,new,keys):
    """Remove all selected non-cover target rows ONLY IN MEMORY; add 2-leg source vector.

    Records are deliberately noncanonical. No output is a deployable data patch.
    """
    oldg,newg,covers=group_rows(old,new)
    keep=[];excluded=0
    for i,r in enumerate(new,1):
        g=doc_group(r,i,'new#')
        if g in keys and key(r.get('VrstaKretanja'))!='ULAZ_TUDJE_AMBALAZE':
            excluded+=1;continue
        keep.append(dict(r))
    constructed=[]
    for i,g in enumerate(sorted(keys),1):
        pair,reason=orig_pair([r for _,r in oldg[g]],g)
        if not pair:raise ValueError(f'Unexpected source pair {g}: {reason}')
        src,dst,q=pair
        constructed.append(dict(AmbID=f'__MIG199_SHADOW_ONLY__{i}',Datum=g[0],
            TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
            OdNalogTip=src[0],OdNalogID=src[1],NaNalogTip=dst[0],NaNalogID=dst[1],
            Kolicina=dec(q),VrstaKretanja='__INVALID_SHADOW_KIND_NOT_AMB10__',StornoOd=''))
    return keep+constructed,excluded,len(constructed)


def count_parity(old,new):
    views=ledger_views(old,new)
    ac,days,docs=reconcile(*views[:6])
    return dict(account_type_mismatches=sum(x['Status']!='PASS' for x in ac),
                cooperant_balance_mismatches=sum(x['Status']!='PASS' and x['NalogTip']=='KOOPERANT' for x in ac),
                station_balance_mismatches=sum(x['Status']!='PASS' and x['NalogTip']=='STANICA' for x in ac),
                daily_close_mismatches=sum(x['Status']!='PASS' for x in days),
                document_account_slice_mismatches=sum(x['Status']!='PASS' for x in docs),
                otkup_document_account_slice_mismatches=sum(x['Status']!='PASS' and x['DokumentTIP']=='OTKUP' for x in docs),
                otpremnica_document_account_slice_mismatches=sum(x['Status']!='PASS' and x['DokumentTIP']=='OTPREMNICA' for x in docs))


def main(argv=None):
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--stage',required=True,type=Path)
    ap.add_argument('--workbook',required=True,type=Path)
    ap.add_argument('--out-dir',required=True,type=Path)
    args=ap.parse_args(argv)
    src=(args.stage/'UNCONVERTED__tblAmbalaza.jsonl').resolve()
    qa=args.workbook.resolve()
    if not src.is_file() or not qa.is_file():ap.error('Missing source staging or QA workbook')
    protected=(src,qa)
    out=args.out_dir.resolve()
    if out in set(protected)|{p.parent for p in protected}:ap.error('Output must be in a separate directory')
    checks={str(f):sha256(f) for f in protected}
    if checks[str(src)]!=STAGE_SHA:raise RuntimeError('Frozen legacy staging SHA does not match')
    if checks[str(qa)]!=PRIOR_TARGET_SHA:raise RuntimeError('Frozen QA workbook SHA does not match')
    old=stage_rows(src)
    reader=XlsmReadOnly(qa)
    try:new=reader.read_table('tblAmbalaza',READ_COLUMNS)
    finally:reader.close()
    before=count_parity(old,new)
    if before!={'account_type_mismatches':4,'cooperant_balance_mismatches':0,
               'station_balance_mismatches':4,'daily_close_mismatches':98,
               'document_account_slice_mismatches':2273,
               'otkup_document_account_slice_mismatches':2272,
               'otpremnica_document_account_slice_mismatches':1}:
        raise RuntimeError('Unexpected frozen parity baseline. No manifest emitted')
    groups,actions,keys,breakdown=doc_reconstruction_plan(old,new)
    if breakdown['divergent_docs']!=1136 or len(keys)!=1136:
        raise RuntimeError('Unexpected divergent OTKUP group count')
    projected,excluded,added=full_source_projection_shadow(old,new,keys)
    after=count_parity(old,projected)
    if after['otkup_document_account_slice_mismatches']!=0:
        raise RuntimeError('In-memory shadow failed source OTKUP document parity')
    if checks!={str(f):sha256(f) for f in protected}:
        raise RuntimeError('A source or workbook file changed during run')
    out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG199_DOCUMENT_RECONSTRUCTION_PRIVATE.csv',groups,GROUP_FIELDS)
    write_csv(out/'MIG199_EXISTING_CANONICAL_ROWS_PRIVATE.csv',actions,ACTION_FIELDS)
    report={
       'format':'AGRIx-MIG19.9-OTKUP-FULL-DOC-SHADOW-REHEARSAL-v1',
       'status':'SHADOW_REHEARSAL_ONLY_INVALID_CANONICAL_ROWS_NO_XLSM_WRITE',
       'source_stage_sha_verified':True,'target_sha_verified':True,
       'target_sha256':checks[str(qa)],'inputs_unchanged':True,
       'legacy_rows':len(old),'canonical_rows_original':len(new),
       'otkup_groups':1625,'untouched_exact_otkup_groups':489,
       'reconstructed_shadow_only_otkup_groups':len(keys),
       'canonical_noncover_rows_removed_in_shadow_only':excluded,
       'noncanonical_shadow_rows_added_in_memory_only':added,
       'component_analysis':breakdown,
       'before':before,
       'after_source_vector_shadow_only':after,
       'seven_legacy_exceptions_not_modified':sorted(SEVEN),
       'historical_cancelled_source_rows_not_replayed':335,
       'event_level_lineage_proven':False,
       'canon_kind_in_shadow_intentionally_invalid':True,
       'xlsm_write_authorized':False,
       'next_decision':'Review entire document-group writer against original SOURCE and AMB-10 coverage/storno invariants, not individual row patch. Seven independent legacy effects still require their own source-grounded mapping.',
    }
    (out/'MIG199_SUMMARY.json').write_text(json.dumps(report,indent=2,ensure_ascii=False),encoding='utf-8')
    print(json.dumps(report,indent=2,ensure_ascii=False))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print(f'MIG199_ERROR {type(exc).__name__}: {exc}',file=sys.stderr)
        sys.exit(3)
